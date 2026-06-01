"""Tests for the ``generate_excel`` Cloud Function.

These exercise the function through the Functions Framework's Flask test
client, which provides the request/response and application context that
``request.get_json`` and ``flask.jsonify`` rely on. The source under
``main.py`` is treated as a black box and is not modified.
"""

import base64
import io
import os

import pytest
from functions_framework import create_app
from openpyxl import load_workbook


SOURCE = os.path.join(os.path.dirname(__file__), "main.py")


@pytest.fixture
def client():
    app = create_app(target="generate_excel", source=SOURCE)
    return app.test_client()


def _decode_workbook(file_data):
    """Decode the base64 payload returned by the function into a workbook."""
    raw = base64.b64decode(file_data)
    return load_workbook(io.BytesIO(raw))


# --- CORS preflight ---------------------------------------------------------

def test_options_preflight_returns_cors_headers(client):
    resp = client.options("/")

    assert resp.status_code == 204
    assert resp.data == b""
    assert resp.headers["Access-Control-Allow-Origin"] == "*"
    assert resp.headers["Access-Control-Allow-Methods"] == "POST"
    assert resp.headers["Access-Control-Allow-Headers"] == "Content-Type"
    assert resp.headers["Access-Control-Max-Age"] == "3600"


# --- Successful generation --------------------------------------------------

def test_post_valid_data_returns_success_payload(client):
    payload = {
        "floor": "2",
        "room_name": "LDK",
        "x_tilt": "3",
        "y_tilt": "5",
    }

    resp = client.post("/", json=payload)

    assert resp.status_code == 200
    assert resp.headers["Access-Control-Allow-Origin"] == "*"

    body = resp.get_json()
    assert body["status"] == "success"
    assert body["fileName"] == "傾斜測定報告書.xlsx"
    assert body["fileData"]


def test_post_valid_data_writes_values_into_template(client):
    payload = {
        "floor": "2",
        "room_name": "LDK",
        "x_tilt": "3",
        "y_tilt": "5",
    }

    resp = client.post("/", json=payload)
    wb = _decode_workbook(resp.get_json()["fileData"])
    ws = wb["傾斜測定"]

    assert ws["A15"].value == "2"
    assert ws["C15"].value == "LDK"
    assert ws["L15"].value == "3"
    assert ws["L17"].value == "5"


def test_returned_file_is_a_valid_xlsx(client):
    resp = client.post("/", json={"floor": "1", "room_name": "和室"})
    raw = base64.b64decode(resp.get_json()["fileData"])

    # XLSX files are zip archives and start with the PK signature.
    assert raw[:2] == b"PK"
    wb = _decode_workbook(resp.get_json()["fileData"])
    assert wb.sheetnames == ["傾斜測定"]


def test_missing_fields_default_to_empty(client):
    # Only ``floor`` is provided; the rest fall back to "" in the source, which
    # openpyxl stores as an empty (None) cell after the save round-trip.
    resp = client.post("/", json={"floor": "3"})
    ws = _decode_workbook(resp.get_json()["fileData"])["傾斜測定"]

    assert ws["A15"].value == "3"
    assert ws["C15"].value is None
    assert ws["L15"].value is None
    assert ws["L17"].value is None


# --- Bad / missing input ----------------------------------------------------

def test_post_without_body_returns_400(client):
    resp = client.post("/", data="", content_type="application/json")

    assert resp.status_code == 400
    assert resp.headers["Access-Control-Allow-Origin"] == "*"
    assert resp.get_json() == {"error": "No data provided"}


def test_post_empty_json_object_returns_400(client):
    # An empty object is falsy, so the function treats it as "no data".
    resp = client.post("/", json={})

    assert resp.status_code == 400
    assert resp.get_json() == {"error": "No data provided"}


def test_non_json_body_returns_400(client):
    # get_json(silent=True) yields None for unparseable bodies.
    resp = client.post("/", data="not-json", content_type="text/plain")

    assert resp.status_code == 400
    assert resp.get_json() == {"error": "No data provided"}


# --- Error handling ---------------------------------------------------------

def test_internal_error_returns_500_with_message(client, monkeypatch, tmp_path):
    # Run from a directory without ``template.xlsx`` so load_workbook raises,
    # exercising the ``except`` branch that returns a 500 with the error text.
    monkeypatch.chdir(tmp_path)

    resp = client.post("/", json={"floor": "1", "room_name": "LDK"})

    assert resp.status_code == 500
    assert resp.headers["Access-Control-Allow-Origin"] == "*"
    assert "template.xlsx" in resp.get_json()["error"]
