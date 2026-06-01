"""Tests for the ``generate_excel`` Cloud Function.

These exercise the function through the Functions Framework's Flask test
client, which provides the request/response and application context that
``request.get_json`` and ``flask.jsonify`` rely on. The source under
``main.py`` is treated as a black box and is not modified.

The Excel template is no longer bundled with the function: callers (the GAS
frontend) read it from Google Drive and pass it as a Base64 ``template``
field in the request body. The helpers below build such a template in memory.
"""

import base64
import io
import os

import pytest
from functions_framework import create_app
from openpyxl import Workbook, load_workbook


SOURCE = os.path.join(os.path.dirname(__file__), "main.py")


@pytest.fixture
def client():
    app = create_app(target="generate_excel", source=SOURCE)
    return app.test_client()


def _make_template_b64(sheet_name="傾斜測定"):
    """Build a minimal template workbook and return it Base64-encoded."""
    wb = Workbook()
    wb.active.title = sheet_name
    buf = io.BytesIO()
    wb.save(buf)
    return base64.b64encode(buf.getvalue()).decode("utf-8")


def _payload(**fields):
    """Build a request body, defaulting to a valid Base64 template."""
    body = {"template": _make_template_b64()}
    body.update(fields)
    return body


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
    resp = client.post("/", json=_payload(
        floor="2", room_name="LDK", x_tilt="3", y_tilt="5",
    ))

    assert resp.status_code == 200
    assert resp.headers["Access-Control-Allow-Origin"] == "*"

    body = resp.get_json()
    assert body["status"] == "success"
    assert body["fileName"] == "傾斜測定報告書.xlsx"
    assert body["fileData"]


def test_post_valid_data_writes_values_into_template(client):
    resp = client.post("/", json=_payload(
        floor="2", room_name="LDK", x_tilt="3", y_tilt="5",
    ))
    wb = _decode_workbook(resp.get_json()["fileData"])
    ws = wb["傾斜測定"]

    assert ws["A15"].value == "2"
    assert ws["C15"].value == "LDK"
    assert ws["L15"].value == "3"
    assert ws["L17"].value == "5"


def test_writes_to_anchor_when_target_is_inside_a_merged_range(client):
    # Build a template where the write targets are NOT the top-left of their
    # merged ranges (A14:A15 makes A15 a read-only MergedCell, L14:L15 makes
    # L15 one). The function must write to the range anchors (A14 / L14) rather
    # than raise "'MergedCell' object attribute 'value' is read-only".
    wb = Workbook()
    ws = wb.active
    ws.title = "傾斜測定"
    ws.merge_cells("A14:A15")
    ws.merge_cells("L14:L15")
    buf = io.BytesIO()
    wb.save(buf)
    template = base64.b64encode(buf.getvalue()).decode("utf-8")

    resp = client.post("/", json={"floor": "7", "x_tilt": "9", "template": template})

    assert resp.status_code == 200
    out = _decode_workbook(resp.get_json()["fileData"])["傾斜測定"]
    assert out["A14"].value == "7"   # anchor of A14:A15
    assert out["L14"].value == "9"   # anchor of L14:L15


def test_returned_file_is_a_valid_xlsx(client):
    resp = client.post("/", json=_payload(floor="1", room_name="和室"))
    raw = base64.b64decode(resp.get_json()["fileData"])

    # XLSX files are zip archives and start with the PK signature.
    assert raw[:2] == b"PK"
    wb = _decode_workbook(resp.get_json()["fileData"])
    assert wb.sheetnames == ["傾斜測定"]


def test_missing_fields_default_to_empty(client):
    # Only ``floor`` is provided; the rest fall back to "" in the source, which
    # openpyxl stores as an empty (None) cell after the save round-trip.
    resp = client.post("/", json=_payload(floor="3"))
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


def test_post_without_template_returns_400(client):
    # Data is present but no template was supplied by the caller.
    resp = client.post("/", json={"floor": "1", "room_name": "LDK"})

    assert resp.status_code == 400
    assert resp.headers["Access-Control-Allow-Origin"] == "*"
    assert resp.get_json() == {"error": "No template provided"}


# --- Error handling ---------------------------------------------------------

def test_invalid_template_bytes_returns_500_with_message(client):
    # A non-xlsx Base64 blob makes load_workbook raise, exercising the
    # ``except`` branch that returns a 500 with the error text.
    bogus = base64.b64encode(b"not a real workbook").decode("utf-8")

    resp = client.post("/", json={"floor": "1", "room_name": "LDK", "template": bogus})

    assert resp.status_code == 500
    assert resp.headers["Access-Control-Allow-Origin"] == "*"
    assert resp.get_json()["error"]


def test_template_without_expected_sheet_returns_500(client):
    # A valid workbook that lacks the 傾斜測定 sheet raises a KeyError.
    resp = client.post("/", json=_payload(floor="1") | {
        "template": _make_template_b64(sheet_name="別のシート"),
    })

    assert resp.status_code == 500
    assert "傾斜測定" in resp.get_json()["error"]
