"""Tests for the ``generate_excel`` Cloud Function.

These exercise the function through the Functions Framework's Flask test
client, which provides the request/response and application context that
``request.get_json`` and ``flask.jsonify`` rely on. The source under
``main.py`` is treated as a black box and is not modified.

The Excel template is no longer bundled with the function: callers (the GAS
frontend) read it from Google Drive and pass it as a Base64 ``template``
field in the request body. The helpers below build such a template in memory.

Where data lands in the sheet is defined in ``mapping.json``; these tests
load the same mapping so they stay in sync if the layout is adjusted there.
"""

import base64
import io
import json
import os

import pytest
from functions_framework import create_app
from openpyxl import Workbook, load_workbook


SOURCE = os.path.join(os.path.dirname(__file__), "main.py")
MAPPING = json.load(open(os.path.join(os.path.dirname(__file__), "mapping.json"), encoding="utf-8"))
SHEET = MAPPING["sheet_name"]
BLOCK = MAPPING["room_block"]
STARTS = MAPPING["block_start_rows"]


@pytest.fixture
def client():
    app = create_app(target="generate_excel", source=SOURCE)
    return app.test_client()


def _make_template_b64(sheet_name=SHEET):
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


def _measurement_cell(room_index, measurement_key, field):
    """Resolve the cell a given room/measurement/field should be written to,
    straight from the mapping, so assertions follow the config."""
    start = STARTS[room_index]
    offset = next(m["row_offset"] for m in BLOCK["measurements"] if m["key"] == measurement_key)
    return f"{BLOCK['fields'][field]}{start + offset}"


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
        property_name="サンプル物件",
        rooms=[{"floor": "2", "room_name": "LDK", "measurements": {}}],
    ))

    assert resp.status_code == 200
    assert resp.headers["Access-Control-Allow-Origin"] == "*"

    body = resp.get_json()
    assert body["status"] == "success"
    assert body["fileName"] == "傾斜測定報告書.xlsx"
    assert body["fileData"]


def test_post_valid_data_writes_values_into_template(client):
    resp = client.post("/", json=_payload(
        property_name="サンプル物件",
        rooms=[{
            "floor": "2",
            "room_name": "LDK",
            "measurements": {
                "floor_x": {"direction": "傾斜無", "diff": "0", "distance": "2000"},
                "floor_y": {"direction": "↑", "diff": "3", "distance": "1500"},
            },
        }],
    ))
    ws = _decode_workbook(resp.get_json()["fileData"])[SHEET]

    # Common field
    assert ws[MAPPING["property_name_cell"]].value == "サンプル物件"
    # Room 0 header
    assert ws[f"{BLOCK['floor_col']}{STARTS[0]}"].value == 2
    assert ws[f"{BLOCK['room_name_col']}{STARTS[0]}"].value == "LDK"
    # floor_x measurement: numeric strings are coerced so the AJ formula computes
    assert ws[_measurement_cell(0, "floor_x", "direction")].value == "傾斜無"
    assert ws[_measurement_cell(0, "floor_x", "diff")].value == 0
    assert ws[_measurement_cell(0, "floor_x", "distance")].value == 2000
    # floor_y measurement
    assert ws[_measurement_cell(0, "floor_y", "direction")].value == "↑"
    assert ws[_measurement_cell(0, "floor_y", "diff")].value == 3
    assert ws[_measurement_cell(0, "floor_y", "distance")].value == 1500


def test_fullwidth_numbers_are_normalized_to_numeric(client):
    # Mobile IME often produces full-width digits/signs; they must be coerced to
    # real numbers so the AJ formula (=1000*AC/AG) can compute.
    resp = client.post("/", json=_payload(rooms=[{
        "floor": "２", "room_name": "LDK",
        "measurements": {"floor_x": {"diff": "－３", "distance": "１５００"}},
    }]))
    ws = _decode_workbook(resp.get_json()["fileData"])[SHEET]

    assert ws[f"{BLOCK['floor_col']}{STARTS[0]}"].value == 2
    assert ws[_measurement_cell(0, "floor_x", "diff")].value == -3
    assert ws[_measurement_cell(0, "floor_x", "distance")].value == 1500


def test_non_dict_room_entries_are_skipped(client):
    # A malformed rooms array (containing a null) must not crash the function.
    resp = client.post("/", json=_payload(rooms=[
        None,
        {"floor": "2", "room_name": "寝室", "measurements": {}},
    ]))
    assert resp.status_code == 200
    ws = _decode_workbook(resp.get_json()["fileData"])[SHEET]
    assert ws[f"{BLOCK['room_name_col']}{STARTS[1]}"].value == "寝室"


def test_multiple_rooms_map_to_successive_blocks(client):
    rooms = [
        {"floor": "1", "room_name": "玄関", "measurements": {}},
        {"floor": "2", "room_name": "寝室", "measurements": {}},
    ]
    resp = client.post("/", json=_payload(rooms=rooms))
    ws = _decode_workbook(resp.get_json()["fileData"])[SHEET]

    assert ws[f"{BLOCK['room_name_col']}{STARTS[0]}"].value == "玄関"
    assert ws[f"{BLOCK['room_name_col']}{STARTS[1]}"].value == "寝室"


def test_too_many_rooms_returns_400(client):
    rooms = [{"floor": "1", "room_name": f"部屋{i}", "measurements": {}}
             for i in range(len(STARTS) + 1)]
    resp = client.post("/", json=_payload(rooms=rooms))

    assert resp.status_code == 400
    assert "上限" in resp.get_json()["error"]


def test_returned_file_is_a_valid_xlsx(client):
    resp = client.post("/", json=_payload(
        rooms=[{"floor": "1", "room_name": "和室", "measurements": {}}]))
    raw = base64.b64decode(resp.get_json()["fileData"])

    # XLSX files are zip archives and start with the PK signature.
    assert raw[:2] == b"PK"
    wb = _decode_workbook(resp.get_json()["fileData"])
    assert wb.sheetnames == [SHEET]


def test_empty_or_missing_fields_default_to_empty(client):
    # No rooms and no property name: nothing is written, file still produced.
    resp = client.post("/", json=_payload())
    ws = _decode_workbook(resp.get_json()["fileData"])[SHEET]

    assert ws[MAPPING["property_name_cell"]].value is None
    assert ws[f"{BLOCK['floor_col']}{STARTS[0]}"].value is None


def test_blank_measurement_fields_leave_cells_empty(client):
    # A room with empty measurement values writes no data to those cells.
    resp = client.post("/", json=_payload(rooms=[{
        "floor": "3", "room_name": "",
        "measurements": {"floor_x": {"direction": "", "diff": "", "distance": ""}},
    }]))
    ws = _decode_workbook(resp.get_json()["fileData"])[SHEET]

    assert ws[f"{BLOCK['floor_col']}{STARTS[0]}"].value == 3
    assert ws[f"{BLOCK['room_name_col']}{STARTS[0]}"].value is None
    assert ws[_measurement_cell(0, "floor_x", "diff")].value is None


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
    resp = client.post("/", json={"property_name": "X", "rooms": []})

    assert resp.status_code == 400
    assert resp.headers["Access-Control-Allow-Origin"] == "*"
    assert resp.get_json() == {"error": "No template provided"}


# --- Error handling ---------------------------------------------------------

def test_invalid_template_bytes_returns_500_with_message(client):
    # A non-xlsx Base64 blob makes load_workbook raise, exercising the
    # ``except`` branch that returns a 500 with the error text.
    bogus = base64.b64encode(b"not a real workbook").decode("utf-8")

    resp = client.post("/", json={"rooms": [], "template": bogus})

    assert resp.status_code == 500
    assert resp.headers["Access-Control-Allow-Origin"] == "*"
    assert resp.get_json()["error"]


def test_template_without_expected_sheet_returns_500(client):
    # A valid workbook that lacks the 傾斜測定 sheet raises a KeyError.
    resp = client.post("/", json={
        "rooms": [{"floor": "1", "room_name": "LDK", "measurements": {}}],
        "template": _make_template_b64(sheet_name="別のシート"),
    })

    assert resp.status_code == 500
    assert SHEET in resp.get_json()["error"]
