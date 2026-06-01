import functions_framework
import base64
import io
import json
from openpyxl import load_workbook
from flask import jsonify


def _set_cell(ws, coord, value):
    """Write ``value`` to ``coord``, handling merged cells.

    In openpyxl, only the top-left (anchor) cell of a merged range is writable;
    the other cells in the range are read-only ``MergedCell`` objects. Writing
    to them raises "'MergedCell' object attribute 'value' is read-only".
    When the target sits inside a merged range, write to that range's anchor
    instead — that is the cell whose value is actually displayed.
    """
    for rng in ws.merged_cells.ranges:
        if coord in rng:
            ws.cell(row=rng.min_row, column=rng.min_col).value = value
            return
    ws[coord] = value


@functions_framework.http
def generate_excel(request):
    if request.method == 'OPTIONS':
        headers = {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'POST',
            'Access-Control-Allow-Headers': 'Content-Type',
            'Access-Control-Max-Age': '3600'
        }
        return ('', 204, headers)

    headers = {'Access-Control-Allow-Origin': '*'}

    try:
        request_json = request.get_json(silent=True)
        if not request_json:
            return jsonify({"error": "No data provided"}), 400, headers

        # 雛形（テンプレート）は GAS 側が Google Drive から読み取り、Base64 で
        # 渡してくる。社外秘フォーマットを Cloud Function に同梱しないことで、
        # フォーマット更新時の再デプロイを不要にし、ファイルのアクセス範囲を
        # 社内の閲覧可能者だけに閉じる。
        template_b64 = request_json.get('template')
        if not template_b64:
            return jsonify({"error": "No template provided"}), 400, headers

        wb = load_workbook(filename=io.BytesIO(base64.b64decode(template_b64)))
        ws = wb['傾斜測定']

        # データの書き込み (セル位置は調整してください)
        # 結合セルでも安全に書き込めるよう _set_cell を経由する。
        _set_cell(ws, 'A15', request_json.get('floor', ''))
        _set_cell(ws, 'C15', request_json.get('room_name', ''))
        _set_cell(ws, 'L15', request_json.get('x_tilt', ''))
        _set_cell(ws, 'L17', request_json.get('y_tilt', ''))

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        encoded_excel = base64.b64encode(output.read()).decode('utf-8')

        return jsonify({
            "status": "success",
            "fileName": "傾斜測定報告書.xlsx",
            "fileData": encoded_excel
        }), 200, headers

    except Exception as e:
        return jsonify({"error": str(e)}), 500, headers
