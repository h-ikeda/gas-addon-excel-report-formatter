import functions_framework
import base64
import io
import json
from openpyxl import load_workbook
from flask import jsonify

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
        ws['A15'] = request_json.get('floor', '')
        ws['C15'] = request_json.get('room_name', '')
        ws['L15'] = request_json.get('x_tilt', '')
        ws['L17'] = request_json.get('y_tilt', '')

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
