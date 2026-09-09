from dotenv import load_dotenv

try:
    load_dotenv()
except Exception:
    pass

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import tempfile
from backend.reports.avatars import generate_avatars
from backend.reports.conversion import generate_conversion
from backend.reports.playtime import generate_playtime
from backend.reports.sessions import generate_sessions
from backend.reports.repeat_visits import generate_repeat_visits
from backend.reports.quest_depth import generate_quest_depth

app = Flask(__name__, static_folder="../frontend", static_url_path="")
CORS(app)

@app.route("/")
@app.route("/index.html")
def serve_index():
    return app.send_static_file("index.html")

@app.route("/api/health")
def health_check():
    from backend.reports.common import get_api_token
    token = get_api_token()
    masked = f"{token[:4]}...{token[-4:]}" if len(token) > 8 else ("EMPTY" if len(token) == 0 else "SHORT")
    return jsonify({
        "status": "ok",
        "token_present": len(token) > 0,
        "token_length": len(token),
        "token_preview": masked
    })

@app.route("/api/generate", methods=["POST"])
def generate_report():
    data = request.json or {}
    date_val = data.get("dateValue")
    period_type = data.get("periodType", "month")
    report_type = data.get("type")
    
    from backend.reports.common import PARK_ORDER
    selected_parks = data.get("parks", PARK_ORDER)

    if not date_val or not report_type:
        return jsonify({"error": "Missing dateValue or type parameter"}), 400

    try:
        # Создаем временный файл
        fd, output_path = tempfile.mkstemp(suffix=".xlsx", prefix=f"{report_type}_{date_val}_")
        os.close(fd)

        if report_type == "avatars":
            generate_avatars(date_val, period_type, selected_parks, output_path)
            filename = f"Avatars_Report_{date_val}.xlsx"
        elif report_type == "conversion":
            generate_conversion(date_val, period_type, selected_parks, output_path)
            filename = f"Conversion_Report_{date_val}.xlsx"
        elif report_type == "playtime":
            generate_playtime(date_val, period_type, selected_parks, output_path)
            filename = f"Playtime_Report_{date_val}.xlsx"
        elif report_type == "sessions":
            generate_sessions(date_val, period_type, selected_parks, output_path)
            filename = f"Sessions_Report_{date_val}.xlsx"
        elif report_type == "repeat_visits":
            generate_repeat_visits(date_val, period_type, selected_parks, output_path)
            filename = f"Repeat_Visits_Report_{date_val}.xlsx"
        elif report_type == "quest_depth":
            generate_quest_depth(date_val, period_type, selected_parks, output_path)
            filename = f"Quest_Depth_Report_{date_val}.xlsx"
        else:
            return jsonify({"error": "Invalid report type"}), 400

        # Отправляем файл пользователю
        return send_file(
            output_path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port, debug=True)
