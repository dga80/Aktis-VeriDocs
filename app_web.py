import os
import sys
import socket
import subprocess
import webbrowser
import threading
from flask import Flask, render_template, request, jsonify
from config import load_config, save_config
from veridocs_core import (
    ExcelManager,
    ScanningSession,
    load_ocr_catalog,
    process_scanned_pdf
)
from ocr_service import is_ocr_available, extract_of_from_header, extract_footer_text, match_footer_code

app = Flask(__name__)
excel_manager = ExcelManager()
session_manager = ScanningSession(excel_manager)


def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/system_info", methods=["GET"])
def system_info():
    cfg = load_config()
    excel_path = cfg.get("excel_path", "")
    escaneos_folder = cfg.get("escaneos_folder", "")
    dest_folder = cfg.get("destination_base_folder", "")

    excel_ok = os.path.exists(excel_path)
    escaneos_ok = os.path.exists(escaneos_folder)
    dest_ok = os.path.exists(dest_folder)

    pending_pdfs = []
    if escaneos_ok:
        try:
            with os.scandir(escaneos_folder) as entries:
                for e in entries:
                    if e.is_file() and e.name.lower().endswith(".pdf") and not e.name.startswith("_"):
                        pending_pdfs.append(e.name)
        except Exception:
            pass

    ocr_ok, ocr_msg = is_ocr_available()
    catalog = load_ocr_catalog()

    return jsonify({
        "hostname": socket.gethostname(),
        "local_ip": get_local_ip(),
        "port": cfg.get("web_port", 5055),
        "excel_path": excel_path,
        "excel_ok": excel_ok,
        "escaneos_folder": escaneos_folder,
        "escaneos_ok": escaneos_ok,
        "destination_base_folder": dest_folder,
        "destination_ok": dest_ok,
        "pending_scans_in_folder": pending_pdfs,
        "pending_count": len(pending_pdfs),
        "ocr_available": ocr_ok,
        "ocr_msg": ocr_msg,
        "ocr_enabled": cfg.get("ocr_enabled", True),
        "catalog_count": len(catalog),
        "watcher_active": session_manager.active
    })


@app.route("/api/catalog", methods=["GET"])
def get_catalog():
    catalog = load_ocr_catalog()
    return jsonify(catalog)


@app.route("/api/session", methods=["GET"])
def session_status():
    status = session_manager.get_status()
    return jsonify(status)


@app.route("/api/watcher/toggle", methods=["POST"])
def toggle_watcher():
    if not session_manager.active:
        ok, msg = session_manager.start_listening()
    else:
        ok, msg = session_manager.stop_listening()
    return jsonify({"success": ok, "message": msg, "active": session_manager.active})


@app.route("/api/watcher/start", methods=["POST"])
def start_watcher():
    ok, msg = session_manager.start_listening()
    return jsonify({"success": ok, "message": msg, "active": session_manager.active})


@app.route("/api/watcher/stop", methods=["POST"])
def stop_watcher():
    ok, msg = session_manager.stop_listening()
    return jsonify({"success": ok, "message": msg, "active": session_manager.active})


@app.route("/api/alerts/dismiss", methods=["POST"])
def dismiss_alert():
    data = request.json or {}
    alert_id = data.get("id")
    if alert_id:
        session_manager.dismiss_alert(alert_id)
        return jsonify({"success": True})
    return jsonify({"success": False, "error": "ID de alerta no proporcionado"}), 400


@app.route("/api/orders", methods=["GET"])
def get_orders():
    q = request.args.get("q", "").strip()
    results = excel_manager.search(q)
    return jsonify(results)


@app.route("/api/orders/reload", methods=["POST"])
def reload_orders():
    orders = excel_manager.get_orders(force_reload=True)
    return jsonify({"success": True, "count": len(orders)})


@app.route("/api/process_pdf_upload", methods=["POST"])
def process_pdf_upload():
    file = request.files.get("file")
    if not file:
        return jsonify({"success": False, "error": "No se ha subido ningún archivo PDF."}), 400

    temp_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"temp_upload_{file.filename}")
    try:
        file.save(temp_path)
        catalog = load_ocr_catalog()
        res = process_scanned_pdf(
            pdf_path=temp_path,
            excel_manager=excel_manager,
            catalog=catalog,
            log_callback=session_manager.log
        )
        if res["success"]:
            session_manager.last_batch = res
            session_manager.batches_history.append(res)
            if res.get("alert"):
                session_manager.add_alert(
                    title=f"Aviso en OF {res['of']}",
                    message=res["alert"],
                    level="warning",
                    dest_dir=res["destination_dir"]
                )
        return jsonify(res)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


@app.route("/api/config", methods=["GET", "POST"])
def manage_config():
    if request.method == "POST":
        new_cfg = request.json or {}
        current = load_config()
        current.update(new_cfg)
        if save_config(current):
            return jsonify({"success": True, "config": current})
        return jsonify({"success": False, "error": "No se pudo guardar la configuración."}), 500
    return jsonify(load_config())


def is_port_available(port, host="0.0.0.0"):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind((host, port))
            return True
        except OSError:
            return False


def find_available_port(start_port, host="0.0.0.0"):
    port = start_port
    while port < start_port + 50:
        if is_port_available(port, host):
            return port
        port += 1
    return start_port


def run_web_server(open_browser=False):
    cfg = load_config()
    configured_port = int(cfg.get("web_port", 5055))
    host = cfg.get("web_host", "0.0.0.0")

    port = find_available_port(configured_port, host)
    if port != configured_port:
        print(f"[Aviso] Puerto {configured_port} ocupado. Usando puerto libre {port}.")

    # Iniciar precarga de Excel y arranque de escucha automática
    threading.Thread(target=lambda: excel_manager.get_orders(force_reload=True), daemon=True).start()
    if cfg.get("auto_listen_on_startup", True):
        session_manager.start_listening()

    if open_browser:
        def _open():
            time.sleep(1.2)
            webbrowser.open(f"http://localhost:{port}")
        threading.Thread(target=_open, daemon=True).start()

    print(f"\n========================================================")
    print(f"  VeriDocs Servidor Web Iniciado")
    print(f"  - Acceso Local: http://localhost:{port}")
    print(f"  - Acceso Red Empresa: http://{get_local_ip()}:{port}")
    print(f"========================================================\n")

    app.run(host=host, port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    should_open = "--no-browser" not in sys.argv
    run_web_server(open_browser=should_open)
