import os
import sys
import time
import threading
import subprocess
import warnings
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from config import load_config, save_config
from veridocs_core import (
    ExcelManager,
    ScanningSession,
    load_ocr_catalog,
    process_scanned_pdf,
    sanitize_filename,
    to_windows_long_path,
    DOCUMENTS_DEF,
    get_effective_documents_def
)
from ocr_service import is_ocr_available

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


def show_system_notification(title: str, message: str, is_error: bool = False):
    """Emite sonido de confirmación y lanza una notificación nativa Toast de Windows."""
    def _run():
        try:
            import winsound
            snd = winsound.MB_ICONHAND if is_error else winsound.MB_ICONASTERISK
            winsound.MessageBeep(snd)
        except Exception:
            pass

        # Intento con winrt nativo
        try:
            import winrt.windows.ui.notifications as notifications
            import winrt.windows.data.xml.dom as dom
            template_type = notifications.ToastTemplateType.TOAST_TEXT02
            xml_doc = notifications.ToastNotificationManager.get_template_content(template_type)
            text_nodes = xml_doc.get_elements_by_tag_name("text")
            text_nodes.item(0).append_child(xml_doc.create_text_node(title))
            text_nodes.item(1).append_child(xml_doc.create_text_node(message))
            toast = notifications.ToastNotification(xml_doc)
            notifier = notifications.ToastNotificationManager.create_toast_notifier("VeriDocs")
            notifier.show(toast)
            return
        except Exception:
            pass

        # Fallback PowerShell con escape seguro
        try:
            t = title.replace("'", "''").replace('"', '""')
            m = message.replace("'", "''").replace('"', '""')
            ps_script = (
                "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null; "
                "$template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02); "
                "$textNodes = $template.GetElementsByTagName('text'); "
                f"$textNodes.Item(0).AppendChild($template.CreateTextNode('{t}')) | Out-Null; "
                f"$textNodes.Item(1).AppendChild($template.CreateTextNode('{m}')) | Out-Null; "
                "$notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('VeriDocs'); "
                "$notification = [Windows.UI.Notifications.ToastNotification]::new($template); "
                "$notifier.Show($notification);"
            )
            subprocess.run(
                ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps_script],
                capture_output=True,
                timeout=4
            )
        except Exception:
            pass

    threading.Thread(target=_run, daemon=True).start()


def copy_to_clipboard(root: tk.Tk, text: str, feedback_label: ctk.CTkLabel = None):
    """Copia texto al portapapeles de Windows y opcionalmente muestra feedback."""
    try:
        root.clipboard_clear()
        root.clipboard_append(text)
        root.update()
        if feedback_label:
            original_text = feedback_label.cget("text")
            feedback_label.configure(text="✓ Copiado!")
            root.after(1500, lambda: feedback_label.configure(text=original_text))
    except Exception as e:
        print(f"[copy_to_clipboard] Error: {e}")


class VeriDocsDesktopApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("VeriDocs - Gestión y Renombrado Automático de Escaneos MES")
        self.geometry("1260x860")
        self.minsize(1060, 740)

        self.excel_manager = ExcelManager()
        self.session_manager = ScanningSession(self.excel_manager)
        self.catalog = load_ocr_catalog()
        self.filtered_catalog = list(self.catalog)
        self.selected_order = None
        self.current_alert_data = None
        self.currently_viewed_batch = None
        self._search_timer = None

        # Variables de control monotónicas para evitar parpadeos y congelaciones
        self.last_rendered_batch_key = None
        self.last_rendered_live_count = -1
        self.last_seen_event_id = 0
        self.last_seen_batch_id = -1
        self._active_alert_index = 0

        # Iniciar escucha automática si está configurada
        cfg = load_config()
        if cfg.get("auto_listen_on_startup", True):
            self.session_manager.start_listening()

        self.setup_ui()
        self.check_system_connectivity()

        # Interceptar cierre de ventana
        self.protocol("WM_DELETE_WINDOW", self.on_close_window)

        # Cargar pedidos de Excel en segundo plano
        threading.Thread(target=self._initial_load_excel, daemon=True).start()

        # Sondeo periódico de la sesión
        self.poll_session_status()

    def on_close_window(self):
        status = self.session_manager.get_status()
        if status.get("current_progress", {}).get("is_processing"):
            if not messagebox.askyesno("Lote en Proceso", "Hay un lote de escaneo procesándose actualmente.\n¿Seguro que deseas cerrar la aplicación?"):
                return
        self.session_manager.stop_listening()
        self.destroy()

    def _initial_load_excel(self):
        self.excel_manager.get_orders(force_reload=True)
        try:
            self.after(0, self.check_system_connectivity)
        except Exception:
            pass

    def setup_ui(self):
        # 1. Cabecera principal
        header = ctk.CTkFrame(self, height=64, corner_radius=8)
        header.pack(fill="x", padx=16, pady=(12, 6))

        # Título y Subtítulo
        title_box = ctk.CTkFrame(header, fg_color="transparent")
        title_box.pack(side="left", padx=14, pady=8)

        title_lbl = ctk.CTkLabel(
            title_box,
            text="📁 VeriDocs",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_lbl.pack(anchor="w")

        subtitle_lbl = ctk.CTkLabel(
            title_box,
            text="Troceado y clasificación automática de dossiers MES",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        subtitle_lbl.pack(anchor="w")

        # Botón de configuración
        btn_settings = ctk.CTkButton(
            header,
            text="⚙️ Configuración",
            width=130,
            height=32,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#3b82f6",
            hover_color="#2563eb",
            command=self.open_settings_dialog
        )
        btn_settings.pack(side="right", padx=(8, 14))

        # Píldoras de estado en la cabecera (Badge chips)
        self.pill_queue = ctk.CTkLabel(
            header,
            text="● Bandeja: 0 pend.",
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="gray25",
            corner_radius=12,
            padx=10,
            pady=4
        )
        self.pill_queue.pack(side="right", padx=4)

        self.pill_ot = ctk.CTkLabel(
            header,
            text="● OT: Comprobando...",
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="gray25",
            corner_radius=12,
            padx=10,
            pady=4
        )
        self.pill_ot.pack(side="right", padx=4)

        self.pill_ocr = ctk.CTkLabel(
            header,
            text="● OCR: Comprobando...",
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="gray25",
            corner_radius=12,
            padx=10,
            pady=4
        )
        self.pill_ocr.pack(side="right", padx=4)

        self.pill_folder = ctk.CTkLabel(
            header,
            text="● Escaneos: Comprobando...",
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="gray25",
            corner_radius=12,
            padx=10,
            pady=4
        )
        self.pill_folder.pack(side="right", padx=4)

        self.pill_excel = ctk.CTkLabel(
            header,
            text="● Excel: Comprobando...",
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="gray25",
            corner_radius=12,
            padx=10,
            pady=4
        )
        self.pill_excel.pack(side="right", padx=4)

        # 2. Banner de Alertas Múltiples (Visible solo cuando hay alerta activa)
        self.alert_frame = ctk.CTkFrame(self, fg_color=("#991b1b", "#7f1d1d"), corner_radius=8)

        self.lbl_alert_icon = ctk.CTkLabel(
            self.alert_frame, text="⚠️", font=ctk.CTkFont(size=22)
        )
        self.lbl_alert_icon.pack(side="left", padx=(14, 8), pady=10)

        self.lbl_alert_counter = ctk.CTkLabel(
            self.alert_frame,
            text="[1/1]",
            font=ctk.CTkFont(family="Consolas", size=11, weight="bold"),
            fg_color="gray30",
            corner_radius=4,
            padx=5
        )
        self.lbl_alert_counter.pack(side="left", padx=(0, 6))

        self.lbl_alert_msg = ctk.CTkLabel(
            self.alert_frame,
            text="",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="white",
            justify="left",
            wraplength=640
        )
        self.lbl_alert_msg.pack(side="left", padx=8, pady=10, fill="x", expand=True)

        self.btn_alert_folder = ctk.CTkButton(
            self.alert_frame,
            text="📂 Abrir Carpeta",
            width=120,
            height=30,
            fg_color="#f59e0b",
            hover_color="#d97706",
            text_color="black",
            font=ctk.CTkFont(weight="bold"),
            command=self.open_alert_folder
        )
        self.btn_alert_folder.pack(side="right", padx=4, pady=10)

        self.btn_alert_next = ctk.CTkButton(
            self.alert_frame,
            text="Siguiente ▶",
            width=90,
            height=30,
            fg_color="gray35",
            hover_color="gray45",
            font=ctk.CTkFont(weight="bold"),
            command=self.cycle_next_alert
        )
        self.btn_alert_next.pack(side="right", padx=4, pady=10)

        self.btn_alert_dismiss = ctk.CTkButton(
            self.alert_frame,
            text="✕ Descartar",
            width=85,
            height=30,
            fg_color="gray30",
            hover_color="gray40",
            font=ctk.CTkFont(weight="bold"),
            command=self.dismiss_current_alert
        )
        self.btn_alert_dismiss.pack(side="right", padx=(4, 10), pady=10)

        # 3. Contenedor dividido en 2 columnas
        self.main_content = ctk.CTkFrame(self, fg_color="transparent")
        self.main_content.pack(fill="both", expand=True, padx=16, pady=(4, 10))

        self.main_content.grid_columnconfigure(0, weight=4) # Columna izquierda
        self.main_content.grid_columnconfigure(1, weight=6) # Columna derecha
        self.main_content.grid_rowconfigure(0, weight=1)

        # Columna Izquierda
        left_col = ctk.CTkFrame(self.main_content)
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.setup_left_column(left_col)

        # Columna Derecha
        right_col = ctk.CTkFrame(self.main_content)
        right_col.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        self.setup_right_column(right_col)

    def setup_left_column(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        self.left_tabs = ctk.CTkTabview(parent)
        self.left_tabs.pack(fill="both", expand=True, padx=8, pady=8)

        tab_listen = self.left_tabs.add("🎧 Escucha")
        tab_history = self.left_tabs.add("📜 Historial")
        tab_aktis = self.left_tabs.add("🔍 Consulta Aktis")
        tab_cat = self.left_tabs.add("📋 Catálogo OCR")

        # --- TAB 1: ESCUCHA AUTOMÁTICA ---
        tab_listen.grid_columnconfigure(0, weight=1)

        box_badge = ctk.CTkFrame(tab_listen, fg_color="transparent")
        box_badge.pack(fill="x", padx=10, pady=(12, 6))

        ctk.CTkLabel(
            box_badge,
            text="Estado del Monitor:",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left")

        is_act = self.session_manager.active
        self.lbl_session_badge = ctk.CTkLabel(
            box_badge,
            text="ESCUCHANDO BANDEJA" if is_act else "INACTIVO",
            fg_color="#10b981" if is_act else "gray30",
            corner_radius=10,
            padx=12,
            pady=3,
            font=ctk.CTkFont(size=11, weight="bold")
        )
        self.lbl_session_badge.pack(side="right")

        self.btn_toggle_listen = ctk.CTkButton(
            tab_listen,
            text="⏹ DETENER ESCUCHA" if is_act else "▶ INICIAR ESCUCHA AUTOMÁTICA",
            height=48,
            fg_color="#ef4444" if is_act else "#059669",
            hover_color="#dc2626" if is_act else "#047857",
            text_color="#ffffff",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.toggle_listening
        )
        self.btn_toggle_listen.pack(fill="x", padx=10, pady=(4, 10))

        card_info = ctk.CTkFrame(tab_listen, fg_color=("gray90", "gray17"), corner_radius=6)
        card_info.pack(fill="x", padx=10, pady=6)

        ctk.CTkLabel(
            card_info,
            text="💡 Funcionamiento Continuo",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=12, pady=(10, 4))

        self.lbl_listen_info = ctk.CTkLabel(
            card_info,
            text="Al escanear un dossier en la fotocopiadora, VeriDocs lo detectará en pocos segundos, leerá la OF de la pág. 1, consultará Aktis, clasificará y troceará cada documento según el catálogo oficial.\n\nSoporta detección y corrección automática de hojas apaisadas y códigos rotados.",
            font=ctk.CTkFont(size=11),
            text_color="gray",
            wraplength=380,
            justify="left"
        )
        self.lbl_listen_info.pack(anchor="w", padx=12, pady=(0, 10))

        btn_open_scan_dir = ctk.CTkButton(
            card_info,
            text="📂 Abrir Bandeja de Escaneos...",
            height=32,
            fg_color="gray30",
            hover_color="gray40",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self.open_scans_folder
        )
        btn_open_scan_dir.pack(fill="x", padx=12, pady=(0, 10))

        card_manual = ctk.CTkFrame(tab_listen, fg_color=("gray90", "gray17"), corner_radius=6)
        card_manual.pack(fill="x", padx=10, pady=(8, 10))

        ctk.CTkLabel(
            card_manual,
            text="📄 Procesamiento Manual",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=12, pady=(10, 4))

        ctk.CTkLabel(
            card_manual,
            text="¿Quieres clasificar un PDF que ya tienes guardado en tu equipo?",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        ).pack(anchor="w", padx=12, pady=(0, 6))

        btn_manual = ctk.CTkButton(
            card_manual,
            text="Seleccionar y Procesar Archivo PDF...",
            height=34,
            fg_color=("#3b82f6", "#2563eb"),
            hover_color=("#1d4ed8", "#1d4ed8"),
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.process_pdf_manually
        )
        btn_manual.pack(fill="x", padx=12, pady=(0, 12))

        # --- TAB 2: HISTORIAL DE LOTES ---
        self.setup_history_tab(tab_history)

        # --- TAB 3: CONSULTA AKTIS ---
        self.setup_aktis_tab(tab_aktis)

        # --- TAB 4: CATÁLOGO OCR ---
        self.setup_catalog_tab(tab_cat)

    def setup_history_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)

        h_header = ctk.CTkFrame(parent, fg_color="transparent")
        h_header.pack(fill="x", padx=10, pady=(10, 4))

        self.lbl_hist_title = ctk.CTkLabel(
            h_header,
            text="Lotes Procesados (0)",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.lbl_hist_title.pack(side="left")

        btn_clear_hist = ctk.CTkButton(
            h_header,
            text="🗑️ Vaciar",
            width=65,
            height=24,
            fg_color="gray30",
            hover_color="gray40",
            font=ctk.CTkFont(size=10),
            command=self.clear_history_data
        )
        btn_clear_hist.pack(side="right")

        ctk.CTkLabel(
            parent,
            text="Haz clic en cualquier lote para abrir su carpeta o verlo en pantalla:",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        ).pack(anchor="w", padx=12, pady=(0, 6))

        self.history_scroll = ctk.CTkScrollableFrame(parent)
        self.history_scroll.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.render_history_list([])

    def render_history_list(self, batches: list):
        for w in self.history_scroll.winfo_children():
            w.destroy()

        if not batches:
            lbl_empty = ctk.CTkLabel(
                self.history_scroll,
                text="Aún no se ha procesado ningún lote en esta sesión.",
                text_color="gray"
            )
            lbl_empty.pack(pady=30)
            return

        for b in reversed(batches):
            card = ctk.CTkFrame(self.history_scroll, fg_color=("gray90", "gray18"), corner_radius=6)
            card.pack(fill="x", pady=3, padx=2)

            top_row = ctk.CTkFrame(card, fg_color="transparent")
            top_row.pack(fill="x", padx=8, pady=(6, 2))

            time_badge = ctk.CTkLabel(
                top_row,
                text=b.get("time", "--:--"),
                font=ctk.CTkFont(family="Consolas", size=10, weight="bold"),
                fg_color="gray30",
                corner_radius=4,
                padx=5
            )
            time_badge.pack(side="left")

            of_lbl = ctk.CTkLabel(
                top_row,
                text=f"OF: {b.get('of', '')}",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=("#1d4ed8", "#38bdf8")
            )
            of_lbl.pack(side="left", padx=6)

            if b.get("ot_replica_path"):
                ot_lbl = ctk.CTkLabel(
                    top_row,
                    text="OT ✓",
                    font=ctk.CTkFont(size=10, weight="bold"),
                    fg_color="#064e3b",
                    text_color="#34d399",
                    corner_radius=4,
                    padx=4
                )
                ot_lbl.pack(side="left", padx=2)

            info_row = ctk.CTkFrame(card, fg_color="transparent")
            info_row.pack(fill="x", padx=8, pady=(0, 4))

            files_count = b.get("files_count", len(b.get("created_files", [])))
            total_pages = b.get("total_pages", 0)
            o_info = b.get("order_info", {})
            odoo = o_info.get("pedido_odoo", "")
            desc = o_info.get("descripcion", "")

            odoo_tag = f"Odoo: {odoo} | " if odoo else ""
            desc_tag = f"\n{desc[:38]}" if desc else ""
            stats_lbl = ctk.CTkLabel(
                info_row,
                text=f"{odoo_tag}{files_count} archivos ({total_pages} págs){desc_tag}",
                font=ctk.CTkFont(size=10),
                text_color="gray",
                justify="left",
                anchor="w"
            )
            stats_lbl.pack(side="left", fill="x", expand=True)

            btns_row = ctk.CTkFrame(card, fg_color="transparent")
            btns_row.pack(fill="x", padx=8, pady=(0, 6))

            target_dir = b.get("destination_dir") or b.get("dest_dir", "")
            btn_open = ctk.CTkButton(
                btns_row,
                text="📂 Carpeta",
                width=80,
                height=24,
                font=ctk.CTkFont(size=10, weight="bold"),
                fg_color="#059669",
                hover_color="#047857",
                command=lambda d=target_dir: os.startfile(d) if os.path.exists(d) else messagebox.showinfo("Carpeta", f"No disponible: {d}")
            )
            btn_open.pack(side="right", padx=(4, 0))

            btn_view = ctk.CTkButton(
                btns_row,
                text="👁️ Ver Lote",
                width=75,
                height=24,
                font=ctk.CTkFont(size=10),
                fg_color=("#3b82f6", "#2563eb"),
                hover_color=("#1d4ed8", "#1d4ed8"),
                command=lambda batch=b: self.load_historical_batch_to_view(batch)
            )
            btn_view.pack(side="right")

    def load_historical_batch_to_view(self, batch: dict):
        self.currently_viewed_batch = batch
        self.render_batch_summary(batch)
        self.render_created_files_list(batch.get("created_files", []))
        self.btn_open_batch_folder.configure(state="normal", fg_color="#059669")
        self.btn_copy_batch_path.configure(state="normal")
        if batch.get("ot_replica_path"):
            self.btn_open_ot_folder.configure(state="normal", fg_color="#4338ca")
        else:
            self.btn_open_ot_folder.configure(state="disabled", fg_color="gray30")

        self.last_rendered_batch_key = f"{batch.get('of')}_{len(batch.get('created_files', []))}"
        if batch.get("alert"):
            self.lbl_alert_msg.configure(text=f"Aviso en lote cargado: {batch['alert']}")
            if not self.alert_frame.winfo_ismapped():
                self.alert_frame.pack(fill="x", padx=16, pady=(0, 6), before=self.main_content)

    def clear_history_data(self):
        if messagebox.askyesno("Vaciar Historial", "¿Deseas vaciar el historial de lotes procesados en esta sesión?"):
            self.session_manager.batches_history.clear()
            self.render_history_list([])
            self.lbl_hist_title.configure(text="Lotes Procesados (0)")

    def setup_aktis_tab(self, tab_aktis):
        tab_aktis.grid_columnconfigure(0, weight=1)

        ak_header = ctk.CTkFrame(tab_aktis, fg_color="transparent")
        ak_header.pack(fill="x", padx=10, pady=(8, 4))

        ctk.CTkLabel(
            ak_header,
            text="Buscar OF en Pedidos en Curso:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(side="left")

        btn_reload_ak = ctk.CTkButton(
            ak_header,
            text="↻ Recargar Excel",
            width=100,
            height=24,
            font=ctk.CTkFont(size=11),
            command=self.reload_excel
        )
        btn_reload_ak.pack(side="right")

        search_row = ctk.CTkFrame(tab_aktis, fg_color="transparent")
        search_row.pack(fill="x", padx=10, pady=(2, 6))

        self.of_entry = ctk.CTkEntry(
            search_row,
            placeholder_text="Escribe OF (ej: AYC/MO/00560 o 560)...",
            height=32
        )
        self.of_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.of_entry.bind("<KeyRelease>", self.on_of_search_typing)
        self.of_entry.bind("<Return>", lambda e: self.do_search_of())

        btn_search = ctk.CTkButton(
            search_row, text="Buscar", width=65, height=32,
            command=self.do_search_of
        )
        btn_search.pack(side="right")

        self.results_frame = ctk.CTkScrollableFrame(tab_aktis, height=120)

        self.details_box = ctk.CTkFrame(tab_aktis, fg_color=("gray85", "gray14"), corner_radius=6)
        self.details_box.pack(fill="both", expand=True, padx=10, pady=6)

        ctk.CTkLabel(
            self.details_box,
            text="Detalle del Pedido Seleccionado:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(anchor="w", padx=12, pady=(10, 4))

        self.lbl_detail_of = ctk.CTkLabel(self.details_box, text="OF: -", font=ctk.CTkFont(weight="bold", size=13))
        self.lbl_detail_of.pack(anchor="w", padx=12, pady=2)

        self.lbl_detail_odoo = ctk.CTkLabel(self.details_box, text="Pedido ODOO: -", font=ctk.CTkFont(size=12))
        self.lbl_detail_odoo.pack(anchor="w", padx=12, pady=2)

        self.lbl_detail_cliente = ctk.CTkLabel(self.details_box, text="Cliente: -", font=ctk.CTkFont(size=12))
        self.lbl_detail_cliente.pack(anchor="w", padx=12, pady=2)

        self.lbl_detail_desc = ctk.CTkLabel(
            self.details_box,
            text="Descripción: -",
            font=ctk.CTkFont(size=11),
            text_color="gray",
            wraplength=380,
            justify="left"
        )
        self.lbl_detail_desc.pack(anchor="w", padx=12, pady=(2, 10))

    def setup_catalog_tab(self, tab_cat):
        tab_cat.grid_columnconfigure(0, weight=1)

        cat_header = ctk.CTkFrame(tab_cat, fg_color="transparent")
        cat_header.pack(fill="x", padx=10, pady=(8, 4))

        self.lbl_catalog_title = ctk.CTkLabel(
            cat_header,
            text=f"Catálogo Oficial ({len(self.catalog)} tipos)",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.lbl_catalog_title.pack(side="left")

        btn_reload_cat = ctk.CTkButton(
            cat_header,
            text="↻ Recargar",
            width=85,
            height=24,
            font=ctk.CTkFont(size=11),
            command=self.reload_catalog
        )
        btn_reload_cat.pack(side="right")

        self.cat_filter_entry = ctk.CTkEntry(
            tab_cat,
            placeholder_text="Filtrar por código o nombre de documento...",
            height=28
        )
        self.cat_filter_entry.pack(fill="x", padx=10, pady=(2, 6))
        self.cat_filter_entry.bind("<KeyRelease>", self.filter_catalog_list)

        self.catalog_scroll = ctk.CTkScrollableFrame(tab_cat)
        self.catalog_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.render_catalog_list()

    def filter_catalog_list(self, event=None):
        q = self.cat_filter_entry.get().strip().lower()
        if not q:
            self.filtered_catalog = list(self.catalog)
        else:
            self.filtered_catalog = [
                d for d in self.catalog
                if q in d.get("codigo", "").lower() or q in d.get("documento", "").lower()
            ]
        self.render_catalog_list()

    def render_catalog_list(self):
        for w in self.catalog_scroll.winfo_children():
            w.destroy()

        for d in self.filtered_catalog:
            code = d.get("codigo", "")
            doc_name = d.get("documento", "")
            is_inc = d.get("is_incidencia", False)

            row = ctk.CTkFrame(self.catalog_scroll, fg_color=("gray90", "gray18"), height=30)
            row.pack(fill="x", pady=2, padx=2)

            code_bg = "#d97706" if is_inc else ("#4338ca" if "ANEXO" in code else "#2563eb")
            code_badge = ctk.CTkLabel(
                row,
                text=code,
                font=ctk.CTkFont(family="Consolas", size=11, weight="bold"),
                fg_color=code_bg,
                text_color="white",
                corner_radius=4,
                padx=6
            )
            code_badge.pack(side="left", padx=5, pady=3)

            tag_inc = " [Incidencia 1-N]" if is_inc else ""
            name_lbl = ctk.CTkLabel(
                row,
                text=doc_name + tag_inc,
                font=ctk.CTkFont(size=11, weight="bold" if is_inc else "normal"),
                text_color="#f59e0b" if is_inc else ("black", "white"),
                anchor="w"
            )
            name_lbl.pack(side="left", fill="x", expand=True, padx=4)

    def reload_catalog(self, show_msg: bool = True):
        self.catalog = load_ocr_catalog()
        self.filter_catalog_list()
        self.lbl_catalog_title.configure(text=f"Catálogo Oficial ({len(self.catalog)} tipos)")
        if show_msg:
            messagebox.showinfo("Catálogo", f"Se han cargado {len(self.catalog)} tipos de documentos desde codigosOCR.xlsx.")

    def setup_right_column(self, parent):
        parent.grid_columnconfigure(0, weight=1)

        # Encabezado del monitor con botones de acción rápida
        header_r = ctk.CTkFrame(parent, fg_color="transparent")
        header_r.pack(fill="x", padx=12, pady=(10, 4))

        ctk.CTkLabel(
            header_r,
            text="Monitor en Tiempo Real y Lotes",
            font=ctk.CTkFont(size=15, weight="bold")
        ).pack(side="left")

        btn_clear_view = ctk.CTkButton(
            header_r,
            text="🧹 Limpiar Vista",
            width=105,
            height=28,
            fg_color="gray30",
            hover_color="gray40",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self.clear_visual_state
        )
        btn_clear_view.pack(side="right", padx=(6, 0))

        btn_go_hist = ctk.CTkButton(
            header_r,
            text="📜 Ver Historial",
            width=105,
            height=28,
            fg_color="#3b82f6",
            hover_color="#2563eb",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: self.left_tabs.set("📜 Historial")
        )
        btn_go_hist.pack(side="right", padx=(6, 0))

        self.lbl_batch_counter = ctk.CTkLabel(
            header_r,
            text="Lotes hoy: 0",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_batch_counter.pack(side="right", padx=6)

        # Tarjeta: Último Lote Procesado
        self.last_batch_card = ctk.CTkFrame(parent, fg_color=("gray85", "gray17"), corner_radius=8)
        self.last_batch_card.pack(fill="x", padx=12, pady=(4, 6))

        batch_head = ctk.CTkFrame(self.last_batch_card, fg_color="transparent")
        batch_head.pack(fill="x", padx=12, pady=(10, 2))

        self.lbl_batch_title = ctk.CTkLabel(
            batch_head,
            text="Último Lote: Ninguno procesado todavía",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.lbl_batch_title.pack(side="left")

        # Botones de acción rápida del lote
        self.btn_open_ot_folder = ctk.CTkButton(
            batch_head,
            text="📁 Abrir en OT",
            height=28,
            width=100,
            fg_color="gray30",
            hover_color="#3730a3",
            font=ctk.CTkFont(size=11, weight="bold"),
            state="disabled",
            command=self.open_current_batch_ot_folder
        )
        self.btn_open_ot_folder.pack(side="right", padx=(4, 0))

        self.btn_copy_batch_path = ctk.CTkButton(
            batch_head,
            text="📋 Copiar Ruta",
            height=28,
            width=95,
            fg_color="gray30",
            hover_color="gray40",
            font=ctk.CTkFont(size=11, weight="bold"),
            state="disabled",
            command=self.copy_current_batch_folder
        )
        self.btn_copy_batch_path.pack(side="right", padx=(4, 0))

        self.btn_open_batch_folder = ctk.CTkButton(
            batch_head,
            text="📂 Abrir Carpeta OF",
            height=28,
            width=130,
            fg_color="gray30",
            hover_color="#047857",
            text_color="#ffffff",
            font=ctk.CTkFont(size=11, weight="bold"),
            state="disabled",
            command=self.open_current_batch_folder
        )
        self.btn_open_batch_folder.pack(side="right")

        self.lbl_batch_sub = ctk.CTkLabel(
            self.last_batch_card,
            text="La escucha automática está lista. Al caer un dossier en la bandeja se procesará automáticamente.",
            font=ctk.CTkFont(size=12),
            text_color="gray",
            justify="left",
            wraplength=600
        )
        self.lbl_batch_sub.pack(anchor="w", padx=12, pady=(0, 4))

        self.lbl_ot_status = ctk.CTkLabel(
            self.last_batch_card,
            text="",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#34d399",
            justify="left"
        )
        self.lbl_ot_status.pack(anchor="w", padx=12, pady=(0, 8))

        # Barra y panel de Progreso en Vivo
        self.progress_frame = ctk.CTkFrame(parent, fg_color=("#dbeafe", "#0f2744"), corner_radius=6)

        self.lbl_progress_text = ctk.CTkLabel(
            self.progress_frame,
            text="⚙️ Procesando escaneo...",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=("#1e40af", "#93c5fd"),
            anchor="w",
            justify="left"
        )
        self.lbl_progress_text.pack(fill="x", padx=12, pady=(8, 4))

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, height=9, corner_radius=5)
        self.progress_bar.pack(fill="x", padx=12, pady=(0, 10))
        self.progress_bar.set(0)

        # Lista de Archivos Creados
        files_header_box = ctk.CTkFrame(parent, fg_color="transparent")
        files_header_box.pack(fill="x", padx=14, pady=(6, 2))

        self.lbl_files_header = ctk.CTkLabel(
            files_header_box,
            text="Documentos PDF Generados del Lote:",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.lbl_files_header.pack(side="left")

        ctk.CTkLabel(
            files_header_box,
            text="(💡 Doble clic en cualquier documento para abrirlo)",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        ).pack(side="left", padx=8)

        self.files_scroll = ctk.CTkScrollableFrame(parent, height=220)
        self.files_scroll.pack(fill="both", expand=True, padx=12, pady=(0, 6))

        # Consola de Eventos
        log_header_box = ctk.CTkFrame(parent, fg_color="transparent")
        log_header_box.pack(fill="x", padx=14, pady=(4, 2))

        ctk.CTkLabel(
            log_header_box,
            text="Registro de Actividad en Directo:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="left")

        btn_copy_log = ctk.CTkButton(
            log_header_box,
            text="📋 Copiar",
            width=65,
            height=22,
            font=ctk.CTkFont(size=10),
            fg_color="gray30",
            hover_color="gray40",
            command=self.copy_log_to_clipboard
        )
        btn_copy_log.pack(side="right", padx=(4, 0))

        btn_clear_log = ctk.CTkButton(
            log_header_box,
            text="🗑️ Limpiar",
            width=65,
            height=22,
            font=ctk.CTkFont(size=10),
            fg_color="gray30",
            hover_color="gray40",
            command=self.clear_log_console
        )
        btn_clear_log.pack(side="right")

        self.log_textbox = ctk.CTkTextbox(
            parent,
            height=140,
            font=ctk.CTkFont(family="Consolas", size=10),
            activate_scrollbars=True
        )
        self.log_textbox.pack(fill="x", padx=12, pady=(0, 10))

        self.log_textbox.tag_config("SUCCESS", foreground="#10b981")
        self.log_textbox.tag_config("WARNING", foreground="#f59e0b")
        self.log_textbox.tag_config("ERROR", foreground="#ef4444")
        self.log_textbox.tag_config("INFO", foreground="#94a3b8")

    def clear_visual_state(self):
        """Limpia todo lo que se ha mostrado visualmente en pantalla (resetea vista)."""
        self.lbl_batch_title.configure(text="Último Lote: Vista restablecida")
        self.lbl_batch_sub.configure(text="Pantalla limpiada. La escucha sigue activa y el próximo escaneo se mostrará aquí.")
        self.lbl_ot_status.configure(text="")
        self.btn_open_batch_folder.configure(state="disabled", fg_color="gray30")
        self.btn_copy_batch_path.configure(state="disabled")
        self.btn_open_ot_folder.configure(state="disabled", fg_color="gray30")
        self.currently_viewed_batch = None

        for w in self.files_scroll.winfo_children():
            w.destroy()
        lbl_empty = ctk.CTkLabel(
            self.files_scroll,
            text="Vista limpia. Los documentos del próximo escaneo aparecerán aquí.",
            text_color="gray"
        )
        lbl_empty.pack(pady=20)
        self.lbl_files_header.configure(text="Documentos PDF Generados del Lote:")

        if self.progress_frame.winfo_ismapped():
            self.progress_frame.pack_forget()

        self.log_textbox.delete("1.0", "end")

        if self.alert_frame.winfo_ismapped():
            self.alert_frame.pack_forget()
        self.current_alert_data = None

        self.last_rendered_batch_key = None
        self.last_rendered_live_count = -1

    def toggle_listening(self):
        if not self.session_manager.active:
            ok, msg = self.session_manager.start_listening()
            if ok:
                self.lbl_session_badge.configure(text="ESCUCHANDO BANDEJA", fg_color="#10b981")
                self.btn_toggle_listen.configure(
                    text="⏹ DETENER ESCUCHA",
                    fg_color="#ef4444",
                    hover_color="#dc2626"
                )
            else:
                messagebox.showerror("Error al iniciar", msg)
        else:
            self.session_manager.stop_listening()
            self.lbl_session_badge.configure(text="INACTIVO", fg_color="gray30")
            self.btn_toggle_listen.configure(
                text="▶ INICIAR ESCUCHA AUTOMÁTICA",
                fg_color="#059669",
                hover_color="#047857"
            )

    def open_scans_folder(self):
        cfg = load_config()
        folder = cfg.get("escaneos_folder", "")
        if os.path.exists(folder):
            os.startfile(folder)
        else:
            messagebox.showinfo("Bandeja", f"La carpeta de escaneos no está accesible:\n{folder}")

    def process_pdf_manually(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo PDF escaneado",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if not file_path:
            return

        def _worker():
            fname = os.path.basename(file_path)
            self.session_manager.log(f"Iniciando procesamiento manual de: '{fname}'...")
            res = self.session_manager.process_file(file_path)

            if res.get("success"):
                show_system_notification(
                    f"OF {res['of']} procesada con éxito",
                    f"Se han generado {len(res['created_files'])} archivos PDF."
                )
                self.after(0, lambda: messagebox.showinfo(
                    "Éxito",
                    f"Lote de la OF '{res['of']}' procesado con éxito.\nSe han creado {len(res['created_files'])} archivos PDF."
                ))
            else:
                err = res.get("error", "Error desconocido")
                self.session_manager.add_alert(
                    title="Error en procesamiento manual",
                    message=err,
                    level="error"
                )
                show_system_notification("Error en procesamiento", err, is_error=True)
                self.after(0, lambda: messagebox.showerror("Error", err))

        threading.Thread(target=_worker, daemon=True).start()

    def poll_session_status(self):
        """Sondeo periódico sin parpadeos ni congelaciones."""
        try:
            status = self.session_manager.get_status()

            # 1. Estado del botón
            if status["active"]:
                if self.lbl_session_badge.cget("text") != "ESCUCHANDO BANDEJA":
                    self.lbl_session_badge.configure(text="ESCUCHANDO BANDEJA", fg_color="#10b981")
                if self.btn_toggle_listen.cget("text") != "⏹ DETENER ESCUCHA":
                    self.btn_toggle_listen.configure(
                        text="⏹ DETENER ESCUCHA",
                        fg_color="#ef4444",
                        hover_color="#dc2626"
                    )
            else:
                if self.lbl_session_badge.cget("text") != "INACTIVO":
                    self.lbl_session_badge.configure(text="INACTIVO", fg_color="gray30")
                if self.btn_toggle_listen.cget("text") != "▶ INICIAR ESCUCHA AUTOMÁTICA":
                    self.btn_toggle_listen.configure(
                        text="▶ INICIAR ESCUCHA AUTOMÁTICA",
                        fg_color="#059669",
                        hover_color="#047857"
                    )

            self.lbl_batch_counter.configure(text=f"Lotes procesados hoy: {status['total_batches_processed']}")

            # 2. Manejo de alertas múltiples
            active_alerts = status.get("active_alerts", [])
            if active_alerts:
                if self._active_alert_index >= len(active_alerts):
                    self._active_alert_index = 0
                alert_item = active_alerts[self._active_alert_index]
                self.current_alert_data = alert_item
                self.lbl_alert_counter.configure(text=f"[{self._active_alert_index + 1}/{len(active_alerts)}]")
                self.lbl_alert_msg.configure(text=f"{alert_item['title']}: {alert_item['message']}")

                if len(active_alerts) > 1:
                    self.btn_alert_next.configure(state="normal")
                else:
                    self.btn_alert_next.configure(state="disabled")

                if not self.alert_frame.winfo_ismapped():
                    self.alert_frame.pack(fill="x", padx=16, pady=(0, 6), before=self.main_content)
            else:
                if self.alert_frame.winfo_ismapped():
                    self.alert_frame.pack_forget()

            # 3. Barra de progreso en vivo
            curr_prog = status.get("current_progress", {})
            is_proc = curr_prog.get("is_processing", False)

            if is_proc:
                if not self.progress_frame.winfo_ismapped():
                    self.progress_frame.pack(fill="x", padx=12, pady=(0, 6), before=self.lbl_files_header.master)

                p_file = curr_prog.get("filename", "")
                p_of = curr_prog.get("of", "")
                p_phase = curr_prog.get("phase", "")
                c_p = curr_prog.get("current_page", 0)
                t_p = max(1, curr_prog.get("total_pages", 1))

                of_tag = f"  |  OF: {p_of}" if p_of else ""
                self.lbl_progress_text.configure(
                    text=f"⚙️ Procesando: {p_file}{of_tag}\n   {p_phase}"
                )
                self.progress_bar.set(min(1.0, c_p / float(t_p)))

                live_files = curr_prog.get("created_files", [])
                if len(live_files) != self.last_rendered_live_count:
                    self.last_rendered_live_count = len(live_files)
                    self.render_created_files_list(live_files)
            else:
                if self.progress_frame.winfo_ismapped():
                    self.progress_frame.pack_forget()

                last_b = status.get("last_batch")
                if last_b and not self.currently_viewed_batch:
                    batch_key = f"{last_b.get('of')}_{len(last_b.get('created_files', []))}_{last_b.get('destination_dir')}"
                    if batch_key != self.last_rendered_batch_key:
                        self.last_rendered_batch_key = batch_key
                        self.render_batch_summary(last_b)
                        self.render_created_files_list(last_b.get("created_files", []))
                        self.btn_open_batch_folder.configure(state="normal", fg_color="#059669")
                        self.btn_copy_batch_path.configure(state="normal")
                        if last_b.get("ot_replica_path"):
                            self.btn_open_ot_folder.configure(state="normal", fg_color="#4338ca")

            # 4. Actualizar historial de lotes con ID monotónico (nunca se congela)
            last_b_id = status.get("last_batch_id", 0)
            if last_b_id != self.last_seen_batch_id:
                self.last_seen_batch_id = last_b_id
                batches = status.get("batches_history", [])
                self.lbl_hist_title.configure(text=f"Lotes Procesados ({len(batches)})")
                self.render_history_list(batches)

            # 5. Actualizar consola de registro incrementalmente con ID (nunca se congela)
            events = status.get("events_log", [])
            new_events = [ev for ev in events if ev.get("id", 0) > self.last_seen_event_id]
            if new_events:
                for ev in new_events:
                    t = ev.get("time", "")
                    m = ev.get("message", "")
                    lvl = ev.get("level", "info").upper()
                    tag = lvl if lvl in ("SUCCESS", "WARNING", "ERROR") else "INFO"
                    line = f"[{t}] [{lvl}] {m}\n"
                    self.log_textbox.insert("end", line, tag)
                    self.last_seen_event_id = max(self.last_seen_event_id, ev.get("id", 0))
                self.log_textbox.see("end")

        except Exception as e:
            print(f"[poll_session_status] Error: {e}")
        finally:
            self.after(1000, self.poll_session_status)

    def cycle_next_alert(self):
        status = self.session_manager.get_status()
        alerts = status.get("active_alerts", [])
        if alerts:
            self._active_alert_index = (self._active_alert_index + 1) % len(alerts)
            alert_item = alerts[self._active_alert_index]
            self.current_alert_data = alert_item
            self.lbl_alert_counter.configure(text=f"[{self._active_alert_index + 1}/{len(alerts)}]")
            self.lbl_alert_msg.configure(text=f"{alert_item['title']}: {alert_item['message']}")

    def render_batch_summary(self, last_b: dict):
        of_str = last_b.get("of", "")
        o_info = last_b.get("order_info", {})
        col_e = o_info.get("pedido_odoo", "")
        col_i = o_info.get("descripcion", "")
        col_cli = o_info.get("cliente", "")
        dest_dir = last_b.get("destination_dir") or last_b.get("dest_dir", "")
        t_pages = last_b.get("total_pages", 0)

        title_text = f"OF: {of_str}  |  Odoo: {col_e or '(N/A)'}  |  Total: {t_pages} páginas"
        self.lbl_batch_title.configure(text=title_text)

        cli_tag = f"  |  Cliente: {col_cli}" if col_cli else ""
        desc_tag = f"Descripción: {col_i or '(Sin descripción)'}{cli_tag}"
        self.lbl_batch_sub.configure(text=f"{desc_tag}\nCarpeta: {dest_dir}")

        ot_replica = last_b.get("ot_replica_path")
        if ot_replica:
            self.lbl_ot_status.configure(text=f"✓ Replicado en Oficina Técnica: {ot_replica}")
            self.btn_open_ot_folder.configure(state="normal", fg_color="#4338ca")
        else:
            self.lbl_ot_status.configure(text="")
            self.btn_open_ot_folder.configure(state="disabled", fg_color="gray30")

    def render_created_files_list(self, files: list):
        for w in self.files_scroll.winfo_children():
            w.destroy()

        if not files:
            lbl_empty = ctk.CTkLabel(
                self.files_scroll,
                text="Los archivos troceados aparecerán aquí conforme se generen.",
                text_color="gray"
            )
            lbl_empty.pack(pady=20)
            self.lbl_files_header.configure(text="Documentos PDF Generados del Lote:")
            return

        self.lbl_files_header.configure(text=f"Documentos PDF Generados del Lote ({len(files)} archivos):")

        for f in files:
            is_warn = f.get("is_warning", False)
            f_path = f.get("path", "")
            f_name = f.get("filename", "")
            doc_n = f.get("doc_name", "")
            pages = f.get("pages_list", [])

            card_bg = ("#fef3c7", "#3b1d06") if is_warn else ("gray90", "gray18")
            card = ctk.CTkFrame(self.files_scroll, fg_color=card_bg, height=44, corner_radius=6)
            card.pack(fill="x", pady=2, padx=2)

            icon = "⚠️ " if is_warn else "📄 "
            pages_tag = f"  [Págs: {pages}]" if pages else ""

            info_box = ctk.CTkFrame(card, fg_color="transparent")
            info_box.pack(side="left", fill="both", expand=True, padx=8, pady=4)

            title_c = ctk.CTkLabel(
                info_box,
                text=f"{icon}{doc_n}{pages_tag}",
                font=ctk.CTkFont(weight="bold", size=12),
                text_color="#f59e0b" if is_warn else ("black", "white"),
                anchor="w"
            )
            title_c.pack(anchor="w")

            file_c = ctk.CTkLabel(
                info_box,
                text=f_name,
                font=ctk.CTkFont(family="Consolas", size=10),
                text_color="gray",
                anchor="w"
            )
            file_c.pack(anchor="w")

            actions_box = ctk.CTkFrame(card, fg_color="transparent")
            actions_box.pack(side="right", padx=6)

            btn_copy_f = ctk.CTkButton(
                actions_box,
                text="📋 Copiar",
                width=65,
                height=26,
                fg_color="gray30",
                hover_color="gray40",
                font=ctk.CTkFont(size=10),
                command=lambda p=f_path: copy_to_clipboard(self, p)
            )
            btn_copy_f.pack(side="right", padx=2)

            btn_open_f = ctk.CTkButton(
                actions_box,
                text="👁️ Abrir",
                width=65,
                height=26,
                fg_color=("#3b82f6", "#2563eb"),
                hover_color=("#1d4ed8", "#1d4ed8"),
                font=ctk.CTkFont(size=10, weight="bold"),
                command=lambda p=f_path: self.open_file_path(p)
            )
            btn_open_f.pack(side="right", padx=2)

            card.bind("<Double-Button-1>", lambda e, p=f_path: self.open_file_path(p))
            title_c.bind("<Double-Button-1>", lambda e, p=f_path: self.open_file_path(p))
            file_c.bind("<Double-Button-1>", lambda e, p=f_path: self.open_file_path(p))

    def open_file_path(self, fpath: str):
        if fpath and os.path.exists(fpath):
            try:
                os.startfile(fpath)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
        else:
            messagebox.showinfo("Aviso", f"El archivo ya no está disponible en la ruta especificada:\n{fpath}")

    def copy_current_batch_folder(self):
        target_b = self.currently_viewed_batch or self.session_manager.get_status().get("last_batch")
        if target_b:
            folder = target_b.get("destination_dir") or target_b.get("dest_dir", "")
            if folder:
                copy_to_clipboard(self, folder, self.btn_copy_batch_path)

    def open_current_batch_folder(self):
        target_b = self.currently_viewed_batch or self.session_manager.get_status().get("last_batch")
        if target_b:
            folder = target_b.get("destination_dir") or target_b.get("dest_dir", "")
            if folder and os.path.exists(folder):
                os.startfile(folder)
            else:
                messagebox.showinfo("Aviso", f"La carpeta aún no está disponible: {folder}")
        else:
            messagebox.showinfo(
                "Información",
                "Aún no se ha procesado ningún lote de escaneo en esta sesión.\n\nEn cuanto se procese un dossier, podrás abrir su carpeta directamente desde aquí."
            )

    def open_current_batch_ot_folder(self):
        target_b = self.currently_viewed_batch or self.session_manager.get_status().get("last_batch")
        if target_b and target_b.get("ot_replica_path"):
            folder = target_b["ot_replica_path"]
            if os.path.exists(folder):
                os.startfile(folder)
            else:
                messagebox.showinfo("Oficina Técnica", f"La ruta de réplica no está accesible:\n{folder}")

    def dismiss_current_alert(self):
        if self.current_alert_data:
            self.session_manager.dismiss_alert(self.current_alert_data["id"])
            self.current_alert_data = None
        if self.alert_frame.winfo_ismapped():
            self.alert_frame.pack_forget()

    def open_alert_folder(self):
        if self.current_alert_data and self.current_alert_data.get("dest_dir"):
            folder = self.current_alert_data["dest_dir"]
            if os.path.exists(folder):
                os.startfile(folder)
            else:
                messagebox.showinfo("Aviso", f"La carpeta no está accesible: {folder}")

    def copy_log_to_clipboard(self):
        content = self.log_textbox.get("1.0", "end-1c")
        copy_to_clipboard(self, content)
        messagebox.showinfo("Registro", "El registro de actividad ha sido copiado al portapapeles.")

    def clear_log_console(self):
        self.log_textbox.delete("1.0", "end")

    # Búsqueda debounced en Aktis
    def on_of_search_typing(self, event=None):
        if event and event.keysym in ("Return", "KP_Enter", "Up", "Down"):
            return
        if self._search_timer:
            self.after_cancel(self._search_timer)
        self._search_timer = self.after(250, self._do_debounced_search)

    def _do_debounced_search(self):
        val = self.of_entry.get().strip()
        if len(val) >= 2:
            matches = self.excel_manager.search(val)
            if matches:
                self.show_results_dropdown(matches)
            else:
                if self.results_frame.winfo_ismapped():
                    self.results_frame.pack_forget()
        else:
            if self.results_frame.winfo_ismapped():
                self.results_frame.pack_forget()

    def do_search_of(self):
        val = self.of_entry.get().strip()
        if not val:
            return
        order = self.excel_manager.find_by_of(val)
        if order:
            self.select_order(order)
            if self.results_frame.winfo_ismapped():
                self.results_frame.pack_forget()
        else:
            matches = self.excel_manager.search(val)
            if len(matches) == 1:
                self.select_order(matches[0])
                if self.results_frame.winfo_ismapped():
                    self.results_frame.pack_forget()
            elif matches:
                self.show_results_dropdown(matches)
            else:
                if self.results_frame.winfo_ismapped():
                    self.results_frame.pack_forget()
                messagebox.showinfo("Búsqueda", f"No se encontró ninguna OF que coincida con '{val}'.")

    def show_results_dropdown(self, matches):
        for widget in self.results_frame.winfo_children():
            widget.destroy()

        if not self.results_frame.winfo_ismapped():
            self.results_frame.pack(fill="x", padx=10, pady=4, before=self.details_box)

        for m in matches[:6]:
            btn = ctk.CTkButton(
                self.results_frame,
                text=f"{m['of']} | {m['pedido_odoo']} | {m['descripcion'][:35]}",
                anchor="w",
                fg_color="transparent",
                hover_color=("gray75", "gray25"),
                text_color=("black", "white"),
                height=26,
                command=lambda order=m: self.select_order(order)
            )
            btn.pack(fill="x", pady=1)

    def select_order(self, order):
        self.selected_order = order
        self.of_entry.delete(0, "end")
        self.of_entry.insert(0, order["of"])

        if self.results_frame.winfo_ismapped():
            self.results_frame.pack_forget()

        self.lbl_detail_of.configure(text=f"OF: {order['of']}")
        self.lbl_detail_odoo.configure(text=f"Pedido ODOO: {order['pedido_odoo'] or '(Vacío)'}")
        self.lbl_detail_cliente.configure(text=f"Cliente: {order.get('cliente', '') or '(No especificado)'}")
        self.lbl_detail_desc.configure(text=f"Descripción: {order['descripcion'] or '(Sin descripción)'}")

    def reload_excel(self):
        def _reload():
            orders = self.excel_manager.get_orders(force_reload=True)
            self.after(0, lambda: messagebox.showinfo("Excel", f"Se han cargado {len(orders)} órdenes correctamente."))
            self.after(0, self.check_system_connectivity)
        threading.Thread(target=_reload, daemon=True).start()

    def check_system_connectivity(self):
        def _worker():
            cfg = load_config()
            excel_path = cfg.get("excel_path", "")
            escaneos_folder = cfg.get("escaneos_folder", "")
            ot_base = cfg.get("oficina_tecnica_base", r"\\172.30.0.10\Oficina_tecnica")

            try:
                ex_ok = bool(excel_path and os.path.exists(excel_path))
            except Exception:
                ex_ok = False

            pending_count = 0
            try:
                fo_ok = bool(escaneos_folder and os.path.exists(escaneos_folder))
                if fo_ok:
                    pending_count = sum(1 for e in os.scandir(escaneos_folder) if e.is_file() and e.name.lower().endswith(".pdf") and not e.name.startswith("_"))
            except Exception:
                fo_ok = False

            try:
                ot_ok = bool(ot_base and os.path.exists(ot_base))
            except Exception:
                ot_ok = False

            ocr_ok, _ = is_ocr_available()
            ocr_en = cfg.get("ocr_enabled", True)

            try:
                self.after(0, lambda: self._apply_connectivity_ui(ex_ok, fo_ok, ot_ok, ocr_ok, ocr_en, pending_count))
            except Exception:
                pass

        threading.Thread(target=_worker, daemon=True).start()

    def _apply_connectivity_ui(self, ex_ok: bool, fo_ok: bool, ot_ok: bool, ocr_ok: bool, ocr_en: bool, pending_count: int):
        self.pill_excel.configure(
            text="● Excel: OK" if ex_ok else "● Excel: No accesible",
            fg_color="#064e3b" if ex_ok else "#7f1d1d",
            text_color="#34d399" if ex_ok else "#fca5a5"
        )
        self.pill_folder.configure(
            text="● Escaneos: OK" if fo_ok else "● Escaneos: No accesible",
            fg_color="#064e3b" if fo_ok else "#7f1d1d",
            text_color="#34d399" if fo_ok else "#fca5a5"
        )
        self.pill_ot.configure(
            text="● OT: OK" if ot_ok else "● OT: No accesible",
            fg_color="#064e3b" if ot_ok else "#7f1d1d",
            text_color="#34d399" if ot_ok else "#fca5a5"
        )
        self.pill_queue.configure(
            text=f"● Bandeja: {pending_count} pend.",
            fg_color="#78350f" if pending_count > 0 else "#064e3b",
            text_color="#fcd34d" if pending_count > 0 else "#34d399"
        )

        if not ocr_ok:
            self.pill_ocr.configure(text="● OCR: No disp.", fg_color="#7f1d1d", text_color="#fca5a5")
        elif not ocr_en:
            self.pill_ocr.configure(text="● OCR: Inactivo", fg_color="#78350f", text_color="#fcd34d")
        else:
            self.pill_ocr.configure(text="● OCR: Activo", fg_color="#064e3b", text_color="#34d399")

    def open_settings_dialog(self):
        cfg = load_config()
        dlg = ctk.CTkToplevel(self)
        dlg.title("Configuración de VeriDocs")
        dlg.geometry("680x620")
        dlg.minsize(580, 560)
        dlg.grab_set()

        tabview = ctk.CTkTabview(dlg)
        tabview.pack(fill="both", expand=True, padx=16, pady=(12, 8))

        tab_paths = tabview.add("📁 Rutas de Red")
        tab_ocr = tabview.add("🔍 Parámetros OCR")

        # --- TAB 1: RUTAS con botones Examinar ---
        ctk.CTkLabel(tab_paths, text="Ruta del archivo Excel (Pedidos en curso Aktis):", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(8, 2))
        row_ex = ctk.CTkFrame(tab_paths, fg_color="transparent")
        row_ex.pack(fill="x", pady=(0, 10))
        entry_excel = ctk.CTkEntry(row_ex)
        entry_excel.insert(0, cfg.get("excel_path", ""))
        entry_excel.pack(side="left", fill="x", expand=True, padx=(0, 6))
        def browse_excel():
            f = filedialog.askopenfilename(title="Seleccionar archivo Excel Aktis", filetypes=[("Excel", "*.xlsm;*.xlsx")])
            if f:
                entry_excel.delete(0, "end")
                entry_excel.insert(0, f)
        ctk.CTkButton(row_ex, text="Examinar...", width=95, command=browse_excel).pack(side="right")

        ctk.CTkLabel(tab_paths, text="Carpeta de entrada (bandeja de escaneos):", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(4, 2))
        row_sc = ctk.CTkFrame(tab_paths, fg_color="transparent")
        row_sc.pack(fill="x", pady=(0, 10))
        entry_scan = ctk.CTkEntry(row_sc)
        entry_scan.insert(0, cfg.get("escaneos_folder", ""))
        entry_scan.pack(side="left", fill="x", expand=True, padx=(0, 6))
        def browse_scan():
            d = filedialog.askdirectory(title="Seleccionar bandeja de escaneos")
            if d:
                entry_scan.delete(0, "end")
                entry_scan.insert(0, d)
        ctk.CTkButton(row_sc, text="Examinar...", width=95, command=browse_scan).pack(side="right")

        ctk.CTkLabel(tab_paths, text="Carpeta base de destino (donde se crearán las carpetas de las OFs):", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(4, 2))
        row_dst = ctk.CTkFrame(tab_paths, fg_color="transparent")
        row_dst.pack(fill="x", pady=(0, 10))
        entry_dest = ctk.CTkEntry(row_dst)
        entry_dest.insert(0, cfg.get("destination_base_folder", ""))
        entry_dest.pack(side="left", fill="x", expand=True, padx=(0, 6))
        def browse_dest():
            d = filedialog.askdirectory(title="Seleccionar carpeta base destino")
            if d:
                entry_dest.delete(0, "end")
                entry_dest.insert(0, d)
        ctk.CTkButton(row_dst, text="Examinar...", width=95, command=browse_dest).pack(side="right")

        ctk.CTkLabel(tab_paths, text="Carpeta base de Oficina Técnica (red de proyectos):", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(4, 2))
        row_ot = ctk.CTkFrame(tab_paths, fg_color="transparent")
        row_ot.pack(fill="x", pady=(0, 10))
        entry_ot = ctk.CTkEntry(row_ot)
        entry_ot.insert(0, cfg.get("oficina_tecnica_base", r"\\172.30.0.10\Oficina_tecnica"))
        entry_ot.pack(side="left", fill="x", expand=True, padx=(0, 6))
        def browse_ot():
            d = filedialog.askdirectory(title="Seleccionar carpeta base de Oficina Técnica")
            if d:
                entry_ot.delete(0, "end")
                entry_ot.insert(0, d)
        ctk.CTkButton(row_ot, text="Examinar...", width=95, command=browse_ot).pack(side="right")

        copy_ot_var = ctk.BooleanVar(value=cfg.get("copy_to_oficina_tecnica", True))
        chk_ot = ctk.CTkSwitch(
            tab_paths,
            text="Replicar copia automáticamente en '05-Documentacion final' de Oficina Técnica",
            variable=copy_ot_var,
            font=ctk.CTkFont(weight="bold")
        )
        chk_ot.pack(anchor="w", pady=(6, 10))

        # --- TAB 2: OCR y Escaneo ---
        ocr_enabled_var = ctk.BooleanVar(value=cfg.get("ocr_enabled", True))
        chk_ocr = ctk.CTkSwitch(
            tab_ocr,
            text="Activar reconocimiento automático OCR (Windows Native Media.Ocr)",
            variable=ocr_enabled_var,
            font=ctk.CTkFont(weight="bold")
        )
        chk_ocr.pack(anchor="w", pady=(10, 8))

        duplex_var = ctk.BooleanVar(value=cfg.get("duplex_mode", True))
        chk_duplex = ctk.CTkSwitch(
            tab_ocr,
            text="Modo Dúplex (Doble Cara) con descarte automático de reversos vacíos",
            variable=duplex_var,
            font=ctk.CTkFont(weight="bold")
        )
        chk_duplex.pack(anchor="w", pady=(4, 12))

        row_header = ctk.CTkFrame(tab_ocr, fg_color="transparent")
        row_header.pack(fill="x", pady=6)
        ctk.CTkLabel(row_header, text="Altura de cabecera para detectar OF (% de la pág 1):").pack(side="left", padx=(0, 8))
        entry_head = ctk.CTkEntry(row_header, width=60)
        entry_head.insert(0, str(cfg.get("ocr_header_height_percent", 30)))
        entry_head.pack(side="left")
        ctk.CTkLabel(row_header, text="(Recomendado: 30%)", text_color="gray").pack(side="left", padx=8)

        row_h = ctk.CTkFrame(tab_ocr, fg_color="transparent")
        row_h.pack(fill="x", pady=6)
        ctk.CTkLabel(row_h, text="Altura de pie de página para códigos (% inferior):").pack(side="left", padx=(0, 8))
        entry_h = ctk.CTkEntry(row_h, width=60)
        entry_h.insert(0, str(cfg.get("ocr_footer_height_percent", 15)))
        entry_h.pack(side="left")
        ctk.CTkLabel(row_h, text="(Recomendado: 12% - 15%)", text_color="gray").pack(side="left", padx=8)

        ctk.CTkLabel(tab_ocr, text="Ruta del Excel de códigos oficiales (codigosOCR.xlsx):", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(10, 2))
        row_cat = ctk.CTkFrame(tab_ocr, fg_color="transparent")
        row_cat.pack(fill="x", pady=(0, 10))
        entry_cat = ctk.CTkEntry(row_cat)
        entry_cat.insert(0, cfg.get("ocr_catalog_path", "codigosOCR.xlsx"))
        entry_cat.pack(side="left", fill="x", expand=True, padx=(0, 6))
        def browse_cat():
            f = filedialog.askopenfilename(title="Seleccionar catálogo de códigos", filetypes=[("Excel", "*.xlsx")])
            if f:
                entry_cat.delete(0, "end")
                entry_cat.insert(0, f)
        ctk.CTkButton(row_cat, text="Examinar...", width=95, command=browse_cat).pack(side="right")

        # Botones de Acción
        btn_box = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_box.pack(fill="x", padx=16, pady=(0, 14))

        def save_and_close():
            try:
                head_val = int(entry_head.get().strip())
            except ValueError:
                head_val = 30

            try:
                foot_val = int(entry_h.get().strip())
            except ValueError:
                foot_val = 15

            new_config = {
                "excel_path": entry_excel.get().strip(),
                "escaneos_folder": entry_scan.get().strip(),
                "destination_base_folder": entry_dest.get().strip(),
                "oficina_tecnica_base": entry_ot.get().strip(),
                "copy_to_oficina_tecnica": copy_ot_var.get(),
                "ocr_enabled": ocr_enabled_var.get(),
                "duplex_mode": duplex_var.get(),
                "ocr_header_height_percent": head_val,
                "ocr_footer_height_percent": foot_val,
                "ocr_catalog_path": entry_cat.get().strip()
            }

            cfg_curr = load_config()
            cfg_curr.update(new_config)
            if save_config(cfg_curr):
                dlg.destroy()
                self.reload_catalog(show_msg=False)
                self.check_system_connectivity()
                messagebox.showinfo("Configuración", "Configuración guardada correctamente.")
            else:
                messagebox.showerror("Error", "No se pudo guardar la configuración.")

        btn_cancel = ctk.CTkButton(btn_box, text="Cancelar", fg_color="gray", hover_color="gray30", width=90, command=dlg.destroy)
        btn_cancel.pack(side="right", padx=(8, 0))

        btn_save = ctk.CTkButton(
            btn_box,
            text="Guardar Cambios",
            fg_color="#059669",
            hover_color="#047857",
            text_color="#ffffff",
            font=ctk.CTkFont(weight="bold"),
            width=140,
            command=save_and_close
        )
        btn_save.pack(side="right")


if __name__ == "__main__":
    app = VeriDocsDesktopApp()
    app.mainloop()
