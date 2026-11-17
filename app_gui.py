import os
import sys
import time
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk

from config import load_config, save_config
from veridocs_core import (
    ExcelManager,
    ScanningSession,
    load_ocr_catalog,
    process_scanned_pdf,
    sanitize_filename,
    DOCUMENTS_DEF,
    get_effective_documents_def
)
from ocr_service import is_ocr_available

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class VeriDocsDesktopApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("VeriDocs - Gestión y Renombrado Automático de Escaneos MES")
        self.geometry("1180x820")
        self.minsize(1020, 720)

        self.excel_manager = ExcelManager()
        self.session_manager = ScanningSession(self.excel_manager)
        self.catalog = load_ocr_catalog()
        self.selected_order = None
        self.poll_timer_id = None
        self.last_batch_time = None

        # Iniciar escucha automática de inmediato al arrancar si está habilitado en configuración
        cfg = load_config()
        if cfg.get("auto_listen_on_startup", True):
            self.session_manager.start_listening()

        self.setup_ui()
        self.check_system_connectivity()

        # Cargar Excel en segundo plano al arrancar
        threading.Thread(target=self._initial_load_excel, daemon=True).start()

        # Iniciar sondeo regular del estado de la sesión
        self.poll_session_status()

    def _initial_load_excel(self):
        self.excel_manager.get_orders(force_reload=True)
        try:
            self.after(0, self.check_system_connectivity)
        except Exception:
            pass

    def setup_ui(self):
        # Header Frame
        header = ctk.CTkFrame(self, height=60, corner_radius=8)
        header.pack(fill="x", padx=16, pady=(12, 8))

        title_lbl = ctk.CTkLabel(
            header,
            text="📁 VeriDocs",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_lbl.pack(side="left", padx=16)

        subtitle_lbl = ctk.CTkLabel(
            header,
            text="Clasificación y troceado automático de dossiers por OF",
            font=ctk.CTkFont(size=13),
            text_color="gray"
        )
        subtitle_lbl.pack(side="left", padx=(0, 20))

        # Action & Indicators
        btn_settings = ctk.CTkButton(
            header,
            text="⚙️ Configuración",
            width=135,
            height=32,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#3b82f6",
            hover_color="#2563eb",
            command=self.open_settings_dialog
        )
        btn_settings.pack(side="right", padx=(8, 14))

        self.lbl_ocr_status = ctk.CTkLabel(
            header, text="OCR: Comprobando...", font=ctk.CTkFont(size=12, weight="bold")
        )
        self.lbl_ocr_status.pack(side="right", padx=10)

        self.lbl_folder_status = ctk.CTkLabel(
            header, text="Escaneos: Comprobando...", font=ctk.CTkFont(size=12, weight="bold")
        )
        self.lbl_folder_status.pack(side="right", padx=10)

        self.lbl_excel_status = ctk.CTkLabel(
            header, text="Excel: Comprobando...", font=ctk.CTkFont(size=12, weight="bold")
        )
        self.lbl_excel_status.pack(side="right", padx=10)

        # Alert Banner (Visible when active alert exists)
        self.alert_frame = ctk.CTkFrame(self, fg_color=("#b91c1c", "#7f1d1d"), corner_radius=8)
        # Not packed initially

        self.lbl_alert_icon = ctk.CTkLabel(
            self.alert_frame, text="⚠️", font=ctk.CTkFont(size=22)
        )
        self.lbl_alert_icon.pack(side="left", padx=(14, 8), pady=10)

        self.lbl_alert_msg = ctk.CTkLabel(
            self.alert_frame,
            text="",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="white",
            justify="left",
            wraplength=700
        )
        self.lbl_alert_msg.pack(side="left", padx=8, pady=10, fill="x", expand=True)

        self.btn_alert_folder = ctk.CTkButton(
            self.alert_frame,
            text="📂 Abrir Carpeta",
            width=120,
            height=30,
            fg_color="#f59e0b",
            hover_color="#d97706",
            font=ctk.CTkFont(weight="bold"),
            command=self.open_alert_folder
        )
        self.btn_alert_folder.pack(side="right", padx=(6, 10), pady=10)

        self.btn_alert_dismiss = ctk.CTkButton(
            self.alert_frame,
            text="✕ Descartar",
            width=90,
            height=30,
            fg_color="gray30",
            hover_color="gray40",
            font=ctk.CTkFont(weight="bold"),
            command=self.dismiss_current_alert
        )
        self.btn_alert_dismiss.pack(side="right", padx=(4, 6), pady=10)

        self.current_alert_data = None

        # Main Split Content
        self.main_content = ctk.CTkFrame(self, fg_color="transparent")
        self.main_content.pack(fill="both", expand=True, padx=16, pady=8)

        self.main_content.grid_columnconfigure(0, weight=4) # Left column
        self.main_content.grid_columnconfigure(1, weight=6) # Right column
        self.main_content.grid_rowconfigure(0, weight=1)

        # Left Column (Controls, Catalog & Aktis Query)
        left_col = ctk.CTkFrame(self.main_content)
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.setup_left_column(left_col)

        # Right Column (Latest Batch Results & Realtime Monitor)
        right_col = ctk.CTkFrame(self.main_content)
        right_col.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        self.setup_right_column(right_col)

    def setup_left_column(self, parent):
        parent.grid_columnconfigure(0, weight=1)

        # Card 1: Escucha Automática en Directo
        card_monitor = ctk.CTkFrame(parent, fg_color=("gray85", "gray17"), corner_radius=6)
        card_monitor.pack(fill="x", padx=12, pady=(12, 8))

        m_header = ctk.CTkFrame(card_monitor, fg_color="transparent")
        m_header.pack(fill="x", padx=12, pady=(10, 4))

        ctk.CTkLabel(
            m_header,
            text="1. Escucha de Escaneos",
            font=ctk.CTkFont(size=15, weight="bold")
        ).pack(side="left")

        is_active = self.session_manager.active
        self.lbl_session_badge = ctk.CTkLabel(
            m_header,
            text="ESCUCHANDO BANDEJA" if is_active else "INACTIVO",
            fg_color="#10b981" if is_active else "gray30",
            corner_radius=10,
            padx=10,
            font=ctk.CTkFont(size=11, weight="bold")
        )
        self.lbl_session_badge.pack(side="right")

        self.btn_toggle_listen = ctk.CTkButton(
            card_monitor,
            text="⏹ DETENER ESCUCHA" if is_active else "▶ INICIAR ESCUCHA AUTOMÁTICA",
            height=44,
            fg_color="#ef4444" if is_active else "#059669",
            hover_color="#dc2626" if is_active else "#047857",
            text_color="#ffffff",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.toggle_listening
        )
        self.btn_toggle_listen.pack(fill="x", padx=12, pady=6)

        self.lbl_listen_info = ctk.CTkLabel(
            card_monitor,
            text="● Modo continuo: al escanear un dossier, VeriDocs leerá la OF de la pág. 1, consultará Aktis, clasificará y troceará según 'codigosOCR.xlsx'.",
            font=ctk.CTkFont(size=11),
            text_color="gray",
            wraplength=420,
            justify="left"
        )
        self.lbl_listen_info.pack(anchor="w", padx=12, pady=(0, 6))

        # Botón de procesamiento manual
        btn_manual = ctk.CTkButton(
            card_monitor,
            text="📄 Procesar Archivo PDF Manualmente...",
            height=32,
            fg_color=("#3b82f6", "#2563eb"),
            hover_color=("#1d4ed8", "#1d4ed8"),
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.process_pdf_manually
        )
        btn_manual.pack(fill="x", padx=12, pady=(4, 10))

        # Card 2: Catálogo de Documentos (codigosOCR.xlsx)
        card_catalog = ctk.CTkFrame(parent, fg_color=("gray85", "gray17"), corner_radius=6)
        card_catalog.pack(fill="both", expand=True, padx=12, pady=6)

        c_header = ctk.CTkFrame(card_catalog, fg_color="transparent")
        c_header.pack(fill="x", padx=10, pady=(10, 4))

        self.lbl_catalog_title = ctk.CTkLabel(
            c_header,
            text=f"2. Catálogo Oficial ({len(self.catalog)} tipos)",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.lbl_catalog_title.pack(side="left")

        btn_reload_cat = ctk.CTkButton(
            c_header,
            text="↻ Recargar",
            width=75,
            height=24,
            command=self.reload_catalog
        )
        btn_reload_cat.pack(side="right")

        self.catalog_scroll = ctk.CTkScrollableFrame(card_catalog, height=180)
        self.catalog_scroll.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.render_catalog_list()

        # Card 3: Buscador de Consulta en Aktis (Opcional)
        card_aktis = ctk.CTkFrame(parent, fg_color=("gray85", "gray17"), corner_radius=6)
        card_aktis.pack(fill="x", padx=12, pady=(6, 12))

        a_header = ctk.CTkFrame(card_aktis, fg_color="transparent")
        a_header.pack(fill="x", padx=10, pady=(8, 2))

        ctk.CTkLabel(
            a_header,
            text="3. Consulta en Aktis (Opcional)",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(side="left")

        btn_reload_ak = ctk.CTkButton(
            a_header, text="↻ Recargar Excel", width=95, height=22,
            font=ctk.CTkFont(size=11),
            command=self.reload_excel
        )
        btn_reload_ak.pack(side="right")

        search_row = ctk.CTkFrame(card_aktis, fg_color="transparent")
        search_row.pack(fill="x", padx=10, pady=4)

        self.of_entry = ctk.CTkEntry(
            search_row,
            placeholder_text="Buscar OF (ej: AYC/MO/00560 o 560)...",
            height=30
        )
        self.of_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.of_entry.bind("<KeyRelease>", self.on_of_search_typing)
        self.of_entry.bind("<Return>", lambda e: self.do_search_of())

        btn_search = ctk.CTkButton(
            search_row, text="Buscar", width=65, height=30,
            command=self.do_search_of
        )
        btn_search.pack(side="right")

        self.results_frame = ctk.CTkScrollableFrame(card_aktis, height=80)
        # Pack only on matches

        self.details_box = ctk.CTkFrame(card_aktis, fg_color=("gray90", "gray13"))
        self.details_box.pack(fill="x", padx=10, pady=(4, 8))

        self.lbl_detail_of = ctk.CTkLabel(self.details_box, text="OF: -", font=ctk.CTkFont(weight="bold", size=12))
        self.lbl_detail_of.grid(row=0, column=0, sticky="w", padx=6, pady=1)

        self.lbl_detail_odoo = ctk.CTkLabel(self.details_box, text="Pedido ODOO: -", font=ctk.CTkFont(size=11))
        self.lbl_detail_odoo.grid(row=0, column=1, sticky="w", padx=6, pady=1)

        self.lbl_detail_desc = ctk.CTkLabel(self.details_box, text="Descripción: -", font=ctk.CTkFont(size=11), wraplength=400, justify="left")
        self.lbl_detail_desc.grid(row=1, column=0, columnspan=2, sticky="w", padx=6, pady=1)

    def render_catalog_list(self):
        for w in self.catalog_scroll.winfo_children():
            w.destroy()

        for d in self.catalog:
            code = d.get("codigo", "")
            doc_name = d.get("documento", "")
            is_inc = d.get("is_incidencia", False)

            row = ctk.CTkFrame(self.catalog_scroll, fg_color=("gray90", "gray19"), height=28)
            row.pack(fill="x", pady=1, padx=2)

            # Badge del código
            code_bg = "#dc2626" if is_inc else ("#4338ca" if "ANEXO" in code else "#2563eb")
            code_badge = ctk.CTkLabel(
                row,
                text=code,
                font=ctk.CTkFont(family="Consolas", size=11, weight="bold"),
                fg_color=code_bg,
                corner_radius=4,
                padx=6
            )
            code_badge.pack(side="left", padx=4, pady=2)

            name_lbl = ctk.CTkLabel(
                row,
                text=doc_name + (" [Páginas individuales (1-N)]" if is_inc else ""),
                font=ctk.CTkFont(size=11, weight="bold" if is_inc else "normal"),
                text_color="#f87171" if is_inc else ("black", "white"),
                anchor="w"
            )
            name_lbl.pack(side="left", fill="x", expand=True, padx=4)

    def reload_catalog(self):
        self.catalog = load_ocr_catalog()
        self.lbl_catalog_title.configure(text=f"2. Catálogo Oficial ({len(self.catalog)} tipos)")
        self.render_catalog_list()
        messagebox.showinfo("Catálogo", f"Se han cargado {len(self.catalog)} tipos de documentos desde codigosOCR.xlsx.")

    def setup_right_column(self, parent):
        parent.grid_columnconfigure(0, weight=1)

        # Header 4: Lotes y Monitor
        header4 = ctk.CTkFrame(parent, fg_color="transparent")
        header4.pack(fill="x", padx=12, pady=(12, 4))

        ctk.CTkLabel(
            header4,
            text="Monitor en Tiempo Real y Lotes Procesados",
            font=ctk.CTkFont(size=15, weight="bold")
        ).pack(side="left")

        self.lbl_batch_counter = ctk.CTkLabel(
            header4,
            text="Lotes hoy: 0",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="gray"
        )
        self.lbl_batch_counter.pack(side="right")

        # Card: Último Lote Procesado
        self.last_batch_card = ctk.CTkFrame(parent, fg_color=("gray85", "gray17"), corner_radius=6)
        self.last_batch_card.pack(fill="x", padx=12, pady=(4, 6))

        batch_header = ctk.CTkFrame(self.last_batch_card, fg_color="transparent")
        batch_header.pack(fill="x", padx=12, pady=(10, 4))

        self.lbl_batch_title = ctk.CTkLabel(
            batch_header,
            text="Último Lote: Ninguno procesado todavía",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.lbl_batch_title.pack(side="left")

        self.btn_open_batch_folder = ctk.CTkButton(
            batch_header,
            text="📂 Abrir Carpeta de la OF",
            height=30,
            fg_color="#059669",
            hover_color="#047857",
            text_color="#ffffff",
            text_color_disabled="#ffffff",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.open_current_batch_folder
        )
        self.btn_open_batch_folder.pack(side="right")

        self.lbl_batch_sub = ctk.CTkLabel(
            self.last_batch_card,
            text="Escucha automática activa. Al colocar un escaneo en la bandeja, se procesará automáticamente.",
            font=ctk.CTkFont(size=12),
            text_color="gray",
            justify="left"
        )
        self.lbl_batch_sub.pack(anchor="w", padx=12, pady=(0, 8))

        # Panel de Progreso en Vivo (visible mientras se procesa un escaneo)
        self.progress_frame = ctk.CTkFrame(parent, fg_color=("#dbeafe", "#1e3a5f"), corner_radius=6)
        # Not packed initially

        self.lbl_progress_text = ctk.CTkLabel(
            self.progress_frame,
            text="⚙️ Procesando escaneo...",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=("#1e40af", "#93c5fd"),
            anchor="w",
            justify="left"
        )
        self.lbl_progress_text.pack(fill="x", padx=12, pady=(8, 4))

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, height=10, corner_radius=5)
        self.progress_bar.pack(fill="x", padx=12, pady=(0, 10))
        self.progress_bar.set(0)

        # Scrollable de Archivos Generados
        self.lbl_files_header = ctk.CTkLabel(
            parent,
            text="Documentos PDF Generados del Lote:",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.lbl_files_header.pack(anchor="w", padx=14, pady=(4, 2))

        self.files_scroll = ctk.CTkScrollableFrame(parent, height=260)
        self.files_scroll.pack(fill="both", expand=True, padx=12, pady=(0, 6))

        # Log de eventos
        ctk.CTkLabel(
            parent,
            text="Registro de Actividad en Directo:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(anchor="w", padx=14, pady=(4, 2))

        self.log_scroll = ctk.CTkScrollableFrame(parent, height=130)
        self.log_scroll.pack(fill="x", padx=12, pady=(0, 10))

    def toggle_listening(self):
        if not self.session_manager.active:
            ok, msg = self.session_manager.start_listening()
            if ok:
                self.lbl_session_badge.configure(text="ESCUCHANDO BANDEJA", fg_color="#10b981")
                self.btn_toggle_listen.configure(
                    text="⏹ DETENER ESCUCHA",
                    fg_color="#ef4444",
                    hover_color="#dc2626",
                    text_color="#ffffff"
                )
            else:
                messagebox.showerror("Error al iniciar", msg)
        else:
            self.session_manager.stop_listening()
            self.lbl_session_badge.configure(text="INACTIVO", fg_color="gray30")
            self.btn_toggle_listen.configure(
                text="▶ INICIAR ESCUCHA AUTOMÁTICA",
                fg_color="#059669",
                hover_color="#047857",
                text_color="#ffffff"
            )

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
            self.session_manager.update_progress(
                is_processing=True,
                filename=fname,
                phase="Iniciando análisis manual...",
                current_page=0,
                total_pages=0,
                created_files=[]
            )

            def _on_prog(pdata):
                self.session_manager.update_progress(
                    is_processing=True,
                    filename=fname,
                    of=pdata.get("of", ""),
                    phase=pdata.get("phase", ""),
                    current_page=pdata.get("current_page", 0),
                    total_pages=pdata.get("total_pages", 0),
                    created_files=pdata.get("created_files", [])
                )

            res = process_scanned_pdf(
                pdf_path=file_path,
                excel_manager=self.excel_manager,
                catalog=self.catalog,
                log_callback=self.session_manager.log,
                progress_callback=_on_prog
            )
            self.session_manager.update_progress(is_processing=False)

            if res["success"]:
                self.session_manager.last_batch = res
                self.session_manager.batches_history.append({
                    "time": datetime.now().strftime("%H:%M:%S"),
                    "of": res["of"],
                    "order_info": res["order_info"],
                    "dest_dir": res["destination_dir"],
                    "total_pages": res["total_pages"],
                    "files_count": len(res["created_files"]),
                    "created_files": res["created_files"],
                    "alert": res["alert"]
                })
                if res.get("alert"):
                    self.session_manager.add_alert(
                        title=f"Aviso en OF {res['of']}",
                        message=res["alert"],
                        level="warning",
                        dest_dir=res["destination_dir"]
                    )
                self.after(0, lambda: messagebox.showinfo("Éxito", f"Lote de la OF '{res['of']}' procesado con éxito.\nSe han creado {len(res['created_files'])} archivos PDF."))
            else:
                err = res.get("error", "Error desconocido")
                self.session_manager.add_alert(
                    title="Error en procesamiento manual",
                    message=err,
                    level="error"
                )
                self.after(0, lambda: messagebox.showerror("Error", err))

        threading.Thread(target=_worker, daemon=True).start()

    def poll_session_status(self):
        try:
            status = self.session_manager.get_status()

            # Actualizar botón según estado real
            if status["active"]:
                self.lbl_session_badge.configure(text="ESCUCHANDO BANDEJA", fg_color="#10b981")
                self.btn_toggle_listen.configure(
                    text="⏹ DETENER ESCUCHA",
                    fg_color="#ef4444",
                    hover_color="#dc2626",
                    text_color="#ffffff"
                )
            else:
                self.lbl_session_badge.configure(text="INACTIVO", fg_color="gray30")
                self.btn_toggle_listen.configure(
                    text="▶ INICIAR ESCUCHA AUTOMÁTICA",
                    fg_color="#059669",
                    hover_color="#047857",
                    text_color="#ffffff"
                )

            self.lbl_batch_counter.configure(text=f"Lotes procesados: {status['total_batches_processed']}")

            # Manejar alertas activas
            active_alerts = status.get("active_alerts", [])
            if active_alerts:
                latest_alert = active_alerts[-1]
                self.current_alert_data = latest_alert
                self.lbl_alert_msg.configure(text=f"{latest_alert['title']}: {latest_alert['message']}")
                if not self.alert_frame.winfo_ismapped():
                    self.alert_frame.pack(fill="x", padx=16, pady=(0, 6), before=self.main_content)
            else:
                if self.alert_frame.winfo_ismapped():
                    self.alert_frame.pack_forget()

            # Progreso en tiempo real si hay un escaneo en curso
            curr_prog = status.get("current_progress", {})
            is_proc = curr_prog.get("is_processing", False)

            if is_proc:
                if not self.progress_frame.winfo_ismapped():
                    self.progress_frame.pack(fill="x", padx=12, pady=(0, 6), before=self.lbl_files_header)

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

                # Renderizar los archivos creados EN DIRECTO a medida que se van generando
                live_files = curr_prog.get("created_files", [])
                current_widgets = len(self.files_scroll.winfo_children())
                if current_widgets != len(live_files):
                    for w in self.files_scroll.winfo_children():
                        w.destroy()

                    for f in live_files:
                        is_warn = f.get("is_warning", False)
                        card_bg = ("#fee2e2", "#450a0a") if is_warn else ("gray90", "gray19")
                        card = ctk.CTkFrame(self.files_scroll, fg_color=card_bg, height=42)
                        card.pack(fill="x", pady=2, padx=2)

                        icon = "⚠️ " if is_warn else "📄 "
                        pages_tag = f"  [Págs: {f.get('pages_list', [])}]"
                        doc_lbl = ctk.CTkLabel(
                            card,
                            text=f"{icon}{f.get('doc_name', '')}{pages_tag}",
                            font=ctk.CTkFont(weight="bold", size=12),
                            text_color="#ef4444" if is_warn else ("black", "white")
                        )
                        doc_lbl.pack(anchor="w", padx=8, pady=(3, 0))

                        file_lbl = ctk.CTkLabel(
                            card,
                            text=f.get("filename", ""),
                            font=ctk.CTkFont(family="Consolas", size=10),
                            text_color="gray"
                        )
                        file_lbl.pack(anchor="w", padx=8, pady=(0, 3))
            else:
                if self.progress_frame.winfo_ismapped():
                    self.progress_frame.pack_forget()

                # Renderizar último lote procesado si ha cambiado
                last_b = status.get("last_batch")
                if last_b:
                    of_str = last_b.get("of", "")
                    o_info = last_b.get("order_info", {})
                    col_e = o_info.get("pedido_odoo", "")
                    col_i = o_info.get("descripcion", "")
                    dest_dir = last_b.get("destination_dir", "")

                    title_text = f"OF: {of_str}  |  Odoo: {col_e or '(N/A)'}  |  Total: {last_b.get('total_pages', 0)} páginas"
                    self.lbl_batch_title.configure(text=title_text)
                    self.lbl_batch_sub.configure(
                        text=f"Descripción: {col_i or '(Sin descripción)'}\nCarpeta: {dest_dir}"
                    )
                    self.btn_open_batch_folder.configure(
                        state="normal",
                        fg_color="#059669",
                        hover_color="#047857",
                        text_color="#ffffff"
                    )

                    # Renderizar archivos creados si es la primera vez que vemos este lote
                    created_files = last_b.get("created_files", [])
                    current_widgets = len(self.files_scroll.winfo_children())
                    if current_widgets != len(created_files):
                        for w in self.files_scroll.winfo_children():
                            w.destroy()

                        for f in created_files:
                            is_warn = f.get("is_warning", False)
                            card_bg = ("#fee2e2", "#450a0a") if is_warn else ("gray90", "gray19")
                            card = ctk.CTkFrame(self.files_scroll, fg_color=card_bg, height=42)
                            card.pack(fill="x", pady=2, padx=2)

                            icon = "⚠️ " if is_warn else "📄 "
                            pages_tag = f"  [Págs: {f.get('pages_list', [])}]"
                            doc_lbl = ctk.CTkLabel(
                                card,
                                text=f"{icon}{f.get('doc_name', '')}{pages_tag}",
                                font=ctk.CTkFont(weight="bold", size=12),
                                text_color="#ef4444" if is_warn else ("black", "white")
                            )
                            doc_lbl.pack(anchor="w", padx=8, pady=(3, 0))

                            file_lbl = ctk.CTkLabel(
                                card,
                                text=f.get("filename", ""),
                                font=ctk.CTkFont(family="Consolas", size=10),
                                text_color="gray"
                            )
                            file_lbl.pack(anchor="w", padx=8, pady=(0, 3))

            # Renderizar últimos logs
            logs = status.get("events_log", [])
            if len(self.log_scroll.winfo_children()) != len(logs):
                for w in self.log_scroll.winfo_children():
                    w.destroy()
                for entry in logs:
                    t = entry.get("time", "")
                    m = entry.get("message", "")
                    lvl = entry.get("level", "info")

                    color = "#10b981" if lvl == "success" else ("#f59e0b" if lvl == "warning" else ("#ef4444" if lvl == "error" else "gray"))
                    row = ctk.CTkLabel(
                        self.log_scroll,
                        text=f"[{t}] {m}",
                        font=ctk.CTkFont(family="Consolas", size=10),
                        text_color=color,
                        anchor="w"
                    )
                    row.pack(fill="x", pady=1, padx=4)

        except Exception as e:
            print(f"[poll_session_status] Error: {e}")
        finally:
            # Reprogramar siguiente comprobación de forma garantizada
            self.after(1000, self.poll_session_status)

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
                messagebox.showinfo("Aviso", f"La carpeta aún no está disponible: {folder}")

    def open_current_batch_folder(self):
        status = self.session_manager.get_status()
        last_b = status.get("last_batch")
        if last_b and last_b.get("destination_dir"):
            folder = last_b["destination_dir"]
            if os.path.exists(folder):
                os.startfile(folder)
            else:
                messagebox.showinfo("Aviso", f"La carpeta no existe: {folder}")
        else:
            messagebox.showinfo("Información", "Aún no se ha procesado ningún lote de escaneo en esta sesión.\n\nEn cuanto se procese un dossier, podrás abrir su carpeta directamente desde aquí.")

    # Búsqueda de OF en Aktis (Opcional para consulta)
    def on_of_search_typing(self, event=None):
        if event and event.keysym in ("Return", "KP_Enter", "Up", "Down"):
            return
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
                height=24,
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
        self.lbl_detail_desc.configure(text=f"Descripción: {order['descripcion'] or '(Sin descripción)'}")

    def reload_excel(self):
        def _reload():
            orders = self.excel_manager.get_orders(force_reload=True)
            self.after(0, lambda: messagebox.showinfo("Excel", f"Se han cargado {len(orders)} órdenes correctamente."))
            self.after(0, self.check_system_connectivity)
        threading.Thread(target=_reload, daemon=True).start()

    def check_system_connectivity(self):
        cfg = load_config()
        excel_path = cfg.get("excel_path", "")
        escaneos_folder = cfg.get("escaneos_folder", "")

        ex_ok = os.path.exists(excel_path)
        fo_ok = os.path.exists(escaneos_folder)

        self.lbl_excel_status.configure(
            text="Excel: Conectado" if ex_ok else "Excel: No accesible",
            text_color="#10b981" if ex_ok else "#ef4444"
        )
        self.lbl_folder_status.configure(
            text="Escaneos: OK" if fo_ok else "Escaneos: No accesible",
            text_color="#10b981" if fo_ok else "#ef4444"
        )

        ocr_ok, _ = is_ocr_available()
        ocr_en = cfg.get("ocr_enabled", True)
        if not ocr_ok:
            self.lbl_ocr_status.configure(text="OCR: No disp.", text_color="#ef4444")
        elif not ocr_en:
            self.lbl_ocr_status.configure(text="OCR: Inactivo", text_color="#f59e0b")
        else:
            self.lbl_ocr_status.configure(text="OCR: Activo", text_color="#10b981")

    def open_settings_dialog(self):
        cfg = load_config()
        dlg = ctk.CTkToplevel(self)
        dlg.title("Configuración de VeriDocs")
        dlg.geometry("640x520")
        dlg.minsize(550, 480)
        dlg.grab_set()

        tabview = ctk.CTkTabview(dlg)
        tabview.pack(fill="both", expand=True, padx=16, pady=(12, 8))

        tab_paths = tabview.add("📁 Rutas de Red")
        tab_ocr = tabview.add("🔍 Parámetros OCR")

        # --- TAB 1: RUTAS ---
        ctk.CTkLabel(tab_paths, text="Ruta del archivo Excel (Pedidos en curso):", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(8, 2))
        entry_excel = ctk.CTkEntry(tab_paths, width=580)
        entry_excel.insert(0, cfg.get("excel_path", ""))
        entry_excel.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(tab_paths, text="Carpeta donde caen los PDFs del escáner:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(4, 2))
        entry_scan = ctk.CTkEntry(tab_paths, width=580)
        entry_scan.insert(0, cfg.get("escaneos_folder", ""))
        entry_scan.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(tab_paths, text="Carpeta base de destino (donde se crea la carpeta de la OF):", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(4, 2))
        entry_dest = ctk.CTkEntry(tab_paths, width=580)
        entry_dest.insert(0, cfg.get("destination_base_folder", ""))
        entry_dest.pack(fill="x", pady=(0, 10))

        # --- TAB 2: OCR ---
        ocr_enabled_var = ctk.BooleanVar(value=cfg.get("ocr_enabled", True))
        chk_ocr = ctk.CTkSwitch(
            tab_ocr,
            text="Activar reconocimiento automático OCR (Windows Native Media.Ocr)",
            variable=ocr_enabled_var,
            font=ctk.CTkFont(weight="bold")
        )
        chk_ocr.pack(anchor="w", pady=(10, 12))

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

        ctk.CTkLabel(tab_ocr, text="Ruta del Excel de códigos oficiales:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(10, 2))
        entry_cat = ctk.CTkEntry(tab_ocr, width=580)
        entry_cat.insert(0, cfg.get("ocr_catalog_path", "codigosOCR.xlsx"))
        entry_cat.pack(fill="x", pady=(0, 10))

        # Action Buttons
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
                "ocr_enabled": ocr_enabled_var.get(),
                "ocr_header_height_percent": head_val,
                "ocr_footer_height_percent": foot_val,
                "ocr_catalog_path": entry_cat.get().strip()
            }

            cfg_curr = load_config()
            cfg_curr.update(new_config)
            if save_config(cfg_curr):
                dlg.destroy()
                self.reload_catalog()
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
