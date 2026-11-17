// VeriDocs - Client Side Application Logic (Batch & OCR Workflow)
let catalogData = [];
let lastBatchData = null;
let currentAlertId = null;

// Audio Chime Generator using Web Audio API
function playChime(isSuccess = true) {
  try {
    const AudioContext = window.AudioContext || window.webkitAudioContext;
    if (!AudioContext) return;
    const ctx = new AudioContext();

    if (isSuccess) {
      const freqs = [523.25, 659.25, 783.99, 1046.50];
      freqs.forEach((freq, idx) => {
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'triangle';
        osc.frequency.value = freq;
        gain.gain.setValueAtTime(0.15, ctx.currentTime + idx * 0.1);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + idx * 0.1 + 0.6);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start(ctx.currentTime + idx * 0.1);
        osc.stop(ctx.currentTime + idx * 0.1 + 0.6);
      });
    } else {
      // Warning double beep
      [440, 370].forEach((freq, idx) => {
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'sawtooth';
        osc.frequency.value = freq;
        gain.gain.setValueAtTime(0.2, ctx.currentTime + idx * 0.18);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + idx * 0.18 + 0.3);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start(ctx.currentTime + idx * 0.18);
        osc.stop(ctx.currentTime + idx * 0.18 + 0.3);
      });
    }
  } catch (e) {
    console.warn("Audio chime error:", e);
  }
}

// Initialization
document.addEventListener("DOMContentLoaded", async () => {
  setupEventListeners();
  await loadCatalog();
  await checkSystemInfo();
  setInterval(checkSystemInfo, 15000);
  setInterval(pollSession, 1200);
  pollSession();
});

function setupEventListeners() {
  // Toggle Watcher
  const btnToggle = document.getElementById("btn-toggle-watcher");
  if (btnToggle) {
    btnToggle.addEventListener("click", toggleWatcher);
  }

  // Manual Upload
  const fileInput = document.getElementById("manual-pdf-upload");
  if (fileInput) {
    fileInput.addEventListener("change", handleManualUpload);
  }

  // Reload Catalog
  const btnReloadCat = document.getElementById("btn-reload-catalog");
  if (btnReloadCat) {
    btnReloadCat.addEventListener("click", loadCatalog);
  }

  // Reload Excel
  const btnReloadExcel = document.getElementById("btn-reload-excel");
  if (btnReloadExcel) {
    btnReloadExcel.addEventListener("click", async () => {
      try {
        const res = await fetch("/api/orders/reload", { method: "POST" });
        const d = await res.json();
        alert(`Se han recargado ${d.count} órdenes desde Aktis.`);
        checkSystemInfo();
      } catch (e) {
        alert("Error recargando Excel: " + e);
      }
    });
  }

  // Alert Actions
  const btnAlertDismiss = document.getElementById("btn-alert-dismiss");
  if (btnAlertDismiss) {
    btnAlertDismiss.addEventListener("click", dismissAlert);
  }

  const btnAlertOpen = document.getElementById("btn-alert-open-folder");
  if (btnAlertOpen) {
    btnAlertOpen.addEventListener("click", openLastBatchFolder);
  }

  const btnBatchOpen = document.getElementById("btn-open-batch-folder");
  if (btnBatchOpen) {
    btnBatchOpen.addEventListener("click", openLastBatchFolder);
  }

  // Toggle Log
  const btnToggleLog = document.getElementById("btn-toggle-log");
  const logContainer = document.getElementById("events-log-container");
  if (btnToggleLog && logContainer) {
    btnToggleLog.addEventListener("click", () => {
      const isVisible = logContainer.style.display !== "none";
      logContainer.style.display = isVisible ? "none" : "block";
      document.getElementById("log-toggle-icon").textContent = isVisible ? "▲" : "▼";
    });
  }

  // Settings Modal
  setupSettingsModal();

  // Aktis OF search
  setupOfSearch();
}

async function loadCatalog() {
  try {
    const res = await fetch("/api/catalog");
    catalogData = await res.json();
    renderCatalogList();
  } catch (e) {
    console.error("Error loading catalog:", e);
  }
}

function renderCatalogList() {
  const container = document.getElementById("catalog-list");
  const countEl = document.getElementById("catalog-count");
  if (!container) return;

  container.innerHTML = "";
  if (countEl) countEl.textContent = catalogData.length;

  catalogData.forEach(item => {
    const code = item.codigo || "";
    const name = item.documento || "";
    const isInc = item.is_incidencia;

    const row = document.createElement("div");
    row.className = "generated-file-item";
    row.style.padding = "6px 10px";
    row.style.marginBottom = "4px";

    const badgeColor = isInc ? "#ef4444" : (code.includes("ANEXO") ? "#8b5cf6" : "#3b82f6");

    row.innerHTML = `
      <div style="display:flex; align-items:center; gap:8px;">
        <span style="background:${badgeColor}; color:white; font-family:'JetBrains Mono', monospace; font-size:0.75rem; font-weight:bold; padding:2px 6px; border-radius:4px;">
          ${code}
        </span>
        <span style="font-size:0.85rem; font-weight:${isInc ? 'bold' : 'normal'}; color:${isInc ? '#f87171' : 'inherit'};">
          ${name} ${isInc ? '<small style="color:#f87171;">(Págs indiv. 1-3)</small>' : ''}
        </span>
      </div>
    `;
    container.appendChild(row);
  });
}

async function toggleWatcher() {
  try {
    const res = await fetch("/api/watcher/toggle", { method: "POST" });
    const data = await res.json();
    pollSession();
  } catch (e) {
    alert("Error cambiando estado de escucha: " + e);
  }
}

async function handleManualUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const formData = new FormData();
  formData.append("file", file);

  const btnUpload = document.querySelector("label[for='manual-pdf-upload']");
  if (btnUpload) {
    btnUpload.textContent = "⏳ Procesando archivo...";
    btnUpload.style.pointerEvents = "none";
  }

  try {
    const res = await fetch("/api/process_pdf_upload", {
      method: "POST",
      body: formData
    });
    const data = await res.json();
    if (data.success) {
      playChime(true);
      alert(`¡Lote procesado con éxito!\nOF: ${data.of}\nGenerados ${data.created_files.length} archivos.`);
    } else {
      playChime(false);
      alert("Error procesando PDF: " + (data.error || "Desconocido"));
    }
  } catch (err) {
    alert("Error de subida: " + err);
  } finally {
    if (btnUpload) {
      btnUpload.textContent = "📄 Subir y Procesar Archivo PDF Manualmente...";
      btnUpload.style.pointerEvents = "auto";
    }
    e.target.value = "";
    pollSession();
  }
}

async function pollSession() {
  try {
    const res = await fetch("/api/session");
    const status = await res.json();
    updateUI(status);
  } catch (e) {
    console.warn("Poll session error:", e);
  }
}

function updateUI(status) {
  const badge = document.getElementById("session-status-badge");
  const btnToggle = document.getElementById("btn-toggle-watcher");
  const batchCount = document.getElementById("batch-counter-badge");

  if (status.active) {
    if (badge) {
      badge.textContent = "ESCUCHANDO BANDEJA";
      badge.className = "badge badge-active";
    }
    if (btnToggle) {
      btnToggle.innerHTML = '<span class="icon">⏹</span> DETENER ESCUCHA';
      btnToggle.className = "btn btn-danger btn-large";
    }
  } else {
    if (badge) {
      badge.textContent = "INACTIVO";
      badge.className = "badge badge-inactive";
    }
    if (btnToggle) {
      btnToggle.innerHTML = '<span class="icon">▶</span> INICIAR ESCUCHA AUTOMÁTICA';
      btnToggle.className = "btn btn-success btn-large";
    }
  }

  if (batchCount) {
    batchCount.textContent = `Lotes hoy: ${status.total_batches_processed || 0}`;
  }

  // Active Alerts
  const alertBanner = document.getElementById("active-alert-banner");
  const alertTitle = document.getElementById("alert-title");
  const alertMsg = document.getElementById("alert-message");

  if (status.active_alerts && status.active_alerts.length > 0) {
    const latest = status.active_alerts[status.active_alerts.length - 1];
    currentAlertId = latest.id;
    if (alertTitle) alertTitle.textContent = latest.title || "Atención";
    if (alertMsg) alertMsg.textContent = latest.message || "";
    if (alertBanner) alertBanner.style.display = "flex";
  } else {
    currentAlertId = null;
    if (alertBanner) alertBanner.style.display = "none";
  }

  // Last Batch
  const lastB = status.last_batch;
  if (lastB) {
    lastBatchData = lastB;
    const titleEl = document.getElementById("last-batch-title");
    const metaEl = document.getElementById("last-batch-meta");
    const openBtn = document.getElementById("btn-open-batch-folder");
    const filesSec = document.getElementById("generated-files-section");
    const filesList = document.getElementById("generated-files-list");

    const ofStr = lastB.of || "Desconocida";
    const odoo = (lastB.order_info && lastB.order_info.pedido_odoo) || "-";
    const desc = (lastB.order_info && lastB.order_info.descripcion) || "";
    const totalP = lastB.total_pages || 0;

    if (titleEl) titleEl.textContent = `OF: ${ofStr}  |  Odoo: ${odoo} (${totalP} páginas)`;
    if (metaEl) metaEl.textContent = `${desc ? desc + ' — ' : ''}Carpeta: ${lastB.destination_dir || ''}`;
    if (openBtn) openBtn.style.display = "inline-block";

    if (filesSec) filesSec.style.display = "block";
    if (filesList && lastB.created_files) {
      filesList.innerHTML = "";
      lastB.created_files.forEach(f => {
        const item = document.createElement("div");
        item.className = "generated-file-item" + (f.is_warning ? " warning" : "");
        item.innerHTML = `
          <div class="file-info">
            <span class="file-name">${f.is_warning ? '⚠️ ' : '📄 '}${f.doc_name}</span>
            <span class="file-sub">${f.filename}</span>
          </div>
          <span class="file-pages-badge">Págs: [${(f.pages_list || []).join(', ')}]</span>
        `;
        filesList.appendChild(item);
      });
    }
  }

  // Live Logs
  const logContainer = document.getElementById("events-log-container");
  if (logContainer && status.events_log) {
    logContainer.innerHTML = "";
    status.events_log.forEach(entry => {
      const row = document.createElement("div");
      row.className = `log-entry log-${entry.level || 'info'}`;
      row.textContent = `[${entry.time || ''}] ${entry.message || ''}`;
      logContainer.appendChild(row);
    });
    logContainer.scrollTop = logContainer.scrollHeight;
  }
}

async function dismissAlert() {
  if (!currentAlertId) return;
  try {
    await fetch("/api/alerts/dismiss", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id: currentAlertId })
    });
    pollSession();
  } catch (e) {
    console.error("Error dismissing alert:", e);
  }
}

async function openLastBatchFolder() {
  try {
    const res = await fetch("/api/session/open_folder", { method: "POST" });
    const d = await res.json();
    if (!d.success) {
      alert("Aviso: " + (d.error || "No se pudo abrir la carpeta."));
    }
  } catch (e) {
    alert("Error: " + e);
  }
}

async function checkSystemInfo() {
  try {
    const res = await fetch("/api/system_info");
    const d = await res.json();

    const exPill = document.getElementById("excel-status-pill");
    if (exPill) {
      const dot = exPill.querySelector(".status-dot");
      const lbl = exPill.querySelector(".status-label");
      if (dot) dot.className = `status-dot ${d.excel_ok ? 'dot-green' : 'dot-red'}`;
      if (lbl) lbl.textContent = `Excel: ${d.excel_ok ? 'Conectado' : 'No accesible'}`;
    }

    const foPill = document.getElementById("folder-status-pill");
    if (foPill) {
      const dot = foPill.querySelector(".status-dot");
      const lbl = foPill.querySelector(".status-label");
      if (dot) dot.className = `status-dot ${d.escaneos_ok ? 'dot-green' : 'dot-red'}`;
      if (lbl) lbl.textContent = `Escaneos: ${d.escaneos_ok ? 'OK' : 'No accesible'}`;
    }

    const ocrPill = document.getElementById("ocr-status-pill");
    if (ocrPill) {
      const dot = ocrPill.querySelector(".status-dot");
      const lbl = ocrPill.querySelector(".status-label");
      if (dot) dot.className = `status-dot ${d.ocr_available ? 'dot-green' : 'dot-red'}`;
      if (lbl) lbl.textContent = `OCR: ${d.ocr_available ? 'Activo' : 'No disp.'}`;
    }

    const ipEl = document.getElementById("cfg-server-ip");
    if (ipEl) ipEl.textContent = d.local_ip || "127.0.0.1";
    const portEl = document.getElementById("cfg-server-port");
    if (portEl) portEl.textContent = d.port || "5055";
  } catch (e) {
    console.warn("System info check failed:", e);
  }
}

function setupOfSearch() {
  const ofInput = document.getElementById("of-input");
  const clearBtn = document.getElementById("btn-clear-of");
  const autocompleteList = document.getElementById("autocomplete-list");
  const detailsCard = document.getElementById("order-details-card");

  let timer = null;
  ofInput.addEventListener("input", (e) => {
    const val = e.target.value.trim();
    if (clearBtn) clearBtn.style.display = val ? "block" : "none";
    clearTimeout(timer);

    if (!val) {
      autocompleteList.style.display = "none";
      if (detailsCard) detailsCard.style.display = "none";
      return;
    }

    timer = setTimeout(async () => {
      try {
        const res = await fetch(`/api/orders?q=${encodeURIComponent(val)}`);
        const orders = await res.json();
        autocompleteList.innerHTML = "";
        if (orders.length > 0) {
          autocompleteList.style.display = "block";
          orders.slice(0, 6).forEach(o => {
            const row = document.createElement("div");
            row.className = "autocomplete-item";
            row.textContent = `${o.of} | ${o.pedido_odoo} | ${o.descripcion.substring(0, 35)}`;
            row.addEventListener("click", () => {
              ofInput.value = o.of;
              autocompleteList.style.display = "none";
              document.getElementById("info-of").textContent = o.of;
              document.getElementById("info-pedido-odoo").textContent = o.pedido_odoo || "-";
              document.getElementById("info-desc").textContent = o.descripcion || "-";
              if (detailsCard) detailsCard.style.display = "block";
            });
            autocompleteList.appendChild(row);
          });
        } else {
          autocompleteList.style.display = "none";
        }
      } catch (err) {
        console.error("Order search error:", err);
      }
    }, 250);
  });

  if (clearBtn) {
    clearBtn.addEventListener("click", () => {
      ofInput.value = "";
      clearBtn.style.display = "none";
      autocompleteList.style.display = "none";
      if (detailsCard) detailsCard.style.display = "none";
    });
  }
}

function setupSettingsModal() {
  const modal = document.getElementById("settings-modal");
  const btnOpen = document.getElementById("btn-open-settings");
  const btnClose = document.getElementById("btn-close-settings");
  const btnSave = document.getElementById("btn-save-settings");

  if (btnOpen) {
    btnOpen.addEventListener("click", async () => {
      try {
        const res = await fetch("/api/config");
        const cfg = await res.json();
        document.getElementById("cfg-excel-path").value = cfg.excel_path || "";
        document.getElementById("cfg-escaneos-folder").value = cfg.escaneos_folder || "";
        document.getElementById("cfg-dest-folder").value = cfg.destination_base_folder || "";
        document.getElementById("cfg-ocr-enabled").checked = (cfg.ocr_enabled !== false);
        document.getElementById("cfg-ocr-header-height").value = cfg.ocr_header_height_percent || 30;
        document.getElementById("cfg-ocr-height").value = cfg.ocr_footer_height_percent || 15;
        document.getElementById("cfg-catalog-path").value = cfg.ocr_catalog_path || "codigosOCR.xlsx";
        modal.style.display = "flex";
      } catch (e) {
        alert("Error cargando configuración: " + e);
      }
    });
  }

  if (btnClose) {
    btnClose.addEventListener("click", () => {
      modal.style.display = "none";
    });
  }

  // Tabs
  const tabBtns = document.querySelectorAll(".tab-btn");
  tabBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      tabBtns.forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".tab-content").forEach(tc => tc.classList.remove("active"));
      btn.classList.add("active");
      const targetId = btn.getAttribute("data-tab");
      const target = document.getElementById(targetId);
      if (target) target.classList.add("active");
    });
  });

  if (btnSave) {
    btnSave.addEventListener("click", async () => {
      const newCfg = {
        excel_path: document.getElementById("cfg-excel-path").value.trim(),
        escaneos_folder: document.getElementById("cfg-escaneos-folder").value.trim(),
        destination_base_folder: document.getElementById("cfg-dest-folder").value.trim(),
        ocr_enabled: document.getElementById("cfg-ocr-enabled").checked,
        ocr_header_height_percent: parseInt(document.getElementById("cfg-ocr-header-height").value) || 30,
        ocr_footer_height_percent: parseInt(document.getElementById("cfg-ocr-height").value) || 15,
        ocr_catalog_path: document.getElementById("cfg-catalog-path").value.trim()
      };

      try {
        const res = await fetch("/api/config", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(newCfg)
        });
        const d = await res.json();
        if (d.success) {
          alert("Configuración guardada correctamente.");
          modal.style.display = "none";
          checkSystemInfo();
          loadCatalog();
        } else {
          alert("Error guardando: " + d.error);
        }
      } catch (e) {
        alert("Error de guardado: " + e);
      }
    });
  }
}
