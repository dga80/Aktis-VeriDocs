// VeriDocs - Client Side Application Logic (Modernized)
let catalogData = [];
let lastBatchData = null;
let currentAlertId = null;
let lastSeenEventId = 0;
let lastSeenBatchCount = -1;

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  const div = document.createElement('div');
  div.textContent = String(str);
  return div.innerHTML;
}

// Toast Notifications
function showToast(message, type = 'info', duration = 4000) {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `
    <span>${escapeHtml(message)}</span>
    <button style="background:none;border:none;color:white;cursor:pointer;font-weight:bold;margin-left:8px;">✕</button>
  `;

  const closeBtn = toast.querySelector('button');
  const removeToast = () => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(100%)';
    toast.style.transition = 'all 0.25s ease';
    setTimeout(() => toast.remove(), 250);
  };

  closeBtn.addEventListener('click', removeToast);
  container.appendChild(toast);
  setTimeout(removeToast, duration);
}

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
  setupDropzone();
  await loadCatalog();
  await checkSystemInfo();
  setInterval(checkSystemInfo, 12000);
  setInterval(pollSession, 1200);
  pollSession();
});

function setupEventListeners() {
  const btnToggle = document.getElementById("btn-toggle-watcher");
  if (btnToggle) btnToggle.addEventListener("click", toggleWatcher);

  const btnReloadCat = document.getElementById("btn-reload-catalog");
  if (btnReloadCat) btnReloadCat.addEventListener("click", loadCatalog);

  const catFilter = document.getElementById("cat-filter-input");
  if (catFilter) catFilter.addEventListener("input", filterCatalog);

  const btnReloadExcel = document.getElementById("btn-reload-excel");
  if (btnReloadExcel) {
    btnReloadExcel.addEventListener("click", async () => {
      try {
        const res = await fetch("/api/orders/reload", { method: "POST" });
        const d = await res.json();
        showToast(`Se han recargado ${d.count} órdenes desde Aktis.`, "success");
        checkSystemInfo();
      } catch (e) {
        showToast("Error recargando Excel: " + e, "error");
      }
    });
  }

  // Alert Actions
  const btnAlertDismiss = document.getElementById("btn-alert-dismiss");
  if (btnAlertDismiss) btnAlertDismiss.addEventListener("click", dismissAlert);

  const btnAlertClearAll = document.getElementById("btn-alert-clear-all");
  if (btnAlertClearAll) btnAlertClearAll.addEventListener("click", clearAllAlerts);

  const btnAlertOpen = document.getElementById("btn-alert-open-folder");
  if (btnAlertOpen) btnAlertOpen.addEventListener("click", openLastBatchFolder);

  const btnBatchOpen = document.getElementById("btn-open-batch-folder");
  if (btnBatchOpen) btnBatchOpen.addEventListener("click", openLastBatchFolder);

  const btnCopyFolder = document.getElementById("btn-copy-batch-folder");
  if (btnCopyFolder) btnCopyFolder.addEventListener("click", copyBatchFolder);

  // Toggle History & Log
  const btnToggleHistory = document.getElementById("btn-toggle-history");
  const historyContainer = document.getElementById("batches-history-container");
  if (btnToggleHistory && historyContainer) {
    btnToggleHistory.addEventListener("click", () => {
      const isVisible = historyContainer.style.display !== "none";
      historyContainer.style.display = isVisible ? "none" : "block";
      document.getElementById("history-toggle-icon").textContent = isVisible ? "▼" : "▲";
    });
  }

  const btnToggleLog = document.getElementById("btn-toggle-log");
  const logContainer = document.getElementById("events-log-container");
  if (btnToggleLog && logContainer) {
    btnToggleLog.addEventListener("click", (e) => {
      if (e.target.id === "btn-clear-log") return;
      const isVisible = logContainer.style.display !== "none";
      logContainer.style.display = isVisible ? "none" : "block";
      document.getElementById("log-toggle-icon").textContent = isVisible ? "▼" : "▲";
    });
  }

  const btnClearLog = document.getElementById("btn-clear-log");
  if (btnClearLog && logContainer) {
    btnClearLog.addEventListener("click", (e) => {
      e.stopPropagation();
      logContainer.innerHTML = '<div class="log-entry log-info">[Listo] Registro limpiado.</div>';
    });
  }

  setupSettingsModal();
  setupOfSearch();
}

function setupDropzone() {
  const dropzone = document.getElementById("pdf-dropzone");
  const fileInput = document.getElementById("manual-pdf-upload");
  if (!dropzone || !fileInput) return;

  dropzone.addEventListener("click", () => fileInput.click());

  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add("dragover");
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove("dragover");
    });
  });

  dropzone.addEventListener("drop", (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0 && files[0].name.toLowerCase().endsWith(".pdf")) {
      processUploadedPdf(files[0]);
    } else {
      showToast("Por favor arrastra un archivo PDF válido.", "warning");
    }
  });

  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
      processUploadedPdf(e.target.files[0]);
      e.target.value = "";
    }
  });
}

async function processUploadedPdf(file) {
  const formData = new FormData();
  formData.append("file", file);

  showToast(`Subiendo y procesando '${file.name}'...`, "info");
  const progContainer = document.getElementById("web-progress-container");
  const progLabel = document.getElementById("web-progress-label");
  const progFill = document.getElementById("web-progress-fill");
  const progPct = document.getElementById("web-progress-pct");

  if (progContainer) {
    progContainer.style.display = "block";
    progLabel.textContent = `Analizando ${file.name}...`;
    progFill.style.width = "20%";
    progPct.textContent = "20%";
  }

  try {
    const res = await fetch("/api/process_pdf_upload", {
      method: "POST",
      body: formData
    });
    const data = await res.json();
    if (data.success) {
      playChime(true);
      showToast(`¡Lote procesado con éxito! OF: ${data.of} (${data.created_files.length} archivos)`, "success");
      if (data.alert) {
        showToast(data.alert, "warning", 6000);
      }
    } else {
      playChime(false);
      showToast("Error procesando PDF: " + (data.error || "Desconocido"), "error", 6000);
    }
  } catch (err) {
    showToast("Error de conexión durante la subida: " + err, "error");
  } finally {
    if (progContainer) progContainer.style.display = "none";
    pollSession();
  }
}

async function loadCatalog() {
  try {
    const res = await fetch("/api/catalog");
    catalogData = await res.json();
    renderCatalogList(catalogData);
  } catch (e) {
    console.error("Error loading catalog:", e);
  }
}

function filterCatalog() {
  const q = (document.getElementById("cat-filter-input")?.value || "").toLowerCase().trim();
  if (!q) {
    renderCatalogList(catalogData);
    return;
  }
  const filtered = catalogData.filter(item =>
    (item.codigo || "").toLowerCase().includes(q) ||
    (item.documento || "").toLowerCase().includes(q)
  );
  renderCatalogList(filtered);
}

function renderCatalogList(items) {
  const container = document.getElementById("catalog-list");
  const countEl = document.getElementById("catalog-count");
  if (!container) return;

  container.innerHTML = "";
  if (countEl) countEl.textContent = items.length;

  items.forEach(item => {
    const code = escapeHtml(item.codigo || "");
    const name = escapeHtml(item.documento || "");
    const isInc = item.is_incidencia;

    const row = document.createElement("div");
    row.className = "generated-file-item";
    row.style.padding = "5px 8px";
    row.style.marginBottom = "3px";

    const badgeColor = isInc ? "#ef4444" : (code.includes("ANEXO") ? "#8b5cf6" : "#2563eb");

    row.innerHTML = `
      <div style="display:flex; align-items:center; gap:8px;">
        <span style="background:${badgeColor}; color:white; font-family:Consolas, monospace; font-size:0.75rem; font-weight:bold; padding:2px 6px; border-radius:4px;">
          ${code}
        </span>
        <span style="font-size:0.83rem; font-weight:${isInc ? '600' : 'normal'}; color:${isInc ? '#f87171' : 'inherit'};">
          ${name} ${isInc ? '<small style="color:#f87171;">(Págs indiv. 1-N)</small>' : ''}
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
    showToast(data.message || (data.active ? "Escucha iniciada" : "Escucha pausada"), data.active ? "success" : "info");
    pollSession();
  } catch (e) {
    showToast("Error cambiando estado de escucha: " + e, "error");
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
  // Watcher Button & Badge
  const badge = document.getElementById("session-status-badge");
  const btnToggle = document.getElementById("btn-toggle-watcher");
  const batchCount = document.getElementById("batch-counter-badge");

  if (status.active) {
    if (badge && badge.textContent !== "ESCUCHANDO BANDEJA") {
      badge.textContent = "ESCUCHANDO BANDEJA";
      badge.className = "badge badge-active";
    }
    if (btnToggle && !btnToggle.classList.contains("btn-danger")) {
      btnToggle.innerHTML = '<span class="icon">⏹</span> DETENER ESCUCHA';
      btnToggle.className = "btn btn-danger btn-large";
    }
  } else {
    if (badge && badge.textContent !== "INACTIVO") {
      badge.textContent = "INACTIVO";
      badge.className = "badge badge-inactive";
    }
    if (btnToggle && !btnToggle.classList.contains("btn-success")) {
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
  const alertCounter = document.getElementById("alert-counter");
  const btnClearAll = document.getElementById("btn-alert-clear-all");

  if (status.active_alerts && status.active_alerts.length > 0) {
    const latest = status.active_alerts[status.active_alerts.length - 1];
    currentAlertId = latest.id;
    if (alertTitle) alertTitle.textContent = latest.title || "Atención";
    if (alertMsg) alertMsg.textContent = latest.message || "";
    if (alertCounter) {
      if (status.active_alerts.length > 1) {
        alertCounter.textContent = `${status.active_alerts.length} alertas`;
        alertCounter.style.display = "inline-block";
        if (btnClearAll) btnClearAll.style.display = "inline-block";
      } else {
        alertCounter.style.display = "none";
        if (btnClearAll) btnClearAll.style.display = "none";
      }
    }
    if (alertBanner) alertBanner.style.display = "flex";
  } else {
    currentAlertId = null;
    if (alertBanner) alertBanner.style.display = "none";
  }

  // Live Progress Bar
  const progContainer = document.getElementById("web-progress-container");
  const progLabel = document.getElementById("web-progress-label");
  const progFill = document.getElementById("web-progress-fill");
  const progPct = document.getElementById("web-progress-pct");
  const currProg = status.current_progress;

  if (currProg && currProg.is_processing) {
    if (progContainer) progContainer.style.display = "block";
    const curP = currProg.current_page || 0;
    const totP = Math.max(1, currProg.total_pages || 1);
    const pct = Math.min(100, Math.round((curP / totP) * 100));
    const ofTag = currProg.of ? ` | OF: ${currProg.of}` : "";

    if (progLabel) progLabel.textContent = `⚙️ ${currProg.filename}${ofTag}: ${currProg.phase}`;
    if (progFill) progFill.style.width = `${pct}%`;
    if (progPct) progPct.textContent = `${pct}%`;
  } else {
    if (progContainer) progContainer.style.display = "none";
  }

  // Last Batch
  const lastB = status.last_batch;
  if (lastB) {
    lastBatchData = lastB;
    renderBatchView(lastB);
  }

  // Batches History
  const batches = status.batches_history || [];
  if (batches.length !== lastSeenBatchCount) {
    lastSeenBatchCount = batches.length;
    renderBatchesHistory(batches);
  }

  // Incremental Live Logs
  const logContainer = document.getElementById("events-log-container");
  if (logContainer && status.events_log) {
    const newEvents = status.events_log.filter(e => (e.id || 0) > lastSeenEventId);
    if (newEvents.length > 0) {
      newEvents.forEach(entry => {
        const row = document.createElement("div");
        row.className = `log-entry log-${escapeHtml(entry.level || 'info')}`;
        row.textContent = `[${entry.time || ''}] ${entry.message || ''}`;
        logContainer.appendChild(row);
        lastSeenEventId = Math.max(lastSeenEventId, entry.id || 0);
      });
      logContainer.scrollTop = logContainer.scrollHeight;
    }
  }
}

function renderBatchView(b) {
  const titleEl = document.getElementById("last-batch-title");
  const metaEl = document.getElementById("last-batch-meta");
  const openBtn = document.getElementById("btn-open-batch-folder");
  const copyBtn = document.getElementById("btn-copy-batch-folder");
  const filesSec = document.getElementById("generated-files-section");
  const filesList = document.getElementById("generated-files-list");
  const otBadge = document.getElementById("ot-replica-badge");

  const ofStr = escapeHtml(b.of || "Desconocida");
  const odoo = escapeHtml((b.order_info && b.order_info.pedido_odoo) || "-");
  const desc = escapeHtml((b.order_info && b.order_info.descripcion) || "");
  const totalP = b.total_pages || 0;
  const destDir = b.destination_dir || b.dest_dir || "";

  if (titleEl) titleEl.textContent = `OF: ${ofStr} | Odoo: ${odoo} (${totalP} págs)`;
  if (metaEl) metaEl.textContent = `${desc ? desc + ' — ' : ''}Carpeta: ${destDir}`;

  if (openBtn) openBtn.style.display = "inline-flex";
  if (copyBtn) copyBtn.style.display = "inline-flex";

  if (otBadge) {
    if (b.ot_replica_path) {
      otBadge.innerHTML = `<span class="badge badge-active">✓ Replicado en Oficina Técnica: ${escapeHtml(b.ot_replica_path)}</span>`;
      otBadge.style.display = "block";
    } else {
      otBadge.style.display = "none";
    }
  }

  if (filesSec) filesSec.style.display = "block";
  if (filesList && b.created_files) {
    filesList.innerHTML = "";
    b.created_files.forEach(f => {
      const item = document.createElement("div");
      item.className = "generated-file-item" + (f.is_warning ? " warning" : "");

      const fPath = f.path || "";
      const downloadUrl = `/api/files/download?path=${encodeURIComponent(fPath)}`;

      item.innerHTML = `
        <div class="file-info">
          <span class="file-name">${f.is_warning ? '⚠️ ' : '📄 '}${escapeHtml(f.doc_name)}</span>
          <span class="file-sub">${escapeHtml(f.filename)}</span>
        </div>
        <div style="display:flex; align-items:center; gap:6px;">
          <span class="file-pages-badge">Págs: [${(f.pages_list || []).join(', ')}]</span>
          <a href="${downloadUrl}" target="_blank" class="btn-sm btn-ghost" title="Abrir / Ver PDF en el navegador">👁️ Ver</a>
          <button class="btn-sm btn-ghost btn-copy-fpath" data-path="${escapeHtml(fPath)}" title="Copiar ruta">📋</button>
        </div>
      `;

      item.querySelector(".btn-copy-fpath").addEventListener("click", function(e) {
        e.stopPropagation();
        navigator.clipboard.writeText(fPath).then(() => {
          this.textContent = "✓";
          setTimeout(() => this.textContent = "📋", 1500);
          showToast("Ruta copiada al portapapeles", "info", 2000);
        });
      });

      filesList.appendChild(item);
    });
  }
}

function renderBatchesHistory(batches) {
  const container = document.getElementById("batches-history-container");
  const countEl = document.getElementById("history-count");
  if (!container) return;

  if (countEl) countEl.textContent = batches.length;
  container.innerHTML = "";

  if (batches.length === 0) {
    container.innerHTML = '<div style="color:var(--text-dim); padding:10px; text-align:center;">No hay lotes en esta sesión.</div>';
    return;
  }

  batches.slice().reverse().forEach(b => {
    const card = document.createElement("div");
    card.className = "history-card";
    const filesCount = b.files_count || (b.created_files ? b.created_files.length : 0);
    const otTag = b.ot_replica_path ? '<span style="color:#34d399; margin-left:6px;">[OT ✓]</span>' : '';

    card.innerHTML = `
      <div>
        <span style="font-family:Consolas, monospace; font-size:0.75rem; background:var(--bg-tertiary); padding:2px 5px; border-radius:3px;">
          ${escapeHtml(b.time || '')}
        </span>
        <strong style="color:#60a5fa; margin-left:6px;">${escapeHtml(b.of || '')}</strong>
        ${otTag}
        <span style="color:var(--text-muted); font-size:0.75rem; margin-left:6px;">${filesCount} PDFs (${b.total_pages || 0} págs)</span>
      </div>
      <div style="display:flex; gap:4px;">
        <button class="btn-sm btn-ghost btn-view-hist">👁️ Ver</button>
      </div>
    `;

    card.querySelector(".btn-view-hist").addEventListener("click", () => {
      renderBatchView(b);
      showToast(`Visualizando lote histórico de la OF ${b.of}`, "info", 2500);
    });

    container.appendChild(card);
  });
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

async function clearAllAlerts() {
  try {
    await fetch("/api/alerts/clear_all", { method: "POST" });
    pollSession();
  } catch (e) {
    console.error("Error clearing all alerts:", e);
  }
}

async function openLastBatchFolder() {
  try {
    const target = lastBatchData ? (lastBatchData.destination_dir || lastBatchData.dest_dir) : null;
    const res = await fetch("/api/session/open_folder", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ folder: target })
    });
    const d = await res.json();
    if (d.success) {
      showToast("Carpeta abierta en el equipo servidor.", "success");
    } else {
      showToast(d.error || "No se pudo abrir la carpeta", "warning");
    }
  } catch (e) {
    showToast("Error: " + e, "error");
  }
}

function copyBatchFolder() {
  if (!lastBatchData) return;
  const folder = lastBatchData.destination_dir || lastBatchData.dest_dir;
  if (!folder) return;

  navigator.clipboard.writeText(folder).then(() => {
    showToast(`Ruta copiada: ${folder}`, "success");
  }).catch(e => {
    showToast("Error al copiar al portapapeles: " + e, "error");
  });
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

    const otPill = document.getElementById("ot-status-pill");
    if (otPill) {
      const dot = otPill.querySelector(".status-dot");
      const lbl = otPill.querySelector(".status-label");
      if (dot) dot.className = `status-dot ${d.oficina_tecnica_ok ? 'dot-green' : 'dot-red'}`;
      if (lbl) lbl.textContent = `OT: ${d.oficina_tecnica_ok ? 'Conectada' : 'No accesible'}`;
    }

    const ocrPill = document.getElementById("ocr-status-pill");
    if (ocrPill) {
      const dot = ocrPill.querySelector(".status-dot");
      const lbl = ocrPill.querySelector(".status-label");
      const active = d.ocr_available && d.ocr_enabled;
      if (dot) dot.className = `status-dot ${active ? 'dot-green' : (d.ocr_available ? 'dot-yellow' : 'dot-red')}`;
      if (lbl) lbl.textContent = `OCR: ${active ? 'Activo' : (d.ocr_available ? 'Inactivo' : 'No disp.')}`;
    }

    const qPill = document.getElementById("queue-status-pill");
    if (qPill) {
      const dot = qPill.querySelector(".status-dot");
      const lbl = qPill.querySelector(".status-label");
      const count = d.pending_count || 0;
      if (dot) dot.className = `status-dot ${count > 0 ? 'dot-yellow' : 'dot-green'}`;
      if (lbl) lbl.textContent = `Bandeja: ${count} pend.`;
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
              document.getElementById("info-cliente").textContent = o.cliente || "-";
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

  document.addEventListener("click", (e) => {
    if (!ofInput.contains(e.target) && !autocompleteList.contains(e.target)) {
      autocompleteList.style.display = "none";
    }
  });
}

function setupSettingsModal() {
  const modal = document.getElementById("settings-modal");
  const btnOpen = document.getElementById("btn-open-settings");
  const btnClose = document.getElementById("btn-close-settings");
  const btnSave = document.getElementById("btn-save-settings");

  const closeModal = () => { modal.style.display = "none"; };

  if (btnOpen) {
    btnOpen.addEventListener("click", async () => {
      try {
        const res = await fetch("/api/config");
        const cfg = await res.json();
        document.getElementById("cfg-excel-path").value = cfg.excel_path || "";
        document.getElementById("cfg-escaneos-folder").value = cfg.escaneos_folder || "";
        document.getElementById("cfg-dest-folder").value = cfg.destination_base_folder || "";
        document.getElementById("cfg-ot-base").value = cfg.oficina_tecnica_base || "";
        document.getElementById("cfg-copy-ot").checked = (cfg.copy_to_oficina_tecnica !== false);
        document.getElementById("cfg-ocr-enabled").checked = (cfg.ocr_enabled !== false);
        document.getElementById("cfg-duplex-mode").checked = (cfg.duplex_mode !== false);
        document.getElementById("cfg-ocr-header-height").value = cfg.ocr_header_height_percent || 30;
        document.getElementById("cfg-ocr-height").value = cfg.ocr_footer_height_percent || 15;
        document.getElementById("cfg-catalog-path").value = cfg.ocr_catalog_path || "codigosOCR.xlsx";
        modal.style.display = "flex";
      } catch (e) {
        showToast("Error cargando configuración: " + e, "error");
      }
    });
  }

  if (btnClose) btnClose.addEventListener("click", closeModal);

  modal.addEventListener("click", (e) => {
    if (e.target === modal) closeModal();
  });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && modal.style.display === "flex") {
      closeModal();
    }
  });

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
        oficina_tecnica_base: document.getElementById("cfg-ot-base").value.trim(),
        copy_to_oficina_tecnica: document.getElementById("cfg-copy-ot").checked,
        ocr_enabled: document.getElementById("cfg-ocr-enabled").checked,
        duplex_mode: document.getElementById("cfg-duplex-mode").checked,
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
          showToast("Configuración guardada correctamente.", "success");
          closeModal();
          checkSystemInfo();
          loadCatalog();
        } else {
          showToast("Error guardando: " + d.error, "error");
        }
      } catch (e) {
        showToast("Error de guardado: " + e, "error");
      }
    });
  }
}
