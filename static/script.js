// Validador de Tareas — interfaz
const ACCEPTED = [".txt", ".docx", ".pdf", ".jpg", ".jpeg", ".png"];

const state = {
  rubric: null,
  solution: null,
  students: [],
  results: [],
  filter: "todas",
  query: "",
  sort: null, // { key: "apellido" | "nota", dir: 1 | -1 }
};

const $ = (sel) => document.querySelector(sel);
const form = $("#gradeForm");
const processBtn = $("#processBtn");
const actionHint = $("#actionHint");
const formMessage = $("#formMessage");
const studentList = $("#studentList");
const clearStudents = $("#clearStudents");
const progress = $("#progress");
const progressClock = $("#progressClock");
const progressTitle = $("#progressTitle");
const resultsSection = $("#results");
const resultsBody = $("#resultsBody");
const emptyFilter = $("#emptyFilter");
const banner = $("#banner");
const engine = $("#engineStatus");
const engineText = $("#engineText");

// ---------- Utilidades ----------

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function extOf(name) {
  const i = name.lastIndexOf(".");
  return i >= 0 ? name.slice(i).toLowerCase() : "";
}

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function formatGrade(n) {
  return n.toLocaleString("es-CO", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
}

function normalize(text) {
  return String(text ?? "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase();
}

function hasGrade(r) {
  return typeof r.nota === "number" && Number.isFinite(r.nota);
}

function needsReview(r) {
  return Array.isArray(r.revisar) && r.revisar.length > 0;
}

function showMessage(text) {
  formMessage.textContent = text;
  formMessage.hidden = !text;
}

function fileRow(file, removeLabel) {
  const ext = extOf(file.name).slice(1);
  return `
    <span class="file">
      <span class="file-icon">
        <svg><use href="#i-file" /></svg>
        <span class="file-ext" data-ext="${escapeHtml(ext)}">${escapeHtml(ext)}</span>
      </span>
      <span class="file-meta">
        <span class="file-name" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</span>
        <span class="file-size">${formatSize(file.size)}</span>
      </span>
      <button type="button" class="icon-btn" data-remove aria-label="${escapeHtml(removeLabel)}">
        <svg><use href="#i-x" /></svg>
      </button>
    </span>`;
}

// ---------- Estado de Ollama ----------

async function checkEngine() {
  engine.dataset.state = "checking";
  engineText.textContent = "Comprobando Ollama…";
  try {
    const res = await fetch("/api/estado");
    const data = await res.json();
    renderOcrNote(data);
    if (data.listo) {
      engine.dataset.state = "ok";
      engineText.innerHTML = `Ollama listo<span class="engine-extra"> con ${escapeHtml(data.modelo)}</span>`;
      engine.title = "Volver a comprobar";
    } else {
      engine.dataset.state = data.error && data.error.includes("no está instalado") ? "warn" : "error";
      engineText.textContent = engine.dataset.state === "warn" ? `Falta el modelo ${data.modelo}` : "Ollama no responde";
      engine.title = `${data.error} (clic para volver a comprobar)`;
    }
  } catch {
    engine.dataset.state = "error";
    engineText.textContent = "Servidor no disponible";
    engine.title = "No se pudo contactar con app.py. ¿Está corriendo?";
  }
}

engine.addEventListener("click", checkEngine);

function renderOcrNote(data) {
  const note = $("#ocrNote");
  if (!data || !data.listo) {
    note.hidden = true;
    return;
  }
  note.hidden = false;
  note.classList.toggle("is-warn", !data.ocr_listo);
  note.innerHTML = data.ocr_listo
    ? `Los PDF escaneados y las fotos, incluida la letra a mano, se leen con <strong>${escapeHtml(data.ocr_modelo)}</strong>. Tarda más que un documento con texto.`
    : `Para leer PDF escaneados y fotos falta el modelo de visión. Instálalo con <code>ollama pull ${escapeHtml(data.ocr_modelo)}</code> y vuelve a comprobar el estado arriba.`;
}

// ---------- Archivos ----------

function acceptFiles(fileList) {
  const ok = [];
  const rejected = [];
  for (const f of fileList) {
    (ACCEPTED.includes(extOf(f.name)) ? ok : rejected).push(f);
  }
  if (rejected.length) {
    const names = rejected.map((f) => f.name).slice(0, 3).join(", ");
    showMessage(
      `${rejected.length === 1 ? "Este archivo no se puede leer" : "Estos archivos no se pueden leer"}: ${names}${rejected.length > 3 ? "…" : ""}. Usa .docx, .pdf, .txt, .jpg o .png.`
    );
  } else {
    showMessage("");
  }
  return ok;
}

function setSlot(slot, files) {
  const accepted = acceptFiles(files);
  if (slot === "students") {
    const seen = new Set(state.students.map((f) => `${f.name}|${f.size}`));
    for (const f of accepted) {
      const key = `${f.name}|${f.size}`;
      if (!seen.has(key)) {
        state.students.push(f);
        seen.add(key);
      }
    }
  } else if (accepted.length) {
    state[slot] = accepted[0];
  }
  render();
}

document.querySelectorAll(".drop").forEach((drop) => {
  const slot = drop.dataset.slot;
  const input = drop.querySelector("input");

  input.addEventListener("change", () => {
    setSlot(slot, input.files);
    input.value = "";
  });

  ["dragenter", "dragover"].forEach((type) =>
    drop.addEventListener(type, (e) => {
      e.preventDefault();
      drop.classList.add("is-over");
    })
  );
  ["dragleave", "drop"].forEach((type) =>
    drop.addEventListener(type, (e) => {
      e.preventDefault();
      if (type === "dragleave" && drop.contains(e.relatedTarget)) return;
      drop.classList.remove("is-over");
    })
  );
  drop.addEventListener("drop", (e) => setSlot(slot, e.dataTransfer.files));

  // Quitar el archivo de rúbrica/solución sin abrir el selector
  drop.addEventListener("click", (e) => {
    if (e.target.closest("[data-remove]")) {
      e.preventDefault();
      state[slot] = null;
      render();
    }
  });
});

// Evita que soltar un archivo fuera de las zonas lo abra en el navegador
["dragover", "drop"].forEach((type) => window.addEventListener(type, (e) => e.preventDefault()));

studentList.addEventListener("click", (e) => {
  const btn = e.target.closest("[data-remove]");
  if (!btn) return;
  const index = Number(btn.closest("li").dataset.index);
  state.students.splice(index, 1);
  render();
});

clearStudents.addEventListener("click", () => {
  state.students = [];
  render();
});

function render() {
  for (const slot of ["rubric", "solution"]) {
    const drop = document.querySelector(`.drop[data-slot="${slot}"]`);
    const file = state[slot];
    drop.classList.toggle("has-file", Boolean(file));
    drop.querySelector(".drop-filled").innerHTML = file ? fileRow(file, `Quitar ${file.name}`) : "";
    drop.closest(".step").classList.toggle("is-done", Boolean(file));
  }

  studentList.innerHTML = state.students
    .map((f, i) => `<li data-index="${i}">${fileRow(f, `Quitar ${f.name}`)}</li>`)
    .join("");
  clearStudents.hidden = state.students.length < 2;
  $("#step3").closest(".step").classList.toggle("is-done", state.students.length > 0);

  const n = state.students.length;
  processBtn.textContent = n ? `Calificar ${n} ${n === 1 ? "tarea" : "tareas"}` : "Calificar tareas";
  const missing = [];
  if (!state.rubric) missing.push("la rúbrica");
  if (!n) missing.push("al menos una tarea");
  processBtn.disabled = missing.length > 0;
  actionHint.textContent = missing.length
    ? `Falta ${missing.join(" y ")}.`
    : "Puede tardar varios minutos; los escaneos y fotos tardan más.";
}

// ---------- Calificar ----------

const FASES = {
  preparando: "Preparando…",
  ocr: "Leyendo escaneos y fotos",
  calificando: "Calificando tareas",
  excel: "Generando el Excel",
  listo: "Listo",
};

let clockTimer = null;
let clockBase = 0; // segundos que ya llevaba el proceso según el servidor
let clockStart = 0;
let pollTimer = null;

function formatClock(s) {
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const ss = String(s % 60).padStart(2, "0");
  return h ? `${h}:${String(m).padStart(2, "0")}:${ss}` : `${m}:${ss}`;
}

function startProgress() {
  progressTitle.textContent = FASES.preparando;
  $("#progressDetail").textContent = "";
  setProgressFill(0);
  progressClock.textContent = "0:00";
  clockBase = 0;
  clockStart = Date.now();
  progress.hidden = false;
  progress.scrollIntoView({ behavior: "smooth", block: "center" });
  clearInterval(clockTimer);
  clockTimer = setInterval(() => {
    progressClock.textContent = formatClock(clockBase + Math.floor((Date.now() - clockStart) / 1000));
  }, 1000);
}

function setProgressFill(pct) {
  $("#progressFill").style.width = `${Math.max(2, Math.min(100, pct))}%`;
  progress.querySelector(".progress-bar").setAttribute("aria-valuenow", String(Math.round(pct)));
}

function updateProgress(job) {
  progressTitle.textContent = FASES[job.fase] || FASES.preparando;
  $("#progressDetail").textContent = job.fase === "preparando" ? "" : job.detalle;
  setProgressFill((job.paso / job.pasos_total) * 100);
  clockBase = job.segundos;
  clockStart = Date.now();
}

function stopProgress() {
  clearInterval(clockTimer);
  clearTimeout(pollTimer);
  progress.hidden = true;
}

function setBusy(busy) {
  form.querySelectorAll("input, button").forEach((el) => (el.disabled = busy));
  form.classList.toggle("form-busy", busy);
  if (!busy) render();
}

function showResults(payload) {
  state.results = payload.results;
  state.filter = "todas";
  state.query = "";
  state.sort = null;
  $("#searchInput").value = "";
  $("#downloadLink").href = payload.excel_url;
  banner.hidden = !payload.ollama_caido;
  banner.textContent = payload.ollama_caido
    ? "Se perdió la conexión con Ollama a mitad del grupo. Las tareas que quedaron sin calificar están marcadas; vuelve a enviarlas cuando Ollama esté disponible."
    : "";
  renderResults();
  resultsSection.hidden = false;
  resultsSection.scrollIntoView({ behavior: "smooth", block: "start" });
}

function finishJob() {
  stopProgress();
  setBusy(false);
}

async function pollJob(id, failures = 0) {
  try {
    const res = await fetch(`/api/trabajo/${id}`);
    const job = await res.json();
    if (!res.ok) {
      showMessage(job.error || "Se perdió el seguimiento del proceso.");
      finishJob();
      return;
    }
    updateProgress(job);
    if (job.estado === "listo") {
      finishJob();
      showResults(job.resultado);
      return;
    }
    if (job.estado === "error") {
      finishJob();
      showMessage(job.error);
      checkEngine();
      return;
    }
    pollTimer = setTimeout(() => pollJob(id), 1500);
  } catch {
    if (failures >= 5) {
      showMessage("No se pudo contactar con el servidor. Revisa que app.py siga corriendo.");
      finishJob();
      return;
    }
    pollTimer = setTimeout(() => pollJob(id, failures + 1), 3000);
  }
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  if (!state.rubric || !state.students.length) return;

  const data = new FormData();
  data.append("rubric", state.rubric);
  if (state.solution) data.append("solution", state.solution);
  state.students.forEach((f) => data.append("students", f));

  showMessage("");
  resultsSection.hidden = true;
  setBusy(true);
  startProgress();

  try {
    const res = await fetch("/api/process", { method: "POST", body: data });
    let payload;
    try {
      payload = await res.json();
    } catch {
      payload = { error: `El servidor respondió con un error (${res.status}).` };
    }
    if (!res.ok || payload.error) {
      showMessage(payload.error || "No se pudo calificar el grupo.");
      checkEngine();
      finishJob();
      return;
    }
    pollJob(payload.trabajo);
  } catch (err) {
    showMessage(`No se pudo contactar con el servidor: ${err.message}. Revisa que app.py siga corriendo.`);
    finishJob();
  }
});

// Si se recarga la página mientras un grupo se está calificando, retoma el avance
async function resumeRunningJob() {
  try {
    const res = await fetch("/api/trabajo");
    const { trabajo } = await res.json();
    if (trabajo && trabajo.estado === "en_curso") {
      setBusy(true);
      startProgress();
      updateProgress(trabajo);
      pollJob(trabajo.id);
    }
  } catch {
    /* sin servidor: el estado de Ollama ya lo indica */
  }
}

$("#newBatchBtn").addEventListener("click", () => {
  state.students = [];
  state.results = [];
  resultsSection.hidden = true;
  render();
  window.scrollTo({ top: 0, behavior: "smooth" });
});

// ---------- Resultados ----------

function renderOverview() {
  const all = state.results;
  const graded = all.filter(hasGrade);
  const review = all.filter(needsReview);
  const avg = graded.length ? graded.reduce((s, r) => s + r.nota, 0) / graded.length : null;

  $("#resultsSummary").textContent =
    `${all.length} ${all.length === 1 ? "tarea procesada" : "tareas procesadas"}, ordenadas por apellido.`;
  $("#figGraded").textContent = `${graded.length} de ${all.length}`;
  $("#figAverage").textContent = avg === null ? "—" : formatGrade(avg);
  const figReview = $("#figReview");
  figReview.textContent = String(review.length);
  figReview.classList.toggle("has-review", review.length > 0);

  // Distribución: una columna por nota entera (1 a 10)
  const bins = Array(10).fill(0);
  graded.forEach((r) => (bins[Math.min(9, Math.max(0, Math.floor(r.nota) - 1))] += 1));
  const max = Math.max(1, ...bins);
  $("#histBars").innerHTML = bins
    .map((count, i) => {
      const grade = i + 1;
      const cls = count === 0 ? "zero" : grade < 6 ? "low" : "";
      const h = count === 0 ? 2 : Math.max(6, Math.round((count / max) * 48));
      return `<div class="hist-col" title="${count} con nota ${grade}">
        <span class="hist-bar ${cls}" style="height:${h}px"></span>
        <span class="hist-label">${grade}</span>
      </div>`;
    })
    .join("");

  const counts = { todas: all.length, revisar: review.length, sinnota: all.length - graded.length };
  document.querySelectorAll(".filter").forEach((btn) => {
    btn.querySelector("span").textContent = counts[btn.dataset.filter];
    btn.setAttribute("aria-selected", String(btn.dataset.filter === state.filter));
  });
}

function visibleResults() {
  const q = normalize(state.query.trim());
  let rows = state.results.filter((r) => {
    if (state.filter === "revisar" && !needsReview(r)) return false;
    if (state.filter === "sinnota" && hasGrade(r)) return false;
    if (q && !normalize(`${r.nombre_completo} ${r.apellido} ${r.archivo}`).includes(q)) return false;
    return true;
  });
  if (state.sort) {
    const { key, dir } = state.sort;
    rows = [...rows].sort((a, b) => {
      if (key === "nota") {
        const av = hasGrade(a) ? a.nota : -Infinity;
        const bv = hasGrade(b) ? b.nota : -Infinity;
        return (av - bv) * dir;
      }
      return normalize(a.apellido).localeCompare(normalize(b.apellido), "es") * dir;
    });
  }
  return rows;
}

function renderRows() {
  const rows = visibleResults();
  resultsBody.innerHTML = rows
    .map((r, i) => {
      const unknown = !r.apellido || r.apellido === "Desconocido";
      const grade = hasGrade(r)
        ? `<span class="grade ${r.nota < 6 ? "low" : ""}">${formatGrade(r.nota)}</span>`
        : `<span class="grade none" title="Sin nota">—</span>`;
      const review = needsReview(r)
        ? `<ul class="review-list">${r.revisar.map((a) => `<li><span>${escapeHtml(a)}</span></li>`).join("")}</ul>`
        : "";
      const why = escapeHtml(r.justificacion || "");
      return `
        <tr class="${needsReview(r) ? "needs-review" : ""}">
          <td class="col-student ${unknown ? "student-unknown" : ""}">
            <span class="student-last">${unknown ? "Sin identificar" : escapeHtml(r.apellido)}</span>
            ${unknown ? "" : `<span class="student-full">${escapeHtml(r.nombre_completo)}</span>`}
            <span class="student-file">${escapeHtml(r.archivo)}</span>
          </td>
          <td class="col-grade">${grade}</td>
          <td class="col-why">
            <div class="why" id="why-${i}">
              <p class="why-text">${why || '<span class="review-ok">Sin justificación.</span>'}</p>
              <button type="button" class="why-toggle" aria-expanded="false" aria-controls="why-${i}" hidden>Ver completa</button>
            </div>
            ${r.transcripcion ? `<details class="ocr-text"><summary>Ver texto leído de la imagen</summary><pre>${escapeHtml(r.transcripcion)}</pre></details>` : ""}
          </td>
          <td class="col-review">${review}</td>
        </tr>`;
    })
    .join("");

  // Muestra "Ver completa" solo cuando el texto está recortado
  resultsBody.querySelectorAll(".why").forEach((box) => {
    const text = box.querySelector(".why-text");
    const toggle = box.querySelector(".why-toggle");
    toggle.hidden = text.scrollHeight <= text.clientHeight + 2;
  });

  emptyFilter.hidden = rows.length > 0;
  resultsBody.closest("table").hidden = rows.length === 0;

  document.querySelectorAll(".sort").forEach((btn) => {
    const active = state.sort && state.sort.key === btn.dataset.sort;
    if (active) btn.setAttribute("aria-sort", state.sort.dir === 1 ? "ascending" : "descending");
    else btn.removeAttribute("aria-sort");
  });
}

function renderResults() {
  renderOverview();
  renderRows();
}

resultsBody.addEventListener("click", (e) => {
  const toggle = e.target.closest(".why-toggle");
  if (!toggle) return;
  const box = toggle.closest(".why");
  const open = box.classList.toggle("is-open");
  toggle.setAttribute("aria-expanded", String(open));
  toggle.textContent = open ? "Ver menos" : "Ver completa";
});

document.querySelectorAll(".filter").forEach((btn) =>
  btn.addEventListener("click", () => {
    state.filter = btn.dataset.filter;
    renderResults();
  })
);

$("#searchInput").addEventListener("input", (e) => {
  state.query = e.target.value;
  renderRows();
});

document.querySelectorAll(".sort").forEach((btn) =>
  btn.addEventListener("click", () => {
    const key = btn.dataset.sort;
    if (!state.sort || state.sort.key !== key) {
      state.sort = { key, dir: key === "nota" ? -1 : 1 };
    } else if ((key === "nota" && state.sort.dir === -1) || (key !== "nota" && state.sort.dir === 1)) {
      state.sort.dir *= -1;
    } else {
      state.sort = null; // vuelve al orden original (por apellido)
    }
    renderRows();
  })
);

render();
checkEngine();
resumeRunningJob();
