const rubricFile = document.getElementById("rubricFile");
const solutionFile = document.getElementById("solutionFile");
const studentFiles = document.getElementById("studentFiles");
const fileCount = document.getElementById("fileCount");
const processBtn = document.getElementById("processBtn");
const statusEl = document.getElementById("status");
const resultsSection = document.getElementById("resultsSection");
const resultsBody = document.getElementById("resultsBody");
const downloadLink = document.getElementById("downloadLink");

studentFiles.addEventListener("change", () => {
  const n = studentFiles.files.length;
  fileCount.textContent = n ? `${n} archivo(s) seleccionado(s)` : "";
});

function setStatus(message, type = "info") {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
  statusEl.classList.remove("hidden");
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.innerText = str == null ? "" : String(str);
  return div.innerHTML;
}

function renderResults(results) {
  resultsBody.innerHTML = "";
  results.forEach((r) => {
    const tr = document.createElement("tr");
    const nota = r.nota === null || r.nota === undefined || r.nota === "" ? "—" : r.nota;
    tr.innerHTML = `
      <td>${escapeHtml(r.apellido)}</td>
      <td>${escapeHtml(r.nombre_completo)}</td>
      <td class="nota">${escapeHtml(nota)}</td>
      <td>${escapeHtml(r.justificacion)}</td>
      <td>${escapeHtml(r.archivo)}</td>
    `;
    resultsBody.appendChild(tr);
  });
  resultsSection.classList.remove("hidden");
}

processBtn.addEventListener("click", async () => {
  if (!rubricFile.files[0]) {
    alert("Debes subir el archivo de la rúbrica.");
    return;
  }
  if (!studentFiles.files.length) {
    alert("Debes subir al menos una tarea de estudiante.");
    return;
  }

  const formData = new FormData();
  formData.append("rubric", rubricFile.files[0]);
  if (solutionFile.files[0]) {
    formData.append("solution", solutionFile.files[0]);
  }
  for (const f of studentFiles.files) {
    formData.append("students", f);
  }

  processBtn.disabled = true;
  resultsSection.classList.add("hidden");
  setStatus(
    `Procesando ${studentFiles.files.length} tarea(s)... esto puede tardar varios minutos según el modelo y la cantidad de archivos. No cierres esta pestaña.`,
    "loading"
  );

  try {
    const res = await fetch("/api/process", {
      method: "POST",
      body: formData,
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      setStatus(`⚠️ ${data.error || "Ocurrió un error al procesar."}`, "error");
      return;
    }

    setStatus(`✅ Listo. Se calificaron ${data.results.length} tarea(s).`, "success");
    renderResults(data.results);
    downloadLink.href = data.excel_url;
  } catch (err) {
    setStatus(`⚠️ Error de conexión: ${err.message}`, "error");
  } finally {
    processBtn.disabled = false;
  }
});
