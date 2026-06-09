const state = {
  bytes: null,
  workbook: null,
  fileName: "",
  source: "local",
  turmas: [],
  turma: null,
  aula: null,
  alunos: [],
  history: [],
  accessToken: ""
};

const $ = id => document.getElementById(id);
const views = ["home", "turmas", "aulas", "chamada"];

function show(view) {
  views.forEach(name => $(`view-${name}`).classList.toggle("active", name === view));
  window.scrollTo({ top: 0, behavior: "instant" });
}

function toast(message) {
  const el = $("toast");
  el.textContent = message;
  el.classList.add("show");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => el.classList.remove("show"), 2800);
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, char => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;"
  })[char]);
}

function dateLabel(value) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) {
    return `${String(value.getDate()).padStart(2, "0")}/${String(value.getMonth() + 1).padStart(2, "0")}/${value.getFullYear()}`;
  }
  if (typeof value === "number") {
    const parts = XLSX.SSF.parse_date_code(value);
    if (parts) return `${String(parts.d).padStart(2, "0")}/${String(parts.m).padStart(2, "0")}/${parts.y}`;
  }
  return String(value || "");
}

function normalizedStatus(value) {
  const v = String(value ?? "").trim().toUpperCase();
  return v === "F" || v === "I" ? v : "0";
}

function parseWorkbook(bytes, fileName, source = "local") {
  const workbook = XLSX.read(bytes, { type: "array", cellDates: true, cellFormula: true, raw: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
  const turmas = [];

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex] || [];
    const marker = String(row[1] || "").trim();
    const isRealHeader = String(row[0] || "").trim().toUpperCase() === "INT"
      && /^TURMA:/i.test(marker)
      && row.slice(4).some(value => value !== null && value !== "");
    if (!isRealHeader) continue;

    const nextHeader = rows.findIndex((candidate, index) =>
      index > rowIndex
      && String(candidate?.[0] || "").trim().toUpperCase() === "INT"
      && /^TURMA:/i.test(String(candidate?.[1] || "").trim())
    );
    const end = nextHeader < 0 ? rows.length : nextHeader;
    const students = [];
    for (let r = rowIndex + 1; r < end; r++) {
      const name = rows[r]?.[2];
      if (typeof name === "string" && name.trim()) {
        students.push({ rowIndex: r, name: name.trim(), row: rows[r] });
      }
    }

    const occurrences = new Map();
    const lessons = [];
    for (let colIndex = 4; colIndex < row.length; colIndex++) {
      const value = row[colIndex];
      if (value === null || value === "") continue;
      const label = dateLabel(value);
      const occurrence = (occurrences.get(label) || 0) + 1;
      occurrences.set(label, occurrence);
      lessons.push({ colIndex, label, occurrence });
    }
    lessons.forEach(lesson => {
      lesson.totalOccurrences = occurrences.get(lesson.label);
    });

    turmas.push({
      name: marker.replace(/^TURMA:\s*-?/i, "").trim(),
      headerRow: rowIndex,
      students,
      lessons
    });
  }

  if (!turmas.length) throw new Error("Nenhuma turma compatível com a Mestre foi encontrada.");
  state.bytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  state.workbook = workbook;
  state.fileName = fileName;
  state.source = source;
  state.turmas = turmas;
  $("arquivo-nome").textContent = fileName;
  renderTurmas();
  show("turmas");
}

function renderTurmas() {
  $("turma-grid").innerHTML = state.turmas.map((turma, index) => `
    <button class="turma-card" data-turma="${index}">
      <strong>${escapeHtml(turma.name)}</strong>
      <span>${turma.students.length} alunos · ${turma.lessons.length} aulas</span>
    </button>
  `).join("");
}

function completedKey(turma, lesson) {
  return `mestre:done:${state.fileName}:${turma.name}:${lesson.colIndex}`;
}

function draftKey(turma, lesson) {
  return `mestre:draft:${state.fileName}:${turma.name}:${lesson.colIndex}`;
}

function openTurma(index) {
  state.turma = state.turmas[index];
  $("turma-nome").textContent = state.turma.name;
  $("aula-list").innerHTML = state.turma.lessons.map((lesson, lessonIndex) => {
    const suffix = lesson.totalOccurrences > 1 ? ` · ${lesson.occurrence}º horário` : "";
    const done = localStorage.getItem(completedKey(state.turma, lesson)) === "1";
    return `
      <button class="aula-card ${done ? "done" : ""}" data-aula="${lessonIndex}">
        <div><div class="date">${escapeHtml(lesson.label)}</div><span>Aula${suffix}</span></div>
      </button>
    `;
  }).join("");
  show("aulas");
}

function openLesson(index) {
  state.aula = state.turma.lessons[index];
  const draft = JSON.parse(localStorage.getItem(draftKey(state.turma, state.aula)) || "null");
  state.alunos = state.turma.students.map(student => ({
    rowIndex: student.rowIndex,
    name: student.name,
    status: draft?.[student.rowIndex] || normalizedStatus(student.row[state.aula.colIndex])
  }));
  state.history = [];
  $("chamada-turma").textContent = state.turma.name;
  $("chamada-data").textContent = state.aula.totalOccurrences > 1
    ? `${state.aula.label} · ${state.aula.occurrence}º horário`
    : state.aula.label;
  $("busca").value = "";
  renderStudents();
  show("chamada");
}

function renderStudents(filter = "") {
  const query = filter.trim().toLocaleLowerCase("pt-BR");
  $("aluno-list").innerHTML = state.alunos.map((student, index) => {
    if (query && !student.name.toLocaleLowerCase("pt-BR").includes(query)) return "";
    const css = student.status === "F" ? "falta" : student.status === "I" ? "justificada" : "";
    return `
      <article class="aluno ${css}" data-student="${index}">
        <span class="numero">${index + 1}</span>
        <span class="nome">${escapeHtml(student.name)}</span>
        <div class="status-buttons">
          <button class="f ${student.status === "F" ? "active" : ""}" data-status="F" aria-label="Falta">F</button>
          <button class="i ${student.status === "I" ? "active" : ""}" data-status="I" aria-label="Justificada">I</button>
        </div>
      </article>
    `;
  }).join("");
  updateSummary();
}

function setStatus(index, requested, recordHistory = true) {
  const student = state.alunos[index];
  const previous = student.status;
  const next = previous === requested ? "0" : requested;
  if (previous === next) return;
  if (recordHistory) state.history.push({ index, previous });
  student.status = next;
  saveDraft();
  renderStudents($("busca").value);
  $("btn-desfazer").disabled = state.history.length === 0;
}

function saveDraft() {
  const values = Object.fromEntries(state.alunos.map(student => [student.rowIndex, student.status]));
  localStorage.setItem(draftKey(state.turma, state.aula), JSON.stringify(values));
  $("draft-state").textContent = `Rascunho salvo às ${new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })}`;
}

function updateSummary() {
  const faltas = state.alunos.filter(student => student.status === "F").length;
  const justificadas = state.alunos.filter(student => student.status === "I").length;
  $("count-faltas").textContent = faltas;
  $("count-justificadas").textContent = justificadas;
  $("count-presentes").textContent = state.alunos.length - faltas - justificadas;
}

function colName(index) {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return name;
}

function patchCellInRow(rowXml, address, status) {
  const cellRegex = new RegExp(`<c\\b([^>]*\\br="${address}"[^>]*)>[\\s\\S]*?<\\/c>`);
  const selfClosingRegex = new RegExp(`<c\\b([^>]*\\br="${address}"[^>]*)\\/>`);
  const match = rowXml.match(cellRegex) || rowXml.match(selfClosingRegex);
  const attributes = match ? match[1].replace(/\s+t="[^"]*"/g, "") : ` r="${address}" s="1"`;
  const cell = status === "F" || status === "I"
    ? `<c${attributes} t="inlineStr"><is><t>${status}</t></is></c>`
    : `<c${attributes}><v>0</v></c>`;
  if (match) return rowXml.replace(match[0], cell);
  return rowXml.replace("</row>", `${cell}</row>`);
}

async function buildPatchedWorkbook() {
  const zip = await JSZip.loadAsync(state.bytes);
  const sheetPath = "xl/worksheets/sheet1.xml";
  const file = zip.file(sheetPath);
  if (!file) throw new Error("A primeira aba da planilha não foi encontrada.");
  let xml = await file.async("string");

  state.alunos.forEach(student => {
    const excelRow = student.rowIndex + 1;
    const address = `${colName(state.aula.colIndex)}${excelRow}`;
    const rowRegex = new RegExp(`<row\\b([^>]*\\br="${excelRow}"[^>]*)>[\\s\\S]*?<\\/row>`);
    const match = xml.match(rowRegex);
    if (!match) throw new Error(`Linha ${excelRow} não encontrada no modelo.`);
    xml = xml.replace(match[0], patchCellInRow(match[0], address, student.status));
  });

  zip.file(sheetPath, xml);
  return zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 }
  });
}

function download(bytes) {
  const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = state.fileName || "frequencia.xlsx";
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1500);
}

function driveConfig() {
  return {
    clientId: localStorage.getItem("mestre:google-client-id") || "",
    fileId: localStorage.getItem("mestre:google-file-id") || ""
  };
}

function requestDriveToken() {
  const { clientId } = driveConfig();
  if (!clientId) return Promise.reject(new Error("Configure o OAuth Client ID."));
  if (!window.google?.accounts?.oauth2) return Promise.reject(new Error("O login Google ainda não carregou."));
  return new Promise((resolve, reject) => {
    const client = google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: "https://www.googleapis.com/auth/drive",
      callback: response => {
        if (response.error) reject(new Error(response.error));
        else {
          state.accessToken = response.access_token;
          resolve(response.access_token);
        }
      }
    });
    client.requestAccessToken({ prompt: state.accessToken ? "" : "consent" });
  });
}

async function openFromDrive() {
  const { fileId } = driveConfig();
  if (!fileId) throw new Error("Configure o ID do arquivo no Google Drive.");
  const token = state.accessToken || await requestDriveToken();
  const metaResponse = await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?fields=name,mimeType`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!metaResponse.ok) throw new Error("Não foi possível localizar o arquivo no Drive.");
  const meta = await metaResponse.json();
  const response = await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?alt=media`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!response.ok) throw new Error("Não foi possível baixar a planilha do Drive.");
  parseWorkbook(new Uint8Array(await response.arrayBuffer()), meta.name, "drive");
}

async function saveToDrive(bytes) {
  const { fileId } = driveConfig();
  const token = state.accessToken || await requestDriveToken();
  const response = await fetch(`https://www.googleapis.com/upload/drive/v3/files/${encodeURIComponent(fileId)}?uploadType=media`, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    },
    body: bytes
  });
  if (!response.ok) throw new Error("O Google Drive recusou a gravação.");
}

async function saveAttendance() {
  const button = $("btn-salvar");
  button.disabled = true;
  button.textContent = "Salvando...";
  try {
    const bytes = await buildPatchedWorkbook();
    if (state.source === "drive") {
      await saveToDrive(bytes);
      state.bytes = bytes;
      toast("Chamada salva no Google Drive.");
    } else {
      download(bytes);
      state.bytes = bytes;
      toast("Planilha salva sem alterar o modelo.");
    }
    localStorage.setItem(completedKey(state.turma, state.aula), "1");
    localStorage.removeItem(draftKey(state.turma, state.aula));
    state.alunos.forEach(student => {
      const sourceStudent = state.turma.students.find(item => item.rowIndex === student.rowIndex);
      if (sourceStudent) sourceStudent.row[state.aula.colIndex] = student.status;
    });
    $("draft-state").textContent = "Chamada salva";
  } catch (error) {
    toast(error.message);
  } finally {
    button.disabled = false;
    button.textContent = "Salvar chamada";
  }
}

document.addEventListener("click", event => {
  const turmaButton = event.target.closest("[data-turma]");
  if (turmaButton) openTurma(Number(turmaButton.dataset.turma));
  const aulaButton = event.target.closest("[data-aula]");
  if (aulaButton) openLesson(Number(aulaButton.dataset.aula));
  const statusButton = event.target.closest("[data-status]");
  if (statusButton) {
    const row = statusButton.closest("[data-student]");
    setStatus(Number(row.dataset.student), statusButton.dataset.status);
  }
  const back = event.target.closest("[data-go]");
  if (back) {
    if (back.dataset.go === "home") show("home");
    if (back.dataset.go === "turmas") show("turmas");
    if (back.dataset.go === "aulas") {
      renderStudents("");
      openTurma(state.turmas.indexOf(state.turma));
    }
  }
});

$("file-input").addEventListener("change", async event => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    parseWorkbook(new Uint8Array(await file.arrayBuffer()), file.name, "local");
  } catch (error) {
    toast(error.message);
  }
});

$("btn-drive").addEventListener("click", async () => {
  try {
    await openFromDrive();
  } catch (error) {
    toast(error.message);
    if (!driveConfig().clientId || !driveConfig().fileId) $("config-dialog").showModal();
  }
});

$("btn-config").addEventListener("click", () => {
  const config = driveConfig();
  $("google-client-id").value = config.clientId;
  $("google-file-id").value = config.fileId;
  $("config-dialog").showModal();
});

$("btn-save-config").addEventListener("click", event => {
  event.preventDefault();
  localStorage.setItem("mestre:google-client-id", $("google-client-id").value.trim());
  localStorage.setItem("mestre:google-file-id", $("google-file-id").value.trim());
  $("config-dialog").close();
  toast("Configuração salva neste aparelho.");
});

$("busca").addEventListener("input", event => renderStudents(event.target.value));
$("btn-salvar").addEventListener("click", saveAttendance);
$("btn-desfazer").addEventListener("click", () => {
  const action = state.history.pop();
  if (!action) return;
  state.alunos[action.index].status = action.previous;
  saveDraft();
  renderStudents($("busca").value);
  $("btn-desfazer").disabled = state.history.length === 0;
});

if ("serviceWorker" in navigator && location.protocol.startsWith("http")) {
  navigator.serviceWorker.register("sw.js").catch(() => {});
}
