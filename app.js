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
  accessToken: "",
  tokenExpiresAt: 0,
  syncStatus: "local",
  lastSyncError: "",
  pendingPatches: []
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
  if (v === "1") return "F";
  if (v === "J") return "I";
  return v === "F" || v === "I" ? v : "0";
}

function openCacheDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open("chamada-mestre", 1);
    request.onupgradeneeded = () => request.result.createObjectStore("files");
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveCachedWorkbook(bytes, fileName, source, pendingPatches = []) {
  const db = await openCacheDb();
  await new Promise((resolve, reject) => {
    const transaction = db.transaction("files", "readwrite");
    transaction.objectStore("files").put({
      bytes: bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes),
      fileName,
      source,
      pendingSync: pendingPatches.length > 0,
      pendingPatches,
      savedAt: Date.now()
    }, "current");
    transaction.oncomplete = resolve;
    transaction.onerror = () => reject(transaction.error);
  });
  db.close();
}

async function loadCachedWorkbook() {
  const db = await openCacheDb();
  const cached = await new Promise((resolve, reject) => {
    const request = db.transaction("files").objectStore("files").get("current");
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
  db.close();
  return cached;
}

function readWorkbookModel(bytes) {
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

    const lessonGroups = new Map();
    for (let colIndex = 4; colIndex < row.length; colIndex++) {
      const value = row[colIndex];
      if (value === null || value === "") continue;
      const label = dateLabel(value);
      if (!lessonGroups.has(label)) lessonGroups.set(label, { label, colIndexes: [] });
      lessonGroups.get(label).colIndexes.push(colIndex);
    }
    const lessons = [...lessonGroups.values()];
    lessons.forEach(lesson => {
      lesson.totalOccurrences = lesson.colIndexes.length;
      lesson.completedInFile = lesson.colIndexes.every(colIndex =>
        students.some(student => {
          const value = student.row[colIndex];
          return value !== null && value !== undefined && String(value).trim() !== "";
        })
      );
    });

    turmas.push({
      name: marker.replace(/^TURMA:\s*-?/i, "").trim(),
      headerRow: rowIndex,
      students,
      lessons
    });
  }

  if (!turmas.length) throw new Error("Nenhuma turma compatível com a Mestre foi encontrada.");
  return { workbook, turmas };
}

function parseWorkbook(bytes, fileName, source = "local", options = {}) {
  const { workbook, turmas } = readWorkbookModel(bytes);
  state.bytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  state.workbook = workbook;
  state.fileName = fileName;
  state.source = source;
  state.turmas = turmas;
  state.pendingPatches = Array.isArray(options.pendingPatches) ? options.pendingPatches : [];
  state.syncStatus = options.fromCache
    ? (source === "drive" ? (state.pendingPatches.length ? "pending" : "cached") : "local")
    : (source === "drive" ? "synced" : "local");
  state.lastSyncError = options.unsafeLegacyPending
    ? "Uma gravação antiga foi bloqueada para não sobrescrever a planilha do Drive. Atualize a planilha antes de continuar."
    : "";
  if (options.unsafeLegacyPending) state.syncStatus = "error";
  $("arquivo-nome").textContent = "Prof. Maurício";
  if (!options.fromCache) saveCachedWorkbook(state.bytes, fileName, source, []).catch(() => {});
  updateSyncUi();
  renderTurmas();
  show("turmas");
}

function refreshWorkbookState(bytes) {
  const currentTurma = state.turma?.name;
  const currentDate = state.aula?.label;
  const { workbook, turmas } = readWorkbookModel(bytes);
  state.bytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  state.workbook = workbook;
  state.turmas = turmas;
  state.turma = turmas.find(item => normalizedName(item.name) === normalizedName(currentTurma)) || null;
  state.aula = state.turma?.lessons.find(item => item.label === currentDate) || null;
}

function updateSyncUi() {
  const card = $("sync-card");
  if (!card) return;
  card.classList.remove("synced", "pending", "error");
  const title = $("sync-title");
  const message = $("sync-message");
  const button = $("btn-sync-drive");

  if (state.source !== "drive") {
    title.textContent = "Arquivo local";
    message.textContent = "As alterações são baixadas como arquivo XLSX.";
    button.hidden = true;
    return;
  }

  button.hidden = false;
  if (state.syncStatus === "synced") {
    card.classList.add("synced");
    title.textContent = "Google Drive sincronizado";
    message.textContent = "A versão aberta corresponde ao arquivo do Drive.";
    button.textContent = "Atualizar";
  } else if (state.syncStatus === "pending") {
    card.classList.add("pending");
    title.textContent = "Salva no iPhone";
    const total = state.pendingPatches.length;
    message.textContent = `${total} chamada${total === 1 ? "" : "s"} aguardando envio seguro ao Google Drive.`;
    button.textContent = "Enviar agora";
  } else if (state.syncStatus === "error") {
    card.classList.add("error");
    title.textContent = "Não sincronizada";
    message.textContent = state.lastSyncError || "Não foi possível acessar o Google Drive.";
    button.textContent = "Tentar novamente";
  } else {
    card.classList.add("pending");
    title.textContent = "Cópia deste iPhone";
    message.textContent = "Toque em Atualizar para buscar a versão mais recente do Drive.";
    button.textContent = "Atualizar";
  }
}

function renderTurmas() {
  $("turma-grid").innerHTML = state.turmas.map((turma, index) => {
    const completed = turma.lessons.filter(isLessonCompleted.bind(null, turma)).length;
    const pending = turma.lessons.length - completed;
    const status = pending === 0
      ? `<span class="turma-status complete">✓ Todas realizadas</span>`
      : `<span class="turma-status">${pending} chamada${pending === 1 ? "" : "s"} pendente${pending === 1 ? "" : "s"}</span>`;
    return `
      <button class="turma-card ${pending === 0 ? "complete" : ""}" data-turma="${index}">
        <span class="turma-icon">${pending === 0 ? "✓" : pending}</span>
        <span class="turma-copy">
          <strong>${escapeHtml(turma.name)}</strong>
          <span>${turma.students.length} alunos · ${completed}/${turma.lessons.length} realizadas</span>
          ${status}
        </span>
        <span class="card-arrow">›</span>
      </button>
    `;
  }).join("");
}

function completedKey(turma, lesson) {
  return `mestre:done:${state.fileName}:${turma.name}:${lesson.label}`;
}

function draftKey(turma, lesson) {
  return `mestre:draft:${state.fileName}:${turma.name}:${lesson.label}`;
}

function isLessonCompleted(turma, lesson) {
  if (state.source === "drive") return lesson.completedInFile;
  return lesson.completedInFile || localStorage.getItem(completedKey(turma, lesson)) === "1";
}

function openTurma(index) {
  state.turma = state.turmas[index];
  $("turma-nome").textContent = state.turma.name;
  $("aula-list").innerHTML = state.turma.lessons.map((lesson, lessonIndex) => {
    const suffix = lesson.totalOccurrences > 1 ? ` · ${lesson.totalOccurrences} horários` : "";
    const done = isLessonCompleted(state.turma, lesson);
    return `
      <button class="aula-card ${done ? "done" : "pending"}" data-aula="${lessonIndex}">
        <span class="lesson-state">${done ? "✓" : lessonIndex + 1}</span>
        <span class="lesson-copy">
          <span class="date">${escapeHtml(lesson.label)}</span>
          <span>Aula${suffix}</span>
        </span>
        <span class="lesson-action">${done ? "Realizada" : "Fazer chamada"} <b>›</b></span>
      </button>
    `;
  }).join("");
  show("aulas");
}

function openLesson(index) {
  const lesson = state.turma.lessons[index];
  if (isLessonCompleted(state.turma, lesson)) {
    const reopen = window.confirm(
      `A chamada de ${lesson.label} já foi realizada.\n\nDeseja abrir mesmo assim para conferir ou corrigir?`
    );
    if (!reopen) return;
  }
  state.aula = lesson;
  const draft = JSON.parse(localStorage.getItem(draftKey(state.turma, state.aula)) || "null");
  const sourceColumn = state.aula.colIndexes.find(colIndex =>
    state.turma.students.some(student => String(student.row[colIndex] ?? "").trim() !== "")
  ) ?? state.aula.colIndexes[0];
  state.alunos = state.turma.students.map(student => {
    const originalStatus = normalizedStatus(student.row[sourceColumn]);
    return {
      rowIndex: student.rowIndex,
      name: student.name,
      originalStatus,
      status: draft?.[student.rowIndex] || originalStatus
    };
  });
  state.history = [];
  $("chamada-turma").textContent = state.turma.name;
  $("chamada-data").textContent = state.aula.totalOccurrences > 1
    ? `${state.aula.label} · ${state.aula.totalOccurrences} horários`
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

function normalizedName(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function createAttendancePatch() {
  return {
    turmaName: state.turma.name,
    dateLabel: state.aula.label,
    statuses: state.alunos.map(student => ({
      name: student.name,
      baseStatuses: state.aula.colIndexes.map(colIndex => {
        const sourceStudent = state.turma.students.find(item => item.rowIndex === student.rowIndex);
        return normalizedStatus(sourceStudent?.row[colIndex]);
      }),
      status: normalizedStatus(student.status)
    }))
  };
}

function upsertPendingPatch(patches, patch) {
  const key = `${normalizedName(patch.turmaName)}|${patch.dateLabel}`;
  const existing = patches.find(item => `${normalizedName(item.turmaName)}|${item.dateLabel}` === key);
  if (existing) {
    const baseByName = new Map(existing.statuses.map(item => [normalizedName(item.name), item.baseStatuses]));
    patch.statuses.forEach(item => {
      item.baseStatuses = baseByName.get(normalizedName(item.name)) ?? item.baseStatuses;
    });
  }
  return [
    ...patches.filter(item => `${normalizedName(item.turmaName)}|${item.dateLabel}` !== key),
    patch
  ];
}

function resolvePatchTarget(model, patch) {
  const turma = model.turmas.find(item => normalizedName(item.name) === normalizedName(patch.turmaName));
  if (!turma) throw new Error(`A turma ${patch.turmaName} não foi encontrada na versão atual do Drive.`);
  const lesson = turma.lessons.find(item => item.label === patch.dateLabel);
  if (!lesson) throw new Error(`A data ${patch.dateLabel} não foi encontrada na turma ${patch.turmaName}.`);

  const remoteStudents = new Map();
  turma.students.forEach(student => {
    const key = normalizedName(student.name);
    if (remoteStudents.has(key)) {
      throw new Error(`Existem alunos com nome repetido na turma ${patch.turmaName}. O envio foi bloqueado.`);
    }
    remoteStudents.set(key, student);
  });
  if (remoteStudents.size !== patch.statuses.length) {
    throw new Error(`A lista de alunos da turma ${patch.turmaName} mudou. Atualize a planilha antes de salvar.`);
  }

  const students = patch.statuses.map(item => {
    const student = remoteStudents.get(normalizedName(item.name));
    if (!student) {
      throw new Error(`O aluno ${item.name} não foi encontrado na versão atual do Drive.`);
    }
    if (!Array.isArray(item.baseStatuses) || item.baseStatuses.length !== lesson.colIndexes.length) {
      throw new Error("Uma alteração pendente antiga foi bloqueada. Atualize a planilha e refaça essa chamada.");
    }
    const currentStatuses = lesson.colIndexes.map(colIndex => normalizedStatus(student.row[colIndex]));
    const requestedStatus = normalizedStatus(item.status);
    const alreadyApplied = currentStatuses.every(status => status === requestedStatus);
    const unchangedSinceOpen = currentStatuses.every((status, index) =>
      status === normalizedStatus(item.baseStatuses[index])
    );
    if (!alreadyApplied && !unchangedSinceOpen) {
      throw new Error(`A chamada de ${patch.dateLabel} foi alterada no Drive. Atualize antes de corrigir novamente.`);
    }
    return { ...student, status: requestedStatus };
  });
  return { lesson, students };
}

async function buildPatchedWorkbook(baseBytes, patches) {
  const model = readWorkbookModel(baseBytes);
  const targets = patches.map(patch => resolvePatchTarget(model, patch));
  const zip = await JSZip.loadAsync(baseBytes);
  const sheetPath = "xl/worksheets/sheet1.xml";
  const file = zip.file(sheetPath);
  if (!file) throw new Error("A primeira aba da planilha não foi encontrada.");
  let xml = await file.async("string");

  targets.forEach(({ lesson, students }) => {
    students.forEach(student => {
      const excelRow = student.rowIndex + 1;
      const rowRegex = new RegExp(`<row\\b([^>]*\\br="${excelRow}"[^>]*)>[\\s\\S]*?<\\/row>`);
      const match = xml.match(rowRegex);
      if (!match) throw new Error(`Linha ${excelRow} não encontrada no modelo.`);
      let rowXml = match[0];
      lesson.colIndexes.forEach(colIndex => {
        const address = `${colName(colIndex)}${excelRow}`;
        rowXml = patchCellInRow(rowXml, address, student.status);
      });
      xml = xml.replace(match[0], rowXml);
    });
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
    apiKey: localStorage.getItem("mestre:google-api-key") || "",
    appId: localStorage.getItem("mestre:google-app-id") || "",
    fileId: localStorage.getItem("mestre:google-file-id") || "",
    fileName: localStorage.getItem("mestre:google-file-name") || ""
  };
}

function updateDriveUi() {
  const config = driveConfig();
  $("btn-drive").textContent = config.fileId
    ? "Atualizar planilha do Google Drive"
    : "Conectar ao Google Drive";
  const status = $("home-status");
  if (config.fileName) {
    status.querySelector("strong").textContent = "Planilha memorizada";
    status.querySelector("p").textContent = `${config.fileName} abre automaticamente neste aparelho.`;
  }
}

function requestDriveToken() {
  const { clientId } = driveConfig();
  if (!clientId) return Promise.reject(new Error("Configure o OAuth Client ID."));
  if (!window.google?.accounts?.oauth2) return Promise.reject(new Error("O login Google ainda não carregou."));
  if (state.accessToken && Date.now() < state.tokenExpiresAt - 60000) {
    return Promise.resolve(state.accessToken);
  }
  return new Promise((resolve, reject) => {
    let finished = false;
    const finish = (handler, value) => {
      if (finished) return;
      finished = true;
      clearTimeout(timer);
      handler(value);
    };
    const timer = setTimeout(() => {
      finish(reject, new Error("O Google não respondeu. Verifique se o popup foi bloqueado e tente novamente."));
    }, 20000);
    const client = google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: "https://www.googleapis.com/auth/drive.file",
      callback: response => {
        if (response.error) finish(reject, new Error(response.error_description || response.error));
        else {
          state.accessToken = response.access_token;
          state.tokenExpiresAt = Date.now() + (Number(response.expires_in) || 3600) * 1000;
          finish(resolve, response.access_token);
        }
      },
      error_callback: error => {
        const message = error?.type === "popup_failed_to_open"
          ? "O iPhone bloqueou a janela do Google. Permita popups para este site."
          : "A autorização do Google foi fechada ou não pôde ser concluída.";
        finish(reject, new Error(message));
      }
    });
    client.requestAccessToken({ prompt: "" });
  });
}

function waitForGoogleLibrary(name, test) {
  if (test()) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const started = Date.now();
    const timer = setInterval(() => {
      if (test()) {
        clearInterval(timer);
        resolve();
      } else if (Date.now() - started > 10000) {
        clearInterval(timer);
        reject(new Error(`${name} não carregou. Verifique sua internet e tente novamente.`));
      }
    }, 100);
  });
}

async function fetchWithTimeout(url, options = {}, timeout = 20000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeout);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error("O Google Drive demorou para responder. Verifique a internet e tente novamente.");
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

async function loadPickerApi() {
  await waitForGoogleLibrary("O seletor do Google", () => Boolean(window.gapi?.load));
  if (window.google?.picker) return;
  await new Promise((resolve, reject) => {
    gapi.load("picker", {
      callback: resolve,
      onerror: () => reject(new Error("Não foi possível carregar o seletor do Google.")),
      timeout: 10000,
      ontimeout: () => reject(new Error("O seletor do Google demorou para responder."))
    });
  });
}

async function pickDriveFile() {
  const { clientId, apiKey, appId } = driveConfig();
  if (!clientId || !apiKey || !appId) {
    throw new Error("Complete a configuração do Google Drive.");
  }
  await waitForGoogleLibrary("O login Google", () => Boolean(window.google?.accounts?.oauth2));
  const token = await requestDriveToken();
  await loadPickerApi();

  return new Promise((resolve, reject) => {
    const view = new google.picker.DocsView(google.picker.ViewId.DOCS)
      .setMimeTypes("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .setIncludeFolders(false)
      .setSelectFolderEnabled(false);
    const picker = new google.picker.PickerBuilder()
      .addView(view)
      .setOAuthToken(token)
      .setDeveloperKey(apiKey)
      .setAppId(appId)
      .setTitle("Escolha a planilha de frequência")
      .setCallback(data => {
        if (data.action === google.picker.Action.PICKED) {
          const document = data.docs?.[0];
          if (!document?.id) {
            reject(new Error("O Google não informou o arquivo selecionado."));
            return;
          }
          resolve({ id: document.id, name: document.name || "frequencia.xlsx" });
        } else if (data.action === google.picker.Action.CANCEL) {
          resolve(null);
        }
      })
      .build();
    picker.setVisible(true);
  });
}

async function getDriveMetadata(token, fileId = driveConfig().fileId) {
  if (!fileId) throw new Error("Escolha uma planilha no Google Drive.");
  const response = await fetchWithTimeout(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?fields=name,mimeType,version,modifiedTime,size`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!response.ok) throw new Error("Não foi possível localizar o arquivo no Drive.");
  return response.json();
}

async function downloadLatestDriveWorkbook(token, fileId = driveConfig().fileId) {
  const before = await getDriveMetadata(token, fileId);
  const response = await fetchWithTimeout(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?alt=media`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!response.ok) throw new Error("Não foi possível baixar a planilha do Drive.");
  const bytes = new Uint8Array(await response.arrayBuffer());
  const after = await getDriveMetadata(token, fileId);
  if (String(before.version) !== String(after.version)) {
    throw new Error("A planilha mudou enquanto era baixada. Tente novamente.");
  }
  return { bytes, meta: after };
}

async function openFromDrive(file) {
  const fileId = file?.id || driveConfig().fileId;
  if (!fileId) throw new Error("Escolha uma planilha no Google Drive.");
  const token = await requestDriveToken();
  const { bytes, meta } = await downloadLatestDriveWorkbook(token, fileId);
  localStorage.setItem("mestre:google-file-id", fileId);
  localStorage.setItem("mestre:google-file-name", meta.name);
  updateDriveUi();
  parseWorkbook(bytes, meta.name, "drive");
}

async function saveToDrive(bytes, authorizedToken = "", expectedVersion = "") {
  const { fileId } = driveConfig();
  if (!fileId) throw new Error("Nenhuma planilha do Drive está memorizada.");
  const token = authorizedToken || await requestDriveToken();
  if (expectedVersion) {
    const current = await getDriveMetadata(token, fileId);
    if (String(current.version) !== String(expectedVersion)) {
      throw new Error("A planilha foi alterada em outro lugar. O envio foi bloqueado; tente novamente.");
    }
  }
  const response = await fetchWithTimeout(`https://www.googleapis.com/upload/drive/v3/files/${encodeURIComponent(fileId)}?uploadType=media&fields=id,version,modifiedTime`, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    },
    body: bytes
  });
  if (!response.ok) {
    if (response.status === 401) {
      state.accessToken = "";
      state.tokenExpiresAt = 0;
      throw new Error("A sessão do Google expirou. Toque em Enviar agora para autorizar novamente.");
    }
    let detail = "";
    try {
      const body = await response.json();
      detail = body?.error?.message || "";
    } catch (_) {}
    throw new Error(detail || `O Google Drive recusou a gravação (${response.status}).`);
  }
  return response.json();
}

async function saveAttendance() {
  const button = $("btn-salvar");
  button.disabled = true;
  button.textContent = state.source === "drive" ? "Conectando..." : "Salvando...";
  try {
    // No Safari/iPhone, o OAuth precisa começar diretamente no toque.
    let driveToken = "";
    let driveAuthError = null;
    if (state.source === "drive") {
      try {
        driveToken = await requestDriveToken();
      } catch (error) {
        driveAuthError = error;
      }
    }
    button.textContent = "Salvando...";
    const patch = createAttendancePatch();
    const pendingPatches = state.source === "drive"
      ? upsertPendingPatch(state.pendingPatches, patch)
      : [];
    const localBytes = await buildPatchedWorkbook(state.bytes, [patch]);
    state.bytes = localBytes;
    state.pendingPatches = pendingPatches;
    await saveCachedWorkbook(localBytes, state.fileName, state.source, pendingPatches);
    localStorage.setItem(completedKey(state.turma, state.aula), "1");
    state.aula.completedInFile = true;
    localStorage.removeItem(draftKey(state.turma, state.aula));
    state.alunos.forEach(student => {
      const sourceStudent = state.turma.students.find(item => item.rowIndex === student.rowIndex);
      if (sourceStudent) {
        state.aula.colIndexes.forEach(colIndex => {
          sourceStudent.row[colIndex] = student.status;
        });
      }
    });

    if (state.source === "drive") {
      try {
        state.syncStatus = "pending";
        updateSyncUi();
        if (driveAuthError) throw driveAuthError;
        const latest = await downloadLatestDriveWorkbook(driveToken);
        const mergedBytes = await buildPatchedWorkbook(latest.bytes, pendingPatches);
        await saveToDrive(mergedBytes, driveToken, latest.meta.version);
        refreshWorkbookState(mergedBytes);
        state.pendingPatches = [];
        state.syncStatus = "synced";
        state.lastSyncError = "";
        await saveCachedWorkbook(mergedBytes, state.fileName, state.source, []);
        toast(`Enviada ao Drive em ${state.aula.totalOccurrences} horário${state.aula.totalOccurrences > 1 ? "s" : ""}.`);
      } catch (error) {
        state.syncStatus = "error";
        state.lastSyncError = error.message;
        toast("Salva no iPhone. O Drive não foi alterado.");
      }
    } else {
      download(localBytes);
      toast("Planilha salva sem alterar o modelo.");
    }
    updateSyncUi();
    $("draft-state").textContent = "Chamada salva";
    renderTurmas();
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
    const config = driveConfig();
    if (config.fileId) {
      await openFromDrive();
    } else {
      const file = await pickDriveFile();
      if (file) await openFromDrive(file);
    }
  } catch (error) {
    toast(error.message);
    const { clientId, apiKey, appId } = driveConfig();
    if (!clientId || !apiKey || !appId) $("config-dialog").showModal();
  }
});

$("btn-sync-drive").addEventListener("click", async () => {
  const button = $("btn-sync-drive");
  button.disabled = true;
  button.textContent = "Aguarde...";
  try {
    if (state.pendingPatches.length) {
      const token = await requestDriveToken();
      const latest = await downloadLatestDriveWorkbook(token);
      const mergedBytes = await buildPatchedWorkbook(latest.bytes, state.pendingPatches);
      await saveToDrive(mergedBytes, token, latest.meta.version);
      refreshWorkbookState(mergedBytes);
      state.pendingPatches = [];
      state.syncStatus = "synced";
      state.lastSyncError = "";
      await saveCachedWorkbook(mergedBytes, state.fileName, state.source, []);
      toast("Alterações enviadas ao Google Drive.");
    } else {
      await openFromDrive();
      toast("Planilha atualizada pelo Google Drive.");
    }
  } catch (error) {
    state.syncStatus = "error";
    state.lastSyncError = error.message;
    toast(error.message);
  } finally {
    button.disabled = false;
    updateSyncUi();
  }
});

$("btn-change-drive-file").addEventListener("click", async () => {
  try {
    const file = await pickDriveFile();
    if (!file) return;
    await openFromDrive(file);
    $("config-dialog").close();
  } catch (error) {
    toast(error.message);
  }
});

$("btn-config").addEventListener("click", () => {
  const config = driveConfig();
  $("google-client-id").value = config.clientId;
  $("google-api-key").value = config.apiKey;
  $("google-app-id").value = config.appId;
  $("selected-drive-file").hidden = !config.fileName;
  $("selected-drive-file").textContent = config.fileName ? `Última planilha: ${config.fileName}` : "";
  $("config-dialog").showModal();
});

$("btn-save-config").addEventListener("click", event => {
  event.preventDefault();
  const clientId = $("google-client-id").value.trim();
  const apiKey = $("google-api-key").value.trim();
  const appId = $("google-app-id").value.trim();
  if (!clientId || !apiKey || !appId) {
    toast("Preencha os três dados do Google Cloud.");
    return;
  }
  localStorage.setItem("mestre:google-client-id", clientId);
  localStorage.setItem("mestre:google-api-key", apiKey);
  localStorage.setItem("mestre:google-app-id", appId);
  state.accessToken = "";
  updateDriveUi();
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

updateDriveUi();

loadCachedWorkbook()
  .then(cached => {
    if (!cached?.bytes || state.bytes) return;
    const pendingPatches = Array.isArray(cached.pendingPatches) ? cached.pendingPatches : [];
    parseWorkbook(new Uint8Array(cached.bytes), cached.fileName, cached.source || "local", {
      fromCache: true,
      pendingPatches,
      unsafeLegacyPending: Boolean(cached.pendingSync) && pendingPatches.length === 0
    });
    toast("Planilha aberta do armazenamento seguro deste aparelho.");
  })
  .catch(() => {});

if ((location.hostname === "localhost" || location.hostname === "127.0.0.1") && location.search === "?test-model") {
  fetch("MODELO -FREQUENCIA.xlsx")
    .then(response => response.arrayBuffer())
    .then(bytes => parseWorkbook(new Uint8Array(bytes), "MODELO -FREQUENCIA.xlsx", "local"))
    .catch(error => toast(error.message));
}
