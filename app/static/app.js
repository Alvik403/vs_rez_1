const form = document.querySelector("#uploadForm");
const dropzone = document.querySelector("#dropzone");
const filesInput = document.querySelector("#files");
const fileList = document.querySelector("#fileList");
const statusText = document.querySelector("#status");
const submitButton = document.querySelector("#submitButton");
let selectedFiles = [];

const formatBytes = (bytes) => {
  if (!bytes) return "0 Б";
  const units = ["Б", "КБ", "МБ", "ГБ"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** index).toFixed(index ? 1 : 0)} ${units[index]}`;
};

const isExcelFile = (file) => /\.(xlsx|xlsm)$/i.test(file.name);

const fileKey = (file) => `${file.name}:${file.size}:${file.lastModified}`;

const EXCEL_MIME_TYPES = new Set([
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel.sheet.macroEnabled.12",
  "application/vnd.ms-excel",
]);

const filesFromItems = (items) =>
  [...(items || [])]
    .filter((item) => item.kind === "file")
    .map((item) => item.getAsFile())
    .filter(Boolean);

const filesFromClipboardApi = async () => {
  if (!navigator.clipboard?.read) {
    return [];
  }

  try {
    const items = await navigator.clipboard.read();
    const files = [];
    for (const item of items) {
      for (const type of item.types) {
        if (!EXCEL_MIME_TYPES.has(type) && !type.includes("spreadsheet") && !type.includes("excel")) {
          continue;
        }
        const blob = await item.getType(type);
        const ext = type.includes("macroEnabled") ? "xlsm" : "xlsx";
        const name = blob.name || `clipboard-${Date.now()}.${ext}`;
        files.push(new File([blob], name, { type, lastModified: Date.now() }));
      }
    }
    return files;
  } catch {
    return [];
  }
};

const setStatus = (className, text) => {
  statusText.className = className;
  statusText.textContent = text;
};

const setFiles = (files) => {
  selectedFiles = [...files];
  const transfer = new DataTransfer();
  for (const file of selectedFiles) {
    transfer.items.add(file);
  }
  filesInput.files = transfer.files;
  updateFileList();
};

const addFiles = (incomingFiles, sourceLabel) => {
  const incoming = [...(incomingFiles || [])].filter(Boolean);
  if (!incoming.length) {
    if (sourceLabel) {
      setStatus(
        "status error",
        `${sourceLabel}: браузер не передал файлы. Попробуйте перетащить файл в область загрузки.`,
      );
    }
    return;
  }

  const nextFiles = [...selectedFiles];
  const existing = new Set(nextFiles.map(fileKey));
  const rejected = [];
  let added = 0;

  for (const file of incoming) {
    if (!isExcelFile(file)) {
      rejected.push(file.name);
      continue;
    }
    const key = fileKey(file);
    if (existing.has(key)) {
      continue;
    }
    nextFiles.push(file);
    existing.add(key);
    added += 1;
  }

  setFiles(nextFiles);

  if (rejected.length) {
    setStatus("status error", `Пропущены не Excel-файлы: ${rejected.join(", ")}`);
    return;
  }

  if (added > 0 && sourceLabel) {
    setStatus("status success", `${sourceLabel}: добавлено файлов ${added}.`);
  } else if (sourceLabel) {
    setStatus("status", `${sourceLabel}: новые файлы не найдены.`);
  }
};

const updateFileList = () => {
  const files = [...filesInput.files];
  fileList.innerHTML = "";

  for (const file of files) {
    const row = document.createElement("div");
    row.className = "file-row";
    row.innerHTML = `<span>${file.name}</span><span>${formatBytes(file.size)}</span>`;
    fileList.append(row);
  }

  setStatus(
    "status",
    files.length ? `Выбрано файлов: ${files.length}. Нажмите "Собрать и скачать".` : "Файлы пока не выбраны.",
  );
};

filesInput.addEventListener("change", () => addFiles(filesInput.files));

dropzone.addEventListener("click", () => filesInput.click());

dropzone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    filesInput.click();
  }
});

document.addEventListener("dragenter", (event) => {
  event.preventDefault();
  dropzone.classList.add("is-dragover");
});

document.addEventListener("dragover", (event) => {
  event.preventDefault();
  if (event.dataTransfer) {
    event.dataTransfer.dropEffect = "copy";
  }
  dropzone.classList.add("is-dragover");
});

document.addEventListener("dragleave", (event) => {
  if (event.target === document.documentElement || event.target === document.body) {
    dropzone.classList.remove("is-dragover");
  }
});

document.addEventListener("drop", (event) => {
  event.preventDefault();
  dropzone.classList.remove("is-dragover");
  const files = event.dataTransfer.files?.length ? [...event.dataTransfer.files] : filesFromItems(event.dataTransfer.items);
  addFiles(files, "Перетаскивание");
});

document.addEventListener("paste", async (event) => {
  const clipboard = event.clipboardData;
  let files = clipboard?.files?.length ? [...clipboard.files] : filesFromItems(clipboard?.items);
  if (!files.length) {
    files = await filesFromClipboardApi();
  }
  if (!files.length) {
    setStatus(
      "status error",
      "Ctrl+V не передал файлы. В браузерах это часто запрещено для файлов из проводника; перетащите файл в область загрузки.",
    );
    return;
  }
  event.preventDefault();
  addFiles(files, "Вставка из буфера");
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  statusText.className = "status";
  statusText.textContent = "Файл собирается на сервере...";
  submitButton.disabled = true;

  try {
    const response = await fetch("/process", {
      method: "POST",
      body: new FormData(form),
    });

    if (!response.ok) {
      let detail = "Не удалось собрать файл.";
      try {
        const payload = await response.json();
        detail = payload.detail || detail;
      } catch {
        detail = await response.text();
      }
      throw new Error(detail);
    }

    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition") || "";
    const filenameMatch = disposition.match(/filename="([^"]+)"/);
    const filename = filenameMatch ? filenameMatch[1] : "consolidated_reserves.xlsx";
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);

    const projectsCount = response.headers.get("X-Projects-Count") || "0";
    const openingsCount = response.headers.get("X-Openings-Count") || "0";
    statusText.className = "status success";
    statusText.textContent = `Готово: проектов ${projectsCount}, строк вскрытий ${openingsCount}.`;
  } catch (error) {
    statusText.className = "status error";
    statusText.textContent = error.message;
  } finally {
    submitButton.disabled = false;
  }
});
