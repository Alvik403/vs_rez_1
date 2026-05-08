const form = document.querySelector("#uploadForm");
const filesInput = document.querySelector("#files");
const fileList = document.querySelector("#fileList");
const statusText = document.querySelector("#status");
const submitButton = document.querySelector("#submitButton");

const formatBytes = (bytes) => {
  if (!bytes) return "0 Б";
  const units = ["Б", "КБ", "МБ", "ГБ"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** index).toFixed(index ? 1 : 0)} ${units[index]}`;
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

  statusText.className = "status";
  statusText.textContent = files.length
    ? `Выбрано файлов: ${files.length}. Нажмите "Собрать и скачать".`
    : "Файлы пока не выбраны.";
};

filesInput.addEventListener("change", updateFileList);

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
