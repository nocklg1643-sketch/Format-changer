const { PDFDocument } = PDFLib;
const { Document, Packer, Paragraph, TextRun } = docx;

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const els = {
  fileInput: document.getElementById("fileInput"),
  dropzone: document.getElementById("dropzone"),
  sourceBadge: document.getElementById("sourceBadge"),
  targetBadge: document.getElementById("targetBadge"),
  fileName: document.getElementById("fileName"),
  fileSize: document.getElementById("fileSize"),
  detectedFormat: document.getElementById("detectedFormat"),
  previewContent: document.getElementById("previewContent"),
  targetSelect: document.getElementById("targetSelect"),
  targetChips: document.getElementById("targetChips"),
  conversionNotes: document.getElementById("conversionNotes"),
  convertBtn: document.getElementById("convertBtn"),
  clearBtn: document.getElementById("clearBtn"),
  resultTitle: document.getElementById("resultTitle"),
  resultSummary: document.getElementById("resultSummary"),
  downloadBtn: document.getElementById("downloadBtn"),
  shareBtn: document.getElementById("shareBtn"),
  shareHint: document.getElementById("shareHint"),
  supportList: document.getElementById("supportList"),
  officeGrid: document.getElementById("officeGrid"),
};

const LOCAL_FORMATS = {
  jpg: {
    label: "JPG / JPEG",
    category: "圖片",
    extensions: ["jpg", "jpeg"],
    mimeTypes: ["image/jpeg"],
    targets: ["png", "webp", "pdf"],
    notes: ["適合轉成 PNG、WEBP 或整理成 PDF。"],
  },
  png: {
    label: "PNG",
    category: "圖片",
    extensions: ["png"],
    mimeTypes: ["image/png"],
    targets: ["jpg", "webp", "pdf"],
    notes: ["轉 JPG 時透明背景會變成白底。"],
  },
  webp: {
    label: "WEBP",
    category: "圖片",
    extensions: ["webp"],
    mimeTypes: ["image/webp"],
    targets: ["jpg", "png", "pdf"],
    notes: ["適合壓縮型圖片工作流。"],
  },
  pdf: {
    label: "PDF",
    category: "文件",
    extensions: ["pdf"],
    mimeTypes: ["application/pdf"],
    targets: ["txt", "docx"],
    notes: [
      "PDF 轉 TXT / DOCX 走的是輕量文字擷取路線。",
      "若你需要完整版面保留，請改用下方高保真入口。",
    ],
  },
  docx: {
    label: "DOCX",
    category: "文件",
    extensions: ["docx"],
    mimeTypes: [
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ],
    targets: ["txt", "html", "pdf"],
    notes: [
      "DOCX 目前可做輕量轉 TXT、HTML、PDF。",
      "若你需要完整保留版面、圖片與字型，請改用下方高保真入口。",
    ],
  },
  pptx: {
    label: "PPTX",
    category: "簡報",
    extensions: ["pptx"],
    mimeTypes: [
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ],
    targets: ["txt", "html", "pdf"],
    notes: [
      "PPTX 目前以投影片中的可擷取文字為主。",
      "若你需要完整保留投影片版面、圖片與字型，請改用下方高保真入口。",
    ],
  },
  txt: {
    label: "TXT",
    category: "文字",
    extensions: ["txt"],
    mimeTypes: ["text/plain"],
    targets: ["md", "html", "pdf"],
    notes: ["純文字會在瀏覽器內直接完成轉換。"],
  },
  md: {
    label: "MD",
    category: "文字",
    extensions: ["md", "markdown"],
    mimeTypes: ["text/markdown"],
    targets: ["txt", "html", "pdf"],
    notes: ["Markdown 會以文字內容為主進行輸出。"],
  },
  html: {
    label: "HTML",
    category: "文字",
    extensions: ["html", "htm"],
    mimeTypes: ["text/html"],
    targets: ["txt", "md", "pdf"],
    notes: ["HTML 會抽出主要文字內容再輸出。"],
  },
  json: {
    label: "JSON",
    category: "資料",
    extensions: ["json"],
    mimeTypes: ["application/json", "text/json"],
    targets: ["txt", "md", "html", "pdf"],
    notes: ["JSON 會先整理格式再輸出。"],
  },
  xml: {
    label: "XML",
    category: "資料",
    extensions: ["xml"],
    mimeTypes: ["application/xml", "text/xml"],
    targets: ["txt", "html", "pdf"],
    notes: ["XML 會以可讀的文字形式輸出。"],
  },
  csv: {
    label: "CSV",
    category: "表格",
    extensions: ["csv"],
    mimeTypes: ["text/csv", "application/csv"],
    targets: ["xlsx", "json", "html", "pdf"],
    notes: ["CSV 可匯出成 XLSX、JSON、HTML 或 PDF 摘要。"],
  },
  xlsx: {
    label: "XLSX",
    category: "表格",
    extensions: ["xlsx"],
    mimeTypes: [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ],
    targets: ["csv", "json", "html", "pdf"],
    notes: ["XLSX 預設使用第一個工作表作為轉換來源。"],
  },
};

const HIGH_FIDELITY_ENTRIES = [
  {
    id: "doc-pdf",
    title: "DOC -> PDF",
    description: "適合舊版 Word 檔，建議交給專業文件引擎處理，避免版面跑掉。",
    label: "使用專業工具",
    url: "https://www.ilovepdf.com/",
  },
  {
    id: "docx-pdf",
    title: "DOCX -> PDF",
    description: "需要完整保留字型、段落、圖片與頁面配置時，建議使用高保真服務。",
    label: "使用專業工具",
    url: "https://www.ilovepdf.com/",
  },
  {
    id: "ppt-pdf",
    title: "PPT -> PDF",
    description: "適合舊版 PowerPoint 簡報，建議使用專業文件引擎保持投影片樣式。",
    label: "使用專業工具",
    url: "https://www.ilovepdf.com/",
  },
  {
    id: "pptx-pdf",
    title: "PPTX -> PDF",
    description: "如果你希望完整保留圖片、字型、位置與投影片版面，請先走高保真路線。",
    label: "使用專業工具",
    url: "https://www.ilovepdf.com/",
  },
];

const supportSummary = [
  ["文字", "TXT / MD / HTML / JSON / XML 可在瀏覽器內直接整理與轉換。"],
  ["表格", "CSV / XLSX 可輸出成 CSV、JSON、HTML、PDF 摘要。"],
  ["圖片", "JPG / PNG / WEBP 可互轉，並可整理成 PDF。"],
  ["PDF / 文件", "PDF 可做輕量文字擷取；DOCX / PPTX 可做輕量文字型轉換。"],
];

let currentFile = null;
let currentFormat = null;
let currentTarget = null;
let activeDownloadUrl = null;
let latestOutput = null;

init();

function init() {
  renderSupportList();
  renderOfficeEntries();
  bindEvents();
}

function bindEvents() {
  els.dropzone.addEventListener("click", () => {
    els.fileInput.click();
  });

  els.dropzone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      els.fileInput.click();
    }
  });

  els.fileInput.addEventListener("change", (event) => {
    const [file] = event.target.files || [];
    if (file) handleFile(file);
  });

  els.dropzone.addEventListener("dragover", (event) => {
    event.preventDefault();
    els.dropzone.classList.add("dragover");
  });

  els.dropzone.addEventListener("dragleave", () => {
    els.dropzone.classList.remove("dragover");
  });

  els.dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    els.dropzone.classList.remove("dragover");
    const [file] = event.dataTransfer.files || [];
    if (file) handleFile(file);
  });

  els.targetSelect.addEventListener("change", () => {
    currentTarget = els.targetSelect.value;
    updateTargetBadge();
    renderTargetChips();
    renderNotes();
    els.convertBtn.disabled = !currentTarget;
  });

  els.convertBtn.addEventListener("click", runConversion);
  els.clearBtn.addEventListener("click", resetApp);
  els.shareBtn.addEventListener("click", handleShareAction);
}

async function handleFile(file) {
  currentFile = file;
  currentFormat = detectLocalFormat(file);
  currentTarget = null;
  clearDownloadUrl();
  resetResult();

  els.fileName.textContent = file.name;
  els.fileSize.textContent = formatFileSize(file.size);
  els.detectedFormat.textContent = currentFormat
    ? `${LOCAL_FORMATS[currentFormat].label} (${LOCAL_FORMATS[currentFormat].category})`
    : "目前不在本地輕量轉換範圍";
  els.sourceBadge.textContent = currentFormat
    ? LOCAL_FORMATS[currentFormat].label
    : "請改用下方高保真入口";

  if (!currentFormat) {
    els.targetSelect.innerHTML = "<option>此檔案請改用高保真入口</option>";
    els.targetSelect.disabled = true;
    els.targetChips.innerHTML = '<span class="chip muted">這類型不在本地輕量轉換中</span>';
    els.conversionNotes.innerHTML =
      "<p>這個檔案類型不建議在目前這版瀏覽器工具裡硬做轉換。若你需要 DOC / DOCX / PPT / PPTX 轉 PDF，請使用下方高保真 Office 入口。</p>";
    els.previewContent.className = "preview-content empty";
    els.previewContent.textContent =
      "這個檔案目前不會在本地輕量轉換區塊處理。請改用下方的高保真 Office 轉 PDF 入口。";
    els.convertBtn.disabled = true;
    updateTargetBadge();
    return;
  }

  populateTargets(currentFormat);
  try {
    renderPreview(await buildPreview(file, currentFormat));
  } catch (error) {
    els.previewContent.className = "preview-content empty";
    els.previewContent.textContent = error.message || "預覽建立失敗。";
  }
  renderNotes();
}

function detectLocalFormat(file) {
  const name = file.name.toLowerCase();
  const extension = name.includes(".") ? name.split(".").pop() : "";
  const mime = file.type;

  return (
    Object.entries(LOCAL_FORMATS).find(([, descriptor]) => {
      return (
        descriptor.extensions.includes(extension) ||
        descriptor.mimeTypes.includes(mime)
      );
    })?.[0] || null
  );
}

function populateTargets(formatKey) {
  const targets = LOCAL_FORMATS[formatKey].targets;
  els.targetSelect.disabled = false;
  els.targetSelect.innerHTML = [
    '<option value="">選擇目標格式</option>',
    ...targets.map((target) => `<option value="${target}">${LOCAL_FORMATS[target].label}</option>`),
  ].join("");
  renderTargetChips();
  updateTargetBadge();
  els.convertBtn.disabled = true;
}

function renderTargetChips() {
  if (!currentFormat) return;
  els.targetChips.innerHTML = LOCAL_FORMATS[currentFormat].targets
    .map((target) => {
      const classes = ["chip"];
      if (currentTarget === target) classes.push("active");
      if (target === "pdf" || target === "txt") classes.push("lossy");
      return `<span class="${classes.join(" ")}">${LOCAL_FORMATS[target].label}</span>`;
    })
    .join("");
}

function renderNotes() {
  if (!currentFormat) return;

  const notes = [...LOCAL_FORMATS[currentFormat].notes];
  if (currentTarget) {
    notes.unshift(
      `即將從 ${LOCAL_FORMATS[currentFormat].label} 轉成 ${LOCAL_FORMATS[currentTarget].label}。`
    );
  } else {
    notes.unshift("請先從右側選單選擇目標格式。");
  }

  els.conversionNotes.innerHTML = `<ul>${notes.map((note) => `<li>${note}</li>`).join("")}</ul>`;
}

function updateTargetBadge() {
  els.targetBadge.textContent = currentTarget
    ? LOCAL_FORMATS[currentTarget].label
    : "請先選擇";
}

function renderSupportList() {
  els.supportList.innerHTML = supportSummary
    .map(([title, summary]) => `<li><strong>${title}</strong><br /><span>${summary}</span></li>`)
    .join("");
}

function renderOfficeEntries() {
  els.officeGrid.innerHTML = HIGH_FIDELITY_ENTRIES.map((entry) => {
    return `
      <article class="office-card">
        <div class="office-card-head">
          <h3>${entry.title}</h3>
          <span class="type-badge muted">高保真</span>
        </div>
        <p>${entry.description}</p>
        <small>目前先導向外部專業服務，未來可在這裡改接自己的 API / 後端。</small>
        <a class="office-link" href="${entry.url}" target="_blank" rel="noreferrer noopener">${entry.label}</a>
      </article>
    `;
  }).join("");
}

async function buildPreview(file, formatKey) {
  switch (formatKey) {
    case "jpg":
    case "png":
    case "webp":
      return { type: "image", url: URL.createObjectURL(file) };
    case "txt":
    case "md":
    case "json":
    case "xml":
    case "html":
      return {
        type: "pagedText",
        ...buildPagedTextPreview(await normalizeTextFile(file, formatKey), {
          title: file.name,
          pageLabel: "頁",
        }),
      };
    case "csv":
      return {
        type: "pagedText",
        ...buildPagedTextPreview(await file.text(), {
          title: file.name,
          pageLabel: "頁",
        }),
      };
    case "xlsx": {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const csvText = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
      return {
        type: "pagedText",
        ...buildPagedTextPreview(csvText, {
          title: `${file.name} - ${sheetName}`,
          pageLabel: "頁",
        }),
      };
    }
    case "docx": {
      const result = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
      return {
        type: "pagedText",
        ...buildPagedTextPreview(result.value, {
          title: file.name,
          pageLabel: "頁",
        }),
      };
    }
    case "pptx": {
      const text = await extractPptxText(file);
      return {
        type: "pagedText",
        ...buildPagedTextPreview(text, {
          title: file.name,
          pageLabel: "投影片",
        }),
      };
    }
    case "pdf":
      return await buildPdfPreview(file);
    default:
      return { type: "text", text: "暫無預覽。" };
  }
}

function renderPreview(preview) {
  els.previewContent.className = "preview-content";

  if (!preview) {
    els.previewContent.classList.add("empty");
    els.previewContent.textContent = "暫無預覽";
    return;
  }

  if (preview.type === "image") {
    els.previewContent.innerHTML = `<img src="${preview.url}" alt="預覽圖片" />`;
    return;
  }

  if (preview.type === "pdf") {
    renderPaperPreview(preview);
    return;
  }

  if (preview.type === "pagedText") {
    renderPagedTextPreview(preview);
    return;
  }

  els.previewContent.innerHTML = `<pre>${escapeHtml(preview.text || "")}</pre>`;
}

async function buildPdfPreview(file) {
  const pdf = await openPdfDocument(await file.arrayBuffer());
  const firstPageText = await extractPageText(pdf, 1);
  const pages = Array.from({ length: Math.min(pdf.numPages, 6) }, (_, index) => index + 1);
  return {
    type: "pdf",
    pdf,
    pages,
    totalPages: pdf.numPages,
    summaryText: truncate(firstPageText || "這份 PDF 沒有可擷取的文字摘要。", 220),
  };
}

async function openPdfDocument(arrayBuffer) {
  return await pdfjsLib.getDocument({
    data: arrayBuffer,
    disableWorker: true,
    useSystemFonts: true,
    isEvalSupported: false,
  }).promise;
}

async function extractPageText(pdf, pageNumber) {
  const page = await pdf.getPage(pageNumber);
  const textContent = await page.getTextContent();
  return textContent.items
    .map((item) => item.str)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

function renderPaperPreview(preview) {
  els.previewContent.className = "preview-content paper-scroll";
  els.previewContent.innerHTML = `
    <div class="preview-meta">
      <span>PDF 頁數：${preview.totalPages}</span>
      <span>預覽頁數：${preview.pages.length}</span>
      <span>摘要：${escapeHtml(preview.summaryText)}</span>
    </div>
    <div class="page-stack"></div>
  `;

  const stack = els.previewContent.querySelector(".page-stack");
  renderPdfPages(stack, preview);
}

async function renderPdfPages(stack, preview) {
  for (const pageNumber of preview.pages) {
    const pageCard = document.createElement("div");
    pageCard.className = "page-card";
    pageCard.innerHTML = `<strong>第 ${pageNumber} 頁</strong>`;

    const page = await preview.pdf.getPage(pageNumber);
    const unscaled = page.getViewport({ scale: 1 });
    const availableWidth = Math.max(360, els.previewContent.clientWidth - 72);
    const scale = Math.min(1.35, availableWidth / unscaled.width);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: context, viewport }).promise;

    pageCard.appendChild(canvas);
    stack.appendChild(pageCard);
  }

  if (preview.totalPages > preview.pages.length) {
    const notice = document.createElement("div");
    notice.className = "page-card";
    notice.innerHTML = `<strong>還有 ${preview.totalPages - preview.pages.length} 頁未顯示</strong>`;
    stack.appendChild(notice);
  }
}

function renderPagedTextPreview(preview) {
  els.previewContent.className = "preview-content paper-scroll";
  els.previewContent.innerHTML = `
    <div class="preview-meta">
      <span>${escapeHtml(preview.title)}</span>
      <span>${escapeHtml(preview.pageLabel)}數：${preview.totalPages}</span>
      <span>摘要：${escapeHtml(preview.summaryText)}</span>
    </div>
    <div class="page-stack"></div>
  `;

  const stack = els.previewContent.querySelector(".page-stack");
  preview.pages.forEach((pageLines, index) => {
    const pageCard = document.createElement("div");
    pageCard.className = "page-card";
    pageCard.innerHTML = `
      <strong>${escapeHtml(preview.pageLabel)} ${index + 1}</strong>
      <div class="text-sheet"><pre>${escapeHtml(pageLines.join("\n"))}</pre></div>
    `;
    stack.appendChild(pageCard);
  });

  if (preview.totalPages > preview.pages.length) {
    const notice = document.createElement("div");
    notice.className = "page-card";
    notice.innerHTML = `<strong>還有 ${preview.totalPages - preview.pages.length} 頁未顯示</strong>`;
    stack.appendChild(notice);
  }
}

async function runConversion() {
  if (!currentFile || !currentFormat || !currentTarget) return;

  els.convertBtn.disabled = true;
  els.convertBtn.textContent = "轉換中...";
  els.resultTitle.textContent = "正在處理";
  els.resultSummary.innerHTML = "系統正在瀏覽器內處理檔案，完成後可直接下載。";

  try {
    const result = await convertFile(currentFile, currentFormat, currentTarget);
    const downloadName = replaceExtension(currentFile.name, currentTarget);
    const downloadUrl = URL.createObjectURL(result.blob);

    clearDownloadUrl();
    activeDownloadUrl = downloadUrl;
    latestOutput = {
      blob: result.blob,
      fileName: downloadName,
      mimeType: result.blob.type || guessMimeTypeFromExtension(currentTarget),
      url: downloadUrl,
      extension: currentTarget,
    };
    els.downloadBtn.href = downloadUrl;
    els.downloadBtn.download = downloadName;
    els.downloadBtn.classList.remove("disabled");
    updateShareUi();
    els.resultTitle.textContent = downloadName;
    els.resultSummary.innerHTML = `
      <p><span class="status-ok">轉換完成</span></p>
      <p>來源格式：${LOCAL_FORMATS[currentFormat].label}</p>
      <p>輸出格式：${LOCAL_FORMATS[currentTarget].label}</p>
      <p>${result.summary}</p>
    `;
  } catch (error) {
    els.resultTitle.textContent = "轉換失敗";
    els.resultSummary.innerHTML = `
      <p><span class="status-warn">這次轉換沒有成功。</span></p>
      <p>${escapeHtml(error.message || "發生未知錯誤。")}</p>
    `;
  } finally {
    els.convertBtn.disabled = false;
    els.convertBtn.textContent = "開始轉換";
  }
}

async function convertFile(file, sourceFormat, targetFormat) {
  if (["jpg", "png", "webp"].includes(sourceFormat)) {
    return await convertImage(file, sourceFormat, targetFormat);
  }

  if (sourceFormat === "pdf") {
    return await convertPdf(file, targetFormat);
  }

  if (sourceFormat === "docx") {
    return await convertDocx(file, targetFormat);
  }

  if (sourceFormat === "pptx") {
    return await convertPptx(file, targetFormat);
  }

  if (["txt", "md", "html", "json", "xml"].includes(sourceFormat)) {
    return await convertTextLike(file, sourceFormat, targetFormat);
  }

  if (["csv", "xlsx"].includes(sourceFormat)) {
    return await convertTable(file, sourceFormat, targetFormat);
  }

  throw new Error("這個格式目前不在本地輕量轉換範圍。");
}

async function convertImage(file, sourceFormat, targetFormat) {
  const image = await loadImage(file);

  if (targetFormat === "pdf") {
    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([image.width, image.height]);
    const bytes = await file.arrayBuffer();
    let embeddedImage;

    if (sourceFormat === "png") {
      embeddedImage = await pdfDoc.embedPng(bytes);
    } else {
      embeddedImage = await pdfDoc.embedJpg(await normalizeImageToJpegBytes(image));
    }

    page.drawImage(embeddedImage, {
      x: 0,
      y: 0,
      width: image.width,
      height: image.height,
    });

    return {
      blob: new Blob([await pdfDoc.save()], { type: "application/pdf" }),
      summary: "已把圖片嵌入單頁 PDF。",
    };
  }

  const mimeMap = {
    jpg: "image/jpeg",
    png: "image/png",
    webp: "image/webp",
  };

  return {
    blob: await canvasToBlob(drawImageToCanvas(image), mimeMap[targetFormat], 0.92),
    summary: `已將圖片轉為 ${LOCAL_FORMATS[targetFormat].label}。`,
  };
}

async function convertPdf(file, targetFormat) {
  const pdf = await openPdfDocument(await file.arrayBuffer());
  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const pageText = await extractPageText(pdf, pageNumber);
    if (pageText) pages.push(pageText);
  }

  const combinedText = pages.join("\n\n").trim();
  if (!combinedText) {
    throw new Error("這份 PDF 沒有可擷取的文字層，這一版無法穩定轉成 TXT。");
  }

  if (targetFormat === "txt") {
    return {
      blob: new Blob([combinedText], { type: "text/plain;charset=utf-8" }),
      summary: `已擷取 ${pdf.numPages} 頁文字內容並輸出為 TXT。`,
    };
  }

  if (targetFormat === "docx") {
    return {
      blob: await textToDocxBlob(combinedText),
      summary: `已擷取 ${pdf.numPages} 頁文字內容並整理為 DOCX。`,
    };
  }

  throw new Error("PDF 目前只支援 TXT 或 DOCX。");
}

async function convertTextLike(file, sourceFormat, targetFormat) {
  const normalizedText = await normalizeTextFile(file, sourceFormat);

  if (targetFormat === "txt" || targetFormat === "md") {
    return {
      blob: new Blob([normalizedText], { type: "text/plain;charset=utf-8" }),
      summary: `已輸出為 ${LOCAL_FORMATS[targetFormat].label}。`,
    };
  }

  if (targetFormat === "html") {
    const html = wrapHtmlDocument(escapeHtml(normalizedText).replace(/\n/g, "<br />"));
    return {
      blob: new Blob([html], { type: "text/html;charset=utf-8" }),
      summary: "已將內容整理成可直接開啟的 HTML。",
    };
  }

  if (targetFormat === "pdf") {
    return {
      blob: await textToPdfBlob(normalizedText, { title: currentFile.name }),
      summary: "已將文字內容排版成 PDF。",
    };
  }

  throw new Error("這組文字格式轉換目前未支援。");
}

async function convertDocx(file, targetFormat) {
  const arrayBuffer = await file.arrayBuffer();
  const htmlResult = await mammoth.convertToHtml({ arrayBuffer });
  const rawTextResult = await mammoth.extractRawText({ arrayBuffer });
  const text = rawTextResult.value.trim();

  if (targetFormat === "txt") {
    return {
      blob: new Blob([text], { type: "text/plain;charset=utf-8" }),
      summary: "已抽出 DOCX 的文字內容。",
    };
  }

  if (targetFormat === "html") {
    return {
      blob: new Blob([htmlResult.value], { type: "text/html;charset=utf-8" }),
      summary: "已把 DOCX 內容整理成 HTML。",
    };
  }

  if (targetFormat === "pdf") {
    return {
      blob: await textToPdfBlob(text || stripHtml(htmlResult.value), { title: currentFile.name }),
      summary: "已將 DOCX 的文字內容重新排版輸出成 PDF。",
    };
  }

  throw new Error("DOCX 目前只支援 TXT、HTML、PDF。");
}

async function convertPptx(file, targetFormat) {
  const text = await extractPptxText(file);
  if (!text.trim()) {
    throw new Error("這份 PPTX 沒有可擷取的投影片文字內容。");
  }

  if (targetFormat === "txt") {
    return {
      blob: new Blob([text], { type: "text/plain;charset=utf-8" }),
      summary: "已輸出為 TXT。",
    };
  }

  if (targetFormat === "html") {
    return {
      blob: new Blob([wrapHtmlDocument(escapeHtml(text).replace(/\n/g, "<br />"))], {
        type: "text/html;charset=utf-8",
      }),
      summary: "已整理成 HTML。",
    };
  }

  if (targetFormat === "pdf") {
    return {
      blob: await textToPdfBlob(text, { title: currentFile.name }),
      summary: "已將投影片文字內容排版成 PDF。",
    };
  }

  throw new Error("PPTX 目前只支援 TXT、HTML、PDF。");
}

async function convertTable(file, sourceFormat, targetFormat) {
  const workbook =
    sourceFormat === "csv"
      ? XLSX.read(await file.text(), { type: "string" })
      : XLSX.read(await file.arrayBuffer(), { type: "array" });

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  if (targetFormat === "csv") {
    return {
      blob: new Blob([XLSX.utils.sheet_to_csv(sheet)], { type: "text/csv;charset=utf-8" }),
      summary: `已將工作表 ${sheetName} 匯出為 CSV。`,
    };
  }

  if (targetFormat === "xlsx") {
    const nextWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(nextWorkbook, sheet, "Sheet1");
    const array = XLSX.write(nextWorkbook, { bookType: "xlsx", type: "array" });
    return {
      blob: new Blob([array], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
      summary: "已輸出為 XLSX 試算表。",
    };
  }

  if (targetFormat === "json") {
    return {
      blob: new Blob([JSON.stringify(rowsToObjects(rows), null, 2)], {
        type: "application/json;charset=utf-8",
      }),
      summary: "已將表格資料整理為 JSON。",
    };
  }

  if (targetFormat === "html") {
    return {
      blob: new Blob([wrapHtmlDocument(rowsToHtmlTable(rows))], {
        type: "text/html;charset=utf-8",
      }),
      summary: "已將表格內容整理成 HTML。",
    };
  }

  if (targetFormat === "pdf") {
    const flatText = rows
      .map((row) => row.map((cell) => String(cell)).join(" | "))
      .join("\n");
    return {
      blob: await textToPdfBlob(flatText, { title: `${sheetName} 表格摘要` }),
      summary: "已將表格內容整理成 PDF 摘要。",
    };
  }

  throw new Error("這組表格格式轉換目前未支援。");
}

async function normalizeTextFile(file, sourceFormat) {
  const raw = await file.text();
  if (sourceFormat === "json") {
    return JSON.stringify(JSON.parse(raw), null, 2);
  }
  if (sourceFormat === "html") {
    return stripHtml(raw);
  }
  return raw;
}

async function textToDocxBlob(text) {
  const paragraphs = String(text || "")
    .replace(/\r/g, "")
    .split(/\n+/)
    .map((line) =>
      new Paragraph({
        children: [new TextRun(line || " ")],
        spacing: { after: 140 },
      })
    );

  const doc = new Document({
    sections: [{ properties: {}, children: paragraphs }],
  });

  return await Packer.toBlob(doc);
}

async function textToPdfBlob(text, options = {}) {
  const pdfDoc = await PDFDocument.create();
  const pageWidth = 595;
  const pageHeight = 842;
  const scale = 2;
  const canvas = document.createElement("canvas");
  canvas.width = pageWidth * scale;
  canvas.height = pageHeight * scale;
  const context = canvas.getContext("2d");
  const margin = 48 * scale;
  const bodySize = 11 * scale;
  const titleSize = 16 * scale;
  const lineHeight = 18 * scale;
  const pageLines = textToWrappedLines(text, context, canvas.width - margin * 2, bodySize, titleSize, options.title);
  const capacity = Math.max(1, Math.floor((canvas.height - margin * 2) / lineHeight));

  for (let index = 0; index < pageLines.length; index += capacity) {
    const lines = pageLines.slice(index, index + capacity);
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.textBaseline = "top";

    let cursorY = margin;
    for (const line of lines) {
      context.fillStyle = "#1c1914";
      context.font =
        line.type === "title"
          ? `700 ${titleSize}px "Noto Sans TC", sans-serif`
          : `${bodySize}px "Noto Sans TC", sans-serif`;
      context.fillText(line.text || " ", margin, cursorY);
      cursorY += line.height;
    }

    const pngBlob = await canvasToBlob(canvas, "image/png", 1);
    const embedded = await pdfDoc.embedPng(await pngBlob.arrayBuffer());
    const page = pdfDoc.addPage([pageWidth, pageHeight]);
    page.drawImage(embedded, { x: 0, y: 0, width: pageWidth, height: pageHeight });
  }

  return new Blob([await pdfDoc.save()], { type: "application/pdf" });
}

function textToWrappedLines(text, context, maxWidth, bodySize, titleSize, title) {
  const lines = [];
  const paragraphs = String(text || " ").replace(/\r/g, "").split("\n");

  if (title) {
    context.font = `700 ${titleSize}px "Noto Sans TC", sans-serif`;
    lines.push({ text: title, height: titleSize + bodySize * 0.8, type: "title" });
    lines.push({ text: " ", height: bodySize, type: "body" });
  }

  context.font = `${bodySize}px "Noto Sans TC", sans-serif`;
  for (const paragraph of paragraphs) {
    const wrapped = wrapCanvasText(paragraph || " ", context, maxWidth);
    for (const line of wrapped) {
      lines.push({ text: line, height: bodySize * 1.6, type: "body" });
    }
    if (!paragraph.trim()) {
      lines.push({ text: " ", height: bodySize * 1.2, type: "body" });
    }
  }

  return lines.length ? lines : [{ text: " ", height: bodySize * 1.6, type: "body" }];
}

function wrapCanvasText(text, context, maxWidth) {
  const lines = [];
  let current = "";

  for (const char of Array.from(text)) {
    const candidate = `${current}${char}`;
    if (!current || context.measureText(candidate).width <= maxWidth) {
      current = candidate;
    } else {
      lines.push(current);
      current = char === " " ? "" : char;
    }
  }

  if (current) lines.push(current);
  return lines.length ? lines : [" "];
}

function buildPagedTextPreview(text, options = {}) {
  const lines = String(text || "").replace(/\r/g, "").split("\n");
  const pageSize = options.pageSize || 28;
  const pages = [];

  for (let index = 0; index < lines.length; index += pageSize) {
    pages.push(lines.slice(index, index + pageSize));
  }

  const normalizedPages = pages.length ? pages : [[""]];
  return {
    title: options.title || "文字預覽",
    pageLabel: options.pageLabel || "頁",
    totalPages: normalizedPages.length,
    pages: normalizedPages.slice(0, 8),
    summaryText: truncate(
      lines.filter((line) => line.trim()).slice(0, 4).join(" "),
      220
    ),
  };
}

async function extractPptxText(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const slideNames = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((left, right) => {
      const a = Number(left.match(/slide(\d+)\.xml$/)?.[1] || 0);
      const b = Number(right.match(/slide(\d+)\.xml$/)?.[1] || 0);
      return a - b;
    });

  const slides = [];
  for (const [index, slideName] of slideNames.entries()) {
    const xml = await zip.file(slideName)?.async("string");
    if (!xml) continue;
    const texts = [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)]
      .map((match) => decodeXmlEntities(match[1]))
      .map((value) => value.trim())
      .filter(Boolean);
    slides.push(`投影片 ${index + 1}\n${texts.join("\n")}`.trim());
  }

  return slides.join("\n\n");
}

function decodeXmlEntities(text) {
  return text
    .replaceAll("&amp;", "&")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'");
}

function loadImage(file) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("無法讀取圖片檔案。"));
    image.src = URL.createObjectURL(file);
  });
}

function drawImageToCanvas(image) {
  const canvas = document.createElement("canvas");
  canvas.width = image.naturalWidth || image.width;
  canvas.height = image.naturalHeight || image.height;
  const context = canvas.getContext("2d");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.drawImage(image, 0, 0);
  return canvas;
}

async function normalizeImageToJpegBytes(image) {
  const blob = await canvasToBlob(drawImageToCanvas(image), "image/jpeg", 0.92);
  return await blob.arrayBuffer();
}

function canvasToBlob(canvas, type, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error("無法產生輸出檔案。"));
    }, type, quality);
  });
}

function rowsToObjects(rows) {
  const [headerRow = [], ...dataRows] = rows;
  return dataRows.map((row) => {
    return headerRow.reduce((accumulator, key, index) => {
      const nextKey = String(key || `column_${index + 1}`);
      accumulator[nextKey] = row[index] ?? "";
      return accumulator;
    }, {});
  });
}

function rowsToHtmlTable(rows) {
  if (!rows.length) {
    return "<p>沒有可輸出的資料。</p>";
  }

  const [headerRow = [], ...dataRows] = rows;
  const head = `<tr>${headerRow
    .map((cell) => `<th>${escapeHtml(String(cell))}</th>`)
    .join("")}</tr>`;
  const body = dataRows
    .map(
      (row) =>
        `<tr>${row
          .map((cell) => `<td>${escapeHtml(String(cell ?? ""))}</td>`)
          .join("")}</tr>`
    )
    .join("");

  return `
    <style>
      body { font-family: "Noto Sans TC", sans-serif; padding: 24px; color: #1c1914; }
      table { width: 100%; border-collapse: collapse; }
      th, td { border: 1px solid #d8d2c8; padding: 10px 12px; text-align: left; }
      th { background: #f2ede3; }
    </style>
    <table>
      <thead>${head}</thead>
      <tbody>${body}</tbody>
    </table>
  `;
}

function wrapHtmlDocument(content) {
  return `
    <!DOCTYPE html>
    <html lang="zh-Hant">
      <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>格式轉換</title>
      </head>
      <body>${content}</body>
    </html>
  `;
}

function stripHtml(html) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  return doc.body.textContent || "";
}

function replaceExtension(name, targetExtension) {
  const base = name.includes(".") ? name.replace(/\.[^.]+$/, "") : name;
  return `${base}.${targetExtension}`;
}

function formatFileSize(bytes) {
  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(value >= 10 || unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
}

function truncate(text, limit) {
  const normalized = String(text || "");
  return normalized.length > limit ? `${normalized.slice(0, limit)}...` : normalized;
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function resetResult() {
  els.resultTitle.textContent = "尚未產生輸出檔";
  els.resultSummary.textContent = "轉換完成後，這裡會顯示輸出格式、檔名與處理摘要。";
  els.downloadBtn.classList.add("disabled");
  els.downloadBtn.removeAttribute("href");
  latestOutput = null;
  els.shareBtn.classList.add("hidden");
  els.shareHint.classList.add("hidden");
  els.shareHint.textContent = "";
}

function clearDownloadUrl() {
  if (activeDownloadUrl) {
    URL.revokeObjectURL(activeDownloadUrl);
    activeDownloadUrl = null;
  }
}

function resetApp() {
  clearDownloadUrl();
  currentFile = null;
  currentFormat = null;
  currentTarget = null;
  els.fileInput.value = "";
  els.sourceBadge.textContent = "尚未上傳";
  els.targetBadge.textContent = "請先選擇";
  els.fileName.textContent = "-";
  els.fileSize.textContent = "-";
  els.detectedFormat.textContent = "-";
  els.targetSelect.innerHTML = "<option>請先上傳檔案</option>";
  els.targetSelect.disabled = true;
  els.targetChips.innerHTML = '<span class="chip muted">等待上傳檔案</span>';
  els.conversionNotes.textContent = "上傳檔案後，這裡會顯示相容性與可能的限制。";
  els.previewContent.className = "preview-content empty";
  els.previewContent.textContent = "上傳檔案後，這裡會顯示預覽、頁面縮圖或文字摘要。";
  els.convertBtn.disabled = true;
  resetResult();
}

async function handleShareAction(event) {
  if (!latestOutput) {
    return;
  }

  try {
    const shareData = buildShareData();
    if (!shareData) {
      triggerDownloadFallback();
      return;
    }

    event.preventDefault();
    await navigator.share(shareData);
  } catch (error) {
    if (error?.name !== "AbortError") {
      triggerDownloadFallback();
    }
  }
}

function buildShareData() {
  if (!supportsFileShare()) {
    return null;
  }

  const shareFile = new File([latestOutput.blob], latestOutput.fileName, {
    type: latestOutput.mimeType,
    lastModified: Date.now(),
  });

  const shareData = {
    files: [shareFile],
    title: latestOutput.fileName,
  };

  if (typeof navigator.canShare === "function" && !navigator.canShare(shareData)) {
    return null;
  }

  return shareData;
}

function supportsFileShare() {
  return isMobileDevice() && typeof navigator.share === "function";
}

function isMobileDevice() {
  const ua = navigator.userAgent || "";
  const platform = navigator.platform || "";
  const maxTouchPoints = navigator.maxTouchPoints || 0;

  return (
    /Android|iPhone|iPad|iPod/i.test(ua) ||
    (/Mac/i.test(platform) && maxTouchPoints > 1)
  );
}

function updateShareUi() {
  const canShareFiles = !!buildShareData();
  const shouldShowShare = canShareFiles;

  els.shareBtn.classList.toggle("hidden", !shouldShowShare);

  if (!latestOutput || !shouldShowShare) {
    els.shareHint.classList.add("hidden");
    els.shareHint.textContent = "";
    return;
  }

  const isImage = /^image\//.test(latestOutput.mimeType);
  els.shareHint.textContent = isImage
    ? "手機上可用分享面板嘗試儲存到照片，或分享到檔案、AirDrop、WhatsApp、Discord 等支援的 App。"
    : "PDF 或文件類型通常會分享到「檔案」或支援該檔案類型的 App，例如 AirDrop、WhatsApp、Discord 等。";
  els.shareHint.classList.remove("hidden");
}

function triggerDownloadFallback() {
  if (!latestOutput?.url) {
    return;
  }
  els.downloadBtn.click();
}

function guessMimeTypeFromExtension(extension) {
  const mimeMap = {
    txt: "text/plain;charset=utf-8",
    md: "text/markdown;charset=utf-8",
    html: "text/html;charset=utf-8",
    json: "application/json;charset=utf-8",
    xml: "application/xml;charset=utf-8",
    csv: "text/csv;charset=utf-8",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    pdf: "application/pdf",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    png: "image/png",
    webp: "image/webp",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  };

  return mimeMap[String(extension || "").toLowerCase()] || "application/octet-stream";
}
