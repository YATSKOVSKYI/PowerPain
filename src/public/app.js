(() => {
  "use strict";

  const MAX_SIZE = 200 * 1024 * 1024;
  const WALLET = "TBxEquczDy6ZSRPAyYrNbczoaP9YThaJuZ";

  // ─── i18n ───────────────────────────────────────────────────

  const translations = {
    en: {
      badge: "Open-source tool",
      heroLine1: "Font won't change",
      heroLine2: "in ",
      heroSub: "Upload your broken .pptx file — we'll normalize all fonts, fix CJK issues, and clean up theme references. In seconds.",
      heroCta: "Fix it now",
      toolLabel: "Upload & Fix",
      toolTitle: "Drop your file below",
      dropTitle: "Drop .pptx file here",
      dropText: "or click to browse",
      fontLabel: "Target font",
      btnProcess: "Fix Fonts",
      btnDownload: "Download Fixed File",
      btnReset: "Process another file",
      statusUploading: "Uploading file...",
      statusProcessing: "Processing — fixing fonts...",
      statusDone: "Done! Your file is ready.",
      statusError: "Processing failed",
      featuresLabel: "How it works",
      featuresTitle: "Three simple steps",
      step1Title: "Upload",
      step1Text: "Drag and drop your .pptx file or click to browse. Up to 200 MB.",
      step2Title: "Process",
      step2Text: "We normalize fonts in slides, layouts, masters, and themes. Fix CJK issues and theme references.",
      step3Title: "Download",
      step3Text: "Get your fixed file instantly. All temporary data is deleted right after processing.",
      privacy: "Your files are processed on the server and deleted immediately after. Nothing is stored, logged, or shared.",
      donateHeadline1: "This takes hours.",
      donateHeadline2: "Your donation is ",
      donateFuel: "fuel",
      donateLabel: "Support the project",
      donateCTitle: "Liked the tool?",
      donateCText: "This is a free open-source project. If it saved you time on a thesis or an article — support the development with a few USDT.",
      cryptoTitle: "Cryptocurrency USDT",
      cryptoDesc: "TRON network (TRC-20). Only USDT — other assets will be lost.",
      btnCopy: "Copy",
      warningLabel: "Warning:",
      warningText: " send only TRON network assets to this address.",
      painkiller: "pain killer",
    },
    ru: {
      badge: "Open-source инструмент",
      heroLine1: "Шрифт не меняется",
      heroLine2: "в ",
      heroSub: "Загрузите ваш .pptx файл — мы нормализуем все шрифты, исправим CJK-проблемы и очистим ссылки на темы. За секунды.",
      heroCta: "Исправить сейчас",
      toolLabel: "Загрузка и исправление",
      toolTitle: "Перетащите файл сюда",
      dropTitle: "Перетащите .pptx файл сюда",
      dropText: "или нажмите для выбора",
      fontLabel: "Целевой шрифт",
      btnProcess: "Исправить шрифты",
      btnDownload: "Скачать исправленный файл",
      btnReset: "Обработать другой файл",
      statusUploading: "Загрузка файла...",
      statusProcessing: "Обработка — исправление шрифтов...",
      statusDone: "Готово! Ваш файл готов.",
      statusError: "Ошибка обработки",
      featuresLabel: "Как это работает",
      featuresTitle: "Три простых шага",
      step1Title: "Загрузите",
      step1Text: "Перетащите ваш .pptx файл или нажмите для выбора. До 200 МБ.",
      step2Title: "Обработка",
      step2Text: "Мы нормализуем шрифты в слайдах, макетах, мастерах и темах. Исправляем CJK и ссылки тем.",
      step3Title: "Скачайте",
      step3Text: "Получите исправленный файл мгновенно. Все временные данные удаляются сразу после обработки.",
      privacy: "Ваши файлы обрабатываются на сервере и удаляются сразу после. Ничего не сохраняется, не логируется и не передаётся.",
      donateHeadline1: "На это уходят часы.",
      donateHeadline2: "Ваш донат — ",
      donateFuel: "топливо",
      donateLabel: "Поддержать проект",
      donateCTitle: "Инструмент понравился?",
      donateCText: "Это бесплатный open-source проект. Если он сэкономил вам время на диплом или статью — поддержите разработку парой USDT.",
      cryptoTitle: "Криптовалюта USDT",
      cryptoDesc: "Сеть TRON (TRC-20). Только USDT — другие активы будут потеряны.",
      btnCopy: "Копировать",
      warningLabel: "Внимание:",
      warningText: " отправляйте только активы сети TRON на этот адрес.",
      painkiller: "обезбол",
    },
    uk: {
      badge: "Open-source інструмент",
      heroLine1: "Шрифт не змінюється",
      heroLine2: "в ",
      heroSub: "Завантажте ваш .pptx файл — ми нормалізуємо всі шрифти, виправимо CJK-проблеми та очистимо посилання на теми. За секунди.",
      heroCta: "Виправити зараз",
      toolLabel: "Завантаження і виправлення",
      toolTitle: "Перетягніть файл сюди",
      dropTitle: "Перетягніть .pptx файл сюди",
      dropText: "або натисніть для вибору",
      fontLabel: "Цільовий шрифт",
      btnProcess: "Виправити шрифти",
      btnDownload: "Завантажити виправлений файл",
      btnReset: "Обробити інший файл",
      statusUploading: "Завантаження файлу...",
      statusProcessing: "Обробка — виправлення шрифтів...",
      statusDone: "Готово! Ваш файл готовий.",
      statusError: "Помилка обробки",
      featuresLabel: "Як це працює",
      featuresTitle: "Три простих кроки",
      step1Title: "Завантажте",
      step1Text: "Перетягніть ваш .pptx файл або натисніть для вибору. До 200 МБ.",
      step2Title: "Обробка",
      step2Text: "Ми нормалізуємо шрифти в слайдах, макетах, мастерах та темах. Виправляємо CJK та посилання тем.",
      step3Title: "Завантажте",
      step3Text: "Отримайте виправлений файл миттєво. Усі тимчасові дані видаляються одразу після обробки.",
      privacy: "Ваші файли обробляються на сервері та видаляються одразу після. Нічого не зберігається, не логується та не передається.",
      donateHeadline1: "На це йдуть години.",
      donateHeadline2: "Ваш донат — ",
      donateFuel: "паливо",
      donateLabel: "Підтримати проєкт",
      donateCTitle: "Сподобався інструмент?",
      donateCText: "Це безкоштовний open-source проєкт. Якщо він зекономив вам час на диплом чи статтю — підтримайте розробку парою USDT.",
      cryptoTitle: "Криптовалюта USDT",
      cryptoDesc: "Мережа TRON (TRC-20). Тільки USDT — інші активи будуть втрачені.",
      btnCopy: "Копіювати",
      warningLabel: "Увага:",
      warningText: " відправляйте тільки активи мережі TRON на цю адресу.",
      painkiller: "знеболююче",
    },
    zh: {
      badge: "开源工具",
      heroLine1: "字体无法更改",
      heroLine2: "— ",
      heroSub: "上传您的 .pptx 文件 — 我们将规范化所有字体、修复 CJK 问题并清理主题引用。仅需几秒。",
      heroCta: "立即修复",
      toolLabel: "上传并修复",
      toolTitle: "将文件拖放到下方",
      dropTitle: "将 .pptx 文件拖放到此处",
      dropText: "或点击浏览",
      fontLabel: "目标字体",
      btnProcess: "修复字体",
      btnDownload: "下载修复后的文件",
      btnReset: "处理另一个文件",
      statusUploading: "正在上传文件...",
      statusProcessing: "处理中 — 修复字体...",
      statusDone: "完成！您的文件已准备好。",
      statusError: "处理失败",
      featuresLabel: "工作原理",
      featuresTitle: "三个简单步骤",
      step1Title: "上传",
      step1Text: "拖放您的 .pptx 文件或点击浏览。最大 200 MB。",
      step2Title: "处理",
      step2Text: "我们在幻灯片、布局、母版和主题中规范化字体。修复 CJK 问题和主题引用。",
      step3Title: "下载",
      step3Text: "立即获取修复后的文件。所有临时数据在处理后立即删除。",
      privacy: "您的文件在服务器上处理并立即删除。不会存储、记录或共享任何内容。",
      donateHeadline1: "这需要数小时。",
      donateHeadline2: "您的捐赠是",
      donateFuel: "动力",
      donateLabel: "支持项目",
      donateCTitle: "喜欢这个工具？",
      donateCText: "这是一个免费的开源项目。如果它为您的论文或文章节省了时间 — 请用几个 USDT 支持开发。",
      cryptoTitle: "加密货币 USDT",
      cryptoDesc: "TRON 网络 (TRC-20)。仅限 USDT — 其他资产将丢失。",
      btnCopy: "复制",
      warningLabel: "警告：",
      warningText: "仅向此地址发送 TRON 网络资产。",
      painkiller: "止痛药",
    },
    de: {
      badge: "Open-Source-Tool",
      heroLine1: "Schriftart ändert sich nicht",
      heroLine2: "in ",
      heroSub: "Laden Sie Ihre defekte .pptx-Datei hoch — wir normalisieren alle Schriftarten, beheben CJK-Probleme und bereinigen Theme-Referenzen. In Sekunden.",
      heroCta: "Jetzt reparieren",
      toolLabel: "Hochladen & Reparieren",
      toolTitle: "Datei hier ablegen",
      dropTitle: ".pptx-Datei hierher ziehen",
      dropText: "oder klicken zum Durchsuchen",
      fontLabel: "Zielschriftart",
      btnProcess: "Schriftarten reparieren",
      btnDownload: "Reparierte Datei herunterladen",
      btnReset: "Weitere Datei verarbeiten",
      statusUploading: "Datei wird hochgeladen...",
      statusProcessing: "Verarbeitung — Schriftarten werden repariert...",
      statusDone: "Fertig! Ihre Datei ist bereit.",
      statusError: "Verarbeitung fehlgeschlagen",
      featuresLabel: "So funktioniert es",
      featuresTitle: "Drei einfache Schritte",
      step1Title: "Hochladen",
      step1Text: "Ziehen Sie Ihre .pptx-Datei per Drag & Drop oder klicken Sie zum Durchsuchen. Bis zu 200 MB.",
      step2Title: "Verarbeiten",
      step2Text: "Wir normalisieren Schriftarten in Folien, Layouts, Mastern und Designs. CJK-Probleme und Theme-Referenzen werden behoben.",
      step3Title: "Herunterladen",
      step3Text: "Erhalten Sie Ihre reparierte Datei sofort. Alle temporären Daten werden direkt nach der Verarbeitung gelöscht.",
      privacy: "Ihre Dateien werden auf dem Server verarbeitet und sofort danach gelöscht. Nichts wird gespeichert, protokolliert oder weitergegeben.",
      donateHeadline1: "Das dauert Stunden.",
      donateHeadline2: "Ihre Spende ist ",
      donateFuel: "Treibstoff",
      donateLabel: "Projekt unterstützen",
      donateCTitle: "Tool gefällt Ihnen?",
      donateCText: "Dies ist ein kostenloses Open-Source-Projekt. Wenn es Ihnen Zeit bei einer Arbeit gespart hat — unterstützen Sie die Entwicklung mit ein paar USDT.",
      cryptoTitle: "Kryptowährung USDT",
      cryptoDesc: "TRON-Netzwerk (TRC-20). Nur USDT — andere Assets gehen verloren.",
      btnCopy: "Kopieren",
      warningLabel: "Achtung:",
      warningText: " senden Sie nur TRON-Netzwerk-Assets an diese Adresse.",
      painkiller: "Schmerzmittel",
    },
    es: {
      badge: "Herramienta open-source",
      heroLine1: "¿La fuente no cambia",
      heroLine2: "en ",
      heroSub: "Sube tu archivo .pptx dañado — normalizaremos todas las fuentes, arreglaremos problemas CJK y limpiaremos las referencias del tema. En segundos.",
      heroCta: "Arreglar ahora",
      toolLabel: "Subir y arreglar",
      toolTitle: "Arrastra tu archivo aquí",
      dropTitle: "Arrastra el archivo .pptx aquí",
      dropText: "o haz clic para buscar",
      fontLabel: "Fuente objetivo",
      btnProcess: "Arreglar fuentes",
      btnDownload: "Descargar archivo arreglado",
      btnReset: "Procesar otro archivo",
      statusUploading: "Subiendo archivo...",
      statusProcessing: "Procesando — arreglando fuentes...",
      statusDone: "¡Listo! Tu archivo está preparado.",
      statusError: "Error de procesamiento",
      featuresLabel: "Cómo funciona",
      featuresTitle: "Tres pasos simples",
      step1Title: "Subir",
      step1Text: "Arrastra tu archivo .pptx o haz clic para buscar. Hasta 200 MB.",
      step2Title: "Procesar",
      step2Text: "Normalizamos las fuentes en diapositivas, diseños, maestros y temas. Arreglamos problemas CJK y referencias de temas.",
      step3Title: "Descargar",
      step3Text: "Obtén tu archivo arreglado al instante. Todos los datos temporales se eliminan justo después del procesamiento.",
      privacy: "Tus archivos se procesan en el servidor y se eliminan inmediatamente. No se almacena, registra ni comparte nada.",
      donateHeadline1: "Esto lleva horas.",
      donateHeadline2: "Tu donación es ",
      donateFuel: "combustible",
      donateLabel: "Apoyar el proyecto",
      donateCTitle: "¿Te gustó la herramienta?",
      donateCText: "Este es un proyecto open-source gratuito. Si te ahorró tiempo en una tesis o artículo — apoya el desarrollo con unos USDT.",
      cryptoTitle: "Criptomoneda USDT",
      cryptoDesc: "Red TRON (TRC-20). Solo USDT — otros activos se perderán.",
      btnCopy: "Copiar",
      warningLabel: "Advertencia:",
      warningText: " envía solo activos de la red TRON a esta dirección.",
      painkiller: "analgésico",
    },
    fr: {
      badge: "Outil open-source",
      heroLine1: "La police ne change pas",
      heroLine2: "dans ",
      heroSub: "Téléchargez votre fichier .pptx — nous normaliserons toutes les polices, corrigerons les problèmes CJK et nettoierons les références de thème. En quelques secondes.",
      heroCta: "Réparer maintenant",
      toolLabel: "Télécharger et réparer",
      toolTitle: "Déposez votre fichier ci-dessous",
      dropTitle: "Déposez le fichier .pptx ici",
      dropText: "ou cliquez pour parcourir",
      fontLabel: "Police cible",
      btnProcess: "Réparer les polices",
      btnDownload: "Télécharger le fichier réparé",
      btnReset: "Traiter un autre fichier",
      statusUploading: "Téléchargement du fichier...",
      statusProcessing: "Traitement — réparation des polices...",
      statusDone: "Terminé ! Votre fichier est prêt.",
      statusError: "Échec du traitement",
      featuresLabel: "Comment ça marche",
      featuresTitle: "Trois étapes simples",
      step1Title: "Télécharger",
      step1Text: "Glissez-déposez votre fichier .pptx ou cliquez pour parcourir. Jusqu'à 200 Mo.",
      step2Title: "Traiter",
      step2Text: "Nous normalisons les polices dans les diapositives, mises en page, masques et thèmes. Correction des problèmes CJK et des références de thème.",
      step3Title: "Télécharger",
      step3Text: "Obtenez votre fichier réparé instantanément. Toutes les données temporaires sont supprimées juste après le traitement.",
      privacy: "Vos fichiers sont traités sur le serveur et supprimés immédiatement après. Rien n'est stocké, enregistré ou partagé.",
      donateHeadline1: "Cela prend des heures.",
      donateHeadline2: "Votre don est du ",
      donateFuel: "carburant",
      donateLabel: "Soutenir le projet",
      donateCTitle: "L'outil vous a plu ?",
      donateCText: "C'est un projet open-source gratuit. S'il vous a fait gagner du temps sur un mémoire ou un article — soutenez le développement avec quelques USDT.",
      cryptoTitle: "Cryptomonnaie USDT",
      cryptoDesc: "Réseau TRON (TRC-20). Uniquement USDT — les autres actifs seront perdus.",
      btnCopy: "Copier",
      warningLabel: "Attention :",
      warningText: " envoyez uniquement des actifs du réseau TRON à cette adresse.",
      painkiller: "antidouleur",
    },
  };

  let currentLang = localStorage.getItem("pp-lang") || "en";

  function setLanguage(lang) {
    currentLang = lang;
    localStorage.setItem("pp-lang", lang);
    const t = translations[lang] || translations.en;

    document.querySelectorAll("[data-i18n]").forEach((el) => {
      const key = el.getAttribute("data-i18n");
      if (t[key] !== undefined) {
        el.textContent = t[key];
      }
    });

    // Update glitch data-text
    const glitch = document.querySelector(".glitch");
    if (glitch) glitch.setAttribute("data-text", "PowerPoint");

    // Update lang switcher display
    document.getElementById("langCurrent").textContent = lang.toUpperCase();
    document.querySelectorAll(".lang-option").forEach((opt) => {
      opt.classList.toggle("active", opt.dataset.lang === lang);
    });
  }

  // ─── Language Switcher ──────────────────────────────────────

  const langSwitcher = document.getElementById("langSwitcher");
  const langBtn = document.getElementById("langBtn");

  langBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    langSwitcher.classList.toggle("open");
  });

  document.querySelectorAll(".lang-option").forEach((opt) => {
    opt.addEventListener("click", () => {
      setLanguage(opt.dataset.lang);
      langSwitcher.classList.remove("open");
    });
  });

  document.addEventListener("click", () => {
    langSwitcher.classList.remove("open");
  });

  // Apply saved language
  setLanguage(currentLang);

  // ─── DOM Elements ───────────────────────────────────────────

  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");
  const fileInfo = document.getElementById("fileInfo");
  const fileName = document.getElementById("fileName");
  const fileSize = document.getElementById("fileSize");
  const fileRemove = document.getElementById("fileRemove");
  const fontSelect = document.getElementById("fontSelect");
  const btnProcess = document.getElementById("btnProcess");
  const status = document.getElementById("status");
  const statusBar = document.getElementById("statusBar");
  const statusIcon = document.getElementById("statusIcon");
  const statusText = document.getElementById("statusText");
  const downloadArea = document.getElementById("downloadArea");
  const btnDownload = document.getElementById("btnDownload");
  const statsText = document.getElementById("statsText");
  const btnReset = document.getElementById("btnReset");
  const btnCopy = document.getElementById("btnCopy");

  let selectedFile = null;
  let resultBlob = null;
  let resultName = "";

  // ─── File Selection ─────────────────────────────────────────

  dropzone.addEventListener("click", () => {
    if (!btnProcess.disabled || !selectedFile) fileInput.click();
  });

  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });

  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
  });

  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
  });

  fileInput.addEventListener("change", () => {
    if (fileInput.files.length > 0) handleFile(fileInput.files[0]);
  });

  fileRemove.addEventListener("click", (e) => {
    e.stopPropagation();
    resetState();
  });

  function handleFile(file) {
    const t = translations[currentLang] || translations.en;

    if (!file.name.toLowerCase().endsWith(".pptx")) {
      showStatus("error", t.statusError + ": only .pptx files accepted");
      return;
    }

    if (file.size > MAX_SIZE) {
      showStatus("error", t.statusError + ": file too large (max 200 MB)");
      return;
    }

    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatSize(file.size);
    fileInfo.classList.add("visible");
    dropzone.classList.add("has-file");
    btnProcess.disabled = false;
    hideStatus();
    downloadArea.classList.remove("visible");
  }

  // ─── Processing ─────────────────────────────────────────────

  btnProcess.addEventListener("click", async () => {
    if (!selectedFile) return;

    const t = translations[currentLang] || translations.en;
    btnProcess.disabled = true;
    downloadArea.classList.remove("visible");

    showStatus("uploading", t.statusUploading);

    const formData = new FormData();
    formData.append("file", selectedFile);
    formData.append("targetFont", fontSelect.value);

    try {
      showStatus("processing", t.statusProcessing);

      const response = await fetch("/api/process", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const err = await response.json().catch(() => ({ error: "Server error" }));
        throw new Error(err.error || `Server error (${response.status})`);
      }

      resultBlob = await response.blob();

      const disposition = response.headers.get("Content-Disposition") || "";
      const match = disposition.match(/filename="?([^";\n]+)"?/);
      resultName = match ? decodeURIComponent(match[1]) : selectedFile.name.replace(/\.pptx$/i, "-fixed.pptx");

      const statsHeader = response.headers.get("X-Stats");
      if (statsHeader) {
        try {
          const s = JSON.parse(statsHeader);
          statsText.textContent = `Fixed: ${s.themes} theme(s), ${s.masters} master(s), ${s.layouts} layout(s), ${s.slides} slide(s)`;
        } catch {
          statsText.textContent = "";
        }
      }

      showStatus("done", t.statusDone);
      downloadArea.classList.add("visible");
    } catch (err) {
      showStatus("error", err.message || t.statusError);
      btnProcess.disabled = false;
    }
  });

  // ─── Download ───────────────────────────────────────────────

  btnDownload.addEventListener("click", () => {
    if (!resultBlob) return;
    const url = URL.createObjectURL(resultBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = resultName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  btnReset.addEventListener("click", resetState);

  // ─── Copy Wallet ────────────────────────────────────────────

  btnCopy.addEventListener("click", () => {
    navigator.clipboard.writeText(WALLET).then(() => {
      const t = translations[currentLang] || translations.en;
      btnCopy.textContent = "✓";
      btnCopy.classList.add("copied");
      setTimeout(() => {
        btnCopy.textContent = t.btnCopy;
        btnCopy.classList.remove("copied");
      }, 2000);
    });
  });

  // ─── Helpers ────────────────────────────────────────────────

  function showStatus(type, text) {
    status.classList.add("visible");
    statusBar.className = "status-bar " + type;

    if (type === "uploading" || type === "processing") {
      statusIcon.innerHTML = '<div class="spinner"></div>';
    } else if (type === "done") {
      statusIcon.textContent = "✓";
    } else {
      statusIcon.textContent = "✕";
    }

    statusText.textContent = text;
  }

  function hideStatus() {
    status.classList.remove("visible");
  }

  function resetState() {
    selectedFile = null;
    resultBlob = null;
    resultName = "";
    fileInput.value = "";
    fileInfo.classList.remove("visible");
    dropzone.classList.remove("has-file");
    btnProcess.disabled = true;
    hideStatus();
    downloadArea.classList.remove("visible");
  }

  function formatSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
  }

  // ─── Pain Killer (scroll to top) ────────────────────────────

  const painkiller = document.getElementById("painkiller");

  window.addEventListener("scroll", () => {
    if (window.scrollY > 600) {
      painkiller.classList.add("visible");
    } else {
      painkiller.classList.remove("visible");
    }
  }, { passive: true });

  painkiller.addEventListener("click", () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
})();
