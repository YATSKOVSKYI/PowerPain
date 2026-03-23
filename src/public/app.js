(() => {
  "use strict";

  const MAX_SIZE = 200 * 1024 * 1024;
  const WALLET = "TBxEquczDy6ZSRPAyYrNbczoaP9YThaJuZ";
  const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

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
      statusReading: "Reading file...",
      statusUnpacking: "Unpacking PPTX archive...",
      statusThemes: "Fixing themes",
      statusSlides: "Fixing slides",
      statusLayouts: "Fixing layouts",
      statusMasters: "Fixing masters",
      statusCompressing: "Compressing fixed file...",
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
      privacy: "Your files never leave your device. All processing happens right in your browser — nothing is uploaded to any server.",
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
      statusReading: "Чтение файла...",
      statusUnpacking: "Распаковка PPTX-архива...",
      statusThemes: "Исправление тем",
      statusSlides: "Исправление слайдов",
      statusLayouts: "Исправление макетов",
      statusMasters: "Исправление мастеров",
      statusCompressing: "Сжатие исправленного файла...",
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
      privacy: "Ваши файлы не покидают ваше устройство. Вся обработка происходит прямо в браузере — ничего не загружается на сервер.",
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
      statusReading: "Читання файлу...",
      statusUnpacking: "Розпакування PPTX-архіву...",
      statusThemes: "Виправлення тем",
      statusSlides: "Виправлення слайдів",
      statusLayouts: "Виправлення макетів",
      statusMasters: "Виправлення мастерів",
      statusCompressing: "Стиснення виправленого файлу...",
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
      privacy: "Ваші файли не покидають ваш пристрій. Уся обробка відбувається прямо у браузері — нічого не завантажується на сервер.",
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
      statusReading: "正在读取文件...",
      statusUnpacking: "正在解压 PPTX 存档...",
      statusThemes: "修复主题",
      statusSlides: "修复幻灯片",
      statusLayouts: "修复布局",
      statusMasters: "修复母版",
      statusCompressing: "正在压缩修复后的文件...",
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
      privacy: "您的文件不会离开您的设备。所有处理都在浏览器中完成 — 没有任何内容上传到服务器。",
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
      statusReading: "Datei wird gelesen...",
      statusUnpacking: "PPTX-Archiv wird entpackt...",
      statusThemes: "Themes reparieren",
      statusSlides: "Folien reparieren",
      statusLayouts: "Layouts reparieren",
      statusMasters: "Master reparieren",
      statusCompressing: "Reparierte Datei wird komprimiert...",
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
      privacy: "Ihre Dateien verlassen Ihr Gerät nicht. Die gesamte Verarbeitung erfolgt direkt in Ihrem Browser — nichts wird auf einen Server hochgeladen.",
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
      statusReading: "Leyendo archivo...",
      statusUnpacking: "Descomprimiendo archivo PPTX...",
      statusThemes: "Arreglando temas",
      statusSlides: "Arreglando diapositivas",
      statusLayouts: "Arreglando diseños",
      statusMasters: "Arreglando maestros",
      statusCompressing: "Comprimiendo archivo arreglado...",
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
      privacy: "Tus archivos no salen de tu dispositivo. Todo el procesamiento ocurre en tu navegador — nada se sube a ningún servidor.",
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
      statusReading: "Lecture du fichier...",
      statusUnpacking: "Décompression de l'archive PPTX...",
      statusThemes: "Réparation des thèmes",
      statusSlides: "Réparation des diapositives",
      statusLayouts: "Réparation des mises en page",
      statusMasters: "Réparation des masques",
      statusCompressing: "Compression du fichier réparé...",
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
      privacy: "Vos fichiers ne quittent pas votre appareil. Tout le traitement se fait dans votre navigateur — rien n'est envoyé à un serveur.",
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
    },
  };

  let currentLang = localStorage.getItem("pp-lang") || "en";
  let currentStatusType = null;
  let currentStatusKey = null;

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

    const glitch = document.querySelector(".glitch");
    if (glitch) glitch.setAttribute("data-text", "PowerPoint");

    document.getElementById("langCurrent").textContent = lang.toUpperCase();
    document.querySelectorAll(".lang-option").forEach((opt) => {
      opt.classList.toggle("active", opt.dataset.lang === lang);
    });

    // Update current status text if visible
    if (currentStatusType && currentStatusKey) {
      const statusKeyMap = {
        reading: "statusReading",
        unzip: "statusUnpacking",
        themes: "statusThemes",
        slides: "statusSlides",
        layouts: "statusLayouts",
        masters: "statusMasters",
        zip: "statusCompressing",
        statusDone: "statusDone",
        statusError: "statusError",
      };
      const translationKey = statusKeyMap[currentStatusKey] || currentStatusKey;
      if (t[translationKey]) {
        const statusTextEl = document.getElementById("statusText");
        if (statusTextEl) {
          statusTextEl.textContent = t[translationKey];
        }
      }
    }
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

  // ─── Donate photo fallback ────────────────────────────────
  const donatePhoto = document.getElementById("donatePhoto");
  if (donatePhoto) {
    donatePhoto.addEventListener("error", () => {
      const wrap = document.getElementById("photoWrap");
      if (wrap) wrap.classList.add("hidden");
    });
  }

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

  // ─── Client-side PPTX Processing ──────────────────────────

  function isCJKLocale(lang) {
    return lang.startsWith("zh") || lang.startsWith("ja") || lang.startsWith("ko") || lang === "yue";
  }

  function patchRunProperties(doc, tagName, targetFont) {
    const elements = doc.getElementsByTagNameNS(NS_A, tagName);
    for (let i = 0; i < elements.length; i++) {
      const rPr = elements[i];

      // Fix CJK language
      const lang = rPr.getAttribute("lang");
      if (lang && isCJKLocale(lang)) {
        rPr.setAttribute("lang", "en-US");
        rPr.removeAttribute("altLang");
      }

      // Remove existing font children
      const toRemove = [];
      for (let j = 0; j < rPr.childNodes.length; j++) {
        const child = rPr.childNodes[j];
        if (child.localName && ["latin", "ea", "cs", "sym"].includes(child.localName)) {
          toRemove.push(child);
        }
      }
      toRemove.forEach((c) => rPr.removeChild(c));

      // Add explicit font
      const latin = doc.createElementNS(NS_A, "a:latin");
      latin.setAttribute("typeface", targetFont);
      rPr.appendChild(latin);

      const ea = doc.createElementNS(NS_A, "a:ea");
      ea.setAttribute("typeface", targetFont);
      rPr.appendChild(ea);
    }
  }

  function normalizeXmlFonts(xmlString, targetFont) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlString, "text/xml");

    patchRunProperties(doc, "rPr", targetFont);
    patchRunProperties(doc, "defRPr", targetFont);
    patchRunProperties(doc, "endParaRPr", targetFont);

    let result = new XMLSerializer().serializeToString(doc);

    result = result.replace(/typeface="\+mj-lt"/g, 'typeface="' + targetFont + '"');
    result = result.replace(/typeface="\+mn-lt"/g, 'typeface="' + targetFont + '"');
    result = result.replace(/typeface="\+mj-ea"/g, 'typeface="' + targetFont + '"');
    result = result.replace(/typeface="\+mn-ea"/g, 'typeface="' + targetFont + '"');

    return result;
  }

  function normalizeThemeFonts(xmlString, targetFont) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlString, "text/xml");

    for (const fontGroup of ["majorFont", "minorFont"]) {
      const group = doc.getElementsByTagNameNS(NS_A, fontGroup)[0];
      if (!group) continue;

      let latin = group.getElementsByTagNameNS(NS_A, "latin")[0];
      if (!latin) {
        latin = doc.createElementNS(NS_A, "a:latin");
        group.insertBefore(latin, group.firstChild);
      }
      latin.setAttribute("typeface", targetFont);
    }

    return new XMLSerializer().serializeToString(doc);
  }

  async function processFile(file, targetFont, onProgress) {
    const stats = { slides: 0, layouts: 0, masters: 0, themes: 0 };

    onProgress("reading", 0, 1, 2);

    const arrayBuffer = await file.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);

    // Validate ZIP magic bytes
    if (data[0] !== 0x50 || data[1] !== 0x4B) {
      throw new Error("File is not a valid PPTX archive");
    }

    onProgress("unzip", 0, 1, 5);
    const zip = await JSZip.loadAsync(data);
    onProgress("unzip", 1, 1, 10);

    // Themes
    const themeFiles = Object.keys(zip.files).filter((f) => /^ppt\/theme\/theme\d+\.xml$/.test(f));
    for (let i = 0; i < themeFiles.length; i++) {
      onProgress("themes", i, themeFiles.length, 10 + (5 * i / Math.max(themeFiles.length, 1)));
      const xml = await zip.file(themeFiles[i]).async("string");
      zip.file(themeFiles[i], normalizeThemeFonts(xml, targetFont));
      stats.themes++;
    }

    // Slides
    const slideFiles = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
    for (let i = 0; i < slideFiles.length; i++) {
      const pct = 15 + (45 * i / Math.max(slideFiles.length, 1));
      onProgress("slides", i + 1, slideFiles.length, pct);
      const xml = await zip.file(slideFiles[i]).async("string");
      zip.file(slideFiles[i], normalizeXmlFonts(xml, targetFont));
      stats.slides++;
      // Yield to UI thread every 5 slides
      if (i % 5 === 0) await new Promise((r) => setTimeout(r, 0));
    }

    // Layouts
    const layoutFiles = Object.keys(zip.files).filter((f) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(f));
    for (let i = 0; i < layoutFiles.length; i++) {
      const pct = 60 + (10 * i / Math.max(layoutFiles.length, 1));
      onProgress("layouts", i + 1, layoutFiles.length, pct);
      const xml = await zip.file(layoutFiles[i]).async("string");
      zip.file(layoutFiles[i], normalizeXmlFonts(xml, targetFont));
      stats.layouts++;
    }

    // Masters
    const masterFiles = Object.keys(zip.files).filter((f) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(f));
    for (let i = 0; i < masterFiles.length; i++) {
      const pct = 70 + (5 * i / Math.max(masterFiles.length, 1));
      onProgress("masters", i + 1, masterFiles.length, pct);
      const xml = await zip.file(masterFiles[i]).async("string");
      zip.file(masterFiles[i], normalizeXmlFonts(xml, targetFont));
      stats.masters++;
    }

    onProgress("zip", 0, 1, 75);
    const output = await zip.generateAsync(
      { type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } },
      (meta) => {
        onProgress("zip", meta.percent, 100, 75 + (meta.percent * 0.25));
      }
    );
    onProgress("zip", 100, 100, 100);

    return { output, stats };
  }

  // ─── Processing ─────────────────────────────────────────────

  btnProcess.addEventListener("click", async () => {
    if (!selectedFile) return;

    const t = translations[currentLang] || translations.en;
    btnProcess.disabled = true;
    downloadArea.classList.remove("visible");

    try {
      const { output, stats } = await processFile(selectedFile, fontSelect.value, (stage, current, total, percent) => {
        const pct = Math.round(percent);
        const msgs = {
          reading: t.statusReading,
          unzip: t.statusUnpacking,
          themes: t.statusThemes + ` (${current}/${total})...`,
          slides: t.statusSlides + ` (${current}/${total})...`,
          layouts: t.statusLayouts + ` (${current}/${total})...`,
          masters: t.statusMasters + ` (${current}/${total})...`,
          zip: t.statusCompressing,
        };
        currentStatusType = "processing";
        currentStatusKey = stage;
        showStatus("processing", msgs[stage] || t.statusCompressing, pct);
      });

      const baseName = selectedFile.name.replace(/\.pptx$/i, "");
      resultBlob = output;
      resultName = baseName + "-fixed.pptx";

      statsText.textContent = `Fixed: ${stats.themes} theme(s), ${stats.masters} master(s), ${stats.layouts} layout(s), ${stats.slides} slide(s)`;

      currentStatusType = "done";
      currentStatusKey = "statusDone";
      showStatus("done", t.statusDone, 100);
      downloadArea.classList.add("visible");
    } catch (err) {
      currentStatusType = "error";
      currentStatusKey = "statusError";
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

  function showStatus(type, text, percent) {
    status.classList.add("visible");
    statusBar.className = "status-bar " + type;

    if (type === "uploading" || type === "processing") {
      statusIcon.textContent = "";
      const spinner = document.createElement("div");
      spinner.className = "spinner";
      statusIcon.appendChild(spinner);
    } else if (type === "done") {
      statusIcon.textContent = "✓";
    } else {
      statusIcon.textContent = "✕";
    }

    let progressBar = statusBar.querySelector(".progress-fill");
    if (type === "uploading" || type === "processing") {
      if (!progressBar) {
        const wrapper = document.createElement("div");
        wrapper.className = "progress-track";
        progressBar = document.createElement("div");
        progressBar.className = "progress-fill";
        wrapper.appendChild(progressBar);
        statusBar.appendChild(wrapper);
      }
      const pct = typeof percent === "number" ? Math.min(percent, 100) : 0;
      progressBar.style.width = pct + "%";
      statusText.textContent = text + (pct > 0 ? " (" + pct + "%)" : "");
    } else {
      const track = statusBar.querySelector(".progress-track");
      if (track) track.remove();
      statusText.textContent = text;
    }
  }

  function hideStatus() {
    status.classList.remove("visible");
  }

  function resetState() {
    selectedFile = null;
    resultBlob = null;
    resultName = "";
    currentStatusType = null;
    currentStatusKey = null;
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

  // ─── Scroll to Top ─────────────────────────────────────────
  {
    const scrollBtn = document.getElementById("scrollTop");
    if (scrollBtn) {
      let visible = false;
      window.addEventListener("scroll", () => {
        const show = window.scrollY > 400;
        if (show !== visible) {
          visible = show;
          scrollBtn.classList.toggle("visible", show);
        }
      }, { passive: true });

      scrollBtn.addEventListener("click", () => {
        window.scrollTo({ top: 0, behavior: "smooth" });
      });
    }
  }
})();
