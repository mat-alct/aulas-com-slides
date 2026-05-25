import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import pdfWorkerUrl from "pdfjs-dist/legacy/build/pdf.worker.mjs?url";
import { AlignmentType, BorderStyle, Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";
import { jsPDF } from "jspdf";
import "./styles.css";

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

const DB_NAME = "aula-slide-player";
const DB_VERSION = 1;
const LESSON_STORE = "lessons";
const MARKER_MERGE_WINDOW_SECONDS = 1.25;
const BACKUP_SCHEMA_VERSION = 1;
const SLIDE_CONVERT_EXT = new Set(["ppt", "pptx", "odp", "ppsx", "pps", "key"]);
const AUDIO_ACCEPT = "audio/mp4,audio/x-m4a,audio/m4a,audio/mpeg,audio/mp3,audio/wav,audio/x-wav,audio/ogg,audio/opus,.m4a,.mp3,.wav,.ogg,.opus,.aac,.flac";

const icons = {
  play: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5v14l11-7z"/></svg>`,
  pause: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 5h4v14H7zM13 5h4v14h-4z"/></svg>`,
  menu: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 12h16M4 18h16"/></svg>`,
  plus: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 5v14M5 12h14"/></svg>`,
  trash: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 6h18M8 6V4h8v2M6 6l1 15h10l1-15"/></svg>`,
  back: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M15 18l-6-6 6-6"/></svg>`,
  prev: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M15 18l-6-6 6-6"/></svg>`,
  next: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 18l6-6-6-6"/></svg>`,
  mark: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v18M5 8h14M7 8l2 13M17 8l-2 13"/></svg>`,
  check: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M20 6L9 17l-5-5"/></svg>`,
  fullscreen: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 3H3v5M16 3h5v5M8 21H3v-5M21 16v5h-5"/></svg>`,
  fullscreenExit: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 3v6H3M15 3v6h6M9 21v-6H3M21 15h-6v6"/></svg>`,
  edit: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 20h9M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>`,
  download: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12M7 10l5 5 5-5M5 21h14"/></svg>`,
  list: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/></svg>`,
  orderedList: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M10 6h11M10 12h11M10 18h11M4 6h1v4M4 10h2M4 14h2l-2 4h2"/></svg>`,
  clock: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 8v5l3 2"/><circle cx="12" cy="12" r="9"/></svg>`,
  quote: `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 17a4 4 0 0 1 4-4V7H5v6h4a4 4 0 0 1-4 4M17 17a4 4 0 0 1 4-4V7h-6v6h4a4 4 0 0 1-4 4"/></svg>`
};

let db;
let lessons = [];
let activeLesson = null;
let activePdf = null;
let activePdfTask = null;
let activeRenderTask = null;
let currentSlide = 1;
let currentAudioUrl = null;
let renderRequest = 0;
let appState = "home";
let isSidebarOpen = false;
let isAutoAdvance = false;
let audioEl = null;
let progressDrag = false;
let lessonChannel = null;
let editorLesson = null;
let editorSlide = 1;
let editorSaveTimer = null;
let latestPlayback = { slide: 1, time: 0 };
let draggedMarkerId = null;
let manualSlideOverride = false;
let audioContext = null;
let audioSource = null;
let gainNode = null;
let voiceHighPassNode = null;
let voiceLowPassNode = null;
let voiceCompressorNode = null;
let volumeBoost = 1;
let voiceCleanupEnabled = false;
let toolsAvailable = { soffice: false, ffmpeg: false };

const app = document.querySelector("#app");

init();

async function init() {
  db = await openDatabase();
  requestPersistentStorage();
  setupAutomaticBackup();
  try {
    toolsAvailable = (await window.aulaTools?.check()) || toolsAvailable;
  } catch (error) {
    console.warn("Falha ao verificar ferramentas externas", error);
  }
  lessons = await getAllLessons();
  const editorLessonId = new URLSearchParams(window.location.search).get("editor");
  if (editorLessonId) {
    await openEditor(editorLessonId);
    return;
  }
  renderHome();
}

function requestPersistentStorage() {
  if (!navigator.storage?.persist) return;
  navigator.storage.persist().catch(() => {});
}

function isCompactViewport() {
  return window.matchMedia("(max-width: 860px)").matches;
}

function setupAutomaticBackup() {
  window.aulaBackup?.onRequest(async () => {
    if (appState === "editor") {
      await saveEditorSlideNow();
    }
    if (activeLesson) {
      await persistActiveLesson();
    }
    const storedLessons = await getAllLessons();
    return createBackupPayload(storedLessons);
  });
}

async function createBackupPayload(rows) {
  return {
    app: "Aulas com Slides",
    schemaVersion: BACKUP_SCHEMA_VERSION,
    createdAt: new Date().toISOString(),
    lessonCount: rows.length,
    lessons: await Promise.all(rows.map(serializeLessonForBackup))
  };
}

async function serializeLessonForBackup(lesson) {
  const normalized = normalizeLesson(lesson);
  return {
    id: normalized.id,
    title: normalized.title,
    pdfName: normalized.pdfName,
    audioName: normalized.audioName,
    slideCount: normalized.slideCount,
    markers: normalized.markers,
    timeline: normalized.timeline,
    notes: normalized.notes.map((note) => ({
      ...note,
      html: sanitizeEditorHtml(note.html || "")
    })),
    createdAt: normalized.createdAt,
    updatedAt: normalized.updatedAt,
    files: {
      pdf: await blobToBackupFile(normalized.pdfBlob),
      audio: await blobToBackupFile(normalized.audioBlob)
    }
  };
}

async function blobToBackupFile(blob) {
  if (!blob) return null;
  const dataUrl = await blobToDataUrl(blob);
  return {
    type: blob.type || "application/octet-stream",
    size: blob.size,
    dataUrl
  };
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const database = request.result;
      if (!database.objectStoreNames.contains(LESSON_STORE)) {
        database.createObjectStore(LESSON_STORE, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function getStore(mode = "readonly") {
  return db.transaction(LESSON_STORE, mode).objectStore(LESSON_STORE);
}

function getAllLessons() {
  return new Promise((resolve, reject) => {
    const request = getStore().getAll();
    request.onsuccess = () => {
      const rows = request.result.map(normalizeLesson).sort((a, b) => b.updatedAt - a.updatedAt);
      resolve(rows);
    };
    request.onerror = () => reject(request.error);
  });
}

function saveLesson(lesson) {
  lesson.updatedAt = Date.now();
  return new Promise((resolve, reject) => {
    const request = getStore("readwrite").put(lesson);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

function deleteLesson(id) {
  return new Promise((resolve, reject) => {
    const request = getStore("readwrite").delete(id);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

function renderHome() {
  appState = "home";
  cleanupLesson();
  app.innerHTML = `
    <main class="home-shell">
      <section class="home-header">
        <div>
          <p class="eyebrow">Aulas sincronizadas</p>
          <h1>Escolha uma aula</h1>
        </div>
      </section>

      <section class="import-panel" aria-label="Cadastrar aula">
        <form id="lessonForm" class="lesson-form">
          <label>
            <span>Nome da aula</span>
            <input name="title" type="text" placeholder="Ex.: Aula 03 - Cardiologia" required />
          </label>
          <label>
            <span>Slides${toolsAvailable.soffice ? " (PDF, PPT, PPTX)" : " (PDF)"}</span>
            <input name="slides" type="file" accept="${toolsAvailable.soffice ? "application/pdf,.pdf,.pptx,.ppt,.odp,.ppsx,.pps,.key" : "application/pdf,.pdf"}" required />
          </label>
          <label>
            <span>Áudio${toolsAvailable.ffmpeg ? " - Parte 1" : ""}</span>
            <input name="audio1" type="file" accept="${AUDIO_ACCEPT}" required />
          </label>
          ${toolsAvailable.ffmpeg ? `
          <label>
            <span>Áudio - Parte 2 <small style="font-weight:500;opacity:0.7">(opcional)</small></span>
            <input name="audio2" type="file" accept="${AUDIO_ACCEPT}" />
          </label>` : ""}
          <button class="primary-button" type="submit">${icons.plus}<span>Cadastrar aula</span></button>
        </form>
        ${renderToolsHint()}
      </section>

      <section class="lesson-grid" aria-label="Aulas cadastradas">
        ${lessons.length ? lessons.map(renderLessonCard).join("") : `<div class="empty-state">Nenhuma aula cadastrada ainda.</div>`}
      </section>
    </main>
  `;

  document.querySelector("#lessonForm").addEventListener("submit", handleCreateLesson);
  document.querySelectorAll("[data-open-lesson]").forEach((button) => {
    button.addEventListener("click", () => openLesson(button.dataset.openLesson));
  });
  document.querySelectorAll("[data-delete-lesson]").forEach((button) => {
    button.addEventListener("click", async () => {
      const lesson = lessons.find((item) => item.id === button.dataset.deleteLesson);
      if (!lesson) return;
      const confirmed = window.confirm(`Excluir "${lesson.title}"?`);
      if (!confirmed) return;
      await deleteLesson(lesson.id);
      lessons = await getAllLessons();
      renderHome();
    });
  });
}

function renderLessonCard(lesson) {
  const markedCount = getSortedTimeline(lesson).length;
  const noteCount = lesson.notes.filter((note) => hasNoteContent(note.html)).length;
  return `
    <article class="lesson-card">
      <button class="lesson-open" type="button" data-open-lesson="${lesson.id}">
        <strong>${escapeHtml(lesson.title)}</strong>
        <span>${lesson.slideCount || "?"} slides · ${markedCount} marcações · ${noteCount} comentários</span>
      </button>
      <button class="icon-button danger" type="button" title="Excluir aula" aria-label="Excluir aula" data-delete-lesson="${lesson.id}">
        ${icons.trash}
      </button>
    </article>
  `;
}

async function handleCreateLesson(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const data = new FormData(form);
  const title = data.get("title").trim();
  const slidesFile = data.get("slides");
  const audioPart1 = data.get("audio1");
  const audioPart2 = data.get("audio2");
  const submit = form.querySelector("button[type='submit']");
  const status = submit.querySelector("span");
  submit.disabled = true;

  try {
    status.textContent = "Preparando slides...";
    const { blob: pdfBlob, name: pdfName } = await resolveSlidesAsPdf(slidesFile);

    status.textContent = audioPart2 && audioPart2.size ? "Mesclando áudios..." : "Lendo áudio...";
    const { blob: audioBlob, name: audioName } = await resolveAudio(audioPart1, audioPart2);

    status.textContent = "Lendo PDF...";
    const pdfArrayBuffer = await pdfBlob.arrayBuffer();
    const task = pdfjsLib.getDocument({ data: pdfArrayBuffer.slice(0) });
    const pdf = await task.promise;
    const slideCount = pdf.numPages;
    await pdf.destroy();

    const lesson = {
      id: crypto.randomUUID(),
      title,
      pdfName,
      audioName,
      pdfBlob,
      audioBlob,
      slideCount,
      markers: Array.from({ length: slideCount }, (_, index) => ({
        slide: index + 1,
        time: index === 0 ? 0 : null
      })),
      timeline: [
        {
          id: crypto.randomUUID(),
          slide: 1,
          time: 0
        }
      ],
      notes: Array.from({ length: slideCount }, (_, index) => ({
        slide: index + 1,
        html: ""
      })),
      createdAt: Date.now(),
      updatedAt: Date.now()
    };

    await saveLesson(lesson);
    lessons = await getAllLessons();
    form.reset();
    renderHome();
  } catch (error) {
    console.error(error);
    window.alert(error?.message || "Não consegui cadastrar essa aula. Verifique os arquivos e tente novamente.");
    submit.disabled = false;
    status.textContent = "Cadastrar aula";
  }
}

function getFileExtension(name) {
  const dot = String(name || "").lastIndexOf(".");
  return dot >= 0 ? name.slice(dot + 1).toLowerCase() : "";
}

async function resolveSlidesAsPdf(file) {
  if (!file) throw new Error("Arquivo de slides não informado.");
  const ext = getFileExtension(file.name);
  if (ext === "pdf" || file.type === "application/pdf") {
    return { blob: file, name: file.name };
  }
  if (!SLIDE_CONVERT_EXT.has(ext)) {
    throw new Error(`Formato de slides não suportado: .${ext || "?"}`);
  }
  if (!window.aulaTools?.convertSlidesToPdf) {
    throw new Error("Conversão de slides indisponível: o app precisa rodar no Electron.");
  }
  if (!toolsAvailable.soffice) {
    throw new Error("LibreOffice (soffice) não foi encontrado. Instale com: sudo apt install libreoffice");
  }
  const inputBuffer = await file.arrayBuffer();
  const pdfBuffer = await window.aulaTools.convertSlidesToPdf(inputBuffer, file.name);
  const pdfBlob = new Blob([pdfBuffer], { type: "application/pdf" });
  const baseName = file.name.replace(/\.[^.]+$/, "");
  return { blob: pdfBlob, name: `${baseName}.pdf` };
}

async function resolveAudio(part1, part2) {
  if (!part1) throw new Error("Áudio (parte 1) não informado.");
  if (!part2 || !part2.size) {
    return { blob: part1, name: part1.name };
  }
  if (!window.aulaTools?.mergeAudio) {
    throw new Error("Mesclagem de áudios indisponível: o app precisa rodar no Electron.");
  }
  if (!toolsAvailable.ffmpeg) {
    throw new Error("ffmpeg não foi encontrado. Instale com: sudo apt install ffmpeg");
  }
  const files = await Promise.all(
    [part1, part2].map(async (file) => ({
      buffer: await file.arrayBuffer(),
      name: file.name,
      mimeType: file.type || "application/octet-stream"
    }))
  );
  const merged = await window.aulaTools.mergeAudio(files);
  const blob = new Blob([merged.buffer], { type: merged.mimeType || "audio/mp4" });
  const baseName = part1.name.replace(/\.[^.]+$/, "");
  return { blob, name: `${baseName}-merged.${merged.extension || "m4a"}` };
}

function renderToolsHint() {
  if (!window.aulaTools) return "";
  const missing = [];
  if (!toolsAvailable.soffice) missing.push("LibreOffice (para PPT/PPTX)");
  if (!toolsAvailable.ffmpeg) missing.push("ffmpeg (para mesclar áudios)");
  if (!missing.length) return "";
  return `<p class="tools-hint">⚠ Não encontrei: ${missing.join(" · ")}. Recursos relacionados ficam desativados.</p>`;
}

async function openLesson(id) {
  const lesson = lessons.find((item) => item.id === id);
  if (!lesson) return;
  cleanupLesson();
  activeLesson = structuredCloneLesson(lesson);
  currentSlide = getSlideForTime(0);
  isSidebarOpen = false;
  isAutoAdvance = false;
  appState = "lesson";

  app.innerHTML = `
    <main class="lesson-shell">
      <header class="topbar">
        <button class="icon-button" id="backHome" type="button" title="Voltar" aria-label="Voltar">${icons.back}</button>
        <div class="lesson-title">
          <strong>${escapeHtml(activeLesson.title)}</strong>
          <span>${activeLesson.slideCount} slides</span>
        </div>
        <button class="icon-button" id="openEditor" type="button" title="Editor de comentários" aria-label="Editor de comentários">${icons.edit}</button>
        <button class="icon-button" id="toggleFullscreen" type="button" title="Tela cheia" aria-label="Tela cheia">${icons.fullscreen}</button>
        <button class="icon-button" id="toggleMenu" type="button" title="Marcação dos slides" aria-label="Marcação dos slides">${icons.menu}</button>
      </header>

      <section class="viewer-layout">
        <aside class="slide-sidebar" id="slideSidebar" aria-label="Navegação dos slides"></aside>
        <div class="sidebar-backdrop" id="sidebarBackdrop" aria-hidden="true"></div>
        <section class="slide-stage">
          <div class="slide-toolbar">
            <button class="icon-button" id="prevSlide" type="button" title="Slide anterior" aria-label="Slide anterior">${icons.prev}</button>
            <span id="slideCounter"></span>
            <button class="icon-button" id="seekCurrentSlide" type="button" title="Ir ao início deste slide no áudio" aria-label="Ir ao início deste slide no áudio">${icons.clock}</button>
            <button class="icon-button" id="nextSlide" type="button" title="Próximo slide" aria-label="Próximo slide">${icons.next}</button>
          </div>
          <div class="canvas-shell">
            <canvas id="pdfCanvas"></canvas>
            <div id="slideLoader" class="slide-loader">Carregando slide...</div>
          </div>
        </section>
      </section>

      <footer class="player-bar">
        <audio id="lessonAudio" preload="metadata"></audio>
        <div class="player-main">
          <button class="icon-button play-button" id="playPause" type="button" title="Reproduzir" aria-label="Reproduzir">${icons.play}</button>
          <span class="time-readout" id="currentTime">0:00</span>
          <div class="progress-wrap">
            <input id="audioProgress" type="range" min="0" max="0" value="0" step="0.01" aria-label="Progresso do áudio" />
            <div id="markerTrack" class="marker-track"></div>
          </div>
          <span class="time-readout" id="durationTime">0:00</span>
        </div>
        <div class="player-actions">
          <label class="toggle">
            <input id="autoAdvance" type="checkbox" />
            <span>Auto</span>
          </label>
          <label class="speed-control">
            <span id="speedLabel">1.00x</span>
            <input id="speedRange" type="range" min="0.75" max="2" value="1" step="0.01" aria-label="Velocidade do áudio" />
          </label>
          <label class="boost-control">
            <span id="boostLabel">100%</span>
            <input id="boostRange" type="range" min="1" max="3" value="1" step="0.05" aria-label="Amplificação do áudio" />
          </label>
          <label class="toggle">
            <input id="voiceCleanup" type="checkbox" />
            <span>Limpar voz</span>
          </label>
          <button class="mark-button accent" id="syncMark" type="button">${icons.mark}<span>Marcar slide atual</span></button>
        </div>
      </footer>
    </main>
  `;

  audioEl = document.querySelector("#lessonAudio");
  currentAudioUrl = URL.createObjectURL(activeLesson.audioBlob);
  audioEl.src = currentAudioUrl;
  setupAudioBoost();

  bindLessonEvents();
  setupLessonChannel();
  await loadPdf(activeLesson.pdfBlob);
  renderLessonUi();
}

function structuredCloneLesson(lesson) {
  return {
    ...lesson,
    markers: lesson.markers.map((marker) => ({ ...marker })),
    timeline: lesson.timeline.map((event) => ({ ...event })),
    notes: lesson.notes.map((note) => ({ ...note }))
  };
}

function normalizeLesson(lesson) {
  if (!lesson) return null;
  const slideCount = lesson.slideCount || lesson.markers?.length || 0;
  const timeline = normalizeTimeline(lesson, slideCount);
  const markers = deriveMarkersFromTimeline(slideCount, timeline);
  const notes = Array.from({ length: slideCount }, (_, index) => {
    const slide = index + 1;
    const existing = lesson.notes?.find((note) => note.slide === slide);
    return {
      slide,
      html: existing?.html || ""
    };
  });
  return { ...lesson, slideCount, markers, timeline, notes };
}

function normalizeTimeline(lesson, slideCount) {
  const source = Array.isArray(lesson.timeline) && lesson.timeline.length
    ? lesson.timeline
    : (lesson.markers || []).filter((marker) => marker.time !== null);

  const timeline = source
    .filter((event) => Number.isFinite(event.time) && event.slide >= 1 && event.slide <= slideCount)
    .map((event) => ({
      id: event.id || crypto.randomUUID(),
      slide: event.slide,
      time: Math.max(0, event.time)
    }))
    .sort((a, b) => a.time - b.time || a.slide - b.slide);

  if (!timeline.length) {
    timeline.push({ id: crypto.randomUUID(), slide: 1, time: 0 });
  }

  return timeline;
}

function deriveMarkersFromTimeline(slideCount, timeline) {
  return Array.from({ length: slideCount }, (_, index) => {
    const slide = index + 1;
    const firstEvent = timeline.find((event) => event.slide === slide);
    return {
      slide,
      time: firstEvent?.time ?? (slide === 1 ? 0 : null)
    };
  });
}

function getSortedTimeline(lesson = activeLesson) {
  return (lesson?.timeline || [])
    .filter((event) => Number.isFinite(event.time))
    .sort((a, b) => a.time - b.time || a.slide - b.slide);
}

async function loadPdf(pdfBlob) {
  const buffer = await pdfBlob.arrayBuffer();
  activePdfTask = pdfjsLib.getDocument({ data: buffer });
  activePdf = await activePdfTask.promise;
  await renderSlide(currentSlide);
}

function bindLessonEvents() {
  document.querySelector("#backHome").addEventListener("click", async () => {
    await persistActiveLesson();
    lessons = await getAllLessons();
    renderHome();
  });

  document.querySelector("#toggleMenu").addEventListener("click", () => {
    isSidebarOpen = !isSidebarOpen;
    renderLessonUi();
  });
  document.querySelector("#sidebarBackdrop")?.addEventListener("click", () => {
    isSidebarOpen = false;
    renderLessonUi();
  });
  document.querySelector("#openEditor").addEventListener("click", openEditorWindow);
  document.querySelector("#toggleFullscreen").addEventListener("click", toggleFullscreen);

  document.querySelector("#prevSlide").addEventListener("click", () => setSlide(currentSlide - 1, { manual: true }));
  document.querySelector("#nextSlide").addEventListener("click", () => setSlide(currentSlide + 1, { manual: true }));
  document.querySelector("#seekCurrentSlide").addEventListener("click", () => seekToSlideStart(currentSlide));

  document.querySelector("#playPause").addEventListener("click", () => {
    toggleAudioPlayback();
  });

  audioEl.addEventListener("play", updatePlayButton);
  audioEl.addEventListener("pause", updatePlayButton);
  audioEl.addEventListener("loadedmetadata", renderAudioState);
  audioEl.addEventListener("timeupdate", () => {
    if (manualSlideOverride && getSlideForTime(audioEl.currentTime) === currentSlide) {
      manualSlideOverride = false;
    }
    if (isAutoAdvance && !progressDrag && !manualSlideOverride) {
      const slide = getSlideForTime(audioEl.currentTime);
      if (slide !== currentSlide) {
        setSlide(slide, { keepAudio: true });
      }
    }
    renderAudioState();
    broadcastPlayback();
  });

  const progress = document.querySelector("#audioProgress");
  progress.addEventListener("input", () => {
    progressDrag = true;
    manualSlideOverride = false;
    audioEl.currentTime = Number(progress.value);
    renderAudioState();
  });
  progress.addEventListener("change", () => {
    progressDrag = false;
    if (isAutoAdvance) setSlide(getSlideForTime(audioEl.currentTime), { keepAudio: true });
  });

  document.querySelector("#autoAdvance").addEventListener("change", (event) => {
    isAutoAdvance = event.target.checked;
    if (isAutoAdvance) setSlide(getSlideForTime(audioEl.currentTime), { keepAudio: true });
  });

  document.querySelector("#speedRange").addEventListener("input", (event) => {
    const rate = Number(event.target.value);
    audioEl.playbackRate = rate;
    document.querySelector("#speedLabel").textContent = `${rate.toFixed(2)}x`;
  });

  document.querySelector("#boostRange").addEventListener("input", (event) => {
    setVolumeBoost(Number(event.target.value));
  });

  document.querySelector("#voiceCleanup").addEventListener("change", (event) => {
    setVoiceCleanup(event.target.checked);
  });

  document.querySelector("#syncMark").addEventListener("click", markCurrentSlideOnly);

  window.addEventListener("resize", handleResize);
  window.addEventListener("keydown", handleLessonKeydown);
  document.addEventListener("fullscreenchange", handleFullscreenChange);
}

function handleResize() {
  if (appState === "lesson" && activePdf) {
    renderSlide(currentSlide);
  }
}

function handleFullscreenChange() {
  updateFullscreenButton();
  handleResize();
}

function handleLessonKeydown(event) {
  if (appState !== "lesson" || isEditableTarget(event.target)) return;

  const key = event.key.toLowerCase();
  if (event.code === "Space") {
    event.preventDefault();
    toggleAudioPlayback();
    return;
  }

  if (event.ctrlKey && !event.altKey && !event.metaKey) {
    event.preventDefault();
    markCurrentSlideOnly();
    return;
  }

  if (key === "arrowleft") {
    event.preventDefault();
    seekAudio(-5);
    return;
  }

  if (key === "arrowright") {
    event.preventDefault();
    seekAudio(5);
    return;
  }

  if (key === "a") {
    event.preventDefault();
    setSlide(currentSlide - 1, { manual: true });
    return;
  }

  if (key === "d") {
    event.preventDefault();
    setSlide(currentSlide + 1, { manual: true });
  }
}

function toggleAudioPlayback() {
  if (!audioEl) return;
  audioContext?.resume();
  if (audioEl.paused) {
    audioEl.play();
  } else {
    audioEl.pause();
  }
}

function setupAudioBoost() {
  if (!audioEl || audioSource) return;
  const AudioContextClass = window.AudioContext || window.webkitAudioContext;
  if (!AudioContextClass) return;
  audioContext = new AudioContextClass();
  audioSource = audioContext.createMediaElementSource(audioEl);

  voiceHighPassNode = audioContext.createBiquadFilter();
  voiceHighPassNode.type = "highpass";
  voiceHighPassNode.frequency.value = 90;
  voiceHighPassNode.Q.value = 0.7;

  voiceLowPassNode = audioContext.createBiquadFilter();
  voiceLowPassNode.type = "lowpass";
  voiceLowPassNode.frequency.value = 7800;
  voiceLowPassNode.Q.value = 0.65;

  voiceCompressorNode = audioContext.createDynamicsCompressor();
  voiceCompressorNode.threshold.value = -34;
  voiceCompressorNode.knee.value = 24;
  voiceCompressorNode.ratio.value = 2.4;
  voiceCompressorNode.attack.value = 0.004;
  voiceCompressorNode.release.value = 0.18;

  gainNode = audioContext.createGain();
  gainNode.gain.value = volumeBoost;
  connectAudioGraph();
}

function setVolumeBoost(value) {
  volumeBoost = Math.max(1, Math.min(3, value || 1));
  if (gainNode) {
    gainNode.gain.value = volumeBoost;
  }
  const label = document.querySelector("#boostLabel");
  if (label) label.textContent = `${Math.round(volumeBoost * 100)}%`;
}

function setVoiceCleanup(enabled) {
  voiceCleanupEnabled = enabled;
  connectAudioGraph();
}

function connectAudioGraph() {
  if (!audioSource || !gainNode || !audioContext) return;
  [audioSource, voiceHighPassNode, voiceLowPassNode, voiceCompressorNode, gainNode].forEach((node) => {
    try {
      node?.disconnect();
    } catch {
      // The node may not be connected yet.
    }
  });

  if (voiceCleanupEnabled && voiceHighPassNode && voiceLowPassNode && voiceCompressorNode) {
    audioSource
      .connect(voiceHighPassNode)
      .connect(voiceLowPassNode)
      .connect(voiceCompressorNode)
      .connect(gainNode)
      .connect(audioContext.destination);
    return;
  }

  audioSource.connect(gainNode).connect(audioContext.destination);
}

function isEditableTarget(target) {
  const tag = target?.tagName?.toLowerCase();
  return target?.isContentEditable || tag === "input" || tag === "textarea" || tag === "select";
}

function seekAudio(deltaSeconds) {
  if (!audioEl) return;
  const duration = Number.isFinite(audioEl.duration) ? audioEl.duration : audioEl.currentTime + deltaSeconds;
  audioEl.currentTime = Math.min(duration, Math.max(0, audioEl.currentTime + deltaSeconds));
  if (isAutoAdvance) setSlide(getSlideForTime(audioEl.currentTime), { keepAudio: true });
  renderAudioState();
}

async function markCurrentSlideOnly() {
  if (!activeLesson || !audioEl) return;
  await markSlide(currentSlide, audioEl.currentTime);
}

async function toggleFullscreen() {
  const shell = document.querySelector(".lesson-shell");
  if (!shell) return;
  if (document.fullscreenElement) {
    await document.exitFullscreen();
  } else {
    await shell.requestFullscreen();
  }
}

function updateFullscreenButton() {
  const button = document.querySelector("#toggleFullscreen");
  if (!button) return;
  const isFullscreen = Boolean(document.fullscreenElement);
  button.innerHTML = isFullscreen ? icons.fullscreenExit : icons.fullscreen;
  button.title = isFullscreen ? "Sair da tela cheia" : "Tela cheia";
  button.setAttribute("aria-label", button.title);
}

async function persistActiveLesson() {
  if (!activeLesson) return;
  await saveLesson(activeLesson);
  broadcastLessonSaved();
}

async function markSlide(slide, time) {
  const safeTime = Math.max(0, time);
  const closeEvent = getSortedTimeline().find((event) => Math.abs(event.time - safeTime) <= MARKER_MERGE_WINDOW_SECONDS);

  if (closeEvent) {
    closeEvent.slide = slide;
    closeEvent.time = safeTime;
  } else {
    activeLesson.timeline.push({
      id: crypto.randomUUID(),
      slide,
      time: safeTime
    });
  }

  activeLesson.timeline = getSortedTimeline();
  activeLesson.markers = deriveMarkersFromTimeline(activeLesson.slideCount, activeLesson.timeline);
  manualSlideOverride = false;
  await persistActiveLesson();
  lessons = lessons.map((lesson) => (lesson.id === activeLesson.id ? structuredCloneLesson(activeLesson) : lesson));
  renderLessonUi();
}

function removeMarker(id) {
  const event = activeLesson.timeline.find((item) => item.id === id);
  if (!event) return;
  if (isInitialMarkerEvent(event)) return;
  activeLesson.timeline = activeLesson.timeline.filter((item) => item.id !== id);
  activeLesson.markers = deriveMarkersFromTimeline(activeLesson.slideCount, activeLesson.timeline);
  persistActiveLesson();
  renderLessonUi();
}

async function updateMarkerTime(id, time, options = {}) {
  const event = activeLesson?.timeline.find((item) => item.id === id);
  if (!event || isInitialMarkerEvent(event)) return;
  const duration = getAudioDuration();
  const safeTime = Math.max(0, Math.min(duration || time, time));
  event.time = Number(safeTime.toFixed(2));
  activeLesson.timeline = getSortedTimeline();
  activeLesson.markers = deriveMarkersFromTimeline(activeLesson.slideCount, activeLesson.timeline);

  if (options.seek && audioEl) {
    audioEl.currentTime = event.time;
    const rendered = await renderSlide(event.slide);
    if (rendered) {
      currentSlide = event.slide;
    }
  }

  if (options.persist !== false) {
    await persistActiveLesson();
    lessons = lessons.map((lesson) => (lesson.id === activeLesson.id ? structuredCloneLesson(activeLesson) : lesson));
  }

  if (options.render !== false) {
    renderLessonUi();
    broadcastPlayback();
  }
}

async function setSlide(slide, options = {}) {
  const bounded = Math.max(1, Math.min(activeLesson.slideCount, slide));
  if (bounded === currentSlide) return;
  const previousSlide = currentSlide;
  const rendered = await renderSlide(bounded);
  if (!rendered) {
    currentSlide = previousSlide;
    renderLessonUi();
    return;
  }
  currentSlide = bounded;
  if (options.manual) {
    manualSlideOverride = getSlideForTime(audioEl.currentTime || 0) !== currentSlide;
  }
  renderLessonUi();
  broadcastPlayback();
}

async function setTimelineEvent(id) {
  const event = activeLesson.timeline.find((item) => item.id === id);
  if (!event) return;
  audioEl.currentTime = event.time;
  const rendered = await renderSlide(event.slide);
  if (!rendered) return;
  currentSlide = event.slide;
  manualSlideOverride = false;
  renderLessonUi();
  broadcastPlayback();
}

function setupLessonChannel() {
  if (!activeLesson || !("BroadcastChannel" in window)) return;
  lessonChannel?.close();
  lessonChannel = new BroadcastChannel(`lesson-${activeLesson.id}`);
  lessonChannel.onmessage = async (event) => {
    if (event.data?.type === "notes-saved") {
      const updated = normalizeLesson(event.data.lesson);
      if (activeLesson?.id === updated.id) {
        activeLesson.notes = updated.notes;
        lessons = lessons.map((lesson) => (lesson.id === updated.id ? structuredCloneLesson(updated) : lesson));
      }
    }
    if (event.data?.type === "request-playback") {
      broadcastPlayback();
    }
  };
  broadcastPlayback();
}

function broadcastPlayback() {
  if (!lessonChannel || !activeLesson || !audioEl) return;
  lessonChannel.postMessage({
    type: "playback",
    lessonId: activeLesson.id,
    slide: currentSlide,
    time: audioEl.currentTime || 0
  });
}

function broadcastLessonSaved() {
  if (!lessonChannel || !activeLesson) return;
  lessonChannel.postMessage({
    type: "lesson-saved",
    lesson: activeLesson
  });
}

function openEditorWindow() {
  if (!activeLesson) return;
  const url = new URL(window.location.href);
  url.search = `?editor=${encodeURIComponent(activeLesson.id)}`;

  if (isCompactViewport()) {
    persistActiveLesson().finally(() => {
      window.location.assign(url.toString());
    });
    return;
  }

  const editor = window.open(url.toString(), `editor-${activeLesson.id}`, "width=1100,height=760");
  if (editor) {
    editor.focus();
  } else {
    window.location.assign(url.toString());
  }
}

async function renderSlide(slideNumber) {
  if (!activePdf) return false;
  const canvas = document.querySelector("#pdfCanvas");
  const loader = document.querySelector("#slideLoader");
  if (!canvas) return false;

  const requestId = ++renderRequest;
  loader.classList.add("visible");
  loader.textContent = "Carregando slide...";
  const page = await activePdf.getPage(slideNumber);
  if (requestId !== renderRequest) return false;

  const shell = document.querySelector(".canvas-shell");
  const baseViewport = page.getViewport({ scale: 1 });
  const availableWidth = Math.max(320, shell.clientWidth - 16);
  const availableHeight = Math.max(240, shell.clientHeight - 16);
  const scale = Math.min(availableWidth / baseViewport.width, availableHeight / baseViewport.height) * window.devicePixelRatio;
  const viewport = page.getViewport({ scale });

  const renderCanvas = document.createElement("canvas");
  renderCanvas.width = Math.floor(viewport.width);
  renderCanvas.height = Math.floor(viewport.height);

  const context = renderCanvas.getContext("2d", { alpha: false });
  const renderTask = page.render({ canvasContext: context, viewport });
  activeRenderTask = renderTask;

  try {
    await renderTask.promise;
    if (requestId === renderRequest) {
      canvas.width = renderCanvas.width;
      canvas.height = renderCanvas.height;
      canvas.style.width = `${Math.floor(viewport.width / window.devicePixelRatio)}px`;
      canvas.style.height = `${Math.floor(viewport.height / window.devicePixelRatio)}px`;
      canvas.getContext("2d", { alpha: false }).drawImage(renderCanvas, 0, 0);
      loader.classList.remove("visible");
      return true;
    }
  } catch (error) {
    console.error(error);
    if (requestId === renderRequest) {
      loader.textContent = "Erro ao carregar slide";
    }
  } finally {
    if (activeRenderTask === renderTask) {
      activeRenderTask = null;
    }
  }
  return false;
}

function renderLessonUi() {
  if (!activeLesson) return;
  const sidebar = document.querySelector("#slideSidebar");
  const shell = document.querySelector(".lesson-shell");
  const counter = document.querySelector("#slideCounter");
  if (!sidebar || !shell || !counter) return;

  shell.classList.toggle("sidebar-open", isSidebarOpen);
  counter.textContent = `Slide ${currentSlide} de ${activeLesson.slideCount}`;
  document.querySelector("#prevSlide").disabled = currentSlide === 1;
  document.querySelector("#nextSlide").disabled = currentSlide === activeLesson.slideCount;
  const currentSlideMarker = getFirstTimelineEventForSlide(activeLesson, currentSlide);
  const seekCurrentSlideButton = document.querySelector("#seekCurrentSlide");
  seekCurrentSlideButton.disabled = !currentSlideMarker;
  seekCurrentSlideButton.title = currentSlideMarker
    ? `Ir para ${formatTime(currentSlideMarker.time)} no áudio`
    : "Este slide ainda não tem marcação no áudio";
  seekCurrentSlideButton.setAttribute("aria-label", seekCurrentSlideButton.title);
  updateFullscreenButton();

  sidebar.innerHTML = `
    <div class="sidebar-header">
      <strong>Linha do tempo</strong>
      <span>${getSortedTimeline().length} marcas</span>
    </div>
    <div class="slide-list">
      ${getSortedTimeline().map(renderSlideRow).join("")}
    </div>
    <div class="sidebar-header compact">
      <strong>Slides</strong>
      <span>${activeLesson.slideCount}</span>
    </div>
    <div class="all-slide-list">
      ${Array.from({ length: activeLesson.slideCount }, (_, index) => renderSlideJumpRow(index + 1)).join("")}
    </div>
  `;

  sidebar.querySelectorAll("[data-event-id]").forEach((button) => {
    button.addEventListener("click", () => setTimelineEvent(button.dataset.eventId));
  });
  sidebar.querySelectorAll("[data-marker-time]").forEach((input) => {
    input.addEventListener("click", (event) => event.stopPropagation());
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        input.blur();
      }
      if (event.key === "Escape") {
        const marker = activeLesson.timeline.find((item) => item.id === input.dataset.markerTime);
        input.value = formatTimeInput(marker?.time || 0);
        input.blur();
      }
    });
    input.addEventListener("change", async () => {
      const seconds = parseTimeInput(input.value);
      if (seconds === null) {
        const marker = activeLesson.timeline.find((item) => item.id === input.dataset.markerTime);
        input.value = formatTimeInput(marker?.time || 0);
        return;
      }
      await updateMarkerTime(input.dataset.markerTime, seconds, { seek: true });
    });
  });
  sidebar.querySelectorAll("[data-remove-marker]").forEach((button) => {
    button.addEventListener("click", (event) => {
      event.stopPropagation();
      removeMarker(button.dataset.removeMarker);
    });
  });
  sidebar.querySelectorAll("[data-show-slide]").forEach((button) => {
    button.addEventListener("click", () => setSlide(Number(button.dataset.showSlide), { keepAudio: true, manual: true }));
  });
  sidebar.querySelectorAll("[data-seek-slide]").forEach((button) => {
    button.addEventListener("click", () => seekToSlideStart(Number(button.dataset.seekSlide)));
  });

  renderAudioState();
}

function renderSlideRow(marker) {
  const activeEvent = getTimelineEventForTime(audioEl?.currentTime || 0);
  const active = marker.id === activeEvent?.id ? "active" : "";
  const time = formatMarkerInterval(marker);
  const locked = isInitialMarkerEvent(marker);
  return `
    <div class="slide-row ${active}">
      <button class="slide-row-main" type="button" data-event-id="${marker.id}">
        <span>Slide ${marker.slide}</span>
        <small>${time}</small>
      </button>
      <input
        class="time-input"
        type="text"
        value="${formatTimeInput(marker.time)}"
        data-marker-time="${marker.id}"
        aria-label="Tempo do slide ${marker.slide}"
        ${locked ? "disabled" : ""}
      />
      ${
        locked
          ? ""
          : `<button class="row-delete" type="button" data-remove-marker="${marker.id}" title="Remover marcação" aria-label="Remover marcação">${icons.trash}</button>`
      }
    </div>
  `;
}

function renderSlideJumpRow(slide) {
  const marker = getFirstTimelineEventForSlide(activeLesson, slide);
  const active = slide === currentSlide ? "active" : "";
  return `
    <div class="slide-jump-row ${active}">
      <button class="slide-jump-main" type="button" data-show-slide="${slide}">
        <span>Slide ${slide}</span>
        <small>${marker ? formatTime(marker.time) : "sem marcação"}</small>
      </button>
      <button
        class="slide-audio-button"
        type="button"
        data-seek-slide="${slide}"
        title="Ir ao início no áudio"
        aria-label="Ir ao início do slide ${slide} no áudio"
        ${marker ? "" : "disabled"}
      >
        ${icons.clock}
      </button>
    </div>
  `;
}

async function seekToSlideStart(slide) {
  const marker = getFirstTimelineEventForSlide(activeLesson, slide);
  if (!marker || !audioEl) return;
  audioEl.currentTime = marker.time;
  const rendered = await renderSlide(slide);
  if (!rendered) return;
  currentSlide = slide;
  manualSlideOverride = false;
  renderLessonUi();
  broadcastPlayback();
}

function formatMarkerInterval(marker) {
  if (marker.time === null) return "sem marcação";
  const timeline = getSortedTimeline();
  const index = timeline.findIndex((item) => item.id === marker.id);
  const nextMarker = timeline[index + 1];
  if (nextMarker) return `${formatTime(marker.time)} -> ${formatTime(nextMarker.time)}`;
  return `desde ${formatTime(marker.time)}`;
}

function renderAudioState() {
  if (!audioEl) return;
  const duration = Number.isFinite(audioEl.duration) ? audioEl.duration : 0;
  const progress = document.querySelector("#audioProgress");
  const current = document.querySelector("#currentTime");
  const durationText = document.querySelector("#durationTime");
  const markerTrack = document.querySelector("#markerTrack");
  if (!progress || !current || !durationText || !markerTrack) return;

  progress.max = duration;
  progress.value = audioEl.currentTime || 0;
  current.textContent = formatTime(audioEl.currentTime || 0);
  durationText.textContent = formatTime(duration);

  if (!draggedMarkerId) {
    markerTrack.innerHTML = getSortedTimeline()
      .filter((marker) => duration > 0)
      .map((marker) => {
        const left = Math.min(100, Math.max(0, (marker.time / duration) * 100));
        const active = marker.id === getTimelineEventForTime(audioEl.currentTime || 0)?.id ? "active" : "";
        const locked = isInitialMarkerEvent(marker) ? " locked" : "";
        return `<button class="timeline-marker ${active}${locked}" style="left:${left}%" type="button" title="Slide ${marker.slide} · ${formatTime(marker.time)}" data-marker-id="${marker.id}"></button>`;
      })
      .join("");

    markerTrack.querySelectorAll("[data-marker-id]").forEach((button) => {
      button.addEventListener("click", () => setTimelineEvent(button.dataset.markerId));
      button.addEventListener("pointerdown", startMarkerDrag);
    });
  }
}

function startMarkerDrag(event) {
  const markerId = event.currentTarget.dataset.markerId;
  const marker = activeLesson?.timeline.find((item) => item.id === markerId);
  if (!marker || isInitialMarkerEvent(marker)) return;
  event.preventDefault();
  event.stopPropagation();
  draggedMarkerId = markerId;
  window.addEventListener("pointermove", handleMarkerDragMove);
  window.addEventListener("pointerup", handleMarkerDragEnd, { once: true });
  previewMarkerDrag(event.clientX);
}

function handleMarkerDragMove(event) {
  if (!draggedMarkerId) return;
  event.preventDefault();
  previewMarkerDrag(event.clientX);
}

async function handleMarkerDragEnd(event) {
  if (!draggedMarkerId) return;
  const markerId = draggedMarkerId;
  const seconds = getTimeFromPointer(event.clientX);
  draggedMarkerId = null;
  window.removeEventListener("pointermove", handleMarkerDragMove);
  await updateMarkerTime(markerId, seconds, { seek: true });
}

function previewMarkerDrag(clientX) {
  const marker = activeLesson?.timeline.find((item) => item.id === draggedMarkerId);
  if (!marker) return;
  const seconds = getTimeFromPointer(clientX);
  marker.time = Number(seconds.toFixed(2));
  activeLesson.timeline = getSortedTimeline();
  activeLesson.markers = deriveMarkersFromTimeline(activeLesson.slideCount, activeLesson.timeline);
  audioEl.currentTime = marker.time;
  currentSlide = marker.slide;
  renderAudioState();
  positionDraggedMarker(marker.time);
  broadcastPlayback();
}

function positionDraggedMarker(time) {
  const duration = getAudioDuration();
  const marker = Array.from(document.querySelectorAll("[data-marker-id]")).find((item) => item.dataset.markerId === draggedMarkerId);
  if (!marker || !duration) return;
  const left = Math.min(100, Math.max(0, (time / duration) * 100));
  marker.style.left = `${left}%`;
  marker.title = `Slide ${activeLesson.timeline.find((item) => item.id === draggedMarkerId)?.slide || ""} · ${formatTime(time)}`;
}

function getTimeFromPointer(clientX) {
  const progress = document.querySelector("#audioProgress");
  const duration = getAudioDuration();
  if (!progress || !duration) return 0;
  const rect = progress.getBoundingClientRect();
  const ratio = Math.min(1, Math.max(0, (clientX - rect.left) / rect.width));
  return ratio * duration;
}

function updatePlayButton() {
  const button = document.querySelector("#playPause");
  if (!button) return;
  button.innerHTML = audioEl.paused ? icons.play : icons.pause;
  button.title = audioEl.paused ? "Reproduzir" : "Pausar";
  button.setAttribute("aria-label", button.title);
}

function getSlideForTime(time) {
  if (!activeLesson) return 1;
  return getTimelineEventForTime(time)?.slide || 1;
}

function getTimelineEventForTime(time, lesson = activeLesson) {
  const marked = getSortedTimeline(lesson).filter((marker) => marker.time <= time + 0.05);
  return marked.at(-1) || getSortedTimeline(lesson)[0] || null;
}

function getFirstTimelineEventForSlide(lesson, slide) {
  return getSortedTimeline(lesson).find((event) => event.slide === slide) || null;
}

function getTimelineEventsForSlide(lesson, slide) {
  return getSortedTimeline(lesson).filter((event) => event.slide === slide);
}

function formatSlideTimes(lesson, slide) {
  const events = getTimelineEventsForSlide(lesson, slide);
  if (!events.length) return "sem marcação";
  return events.map((event) => formatTime(event.time)).join(", ");
}

function formatTime(seconds) {
  if (!Number.isFinite(seconds)) return "0:00";
  const safe = Math.max(0, Math.floor(seconds));
  const minutes = Math.floor(safe / 60);
  const remainder = String(safe % 60).padStart(2, "0");
  return `${minutes}:${remainder}`;
}

function formatTimeInput(seconds) {
  if (!Number.isFinite(seconds)) return "0:00.00";
  const safe = Math.max(0, seconds);
  const minutes = Math.floor(safe / 60);
  const remainder = (safe % 60).toFixed(2).padStart(5, "0");
  return `${minutes}:${remainder}`;
}

function parseTimeInput(value) {
  const raw = String(value || "").trim().replace(",", ".");
  if (!raw) return null;
  if (!raw.includes(":")) {
    const seconds = Number(raw);
    return Number.isFinite(seconds) && seconds >= 0 ? seconds : null;
  }

  const parts = raw.split(":").map((part) => Number(part));
  if (parts.some((part) => !Number.isFinite(part) || part < 0) || parts.length > 3) return null;
  return parts.reduce((total, part) => total * 60 + part, 0);
}

function getAudioDuration() {
  return Number.isFinite(audioEl?.duration) ? audioEl.duration : 0;
}

function isInitialMarkerEvent(event) {
  return event?.slide === 1 && event?.time === 0 && getSortedTimeline()[0]?.id === event.id;
}

function cleanupLesson() {
  window.removeEventListener("resize", handleResize);
  window.removeEventListener("keydown", handleLessonKeydown);
  document.removeEventListener("fullscreenchange", handleFullscreenChange);
  lessonChannel?.close();
  lessonChannel = null;
  if (currentAudioUrl) {
    URL.revokeObjectURL(currentAudioUrl);
    currentAudioUrl = null;
  }
  if (activeRenderTask) {
    activeRenderTask.cancel();
    activeRenderTask = null;
  }
  if (activePdf) {
    activePdf.destroy();
  }
  if (!activePdf && activePdfTask) {
    activePdfTask.destroy();
  }
  activeLesson = null;
  activePdf = null;
  activePdfTask = null;
  activeRenderTask = null;
  draggedMarkerId = null;
  manualSlideOverride = false;
  gainNode = null;
  voiceHighPassNode = null;
  voiceLowPassNode = null;
  voiceCompressorNode = null;
  audioSource = null;
  audioContext?.close();
  audioContext = null;
  volumeBoost = 1;
  voiceCleanupEnabled = false;
  audioEl = null;
}

async function openEditor(id) {
  appState = "editor";
  const storedLessons = await getAllLessons();
  editorLesson = normalizeLesson(storedLessons.find((lesson) => lesson.id === id));
  if (!editorLesson) {
    app.innerHTML = `<main class="editor-missing">Aula não encontrada.</main>`;
    return;
  }

  latestPlayback = { slide: 1, time: 0 };
  editorSlide = 1;

  app.innerHTML = `
    <main class="editor-shell">
      <button class="editor-mobile-back" id="editorMobileBack" type="button" aria-label="Voltar à aula">${icons.back}<span>Aula</span></button>
      <button class="editor-mobile-toggle" id="editorMobileToggle" type="button" aria-label="Slides">${icons.menu}</button>
      <aside class="editor-sidebar" id="editorSidebar">
        <div class="editor-title">
          <strong>${escapeHtml(editorLesson.title)}</strong>
          <span id="editorSyncStatus">Slide 1 · 0:00</span>
        </div>
        <div class="editor-slide-list" id="editorSlideList"></div>
      </aside>
      <section class="editor-workspace">
        <header class="editor-toolbar">
          <select id="formatBlock" title="Estilo do texto" aria-label="Estilo do texto">
            <option value="p">Texto</option>
            <option value="h1">Título 1</option>
            <option value="h2">Título 2</option>
            <option value="h3">Título 3</option>
          </select>
          <select id="fontSize" title="Tamanho do texto" aria-label="Tamanho do texto">
            <option value="3">12 pt</option>
            <option value="4">14 pt</option>
            <option value="5">17 pt</option>
            <option value="6">22 pt</option>
          </select>
          <button class="editor-tool" data-command="bold" type="button" title="Negrito" aria-label="Negrito"><strong>B</strong></button>
          <button class="editor-tool" data-command="italic" type="button" title="Itálico" aria-label="Itálico"><em>I</em></button>
          <button class="editor-tool" data-command="insertUnorderedList" type="button" title="Lista" aria-label="Lista">${icons.list}</button>
          <button class="editor-tool" data-command="insertOrderedList" type="button" title="Lista numerada" aria-label="Lista numerada">${icons.orderedList}</button>
          <button class="editor-tool" id="quoteBlock" type="button" title="Bloco de observação" aria-label="Bloco de observação">${icons.quote}</button>
          <button class="editor-tool" id="insertMoment" type="button" title="Inserir momento do áudio" aria-label="Inserir momento do áudio">${icons.clock}</button>
          <span class="editor-spacer"></span>
          <button class="editor-export" id="exportDocx" type="button">${icons.download}<span>DOCX</span></button>
          <button class="editor-export" id="exportPdf" type="button">${icons.download}<span>PDF</span></button>
        </header>
        <div class="editor-page-wrap">
          <article class="editor-page">
            <div class="editor-page-heading">
              <div class="editor-heading-main">
                <span id="editorSlideHeading">Slide 1</span>
                <small id="editorSlideTime"></small>
              </div>
              <div class="editor-slide-nav">
                <button class="editor-tool" id="editorPrevSlide" type="button" title="Slide anterior" aria-label="Slide anterior">${icons.prev}</button>
                <button class="editor-tool" id="editorCurrentAudioSlide" type="button" title="Ir para o slide do áudio" aria-label="Ir para o slide do áudio">${icons.clock}</button>
                <button class="editor-tool" id="editorNextSlide" type="button" title="Próximo slide" aria-label="Próximo slide">${icons.next}</button>
              </div>
            </div>
            <div id="richEditor" class="rich-editor" contenteditable="true" spellcheck="true"></div>
          </article>
        </div>
      </section>
    </main>
  `;

  bindEditorEvents();
  setupEditorChannel();
  renderEditorSlideList();
  loadEditorSlide(1);
}

function bindEditorEvents() {
  const editor = document.querySelector("#richEditor");
  editor.addEventListener("input", scheduleEditorSave);
  editor.addEventListener("paste", handleEditorPaste);

  document.querySelectorAll("[data-command]").forEach((button) => {
    button.addEventListener("click", () => runEditorCommand(button.dataset.command));
  });

  document.querySelector("#formatBlock").addEventListener("change", (event) => {
    document.execCommand("formatBlock", false, event.target.value);
    document.querySelector("#richEditor").focus();
    scheduleEditorSave();
  });

  document.querySelector("#fontSize").addEventListener("change", (event) => {
    document.execCommand("fontSize", false, event.target.value);
    document.querySelector("#richEditor").focus();
    scheduleEditorSave();
  });

  document.querySelector("#insertMoment").addEventListener("click", insertAudioMoment);
  document.querySelector("#quoteBlock").addEventListener("click", formatQuoteBlock);
  document.querySelector("#exportDocx").addEventListener("click", exportDocx);
  document.querySelector("#exportPdf").addEventListener("click", exportPdf);
  document.querySelector("#editorPrevSlide").addEventListener("click", () => loadEditorSlide(editorSlide - 1));
  document.querySelector("#editorNextSlide").addEventListener("click", () => loadEditorSlide(editorSlide + 1));
  document.querySelector("#editorCurrentAudioSlide").addEventListener("click", () => loadEditorSlide(latestPlayback.slide || editorSlide));

  document.querySelector("#editorMobileBack").addEventListener("click", async () => {
    await saveEditorSlideNow();
    if (window.history.length > 1) {
      window.history.back();
    } else {
      window.location.assign(window.location.pathname);
    }
  });

  document.querySelector("#editorMobileToggle").addEventListener("click", () => {
    document.querySelector(".editor-shell")?.classList.toggle("slides-open");
  });

  document.querySelector("#editorSlideList").addEventListener("click", () => {
    document.querySelector(".editor-shell")?.classList.remove("slides-open");
  });

  window.addEventListener("beforeunload", () => {
    saveEditorSlideNow();
    lessonChannel?.close();
  });
}

function setupEditorChannel() {
  if (!("BroadcastChannel" in window)) return;
  lessonChannel?.close();
  lessonChannel = new BroadcastChannel(`lesson-${editorLesson.id}`);
  lessonChannel.onmessage = (event) => {
    if (event.data?.type === "playback") {
      latestPlayback = { slide: event.data.slide, time: event.data.time };
      document.querySelector("#editorSyncStatus").textContent = `Slide ${latestPlayback.slide} · ${formatTime(latestPlayback.time)}`;
    }
    if (event.data?.type === "lesson-saved") {
      const updated = normalizeLesson(event.data.lesson);
      editorLesson.markers = updated.markers;
      editorLesson.timeline = updated.timeline;
      renderEditorSlideList();
      updateEditorHeading();
    }
  };
  lessonChannel.postMessage({ type: "request-playback" });
}

function runEditorCommand(command) {
  document.execCommand(command, false, null);
  document.querySelector("#richEditor").focus();
  scheduleEditorSave();
}

function insertAudioMoment() {
  const slideFromAudio = latestPlayback.slide || editorSlide;
  if (slideFromAudio !== editorSlide) {
    loadEditorSlide(slideFromAudio);
  }
  document.querySelector("#richEditor").focus();
  const label = `[${formatTime(latestPlayback.time || 0)}] `;
  document.execCommand("insertHTML", false, `<p><strong>${escapeHtml(label)}</strong></p>`);
  scheduleEditorSave();
}

function formatQuoteBlock() {
  document.execCommand("formatBlock", false, "blockquote");
  document.querySelector("#richEditor").focus();
  scheduleEditorSave();
}

function handleEditorPaste(event) {
  event.preventDefault();
  const html = event.clipboardData?.getData("text/html");
  const text = event.clipboardData?.getData("text/plain") || "";
  const content = html
    ? sanitizeEditorHtml(html)
    : text
        .split(/\n{2,}/)
        .map((block) => `<p>${escapeHtml(block).replace(/\n/g, "<br>")}</p>`)
        .join("");
  document.execCommand("insertHTML", false, content);
  scheduleEditorSave();
}

function renderEditorSlideList() {
  const list = document.querySelector("#editorSlideList");
  if (!list || !editorLesson) return;
  list.innerHTML = editorLesson.notes
    .map((note) => {
      const active = note.slide === editorSlide ? "active" : "";
      const hasNotes = hasNoteContent(note.html) ? "comentado" : "vazio";
      const time = formatSlideTimes(editorLesson, note.slide);
      return `
        <button class="editor-slide-row ${active}" type="button" data-editor-slide="${note.slide}">
          <span>Slide ${note.slide}</span>
          <small>${time} · ${hasNotes}</small>
        </button>
      `;
    })
    .join("");
  list.querySelectorAll("[data-editor-slide]").forEach((button) => {
    button.addEventListener("click", () => loadEditorSlide(Number(button.dataset.editorSlide)));
  });
}

function loadEditorSlide(slide) {
  saveEditorSlideNow();
  editorSlide = Math.max(1, Math.min(editorLesson.slideCount, slide));
  const note = getEditorNote(editorSlide);
  document.querySelector("#richEditor").innerHTML = sanitizeEditorHtml(note.html || "");
  updateEditorHeading();
  renderEditorSlideList();
}

function updateEditorHeading() {
  document.querySelector("#editorSlideHeading").textContent = `Slide ${editorSlide}`;
  document.querySelector("#editorSlideTime").textContent = formatSlideTimes(editorLesson, editorSlide);
  document.querySelector("#editorPrevSlide").disabled = editorSlide === 1;
  document.querySelector("#editorNextSlide").disabled = editorSlide === editorLesson.slideCount;
}

function scheduleEditorSave() {
  window.clearTimeout(editorSaveTimer);
  editorSaveTimer = window.setTimeout(saveEditorSlideNow, 450);
}

async function saveEditorSlideNow() {
  if (!editorLesson || appState !== "editor") return;
  window.clearTimeout(editorSaveTimer);
  const editor = document.querySelector("#richEditor");
  if (!editor) return;
  getEditorNote(editorSlide).html = sanitizeEditorHtml(editor.innerHTML);
  editorLesson.updatedAt = Date.now();
  await saveLesson(editorLesson);
  renderEditorSlideList();
  lessonChannel?.postMessage({ type: "notes-saved", lesson: editorLesson });
}

function getEditorNote(slide) {
  let note = editorLesson.notes.find((item) => item.slide === slide);
  if (!note) {
    note = { slide, html: "" };
    editorLesson.notes.push(note);
    editorLesson.notes.sort((a, b) => a.slide - b.slide);
  }
  return note;
}

async function exportDocx() {
  await saveEditorSlideNow();
  const filledNotes = getFilledNotes();
  if (!filledNotes.length) {
    window.alert("Nenhum slide com texto para exportar.");
    return;
  }
  const children = [
    new Paragraph({
      text: editorLesson.title,
      heading: HeadingLevel.TITLE
    })
  ];

  filledNotes.forEach((note) => {
    const times = formatSlideTimes(editorLesson, note.slide);
    children.push(
      new Paragraph({
        text: `Slide ${note.slide}${times !== "sem marcação" ? ` - ${times}` : ""}`,
        heading: HeadingLevel.HEADING_1
      }),
      ...htmlToDocxParagraphs(note.html)
    );
  });

  const documentFile = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Inter",
            size: 24,
            color: "222222"
          },
          paragraph: {
            alignment: AlignmentType.LEFT,
            spacing: { line: 360, after: 160 }
          }
        }
      },
      paragraphStyles: [
        {
          id: "Title",
          name: "Title",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { font: "Inter", size: 44, bold: true, color: "111111" },
          paragraph: { spacing: { before: 0, after: 240, line: 276 }, alignment: AlignmentType.LEFT }
        },
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { font: "Inter", size: 44, bold: true, color: "111111" },
          paragraph: { spacing: { before: 480, after: 240, line: 276 }, alignment: AlignmentType.LEFT }
        },
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { font: "Inter", size: 34, bold: true, color: "222222" },
          paragraph: { spacing: { before: 400, after: 160, line: 276 }, alignment: AlignmentType.LEFT }
        },
        {
          id: "Heading3",
          name: "Heading 3",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { font: "Inter", size: 28, bold: true, color: "333333" },
          paragraph: { spacing: { before: 280, after: 120, line: 276 }, alignment: AlignmentType.LEFT }
        },
        {
          id: "Observation",
          name: "Observation",
          basedOn: "Normal",
          run: { font: "Inter", size: 22, color: "444444" },
          paragraph: {
            spacing: { before: 200, after: 200, line: 330 },
            indent: { left: 567 },
            border: {
              left: { color: "DDDDDD", space: 8, style: BorderStyle.SINGLE, size: 6 }
            }
          }
        }
      ]
    },
    numbering: {
      config: [
        {
          reference: "ordered-list",
          levels: [{ level: 0, format: "decimal", text: "%1.", alignment: "left" }]
        }
      ]
    },
    sections: [
      {
        properties: {
          page: {
            margin: { top: 907, right: 1021, bottom: 907, left: 1021 }
          }
        },
        children
      }
    ]
  });
  const blob = await Packer.toBlob(documentFile);
  downloadBlob(blob, `${safeFileName(editorLesson.title)}-comentarios.docx`);
}

async function exportPdf() {
  await saveEditorSlideNow();
  const filledNotes = getFilledNotes();
  if (!filledNotes.length) {
    window.alert("Nenhum slide com texto para exportar.");
    return;
  }
  const pdf = new jsPDF("p", "pt", "a4");
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const margin = 54;
  const maxWidth = pageWidth - margin * 2;
  let cursorY = margin;

  const writeBlock = (text, options = {}) => {
    const cleanText = String(text || "").replace(/\s+/g, " ").trim();
    if (!cleanText) return;
    const fontSize = options.fontSize || 12;
    const lineHeight = options.lineHeight || fontSize * 1.45;
    const before = options.before || 0;
    const after = options.after ?? 8;
    const x = options.x || margin;
    const width = options.width || maxWidth - (x - margin);

    cursorY += before;
    pdf.setFont("helvetica", options.style || "normal");
    pdf.setFontSize(fontSize);
    pdf.setTextColor(options.color || "#222222");

    const lines = pdf.splitTextToSize(cleanText, width);
    lines.forEach((line) => {
      if (cursorY + lineHeight > pageHeight - margin) {
        pdf.addPage();
        cursorY = margin;
      }
      pdf.text(line, x, cursorY);
      cursorY += lineHeight;
    });
    cursorY += after;
  };

  writeBlock(editorLesson.title, { fontSize: 22, style: "bold", lineHeight: 28, after: 20 });

  filledNotes.forEach((note) => {
    const times = formatSlideTimes(editorLesson, note.slide);
    const suffix = times !== "sem marcação" ? ` - ${times}` : "";
    writeBlock(`Slide ${note.slide}${suffix}`, { fontSize: 17, style: "bold", lineHeight: 22, before: 12, after: 10 });

    const blocks = htmlToPdfBlocks(note.html);

    blocks.forEach((block, index) => {
      if (block.type === "h1") writeBlock(block.text, { fontSize: 17, style: "bold", lineHeight: 22, before: index ? 10 : 0 });
      else if (block.type === "h2") writeBlock(block.text, { fontSize: 15, style: "bold", lineHeight: 20, before: 8 });
      else if (block.type === "h3") writeBlock(block.text, { fontSize: 13, style: "bold", lineHeight: 18, before: 6 });
      else if (block.type === "blockquote") writeBlock(block.text, { x: margin + 18, width: maxWidth - 18, fontSize: 11, color: "#444444" });
      else if (block.type === "li") writeBlock(`${block.prefix} ${block.text}`, { x: margin + 14, width: maxWidth - 14, after: 4 });
      else writeBlock(block.text);
    });
  });

  pdf.save(`${safeFileName(editorLesson.title)}-comentarios.pdf`);
}

function getFilledNotes() {
  return editorLesson.notes.filter((note) => hasNoteContent(note.html));
}

function buildPdfExportNode() {
  const node = document.createElement("article");
  node.className = "pdf-export";
  node.innerHTML = `<h1>${escapeHtml(editorLesson.title)}</h1>${editorLesson.notes
    .map((note) => {
      const times = formatSlideTimes(editorLesson, note.slide);
      const time = times !== "sem marcação" ? ` - ${times}` : "";
      const content = hasNoteContent(note.html) ? sanitizeEditorHtml(note.html) : "<p>Sem comentarios.</p>";
      return `<section><h2>Slide ${note.slide}${time}</h2>${content}</section>`;
    })
    .join("")}`;
  return node;
}

function htmlToPdfBlocks(html) {
  const template = document.createElement("template");
  template.innerHTML = sanitizeEditorHtml(html);
  const blocks = [];

  const visit = (node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent?.trim();
      if (text) blocks.push({ type: "p", text });
      return;
    }

    const tag = node.tagName?.toLowerCase();
    if (!tag) {
      Array.from(node.childNodes).forEach(visit);
      return;
    }

    if (tag === "ul" || tag === "ol") {
      Array.from(node.children).forEach((child, index) => {
        blocks.push({
          type: "li",
          prefix: tag === "ol" ? `${index + 1}.` : "-",
          text: child.textContent || ""
        });
      });
      return;
    }

    if (tag === "table") {
      Array.from(node.querySelectorAll("tr")).forEach((row) => {
        blocks.push({
          type: "p",
          text: Array.from(row.children).map((cell) => cell.textContent?.trim()).filter(Boolean).join(" | ")
        });
      });
      return;
    }

    if (["p", "blockquote", "h1", "h2", "h3", "li"].includes(tag)) {
      blocks.push({ type: tag, text: node.textContent || "" });
      return;
    }

    Array.from(node.childNodes).forEach(visit);
  };

  Array.from(template.content.childNodes).forEach(visit);
  return blocks.filter((block) => block.text.trim());
}

function htmlToDocxParagraphs(html) {
  const template = document.createElement("template");
  template.innerHTML = sanitizeEditorHtml(html);
  const blocks = Array.from(template.content.childNodes).filter((node) => node.textContent?.trim());
  if (!blocks.length) return [new Paragraph("")];
  return blocks.flatMap((node) => nodeToDocxParagraphs(node));
}

function nodeToDocxParagraphs(node) {
  if (node.nodeType === Node.TEXT_NODE) {
    return [createAcademicParagraph({ children: [new TextRun(node.textContent || "")] })];
  }

  const tag = node.tagName?.toLowerCase();
  if (tag === "ul" || tag === "ol") {
    return Array.from(node.children).map(
      (child) =>
        createAcademicParagraph({
          children: inlineRuns(child),
          bullet: tag === "ul" ? { level: 0 } : undefined,
          numbering: tag === "ol" ? { reference: "ordered-list", level: 0 } : undefined,
          indent: { left: 284 },
          spacing: { line: 324, after: 80 }
        })
    );
  }

  const heading = tag === "h1" ? HeadingLevel.HEADING_1 : tag === "h2" ? HeadingLevel.HEADING_2 : tag === "h3" ? HeadingLevel.HEADING_3 : undefined;
  if (tag === "blockquote") {
    return [
      new Paragraph({
        style: "Observation",
        children: inlineRuns(node, { size: 22, color: "444444" })
      })
    ];
  }

  return [
    createAcademicParagraph({
      heading,
      children: inlineRuns(node)
    })
  ];
}

function createAcademicParagraph(options = {}) {
  return new Paragraph({
    spacing: { line: 360, after: 160, ...options.spacing },
    alignment: AlignmentType.LEFT,
    ...options
  });
}

function inlineRuns(node, style = {}) {
  if (node.nodeType === Node.TEXT_NODE) {
    return [new TextRun({ text: node.textContent || "", ...style })];
  }
  const tag = node.tagName?.toLowerCase();
  const fontSize = getDocxFontSize(node);
  const nextStyle = {
    ...style,
    bold: style.bold || tag === "strong" || tag === "b",
    italics: style.italics || tag === "em" || tag === "i",
    size: fontSize || style.size
  };
  const runs = Array.from(node.childNodes).flatMap((child) => inlineRuns(child, nextStyle));
  return runs.length ? runs : [new TextRun("")];
}

function getDocxFontSize(node) {
  const tag = node.tagName?.toLowerCase();
  if (tag === "font") {
    const size = Number(node.getAttribute("size"));
    const sizeMap = { 1: 18, 2: 20, 3: 24, 4: 28, 5: 32, 6: 40, 7: 48 };
    return sizeMap[size];
  }
  const inlineSize = node.style?.fontSize;
  if (inlineSize?.endsWith("px")) return Math.round(Number.parseFloat(inlineSize) * 1.5);
  if (inlineSize?.endsWith("pt")) return Math.round(Number.parseFloat(inlineSize) * 2);
  return undefined;
}

function sanitizeEditorHtml(html) {
  const template = document.createElement("template");
  template.innerHTML = html || "";
  sanitizeEditorNode(template.content);
  return template.innerHTML;
}

function sanitizeEditorNode(parent) {
  const allowedTags = new Set(["P", "BR", "STRONG", "B", "EM", "I", "UL", "OL", "LI", "BLOCKQUOTE", "H1", "H2", "H3", "TABLE", "THEAD", "TBODY", "TR", "TH", "TD", "FONT"]);
  const removeWithChildren = new Set(["SCRIPT", "STYLE", "IFRAME", "OBJECT", "EMBED", "META", "LINK"]);

  Array.from(parent.childNodes).forEach((node) => {
    if (node.nodeType === Node.TEXT_NODE) return;
    if (node.nodeType !== Node.ELEMENT_NODE) {
      node.remove();
      return;
    }

    if (removeWithChildren.has(node.tagName)) {
      node.remove();
      return;
    }

    sanitizeEditorNode(node);

    if (!allowedTags.has(node.tagName)) {
      while (node.firstChild) {
        parent.insertBefore(node.firstChild, node);
      }
      node.remove();
      return;
    }

    Array.from(node.attributes).forEach((attribute) => {
      const name = attribute.name.toLowerCase();
      const value = attribute.value;
      const isFontSize = node.tagName === "FONT" && name === "size" && /^[1-7]$/.test(value);
      const isTableSpan = ["TD", "TH"].includes(node.tagName) && ["colspan", "rowspan"].includes(name) && /^\d{1,2}$/.test(value);
      if (!isFontSize && !isTableSpan) {
        node.removeAttribute(attribute.name);
      }
    });
  });
}

function hasNoteContent(html) {
  const template = document.createElement("template");
  template.innerHTML = html || "";
  return Boolean(template.content.textContent?.trim());
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function safeFileName(value) {
  return String(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9._-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
