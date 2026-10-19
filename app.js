const main = document.querySelector("#main");
const typeNames = { lecture: "Лекция", practice: "Практическая работа", lab: "Лабораторная работа", slides: "Презентация", file: "Файл" };
const typeMarks = { lecture: "Л", practice: "П", lab: "ЛР", slides: "С", file: "↓" };
let catalogue;
let materialContent;
let slideIndex = 0;
const storage = {
  getItem(key) {
    try {
      return window.localStorage.getItem(key);
    } catch {
      return null;
    }
  },
  setItem(key, value) {
    try {
      window.localStorage.setItem(key, value);
    } catch {
      // Direct file previews may not allow browser storage.
    }
  },
};
let readingSize = Math.min(26, Math.max(15, Number(storage.getItem("electrobaza-reading-size")) || 19));

const escapeHTML = (value) => String(value).replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char]);
const findById = (items, id) => items.find((item) => item.id === id);
const route = () => decodeURIComponent(location.hash.replace(/^#\/?/, "")).split("/").filter(Boolean);
const href = (parts) => `#/${parts.map(encodeURIComponent).join("/")}`;

function breadcrumb(items) {
  return `<nav class="breadcrumb" aria-label="Навигация">${items.map((item, index) => `${index ? '<span class="crumb-separator" aria-hidden="true">/</span>' : ""}${index === items.length - 1 ? `<span class="current" aria-current="page">${escapeHTML(item.label)}</span>` : `<a href="${href(item.path)}">${escapeHTML(item.label)}</a>`}`).join("")}</nav>`;
}

function intro(eyebrow, title, subtitle = "") {
  return `<div class="intro">${eyebrow ? `<span class="eyebrow">${escapeHTML(eyebrow)}</span>` : ""}<h1>${escapeHTML(title)}</h1>${subtitle ? `<p>${escapeHTML(subtitle)}</p>` : ""}</div>`;
}

function groupPage() {
  main.innerHTML = `${intro("", "Материалы к занятиям")}<div class="group-grid">${catalogue.groups.map((group) => {
    const subjects = group.disciplines.map((id) => findById(catalogue.disciplines, id)).filter(Boolean);
    return `<a class="group-card" href="${href(["group", group.id])}"><span class="group-icon" aria-hidden="true">${escapeHTML(group.mark || "")}</span><span class="card-copy"><strong>${escapeHTML(group.name)}</strong><small>${subjects.length} ${subjects.length === 1 ? "дисциплина" : "дисциплины"} · открыть материалы</small></span><span class="arrow" aria-hidden="true">→</span></a>`;
  }).join("")}</div><a class="trainer-card" href="https://electrosim.github.io/" target="_blank" rel="noopener noreferrer" aria-label="Открыть тренажёр электромонтажа на electrosim.github.io в новой вкладке"><span class="trainer-icon" aria-hidden="true"><svg viewBox="0 0 24 24" width="23" height="23" focusable="false"><path fill="currentColor" d="M14.2 1.5 4.9 13h5.8l-.8 9.5L19.1 11h-5.8l.9-9.5Z"/></svg></span><span class="trainer-copy"><strong>Тренажёр электромонтажа</strong><small>Интерактивная практика в браузере · electrosim.github.io</small></span><span class="trainer-action">Открыть <span aria-hidden="true">↗</span></span></a>`;
  document.title = "Электробаза — учебные материалы";
}

function groupPageById(groupId) {
  const group = findById(catalogue.groups, groupId);
  if (!group) return notFound();
  const subjects = group.disciplines.map((id) => findById(catalogue.disciplines, id)).filter(Boolean);
  const currentSubject = subjects.find((subject) => subject.topics.some((topic) => topic.id === catalogue.currentTopic));
  const currentTopic = currentSubject?.topics.find((topic) => topic.id === catalogue.currentTopic);
  const current = currentTopic ? `<section class="current-callout" aria-label="Материал текущей пары"><div><span class="eyebrow">На текущую пару</span><h2>${escapeHTML(currentTopic.name)}</h2><p>${escapeHTML(currentSubject.name)} · открыть материал</p></div><a class="button button-light" href="${href(["group", group.id, "subject", currentSubject.id, "topic", currentTopic.id])}">Открыть тему →</a></section>` : "";
  main.innerHTML = `${breadcrumb([{ label: "Группы", path: [] }, { label: group.name, path: ["group", group.id] }])}${intro("Ваша группа", group.name)}${current}<h2 class="section-heading">Дисциплины</h2><div class="list">${subjects.map((subject) => `<a class="list-link" href="${href(["group", group.id, "subject", subject.id])}"><span class="row-icon" aria-hidden="true">EE</span><span class="row-copy"><strong>${escapeHTML(subject.name)}</strong><small>${subject.semester ? `${subject.semester}-й семестр · ` : ""}${subject.topics.length ? `${subject.topics.length} ${subject.topics.length === 1 ? "тема" : "темы"} · весь курс` : "Темы скоро появятся"}</small></span><span class="arrow" aria-hidden="true">→</span></a>`).join("")}</div>`;
    document.title = `${group.name} — Электробаза`;
}

function subjectPage(groupId, subjectId) {
  const group = findById(catalogue.groups, groupId);
  const subject = findById(catalogue.disciplines, subjectId);
  if (!group || !subject || !group.disciplines.includes(subject.id)) return notFound();
  const topics = subject.topics.length
    ? `<div class="list">${subject.topics.map((topic, index) => `<a class="list-link" href="${href(["group", group.id, "subject", subject.id, "topic", topic.id])}"><span class="topic-number">${String(index + 1).padStart(2, "0")}</span><span class="row-copy"><strong>${escapeHTML(topic.name)}</strong><small>${topic.materials.map((id) => findById(catalogue.materials, id)).filter(Boolean).map((material) => typeNames[material.type] || "Материал").join(" · ")}</small></span><span class="arrow" aria-hidden="true">→</span></a>`).join("")}</div>`
    : `<p class="course-note">Темы и материалы по этой дисциплине будут добавлены позже.</p>`;
  main.innerHTML = `${intro(`${group.name} · дисциплина${subject.semester ? ` · ${subject.semester}-й семестр` : ""}`, subject.name)}${topics}`;
  document.title = `${subject.name} — Электробаза`;
}

function topicPage(groupId, subjectId, topicId) {
  const group = findById(catalogue.groups, groupId);
  const subject = findById(catalogue.disciplines, subjectId);
  const topic = subject && findById(subject.topics, topicId);
  if (!group || !subject || !group.disciplines.includes(subject.id) || !topic) return notFound();
  const topicIndex = subject.topics.findIndex((item) => item.id === topic.id);
  const materials = topic.materials.map((id) => findById(catalogue.materials, id)).filter(Boolean);
  main.innerHTML = `${intro(`${subject.name} · тема ${topicIndex + 1}`, topic.name, "Материалы к занятию.")}<div class="list">${materials.map((material) => `<a class="list-link" href="${href(["material", group.id, subject.id, topic.id, material.id])}"><span class="row-icon" aria-hidden="true">${escapeHTML(typeMarks[material.type] || "М")}</span><span class="row-copy"><strong>${escapeHTML(material.name)}</strong><small>${escapeHTML(material.description || typeNames[material.type] || "Учебный материал")}</small></span><span class="arrow" aria-hidden="true">→</span></a>`).join("")}</div>`;
  document.title = `${topic.name} — Электробаза`;
}

function materialPage(groupId, subjectId, topicId, materialId) {
  const group = findById(catalogue.groups, groupId);
  const subject = findById(catalogue.disciplines, subjectId);
  const topic = subject && findById(subject.topics, topicId);
  const material = topic && topic.materials.includes(materialId) && findById(catalogue.materials, materialId);
  if (!group || !subject || !group.disciplines.includes(subject.id) || !topic || !material) return notFound();
  main.innerHTML = `${intro(typeNames[material.type] || "Материал", topic.name)}<div class="reader-layout"><div class="reader-tools"><div class="text-size-control" role="group" aria-label="Размер текста"><button type="button" class="text-size-step" id="text-size-decrease" aria-label="Уменьшить размер текста">−</button><span class="text-size-value" id="text-size-value" aria-live="polite">Размер текста: ${readingSize}px</span><button type="button" class="text-size-step" id="text-size-increase" aria-label="Увеличить размер текста">+</button></div></div><article class="material-reader" id="reader" aria-live="polite"><p class="loading">Открываем материал…</p></article><div class="pager" id="slide-controls" hidden><button type="button" class="button button-outline" id="slide-prev">← Назад</button><button type="button" class="button" id="slide-next">Далее →</button></div></div><p class="course-note"><a href="${href(["group", group.id, "subject", subject.id, "topic", topic.id])}">← Вернуться ко всем материалам темы</a></p>`;
  document.title = `${material.name}: ${topic.name} — Электробаза`;
  const updateReadingSize = (delta) => {
    readingSize = Math.min(26, Math.max(15, readingSize + delta));
    storage.setItem("electrobaza-reading-size", String(readingSize));
    document.querySelector("#reader").style.setProperty("--reading-size", `${readingSize}px`);
    document.querySelector("#text-size-value").textContent = `Размер текста: ${readingSize}px`;
  };
  document.querySelector("#text-size-decrease").addEventListener("click", () => updateReadingSize(-1));
  document.querySelector("#text-size-increase").addEventListener("click", () => updateReadingSize(1));
  const content = materialContent[material.id];
  if (typeof content !== "string") {
    document.querySelector("#reader").innerHTML = `<p class="error">Материал пока недоступен. Обновите страницу немного позже.</p>`;
    return;
  }
  document.querySelector("#reader").innerHTML = content;
  document.querySelector("#reader").style.setProperty("--reading-size", `${readingSize}px`);
  if (material.type === "slides") setupSlides();
}

function setupSlides() {
  const reader = document.querySelector("#reader");
  const slides = [...reader.querySelectorAll("[data-slide]")];
  if (!slides.length) return;
  const controls = document.querySelector("#slide-controls");
  const previous = document.querySelector("#slide-prev");
  const next = document.querySelector("#slide-next");
  controls.hidden = false;
  slideIndex = 0;
  const update = () => {
    slides.forEach((slide, index) => { slide.hidden = index !== slideIndex; });
    previous.disabled = slideIndex === 0;
    next.disabled = slideIndex === slides.length - 1;
    previous.setAttribute("aria-label", `Предыдущий слайд, ${slideIndex} из ${slides.length}`);
    next.setAttribute("aria-label", `Следующий слайд, ${slideIndex + 2} из ${slides.length}`);
  };
  previous.addEventListener("click", () => { if (slideIndex > 0) { slideIndex -= 1; update(); } });
  next.addEventListener("click", () => { if (slideIndex < slides.length - 1) { slideIndex += 1; update(); } });
  update();
}

function notFound() {
  main.innerHTML = `${breadcrumb([{ label: "Группы", path: [] }, { label: "Материал не найден", path: ["not-found"] }])}<div class="intro"><span class="eyebrow">Такой страницы нет</span><h1>Материал не найден</h1><p>Возможно, ссылка устарела. Вернитесь к списку групп.</p></div><a class="button" href="#/">К группам</a>`;
  document.title = "Материал не найден — Электробаза";
}

function render() {
  const parts = route();
  if (parts.length === 0) return groupPage();
  if (parts[0] === "group" && parts.length === 2) return groupPageById(parts[1]);
  if (parts[0] === "group" && parts[2] === "subject" && parts.length === 4) return subjectPage(parts[1], parts[3]);
  if (parts[0] === "group" && parts[2] === "subject" && parts[4] === "topic" && parts.length === 6) return topicPage(parts[1], parts[3], parts[5]);
  if (parts[0] === "material" && parts.length === 5) return materialPage(parts[1], parts[2], parts[3], parts[4]);
  return notFound();
}

function start() {
  catalogue = window.ELECTROBAZA_CATALOG;
  materialContent = window.ELECTROBAZA_CONTENT;
  if (!catalogue || !materialContent) {
    main.innerHTML = `<p class="error">Не найден каталог сайта. Запустите build-catalogue.cmd и откройте index.html заново.</p>`;
    return;
  }
  window.addEventListener("hashchange", render);
  render();
}

start();
