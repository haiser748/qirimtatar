let dictionary = [];
let favorites = new Set();
let currentView = "all";
let currentCategory = "";

const $ = id => document.getElementById(id);
const input = $("searchInput");
const results = $("results");
const status = $("status");
const empty = $("empty");

// Разные написания буквы каф считаем одной буквой:
// ك — арабская каф
// ک — персидская/крымскотатарская каф
function norm(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    // Каф: ك = ک
    // Йа/я: ى = ي = ی
    .replaceAll("ك", "ک")
    .replaceAll("ى", "ي")
    .replaceAll("ی", "ي");
}

function esc(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

try {
  favorites = new Set(
    JSON.parse(localStorage.getItem("dictionaryFavorites") || "[]")
  );
} catch (_) {}

async function loadDictionary() {
  try {
    // Именно () — иначе браузер может получить устаревший файл из кэша.
    const response = await fetch("dictionary.xlsx?v=" + Date.now(), {
      cache: "no-store"
    });

    if (!response.ok) throw new Error("Excel не найден");

    const buffer = await response.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];

    dictionary = XLSX.utils
      .sheet_to_json(sheet, { defval: "" })
      .map((item, index) => ({ ...item, _id: String(index) }))
      .filter(item =>
        Object.values(item).some(value => String(value ?? "").trim() !== "")
      );

    buildCategories();
    updateCounts();
    render();
  } catch (error) {
    status.textContent = "Не удалось загрузить слова.";
    console.error(error);
  }
}

function buildCategories() {
  const map = {};

  dictionary.forEach(item => {
    const category = String(item.category || "Без категории");
    map[category] = (map[category] || 0) + 1;
  });

  const categories = $("categories");
  if (!categories) return;

  categories.innerHTML = Object.entries(map)
    .sort(([a], [b]) => a.localeCompare(b, "ru"))
    .map(
      ([category, count]) =>
        `<button class="nav-item" data-category="${esc(category)}">• <span>${esc(category)}</span><b>${count}</b></button>`
    )
    .join("");

  categories.querySelectorAll("[data-category]").forEach(button => {
    button.addEventListener("click", () => {
      currentView = "category";
      currentCategory = button.dataset.category;
      activateNav(button);
      render();
    });
  });
}

function updateCounts() {
  const allCount = $("countAll");
  const favoritesCount = $("countFav");

  if (allCount) allCount.textContent = dictionary.length;
  if (favoritesCount) favoritesCount.textContent = favorites.size;
}

function activateNav(activeButton) {
  document.querySelectorAll(".nav-item").forEach(button => {
    button.classList.remove("active");
  });

  if (activeButton) activeButton.classList.add("active");
}

function searchScore(item, query) {
  if (!query) return 0;

  const word = norm(item.word);
  const translation = norm(item.translation);
  const transcription = norm(item.transcription);
  const example = norm(item.example);
  const category = norm(item.category);

  // Поиск только по вхождению.
  // Совпадения в колонке word имеют значительно больший вес.
  let score = 0;

  if (word.includes(query)) {
    score += 1000;
    if (word === query) score += 500;
    else if (word.startsWith(query)) score += 250;
  }

  if (translation.includes(query)) {
    score += 200;
    if (translation === query) score += 50;
    else if (translation.startsWith(query)) score += 25;
  }

  if (transcription.includes(query)) {
    score += 100;
    if (transcription === query) score += 25;
    else if (transcription.startsWith(query)) score += 15;
  }

  if (example.includes(query)) score += 50;
  if (category.includes(query)) score += 25;

  return score;
}

function render() {
  let arr = [...dictionary];
  const query = norm(input?.value || "");

  if (currentView === "favorites") {
    arr = arr.filter(item => favorites.has(item._id));
  }

  if (currentView === "category") {
    arr = arr.filter(
      item => String(item.category || "Без категории") === currentCategory
    );
  }

  if (query) {
    // Сначала более релевантные совпадения.
    // word имеет максимальный вес, остальные поля — меньший.
    arr = arr
      .map(item => ({ item, score: searchScore(item, query) }))
      .filter(entry => entry.score > 0)
      .sort((a, b) => b.score - a.score)
      .map(entry => entry.item);
  }

  const sort = $("sortSelect")?.value;

  // При активном поиске сохраняем сортировку по релевантности.
  // A-Z / Z-A применяются, когда поиска нет.
  if (!query && sort === "az") {
    arr.sort((a, b) =>
      norm(a.word).localeCompare(norm(b.word), "ar")
    );
  } else if (!query && sort === "za") {
    arr.sort((a, b) =>
      norm(b.word).localeCompare(norm(a.word), "ar")
    );
  }

  const viewTitle = $("viewTitle");
  if (viewTitle) {
    viewTitle.textContent = query
      ? `Поиск: «${input.value}»`
      : currentView === "favorites"
        ? "Избранное"
        : currentView === "category"
          ? currentCategory
          : "Все слова";
  }

  if (status) status.textContent = `Найдено: ${arr.length}`;

  // Ограничиваем количество одновременно отображаемых карточек.
  results.innerHTML = arr.slice(0, 100).map(card).join("");

  empty.hidden = arr.length !== 0;
  results.hidden = arr.length === 0;

  results.querySelectorAll(".favorite").forEach(button => {
    button.addEventListener("click", event => {
      event.stopPropagation();
      toggleFav(button.dataset.id);
    });
  });
}

function card(item) {
  const isFavorite = favorites.has(item._id);

  return `<article class="card">
    <div class="card-top">
      <div>
        <div class="word">${esc(item.word)}</div>
        <div class="translation">${esc(item.translation)}</div>
        ${
          item.category
            ? `<span class="category">${esc(item.category)}</span>`
            : ""
        }
      </div>
      <button class="favorite ${isFavorite ? "on" : ""}" data-id="${item._id}" title="Избранное" aria-label="Добавить в избранное">
        ${isFavorite ? "♥" : "♡"}
      </button>
    </div>
    ${
      item.transcription
        ? `<div class="line"></div><div class="label">Транскрипция</div><div class="transcription">${esc(item.transcription)}</div>`
        : ""
    }
    ${
      item.example
        ? `<div class="line"></div><div class="label">Пример</div><div class="example">${esc(item.example)}</div>`
        : ""
    }
  </article>`;
}

function toggleFav(id) {
  if (favorites.has(id)) {
    favorites.delete(id);
  } else {
    favorites.add(id);
  }

  localStorage.setItem(
    "dictionaryFavorites",
    JSON.stringify([...favorites])
  );

  updateCounts();
  render();
}

// Навигация: Все слова / Избранное
// Данные кнопки находятся непосредственно в index.html.
document.querySelectorAll("[data-view]").forEach(button => {
  button.addEventListener("click", () => {
    currentView = button.dataset.view;
    currentCategory = "";
    activateNav(button);
    render();
  });
});

// Сортировка
$("sortSelect")?.addEventListener("change", render);

// Основной поиск
input?.addEventListener("input", () => {
  if ($("clearBtn")) {
    $("clearBtn").hidden = !input.value;
  }
  render();
});

// Очистка поиска
$("clearBtn")?.addEventListener("click", () => {
  input.value = "";
  $("clearBtn").hidden = true;
  render();
  input.focus();
});

// Кнопки быстрых слов: كتاب / بيلگي / مدرسة
// Они просто подставляют слово в поле поиска и запускают обычный поиск по вхождению.
document.querySelectorAll(".quick button[data-query]").forEach(button => {
  button.addEventListener("click", () => {
    input.value = button.dataset.query || "";
    if ($("clearBtn")) $("clearBtn").hidden = !input.value;
    render();
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
});

// Ctrl + K / Cmd + K — перейти в поле поиска
// (если браузер не перехватил эту комбинацию раньше)
document.addEventListener("keydown", event => {
  if (
    (event.ctrlKey || event.metaKey) &&
    event.key.toLowerCase() === "k"
  ) {
    event.preventDefault();
    input?.focus();
  }
});

// Светлая / тёмная тема
const savedTheme = localStorage.getItem("dictionaryTheme");
if (savedTheme) {
  document.documentElement.dataset.theme = savedTheme;
}

$("themeBtn")?.addEventListener("click", () => {
  const dark = document.documentElement.dataset.theme === "dark";
  const nextTheme = dark ? "" : "dark";

  document.documentElement.dataset.theme = nextTheme;
  localStorage.setItem("dictionaryTheme", nextTheme);
  $("themeBtn").textContent = dark ? "☾" : "☀";
});

if ($("themeBtn")) {
  $("themeBtn").textContent =
    document.documentElement.dataset.theme === "dark" ? "☀" : "☾";
}

loadDictionary();
