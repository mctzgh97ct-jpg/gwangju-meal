const button = document.querySelector("#loadButton");
const statusBox = document.querySelector("#status");
const sourceCard = document.querySelector("#sourceCard");
const postTitle = document.querySelector("#postTitle");
const postLink = document.querySelector("#postLink");
const mealSection = document.querySelector("#mealSection");
const dateTabs = document.querySelector("#dateTabs");
const selectedDate = document.querySelector("#selectedDate");
const mealList = document.querySelector("#mealList");
const emptyMessage = document.querySelector("#emptyMessage");
const parserNote = document.querySelector("#parserNote");
const studentTab = document.querySelector("#studentTab");
const foodTab = document.querySelector("#foodTab");
const mealCardLabel = document.querySelector("#mealCardLabel");

let mealDays = [];
let activeDate = null;
let activeCategory = "student";

button.addEventListener("click", loadLatest);
studentTab.addEventListener("click", () => setCategory("student"));
foodTab.addEventListener("click", () => setCategory("food"));
window.addEventListener("DOMContentLoaded", loadLatest);

async function loadLatest() {
  button.disabled = true;
  statusBox.classList.remove("error", "success");
  statusBox.textContent = "광주대학교 공식 게시판에서 최신 식단을 확인하고 있어요…";

  try {
    const response = await fetch("/api/latest", { cache: "no-store" });
    const data = await response.json();

    if (!response.ok || !data.ok) {
      throw new Error(data.error || "식단을 불러오지 못했습니다.");
    }

    postTitle.textContent = data.post?.title || "최신 식단 게시글";
    postLink.href = data.post?.url || data.source || "#";
    sourceCard.classList.remove("hidden");

    mealDays = Array.isArray(data.days) ? data.days : [];
    activeDate = data.default_date || mealDays[0]?.date || null;

    // 이번 주 전체에 푸드단품이 하나라도 있으면 탭을 정상 표시합니다.
    const hasFood = mealDays.some((day) => Array.isArray(day.food_items) && day.food_items.length > 0);
    foodTab.classList.toggle("category-tab-muted", !hasFood);

    renderDateTabs();
    renderMeal(activeDate);
    mealSection.classList.remove("hidden");

    parserNote.textContent = data.parse_note || "광주대학교 공식 식단표에서 불러온 메뉴입니다.";
    statusBox.classList.add("success");
    statusBox.textContent = "최신 학생정식·푸드단품 메뉴를 불러왔어요.";
  } catch (error) {
    mealDays = [];
    activeDate = null;
    mealSection.classList.add("hidden");
    statusBox.classList.add("error");
    statusBox.textContent = `오류: ${error.message}`;
  } finally {
    button.disabled = false;
  }
}

function setCategory(category) {
  activeCategory = category;
  studentTab.classList.toggle("active", category === "student");
  foodTab.classList.toggle("active", category === "food");
  studentTab.setAttribute("aria-pressed", category === "student" ? "true" : "false");
  foodTab.setAttribute("aria-pressed", category === "food" ? "true" : "false");
  renderMeal(activeDate);
}

function renderDateTabs() {
  dateTabs.innerHTML = "";

  if (!mealDays.length) {
    const message = document.createElement("p");
    message.className = "no-days";
    message.textContent = "이번 주 날짜 정보를 찾지 못했어요.";
    dateTabs.appendChild(message);
    return;
  }

  mealDays.forEach((day) => {
    const tab = document.createElement("button");
    tab.type = "button";
    tab.className = "date-tab";
    tab.dataset.date = day.date;
    tab.setAttribute("aria-pressed", day.date === activeDate ? "true" : "false");

    const date = parseLocalDate(day.date);
    const monthDay = Number.isNaN(date.getTime())
      ? day.date
      : `${date.getMonth() + 1}/${date.getDate()}`;

    tab.innerHTML = `<span>${monthDay}</span><strong>${escapeHtml(day.weekday || "")}</strong>`;
    if (day.date === activeDate) tab.classList.add("active");

    tab.addEventListener("click", () => {
      activeDate = day.date;
      renderDateTabs();
      renderMeal(activeDate);
    });

    dateTabs.appendChild(tab);
  });
}

function renderMeal(dateString) {
  mealList.innerHTML = "";
  const day = mealDays.find((item) => item.date === dateString) || mealDays[0];

  if (!day) {
    selectedDate.textContent = "식단 정보 없음";
    emptyMessage.textContent = "표시할 식단이 없습니다.";
    emptyMessage.classList.remove("hidden");
    return;
  }

  const parsed = parseLocalDate(day.date);
  selectedDate.textContent = Number.isNaN(parsed.getTime())
    ? `${day.date} ${day.weekday || ""}`
    : `${parsed.getMonth() + 1}월 ${parsed.getDate()}일 ${day.weekday || ""}요일`;

  const isFood = activeCategory === "food";
  const categoryLabel = isFood ? "푸드단품" : "학생정식";
  mealCardLabel.textContent = categoryLabel;

  const rawItems = isFood ? day.food_items : (day.student_items || day.items);
  const items = Array.isArray(rawItems) ? rawItems.filter(Boolean) : [];

  if (!items.length) {
    emptyMessage.textContent = `이 날짜에는 등록된 ${categoryLabel} 메뉴가 없어요.`;
    emptyMessage.classList.remove("hidden");
    return;
  }

  emptyMessage.classList.add("hidden");
  items.forEach((item, index) => {
    const li = document.createElement("li");
    li.innerHTML = `<span class="meal-number">${String(index + 1).padStart(2, "0")}</span><span>${escapeHtml(item)}</span>`;
    mealList.appendChild(li);
  });
}

function parseLocalDate(value) {
  if (!value || typeof value !== "string") return new Date(NaN);
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) return new Date(NaN);
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
