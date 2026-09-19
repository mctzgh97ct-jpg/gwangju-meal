const button = document.querySelector("#loadButton");
const statusBox = document.querySelector("#status");
const sourceCard = document.querySelector("#sourceCard");
const postTitle = document.querySelector("#postTitle");
const postLink = document.querySelector("#postLink");
const mealWrap = document.querySelector("#mealWrap");
const dateTabs = document.querySelector("#dateTabs");
const mealDate = document.querySelector("#mealDate");
const mealItems = document.querySelector("#mealItems");
const mealEmpty = document.querySelector("#mealEmpty");

const mealState = {
  days: [],
  selected: null,
};

button.addEventListener("click", loadLatest);

async function loadLatest() {
  button.disabled = true;
  statusBox.classList.remove("error");
  statusBox.textContent = "광주대학교 공식 게시판에서 최신 식단을 확인하고 있어요…";

  try {
    const response = await fetch("/api/latest");
    const data = await response.json();

    if (!response.ok || !data.ok) {
      throw new Error(data.error || "식단을 불러오지 못했습니다.");
    }

    postTitle.textContent = data.post.title;
    postLink.href = data.post.url;
    sourceCard.classList.remove("hidden");

    mealState.days = data.days || [];
    mealState.selected = data.default_date;
    renderDateTabs();
    renderMealCard();
    mealWrap.classList.remove("hidden");

    statusBox.textContent = "최신 게시글의 학생정식 식단을 불러왔어요.";
  } catch (error) {
    statusBox.classList.add("error");
    statusBox.textContent = `오류: ${error.message}`;
  } finally {
    button.disabled = false;
  }
}

function renderDateTabs() {
  dateTabs.innerHTML = "";

  mealState.days.forEach(day => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "date-tab" + (day.date === mealState.selected ? " selected" : "");
    const month = Number(day.date.slice(5, 7));
    const dateNum = Number(day.date.slice(8, 10));
    btn.textContent = `${month}/${dateNum} (${day.weekday})`;
    btn.addEventListener("click", () => {
      mealState.selected = day.date;
      renderDateTabs();
      renderMealCard();
    });
    dateTabs.appendChild(btn);
  });
}

function renderMealCard() {
  const day = mealState.days.find(item => item.date === mealState.selected);
  mealItems.innerHTML = "";

  if (!day) {
    mealDate.textContent = "";
    mealEmpty.classList.remove("hidden");
    return;
  }

  mealDate.textContent = `${day.date} (${day.weekday})`;

  if (!day.items.length) {
    mealEmpty.classList.remove("hidden");
    return;
  }

  mealEmpty.classList.add("hidden");
  day.items.forEach(name => {
    const li = document.createElement("li");
    li.textContent = name;
    mealItems.appendChild(li);
  });
}
