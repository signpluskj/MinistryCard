const state = {
  apiUrl: "",
  user: null,
  data: {
    cards: [],
    areas: [],
    completions: [],
    visits: [],
    evangelists: []
  },
  selectedArea: null,
  selectedCard: null,
  expandedAreaId: null,
  view: "cards",
  sortOrder: "number",
  searchQuery: "",
  filterArea: "all",
  filterStatus: "all",
  filterVisit: "all",
  editingVisit: null,
  statusTimer: null
};

const elements = {
  configPanel: document.getElementById("config-panel"),
  loginPanel: document.getElementById("login-panel"),
  dashboard: document.getElementById("dashboard"),
  nameInput: document.getElementById("name-input"),
  passwordInput: document.getElementById("password-input"),
  loginButton: document.getElementById("login-button"),
  apiUrlInput: document.getElementById("api-url-input"),
  saveApiUrl: document.getElementById("save-api-url"),
  userInfo: document.getElementById("user-info"),
  areaList: document.getElementById("area-list"),
  areaTitle: document.getElementById("area-title"),
  cardList: document.getElementById("card-list"),
  cardListHome: document.getElementById("card-list").parentElement,
  visitForm: document.getElementById("visit-form"),
  visitDate: document.getElementById("visit-date"),
  visitWorker: document.getElementById("visit-worker"),
  visitResult: document.getElementById("visit-result"),
  visitNote: document.getElementById("visit-note"),
  visitTitle: document.getElementById("visit-title"),
  statusMessage: document.getElementById("status-message"),
  leaderActions: document.getElementById("leader-actions"),
  startService: document.getElementById("start-service"),
  areaOverlay: document.getElementById("area-overlay"),
  closeAreas: document.getElementById("close-areas"),
  areaListOverlay: document.getElementById("area-list-overlay"),
  adminPanel: document.getElementById("admin-panel"),
  completionList: document.getElementById("completion-list"),
  visitList: document.getElementById("visit-list"),
  evangelistList: document.getElementById("evangelist-list"),
  searchInput: document.getElementById("search-input"),
  searchButton: document.getElementById("search-button"),
  filterArea: document.getElementById("filter-area"),
  filterVisit: document.getElementById("filter-visit"),
  areaListInline: document.getElementById("area-list-inline"),
  loadingIndicator: document.getElementById("loading-indicator"),
  loadingText: document.getElementById("loading-text")
};

const todayISO = () => new Date().toISOString().slice(0, 10);

const isSameDay = (value, isoString) => {
  if (!value) {
    return false;
  }
  const d1 = parseVisitDate(value);
  const d2 = new Date(isoString);
  if (!d1 || Number.isNaN(d2.getTime())) {
    return false;
  }
  return (
    d1.getFullYear() === d2.getFullYear() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate()
  );
};

const formatDate = (value) => {
  if (!value) {
    return "";
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  const yy = String(date.getFullYear()).slice(-2);
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yy}/${mm}/${dd}`;
};

const toISODate = (value) => {
  const d = parseVisitDate(value);
  if (!d) {
    return todayISO();
  }
  const z = new Date(d.getTime() - d.getTimezoneOffset() * 60000);
  return z.toISOString().slice(0, 10);
};

const isTrueValue = (value) =>
  value === true ||
  value === "TRUE" ||
  value === "true" ||
  value === "Y" ||
  value === "1" ||
  value === 1;

const getVisitDateValue = (row) =>
  row["방문일"] ||
  row["방문날짜"] ||
  row["방문일자"] ||
  row["날짜"] ||
  row["방문일시"] ||
  row["일자"] ||
  "";

const getVisitCardNumber = (row) => row["카드번호"] || row["구역카드"] || "";

const parseVisitDate = (value) => {
  if (!value) {
    return null;
  }
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }
  const text = String(value).trim().replace(/\s+/g, "");
  const m = text.match(/^(\d{2,4})[./-](\d{1,2})[./-](\d{1,2})$/);
  if (m) {
    const y = m[1].length === 2 ? 2000 + Number(m[1]) : Number(m[1]);
    const mo = Number(m[2]) - 1;
    const d = Number(m[3]);
    const dt = new Date(y, mo, d);
    return Number.isNaN(dt.getTime()) ? null : dt;
  }
  const parsed = new Date(text.replace(/\./g, "-").replace(/\//g, "-"));
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const setStatus = (message) => {
  if (state.statusTimer) {
    clearTimeout(state.statusTimer);
    state.statusTimer = null;
  }
  elements.statusMessage.textContent = message || "";
  if (message) {
    state.statusTimer = setTimeout(() => {
      elements.statusMessage.textContent = "";
      state.statusTimer = null;
    }, 3000);
  }
};

const setLoading = (isLoading, message) => {
  if (message) {
    elements.loadingText.textContent = message;
  }
  elements.loadingIndicator.classList.toggle("hidden", !isLoading);
};

const loadApiUrl = () => {
  const stored = localStorage.getItem("ministry_api_url") || "";
  state.apiUrl = stored;
  elements.apiUrlInput.value = stored;
  elements.configPanel.classList.toggle("hidden", Boolean(stored));
};

const saveApiUrl = () => {
  const value = elements.apiUrlInput.value.trim();
  if (!value) {
    alert("앱스 스크립트 URL을 입력해 주세요.");
    return;
  }
  localStorage.setItem("ministry_api_url", value);
  state.apiUrl = value;
  elements.configPanel.classList.add("hidden");
};

const apiRequest = async (action, payload = {}, method = "POST") => {
  if (!state.apiUrl) {
    elements.configPanel.classList.remove("hidden");
    throw new Error("API URL이 설정되지 않았습니다.");
  }
  try {
    if (method === "GET") {
      const url = new URL(state.apiUrl);
      url.searchParams.set("action", action);
      const res = await fetch(url.toString(), { method: "GET" });
      return res.json();
    }
    const body = new URLSearchParams({ action, ...payload }).toString();
    const res = await fetch(state.apiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8" },
      body
    });
    return res.json();
  } catch (error) {
    alert("서버에 연결할 수 없습니다. 나중에 다시 시도해 주세요.");
    throw error;
  }
};

const groupCardsByArea = () => {
  const areas = {};
  state.data.cards.forEach((card) => {
    const key = String(card["구역번호"] || card.area || "");
    if (!areas[key]) {
      areas[key] = [];
    }
    areas[key].push(card);
  });
  return areas;
};

const areaCompletionStatus = (cards) => cards.every((card) => Boolean(card["최근방문일"]));

const getFirstInProgressArea = () => {
  const grouped = groupCardsByArea();
  const areaIds = Object.keys(grouped);
  for (const areaId of areaIds) {
    if (!areaCompletionStatus(grouped[areaId])) {
      return areaId;
    }
  }
  return null;
};

const renderAreas = () => {
  const grouped = groupCardsByArea();
  elements.areaList.innerHTML = "";
  elements.areaListOverlay.innerHTML = "";
  elements.areaListInline.innerHTML = "";
  elements.filterArea.innerHTML = "";
  const allOption = document.createElement("option");
  allOption.value = "all";
  allOption.textContent = "전체구역";
  elements.filterArea.appendChild(allOption);
  const today = todayISO();
  const areaIds = Object.keys(grouped);
  areaIds.forEach((areaId) => {
    const createItem = () => {
      const item = document.createElement("div");
      item.className = "list-item";
      item.dataset.areaId = areaId;
      if (state.selectedArea === areaId) {
        item.classList.add("active");
      }
      const areaInfo = state.data.areas.find(
        (row) => String(row["구역번호"]) === String(areaId)
      );
      let dateSuffix = "";
      if (areaInfo) {
        const start = areaInfo["시작날짜"] ? formatDate(areaInfo["시작날짜"]) : "";
        const end = areaInfo["완료날짜"] ? formatDate(areaInfo["완료날짜"]) : "";
        const range =
          start && end ? `${start} ~ ${end}` : start ? start : end ? end : "";
        if (range) {
          dateSuffix = ` [${range}]`;
        }
      }
      const isToday = areaInfo && isSameDay(areaInfo["시작날짜"], today);
      const title = document.createElement("span");
      title.textContent = `${areaId}${dateSuffix}`;
      const badge = document.createElement("span");
      badge.className = "status-badge";
      const isComplete = areaCompletionStatus(grouped[areaId]);
      badge.textContent = isComplete ? "완료" : "진행중";
      if (isToday) {
        item.classList.add("area-today");
      }
      item.append(title, badge);
      item.addEventListener("click", (event) => {
        if (
          event.target.closest(".area-cards") ||
          event.target.closest("#card-list") ||
          event.target.closest(".card") ||
          event.target.closest(".card-history") ||
          event.target.closest("#visit-form")
        ) {
          return;
        }
        if (state.expandedAreaId === areaId) {
          state.expandedAreaId = null;
          state.filterArea = "all";
          elements.filterArea.value = "all";
          state.selectedArea = null;
        } else {
          state.expandedAreaId = areaId;
          state.filterArea = areaId;
          elements.filterArea.value = areaId;
          selectArea(areaId);
        }
        renderAreas();
        renderCards();
      });
      return item;
    };
    const createStartButton = () => {
      const startBtn = document.createElement("button");
      startBtn.type = "button";
      startBtn.textContent = "봉사 시작";
      startBtn.style.marginLeft = "auto";
      startBtn.className = "start-service-btn";
      startBtn.addEventListener("click", (event) => {
        event.stopPropagation();
        state.expandedAreaId = areaId;
        state.filterArea = areaId;
        elements.filterArea.value = areaId;
        selectArea(areaId);
        startServiceForArea(areaId);
        elements.areaOverlay.classList.add("hidden");
        renderAreas();
        renderCards();
      });
      return startBtn;
    };
    const leftItem = createItem();
    const overlayItem = createItem();
    const inlineItem = createItem();
    if (areaCompletionStatus(grouped[areaId])) {
      overlayItem.appendChild(createStartButton());
      inlineItem.appendChild(createStartButton());
    }
    elements.areaList.appendChild(leftItem);
    elements.areaListOverlay.appendChild(overlayItem);
    elements.areaListInline.appendChild(inlineItem);
    const option = document.createElement("option");
    option.value = areaId;
    option.textContent = areaId;
    elements.filterArea.appendChild(option);
  });
  elements.filterArea.value = state.filterArea;
};

const renderCards = () => {
  elements.cardList.innerHTML = "";
  let rawCards = state.data.cards.slice();
  if (state.filterArea !== "all") {
    rawCards = rawCards.filter(
      (card) => String(card["구역번호"]) === String(state.filterArea)
    );
  }
  const areaInfo =
    state.selectedArea && state.filterArea !== "all"
      ? state.data.areas.find(
          (row) => String(row["구역번호"]) === String(state.selectedArea)
        )
      : null;
  let startDate = null;
  if (areaInfo && areaInfo["시작날짜"]) {
    const parsed = parseVisitDate(areaInfo["시작날짜"]);
    if (parsed) {
      startDate = parsed;
    }
  }
  let cards = rawCards.slice();
  const query = state.searchQuery.trim();
  const shouldHideAll =
    state.filterArea === "all" &&
    !state.expandedAreaId &&
    !query &&
    state.filterStatus === "all" &&
    state.filterVisit === "all";
  if (shouldHideAll) {
    elements.cardListHome.appendChild(elements.cardList);
    elements.cardList.classList.add("hidden");
    elements.visitForm.classList.add("hidden");
    elements.cardListHome.appendChild(elements.visitForm);
    return;
  }
  if (query) {
    cards = cards.filter((card) =>
      `${card["주소"] || ""} ${card["상세주소"] || ""}`.includes(query)
    );
  }
  if (state.filterVisit === "meet") {
    cards = cards.filter((card) => isTrueValue(card["만남"]));
  } else if (state.filterVisit === "absent") {
    cards = cards.filter(
      (card) => card["만남"] === false || isTrueValue(card["부재"])
    );
  } else if (state.filterVisit === "invite") {
    cards = cards.filter((card) => isTrueValue(card["초대장"]));
  } else if (state.filterVisit === "six") {
    cards = cards.filter((card) => isTrueValue(card["6개월"]));
  } else if (state.filterVisit === "banned") {
    cards = cards.filter((card) => isTrueValue(card["방문금지"]));
  }
  cards.sort((a, b) => {
    const getUnvisited = (card) => {
      if (!startDate) {
        return false;
      }
      if (!card["최근방문일"]) {
        return true;
      }
      const d = parseVisitDate(card["최근방문일"]);
      if (!d) {
        return false;
      }
      return d < startDate;
    };
    if (state.sortOrder === "visitDate") {
      const da = parseVisitDate(a["최근방문일"]);
      const db = parseVisitDate(b["최근방문일"]);
      const va = da ? da.getTime() : 0;
      const vb = db ? db.getTime() : 0;
      if (va !== vb) {
        return vb - va;
      }
    } else if (state.sortOrder === "worker") {
      const wa = String(a["정기방문자"] || "");
      const wb = String(b["정기방문자"] || "");
      if (wa !== wb) {
        return wa.localeCompare(wb);
      }
    } else if (state.sortOrder === "address") {
      const aa = `${a["주소"] || ""} ${a["상세주소"] || ""}`;
      const ab = `${b["주소"] || ""} ${b["상세주소"] || ""}`;
      if (aa !== ab) {
        return aa.localeCompare(ab);
      }
    } else if (state.sortOrder === "number") {
      const na = String(a["카드번호"] || "");
      const nb = String(b["카드번호"] || "");
      if (na !== nb) {
        return na.localeCompare(nb);
      }
    }
    const na = String(a["카드번호"] || "");
    const nb = String(b["카드번호"] || "");
    return na.localeCompare(nb);
  });
  cards.forEach((card) => {
    const cardEl = document.createElement("div");
    cardEl.className = "card";
    if (state.selectedCard && card["카드번호"] === state.selectedCard["카드번호"]) {
      cardEl.classList.add("active");
    }
    let unvisited = false;
    if (startDate) {
      if (card["최근방문일"]) {
        const recentDate = parseVisitDate(card["최근방문일"]);
        if (recentDate && recentDate < startDate) {
          unvisited = true;
        }
      }
    }
    if (unvisited) {
      cardEl.classList.add("card-unvisited");
    }
    const title = document.createElement("div");
    title.className = "card-header";
    const mainTitle = document.createElement("strong");
    mainTitle.textContent = card["카드번호"] || "";
    const badgeRow = document.createElement("div");
    badgeRow.className = "card-badges";
    if (card["방문금지"]) {
      const ban = document.createElement("span");
      ban.className = "badge badge-danger";
      ban.textContent = "방문금지";
      badgeRow.appendChild(ban);
    }
    const meetVal = card["만남"];
    const absentVal = card["부재"];
    const recentVal = card["최근방문일"];
    if (!recentVal) {
      const unv = document.createElement("span");
      unv.className = "badge badge-unvisited";
      unv.textContent = "미방문";
      badgeRow.appendChild(unv);
    } else if (isTrueValue(meetVal)) {
      const meet = document.createElement("span");
      meet.className = "badge badge-success";
      meet.textContent = "만남";
      badgeRow.appendChild(meet);
    } else if (meetVal === false || isTrueValue(absentVal)) {
      const absent = document.createElement("span");
      absent.className = "badge";
      absent.textContent = "부재";
      badgeRow.appendChild(absent);
    } else if (card["초대장"]) {
      const invite = document.createElement("span");
      invite.className = "badge badge-invite";
      invite.textContent = "초대장";
      badgeRow.appendChild(invite);
    } else if (card["6개월"]) {
      const six = document.createElement("span");
      six.className = "badge badge-sixmonths";
      six.textContent = "6개월";
      badgeRow.appendChild(six);
    }
    title.append(mainTitle, badgeRow);
    const address = document.createElement("div");
    address.className = "card-line";
    const addrText = [card["주소"], card["상세주소"]].filter(Boolean).join(" ");
    address.textContent = addrText;
    const visitInfo = document.createElement("div");
    visitInfo.className = "card-line";
    visitInfo.textContent = card["최근방문일"]
      ? `최근방문: ${formatDate(card["최근방문일"])}`
      : "최근방문 없음";
    const regular = document.createElement("div");
    regular.className = "card-line";
    if (card["정기방문자"]) {
      regular.textContent = `정기방문자: ${card["정기방문자"]}`;
    }
    const info = document.createElement("div");
    info.className = "card-line";
    if (card["정보"]) {
      info.textContent = `정보: ${card["정보"]}`;
    }
    const nav = document.createElement("div");
    nav.className = "nav-links";
    const searchText = card["주소"] || "";
    const encoded = encodeURIComponent(searchText || card["카드번호"] || "");
    const kakao = document.createElement("a");
    kakao.href = `https://map.kakao.com/link/search/${encoded}`;
    kakao.target = "_blank";
    kakao.rel = "noopener noreferrer";
    kakao.textContent = "카카오";
    const naver = document.createElement("a");
    naver.href = `https://m.map.naver.com/search2/search.naver?query=${encoded}`;
    naver.target = "_blank";
    naver.rel = "noopener noreferrer";
    naver.textContent = "네이버";
    const tmap = document.createElement("a");
    tmap.href = "https://www.tmap.co.kr";
    tmap.target = "_blank";
    tmap.rel = "noopener noreferrer";
    tmap.textContent = "티맵";
    const google = document.createElement("a");
    google.href = `https://www.google.com/maps/search/?api=1&query=${encoded}`;
    google.target = "_blank";
    google.rel = "noopener noreferrer";
    google.textContent = "구글";
    nav.append(kakao, naver, tmap, google);
    cardEl.append(title, address, visitInfo);
    if (regular.textContent) {
      cardEl.appendChild(regular);
    }
    if (info.textContent) {
      cardEl.appendChild(info);
    }
    cardEl.appendChild(nav);
    if (state.selectedCard && card["카드번호"] === state.selectedCard["카드번호"]) {
      const oneYearAgo = new Date();
      oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
    const normalizeId = (v) =>
      String(v == null ? "" : v).trim().replace(/\s+/g, "");
    const allRows = (state.data.visits || [])
        .filter(
          (row) =>
          normalizeId(row["구역카드"]) === normalizeId(card["카드번호"])
      );
      const lastYearRows = allRows.filter((row) => {
        const d = parseVisitDate(getVisitDateValue(row));
        return d && d >= oneYearAgo;
      });
    const historyRows = (lastYearRows.length ? lastYearRows : allRows).sort(
      (a, b) => {
        const db = parseVisitDate(getVisitDateValue(b));
        const da = parseVisitDate(getVisitDateValue(a));
        const tb = db ? db.getTime() : 0;
        const ta = da ? da.getTime() : 0;
        return tb - ta;
      }
    );
      const history = document.createElement("div");
      history.className = "card-history";
      if (!historyRows.length) {
        const empty = document.createElement("div");
        empty.className = "card-history-empty";
        empty.textContent = "방문내역 없음";
        history.appendChild(empty);
      } else {
        const maxItems = 200;
        historyRows.slice(0, maxItems).forEach((row) => {
          const resultText = row["결과"] || row["방문결과"] || "";
          const workerText = row["전도인"] || row["방문자"] || "";
          const memoText = row["메모"] || row["비고"] || "";
          const item = document.createElement("div");
          item.className = "card-history-item";
          item.textContent = `${formatDate(getVisitDateValue(row))} · ${workerText} · ${resultText}${memoText ? " · " + memoText : ""}`;
          item.addEventListener("click", (event) => {
            event.stopPropagation();
            state.editingVisit = {
              areaId: String(card["구역번호"]),
              cardNumber: String(card["카드번호"]),
              oldVisitDate: String(getVisitDateValue(row)),
              oldWorker: String(workerText),
              oldResult: String(resultText),
              oldNote: String(memoText)
            };
            elements.visitTitle.textContent = `카드 ${card["카드번호"]} 방문 내역 수정`;
            elements.visitDate.value = toISODate(getVisitDateValue(row));
            elements.visitWorker.value = workerText;
            elements.visitResult.value = resultText || "만남";
            elements.visitNote.value = memoText || "";
          });
          history.appendChild(item);
        });
      }
      cardEl.appendChild(history);
      elements.visitForm.classList.remove("hidden");
      cardEl.appendChild(elements.visitForm);
    }
    cardEl.addEventListener("click", (event) => {
      if (event.target.closest("#visit-form")) {
        return;
      }
      if (
        state.selectedCard &&
        state.selectedCard["카드번호"] === card["카드번호"]
      ) {
        state.selectedCard = null;
        elements.visitForm.classList.add("hidden");
        elements.cardList.parentElement.appendChild(elements.visitForm);
        renderCards();
        return;
      }
      selectCard(card);
    });
    elements.cardList.appendChild(cardEl);
  });
  if (!state.selectedCard) {
    elements.visitForm.classList.add("hidden");
    elements.cardList.parentElement.appendChild(elements.visitForm);
  }
  const expanded =
    state.filterArea !== "all"
      ? elements.areaListInline.querySelector(
          `[data-area-id="${state.expandedAreaId || ""}"]`
        )
      : null;
  if (expanded) {
    let container = expanded.querySelector(".area-cards");
    if (!container) {
      container = document.createElement("div");
      container.className = "area-cards";
      expanded.appendChild(container);
    }
    container.appendChild(elements.cardList);
    expanded.scrollIntoView({ behavior: "smooth", block: "start" });
  } else {
    elements.cardListHome.appendChild(elements.cardList);
  }
  elements.cardList.classList.remove("hidden");
};

const renderVisitsView = () => {
  elements.cardList.innerHTML = "";
  elements.visitForm.classList.add("hidden");
  const query = state.searchQuery.trim();
  const filterResult = state.visitFilter;
  const getCardKey = (row) =>
    `${row["구역번호"] || ""}__${getVisitCardNumber(row)}`;
  const cardMap = {};
  state.data.cards.forEach((card) => {
    const key = `${card["구역번호"] || ""}__${card["카드번호"] || ""}`;
    cardMap[key] = card;
  });
  let rows = (state.data.visits || []).slice();
  if (filterResult === "meet") {
    rows = rows.filter((r) => String(r["결과"] || r["방문결과"] || "") === "만남");
  } else if (filterResult === "absent") {
    rows = rows.filter((r) => String(r["결과"] || r["방문결과"] || "") === "부재");
  } else if (filterResult === "invite") {
    rows = rows.filter((r) => String(r["결과"] || r["방문결과"] || "") === "초대장");
  } else if (filterResult === "six") {
    rows = rows.filter((r) => String(r["결과"] || r["방문결과"] || "") === "6개월");
  } else if (filterResult === "banned") {
    rows = rows.filter((r) => String(r["결과"] || r["방문결과"] || "") === "방문금지");
  }
  if (query) {
    rows = rows.filter((row) => {
      const card = cardMap[getCardKey(row)] || {};
      const addr = `${card["주소"] || ""} ${card["상세주소"] || ""}`;
      return addr.includes(query);
    });
  }
  rows.sort((a, b) => {
    if (state.sortOrder === "visitDate") {
      const da = parseVisitDate(getVisitDateValue(a));
      const db = parseVisitDate(getVisitDateValue(b));
      const va = da ? da.getTime() : 0;
      const vb = db ? db.getTime() : 0;
      if (va !== vb) {
        return vb - va;
      }
    } else if (state.sortOrder === "worker") {
      const wa = String(a["전도인"] || "");
      const wb = String(b["전도인"] || "");
      if (wa !== wb) {
        return wa.localeCompare(wb);
      }
    } else if (state.sortOrder === "address") {
      const cardA = cardMap[getCardKey(a)] || {};
      const cardB = cardMap[getCardKey(b)] || {};
      const aa = `${cardA["주소"] || ""} ${cardA["상세주소"] || ""}`;
      const ab = `${cardB["주소"] || ""} ${cardB["상세주소"] || ""}`;
      if (aa !== ab) {
        return aa.localeCompare(ab);
      }
    }
    const da = parseVisitDate(getVisitDateValue(a));
    const db = parseVisitDate(getVisitDateValue(b));
    const va = da ? da.getTime() : 0;
    const vb = db ? db.getTime() : 0;
    return vb - va;
  });
  elements.cardList.classList.remove("card-grid");
  rows.slice(0, 200).forEach((row) => {
    const item = document.createElement("div");
    item.className = "list-item";
    const title = document.createElement("div");
    title.textContent = `${row["구역번호"] || ""} - ${getVisitCardNumber(row)} (${formatDate(
      getVisitDateValue(row)
    )})`;
    const sub = document.createElement("div");
    sub.textContent = `${row["전도인"] || row["방문자"] || ""} · ${
      row["결과"] || row["방문결과"] || ""
    }${row["메모"] || row["비고"] ? " · " + (row["메모"] || row["비고"]) : ""}`;
    item.append(title, sub);
    elements.cardList.appendChild(item);
  });
};

const startServiceForArea = async (areaId) => {
  if (!areaId) {
    alert("구역을 선택해 주세요.");
    return;
  }
  setLoading(true, "봉사 시작 중...");
  try {
    await apiRequest("startService", {
      areaId,
      leaderName: state.user.name,
      date: todayISO()
    });
    setStatus("오늘 봉사가 시작되었습니다.");
    await loadData();
    renderAreas();
    renderAdminPanel();
    state.selectedArea = areaId;
    state.view = "area";
    elements.areaTitle.textContent = `구역 ${areaId}`;
    renderCards();
  } finally {
    setLoading(false);
  }
};

const renderAdminPanel = () => {
  if (!state.user || state.user.role !== "관리자") {
    elements.adminPanel.classList.add("hidden");
    return;
  }
  elements.adminPanel.classList.remove("hidden");
  elements.completionList.innerHTML = state.data.completions
    .map(
      (row) =>
        `<div class="list-item">${row["구역번호"]} | ${formatDate(
          row["시작날짜"]
        )} → ${formatDate(row["완료날짜"])}</div>`
    )
    .join("");
  elements.visitList.innerHTML = state.data.visits
    .slice(-20)
    .map(
      (row) =>
        `<div class="list-item">${row["구역번호"]} ${getVisitCardNumber(
          row
        )} | ${formatDate(
          getVisitDateValue(row)
        )} | ${row["전도인"] || row["방문자"] || ""}</div>`
    )
    .join("");
  elements.evangelistList.innerHTML = state.data.evangelists
    .map((row) => `<div class="list-item">${row["이름"]} | ${row["역할"]}</div>`)
    .join("");
};

const selectArea = (areaId) => {
  state.selectedArea = areaId;
  state.selectedCard = null;
  elements.areaTitle.textContent = `구역 ${areaId}`;
  elements.leaderActions.classList.toggle("hidden", state.user.role === "전도인");
  renderAreas();
  renderCards();
  setStatus("");
};

const selectCard = (card) => {
  state.selectedCard = card;
  state.editingVisit = null;
  elements.visitTitle.textContent = `카드 ${card["카드번호"]} 방문 기록`;
  elements.visitDate.value = todayISO();
  elements.visitWorker.value = state.user.name;
  const meetVal = card["만남"];
  const absentVal = card["부재"];
  elements.visitResult.value = isTrueValue(meetVal)
    ? "만남"
    : meetVal === false || isTrueValue(absentVal)
    ? "부재"
    : card["초대장"]
    ? "초대장"
    : card["6개월"]
    ? "6개월"
    : card["방문금지"]
    ? "방문금지"
    : "만남";
  elements.visitNote.value = "";
  renderCards();
};

const loadData = async () => {
  setLoading(true, "데이터 불러오는 중...");
  try {
    const data = await apiRequest("bootstrap", {}, "GET");
    if (data.error) {
      throw new Error(data.error);
    }
    state.data.cards = data.cards || [];
    state.data.areas = data.areas || data.areaStatus || [];
    state.data.completions = data.completions || [];
    state.data.visits = data.visits || [];
    state.data.evangelists = data.evangelists || [];
  } finally {
    setLoading(false);
  }
};

const login = async () => {
  const name = elements.nameInput.value.trim();
  const password = elements.passwordInput.value.trim();
  if (!name) {
    alert("이름을 입력해 주세요.");
    return;
  }
  const res = await apiRequest("login", { name, password });
  if (!res.success) {
    alert(res.message || "로그인 실패");
    return;
  }
  state.user = { role: res.role, name: res.name };
  elements.userInfo.textContent = `${state.user.name} (${state.user.role})`;
  elements.loginPanel.classList.add("hidden");
  elements.dashboard.classList.remove("hidden");
  await loadData();
  if (!state.expandedAreaId && state.filterArea === "all") {
    const inProgressAreaId = getFirstInProgressArea();
    if (inProgressAreaId) {
      state.expandedAreaId = inProgressAreaId;
      state.filterArea = inProgressAreaId;
      state.selectedArea = inProgressAreaId;
    }
  }
  renderAreas();
  renderCards();
  renderAdminPanel();
};

const startService = async () => {
  if (!state.selectedArea) {
    alert("구역을 선택해 주세요.");
    return;
  }
  await startServiceForArea(state.selectedArea);
};

const resetRecentVisits = async () => {
  if (!state.selectedArea) {
    alert("구역을 선택해 주세요.");
    return;
  }
  await apiRequest("resetRecentVisits", { areaId: state.selectedArea });
  setStatus("최근방문일이 초기화되었습니다.");
  await loadData();
  renderAreas();
  renderCards();
  renderAdminPanel();
};

const saveVisit = async (event) => {
  event.preventDefault();
  if (!state.selectedArea || !state.selectedCard) {
    return;
  }
  const areaId = state.selectedArea;
  const cardNumber = state.selectedCard["카드번호"];
  const visitDate = elements.visitDate.value;
  const worker = elements.visitWorker.value.trim();
  const result = elements.visitResult.value;
  const note = elements.visitNote.value.trim();
  if (!visitDate || !worker || !result) {
    alert("방문일, 전도인, 결과를 입력해 주세요.");
    return;
  }
  const oldISO =
    state.editingVisit && state.editingVisit.oldVisitDate
      ? toISODate(state.editingVisit.oldVisitDate)
      : null;
  const isEdit =
    Boolean(state.editingVisit) &&
    visitDate === oldISO &&
    worker === String(state.editingVisit.oldWorker || "");
  setLoading(true, isEdit ? "방문 기록 수정 중..." : "방문 기록 저장 중...");
  try {
    const res = isEdit
      ? await apiRequest("updateVisit", {
          areaId,
          cardNumber,
          oldVisitDate: state.editingVisit.oldVisitDate,
          oldWorker: state.editingVisit.oldWorker,
          newVisitDate: visitDate,
          newWorker: worker,
          newResult: result,
          newNote: note,
          leaderName: state.user.name
        })
      : await apiRequest("recordVisit", {
          areaId,
          cardNumber,
          visitDate,
          worker,
          result,
          note,
          leaderName: state.user.name
        });
    if (res && res.visit) {
      if (isEdit) {
        const oldKey = (row) =>
          `${row["구역번호"] || ""}__${getVisitCardNumber(row) || ""}__${String(getVisitDateValue(row) || "")}__${String(row["전도인"] || row["방문자"] || "")}`;
        const newKey = `${String(areaId)}__${String(cardNumber)}__${String(res.visit["방문날짜"] || res.visit["방문일"] || res.visit["방문일자"] || "")}__${String(res.visit["전도인"] || res.visit["방문자"] || "")}`;
        state.data.visits = (state.data.visits || []).map((row) => (oldKey(row) === `${String(state.editingVisit.areaId)}__${String(state.editingVisit.cardNumber)}__${String(state.editingVisit.oldVisitDate)}__${String(state.editingVisit.oldWorker)}` ? res.visit : row));
      } else {
        state.data.visits = [res.visit, ...(state.data.visits || [])];
      }
    }
    const updatedCard = state.data.cards.find(
      (card) =>
        String(card["구역번호"]) === String(areaId) &&
        String(card["카드번호"]) === String(cardNumber)
    );
    if (updatedCard) {
      updatedCard["최근방문일"] = res?.cardUpdate?.recentVisitDate || visitDate;
      updatedCard["만남"] = res?.cardUpdate?.meet ?? (result === "만남");
      updatedCard["부재"] = res?.cardUpdate?.absent ?? (result === "부재");
      updatedCard["초대장"] = res?.cardUpdate?.invite ?? (result === "초대장");
      updatedCard["6개월"] = res?.cardUpdate?.sixMonths ?? (result === "6개월");
      updatedCard["방문금지"] = res?.cardUpdate?.banned ?? (result === "방문금지");
      state.selectedCard = updatedCard;
    }
    if (res.complete) {
      await loadData();
    }
    renderAreas();
    renderCards();
    renderAdminPanel();
    if (isEdit) {
      setStatus("방문내역이 수정되었습니다.");
      state.editingVisit = null;
    } else {
      if (res.complete) {
        setStatus("모든 카드의 최근방문일이 기록되었습니다. 완료내역이 업데이트되었습니다.");
      } else {
        setStatus("방문내역이 기록되었습니다.");
      }
    }
  } finally {
    setLoading(false);
  }
};

elements.saveApiUrl.addEventListener("click", saveApiUrl);
elements.loginButton.addEventListener("click", login);
elements.startService.addEventListener("click", startService);
elements.closeAreas.addEventListener("click", () => {
  elements.areaOverlay.classList.add("hidden");
});
elements.cardList.addEventListener("click", (event) => {
  event.stopPropagation();
});
elements.visitForm.addEventListener("click", (event) => {
  event.stopPropagation();
});
elements.searchButton.addEventListener("click", () => {
  state.searchQuery = elements.searchInput.value;
  renderCards();
});
elements.searchInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    elements.searchButton.click();
  }
});
elements.filterArea.addEventListener("change", () => {
  state.filterArea = elements.filterArea.value;
  if (state.filterArea === "all") {
    state.selectedArea = null;
    state.expandedAreaId = null;
  } else {
    state.selectedArea = state.filterArea;
    state.expandedAreaId = state.filterArea;
  }
  renderAreas();
  renderCards();
});
elements.filterVisit.addEventListener("change", () => {
  state.filterVisit = elements.filterVisit.value;
  renderCards();
});
elements.visitForm.addEventListener("submit", saveVisit);

loadApiUrl();
