const DEFAULT_API_URL =
  (typeof window !== "undefined" &&
    window.__ENV__ &&
    window.__ENV__.API_URL) ||
  "https://script.google.com/macros/s/AKfycbx1O_Ab6n1j2py4A6Qck48fSY8N_J1wTiZF97y09HzW21kHgCinR1K2rCWiZmONG8mJ/exec";

const SUPABASE_URL = (window.__ENV__ && window.__ENV__.SUPABASE_URL) || "";
const SUPABASE_KEY = (window.__ENV__ && window.__ENV__.SUPABASE_KEY) || "";

let supabaseClient = null;
if (typeof window.supabase !== "undefined" && SUPABASE_URL && SUPABASE_KEY) {
  supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);
} else if (typeof supabasejs !== "undefined" && SUPABASE_URL && SUPABASE_KEY) {
  supabaseClient = supabasejs.createClient(SUPABASE_URL, SUPABASE_KEY);
}

const state = {
  apiUrl: DEFAULT_API_URL,
  user: null,
  data: {
    cards: [],
    areas: [],
    completions: [],
    visits: [],
    evangelists: [],
    assignments: [],
    deletedCards: []
  },
  selectedArea: null,
  selectedCard: null,
  expandedAreaId: null,
  isSuperAdmin: false,
  view: "cards",
  sortOrder: "number",
  searchQuery: "",
  filterArea: "all",
  filterStatus: "all",
  filterVisit: "all",
  adminCardsAreaId: null,
  adminCardsSelectedKey: null,
  adminBannedAreaId: null,
  adminBannedSelectedKey: null,
  completionExpandedAreaId: null,
  editingVisit: null,
  statusTimer: null,
  scrollToSelectedCard: false,
  scrollAreaToTop: false,
  participantsToday: [],
  carAssignments: [],
  selectedCards: [],
  currentMenu: "cards",
  inviteCampaign: null,
  inviteStats: null,
  volunteerWeeks: [],
  selectedVolunteerWeekStart: "",
  selectedVolunteerDate: "",
  selectedVolunteerSlot: "오전",
  carAssignDate: "",
  carAssignSlot: "오전",
  carAssignAssignmentsDate: "",
  carAssignAssignmentsAll: [],
  volunteerDayRowScrollLeft: 0
};

let carAssignTapSelection = null;
let carAssignCardSelection = null;
let carAssignCardMultiSelection = [];

const elements = {
  configPanel: document.getElementById("config-panel"),
  loginPanel: document.getElementById("login-panel"),
  dashboard: document.getElementById("dashboard"),
  menuToggle: document.getElementById("menu-toggle"),
  menuClose: document.getElementById("menu-close"),
  sideMenu: document.getElementById("side-menu"),
  appTitle: document.getElementById("app-title"),
  nameInput: document.getElementById("name-input"),
  passwordInput: document.getElementById("password-input"),
  loginButton: document.getElementById("login-button"),
  apiUrlInput: document.getElementById("api-url-input"),
  saveApiUrl: document.getElementById("save-api-url"),
  closeConfig: document.getElementById("close-config"),
  syncToSupabase: document.getElementById("sync-to-supabase"),
  syncToSheets: document.getElementById("sync-to-sheets"),
  userInfo: document.getElementById("user-info"),
  carInfo: document.getElementById("car-info"),
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
  visitClearRevisit: document.getElementById("visit-clear-revisit"),
  visitClearStudy: document.getElementById("visit-clear-study"),
  visitClearSix: document.getElementById("visit-clear-six"),
  visitClearBanned: document.getElementById("visit-clear-banned"),
  statusMessage: document.getElementById("status-message"),
  areaOverlay: document.getElementById("area-overlay"),
  closeAreas: document.getElementById("close-areas"),
  areaListOverlay: document.getElementById("area-list-overlay"),
  completionOverlay: document.getElementById("completion-overlay"),
  closeCompletion: document.getElementById("close-completion"),
  completionAreaList: document.getElementById("completion-area-list"),
  volunteerOverlay: document.getElementById("volunteer-overlay"),
  closeVolunteer: document.getElementById("close-volunteer"),
  volunteerBody: document.getElementById("volunteer-body"),
  carAssignOverlay: document.getElementById("car-assign-overlay"),
  closeCarAssign: document.getElementById("close-car-assign"),
  adminOverlay: document.getElementById("admin-overlay"),
  closeAdmin: document.getElementById("close-admin"),
  carAssignEvangelistList: document.getElementById("car-assign-evangelist-list"),
  carAssignSelected: document.getElementById("car-assign-selected"),
  carAssignMeta: document.getElementById("car-assign-meta"),
  carAssignTodayDisplay: document.getElementById("car-assign-today-display"),
  carAssignDayList: document.getElementById("car-assign-day-list"),
  carAssignSlotList: document.getElementById("car-assign-slot-list"),
  carAssignAuto: document.getElementById("car-assign-auto"),
  carAssignAdd: document.getElementById("car-assign-add"),
  carAssignReset: document.getElementById("car-assign-reset"),
  carAssignSave: document.getElementById("car-assign-save"),
  carAssignAssignCards: document.getElementById("car-assign-assign-cards"),
  carSelectOverlay: document.getElementById("car-select-overlay"),
  closeCarSelect: document.getElementById("close-car-select"),
  carSelectList: document.getElementById("car-select-list"),
  adminPanel: document.getElementById("admin-panel"),
  completionList: document.getElementById("completion-list"),
  visitList: document.getElementById("visit-list"),
  evangelistList: document.getElementById("evangelist-list"),
  carAssignPanel: document.getElementById("car-assign-panel"),
  bannedCardList: document.getElementById("banned-card-list"),
  deletedCardList: document.getElementById("deleted-card-list"),
  searchInput: document.getElementById("search-input"),
  searchButton: document.getElementById("search-button"),
  filterArea: document.getElementById("filter-area"),
  filterVisit: document.getElementById("filter-visit"),
  areaListInline: document.getElementById("area-list-inline"),
  loadingIndicator: document.getElementById("loading-indicator"),
  loadingText: document.getElementById("loading-text"),
  adminEvAdd: document.getElementById("admin-ev-add"),
  adminEvDelete: document.getElementById("admin-ev-delete"),
  inviteOverlay: document.getElementById("invite-overlay"),
  closeInvite: document.getElementById("close-invite"),
  inviteMeta: document.getElementById("invite-meta"),
  inviteStart: document.getElementById("invite-start"),
  inviteStop: document.getElementById("invite-stop"),
  inviteRefresh: document.getElementById("invite-refresh"),
  inviteStats: document.getElementById("invite-stats"),
  backupToExcel: document.getElementById("backup-to-excel")
};

const openCarSelectPopup = (areaId, cardNumbers) => {
  const cars = state.carAssignments || [];
  if (!elements.carSelectOverlay || !elements.carSelectList || !cars.length) {
    return;
  }
  elements.carSelectList.innerHTML = "";
  cars.forEach((car) => {
    const driverName = car.driver ? String(car.driver) : "";
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "car-select-btn";
    btn.textContent = driverName
      ? `차량 ${car.carId} (${driverName})`
      : `차량 ${car.carId}`;
    btn.addEventListener("click", async () => {
      const pairs = cardNumbers.map((cardNumber) => ({
        cardNumber: String(cardNumber),
        carId: String(car.carId || "")
      }));
      const assignDate = state.carAssignDate || todayISO();
      setLoading(true, "카드 차량 배정 중...");
      try {
        if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
        
        for (const pair of pairs) {
          const { error: updateError } = await supabaseClient
            .from("cards")
            .update({ car_id: pair.carId, assignment_date: assignDate })
            .eq("area_id", String(areaId))
            .eq("card_number", String(pair.cardNumber));
            
          if (updateError) throw updateError;
          
          // 로컬 상태 즉시 업데이트
          const localCard = state.data.cards.find(c => 
            String(c["구역번호"]) === String(areaId) && String(c["카드번호"]) === String(pair.cardNumber)
          );
          if (localCard) {
            localCard["차량"] = pair.carId;
            localCard["배정날짜"] = assignDate;
          }
        }
        
        state.selectedCards = [];
        renderAreas();
        renderCards();
        renderAdminPanel();
        renderMyCarInfo();
        setStatus("선택한 카드가 차량에 배정되었습니다.");
        elements.carSelectOverlay.classList.add("hidden");
      } catch (err) {
        console.error("Assign cards to cars error:", err);
        alert("카드 배정에 실패했습니다: " + err.message);
      } finally {
        setLoading(false);
      }
    });
    elements.carSelectList.appendChild(btn);
  });
  elements.carSelectOverlay.classList.remove("hidden");
};

const todayISO = () => new Date().toISOString().slice(0, 10);

const toAssignmentDateText = (value) => {
  const d = parseVisitDate(value);
  if (!d) {
    return "";
  }
  const yy = String(d.getFullYear()).slice(-2);
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yy}/${mm}/${dd}`;
};

const getCardAssignedCarIdForDate = (card, isoDate) => {
  const carId = String((card && card["차량"]) || "");
  if (!carId) {
    return "";
  }
  const cardDateText = toAssignmentDateText(card && card["배정날짜"]);
  const targetDateText = toAssignmentDateText(isoDate);
  if (!cardDateText || !targetDateText || cardDateText !== targetDateText) {
    return "";
  }
  return carId;
};

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
    return "";
  }
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
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

const parseCardNumber = (value) => {
  const raw = String(value || "").trim();
  const mPair = raw.match(/(\d+)\s*-\s*(\d+)/);
  if (mPair) {
    const area = Number(mPair[1]);
    const num = Number(mPair[2]);
    if (!Number.isNaN(area) && !Number.isNaN(num)) {
      return { raw, area, num };
    }
  }
  const mSingle = raw.match(/(\d+)/);
  if (mSingle) {
    const area = Number(mSingle[1]);
    if (!Number.isNaN(area)) {
      return { raw, area, num: null };
    }
  }
  return { raw, area: null, num: null };
};

const compareCardNumbers = (a, b) => {
  const pa = parseCardNumber(a);
  const pb = parseCardNumber(b);
  if (pa.area != null && pb.area != null) {
    if (pa.area !== pb.area) {
      return pa.area - pb.area;
    }
    if (pa.num != null && pb.num != null && pa.num !== pb.num) {
      return pa.num - pb.num;
    }
    if (pa.num != null && pb.num == null) {
      return 1;
    }
    if (pa.num == null && pb.num != null) {
      return -1;
    }
  }
  return pa.raw.localeCompare(pb.raw, "ko-KR");
};

const getCardTownLabel = (card) => {
  const direct = String(card["읍면동"] || "").trim();
  if (direct) {
    return direct;
  }
  const areaId = String(card["구역번호"] || "").trim();
  if (!areaId) {
    return "";
  }
  const areaRow =
    (state.data.areas || []).find((row) => {
      const raw = String(row["구역번호"] || "").trim();
      return raw === areaId || raw.startsWith(areaId + " ");
    }) || null;
  if (!areaRow) {
    return "";
  }
  const rawLabel = String(areaRow["구역번호"] || "");
  const m = rawLabel.match(/\((.+)\)/);
  return m ? m[1].trim() : "";
};

const isKslArea = (areaId) => {
  const id = String(areaId || "").trim();
  if (!id) {
    return false;
  }
  if (id.includes("찾기봉사")) {
    return true;
  }
  const areaRows = state.data.areas || [];
  const row =
    areaRows.find((r) => {
      const raw = String(r["구역번호"] || "").trim();
      if (!raw) {
        return false;
      }
      return raw === id || raw.startsWith(id + " ") || id.startsWith(raw + " ");
    }) || null;
  if (row) {
    const rowHasKsl = Object.keys(row).some((key) => {
      const v = row[key];
      return v != null && String(v).includes("찾기봉사");
    });
    if (rowHasKsl) {
      return true;
    }
  }
  const cards = (state.data.cards || []).filter(
    (card) => String(card["구역번호"] || card.area || "").trim() === id
  );
  if (cards.length) {
    const cardHasKsl = cards.some((card) => {
      const num = String(card["카드번호"] || "").trim();
      const name = String(
        card["구역이름"] || card["구역명"] || card["구역"] || ""
      ).trim();
      return num.includes("찾기") || name.includes("찾기봉사");
    });
    if (cardHasKsl) {
      return true;
    }
  }
  return false;
};

const parseVisitDate = (value) => {
  if (!value) {
    return null;
  }
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }
  const text = String(value).trim();
  const direct = new Date(text);
  if (!Number.isNaN(direct.getTime())) {
    return direct;
  }
  const normalized = text.replace(/\s+/g, "");
  const m = normalized.match(/^(\d{2,4})[./-](\d{1,2})[./-](\d{1,2})$/);
  if (m) {
    const y = m[1].length === 2 ? 2000 + Number(m[1]) : Number(m[1]);
    const mo = Number(m[2]) - 1;
    const d = Number(m[3]);
    const dt = new Date(y, mo, d);
    return Number.isNaN(dt.getTime()) ? null : dt;
  }
  const parsed = new Date(
    normalized.replace(/\./g, "-").replace(/\//g, "-")
  );
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const formatVolunteerDateText = (value) => {
  if (!value) {
    return "";
  }
  const d = parseVisitDate(value);
  if (!d) {
    return "";
  }
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${m}/${day}`;
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

const renderMyCarInfo = () => {
  const box = elements.carInfo;
  if (!box || !state.user) {
    return;
  }
  const rows = state.data.assignments || [];
  const name = state.user.name;
  
  const showNoInfo = () => {
    box.innerHTML = '<div class="car-info-main">오늘 배정된 차량 정보가 없습니다.</div>';
    box.classList.remove("hidden");
  };

  if (!rows.length || !name) {
    showNoInfo();
    return;
  }
  const myRow =
    rows.find(
      (row) =>
        normalizeAssignmentName(row["이름"]) === normalizeAssignmentName(name)
    ) || null;
  if (!myRow) {
    showNoInfo();
    return;
  }
  const carId = String(myRow["차량"] || "");
  if (!carId) {
    showNoInfo();
    return;
  }
  const sameCar = rows.filter(
    (row) => String(row["차량"] || "") === carId
  );
  const driverRow =
    sameCar.find((row) => String(row["역할"] || "") === "운전자") || null;
  const driverName = driverRow ? String(driverRow["이름"] || "") : "";
  const passengerNames = sameCar
    .filter((row) => String(row["역할"] || "") !== "운전자")
    .map((row) => normalizeAssignmentName(row["이름"]))
    .filter(Boolean);
  const passengers = Array.from(new Set(passengerNames));
  const textParts = [];
  textParts.push(`차량 ${carId}`);
  if (driverName) {
    textParts.push(`운전자: ${driverName}`);
  }
  if (passengers.length) {
    textParts.push(passengers.join(", "));
  }
  const lines = [];
  lines.push(`<div class="car-info-main">${textParts.join(" | ")}</div>`);
  
  const cards = (state.data.cards || []).filter(
    (card) => getCardAssignedCarIdForDate(card, todayISO()) === carId
  );
  if (cards.length) {
    const labels = cards
      .slice()
      .sort((a, b) =>
        compareCardNumbers(a["카드번호"] || "", b["카드번호"] || "")
      )
      .map((card) => String(card["카드번호"] || ""));
    lines.push(`<div class="car-info-separator"></div>`);
    lines.push(`<div class="card-info-main">${labels.join(", ")}</div>`);
  }
  box.innerHTML = lines.join("");
  box.classList.remove("hidden");
};

const getEvangelistByName = (name) =>
  (state.data.evangelists || []).find(
    (row) => String(row["이름"] || "") === String(name || "")
  ) || null;

const normalizeAssignmentName = (value) => {
  return String(value || "")
    .replace(/\s*\((오전|오후)\)\s*$/g, "")
    .trim();
};

const sanitizeAssignmentRows = (rows) => {
  return (rows || []).map((row) => {
    if (!row || typeof row !== "object") {
      return row;
    }
    const normalized = Object.assign({}, row);
    normalized["이름"] = normalizeAssignmentName(row["이름"]);
    return normalized;
  });
};

const buildInitialParticipantsFromAssignments = () => {
  const names =
    (state.data.assignments || []).map((row) =>
      normalizeAssignmentName(row["이름"])
    ) || [];
  state.participantsToday = Array.from(new Set(names));
};

const buildCarAssignmentsFromServer = () => {
  const rows = state.data.assignments || [];
  const byCar = {};
  rows.forEach((row) => {
    const carId = String(row["차량"] || "");
    const name = normalizeAssignmentName(row["이름"]);
    const role = String(row["역할"] || "");
    if (!carId || !name) {
      return;
    }
    if (!byCar[carId]) {
      byCar[carId] = {
        carId,
        driver: "",
        capacity: 0,
        members: []
      };
    }
    if (role === "운전자") {
      byCar[carId].driver = name;
      const ev = getEvangelistByName(name);
      const cap = ev && ev["차량"] != null ? Number(ev["차량"]) || 0 : 0;
      byCar[carId].capacity = cap || byCar[carId].capacity || 0;
    }
    if (!byCar[carId].members.includes(name)) {
      byCar[carId].members.push(name);
    }
  });
  state.carAssignments = Object.keys(byCar).map((key) => byCar[key]);
};

const ensureAssignmentState = () => {
  if (!state.participantsToday.length && (state.data.assignments || []).length) {
    buildInitialParticipantsFromAssignments();
  }
  if (!state.carAssignments.length && (state.data.assignments || []).length) {
    buildCarAssignmentsFromServer();
  }
};

const resetCarAssignmentsFromSaved = () => {
  const has = (state.data.assignments || []).length > 0;
  if (has) {
    buildInitialParticipantsFromAssignments();
    buildCarAssignmentsFromServer();
  } else {
    state.participantsToday = [];
    state.carAssignments = [];
  }
};

const autoAssignCars = () => {
  const all = state.participantsToday.slice();
  const evangelists = all
    .map((name) => {
      const found = getEvangelistByName(name);
      if (found) {
        return found;
      }
      return { 이름: name };
    })
    .filter((row) => String(row["이름"] || ""));
  if (!evangelists.length) {
    alert("선택된 전도인이 없습니다.");
    return;
  }
  const driverEv = [];
  const nonDriverEv = [];
  evangelists.forEach((e) => {
    const isDriver = isTrueValue(e["운전자"]);
    const cap = e["차량"] != null ? Number(e["차량"]) || 0 : 0;
    if (isDriver && cap > 0) {
      driverEv.push(e);
    } else {
      nonDriverEv.push(e);
    }
  });
  if (!driverEv.length) {
    alert("운전자로 설정된 전도인이 없습니다.");
    return;
  }
  const sortedDrivers = driverEv.slice().sort((a, b) => {
    const capA = a["차량"] != null ? Number(a["차량"]) || 0 : 0;
    const capB = b["차량"] != null ? Number(b["차량"]) || 0 : 0;
    return capB - capA;
  });
  const activeDrivers = [];
  let capacitySum = 0;
  const totalPeople = evangelists.length;
  sortedDrivers.forEach((e) => {
    if (capacitySum >= totalPeople) {
      nonDriverEv.push(e);
      return;
    }
    const cap = e["차량"] != null ? Number(e["차량"]) || 0 : 0;
    activeDrivers.push(e);
    capacitySum += cap || 0;
  });
  const evByName = {};
  evangelists.forEach((e) => {
    const n = String(e["이름"] || "");
    if (n) {
      evByName[n] = e;
    }
  });
  const cars = activeDrivers.map((d, idx) => {
    const cap = d["차량"] != null ? Number(d["차량"]) || 0 : 0;
    return {
      carId: String(idx + 1),
      driver: d["이름"],
      capacity: cap || 0,
      members: [d["이름"]],
      hasDeaf: isTrueValue(d["농인"]),
      maleCount: String(d["성별"] || "") === "남" ? 1 : 0,
      femaleCount: String(d["성별"] || "") === "여" ? 1 : 0
    };
  });
  const assignedNames = new Set();
  cars.forEach((c) => {
    (c.members || []).forEach((n) => {
      assignedNames.add(String(n));
    });
  });
  const driverSpouseNames = new Set();
  cars.forEach((car) => {
    const driverName = String(car.driver || "");
    if (!driverName) {
      return;
    }
    const ev = evByName[driverName];
    if (!ev) {
      return;
    }
    const spouseName = String(ev["부부"] || "");
    if (!spouseName) {
      return;
    }
    if (driverSpouseNames.has(spouseName)) {
      return;
    }
    const spouseEv = evByName[spouseName];
    if (!spouseEv) {
      return;
    }
    if (!all.includes(spouseName)) {
      return;
    }
    const cap = car.capacity || 0;
    if (cap && car.members.length >= cap) {
      return;
    }
    if (!car.members.includes(spouseName)) {
      car.members.push(spouseName);
      assignedNames.add(spouseName);
      driverSpouseNames.add(spouseName);
      const gender = String(spouseEv["성별"] || "");
      if (gender === "남") {
        car.maleCount += 1;
      } else if (gender === "여") {
        car.femaleCount += 1;
      }
      if (isTrueValue(spouseEv["농인"])) {
        car.hasDeaf = true;
      }
    }
  });
  const makePerson = (row) => ({
    name: row["이름"],
    isDeaf: isTrueValue(row["농인"]),
    gender: String(row["성별"] || ""),
    spouse: String(row["부부"] || "")
  });
  const others = evangelists.filter(
    (e) =>
      !activeDrivers.includes(e) &&
      !driverSpouseNames.has(String(e["이름"] || ""))
  );
  const nonDriverPeople = others.map(makePerson);
  const deafPeople = nonDriverPeople.filter((p) => p.isDeaf);
  const normalPeople = nonDriverPeople.filter((p) => !p.isDeaf);
  const assignToCar = (person) => {
    if (!person || !person.name) {
      return;
    }
    if (assignedNames.has(String(person.name))) {
      return;
    }
    const spouseName = String(person.spouse || "");
    const spouseEv = spouseName ? evByName[spouseName] : null;
    const spouseAvailable =
      spouseEv &&
      all.includes(spouseName) &&
      !assignedNames.has(spouseName) &&
      others.includes(spouseEv);
    const requiredSeats = spouseAvailable ? 2 : 1;
    const candidates = cars.filter(
      (c) =>
        c.capacity === 0 ||
        (c.members.length + requiredSeats <= c.capacity)
    );
    if (!candidates.length) {
      return;
    }
    let target = null;
    if (person.isDeaf) {
      const preferred = candidates.filter((c) => !c.hasDeaf);
      const list = preferred.length ? preferred : candidates;
      target = list.reduce((best, c) => {
        if (!best) return c;
        const seatsBest =
          (best.capacity || 0) - (best.members.length || 0);
        const seatsC = (c.capacity || 0) - (c.members.length || 0);
        return seatsC > seatsBest ? c : best;
      }, null);
    } else if (person.gender === "여") {
      const preferred = candidates.filter(
        (c) => c.femaleCount >= 1 || c.maleCount === 0
      );
      const list = preferred.length ? preferred : candidates;
      target = list.reduce((best, c) => {
        if (!best) return c;
        const seatsBest =
          (best.capacity || 0) - (best.members.length || 0);
        const seatsC = (c.capacity || 0) - (c.members.length || 0);
        return seatsC > seatsBest ? c : best;
      }, null);
    } else {
      target = candidates.reduce((best, c) => {
        if (!best) return c;
        const seatsBest =
          (best.capacity || 0) - (best.members.length || 0);
        const seatsC = (c.capacity || 0) - (c.members.length || 0);
        return seatsC > seatsBest ? c : best;
      }, null);
    }
    if (!target) {
      return;
    }
    const addPersonToCar = (name, evRow, info) => {
      if (!name) {
        return;
      }
      if (assignedNames.has(String(name))) {
        return;
      }
      if (!target.members.includes(name)) {
        target.members.push(name);
      }
      assignedNames.add(String(name));
      const gender = String(info.gender || "");
      if (gender === "남") {
        target.maleCount += 1;
      } else if (gender === "여") {
        target.femaleCount += 1;
      }
      if (info.isDeaf) {
        target.hasDeaf = true;
      }
    };
    addPersonToCar(person.name, evByName[person.name] || null, person);
    if (spouseAvailable) {
      const spouseInfo = makePerson(spouseEv);
      addPersonToCar(spouseName, spouseEv, spouseInfo);
    }
  };
  deafPeople.forEach(assignToCar);
  normalPeople.forEach(assignToCar);
  const seenNames = new Set();
  cars.forEach((car) => {
    const driverName = String(car.driver || "");
    const uniqueMembers = [];
    const localSeen = new Set();
    if (driverName) {
      if (!seenNames.has(driverName)) {
        uniqueMembers.push(driverName);
        seenNames.add(driverName);
        localSeen.add(driverName);
      } else {
        car.driver = "";
      }
    }
    (car.members || []).forEach((name) => {
      const key = String(name || "");
      if (!key || key === driverName) {
        return;
      }
      if (localSeen.has(key) || seenNames.has(key)) {
        return;
      }
      localSeen.add(key);
      seenNames.add(key);
      uniqueMembers.push(key);
    });
    car.members = uniqueMembers;
  });
  state.carAssignments = cars.map((c) => ({
    carId: c.carId,
    driver: c.driver,
    capacity: c.capacity,
    members: c.members || []
  }));
  renderCarAssignmentsPanel();
};

const saveCarAssignments = async () => {
  if (!supabaseClient) return;
  const validCars =
    state.carAssignments.filter((car) => {
      const hasDriver = Boolean(car.driver);
      const hasMembers = (car.members || []).some((name) => Boolean(name));
      return hasDriver || hasMembers;
    }) || [];

  const currentDate = state.carAssignDate || todayISO();
  const currentSlot = state.carAssignSlot || "오전";

  const newAssignments = [];
  validCars.forEach((car) => {
    newAssignments.push({
      date: currentDate,
      slot: currentSlot,
      car_id: String(car.carId),
      driver: car.driver || "",
      passengers: car.members || []
    });
  });

  setLoading(true, "차량 배정 저장 중...");
  try {
    // 1. 기존 해당 날짜/시간대 배정 삭제
    const { error: deleteError } = await supabaseClient
      .from("car_assignments")
      .delete()
      .eq("date", currentDate)
      .eq("slot", currentSlot);

    if (deleteError) throw deleteError;

    // 2. 새 배정 저장
    if (newAssignments.length > 0) {
      const { error: insertError } = await supabaseClient
        .from("car_assignments")
        .insert(newAssignments);
      if (insertError) throw insertError;
    }

    // 3. 로컬 상태 업데이트를 위해 다시 로드
    const { data: allAssignments, error: fetchError } = await supabaseClient
      .from("car_assignments")
      .select("*")
      .eq("date", currentDate);

    if (fetchError) throw fetchError;

    state.carAssignAssignmentsDate = currentDate;
    state.carAssignAssignmentsAll = (allAssignments || []).map(a => ({
      "날짜": a.date,
      "시간대": a.slot,
      "차량": a.car_id,
      "이름": a.driver,
      "동승자": a.passengers || []
    }));

    // 4. 구역카드 차량 배정 정보 업데이트
    const cards = state.data.cards || [];
    const toUpdateCards = [];
    cards.forEach((card) => {
      const carId = getCardAssignedCarIdForDate(card, currentDate);
      if (carId) {
        toUpdateCards.push({
          area_id: String(card["구역번호"]),
          card_number: String(card["카드번호"]),
          car_id: String(carId),
          assignment_date: currentDate
        });
      }
    });

    if (toUpdateCards.length > 0) {
      for (const cardData of toUpdateCards) {
        await supabaseClient
          .from("cards")
          .update({ car_id: cardData.car_id, assignment_date: cardData.assignment_date })
          .eq("area_id", cardData.area_id)
          .eq("card_number", cardData.card_number);
      }
    }

    applyCarAssignDataForSlot(currentDate, currentSlot);
    renderAreas();
    renderCards();
    renderAdminPanel();
    renderMyCarInfo();
    setStatus("차량 배정이 저장되었습니다.");
    renderCarAssignPopup();
  } catch (err) {
    console.error("Save car assignments error:", err);
    alert("차량 배정 저장에 실패했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

const resetCarAssignmentsOnServer = async () => {
  setLoading(true, "차량 배정 초기화 중...");
  try {
    const targetDate = state.carAssignDate || state.selectedVolunteerDate || todayISO();
    const targetSlot = state.carAssignSlot || "오전";
    
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
    
    // 1. 해당 날짜/시간대의 차량 배정 삭제
    const { error: deleteError } = await supabaseClient
      .from("car_assignments")
      .delete()
      .eq("date", targetDate)
      .eq("slot", targetSlot);
      
    if (deleteError) throw deleteError;
    
    // 2. 해당 날짜에 배정된 카드의 차량 정보 초기화
    const { error: updateError } = await supabaseClient
      .from("cards")
      .update({ car_id: null, assignment_date: null })
      .eq("assignment_date", targetDate);
      
    if (updateError) throw updateError;
    
    // 3. 데이터 다시 불러오기 및 UI 갱신
    await refreshAll();
    setStatus("차량 배정이 초기화되었습니다.");
  } catch (err) {
    console.error("Reset car assignments error:", err);
    alert("차량 배정 초기화에 실패했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

const renderCarAssignDayRow = () => {
  const container = elements.carAssignDayList;
  if (!container) {
    return;
  }
  
  if (elements.carAssignTodayDisplay) {
    const weekdayFullMap = {
      월: "월요일", 화: "화요일", 수: "수요일", 목: "목요일",
      금: "금요일", 토: "토요일", 일: "일요일"
    };
    const todayDate = new Date();
    const todayYear = todayDate.getFullYear();
    const todayMonth = todayDate.getMonth() + 1;
    const todayDay = todayDate.getDate();
    const todayWeekday = weekdayFullMap[["일", "월", "화", "수", "목", "금", "토"][todayDate.getDay()]] || "";
    elements.carAssignTodayDisplay.innerHTML = `<div class="volunteer-today-text">오늘은</div><div class="volunteer-today-date">${todayYear}년 ${todayMonth}월 ${todayDay}일 ${todayWeekday}</div>`;
  }

  container.innerHTML = "";
  const days = buildVolunteerDayList();
  const current = state.carAssignDate || todayISO();
  if (!days.length) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "weekday-chip active";
    btn.textContent = "오늘";
    btn.addEventListener("click", () => {
      state.carAssignDate = todayISO();
    });
    container.appendChild(btn);
    return;
  }
  
  let activeBtn = null;
  days.forEach((d) => {
    if (!d.isoDate) {
      return;
    }
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "weekday-chip";
    if (d.isoDate === current) {
      btn.classList.add("active");
      activeBtn = btn;
    }
    const weekdayShortMap = {
      월: "월",
      화: "화",
      수: "수",
      목: "목",
      금: "금",
      토: "토",
      일: "일"
    };
    const dateLabel = formatVolunteerDateText(d.dateText || d.isoDate);
    const weekdayRaw = String(d.weekday || "").trim();
    const weekdayLabel = weekdayRaw ? weekdayShortMap[weekdayRaw] || weekdayRaw : "";
    btn.innerHTML = `<span class="weekday-chip-date">${dateLabel}</span><span class="weekday-chip-weekday">${weekdayLabel}</span>`;
    btn.dataset.isoDate = d.isoDate;
    btn.addEventListener("click", () => {
      if (state.carAssignDate === d.isoDate) {
        return;
      }
      setCarAssignDate(d.isoDate);
    });
    container.appendChild(btn);
  });

  if (activeBtn) {
    requestAnimationFrame(() => {
      activeBtn.scrollIntoView({ behavior: "auto", inline: "start", block: "nearest" });
    });
  }
};

const renderCarAssignSlotRow = () => {
  const container = elements.carAssignSlotList;
  if (!container) {
    return;
  }
  container.innerHTML = "";
  const days = buildVolunteerDayList();
  const currentDate = state.carAssignDate || todayISO();
  const day = days.find((d) => d.isoDate === currentDate) || null;
  const slots = [];
  if (!day || day.slotAMEnabled !== false) {
    slots.push("오전");
  }
  if (!day || day.slotPMEnabled !== false) {
    slots.push("오후");
  }
  const availableSlots = slots.length ? slots : ["오전"];
  if (availableSlots.indexOf(state.carAssignSlot) === -1) {
    state.carAssignSlot = availableSlots[0];
  }
  availableSlots.forEach((slot) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "car-assign-slot-btn";
    btn.classList.add(slot === "오후" ? "slot-pm" : "slot-am");
    if (slot === state.carAssignSlot) {
      btn.classList.add("active");
    }
    btn.textContent = slot;
    btn.dataset.slot = slot;
    btn.addEventListener("click", () => {
      if (state.carAssignSlot === slot) {
        return;
      }
      setCarAssignSlot(slot);
    });
    container.appendChild(btn);
  });
};

const applyCarAssignDataForSlot = (date, slot, dayOverride) => {
  const selectedSlot = slot === "오후" ? "오후" : "오전";
  const day =
    dayOverride || buildVolunteerDayList().find((d) => d.isoDate === date) || null;
  const allRows =
    state.carAssignAssignmentsDate === date
      ? state.carAssignAssignmentsAll || []
      : [];
  const filteredRows = allRows.filter((row) => {
    const rowSlot = String(row["시간대"] || "").trim() || "오전";
    return rowSlot === selectedSlot;
  });
  state.data.assignments = sanitizeAssignmentRows(filteredRows);
  const signupNames = day
    ? (day.participantEntries || [])
        .filter((entry) => String(entry.slot || "오전") === selectedSlot)
        .map((entry) => normalizeAssignmentName(entry.name))
    : [];
  const assignmentNames =
    (state.data.assignments || []).map((row) =>
      normalizeAssignmentName(row["이름"])
    ) || [];
  const baseNames = []
    .concat(signupNames)
    .concat(assignmentNames)
    .filter((name) => name);
  state.participantsToday = Array.from(new Set(baseNames));
  buildCarAssignmentsFromServer();
  renderCarAssignDayRow();
  renderCarAssignSlotRow();
  renderCarAssignPopup();
};

const setCarAssignSlot = async (slotName) => {
  const normalizedSlot = slotName === "오후" ? "오후" : "오전";
  state.carAssignSlot = normalizedSlot;
  const targetDate = state.carAssignDate || todayISO();
  if (state.carAssignAssignmentsDate === targetDate) {
    applyCarAssignDataForSlot(targetDate, normalizedSlot);
    return;
  }
  await setCarAssignDate(targetDate);
};

const setCarAssignDate = async (isoDate) => {
  const date = getNearestVolunteerDateISO(isoDate || todayISO());
  state.carAssignDate = date;
  setLoading(true, "차량 배정 정보를 불러오는 중...");
  try {
    await loadVolunteerConfig();
    const days = buildVolunteerDayList();
    const day =
      days.find((d) => d.isoDate === date) || null;
    const availableSlots = [];
    if (!day || day.slotAMEnabled !== false) {
      availableSlots.push("오전");
    }
    if (!day || day.slotPMEnabled !== false) {
      availableSlots.push("오후");
    }
    if (!availableSlots.length) {
      availableSlots.push("오전");
    }
    if (availableSlots.indexOf(state.carAssignSlot) === -1) {
      state.carAssignSlot = availableSlots[0];
    }
    const selectedSlot = state.carAssignSlot || "오전";

    if (!supabaseClient) {
      throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
    }

    const { data: assignments, error } = await supabaseClient
      .from("car_assignments")
      .select("*")
      .eq("date", date);

    if (error) throw error;

    state.carAssignAssignmentsDate = date;
    state.carAssignAssignmentsAll = (assignments || []).map(a => ({
      "날짜": a.date,
      "시간대": a.slot,
      "차량": a.car_id,
      "이름": a.driver,
      "동승자": a.passengers || []
    }));
    
    applyCarAssignDataForSlot(date, selectedSlot, day);
  } catch (err) {
    console.error("Failed to fetch car assignments:", err);
    alert("차량 배정 정보를 불러오는 데 실패했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

const renderCarAssignmentsPanel = () => {
  const panel = elements.carAssignPanel;
  if (!panel) {
    return;
  }
  if (!state.user || state.user.role === "전도인") {
    panel.innerHTML = "";
    return;
  }
  panel.innerHTML = "";
  const grid = document.createElement("div");
  grid.className = "car-assign-grid";
  const cars = state.carAssignments || [];
  const allCards = state.data.cards || [];
  cars.forEach((car) => {
    const driverName = car.driver ? String(car.driver) : "";
    const col = document.createElement("div");
    col.className = "car-column";
    col.dataset.carId = car.carId;
    const header = document.createElement("div");
    header.className = "car-column-header";
    header.innerHTML = driverName
      ? `차량 ${car.carId}<br />(${driverName})`
      : `차량 ${car.carId}`;
    header.addEventListener("dblclick", async () => {
      let areaId = state.selectedArea || state.filterArea;
      if (!areaId || areaId === "all") {
        const input = window.prompt("배정할 구역번호를 입력해 주세요.");
        if (!input) return;
        areaId = input.trim();
      }
      const cardNumber = window.prompt("배정할 카드번호를 입력해 주세요.");
      if (!cardNumber) return;
      setLoading(true, "카드 수동 배정 중...");
      const assignDate = state.carAssignDate || todayISO();
      try {
        if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
        
        const { error: updateError } = await supabaseClient
          .from("cards")
          .update({ car_id: String(car.carId || ""), assignment_date: assignDate })
          .eq("area_id", String(areaId))
          .eq("card_number", String(cardNumber));
          
        if (updateError) throw updateError;
        
        // 로컬 상태 즉시 업데이트
        const localCard = state.data.cards.find(c => 
          String(c["구역번호"]) === String(areaId) && String(c["카드번호"]) === String(cardNumber)
        );
        if (localCard) {
          localCard["차량"] = String(car.carId || "");
          localCard["배정날짜"] = assignDate;
        }

        renderAreas();
        renderCards();
        renderAdminPanel();
        renderMyCarInfo();
        setStatus(`카드 ${String(cardNumber)}가 차량 ${String(car.carId)}에 배정되었습니다.`);
      } catch (err) {
        console.error("Assign cards to cars error:", err);
        alert("카드 배정에 실패했습니다: " + err.message);
      } finally {
        setLoading(false);
      }
    });
    const membersBox = document.createElement("div");
    membersBox.className = "car-members";
    membersBox.dataset.carId = car.carId;
    (car.members || []).forEach((name) => {
      const item = document.createElement("div");
      item.className = "car-member";
      item.draggable = true;
      item.dataset.name = name;
      item.dataset.carId = car.carId;
      item.textContent = name;
      membersBox.appendChild(item);
    });
    const cardTagsBox = document.createElement("div");
    cardTagsBox.className = "car-card-list";
    const assignedCards = allCards.filter(
      (card) =>
        getCardAssignedCarIdForDate(
          card,
          state.carAssignDate || todayISO()
        ) === String(car.carId || "")
    );
    if (assignedCards.length) {
      const label = document.createElement("div");
      label.className = "car-card-list-label";
      label.textContent = "배정된 카드";
      cardTagsBox.appendChild(label);
      const tagsRow = document.createElement("div");
      tagsRow.className = "car-card-list-tags";
      assignedCards.forEach((card) => {
        const span = document.createElement("span");
        span.className = "car-card-tag";
        const cardNumber = String(card["카드번호"] || "");
        const townText = getCardTownLabel(card);
        const tagLabel = townText ? `${cardNumber} (${townText})` : cardNumber;
        const areaId = String(card["구역번호"] || "");
        span.textContent = tagLabel;
        span.dataset.cardNumber = cardNumber;
        span.dataset.areaId = areaId;
        span.dataset.carId = String(car.carId || "");

        let cardPressTimer = null;
        const clearCardPressTimer = () => {
          if (cardPressTimer) {
            clearTimeout(cardPressTimer);
            cardPressTimer = null;
          }
        };
        const startCardPressTimer = () => {
          clearCardPressTimer();
          cardPressTimer = setTimeout(() => {
            cardPressTimer = null;
            const area = span.dataset.areaId || "";
            const num = span.dataset.cardNumber || "";
            const fromCar = span.dataset.carId || "";
            if (!area || !num || !fromCar) {
              return;
            }
            const ok = window.confirm(
              `차량 ${fromCar}에서 카드 ${num} 배정을 삭제할까요?`
            );
            if (!ok) {
              return;
            }
            const cardsAll = state.data.cards || [];
            const target = cardsAll.find(
              (card) =>
                String(card["구역번호"] || "") === String(area) &&
                String(card["카드번호"] || "") === String(num)
            );
            if (target) {
              target["차량"] = "";
              target["배정날짜"] = "";
            }
            renderCards();
            renderCarAssignPopup();
          }, 800);
        };

        span.addEventListener("mousedown", (event) => {
          if (event.button !== 0) {
            return;
          }
          startCardPressTimer();
        });
        span.addEventListener("touchstart", () => {
          startCardPressTimer();
        });
        ["mouseup", "mouseleave", "touchend", "touchcancel"].forEach(
          (type) => {
            span.addEventListener(type, clearCardPressTimer);
          }
        );

        tagsRow.appendChild(span);
      });
      cardTagsBox.appendChild(tagsRow);
    }
    let pressTimer = null;
    const clearPressTimer = () => {
      if (pressTimer) {
        clearTimeout(pressTimer);
        pressTimer = null;
      }
    };
    const startPressTimer = () => {
      clearPressTimer();
      pressTimer = setTimeout(() => {
        pressTimer = null;
        const hasDriver = Boolean(car.driver);
        const hasMembers = (car.members || []).some((name) => Boolean(name));
        if (hasDriver || hasMembers) {
          return;
        }
        if (
          window.confirm(
            `차량 ${car.carId}에 배정된 사람이 없습니다.\n이 차량을 삭제하시겠습니까?`
          )
        ) {
          state.carAssignments = (state.carAssignments || []).filter(
            (c) => String(c.carId) !== String(car.carId)
          );
          renderCarAssignmentsPanel();
        }
      }, 800);
    };
    col.addEventListener("mousedown", (event) => {
      if (event.button !== 0) {
        return;
      }
      startPressTimer();
    });
    col.addEventListener("touchstart", () => {
      startPressTimer();
    });
    ["mouseup", "mouseleave", "touchend", "touchcancel"].forEach((type) => {
      col.addEventListener(type, clearPressTimer);
    });
    col.append(header, membersBox);
    col.appendChild(cardTagsBox);
    grid.appendChild(col);
  });
  panel.appendChild(grid);
};

const renderSelectedParticipants = () => {
  const box = elements.carAssignSelected;
  if (!box) {
    return;
  }
  const names = state.participantsToday.slice();
  box.innerHTML = "";
  if (!names.length) {
    return;
  }
  const assigned = new Set();
  (state.carAssignments || []).forEach((car) => {
    if (car.driver) {
      assigned.add(normalizeAssignmentName(car.driver));
    }
    (car.members || []).forEach((name) => {
      if (name) {
        assigned.add(normalizeAssignmentName(name));
      }
    });
  });
  const unassigned = names.filter(
    (name) => !assigned.has(normalizeAssignmentName(name))
  );
  const frag = document.createDocumentFragment();
  if (!unassigned.length) {
    box.appendChild(frag);
    return;
  }
  unassigned.forEach((name) => {
    const item = document.createElement("div");
    item.className = "selected-person";
    item.textContent = name;
    frag.appendChild(item);
  });
  box.appendChild(frag);
};

const formatAssignmentDate = (value) => {
  if (!value) {
    return "";
  }
  if (typeof value === "string") {
    if (/^\d{2}\/\d{2}\/\d{2}$/.test(value)) {
      return value;
    }
    const isoMatch = value.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) {
      const yy = isoMatch[1].slice(-2);
      const mm = isoMatch[2];
      const dd = isoMatch[3];
      return yy + "/" + mm + "/" + dd;
    }
  }
  try {
    const d = new Date(value);
    if (!isNaN(d.getTime())) {
      const yy = String(d.getFullYear()).slice(-2);
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const dd = String(d.getDate()).padStart(2, "0");
      return yy + "/" + mm + "/" + dd;
    }
  } catch (e) {
  }
  return String(value);
};

const renderCarAssignPopup = () => {
  if (!elements.carAssignOverlay || !state.user) {
    return;
  }
  const userRole = String(state.user.role || "").trim();
  const isLeader = userRole === "관리자" || userRole === "인도자";
  if (!isLeader) {
    elements.carAssignOverlay.classList.add("hidden");
    return;
  }
  const listEl = elements.carAssignEvangelistList;
  if (listEl) {
    const list = (state.data.evangelists || [])
      .slice()
      .sort((a, b) =>
        String(a["이름"] || "").localeCompare(
          String(b["이름"] || ""),
          "ko-KR"
        )
      );
    const itemsHtml = list
      .map((row) => {
        const name = row["이름"] || "";
        const isParticipant = state.participantsToday.includes(
          normalizeAssignmentName(name)
        );
        return `<div class="ev-item${
          isParticipant ? " selected" : ""
        }" data-name="${name}">${name}</div>`;
      })
      .join("");
    const tempHtml =
      '<div class="ev-item ev-temp-add" data-role="temp-add">+ 추가</div>';
    listEl.innerHTML = itemsHtml + tempHtml;
  }
  const meta = elements.carAssignMeta;
  renderCarAssignDayRow();
  renderCarAssignSlotRow();
  if (meta) {
    const rows = state.data.assignments || [];
    if (rows.length) {
      const date =
        rows[0]["날짜"] ||
        rows[0]["date"] ||
        rows[0]["Date"] ||
        "";
      const dateText = formatAssignmentDate(date);
      const slot = String(rows[0]["시간대"] || "").trim();
      const slotText = slot ? ` (${slot})` : "";
      meta.textContent = dateText
        ? `배정된 날짜: ${dateText}${slotText}`
        : "오늘 차량 배정이 저장되어 있습니다.";
    } else {
      meta.textContent = "아직 저장된 차량 배정이 없습니다. 배정 후 저장해 주세요.";
    }
  }
  renderSelectedParticipants();
  renderCarAssignmentsPanel();
};

const loadApiUrl = () => {
  let url = DEFAULT_API_URL;
  try {
    if (typeof window !== "undefined" && window.localStorage) {
      const stored = window.localStorage.getItem("mc_api_url");
      if (stored) {
        url = stored;
      }
    }
  } catch (e) {}
  state.apiUrl = url;
  if (elements.apiUrlInput) {
    elements.apiUrlInput.value = url;
  }
  elements.configPanel.classList.add("hidden");
};

const saveApiUrl = () => {
  let url = DEFAULT_API_URL;
  if (elements.apiUrlInput) {
    const value = elements.apiUrlInput.value.trim();
    if (value) {
      url = value;
    }
  }
  state.apiUrl = url;
  try {
    if (typeof window !== "undefined" && window.localStorage) {
      window.localStorage.setItem("mc_api_url", url);
    }
  } catch (e) {}
  elements.configPanel.classList.add("hidden");
};

const refreshAll = async () => {
  setLoading(true, "데이터 새로고침 중...");
  try {
    await loadData();
    renderAreas();
    renderCards();
    renderAdminPanel();
    renderMyCarInfo();
    renderVolunteerOverlay();
    if (state.currentMenu === "visits" && state.completionExpandedAreaId) {
      renderVisitsView();
    }
    if (state.currentMenu === "car-assign") {
      await setCarAssignDate(state.carAssignDate || todayISO());
    }
  } finally {
    setLoading(false);
  }
};

const saveVolunteerWeekToSupabase = async (weekStart, weekData) => {
  if (!supabaseClient) return { success: false, message: "Supabase client not initialized" };
  try {
    // weekStart는 ISO 날짜 형식 (YYYY-MM-DD)
    // weekData는 JSON 객체 (데이터 필드에 저장될 내용)
    const { error } = await supabaseClient
      .from("volunteer_weeks")
      .upsert({
        week_start: weekStart,
        data: weekData
      });

    if (error) throw error;
    return { success: true };
  } catch (err) {
    console.error("Failed to save volunteer week:", err);
    return { success: false, message: err.message };
  }
};

const updateCardFlagsInSupabase = async (areaId, cardNumber, flags) => {
  if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
  
  const payload = {};
  if (flags.hasOwnProperty("revisit")) payload.revisit = flags.revisit;
  if (flags.hasOwnProperty("study")) payload.study = flags.study;
  if (flags.hasOwnProperty("sixMonths")) payload.six_months = flags.sixMonths;
  if (flags.hasOwnProperty("banned")) payload.banned = flags.banned;
  if (flags.hasOwnProperty("invite")) payload.invite = flags.invite;
  
  const { error } = await supabaseClient
    .from("cards")
    .update(payload)
    .eq("area_id", String(areaId))
    .eq("card_number", String(cardNumber));
    
  if (error) throw error;
  return { success: true, ...flags };
};

const deleteCardInSupabase = async (areaId, cardNumber) => {
  if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

  // 1. 기존 카드 정보 조회
  const { data: card, error: selectError } = await supabaseClient
    .from("cards")
    .select("*")
    .eq("area_id", String(areaId))
    .eq("card_number", String(cardNumber))
    .single();

  if (selectError) {
    console.warn("Card not found in Supabase or already deleted:", selectError);
    // 이미 Supabase에 없어도 계속 진행할 수 있도록 무시하거나 에러 처리
  }

  if (card) {
    // 2. 삭제된 카드 테이블로 이동 (보관)
    const { error: insertError } = await supabaseClient
      .from("deleted_cards")
      .insert({
        area_id: String(card.area_id),
        card_number: String(card.card_number),
        address: card.address,
        deleted_at: new Date().toISOString()
      });

    if (insertError) {
      console.error("Failed to archive deleted card:", insertError);
      throw insertError;
    }

    // 3. 원본 테이블에서 삭제
    const { error: deleteError } = await supabaseClient
      .from("cards")
      .delete()
      .eq("area_id", String(areaId))
      .eq("card_number", String(cardNumber));

    if (deleteError) {
      console.error("Failed to delete card from cards table:", deleteError);
      throw deleteError;
    }
  }

  return { success: true };
};

const restoreDeletedCardInSupabase = async (areaId, cardNumber) => {
  if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

  // 1. 삭제된 카드 정보 조회
  const { data: deletedCard, error: selectError } = await supabaseClient
    .from("deleted_cards")
    .select("*")
    .eq("area_id", String(areaId))
    .eq("card_number", String(cardNumber))
    .single();

  if (selectError) throw selectError;

  if (deletedCard) {
    // 2. cards 테이블로 복구 (기본 정보만)
    const { error: insertError } = await supabaseClient
      .from("cards")
      .upsert({
        area_id: String(deletedCard.area_id),
        card_number: String(deletedCard.card_number),
        address: deletedCard.address
      }, { onConflict: "area_id, card_number" });

    if (insertError) throw insertError;

    // 3. deleted_cards 테이블에서 삭제
    const { error: deleteError } = await supabaseClient
      .from("deleted_cards")
      .delete()
      .eq("area_id", String(areaId))
      .eq("card_number", String(cardNumber));

    if (deleteError) throw deleteError;
  }

  return { success: true };
};

const purgeDeletedCardInSupabase = async (areaId, cardNumber) => {
  if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

  const { error } = await supabaseClient
    .from("deleted_cards")
    .delete()
    .eq("area_id", String(areaId))
    .eq("card_number", String(cardNumber));

  if (error) throw error;
  return { success: true };
};

const apiRequest = async (action, payload = {}, method = "POST") => {
  if (!state.apiUrl) {
    if (state.isSuperAdmin) {
      elements.configPanel.classList.remove("hidden");
    }
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

const buildVolunteerDayList = () => {
  const weeks = state.volunteerWeeks || [];
  const days = [];
  weeks.forEach((week) => {
    const weekStartText = week.weekStartText || "";
    const weekStartISO = week.weekStartISO || "";
    (week.days || []).forEach((day) => {
      days.push({
        weekStartText,
        weekStartISO,
        dateText: day.dateText || "",
        isoDate: day.isoDate || "",
        weekday: day.weekday || "",
        participants: day.participants || [],
        participantEntries: day.participantEntries || [],
        slotAMEnabled: day.slotAMEnabled,
        slotPMEnabled: day.slotPMEnabled,
        fixedMemo: day.fixedMemo || "",
        extraMemo: day.extraMemo || "",
        fixedMemoAM: day.fixedMemoAM || "",
        fixedMemoPM: day.fixedMemoPM || "",
        extraMemoAM: day.extraMemoAM || "",
        extraMemoPM: day.extraMemoPM || ""
      });
    });
  });
  days.sort((a, b) => {
    const da = a.isoDate || "";
    const db = b.isoDate || "";
    return da.localeCompare(db);
  });
  return days;
};

const getNearestVolunteerDateISO = (baseDate) => {
  const days = buildVolunteerDayList().filter((d) => Boolean(d.isoDate));
  if (!days.length) {
    return baseDate || todayISO();
  }
  const requested = String(baseDate || "").trim();
  if (requested) {
    const exact = days.find((d) => d.isoDate === requested);
    if (exact) {
      return exact.isoDate;
    }
  }
  const base = parseVisitDate(requested || todayISO()) || new Date();
  const baseTime = new Date(
    base.getFullYear(),
    base.getMonth(),
    base.getDate()
  ).getTime();
  const sorted = days
    .slice()
    .sort((a, b) => {
      const da = parseVisitDate(a.isoDate);
      const db = parseVisitDate(b.isoDate);
      const ta = da
        ? new Date(da.getFullYear(), da.getMonth(), da.getDate()).getTime()
        : baseTime;
      const tb = db
        ? new Date(db.getFullYear(), db.getMonth(), db.getDate()).getTime()
        : baseTime;
      const diffA = Math.abs(ta - baseTime);
      const diffB = Math.abs(tb - baseTime);
      if (diffA !== diffB) {
        return diffA - diffB;
      }
      return String(a.isoDate || "").localeCompare(String(b.isoDate || ""));
    });
  return sorted[0].isoDate || todayISO();
};

const ensureVolunteerSelection = () => {
  const weeks = state.volunteerWeeks || [];
  if (!weeks.length) {
    state.selectedVolunteerWeekStart = "";
    state.selectedVolunteerDate = "";
    state.selectedVolunteerSlot = "오전";
    return;
  }
  const allDays = buildVolunteerDayList();
  if (!allDays.length) {
    state.selectedVolunteerWeekStart = "";
    state.selectedVolunteerDate = "";
    state.selectedVolunteerSlot = "오전";
    return;
  }
  const today = todayISO();
  const futureDays = allDays.filter(
    (d) => d.isoDate && d.isoDate >= today
  );
  const days = futureDays.length ? futureDays : allDays;
  if (state.selectedVolunteerDate) {
    const found = days.find(
      (d) => d.isoDate === state.selectedVolunteerDate
    );
    if (found) {
      state.selectedVolunteerWeekStart = found.weekStartText || "";
      return;
    }
  }
  const todayDay = days.find((d) => d.isoDate === today);
  const target = todayDay || days[0];
  state.selectedVolunteerDate = target.isoDate || "";
  state.selectedVolunteerWeekStart = target.weekStartText || "";
  if (!state.selectedVolunteerSlot) {
    state.selectedVolunteerSlot = "오전";
  }
};

const loadVolunteerConfig = async () => {
  if (!supabaseClient) return;
  try {
    const { data, error } = await supabaseClient
      .from("volunteer_weeks")
      .select("*")
      .order("week_start", { ascending: true });

    if (error) throw error;

    if (data && data.length > 0) {
      state.volunteerWeeks = data.map(row => ({
        ...row.data,
        weekStartText: row.week_start
      }));
    } else {
      state.volunteerWeeks = [];
    }
    ensureVolunteerSelection();
  } catch (err) {
    console.error("Failed to load volunteer config:", err);
  }
};

const renderVolunteerOverlay = () => {
  const overlay = elements.volunteerOverlay;
  const body = elements.volunteerBody;
  if (!overlay || !body || !state.user) {
    return;
  }
  const prevDayRow = body.querySelector(".weekday-scroll");
  if (prevDayRow) {
    state.volunteerDayRowScrollLeft = prevDayRow.scrollLeft || 0;
  }
  const weeks = state.volunteerWeeks || [];
  body.innerHTML = "";
  const userRole = String(state.user.role || "").trim();
  const isAdmin = userRole === "관리자";
  const isLeader = userRole === "인도자";
  const canManageVolunteerSignups = isAdmin || isLeader;

  const allDays = buildVolunteerDayList();
  ensureVolunteerSelection();

  if (!weeks.length && !isAdmin) {
    const msg = document.createElement("div");
    msg.className = "volunteer-empty-msg";
    msg.textContent = "등록된 봉사 주간이 없습니다.";
    body.appendChild(msg);
    return;
  }

  const today = todayISO();
  const selectableDays = allDays.filter((d) => d.isoDate && d.isoDate >= today);
  const days = selectableDays.length ? selectableDays : allDays;
  
  let selectedDay = allDays.find((d) => d.isoDate === state.selectedVolunteerDate) || null;
  if (!selectedDay && allDays.length) {
    selectedDay = allDays[0];
    state.selectedVolunteerDate = selectedDay.isoDate || "";
    state.selectedVolunteerWeekStart = selectedDay.weekStartText || "";
  }

  const getAvailableVolunteerSlots = (day) => {
    if (!day) return ["오전"];
    const slots = [];
    if (day.slotAMEnabled !== false) slots.push("오전");
    if (day.slotPMEnabled !== false) slots.push("오후");
    return slots.length ? slots : ["오전"];
  };

  const availableSlots = getAvailableVolunteerSlots(selectedDay);
  if (availableSlots.indexOf(state.selectedVolunteerSlot) === -1) {
    state.selectedVolunteerSlot = availableSlots[0];
  }

  const selectedWeek = (selectedDay && weeks.length)
    ? weeks.find(w => String(w.weekStartText || "") === String(selectedDay.weekStartText || "")) || null
    : null;

  const layout = document.createElement("div");
  layout.className = "volunteer-layout";
  
  const mainBox = document.createElement("div");
  mainBox.className = "volunteer-main";
  layout.appendChild(mainBox);

  if (weeks.length === 0) {
    const emptyMsg = document.createElement("div");
    emptyMsg.className = "volunteer-empty-msg";
    emptyMsg.textContent = "등록된 봉사 주간이 없습니다. 하단 관리자 메뉴에서 주간을 등록해 주세요.";
    mainBox.appendChild(emptyMsg);
  } else if (allDays.length === 0) {
    const emptyMsg = document.createElement("div");
    emptyMsg.className = "volunteer-empty-msg";
    emptyMsg.textContent = "등록된 봉사 날짜가 없습니다.";
    mainBox.appendChild(emptyMsg);
  } else {
    // Standard rendering logic
    const header = document.createElement("div");
    header.className = "volunteer-week-header";
    const range = document.createElement("div");
    range.className = "volunteer-week-range";
    range.textContent = "";
    header.appendChild(range);
    mainBox.appendChild(header);

    const grid = document.createElement("div");
    grid.className = "volunteer-main-grid";
    const dayCard = document.createElement("div");
    dayCard.className = "volunteer-card volunteer-card-days";
    
    const bindLongPress = (element, onLongPress) => {
      let timer = null;
      let longPressed = false;
      const start = (event) => {
        if (event) event.stopPropagation();
        longPressed = false;
        timer = setTimeout(async () => {
          longPressed = true;
          timer = null;
          await onLongPress();
        }, 700);
      };
      const cancel = () => {
        if (timer) {
          clearTimeout(timer);
          timer = null;
        }
      };
      element.addEventListener("mousedown", (event) => {
        if (event.button !== 0) return;
        start(event);
      });
      element.addEventListener("touchstart", start);
      ["mouseup", "mouseleave", "touchend", "touchcancel"].forEach((type) => {
        element.addEventListener(type, cancel);
      });
      element.addEventListener("click", (event) => {
        if (!longPressed) return;
        event.preventDefault();
        event.stopPropagation();
        longPressed = false;
      });
    };

    const weekdayFullMap = {
      월: "월요일", 화: "화요일", 수: "수요일", 목: "목요일",
      금: "금요일", 토: "토요일", 일: "일요일"
    };

    const todayDate = new Date();
    const todayYear = todayDate.getFullYear();
    const todayMonth = todayDate.getMonth() + 1;
    const todayDay = todayDate.getDate();
    const todayWeekday = weekdayFullMap[["일", "월", "화", "수", "목", "금", "토"][todayDate.getDay()]] || "";

    const todayLabel = document.createElement("div");
    todayLabel.className = "volunteer-today-label";
    todayLabel.innerHTML = `<div class="volunteer-today-text">오늘은</div><div class="volunteer-today-date">${todayYear}년 ${todayMonth}월 ${todayDay}일 ${todayWeekday}</div>`;
    mainBox.appendChild(todayLabel);

    const dayRow = document.createElement("div");
    dayRow.className = "weekday-scroll";
    
    if (days.length) {
      days.forEach((d) => {
        if (!d.isoDate) return;
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "weekday-chip";
        if (d.isoDate === state.selectedVolunteerDate) {
          btn.classList.add("active");
        }
        const dateLabel = formatVolunteerDateText(d.dateText || d.isoDate);
        const weekdayRaw = String(d.weekday || "").trim();
        const weekdayLabel = weekdayRaw ? (weekdayFullMap[weekdayRaw] || `${weekdayRaw}요일`).replace("요일", "") : "";
        btn.innerHTML = `<span class="weekday-chip-date">${dateLabel}</span><span class="weekday-chip-weekday">${weekdayLabel}</span>`;
        btn.dataset.isoDate = d.isoDate;
        btn.addEventListener("click", () => {
          state.selectedVolunteerDate = d.isoDate || "";
          state.selectedVolunteerWeekStart = d.weekStartText || "";
          const slots = getAvailableVolunteerSlots(d);
          if (slots.indexOf(state.selectedVolunteerSlot) === -1) {
            state.selectedVolunteerSlot = slots[0] || "오전";
          }
          renderVolunteerOverlay();
        });
        if (isAdmin) {
          bindLongPress(btn, async () => {
            const ok = window.confirm(`${dateLabel} ${weekdayLabel}의 오전/오후 신청 전체를 삭제할까요?`);
            if (!ok) return;
            setLoading(true, "날짜를 삭제하는 중...");
            try {
              const weekStart = d.weekStartText || d.weekStartISO;
              const { data: weekRow, error: fetchError } = await supabaseClient
                .from("volunteer_weeks")
                .select("data")
                .eq("week_start", weekStart)
                .single();
                
              if (fetchError) throw fetchError;
              
              const updatedData = { ...weekRow.data };
              updatedData.days = (updatedData.days || []).filter(day => day.isoDate !== d.isoDate);
              
              const saveRes = await saveVolunteerWeekToSupabase(weekStart, updatedData);
              if (!saveRes.success) throw new Error(saveRes.message);
              
              await loadVolunteerConfig();
              ensureVolunteerSelection();
              renderVolunteerOverlay();
            } catch (err) {
              console.error("Remove volunteer day error:", err);
              alert("날짜 삭제에 실패했습니다: " + err.message);
            } finally {
              setLoading(false);
            }
          });
        }
        dayRow.appendChild(btn);
      });
    } else {
      const weekdayNames = ["월", "화", "수", "목", "금", "토", "일"];
      weekdayNames.forEach((name) => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "weekday-chip";
        btn.textContent = name;
        dayRow.appendChild(btn);
      });
    }
    dayCard.appendChild(dayRow);
    dayRow.addEventListener(
      "scroll",
      () => {
        state.volunteerDayRowScrollLeft = dayRow.scrollLeft || 0;
      },
      { passive: true }
    );
    const savedVolunteerDayRowScrollLeft = Number(state.volunteerDayRowScrollLeft) || 0;
    requestAnimationFrame(() => {
      dayRow.scrollLeft = savedVolunteerDayRowScrollLeft;
    });
    grid.appendChild(dayCard);

    const selectedCard = document.createElement("div");
    selectedCard.className = "volunteer-card volunteer-card-selected";
    
    if (selectedDay) {
      const dateText = formatVolunteerDateText(selectedDay.dateText || selectedDay.isoDate);
      const weekdayRaw = String(selectedDay.weekday || "").trim();
      const weekdayLabel = weekdayRaw ? (weekdayFullMap[weekdayRaw] || `${weekdayRaw}요일`) : "";
      
      const daySlots = getAvailableVolunteerSlots(selectedDay);
      daySlots.forEach((slotName) => {
        const fixedMemoText = String(
          slotName === "오후" ? selectedDay.fixedMemoPM || selectedDay.fixedMemo || "" : selectedDay.fixedMemoAM || selectedDay.fixedMemo || ""
        ).trim();
        const extraMemoText = String(
          slotName === "오후" ? selectedDay.extraMemoPM || selectedDay.extraMemo || "" : selectedDay.extraMemoAM || selectedDay.extraMemo || ""
        ).trim();
        const memoParts = [];
        if (fixedMemoText) memoParts.push(fixedMemoText);
        if (extraMemoText) memoParts.push(extraMemoText);
        const memoText = memoParts.join("\n");
        const section = document.createElement("div");
        section.className = "volunteer-slot-section";
        const slotHeader = document.createElement("div");
        slotHeader.className = "volunteer-slot-header";
        slotHeader.textContent = `${dateText} ${weekdayLabel} ${slotName}`;
        section.appendChild(slotHeader);
        if (memoText) {
          const slotMemo = document.createElement("div");
          slotMemo.className = "volunteer-slot-memo";
          slotMemo.textContent = memoText;
          section.appendChild(slotMemo);
        }
        const memoDivider = document.createElement("div");
        memoDivider.className = "volunteer-slot-memo-divider";
        section.appendChild(memoDivider);
        const list = document.createElement("div");
        list.className = "volunteer-participant-list";
        const entries = (selectedDay.participantEntries || []).filter(e => (e.slot || "오전") === slotName);
        if (entries.length) {
          entries.forEach((entry) => {
            const entryName = String(entry.name || "");
            const chip = document.createElement("div");
            chip.className = "volunteer-participant";
            chip.textContent = entryName;
            if (canManageVolunteerSignups) {
              bindLongPress(chip, async () => {
                const ok = window.confirm(`${entryName} (${slotName}) 신청자를 삭제할까요?`);
                if (!ok) return;
                setLoading(true, "신청자를 삭제하는 중...");
                try {
                  const weekStart = selectedDay.weekStartText || selectedDay.weekStartISO;
                  const { data: weekRow, error: fetchError } = await supabaseClient
                    .from("volunteer_weeks")
                    .select("data")
                    .eq("week_start", weekStart)
                    .single();
                    
                  if (fetchError) throw fetchError;
                  
                  const updatedData = { ...weekRow.data };
                  const dayObj = (updatedData.days || []).find(d => d.isoDate === selectedDay.isoDate);
                  if (dayObj && dayObj.participantEntries) {
                    dayObj.participantEntries = dayObj.participantEntries.filter(e => !(String(e.name || "") === String(entryName) && (e.slot || "오전") === slotName));
                  }
                  
                  const saveRes = await saveVolunteerWeekToSupabase(weekStart, updatedData);
                  if (!saveRes.success) throw new Error(saveRes.message);
                  
                  await loadVolunteerConfig();
                  ensureVolunteerSelection();
                  renderVolunteerOverlay();
                } catch (err) {
                  console.error("Remove volunteer by admin error:", err);
                  alert("신청자 삭제에 실패했습니다: " + err.message);
                } finally {
                  setLoading(false);
                }
              });
            }
            list.appendChild(chip);
          });
        } else {
          const none = document.createElement("div");
          none.className = "volunteer-none";
          none.textContent = "아직 신청한 사람이 없습니다.";
          list.appendChild(none);
        }
        section.appendChild(list);
        if (state.user) {
          const isoDate = selectedDay.isoDate || "";
          if (isoDate && isoDate >= today) {
            const currentName = state.user.name;
            const already = entries.some(e => String(e.name || "") === String(currentName));
            const actionBtn = document.createElement("button");
            actionBtn.type = "button";
            actionBtn.className = "volunteer-slot-action-btn";
            if (already) actionBtn.classList.add("cancel");
            actionBtn.textContent = already ? "취소하기" : "신청하기";
            actionBtn.addEventListener("click", async () => {
              const applying = !already;
              setLoading(true, applying ? "봉사 신청 중..." : "봉사 신청 취소 중...");
              try {
                const weekStart = selectedDay.weekStartText || selectedDay.weekStartISO;
                const { data: weekRow, error: fetchError } = await supabaseClient
                  .from("volunteer_weeks")
                  .select("data")
                  .eq("week_start", weekStart)
                  .single();
                  
                if (fetchError) throw fetchError;
                
                const updatedData = { ...weekRow.data };
                const dayObj = updatedData.days.find(d => d.isoDate === selectedDay.isoDate);
                if (!dayObj) throw new Error("날짜 정보를 찾을 수 없습니다.");
                
                if (!dayObj.participantEntries) dayObj.participantEntries = [];
                
                if (applying) {
                  const exists = dayObj.participantEntries.some(e => String(e.name || "") === String(state.user.name) && (e.slot || "오전") === slotName);
                  if (!exists) dayObj.participantEntries.push({ name: state.user.name, slot: slotName });
                } else {
                  dayObj.participantEntries = dayObj.participantEntries.filter(e => !(String(e.name || "") === String(state.user.name) && (e.slot || "오전") === slotName));
                }
                
                const saveRes = await saveVolunteerWeekToSupabase(weekStart, updatedData);
                if (!saveRes.success) throw new Error(saveRes.message);
                
                await loadVolunteerConfig();
                renderVolunteerOverlay();
                setStatus(applying ? "봉사 신청이 등록되었습니다." : "봉사 신청이 취소되었습니다.");
              } catch (err) {
                console.error("Volunteer apply/cancel error:", err);
                alert("작업에 실패했습니다: " + err.message);
              } finally {
                setLoading(false);
              }
            });
            section.appendChild(actionBtn);
          }
        }
        if (canManageVolunteerSignups) {
          const addBtn = document.createElement("button");
          addBtn.type = "button";
          addBtn.className = "volunteer-admin-add-btn";
          addBtn.textContent = "+ 신청자 추가/취소";
          addBtn.addEventListener("click", async () => {
            const evangelists = (state.data.evangelists || []).map((row) => String(row["이름"] || "").trim()).filter((name) => !!name);
            const uniqueEvangelists = Array.from(new Set(evangelists));
            const initialSelectedNames = new Set(
              entries.map((entry) => String(entry.name || "").trim()).filter((name) => !!name)
            );
            const allSelectableNames = Array.from(
              new Set(uniqueEvangelists.concat(Array.from(initialSelectedNames)))
            ).filter((name) => !!name);
            const workingSelectedNames = new Set(initialSelectedNames);
            const existingPicker = section.querySelector(".volunteer-admin-add-picker");
            if (existingPicker) {
              existingPicker.remove();
              return;
            }
            const picker = document.createElement("div");
            picker.className = "volunteer-admin-add-picker";
            const title = document.createElement("div");
            title.className = "volunteer-admin-add-picker-title";
            title.textContent = `${slotName} 신청자 선택 (선택=신청됨, 해제=취소)`;
            picker.appendChild(title);
            const customRow = document.createElement("div");
            customRow.className = "volunteer-admin-add-custom-row";
            const customInput = document.createElement("input");
            customInput.type = "text";
            customInput.className = "volunteer-admin-add-custom-input";
            customInput.placeholder = "이름 입력";
            const customAddBtn = document.createElement("button");
            customAddBtn.type = "button";
            customAddBtn.className = "volunteer-admin-add-custom-btn";
            customAddBtn.textContent = "+추가";
            customRow.appendChild(customInput);
            customRow.appendChild(customAddBtn);
            picker.appendChild(customRow);
            const listWrap = document.createElement("div");
            listWrap.className = "volunteer-admin-add-list";
            const renderNameButtons = () => {
              listWrap.innerHTML = "";
              const orderedNames = allSelectableNames
                .slice()
                .sort((a, b) => a.localeCompare(b, "ko-KR"));
              orderedNames.forEach((name) => {
              const nameBtn = document.createElement("button");
              nameBtn.type = "button";
              nameBtn.className = "volunteer-admin-add-name";
              nameBtn.classList.add(slotName === "오후" ? "slot-pm" : "slot-am");
              if (workingSelectedNames.has(name)) {
                nameBtn.classList.add("selected");
              }
              nameBtn.textContent = name;
              nameBtn.addEventListener("click", () => {
                if (workingSelectedNames.has(name)) {
                  workingSelectedNames.delete(name);
                  nameBtn.classList.remove("selected");
                } else {
                  workingSelectedNames.add(name);
                  nameBtn.classList.add("selected");
                }
              });
              listWrap.appendChild(nameBtn);
              });
            };
            renderNameButtons();
            const addCustomName = () => {
              const customName = String(customInput.value || "").trim();
              if (!customName) {
                return;
              }
              if (!allSelectableNames.includes(customName)) {
                allSelectableNames.push(customName);
              }
              workingSelectedNames.add(customName);
              customInput.value = "";
              renderNameButtons();
            };
            customAddBtn.addEventListener("click", addCustomName);
            customInput.addEventListener("keydown", (event) => {
              if (event.key === "Enter") {
                event.preventDefault();
                addCustomName();
              }
            });
            picker.appendChild(listWrap);
            const saveBtn = document.createElement("button");
            saveBtn.type = "button";
            saveBtn.className = "volunteer-admin-add-save";
            saveBtn.textContent = "저장";
            saveBtn.addEventListener("click", async () => {
              const namesToAdd = [];
              const namesToRemove = [];
              allSelectableNames.forEach((name) => {
                const before = initialSelectedNames.has(name);
                const after = workingSelectedNames.has(name);
                if (!before && after) {
                  namesToAdd.push(name);
                } else if (before && !after) {
                  namesToRemove.push(name);
                }
              });
              if (!namesToAdd.length && !namesToRemove.length) {
                picker.remove();
                return;
              }
              setLoading(true, "신청자 변경사항을 저장하는 중...");
              try {
                const weekStart = selectedDay.weekStartText || selectedDay.weekStartISO;
                const { data: weekRow, error: fetchError } = await supabaseClient
                  .from("volunteer_weeks")
                  .select("data")
                  .eq("week_start", weekStart)
                  .single();
                  
                if (fetchError) throw fetchError;
                
                const updatedData = { ...weekRow.data };
                const dayObj = (updatedData.days || []).find(d => d.isoDate === selectedDay.isoDate);
                if (dayObj) {
                  if (!dayObj.participantEntries) dayObj.participantEntries = [];
                  
                  // Remove
                  namesToRemove.forEach(name => {
                    dayObj.participantEntries = dayObj.participantEntries.filter(e => !(String(e.name || "") === String(name) && (e.slot || "오전") === slotName));
                  });
                  
                  // Add
                  namesToAdd.forEach(name => {
                    const exists = dayObj.participantEntries.some(e => String(e.name || "") === String(name) && (e.slot || "오전") === slotName);
                    if (!exists) dayObj.participantEntries.push({ name, slot: slotName });
                  });
                }
                
                const saveRes = await saveVolunteerWeekToSupabase(weekStart, updatedData);
                if (!saveRes.success) throw new Error(saveRes.message);
                
                await loadVolunteerConfig();
                ensureVolunteerSelection();
                renderVolunteerOverlay();
                setStatus("신청자 변경사항이 저장되었습니다.");
              } catch (err) {
                console.error("Save volunteer by admin error:", err);
                alert("변경사항 저장에 실패했습니다: " + err.message);
              } finally {
                setLoading(false);
              }
            });
            const closeBtn = document.createElement("button");
            closeBtn.type = "button";
            closeBtn.className = "volunteer-admin-add-close";
            closeBtn.textContent = "닫기";
            closeBtn.addEventListener("click", () => {
              picker.remove();
            });
            const actionRow = document.createElement("div");
            actionRow.className = "volunteer-admin-add-actions";
            actionRow.appendChild(saveBtn);
            actionRow.appendChild(closeBtn);
            picker.appendChild(actionRow);
            section.appendChild(picker);
          });
          section.appendChild(addBtn);
        }
        selectedCard.appendChild(section);
      });
    } else {
      const none = document.createElement("div");
      none.className = "volunteer-selected-date";
      none.textContent = "날짜를 선택해 주세요";
      selectedCard.appendChild(none);
    }
    grid.appendChild(selectedCard);
    mainBox.appendChild(grid);
  }

  if (isAdmin) {
    const adminBox = document.createElement("div");
    adminBox.className = "volunteer-admin";
    const title = document.createElement("div");
    title.textContent = "관리자 메뉴";
    adminBox.appendChild(title);
    const weekRow = document.createElement("div");
    weekRow.className = "volunteer-admin-row";
    const weekLabel = document.createElement("span");
    weekLabel.className = "volunteer-admin-label";
    weekLabel.textContent = "주 선택";
    weekRow.appendChild(weekLabel);
    const weekSelect = document.createElement("select");
    const sortedWeeks = (weeks || []).slice().sort((a, b) => {
      const aISO = a.weekStartISO || "";
      const bISO = b.weekStartISO || "";
      return aISO.localeCompare(bISO);
    });
    sortedWeeks.forEach((w) => {
      const opt = document.createElement("option");
      const ds = w.days || [];
      let label = "";
      if (ds.length) {
        const first = ds[0];
        const last = ds[ds.length - 1];
        label =
          `${formatVolunteerDateText(first.dateText || first.isoDate)} ~ ` +
          `${formatVolunteerDateText(last.dateText || last.isoDate)}`;
      } else {
        label = formatVolunteerDateText(
          w.weekStartText || w.weekStartISO || ""
        );
      }
      opt.value = w.weekStartText || "";
      opt.textContent = label;
      if (
        String(w.weekStartText || "") ===
        String(state.selectedVolunteerWeekStart || "")
      ) {
        opt.selected = true;
      }
      weekSelect.appendChild(opt);
    });
    weekSelect.addEventListener("change", () => {
      const value = weekSelect.value;
      state.selectedVolunteerWeekStart = value;
      const target =
        days.find((d) => String(d.weekStartText || "") === String(value)) ||
        null;
      if (target) {
        state.selectedVolunteerDate = target.isoDate || "";
      }
      renderVolunteerOverlay();
    });
    weekRow.appendChild(weekSelect);
    adminBox.appendChild(weekRow);
    const row1 = document.createElement("div");
    row1.className = "volunteer-admin-row";
    const dateLabel = document.createElement("span");
    dateLabel.className = "volunteer-admin-label";
    dateLabel.textContent = "주 시작";
    row1.appendChild(dateLabel);
    const dateInput = document.createElement("input");
    dateInput.type = "date";
   dateInput.id = "volunteer-week-start";
   dateInput.name = "volunteer-week-start";
    let baseISO =
      (selectedWeek && selectedWeek.weekStartISO) ||
      (selectedDay && selectedDay.isoDate) ||
      todayISO();
    try {
      const d = new Date(baseISO);
      if (!Number.isNaN(d.getTime())) {
        const day = (d.getDay() + 6) % 7;
        d.setDate(d.getDate() - day);
        baseISO = d.toISOString().slice(0, 10);
      }
    } catch (e) {}
    dateInput.value = baseISO;
    row1.appendChild(dateInput);
    adminBox.appendChild(row1);
    const row2 = document.createElement("div");
    row2.className = "volunteer-admin-row";
    const dayButtonsContainer = document.createElement("div");
    dayButtonsContainer.className = "volunteer-day-grid-admin";
    const weekdayNames = ["월", "화", "수", "목", "금", "토", "일"];
    
    // slotsConfig holds { [dayIndex]: { am: boolean, pm: boolean } }
    const slotsConfig = {};
    for (let i = 1; i <= 7; i++) {
      slotsConfig[i] = { am: false, pm: false };
    }

    const weekByISO =
      weeks.find((w) => String(w.weekStartISO || "") === String(baseISO)) ||
      null;
    if (weekByISO && (weekByISO.days || []).length) {
      try {
        const baseDate = new Date(baseISO);
        const ms = 24 * 60 * 60 * 1000;
        weekByISO.days.forEach((d) => {
          const dd = new Date(d.isoDate);
          if (Number.isNaN(dd.getTime())) {
            return;
          }
          const diff = Math.round((dd.getTime() - baseDate.getTime()) / ms);
          const idx = diff + 1;
          if (idx >= 1 && idx <= 7) {
            slotsConfig[idx] = {
              am: d.slotAMEnabled !== false,
              pm: d.slotPMEnabled === true
            };
          }
        });
      } catch (e) {}
    }

    for (let i = 1; i <= 7; i++) {
      const col = document.createElement("div");
      col.className = "volunteer-admin-day-col";
      
      const label = document.createElement("div");
      label.className = "volunteer-admin-day-label";
      label.textContent = weekdayNames[i - 1];
      col.appendChild(label);

      const amBtn = document.createElement("button");
      amBtn.type = "button";
      amBtn.className = "volunteer-slot-toggle-btn";
      amBtn.classList.add("slot-am");
      amBtn.textContent = "오전";
      if (slotsConfig[i].am) amBtn.classList.add("active");
      amBtn.addEventListener("click", () => {
        slotsConfig[i].am = !slotsConfig[i].am;
        amBtn.classList.toggle("active");
      });
      col.appendChild(amBtn);

      const pmBtn = document.createElement("button");
      pmBtn.type = "button";
      pmBtn.className = "volunteer-slot-toggle-btn";
      pmBtn.classList.add("slot-pm");
      pmBtn.textContent = "오후";
      if (slotsConfig[i].pm) pmBtn.classList.add("active");
      pmBtn.addEventListener("click", () => {
        slotsConfig[i].pm = !slotsConfig[i].pm;
        pmBtn.classList.toggle("active");
      });
      col.appendChild(pmBtn);

      dayButtonsContainer.appendChild(col);
    }
    row2.appendChild(dayButtonsContainer);
    adminBox.appendChild(row2);

    const row3 = document.createElement("div");
    row3.className = "volunteer-admin-row volunteer-admin-row-center";
    const saveBtn = document.createElement("button");
    saveBtn.type = "button";
    saveBtn.textContent = "주간/요일 저장";
    saveBtn.addEventListener("click", async () => {
      const startDate = dateInput.value;
      if (!startDate) {
        alert("주 시작일을 선택해 주세요.");
        return;
      }
      
      const activeSlots = {};
      let hasAny = false;
      for (let i = 1; i <= 7; i++) {
        if (slotsConfig[i].am || slotsConfig[i].pm) {
          activeSlots[i] = slotsConfig[i];
          hasAny = true;
        }
      }

      if (!hasAny) {
        alert("봉사 가능한 요일과 시간대를 한 개 이상 선택해 주세요.");
        return;
      }

      setLoading(true, "봉사 가능 요일을 저장하는 중...");
      try {
        const weekStartDate = parseVisitDate(startDate);
        if (!weekStartDate) throw new Error("Invalid start date");
        
        // Find Monday of that week
        const day = weekStartDate.getDay();
        const diff = (day + 6) % 7;
        const monday = new Date(weekStartDate.getTime() - diff * 24 * 60 * 60 * 1000);
        const mondayISO = toISODate(monday);
        
        const weekdayOrder = ["월", "화", "수", "목", "금", "토", "일"];
        const newDays = [];
        
        for (let i = 0; i < 7; i++) {
          const current = new Date(monday.getTime() + i * 24 * 60 * 60 * 1000);
          const weekday = weekdayOrder[i];
          const config = activeSlots.find(s => s.weekday === weekday);
          
          if (config && (config.am || config.pm)) {
            newDays.push({
              isoDate: toISODate(current),
              dateText: formatVolunteerDateText(current),
              weekday: weekday,
              slotAMEnabled: config.am,
              slotPMEnabled: config.pm,
              participantEntries: []
            });
          }
        }
        
        const { data: existingWeek, error: fetchError } = await supabaseClient
          .from("volunteer_weeks")
          .select("data")
          .eq("week_start", mondayISO)
          .single();
          
        let finalDays = newDays;
        if (!fetchError && existingWeek && existingWeek.data && existingWeek.data.days) {
          finalDays = newDays.map(newDay => {
            const oldDay = existingWeek.data.days.find(d => d.isoDate === newDay.isoDate);
            if (oldDay) {
              return { 
                ...oldDay, 
                slotAMEnabled: newDay.slotAMEnabled, 
                slotPMEnabled: newDay.slotPMEnabled 
              };
            }
            return newDay;
          });
        }
        
        const weekData = { days: finalDays };
        const saveRes = await saveVolunteerWeekToSupabase(mondayISO, weekData);
        if (!saveRes.success) throw new Error(saveRes.message);
        
        await loadVolunteerConfig();
        ensureVolunteerSelection();
        renderVolunteerOverlay();
        setStatus("봉사 가능 요일이 저장되었습니다.");
      } catch (err) {
        console.error("Set volunteer week error:", err);
        alert("저장에 실패했습니다: " + err.message);
      } finally {
        setLoading(false);
      }
    });
    row3.appendChild(saveBtn);
    adminBox.appendChild(row3);
    if (selectedDay && selectedDay.isoDate) {
      const memoSlots = getAvailableVolunteerSlots(selectedDay);
      memoSlots.forEach((slotName) => {
        const memoBox = document.createElement("div");
        memoBox.className = "volunteer-admin";
        const memoRow = document.createElement("div");
        memoRow.className = "volunteer-admin-row";
        const slotLabel = document.createElement("span");
        slotLabel.className = "volunteer-admin-label";
        slotLabel.textContent = `${slotName} 메모`;
        memoRow.appendChild(slotLabel);
        memoBox.appendChild(memoRow);
        const fixedRow = document.createElement("div");
        fixedRow.className = "volunteer-admin-row";
        const fixedLabel = document.createElement("span");
        fixedLabel.className = "volunteer-admin-label";
        fixedLabel.textContent = "요일 메모";
        fixedRow.appendChild(fixedLabel);
        const fixedMemoInput = document.createElement("input");
        fixedMemoInput.type = "text";
        fixedMemoInput.placeholder = `${selectedDay.weekday || ""} 공통 메모`;
        fixedMemoInput.value = String(slotName === "오후" ? selectedDay.fixedMemoPM || "" : selectedDay.fixedMemoAM || "");
        fixedRow.appendChild(fixedMemoInput);
        memoBox.appendChild(fixedRow);
        const extraRow = document.createElement("div");
        extraRow.className = "volunteer-admin-row";
        const extraLabel = document.createElement("span");
        extraLabel.className = "volunteer-admin-label";
        extraLabel.textContent = "날짜 메모";
        extraRow.appendChild(extraLabel);
        const extraMemoInput = document.createElement("input");
        extraMemoInput.type = "text";
        extraMemoInput.placeholder = `${slotName} 추가 메모`;
        extraMemoInput.value = String(slotName === "오후" ? selectedDay.extraMemoPM || "" : selectedDay.extraMemoAM || "");
        extraRow.appendChild(extraMemoInput);
        memoBox.appendChild(extraRow);
        const actionRow = document.createElement("div");
        actionRow.className = "volunteer-admin-row volunteer-admin-row-center";
        const memoSaveBtn = document.createElement("button");
        memoSaveBtn.type = "button";
        memoSaveBtn.textContent = "저장";
        memoSaveBtn.addEventListener("click", async () => {
          setLoading(true, "메모를 저장하는 중...");
          try {
            const weekStart = selectedDay.weekStartText || selectedDay.weekStartISO;
            const { data: weekRow, error: fetchError } = await supabaseClient
              .from("volunteer_weeks")
              .select("data")
              .eq("week_start", weekStart)
              .single();
              
            if (fetchError) throw fetchError;
            
            const updatedData = { ...weekRow.data };
            const dayObj = (updatedData.days || []).find(d => d.isoDate === selectedDay.isoDate);
            if (dayObj) {
              if (slotName === "오후") {
                dayObj.fixedMemoPM = fixedMemoInput.value || "";
                dayObj.extraMemoPM = extraMemoInput.value || "";
              } else {
                dayObj.fixedMemoAM = fixedMemoInput.value || "";
                dayObj.extraMemoAM = extraMemoInput.value || "";
              }
            }
            
            const saveRes = await saveVolunteerWeekToSupabase(weekStart, updatedData);
            if (!saveRes.success) throw new Error(saveRes.message);
            
            await loadVolunteerConfig();
            ensureVolunteerSelection();
            renderVolunteerOverlay();
            setStatus("메모가 저장되었습니다.");
          } catch (err) {
            console.error("Set volunteer memo error:", err);
            alert("메모 저장에 실패했습니다: " + err.message);
          } finally {
            setLoading(false);
          }
        });
        actionRow.appendChild(memoSaveBtn);
        memoBox.appendChild(actionRow);
        adminBox.appendChild(memoBox);
      });
    }
    if (selectedWeek && (selectedWeek.days || []).length) {
      const statsTitle = document.createElement("div");
      statsTitle.textContent = "선택 주간 신청 현황";
      adminBox.appendChild(statsTitle);
      const table = document.createElement("table");
      table.className = "volunteer-admin-table";
      const thead = document.createElement("thead");
      const headRow = document.createElement("tr");
      ["날짜", "요일", "오전", "오후"].forEach((text) => {
        const th = document.createElement("th");
        th.textContent = text;
        headRow.appendChild(th);
      });
      thead.appendChild(headRow);
      table.appendChild(thead);
      const tbody = document.createElement("tbody");
      const sortedDays = selectedWeek.days.slice().sort((a, b) => {
        const aISO = a.isoDate || "";
        const bISO = b.isoDate || "";
        return aISO.localeCompare(bISO);
      });
      sortedDays.forEach((d) => {
        const tr = document.createElement("tr");
        const tdDate = document.createElement("td");
        tdDate.textContent = formatVolunteerDateText(d.dateText || d.isoDate);
        tr.appendChild(tdDate);
        const tdWeekday = document.createElement("td");
        tdWeekday.textContent = d.weekday || "";
        tr.appendChild(tdWeekday);
        
        const entries = d.participantEntries || [];
        const amNames = entries.filter(e => (e.slot || "오전") === "오전").map(e => e.name);
        const pmNames = entries.filter(e => e.slot === "오후").map(e => e.name);
        
        const tdAM = document.createElement("td");
        tdAM.textContent = amNames.join(", ");
        tr.appendChild(tdAM);
        
        const tdPM = document.createElement("td");
        tdPM.textContent = pmNames.join(", ");
        tr.appendChild(tdPM);
        
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      adminBox.appendChild(table);
    }
    layout.appendChild(adminBox);
  }
  body.appendChild(layout);
};

const openVolunteerOverlay = async () => {
  if (!elements.volunteerOverlay || !elements.volunteerBody) {
    return;
  }
  await loadVolunteerConfig();
  elements.volunteerOverlay.classList.remove("hidden");
  document.body.style.overflow = "hidden";
  renderVolunteerOverlay();
};

const groupCardsByArea = () => {
  const byArea = {};
  (state.data.cards || []).forEach((card) => {
    const areaId = String(card["구역번호"] || card.area || "");
    if (!areaId) return;
    if (!byArea[areaId]) byArea[areaId] = [];
    byArea[areaId].push(card);
  });
  return byArea;
};

const normalizeVisitDateText = (value) => {
  if (!value) {
    return "";
  }
  const d = toISODate(value);
  if (!d) {
    return "";
  }
  const yy = String(d.getFullYear()).slice(-2);
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yy}/${mm}/${dd}`;
};

const areaCompletionStatus = (cards) =>
  cards.every(
    (card) =>
      Boolean(card["최근방문일"]) ||
      isTrueValue(card["재방"]) ||
      isTrueValue(card["연구"]) ||
      isTrueValue(card["6개월"]) ||
      isTrueValue(card["방문금지"])
  );

const getFirstInProgressArea = () => {
  const grouped = groupCardsByArea();
  const areaIds = Object.keys(grouped);
  for (const areaId of areaIds) {
    if (isKslArea(areaId)) {
      continue;
    }
    if (!areaCompletionStatus(grouped[areaId])) {
      return areaId;
    }
  }
  return null;
};

const collapseExpandedArea = () => {
  state.expandedAreaId = null;
  state.selectedArea = null;
  state.filterArea = "all";
  state.scrollAreaToTop = false;
  if (elements.filterArea) {
    elements.filterArea.value = "all";
  }
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
  const isLeader =
    state.user &&
    (state.user.role === "관리자" || state.user.role === "인도자");
  const areasRows = state.data.areas || [];
  let latestServiceDate = null;
  let hasTodayInProgress = false;
  areasRows.forEach((row) => {
    const startD = row["시작날짜"] ? parseVisitDate(row["시작날짜"]) : null;
    const endD = row["완료날짜"] ? parseVisitDate(row["완료날짜"]) : null;
    const d = endD || startD;
    if (d && !Number.isNaN(d.getTime())) {
      if (!latestServiceDate || d.getTime() > latestServiceDate.getTime()) {
        latestServiceDate = d;
      }
    }
     if (
       row["시작날짜"] &&
       !row["완료날짜"] &&
       isSameDay(row["시작날짜"], today)
     ) {
       hasTodayInProgress = true;
     }
  });
  const areaIds = Object.keys(grouped).sort((a, b) => {
    const sheetAreaOrder = (state.data.areas || []).map(row => String(row["구역번호"]));
    const idxA = sheetAreaOrder.indexOf(String(a));
    const idxB = sheetAreaOrder.indexOf(String(b));
    if (idxA !== -1 && idxB !== -1) return idxA - idxB;
    if (idxA !== -1) return -1;
    if (idxB !== -1) return 1;
    
    const na = Number(a);
    const nb = Number(b);
    if (!Number.isNaN(na) && !Number.isNaN(nb)) {
      return na - nb;
    }
    return String(a).localeCompare(String(b), "ko-KR");
  });
  areaIds.forEach((areaId) => {
    const createItem = (withAssignButton) => {
      const item = document.createElement("div");
      let justSwiped = false;
      item.className = "list-item";
      item.dataset.areaId = areaId;
      if (state.selectedArea === areaId) {
        item.classList.add("active");
      }
      const areaInfo = state.data.areas.find(
        (row) => String(row["구역번호"]) === String(areaId)
      );
      const startText =
        areaInfo && areaInfo["시작날짜"]
          ? formatDate(areaInfo["시작날짜"])
          : "";
      const doneText =
        areaInfo && areaInfo["완료날짜"]
          ? formatDate(areaInfo["완료날짜"])
          : "";
      const title = document.createElement("div");
      title.className = "area-title";
      const mainRow = document.createElement("div");
      mainRow.className = "area-title-main";
      const idSpan = document.createElement("span");
      idSpan.className = "area-id";
      idSpan.textContent = `${areaId}`;
      mainRow.appendChild(idSpan);
      title.appendChild(mainRow);
      const metaRow = document.createElement("div");
      metaRow.className = "area-meta-row";
      const cardsInArea = grouped[areaId];
      const isComplete = areaCompletionStatus(cardsInArea);
      const hasStart = Boolean(startText);
      const hasDone = Boolean(doneText);
      const inProgress = hasStart && !hasDone;
      const areaCompleteDate =
        areaInfo && areaInfo["완료날짜"]
          ? parseVisitDate(areaInfo["완료날짜"])
          : null;
      const isToday =
        areaInfo && areaInfo["시작날짜"]
          ? isSameDay(areaInfo["시작날짜"], today)
          : false;
      const inviteInfo = state.inviteCampaign;
      const isInviteArea =
        inviteInfo &&
        inviteInfo.startDate &&
        inProgress &&
        startText &&
        startText >= inviteInfo.startDate;
      const range = isComplete ? doneText || startText : startText || doneText;
      const isKsl = isKslArea(areaId);
      if (isComplete || inProgress) {
        const labelParts = [];
        if (range) {
          labelParts.push(range);
        }
        labelParts.push(inProgress ? "시작" : "완료");
        const stateChip = document.createElement("span");
        stateChip.className = "area-date-range";
        stateChip.textContent = labelParts.join(" ");
        metaRow.appendChild(stateChip);
      } else if (range) {
        const dateSpan = document.createElement("span");
        dateSpan.className = "area-date-range";
        dateSpan.textContent = range;
        metaRow.appendChild(dateSpan);
      }
      if (metaRow.childNodes.length) {
        title.appendChild(metaRow);
      }
      const headerRow = document.createElement("div");
      headerRow.className = "area-header-row";
      const leftBox = document.createElement("div");
      leftBox.className = "area-header-left";
      leftBox.appendChild(title);
      const rightBox = document.createElement("div");
      rightBox.className = "area-header-right";
      let badge = null;
      if (inProgress) {
        badge = document.createElement("span");
        badge.className = "status-badge status-badge-progress";
        badge.textContent = "진행중";
        item.classList.add("area-in-progress");
      }
      if (isInviteArea) {
        item.classList.add("area-invite-campaign");
      }
      const activityDate = areaCompleteDate || (startText ? parseVisitDate(startText) : null);
      const isLastService = activityDate && latestServiceDate && activityDate.getTime() === latestServiceDate.getTime();

      if (isToday) {
        if (inProgress) {
          item.classList.add("area-today");
        } else if (isComplete) {
          item.classList.add("area-today-complete");
        }
      } else if (isLastService) {
        if (isComplete) {
          item.classList.add("area-last-complete");
        } else if (inProgress) {
          item.classList.add("area-last-in-progress");
        }
      }

      const fg = document.createElement("div");
      fg.className = "list-item-fg";

      if (
        withAssignButton &&
        isLeader &&
        ((inProgress && !isKsl) || isKsl) &&
        String(state.expandedAreaId) === String(areaId)
      ) {
        const areaCards = (state.data.cards || []).filter(
          (c) =>
            String(c["구역번호"] || "") === String(areaId) &&
            !isTrueValue(c["6개월"]) &&
            !isTrueValue(c["방문금지"]) &&
            !isTrueValue(c["재방"]) &&
            !isTrueValue(c["연구"])
        );
        const areaCardNumbers = areaCards.map((c) => String(c["카드번호"] || ""));
        
        const bulkCheck = document.createElement("input");
        bulkCheck.type = "checkbox";
        bulkCheck.className = "area-bulk-check";
        bulkCheck.title = "구역 전체 선택";
        
        const selectedForArea = areaCardNumbers.filter(cn => (state.selectedCards || []).includes(cn));
        bulkCheck.checked = areaCardNumbers.length > 0 && selectedForArea.length === areaCardNumbers.length;
        bulkCheck.indeterminate = selectedForArea.length > 0 && selectedForArea.length < areaCardNumbers.length;

        ["click", "mousedown", "touchstart"].forEach(type => {
          bulkCheck.addEventListener(type, (e) => {
            e.stopPropagation();
          });
        });
        bulkCheck.addEventListener("change", (e) => {
          e.stopPropagation();
          const set = new Set(state.selectedCards || []);
          if (e.target.checked) {
            areaCardNumbers.forEach(cn => set.add(cn));
          } else {
            areaCardNumbers.forEach(cn => set.delete(cn));
          }
          state.selectedCards = Array.from(set);
          renderAreas();
          renderCards(); 
        });
        rightBox.appendChild(bulkCheck);

        const assignBtn = document.createElement("button");
        assignBtn.type = "button";
        assignBtn.textContent = "차량 배정";
        assignBtn.className = "assign-cards-btn";
        let assignPressTimer = null;
        let assignLongPress = false;
        const clearAssignPress = () => {
          if (assignPressTimer) {
            clearTimeout(assignPressTimer);
            assignPressTimer = null;
          }
        };
        const startAssignPress = () => {
          clearAssignPress();
          assignPressTimer = setTimeout(() => {
            assignPressTimer = null;
            assignLongPress = true;
              const targetAssignDate = state.carAssignDate || todayISO();
            const selectedCards = (state.selectedCards || []).filter(
              (card) =>
                String(card["구역번호"] || "") === String(areaId) &&
                  getCardAssignedCarIdForDate(card, targetAssignDate)
            );
            if (!selectedCards.length) {
              alert(
                "차량 배정을 삭제할 카드를 먼저 선택해 주세요."
              );
              return;
            }
            if (
              !window.confirm(
                `구역 ${areaId}에서 선택된 카드들의 차량 배정을 모두 삭제할까요?`
              )
            ) {
              return;
            }
            const allCards = state.data.cards || [];
            selectedCards.forEach((sel) => {
              const cn = String(sel["카드번호"] || "");
              allCards.forEach((card) => {
                if (
                  String(card["구역번호"] || "") === String(areaId) &&
                  String(card["카드번호"] || "") === cn
                ) {
                  card["차량"] = "";
                  card["배정날짜"] = "";
                }
              });
            });
            renderCards();
            renderCarAssignPopup();
            setStatus("선택된 카드들의 차량 배정이 삭제되었습니다.");
          }, 800);
        };
        assignBtn.addEventListener("mousedown", (event) => {
          if (event.button !== 0) {
            return;
          }
          event.stopPropagation();
          startAssignPress();
        });
        assignBtn.addEventListener("touchstart", (event) => {
          event.stopPropagation();
          startAssignPress();
        });
        ["mouseup", "mouseleave", "touchend", "touchcancel"].forEach(
          (type) => {
            assignBtn.addEventListener(type, clearAssignPress);
          }
        );
        assignBtn.addEventListener("click", (event) => {
          event.stopPropagation();
          if (assignLongPress) {
            assignLongPress = false;
            return;
          }
          if (!state.selectedArea || state.selectedArea !== areaId) {
            alert("해당 구역의 카드를 먼저 펼쳐서 선택해 주세요.");
            return;
          }
          const selected = state.selectedCards.slice();
          if (!selected.length) {
            alert("배정할 카드를 선택해 주세요.");
            return;
          }
          const cars = state.carAssignments || [];
          if (!cars.length) {
            alert("먼저 차량 배정 패널에서 차량을 설정해 주세요.");
            return;
          }
          openCarSelectPopup(areaId, selected);
        });
        rightBox.appendChild(assignBtn);
      }
      const showStartButton =
        withAssignButton &&
        hasDone &&
        !isKsl &&
        state.user &&
        (state.user.role === "인도자" || state.user.role === "관리자");
      if (showStartButton) {
        const startBtn = document.createElement("button");
        startBtn.type = "button";
        startBtn.textContent = "봉사 시작";
        startBtn.className = "start-service-btn";
        startBtn.addEventListener("click", (event) => {
          event.stopPropagation();
          if (!window.confirm(`${areaId} 봉사를 시작할까요?`)) {
            return;
          }
          state.expandedAreaId = areaId;
          state.filterArea = areaId;
          elements.filterArea.value = areaId;
          selectArea(areaId);
          startServiceForArea(areaId);
          elements.areaOverlay.classList.add("hidden");
          renderAreas();
          renderCards();
        });
        rightBox.appendChild(startBtn);
      }
      if (badge) {
        rightBox.appendChild(badge);
      }
      headerRow.append(leftBox, rightBox);
      fg.appendChild(headerRow);

      if (inProgress && String(state.expandedAreaId) !== String(areaId) && isLeader) {
        const bg = document.createElement("div");
        bg.className = "list-item-bg";
        const completeAction = document.createElement("div");
        completeAction.className = "swipe-action swipe-action-left";
        completeAction.textContent = "완료처리";
        const cancelAction = document.createElement("div");
        cancelAction.className = "swipe-action swipe-action-right";
        cancelAction.textContent = "봉사취소";
        bg.append(completeAction, cancelAction);
        item.appendChild(bg);

        let startX = null;
        let startY = null;
        let currentX = 0;
        let isSwiping = false;
        const threshold = 70;

        const handleStart = (e) => {
          if (state.expandedAreaId === areaId) return;
          startX = e.type.startsWith("touch") ? e.touches[0].clientX : e.clientX;
          startY = e.type.startsWith("touch") ? e.touches[0].clientY : e.clientY;
          isSwiping = false;
          justSwiped = false;
          fg.style.transition = "none";
        };

        const handleMove = (e) => {
          if (startX === null || state.expandedAreaId === areaId) return;
          const x = e.type.startsWith("touch") ? e.touches[0].clientX : e.clientX;
          const y = e.type.startsWith("touch") ? e.touches[0].clientY : e.clientY;
          const dx = x - startX;
          const dy = y - startY;

          if (!isSwiping) {
            if (Math.abs(dx) > 10 && Math.abs(dx) > Math.abs(dy)) {
              isSwiping = true;
            }
          }

          if (isSwiping) {
            if (e.cancelable) e.preventDefault();
            currentX = dx;
            fg.style.transform = `translateX(${currentX}px)`;
          }
        };

        const handleEnd = async (e) => {
          if (startX === null || state.expandedAreaId === areaId) return;
          fg.style.transition = "transform 0.2s ease-out";
          fg.style.transform = "translateX(0)";

          if (isSwiping) {
            justSwiped = true;
            setTimeout(() => { justSwiped = false; }, 100);
            if (currentX > threshold) {
              if (window.confirm(`구역 ${areaId}를 완료 처리할까요?\n(방문 기록이 없는 카드는 '부재'로 기록됩니다)`)) {
                await finishAreaWithoutVisits(areaId);
              }
            } else if (currentX < -threshold) {
              if (window.confirm(`구역 ${areaId}의 봉사 시작을 취소할까요?`)) {
                await cancelServiceForArea(areaId);
              }
            }
          }
          startX = null;
          startY = null;
          isSwiping = false;
          currentX = 0;
        };

        item.addEventListener("touchstart", handleStart, { passive: true });
        item.addEventListener("touchmove", handleMove, { passive: false });
        item.addEventListener("touchend", handleEnd);
        item.addEventListener("mousedown", handleStart);
        const moveHandler = (e) => handleMove(e);
        const endHandler = (e) => {
          handleEnd(e);
          window.removeEventListener("mousemove", moveHandler);
          window.removeEventListener("mouseup", endHandler);
        };
        item.addEventListener("mousedown", () => {
          window.addEventListener("mousemove", moveHandler);
          window.addEventListener("mouseup", endHandler);
        });
      }

      item.appendChild(fg);
      item.addEventListener("click", (event) => {
        if (justSwiped) return;
        if (
          event.target.closest(".area-cards") ||
          event.target.closest("#card-list") ||
          event.target.closest(".card") ||
          event.target.closest(".card-history") ||
          event.target.closest("#visit-form") ||
          event.target.closest(".area-bulk-check") ||
          event.target.closest(".status-badge-progress") ||
          event.target.closest(".assign-cards-btn") ||
          event.target.closest(".start-service-btn")
        ) {
          return;
        }
        if (state.expandedAreaId === areaId) {
          state.expandedAreaId = null;
          state.filterArea = "all";
          elements.filterArea.value = "all";
          state.selectedArea = null;
          state.scrollAreaToTop = false;
          renderAreas();
          renderCards();
        } else {
          state.expandedAreaId = areaId;
          state.filterArea = areaId;
          elements.filterArea.value = areaId;
          selectArea(areaId);
        }
      });
      return item;
    };
    const leftItem = createItem(false);
    const overlayItem = createItem(true);
    const inlineItem = createItem(true);
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
  const today = todayISO();
  let startDate = null;
  const isInProgress =
    areaInfo && areaInfo["시작날짜"] && !areaInfo["완료날짜"];
  if (isInProgress) {
    const parsed = parseVisitDate(areaInfo["시작날짜"]);
    if (parsed) {
      startDate = parsed;
    }
  }
  const currentAreaId =
    state.filterArea && state.filterArea !== "all"
      ? String(state.filterArea)
      : state.selectedArea
      ? String(state.selectedArea)
      : "";
  const isKslCurrentArea =
    currentAreaId && isKslArea(currentAreaId) ? true : false;
  const canAssignFromCardsPanel =
    state.user &&
    (state.user.role === "관리자" || state.user.role === "인도자") &&
    currentAreaId &&
    (isKslCurrentArea || isInProgress);
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
  } else if (state.filterVisit === "revisit") {
    cards = cards.filter((card) => isTrueValue(card["재방"]));
  } else if (state.filterVisit === "study") {
    cards = cards.filter((card) => isTrueValue(card["연구"]));
  } else if (state.filterVisit === "six") {
    cards = cards.filter((card) => isTrueValue(card["6개월"]));
  } else if (state.filterVisit === "banned") {
    cards = cards.filter((card) => isTrueValue(card["방문금지"]));
  } else if (state.filterVisit === "campaign-unvisited") {
    const info = state.inviteCampaign;
    if (!info || !info.startDate) {
      cards = [];
    } else {
      const start = info.startDate;
      const end = info.endDate || start;
      const visits = state.data.visits || [];
      cards = cards.filter((card) => {
        const a = String(card["구역번호"] || "");
        const c = String(card["카드번호"] || "");
        if (!a || !c) {
          return false;
        }
        const hasVisit = visits.some((row) => {
          const ra = String(row["구역번호"] || row["areaId"] || "");
          const rc = String(
            row["카드번호"] || row["구역카드"] || row["cardNumber"] || ""
          );
          if (ra !== a || rc !== c) {
            return false;
          }
          const dText =
            normalizeVisitDateText(
              row["방문날짜"] ||
                row["방문일"] ||
                row["방문일자"] ||
                row["날짜"] ||
                row["방문일시"] ||
                row["일자"]
            ) || "";
          if (!dText) {
            return false;
          }
          return dText >= start && dText <= end;
        });
        return !hasVisit;
      });
    }
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
        return compareCardNumbers(na, nb);
      }
    }
    const na = String(a["카드번호"] || "");
    const nb = String(b["카드번호"] || "");
    return compareCardNumbers(na, nb);
  });
  cards.forEach((card) => {
    const cardEl = document.createElement("div");
    cardEl.className = "card";
    const isSelected =
      state.selectedCard &&
      card["카드번호"] === state.selectedCard["카드번호"];
    if (isSelected) {
      cardEl.classList.add("active");
      if (isInProgress) {
        cardEl.classList.add("card-active-progress");
      }
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
    }    const title = document.createElement("div");
    title.className = "card-header";
    const mainTitle = document.createElement("strong");
  const cardNumText = String(card["카드번호"] || "");
  const townText = getCardTownLabel(card);
  mainTitle.textContent = townText
    ? `${cardNumText} (${townText})`
    : cardNumText;
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
    }
    if (card["재방"]) {
      const rv = document.createElement("span");
      rv.className = "badge badge-revisit";
      rv.textContent = "재방";
      badgeRow.appendChild(rv);
    }
    if (card["연구"]) {
      const st = document.createElement("span");
      st.className = "badge badge-study";
      st.textContent = "연구";
      badgeRow.appendChild(st);
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
    const addressText = String(card["주소"] || "").trim();
    const query = addressText || String(card["카드번호"] || "").trim();
    const encoded = encodeURIComponent(query);
    const kakao = document.createElement("a");
    kakao.className = "nav-kakao";
    // 카카오맵 앱 딥링크: 검색어를 바로 전달
    kakao.href = `kakaomap://search?q=${encoded}`;
    kakao.target = "_blank";
    kakao.rel = "noopener noreferrer";
    kakao.textContent = "카카오";
    const naver = document.createElement("a");
    naver.className = "nav-naver";
    naver.href = `https://m.map.naver.com/search2/search.naver?query=${encoded}`;
    naver.target = "_blank";
    naver.rel = "noopener noreferrer";
    naver.textContent = "네이버";
    const google = document.createElement("a");
    google.href = `https://www.google.com/maps/search/?api=1&query=${encoded}`;
    google.target = "_blank";
    google.rel = "noopener noreferrer";
    google.textContent = "구글";
    nav.append(kakao, naver, google);
    if (canAssignFromCardsPanel) {
    const assignBox = document.createElement("span");
    assignBox.className = "card-assign-box";
    const label = document.createElement("label");
    label.className = "card-assign-label";
    const carId = getCardAssignedCarIdForDate(card, todayISO());
    if (carId) {
      const rows = state.data.assignments || [];
      const sameCar = rows.filter(
        (row) => String(row["차량"] || "") === carId
      );
      const driverRow =
        sameCar.find(
          (row) => String(row["역할"] || "") === "운전자"
        ) || null;
      const driverName = driverRow ? String(driverRow["이름"] || "") : "";
      const infoSpan = document.createElement("span");
      infoSpan.className = "card-car-info";
      infoSpan.textContent = driverName
        ? `차량 ${carId} · ${driverName}`
        : `차량 ${carId}`;
      label.appendChild(infoSpan);
    }
    const check = document.createElement("input");
    check.type = "checkbox";
    check.className = "card-select";
    check.dataset.cardNumber = String(card["카드번호"] || "");
    check.checked = state.selectedCards.includes(
      String(card["카드번호"] || "")
    );
    check.addEventListener("change", (e) => {
      const id = String(card["카드번호"] || "");
      const set = new Set(state.selectedCards);
      if (e.target.checked) {
        set.add(id);
      } else {
        set.delete(id);
      }
      state.selectedCards = Array.from(set);
    });
    label.appendChild(check);
    assignBox.appendChild(label);
    nav.appendChild(assignBox);
    }
    cardEl.append(title, address, visitInfo);
    if (regular.textContent) {
      cardEl.appendChild(regular);
    }
    if (info.textContent) {
      cardEl.appendChild(info);
    }
    cardEl.appendChild(nav);
    if (state.user && state.user.role === "관리자") {
      let swipeStartX = null;
      let swipeStartY = null;
      let swipeStartTime = 0;
      cardEl.addEventListener("touchstart", (event) => {
        if (!event.touches || !event.touches.length) {
          return;
        }
        const t = event.touches[0];
        swipeStartX = t.clientX;
        swipeStartY = t.clientY;
        swipeStartTime = Date.now();
      });
      cardEl.addEventListener("touchend", async (event) => {
        if (
          swipeStartX == null ||
          !event.changedTouches ||
          !event.changedTouches.length
        ) {
          return;
        }
        const t = event.changedTouches[0];
        const dx = t.clientX - swipeStartX;
        const dy = t.clientY - swipeStartY;
        swipeStartX = null;
        swipeStartY = null;
        const dt = Date.now() - swipeStartTime;
        const absDx = Math.abs(dx);
        const absDy = Math.abs(dy);
        if (
          absDx < 60 ||
          absDx < absDy * 2 ||
          dt > 800
        ) {
          return;
        }
        if (
          !state.selectedCard ||
          String(state.selectedCard["카드번호"] || "") !==
            String(card["카드번호"] || "")
        ) {
          return;
        }
        const areaId = String(card["구역번호"] || "");
        const cardNumber = String(card["카드번호"] || "");
        if (
          !areaId ||
          !cardNumber ||
          !window.confirm(
            `구역 ${areaId}, 카드 ${cardNumber}를 삭제하고 삭제 시트로 이동할까요?`
          )
        ) {
          return;
        }
        try {
          setLoading(true, "카드를 삭제하는 중...");

          // 1. Supabase에서 삭제/이동 처리
          if (supabaseClient) {
            try {
              await deleteCardInSupabase(areaId, cardNumber);
            } catch (supaErr) {
              console.warn("Supabase delete failed (non-critical):", supaErr);
            }
          }

          // 2. Google Sheets에서 삭제/이동 처리
          const res = await apiRequest("deleteCard", {
            areaId,
            cardNumber
          });
          if (!res.success) {
            alert(res.message || "구역카드 삭제에 실패했습니다.");
            return;
          }
          await loadData();          renderAreas();
          renderCards();
          renderAdminPanel();
          setStatus("구역카드가 삭제되었습니다.");
        } catch (e) {
          alert("구역카드 삭제 중 오류가 발생했습니다.");
        } finally {
          setLoading(false);
        }
      });
    }
    if (state.selectedCard && card["카드번호"] === state.selectedCard["카드번호"]) {
      const oneYearAgo = new Date();
      oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
      const normalizeId = (v) =>
        String(v == null ? "" : v).trim().replace(/\s+/g, "");
      const cardId = normalizeId(card["카드번호"]);
      const allRows = (state.data.visits || []).filter((row) => {
        const rawCard =
          row["구역카드"] || row["카드번호"] || row["구역카드번호"] || "";
        const rowCardId = normalizeId(rawCard);
        return rowCardId && rowCardId === cardId;
      });
      const historyRows = allRows
        .map((row, index) => ({ row, index }))
        .sort((a, b) => {
          const db = parseVisitDate(getVisitDateValue(b.row));
          const da = parseVisitDate(getVisitDateValue(a.row));
          const tb = db ? db.getTime() : 0;
          const ta = da ? da.getTime() : 0;
          if (tb !== ta) {
            return tb - ta;
          }
          // 같은 날짜일 때는 원래 순서 기준으로 나중에 추가된 것을 위로
          return b.index - a.index;
        })
        .map((x) => x.row);
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
            const currentUser = state.user;
            const isLeader =
              currentUser &&
              (currentUser.role === "관리자" || currentUser.role === "인도자");
            if (!isLeader) {
              const visitISO = toISODate(getVisitDateValue(row));
              if (visitISO !== todayISO()) {
                alert("오늘 날짜의 방문만 수정할 수 있습니다.");
                return;
              }
              const areaIdText = String(card["구역번호"] || "");
              const areaInfo =
                (state.data.areas || []).find(
                  (a) => String(a["구역번호"] || "") === areaIdText
                ) || null;
              const hasStart =
                areaInfo && areaInfo["시작날짜"] ? true : false;
              const hasDone =
                areaInfo && areaInfo["완료날짜"] ? true : false;
              const inProgress = hasStart && !hasDone;
              if (!inProgress) {
                alert("봉사가 진행 중인 구역의 방문만 수정할 수 있습니다.");
                return;
              }
              const rows = state.data.assignments || [];
              const myRow =
                rows.find(
                  (r) =>
                    String(r["이름"] || "") === String(currentUser.name || "")
                ) || null;
              const myCarId = myRow ? String(myRow["차량"] || "") : "";
              const cardCarId = getCardAssignedCarIdForDate(card, todayISO());
              if (!myCarId) {
                alert("오늘 봉사 중인 차량 정보가 없습니다.");
                return;
              }
              if (!cardCarId || myCarId !== cardCarId) {
                alert("오늘 봉사 중인 카드의 방문만 수정할 수 있습니다.");
                return;
              }
            }
            state.editingVisit = {
              id: row.id,
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
      if (
        event.target.closest("#visit-form") ||
        event.target.closest(".card-assign-label") ||
        event.target.closest(".card-select") ||
        event.target.closest(".card-assign-box")
      ) {
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
    if (state.scrollAreaToTop && !state.scrollToSelectedCard) {
      const scrollTarget = expanded;
      const scrollOptions = { behavior: "smooth", block: "start" };
      scrollTarget.scrollIntoView(scrollOptions);
      if (typeof window !== "undefined" && window.requestAnimationFrame) {
        window.requestAnimationFrame(() => {
          scrollTarget.scrollIntoView(scrollOptions);
        });
      }
    }
  } else {
    elements.cardListHome.appendChild(elements.cardList);
  }
  elements.cardList.classList.remove("hidden");
  if (state.scrollToSelectedCard && state.selectedCard) {
    const targetNumber = String(state.selectedCard["카드번호"] || "");
    const cardsEls = elements.cardList.querySelectorAll(".card");
    let target = null;
    cardsEls.forEach((el) => {
      if (target) {
        return;
      }
      const strong = el.querySelector(".card-header strong");
      if (strong && String(strong.textContent || "") === targetNumber) {
        target = el;
      }
    });
    if (target) {
      target.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }
  state.scrollToSelectedCard = false;
  state.scrollAreaToTop = false;
};

const renderVisitsView = () => {
  elements.cardList.innerHTML = "";
  elements.visitForm.classList.add("hidden");
  elements.cardList.classList.remove("card-grid");
  elements.areaListInline.innerHTML = "";
  elements.statusMessage.textContent = "";
  const areaId = state.completionExpandedAreaId;
  if (!areaId) {
    elements.statusMessage.textContent = "구역을 선택해 주세요.";
    return;
  }
  elements.areaTitle.textContent = `완료 내역 · ${areaId}`;
  const completions = (state.data.completions || []).filter(
    (row) => String(row["구역번호"] || row["areaId"] || "") === String(areaId)
  );
  completions.sort((a, b) => {
    const da = parseVisitDate(a["완료날짜"] || a["completeDate"]);
    const db = parseVisitDate(b["완료날짜"] || b["completeDate"]);
    const va = da ? da.getTime() : 0;
    const vb = db ? db.getTime() : 0;
    return vb - va;
  });
  const list = document.createElement("div");
  list.className = "list";
  if (!completions.length) {
    const empty = document.createElement("div");
    empty.className = "list-item";
    empty.textContent = "완료 내역이 없습니다.";
    list.appendChild(empty);
  } else {
    completions.forEach((row) => {
      const item = document.createElement("div");
      item.className = "list-item";
      const title = document.createElement("div");
      const startText = row["시작날짜"]
        ? formatDate(row["시작날짜"])
        : "";
      const doneText = row["완료날짜"]
        ? formatDate(row["완료날짜"])
        : "";
      title.textContent =
        startText && doneText
          ? `${startText} → ${doneText}`
          : doneText || startText || "";
      const leader = document.createElement("div");
      leader.textContent = row["인도자"]
        ? `인도자: ${row["인도자"]}`
        : "";
      const memo = document.createElement("div");
      memo.textContent = row["비고"] || "";
      item.appendChild(title);
      if (leader.textContent) {
        item.appendChild(leader);
      }
      if (memo.textContent) {
        item.appendChild(memo);
      }
      list.appendChild(item);
    });
  }
  elements.cardListHome.appendChild(list);
};

const renderCompletionOverlayList = () => {
  if (!elements.completionAreaList) {
    return;
  }
  const box = elements.completionAreaList;
  box.innerHTML = "";
  const completions = state.data.completions || [];
  const byArea = {};
  completions.forEach((row) => {
    const areaId = String(row["구역번호"] || row["areaId"] || "");
    if (!areaId) {
      return;
    }
    if (!byArea[areaId]) {
      byArea[areaId] = [];
    }
    byArea[areaId].push(row);
  });
  const areaIds = Object.keys(byArea).sort((a, b) => {
    const sheetAreaOrder = (state.data.areas || []).map(row => String(row["구역번호"]));
    const idxA = sheetAreaOrder.indexOf(String(a));
    const idxB = sheetAreaOrder.indexOf(String(b));
    if (idxA !== -1 && idxB !== -1) return idxA - idxB;
    if (idxA !== -1) return -1;
    if (idxB !== -1) return 1;
    
    const na = Number(a);
    const nb = Number(b);
    if (!Number.isNaN(na) && !Number.isNaN(nb)) {
      return na - nb;
    }
    return String(a).localeCompare(String(b), "ko-KR");
  });
  if (!areaIds.length) {
    const empty = document.createElement("div");
    empty.className = "list-item";
    empty.textContent = "완료 내역이 없습니다.";
    box.appendChild(empty);
    return;
  }
  const expandedId = state.completionExpandedAreaId;
  areaIds.forEach((areaId) => {
    const item = document.createElement("div");
    item.className = "list-item";
    if (expandedId === areaId) {
      item.classList.add("active");
    }
    item.dataset.areaId = areaId;
    const header = document.createElement("div");
    header.textContent = `${areaId} · ${byArea[areaId].length}회 완료`;
    item.appendChild(header);
    const detail = document.createElement("div");
    const list = byArea[areaId].slice().sort((a, b) => {
      const da = parseVisitDate(a["완료날짜"] || a["completeDate"]);
      const db = parseVisitDate(b["완료날짜"] || b["completeDate"]);
      const va = da ? da.getTime() : 0;
      const vb = db ? db.getTime() : 0;
      return vb - va;
    });
    if (expandedId === areaId) {
      if (!list.length) {
        const empty = document.createElement("div");
        empty.className = "card-history-empty";
        empty.textContent = "완료 내역이 없습니다.";
        detail.appendChild(empty);
      } else {
        list.forEach((row) => {
          const entry = document.createElement("div");
          entry.className = "card-history-item";
          const startText = row["시작날짜"]
            ? formatDate(row["시작날짜"])
            : "";
          const doneText = row["완료날짜"]
            ? formatDate(row["완료날짜"])
            : "";
          const title = startText && doneText
            ? `${startText} → ${doneText}`
            : doneText || startText || "";
          const leader = row["인도자"]
            ? ` · 인도자: ${row["인도자"]}`
            : "";
          const memo = row["비고"] ? ` · ${row["비고"]}` : "";
          entry.textContent = `${title}${leader}${memo}`;
          detail.appendChild(entry);
        });
      }
    } else {
      detail.style.display = "none";
    }
    item.appendChild(detail);
    item.addEventListener("click", () => {
      const id = item.dataset.areaId || "";
      state.completionExpandedAreaId =
        state.completionExpandedAreaId === id ? null : id;
      renderCompletionOverlayList();
    });
    box.appendChild(item);
  });
};

const startServiceForArea = async (areaId) => {
  if (!areaId) {
    alert("구역을 선택해 주세요.");
    return;
  }
  setLoading(true, "봉사 시작 중...");
  try {
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
    
    const { error } = await supabaseClient
      .from("areas")
      .update({
        start_date: todayISO(),
        leader: state.user.name,
        end_date: null
      })
      .eq("area_id", String(areaId));
      
    if (error) throw error;
    
    setStatus("오늘 봉사가 시작되었습니다.");
    await loadData();
    renderAreas();
    renderAdminPanel();
    state.selectedArea = areaId;
    state.view = "area";
    elements.areaTitle.textContent = `구역 ${areaId}`;
    renderCards();
  } catch (err) {
    console.error("Start service error:", err);
    alert("봉사 시작에 실패했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

const cancelServiceForArea = async (areaId) => {
  if (!areaId) return;
  setLoading(true, "봉사 취소 중...");
  try {
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
    
    const { error } = await supabaseClient
      .from("areas")
      .update({
        start_date: null,
        leader: null
      })
      .eq("area_id", String(areaId));
      
    if (error) throw error;
    
    setStatus("봉사 시작이 취소되었습니다.");
    await loadData();
    renderAreas();
    renderAdminPanel();
    renderCards();
    renderMyCarInfo();
  } catch (err) {
    console.error("Cancel service error:", err);
    alert("봉사 취소에 실패했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

const finishAreaWithoutVisits = async (areaId) => {
  if (!areaId) return;
  setLoading(true, "구역 완료 처리 중...");
  try {
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
    
    const { data: areaRow, error: fetchError } = await supabaseClient
      .from("areas")
      .select("*")
      .eq("area_id", areaId)
      .single();
      
    if (fetchError) throw fetchError;
    
    const endDate = todayISO();
    
    const { error: updateError } = await supabaseClient
      .from("areas")
      .update({ end_date: endDate })
      .eq("area_id", areaId);
      
    if (updateError) throw updateError;
    
    const { error: insertError } = await supabaseClient
      .from("completions")
      .insert({
        area_id: areaId,
        start_date: areaRow.start_date,
        end_date: endDate,
        leader: areaRow.leader
      });
      
    if (insertError) throw insertError;
    
    setStatus("구역 봉사가 완료되었습니다.");
    await loadData();
    collapseExpandedArea();
    renderAreas();
    renderCards();
    renderAdminPanel();
  } catch (err) {
    console.error("Finish area error:", err);
    alert("구역 완료 처리에 실패했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

const compareAreaIds = (a, b) => {
  const areas = state.data.areas || [];
  const findIndex = (id) => areas.findIndex(row => {
    const raw = String(row["구역번호"] || "").trim();
    return raw === String(id) || raw.startsWith(String(id) + " ");
  });
  const idxA = findIndex(a);
  const idxB = findIndex(b);
  if (idxA !== -1 && idxB !== -1) return idxA - idxB;
  if (idxA !== -1) return -1;
  if (idxB !== -1) return 1;
  const na = Number(a), nb = Number(b);
  if (!isNaN(na) && !isNaN(nb)) return na - nb;
  return String(a).localeCompare(String(b), "ko-KR");
};

const renderAdminCards = () => {
  const box = document.getElementById("admin-cards-list");
  if (!box) return;
  box.innerHTML = "";
  const byAreaCards = groupCardsByArea();
  const areaIds = Object.keys(byAreaCards).sort((a, b) => compareAreaIds(a, b));
  const selectedAreaId = state.adminCardsAreaId;

  areaIds.forEach((areaId) => {
    const cards = byAreaCards[areaId].sort((ca, cb) => compareCardNumbers(ca["카드번호"], cb["카드번호"]));
    const item = document.createElement("div");
    item.className = "list-item" + (selectedAreaId === areaId ? " active" : "");
    item.innerHTML = `<div><strong>${areaId} · 카드 ${cards.length}장</strong></div>`;
    
    const cardsBox = document.createElement("div");
    cardsBox.className = "area-cards";
    if (selectedAreaId !== areaId) cardsBox.style.display = "none";
    
    if (selectedAreaId === areaId) {
      cards.forEach(card => {
        const cardWrapper = document.createElement("div");
        cardWrapper.className = "card admin-card-row";
        const cardNo = String(card["카드번호"] || "");
        const townText = getCardTownLabel(card);
        
        cardWrapper.innerHTML = `
          <div class="card-header">
            <strong>${cardNo} ${townText ? `(${townText})` : ""}</strong>
            <div class="card-badges">
              <button class="status-badge" data-action="save-card" data-area-id="${areaId}" data-card-number="${cardNo}">저장</button>
              <button class="status-badge" data-action="delete-card" data-area-id="${areaId}" data-card-number="${cardNo}">삭제</button>
            </div>
          </div>
          <div class="card-line">주소: <input type="text" value="${card["주소"] || ""}" data-field="address"></div>
          <div class="card-line">상세: <input type="text" value="${card["상세주소"] || ""}" data-field="detailAddress"></div>
          <div class="card-check-group">
            <label><input type="checkbox" ${isTrueValue(card["6개월"]) ? "checked" : ""} data-field="sixMonths"> 6개월</label>
            <label><input type="checkbox" ${isTrueValue(card["방문금지"]) ? "checked" : ""} data-field="banned"> 방문금지</label>
            <label><input type="checkbox" ${isTrueValue(card["재방"]) ? "checked" : ""} data-field="revisit"> 재방</label>
            <label><input type="checkbox" ${isTrueValue(card["연구"]) ? "checked" : ""} data-field="study"> 연구</label>
          </div>
        `;
        cardsBox.appendChild(cardWrapper);
      });

      // --- 새 카드 추가 폼 복구 ---
      const newCardWrapper = document.createElement("div");
      newCardWrapper.className = "card admin-card-row";
      newCardWrapper.dataset.new = "true";
      newCardWrapper.dataset.areaId = areaId;
      newCardWrapper.innerHTML = `
        <div class="card-header">
          <strong>새 카드 추가</strong>
          <div class="card-badges">
            <button class="status-badge" data-card-action="create-card">추가</button>
          </div>
        </div>
        <div class="card-line">카드번호: <input type="text" placeholder="예: 101-1" data-field="cardNumber"></div>
        <div class="card-line">주소: <input type="text" data-field="address"></div>
        <div class="card-line">상세: <input type="text" data-field="detailAddress"></div>
        <div class="card-check-group">
          <label><input type="checkbox" data-field="sixMonths"> 6개월</label>
          <label><input type="checkbox" data-field="banned"> 방문금지</label>
          <label><input type="checkbox" data-field="revisit"> 재방</label>
          <label><input type="checkbox" data-field="study"> 연구</label>
        </div>
      `;
      cardsBox.appendChild(newCardWrapper);
    }
    
    item.appendChild(cardsBox);
    item.addEventListener("click", (e) => {
      if (e.target.closest("button") || e.target.closest("input")) return;
      state.adminCardsAreaId = state.adminCardsAreaId === areaId ? null : areaId;
      renderAdminCards();
    });
    box.appendChild(item);
  });
};

const renderAdminCompletions = () => {
  const box = document.getElementById("admin-completions-list");
  if (!box) return;
  box.innerHTML = "";
  const completions = [...(state.data.completions || [])].sort((a, b) => new Date(b["완료날짜"]) - new Date(a["완료날짜"]));
  
  if (!completions.length) {
    box.innerHTML = '<div class="list-item">완료 내역이 없습니다.</div>';
    return;
  }

  completions.forEach(c => {
    const div = document.createElement("div");
    div.className = "list-item";
    div.innerHTML = `
      <div style="display: flex; justify-content: space-between;">
        <strong>${c["구역번호"]}</strong>
        <span>${formatDate(c["완료날짜"])}</span>
      </div>
      <div style="font-size: 0.9em; margin-top: 4px;">인도자: ${c["인도자"] || "없음"}</div>
    `;
    box.appendChild(div);
  });
};

const renderAdminEvangelists = () => {
  const box = document.getElementById("admin-evangelist-list");
  if (!box) return;
  box.innerHTML = "";
  // 이름순(가나다순)으로 정렬
  const evangelists = [...(state.data.evangelists || [])].sort((a, b) => 
    String(a["이름"] || "").localeCompare(String(b["이름"] || ""), "ko-KR")
  );
  
  const table = document.createElement("table");
  table.className = "admin-table admin-table-wide";
  table.innerHTML = `
    <thead>
      <tr>
        <th>이름</th><th>성별</th><th>농인</th><th>역할</th><th>운전자</th><th>정원</th><th>부부</th><th>작업</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;
  
  const tbody = table.querySelector("tbody");
  evangelists.forEach(ev => {
    const tr = document.createElement("tr");
    tr.dataset.name = ev["이름"];
    tr.innerHTML = `
      <td><input type="text" value="${ev["이름"] || ""}" readOnly></td>
      <td><input type="text" value="${ev["성별"] || ""}" data-field="gender" style="width:30px"></td>
      <td><input type="checkbox" ${isTrueValue(ev["농인"]) ? "checked" : ""} data-field="deaf"></td>
      <td><input type="text" value="${ev["역할"] || ""}" data-field="role" style="width:60px"></td>
      <td><input type="checkbox" ${isTrueValue(ev["운전자"]) ? "checked" : ""} data-field="driver"></td>
      <td><input type="number" value="${ev["차량"] || 0}" data-field="capacity" style="width:40px"></td>
      <td><input type="text" value="${ev["부부"] || ""}" data-field="spouse" style="width:60px"></td>
      <td class="admin-actions-cell">
        <button data-action="save-ev" data-name="${ev["이름"]}">저장</button>
        <button data-action="delete-ev" data-name="${ev["이름"]}">삭제</button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  // --- 새 전도인 추가 행 ---
  const newTr = document.createElement("tr");
  newTr.dataset.new = "true";
  newTr.innerHTML = `
    <td><input type="text" placeholder="이름" data-field="name"></td>
    <td><input type="text" data-field="gender" style="width:30px"></td>
    <td><input type="checkbox" data-field="deaf"></td>
    <td><input type="text" value="전도인" data-field="role" style="width:60px"></td>
    <td><input type="checkbox" data-field="driver"></td>
    <td><input type="number" data-field="capacity" style="width:40px"></td>
    <td><input type="text" data-field="spouse" style="width:60px"></td>
    <td class="admin-actions-cell"><button data-action="create-ev">추가</button></td>
  `;
  tbody.appendChild(newTr);
  box.appendChild(table);
};

const renderAdminDeletedCards = () => {
  const box = document.getElementById("admin-deleted-card-list");
  if (!box) return;
  box.innerHTML = "";
  const deleted = state.data.deletedCards || [];
  
  if (!deleted.length) {
    box.innerHTML = '<div class="list-item">삭제된 카드가 없습니다.</div>';
    return;
  }

  const byArea = {};
  deleted.forEach(c => {
    const aid = c["구역번호"];
    if (!aid) return;
    if (!byArea[aid]) byArea[aid] = [];
    byArea[aid].push(c);
  });

  Object.keys(byArea).sort((a,b) => compareAreaIds(a,b)).forEach(areaId => {
    const item = document.createElement("div");
    item.className = "list-item";
    item.innerHTML = `<strong>${areaId} · 삭제 카드 ${byArea[areaId].length}장</strong>`;
    const innerBox = document.createElement("div");
    innerBox.className = "deleted-card-box";
    byArea[areaId].forEach(c => {
      const cardEl = document.createElement("div");
      cardEl.className = "card";
      cardEl.dataset.areaId = areaId;
      cardEl.dataset.cardNumber = c["카드번호"];
      
      cardEl.innerHTML = `
        <div class="card-header">
          <strong>${c["카드번호"]}</strong> 
          <button class="status-badge" data-action="restore-deleted" data-area-id="${areaId}" data-card-number="${c["카드번호"]}">복원</button>
        </div>
        <div class="card-line">${c["주소"] || ""}</div>
        <div class="card-line" style="font-size:0.8em; color:#888;">삭제일: ${formatDate(c["삭제일"])}</div>
        <div class="card-history deleted-card-history" style="display:none; margin-top:8px; border-top:1px solid #444; padding-top:8px;"></div>
      `;

      cardEl.addEventListener("click", (e) => {
        if (e.target.closest("button")) return;
        const historyBox = cardEl.querySelector(".deleted-card-history");
        const isVisible = historyBox.style.display !== "none";
        
        if (isVisible) {
          historyBox.style.display = "none";
        } else {
          const areaIdText = cardEl.dataset.areaId;
          const cardNumberText = cardEl.dataset.cardNumber;
          const visits = (state.data.visits || []).filter(v => 
            String(v["구역번호"]) === String(areaIdText) && String(v["카드번호"]) === String(cardNumberText)
          );
          
          historyBox.innerHTML = visits.length 
            ? visits.map(v => `<div style="font-size:0.85em; margin-bottom:4px;">${formatDate(v["방문날짜"])} · ${v["전도인"]} · ${v["결과"]} ${v["메모"] ? `· ${v["메모"]}` : ""}</div>`).join("")
            : '<div style="font-size:0.85em; color:#888;">방문 기록 없음</div>';
          historyBox.style.display = "block";
        }
      });

      innerBox.appendChild(cardEl);
    });
    item.appendChild(innerBox);
    box.appendChild(item);
  });
};

const renderAdminPanel = () => {
  if (!state.user) {
    elements.adminPanel.classList.add("hidden");
    return;
  }

  const userRole = String(state.user.role || "").trim();
  const isAdmin = userRole === "관리자";
  const isLeader = userRole === "인도자";
  const canAccessAdmin = isAdmin || isLeader;

  if (!canAccessAdmin) {
    elements.adminPanel.classList.add("hidden");
    return;
  }

  // 섹션 요소들 정의
  const sections = {
    "admin-cards": document.getElementById("admin-cards-section"),
    "admin-ev": document.getElementById("admin-evangelists-section"),
    "visits": document.getElementById("admin-completions-section"),
    "admin-deleted": document.getElementById("admin-deleted-section")
  };

  // 모든 섹션 숨기기
  Object.values(sections).forEach(el => {
    if (el) el.classList.add("hidden");
  });

  const menu = state.currentMenu;
  const overlayTitleEl = elements.adminOverlay ? elements.adminOverlay.querySelector("#admin-overlay-title") : null;

  // 선택된 메뉴에 따른 타이틀 및 섹션 표시
  if (menu === "admin-cards") {
    if (overlayTitleEl) overlayTitleEl.textContent = "구역카드 관리";
    if (sections["admin-cards"]) sections["admin-cards"].classList.remove("hidden");
    renderAdminCards();
  } else if (menu === "visits") {
    if (overlayTitleEl) overlayTitleEl.textContent = "완료 내역";
    if (sections["visits"]) sections["visits"].classList.remove("hidden");
    renderAdminCompletions();
  } else if (menu === "admin-ev") {
    if (overlayTitleEl) overlayTitleEl.textContent = "전도인 명단 관리";
    if (sections["admin-ev"]) sections["admin-ev"].classList.remove("hidden");
    renderAdminEvangelists();
  } else if (menu === "admin-deleted") {
    if (overlayTitleEl) overlayTitleEl.textContent = "삭제 카드";
    if (sections["admin-deleted"]) sections["admin-deleted"].classList.remove("hidden");
    renderAdminDeletedCards();
  }

  elements.adminPanel.classList.remove("hidden");
};




const renderInviteCampaignOverlay = () => {
  if (!elements.inviteOverlay) {
    return;
  }
  const info = state.inviteCampaign;
  const meta = elements.inviteMeta;
  const statsBox = elements.inviteStats;
  if (meta) {
    if (!info || !info.startDate) {
      meta.textContent = "초대장 배부가 아직 시작되지 않았습니다.";
    } else {
      const statusText = info.active ? "진행중" : "종료됨";
      const rangeText = info.endDate
        ? `${info.startDate} ~ ${info.endDate}`
        : `${info.startDate} ~ 오늘`;
      const memoText = info.memo ? ` (${info.memo})` : "";
      meta.textContent = `${statusText} · ${rangeText}${memoText}`;
    }
  }
  if (elements.inviteStart) {
    elements.inviteStart.disabled = info && info.active;
  }
  if (elements.inviteStop) {
    elements.inviteStop.disabled = !info || !info.active;
  }
  if (statsBox) {
    const summary = state.inviteStats;
    if (!summary) {
      statsBox.textContent = "통계를 불러오려면 통계 새로고침을 눌러 주세요.";
    } else {
      const parts = [];
      parts.push(
        `기간: ${summary.startDate} ~ ${summary.endDate}`
      );
      parts.push(
        `방문한 카드 수: ${summary.totalCards}장, 방문 횟수: ${summary.totalVisits}회`
      );
      let html = `<div>${parts.join("<br>")}</div>`;
      if (summary.byArea && summary.byArea.length) {
        html += '<table class="invite-stats-table"><thead><tr>';
        html += "<th>구역</th><th>방문 카드 수</th><th>방문 횟수</th>";
        html += "</tr></thead><tbody>";
        summary.byArea.forEach((row) => {
          html += `<tr><td>${row.areaId}</td><td>${row.cardCount}</td><td>${row.visitCount}</td></tr>`;
        });
        html += "</tbody></table>";
      }
      statsBox.innerHTML = html;
    }
  }
};

const selectArea = (areaId) => {
  state.selectedArea = areaId;
  state.selectedCard = null;
   state.selectedCards = [];
  state.scrollAreaToTop = true;
  elements.areaTitle.textContent = `구역 ${areaId}`;
  renderAreas();
  renderCards();
  setStatus("");
};

const selectCard = (card) => {
  state.selectedCard = card;
  state.scrollToSelectedCard = true;
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
    : card["재방"]
    ? "재방"
    : card["연구"]
    ? "연구"
    : card["6개월"]
    ? "6개월"
    : card["방문금지"]
    ? "방문금지"
    : "만남";
  elements.visitNote.value = "";
  updateVisitFlagButtons();
  renderCards();
};

const toIsoDate = (dateStr) => {
  if (!dateStr) return null;
  const s = String(dateStr).trim();
  if (!s) return null;
  // yy/mm/dd -> 20yy-mm-dd
  if (/^\d{2}\/\d{2}\/\d{2}$/.test(s)) {
    return "20" + s.replace(/\//g, "-");
  }
  // yyyy-mm-dd... -> yyyy-mm-dd
  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) return isoMatch[0];
  try {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d.toISOString().split("T")[0];
  } catch (e) {}
  return null;
};

const migrateToSupabase = async () => {
  if (!supabaseClient) {
    alert("Supabase 설정이 올바르지 않습니다.");
    return;
  }
  if (!confirm("구글 시트의 데이터를 Supabase로 복사하시겠습니까?")) {
    return;
  }

  setLoading(true, "구글 시트에서 데이터 가져오는 중...");
  try {
    const data = await apiRequest("bootstrap", {}, "GET");
    if (data.error) throw new Error(data.error);
    
    console.log("Data from Google Sheets:", data);

    const checkError = (res, stage) => {
      if (res.error) {
        console.error(`Error at ${stage}:`, res.error);
        throw new Error(`[${stage}] ${res.error.message}\n${res.error.details || ""}`);
      }
    };

    setLoading(true, "구역번호(areas) 동기화 중...");
    const areas = (data.areas || []).map(a => ({
      area_id: String(a["구역번호"] || ""),
      start_date: toIsoDate(a["시작날짜"]),
      end_date: toIsoDate(a["완료날짜"]),
      leader: a["인도자"],
      start_date_backup: toIsoDate(a["시작날짜백업"]),
      end_date_backup: toIsoDate(a["완료날짜백업"]),
      leader_backup: a["인도자백업"]
    })).filter(a => a.area_id);
    console.log("Processed areas:", areas);
    if (areas.length) {
      const res = await supabaseClient.from("areas").upsert(areas, { onConflict: "area_id" });
      checkError(res, "areas");
    }

    setLoading(true, "구역카드(cards) 동기화 중...");
    const cards = (data.cards || []).map(c => ({
      area_id: String(c["구역번호"] || ""),
      card_number: String(c["카드번호"] || ""),
      address: c["주소"],
      recent_visit_date: toIsoDate(c["최근방문일"]),
      prev_visit_date: toIsoDate(c["이전봉사일"]),
      meet: !!c["만남"],
      absent: !!c["부재"],
      revisit: !!c["재방"],
      study: !!c["연구"],
      six_months: !!c["6개월"],
      banned: !!c["방문금지"],
      car_id: String(c["차량"] || ""),
      assignment_date: toIsoDate(c["배정날짜"]),
      invite: !!c["초대장"]
    })).filter(c => c.area_id && c.card_number);
    console.log("Processed cards:", cards);
    if (cards.length) {
      const res = await supabaseClient.from("cards").upsert(cards, { onConflict: "area_id, card_number" });
      checkError(res, "cards");
    }

    setLoading(true, "삭제된카드(deleted_cards) 동기화 중...");
    const deletedCards = (data.deletedCards || []).map(dc => ({
      area_id: String(dc["구역번호"] || ""),
      card_number: String(dc["카드번호"] || ""),
      address: dc["주소"],
      deleted_at: toIsoDate(dc["삭제일"]) || new Date().toISOString()
    })).filter(dc => dc.area_id && dc.card_number);
    console.log("Processed deleted cards:", deletedCards);
    if (deletedCards.length) {
      const res = await supabaseClient.from("deleted_cards").upsert(deletedCards, { onConflict: "area_id, card_number" });
      checkError(res, "deleted_cards");
    }

    setLoading(true, "전도인명단(evangelists) 동기화 중...");
    const evangelists = (data.evangelists || []).map(e => ({
      name: String(e["이름"] || ""),
      password: String(e["비밀번호"] || ""),
      role: e["역할"] || e["권한"] || "전도인",
      gender: e["성별"] || "",
      driver: isTrueValue(e["운전자"]),
      capacity: Number(e["차량"]) || 0,
      spouse: e["부부"] || "",
      is_deaf: isTrueValue(e["농인"])
    })).filter(e => e.name);
    console.log("Processed evangelists:", evangelists);
    if (evangelists.length) {
      const res = await supabaseClient.from("evangelists").upsert(evangelists, { onConflict: "name" });
      checkError(res, "evangelists");
    }

    setLoading(true, "완료내역(completions) 동기화 중...");
    const completions = (data.completions || []).map(c => ({
      area_id: String(c["구역번호"] || c["areaId"] || ""),
      start_date: toIsoDate(c["시작날짜"] || c["startDate"]),
      end_date: toIsoDate(c["완료날짜"] || c["completionDate"] || c["endDate"]),
      leader: c["인도자"] || c["leader"]
    })).filter(c => c.area_id && c.end_date);
    console.log("Processed completions:", completions);
    if (completions.length) {
      const res = await supabaseClient.from("completions").upsert(completions, { onConflict: "area_id, end_date" });
      checkError(res, "completions");
    }

    setLoading(true, "방문기록(visits) 동기화 중...");
    const visits = (data.visits || []).map(v => ({
      area_id: String(v["구역번호"] || v["areaId"] || ""),
      card_number: String(v["카드번호"] || v["구역카드"] || v["cardNumber"] || ""),
      visit_date: toIsoDate(v["방문날짜"] || v["방문일"] || v["날짜"]),
      worker: v["전도인"] || v["방문자"],
      result: v["결과"] || v["방문결과"],
      note: v["메모"] || v["비고"]
    })).filter(v => v.area_id && v.card_number);
    if (visits.length) {
      for (let i = 0; i < visits.length; i += 500) {
        const res = await supabaseClient.from("visits").insert(visits.slice(i, i + 500));
        checkError(res, `visits (chunk ${i/500 + 1})`);
      }
    }

    setLoading(true, "차량배정(assignments) 동기화 중...");
    const assignments = (data.assignments || []).map(a => ({
      date: toIsoDate(a["날짜"]),
      slot: a["시간대"] || "오전",
      car_id: String(a["차량"] || ""),
      driver: a["운전자"] || a["이름"],
      passengers: Array.isArray(a["동승자"]) ? a["동승자"] : []
    }));
    if (assignments.length) {
      const res = await supabaseClient.from("car_assignments").insert(assignments);
      checkError(res, "car_assignments");
    }

    setLoading(true, "봉사신청(volunteer_weeks) 동기화 중...");
    try {
      const volData = await apiRequest("getVolunteerConfig", {});
      if (volData && volData.weeks && volData.weeks.length > 0) {
        for (const week of volData.weeks) {
          const weekStart = week.weekStartISO || toIsoDate(week.weekStartText);
          if (weekStart) {
            const { error: volError } = await supabaseClient
              .from("volunteer_weeks")
              .upsert({
                week_start: weekStart,
                data: week
              });
            if (volError) console.error("Volunteer sync error for week " + weekStart, volError);
          }
        }
      }
    } catch (err) {
      console.warn("Volunteer data sync failed (optional):", err);
    }

    alert("동기화가 완료되었습니다!");
    location.reload();
  } catch (err) {
    console.error("Migration error:", err);
    alert("동기화 실패: " + err.message);
  } finally {
    setLoading(false);
  }
};

const migrateToSheets = async () => {
  if (!state.isSuperAdmin) {
    alert("최고관리자 권한이 필요합니다.");
    return;
  }
  if (!confirm("Supabase의 데이터를 구글 시트로 '전체 덮어쓰기' 하시겠습니까?\n이 작업은 구글 시트의 기존 내용을 삭제하고 Supabase 내용으로 교체합니다.")) {
    return;
  }

  setLoading(true, "Supabase에서 데이터 가져오는 중...");
  try {
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

    // 1. Supabase에서 전체 데이터 읽기
    const { data: areas } = await supabaseClient.from("areas").select("*");
    const { data: cards } = await supabaseClient.from("cards").select("*");
    const { data: deletedCards } = await supabaseClient.from("deleted_cards").select("*");
    const { data: evangelists } = await supabaseClient.from("evangelists").select("*");
    const { data: completions } = await supabaseClient.from("completions").select("*");

    // 2. Apps Script 포맷으로 변환 (한글 키 사용)
    const payload = {
      areas: (areas || []).map(a => ({
        "구역번호": a.area_id, "시작날짜": a.start_date, "완료날짜": a.end_date, "인도자": a.leader
      })),
      cards: (cards || []).map(c => ({
        "구역번호": c.area_id, "카드번호": c.card_number, "주소": c.address,
        "최근방문일": c.recent_visit_date, "이전봉사일": c.prev_visit_date,
        "만남": c.meet, "부재": c.absent, "재방": c.revisit, "연구": c.study,
        "6개월": c.six_months, "방문금지": c.banned, "차량": c.car_id, "배정날짜": c.assignment_date,
        "초대장": c.invite
      })),
      deletedCards: (deletedCards || []).map(dc => ({
        "구역번호": dc.area_id, "카드번호": dc.card_number, "주소": dc.address, "삭제일": dc.deleted_at
      })),
      evangelists: (evangelists || []).map(e => ({
        "이름": e.name, "비밀번호": e.password, "역할": e.role
      })),
      completions: (completions || []).map(c => ({
        "구역번호": c.area_id, "완료날짜": c.completion_date, "인도자": c.leader
      }))
    };

    setLoading(true, "구글 시트로 데이터 전송 중...");
    const res = await apiRequest("syncFromSupabase", { data: JSON.stringify(payload) });
    
    if (res.success) {
      alert("구글 시트 덮어쓰기가 완료되었습니다!");
    } else {
      throw new Error(res.message);
    }
  } catch (err) {
    console.error("Sync to Sheets error:", err);
    alert("동기화 실패: " + err.message);
  } finally {
    setLoading(false);
  }
};

const loadData = async () => {
  setLoading(true, "데이터 불러오는 중...");
  try {
    // 1. Supabase에서 데이터 먼저 시도
    if (supabaseClient) {
      const { data: areas } = await supabaseClient.from("areas").select("*");
      if (areas && areas.length > 0) {
        const { data: cards } = await supabaseClient.from("cards").select("*");
        const { data: evangelists } = await supabaseClient.from("evangelists").select("*");
        const { data: assignments } = await supabaseClient.from("car_assignments").select("*").gte("date", toIsoDate(new Date()));
        const { data: inviteCampaign } = await supabaseClient.from("invite_campaign").select("*").order("created_at", { ascending: false }).limit(1);
        const { data: completions } = await supabaseClient.from("completions").select("*");
        const { data: deletedCards } = await supabaseClient.from("deleted_cards").select("*").order("deleted_at", { ascending: false });

        // 데이터 형식 변환 (Supabase 스네이크 케이스 -> 앱 내부 한글/카멜 케이스 유지)
        state.data.areas = areas.map(a => ({
          "구역번호": a.area_id, "시작날짜": a.start_date, "완료날짜": a.end_date, "인도자": a.leader
        }));
        state.data.cards = cards.map(c => ({
          id: c.id,
          "구역번호": c.area_id, "카드번호": c.card_number, "주소": c.address,
          "최근방문일": c.recent_visit_date, "이전봉사일": c.prev_visit_date,
          "만남": c.meet, "부재": c.absent, "재방": c.revisit, "연구": c.study,
          "6개월": c.six_months, "방문금지": c.banned, "차량": c.car_id, "배정날짜": c.assignment_date,
          "초대장": c.invite
        }));
        state.data.deletedCards = (deletedCards || []).map(dc => ({
          id: dc.id,
          "구역번호": dc.area_id, "카드번호": dc.card_number, "주소": dc.address, "삭제일": dc.deleted_at
        }));        state.data.evangelists = evangelists.map(e => ({
          "이름": e.name, "비밀번호": e.password, "역할": e.role,
          "성별": e.gender, "농인": isTrueValue(e.is_deaf), "운전자": isTrueValue(e.driver),
          "차량": e.capacity, "부부": e.spouse
        }));
        state.data.assignments = assignments.map(a => ({
          id: a.id,
          "날짜": a.date, "시간대": a.slot, "차량": a.car_id, "이름": a.driver, "동승자": a.passengers || []
        }));
        state.inviteCampaign = inviteCampaign && inviteCampaign[0] ? {
          id: inviteCampaign[0].id,
          active: inviteCampaign[0].status === "active",
          startDate: inviteCampaign[0].start_date,
          endDate: inviteCampaign[0].end_date,
          memo: inviteCampaign[0].memo
        } : null;
        state.data.completions = (completions || []).map(c => ({
          "구역번호": c.area_id, "시작날짜": c.start_date, "완료날짜": c.end_date, "인도자": c.leader
        }));

        // 방문기록은 너무 많으면 나중에 페이징 처리 필요 (일단 최근 1000개만)
        const { data: visits } = await supabaseClient.from("visits").select("*").order("visit_date", { ascending: false }).limit(1000);
        state.data.visits = (visits || []).map(v => ({
          id: v.id,
          "구역번호": v.area_id, "카드번호": v.card_number, "방문날짜": v.visit_date,
          "전도인": v.worker, "결과": v.result, "메모": v.note
        }));
        console.log("Supabase data loaded");
        return;
      }
    }

    // 2. Supabase에 데이터가 없거나 실패한 경우 구글 시트 사용 (최초 1회)
    const data = await apiRequest("bootstrap", {}, "GET");
    if (data.error) throw new Error(data.error);
    
    state.data.cards = data.cards || [];
    state.data.areas = data.areas || data.areaStatus || [];
    state.data.completions = data.completions || [];
    state.data.visits = data.visits || [];
    state.data.evangelists = data.evangelists || [];
    state.data.assignments = data.assignments || [];
    state.data.assignments = sanitizeAssignmentRows(state.data.assignments);
    state.inviteCampaign = data.inviteCampaign || null;
    state.data.deletedCards = data.deletedCards || [];
    
    console.log("Google Sheets data loaded (Fallback)");
  } finally {
    setLoading(false);
  }
};

const updateMenuVisibility = () => {
  if (!state.user) return;
  const userRole = String(state.user.role || "").trim();
  const isAdmin = userRole === "관리자";
  const isLeader = userRole === "인도자";
  const isSuper = state.isSuperAdmin === true;
  
  document.querySelectorAll(".admin-only").forEach(el => {
    el.classList.toggle("hidden", !isAdmin);
  });
  
  document.querySelectorAll(".leader-only").forEach(el => {
    el.classList.toggle("hidden", !isAdmin && !isLeader);
  });

  document.querySelectorAll(".super-admin-only").forEach(el => {
    el.classList.toggle("hidden", !isSuper);
  });
};

const enterDashboard = async (user) => {
  state.user = user;
  elements.userInfo.textContent = `${state.user.name} (${state.user.role})`;
  elements.menuToggle.style.display = "inline-block";
  updateMenuVisibility();
  elements.loginPanel.classList.add("hidden");
  elements.dashboard.classList.remove("hidden");
  await loadData();
  if (!state.expandedAreaId && state.filterArea === "all") {
    const inProgressAreaId = getFirstInProgressArea();
    if (inProgressAreaId) {
      state.expandedAreaId = inProgressAreaId;
      state.filterArea = inProgressAreaId;
      state.selectedArea = inProgressAreaId;
      state.scrollAreaToTop = true;
    }
  }
  renderAreas();
  renderCards();
  state.carAssignDate = todayISO();
  state.carAssignSlot = "오전";
  ensureAssignmentState();
  renderAdminPanel();
  renderMyCarInfo();
};

const login = async () => {
  const name = elements.nameInput.value.trim();
  const password = elements.passwordInput.value.trim();
  if (!name) {
    alert("이름을 입력해 주세요.");
    return;
  }
  setLoading(true, "로그인 중입니다...");
  try {
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

    // 1. 최고관리자 체크 (Supabase site_config)
    try {
      const { data: config } = await supabaseClient
        .from("site_config")
        .select("config_value")
        .eq("config_key", "super_admin")
        .single();
      
      if (config && config.config_value) {
        const admin = config.config_value;
        if (name === admin.id && password === admin.pass) {
          state.isSuperAdmin = true;
          const superUser = { name: "최고관리자", role: "관리자" };
          try {
            window.localStorage.setItem("mcUser", JSON.stringify({ name: "최고관리자", role: "관리자", isSuper: true }));
          } catch (e) {}
          await enterDashboard(superUser);
          return;
        }
      }
    } catch (adminErr) {
      console.warn("Super admin check skipped:", adminErr);
    }

    // 2. 일반 전도인 체크 (Supabase evangelists)
    const { data, error } = await supabaseClient
      .from("evangelists")
      .select("*")
      .eq("name", name)
      .eq("password", password)
      .single();
      
    if (error || !data) {
      alert("이름 또는 비밀번호가 올바르지 않습니다.");
      return;
    }
    
    state.isSuperAdmin = false;
    const user = { role: data.role || "전도인", name: data.name };
    try {
      window.localStorage.setItem(
        "mcUser",
        JSON.stringify({ name: user.name, role: user.role })
      );
    } catch (e) {}
    await enterDashboard(user);
  } catch (err) {
    console.error("Login error:", err);
    alert("로그인 중 오류가 발생했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

const resetRecentVisits = async () => {
  if (!state.selectedArea) {
    alert("구역을 선택해 주세요.");
    return;
  }
  setLoading(true, "최근방문일 초기화 중...");
  try {
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
    
    const { error } = await supabaseClient
      .from("cards")
      .update({
        recent_visit_date: null,
        meet: false,
        absent: false,
        invite: false
      })
      .eq("area_id", String(state.selectedArea));
      
    if (error) throw error;
    
    setStatus("최근방문일이 초기화되었습니다.");
    await loadData();
    renderAreas();
    renderCards();
    renderAdminPanel();
  } catch (err) {
    console.error(err);
    alert("초기화에 실패했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

const updateVisitFlagButtons = () => {
  const isAdmin =
    state.user &&
    (state.user.role === "관리자" || state.user.role === "인도자");
  const card = state.selectedCard;
  const buttons = [
    elements.visitClearRevisit,
    elements.visitClearStudy,
    elements.visitClearSix,
    elements.visitClearBanned
  ];
  buttons.forEach((btn) => {
    if (!btn) return;
    if (!isAdmin || !card) {
      btn.classList.add("hidden");
    } else {
      btn.classList.remove("hidden");
    }
  });
  if (!isAdmin || !card) {
    return;
  }
  if (elements.visitClearRevisit) {
    if (isTrueValue(card["재방"])) {
      elements.visitClearRevisit.classList.remove("hidden");
    } else {
      elements.visitClearRevisit.classList.add("hidden");
    }
  }
  if (elements.visitClearStudy) {
    if (isTrueValue(card["연구"])) {
      elements.visitClearStudy.classList.remove("hidden");
    } else {
      elements.visitClearStudy.classList.add("hidden");
    }
  }
  if (elements.visitClearSix) {
    if (isTrueValue(card["6개월"])) {
      elements.visitClearSix.classList.remove("hidden");
    } else {
      elements.visitClearSix.classList.add("hidden");
    }
  }
  if (elements.visitClearBanned) {
    if (isTrueValue(card["방문금지"])) {
      elements.visitClearBanned.classList.remove("hidden");
    } else {
      elements.visitClearBanned.classList.add("hidden");
    }
  }
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
  const isEdit = Boolean(state.editingVisit);
  
  setLoading(true, isEdit ? "방문 기록 수정 중..." : "방문 기록 저장 중...");
  try {
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
    
    if (isEdit) {
      const { error: updateError } = await supabaseClient
        .from("visits")
        .update({
          visit_date: visitDate,
          worker: worker,
          result: result,
          note: note
        })
        .eq("id", state.editingVisit.id);
        
      if (updateError) throw updateError;
    } else {
      const { error: insertError } = await supabaseClient
        .from("visits")
        .insert({
          area_id: areaId,
          card_number: cardNumber,
          visit_date: visitDate,
          worker: worker,
          result: result,
          note: note
        });
        
      if (insertError) throw insertError;
    }

    // 카드 상태 업데이트
    const { data: allVisits, error: visitsError } = await supabaseClient
      .from("visits")
      .select("*")
      .eq("area_id", areaId)
      .eq("card_number", cardNumber)
      .order("visit_date", { ascending: false });
      
    if (visitsError) throw visitsError;
    
    const latest = allVisits[0];
    const latestResult = latest ? latest.result : "";
    const latestDate = latest ? latest.visit_date : null;
    
    const { data: cardRow, error: cardError } = await supabaseClient
      .from("cards")
      .select("*")
      .eq("area_id", areaId)
      .eq("card_number", cardNumber)
      .single();
      
    if (cardError) throw cardError;
    
    const updatePayload = {
      recent_visit_date: latestDate,
      meet: latestResult === "만남",
      absent: latestResult === "부재",
      revisit: latestResult === "재방",
      study: latestResult === "연구",
      invite: latestResult === "초대장",
      six_months: (cardRow.six_months || latestResult === "6개월"),
      banned: (cardRow.banned || latestResult === "방문금지")
    };
    
    const { error: cardUpdateError } = await supabaseClient
      .from("cards")
      .update(updatePayload)
      .eq("area_id", areaId)
      .eq("card_number", cardNumber);
      
    if (cardUpdateError) throw cardUpdateError;
    
    // 구역 완료 여부 확인
    const { data: areaCards, error: areaCardsError } = await supabaseClient
      .from("cards")
      .select("*")
      .eq("area_id", areaId);
      
    if (areaCardsError) throw areaCardsError;
    
    const isComplete = areaCards.every(c => {
      if (c.revisit || c.study || c.six_months || c.banned) return true;
      return !!c.recent_visit_date;
    });
    
    let completeResult = false;
    if (isComplete) {
       const { data: areaRow, error: areaRowError } = await supabaseClient
         .from("areas")
         .select("*")
         .eq("area_id", areaId)
         .single();
         
       if (!areaRowError && areaRow.start_date && !areaRow.end_date) {
         const completeDate = visitDate;
         const leaderName = state.user.name;
         
         await supabaseClient
           .from("areas")
           .update({ end_date: completeDate, leader: leaderName })
           .eq("area_id", areaId);
           
         await supabaseClient
           .from("completions")
           .insert({
             area_id: areaId,
             start_date: areaRow.start_date,
             end_date: completeDate,
             leader: leaderName
           });
           
         completeResult = true;
       }
    }

    await loadData();
    if (!isEdit && completeResult) {
      collapseExpandedArea();
    }
    renderAreas();
    renderCards();
    renderAdminPanel();
    
    if (isEdit) {
      setStatus("방문내역이 수정되었습니다.");
      state.editingVisit = null;
    } else {
      if (completeResult) {
        setStatus("모든 카드의 최근방문일이 기록되었습니다. 완료내역이 업데이트되었습니다.");
      } else {
        setStatus("방문내역이 기록되었습니다.");
      }
    }
  } catch (err) {
    console.error("Save visit error:", err);
    alert("오류가 발생했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

const exportToExcel = async () => {
  if (!supabaseClient) {
    alert("Supabase 클라이언트가 초기화되지 않았습니다.");
    return;
  }
  if (!window.XLSX) {
    alert("Excel 라이브러리를 불러오지 못했습니다.");
    return;
  }

  setLoading(true, "Supabase 데이터를 백업용 엑셀로 생성 중...");
  try {
    const tables = [
      { name: "areas", label: "구역정보" },
      { name: "cards", label: "구역카드" },
      { name: "visits", label: "방문기록" },
      { name: "evangelists", label: "전도인명단" },
      { name: "car_assignments", label: "차량배정" },
      { name: "volunteer_weeks", label: "봉사신청설정" },
      { name: "completions", label: "완료내역" },
      { name: "invite_campaign", label: "초대장캠페인" }
    ];

    const wb = XLSX.utils.book_new();

    for (const table of tables) {
      const { data, error } = await supabaseClient.from(table.name).select("*");
      if (error) {
        console.warn(`Table ${table.name} backup error:`, error);
        continue;
      }
      if (data && data.length > 0) {
        const ws = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, table.label);
      }
    }

    const fileName = `MinistryCard_Backup_${new Date().toISOString().split("T")[0]}.xlsx`;
    XLSX.writeFile(wb, fileName);
    setStatus("데이터 백업이 완료되었습니다.");
  } catch (err) {
    console.error("Excel export error:", err);
    alert("엑셀 백업 중 오류가 발생했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

elements.saveApiUrl.addEventListener("click", saveApiUrl);
if (elements.syncToSupabase) {
  elements.syncToSupabase.addEventListener("click", migrateToSupabase);
}
if (elements.syncToSheets) {
  elements.syncToSheets.addEventListener("click", migrateToSheets);
}
if (elements.backupToExcel) {
  elements.backupToExcel.addEventListener("click", exportToExcel);
}
elements.loginButton.addEventListener("click", login);
if (elements.appTitle) {
  elements.appTitle.addEventListener("click", () => {
    if (!state.user) {
      return;
    }
    refreshAll();
  });
}
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

if (elements.userInfo) {
  elements.userInfo.addEventListener("click", () => {
    if (!state.user) {
      return;
    }
    const ok = window.confirm("로그아웃 하시겠습니까?");
    if (!ok) {
      return;
    }
    logout();
  });
}

if (elements.carAssignSelected) {
  elements.carAssignSelected.addEventListener("click", (event) => {
    const item = event.target.closest(".selected-person");
    if (!carAssignTapSelection) {
      if (!item) {
        return;
      }
      const name = (item.textContent || "").trim();
      if (!name) {
        return;
      }
      const prev = elements.carAssignSelected.querySelector(
        ".selected-person.tap-selected"
      );
      if (prev) {
        prev.classList.remove("tap-selected");
      }
      item.classList.add("tap-selected");
      carAssignTapSelection = { name, carId: "" };
      return;
    }
    const selection = carAssignTapSelection;
    carAssignTapSelection = null;
    const prevCar = elements.carAssignPanel.querySelector(
      ".car-member.tap-selected"
    );
    if (prevCar) {
      prevCar.classList.remove("tap-selected");
    }
    const prevUn = elements.carAssignSelected.querySelector(
      ".selected-person.tap-selected"
    );
    if (prevUn) {
      prevUn.classList.remove("tap-selected");
    }
    if (!selection.name) {
      return;
    }
    if (!selection.carId) {
      return;
    }
    const cars = state.carAssignments || [];
    const fromCar = cars.find(
      (c) => String(c.carId) === String(selection.carId)
    );
    if (!fromCar) {
      return;
    }
    fromCar.members = (fromCar.members || []).filter(
      (n) => n !== selection.name
    );
    cars.forEach((car) => {
      const first = (car.members || [])[0];
      car.driver = first || "";
    });
    state.carAssignments = cars;
    renderCarAssignPopup();
  });
}

const tryAutoLogin = async () => {
  try {
    const raw = window.localStorage.getItem("mcUser");
    if (!raw) {
      return;
    }
    const data = JSON.parse(raw);
    if (!data || !data.name || !data.role) {
      return;
    }
    state.isSuperAdmin = data.isSuper === true;
    await enterDashboard({ name: data.name, role: data.role });
  } catch (e) {}
};

const logout = () => {
  state.user = null;
  state.isSuperAdmin = false;
  elements.configPanel.classList.add("hidden");
  try {
    window.localStorage.removeItem("mcUser");
  } catch (e) {}
  elements.userInfo.textContent = "";
  elements.menuToggle.style.display = "none";
  elements.dashboard.classList.add("hidden");
  elements.loginPanel.classList.remove("hidden");
  elements.nameInput.value = "";
  elements.passwordInput.value = "";
};

tryAutoLogin();

if (elements.visitClearRevisit) {
  elements.visitClearRevisit.addEventListener("click", async () => {
    if (!state.selectedArea || !state.selectedCard) {
      return;
    }
    if (!window.confirm("이 카드의 재방 표시를 해제할까요?")) {
      return;
    }
    try {
      setLoading(true, "재방 표시 해제 중...");
      const res = await updateCardFlagsInSupabase(
        state.selectedArea,
        state.selectedCard["카드번호"],
        { revisit: false }
      );
      if (!res.success) {
        alert("재방 해제에 실패했습니다.");
        return;
      }
      state.selectedCard["재방"] = false;
      const found = state.data.cards.find(
        (card) =>
          String(card["구역번호"]) === String(state.selectedArea) &&
          String(card["카드번호"]) === String(state.selectedCard["카드번호"])
      );
      if (found) {
        found["재방"] = false;
      }
      renderAreas();
      renderCards();
      renderAdminPanel();
      updateVisitFlagButtons();
      setStatus("재방 표시가 해제되었습니다.");
    } catch (e) {
      console.error(e);
      alert("재방 해제 중 오류가 발생했습니다.");
    } finally {
      setLoading(false);
    }
  });
}

if (elements.visitClearStudy) {
  elements.visitClearStudy.addEventListener("click", async () => {
    if (!state.selectedArea || !state.selectedCard) {
      return;
    }
    if (!window.confirm("이 카드의 연구 표시를 해제할까요?")) {
      return;
    }
    try {
      setLoading(true, "연구 표시 해제 중...");
      const res = await updateCardFlagsInSupabase(
        state.selectedArea,
        state.selectedCard["카드번호"],
        { study: false }
      );
      if (!res.success) {
        alert("연구 해제에 실패했습니다.");
        return;
      }
      state.selectedCard["연구"] = false;
      const found = state.data.cards.find(
        (card) =>
          String(card["구역번호"]) === String(state.selectedArea) &&
          String(card["카드번호"]) === String(state.selectedCard["카드번호"])
      );
      if (found) {
        found["연구"] = false;
      }
      renderAreas();
      renderCards();
      renderAdminPanel();
      updateVisitFlagButtons();
      setStatus("연구 표시가 해제되었습니다.");
    } catch (e) {
      console.error(e);
      alert("연구 해제 중 오류가 발생했습니다.");
    } finally {
      setLoading(false);
    }
  });
}

if (elements.visitClearSix) {
  elements.visitClearSix.addEventListener("click", async () => {
    if (!state.selectedArea || !state.selectedCard) {
      return;
    }
    if (!window.confirm("이 카드의 6개월 표시를 해제할까요?")) {
      return;
    }
    try {
      setLoading(true, "6개월 표시 해제 중...");
      const res = await updateCardFlagsInSupabase(
        state.selectedArea,
        state.selectedCard["카드번호"],
        { sixMonths: false }
      );
      if (!res.success) {
        alert("6개월 해제에 실패했습니다.");
        return;
      }
      state.selectedCard["6개월"] = false;
      const found = state.data.cards.find(
        (card) =>
          String(card["구역번호"]) === String(state.selectedArea) &&
          String(card["카드번호"]) === String(state.selectedCard["카드번호"])
      );
      if (found) {
        found["6개월"] = false;
      }
      renderAreas();
      renderCards();
      renderAdminPanel();
      updateVisitFlagButtons();
      setStatus("6개월 표시가 해제되었습니다.");
    } catch (e) {
      console.error(e);
      alert("6개월 해제 중 오류가 발생했습니다.");
    } finally {
      setLoading(false);
    }
  });
}

if (elements.visitClearBanned) {
  elements.visitClearBanned.addEventListener("click", async () => {
    if (!state.selectedArea || !state.selectedCard) {
      return;
    }
    if (!window.confirm("이 카드의 방문금지 표시를 해제할까요?")) {
      return;
    }
    try {
      setLoading(true, "방문금지 표시 해제 중...");
      const res = await updateCardFlagsInSupabase(
        state.selectedArea,
        state.selectedCard["카드번호"],
        { banned: false }
      );
      if (!res.success) {
        alert("방문금지 해제에 실패했습니다.");
        return;
      }
      state.selectedCard["방문금지"] = false;
      const found = state.data.cards.find(
        (card) =>
          String(card["구역번호"]) === String(state.selectedArea) &&
          String(card["카드번호"]) === String(state.selectedCard["카드번호"])
      );
      if (found) {
        found["방문금지"] = false;
      }
      renderAreas();
      renderCards();
      renderAdminPanel();
      updateVisitFlagButtons();
      setStatus("방문금지 표시가 해제되었습니다.");
    } catch (e) {
      console.error(e);
      alert("방문금지 해제 중 오류가 발생했습니다.");
    } finally {
      setLoading(false);
    }
  });
}

elements.carAssignEvangelistList.addEventListener("click", (event) => {
  const item = event.target.closest(".ev-item");
  if (!item) {
    return;
  }
  const role = item.dataset.role || "";
  if (role === "temp-add") {
    const input = window.prompt("임시로 추가할 이름을 입력해 주세요.");
    if (!input) {
      return;
    }
    const name = input.trim();
    if (!name) {
      return;
    }
    const normalizedName = normalizeAssignmentName(name);
    if (!state.participantsToday.includes(normalizedName)) {
      state.participantsToday.push(normalizedName);
    }
    renderSelectedParticipants();
    renderCarAssignmentsPanel();
    const listEl = elements.carAssignEvangelistList;
    if (listEl) {
      const existing = listEl.querySelector(
        '.ev-item[data-name="' + name + '"]'
      );
      if (!existing) {
        const tempItem = document.createElement("div");
        tempItem.className = "ev-item selected";
        tempItem.dataset.name = name;
        tempItem.textContent = name;
        const addBtn = listEl.querySelector(".ev-temp-add");
        if (addBtn && addBtn.parentElement === listEl) {
          listEl.insertBefore(tempItem, addBtn);
        } else {
          listEl.appendChild(tempItem);
        }
      }
    }
    return;
  }
  const name = normalizeAssignmentName(item.dataset.name || "");
  if (!name) {
    return;
  }
  const exists = state.participantsToday.includes(name);
  if (exists) {
    state.participantsToday = state.participantsToday.filter((n) => n !== name);
    item.classList.remove("selected");
    const cars = state.carAssignments || [];
    cars.forEach((car) => {
      car.members = (car.members || []).filter(
        (n) => normalizeAssignmentName(n) !== name
      );
    });
    state.carAssignments = cars;
  } else {
    state.participantsToday.push(name);
    item.classList.add("selected");
  }
  renderSelectedParticipants();
  renderCarAssignmentsPanel();
});

elements.carAssignPanel.addEventListener("dragstart", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) {
    return;
  }
  if (!target.classList.contains("car-member")) {
    return;
  }
  target.classList.add("dragging");
  const name = target.dataset.name || "";
  const carId = target.dataset.carId || "";
  event.dataTransfer.setData(
    "text/plain",
    JSON.stringify({ name, carId })
  );
});

elements.carAssignPanel.addEventListener("dragend", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) {
    return;
  }
  if (target.classList.contains("car-member")) {
    target.classList.remove("dragging");
  }
});

elements.carAssignPanel.addEventListener("dragover", (event) => {
  const column = event.target.closest(".car-column");
  if (!column) {
    return;
  }
  const zone = column.querySelector(".car-members");
  if (!zone) {
    return;
  }
  event.preventDefault();
});

elements.carAssignPanel.addEventListener("drop", (event) => {
  const column = event.target.closest(".car-column");
  if (!column) {
    return;
  }
  const zone = column.querySelector(".car-members");
  if (!zone) {
    return;
  }
  event.preventDefault();
  const data = event.dataTransfer.getData("text/plain");
  if (!data) {
    return;
  }
  let parsed;
  try {
    parsed = JSON.parse(data);
  } catch (e) {
    return;
  }
  const name = normalizeAssignmentName(parsed.name);
  const fromCarId = parsed.carId;
  const toCarId = zone.dataset.carId || "";
  if (!name || !fromCarId || !toCarId) {
    return;
  }
  const cars = state.carAssignments || [];
  const fromCar = cars.find((c) => String(c.carId) === String(fromCarId));
  const toCar = cars.find((c) => String(c.carId) === String(toCarId));
  if (!fromCar || !toCar) {
    return;
  }
  if (fromCarId === toCarId) {
    const rest = (fromCar.members || []).filter((n) => n !== name);
    fromCar.members = [name].concat(rest);
  } else {
    fromCar.members = (fromCar.members || []).filter(
      (n) => normalizeAssignmentName(n) !== name
    );
    if (!toCar.members) {
      toCar.members = [];
    }
    if (
      !(toCar.members || []).some(
        (member) => normalizeAssignmentName(member) === name
      )
    ) {
      toCar.members.push(name);
    }
  }
  cars.forEach((car) => {
    const first = (car.members || [])[0];
    car.driver = first || "";
  });
  renderCarAssignPopup();
});

elements.carAssignPanel.addEventListener("click", async (event) => {
  const column = event.target.closest(".car-column");
  if (!column) {
    return;
  }
  const zone = column.querySelector(".car-members");
  if (!zone) {
    return;
  }
  const toCarId = zone.dataset.carId || "";
  if (!toCarId) {
    return;
  }
  const cardTag = event.target.closest(".car-card-tag");

  // 카드 이동 처리
  if (cardTag) {
    const cardNumber = cardTag.dataset.cardNumber || "";
    const areaId = cardTag.dataset.areaId || "";
    const fromCarId = cardTag.dataset.carId || "";
    if (!cardNumber || !areaId || !fromCarId) {
      return;
    }
    if (event.shiftKey || event.ctrlKey || event.metaKey) {
      const key = `${areaId}:${cardNumber}:${fromCarId}`;
      const idx = carAssignCardMultiSelection.indexOf(key);
      if (idx === -1) {
        carAssignCardMultiSelection.push(key);
        cardTag.classList.add("car-card-tag-multi");
      } else {
        carAssignCardMultiSelection.splice(idx, 1);
        cardTag.classList.remove("car-card-tag-multi");
      }
      carAssignCardSelection = null;
      carAssignTapSelection = null;
      return;
    }
    if (!carAssignCardSelection) {
      const prev = elements.carAssignPanel.querySelector(
        ".car-card-tag-selected"
      );
      if (prev) {
        prev.classList.remove("car-card-tag-selected");
      }
      cardTag.classList.add("car-card-tag-selected");
      carAssignCardSelection = { cardNumber, areaId, carId: fromCarId };
      carAssignTapSelection = null;
      carAssignCardMultiSelection = [];
      const multiPrev = elements.carAssignPanel.querySelectorAll(
        ".car-card-tag-multi"
      );
      multiPrev.forEach((el) => el.classList.remove("car-card-tag-multi"));
      return;
    }
    const selection = carAssignCardSelection;
    if (toCarId === (selection.carId || "")) {
      const prev = elements.carAssignPanel.querySelector(
        ".car-card-tag-selected"
      );
      if (prev) {
        prev.classList.remove("car-card-tag-selected");
      }
      cardTag.classList.add("car-card-tag-selected");
      carAssignCardSelection = { cardNumber, areaId, carId: fromCarId };
      return;
    }
    carAssignCardSelection = null;
    const prevCard = elements.carAssignPanel.querySelector(
      ".car-card-tag-selected"
    );
    if (prevCard) {
      prevCard.classList.remove("car-card-tag-selected");
    }
    const moveList = [];
    if (carAssignCardMultiSelection.length) {
      carAssignCardMultiSelection.forEach((key) => {
        const parts = key.split(":");
        if (parts.length === 3) {
          moveList.push({
            areaId: parts[0],
            cardNumber: parts[1],
            fromCarId: parts[2]
          });
        }
      });
    } else if (selection.cardNumber && selection.areaId) {
      moveList.push({
        areaId: selection.areaId,
        cardNumber: selection.cardNumber,
        fromCarId: selection.carId || ""
      });
    }
    if (!moveList.length || !toCarId) {
      return;
    }
    const assignDate = toAssignmentDateText(state.carAssignDate || todayISO());
    const allCards = state.data.cards || [];
    moveList.forEach((item) => {
      if (!item.areaId || !item.cardNumber) return;
      allCards.forEach((card) => {
        if (
          String(card["구역번호"] || "") === String(item.areaId) &&
          String(card["카드번호"] || "") === String(item.cardNumber)
        ) {
          card["차량"] = String(toCarId);
          card["배정날짜"] = assignDate;
        }
      });
    });
    carAssignCardMultiSelection = [];
    const multiPrev = elements.carAssignPanel.querySelectorAll(
      ".car-card-tag-multi"
    );
    multiPrev.forEach((el) => el.classList.remove("car-card-tag-multi"));
    renderCards();
    renderCarAssignPopup();
    return;
  }

  if (carAssignCardSelection || carAssignCardMultiSelection.length) {
    const selection = carAssignCardSelection || {};
    carAssignCardSelection = null;
    const prevCard = elements.carAssignPanel.querySelector(
      ".car-card-tag-selected"
    );
    if (prevCard) {
      prevCard.classList.remove("car-card-tag-selected");
    }
    const moveList = [];
    if (carAssignCardMultiSelection.length) {
      carAssignCardMultiSelection.forEach((key) => {
        const parts = key.split(":");
        if (parts.length === 3) {
          moveList.push({
            areaId: parts[0],
            cardNumber: parts[1],
            fromCarId: parts[2]
          });
        }
      });
    } else if (selection.cardNumber && selection.areaId) {
      moveList.push({
        areaId: selection.areaId,
        cardNumber: selection.cardNumber,
        fromCarId: selection.carId || ""
      });
    }
    if (!moveList.length || !toCarId) {
      return;
    }
    const assignDate = toAssignmentDateText(state.carAssignDate || todayISO());
    const allCards = state.data.cards || [];
    moveList.forEach((item) => {
      if (!item.areaId || !item.cardNumber) return;
      allCards.forEach((card) => {
        if (
          String(card["구역번호"] || "") === String(item.areaId) &&
          String(card["카드번호"] || "") === String(item.cardNumber)
        ) {
          card["차량"] = String(toCarId);
          card["배정날짜"] = assignDate;
        }
      });
    });
    carAssignCardMultiSelection = [];
    const multiPrev = elements.carAssignPanel.querySelectorAll(
      ".car-card-tag-multi"
    );
    multiPrev.forEach((el) => el.classList.remove("car-card-tag-multi"));
    renderCards();
    renderCarAssignPopup();
    return;
  }

  const member = event.target.closest(".car-member");
  if (!carAssignTapSelection) {
    if (!member) {
      return;
    }
    const name = member.dataset.name || "";
    const fromCarId = member.dataset.carId || "";
    if (!name || !fromCarId) {
      return;
    }
    const prev = elements.carAssignPanel.querySelector(
      ".car-member.tap-selected"
    );
    if (prev) {
      prev.classList.remove("tap-selected");
    }
    member.classList.add("tap-selected");
    carAssignTapSelection = { name, carId: fromCarId };
    return;
  }
  const selection = carAssignTapSelection;
  carAssignTapSelection = null;
  const prev = elements.carAssignPanel.querySelector(
    ".car-member.tap-selected"
  );
  if (prev) {
    prev.classList.remove("tap-selected");
  }
  if (!selection.name) {
    return;
  }
  const name = selection.name;
  const fromCarId = selection.carId || "";
  const cars = state.carAssignments || [];
  const toCar = cars.find((c) => String(c.carId) === String(toCarId));
  if (!name || !toCar) {
    return;
  }
  const fromCar = fromCarId
    ? cars.find((c) => String(c.carId) === String(fromCarId))
    : null;
  if (fromCarId && !fromCar) {
    return;
  }
  if (fromCarId && fromCarId === toCarId) {
    const rest = (fromCar.members || []).filter((n) => n !== name);
    fromCar.members = [name].concat(rest);
  } else {
    if (fromCar) {
      fromCar.members = (fromCar.members || []).filter((n) => n !== name);
    }
    if (!toCar.members) {
      toCar.members = [];
    }
    if (!toCar.members.includes(name)) {
      toCar.members.push(name);
    }
  }
  cars.forEach((car) => {
    const first = (car.members || [])[0];
    car.driver = first || "";
  });
  renderCarAssignPopup();
});

elements.menuToggle.addEventListener("click", () => {
  elements.sideMenu.classList.remove("hidden");
});

elements.menuClose.addEventListener("click", () => {
  elements.sideMenu.classList.add("hidden");
});

elements.sideMenu.addEventListener("click", (event) => {
  if (event.target === elements.sideMenu) {
    elements.sideMenu.classList.add("hidden");
  }
});

elements.sideMenu.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-menu]");
  if (!button) {
    return;
  }
  const key = button.dataset.menu;
  if (!key) {
    return;
  }
  
  const isAdmin = state.user && state.user.role === "관리자";
  const isLeader = state.user && state.user.role === "인도자";
  
  if (
    (key === "admin-cards" ||
      key === "admin-ev" ||
      key === "admin-banned" ||
      key === "admin-deleted" ||
      key === "invite-campaign") &&
    !isAdmin
  ) {
    alert("관리자만 사용할 수 있습니다.");
    return;
  }
  
  if (key === "car-assign" && !isAdmin && !isLeader) {
    alert("관리자 또는 인도자만 사용할 수 있습니다.");
    return;
  }
  if (key === "visits" && !isAdmin) {
    alert("관리자만 사용할 수 있습니다.");
    return;
  }
  
  state.currentMenu = key;
  elements.sideMenu.classList.add("hidden");
  if (key === "cards") {
    renderAreas();
    renderCards();
    renderAdminPanel();
  } else if (key === "visits") {
    state.completionExpandedAreaId = null;
    if (elements.completionOverlay) {
      elements.completionOverlay.classList.remove("hidden");
      document.body.style.overflow = "hidden";
      renderCompletionOverlayList();
    } else {
      renderVisitsView();
      renderAdminPanel();
    }
  } else if (key === "volunteer") {
    openVolunteerOverlay();
  } else if (
    key === "admin-cards" ||
    key === "admin-ev" ||
    key === "admin-banned" ||
    key === "admin-deleted"
  ) {
    if (elements.adminOverlay) {
      elements.adminOverlay.classList.remove("hidden");
      document.body.style.overflow = "hidden";
    }
    renderAdminPanel();
  } else if (key === "car-assign") {
    elements.carAssignOverlay.classList.remove("hidden");
    document.body.style.overflow = "hidden";
    (async () => {
      await loadVolunteerConfig();
      const defaultDate = getNearestVolunteerDateISO(todayISO());
      await setCarAssignDate(defaultDate);
    })();
  } else if (key === "invite-campaign") {
    if (elements.inviteOverlay) {
      elements.inviteOverlay.classList.remove("hidden");
      document.body.style.overflow = "hidden";
      renderInviteCampaignOverlay();
    }
  } else if (key === "config") {
    if (state.isSuperAdmin) {
      elements.configPanel.classList.remove("hidden");
    } else {
      alert("최고관리자만 설정에 접근할 수 있습니다.");
    }
  }
});

if (elements.closeConfig && elements.configPanel) {
  elements.closeConfig.addEventListener("click", () => {
    elements.configPanel.classList.add("hidden");
  });
}

if (elements.closeVolunteer && elements.volunteerOverlay) {
  elements.closeVolunteer.addEventListener("click", () => {
    elements.volunteerOverlay.classList.add("hidden");
    document.body.style.overflow = "";
  });
  elements.volunteerOverlay.addEventListener("click", (event) => {
    if (event.target === elements.volunteerOverlay) {
      elements.volunteerOverlay.classList.add("hidden");
      document.body.style.overflow = "";
    }
  });
}

elements.closeCarAssign.addEventListener("click", () => {
  elements.carAssignOverlay.classList.add("hidden");
  document.body.style.overflow = "";
});

elements.carAssignOverlay.addEventListener("click", (event) => {
  if (event.target === elements.carAssignOverlay) {
    elements.carAssignOverlay.classList.add("hidden");
    document.body.style.overflow = "";
  }
});

if (elements.closeInvite && elements.inviteOverlay) {
  elements.closeInvite.addEventListener("click", () => {
    elements.inviteOverlay.classList.add("hidden");
    document.body.style.overflow = "";
  });
  elements.inviteOverlay.addEventListener("click", (event) => {
    if (event.target === elements.inviteOverlay) {
      elements.inviteOverlay.classList.add("hidden");
      document.body.style.overflow = "";
    }
  });
}

if (elements.inviteStart) {
  elements.inviteStart.addEventListener("click", async () => {
    setLoading(true, "초대장 배부를 시작하는 중...");
    try {
      if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
      
      const { data, error } = await supabaseClient
        .from("invite_campaign")
        .insert({
          status: "active",
          start_date: todayISO()
        })
        .select()
        .single();
        
      if (error) throw error;
      
      state.inviteCampaign = {
        id: data.id,
        active: true,
        startDate: data.start_date,
        memo: data.memo
      };
      state.inviteStats = null;
      renderInviteCampaignOverlay();
    } catch (err) {
      console.error(err);
      alert("초대장 배부 시작에 실패했습니다: " + err.message);
    } finally {
      setLoading(false);
    }
  });
}

if (elements.inviteStop) {
  elements.inviteStop.addEventListener("click", async () => {
    setLoading(true, "초대장 배부를 종료하는 중...");
    try {
      if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
      
      const { data, error } = await supabaseClient
        .from("invite_campaign")
        .update({
          status: "finished",
          end_date: todayISO()
        })
        .eq("status", "active")
        .select()
        .single();
        
      if (error) throw error;
      
      state.inviteCampaign = {
        id: data.id,
        active: false,
        startDate: data.start_date,
        endDate: data.end_date,
        memo: data.memo
      };
      renderInviteCampaignOverlay();
    } catch (err) {
      console.error(err);
      alert("초대장 배부 종료에 실패했습니다: " + err.message);
    } finally {
      setLoading(false);
    }
  });
}

if (elements.inviteRefresh) {
  elements.inviteRefresh.addEventListener("click", async () => {
    setLoading(true, "초대장 배부 통계를 불러오는 중...");
    try {
      if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
      
      if (!state.inviteCampaign) {
        alert("진행 중이거나 종료된 캠페인이 없습니다.");
        return;
      }
      
      const start = state.inviteCampaign.startDate;
      const end = state.inviteCampaign.endDate || todayISO();
      
      const { data: visits, error } = await supabaseClient
        .from("visits")
        .select("*")
        .gte("visit_date", start)
        .lte("visit_date", end)
        .eq("result", "초대장");
        
      if (error) throw error;
      
      const totalVisits = (visits || []).length;
      const uniqueCards = new Set((visits || []).map(v => `${v.area_id}-${v.card_number}`)).size;
      
      const byAreaMap = {};
      (visits || []).forEach(v => {
        if (!byAreaMap[v.area_id]) byAreaMap[v.area_id] = { areaId: v.area_id, cardCount: 0, visitCount: 0, cards: new Set() };
        byAreaMap[v.area_id].visitCount++;
        byAreaMap[v.area_id].cards.add(v.card_number);
      });
      
      const areaList = Object.values(byAreaMap).map(a => ({
        areaId: a.areaId,
        visitCount: a.visitCount,
        cardCount: a.cards.size
      }));
      
      state.inviteStats = {
        totalVisits,
        totalCards: uniqueCards,
        byArea: areaList
      };
      
      renderInviteCampaignOverlay();
    } catch (err) {
      console.error(err);
      alert("통계를 불러오는 데 실패했습니다: " + err.message);
    } finally {
      setLoading(false);
    }
  });
}

if (elements.closeAdmin) {
  elements.closeAdmin.addEventListener("click", () => {
    elements.adminOverlay.classList.add("hidden");
    document.body.style.overflow = "";
  });
}

if (elements.adminOverlay) {
  elements.adminOverlay.addEventListener("click", (event) => {
    if (event.target === elements.adminOverlay) {
      elements.adminOverlay.classList.add("hidden");
      document.body.style.overflow = "";
    }
  });
}

if (elements.closeCompletion) {
  elements.closeCompletion.addEventListener("click", () => {
    elements.completionOverlay.classList.add("hidden");
    document.body.style.overflow = "";
  });
}

if (elements.completionOverlay) {
  elements.completionOverlay.addEventListener("click", (event) => {
    if (event.target === elements.completionOverlay) {
      elements.completionOverlay.classList.add("hidden");
      document.body.style.overflow = "";
    }
  });
}

if (elements.adminEvAdd) {
  elements.adminEvAdd.addEventListener("click", async () => {
    const name = window.prompt("전도인 이름을 입력해 주세요.");
    if (!name) {
      return;
    }
    const gender = window.prompt("성별을 입력해 주세요. (예: 남, 여)");
    const role = window.prompt("역할을 입력해 주세요. (예: 전도인, 인도자, 관리자)");
    const driver = window.prompt("운전자 여부를 입력해 주세요. (Y/N)");
    const capacityText = window.prompt("차량 정원 인원을 입력해 주세요. (숫자, 없으면 빈칸)");
    const spouse = window.prompt("부부 이름을 입력해 주세요. (없으면 빈칸)");
    const password = window.prompt("비밀번호를 입력해 주세요. (없으면 기존 유지 또는 빈칸)");
    const driverFlag =
      driver && (driver.toUpperCase() === "Y" || driver === "1" || driver.toLowerCase() === "true");
    const capacity = capacityText ? Number(capacityText) || 0 : "";
    try {
      if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
      
      const dbData = {
        name: name,
        role: role || "전도인",
        gender: gender || "",
        driver: !!driverFlag,
        capacity: capacity === "" ? 0 : Number(capacity),
        spouse: spouse || ""
      };
      if (password) {
        dbData.password = password;
      }
      
      const { error } = await supabaseClient
        .from("evangelists")
        .upsert(dbData);
        
      if (error) throw error;
      
      await loadData();
      renderAdminPanel();
      setStatus("전도인 정보가 저장되었습니다.");
    } catch (err) {
      console.error(err);
      alert("전도인 저장에 실패했습니다: " + err.message);
    }
  });
}

if (elements.adminEvDelete) {
  elements.adminEvDelete.addEventListener("click", async () => {
    const name = window.prompt("삭제할 전도인 이름을 입력해 주세요.");
    if (!name) {
      return;
    }
    if (!window.confirm(`${name} 전도인을 삭제하시겠습니까?`)) {
      return;
    }
    try {
      if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
      
      const { error } = await supabaseClient
        .from("evangelists")
        .delete()
        .eq("name", name);
        
      if (error) throw error;
      
      await loadData();
      renderAdminPanel();
      setStatus("전도인이 삭제되었습니다.");
    } catch (err) {
      console.error(err);
      alert("전도인 삭제에 실패했습니다: " + err.message);
    }
  });
}

if (elements.adminCarEdit) {
  elements.adminCarEdit.addEventListener("click", () => {
    alert("차량 정보는 전도인 표에서 직접 수정할 수 있습니다.");
  });
}

if (elements.evangelistList) {
  elements.evangelistList.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) {
      return;
    }
    const action = button.dataset.action;
    const rowEl = button.closest("tr");
    if (!rowEl) {
      return;
    }
    const isNew = rowEl.dataset.new === "true";
    const getInput = (field) =>
      rowEl.querySelector('input[data-field="' + field + '"]') ||
      rowEl.querySelector('select[data-field="' + field + '"]');
    if (action === "create-ev") {
      const nameInput = getInput("name");
      const name = nameInput ? nameInput.value.trim() : "";
      if (!name) {
        alert("이름을 입력해 주세요.");
        return;
      }
      const genderInput = getInput("gender");
      const deafInput = getInput("deaf");
      const roleInput = getInput("role");
      const driverInput = getInput("driver");
      const capInput = getInput("capacity");
      const spouseInput = getInput("spouse");
      const pwInput = getInput("password");
      const payload = {
        name,
        gender: genderInput ? genderInput.value.trim() : "",
        deaf: deafInput ? !!deafInput.checked : false,
        role: roleInput ? roleInput.value.trim() : "",
        driver: driverInput ? driverInput.checked : false,
        capacity: capInput && capInput.value ? String(Number(capInput.value) || 0) : "",
        spouse: spouseInput ? spouseInput.value.trim() : "",
        password: pwInput ? pwInput.value : ""
      };
      try {
        if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
        
        const dbData = {
          name: payload.name,
          gender: payload.gender,
          is_deaf: !!payload.deaf,
          role: payload.role || "전도인",
          driver: !!payload.driver,
          capacity: payload.capacity ? Number(payload.capacity) : 0,
          spouse: payload.spouse
        };
        if (payload.password) dbData.password = payload.password;
        
        const { error } = await supabaseClient
          .from("evangelists")
          .upsert(dbData);
          
        if (error) throw error;
        
        await loadData();
        renderAdminPanel();
        setStatus("전도인이 추가되었습니다.");
      } catch (e) {
        console.error(e);
        alert("전도인 저장 중 오류가 발생했습니다: " + e.message);
      }
      return;
    }
    const name = rowEl.dataset.name || button.dataset.name || "";
    if (!name) {
      alert("이름 정보를 찾을 수 없습니다.");
      return;
    }
    if (action === "save-ev") {
      const genderInput = getInput("gender");
      const roleInput = getInput("role");
      const driverInput = getInput("driver");
      const capInput = getInput("capacity");
      const spouseInput = getInput("spouse");
      const pwInput = getInput("password");
      const payload = {
        name,
        gender: genderInput ? genderInput.value.trim() : "",
        deaf: getInput("deaf") ? !!getInput("deaf").checked : false,
        role: roleInput ? roleInput.value.trim() : "",
        driver: driverInput ? driverInput.checked : false,
        capacity: capInput && capInput.value ? String(Number(capInput.value) || 0) : "",
        spouse: spouseInput ? spouseInput.value.trim() : "",
        password: pwInput ? pwInput.value : ""
      };
      try {
        if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
        
        const dbData = {
          name: payload.name,
          gender: payload.gender,
          is_deaf: !!payload.deaf,
          role: payload.role || "전도인",
          driver: !!payload.driver,
          capacity: payload.capacity ? Number(payload.capacity) : 0,
          spouse: payload.spouse
        };
        if (payload.password) dbData.password = payload.password;
        
        const { error } = await supabaseClient
          .from("evangelists")
          .upsert(dbData);
          
        if (error) throw error;
        
        await loadData();
        renderAdminPanel();
        setStatus("전도인 정보가 저장되었습니다.");
      } catch (e) {
        console.error(e);
        alert("전도인 저장 중 오류가 발생했습니다: " + e.message);
      }
    } else if (action === "delete-ev") {
      if (!window.confirm(`${name} 전도인을 삭제하시겠습니까?`)) {
        return;
      }
      try {
        if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
        
        const { error } = await supabaseClient
          .from("evangelists")
          .delete()
          .eq("name", name);
          
        if (error) throw error;
        
        await loadData();
        renderAdminPanel();
        setStatus("전도인이 삭제되었습니다.");
      } catch (e) {
        console.error(e);
        alert("전도인 삭제 중 오류가 발생했습니다: " + e.message);
      }
    }
  });
}

if (elements.completionList) {
  elements.completionList.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-card-action]");
    if (!button) {
      return;
    }
    const action = button.dataset.cardAction;
    const rowEl =
      button.closest("tr") || button.closest(".admin-card-row");
    if (!rowEl) {
      return;
    }
    const tableEl = rowEl.closest("table");
    const areaId =
      rowEl.dataset.areaId ||
      (tableEl && tableEl.dataset.areaId) ||
      state.adminCardsAreaId;
    if (!areaId) {
      alert("구역 정보를 찾을 수 없습니다.");
      return;
    }
    const getInput = (field) =>
      rowEl.querySelector('input[data-field="' + field + '"]');
    if (action === "create-card") {
      const cardInput = getInput("cardNumber");
      const cardNumber = cardInput ? cardInput.value.trim() : "";
      if (!cardNumber) {
        alert("카드번호를 입력해 주세요.");
        return;
      }
      const addressInput = getInput("address");
      const detailInput = getInput("detailAddress");
      const memoInput = getInput("memo");
      const townInput = getInput("town");
      const payload = {
        areaId,
        cardNumber,
        isNew: true,
        town: townInput ? townInput.value : "",
        address: addressInput ? addressInput.value : "",
        detailAddress: detailInput ? detailInput.value : "",
        memo: memoInput ? memoInput.value : "",
        sixMonths: getInput("sixMonths")?.checked || false,
        banned: getInput("banned")?.checked || false,
        revisit: getInput("revisit")?.checked || false,
        study: getInput("study")?.checked || false
      };
      const originalText = button.textContent;
      button.textContent = "저장 중...";
      button.disabled = true;
      setLoading(true, "구역카드를 추가하는 중...");
      try {
        const res = await apiRequest("upsertCard", payload);
        if (!res.success) {
          alert(res.message || "구역카드 저장에 실패했습니다.");
          return;
        }
        state.data.cards = res.cards || [];
        renderAreas();
        renderCards();
        renderAdminPanel();
        setStatus("구역카드가 추가되었습니다.");
      } catch (e) {
        alert("구역카드 저장 중 오류가 발생했습니다.");
      } finally {
        button.textContent = originalText;
        button.disabled = false;
        setLoading(false);
      }
      return;
    }
    const baseCardNumber = rowEl.dataset.cardNumber || "";
    if (!baseCardNumber) {
      alert("카드번호 정보를 찾을 수 없습니다.");
      return;
    }
    if (action === "save-card") {
      const addressInput = getInput("address");
      const detailInput = getInput("detailAddress");
      const memoInput = getInput("memo");
      const townInput = getInput("town");
      const payload = {
        areaId,
        cardNumber: baseCardNumber,
        town: townInput ? townInput.value : "",
        address: addressInput ? addressInput.value : "",
        detailAddress: detailInput ? detailInput.value : "",
        memo: memoInput ? memoInput.value : "",
        sixMonths: getInput("sixMonths")?.checked || false,
        banned: getInput("banned")?.checked || false,
        revisit: getInput("revisit")?.checked || false,
        study: getInput("study")?.checked || false
      };
      const originalText = button.textContent;
      button.textContent = "저장 중...";
      button.disabled = true;
      setLoading(true, "구역카드를 저장하는 중...");
      try {
        if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
        
        const dbData = {
          area_id: String(payload.areaId),
          card_number: String(payload.cardNumber),
          address: payload.address,
          meet: !!payload.meet,
          absent: !!payload.absent,
          revisit: !!payload.revisit,
          study: !!payload.study,
          six_months: !!payload.sixMonths,
          banned: !!payload.banned
        };
        
        const { error } = await supabaseClient
          .from("cards")
          .upsert(dbData, { onConflict: "area_id, card_number" });
          
        if (error) throw error;
        
        await loadData();
        renderAreas();
        renderCards();
        renderAdminPanel();
        setStatus("구역카드가 저장되었습니다.");
      } catch (e) {
        console.error(e);
        alert("구역카드 저장 중 오류가 발생했습니다: " + e.message);
      } finally {
        button.textContent = originalText;
        button.disabled = false;
        setLoading(false);
      }
    } else if (action === "delete-card") {
      if (
        !window.confirm(
          `구역 ${areaId}, 카드 ${baseCardNumber}를 삭제하시겠습니까?`
        )
      ) {
        return;
      }
      const originalText = button.textContent;
      button.textContent = "삭제 중...";
      button.disabled = true;
      setLoading(true, "구역카드를 삭제하는 중...");
      try {
        if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

        // 1. Supabase에서 삭제/이동 처리
        await deleteCardInSupabase(areaId, baseCardNumber);

        // 2. Google Sheets에서 삭제/이동 처리 (동기화 유지)
        try {
          await apiRequest("deleteCard", {
            areaId: String(areaId),
            cardNumber: String(baseCardNumber)
          });
        } catch (gasErr) {
          console.warn("GAS delete failed (non-critical):", gasErr);
        }

        await loadData();
        renderAreas();
        renderCards();
        renderAdminPanel();
        setStatus("구역카드가 삭제되었습니다.");
      } catch (e) {
        console.error(e);
        alert("구역카드 삭제 중 오류가 발생했습니다: " + e.message);
      } finally {        button.textContent = originalText;
        button.disabled = false;
        setLoading(false);
      }
    } else if (action === "clear-flags") {
      const cards = state.data.cards || [];
      const target = cards.find(
        (card) =>
          String(card["구역번호"]) === String(areaId) &&
          String(card["카드번호"]) === String(baseCardNumber)
      );
      if (!target) {
        alert("카드를 찾을 수 없습니다.");
        return;
      }
      const flags = {
        revisit: isTrueValue(target["재방"]),
        study: isTrueValue(target["연구"]),
        sixMonths: isTrueValue(target["6개월"]),
        banned: isTrueValue(target["방문금지"])
      };
      if (!flags.revisit && !flags.study && !flags.sixMonths && !flags.banned) {
        alert("해제할 표시가 없습니다.");
        return;
      }
      const payload = { areaId, cardNumber: baseCardNumber };
      if (flags.revisit && window.confirm("재방 표시를 해제할까요?")) {
        payload.revisit = false;
      }
      if (flags.study && window.confirm("연구 표시를 해제할까요?")) {
        payload.study = false;
      }
      if (flags.sixMonths && window.confirm("6개월 표시를 해제할까요?")) {
        payload.sixMonths = false;
      }
      if (flags.banned && window.confirm("방문금지 표시를 해제할까요?")) {
        payload.banned = false;
      }
      const hasChange = Object.keys(payload).some(
        (key) => !["areaId", "cardNumber"].includes(key)
      );
      if (!hasChange) {
        return;
      }
      setLoading(true, "카드 상태 해제 중...");
      try {
        const res = await updateCardFlagsInSupabase(areaId, baseCardNumber, payload);
        if (!res.success) {
          alert("상태 해제에 실패했습니다.");
          return;
        }
        const found = state.data.cards.find(
          (card) =>
            String(card["구역번호"]) === String(areaId) &&
            String(card["카드번호"]) === String(baseCardNumber)
        );
        if (found) {
          if ("revisit" in res) {
            found["재방"] = !!res.revisit;
          }
          if ("study" in res) {
            found["연구"] = !!res.study;
          }
          if ("sixMonths" in res) {
            found["6개월"] = !!res.sixMonths;
          }
          if ("banned" in res) {
            found["방문금지"] = !!res.banned;
          }
        }
        renderAreas();
        renderCards();
        renderAdminPanel();
        setStatus("카드 상태가 해제되었습니다.");
      } catch (e) {
        console.error(e);
        alert("상태 해제 중 오류가 발생했습니다.");
      } finally {
        setLoading(false);
      }
    }
  });
}

if (elements.deletedCardList) {
  elements.deletedCardList.addEventListener("click", async (event) => {
    const btn = event.target.closest("button");
    if (btn && btn.dataset.cardNumber) {
      const areaId = btn.dataset.areaId || "";
      const cardNumber = btn.dataset.cardNumber || "";
      if (!areaId || !cardNumber) {
        return;
      }
      const action = btn.dataset.action || "restore-deleted";
      if (action === "restore-deleted") {
        if (
          !window.confirm(
            `구역 ${areaId}, 카드 ${cardNumber}를 삭제 목록에서 복원할까요?`
          )
        ) {
          return;
        }
        try {
          setLoading(true, "삭제된 카드를 복원하는 중...");

          // 1. Supabase에서 복원 처리
          if (supabaseClient) {
            try {
              await restoreDeletedCardInSupabase(areaId, cardNumber);
            } catch (supaErr) {
              console.warn("Supabase restore failed:", supaErr);
            }
          }

          // 2. Google Sheets에서 복원 처리 (동기화)
          const res = await apiRequest("restoreDeletedCard", {
            areaId,
            cardNumber
          });
          if (!res.success) {
            alert(res.message || "삭제 카드 복원에 실패했습니다.");
            return;
          }
          await loadData();
          renderAreas();
          renderCards();
          renderAdminPanel();
          setStatus("삭제된 카드가 복원되었습니다.");
        } catch (e) {
          alert("삭제 카드 복원 중 오류가 발생했습니다.");
        } finally {
          setLoading(false);
        }
      } else if (action === "purge-deleted") {
        if (
          !window.confirm(
            `구역 ${areaId}, 카드 ${cardNumber}를 삭제 카드 목록에서 영구적으로 삭제할까요?`
          )
        ) {
          return;
        }
        try {
          setLoading(true, "삭제된 카드를 영구 삭제하는 중...");

          // 1. Supabase에서 영구 삭제 처리
          if (supabaseClient) {
            try {
              await purgeDeletedCardInSupabase(areaId, cardNumber);
            } catch (supaErr) {
              console.warn("Supabase purge failed:", supaErr);
            }
          }

          // 2. Google Sheets에서 영구 삭제 처리 (동기화)
          const res = await apiRequest("purgeDeletedCard", {
            areaId,
            cardNumber
          });
          if (!res.success) {
            alert(res.message || "삭제 카드 영구 삭제에 실패했습니다.");
            return;
          }
          await loadData();
          renderAdminPanel();
          setStatus("삭제된 카드가 영구 삭제되었습니다.");
        } catch (e) {
          alert("삭제 카드 영구 삭제 중 오류가 발생했습니다.");
        } finally {
          setLoading(false);
        }
      }
      return;
    }
  });
}

elements.carAssignAuto.addEventListener("click", () => {
  autoAssignCars();
  renderCarAssignPopup();
});

if (elements.carAssignAssignCards) {
  elements.carAssignAssignCards.addEventListener("click", async () => {
    const cars = state.carAssignments || [];
    if (!cars.length) {
      alert("먼저 차량을 배정해 주세요.");
      return;
    }
    const areas = state.data.areas || [];
    let targetAreaIds = areas
      .filter((row) => {
        const start = row["시작날짜"];
        const done = row["완료날짜"];
        const areaId = String(row["구역번호"] || "");
        return start && !done && !isKslArea(areaId);
      })
      .map((row) => String(row["구역번호"] || ""));
    if (!targetAreaIds.length) {
      const input = window.prompt(
        "카드를 배정할 구역번호를 입력해 주세요."
      );
      if (!input) {
        return;
      }
      targetAreaIds = [input.trim()];
    }
    const allCards = state.data.cards || [];
    const assignDate = toAssignmentDateText(state.carAssignDate || todayISO());
    const areaCardsList = targetAreaIds.map((areaId) => ({
      areaId,
      cards: allCards.filter((card) => {
        const area = String(card["구역번호"] || "");
        const carId = getCardAssignedCarIdForDate(
          card,
          state.carAssignDate || todayISO()
        );
        const isRevisit = isTrueValue(card["재방"]);
        const isStudy = isTrueValue(card["연구"]);
        const isSixMonths = isTrueValue(card["6개월"]);
        const isBanned = isTrueValue(card["방문금지"]);
        return (
          area === String(areaId) &&
          !carId &&
          !isRevisit &&
          !isStudy &&
          !isSixMonths &&
          !isBanned
        );
      })
        .sort((a, b) =>
          compareCardNumbers(a["카드번호"] || "", b["카드번호"] || "")
        )
    }));
    const totalCards = areaCardsList.reduce(
      (sum, item) => sum + item.cards.length,
      0
    );
    if (!totalCards) {
      alert("선택된 구역에 해당하는 구역카드가 없습니다.");
      return;
    }
    const areaLabel =
      targetAreaIds.length === 1
        ? `구역 ${targetAreaIds[0]}`
        : `진행중 구역 ${targetAreaIds.join(", ")}`;
    if (
      !window.confirm(
        `${areaLabel}의 구역카드 ${totalCards}장을 차량 ${cars.length}대에 자동 배정할까요?`
      )
    ) {
      return;
    }
    const flatCards = [];
    areaCardsList.forEach(({ areaId, cards }) => {
      cards.forEach((card) => {
        flatCards.push(card);
      });
    });
    if (!flatCards.length) {
      return;
    }
    const carsCount = cars.length;
    const base = Math.floor(totalCards / carsCount);
    const extra = totalCards % carsCount;
    const capacities = cars.map((_, index) =>
      index >= carsCount - extra ? base + 1 : base
    );
    let cursor = 0;
    for (let i = 0; i < carsCount; i++) {
      const car = cars[i];
      const limit = capacities[i];
      for (let n = 0; n < limit && cursor < flatCards.length; n += 1) {
        const card = flatCards[cursor];
        cursor += 1;
        const cardNumber = String(card["카드번호"] || "");
        if (!cardNumber) {
          continue;
        }
        card["차량"] = String(car.carId || "");
        card["배정날짜"] = assignDate;
      }
    }
    renderAreas();
    renderCards();
    renderAdminPanel();
    renderMyCarInfo();
    setStatus("구역카드가 차량에 자동 배정되었습니다.");
    renderCarAssignPopup();
  });
}

if (elements.closeCarSelect && elements.carSelectOverlay) {
  elements.closeCarSelect.addEventListener("click", () => {
    elements.carSelectOverlay.classList.add("hidden");
  });
}

elements.carAssignSave.addEventListener("click", () => {
  saveCarAssignments();
});

elements.carAssignAdd.addEventListener("click", () => {
  const cars = state.carAssignments || [];
  const nextId =
    cars.reduce((max, c) => {
      const v = Number(c.carId) || 0;
      return v > max ? v : max;
    }, 0) + 1;
  cars.push({
    carId: String(nextId),
    driver: "",
    capacity: 0,
    members: []
  });
  state.carAssignments = cars;
  renderCarAssignPopup();
});

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker
      .getRegistrations()
      .then((registrations) => {
        registrations.forEach((reg) => reg.unregister());
      })
      .catch(() => {});
    if (window.caches && caches.keys) {
      caches
        .keys()
        .then((keys) => Promise.all(keys.map((key) => caches.delete(key))))
        .catch(() => {});
    }
  });
}

elements.carAssignReset.addEventListener("click", async () => {
  await resetCarAssignmentsOnServer();
  renderCarAssignPopup();
});

if (elements.carAssignTempAdd) {
  elements.carAssignTempAdd.addEventListener("click", () => {
    const input = window.prompt("임시로 추가할 이름을 입력해 주세요.");
    if (!input) {
      return;
    }
    const name = input.trim();
    if (!name) {
      return;
    }
    if (!state.participantsToday.includes(name)) {
      state.participantsToday.push(name);
    }
    renderSelectedParticipants();
    renderCarAssignmentsPanel();
  });
}

loadApiUrl();
