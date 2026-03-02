const DEFAULT_API_URL =
  (typeof window !== "undefined" &&
    window.__ENV__ &&
    window.__ENV__.API_URL) ||
  "https://script.google.com/macros/s/AKfycbx1O_Ab6n1j2py4A6Qck48fSY8N_J1wTiZF97y09HzW21kHgCinR1K2rCWiZmONG8mJ/exec";
const state = {
  apiUrl: DEFAULT_API_URL,
  user: null,
  data: {
    cards: [],
    areas: [],
    completions: [],
    visits: [],
    evangelists: [],
    assignments: []
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
  adminCardsAreaId: null,
  adminCardsSelectedKey: null,
  adminBannedAreaId: null,
  adminBannedSelectedKey: null,
  completionExpandedAreaId: null,
  editingVisit: null,
  statusTimer: null,
  scrollToSelectedCard: false,
  participantsToday: [],
  carAssignments: [],
  currentMenu: "cards"
};

let carAssignTapSelection = null;

const elements = {
  configPanel: document.getElementById("config-panel"),
  loginPanel: document.getElementById("login-panel"),
  dashboard: document.getElementById("dashboard"),
  menuToggle: document.getElementById("menu-toggle"),
  menuClose: document.getElementById("menu-close"),
  sideMenu: document.getElementById("side-menu"),
  nameInput: document.getElementById("name-input"),
  passwordInput: document.getElementById("password-input"),
  loginButton: document.getElementById("login-button"),
  refreshButton: document.getElementById("refresh-button"),
  apiUrlInput: document.getElementById("api-url-input"),
  saveApiUrl: document.getElementById("save-api-url"),
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
  carAssignOverlay: document.getElementById("car-assign-overlay"),
  closeCarAssign: document.getElementById("close-car-assign"),
  adminOverlay: document.getElementById("admin-overlay"),
  closeAdmin: document.getElementById("close-admin"),
  carAssignEvangelistList: document.getElementById("car-assign-evangelist-list"),
  carAssignSelected: document.getElementById("car-assign-selected"),
  carAssignMeta: document.getElementById("car-assign-meta"),
  carAssignAuto: document.getElementById("car-assign-auto"),
  carAssignAdd: document.getElementById("car-assign-add"),
  carAssignReset: document.getElementById("car-assign-reset"),
  carAssignSave: document.getElementById("car-assign-save"),
  adminPanel: document.getElementById("admin-panel"),
  completionList: document.getElementById("completion-list"),
  visitList: document.getElementById("visit-list"),
  evangelistList: document.getElementById("evangelist-list"),
  carAssignPanel: document.getElementById("car-assign-panel"),
  bannedCardList: document.getElementById("banned-card-list"),
  searchInput: document.getElementById("search-input"),
  searchButton: document.getElementById("search-button"),
  filterArea: document.getElementById("filter-area"),
  filterVisit: document.getElementById("filter-visit"),
  areaListInline: document.getElementById("area-list-inline"),
  loadingIndicator: document.getElementById("loading-indicator"),
  loadingText: document.getElementById("loading-text"),
  adminCardAdd: document.getElementById("admin-card-add"),
  adminCardDelete: document.getElementById("admin-card-delete"),
  adminCarEdit: document.getElementById("admin-car-edit"),
  adminEvAdd: document.getElementById("admin-ev-add"),
  adminEvDelete: document.getElementById("admin-ev-delete")
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
    return "";
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

const renderMyCarInfo = () => {
  const box = elements.carInfo;
  if (!box || !state.user) {
    return;
  }
  const rows = state.data.assignments || [];
  const name = state.user.name;
  if (!rows.length || !name) {
    box.textContent = "";
    box.classList.add("hidden");
    return;
  }
  const myRow =
    rows.find((row) => String(row["이름"] || "") === String(name)) || null;
  if (!myRow) {
    box.textContent = "";
    box.classList.add("hidden");
    return;
  }
  const carId = String(myRow["차량"] || "");
  if (!carId) {
    box.textContent = "";
    box.classList.add("hidden");
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
    .map((row) => String(row["이름"] || ""))
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
  box.textContent = textParts.join(" | ");
  box.classList.remove("hidden");
  const defaultSize = 15;
  const minSize = 11;
  let size = defaultSize;
  box.style.fontSize = defaultSize + "px";
  if (box.scrollWidth > box.clientWidth) {
    while (size > minSize && box.scrollWidth > box.clientWidth) {
      size -= 1;
      box.style.fontSize = size + "px";
    }
  }
};

const getEvangelistByName = (name) =>
  (state.data.evangelists || []).find(
    (row) => String(row["이름"] || "") === String(name || "")
  ) || null;

const buildInitialParticipantsFromAssignments = () => {
  const names =
    (state.data.assignments || []).map((row) => String(row["이름"] || "")) || [];
  state.participantsToday = Array.from(new Set(names));
};

const buildCarAssignmentsFromServer = () => {
  const rows = state.data.assignments || [];
  const byCar = {};
  rows.forEach((row) => {
    const carId = String(row["차량"] || "");
    const name = String(row["이름"] || "");
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
  const validCars =
    state.carAssignments.filter((car) => {
      const hasDriver = Boolean(car.driver);
      const hasMembers = (car.members || []).some((name) => Boolean(name));
      return hasDriver || hasMembers;
    }) || [];
  const rows = [];
  validCars.forEach((car) => {
    const carId = car.carId;
    const driver = car.driver;
    const members = car.members || [];
    if (driver) {
      rows.push({ carId, name: driver, role: "운전자" });
    }
    members.forEach((name) => {
      if (!name || name === driver) {
        return;
      }
      rows.push({ carId, name, role: "탑승자" });
    });
  });
  setLoading(true, "차량 배정 저장 중...");
  try {
    const res = await apiRequest("saveCarAssignments", {
      assignments: JSON.stringify(rows)
    });
    if (!res.success) {
      alert(res.message || "차량 배정 저장에 실패했습니다.");
      return;
    }
    state.data.assignments = res.assignments || [];
    buildInitialParticipantsFromAssignments();
    buildCarAssignmentsFromServer();
    renderAdminPanel();
    renderMyCarInfo();
  } finally {
    setLoading(false);
  }
};

const resetCarAssignmentsOnServer = async () => {
  setLoading(true, "차량 배정 초기화 중...");
  try {
    const res = await apiRequest("resetCarAssignments", {});
    if (!res.success) {
      alert(res.message || "차량 배정 초기화에 실패했습니다.");
      return;
    }
    state.data.assignments = [];
    state.participantsToday = [];
    state.carAssignments = [];
    renderAdminPanel();
    renderMyCarInfo();
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
      assigned.add(String(car.driver));
    }
    (car.members || []).forEach((name) => {
      if (name) {
        assigned.add(String(name));
      }
    });
  });
  const unassigned = names.filter((name) => !assigned.has(String(name)));
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
  const isLeader =
    state.user.role === "관리자" || state.user.role === "인도자";
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
        const isParticipant = state.participantsToday.includes(name);
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
  if (meta) {
    const rows = state.data.assignments || [];
    if (rows.length) {
      const date =
        rows[0]["날짜"] ||
        rows[0]["date"] ||
        rows[0]["Date"] ||
        "";
      const dateText = formatAssignmentDate(date);
      meta.textContent = dateText
        ? `배정된 날짜: ${dateText}`
        : "오늘 차량 배정이 저장되어 있습니다.";
    } else {
      meta.textContent = "저장된 차량 배정이 없습니다.";
    }
  }
  renderSelectedParticipants();
  renderCarAssignmentsPanel();
};

const loadApiUrl = () => {
  state.apiUrl = DEFAULT_API_URL;
  if (elements.apiUrlInput) {
    elements.apiUrlInput.value = DEFAULT_API_URL;
  }
  elements.configPanel.classList.add("hidden");
};

const saveApiUrl = () => {
  state.apiUrl = DEFAULT_API_URL;
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
    if (state.currentMenu === "visits" && state.completionExpandedAreaId) {
      renderVisitsView();
    }
    if (state.currentMenu === "car-assign") {
      resetCarAssignmentsFromSaved();
      renderCarAssignPopup();
    }
  } finally {
    setLoading(false);
  }
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
    if (
      areaCompletionStatus(grouped[areaId]) &&
      state.user &&
      (state.user.role === "인도자" || state.user.role === "관리자")
    ) {
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
    const google = document.createElement("a");
    google.href = `https://www.google.com/maps/search/?api=1&query=${encoded}`;
    google.target = "_blank";
    google.rel = "noopener noreferrer";
    google.textContent = "구글";
    nav.append(kakao, naver, google);
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
      const cardId = normalizeId(card["카드번호"]);
      const allRows = (state.data.visits || []).filter((row) => {
        const rawCard =
          row["구역카드"] || row["카드번호"] || row["구역카드번호"] || "";
        const rowCardId = normalizeId(rawCard);
        return rowCardId && rowCardId === cardId;
      });
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
    if (!state.scrollToSelectedCard) {
      expanded.scrollIntoView({ behavior: "smooth", block: "start" });
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
  if (!state.user) {
    elements.adminPanel.classList.add("hidden");
    return;
  }
  const showAdmin =
    state.user.role === "관리자" || state.user.role === "인도자";
  if (
    !showAdmin ||
    !["admin-cards", "admin-ev", "admin-banned"].includes(state.currentMenu)
  ) {
    elements.adminPanel.classList.add("hidden");
    return;
  }
  const overlayTitleEl =
    elements.adminOverlay &&
    elements.adminOverlay.querySelector("#admin-overlay-title");
  if (overlayTitleEl) {
    if (state.currentMenu === "admin-cards") {
      overlayTitleEl.textContent = "구역카드 관리";
    } else if (state.currentMenu === "admin-ev") {
      overlayTitleEl.textContent = "전도인 명단 관리";
    } else if (state.currentMenu === "admin-banned") {
      overlayTitleEl.textContent = "방문금지 관리";
    } else {
      overlayTitleEl.textContent = "";
    }
  }
  elements.adminPanel.classList.remove("hidden");
  const byAreaCards = groupCardsByArea();
  const areaIds = Object.keys(byAreaCards).sort((a, b) => {
    const na = Number(a);
    const nb = Number(b);
    if (!Number.isNaN(na) && !Number.isNaN(nb)) {
      return na - nb;
    }
    return String(a).localeCompare(String(b), "ko-KR");
  });
  elements.completionList.innerHTML = "";
  const selectedAreaId = state.adminCardsAreaId;
  const selectedCardKey = state.adminCardsSelectedKey;
  areaIds.forEach((areaId) => {
    const cards = byAreaCards[areaId] || [];
    const item = document.createElement("div");
    item.className = "list-item";
    item.dataset.areaId = areaId;
    if (selectedAreaId === areaId) {
      item.classList.add("active");
    }
    const header = document.createElement("div");
    header.textContent = `${areaId} · 카드 ${cards.length}장`;
    item.appendChild(header);
    const cardsBox = document.createElement("div");
    const expanded = selectedAreaId === areaId;
    if (!expanded) {
      cardsBox.style.display = "none";
    }
    if (expanded) {
      const cardTable = document.createElement("table");
      cardTable.className = "admin-table admin-card-table";
      cardTable.dataset.areaId = areaId;
      const cardThead = document.createElement("thead");
      const cardHeadRow = document.createElement("tr");
      ["카드번호", "주소", "상세주소", "비고", "작업"].forEach((text) => {
        const th = document.createElement("th");
        th.textContent = text;
        cardHeadRow.appendChild(th);
      });
      cardThead.appendChild(cardHeadRow);
      const cardTbody = document.createElement("tbody");
      cards.forEach((card) => {
        const tr = document.createElement("tr");
        const cardNo = String(card["카드번호"] || "");
        tr.dataset.areaId = areaId;
        tr.dataset.cardNumber = cardNo;
        const noTd = document.createElement("td");
        const noInput = document.createElement("input");
        noInput.type = "text";
        noInput.value = cardNo;
        noInput.readOnly = true;
        noTd.appendChild(noInput);
        const addrTd = document.createElement("td");
        const addrInput = document.createElement("input");
        addrInput.type = "text";
        addrInput.value = String(card["주소"] || "");
        addrInput.dataset.field = "address";
        addrTd.appendChild(addrInput);
        const detailTd = document.createElement("td");
        const detailInput = document.createElement("input");
        detailInput.type = "text";
        detailInput.value = String(card["상세주소"] || "");
        detailInput.dataset.field = "detailAddress";
        detailTd.appendChild(detailInput);
        const memoTd = document.createElement("td");
        const memoInput = document.createElement("input");
        memoInput.type = "text";
        memoInput.value = String(card["비고"] || "");
        memoInput.dataset.field = "memo";
        memoTd.appendChild(memoInput);
        const actionTd = document.createElement("td");
        actionTd.className = "admin-actions-cell";
        const saveBtn = document.createElement("button");
        saveBtn.type = "button";
        saveBtn.textContent = "저장";
        saveBtn.dataset.cardAction = "save-card";
        const deleteBtn = document.createElement("button");
        deleteBtn.type = "button";
        deleteBtn.textContent = "삭제";
        deleteBtn.dataset.cardAction = "delete-card";
        actionTd.appendChild(saveBtn);
        actionTd.appendChild(deleteBtn);
      const hasFlags =
        isTrueValue(card["재방"]) ||
        isTrueValue(card["연구"]) ||
        isTrueValue(card["6개월"]) ||
        isTrueValue(card["방문금지"]);
      if (hasFlags) {
        const clearBtn = document.createElement("button");
        clearBtn.type = "button";
        clearBtn.textContent = "해제";
        clearBtn.dataset.cardAction = "clear-flags";
        actionTd.appendChild(clearBtn);
      }
        tr.append(noTd, addrTd, detailTd, memoTd, actionTd);
        cardTbody.appendChild(tr);
      });
      const newTr = document.createElement("tr");
      newTr.dataset.new = "true";
      newTr.dataset.areaId = areaId;
      const newNoTd = document.createElement("td");
      const newNoInput = document.createElement("input");
      newNoInput.type = "text";
      newNoInput.placeholder = "새 카드번호";
      newNoInput.dataset.field = "cardNumber";
      newNoTd.appendChild(newNoInput);
      const newAddrTd = document.createElement("td");
      const newAddrInput = document.createElement("input");
      newAddrInput.type = "text";
      newAddrInput.dataset.field = "address";
      newAddrTd.appendChild(newAddrInput);
      const newDetailTd = document.createElement("td");
      const newDetailInput = document.createElement("input");
      newDetailInput.type = "text";
      newDetailInput.dataset.field = "detailAddress";
      newDetailTd.appendChild(newDetailInput);
      const newMemoTd = document.createElement("td");
      const newMemoInput = document.createElement("input");
      newMemoInput.type = "text";
      newMemoInput.dataset.field = "memo";
      newMemoTd.appendChild(newMemoInput);
      const newActionTd = document.createElement("td");
      newActionTd.className = "admin-actions-cell";
      const newSaveBtn = document.createElement("button");
      newSaveBtn.type = "button";
      newSaveBtn.textContent = "추가";
      newSaveBtn.dataset.cardAction = "create-card";
      newActionTd.appendChild(newSaveBtn);
      newTr.append(newNoTd, newAddrTd, newDetailTd, newMemoTd, newActionTd);
      cardTbody.appendChild(newTr);
      cardTable.append(cardThead, cardTbody);
      cardsBox.appendChild(cardTable);
    }
    item.appendChild(cardsBox);
    header.addEventListener("click", () => {
      if (state.adminCardsAreaId === areaId) {
        state.adminCardsAreaId = null;
        state.adminCardsSelectedKey = null;
      } else {
        state.adminCardsAreaId = areaId;
        state.adminCardsSelectedKey = null;
      }
      renderAdminPanel();
    });
    elements.completionList.appendChild(item);
  });

  const adminSections =
    elements.adminPanel.querySelectorAll(".admin-grid > div");
  if (adminSections.length === 3) {
    const [cardsSection, evSection, bannedSection] = adminSections;
    cardsSection.style.display =
      state.currentMenu === "admin-cards" ? "" : "none";
    evSection.style.display =
      state.currentMenu === "admin-ev" ? "" : "none";
    bannedSection.style.display =
      state.currentMenu === "admin-banned" ? "" : "none";
  }

  // 전도인 명단 관리: 표 형식 편집 (차량 정보 포함)
  const evangelists = state.data.evangelists || [];
  elements.evangelistList.innerHTML = "";
  const table = document.createElement("table");
  table.className = "admin-table admin-table-wide";
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  [
    "이름",
    "성별",
    "농인",
    "역할",
    "운전자",
    "정원",
    "부부",
    "비밀번호",
    "작업"
  ].forEach((text) => {
    const th = document.createElement("th");
    th.textContent = text;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  const tbody = document.createElement("tbody");
  evangelists.forEach((row) => {
    const name = row["이름"] || "";
    const role = row["역할"] || row["권한"] || "";
    const gender = row["성별"] || "";
    const isDeaf = isTrueValue(row["농인"]);
    const isDriver = isTrueValue(row["운전자"]);
    const cap = row["차량"] != null ? Number(row["차량"]) || 0 : 0;
    const spouse = row["부부"] || "";
    const tr = document.createElement("tr");
    tr.dataset.name = name;
    const nameTd = document.createElement("td");
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.value = name;
    nameInput.readOnly = true;
    nameTd.appendChild(nameInput);
    const genderTd = document.createElement("td");
    const genderInput = document.createElement("input");
    genderInput.type = "text";
    genderInput.value = gender;
    genderInput.dataset.field = "gender";
    genderTd.appendChild(genderInput);
    const deafTd = document.createElement("td");
    const deafInput = document.createElement("input");
    deafInput.type = "checkbox";
    deafInput.checked = isDeaf;
    deafInput.dataset.field = "deaf";
    deafTd.appendChild(deafInput);
    const roleTd = document.createElement("td");
    const roleInput = document.createElement("input");
    roleInput.type = "text";
    roleInput.value = role;
    roleInput.dataset.field = "role";
    roleTd.appendChild(roleInput);
    const driverTd = document.createElement("td");
    const driverInput = document.createElement("input");
    driverInput.type = "checkbox";
    driverInput.checked = isDriver;
    driverInput.dataset.field = "driver";
    driverTd.appendChild(driverInput);
    const capTd = document.createElement("td");
    const capInput = document.createElement("input");
    capInput.type = "number";
    capInput.min = "0";
    capInput.value = cap > 0 ? String(cap) : "";
    capInput.dataset.field = "capacity";
    capTd.appendChild(capInput);
    const spouseTd = document.createElement("td");
    const spouseInput = document.createElement("input");
    spouseInput.type = "text";
    spouseInput.value = spouse;
    spouseInput.dataset.field = "spouse";
    spouseTd.appendChild(spouseInput);
    const pwTd = document.createElement("td");
    const pwInput = document.createElement("input");
    pwInput.type = "text";
    pwInput.value = row["비밀번호"] != null ? String(row["비밀번호"]) : "";
    pwInput.dataset.field = "password";
    pwTd.appendChild(pwInput);
    const actionTd = document.createElement("td");
    actionTd.className = "admin-actions-cell";
    const saveBtn = document.createElement("button");
    saveBtn.type = "button";
    saveBtn.textContent = "저장";
    saveBtn.dataset.action = "save-ev";
    saveBtn.dataset.name = name;
    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.textContent = "삭제";
    deleteBtn.dataset.action = "delete-ev";
    deleteBtn.dataset.name = name;
    actionTd.appendChild(saveBtn);
    actionTd.appendChild(deleteBtn);
    tr.appendChild(nameTd);
    tr.appendChild(genderTd);
    tr.appendChild(deafTd);
    tr.appendChild(roleTd);
    tr.appendChild(driverTd);
    tr.appendChild(capTd);
    tr.appendChild(spouseTd);
    tr.appendChild(pwTd);
    tr.appendChild(actionTd);
    tbody.appendChild(tr);
  });
  const newTr = document.createElement("tr");
  newTr.dataset.new = "true";
  const newNameTd = document.createElement("td");
  const newNameInput = document.createElement("input");
  newNameInput.type = "text";
  newNameInput.placeholder = "새 전도인 이름";
  newNameInput.dataset.field = "name";
  newNameTd.appendChild(newNameInput);
  const newGenderTd = document.createElement("td");
  const newGenderInput = document.createElement("input");
  newGenderInput.type = "text";
  newGenderInput.dataset.field = "gender";
  newGenderTd.appendChild(newGenderInput);
  const newDeafTd = document.createElement("td");
  const newDeafInput = document.createElement("input");
  newDeafInput.type = "checkbox";
  newDeafInput.dataset.field = "deaf";
  newDeafTd.appendChild(newDeafInput);
  const newRoleTd = document.createElement("td");
  const newRoleInput = document.createElement("input");
  newRoleInput.type = "text";
  newRoleInput.dataset.field = "role";
  newRoleTd.appendChild(newRoleInput);
  const newDriverTd = document.createElement("td");
  const newDriverInput = document.createElement("input");
  newDriverInput.type = "checkbox";
  newDriverInput.dataset.field = "driver";
  newDriverTd.appendChild(newDriverInput);
  const newCapTd = document.createElement("td");
  const newCapInput = document.createElement("input");
  newCapInput.type = "number";
  newCapInput.min = "0";
  newCapInput.dataset.field = "capacity";
  newCapTd.appendChild(newCapInput);
  const newSpouseTd = document.createElement("td");
  const newSpouseInput = document.createElement("input");
  newSpouseInput.type = "text";
  newSpouseInput.dataset.field = "spouse";
  newSpouseTd.appendChild(newSpouseInput);
  const newPwTd = document.createElement("td");
  const newPwInput = document.createElement("input");
  newPwInput.type = "text";
  newPwInput.dataset.field = "password";
  newPwTd.appendChild(newPwInput);
  const newActionTd = document.createElement("td");
  newActionTd.className = "admin-actions-cell";
  const newSaveBtn = document.createElement("button");
  newSaveBtn.type = "button";
  newSaveBtn.textContent = "추가";
  newSaveBtn.dataset.action = "create-ev";
  newActionTd.appendChild(newSaveBtn);
  newTr.appendChild(newNameTd);
  newTr.appendChild(newGenderTd);
  newTr.appendChild(newDeafTd);
  newTr.appendChild(newRoleTd);
  newTr.appendChild(newDriverTd);
  newTr.appendChild(newCapTd);
  newTr.appendChild(newSpouseTd);
  newTr.appendChild(newPwTd);
  newTr.appendChild(newActionTd);
  tbody.appendChild(newTr);
  table.appendChild(thead);
  table.appendChild(tbody);
  elements.evangelistList.appendChild(table);
  const bannedCards = state.data.cards.filter(
    (card) => isTrueValue(card["방문금지"]) || isTrueValue(card["6개월"])
  );
  elements.bannedCardList.innerHTML = "";
  const bannedByArea = {};
  bannedCards.forEach((card) => {
    const areaId = String(card["구역번호"] || "");
    if (!areaId) {
      return;
    }
    if (!bannedByArea[areaId]) {
      bannedByArea[areaId] = [];
    }
    bannedByArea[areaId].push(card);
  });
  const bannedAreaIds = Object.keys(bannedByArea).sort((a, b) => {
    const na = Number(a);
    const nb = Number(b);
    if (!Number.isNaN(na) && !Number.isNaN(nb)) {
      return na - nb;
    }
    return String(a).localeCompare(String(b), "ko-KR");
  });
  const selectedBannedArea = state.adminBannedAreaId;
  const selectedBannedKey = state.adminBannedSelectedKey;
  bannedAreaIds.forEach((areaId) => {
    const list = bannedByArea[areaId] || [];
    const item = document.createElement("div");
    item.className = "list-item";
    item.dataset.areaId = areaId;
    if (selectedBannedArea === areaId) {
      item.classList.add("active");
    }
    const header = document.createElement("div");
    header.textContent = `${areaId} · 카드 ${list.length}장`;
    item.appendChild(header);
    const box = document.createElement("div");
    const expanded = selectedBannedArea === areaId;
    if (!expanded) {
      box.style.display = "none";
    }
    if (expanded) {
      const table = document.createElement("table");
      table.className = "admin-table";
      const thead = document.createElement("thead");
      const headRow = document.createElement("tr");
      ["카드번호", "상태", "선택"].forEach((text) => {
        const th = document.createElement("th");
        th.textContent = text;
        headRow.appendChild(th);
      });
      thead.appendChild(headRow);
      const tbody = document.createElement("tbody");
      list.forEach((card) => {
        const tr = document.createElement("tr");
        const cardNo = String(card["카드번호"] || "");
        const key = `${areaId}__${cardNo}`;
        const cardTd = document.createElement("td");
        cardTd.textContent = cardNo;
        const stateTd = document.createElement("td");
        const tags = [];
        if (isTrueValue(card["방문금지"])) {
          tags.push("방문금지");
        }
        if (isTrueValue(card["6개월"])) {
          tags.push("6개월");
        }
        stateTd.textContent = tags.join(", ");
        const actionTd = document.createElement("td");
        const selectBtn = document.createElement("button");
        selectBtn.type = "button";
        selectBtn.textContent = "선택";
        selectBtn.addEventListener("click", () => {
          state.adminBannedAreaId = areaId;
          state.adminBannedSelectedKey = key;
          renderAdminPanel();
        });
        actionTd.appendChild(selectBtn);
        tr.append(cardTd, stateTd, actionTd);
        tbody.appendChild(tr);
      });
      table.append(thead, tbody);
      box.appendChild(table);
      const editorKey =
        selectedBannedKey && selectedBannedKey.startsWith(`${areaId}__`)
          ? selectedBannedKey
          : null;
      if (editorKey) {
        const target = list.find(
          (card) =>
            `${areaId}__${String(card["카드번호"] || "")}` === editorKey
        );
        if (target) {
          const editor = document.createElement("div");
          const info = document.createElement("div");
          info.textContent = `${areaId}, 카드 ${String(
            target["카드번호"] || ""
          )}`;
          const stateText = document.createElement("div");
          const tags = [];
          if (isTrueValue(target["방문금지"])) {
            tags.push("방문금지");
          }
          if (isTrueValue(target["6개월"])) {
            tags.push("6개월");
          }
          stateText.textContent = `현재 상태: ${tags.join(", ")}`;
          const clearBtn = document.createElement("button");
          clearBtn.type = "button";
          clearBtn.textContent = "해제";
          clearBtn.addEventListener("click", async () => {
            try {
              const res = await apiRequest("updateCardFlags", {
                areaId,
                cardNumber: String(target["카드번호"] || ""),
                sixMonths: false,
                banned: false
              });
              if (!res.success) {
                alert(res.message || "카드 상태 변경에 실패했습니다.");
                return;
              }
              const cards = state.data.cards || [];
              const found = cards.find(
                (c) =>
                  String(c["구역번호"] || "") === String(areaId) &&
                  String(c["카드번호"] || "") ===
                    String(target["카드번호"] || "")
              );
              if (found) {
                found["6개월"] = res.sixMonths;
                found["방문금지"] = res.banned;
              }
              renderAreas();
              renderCards();
              state.adminBannedSelectedKey = null;
              renderAdminPanel();
              setStatus("방문금지/6개월 상태가 해제되었습니다.");
            } catch (e) {
              alert("카드 상태 변경 중 오류가 발생했습니다.");
            }
          });
          editor.appendChild(info);
          editor.appendChild(stateText);
          editor.appendChild(clearBtn);
          box.appendChild(editor);
        }
      }
    }
    item.appendChild(box);
    header.addEventListener("click", () => {
      if (state.adminBannedAreaId === areaId) {
        state.adminBannedAreaId = null;
        state.adminBannedSelectedKey = null;
      } else {
        state.adminBannedAreaId = areaId;
        state.adminBannedSelectedKey = null;
      }
      renderAdminPanel();
    });
    elements.bannedCardList.appendChild(item);
  });
};

const selectArea = (areaId) => {
  state.selectedArea = areaId;
  state.selectedCard = null;
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
    : card["초대장"]
    ? "초대장"
    : card["6개월"]
    ? "6개월"
    : card["방문금지"]
    ? "방문금지"
    : "만남";
  elements.visitNote.value = "";
  updateVisitFlagButtons();
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
    state.data.assignments = data.assignments || [];
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
  if (state.user.role === "관리자" || state.user.role === "인도자") {
    elements.menuToggle.style.display = "inline-block";
  } else {
    elements.menuToggle.style.display = "none";
  }
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
  ensureAssignmentState();
  renderAdminPanel();
  renderMyCarInfo();
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
      updatedCard["6개월"] =
        res?.cardUpdate?.sixMonths ?? (result === "6개월");
      updatedCard["방문금지"] =
        res?.cardUpdate?.banned ?? (result === "방문금지");
      updatedCard["재방"] =
        res?.cardUpdate?.revisit ?? (result === "재방");
      updatedCard["연구"] =
        res?.cardUpdate?.study ?? (result === "연구");
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
if (elements.refreshButton) {
  elements.refreshButton.addEventListener("click", () => {
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
      const res = await apiRequest("updateCardFlags", {
        areaId: state.selectedArea,
        cardNumber: state.selectedCard["카드번호"],
        revisit: false
      });
      if (!res.success) {
        alert(res.message || "재방 해제에 실패했습니다.");
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
      const res = await apiRequest("updateCardFlags", {
        areaId: state.selectedArea,
        cardNumber: state.selectedCard["카드번호"],
        study: false
      });
      if (!res.success) {
        alert(res.message || "연구 해제에 실패했습니다.");
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
      const res = await apiRequest("updateCardFlags", {
        areaId: state.selectedArea,
        cardNumber: state.selectedCard["카드번호"],
        sixMonths: false
      });
      if (!res.success) {
        alert(res.message || "6개월 해제에 실패했습니다.");
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
      const res = await apiRequest("updateCardFlags", {
        areaId: state.selectedArea,
        cardNumber: state.selectedCard["카드번호"],
        banned: false
      });
      if (!res.success) {
        alert(res.message || "방문금지 해제에 실패했습니다.");
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
    if (!state.participantsToday.includes(name)) {
      state.participantsToday.push(name);
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
  const name = item.dataset.name || "";
  if (!name) {
    return;
  }
  const exists = state.participantsToday.includes(name);
  if (exists) {
    state.participantsToday = state.participantsToday.filter((n) => n !== name);
    item.classList.remove("selected");
    const cars = state.carAssignments || [];
    cars.forEach((car) => {
      car.members = (car.members || []).filter((n) => n !== name);
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
  const name = parsed.name;
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
    fromCar.members = (fromCar.members || []).filter((n) => n !== name);
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

elements.carAssignPanel.addEventListener("click", (event) => {
  const column = event.target.closest(".car-column");
  if (!column) {
    return;
  }
  const zone = column.querySelector(".car-members");
  if (!zone) {
    return;
  }
  const member = event.target.closest(".car-member");
  const toCarId = zone.dataset.carId || "";
  if (!toCarId) {
    return;
  }
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
  if (
    (key === "admin-cards" ||
      key === "admin-ev" ||
      key === "admin-banned" ||
      key === "car-assign") &&
    (!state.user ||
      (state.user.role !== "관리자" && state.user.role !== "인도자"))
  ) {
    alert("관리자 또는 인도자만 사용할 수 있습니다.");
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
  } else if (
    key === "admin-cards" ||
    key === "admin-ev" ||
    key === "admin-banned"
  ) {
    if (elements.adminOverlay) {
      elements.adminOverlay.classList.remove("hidden");
      document.body.style.overflow = "hidden";
    }
    renderAdminPanel();
  } else if (key === "car-assign") {
    resetCarAssignmentsFromSaved();
    elements.carAssignOverlay.classList.remove("hidden");
    document.body.style.overflow = "hidden";
    renderCarAssignPopup();
  }
});

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

if (elements.adminCardAdd) {
  elements.adminCardAdd.addEventListener("click", async () => {
    const areaId = window.prompt("구역번호를 입력해 주세요.");
    if (!areaId) {
      return;
    }
    const cardNumber = window.prompt("카드번호를 입력해 주세요.");
    if (!cardNumber) {
      return;
    }
    const address = window.prompt("주소를 입력해 주세요. (없으면 빈칸)");
    const detailAddress = window.prompt("상세주소를 입력해 주세요. (없으면 빈칸)");
    const memo = window.prompt("비고를 입력해 주세요. (없으면 빈칸)");
    try {
      const res = await apiRequest("upsertCard", {
        areaId,
        cardNumber,
        address: address || "",
        detailAddress: detailAddress || "",
        memo: memo || ""
      });
      if (!res.success) {
        alert(res.message || "구역카드 저장에 실패했습니다.");
        return;
      }
      state.data.cards = res.cards || [];
      renderAreas();
      renderCards();
      renderAdminPanel();
      setStatus("구역카드가 저장되었습니다.");
    } catch (e) {
      alert("구역카드 저장 중 오류가 발생했습니다.");
    }
  });
}

if (elements.adminCardDelete) {
  elements.adminCardDelete.addEventListener("click", async () => {
    const areaId = window.prompt("삭제할 구역번호를 입력해 주세요.");
    if (!areaId) {
      return;
    }
    const cardNumber = window.prompt("삭제할 카드번호를 입력해 주세요.");
    if (!cardNumber) {
      return;
    }
    if (!window.confirm(`${areaId}, 카드 ${cardNumber}를 삭제하시겠습니까?`)) {
      return;
    }
    try {
      const res = await apiRequest("deleteCard", {
        areaId,
        cardNumber
      });
      if (!res.success) {
        alert(res.message || "구역카드 삭제에 실패했습니다.");
        return;
      }
      state.data.cards = res.cards || [];
      renderAreas();
      renderCards();
      renderAdminPanel();
      setStatus("구역카드가 삭제되었습니다.");
    } catch (e) {
      alert("구역카드 삭제 중 오류가 발생했습니다.");
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
      const res = await apiRequest("upsertEvangelist", {
        name,
        gender: gender || "",
        role: role || "",
        driver: driverFlag,
        capacity: capacity === "" ? "" : String(capacity),
        spouse: spouse || "",
        password: password || ""
      });
      if (!res.success) {
        alert(res.message || "전도인 저장에 실패했습니다.");
        return;
      }
      state.data.evangelists = res.evangelists || [];
      renderAdminPanel();
      setStatus("전도인 정보가 저장되었습니다.");
    } catch (e) {
      alert("전도인 저장 중 오류가 발생했습니다.");
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
      const res = await apiRequest("deleteEvangelist", { name });
      if (!res.success) {
        alert(res.message || "전도인 삭제에 실패했습니다.");
        return;
      }
      state.data.evangelists = res.evangelists || [];
      renderAdminPanel();
      setStatus("전도인이 삭제되었습니다.");
    } catch (e) {
      alert("전도인 삭제 중 오류가 발생했습니다.");
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
        const res = await apiRequest("upsertEvangelist", payload);
        if (!res.success) {
          alert(res.message || "전도인 저장에 실패했습니다.");
          return;
        }
        state.data.evangelists = res.evangelists || [];
        renderAdminPanel();
        setStatus("전도인이 추가되었습니다.");
      } catch (e) {
        alert("전도인 저장 중 오류가 발생했습니다.");
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
        const res = await apiRequest("upsertEvangelist", payload);
        if (!res.success) {
          alert(res.message || "전도인 저장에 실패했습니다.");
          return;
        }
        state.data.evangelists = res.evangelists || [];
        renderAdminPanel();
        setStatus("전도인 정보가 저장되었습니다.");
      } catch (e) {
        alert("전도인 저장 중 오류가 발생했습니다.");
      }
    } else if (action === "delete-ev") {
      if (!window.confirm(`${name} 전도인을 삭제하시겠습니까?`)) {
        return;
      }
      try {
        const res = await apiRequest("deleteEvangelist", { name });
        if (!res.success) {
          alert(res.message || "전도인 삭제에 실패했습니다.");
          return;
        }
        state.data.evangelists = res.evangelists || [];
        renderAdminPanel();
        setStatus("전도인이 삭제되었습니다.");
      } catch (e) {
        alert("전도인 삭제 중 오류가 발생했습니다.");
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
    const rowEl = button.closest("tr");
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
      const payload = {
        areaId,
        cardNumber,
        address: addressInput ? addressInput.value : "",
        detailAddress: detailInput ? detailInput.value : "",
        memo: memoInput ? memoInput.value : ""
      };
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
      const payload = {
        areaId,
        cardNumber: baseCardNumber,
        address: addressInput ? addressInput.value : "",
        detailAddress: detailInput ? detailInput.value : "",
        memo: memoInput ? memoInput.value : ""
      };
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
        setStatus("구역카드가 저장되었습니다.");
      } catch (e) {
        alert("구역카드 저장 중 오류가 발생했습니다.");
      }
    } else if (action === "delete-card") {
      if (
        !window.confirm(
          `구역 ${areaId}, 카드 ${baseCardNumber}를 삭제하시겠습니까?`
        )
      ) {
        return;
      }
      try {
        const res = await apiRequest("deleteCard", {
          areaId,
          cardNumber: baseCardNumber
        });
        if (!res.success) {
          alert(res.message || "구역카드 삭제에 실패했습니다.");
          return;
        }
        state.data.cards = res.cards || [];
        renderAreas();
        renderCards();
        renderAdminPanel();
        setStatus("구역카드가 삭제되었습니다.");
      } catch (e) {
        alert("구역카드 삭제 중 오류가 발생했습니다.");
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
        const res = await apiRequest("updateCardFlags", payload);
        if (!res.success) {
          alert(res.message || "상태 해제에 실패했습니다.");
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
        alert("상태 해제 중 오류가 발생했습니다.");
      } finally {
        setLoading(false);
      }
    }
  });
}

elements.carAssignAuto.addEventListener("click", () => {
  autoAssignCars();
  renderCarAssignPopup();
});

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
      .register("service-worker.js")
      .catch(() => {});
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
