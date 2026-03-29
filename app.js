const DEFAULT_API_URL =
  (typeof window !== "undefined" &&
    window.__ENV__ &&
    window.__ENV__.API_URL) ||
  "https://script.google.com/macros/s/AKfycbx1O_Ab6n1j2py4A6Qck48fSY8N_J1wTiZF97y09HzW21kHgCinR1K2rCWiZmONG8mJ/exec";

let SUPABASE_URL = (window.__ENV__ && window.__ENV__.SUPABASE_URL) || "";
let SUPABASE_KEY = (window.__ENV__ && window.__ENV__.SUPABASE_KEY) || "";

// Remove common prefixes if present (sb_publishable_ for example)
if (SUPABASE_KEY.startsWith("sb_publishable_")) {
  SUPABASE_KEY = SUPABASE_KEY.replace("sb_publishable_", "");
}
if (SUPABASE_KEY.startsWith("sb_secret_")) {
  SUPABASE_KEY = SUPABASE_KEY.replace("sb_secret_", "");
}

let supabaseClient = null;
if (typeof window.supabase !== "undefined" && SUPABASE_URL && SUPABASE_KEY) {
  supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);
} else if (typeof supabasejs !== "undefined" && SUPABASE_URL && SUPABASE_KEY) {
  supabaseClient = supabasejs.createClient(SUPABASE_URL, SUPABASE_KEY);
}

const state = {
  apiUrl: (window.__ENV__ && window.__ENV__.API_URL) || DEFAULT_API_URL,
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
  congregationName: "",
  areaOrder: [],
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
  scrollCarAssignToActive: false,
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
  congregationNameInput: document.getElementById("congregation-name-input"),
  areaOrderInput: document.getElementById("area-order-input"),
  saveConfig: document.getElementById("save-config"),
  closeConfig: document.getElementById("close-config"),
  syncToSupabase: document.getElementById("sync-to-supabase"),
  syncCardInfoOnly: document.getElementById("sync-card-info-only"),
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
  visitDelete: document.getElementById("visit-delete"),
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
  adminCardsList: document.getElementById("admin-cards-list"),
  completionList: document.getElementById("admin-completions-list"),
  visitList: document.getElementById("visit-list"),

  evangelistList: document.getElementById("admin-evangelist-list"),
  carAssignPanel: document.getElementById("car-assign-panel"),
  bannedCardList: document.getElementById("banned-card-list"),
  deletedCardList: document.getElementById("admin-deleted-card-list"),
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

const updateAppTitle = () => {
  const titleTextEl = document.getElementById("app-title-text");
  if (titleTextEl) {
    const cong = state.congregationName || "";
    // Requirement: '회중 이름 + 봉사 카드'
    const titleText = cong ? `${cong} 봉사 카드` : "봉사 카드 정보";
    titleTextEl.textContent = titleText;
  }
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

  // '배정 삭제' 버튼 추가: 선택된 카드 중 하나라도 차량에 배정되어 있다면 표시
  const anyAssigned = cardNumbers.some(num => {
    const card = state.data.cards.find(c => 
      String(c["구역번호"]) === String(areaId) && String(c["카드번호"]) === String(num)
    );
    return card && getCardAssignedCarIdForDate(card, todayISO());
  });

  if (anyAssigned) {
    const unassignBtn = document.createElement("button");
    unassignBtn.type = "button";
    unassignBtn.className = "car-select-btn unassign-btn";
    unassignBtn.style.backgroundColor = "#fee2e2"; // 연한 빨간색 배경
    unassignBtn.style.color = "#b91c1c"; // 어두운 빨간색 글자
    unassignBtn.style.borderColor = "#fecaca";
    unassignBtn.textContent = "배정 삭제";
    unassignBtn.addEventListener("click", async () => {
      if (!confirm("선택한 카드들의 차량 배정을 삭제할까요?")) return;
      
      setLoading(true, "차량 배정 삭제 중...");
      try {
        if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
        
        for (const num of cardNumbers) {
          const { error: updateError } = await supabaseClient
            .from("cards")
            .update({ car_id: "", assignment_date: null })
            .eq("area_id", String(areaId))
            .eq("card_number", String(num));
            
          if (updateError) throw updateError;
          
          // 로컬 상태 즉시 업데이트
          const localCard = state.data.cards.find(c => 
            String(c["구역번호"]) === String(areaId) && String(c["카드번호"]) === String(num)
          );
          if (localCard) {
            localCard["차량"] = "";
            localCard["배정날짜"] = "";
          }
        }
        
        state.selectedCards = [];
        renderAreas();
        renderCards();
        renderAdminPanel();
        renderMyCarInfo();
        setStatus("차량 배정이 삭제되었습니다.");
        elements.carSelectOverlay.classList.add("hidden");
      } catch (err) {
        console.error("Unassign cards error:", err);
        alert("배정 삭제에 실패했습니다: " + err.message);
      } finally {
        setLoading(false);
      }
    });
    elements.carSelectList.appendChild(unassignBtn);
  }

  elements.carSelectOverlay.classList.remove("hidden");
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
  const name = normalizeAssignmentName(state.user.name);
  
  const showNoInfo = () => {
    box.innerHTML = '<div class="car-info-main">오늘 배정된 차량 정보가 없습니다.</div>';
    box.classList.remove("car-info-clickable");
    box.classList.remove("hidden");
    box.onclick = null;
  };

  if (!rows.length || !name) {
    showNoInfo();
    return;
  }
  
  // 현재 사용자가 운전자이거나 동승자인 행 찾기
  const myRow = rows.find(row => {
    const driver = normalizeAssignmentName(row["이름"]);
    const passengers = (row["동승자"] || []).map(p => normalizeAssignmentName(p));
    return driver === name || passengers.includes(name);
  }) || null;

  if (!myRow) {
    showNoInfo();
    return;
  }

  // 배정 정보가 있는 경우 클릭 가능하도록 설정
  box.classList.add("car-info-clickable");
  box.onclick = () => {
    state.carDashboardExpanded = !state.carDashboardExpanded;
    renderMyCarInfo();
  };

  if (state.carDashboardExpanded) {
    renderFullCarDashboard(box, rows);
    return;
  }
  
  const carId = String(myRow["차량"] || "");
  const driverName = normalizeAssignmentName(myRow["이름"]);
  const passengers = (myRow["동승자"] || []).map(p => normalizeAssignmentName(p)).filter(Boolean);
  
  const textParts = [];
  textParts.push(`차량 ${carId}`);
  if (driverName) {
    textParts.push(`운전자: ${driverName}`);
  }
  if (passengers.length) {
    textParts.push(passengers.join(", "));
  }
  const lines = [];
  lines.push(`<div class="car-info-main" style="display:flex; justify-content:space-between; align-items:center;">
    <span>${textParts.join(" | ")}</span>
    <span style="font-size:10px; font-weight:bold; color:#16a34a;">▼</span>
  </div>`);
  
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

const renderFullCarDashboard = (box, rows) => {
  const lines = [];
  lines.push(`<div class="car-info-main" style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; border-bottom:1px solid #bbf7d0; padding-bottom:6px;">
    <span style="color:#166534">오늘의 전체 배정 현황</span>
    <span style="font-size:10px; font-weight:normal; color:#16a34a;">▲</span>
  </div>`);

  const carGroups = {};
  rows.forEach(row => {
    const cId = String(row["차량"] || "미지정");
    if (!carGroups[cId]) carGroups[cId] = { driver: "", passengers: [], cards: [] };
    carGroups[cId].driver = normalizeAssignmentName(row["이름"]);
    carGroups[cId].passengers = (row["동승자"] || []).map(p => normalizeAssignmentName(p)).filter(Boolean);
  });

  // 카드 정보 매칭
  const allCards = state.data.cards || [];
  allCards.forEach(card => {
    const cId = getCardAssignedCarIdForDate(card, todayISO());
    if (cId && carGroups[cId]) {
      carGroups[cId].cards.push(String(card["카드번호"] || ""));
    }
  });

  const sortedCarIds = Object.keys(carGroups).sort((a, b) => {
    const na = parseInt(a), nb = parseInt(b);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return a.localeCompare(b);
  });

  sortedCarIds.forEach(cId => {
    const g = carGroups[cId];
    lines.push(`<div class="car-dashboard-row" style="margin-bottom:12px;">`);
    lines.push(`  <div style="font-size:15px; color:black;"><strong>차량 ${cId}</strong> | 운전자: ${g.driver}${g.passengers.length ? " | " + g.passengers.join(", ") : ""}</div>`);
    if (g.cards.length) {
      g.cards.sort((a, b) => compareCardNumbers(a, b));
      lines.push(`  <div style="font-size:14px; color:#166534; margin-top:2px; padding-left:10px; border-left:2px solid #86efac;">${g.cards.join(", ")}</div>`);
    } else {
      lines.push(`  <div style="font-size:14px; color:#94a3b8; margin-top:2px; padding-left:10px; border-left:2px solid #e2e8f0;">배정된 카드 없음</div>`);
    }
    lines.push(`</div>`);
  });

  box.innerHTML = lines.join("");
  box.classList.remove("hidden");
};

const getEvangelistByName = (name) =>
  (state.data.evangelists || []).find(
    (row) => String(row["이름"] || "") === String(name || "")
  ) || null;

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
    const driver = normalizeAssignmentName(row["이름"] || "");
    const passengers = (row["동승자"] || []).map(n => normalizeAssignmentName(n));
    
    if (!carId || !driver) {
      return;
    }
    
    if (!byCar[carId]) {
      byCar[carId] = {
        carId,
        driver: driver,
        capacity: 0,
        members: [driver, ...passengers]
      };
      const ev = getEvangelistByName(driver);
      const cap = ev && ev["차량"] != null ? Number(ev["차량"]) || 0 : 0;
      byCar[carId].capacity = cap || 0;
    }
  });
  state.carAssignments = Object.values(byCar);
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
    // 운전자가 members에 포함되어 있는 경우, 동승자 명단에서는 제외하여 중복 저장 방지
    const driverName = normalizeAssignmentName(car.driver || "");
    const passengerList = (car.members || [])
      .map(n => normalizeAssignmentName(n))
      .filter(n => n && n !== driverName);

    newAssignments.push({
      date: currentDate,
      slot: currentSlot,
      car_id: String(car.carId),
      driver: car.driver || "",
      passengers: passengerList
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

  if (activeBtn && state.scrollCarAssignToActive) {
    requestAnimationFrame(() => {
      activeBtn.scrollIntoView({ behavior: "auto", inline: "start", block: "nearest" });
      state.scrollCarAssignToActive = false;
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
  
  const assignmentNames = [];
  (state.data.assignments || []).forEach(row => {
    if (row["이름"]) assignmentNames.push(normalizeAssignmentName(row["이름"]));
    if (row["동승자"]) {
      (row["동승자"] || []).forEach(p => assignmentNames.push(normalizeAssignmentName(p)));
    }
  });

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
  state.scrollCarAssignToActive = true;
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
  state.scrollCarAssignToActive = true;
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

const moveMember = (name, fromCarId, toCarId) => {
  if (!name) return;

  const cars = state.carAssignments || [];
  const fromCar = fromCarId ? cars.find((c) => String(c.carId) === String(fromCarId)) : null;
  const toCar = cars.find((c) => String(c.carId) === String(toCarId));

  if (!toCar) return;

  if (fromCar) {
    if (fromCarId === toCarId) {
      // 같은 차량 내에서의 이동: 해당 인원을 가장 앞으로 보내서 운전자로 변경
      const otherMembers = (fromCar.members || []).filter((m) => m !== name);
      fromCar.members = [name, ...otherMembers];
      fromCar.driver = name;
    } else {
      // 다른 차량으로 이동
      fromCar.members = (fromCar.members || []).filter((m) => m !== name);
      if (fromCar.driver === name) {
        fromCar.driver = fromCar.members.length > 0 ? fromCar.members[0] : "";
      }

      if (!toCar.members.includes(name)) {
        toCar.members.push(name);
      }
      
      // 대상 차량에 운전자가 없으면 운전자로 설정
      if (!toCar.driver) {
        toCar.driver = name;
      }
    }
  } else {
    // 미배정 명단에서 차량으로 추가
    if (!toCar.members.includes(name)) {
      toCar.members.push(name);
    }
    if (!toCar.driver) {
      toCar.driver = name;
    }
  }

  state.carAssignments = cars;
  renderCarAssignmentsPanel();
  renderSelectedParticipants(); // 미배정 명단 갱신
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
    header.addEventListener("click", async () => {
      // 모바일 지원을 위해 dblclick 대신 click으로 변경하고 안내 문구 추가
      if (!window.confirm(`차량 ${car.carId}에 카드를 수동으로 배정하시겠습니까?`)) return;
      
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
    (car.members || []).forEach((name, index) => {
      const item = document.createElement("div");
      item.className = "car-member";
      if (index === 0) {
        item.classList.add("is-driver");
        item.innerHTML = `<strong>${name}</strong>`;
      } else {
        item.textContent = name;
      }
      item.draggable = true;
      item.dataset.name = name;
      item.dataset.carId = car.carId;

      // 드래그 시작
      item.addEventListener("dragstart", (e) => {
        item.classList.add("dragging");
        e.dataTransfer.setData("text/plain", JSON.stringify({
          name: name,
          fromCarId: car.carId
        }));
      });

      item.addEventListener("dragend", () => {
        item.classList.remove("dragging");
      });

      // 터치/클릭 선택 및 이동 처리
      item.addEventListener("click", (e) => {
        e.stopPropagation();
        
        // 이미 다른 차량의 인원이 선택되어 있다면 이 차량으로 이동
        if (carAssignTapSelection && carAssignTapSelection.name !== name) {
          moveMember(carAssignTapSelection.name, carAssignTapSelection.fromCarId, car.carId);
          carAssignTapSelection = null;
          return;
        }

        // 선택 해제 또는 새로운 선택
        const isCurrentlySelected = item.classList.contains("tap-selected");
        document.querySelectorAll(".car-member, .selected-person").forEach(el => el.classList.remove("tap-selected"));
        
        if (isCurrentlySelected) {
          carAssignTapSelection = null;
        } else {
          item.classList.add("tap-selected");
          carAssignTapSelection = { name, fromCarId: car.carId };
        }
      });

      membersBox.appendChild(item);
    });

    // ... (중략: 드래그 이벤트는 유지됨)

    // 터치 이동 처리 (차량 헤더나 빈 공간 클릭 시 이동 대상 확정)
    col.addEventListener("click", () => {
      if (carAssignTapSelection) {
        moveMember(carAssignTapSelection.name, carAssignTapSelection.fromCarId, car.carId);
        carAssignTapSelection = null;
      }
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
    
    // 클릭하여 선택
    item.addEventListener("click", (e) => {
      e.stopPropagation();
      const isCurrentlySelected = item.classList.contains("tap-selected");
      document.querySelectorAll(".car-member, .selected-person").forEach(el => el.classList.remove("tap-selected"));
      
      if (isCurrentlySelected) {
        carAssignTapSelection = null;
      } else {
        item.classList.add("tap-selected");
        carAssignTapSelection = { name, fromCarId: null }; // fromCarId가 null이면 미배정에서 이동함을 의미
      }
    });
    
    frag.appendChild(item);
  });
  box.appendChild(frag);
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

  // Add volunteer signup features below the assigned names
  const selectedParticipantsBox = elements.carAssignSelected;
  if (selectedParticipantsBox) {
    renderSelectedParticipants();
  }

  // 기존 버튼 제거 후 새로 추가 (중복 방지)
  const existingAction = elements.carAssignOverlay.querySelector(".car-assign-signup-actions");
  if (existingAction) existingAction.remove();

  // '신청자 추가/취소' 버튼을 car-assign-selected 박스 아래에 추가
  const signupActionContainer = document.createElement("div");
  signupActionContainer.className = "car-assign-signup-actions";
  signupActionContainer.style.marginTop = "10px";
  signupActionContainer.style.padding = "0 15px"; // 좌우 여백
  signupActionContainer.style.width = "100%";
  signupActionContainer.style.boxSizing = "border-box";

  const addBtn = document.createElement("button");
  addBtn.type = "button";
  addBtn.className = "volunteer-admin-add-btn";
  addBtn.textContent = "+ 신청자 추가/취소";
  addBtn.style.fontSize = "14px";
  addBtn.style.padding = "10px";
  addBtn.style.width = "100%"; // 가로 전체 사이즈
  addBtn.style.display = "block";

  addBtn.addEventListener("click", async () => {
    const allDays = buildVolunteerDayList();
    const selectedDay = allDays.find(d => d.isoDate === (state.carAssignDate || todayISO())) || null;
    if (!selectedDay) {
      alert("선택된 날짜의 정보를 찾을 수 없습니다.");
      return;
    }
    const slotName = state.carAssignSlot || "오전";
    const entries = (selectedDay.participantEntries || []).filter(e => (e.slot || "오전") === slotName);

    const evangelists = (state.data.evangelists || []).map((row) => String(row["이름"] || "").trim()).filter((name) => !!name);
    const uniqueEvangelists = Array.from(new Set(evangelists));
    const initialSelectedNames = new Set(
      entries.map((entry) => String(entry.name || "").trim()).filter((name) => !!name)
    );
    const allSelectableNames = Array.from(
      new Set(uniqueEvangelists.concat(Array.from(initialSelectedNames)))
    ).filter((name) => !!name);
    const workingSelectedNames = new Set(initialSelectedNames);

    const existingPicker = elements.carAssignOverlay.querySelector(".volunteer-admin-add-picker");
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
      if (!customName) return;
      if (!allSelectableNames.includes(customName)) allSelectableNames.push(customName);
      workingSelectedNames.add(customName);
      customInput.value = "";
      renderNameButtons();
    };
    customAddBtn.addEventListener("click", addCustomName);
    customInput.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); addCustomName(); } });

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
        if (!before && after) namesToAdd.push(name);
        else if (before && !after) namesToRemove.push(name);
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
          namesToRemove.forEach(name => {
            dayObj.participantEntries = dayObj.participantEntries.filter(e => !(String(e.name || "") === String(name) && (e.slot || "오전") === slotName));
          });
          namesToAdd.forEach(name => {
            const exists = dayObj.participantEntries.some(e => String(e.name || "") === String(name) && (e.slot || "오전") === slotName);
            if (!exists) dayObj.participantEntries.push({ name, slot: slotName });
          });
        }
        const saveRes = await saveVolunteerWeekToSupabase(weekStart, updatedData);
        if (!saveRes.success) throw new Error(saveRes.message);
        await loadVolunteerConfig();
        ensureVolunteerSelection();
        // 신청자 명단(state.participantsToday) 갱신 및 UI 리렌더링
        applyCarAssignDataForSlot(selectedDay.isoDate, slotName);
        setStatus("신청자 명단이 업데이트되었습니다.");
      } catch (err) {
        console.error("Update volunteers in car assign error:", err);
        alert("저장에 실패했습니다: " + err.message);
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

    signupActionContainer.after(picker);
  });

  signupActionContainer.appendChild(addBtn);
  // selectedParticipantsBox 다음에 배치 (비어있어도 표시되도록)
  selectedParticipantsBox.after(signupActionContainer);

  renderCarAssignmentsPanel();
};

const loadConfig = async () => {
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

  // Load other configs if Supabase is ready
  if (supabaseClient) {
    try {
      const { data: configs, error: configError } = await supabaseClient.from("site_config").select("*");
      if (configError) throw configError;
      
      if (configs && configs.length > 0) {
        const congConfig = configs.find(c => c.config_key === "congregation_name");
        if (congConfig && congConfig.config_value) {
          // JSONB에서 문자열을 가져올 때 따옴표가 포함될 수 있으므로 처리
          let val = congConfig.config_value;
          if (typeof val === 'string' && val.startsWith('"') && val.endsWith('"')) {
             val = JSON.parse(val);
          }
          state.congregationName = String(val || "");
          if (elements.congregationNameInput) {
            elements.congregationNameInput.value = state.congregationName;
          }
        }
        const orderConfig = configs.find(c => c.config_key === "area_order");
        if (orderConfig && orderConfig.config_value) {
          state.areaOrder = Array.isArray(orderConfig.config_value) ? orderConfig.config_value : [];
          if (elements.areaOrderInput) {
            elements.areaOrderInput.value = state.areaOrder.join(", ");
          }
        }
        updateAppTitle();
      }
    } catch (e) {
      console.warn("Failed to load site_config in loadConfig:", e);
    }
  }
  // 설정을 불러온 후 설정창을 숨김 (자동 로그인 시 잔상 제거)
  if (elements.configPanel) {
    elements.configPanel.classList.add("hidden");
  }
};

const saveConfig = async () => {
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

  const congName = elements.congregationNameInput.value.trim();
  const areaOrderText = elements.areaOrderInput ? elements.areaOrderInput.value.trim() : "";

  if (supabaseClient) {
    setLoading(true, "설정 저장 중...");
    try {
      // 회중 이름 저장
      const { error: congError } = await supabaseClient
        .from("site_config")
        .upsert(
          { config_key: "congregation_name", config_value: congName },
          { onConflict: "config_key" }
        );
      if (congError) throw congError;
      state.congregationName = congName;

      // 구역 정렬 순서 저장
      const orderArr = areaOrderText ? areaOrderText.split(",").map(s => s.trim()).filter(s => !!s) : [];
      const { error: orderError } = await supabaseClient
        .from("site_config")
        .upsert(
          { config_key: "area_order", config_value: orderArr },
          { onConflict: "config_key" }
        );
      if (orderError) throw orderError;
      state.areaOrder = orderArr;

      alert("설정이 저장되었습니다.");
      updateAppTitle();
      renderAreas();
    } catch (err) {
      console.error("Failed to save config to Supabase:", err);
      const detail = err.message || JSON.stringify(err);
      alert(`Supabase 설정 저장에 실패했습니다.\n사유: ${detail}`);
    } finally {
      setLoading(false);
    }
  }
 else {
    alert("API URL이 로컬에 저장되었습니다. (Supabase 연결 안됨)");
  }
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
      if (elements.completionOverlay && !elements.completionOverlay.classList.contains("hidden")) {
        renderCompletionOverlayList();
      } else {
        renderVisitsView();
      }
    }
    if (state.currentMenu === "car-assign") {
      await setCarAssignDate(state.carAssignDate || todayISO());
    }
  } finally {
    setLoading(false);
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
    
    const weekChipsContainer = document.createElement("div");
    weekChipsContainer.className = "volunteer-admin-week-chips-container";

    const sortedWeeks = (weeks || []).slice().sort((a, b) => {
      const aISO = a.weekStartISO || a.weekStartText || "";
      const bISO = b.weekStartISO || b.weekStartText || "";
      return aISO.localeCompare(bISO);
    });
    sortedWeeks.forEach((w) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "volunteer-admin-week-chip";
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
      chip.textContent = label;
      if (
        String(w.weekStartText || "") ===
        String(state.selectedVolunteerWeekStart || "")
      ) {
        chip.classList.add("active");
      }
      chip.addEventListener("click", () => {
        const value = w.weekStartText || "";
        state.selectedVolunteerWeekStart = value;
        const target =
          days.find((d) => String(d.weekStartText || "") === String(value)) ||
          null;
        if (target) {
          state.selectedVolunteerDate = target.isoDate || "";
        }
        renderVolunteerOverlay();
      });

      bindLongPress(chip, async () => {
        const ok = window.confirm(`[${label}] 주간의 설정을 전체 삭제할까요?\n(해당 주간의 모든 날짜와 신청 기록이 삭제됩니다.)`);
        if (!ok) return;
        setLoading(true, "주간 설정을 삭제하는 중...");
        try {
          const weekStart = w.weekStartText || w.weekStartISO;
          const res = await deleteVolunteerWeekFromSupabase(weekStart);
          if (!res.success) throw new Error(res.message);
          
          await loadVolunteerConfig();
          ensureVolunteerSelection();
          renderVolunteerOverlay();
        } catch (err) {
          console.error("Delete volunteer week error:", err);
          alert("삭제에 실패했습니다: " + err.message);
        } finally {
          setLoading(false);
        }
      });
      weekChipsContainer.appendChild(chip);
    });
    weekRow.appendChild(weekChipsContainer);
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
      
      const activeSlots = [];
      const weekdayOrder = ["월", "화", "수", "목", "금", "토", "일"];
      for (let i = 1; i <= 7; i++) {
        if (slotsConfig[i].am || slotsConfig[i].pm) {
          activeSlots.push({
            ...slotsConfig[i],
            weekday: weekdayOrder[i - 1]
          });
        }
      }

      if (activeSlots.length === 0) {
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
            const weekday = selectedDay.weekday;
            
            // 1. Update weekday memo (fixedMemo) for ALL weeks
            const batchRes = await batchUpdateVolunteerWeekdayMemos(
              weekday, 
              slotName, 
              fixedMemoInput.value || ""
            );
            if (!batchRes.success) throw new Error(batchRes.message);

            // 2. Update date memo (extraMemo) for current week
            const { data: weekRow, error: fetchError } = await supabaseClient
              .from("volunteer_weeks")
              .select("data")
              .eq("week_start", weekStart)
              .single();
              
            if (fetchError) throw fetchError;
            
            const updatedData = { ...weekRow.data };
            (updatedData.days || []).forEach(d => {
              if (d.isoDate === selectedDay.isoDate) {
                if (slotName === "오후") {
                  d.extraMemoPM = extraMemoInput.value || "";
                } else {
                  d.extraMemoAM = extraMemoInput.value || "";
                }
              }
            });
            
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
    const areaInfo = state.data.areas.find(a => String(a["구역번호"]) === String(areaId));
    const inProgress = areaInfo && areaInfo["시작날짜"] && !areaInfo["완료날짜"];
    if (inProgress) {
      return areaId;
    }
  }
  return null;
};

const getAllInProgressAreas = () => {
  const grouped = groupCardsByArea();
  const areaIds = Object.keys(grouped);
  const result = [];
  for (const areaId of areaIds) {
    if (isKslArea(areaId)) {
      continue;
    }
    const areaInfo = state.data.areas.find(a => String(a["구역번호"]) === String(areaId));
    const inProgress = areaInfo && areaInfo["시작날짜"] && !areaInfo["완료날짜"];
    if (inProgress) {
      result.push(areaId);
    }
  }
  return result;
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
  const areaIds = Object.keys(grouped).sort((a, b) => compareAreaIds(a, b, state.areaOrder));
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
      // 구역 이름 뒤 인도자 이름 제거
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
      const range = isComplete || hasDone ? doneText || startText : startText || doneText;
      const isKsl = isKslArea(areaId);
      if (isComplete || hasDone) {
        const labelParts = [];
        if (range) {
          labelParts.push(range);
        }
        labelParts.push("완료");
        const stateChip = document.createElement("span");
        stateChip.className = "area-date-range";
        stateChip.textContent = labelParts.join(" ");
        metaRow.appendChild(stateChip);
      } else if (inProgress) {
        const labelParts = [];
        if (range) {
          labelParts.push(range);
        }
        labelParts.push("시작");
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

// 앱 초기화 실행
loadConfig().then(() => {
  if (typeof tryAutoLogin === "function") {
    tryAutoLogin();
  }
});
