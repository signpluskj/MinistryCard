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

        await deleteCardInSupabase(areaId, baseCardNumber);

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
      } finally {
        button.textContent = originalText;
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

          if (supabaseClient) {
            try {
              await restoreDeletedCardInSupabase(areaId, cardNumber);
            } catch (supaErr) {
              console.warn("Supabase restore failed:", supaErr);
            }
          }

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

          if (supabaseClient) {
            try {
              await purgeDeletedCardInSupabase(areaId, cardNumber);
            } catch (supaErr) {
              console.warn("Supabase purge failed:", supaErr);
            }
          }

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
