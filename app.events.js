if (elements.saveConfig) elements.saveConfig.addEventListener("click", saveConfig);
if (elements.syncToSupabase) elements.syncToSupabase.addEventListener("click", migrateToSupabase);
if (elements.syncToSheets) elements.syncToSheets.addEventListener("click", migrateToSheets);
if (elements.backupToExcel) elements.backupToExcel.addEventListener("click", exportToExcel);
if (elements.loginButton) elements.loginButton.addEventListener("click", login);

if (elements.appTitle) {
  elements.appTitle.addEventListener("click", () => {
    if (!state.user) return;
    refreshAll();
  });
}

if (elements.menuToggle) {
  elements.menuToggle.addEventListener("click", (e) => {
    e.stopPropagation();
    elements.sideMenu.classList.remove("hidden");
  });
}

if (elements.menuClose) {
  elements.menuClose.addEventListener("click", () => {
    elements.sideMenu.classList.add("hidden");
  });
}

if (elements.sideMenu) {
  elements.sideMenu.addEventListener("click", (event) => {
    if (event.target === elements.sideMenu) {
      elements.sideMenu.classList.add("hidden");
    }
  });

  elements.sideMenu.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-menu]");
    if (!button) return;

    const key = button.dataset.menu;
    if (!key) return;

    const isAdmin = state.user && state.user.role === "관리자";
    const isLeader = state.user && state.user.role === "인도자";

    if ((["admin-cards", "admin-ev", "admin-banned", "admin-deleted", "invite-campaign"].includes(key)) && !isAdmin) {
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
    } else if (["admin-cards", "admin-ev", "admin-banned", "admin-deleted"].includes(key)) {
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
}

if (elements.closeAreas) {
  elements.closeAreas.addEventListener("click", () => {
    elements.areaOverlay.classList.add("hidden");
  });
}

if (elements.areaOverlay) {
  elements.areaOverlay.addEventListener("click", (event) => {
    if (event.target === elements.areaOverlay) {
      elements.areaOverlay.classList.add("hidden");
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

if (elements.closeVolunteer) {
  elements.closeVolunteer.addEventListener("click", () => {
    elements.volunteerOverlay.classList.add("hidden");
    document.body.style.overflow = "";
  });
}

if (elements.volunteerOverlay) {
  elements.volunteerOverlay.addEventListener("click", (event) => {
    if (event.target === elements.volunteerOverlay) {
      elements.volunteerOverlay.classList.add("hidden");
      document.body.style.overflow = "";
    }
  });
}

if (elements.closeCarAssign) {
  elements.closeCarAssign.addEventListener("click", () => {
    elements.carAssignOverlay.classList.add("hidden");
    document.body.style.overflow = "";
  });
}

if (elements.carAssignOverlay) {
  elements.carAssignOverlay.addEventListener("click", (event) => {
    if (event.target === elements.carAssignOverlay) {
      elements.carAssignOverlay.classList.add("hidden");
      document.body.style.overflow = "";
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

if (elements.closeCarSelect) {
  elements.closeCarSelect.addEventListener("click", () => {
    elements.carSelectOverlay.classList.add("hidden");
  });
}

if (elements.carSelectOverlay) {
  elements.carSelectOverlay.addEventListener("click", (event) => {
    if (event.target === elements.carSelectOverlay) {
      elements.carSelectOverlay.classList.add("hidden");
    }
  });
}

if (elements.closeConfig) {
  elements.closeConfig.addEventListener("click", () => {
    elements.configPanel.classList.add("hidden");
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

if (elements.userInfo) {
  elements.userInfo.addEventListener("click", () => {
    if (!state.user) return;
    if (window.confirm("로그아웃 하시겠습니까?")) {
      if (typeof window.logout === "function") {
        window.logout();
      }
    }
  });
}

loadApiUrl();
