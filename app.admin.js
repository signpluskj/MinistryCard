var escapeHtml = (typeof window !== "undefined" && window.escapeHtml) ? window.escapeHtml : (str) => {
  if (str == null) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
};

// ===================================================
// 1. 완료 내역 (Completion History) 통합 뷰 컴포넌트
// ===================================================
const renderCompletionHistoryView = (container) => {
  if (!container) return;
  container.innerHTML = "";

  const completions = (state.data.completions || []).slice();
  if (!completions.length) {
    container.innerHTML = '<div class="list-item" style="text-align: center; color: #64748b; padding: 24px;">등록된 완료 내역이 없습니다.</div>';
    return;
  }

  // 상태 초기화
  if (!state.completionsViewMode) state.completionsViewMode = "by-area";
  if (!state.completionsYearFilter) state.completionsYearFilter = "all";
  if (!state.completionsSearch) state.completionsSearch = "";

  const currentYear = new Date().getFullYear().toString();
  const allYears = Array.from(
    new Set(
      completions
        .map((c) => {
          const d = parseVisitDate(c["완료날짜"] || c["completeDate"]);
          return d ? d.getFullYear().toString() : "";
        })
        .filter(Boolean)
    )
  ).sort((a, b) => b.localeCompare(a));

  // 통계 계산
  const thisYearCount = completions.filter((c) => {
    const d = parseVisitDate(c["완료날짜"] || c["completeDate"]);
    return d && d.getFullYear().toString() === currentYear;
  }).length;

  const sortedByRecent = completions.slice().sort((a, b) => {
    const da = parseVisitDate(a["완료날짜"] || a["completeDate"]);
    const db = parseVisitDate(b["완료날짜"] || b["completeDate"]);
    return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
  });
  const latestComp = sortedByRecent[0] || null;
  const latestAreaId = latestComp ? (latestComp["구역번호"] || latestComp["areaId"] || "") : "-";
  const latestDateText = latestComp && latestComp["완료날짜"] ? formatDate(latestComp["완료날짜"]) : "-";

  const root = document.createElement("div");
  root.className = "completions-dashboard";

  // 1. 상단 통계 카드 (줄바꿈 없이 한 줄에 알맞게 표시)
  const statsGrid = document.createElement("div");
  statsGrid.className = "completions-stats-grid";
  const areaLen = (latestAreaId || "").length;
  const areaFontSize = areaLen > 11 ? "14px" : (areaLen > 8 ? "15px" : "17px");

  statsGrid.innerHTML = `
    <div class="completions-stat-box">
      <div class="completions-stat-num">${thisYearCount}회</div>
      <div class="completions-stat-label">${currentYear}년 완료 구역</div>
    </div>
    <div class="completions-stat-box completions-stat-box-center">
      <div class="completions-stat-num area-stat-num" style="font-size: ${areaFontSize};" title="${escapeHtml(latestAreaId)}">${escapeHtml(latestAreaId)}</div>
      <div class="completions-stat-label">최근 완료 (${latestDateText})</div>
    </div>
    <div class="completions-stat-box">
      <div class="completions-stat-num">${completions.length}회</div>
      <div class="completions-stat-label">누적 전체 완료</div>
    </div>
  `;
  root.appendChild(statsGrid);

  // 2. 툴바 (보기 탭 & 필터)
  const toolbar = document.createElement("div");
  toolbar.className = "completions-toolbar";

  const viewTabs = document.createElement("div");
  viewTabs.className = "completions-view-tabs";

  const tabByArea = document.createElement("button");
  tabByArea.className = `completions-view-tab-btn ${state.completionsViewMode === "by-area" ? "active" : ""}`;
  tabByArea.textContent = "구역별 모아보기";
  tabByArea.onclick = () => {
    state.completionsViewMode = "by-area";
    renderCompletionHistoryView(container);
  };

  const tabTimeline = document.createElement("button");
  tabTimeline.className = `completions-view-tab-btn ${state.completionsViewMode === "timeline" ? "active" : ""}`;
  tabTimeline.textContent = "최근 완료순 타임라인";
  tabTimeline.onclick = () => {
    state.completionsViewMode = "timeline";
    renderCompletionHistoryView(container);
  };

  viewTabs.append(tabByArea, tabTimeline);
  toolbar.appendChild(viewTabs);

  const filterRow = document.createElement("div");
  filterRow.className = "completions-filters-row";

  const yearSelect = document.createElement("select");
  yearSelect.className = "admin-cards-filter-select";
  yearSelect.innerHTML = `<option value="all">전체 연도</option>` +
    allYears.map((y) => `<option value="${y}" ${state.completionsYearFilter === y ? "selected" : ""}>${y}년</option>`).join("");
  yearSelect.onchange = (e) => {
    state.completionsYearFilter = e.target.value;
    renderCompletionHistoryView(container);
  };

  const searchInput = document.createElement("input");
  searchInput.type = "text";
  searchInput.className = "admin-cards-search-box";
  searchInput.placeholder = "구역번호 또는 인도자 검색...";
  searchInput.value = state.completionsSearch || "";
  searchInput.oninput = (e) => {
    state.completionsSearch = e.target.value.trim().toLowerCase();
    renderCompletionListContent();
  };

  filterRow.append(yearSelect, searchInput);
  toolbar.appendChild(filterRow);
  root.appendChild(toolbar);

  // 3. 목록 컨테이너
  const listContainer = document.createElement("div");
  listContainer.className = "completions-list-container";
  root.appendChild(listContainer);

  const canEdit = isAdmin() || isLeader();

  const renderCompletionListContent = () => {
    listContainer.innerHTML = "";

    // 필터링 적용
    let filtered = completions.filter((c) => {
      const compDate = parseVisitDate(c["완료날짜"] || c["completeDate"]);
      const year = compDate ? compDate.getFullYear().toString() : "";
      if (state.completionsYearFilter !== "all" && year !== state.completionsYearFilter) {
        return false;
      }
      if (state.completionsSearch) {
        const areaId = String(c["구역번호"] || c["areaId"] || "").toLowerCase();
        const leader = String(c["인도자"] || "").toLowerCase();
        const memo = String(c["비고"] || "").toLowerCase();
        if (!areaId.includes(state.completionsSearch) && !leader.includes(state.completionsSearch) && !memo.includes(state.completionsSearch)) {
          return false;
        }
      }
      return true;
    });

    if (!filtered.length) {
      listContainer.innerHTML = '<div class="list-item" style="text-align: center; color: #94a3b8; padding: 20px;">조건에 맞는 완료 내역이 없습니다.</div>';
      return;
    }

    if (state.completionsViewMode === "timeline") {
      // 타임라인 뷰
      const timelineWrap = document.createElement("div");
      timelineWrap.className = "completions-timeline-list";

      const timelineSorted = filtered.sort((a, b) => {
        const da = parseVisitDate(a["완료날짜"] || a["completeDate"]);
        const db = parseVisitDate(b["완료날짜"] || b["completeDate"]);
        return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
      });

      timelineSorted.forEach((row) => {
        const card = document.createElement("div");
        card.className = "completions-timeline-card";

        const startD = parseVisitDate(row["시작날짜"]);
        const endD = parseVisitDate(row["완료날짜"] || row["completeDate"]);
        let durationText = "";
        if (startD && endD) {
          const days = Math.round((endD.getTime() - startD.getTime()) / (1000 * 60 * 60 * 24)) + 1;
          if (days > 0) durationText = `${days}일 소요`;
        }

        const areaId = row["구역번호"] || row["areaId"] || "";
        const dateRangeStr = row["시작날짜"] && row["완료날짜"]
          ? `${formatDate(row["시작날짜"])} → ${formatDate(row["완료날짜"])}`
          : formatDate(row["완료날짜"] || row["시작날짜"] || "");

        const infoBox = document.createElement("div");
        infoBox.className = "completions-timeline-info";
        infoBox.innerHTML = `
          <div class="completions-timeline-area-badge">
            <span>${areaId}구역</span>
            ${durationText ? `<span class="completions-days-badge">${durationText}</span>` : ""}
          </div>
          <div class="completions-timeline-date">📅 ${dateRangeStr}</div>
          <div class="completions-timeline-leader">👤 인도자: ${row["인도자"] || "없음"}${row["비고"] ? ` · 📝 ${row["비고"]}` : ""}</div>
        `;
        card.appendChild(infoBox);

        if (canEdit) {
          const actBox = document.createElement("div");
          actBox.className = "completions-actions";

          const editBtn = document.createElement("button");
          editBtn.className = "btn-small";
          editBtn.textContent = "수정";
          editBtn.style.padding = "3px 8px";
          editBtn.onclick = (e) => {
            e.stopPropagation();
            editCompletion(row);
          };

          const delBtn = document.createElement("button");
          delBtn.className = "btn-small btn-danger";
          delBtn.textContent = "삭제";
          delBtn.style.padding = "3px 8px";
          delBtn.onclick = (e) => {
            e.stopPropagation();
            deleteCompletion(row);
          };

          actBox.append(editBtn, delBtn);
          card.appendChild(actBox);
        }

        timelineWrap.appendChild(card);
      });
      listContainer.appendChild(timelineWrap);
    } else {
      // 구역별 모아보기 뷰 (넓고 시원한 행간격 적용)
      const byArea = {};
      filtered.forEach((row) => {
        const areaId = String(row["구역번호"] || row["areaId"] || "").trim();
        if (!areaId) return;
        if (!byArea[areaId]) byArea[areaId] = [];
        byArea[areaId].push(row);
      });

      const areaIds = Object.keys(byArea).sort((a, b) => compareAreaIds(a, b, state.areaOrder));
      const expandedId = state.completionExpandedAreaId;

      const byAreaWrap = document.createElement("div");
      byAreaWrap.className = "completions-by-area-list";

      areaIds.forEach((areaId) => {
        const item = document.createElement("div");
        item.className = "completions-area-item" + (expandedId === areaId ? " active" : "");

        const rows = byArea[areaId].slice().sort((a, b) => {
          const da = parseVisitDate(a["완료날짜"] || a["completeDate"]);
          const db = parseVisitDate(b["완료날짜"] || b["completeDate"]);
          return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
        });

        const recentDone = rows[0] && rows[0]["완료날짜"] ? formatDate(rows[0]["완료날짜"]) : "";

        const header = document.createElement("div");
        header.className = "completions-area-header";
        header.innerHTML = `
          <div class="completions-area-title-group">
            <span class="completions-area-title">${escapeHtml(areaId)}</span>
            <span class="completions-count-pill">${rows.length}회 완료</span>
          </div>
          <div class="completions-area-recent">
            <span>최근: <b>${recentDone || "-"}</b></span>
            <span class="completions-area-toggle-arrow">${expandedId === areaId ? "▲" : "▼"}</span>
          </div>
        `;
        item.appendChild(header);

        if (expandedId === areaId) {
          const detail = document.createElement("div");
          detail.className = "completions-area-detail-list";

          rows.forEach((row) => {
            const entry = document.createElement("div");
            entry.className = "completions-detail-entry";

            const startText = row["시작날짜"] ? formatDate(row["시작날짜"]) : "";
            const doneText = row["완료날짜"] ? formatDate(row["완료날짜"]) : "";
            const dateStr = startText && doneText ? `${startText} → ${doneText}` : doneText || startText;

            let durationBadge = "";
            const startD = parseVisitDate(row["시작날짜"]);
            const endD = parseVisitDate(row["완료날짜"]);
            if (startD && endD) {
              const days = Math.round((endD.getTime() - startD.getTime()) / (1000 * 60 * 60 * 24)) + 1;
              if (days > 0) {
                durationBadge = `<span class="completions-days-badge">${days}일 소요</span>`;
              }
            }

            const textSpan = document.createElement("div");
            textSpan.innerHTML = `
              <div style="font-weight: 700; font-size: 14px; color: #1e293b; display: flex; align-items: center; gap: 8px;">
                <span>📅 ${dateStr}</span>
                ${durationBadge}
              </div>
              <div style="font-size: 13px; color: #64748b; margin-top: 4px;">
                👤 인도자: <b>${escapeHtml(row["인도자"] || "없음")}</b>${row["비고"] ? ` · 📝 ${escapeHtml(row["비고"])}` : ""}
              </div>
            `;
            entry.appendChild(textSpan);

            if (canEdit) {
              const acts = document.createElement("div");
              acts.className = "completions-actions";

              const editBtn = document.createElement("button");
              editBtn.className = "admin-card-icon-btn edit-btn";
              editBtn.title = "수정";
              editBtn.textContent = "✏️";
              editBtn.onclick = (e) => {
                e.stopPropagation();
                editCompletion(row);
              };

              const delBtn = document.createElement("button");
              delBtn.className = "admin-card-icon-btn delete-btn";
              delBtn.title = "삭제";
              delBtn.textContent = "🗑️";
              delBtn.onclick = (e) => {
                e.stopPropagation();
                deleteCompletion(row);
              };

              acts.append(editBtn, delBtn);
              entry.appendChild(acts);
            }

            detail.appendChild(entry);
          });
          item.appendChild(detail);
        }

        item.onclick = (e) => {
          if (e.target.closest("button")) return;
          state.completionExpandedAreaId = state.completionExpandedAreaId === areaId ? null : areaId;
          renderCompletionListContent();
        };

        byAreaWrap.appendChild(item);
      });

      listContainer.appendChild(byAreaWrap);
    }
  };

  renderCompletionListContent();
  container.appendChild(root);
};

const renderCompletionOverlayList = () => {
  renderCompletionHistoryView(elements.completionAreaList);
};

const renderAdminCompletions = () => {
  renderCompletionHistoryView(document.getElementById("admin-completions-list"));
};

// ===================================================
// 2. 구역카드 관리 (Admin Cards Management) 전면 개선
// ===================================================
const openCardEditModal = (card = null, defaultAreaId = "") => {
  const isNew = !card;
  const overlay = document.createElement("div");
  overlay.className = "area-overlay";
  overlay.style.zIndex = "3500";

  const content = document.createElement("div");
  content.className = "area-overlay-content";
  content.style.maxWidth = "420px";
  content.style.margin = "auto";
  content.style.borderRadius = "12px";

  const areaIds = Object.keys(groupCardsByArea()).sort((a, b) => compareAreaIds(a, b, state.areaOrder));
  const currentAreaId = card ? String(card["구역번호"] || "") : defaultAreaId || (areaIds[0] || "");
  const cardNo = card ? String(card["카드번호"] || "") : "";

  content.innerHTML = `
    <div class="area-overlay-header">
      <span>${isNew ? "새 구역카드 등록" : `구역카드 수정 (${currentAreaId} - ${cardNo})`}</span>
      <button class="icon-button modal-close-btn" type="button">✕</button>
    </div>
    <div class="panel-body" style="padding: 16px; display: flex; flex-direction: column; gap: 10px;">
      <div style="display: flex; gap: 8px;">
        <label class="field" style="flex: 1;">
          <span>구역번호</span>
          ${isNew ? `
            <input type="text" id="modal-card-area" value="${currentAreaId}" placeholder="예: 101" list="modal-card-area-list" />
            <datalist id="modal-card-area-list">
              ${areaIds.map(a => `<option value="${a}">`).join("")}
            </datalist>
          ` : `
            <input type="text" id="modal-card-area" value="${currentAreaId}" readonly style="background: #f1f5f9;" />
          `}
        </label>
        <label class="field" style="flex: 1;">
          <span>카드번호</span>
          <input type="text" id="modal-card-num" value="${cardNo}" ${!isNew ? "readonly style='background: #f1f5f9;'" : "placeholder='예: 1, 101-1'"} />
        </label>
      </div>
      <label class="field">
        <span>읍면동</span>
        <input type="text" id="modal-card-town" value="${card ? (card["읍면동"] || "") : ""}" placeholder="예: 퇴계동, 신동" />
      </label>
      <label class="field">
        <span>주소</span>
        <input type="text" id="modal-card-addr" value="${card ? (card["주소"] || "") : ""}" placeholder="도로명 또는 지번 주소" />
      </label>
      <label class="field">
        <span>상세주소</span>
        <input type="text" id="modal-card-detail" value="${card ? (card["상세주소"] || "") : ""}" placeholder="동/호수, 건물명 등" />
      </label>
      <label class="field">
        <span>정보 / 메모</span>
        <input type="text" id="modal-card-memo" value="${card ? (card["정보"] || "") : ""}" placeholder="특이사항, 메모" />
      </label>
      <div style="margin-top: 4px;">
        <span style="font-size: 12px; font-weight: 700; color: #475569; display: block; margin-bottom: 6px;">특수 상태 설정</span>
        <div style="display: flex; flex-wrap: wrap; gap: 10px;">
          <label style="display: flex; align-items: center; gap: 4px; font-size: 13px; cursor: pointer;">
            <input type="checkbox" id="modal-card-revisit" ${card && isTrueValue(card["재방"]) ? "checked" : ""} /> 📌 재방
          </label>
          <label style="display: flex; align-items: center; gap: 4px; font-size: 13px; cursor: pointer;">
            <input type="checkbox" id="modal-card-study" ${card && isTrueValue(card["연구"]) ? "checked" : ""} /> 📖 연구
          </label>
          <label style="display: flex; align-items: center; gap: 4px; font-size: 13px; cursor: pointer;">
            <input type="checkbox" id="modal-card-six" ${card && isTrueValue(card["6개월"]) ? "checked" : ""} /> ⏳ 6개월
          </label>
          <label style="display: flex; align-items: center; gap: 4px; font-size: 13px; cursor: pointer; color: #ef4444;">
            <input type="checkbox" id="modal-card-banned" ${card && isTrueValue(card["방문금지"]) ? "checked" : ""} /> 🚫 방문금지
          </label>
        </div>
      </div>
      <div style="display: flex; gap: 8px; margin-top: 14px;">
        <button id="modal-card-save" style="flex: 1; padding: 10px; background: #2563eb; color: #fff; font-weight: 700; border-radius: 6px;">저장</button>
        <button id="modal-card-cancel" style="flex: 1; padding: 10px; background: #64748b; color: #fff; font-weight: 700; border-radius: 6px;">취소</button>
      </div>
    </div>
  `;

  overlay.appendChild(content);
  document.body.appendChild(overlay);

  const closeModal = () => overlay.remove();
  content.querySelector(".modal-close-btn").onclick = closeModal;
  content.querySelector("#modal-card-cancel").onclick = closeModal;

  content.querySelector("#modal-card-save").onclick = async () => {
    const targetArea = content.querySelector("#modal-card-area").value.trim();
    const targetNum = content.querySelector("#modal-card-num").value.trim();
    if (!targetArea) {
      alert("구역번호를 입력해 주세요.");
      return;
    }
    if (!targetNum) {
      alert("카드번호를 입력해 주세요.");
      return;
    }

    const payload = {
      area_id: String(targetArea),
      card_number: String(targetNum),
      town: content.querySelector("#modal-card-town").value.trim(),
      address: content.querySelector("#modal-card-addr").value.trim(),
      detail_address: content.querySelector("#modal-card-detail").value.trim(),
      memo: content.querySelector("#modal-card-memo").value.trim(),
      revisit: content.querySelector("#modal-card-revisit").checked,
      study: content.querySelector("#modal-card-study").checked,
      six_months: content.querySelector("#modal-card-six").checked,
      banned: content.querySelector("#modal-card-banned").checked,
      meet: card ? !!card["만남"] : false,
      absent: card ? !!card["부재"] : false
    };

    closeModal();
    setLoading(true, "구역카드를 저장하는 중...");
    try {
      if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

      const { error } = await supabaseClient
        .from("cards")
        .upsert(payload, { onConflict: "area_id, card_number" });

      if (error) throw error;

      await loadData();
      renderAreas();
      renderCards();
      renderAdminCards();
      setStatus(`구역 ${targetArea}의 카드 ${targetNum}이(가) 저장되었습니다.`);
    } catch (err) {
      console.error(err);
      alert("저장에 실패했습니다: " + err.message);
    } finally {
      setLoading(false);
    }
  };
};

const renderAdminCards = () => {
  const box = document.getElementById("admin-cards-list");
  if (!box) return;
  box.innerHTML = "";

  const allCards = (state.data.cards || []).slice();
  const byAreaCards = groupCardsByArea();
  const areaIds = Object.keys(byAreaCards).sort((a, b) => compareAreaIds(a, b, state.areaOrder));

  if (!state.adminCardsAreaId) state.adminCardsAreaId = "all";
  if (!state.adminCardsSearch) state.adminCardsSearch = "";
  if (!state.adminCardsFilter) state.adminCardsFilter = "all";

  // 통계 계산
  const totalCardsCount = allCards.length;
  const revisitCount = allCards.filter(c => isTrueValue(c["재방"])).length;
  const studyCount = allCards.filter(c => isTrueValue(c["연구"])).length;
  const sixCount = allCards.filter(c => isTrueValue(c["6개월"])).length;
  const bannedCount = allCards.filter(c => isTrueValue(c["방문금지"])).length;

  const root = document.createElement("div");
  root.className = "admin-cards-container";

  // 1. 상단 통계 띠
  const statsStrip = document.createElement("div");
  statsStrip.className = "admin-cards-stats-strip";
  statsStrip.innerHTML = `
    <span class="admin-cards-stat-pill">🗺️ 구역 <b>${areaIds.length}</b>개</span>
    <span class="admin-cards-stat-pill">🏷️ 카드 <b>${totalCardsCount}</b>장</span>
    <span class="admin-cards-stat-pill" style="color:#1d4ed8;">📌 재방 <b>${revisitCount}</b></span>
    <span class="admin-cards-stat-pill" style="color:#b45309;">📖 연구 <b>${studyCount}</b></span>
    <span class="admin-cards-stat-pill" style="color:#475569;">⏳ 6개월 <b>${sixCount}</b></span>
    <span class="admin-cards-stat-pill" style="color:#b91c1c;">🚫 방문금지 <b>${bannedCount}</b></span>
  `;
  root.appendChild(statsStrip);

  // 2. 툴바 (구역 선택, 검색, 새 카드 버튼)
  const toolbar = document.createElement("div");
  toolbar.className = "admin-cards-toolbar";

  const areaSelect = document.createElement("select");
  areaSelect.className = "admin-cards-filter-select";
  areaSelect.innerHTML = `<option value="all">전체 구역 (${areaIds.length})</option>` +
    areaIds.map(id => {
      const cnt = (byAreaCards[id] || []).length;
      return `<option value="${id}" ${state.adminCardsAreaId === id ? "selected" : ""}>${id}구역 (${cnt}장)</option>`;
    }).join("");
  areaSelect.onchange = (e) => {
    state.adminCardsAreaId = e.target.value;
    renderCardsGridContent();
  };

  const searchInput = document.createElement("input");
  searchInput.type = "text";
  searchInput.className = "admin-cards-search-box";
  searchInput.placeholder = "카드번호, 읍면동, 주소, 메모 실시간 검색...";
  searchInput.value = state.adminCardsSearch || "";
  searchInput.oninput = (e) => {
    state.adminCardsSearch = e.target.value.trim().toLowerCase();
    renderCardsGridContent();
  };

  const addCardBtn = document.createElement("button");
  addCardBtn.type = "button";
  addCardBtn.className = "admin-cards-add-btn";
  addCardBtn.innerHTML = `<span>+</span> 새 카드 등록`;
  addCardBtn.onclick = () => {
    const curArea = state.adminCardsAreaId !== "all" ? state.adminCardsAreaId : (areaIds[0] || "");
    openCardEditModal(null, curArea);
  };

  toolbar.append(areaSelect, searchInput, addCardBtn);
  root.appendChild(toolbar);

  // 3. 필터 탭 (전체, 재방, 연구, 6개월, 방문금지)
  const filterTabs = document.createElement("div");
  filterTabs.className = "invite-board-tabs";
  const filterOptions = [
    { key: "all", label: "전체" },
    { key: "revisit", label: `재방 (${revisitCount})` },
    { key: "study", label: `연구 (${studyCount})` },
    { key: "six", label: `6개월 (${sixCount})` },
    { key: "banned", label: `방문금지 (${bannedCount})` }
  ];
  filterOptions.forEach(opt => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `invite-tab-btn ${state.adminCardsFilter === opt.key ? "active" : ""}`;
    btn.textContent = opt.label;
    btn.onclick = () => {
      state.adminCardsFilter = opt.key;
      renderAdminCards();
    };
    filterTabs.appendChild(btn);
  });
  root.appendChild(filterTabs);

  // 4. 구역카드 목록 컨테이너 (구역카드 목록과 동일한 묶음 방식)
  const listContainer = document.createElement("div");
  listContainer.className = "admin-cards-list-view";
  root.appendChild(listContainer);

  const renderCardsGridContent = () => {
    listContainer.innerHTML = "";

    const filtered = allCards.filter((c) => {
      const aId = String(c["구역번호"] || "").trim();
      if (state.adminCardsAreaId !== "all" && aId !== state.adminCardsAreaId) {
        return false;
      }
      if (state.adminCardsFilter === "revisit" && !isTrueValue(c["재방"])) return false;
      if (state.adminCardsFilter === "study" && !isTrueValue(c["연구"])) return false;
      if (state.adminCardsFilter === "six" && !isTrueValue(c["6개월"])) return false;
      if (state.adminCardsFilter === "banned" && !isTrueValue(c["방문금지"])) return false;

      if (state.adminCardsSearch) {
        const num = String(c["카드번호"] || "").toLowerCase();
        const town = String(c["읍면동"] || "").toLowerCase();
        const addr = String(c["주소"] || "").toLowerCase();
        const detail = String(c["상세주소"] || "").toLowerCase();
        const memo = String(c["정보"] || "").toLowerCase();
        if (
          !num.includes(state.adminCardsSearch) &&
          !town.includes(state.adminCardsSearch) &&
          !addr.includes(state.adminCardsSearch) &&
          !detail.includes(state.adminCardsSearch) &&
          !memo.includes(state.adminCardsSearch)
        ) {
          return false;
        }
      }
      return true;
    });

    if (!filtered.length) {
      listContainer.innerHTML =
        '<div style="text-align: center; color: #94a3b8; padding: 36px 20px; background: #fff; border-radius: 8px; border: 1px dashed #cbd5e1;">조건에 맞는 구역카드가 없습니다.</div>';
      return;
    }

    // 구역별로 그룹화
    const groupedByArea = {};
    filtered.forEach((card) => {
      const aId = String(card["구역번호"] || "").trim();
      if (!groupedByArea[aId]) groupedByArea[aId] = [];
      groupedByArea[aId].push(card);
    });

    // 구역 정렬 (state.areaOrder 기준)
    const sortedAreaIds = Object.keys(groupedByArea).sort((a, b) =>
      compareAreaIds(a, b, state.areaOrder)
    );

    sortedAreaIds.forEach((areaId) => {
      const areaCards = groupedByArea[areaId];
      // 구역 내 카드 정렬 (카드번호 기준)
      areaCards.sort((ca, cb) => compareCardNumbers(ca["카드번호"], cb["카드번호"]));

      // 구역 메타정보 조회
      const areaInfo = (state.data.areas || []).find(
        (a) => String(a["구역번호"] || "").trim() === areaId
      );
      const startText = areaInfo && areaInfo["시작날짜"] ? formatDate(areaInfo["시작날짜"]) : "";
      const doneText = areaInfo && areaInfo["완료날짜"] ? formatDate(areaInfo["완료날짜"]) : "";
      const leader = areaInfo ? areaInfo["인도자"] || "" : "";

      let statusBadgeHtml = "";
      if (doneText) {
        statusBadgeHtml = `<span class="area-date-range">${doneText} 완료</span>`;
      } else if (startText) {
        statusBadgeHtml = `<span class="area-date-range" style="border-color:#f97316; color:#ea580c; background:#fffbeb;">${startText} 시작</span>`;
      }

      const groupEl = document.createElement("div");
      groupEl.className = "admin-area-group";

      // 1. 구역 상단 헤더 (구역카드 목록과 동일한 구조)
      const headerEl = document.createElement("div");
      headerEl.className = "admin-area-group-header";
      headerEl.innerHTML = `
        <div class="admin-area-group-left">
          <div class="admin-area-group-title">
            <span>${escapeHtml(areaId)}</span>
          </div>
          <div class="admin-area-group-sub">
            ${statusBadgeHtml}
            <span class="area-date-range">${areaCards.length}장</span>
            ${leader ? `<span class="area-date-range">👤 ${escapeHtml(leader)}</span>` : ""}
          </div>
        </div>
        <button type="button" class="admin-area-add-card-btn" title="이 구역에 새 카드 추가">+ 새 카드</button>
      `;

      headerEl.querySelector(".admin-area-add-card-btn").onclick = () => {
        openCardEditModal(null, areaId);
      };
      groupEl.appendChild(headerEl);

      // 2. 구역 내 카드 목록
      const cardsListEl = document.createElement("div");
      cardsListEl.className = "admin-area-cards-list";

      areaCards.forEach((card) => {
        const cardNo = String(card["카드번호"] || "");
        const town = String(card["읍면동"] || "");
        const addr = String(card["주소"] || "");
        const detail = String(card["상세주소"] || "");
        const memo = String(card["정보"] || "");
        const worker = String(card["정기방문자"] || "");

        const isRev = isTrueValue(card["재방"]);
        const isStd = isTrueValue(card["연구"]);
        const isSix = isTrueValue(card["6개월"]);
        const isBan = isTrueValue(card["방문금지"]);

        const fullAddr = [addr, detail].filter(Boolean).join(" ");
        const encoded = encodeURIComponent(fullAddr || cardNo);
        const kakaoUrl = `kakaomap://search?q=${encoded}`;
        const naverUrl = `https://m.map.naver.com/search2/search.naver?query=${encoded}`;
        const googleUrl = `https://www.google.com/maps/search/?api=1&query=${encoded}`;

        const cardRow = document.createElement("div");
        cardRow.className = `admin-card-row ${isBan ? "banned-card" : ""}`;

        cardRow.innerHTML = `
          <div class="admin-card-row-header">
            <span class="admin-card-row-title">${escapeHtml(cardNo)}${town ? ` <span style="font-weight: 500; font-size: 13px; color: #64748b;">(${escapeHtml(town)})</span>` : ""}</span>
            <div class="admin-card-row-actions">
              ${isBan ? '<span class="badge badge-danger">방문금지</span>' : ""}
              ${isStd ? '<span class="badge badge-study">연구</span>' : ""}
              ${isRev ? '<span class="badge badge-revisit">재방</span>' : ""}
              ${isSix ? '<span class="badge badge-sixmonths">6개월</span>' : ""}
              <button type="button" class="admin-card-icon-btn edit-btn" title="수정">✏️</button>
              <button type="button" class="admin-card-icon-btn delete-btn" title="삭제">🗑️</button>
            </div>
          </div>
          <div class="card-line" style="color: #334155;">
            ${addr ? escapeHtml(addr) : '<span style="color:#94a3b8;">주소 미등록</span>'} ${detail ? `<b>${escapeHtml(detail)}</b>` : ""}
          </div>
          <div class="card-line" style="color: #64748b; font-size: 13px;">
            최근방문: ${card["최근방문일"] ? formatDate(card["최근방문일"]) : "최근방문 없음"}${card["이전봉사일"] ? ` (이전: ${formatDate(card["이전봉사일"])})` : ""}
          </div>
          ${memo ? `<div class="card-line" style="color: #334155; font-size: 13px;">정보: ${escapeHtml(memo)}</div>` : ""}
          ${worker ? `<div class="card-line" style="color: #334155; font-size: 13px;">정기방문자: ${escapeHtml(worker)}</div>` : ""}
          <div class="nav-links">
            <a class="nav-kakao" href="${kakaoUrl}" target="_blank" rel="noopener noreferrer">카카오</a>
            <a class="nav-naver" href="${naverUrl}" target="_blank" rel="noopener noreferrer">네이버</a>
            <a href="${googleUrl}" target="_blank" rel="noopener noreferrer">구글</a>
          </div>
        `;

        // 수정 아이콘 클릭 이벤트
        cardRow.querySelector(".edit-btn").onclick = (e) => {
          e.stopPropagation();
          openCardEditModal(card, areaId);
        };

        // 삭제 아이콘 클릭 이벤트
        cardRow.querySelector(".delete-btn").onclick = async (e) => {
          e.stopPropagation();
          if (!confirm(`[${areaId}] 카드 ${cardNo}를 완전히 삭제할까요?`)) return;
          setLoading(true, "구역카드를 삭제하는 중...");
          try {
            if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");
            await deleteCardInSupabase(areaId, cardNo);
            await loadData();
            renderAreas();
            renderCards();
            renderAdminCards();
            setStatus(`${areaId} 카드 ${cardNo}가 삭제되었습니다.`);
          } catch (err) {
            console.error(err);
            alert("삭제에 실패했습니다: " + err.message);
          } finally {
            setLoading(false);
          }
        };

        cardsListEl.appendChild(cardRow);
      });

      groupEl.appendChild(cardsListEl);
      listContainer.appendChild(groupEl);
    });
  };

  renderCardsGridContent();
  box.appendChild(root);
};

const renderAdminEvangelists = () => {
  const box = document.getElementById("admin-evangelist-list");
  if (!box) return;
  box.innerHTML = "";

  const isSuper = isSuperAdmin();

  const allEvangelists = [...(state.data.evangelists || [])].sort((a, b) =>
    String(a["이름"] || "").localeCompare(String(b["이름"] || ""), "ko-KR")
  );

  const totalCount = allEvangelists.length;
  const superCount = allEvangelists.filter(
    (e) => (e["역할"] || "").trim() === "최고관리자"
  ).length;
  const adminCount = allEvangelists.filter(
    (e) => (e["역할"] || "").trim() === "관리자"
  ).length;
  const leaderCount = allEvangelists.filter(
    (e) => (e["역할"] || "").trim() === "인도자"
  ).length;
  const memberCount = allEvangelists.filter((e) => {
    const r = (e["역할"] || "").trim();
    return r !== "최고관리자" && r !== "관리자" && r !== "인도자";
  }).length;

  const currentFilter = state.adminEvRoleFilter || "all";
  const currentSearch = (state.adminEvSearch || "").trim().toLowerCase();

  // 필터 적용
  let filtered = allEvangelists;
  if (currentFilter === "최고관리자") {
    filtered = filtered.filter(
      (e) => (e["역할"] || "").trim() === "최고관리자"
    );
  } else if (currentFilter === "관리자") {
    filtered = filtered.filter((e) => (e["역할"] || "").trim() === "관리자");
  } else if (currentFilter === "인도자") {
    filtered = filtered.filter((e) => (e["역할"] || "").trim() === "인도자");
  } else if (currentFilter === "전도인") {
    filtered = filtered.filter((e) => {
      const r = (e["역할"] || "").trim();
      return r !== "최고관리자" && r !== "관리자" && r !== "인도자";
    });
  }

  if (currentSearch) {
    filtered = filtered.filter((e) => {
      const name = String(e["이름"] || "").toLowerCase();
      const spouse = String(e["부부"] || "").toLowerCase();
      const role = String(e["역할"] || "").toLowerCase();
      return (
        name.includes(currentSearch) ||
        spouse.includes(currentSearch) ||
        role.includes(currentSearch)
      );
    });
  }

  // 상단 컨트롤 래퍼
  const controlWrapper = document.createElement("div");
  controlWrapper.className = "admin-ev-controls";

  // 1. 탭 필터 & 실시간 검색 바
  const topBar = document.createElement("div");
  topBar.className = "admin-ev-top-bar";
  topBar.innerHTML = `
    <div class="admin-ev-filter-tabs">
      <button type="button" class="admin-ev-tab-btn ${currentFilter === "all" ? "active" : ""}" data-ev-filter="all">
        전체 (${totalCount})
      </button>
      <button type="button" class="admin-ev-tab-btn tab-super ${currentFilter === "최고관리자" ? "active" : ""}" data-ev-filter="최고관리자">
        👑 최고관리자 (${superCount})
      </button>
      <button type="button" class="admin-ev-tab-btn tab-admin ${currentFilter === "관리자" ? "active" : ""}" data-ev-filter="관리자">
        ⭐ 관리자 (${adminCount})
      </button>
      <button type="button" class="admin-ev-tab-btn tab-leader ${currentFilter === "인도자" ? "active" : ""}" data-ev-filter="인도자">
        🧭 인도자 (${leaderCount})
      </button>
      <button type="button" class="admin-ev-tab-btn tab-member ${currentFilter === "전도인" ? "active" : ""}" data-ev-filter="전도인">
        👤 전도인 (${memberCount})
      </button>
    </div>
    <div class="admin-ev-search-box">
      <input type="text" id="admin-ev-search" placeholder="🔍 이름/역할 검색..." value="${escapeAttr(state.adminEvSearch || "")}">
    </div>
  `;
  controlWrapper.appendChild(topBar);

  // 2. 신규 인원 등록 카드 (상단 배치)
  let defaultRoleForAdd = "전도인";
  if (currentFilter === "최고관리자" && isSuper) defaultRoleForAdd = "최고관리자";
  else if (currentFilter === "관리자") defaultRoleForAdd = "관리자";
  else if (currentFilter === "인도자") defaultRoleForAdd = "인도자";

  const addCard = document.createElement("div");
  addCard.className = "admin-ev-quick-add-card";
  addCard.innerHTML = `
    <div class="admin-ev-quick-add-title">
      <strong>➕ 신규 등록</strong>
      <span class="admin-ev-hint">이름과 권한, 봉사 정보를 등록합니다.${!isSuper ? " (비밀번호 설정은 최고관리자 전용)" : ""}</span>
    </div>
    <div class="admin-ev-quick-add-grid">
      <div class="ev-add-field">
        <label>이름 *</label>
        <input type="text" data-field="name" placeholder="이름 입력" required>
      </div>
      ${
        isSuper
          ? `
      <div class="ev-add-field">
        <label>비밀번호</label>
        <input type="text" data-field="password" placeholder="비밀번호 설정" autocomplete="off">
      </div>
      `
          : `
      <div class="ev-add-field" title="비밀번호 관리는 최고관리자만 가능합니다">
        <label>비밀번호 <span style="font-size:10px; color:#ef4444;">(최고관리자 전용)</span></label>
        <input type="text" data-field="password" placeholder="최고관리자만 설정 가능" disabled style="background:#f1f5f9; cursor:not-allowed;">
      </div>
      `
      }
      <div class="ev-add-field">
        <label>권한(역할) *</label>
        <select data-field="role" class="ev-role-select role-${defaultRoleForAdd}">
          <option value="전도인" ${defaultRoleForAdd === "전도인" ? "selected" : ""}>👤 전도인</option>
          <option value="인도자" ${defaultRoleForAdd === "인도자" ? "selected" : ""}>🧭 인도자</option>
          <option value="관리자" ${defaultRoleForAdd === "관리자" ? "selected" : ""}>⭐ 관리자</option>
          ${isSuper ? `<option value="최고관리자" ${defaultRoleForAdd === "최고관리자" ? "selected" : ""}>👑 최고관리자</option>` : ""}
        </select>
      </div>
      <div class="ev-add-field" style="max-width: 80px;">
        <label>성별</label>
        <select data-field="gender">
          <option value="">-</option>
          <option value="남">남</option>
          <option value="여">여</option>
        </select>
      </div>
      <div class="ev-add-field" style="max-width: 80px;">
        <label>차량정원</label>
        <input type="number" data-field="capacity" placeholder="0" min="0" max="20">
      </div>
      <div class="ev-add-field">
        <label>부부(배우자)</label>
        <input type="text" data-field="spouse" placeholder="배우자 이름">
      </div>
      <div class="ev-add-checkboxes">
        <label><input type="checkbox" data-field="driver"> 운전자</label>
        <label><input type="checkbox" data-field="deaf"> 농인</label>
      </div>
      <div class="ev-add-btn-wrap">
        <button type="button" data-action="create-ev" class="btn-create-ev">➕ 등록</button>
      </div>
    </div>
  `;
  controlWrapper.appendChild(addCard);
  box.appendChild(controlWrapper);

  // 3. 목록 테이블
  const tableContainer = document.createElement("div");
  tableContainer.className = "admin-table-container";

  const table = document.createElement("table");
  table.className = "admin-table admin-table-wide admin-ev-table";
  table.innerHTML = `
    <thead>
      <tr>
        <th style="width: 14%;">이름</th>
        <th style="width: 14%;">비밀번호</th>
        <th style="width: 14%;">권한(역할)</th>
        <th style="width: 7%;">성별</th>
        <th style="width: 7%;">운전자</th>
        <th style="width: 7%;">정원</th>
        <th style="width: 7%;">농인</th>
        <th style="width: 18%;">부부</th>
        <th style="width: 12%; min-width: 72px;">작업</th>
      </tr>
    </thead>
    <tbody></tbody>
  `;

  const tbody = table.querySelector("tbody");
  if (filtered.length === 0) {
    const emptyTr = document.createElement("tr");
    emptyTr.innerHTML = `<td colspan="9" style="text-align:center; padding: 24px; color:#64748b;">조건에 맞는 등록 인원이 없습니다.</td>`;
    tbody.appendChild(emptyTr);
  } else {
    filtered.forEach((ev) => {
      const tr = document.createElement("tr");
      tr.dataset.name = ev["이름"];
      tr.dataset.originalName = ev["이름"];

      const rawRole = (ev["역할"] || "전도인").trim();
      const isRowSuper = rawRole === "최고관리자";
      const currentRole = ["최고관리자", "관리자", "인도자"].includes(rawRole)
        ? rawRole
        : "전도인";

      tr.innerHTML = `
        <td><input type="text" value="${escapeAttr(ev["이름"] || "")}" data-field="name" class="ev-input-name" title="이름 수정 가능" ${!isSuper && isRowSuper ? "disabled" : ""}></td>
        ${
          isSuper
            ? `
        <td><input type="text" value="${escapeAttr(ev["비밀번호"] || "")}" data-field="password" class="ev-input-pw" placeholder="비번 없음" autocomplete="off"></td>
        `
            : `
        <td><div class="ev-pw-locked" title="비밀번호 관리는 최고관리자만 가능합니다">🔒 ••••••</div></td>
        `
        }
        <td>
          ${
            isSuper
              ? `
          <select data-field="role" class="ev-role-select role-${currentRole}">
            <option value="전도인" ${currentRole === "전도인" ? "selected" : ""}>👤 전도인</option>
            <option value="인도자" ${currentRole === "인도자" ? "selected" : ""}>🧭 인도자</option>
            <option value="관리자" ${currentRole === "관리자" ? "selected" : ""}>⭐ 관리자</option>
            <option value="최고관리자" ${currentRole === "최고관리자" ? "selected" : ""}>👑 최고관리자</option>
          </select>
          `
              : isRowSuper
              ? `
          <span class="ev-role-badge role-최고관리자">👑 최고관리자</span>
          <input type="hidden" data-field="role" value="최고관리자">
          `
              : `
          <select data-field="role" class="ev-role-select role-${currentRole}">
            <option value="전도인" ${currentRole === "전도인" ? "selected" : ""}>👤 전도인</option>
            <option value="인도자" ${currentRole === "인도자" ? "selected" : ""}>🧭 인도자</option>
            <option value="관리자" ${currentRole === "관리자" ? "selected" : ""}>⭐ 관리자</option>
          </select>
          `
          }
        </td>
        <td>
          <select data-field="gender" style="width:100%" ${!isSuper && isRowSuper ? "disabled" : ""}>
            <option value="" ${!ev["성별"] ? "selected" : ""}>-</option>
            <option value="남" ${ev["성별"] === "남" ? "selected" : ""}>남</option>
            <option value="여" ${ev["성별"] === "여" ? "selected" : ""}>여</option>
          </select>
        </td>
        <td><input type="checkbox" ${isTrueValue(ev["운전자"]) ? "checked" : ""} data-field="driver" ${!isSuper && isRowSuper ? "disabled" : ""}></td>
        <td><input type="number" value="${ev["차량"] || 0}" data-field="capacity" style="width:100%" min="0" max="20" ${!isSuper && isRowSuper ? "disabled" : ""}></td>
        <td><input type="checkbox" ${isTrueValue(ev["농인"]) ? "checked" : ""} data-field="deaf" ${!isSuper && isRowSuper ? "disabled" : ""}></td>
        <td><input type="text" value="${escapeAttr(ev["부부"] || "")}" data-field="spouse" style="width:100%" placeholder="배우자" ${!isSuper && isRowSuper ? "disabled" : ""}></td>
        <td class="admin-actions-cell">
          ${
            !isSuper && isRowSuper
              ? `
            <span style="font-size:11px; color:#94a3b8; font-weight:600;">최고관리자 전용</span>
          `
              : `
            <div style="display:flex; gap:4px; justify-content:center; align-items:center;">
              <button type="button" data-action="save-ev" data-name="${escapeAttr(ev["이름"])}" class="btn-save-ev" title="저장" aria-label="저장">💾</button>
              <button type="button" data-action="delete-ev" data-name="${escapeAttr(ev["이름"])}" class="btn-delete-ev" title="삭제" aria-label="삭제">🗑️</button>
            </div>
          `
          }
        </td>
      `;
      tbody.appendChild(tr);
    });
  }

  tableContainer.appendChild(table);
  box.appendChild(tableContainer);

  // 이벤트 바인딩: 검색 및 역할 변경 시 색상 즉시 업데이트
  const searchInput = box.querySelector("#admin-ev-search");
  if (searchInput) {
    searchInput.addEventListener("input", (e) => {
      state.adminEvSearch = e.target.value.trim().toLowerCase();
      const rows = tbody.querySelectorAll("tr");
      rows.forEach((tr) => {
        if (!tr.dataset.name) return;
        const nameVal = (
          tr.querySelector('[data-field="name"]')?.value ||
          tr.dataset.name ||
          ""
        ).toLowerCase();
        const spouseVal = (
          tr.querySelector('[data-field="spouse"]')?.value || ""
        ).toLowerCase();
        const roleVal = (
          tr.querySelector('[data-field="role"]')?.value || ""
        ).toLowerCase();
        const matches =
          !state.adminEvSearch ||
          nameVal.includes(state.adminEvSearch) ||
          spouseVal.includes(state.adminEvSearch) ||
          roleVal.includes(state.adminEvSearch);
        tr.style.display = matches ? "" : "none";
      });
    });
  }

  box.querySelectorAll("select[data-field='role']").forEach((sel) => {
    sel.addEventListener("change", (e) => {
      const val = e.target.value;
      e.target.className = `ev-role-select role-${val}`;
    });
  });
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
  deleted.forEach((c) => {
    const aid = c["구역번호"];
    if (!aid) return;
    if (!byArea[aid]) byArea[aid] = [];
    byArea[aid].push(c);
  });

  Object.keys(byArea)
    .sort((a, b) => compareAreaIds(a, b, state.areaOrder))
    .forEach((areaId) => {
      const item = document.createElement("div");
      item.className = "list-item";
      item.innerHTML = `<strong>${areaId} · 삭제 카드 ${byArea[areaId].length}장</strong>`;
      const innerBox = document.createElement("div");
      innerBox.className = "deleted-card-box";
      
      // 구역 내 카드 번호 순으로 정렬 추가
      const sortedCards = byArea[areaId].sort((ca, cb) => 
        compareCardNumbers(ca["카드번호"], cb["카드번호"])
      );

      sortedCards.forEach((c) => {
        const cardEl = document.createElement("div");
        cardEl.className = "card";
        cardEl.dataset.areaId = areaId;
        cardEl.dataset.cardNumber = c["카드번호"];

        cardEl.innerHTML = `
        <div class="card-header">
          <strong>${c["카드번호"]}</strong> 
          <div class="card-badges">
            <button class="status-badge" data-action="restore-deleted" data-area-id="${areaId}" data-card-number="${c["카드번호"]}" style="background-color: #3b82f6; color: white; border: none;">복원</button>
            <button class="status-badge" data-action="purge-deleted" data-area-id="${areaId}" data-card-number="${c["카드번호"]}" style="background-color: #ef4444; color: white; border: none;">삭제</button>
          </div>
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
            const visits = (state.data.visits || []).filter(
              (v) =>
                String(v["구역번호"]) === String(areaIdText) &&
                String(v["카드번호"]) === String(cardNumberText)
            );

            historyBox.innerHTML = visits.length
              ? visits
                  .map(
                    (v) =>
                      `<div style="font-size:0.85em; margin-bottom:4px;">${formatDate(v["방문날짜"])} · ${v["전도인"]} · ${v["결과"]} ${v["메모"] ? `· ${v["메모"]}` : ""}</div>`
                  )
                  .join("")
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

  const canAccess = isAdmin() || isLeader();

  if (!canAccess) {
    elements.adminPanel.classList.add("hidden");
    return;
  }

  const sections = {
    "admin-cards": document.getElementById("admin-cards-section"),
    "admin-ev": document.getElementById("admin-evangelists-section"),
    visits: document.getElementById("admin-completions-section"),
    "admin-deleted": document.getElementById("admin-deleted-section")
  };

  Object.values(sections).forEach((el) => {
    if (el) el.classList.add("hidden");
  });

  const menu = state.currentMenu;
  const overlayTitleEl = elements.adminOverlay
    ? elements.adminOverlay.querySelector("#admin-overlay-title")
    : null;

  if (menu === "admin-cards") {
    if (overlayTitleEl) overlayTitleEl.textContent = "구역카드 관리";
    if (sections["admin-cards"]) sections["admin-cards"].classList.remove("hidden");
    renderAdminCards();
  } else if (menu === "visits") {
    if (overlayTitleEl) overlayTitleEl.textContent = "완료 내역";
    if (sections.visits) sections.visits.classList.remove("hidden");
    renderAdminCompletions();
  } else if (menu === "admin-ev") {
    if (overlayTitleEl) overlayTitleEl.textContent = "인원 관리";
    if (sections["admin-ev"]) sections["admin-ev"].classList.remove("hidden");
    renderAdminEvangelists();
  } else if (menu === "admin-deleted") {
    if (overlayTitleEl) overlayTitleEl.textContent = "삭제 카드";
    if (sections["admin-deleted"]) sections["admin-deleted"].classList.remove("hidden");
    renderAdminDeletedCards();
  }

  elements.adminPanel.classList.remove("hidden");
};

const computeInviteCampaignData = () => {
  const campaign = state.inviteCampaign;
  if (!campaign || !campaign.startDate) return null;

  const startDate = campaign.startDate;
  const isActive = Boolean(campaign.active);
  const actualEndDate = campaign.endDate || null;
  const memo = campaign.memo || "";

  // 3주 (21일) 목표 마감일 계산 (시작일 포함 21일)
  const sDate = new Date(startDate);
  const targetEnd = new Date(sDate);
  targetEnd.setDate(targetEnd.getDate() + 20);
  const targetEndDateStr = targetEnd.toISOString().split("T")[0];

  const todayStr = todayISO();
  const evalDateStr = isActive ? todayStr : (actualEndDate || todayStr);
  const evalDate = new Date(evalDateStr);

  // 경과 일수 계산 (1 ~ 21일)
  const diffTime = evalDate.getTime() - sDate.getTime();
  const rawElapsed = Math.floor(diffTime / 86400000) + 1;
  const elapsedDays = Math.max(1, Math.min(21, rawElapsed));

  // D-Day 계산
  const diffToTarget = Math.ceil((targetEnd.getTime() - new Date(todayStr).getTime()) / 86400000);
  let dDayText = "";
  let dDayChipClass = "";
  if (!isActive) {
    dDayText = "종료됨";
    dDayChipClass = "finished";
  } else if (diffToTarget > 0) {
    dDayText = `D-${diffToTarget} (${diffToTarget}일 남음)`;
    dDayChipClass = "active";
  } else if (diffToTarget === 0) {
    dDayText = "D-Day (오늘 마감!)";
    dDayChipClass = "warning";
  } else {
    dDayText = `기한 초과 (+${Math.abs(diffToTarget)}일)`;
    dDayChipClass = "danger";
  }

  // 3주 기준 경과 권장 진도율 (%)
  const targetPaceRate = Math.min(100, Math.max(0, Math.round((elapsedDays / 21) * 1000) / 10));

  // 캠페인 기간 내 카드 방문 여부 확인
  const isCardVisited = (c) => {
    const cardArea = String(c["구역번호"] || "").trim();
    const cardNo = String(c["카드번호"] || "").trim();

    // 1) visits 테이블 내 캠페인 기간 방문 기록 확인
    const hasVisit = (state.data.visits || []).some((v) => {
      const vArea = String(v["구역번호"] || "").trim();
      const vCard = String(v["카드번호"] || "").trim();
      if (vArea !== cardArea || vCard !== cardNo) return false;
      const vDate = v["방문날짜"] || "";
      if (!vDate) return false;
      if (vDate < startDate) return false;
      if (actualEndDate && vDate > actualEndDate) return false;
      return true;
    });
    if (hasVisit) return true;

    // 2) 카드의 최근방문일이 캠페인 시작일 이후인지 확인
    const recent = c["최근방문일"];
    if (recent && recent >= startDate) {
      if (!actualEndDate || recent <= actualEndDate) return true;
    }

    // 3) 카드 초대장 체크박스 플래그
    if (isTrueValue(c["초대장"])) return true;

    return false;
  };

  // 모든 고유 구역 목록
  const allAreaIds = [
    ...new Set(
      (state.data.areas || [])
        .map((a) => String(a["구역번호"] || "").trim())
        .filter(Boolean)
        .concat(
          (state.data.cards || [])
            .map((c) => String(c["구역번호"] || "").trim())
            .filter(Boolean)
        )
    ),
  ].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  const areaSummaries = allAreaIds.map((areaId) => {
    const isExcluded = isKslArea(areaId); // 찾기봉사 여부 판별 (자동 제외)
    const areaCards = (state.data.cards || []).filter(
      (c) => String(c["구역번호"] || "").trim() === areaId
    );
    const validCards = areaCards.filter((c) => !isTrueValue(c["방문금지"]));
    const bannedCount = areaCards.filter((c) => isTrueValue(c["방문금지"])).length;

    const visitedCards = validCards.filter(isCardVisited);
    const unvisitedCards = validCards.filter((c) => !isCardVisited(c));

    const areaMeta =
      (state.data.areas || []).find(
        (a) => String(a["구역번호"] || "").trim() === areaId
      ) || {};
    const leader = areaMeta["인도자"] || "";

    const towns = [
      ...new Set(
        areaCards
          .map((c) => String(c["읍면동"] || "").trim())
          .filter(Boolean)
      ),
    ].join(", ");

    let status = "in_progress";
    if (isExcluded) {
      status = "excluded";
    } else if (validCards.length > 0 && visitedCards.length >= validCards.length) {
      status = "completed";
    }

    const rate =
      validCards.length > 0
        ? Math.round((visitedCards.length / validCards.length) * 1000) / 10
        : 0;

    return {
      areaId,
      isExcluded,
      leader,
      towns,
      totalCount: validCards.length,
      visitedCount: visitedCards.length,
      unvisitedCards,
      bannedCount,
      status,
      rate,
    };
  });

  // 대상 구역 (찾기봉사 제외) vs 제외 구역
  const targetAreas = areaSummaries.filter((a) => !a.isExcluded);
  const excludedAreas = areaSummaries.filter((a) => a.isExcluded);

  const totalTargetAreasCount = targetAreas.length;
  const completedTargetAreasCount = targetAreas.filter(
    (a) => a.status === "completed"
  ).length;
  const inProgressTargetAreasCount =
    totalTargetAreasCount - completedTargetAreasCount;

  const areaCompletionRate =
    totalTargetAreasCount > 0
      ? Math.round((completedTargetAreasCount / totalTargetAreasCount) * 1000) / 10
      : 0;

  const totalValidCardsCount = targetAreas.reduce(
    (sum, a) => sum + a.totalCount,
    0
  );
  const totalVisitedCardsCount = targetAreas.reduce(
    (sum, a) => sum + a.visitedCount,
    0
  );
  const cardCompletionRate =
    totalValidCardsCount > 0
      ? Math.round((totalVisitedCardsCount / totalValidCardsCount) * 1000) / 10
      : 0;

  // 캠페인 기간 총 방문 횟수
  const totalCampaignVisits = (state.data.visits || []).filter((v) => {
    const d = v["방문날짜"];
    if (!d || d < startDate) return false;
    if (actualEndDate && d > actualEndDate) return false;
    return true;
  }).length;

  // 진도 현황 평가
  let paceStatus = "good";
  let paceLabel = "🚀 목표 진도 달성 (순항 중)";
  if (areaCompletionRate >= targetPaceRate) {
    paceStatus = "good";
    paceLabel = "🚀 목표 진도 달성 (순항 중)";
  } else if (areaCompletionRate >= targetPaceRate - 15) {
    paceStatus = "warning";
    const diff = (targetPaceRate - areaCompletionRate).toFixed(1);
    paceLabel = `⚠️ 진도 분발 필요 (목표 대비 ${diff}%p 지연)`;
  } else {
    paceStatus = "danger";
    const diff = (targetPaceRate - areaCompletionRate).toFixed(1);
    paceLabel = `🚨 배부 집중 필요 (목표 대비 ${diff}%p 지연)`;
  }

  return {
    campaign,
    startDate,
    targetEndDateStr,
    actualEndDate,
    isActive,
    memo,
    elapsedDays,
    dDayText,
    dDayChipClass,
    targetPaceRate,
    areaCompletionRate,
    cardCompletionRate,
    totalTargetAreasCount,
    completedTargetAreasCount,
    inProgressTargetAreasCount,
    totalValidCardsCount,
    totalVisitedCardsCount,
    totalCampaignVisits,
    excludedCount: excludedAreas.length,
    paceStatus,
    paceLabel,
    areaSummaries,
  };
};

const openInviteStartModal = () => {
  const existing = document.getElementById("invite-start-modal");
  if (existing) existing.remove();

  const overlay = document.createElement("div");
  overlay.id = "invite-start-modal";
  overlay.className = "area-overlay";
  overlay.style.zIndex = "2500";

  const content = document.createElement("div");
  content.className = "area-overlay-content";
  content.style.maxWidth = "480px";
  content.style.borderRadius = "12px";
  content.style.overflow = "hidden";

  const startDef = todayISO();
  const d = new Date(startDef);
  d.setDate(d.getDate() + 20);
  const endDef = d.toISOString().split("T")[0];

  content.innerHTML = `
    <div class="area-overlay-header">
      <span style="font-weight: 700; font-size: 16px;">✉️ 새 3주 초대장 배부 캠페인 시작</span>
      <button class="icon-button modal-close-btn" type="button">✕</button>
    </div>
    <div style="padding: 18px 20px; display: flex; flex-direction: column; gap: 14px;">
      <div style="background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 8px; padding: 12px; font-size: 13px; color: #1e40af; line-height: 1.55;">
        📢 <b>초대장 배부 규칙 안내</b><br>
        • 시작일로부터 <b>3주(21일)</b> 기간 동안 전 구역 배부를 목표로 합니다.<br>
        • <b>찾기봉사 구역은 자동 제외</b>되며, 일반 구역 카드가 1회 이상 방문되면 자동 집계됩니다.<br>
        • 해당 구역의 모든 카드가 방문 완료되면 구역 완료로 처리됩니다.
      </div>

      <div class="form-group">
        <label style="font-size: 13px; font-weight: 700; color: #334155; display: block; margin-bottom: 4px;">배부 시작일</label>
        <input type="date" id="invite-input-start" value="${startDef}" style="width: 100%; padding: 8px 10px; border: 1px solid #cbd5e1; border-radius: 6px; box-sizing: border-box;" />
      </div>

      <div class="form-group">
        <label style="font-size: 13px; font-weight: 700; color: #334155; display: block; margin-bottom: 4px;">3주 마감 목표일</label>
        <input type="text" id="invite-display-end" value="${endDef} (21일간)" readonly style="width: 100%; padding: 8px 10px; border: 1px solid #e2e8f0; background: #f8fafc; border-radius: 6px; box-sizing: border-box; color: #64748b; font-weight: 600;" />
      </div>

      <div class="form-group">
        <label style="font-size: 13px; font-weight: 700; color: #334155; display: block; margin-bottom: 4px;">캠페인 명칭 / 메모 (선택)</label>
        <input type="text" id="invite-input-memo" placeholder="예: 2026 기념식 초대장 배부" style="width: 100%; padding: 8px 10px; border: 1px solid #cbd5e1; border-radius: 6px; box-sizing: border-box;" />
      </div>

      <div style="display: flex; justify-content: flex-end; gap: 8px; margin-top: 10px;">
        <button type="button" class="btn btn-secondary modal-close-btn" style="padding: 8px 14px; border-radius: 6px; border: 1px solid #cbd5e1; background: #fff; cursor: pointer;">취소</button>
        <button type="button" id="invite-modal-submit" style="padding: 8px 16px; border-radius: 6px; border: none; background: #2563eb; color: #fff; font-weight: 700; cursor: pointer;">🚀 캠페인 시작</button>
      </div>
    </div>
  `;

  overlay.appendChild(content);
  document.body.appendChild(overlay);

  const startInput = content.querySelector("#invite-input-start");
  const endDisplay = content.querySelector("#invite-display-end");
  startInput.addEventListener("change", () => {
    const val = startInput.value || todayISO();
    const ed = new Date(val);
    ed.setDate(ed.getDate() + 20);
    endDisplay.value = `${ed.toISOString().split("T")[0]} (21일간)`;
  });

  const close = () => overlay.remove();
  content.querySelectorAll(".modal-close-btn").forEach((btn) => btn.addEventListener("click", close));
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) close();
  });

  content.querySelector("#invite-modal-submit").addEventListener("click", async () => {
    const startDate = startInput.value || todayISO();
    const memo = content.querySelector("#invite-input-memo").value.trim();

    setLoading(true, "초대장 배부를 시작하는 중...");
    try {
      if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

      const { data, error } = await supabaseClient
        .from("invite_campaign")
        .insert({
          status: "active",
          start_date: startDate,
          memo: memo || null,
        })
        .select()
        .single();

      if (error) throw error;

      state.inviteCampaign = {
        id: data.id,
        active: true,
        startDate: data.start_date,
        memo: data.memo,
      };

      close();
      renderInviteCampaignOverlay();
      alert("🚀 3주 초대장 배부 캠페인이 시작되었습니다!");
    } catch (err) {
      console.error(err);
      alert("캠페인 시작 실패: " + err.message);
    } finally {
      setLoading(false);
    }
  });
};

const handleStopInviteCampaign = async () => {
  if (
    !confirm(
      "현재 진행 중인 초대장 배부 캠페인을 종료하시겠습니까?\n종료일이 오늘로 기록됩니다."
    )
  ) {
    return;
  }
  setLoading(true, "초대장 배부를 종료하는 중...");
  try {
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

    const { data, error } = await supabaseClient
      .from("invite_campaign")
      .update({
        status: "finished",
        end_date: todayISO(),
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
      memo: data.memo,
    };
    renderInviteCampaignOverlay();
    alert("초대장 배부 캠페인이 종료되었습니다.");
  } catch (err) {
    console.error(err);
    alert("캠페인 종료 실패: " + err.message);
  } finally {
    setLoading(false);
  }
};

const handleRefreshInviteCampaign = async () => {
  setLoading(true, "초대장 배부 통계를 최신화하는 중...");
  try {
    if (supabaseClient) {
      const { data: visits } = await supabaseClient
        .from("visits")
        .select("*")
        .order("visit_date", { ascending: false })
        .limit(1000);
      if (visits) {
        state.data.visits = visits.map((v) => ({
          id: v.id,
          "구역번호": v.area_id,
          "카드번호": v.card_number,
          "방문날짜": v.visit_date,
          "전도인": v.worker,
          "결과": v.result,
          "메모": v.note,
        }));
      }

      const { data: inviteCampaign } = await supabaseClient
        .from("invite_campaign")
        .select("*")
        .order("created_at", { ascending: false })
        .limit(1);
      if (inviteCampaign && inviteCampaign[0]) {
        state.inviteCampaign = {
          id: inviteCampaign[0].id,
          active: inviteCampaign[0].status === "active",
          startDate: inviteCampaign[0].start_date,
          endDate: inviteCampaign[0].end_date,
          memo: inviteCampaign[0].memo,
        };
      }
    }
    renderInviteCampaignOverlay();
  } catch (err) {
    console.error(err);
    alert("통계 새로고침 중 오류: " + err.message);
  } finally {
    setLoading(false);
  }
};

const renderInviteCampaignOverlay = () => {
  if (!elements.inviteOverlay) return;

  const body = elements.inviteOverlay.querySelector(".invite-body");
  if (!body) return;

  const data = computeInviteCampaignData();

  if (!data) {
    body.innerHTML = `
      <div class="invite-empty-container">
        <div class="invite-empty-icon">✉️</div>
        <div class="invite-empty-title">초대장 배부 캠페인</div>
        <div class="invite-empty-desc">
          초대장 배부는 시작일로부터 <b>3주(21일)</b> 동안 <b>찾기봉사를 제외한 전 구역</b>을<br>
          1회씩 방문하여 완료하는 특별 활동입니다.<br>
          새 캠페인을 시작하여 구역별 배부 진도를 실시간으로 관리하세요.
        </div>
        <button type="button" id="btn-invite-new-start" class="invite-btn invite-btn-primary" style="padding: 10px 20px; font-size: 15px;">
          🚀 3주 초대장 배부 시작
        </button>
      </div>
    `;
    const btn = body.querySelector("#btn-invite-new-start");
    if (btn) btn.addEventListener("click", openInviteStartModal);
    return;
  }

  // 필터 및 검색어 상태 유지
  if (!state.inviteBoardFilter) state.inviteBoardFilter = "all";
  if (!state.inviteBoardSearch) state.inviteBoardSearch = "";

  body.innerHTML = `
    <div class="invite-dashboard">
      <!-- 1. 상단 헤더 배너 -->
      <div class="invite-header-banner">
        <div class="invite-banner-top">
          <div class="invite-banner-title">
            ✉️ 초대장 배부 캠페인 ${data.memo ? `· <span style="font-weight: 500; font-size: 16px;">${escapeHtml(data.memo)}</span>` : ""}
          </div>
          <div style="display: flex; gap: 8px; align-items: center;">
            <span class="invite-status-badge ${data.isActive ? "active" : "finished"}">
              ${data.isActive ? "● 진행중 (3주 캠페인)" : "종료됨"}
            </span>
            <span class="invite-dday-chip">${data.dDayText}</span>
          </div>
        </div>
        <div class="invite-banner-meta">
          <span>📅 캠페인 기간: <b>${data.startDate} ~ ${data.targetEndDateStr} (21일간)</b></span>
          <span>⏱️ 경과: <b>${data.elapsedDays}일차</b> / 총 21일</span>
          <span>🚫 <b>찾기봉사(${data.excludedCount}개 구역)</b> 자동 제외</span>
        </div>
      </div>

      <!-- 2. 컨트롤 버튼 행 -->
      <div class="invite-actions-row">
        ${
          data.isActive
            ? `<button type="button" id="btn-invite-stop-act" class="invite-btn invite-btn-danger">⏹️ 배부 종료</button>`
            : `<button type="button" id="btn-invite-start-act" class="invite-btn invite-btn-primary">🚀 새 3주 캠페인 시작</button>`
        }
        <button type="button" id="btn-invite-refresh-act" class="invite-btn invite-btn-secondary">🔄 통계 새로고침</button>
      </div>

      <!-- 3. 진행도 게이지 카드 (실제 진도율 vs 3주 권장 목표선) -->
      <div class="invite-progress-card">
        <div class="invite-progress-header">
          <div class="invite-progress-title">전체 구역 배부 완료율</div>
          <div class="invite-progress-rate">${data.areaCompletionRate}%</div>
        </div>
        <div class="invite-progress-bar-wrap">
          <div class="invite-progress-bar-fill" style="width: ${Math.min(100, data.areaCompletionRate)}%;"></div>
          <div class="invite-progress-target-line" style="left: ${data.targetPaceRate}%;">
            <span class="invite-progress-target-label">3주 권장 목표 (${data.targetPaceRate}%)</span>
          </div>
        </div>
        <div class="invite-progress-footer">
          <div>
            <b>구역:</b> ${data.completedTargetAreasCount} / ${data.totalTargetAreasCount} 구역 완료 &nbsp;|&nbsp;
            <b>카드:</b> ${data.totalVisitedCardsCount} / ${data.totalValidCardsCount} 장 방문 (${data.cardCompletionRate}%)
          </div>
          <span class="invite-pace-badge ${data.paceStatus}">${data.paceLabel}</span>
        </div>
      </div>

      <!-- 4. 4개 KPI 지표 그리드 -->
      <div class="invite-kpi-grid">
        <div class="invite-kpi-tile">
          <div class="invite-kpi-icon">⏱️</div>
          <div class="invite-kpi-value">${data.dDayText.split(" ")[0]}</div>
          <div class="invite-kpi-label">${data.elapsedDays}일 경과 / 총 21일</div>
        </div>
        <div class="invite-kpi-tile">
          <div class="invite-kpi-icon">🗺️</div>
          <div class="invite-kpi-value">${data.completedTargetAreasCount} / ${data.totalTargetAreasCount}</div>
          <div class="invite-kpi-label">완료 구역 (${data.areaCompletionRate}%)</div>
        </div>
        <div class="invite-kpi-tile">
          <div class="invite-kpi-icon">🗂️</div>
          <div class="invite-kpi-value">${data.totalVisitedCardsCount} / ${data.totalValidCardsCount}</div>
          <div class="invite-kpi-label">방문 카드 (${data.cardCompletionRate}%)</div>
        </div>
        <div class="invite-kpi-tile">
          <div class="invite-kpi-icon">👥</div>
          <div class="invite-kpi-value">${data.totalCampaignVisits}회</div>
          <div class="invite-kpi-label">캠페인 총 방문 횟수</div>
        </div>
      </div>

      <!-- 5. 구역별 현황 보드 -->
      <div class="invite-board-container">
        <div class="invite-board-header">
          <div class="invite-board-tabs" id="invite-board-tabs">
            <button type="button" class="invite-tab-btn ${state.inviteBoardFilter === "all" ? "active" : ""}" data-filter="all">전체 (${data.areaSummaries.length})</button>
            <button type="button" class="invite-tab-btn ${state.inviteBoardFilter === "in_progress" ? "active" : ""}" data-filter="in_progress">미완료 (${data.inProgressTargetAreasCount})</button>
            <button type="button" class="invite-tab-btn ${state.inviteBoardFilter === "completed" ? "active" : ""}" data-filter="completed">완료 (${data.completedTargetAreasCount})</button>
            <button type="button" class="invite-tab-btn ${state.inviteBoardFilter === "excluded" ? "active" : ""}" data-filter="excluded">찾기봉사 제외 (${data.excludedCount})</button>
          </div>
          <input type="text" id="invite-board-search" class="invite-search-input" placeholder="🔍 구역번호, 인도자, 읍면동 검색..." value="${escapeHtml(state.inviteBoardSearch)}" />
        </div>
        <div class="invite-area-cards-grid" id="invite-area-cards-grid"></div>
      </div>
    </div>
  `;

  // 액션 버튼 이벤트 연결
  const stopBtn = body.querySelector("#btn-invite-stop-act");
  if (stopBtn) stopBtn.addEventListener("click", handleStopInviteCampaign);

  const startBtn = body.querySelector("#btn-invite-start-act");
  if (startBtn) startBtn.addEventListener("click", openInviteStartModal);

  const refreshBtn = body.querySelector("#btn-invite-refresh-act");
  if (refreshBtn) refreshBtn.addEventListener("click", handleRefreshInviteCampaign);

  // 탭 및 검색어 필터링 렌더러
  const grid = body.querySelector("#invite-area-cards-grid");
  const renderBoardGrid = () => {
    if (!grid) return;
    const filter = state.inviteBoardFilter || "all";
    const q = (state.inviteBoardSearch || "").trim().toLowerCase();

    let list = data.areaSummaries;
    if (filter === "in_progress") {
      list = list.filter((a) => !a.isExcluded && a.status !== "completed");
    } else if (filter === "completed") {
      list = list.filter((a) => !a.isExcluded && a.status === "completed");
    } else if (filter === "excluded") {
      list = list.filter((a) => a.isExcluded);
    }

    if (q) {
      list = list.filter((a) => {
        return (
          a.areaId.toLowerCase().includes(q) ||
          a.leader.toLowerCase().includes(q) ||
          a.towns.toLowerCase().includes(q)
        );
      });
    }

    if (list.length === 0) {
      grid.innerHTML = `<div style="grid-column: 1 / -1; padding: 24px; text-align: center; color: #94a3b8;">조건에 해당하는 구역이 없습니다.</div>`;
      return;
    }

    grid.innerHTML = "";
    list.forEach((a) => {
      const cardEl = document.createElement("div");
      cardEl.className = "invite-area-card";

      let statusBadge = "";
      if (a.isExcluded) {
        statusBadge = `<span style="background: #f1f5f9; color: #64748b; font-size: 11px; padding: 2px 7px; border-radius: 10px; font-weight: 700;">찾기봉사 제외 🚫</span>`;
      } else if (a.status === "completed") {
        statusBadge = `<span style="background: #dcfce7; color: #166534; font-size: 11px; padding: 2px 7px; border-radius: 10px; font-weight: 700;">완료 ✅</span>`;
      } else {
        const remainingCount = a.totalCount - a.visitedCount;
        statusBadge = `<span style="background: #fef3c7; color: #92400e; font-size: 11px; padding: 2px 7px; border-radius: 10px; font-weight: 700;">${remainingCount}장 남음 ⏳</span>`;
      }

      const unvisitedChips = a.unvisitedCards
        .slice(0, 6)
        .map((c) => `#${c["카드번호"]}`)
        .join(", ");
      const moreText = a.unvisitedCards.length > 6 ? ` 외 ${a.unvisitedCards.length - 6}장` : "";

      cardEl.innerHTML = `
        <div style="display: flex; justify-content: space-between; align-items: center;">
          <span style="font-weight: 800; font-size: 15px; color: #1e293b;">${escapeHtml(a.areaId)}구역</span>
          ${statusBadge}
        </div>
        <div style="font-size: 12px; color: #64748b; display: flex; justify-content: space-between;">
          <span>👤 ${a.leader ? escapeHtml(a.leader) : "인도자 미지정"}</span>
          <span>${a.towns ? escapeHtml(a.towns) : ""}</span>
        </div>
        ${
          !a.isExcluded
            ? `
          <div style="margin-top: 4px;">
            <div style="display: flex; justify-content: space-between; font-size: 11px; color: #475569; margin-bottom: 3px;">
              <span>방문: <b>${a.visitedCount}</b> / ${a.totalCount}장</span>
              <span><b>${a.rate}%</b></span>
            </div>
            <div style="height: 6px; background: #e2e8f0; border-radius: 3px; overflow: hidden;">
              <div style="height: 100%; width: ${Math.min(100, a.rate)}%; background: ${a.status === "completed" ? "#10b981" : "#3b82f6"};"></div>
            </div>
          </div>
          <div style="font-size: 11px; color: #64748b; min-height: 16px; margin-top: 4px;">
            ${
              a.status === "completed"
                ? `<span style="color: #10b981; font-weight: 700;">모든 카드 방문 완료! 🎉</span>`
                : `<span style="color: #d97706;">미방문: <b>${unvisitedChips}${moreText}</b></span>`
            }
          </div>
        `
            : `
          <div style="font-size: 11px; color: #94a3b8; padding: 8px 0;">
            찾기봉사 전용 구역 (초대장 배부 대상에서 제외됨)
          </div>
        `
        }
      `;

      // 구역 카드 클릭 시 해당 구역 바로 열기
      cardEl.style.cursor = "pointer";
      cardEl.addEventListener("click", () => {
        elements.inviteOverlay.classList.add("hidden");
        selectArea(a.areaId);
      });

      grid.appendChild(cardEl);
    });
  };

  renderBoardGrid();

  // 탭 클릭 이벤트
  body.querySelectorAll("#invite-board-tabs .invite-tab-btn").forEach((tab) => {
    tab.addEventListener("click", () => {
      body.querySelectorAll("#invite-board-tabs .invite-tab-btn").forEach((t) => t.classList.remove("active"));
      tab.classList.add("active");
      state.inviteBoardFilter = tab.getAttribute("data-filter");
      renderBoardGrid();
    });
  });

  // 검색어 입력 이벤트
  const searchInput = body.querySelector("#invite-board-search");
  if (searchInput) {
    searchInput.addEventListener("input", () => {
      state.inviteBoardSearch = searchInput.value;
      renderBoardGrid();
    });
  }
};

window.computeInviteCampaignData = computeInviteCampaignData;
window.openInviteStartModal = openInviteStartModal;
window.handleStopInviteCampaign = handleStopInviteCampaign;
window.handleRefreshInviteCampaign = handleRefreshInviteCampaign;
window.renderInviteCampaignOverlay = renderInviteCampaignOverlay;
window.renderAdminCards = renderAdminCards;
window.renderAdminCompletions = renderAdminCompletions;
window.renderCompletionHistoryView = renderCompletionHistoryView;
window.openCardEditModal = openCardEditModal;
window.renderAdminPanel = renderAdminPanel;

