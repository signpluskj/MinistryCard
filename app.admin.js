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
  const areaIds = Object.keys(byArea).sort((a, b) => compareAreaIds(a, b, state.areaOrder));
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
          entry.style.display = "flex";
          entry.style.justifyContent = "space-between";
          entry.style.alignItems = "center";
          entry.style.gap = "8px";

          const textSpan = document.createElement("span");
          const startText = row["시작날짜"] ? formatDate(row["시작날짜"]) : "";
          const doneText = row["완료날짜"] ? formatDate(row["완료날짜"]) : "";
          const titleText =
            startText && doneText
              ? `${startText} → ${doneText}`
              : doneText || startText || "";
          const leaderText = row["인도자"] ? ` · 인도자: ${row["인도자"]}` : "";
          const memoText = row["비고"] ? ` · ${row["비고"]}` : "";
          textSpan.textContent = `${titleText}${leaderText}${memoText}`;
          entry.appendChild(textSpan);

          const isAdmin = state.user && (state.user.role === "관리자" || state.user.role === "인도자");
          if (isAdmin) {
            const actions = document.createElement("div");
            actions.className = "completion-item-actions";
            actions.style.display = "flex";
            actions.style.gap = "4px";
            actions.style.flexShrink = "0";

            const editBtn = document.createElement("button");
            editBtn.textContent = "수정";
            editBtn.className = "btn-small";
            editBtn.style.padding = "2px 6px";
            editBtn.style.fontSize = "inherit";
            editBtn.onclick = (e) => {
              e.stopPropagation();
              editCompletion(row);
            };

            const delBtn = document.createElement("button");
            delBtn.textContent = "삭제";
            delBtn.className = "btn-small btn-danger";
            delBtn.style.backgroundColor = "#ef4444";
            delBtn.style.color = "white";
            delBtn.style.border = "none";
            delBtn.style.borderRadius = "4px";
            delBtn.style.padding = "2px 6px";
            delBtn.style.fontSize = "inherit";
            delBtn.onclick = (e) => {
              e.stopPropagation();
              deleteCompletion(row);
            };

            actions.append(editBtn, delBtn);
            entry.appendChild(actions);
          }
          
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

const renderAdminCards = () => {
  const box = document.getElementById("admin-cards-list");
  if (!box) return;
  box.innerHTML = "";
  const byAreaCards = groupCardsByArea();
  const areaIds = Object.keys(byAreaCards).sort((a, b) => compareAreaIds(a, b, state.areaOrder));
  const selectedAreaId = state.adminCardsAreaId;

  areaIds.forEach((areaId) => {
    const cards = byAreaCards[areaId].sort((ca, cb) =>
      compareCardNumbers(ca["카드번호"], cb["카드번호"])
    );
    const item = document.createElement("div");
    item.className = "list-item" + (selectedAreaId === areaId ? " active" : "");
    item.innerHTML = `<div><strong>${areaId} · 카드 ${cards.length}장</strong></div>`;

    const cardsBox = document.createElement("div");
    cardsBox.className = "area-cards";
    if (selectedAreaId !== areaId) cardsBox.style.display = "none";

    if (selectedAreaId === areaId) {
      cards.forEach((card) => {
        const cardWrapper = document.createElement("div");
        cardWrapper.className = "card admin-card-row";
        const cardNo = String(card["카드번호"] || "");
        const townText = getCardTownLabel(card);

        cardWrapper.innerHTML = `
          <div class="card-header">
            <strong>${cardNo} ${townText ? `(${townText})` : ""}</strong>
            <div class="card-badges">
              <button class="status-badge" data-action="save-card" data-area-id="${areaId}" data-card-number="${cardNo}" style="background-color: #3b82f6; color: white; border: none;">저장</button>
              <button class="status-badge" data-action="delete-card" data-area-id="${areaId}" data-card-number="${cardNo}" style="background-color: #ef4444; color: white; border: none;">삭제</button>
            </div>
          </div>
          <div class="card-line">읍면동: <input type="text" value="${card["읍면동"] || ""}" data-field="town"></div>
          <div class="card-line">주소: <input type="text" value="${card["주소"] || ""}" data-field="address"></div>
          <div class="card-line">상세: <input type="text" value="${card["상세주소"] || ""}" data-field="detailAddress"></div>
          <div class="card-line">정보: <input type="text" value="${card["정보"] || ""}" data-field="memo"></div>
          <div class="card-check-group">
            <label><input type="checkbox" ${isTrueValue(card["6개월"]) ? "checked" : ""} data-field="sixMonths"> 6개월</label>
            <label><input type="checkbox" ${isTrueValue(card["방문금지"]) ? "checked" : ""} data-field="banned"> 방문금지</label>
            <label><input type="checkbox" ${isTrueValue(card["재방"]) ? "checked" : ""} data-field="revisit"> 재방</label>
            <label><input type="checkbox" ${isTrueValue(card["연구"]) ? "checked" : ""} data-field="study"> 연구</label>
          </div>
        `;
        cardsBox.appendChild(cardWrapper);
        });

        const newCardWrapper = document.createElement("div");
        newCardWrapper.className = "card admin-card-row";
        newCardWrapper.dataset.new = "true";
        newCardWrapper.dataset.areaId = areaId;
        newCardWrapper.innerHTML = `
        <div class="card-header">
        <strong>새 카드 추가</strong>
        <div class="card-badges">
          <button class="status-badge" data-card-action="create-card" style="background-color: #22c55e; color: white; border: none;">추가</button>
        </div>
        </div>
        <div class="card-line">카드번호: <input type="text" placeholder="예: 101-1" data-field="cardNumber"></div>
        <div class="card-line">읍면동: <input type="text" data-field="town"></div>
        <div class="card-line">주소: <input type="text" data-field="address"></div>
        <div class="card-line">상세: <input type="text" data-field="detailAddress"></div>
        <div class="card-line">정보: <input type="text" data-field="memo"></div>
        <div class="card-check-group">
        <label><input type="checkbox" data-field="sixMonths"> 6개월</label>
        <label><input type="checkbox" data-field="banned"> 방문금지</label>
        <label><input type="checkbox" data-field="revisit"> 재방</label>
        <label><input type="checkbox" data-field="study"> 연구</label>
        </div>
        `;      cardsBox.appendChild(newCardWrapper);
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
  const completions = [...(state.data.completions || [])].sort(
    (a, b) => new Date(b["완료날짜"] || 0) - new Date(a["완료날짜"] || 0)
  );

  if (!completions.length) {
    box.innerHTML = '<div class="list-item">완료 내역이 없습니다.</div>';
    return;
  }

  const userRole = state.user ? state.user.role : "";
  const canEdit = userRole === "관리자" || state.isSuperAdmin === true;

  completions.forEach((c) => {
    const div = document.createElement("div");
    div.className = "list-item";
    div.style.display = "flex";
    div.style.justifyContent = "space-between";
    div.style.alignItems = "flex-start";
    div.style.gap = "8px";

    const content = document.createElement("div");
    content.style.flex = "1";
    content.innerHTML = `
      <div style="display: flex; justify-content: space-between;">
        <strong>${c["구역번호"]}</strong>
        <span>${formatDate(c["완료날짜"])}</span>
      </div>
      <div style="font-size: 0.9em; margin-top: 4px;">인도자: ${c["인도자"] || "없음"}</div>
    `;
    div.appendChild(content);

    if (canEdit) {
      const actions = document.createElement("div");
      actions.style.display = "flex";
      actions.style.gap = "5px";
      actions.style.flexShrink = "0";

      const editBtn = document.createElement("button");
      editBtn.textContent = "수정";
      editBtn.className = "admin-btn-small";
      editBtn.style.padding = "2px 6px";
      editBtn.style.fontSize = "inherit";
      editBtn.addEventListener("click", () => editCompletion(c));

      const delBtn = document.createElement("button");
      delBtn.textContent = "삭제";
      delBtn.className = "admin-btn-small";
      delBtn.style.padding = "2px 6px";
      delBtn.style.fontSize = "inherit";
      delBtn.style.backgroundColor = "#ef4444";
      delBtn.style.color = "white";
      delBtn.style.border = "none";
      delBtn.addEventListener("click", () => deleteCompletion(c));

      actions.append(editBtn, delBtn);
      div.appendChild(actions);
    }

    box.appendChild(div);
  });
};

const renderAdminEvangelists = () => {
  const box = document.getElementById("admin-evangelist-list");
  if (!box) return;
  box.innerHTML = "";
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
  evangelists.forEach((ev) => {
    const tr = document.createElement("tr");
    tr.dataset.name = ev["이름"];

    tr.innerHTML = `
      <td><input type="text" value="${ev["이름"] || ""}" readOnly></td>
      <td><input type="text" value="${ev["성별"] || ""}" data-field="gender" style="width:100%"></td>
      <td><input type="checkbox" ${isTrueValue(ev["농인"]) ? "checked" : ""} data-field="deaf"></td>
      <td><input type="text" value="${ev["역할"] || ""}" data-field="role" style="width:100%"></td>
      <td><input type="checkbox" ${isTrueValue(ev["운전자"]) ? "checked" : ""} data-field="driver"></td>
      <td><input type="number" value="${ev["차량"] || 0}" data-field="capacity" style="width:100%"></td>
      <td><input type="text" value="${ev["부부"] || ""}" data-field="spouse" style="width:100%"></td>
      <td class="admin-actions-cell">
        <div style="display:flex; gap:2px; justify-content:center">
          <button data-action="save-ev" data-name="${ev["이름"]}" style="flex:1; padding:4px 0">저장</button>
          <button data-action="delete-ev" data-name="${ev["이름"]}" style="flex:1; padding:4px 0; background-color:#ef4444">삭제</button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });

  const newTr = document.createElement("tr");
  newTr.dataset.new = "true";
  newTr.innerHTML = `
    <td><input type="text" placeholder="이름" data-field="name"></td>
    <td><input type="text" data-field="gender" style="width:100%"></td>
    <td><input type="checkbox" data-field="deaf"></td>
    <td><input type="text" value="전도인" data-field="role" style="width:100%"></td>
    <td><input type="checkbox" data-field="driver"></td>
    <td><input type="number" data-field="capacity" style="width:100%"></td>
    <td><input type="text" data-field="spouse" style="width:100%"></td>
    <td class="admin-actions-cell">
      <div style="display:flex; gap:2px; justify-content:center">
        <button data-action="create-ev" style="width:100%; padding:4px 0">추가</button>
      </div>
    </td>
  `;  tbody.appendChild(newTr);
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
      byArea[areaId].forEach((c) => {
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

  const userRole = String(state.user.role || "").trim();
  const isAdmin = userRole === "관리자";
  const isLeader = userRole === "인도자";
  const canAccessAdmin = isAdmin || isLeader;

  if (!canAccessAdmin) {
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
      parts.push(`기간: ${summary.startDate} ~ ${summary.endDate}`);
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
