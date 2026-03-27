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
      const startText = row["시작날짜"] ? formatDate(row["시작날짜"]) : "";
      const doneText = row["완료날짜"] ? formatDate(row["완료날짜"]) : "";
      title.textContent =
        startText && doneText
          ? `${startText} → ${doneText}`
          : doneText || startText || "";
      const leader = document.createElement("div");
      leader.textContent = row["인도자"] ? `인도자: ${row["인도자"]}` : "";
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
