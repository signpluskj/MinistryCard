const renderCards = () => {
  const getStickyOffset = () => {
    const header = document.querySelector(".app-header");
    const info = document.querySelector(".global-info");
    let offset = 0;
    if (header && !header.classList.contains("hidden")) {
      offset += header.getBoundingClientRect().height || 0;
    }
    if (info && !info.classList.contains("hidden")) {
      offset += info.getBoundingClientRect().height || 0;
    }
    return offset + 10;
  };
  const scrollWithOffset = (target) => {
    if (!target) {
      return;
    }
    const y =
      window.pageYOffset + target.getBoundingClientRect().top - getStickyOffset();
    window.scrollTo({
      top: Math.max(0, y),
      behavior: "smooth"
    });
  };
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
    const lowerQuery = query.toLowerCase().replace(/\s+/g, "");
    cards = cards.filter((card) => {
      const areaId = String(card["구역번호"] || "");
      const area = state.data.areas.find(a => String(a["구역번호"]) === areaId);
      const leader = area ? (area["인도자"] || "") : "";
      const searchTarget = `${card["주소"] || ""} ${card["상세주소"] || ""} ${card["정기방문자"] || ""} ${leader}`.toLowerCase().replace(/\s+/g, "");
      return searchTarget.includes(lowerQuery);
    });
  }
  if (state.filterVisit === "meet") {
    cards = cards.filter((card) => isTrueValue(card["만남"]));
  } else if (state.filterVisit === "absent") {
    cards = cards.filter(
      (card) => (isTrueValue(card["부재"]) || card["만남"] === false) && 
                !isTrueValue(card["재방"]) && 
                !isTrueValue(card["연구"])
    );
  } else if (state.filterVisit === "revisit") {
    cards = cards.filter((card) => isTrueValue(card["재방"]));
  } else if (state.filterVisit === "study") {
    cards = cards.filter((card) => isTrueValue(card["연구"]));
  } else if (state.filterVisit === "six") {
    cards = cards.filter((card) => isTrueValue(card["6개월"]));
  } else if (state.filterVisit === "banned") {
    cards = cards.filter((card) => isTrueValue(card["방문금지"]));
  }
  cards.sort((a, b) => {
    // 1순위: 차량 배정 여부 (오늘 날짜 기준)
    const carA = getCardAssignedCarIdForDate(a, today);
    const carB = getCardAssignedCarIdForDate(b, today);
    if (carA && !carB) return -1;
    if (!carA && carB) return 1;

    // 2순위: 구역 번호 (전체 구역 보기일 때)
    if (state.filterArea === "all") {
      const aa = String(a["구역번호"] || "");
      const ab = String(b["구역번호"] || "");
      if (aa !== ab) {
        return compareAreaIds(aa, ab, state.areaOrder);
      }
    }

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

    // Requirement: Long-press to open Google Maps
    let pressTimer;
    const startPress = () => {
      pressTimer = setTimeout(() => {
        const addressText = String(card["주소"] || "").trim();
        const queryText = addressText || String(card["카드번호"] || "").trim();
        const encoded = encodeURIComponent(queryText);
        window.open(`https://www.google.com/maps/search/?api=1&query=${encoded}`, "_blank");
      }, 600);
    };
    const cancelPress = () => {
      clearTimeout(pressTimer);
    };
    cardEl.addEventListener("mousedown", startPress);
    cardEl.addEventListener("mouseup", cancelPress);
    cardEl.addEventListener("mouseleave", cancelPress);
    cardEl.addEventListener("touchstart", (e) => {
      // Don't trigger if it's a multi-touch
      if (e.touches.length === 1) startPress();
    }, { passive: true });
    cardEl.addEventListener("touchend", cancelPress);
    cardEl.addEventListener("touchmove", cancelPress);
    cardEl.addEventListener("contextmenu", (e) => {
      // Prevent default context menu to not interfere with long-press
      // But only if we want to fully take over. For now, let's just let it be.
    });
    const isSelected =
      state.selectedCard &&
      String(card["구역번호"]) === String(state.selectedCard["구역번호"]) &&
      String(card["카드번호"]) === String(state.selectedCard["카드번호"]);
    
    // 오직 선택된(방문 기록 창이 열린) 카드만 active 클래스를 가짐
    if (isSelected) {
      cardEl.classList.add("active");
      if (isInProgress) {
        cardEl.classList.add("card-active-progress");
      }
    }

    const title = document.createElement("div");
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

    // 차량 배정 정보 표시 (항상 표시하되, 체크박스가 있을 때는 그 옆에 배치)
    const assignedCarId = getCardAssignedCarIdForDate(card, todayISO());
    let carInfoEl = null;
    if (assignedCarId) {
      const assignments = state.data.assignments || [];
      // 오늘 날짜와 해당 차량 번호에 맞는 배정 정보를 찾음
      const assignment = assignments.find(
        (a) =>
          String(a["차량"]) === String(assignedCarId) &&
          toAssignmentDateText(a["날짜"]) === toAssignmentDateText(todayISO())
      );
      const driverName = assignment ? String(assignment["이름"] || "") : "";
      carInfoEl = document.createElement("span");
      carInfoEl.className = "card-car-info";
      carInfoEl.style.fontSize = "14px";
      carInfoEl.style.color = "#1d4ed8";
      carInfoEl.style.fontWeight = "600";
      carInfoEl.style.marginRight = "8px";
      carInfoEl.innerHTML = `🚗${driverName ? ` (${driverName})` : ""}`;
    }

    if (canAssignFromCardsPanel) {
      const assignBox = document.createElement("span");
      assignBox.className = "card-assign-box";
      const label = document.createElement("label");
      label.className = "card-assign-label";
      
      if (carInfoEl) {
        label.appendChild(carInfoEl);
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
    } else if (carInfoEl) {
      // 체크박스를 볼 권한이 없는 일반 전도인에게도 차량 정보는 보여줌
      const assignBox = document.createElement("span");
      assignBox.className = "card-assign-box";
      assignBox.appendChild(carInfoEl);
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

          if (supabaseClient) {
            try {
              await deleteCardInSupabase(areaId, cardNumber);
            } catch (supaErr) {
              console.warn("Supabase delete failed (non-critical):", supaErr);
            }
          }

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

            if (elements.visitDelete) {
              const isAdmin = state.user && (state.user.role === "관리자" || state.user.role === "인도자");
              if (isAdmin) {
                elements.visitDelete.classList.remove("hidden");
              } else {
                elements.visitDelete.classList.add("hidden");
              }
            }

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
      scrollWithOffset(scrollTarget);
      if (typeof window !== "undefined" && window.requestAnimationFrame) {
        window.requestAnimationFrame(() => {
          scrollWithOffset(scrollTarget);
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
      scrollWithOffset(target);
    }
  }
  state.scrollToSelectedCard = false;
  state.scrollAreaToTop = false;
};
