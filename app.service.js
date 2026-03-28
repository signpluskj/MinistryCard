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

    // 마지막 완료 날짜 찾기
    const { data: latestComp, error: compError } = await supabaseClient
      .from("completions")
      .select("end_date")
      .eq("area_id", String(areaId))
      .order("end_date", { ascending: false })
      .limit(1);

    if (compError) {
      console.warn("Could not fetch last completion date:", compError);
    }

    const lastEndDate = (latestComp && latestComp.length > 0) ? latestComp[0].end_date : null;

    const { error } = await supabaseClient
      .from("areas")
      .update({
        start_date: null,
        leader: null,
        end_date: lastEndDate // 이전 완료 날짜 복원
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

    // 방문 기록이 없는 카드를 찾아 '부재'로 기록
    const { data: areaCards, error: areaCardsError } = await supabaseClient
      .from("cards")
      .select("*")
      .eq("area_id", areaId);

    if (areaCardsError) throw areaCardsError;

    const startDate = areaRow.start_date;
    const unvisitedCards = areaCards.filter((c) => {
      if (c.revisit || c.study || c.six_months || c.banned) return false;
      if (!c.recent_visit_date) return true;
      if (startDate) {
        const rv = parseVisitDate(c.recent_visit_date);
        const sd = parseVisitDate(startDate);
        return rv && sd && rv < sd;
      }
      return false;
    });

    if (unvisitedCards.length > 0) {
      const completionVisits = unvisitedCards.map((c) => ({
        area_id: areaId,
        card_number: c.card_number,
        visit_date: endDate,
        worker: `${state.user.name}`,
        result: "부재",
        note: "완료 처리"
      }));

      const { error: batchVisitError } = await supabaseClient
        .from("visits")
        .insert(completionVisits);

      if (batchVisitError) throw batchVisitError;

      // 카드 상태 업데이트
      for (const c of unvisitedCards) {
        await supabaseClient
          .from("cards")
          .update({
            recent_visit_date: endDate,
            absent: true,
            meet: false
          })
          .eq("area_id", areaId)
          .eq("card_number", c.card_number);
      }
    }

    const { error: updateError } = await supabaseClient
      .from("areas")
      .update({ end_date: endDate })
      .eq("area_id", areaId);

    if (updateError) throw updateError;

    const { error: insertError } = await supabaseClient
      .from("completions")
      .upsert({
        area_id: areaId,
        start_date: areaRow.start_date || endDate,
        end_date: endDate,
        leader: areaRow.leader || (state.user ? state.user.name : "시스템")
      }, { onConflict: "area_id, end_date" });

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

const editCompletion = async (row) => {
  const newStart = window.prompt("시작날짜를 입력해 주세요 (YYYY-MM-DD)", row["시작날짜"] || "");
  if (newStart === null) return;
  const newDone = window.prompt("완료날짜를 입력해 주세요 (YYYY-MM-DD)", row["완료날짜"] || "");
  if (newDone === null) return;
  const newLeader = window.prompt("인도자를 입력해 주세요", row["인도자"] || "");
  if (newLeader === null) return;

  setLoading(true, "완료 내역 수정 중...");
  try {
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

    const { error } = await supabaseClient
      .from("completions")
      .update({
        start_date: newStart || null,
        end_date: newDone || null,
        leader: newLeader || null
      })
      .eq("id", row.id);

    if (error) throw error;

    setStatus("완료 내역이 수정되었습니다.");
    await loadData();
    renderVisitsView();
  } catch (err) {
    console.error("Edit completion error:", err);
    alert("수정에 실패했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

const deleteCompletion = async (row) => {
  if (!window.confirm("이 완료 내역을 삭제하시겠습니까?")) return;

  setLoading(true, "완료 내역 삭제 중...");
  try {
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

    const { error } = await supabaseClient
      .from("completions")
      .delete()
      .eq("id", row.id);

    if (error) throw error;

    setStatus("완료 내역이 삭제되었습니다.");
    await loadData();
    renderVisitsView();
  } catch (err) {
    console.error("Delete completion error:", err);
    alert("삭제에 실패했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

const deleteVisit = async () => {
  if (!state.editingVisit) return;
  if (!window.confirm("이 방문 기록을 삭제하시겠습니까?")) return;

  setLoading(true, "방문 기록 삭제 중...");
  try {
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

    const { error: deleteError } = await supabaseClient
      .from("visits")
      .delete()
      .eq("id", state.editingVisit.id);

    if (deleteError) throw deleteError;

    const { data: allVisits, error: visitsError } = await supabaseClient
      .from("visits")
      .select("*")
      .eq("area_id", state.editingVisit.areaId)
      .eq("card_number", state.editingVisit.cardNumber)
      .order("visit_date", { ascending: false });

    if (visitsError) throw visitsError;

    const last = allVisits && allVisits.length ? allVisits[0] : null;
    
    // 6개월 및 방문금지는 과거 기록 중 하나라도 있으면 유지되는 특성이 있음 (기존 saveVisit 로직 참고)
    const hasSixMonths = (allVisits || []).some(v => v.result === "6개월");
    const hasBanned = (allVisits || []).some(v => v.result === "방문금지");

    const { error: cardError } = await supabaseClient
      .from("cards")
      .update({
        recent_visit_date: last ? last.visit_date : null,
        meet: last ? last.result === "만남" : false,
        absent: last ? last.result === "부재" : false,
        revisit: last ? last.result === "재방" : false,
        study: last ? last.result === "연구" : false,
        invite: last ? last.result === "초대장" : false,
        six_months: hasSixMonths,
        banned: hasBanned
      })
      .eq("area_id", state.editingVisit.areaId)
      .eq("card_number", state.editingVisit.cardNumber);

    if (cardError) throw cardError;

    setStatus("방문 기록이 삭제되었습니다.");
    
    if (last) {
      // 삭제 후 남은 기록 중 가장 최근 기록을 자동으로 선택
      state.editingVisit = {
        id: last.id,
        areaId: String(state.editingVisit.areaId),
        cardNumber: String(state.editingVisit.cardNumber),
        oldVisitDate: String(last.visit_date),
        oldWorker: String(last.worker || ""),
        oldResult: String(last.result || ""),
        oldNote: String(last.note || "")
      };

      if (elements.visitTitle) elements.visitTitle.textContent = `카드 ${state.editingVisit.cardNumber} 방문 내역 수정`;
      if (elements.visitDate) elements.visitDate.value = toISODate(last.visit_date);
      if (elements.visitWorker) elements.visitWorker.value = last.worker || "";
      if (elements.visitResult) elements.visitResult.value = last.result || "만남";
      if (elements.visitNote) elements.visitNote.value = last.note || "";
      
      // 삭제 버튼 유지 (관리자 권한 확인은 이미 되어 있을 것이므로)
      if (elements.visitDelete) elements.visitDelete.classList.remove("hidden");
    } else {
      // 남은 기록이 없으면 초기화
      state.editingVisit = null;
      if (elements.visitTitle) elements.visitTitle.textContent = "방문 내역 기록";
      if (elements.visitDate) elements.visitDate.value = todayISO();
      if (elements.visitWorker) elements.visitWorker.value = state.user ? state.user.name : "";
      if (elements.visitResult) elements.visitResult.value = "만남";
      if (elements.visitNote) elements.visitNote.value = "";
      if (elements.visitDelete) elements.visitDelete.classList.add("hidden");
    }

    await loadData();
    renderCards();
  } catch (err) {
    console.error("Delete visit error:", err);
    alert("삭제에 실패했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

const saveVisit = async (event) => {
  if (event) event.preventDefault();
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
      six_months: cardRow.six_months || latestResult === "6개월",
      banned: cardRow.banned || latestResult === "방문금지"
    };

    const { error: cardUpdateError } = await supabaseClient
      .from("cards")
      .update(updatePayload)
      .eq("area_id", areaId)
      .eq("card_number", cardNumber);

    if (cardUpdateError) throw cardUpdateError;

    const { data: areaCards, error: areaCardsError } = await supabaseClient
      .from("cards")
      .select("*")
      .eq("area_id", areaId);

    if (areaCardsError) throw areaCardsError;

    const isComplete = areaCards.every((c) => {
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

      // Ensure we save completion if not already done, even if start_date is missing
      if (!areaRowError && !areaRow.end_date) {
        const completeDate = visitDate || todayISO();
        // 봉사 시작 시 등록된 인도자 이름을 유지, 없으면 현재 사용자
        const leaderName = areaRow.leader || (state.user ? state.user.name : "시스템");

        await supabaseClient
          .from("areas")
          .update({ end_date: completeDate, leader: leaderName })
          .eq("area_id", areaId);

        await supabaseClient
          .from("completions")
          .upsert({
            area_id: areaId,
            start_date: areaRow.start_date || completeDate,
            end_date: completeDate,
            leader: leaderName
          }, { onConflict: "area_id, end_date" });

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
    } else if (completeResult) {
      setStatus("모든 카드의 최근방문일이 기록되었습니다. 완료내역이 업데이트되었습니다.");
    } else {
      setStatus("방문내역이 기록되었습니다.");
    }
  } catch (err) {
    console.error("Save visit error:", err);
    alert("오류가 발생했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};
