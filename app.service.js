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
