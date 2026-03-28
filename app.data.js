const saveVolunteerWeekToSupabase = async (weekStart, weekData) => {
  if (!supabaseClient) return { success: false, message: "Supabase client not initialized" };
  try {
    const { error } = await supabaseClient
      .from("volunteer_weeks")
      .upsert({
        week_start: weekStart,
        data: weekData
      });

    if (error) throw error;
    return { success: true };
  } catch (err) {
    console.error("Failed to save volunteer week:", err);
    return { success: false, message: err.message };
  }
};

const updateCardFlagsInSupabase = async (areaId, cardNumber, flags) => {
  if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

  const payload = {};
  if (flags.hasOwnProperty("revisit")) payload.revisit = flags.revisit;
  if (flags.hasOwnProperty("study")) payload.study = flags.study;
  if (flags.hasOwnProperty("sixMonths")) payload.six_months = flags.sixMonths;
  if (flags.hasOwnProperty("banned")) payload.banned = flags.banned;
  if (flags.hasOwnProperty("invite")) payload.invite = flags.invite;

  const { error } = await supabaseClient
    .from("cards")
    .update(payload)
    .eq("area_id", String(areaId))
    .eq("card_number", String(cardNumber));

  if (error) throw error;
  return { success: true, ...flags };
};

const deleteCardInSupabase = async (areaId, cardNumber) => {
  if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

  const { data: card, error: selectError } = await supabaseClient
    .from("cards")
    .select("*")
    .eq("area_id", String(areaId))
    .eq("card_number", String(cardNumber))
    .single();

  if (selectError) {
    console.warn("Card not found in Supabase or already deleted:", selectError);
  }

  if (card) {
    const { error: insertError } = await supabaseClient
      .from("deleted_cards")
      .insert({
        area_id: String(card.area_id),
        card_number: String(card.card_number),
        address: card.address,
        deleted_at: new Date().toISOString()
      });

    if (insertError) {
      console.error("Failed to archive deleted card:", insertError);
      throw insertError;
    }

    const { error: deleteError } = await supabaseClient
      .from("cards")
      .delete()
      .eq("area_id", String(areaId))
      .eq("card_number", String(cardNumber));

    if (deleteError) {
      console.error("Failed to delete card from cards table:", deleteError);
      throw deleteError;
    }
  }

  return { success: true };
};

const restoreDeletedCardInSupabase = async (areaId, cardNumber) => {
  if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

  const { data: deletedCard, error: selectError } = await supabaseClient
    .from("deleted_cards")
    .select("*")
    .eq("area_id", String(areaId))
    .eq("card_number", String(cardNumber))
    .single();

  if (selectError) throw selectError;

  if (deletedCard) {
    const { error: insertError } = await supabaseClient
      .from("cards")
      .upsert({
        area_id: String(deletedCard.area_id),
        card_number: String(deletedCard.card_number),
        address: deletedCard.address
      }, { onConflict: "area_id, card_number" });

    if (insertError) throw insertError;

    const { error: deleteError } = await supabaseClient
      .from("deleted_cards")
      .delete()
      .eq("area_id", String(areaId))
      .eq("card_number", String(cardNumber));

    if (deleteError) throw deleteError;
  }

  return { success: true };
};

const purgeDeletedCardInSupabase = async (areaId, cardNumber) => {
  if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

  const { error } = await supabaseClient
    .from("deleted_cards")
    .delete()
    .eq("area_id", String(areaId))
    .eq("card_number", String(cardNumber));

  if (error) throw error;
  return { success: true };
};

const apiRequest = async (action, payload = {}, method = "POST") => {
  if (!state.apiUrl) {
    if (state.isSuperAdmin) {
      elements.configPanel.classList.remove("hidden");
    }
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

const loadVolunteerConfig = async () => {
  if (!supabaseClient) return;
  try {
    const { data, error } = await supabaseClient
      .from("volunteer_weeks")
      .select("*")
      .order("week_start", { ascending: true });

    if (error) throw error;

    if (data && data.length > 0) {
      state.volunteerWeeks = data.map(row => ({
        ...row.data,
        weekStartText: row.week_start
      }));
    } else {
      state.volunteerWeeks = [];
    }
    ensureVolunteerSelection();
  } catch (err) {
    console.error("Failed to load volunteer config:", err);
  }
};

const migrateToSupabase = async () => {
  if (!supabaseClient) {
    alert("Supabase 설정이 올바르지 않습니다.");
    return;
  }
  if (!confirm("구글 시트의 데이터를 Supabase로 복사하시겠습니까?")) {
    return;
  }

  setLoading(true, "구글 시트에서 데이터 가져오는 중...");
  try {
    const data = await apiRequest("bootstrap", {}, "GET");
    if (data.error) throw new Error(data.error);

    console.log("Data from Google Sheets:", data);

    const checkError = (res, stage) => {
      if (res.error) {
        console.error(`Error at ${stage}:`, res.error);
        throw new Error(`[${stage}] ${res.error.message}\n${res.error.details || ""}`);
      }
    };

    setLoading(true, "구역번호(areas) 동기화 중...");
    const areas = (data.areas || []).map(a => ({
      area_id: String(a["구역번호"] || ""),
      start_date: toIsoDate(a["시작날짜"]),
      end_date: toIsoDate(a["완료날짜"]),
      leader: a["인도자"],
      start_date_backup: toIsoDate(a["시작날짜백업"]),
      end_date_backup: toIsoDate(a["완료날짜백업"]),
      leader_backup: a["인도자백업"]
    })).filter(a => a.area_id);
    console.log("Processed areas:", areas);
    if (areas.length) {
      const res = await supabaseClient.from("areas").upsert(areas, { onConflict: "area_id" });
      checkError(res, "areas");
    }

    setLoading(true, "구역카드(cards) 동기화 중...");
    const cards = (data.cards || []).map(c => ({
      area_id: String(c["구역번호"] || ""),
      card_number: String(c["카드번호"] || ""),
      address: c["주소"],
      detail_address: c["상세주소"],
      memo: c["정보"],
      town: c["읍면동"],
      recent_visit_date: toIsoDate(c["최근방문일"]),
      prev_visit_date: toIsoDate(c["이전봉사일"]),
      meet: !!c["만남"],
      absent: !!c["부재"],
      revisit: !!c["재방"],
      study: !!c["연구"],
      six_months: !!c["6개월"],
      banned: !!c["방문금지"],
      car_id: String(c["차량"] || ""),
      assignment_date: toIsoDate(c["배정날짜"]),
      invite: !!c["초대장"]
    })).filter(c => c.area_id && c.card_number);
    console.log("Processed cards:", cards);
    if (cards.length) {
      const res = await supabaseClient.from("cards").upsert(cards, { onConflict: "area_id, card_number" });
      checkError(res, "cards");
    }

    setLoading(true, "삭제된카드(deleted_cards) 동기화 중...");
    const deletedCards = (data.deletedCards || []).map(dc => ({
      area_id: String(dc["구역번호"] || ""),
      card_number: String(dc["카드번호"] || ""),
      address: dc["주소"],
      deleted_at: toIsoDate(dc["삭제일"]) || new Date().toISOString()
    })).filter(dc => dc.area_id && dc.card_number);
    console.log("Processed deleted cards:", deletedCards);
    if (deletedCards.length) {
      const res = await supabaseClient.from("deleted_cards").upsert(deletedCards, { onConflict: "area_id, card_number" });
      checkError(res, "deleted_cards");
    }

    setLoading(true, "전도인명단(evangelists) 동기화 중...");
    const evangelists = (data.evangelists || []).map(e => ({
      name: String(e["이름"] || ""),
      password: String(e["비밀번호"] || ""),
      role: e["역할"] || e["권한"] || "전도인",
      gender: e["성별"] || "",
      driver: isTrueValue(e["운전자"]),
      capacity: Number(e["차량"]) || 0,
      spouse: e["부부"] || "",
      is_deaf: isTrueValue(e["농인"])
    })).filter(e => e.name);
    console.log("Processed evangelists:", evangelists);
    if (evangelists.length) {
      const res = await supabaseClient.from("evangelists").upsert(evangelists, { onConflict: "name" });
      checkError(res, "evangelists");
    }

    setLoading(true, "완료내역(completions) 동기화 중...");
    const completions = (data.completions || []).map(c => ({
      area_id: String(c["구역번호"] || c["areaId"] || ""),
      start_date: toIsoDate(c["시작날짜"] || c["startDate"]),
      end_date: toIsoDate(c["완료날짜"] || c["completionDate"] || c["endDate"]),
      leader: c["인도자"] || c["leader"]
    })).filter(c => c.area_id && c.end_date);
    console.log("Processed completions:", completions);
    if (completions.length) {
      const res = await supabaseClient.from("completions").upsert(completions, { onConflict: "area_id, end_date" });
      checkError(res, "completions");
    }

    setLoading(true, "방문기록(visits) 동기화 중...");
    const visits = (data.visits || []).map(v => ({
      area_id: String(v["구역번호"] || v["areaId"] || ""),
      card_number: String(v["카드번호"] || v["구역카드"] || v["cardNumber"] || ""),
      visit_date: toIsoDate(v["방문날짜"] || v["방문일"] || v["날짜"]),
      worker: v["전도인"] || v["방문자"],
      result: v["결과"] || v["방문결과"],
      note: v["메모"] || v["비고"]
    })).filter(v => v.area_id && v.card_number);
    if (visits.length) {
      for (let i = 0; i < visits.length; i += 500) {
        const res = await supabaseClient.from("visits").insert(visits.slice(i, i + 500));
        checkError(res, `visits (chunk ${i / 500 + 1})`);
      }
    }

    setLoading(true, "차량배정(assignments) 동기화 중...");
    const assignments = (data.assignments || []).map(a => ({
      date: toIsoDate(a["날짜"]),
      slot: a["시간대"] || "오전",
      car_id: String(a["차량"] || ""),
      driver: a["운전자"] || a["이름"],
      passengers: Array.isArray(a["동승자"]) ? a["동승자"] : []
    }));
    if (assignments.length) {
      const res = await supabaseClient.from("car_assignments").insert(assignments);
      checkError(res, "car_assignments");
    }

    setLoading(true, "봉사신청(volunteer_weeks) 동기화 중...");
    try {
      const volData = await apiRequest("getVolunteerConfig", {});
      if (volData && volData.weeks && volData.weeks.length > 0) {
        for (const week of volData.weeks) {
          const weekStart = week.weekStartISO || toIsoDate(week.weekStartText);
          if (weekStart) {
            const { error: volError } = await supabaseClient
              .from("volunteer_weeks")
              .upsert({
                week_start: weekStart,
                data: week
              });
            if (volError) console.error("Volunteer sync error for week " + weekStart, volError);
          }
        }
      }
    } catch (err) {
      console.warn("Volunteer data sync failed (optional):", err);
    }

    setLoading(true, "초대장배부(invite_campaign) 동기화 중...");
    try {
      if (data.inviteCampaign) {
        const ic = data.inviteCampaign;
        if (ic.startDate) {
          await supabaseClient.from("invite_campaign").upsert({
            status: ic.active ? "active" : "finished",
            start_date: toIsoDate(ic.startDate),
            end_date: toIsoDate(ic.endDate),
            memo: ic.memo || ""
          }, { onConflict: "start_date" });
        }
      }
    } catch (err) {
      console.warn("Invite campaign sync failed (optional):", err);
    }

    alert("동기화가 완료되었습니다!");
    location.reload();
  } catch (err) {
    console.error("Migration error:", err);
    alert("동기화 실패: " + err.message);
  } finally {
    setLoading(false);
  }
};

const migrateToSheets = async () => {
  if (!state.isSuperAdmin) {
    alert("최고관리자 권한이 필요합니다.");
    return;
  }
  if (!confirm("Supabase의 데이터를 구글 시트로 '전체 덮어쓰기' 하시겠습니까?\n이 작업은 구글 시트의 기존 내용을 삭제하고 Supabase 내용으로 교체합니다.")) {
    return;
  }

  setLoading(true, "Supabase에서 데이터 가져오는 중...");
  try {
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

    const { data: areas } = await supabaseClient.from("areas").select("*");
    const { data: cards } = await supabaseClient.from("cards").select("*");
    const { data: deletedCards } = await supabaseClient.from("deleted_cards").select("*");
    const { data: evangelists } = await supabaseClient.from("evangelists").select("*");
    const { data: completions } = await supabaseClient.from("completions").select("*");
    const { data: visits } = await supabaseClient.from("visits").select("*").order("visit_date", { ascending: false });
    const { data: assignments } = await supabaseClient.from("car_assignments").select("*");
    const { data: volunteerWeeks } = await supabaseClient.from("volunteer_weeks").select("*");
    const { data: inviteCampaign } = await supabaseClient.from("invite_campaign").select("*");

    const payload = {
      areas: (areas || []).map(a => ({
        "구역번호": a.area_id, 
        "시작날짜": a.start_date, 
        "완료날짜": a.end_date, 
        "인도자": a.leader,
        "시작날짜백업": a.start_date_backup,
        "완료날짜백업": a.end_date_backup,
        "인도자백업": a.leader_backup
      })),
      cards: (cards || []).map(c => ({
        "구역번호": c.area_id, 
        "카드번호": c.card_number, 
        "읍면동": c.town, 
        "주소": c.address, 
        "상세주소": c.detail_address,
        "정보": c.memo,
        "정기방문자": c.regular_worker,
        "최근방문일": c.recent_visit_date, 
        "이전봉사일": c.prev_visit_date,
        "만남": c.meet, 
        "부재": c.absent, 
        "재방": c.revisit, 
        "연구": c.study,
        "6개월": c.six_months, 
        "방문금지": c.banned, 
        "차량": c.car_id, 
        "배정날짜": c.assignment_date,
        "초대장": c.invite
      })),
      deletedCards: (deletedCards || []).map(dc => ({
        "구역번호": dc.area_id, "카드번호": dc.card_number, "주소": dc.address, "삭제일": dc.deleted_at
      })),
      evangelists: (evangelists || []).map(e => ({
        "이름": e.name, 
        "비밀번호": e.password, 
        "역할": e.role,
        "성별": e.gender,
        "농인": e.is_deaf,
        "운전자": e.driver,
        "차량": e.capacity,
        "부부": e.spouse
      })),
      completions: (completions || []).map(c => ({
        "구역번호": c.area_id, 
        "완료날짜": c.end_date, 
        "인도자": c.leader,
        "비고": ""
      })),
      visits: (visits || []).map(v => ({
        "구역번호": v.area_id,
        "카드번호": v.card_number,
        "방문날짜": v.visit_date,
        "전도인": v.worker,
        "결과": v.result,
        "메모": v.note
      })),
      assignments: (assignments || []).map(a => ({
        "날짜": a.date,
        "시간대": a.slot,
        "차량": a.car_id,
        "이름": a.driver,
        "역할": (a.passengers || []).join(", ")
      })),
      volunteerWeeks: (volunteerWeeks || []).map(vw => ({
        "주차시작": vw.week_start,
        "데이터": JSON.stringify(vw.data)
      })),
      inviteCampaign: (inviteCampaign || []).map(ic => ({
        "상태": ic.status,
        "시작일": ic.start_date,
        "종료일": ic.end_date,
        "메모": ic.memo
      }))
    };

    setLoading(true, "구글 시트로 데이터 전송 중...");
    const res = await apiRequest("syncFromSupabase", { data: JSON.stringify(payload) });

    if (res.success) {
      alert("구글 시트 덮어쓰기가 완료되었습니다!");
    } else {
      throw new Error(res.message);
    }
  } catch (err) {
    console.error("Sync to Sheets error:", err);
    alert("동기화 실패: " + err.message);
  } finally {
    setLoading(false);
  }
};

const loadData = async () => {
  setLoading(true, "데이터 불러오는 중...");
  try {
    if (supabaseClient) {
      const { data: areas, error: areaError } = await supabaseClient.from("areas").select("*");
      if (areaError) {
        console.error("Supabase load error (areas):", areaError);
      }
      if (areas && areas.length > 0) {
        const { data: cards, error: cardsError } = await supabaseClient.from("cards").select("*");
        if (cardsError) console.error("Supabase load error (cards):", cardsError);
        
        const { data: evangelists } = await supabaseClient.from("evangelists").select("*");
        const { data: assignments } = await supabaseClient.from("car_assignments").select("*").gte("date", toIsoDate(new Date()));
        const { data: inviteCampaign } = await supabaseClient.from("invite_campaign").select("*").order("created_at", { ascending: false }).limit(1);
        const { data: completions } = await supabaseClient.from("completions").select("*");
        const { data: deletedCards } = await supabaseClient.from("deleted_cards").select("*").order("deleted_at", { ascending: false });

        updateAppTitle();

        state.data.areas = areas.map(a => ({
          "구역번호": a.area_id, "시작날짜": a.start_date, "완료날짜": a.end_date, "인도자": a.leader
        }));
        state.data.cards = cards.map(c => ({
          id: c.id,
          "구역번호": c.area_id, "카드번호": c.card_number, "주소": c.address,
          "상세주소": c.detail_address, "정보": c.memo, "읍면동": c.town,
          "최근방문일": c.recent_visit_date, "이전봉사일": c.prev_visit_date,
          "만남": c.meet, "부재": c.absent, "재방": c.revisit, "연구": c.study,
          "6개월": c.six_months, "방문금지": c.banned, "차량": c.car_id, "배정날짜": c.assignment_date,
          "초대장": c.invite
        }));
        state.data.deletedCards = (deletedCards || []).map(dc => ({
          id: dc.id,
          "구역번호": dc.area_id, "카드번호": dc.card_number, "주소": dc.address, "삭제일": dc.deleted_at
        }));
        state.data.evangelists = evangelists.map(e => ({
          "이름": e.name, "비밀번호": e.password, "역할": e.role,
          "성별": e.gender, "농인": isTrueValue(e.is_deaf), "운전자": isTrueValue(e.driver),
          "차량": e.capacity, "부부": e.spouse
        }));
        state.data.assignments = assignments.map(a => ({
          id: a.id,
          "날짜": a.date, "시간대": a.slot, "차량": a.car_id, "이름": a.driver, "동승자": a.passengers || []
        }));
        state.inviteCampaign = inviteCampaign && inviteCampaign[0] ? {
          id: inviteCampaign[0].id,
          active: inviteCampaign[0].status === "active",
          startDate: inviteCampaign[0].start_date,
          endDate: inviteCampaign[0].end_date,
          memo: inviteCampaign[0].memo
        } : null;
        state.data.completions = (completions || []).map(c => ({
          "구역번호": c.area_id, "시작날짜": c.start_date, "완료날짜": c.end_date, "인도자": c.leader
        }));

        const { data: visits } = await supabaseClient.from("visits").select("*").order("visit_date", { ascending: false }).limit(1000);
        state.data.visits = (visits || []).map(v => ({
          id: v.id,
          "구역번호": v.area_id, "카드번호": v.card_number, "방문날짜": v.visit_date,
          "전도인": v.worker, "결과": v.result, "메모": v.note
        }));
        console.log("Supabase data loaded");
        return;
      }
    }

    const data = await apiRequest("bootstrap", {}, "GET");
    if (data.error) throw new Error(data.error);

    state.data.cards = data.cards || [];
    state.data.areas = data.areas || data.areaStatus || [];
    state.data.completions = data.completions || [];
    state.data.visits = data.visits || [];
    state.data.evangelists = data.evangelists || [];
    state.data.assignments = data.assignments || [];
    state.data.assignments = sanitizeAssignmentRows(state.data.assignments);
    state.inviteCampaign = data.inviteCampaign || null;
    state.data.deletedCards = data.deletedCards || [];

    console.log("Google Sheets data loaded (Fallback)");
  } finally {
    setLoading(false);
  }
};
