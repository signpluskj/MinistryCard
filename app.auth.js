const updateMenuVisibility = () => {
  if (!state.user) return;
  const isAdm = isAdmin();
  const isLdr = isLeader();
  const isSuper = isSuperAdmin();

  document.querySelectorAll(".admin-only").forEach((el) => {
    el.classList.toggle("hidden", !isAdm);
  });

  document.querySelectorAll(".leader-only").forEach((el) => {
    el.classList.toggle("hidden", !isLdr);
  });

  document.querySelectorAll(".super-admin-only").forEach((el) => {
    el.classList.toggle("hidden", !isSuper);
  });
};

const enterDashboard = async (user) => {
  state.user = user;
  elements.userInfo.textContent = `${state.user.name} (${state.user.role})`;
  elements.menuToggle.classList.remove("hidden");
  updateMenuVisibility();
  elements.loginPanel.classList.add("hidden");
  elements.dashboard.classList.remove("hidden");
  await loadData();
  if (state.user && state.data.evangelists && state.data.evangelists.length) {
    const freshUser = state.data.evangelists.find((e) => e["이름"] === state.user.name);
    if (freshUser && freshUser["역할"] && freshUser["역할"] !== state.user.role) {
      state.user.role = freshUser["역할"];
      state.isSuperAdmin = state.user.role === "최고관리자";
      elements.userInfo.textContent = `${state.user.name} (${state.user.role})`;
      updateMenuVisibility();
      try {
        window.localStorage.setItem(
          "mcUser",
          JSON.stringify({ name: state.user.name, role: state.user.role, isSuper: state.isSuperAdmin })
        );
      } catch (e) {}
    }
  }
  if (!state.expandedAreaId && state.filterArea === "all") {
    const inProgressAreas = getAllInProgressAreas();
    if (inProgressAreas.length === 1) {
      // 진행중 구역이 1개면 기존처럼 펼침
      const inProgressAreaId = inProgressAreas[0];
      state.expandedAreaId = inProgressAreaId;
      state.filterArea = inProgressAreaId;
      state.selectedArea = inProgressAreaId;
      state.scrollAreaToTop = true;
    } else if (inProgressAreas.length > 1) {
      // 진행중 구역이 2개 이상이면 펼치지 않음
      state.scrollAreaToTop = false;
      state.expandedAreaId = null;
      state.filterArea = "all";
      state.selectedArea = null;
    }
  }
  renderAreas();
  renderCards();
  state.carAssignDate = todayISO();
  state.carAssignSlot = "오전";
  ensureAssignmentState();
  renderAdminPanel();
  renderMyCarInfo();
};

const login = async () => {
  const name = elements.nameInput.value.trim();
  const password = elements.passwordInput.value.trim();
  if (!name) {
    alert("이름을 입력해 주세요.");
    return;
  }
  setLoading(true, "로그인 중입니다...");
  try {
    if (!supabaseClient) throw new Error("Supabase 클라이언트가 초기화되지 않았습니다.");

    try {
      const { data: config } = await supabaseClient
        .from("site_config")
        .select("config_value")
        .eq("config_key", "super_admin")
        .single();

      if (config && config.config_value) {
        const admin = config.config_value;
        if (name === admin.id && password === admin.pass) {
          state.isSuperAdmin = true;
          const superUser = { name: "최고관리자", role: "최고관리자" };
          try {
            window.localStorage.setItem(
              "mcUser",
              JSON.stringify({ name: "최고관리자", role: "최고관리자", isSuper: true })
            );
          } catch (e) {}
          await enterDashboard(superUser);
          return;
        }
      }
    } catch (adminErr) {
      console.warn("Super admin check skipped:", adminErr);
    }

    const { data, error } = await supabaseClient
      .from("evangelists")
      .select("*")
      .eq("name", name)
      .eq("password", password)
      .single();

    if (error || !data) {
      alert("이름 또는 비밀번호가 올바르지 않습니다.");
      return;
    }

    const userRole = String(data.role || "전도인").trim();
    state.isSuperAdmin = userRole === "최고관리자";
    const user = { role: userRole, name: data.name };
    try {
      window.localStorage.setItem(
        "mcUser",
        JSON.stringify({ name: user.name, role: user.role, isSuper: state.isSuperAdmin })
      );
    } catch (e) {}
    await enterDashboard(user);
  } catch (err) {
    console.error("Login error:", err);
    alert("로그인 중 오류가 발생했습니다: " + err.message);
  } finally {
    setLoading(false);
  }
};

const tryAutoLogin = async () => {
  try {
    const raw = window.localStorage.getItem("mcUser");
    if (!raw) {
      return;
    }
    const data = JSON.parse(raw);
    if (!data || !data.name || !data.role) {
      return;
    }
    const userRole = String(data.role || "").trim();
    state.isSuperAdmin = data.isSuper === true || userRole === "최고관리자";
    await enterDashboard({ name: data.name, role: userRole });
  } catch (e) {}
};

window.logout = () => {
  state.user = null;
  state.isSuperAdmin = false;
  if (elements.configPanel) elements.configPanel.classList.add("hidden");
  try {
    window.localStorage.removeItem("mcUser");
  } catch (e) {}
  if (elements.userInfo) elements.userInfo.textContent = "";
  if (elements.menuToggle) elements.menuToggle.classList.add("hidden");
  if (elements.dashboard) elements.dashboard.classList.add("hidden");
  if (elements.loginPanel) elements.loginPanel.classList.remove("hidden");
  if (elements.nameInput) elements.nameInput.value = "";
  if (elements.passwordInput) elements.passwordInput.value = "";
  
  // 사이드 메뉴도 닫기
  if (elements.sideMenu) elements.sideMenu.classList.add("hidden");
};
