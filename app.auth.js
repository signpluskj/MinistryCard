const updateMenuVisibility = () => {
  if (!state.user) return;
  const userRole = String(state.user.role || "").trim();
  const isAdmin = userRole === "관리자";
  const isLeader = userRole === "인도자";
  const isSuper = state.isSuperAdmin === true;

  document.querySelectorAll(".admin-only").forEach((el) => {
    el.classList.toggle("hidden", !isAdmin);
  });

  document.querySelectorAll(".leader-only").forEach((el) => {
    el.classList.toggle("hidden", !isAdmin && !isLeader);
  });

  document.querySelectorAll(".super-admin-only").forEach((el) => {
    el.classList.toggle("hidden", !isSuper);
  });
};

const enterDashboard = async (user) => {
  state.user = user;
  elements.userInfo.textContent = `${state.user.name} (${state.user.role})`;
  elements.menuToggle.style.display = "inline-block";
  updateMenuVisibility();
  elements.loginPanel.classList.add("hidden");
  elements.dashboard.classList.remove("hidden");
  await loadData();
  if (!state.expandedAreaId && state.filterArea === "all") {
    const inProgressAreaId = getFirstInProgressArea();
    if (inProgressAreaId) {
      state.expandedAreaId = inProgressAreaId;
      state.filterArea = inProgressAreaId;
      state.selectedArea = inProgressAreaId;
      state.scrollAreaToTop = true;
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
          const superUser = { name: "최고관리자", role: "관리자" };
          try {
            window.localStorage.setItem(
              "mcUser",
              JSON.stringify({ name: "최고관리자", role: "관리자", isSuper: true })
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

    state.isSuperAdmin = false;
    const user = { role: data.role || "전도인", name: data.name };
    try {
      window.localStorage.setItem(
        "mcUser",
        JSON.stringify({ name: user.name, role: user.role })
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
    state.isSuperAdmin = data.isSuper === true;
    await enterDashboard({ name: data.name, role: data.role });
  } catch (e) {}
};

const logout = () => {
  state.user = null;
  state.isSuperAdmin = false;
  elements.configPanel.classList.add("hidden");
  try {
    window.localStorage.removeItem("mcUser");
  } catch (e) {}
  elements.userInfo.textContent = "";
  elements.menuToggle.style.display = "none";
  elements.dashboard.classList.add("hidden");
  elements.loginPanel.classList.remove("hidden");
  elements.nameInput.value = "";
  elements.passwordInput.value = "";
};

tryAutoLogin();
