async function saveUIState(updates) {
  try {
    const result = await chrome.storage.local.get([UI_STATE_STORAGE_KEY]);
    const currentState = result[UI_STATE_STORAGE_KEY] || {};
    const newState = { ...currentState, ...updates };
    await chrome.storage.local.set({ [UI_STATE_STORAGE_KEY]: newState });
  } catch (error) {
    console.error("Failed to save UI state:", error);
  }
}

async function restoreUIState() {
  try {
    const result = await chrome.storage.local.get([UI_STATE_STORAGE_KEY]);
    const state = result[UI_STATE_STORAGE_KEY];

    if (!state) {
      return;
    }

    if (
      state.activeSessionIndex !== undefined &&
      state.activeSessionIndex !== null &&
      m365Sessions[state.activeSessionIndex]
    ) {
      window._restoringSession = true;
      loadM365Session(state.activeSessionIndex);
      window._restoringSession = false;
    }

    if (state.activeTab) {
      switchTab(state.activeTab);
    }
  } catch (error) {
    console.error("Failed to restore UI state:", error);
  }
}

function updateTabsScroll() {
  const tabs = document.querySelector(".tabs");
  const wrapper = document.querySelector(".tabs-wrapper");
  const leftBtn = document.getElementById("tabScrollLeft");
  const rightBtn = document.getElementById("tabScrollRight");
  if (!tabs || !wrapper) return;

  const canLeft = tabs.scrollLeft > 1;
  const canRight = tabs.scrollLeft + tabs.clientWidth < tabs.scrollWidth - 1;

  wrapper.classList.toggle("can-scroll-left", canLeft);
  wrapper.classList.toggle("can-scroll-right", canRight);
  if (leftBtn) leftBtn.hidden = !canLeft;
  if (rightBtn) rightBtn.hidden = !canRight;
}

function switchTab(tabName) {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.tab === tabName);
  });

  const activeBtn = document.querySelector(`.tab-btn[data-tab="${tabName}"]`);
  if (activeBtn) {
    activeBtn.scrollIntoView({ inline: "nearest", block: "nearest", behavior: "smooth" });
    setTimeout(updateTabsScroll, 320);
  }

  document.querySelectorAll(".tab-content").forEach((content) => {
    const isActive = content.id === `${tabName}-tab`;
    content.classList.toggle("active", isActive);
  });

  const headerImport = document.getElementById("headerImport");
  const headerExport = document.getElementById("headerExport");
  const clearAllBtn = document.getElementById("clearAllCookies");

  if (tabName === "cookies" || tabName === "sessions") {
    headerImport.classList.remove("hidden");
    headerExport.classList.remove("hidden");
    clearAllBtn.classList.remove("hidden");
  } else {
    headerImport.classList.add("hidden");
    headerExport.classList.add("hidden");
    clearAllBtn.classList.add("hidden");
  }

  saveUIState({ activeTab: tabName });

  if (tabName === "mailbox") {
    initializeMailbox();
  }

  if (tabName === "onedrive") {
    initializeOneDrive();
  }

  if (tabName === "user") {
    initializeUser();
  }

  if (tabName === "directory") {
    initializeDirectory();
  }

  if (tabName === "calendar") {
    initializeCalendar();
  }

  if (tabName === "sharepoint") {
    initializeSharePoint();
  }

  if (tabName === "teams") {
    initializeTeams();
    setupTeamsListeners();
  }
}
