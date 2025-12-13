async function initializeTheme() {
  try {
    const result = await chrome.storage.local.get([THEME_STORAGE_KEY]);
    let theme = result[THEME_STORAGE_KEY];

    if (!theme) {
      const prefersDark =
        window.matchMedia &&
        window.matchMedia("(prefers-color-scheme: dark)").matches;
      theme = prefersDark ? "dark" : "light";
    }

    applyTheme(theme);
  } catch (error) {
    console.error("Error initializing theme:", error);
    applyTheme("light");
  }
}

function applyTheme(theme) {
  document.documentElement.setAttribute("data-theme", theme);
  updateThemeToggleIcon(theme);

  chrome.storage.local.set({ [THEME_STORAGE_KEY]: theme }).catch((err) => {
    console.error("Error saving theme:", err);
  });
}

function updateThemeToggleIcon(theme) {
  const toggleBtn = document.getElementById("themeToggle");
  if (toggleBtn) {
    toggleBtn.textContent = theme === "dark" ? "☀️" : "🌙";
    toggleBtn.title =
      theme === "dark" ? "Switch to light mode" : "Switch to dark mode";
  }
}

function toggleTheme() {
  const currentTheme =
    document.documentElement.getAttribute("data-theme") || "light";
  const newTheme = currentTheme === "dark" ? "light" : "dark";
  applyTheme(newTheme);

  showToast(`Switched to ${newTheme} mode`, "info");
}

if (window.matchMedia) {
  window
    .matchMedia("(prefers-color-scheme: dark)")
    .addEventListener("change", async (e) => {
      const result = await chrome.storage.local.get([THEME_STORAGE_KEY]);
      if (!result[THEME_STORAGE_KEY]) {
        applyTheme(e.matches ? "dark" : "light");
      }
    });
}
