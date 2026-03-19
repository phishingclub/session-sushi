async function loadTabHTML(tabName, containerId) {
  try {
    const response = await fetch(`popup/tabs/${tabName}/${tabName}.html`);
    if (!response.ok) {
      console.error(`Failed to load ${tabName} tab HTML`);
      return;
    }
    const html = await response.text();
    const container = document.getElementById(containerId);
    if (container) {
      container.innerHTML = html;
    }
  } catch (error) {
    console.error(`Error loading ${tabName} tab:`, error);
  }
}

async function loadAllTabs() {
  await Promise.all([
    loadTabHTML("cookies", "cookies-tab"),
    loadTabHTML("sessions", "sessions-tab"),
    loadTabHTML("graph", "graph-tab"),
    loadTabHTML("user", "user-tab"),
    loadTabHTML("calendar", "calendar-tab"),
    loadTabHTML("sharepoint", "sharepoint-tab"),
    loadTabHTML("teams", "teams-tab"),
    loadTabHTML("directory", "directory-tab"),
    loadTabHTML("mailbox", "mailbox-tab"),
    loadTabHTML("onedrive", "onedrive-tab"),
    loadTabHTML("settings", "settings-tab"),
  ]);
}

loadAllTabs().then(() => {
  if (typeof initializeApp === "function") {
    initializeApp();
  }
});
