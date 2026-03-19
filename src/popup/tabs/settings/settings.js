const SETTINGS_STORAGE_KEY = "session_sushi_settings";

async function loadSettings() {
  const result = await chrome.storage.local.get([SETTINGS_STORAGE_KEY]);
  return result[SETTINGS_STORAGE_KEY] || {};
}

async function saveSettings(updates) {
  const current = await loadSettings();
  await chrome.storage.local.set({
    [SETTINGS_STORAGE_KEY]: { ...current, ...updates },
  });
}

function updateProxyBadge(host, port) {
  const badge = document.getElementById("proxyStatusBadge");
  if (!badge) return;
  if (host && port) {
    badge.textContent = `${host}:${port}`;
    badge.className = "proxy-badge proxy-badge-on";
  } else {
    badge.textContent = "Off";
    badge.className = "proxy-badge proxy-badge-off";
  }
}

function updateUABadge(ua) {
  const badge = document.getElementById("uaStatusBadge");
  if (!badge) return;
  if (ua) {
    badge.textContent = "Active";
    badge.className = "proxy-badge proxy-badge-on";
  } else {
    badge.textContent = "Off";
    badge.className = "proxy-badge proxy-badge-off";
  }
}

async function initializeSettings() {
  const settings = await loadSettings();

  const hostInput = document.getElementById("proxyHost");
  const portInput = document.getElementById("proxyPort");
  const uaInput = document.getElementById("userAgentInput");

  if (hostInput && settings.proxyHost) hostInput.value = settings.proxyHost;
  if (portInput && settings.proxyPort) portInput.value = settings.proxyPort;
  if (uaInput && settings.userAgent) uaInput.value = settings.userAgent;

  updateProxyBadge(settings.proxyHost, settings.proxyPort);
  updateUABadge(settings.userAgent);
}

async function applyProxy() {
  const host = document.getElementById("proxyHost")?.value.trim();
  const portRaw = document.getElementById("proxyPort")?.value.trim();

  if (!host) {
    showToast("Proxy host is required", "error");
    return;
  }

  const port = parseInt(portRaw, 10);
  if (!portRaw || isNaN(port) || port < 1 || port > 65535) {
    showToast("A valid port (1–65535) is required", "error");
    return;
  }

  try {
    await chrome.proxy.settings.set({
      value: {
        mode: "fixed_servers",
        rules: {
          singleProxy: {
            scheme: "socks5",
            host,
            port,
          },
        },
      },
      scope: "regular",
    });

    await saveSettings({ proxyHost: host, proxyPort: portRaw });
    updateProxyBadge(host, portRaw);
    showToast(
      "SOCKS5 proxy applied — all browser traffic is now routed through it",
      "info",
    );
  } catch (err) {
    console.error("Failed to apply proxy:", err);
    showToast("Failed to apply proxy: " + err.message, "error");
  }
}

async function clearProxy() {
  try {
    await chrome.proxy.settings.clear({ scope: "regular" });
    await saveSettings({ proxyHost: "", proxyPort: "" });

    const hostInput = document.getElementById("proxyHost");
    const portInput = document.getElementById("proxyPort");
    if (hostInput) hostInput.value = "";
    if (portInput) portInput.value = "";

    updateProxyBadge("", "");
    showToast("Proxy cleared", "success");
  } catch (err) {
    console.error("Failed to clear proxy:", err);
    showToast("Failed to clear proxy: " + err.message, "error");
  }
}

async function applyUserAgent() {
  const ua = document.getElementById("userAgentInput")?.value.trim();

  try {
    if (ua) {
      await chrome.declarativeNetRequest.updateDynamicRules({
        removeRuleIds: [2],
        addRules: [
          {
            id: 2,
            priority: 1,
            action: {
              type: "modifyHeaders",
              requestHeaders: [
                { header: "User-Agent", operation: "set", value: ua },
              ],
            },
            condition: {
              urlFilter: "*",
              resourceTypes: [
                "main_frame",
                "sub_frame",
                "xmlhttprequest",
                "script",
                "stylesheet",
                "image",
                "font",
                "object",
                "media",
                "websocket",
                "other",
              ],
            },
          },
        ],
      });
    } else {
      await chrome.declarativeNetRequest.updateDynamicRules({
        removeRuleIds: [2],
        addRules: [],
      });
    }

    await saveSettings({ userAgent: ua });
    updateUABadge(ua);
    showToast(
      ua
        ? "User agent applied — all browser traffic will use this user agent"
        : "User agent cleared",
      ua ? "info" : "success",
    );
  } catch (err) {
    console.error("Failed to apply user agent:", err);
    showToast("Failed to apply user agent: " + err.message, "error");
  }
}

async function clearUserAgent() {
  const uaInput = document.getElementById("userAgentInput");
  if (uaInput) uaInput.value = "";

  try {
    await chrome.declarativeNetRequest.updateDynamicRules({
      removeRuleIds: [2],
      addRules: [],
    });
    await saveSettings({ userAgent: "" });
    updateUABadge("");
    showToast("User agent cleared", "success");
  } catch (err) {
    console.error("Failed to clear user agent:", err);
    showToast("Failed to clear user agent: " + err.message, "error");
  }
}

function setupSettingsListeners() {
  document
    .getElementById("applyProxyBtn")
    ?.addEventListener("click", applyProxy);
  document
    .getElementById("clearProxyBtn")
    ?.addEventListener("click", clearProxy);
  document
    .getElementById("applyUABtn")
    ?.addEventListener("click", applyUserAgent);
  document
    .getElementById("clearUABtn")
    ?.addEventListener("click", clearUserAgent);
}
