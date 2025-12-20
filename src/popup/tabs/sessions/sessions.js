// Known FOCI (Family of Client IDs) clients from Secureworks research
// Source: https://github.com/secureworks/family-of-client-ids-research
const FOCI_CLIENTS = {
  "14638111-3389-403d-b206-a6a71d9f8f16": "Copilot App",
  "598ab7bb-a59c-4d31-ba84-ded22c220dbd": "Designer App",
  "cde6adac-58fd-4b78-8d6d-9beaf1b0d668": "Global Secure Access Client",
  "0922ef46-e1b9-4f7e-9134-9ad00547eb41": "Loop",
  "eb20f3e3-3dce-4d2c-b721-ebb8d4414067": "Managed Meeting Rooms",
  "4813382a-8fa7-425e-ab75-3b753aab3abb": "Microsoft Authenticator App",
  "04b07795-8ddb-461a-bbee-02f9e1bf7b46": "Microsoft Azure CLI",
  "1950a258-227b-4e31-a9cf-717495945fc2": "Microsoft Azure PowerShell",
  "cf36b471-5b44-428c-9ce7-313bf84528de": "Microsoft Bing Search",
  "dd47d17a-3194-4d86-bfd5-c6ae6f5651e3": "Microsoft Defender for Mobile",
  "e9c51622-460d-4d3d-952d-966a5b1da34c": "Microsoft Edge",
  "ecd6b820-32c2-49b6-98a6-444530e5a77a": "Microsoft Edge (2)",
  "f44b1140-bc5e-48c6-8dc0-5cf5a53c0e34": "Microsoft Edge (3)",
  "d7b530a4-7680-4c23-a8bf-c52c121d2e87":
    "Microsoft Edge Enterprise New Tab Page",
  "82864fa0-ed49-4711-8395-a0e6003dca1f": "Microsoft Edge MSAv2",
  "57fcbcfa-7cee-4eb1-8b25-12d2030b4ee0": "Microsoft Flow",
  "9ba1a5c7-f17a-4de9-a1f1-6178c8d51223": "Microsoft Intune Company Portal",
  "a670efe7-64b6-454f-9ae9-4f1cf27aba58": "Microsoft Lists App on Android",
  "d3590ed6-52b3-4102-aeff-aad2292ab01c": "Microsoft Office",
  "66375f6b-983f-4c2c-9701-d680650f588f": "Microsoft Planner",
  "c0d2a505-13b8-4ae0-aa9e-cddd5eab0b12": "Microsoft Power BI",
  "1fec8e78-bce4-4aaf-ab1b-5451cc387264": "Microsoft Teams",
  "8ec6bc83-69c8-4392-8f08-b3c986009232": "Microsoft Teams-T4L",
  "22098786-6e16-43cc-a27d-191a01a1e3b5": "Microsoft To-Do client",
  "57336123-6e14-4acc-8dcf-287b6088aa28": "Microsoft Whiteboard Client",
  "540d4ff4-b4c0-44c1-bd06-cab1782d582a": "ODSP Mobile Lists App",
  "00b41c95-dab0-4487-9791-b9d2c32c80f2": "Office 365 Management",
  "0ec893e0-5785-4de6-99da-4ed124e5296c": "Office UWP PWA",
  "b26aadf8-566f-4478-926f-589f601d9c74": "OneDrive",
  "af124e86-4e96-495a-b70a-90f90ab96707": "OneDrive iOS App",
  "ab9b8c07-8f02-4f72-87fa-80105867a763": "OneDrive SyncEngine",
  "27922004-5251-4030-b22d-91ecd9a37ea4": "Outlook Mobile",
  "4e291c71-d680-4d0e-9640-0a3358e31177": "PowerApps",
  "d326c1ce-6cc6-4de2-bebc-4591e5e13ef0": "SharePoint",
  "f05ff7c9-f75a-4acd-a3b5-f4b6a870245d": "SharePoint Android",
  "872cd9fa-d31f-45e0-9eab-6e460a02d1f1": "Visual Studio",
  "26a7ee05-5602-4d76-a7ba-eae8b7b67941": "Windows Search",
  "a569458c-7f2b-45cb-bab9-b7dee514d112": "Yammer iPhone",
  "038ddad9-5bbe-4f64-b0cd-12434d1e633b": "ZTNA Network Access Client",
  "d5e23a82-d7e1-4886-af25-27037a0fdc2a": "ZTNA Network Access Client -- M365",
  "760282b4-0cfc-4952-b467-c8e0298fee16":
    "ZTNA Network Access Client -- Private",
};

const M365_PRESETS = {
  "graph-powershell-user": {
    clientId: "14d82eec-204b-4c2f-b7e8-296a70dab67e",
    redirectUri: "https://login.microsoftonline.com/common/oauth2/nativeclient",
    scope: "https://graph.microsoft.com/.default offline_access",
    auth_url: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    token_url: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    description: "Graph PowerShell (FOCI)",
  },
  "azure-cli": {
    clientId: "04b07795-8ddb-461a-bbee-02f9e1bf7b46",
    redirectUri: "http://localhost",
    scope: "https://graph.microsoft.com/.default offline_access",
    auth_url: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    token_url: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    description: "Azure CLI (FOCI)",
  },
  "azure-powershell": {
    clientId: "1950a258-227b-4e31-a9cf-717495945fc2",
    redirectUri: "urn:ietf:wg:oauth:2.0:oob",
    scope: "https://graph.microsoft.com/.default offline_access",
    auth_url: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    token_url: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    description: "Azure PowerShell (FOCI)",
  },
  "ms-teams": {
    clientId: "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
    redirectUri: "https://login.microsoftonline.com/common/oauth2/nativeclient",
    scope: "https://graph.microsoft.com/.default offline_access",
    auth_url: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    token_url: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    description: "Microsoft Teams (FOCI)",
  },
  intune: {
    clientId: "9ba1a5c7-f17a-4de9-a1f1-6178c8d51223",
    redirectUri:
      "ms-appx-web://Microsoft.AAD.BrokerPlugin/S-1-15-2-2666988183-1750391847-2906264630-3525785777-2857982319-3063633125-1907478113",
    scope: "openid offline_access https://graph.microsoft.com/.default",
    auth_url: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    token_url: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    description: "Intune (FOCI)",
  },
  "graph-explorer": {
    clientId: "de8bc8b5-d9f9-48b1-a8ad-b748da725064",
    redirectUri: "https://developer.microsoft.com/en-us/graph/graph-explorer",
    scope: "https://graph.microsoft.com/.default offline_access",
    auth_url: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    token_url: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    description: "Graph Explorer",
  },
};

const DEFAULT_PRESET = "ms-teams";

let isRefreshing = false;

function populatePresetDropdown() {
  const presetSelector = document.getElementById("presetSelector");
  if (!presetSelector) return;

  presetSelector.innerHTML = "";

  const customOption = document.createElement("option");
  customOption.value = "custom";
  customOption.textContent = "Custom Configuration";
  presetSelector.appendChild(customOption);

  Object.keys(M365_PRESETS).forEach((key) => {
    const option = document.createElement("option");
    option.value = key;

    option.textContent = M365_PRESETS[key].description;

    if (key === DEFAULT_PRESET) {
      option.selected = true;
    }

    presetSelector.appendChild(option);
  });
}

function applyPreset(preset) {
  const clientIdInput = document.getElementById("clientIdInput");
  const redirectUriInput = document.getElementById("redirectUriInput");
  const scopeInput = document.getElementById("scopeInput");
  const authUrlInput = document.getElementById("authUrlInput");
  const tokenUrlInput = document.getElementById("tokenUrlInput");

  if (preset === "custom") {
    // Enable inputs for custom configuration
    clientIdInput.readOnly = false;
    redirectUriInput.readOnly = false;
    scopeInput.readOnly = false;
    authUrlInput.readOnly = false;
    tokenUrlInput.readOnly = false;
    // Set default values for custom configuration
    clientIdInput.value = "";
    redirectUriInput.value =
      "https://login.microsoftonline.com/common/oauth2/nativeclient";
    scopeInput.value = "https://graph.microsoft.com/.default offline_access";
    authUrlInput.value =
      "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
    tokenUrlInput.value =
      "https://login.microsoftonline.com/common/oauth2/v2.0/token";
  } else if (M365_PRESETS[preset]) {
    // Make inputs readonly and populate with preset values
    clientIdInput.readOnly = true;
    redirectUriInput.readOnly = true;
    scopeInput.readOnly = true;
    authUrlInput.readOnly = true;
    tokenUrlInput.readOnly = true;
    clientIdInput.value = M365_PRESETS[preset].clientId;
    redirectUriInput.value = M365_PRESETS[preset].redirectUri;
    scopeInput.value = M365_PRESETS[preset].scope;
    authUrlInput.value = M365_PRESETS[preset].auth_url;
    tokenUrlInput.value = M365_PRESETS[preset].token_url;
  }
}

async function copyAuthUrl() {
  try {
    const clientId =
      document.getElementById("clientIdInput")?.value?.trim() ||
      "1b730954-1685-4b74-9bfd-dac224a7b894";
    const redirectUri =
      document.getElementById("redirectUriInput")?.value?.trim() ||
      "https://login.microsoftonline.com/common/oauth2/nativeclient";
    const scope =
      document.getElementById("scopeInput")?.value?.trim() ||
      "https://graph.microsoft.com/.default offline_access";
    const authUrl =
      document.getElementById("authUrlInput")?.value?.trim() ||
      "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";

    const codeVerifier = generateCodeVerifier();
    const codeChallenge = await generateCodeChallenge(codeVerifier);

    const authUrlObj = new URL(authUrl);
    authUrlObj.searchParams.set("client_id", clientId);
    authUrlObj.searchParams.set("response_type", "code");
    authUrlObj.searchParams.set("redirect_uri", redirectUri);
    authUrlObj.searchParams.set("scope", scope);
    authUrlObj.searchParams.set("response_mode", "query");
    authUrlObj.searchParams.set("code_challenge", codeChallenge);
    authUrlObj.searchParams.set("code_challenge_method", "S256");

    const fullAuthUrl = authUrlObj.toString();
    await navigator.clipboard.writeText(fullAuthUrl);
    showToast("Authorization URL copied to clipboard!", "success");
  } catch (error) {
    console.error("Error copying auth URL:", error);
    showToast("Failed to copy authorization URL", "error");
  }
}

function generateCodeVerifier() {
  const array = new Uint8Array(32);
  crypto.getRandomValues(array);
  return base64UrlEncode(array);
}

async function generateCodeChallenge(verifier) {
  const encoder = new TextEncoder();
  const data = encoder.encode(verifier);
  const hash = await crypto.subtle.digest("SHA-256", data);
  return base64UrlEncode(new Uint8Array(hash));
}

function base64UrlEncode(array) {
  let binary = "";
  for (let i = 0; i < array.length; i++) {
    binary += String.fromCharCode(array[i]);
  }
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
}

async function getGraphToken() {
  const getGraphTokenBtn = document.getElementById("getGraphToken");

  if (isPopupMode && !isConvertingToWindow) {
    isConvertingToWindow = true;

    await chrome.storage.session.set({
      pendingAction: "getGraphToken",
      graphTokenConfig: {
        clientId: document.getElementById("clientIdInput")?.value?.trim(),
        redirectUri: document.getElementById("redirectUriInput")?.value?.trim(),
        scope: document.getElementById("scopeInput")?.value?.trim(),
        authUrl: document.getElementById("authUrlInput")?.value?.trim(),
        tokenUrl: document.getElementById("tokenUrlInput")?.value?.trim(),
      },
    });

    await openInWindow();
    return;
  }

  getGraphTokenBtn.disabled = true;
  getGraphTokenBtn.textContent = "⏳ Authorizing...";

  try {
    const clientId =
      document.getElementById("clientIdInput")?.value?.trim() ||
      "1b730954-1685-4b74-9bfd-dac224a7b894";
    const redirectUri =
      document.getElementById("redirectUriInput")?.value?.trim() ||
      "https://login.microsoftonline.com/common/oauth2/nativeclient";
    const scope =
      document.getElementById("scopeInput")?.value?.trim() ||
      "https://graph.microsoft.com/.default offline_access";
    const authUrl =
      document.getElementById("authUrlInput")?.value?.trim() ||
      "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
    const tokenUrl =
      document.getElementById("tokenUrlInput")?.value?.trim() ||
      "https://login.microsoftonline.com/common/oauth2/v2.0/token";

    const presetSelector = document.getElementById("presetSelector");
    const presetValue = presetSelector ? presetSelector.value : "custom";

    let appName = "Custom App";
    if (presetValue === "custom") {
      appName = "Custom Configuration";
    } else if (M365_PRESETS[presetValue]) {
      appName = M365_PRESETS[presetValue].description;
    }

    const response = await chrome.runtime.sendMessage({
      action: "getGraphToken",
      clientId: clientId,
      redirectUri: redirectUri,
      scope: scope,
      authUrl: authUrl,
      tokenUrl: tokenUrl,
    });

    if (!response.success) {
      throw new Error(response.error || "Failed to get token");
    }

    const tokenData = response.tokenData;

    let userEmail = "Unknown";
    try {
      const payload = JSON.parse(atob(tokenData.access_token.split(".")[1]));
      userEmail =
        payload.upn || payload.unique_name || payload.email || "Unknown";
    } catch (e) {
      // Could not decode user from token
    }

    const newSession = {
      name: `${userEmail} (${appName})`,
      user: userEmail,
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token,
      expires_at: Date.now() + tokenData.expires_in * 1000,
      created_at: Date.now(),
      client_id: clientId,
      redirect_uri: redirectUri,
      scope: scope,
      auth_url: authUrl,
      token_url: tokenUrl,
    };

    await saveM365SessionToList(newSession);
    await loadM365Sessions();

    if (!activeM365Session || activeM365Session === newSession) {
      saveUIState({ activeSessionIndex: m365Sessions.length - 1 });
    }

    showToast("✅ Successfully obtained Graph API tokens!");
  } catch (error) {
    if (error.message.includes("closed by user")) {
      // Auth window closed by user
      return;
    }

    console.error("Error getting Graph token:", error);

    let errorMsg = error.message;

    if (error.message.includes("AADSTS65002")) {
      errorMsg = "This client requires special authentication flow";
      troubleshooting =
        "💡 Try the 'Graph PowerShell - Read Only' preset instead.";
    } else if (error.message.includes("AADSTS50011")) {
      errorMsg = "Redirect URI mismatch";
      troubleshooting =
        "💡 Make sure redirect URI is: https://login.microsoftonline.com/common/oauth2/nativeclient";
    } else if (
      error.message.includes("AADSTS65001") ||
      error.message.includes("consent")
    ) {
      errorMsg = "User consent required";
    } else if (error.message.includes("AADSTS700016")) {
      errorMsg = "Application not found in directory";
      troubleshooting = "💡 The client ID may not be valid for your tenant.";
    }

    showToast(`Failed to get Graph token: ${errorMsg}`, "error");
  } finally {
    getGraphTokenBtn.disabled = false;
    getGraphTokenBtn.textContent = "🔑 Authorize";
  }
}

async function checkCompletedAuth() {
  try {
    const result = await chrome.storage.local.get("lastTokenResult");
    const lastResult = result.lastTokenResult;

    if (!lastResult) return;

    const age = Date.now() - lastResult.timestamp;
    if (age > 30000) {
      chrome.storage.local.remove("lastTokenResult");
      return;
    }

    chrome.storage.local.remove("lastTokenResult");

    if (!lastResult.success) {
      console.error("Auth flow failed:", lastResult.error);
      let errorMsg = lastResult.error || "Unknown error";

      if (errorMsg.includes("AADSTS65002")) {
        showToast(
          `Error: ${errorMsg}. 💡 Try the 'Graph PowerShell - Read Only' preset instead.`,
          "error",
        );
      } else if (errorMsg.includes("closed by user")) {
        return;
      } else {
        showToast(`Error: ${errorMsg}`, "error");
      }
    }
  } catch (error) {
    console.error("Error checking completed auth:", error);
  }
}

async function refreshActiveM365Session() {
  if (isRefreshing) {
    return;
  }

  if (!activeM365Session || !activeM365Session.refresh_token) {
    showToast("No active session to refresh");
    return;
  }

  isRefreshing = true;

  const refreshBtn = document.getElementById("refreshActiveSession");
  if (refreshBtn) {
    refreshBtn.disabled = true;
    refreshBtn.textContent = "⏳ Refreshing...";
  }

  try {
    const clientId =
      activeM365Session.client_id || "1b730954-1685-4b74-9bfd-dac224a7b894";
    const scope =
      activeM365Session.scope ||
      "https://graph.microsoft.com/.default offline_access";

    const tokenUrl =
      activeM365Session.token_url ||
      "https://login.microsoftonline.com/common/oauth2/v2.0/token";

    const response = await chrome.runtime.sendMessage({
      action: "fetchWithoutOrigin",
      url: tokenUrl,
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: new URLSearchParams({
        client_id: clientId,
        scope: scope,
        refresh_token: activeM365Session.refresh_token,
        grant_type: "refresh_token",
      }).toString(),
    });

    if (!response.ok) {
      const errorData = JSON.parse(response.body);
      throw new Error(errorData.error_description || "Failed to refresh token");
    }

    const tokenData = JSON.parse(response.body);

    activeM365Session.access_token = tokenData.access_token;
    activeM365Session.refresh_token = tokenData.refresh_token;
    activeM365Session.expires_at = Date.now() + tokenData.expires_in * 1000;

    await chrome.storage.local.set({ [TOKEN_STORAGE_KEY]: activeM365Session });

    displayActiveSession(activeM365Session);

    // Find if this session already exists in the saved sessions list
    const sessionIndex = m365Sessions.findIndex(
      (s) =>
        s.user === activeM365Session.user &&
        s.created_at === activeM365Session.created_at,
    );

    if (sessionIndex >= 0) {
      // Update existing session in place
      m365Sessions[sessionIndex] = activeM365Session;
      await chrome.storage.local.set({ [SESSIONS_STORAGE_KEY]: m365Sessions });
      renderM365Sessions();
      saveUIState({ activeSessionIndex: sessionIndex });
    } else {
      // Session doesn't exist in list, add it
      await saveM365SessionToList(activeM365Session);
      await loadM365Sessions();
      saveUIState({ activeSessionIndex: m365Sessions.length - 1 });
    }

    showToast("✅ Token refreshed successfully!");

    const mailboxTab = document.getElementById("mailbox-tab");
    if (mailboxTab && mailboxTab.classList.contains("active")) {
      if (typeof initializeMailbox === "function") {
        initializeMailbox();
      }
    }
  } catch (error) {
    console.error("Error refreshing token:", error);
    showToast("Failed to refresh token: " + error.message, "error");
  } finally {
    if (refreshBtn) {
      refreshBtn.disabled = false;
      refreshBtn.textContent = "🔄 Refresh";
    }
    isRefreshing = false;
  }
}

function displayActiveSession(session) {
  const display = document.getElementById("activeSessionDisplay");
  const noActiveMessage = document.getElementById("noActiveSessionMessage");
  const nameEl = document.getElementById("activeSessionName");
  const userEl = document.getElementById("activeSessionUser");
  const expiryEl = document.getElementById("activeSessionExpiry");
  const createdAtEl = document.getElementById("activeSessionCreatedAt");
  const clientIdEl = document.getElementById("activeSessionClientId");
  const redirectUriEl = document.getElementById("activeSessionRedirectUri");
  const authUrlEl = document.getElementById("activeSessionAuthUrl");
  const tokenUrlEl = document.getElementById("activeSessionTokenUrl");

  if (!session || !session.name || session.name === "Unknown") {
    display.classList.remove("visible");
    if (noActiveMessage) noActiveMessage.classList.remove("hidden");
    updateQuickActionButtons([]);
    stopAutoRefreshTimer();
    return;
  }

  display.classList.add("visible");
  if (noActiveMessage) noActiveMessage.classList.add("hidden");
  nameEl.textContent = session.name || "Unnamed Session";
  userEl.textContent = session.user || "Unknown";

  // Display created_at timestamp
  if (session.created_at) {
    const createdDate = new Date(session.created_at);
    createdAtEl.textContent = createdDate.toLocaleString();
  } else {
    createdAtEl.textContent = "Not stored";
  }

  // Display client_id with dropdown for FOCI switching
  if (session.client_id) {
    clientIdEl.innerHTML = "";
    const select = document.createElement("select");
    select.className = "client-id-selector";

    // Get FOCI clients in original object order
    const fociClientEntries = Object.entries(FOCI_CLIENTS);

    // Add all FOCI clients in original order
    fociClientEntries.forEach(([clientId, name]) => {
      const option = document.createElement("option");
      option.value = clientId;
      option.textContent = `${name} (${clientId.substring(0, 8)}...)`;
      if (clientId === session.client_id) {
        option.selected = true;
      }
      select.appendChild(option);
    });

    select.addEventListener("change", async (e) => {
      const newClientId = e.target.value;
      if (newClientId !== session.client_id) {
        await switchActiveSessionClientId(newClientId);
      }
    });

    clientIdEl.appendChild(select);
  } else {
    clientIdEl.textContent = "Not stored";
  }

  // Display redirect_uri
  redirectUriEl.textContent = session.redirect_uri || "Not stored";
  redirectUriEl.style.wordBreak = "break-all";

  // Display auth_url
  authUrlEl.textContent = session.auth_url || "Not stored";
  authUrlEl.style.wordBreak = "break-all";

  // Display token_url
  tokenUrlEl.textContent = session.token_url || "Not stored";
  tokenUrlEl.style.wordBreak = "break-all";

  const expiryDate = new Date(session.expires_at);
  const now = new Date();
  const remainingMs = expiryDate - now;
  const expiryMinutes = Math.floor(remainingMs / 1000 / 60);

  // Check if auto-refresh is enabled
  chrome.storage.local.get([AUTO_REFRESH_STORAGE_KEY]).then((result) => {
    const autoRefreshEnabled = result[AUTO_REFRESH_STORAGE_KEY] === true;

    if (remainingMs > 0) {
      let expiryText = `${expiryMinutes} minutes (${expiryDate.toLocaleString()})`;
      if (autoRefreshEnabled && expiryMinutes < 5) {
        expiryText += " 🔄 Auto-refreshing...";
        expiryEl.className = "color-warning";
      } else if (autoRefreshEnabled) {
        expiryEl.className = "";
      } else {
        expiryEl.className = "";
      }
      expiryEl.textContent = expiryText;
    } else {
      let expiryText = `Expired at ${expiryDate.toLocaleString()}`;
      if (autoRefreshEnabled) {
        expiryText += " 🔄 Auto-refresh will attempt...";
      }
      expiryEl.textContent = expiryText;
      expiryEl.className = "color-danger";
    }
  });

  let scopes = [];
  try {
    const tokenParts = session.access_token.split(".");
    if (tokenParts.length === 3) {
      const payload = JSON.parse(atob(tokenParts[1]));
      const scopeString = payload.scp || payload.roles || "";
      scopes = scopeString.split(" ").filter((s) => s.length > 0);
    }
  } catch (e) {
    // Could not decode token
  }

  const toggleBtn = document.getElementById("toggleActiveScopes");
  const scopesDisplay = document.getElementById("activeScopesDisplay");

  if (toggleBtn && scopesDisplay) {
    // Check if scopes were already expanded
    const wasExpanded = scopesDisplay.style.display === "block";

    const newToggleBtn = toggleBtn.cloneNode(true);
    toggleBtn.parentNode.replaceChild(newToggleBtn, toggleBtn);

    const scopeCount = scopes.length;

    // Update scopes display content
    const updateScopesContent = () => {
      scopesDisplay.innerHTML = "";
      if (scopes.length > 0) {
        scopes.forEach((s) => {
          const scopeDiv = document.createElement("div");
          scopeDiv.className = "font-size-12 border-bottom";
          scopeDiv.style.padding = "3px 0";
          scopeDiv.textContent = s;
          scopesDisplay.appendChild(scopeDiv);
        });
      } else {
        const noScopesDiv = document.createElement("div");
        noScopesDiv.className = "color-text-secondary text-italic font-size-12";
        noScopesDiv.textContent = "No scopes found in token";
        scopesDisplay.appendChild(noScopesDiv);
      }
    };

    // If scopes were expanded, keep them expanded and update content
    if (wasExpanded) {
      scopesDisplay.style.display = "block";
      newToggleBtn.textContent = "Hide Scopes";
      updateScopesContent();
    } else {
      scopesDisplay.style.display = "none";
      newToggleBtn.textContent = `${scopeCount} Scope${scopeCount !== 1 ? "s" : ""}`;
    }

    newToggleBtn.addEventListener("click", () => {
      if (
        scopesDisplay.style.display === "none" ||
        scopesDisplay.style.display === ""
      ) {
        scopesDisplay.style.display = "block";
        newToggleBtn.textContent = "Hide Scopes";
        updateScopesContent();
      } else {
        scopesDisplay.style.display = "none";
        newToggleBtn.textContent = `${scopeCount} Scope${scopeCount !== 1 ? "s" : ""}`;
      }
    });
  }

  updateQuickActionButtons(scopes);

  // Update session status bar
  updateSessionStatusBar(session);

  // Start auto-refresh if globally enabled (but not when manually loading a session)
  if (!window._loadingSession) {
    loadAutoRefreshSetting();
  }
}

function updateSessionStatusBar(session) {
  const statusBar = document.getElementById("sessionStatusBar");
  const statusText = document.getElementById("sessionStatusText");

  if (!statusBar || !statusText) return;

  if (!session || !session.name || session.name === "Unknown") {
    statusBar.classList.add("hidden");
    return;
  }

  const expiryDate = new Date(session.expires_at);
  const now = new Date();
  const remainingMs = expiryDate - now;
  const remainingMinutes = Math.floor(remainingMs / 1000 / 60);
  const remainingHours = Math.floor(remainingMinutes / 60);

  statusText.textContent = `Active Session: ${session.name} (${session.user})`;
  statusBar.className = "session-status-bar";
  statusBar.classList.remove("hidden");
}

function updateQuickActionButtons(scopes) {
  const buttons = document.querySelectorAll(".quick-action-btn");

  buttons.forEach((btn) => {
    btn.disabled = false;
    btn.style.opacity = "1";
    btn.style.cursor = "pointer";
    btn.title = "";
    btn.classList.remove("scope-warning");
  });
}

async function saveTokens(tokenData) {
  try {
    const storageData = {
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token,
      expires_at: Date.now() + tokenData.expires_in * 1000,
      saved_at: Date.now(),
    };

    await chrome.storage.local.set({ [TOKEN_STORAGE_KEY]: storageData });
  } catch (error) {
    console.error("Failed to save tokens:", error);
  }
}

async function loadSavedTokens() {
  try {
    const result = await chrome.storage.local.get(TOKEN_STORAGE_KEY);
    const savedTokens = result[TOKEN_STORAGE_KEY];

    if (savedTokens && savedTokens.access_token && savedTokens.refresh_token) {
      m365TokenData = savedTokens;
      activeM365Session = savedTokens;

      // Display the active session in the UI
      displayActiveSession(activeM365Session);

      const now = Date.now();
      if (savedTokens.expires_at && savedTokens.expires_at < now) {
        const tokenError = document.getElementById("tokenError");
        if (tokenError) {
          tokenError.textContent =
            '⚠️ Loaded saved token is expired. Click "Refresh Token" to get a new one.';
          tokenError.style.display = "block";
          tokenError.style.background = "var(--warning-bg, #fff3cd)";
        }
      }
    }
  } catch (error) {
    console.error("Failed to load saved tokens:", error);
  }
}

async function clearActiveM365Session() {
  if (
    !confirm(
      "Are you sure you want to clear the active session?\n\nThis will remove the active tokens but saved sessions will remain.",
    )
  ) {
    return;
  }

  activeM365Session = null;
  stopAutoRefreshTimer();
  displayActiveSession(null);
  await chrome.storage.local.remove(TOKEN_STORAGE_KEY);

  saveUIState({ activeSessionIndex: null });

  // Hide session status bar
  const statusBar = document.getElementById("sessionStatusBar");
  if (statusBar) {
    statusBar.style.display = "none";
  }

  showToast("Active session cleared");

  // Clear mailbox if on mailbox tab
  const mailboxTab = document.getElementById("mailbox-tab");
  if (mailboxTab && mailboxTab.classList.contains("active")) {
    if (typeof showMailboxNoSession === "function") {
      showMailboxNoSession();
    }
  }
}

async function loadM365Sessions() {
  try {
    const result = await chrome.storage.local.get(SESSIONS_STORAGE_KEY);
    m365Sessions = result[SESSIONS_STORAGE_KEY] || [];
    renderM365Sessions();
  } catch (error) {
    console.error("Failed to load M365 sessions:", error);
  }
}

async function saveM365SessionToList(session) {
  try {
    m365Sessions.push(session);
    await chrome.storage.local.set({ [SESSIONS_STORAGE_KEY]: m365Sessions });
    renderM365Sessions();
    showToast("✅ Session saved");
  } catch (error) {
    console.error("Failed to save session:", error);
    showToast("Failed to save session: " + error.message, "error");
  }
}

function renderM365Sessions() {
  const container = document.getElementById("m365SessionsContainer");
  if (!container) return;

  if (m365Sessions.length === 0) {
    container.innerHTML =
      '<div class="empty-state-centered">No saved sessions</div>';
    return;
  }

  container.innerHTML = "";

  m365Sessions.forEach((session, index) => {
    const expiryDate = new Date(session.expires_at);
    const isExpired = expiryDate < new Date();
    const now = new Date();
    const remainingMs = expiryDate - now;
    const remainingMinutes = Math.floor(remainingMs / 1000 / 60);

    let scopes = [];
    try {
      const tokenParts = session.access_token.split(".");
      if (tokenParts.length === 3) {
        const payload = JSON.parse(atob(tokenParts[1]));
        const scopeString = payload.scp || payload.roles || "";
        scopes = scopeString.split(" ").filter((s) => s.length > 0);
      }
    } catch (e) {
      // Could not decode token for session
    }

    const scopeCount = scopes.length;

    // Create card
    const card = document.createElement("div");
    card.className = "session-card";
    card.style.background = "var(--card-bg)";
    card.style.padding = "12px";
    card.style.marginBottom = "8px";
    card.style.borderRadius = "6px";
    card.style.border = "1px solid var(--border-color)";
    card.style.transition = "all 0.2s";

    // Header with name and user
    const headerDiv = document.createElement("div");
    headerDiv.style.marginBottom = "8px";

    const nameStrong = document.createElement("strong");
    nameStrong.style.fontSize = "13px";
    nameStrong.style.color = "var(--text-color)";
    nameStrong.style.display = "block";
    nameStrong.style.marginBottom = "4px";
    nameStrong.style.wordBreak = "break-word";
    nameStrong.textContent = session.name;

    const userSmall = document.createElement("small");
    userSmall.style.color = "var(--text-secondary)";
    userSmall.style.fontSize = "11px";
    userSmall.textContent = session.user;

    headerDiv.appendChild(nameStrong);
    headerDiv.appendChild(userSmall);

    // Status
    const statusDiv = document.createElement("div");
    statusDiv.style.marginBottom = "10px";
    statusDiv.style.fontSize = "11px";

    const statusSpan = document.createElement("span");
    statusSpan.style.color = isExpired
      ? "var(--danger-color)"
      : "var(--text-color)";
    statusSpan.style.fontWeight = "500";
    statusSpan.textContent = isExpired
      ? "⚠️ Expired"
      : `✓ ${remainingMinutes}m left`;
    statusDiv.appendChild(statusSpan);

    // Scopes section
    const scopesSection = document.createElement("div");
    scopesSection.style.marginBottom = "10px";

    const scopesToggleBtn = document.createElement("button");
    scopesToggleBtn.className = "toggle-session-scopes-btn";
    scopesToggleBtn.setAttribute("data-index", index);
    scopesToggleBtn.style.background = "none";
    scopesToggleBtn.style.border = "none";
    scopesToggleBtn.style.padding = "0";
    scopesToggleBtn.style.color = "var(--text-color)";
    scopesToggleBtn.style.cursor = "pointer";
    scopesToggleBtn.style.fontSize = "11px";
    scopesToggleBtn.style.textDecoration = "underline";
    scopesToggleBtn.textContent = `${scopeCount} Scope${scopeCount !== 1 ? "s" : ""}`;

    const scopesDisplay = document.createElement("div");
    scopesDisplay.className = "session-scopes-display";
    scopesDisplay.setAttribute("data-index", index);
    scopesDisplay.style.display = "none";
    scopesDisplay.style.marginTop = "8px";
    scopesDisplay.style.padding = "8px";
    scopesDisplay.style.background = "var(--bg-secondary)";
    scopesDisplay.style.border = "1px solid var(--border-color)";
    scopesDisplay.style.borderRadius = "6px";
    scopesDisplay.style.maxHeight = "120px";
    scopesDisplay.style.overflowY = "auto";

    if (scopes.length > 0) {
      scopes.forEach((s) => {
        const scopeDiv = document.createElement("div");
        scopeDiv.style.padding = "3px 0";
        scopeDiv.style.borderBottom = "1px solid var(--border-color)";
        scopeDiv.style.fontSize = "12px";
        scopeDiv.textContent = s;
        scopesDisplay.appendChild(scopeDiv);
      });
    } else {
      const noScopesDiv = document.createElement("div");
      noScopesDiv.style.color = "var(--text-secondary)";
      noScopesDiv.style.fontStyle = "italic";
      noScopesDiv.style.fontSize = "12px";
      noScopesDiv.textContent = "No scopes found";
      scopesDisplay.appendChild(noScopesDiv);
    }

    scopesSection.appendChild(scopesToggleBtn);
    scopesSection.appendChild(scopesDisplay);

    // Action buttons
    const actionsDiv = document.createElement("div");
    actionsDiv.style.display = "flex";
    actionsDiv.style.flexWrap = "wrap";
    actionsDiv.style.gap = "6px";

    const loadBtn = document.createElement("button");
    loadBtn.className = "btn btn-primary btn-small load-session-btn";
    loadBtn.setAttribute("data-index", index);
    loadBtn.style.fontSize = "11px";
    loadBtn.style.padding = "6px 10px";
    loadBtn.textContent = "📥 Load";

    const editBtn = document.createElement("button");
    editBtn.className = "btn btn-secondary btn-small edit-session-btn";
    editBtn.setAttribute("data-index", index);
    editBtn.style.fontSize = "11px";
    editBtn.style.padding = "6px 10px";
    editBtn.textContent = "✏️ Edit";

    const exportBtn = document.createElement("button");
    exportBtn.className = "btn btn-secondary btn-small export-session-btn";
    exportBtn.setAttribute("data-index", index);
    exportBtn.style.fontSize = "11px";
    exportBtn.style.padding = "6px 10px";
    exportBtn.textContent = "📤 Export";

    const copyBtn = document.createElement("button");
    copyBtn.className = "btn btn-secondary btn-small copy-session-btn";
    copyBtn.setAttribute("data-index", index);
    copyBtn.style.fontSize = "11px";
    copyBtn.style.padding = "6px 10px";
    copyBtn.textContent = "📋 Copy";

    const deleteBtn = document.createElement("button");
    deleteBtn.className = "btn btn-danger-outline btn-small delete-session-btn";
    deleteBtn.setAttribute("data-index", index);
    deleteBtn.style.fontSize = "11px";
    deleteBtn.style.padding = "6px 10px";
    deleteBtn.textContent = "❌ Delete";

    actionsDiv.appendChild(loadBtn);
    actionsDiv.appendChild(editBtn);
    actionsDiv.appendChild(exportBtn);
    actionsDiv.appendChild(copyBtn);
    actionsDiv.appendChild(deleteBtn);

    // Assemble card
    card.appendChild(headerDiv);
    card.appendChild(statusDiv);
    card.appendChild(scopesSection);
    card.appendChild(actionsDiv);

    container.appendChild(card);
  });

  container.querySelectorAll(".load-session-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const index = parseInt(e.target.getAttribute("data-index"));
      loadM365Session(index);
    });
  });

  container.querySelectorAll(".edit-session-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const index = parseInt(e.target.getAttribute("data-index"));
      editM365Session(index);
    });
  });

  container.querySelectorAll(".export-session-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const index = parseInt(e.target.getAttribute("data-index"));
      exportSingleM365Session(index);
    });
  });

  container.querySelectorAll(".copy-session-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const index = parseInt(e.target.getAttribute("data-index"));
      copyM365Session(index);
    });
  });

  container.querySelectorAll(".delete-session-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const index = parseInt(e.target.getAttribute("data-index"));
      deleteM365Session(index);
    });
  });

  container.querySelectorAll(".toggle-session-scopes-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const index = parseInt(e.target.getAttribute("data-index"));
      const scopesDisplay = container.querySelector(
        `.session-scopes-display[data-index="${index}"]`,
      );

      if (scopesDisplay) {
        if (scopesDisplay.style.display === "none") {
          scopesDisplay.style.display = "block";
          btn.innerHTML = "Hide Scopes";
        } else {
          scopesDisplay.style.display = "none";
          const scopeCount = scopesDisplay.querySelectorAll("div").length;
          btn.innerHTML = `${scopeCount} Scope${scopeCount !== 1 ? "s" : ""}`;
        }
      }
    });
  });
}

async function loadM365Session(index) {
  activeM365Session = m365Sessions[index];
  m365TokenData = activeM365Session;

  await chrome.storage.local.set({ [TOKEN_STORAGE_KEY]: m365TokenData });

  // Prevent auto-refresh from starting when manually loading a session
  window._loadingSession = true;
  displayActiveSession(activeM365Session);
  window._loadingSession = false;

  if (!window._restoringSession) {
    showToast(`✓ Loaded session: ${activeM365Session.name}`);
  }

  saveUIState({ activeSessionIndex: index });

  // Initialize mailbox if on mailbox tab
  const mailboxTab = document.getElementById("mailbox-tab");
  if (mailboxTab && mailboxTab.classList.contains("active")) {
    if (typeof initializeMailbox === "function") {
      initializeMailbox();
    }
  }
}

async function copyM365Session(sessionIndex) {
  const originalSession = m365Sessions[sessionIndex];
  if (!originalSession) return;

  const copiedSession = {
    ...originalSession,
    name: `${originalSession.name} (Copy)`,
    created_at: new Date().toISOString(),
  };

  m365Sessions.push(copiedSession);
  await chrome.storage.local.set({
    [SESSIONS_STORAGE_KEY]: m365Sessions,
  });

  renderM365Sessions();
  showToast("Session copied successfully", "success");
}

async function deleteM365Session(sessionIndex) {
  if (confirm(`Delete session "${m365Sessions[sessionIndex].name}"?`)) {
    const sessionToDelete = m365Sessions[sessionIndex];

    // If deleting the active session, clear it
    if (
      activeM365Session &&
      activeM365Session.access_token === sessionToDelete.access_token
    ) {
      activeM365Session = null;
      m365TokenData = null;
      await chrome.storage.local.remove(TOKEN_STORAGE_KEY);
      displayActiveSession(null);
      saveUIState({ activeSessionIndex: -1 });
    }

    m365Sessions.splice(sessionIndex, 1);
    await chrome.storage.local.set({ [SESSIONS_STORAGE_KEY]: m365Sessions });
    renderM365Sessions();
    showToast("Session deleted");
  }
}

function editM365Session(index) {
  const session = m365Sessions[index];
  if (!session) {
    showToast("Session not found");
    return;
  }

  // Store the index for later use
  window._editingSessionIndex = index;

  document.getElementById("editM365SessionName").value = session.name || "";
  document.getElementById("editM365SessionUser").value = session.user || "";
  document.getElementById("editM365SessionClientId").value =
    session.client_id || "";
  document.getElementById("editM365SessionRedirectUri").value =
    session.redirect_uri || "";
  document.getElementById("editM365SessionScope").value = session.scope || "";
  document.getElementById("editM365SessionAuthUrl").value =
    session.auth_url || "";
  document.getElementById("editM365SessionTokenUrl").value =
    session.token_url || "";

  document.getElementById("editM365SessionModal").style.display = "flex";
}

function showEditM365SessionModal() {
  if (!activeM365Session) {
    showToast("No active session to edit");
    return;
  }

  // Clear the editing index to indicate we're editing the active session
  window._editingSessionIndex = null;

  document.getElementById("editM365SessionName").value =
    activeM365Session.name || "";
  document.getElementById("editM365SessionUser").value =
    activeM365Session.user || "";
  document.getElementById("editM365SessionClientId").value =
    activeM365Session.client_id || "";
  document.getElementById("editM365SessionRedirectUri").value =
    activeM365Session.redirect_uri || "";
  document.getElementById("editM365SessionScope").value =
    activeM365Session.scope || "";
  document.getElementById("editM365SessionAuthUrl").value =
    activeM365Session.auth_url || "";
  document.getElementById("editM365SessionTokenUrl").value =
    activeM365Session.token_url || "";

  document.getElementById("editM365SessionModal").style.display = "flex";
}

function closeEditM365SessionModal() {
  document.getElementById("editM365SessionModal").style.display = "none";
}

async function confirmEditM365Session() {
  const name = document.getElementById("editM365SessionName").value.trim();
  const user = document.getElementById("editM365SessionUser").value.trim();
  const clientId = document
    .getElementById("editM365SessionClientId")
    .value.trim();
  const redirectUri = document
    .getElementById("editM365SessionRedirectUri")
    .value.trim();
  const scope = document.getElementById("editM365SessionScope").value.trim();
  const authUrl = document
    .getElementById("editM365SessionAuthUrl")
    .value.trim();
  const tokenUrl = document
    .getElementById("editM365SessionTokenUrl")
    .value.trim();

  if (!name) {
    showToast("Please enter a session name");
    return;
  }

  // Check if we're editing a saved session or the active session
  if (
    window._editingSessionIndex !== null &&
    window._editingSessionIndex !== undefined
  ) {
    // Editing a saved session
    const sessionIndex = window._editingSessionIndex;
    const session = m365Sessions[sessionIndex];

    if (session) {
      session.name = name;
      session.user = user || session.user;
      session.client_id = clientId || session.client_id;
      session.redirect_uri = redirectUri || session.redirect_uri;
      session.scope = scope || session.scope;
      session.auth_url = authUrl || session.auth_url;
      session.token_url = tokenUrl || session.token_url;

      m365Sessions[sessionIndex] = session;
      await chrome.storage.local.set({ [SESSIONS_STORAGE_KEY]: m365Sessions });

      // If this is also the active session, update it
      if (
        activeM365Session &&
        activeM365Session.created_at === session.created_at
      ) {
        activeM365Session = session;
        await chrome.storage.local.set({
          [TOKEN_STORAGE_KEY]: activeM365Session,
        });
        displayActiveSession(activeM365Session);
      }

      await loadM365Sessions();
      closeEditM365SessionModal();
      showToast("✅ Session updated");
      window._editingSessionIndex = null;
    }
  } else {
    // Editing the active session
    if (!activeM365Session) {
      showToast("No active session to edit");
      return;
    }

    // Find session index before updating user field
    const sessionIndex = m365Sessions.findIndex(
      (s) =>
        s.user === activeM365Session.user &&
        s.created_at === activeM365Session.created_at,
    );

    // Update session fields
    activeM365Session.name = name;
    activeM365Session.user = user || activeM365Session.user;
    activeM365Session.client_id = clientId || activeM365Session.client_id;
    activeM365Session.redirect_uri =
      redirectUri || activeM365Session.redirect_uri;
    activeM365Session.scope = scope || activeM365Session.scope;
    activeM365Session.auth_url = authUrl || activeM365Session.auth_url;
    activeM365Session.token_url = tokenUrl || activeM365Session.token_url;

    if (sessionIndex >= 0) {
      m365Sessions[sessionIndex] = activeM365Session;
      await chrome.storage.local.set({ [SESSIONS_STORAGE_KEY]: m365Sessions });
    } else {
      await saveM365SessionToList(activeM365Session);
    }

    await chrome.storage.local.set({ [TOKEN_STORAGE_KEY]: activeM365Session });
    await loadM365Sessions();
    displayActiveSession(activeM365Session);
    closeEditM365SessionModal();
    showToast("✅ Session updated");
  }
}

function showImportM365SessionModal() {
  document.getElementById("importM365SessionModal").style.display = "flex";
}

function closeImportM365SessionModal() {
  document.getElementById("importM365SessionModal").style.display = "none";
  document.getElementById("importM365SessionData").value = "";
}

function handleM365SessionFileSelect(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (event) => {
    document.getElementById("importM365SessionData").value =
      event.target.result;
  };
  reader.readAsText(file);
}

async function confirmImportM365Session() {
  const data = document.getElementById("importM365SessionData").value.trim();
  if (!data) {
    showToast("Please paste session data");
    return;
  }

  try {
    const sessions = JSON.parse(data);
    const sessionArray = Array.isArray(sessions) ? sessions : [sessions];

    let imported = 0;
    let skipped = 0;
    let fetched = 0;
    const errors = [];

    for (let i = 0; i < sessionArray.length; i++) {
      const session = sessionArray[i];
      const missing = [];

      if (!session.refresh_token) {
        missing.push("refresh_token");
      }

      if (missing.length > 0) {
        skipped++;
        errors.push(`Session ${i + 1}: Missing ${missing.join(", ")}`);
        continue;
      }

      // If access_token is missing, try to fetch it using refresh_token
      if (!session.access_token) {
        try {
          const clientId =
            session.client_id || "1b730954-1685-4b74-9bfd-dac224a7b894";
          const tokenUrl =
            session.token_url ||
            "https://login.microsoftonline.com/common/oauth2/v2.0/token";

          const params = {
            client_id: clientId,
            refresh_token: session.refresh_token,
            grant_type: "refresh_token",
          };

          if (session.client_secret) {
            params.client_secret = session.client_secret;
          }

          if (session.scope) {
            params.scope = session.scope;
          }

          const response = await chrome.runtime.sendMessage({
            action: "fetchWithoutOrigin",
            url: tokenUrl,
            method: "POST",
            headers: {
              "Content-Type": "application/x-www-form-urlencoded",
            },
            body: new URLSearchParams(params).toString(),
          });

          if (!response.ok) {
            const errorData = JSON.parse(response.body);
            const errorMsg =
              errorData.error_description ||
              errorData.error ||
              "Failed to get access token";
            throw new Error(errorMsg);
          }

          const tokenData = JSON.parse(response.body);
          session.access_token = tokenData.access_token;
          if (tokenData.refresh_token) {
            session.refresh_token = tokenData.refresh_token;
          }
          session.expires_at =
            Date.now() + (tokenData.expires_in || 3600) * 1000;
          fetched++;
        } catch (error) {
          skipped++;
          errors.push(
            `Session ${i + 1}: Failed to fetch access token - ${error.message}`,
          );
          continue;
        }
      }

      // Add default values for missing fields when importing old sessions
      if (!session.created_at) {
        session.created_at = Date.now();
      }
      if (!session.name) {
        const randomStr = Math.random().toString(36).substring(2, 10);
        session.name = `Session-${randomStr}`;
      }
      if (!session.user) {
        try {
          const tokenParts = session.access_token.split(".");
          if (tokenParts.length === 3) {
            const payload = JSON.parse(atob(tokenParts[1]));
            session.user =
              payload.upn ||
              payload.unique_name ||
              payload.email ||
              "Unknown User";
          } else {
            session.user = "Unknown User";
          }
        } catch (e) {
          session.user = "Unknown User";
        }
      }
      if (!session.client_id) {
        session.client_id = "1b730954-1685-4b74-9bfd-dac224a7b894";
      }
      if (!session.scope) {
        session.scope = "https://graph.microsoft.com/.default offline_access";
      }

      if (!session.token_url) {
        session.token_url =
          "https://login.microsoftonline.com/common/oauth2/v2.0/token";
      }
      if (!session.redirect_uri) {
        session.redirect_uri =
          "https://login.microsoftonline.com/common/oauth2/nativeclient";
      }
      if (!session.expires_at && session.access_token) {
        session.expires_at = Date.now() + 3600 * 1000;
      }

      await saveM365SessionToList(session);
      imported++;
    }

    closeImportM365SessionModal();

    if (imported === 0 && skipped > 0) {
      const errorMsg = errors.length > 0 ? `\n${errors[0]}` : "";
      showToast(
        `❌ No sessions imported. ${skipped} skipped.${errorMsg}`,
        "error",
      );
    } else if (imported > 0 && skipped > 0) {
      let message = `⚠️ Imported ${imported} session(s), skipped ${skipped}`;
      if (fetched > 0) {
        message += ` (${fetched} fetched new tokens)`;
      }
      showToast(message);
    } else if (imported > 0) {
      let message = `✅ Imported ${imported} session(s)`;
      if (fetched > 0) {
        message += ` (${fetched} fetched new tokens)`;
      }
      showToast(message);
    } else {
      showToast("No valid sessions found to import", "error");
    }
  } catch (error) {
    console.error("Import error:", error);
    showToast("Failed to import: " + error.message, "error");
  }
}

async function exportM365Sessions() {
  if (m365Sessions.length === 0) {
    showToast("No sessions to export");
    return;
  }

  const json = JSON.stringify(m365Sessions, null, 2);
  const filename = `sushi_sessions_${new Date().toISOString().split("T")[0]}.json`;
  downloadFile(json, filename, "application/json");
  showToast(`✅ Exported ${m365Sessions.length} session(s)`);
}

async function exportSingleM365Session(index) {
  if (index < 0 || index >= m365Sessions.length) {
    showToast("❌ Invalid session");
    return;
  }

  const session = m365Sessions[index];
  const json = JSON.stringify([session], null, 2);
  const safeName = session.name.replace(/[^a-z0-9]/gi, "_").toLowerCase();
  const filename = `sushi_session_${safeName}_${new Date().toISOString().split("T")[0]}.json`;
  downloadFile(json, filename, "application/json");
  showToast(`✓ Exported session: ${session.name}`);
}

function showClearAllM365Modal() {
  const totalSessions = m365Sessions.length;
  const hasActiveSession = activeM365Session !== null;

  if (totalSessions === 0 && !hasActiveSession) {
    showToast("No M365 sessions to clear", "warning");
    return;
  }

  const modal = document.getElementById("clearAllM365Modal");
  const confirmText = document.getElementById("clearAllM365ConfirmText");
  const confirmBtn = document.getElementById("confirmClearAllM365");

  confirmText.value = "";
  confirmBtn.disabled = true;

  const warningText = modal.querySelector("p");
  let message = "This will <strong>permanently delete ";

  if (hasActiveSession && totalSessions > 0) {
    message += `the active session and all ${totalSessions} saved session(s)</strong>`;
  } else if (hasActiveSession) {
    message += "the active session</strong>";
  } else {
    message += `all ${totalSessions} saved session(s)</strong>`;
  }

  message += ". This action cannot be undone.";
  warningText.innerHTML = message;

  modal.style.display = "flex";

  setTimeout(() => confirmText.focus(), 100);
}

function closeClearAllM365Modal() {
  const modal = document.getElementById("clearAllM365Modal");
  const confirmText = document.getElementById("clearAllM365ConfirmText");
  const confirmBtn = document.getElementById("confirmClearAllM365");

  modal.style.display = "none";
  confirmText.value = "";
  confirmBtn.disabled = true;
}

// Auto-refresh functionality
function startAutoRefreshTimer() {
  stopAutoRefreshTimer();

  autoRefreshTimer = setInterval(() => {
    checkAndAutoRefresh();
    // Update UI display to show current expiration status for active session only
    if (activeM365Session) {
      updateSessionStatusBar(activeM365Session);
    }
  }, 30000); // Check every 30 seconds

  // Also check immediately
  checkAndAutoRefresh();
}

function stopAutoRefreshTimer() {
  if (autoRefreshTimer) {
    clearInterval(autoRefreshTimer);
    autoRefreshTimer = null;
  }
}

async function checkAndAutoRefresh() {
  if (isRefreshing) {
    return;
  }

  // Don't auto-refresh if we're in the middle of loading a session
  if (window._loadingSession) {
    return;
  }

  // Check if auto-refresh is globally enabled
  const result = await chrome.storage.local.get([AUTO_REFRESH_STORAGE_KEY]);
  const autoRefreshEnabled = result[AUTO_REFRESH_STORAGE_KEY] === true;

  if (!autoRefreshEnabled) {
    return;
  }

  const now = new Date();
  let refreshedAny = false;
  let hasActiveSession = false;

  // Check all saved sessions
  for (let i = 0; i < m365Sessions.length; i++) {
    const session = m365Sessions[i];

    if (!session || !session.refresh_token) {
      continue;
    }

    const expiryDate = new Date(session.expires_at);
    const remainingMs = expiryDate - now;
    const remainingMinutes = Math.floor(remainingMs / 1000 / 60);

    const isActiveSession =
      activeM365Session &&
      session.name === activeM365Session.name &&
      session.user === activeM365Session.user;

    if (isActiveSession) {
      hasActiveSession = true;
    }

    // Auto-refresh if less than 5 minutes remaining OR already expired
    // The refresh token can get a new access token regardless of how long the access token has been expired
    if (remainingMinutes < 5) {
      isRefreshing = true;
      try {
        const clientId =
          session.client_id || "1b730954-1685-4b74-9bfd-dac224a7b894";
        const scope =
          session.scope ||
          "https://graph.microsoft.com/.default offline_access";

        const tokenUrl =
          session.token_url ||
          "https://login.microsoftonline.com/common/oauth2/v2.0/token";

        const response = await chrome.runtime.sendMessage({
          action: "fetchWithoutOrigin",
          url: tokenUrl,
          method: "POST",
          headers: {
            "Content-Type": "application/x-www-form-urlencoded",
          },
          body: new URLSearchParams({
            client_id: clientId,
            scope: scope,
            refresh_token: session.refresh_token,
            grant_type: "refresh_token",
          }).toString(),
        });

        if (!response.ok) {
          const errorData = JSON.parse(response.body);
          throw new Error(
            errorData.error_description || "Failed to refresh token",
          );
        }

        const tokenData = JSON.parse(response.body);

        // Update the session with new tokens
        session.access_token = tokenData.access_token;
        session.refresh_token = tokenData.refresh_token;
        session.expires_at = Date.now() + tokenData.expires_in * 1000;

        // Update in saved sessions array
        m365Sessions[i] = session;

        // If this is the active session, update it too
        if (isActiveSession) {
          activeM365Session = session;
          await chrome.storage.local.set({
            [TOKEN_STORAGE_KEY]: activeM365Session,
          });
          displayActiveSession(activeM365Session);
          updateSessionStatusBar(activeM365Session);
        }

        refreshedAny = true;
      } catch (error) {
        console.error(
          `[Auto-refresh] Failed to refresh session "${session.name}":`,
          error,
        );
        // Continue trying other sessions even if one fails
      } finally {
        isRefreshing = false;
      }
    }
  }

  // Save all updated sessions if any were refreshed
  if (refreshedAny) {
    await chrome.storage.local.set({ [SESSIONS_STORAGE_KEY]: m365Sessions });
    renderM365Sessions();

    // Count how many sessions were refreshed
    const refreshedCount = m365Sessions.filter((s) => {
      const expiryDate = new Date(s.expires_at);
      const remainingMs = expiryDate - now;
      const remainingMinutes = Math.floor(remainingMs / 1000 / 60);
      return remainingMinutes >= 55; // Recently refreshed (tokens usually last ~60 minutes)
    }).length;

    if (refreshedCount > 1) {
      showToast(`✅ Auto-refreshed ${refreshedCount} sessions`);
    } else if (refreshedCount === 1) {
      showToast(`✅ Auto-refreshed session`);
    }
  }

  // If no sessions were found at all, stop the timer
  if (m365Sessions.length === 0 && !hasActiveSession) {
    stopAutoRefreshTimer();
  }
}

async function toggleAutoRefresh() {
  const autoRefreshCheckbox = document.getElementById("autoRefreshCheckbox");

  if (!autoRefreshCheckbox) {
    return;
  }

  const isEnabled = autoRefreshCheckbox.checked;

  // Save the global auto-refresh setting
  await chrome.storage.local.set({
    [AUTO_REFRESH_STORAGE_KEY]: isEnabled,
  });

  if (isEnabled) {
    showToast("✅ Auto-refresh enabled for all sessions");
    if (activeM365Session || m365Sessions.length > 0) {
      startAutoRefreshTimer();
    }
  } else {
    showToast("Auto-refresh disabled");
    stopAutoRefreshTimer();
  }
}

async function loadAutoRefreshSetting() {
  const autoRefreshCheckbox = document.getElementById("autoRefreshCheckbox");
  if (!autoRefreshCheckbox) {
    return;
  }

  const result = await chrome.storage.local.get([AUTO_REFRESH_STORAGE_KEY]);
  const autoRefreshEnabled = result[AUTO_REFRESH_STORAGE_KEY] === true;

  autoRefreshCheckbox.checked = autoRefreshEnabled;

  // Start timer if enabled and there are any sessions (active or saved)
  if (autoRefreshEnabled && (activeM365Session || m365Sessions.length > 0)) {
    startAutoRefreshTimer();
  } else if (autoRefreshEnabled) {
  }
}

async function switchActiveSessionClientId(newClientId) {
  try {
    const session = activeM365Session;
    if (!session) {
      showToast("No active session to switch client ID", "error");
      return;
    }

    session.client_id = newClientId;

    await chrome.storage.local.set({ [TOKEN_STORAGE_KEY]: session });

    const sessionIndex = m365Sessions.findIndex(
      (s) => s.user === session.user && s.created_at === session.created_at,
    );

    if (sessionIndex >= 0) {
      m365Sessions[sessionIndex] = session;
      await chrome.storage.local.set({ [SESSIONS_STORAGE_KEY]: m365Sessions });
      renderM365Sessions();
    }

    displayActiveSession(session);

    const clientName = FOCI_CLIENTS[newClientId] || "Custom Client";
    showToast(`Client ID switched to ${clientName}`);
  } catch (error) {
    console.error("Error switching client ID:", error);
    showToast("Failed to switch client ID: " + error.message, "error");
    displayActiveSession(activeM365Session);
  }
}

async function clearAllM365Sessions() {
  try {
    showToast("Clearing M365 sessions...", "info");
    closeClearAllM365Modal();

    // Clear active session
    if (activeM365Session !== null) {
      activeM365Session = null;
      stopAutoRefreshTimer();
      displayActiveSession(null);
      await chrome.storage.local.remove(TOKEN_STORAGE_KEY);

      // Hide session status bar
      const statusBar = document.getElementById("sessionStatusBar");
      if (statusBar) {
        statusBar.style.display = "none";
      }
    }

    // Clear all saved sessions
    const sessionCount = m365Sessions.length;
    m365Sessions = [];
    await chrome.storage.local.remove(SESSIONS_STORAGE_KEY);
    renderM365Sessions();

    // Clear UI state
    saveUIState({ activeSessionIndex: null });

    // Clear mailbox if on mailbox tab
    const mailboxTab = document.getElementById("mailbox-tab");
    if (mailboxTab && mailboxTab.classList.contains("active")) {
      if (typeof showMailboxNoSession === "function") {
        showMailboxNoSession();
      }
    }

    showToast(
      `Cleared ${sessionCount} session(s) and active session`,
      "success",
    );
  } catch (error) {
    console.error("Error clearing M365 sessions:", error);
    showToast("Failed to clear M365 sessions", "error");
  }
}
