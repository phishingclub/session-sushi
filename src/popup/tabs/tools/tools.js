let _toolsLastTokenData = null;

async function exchangeAuthCode() {
    const code = document.getElementById("toolsAuthCodeInput").value.trim();
    const tokenUrl = document.getElementById("toolsTokenUrlInput").value.trim();
    const clientId = document.getElementById("toolsClientIdInput").value.trim();
    const redirectUri = document.getElementById("toolsRedirectUriInput").value.trim();
    const scope = document.getElementById("toolsScopeInput").value.trim();
    const codeVerifier = document.getElementById("toolsCodeVerifierInput").value.trim();
    const clientSecret = document.getElementById("toolsClientSecretInput").value.trim();

    if (!code) {
        showToast("Authorization code is required", "error");
        return;
    }
    if (!tokenUrl) {
        showToast("Token URL is required", "error");
        return;
    }
    if (!clientId) {
        showToast("Client ID is required", "error");
        return;
    }
    if (!redirectUri) {
        showToast("Redirect URI is required", "error");
        return;
    }

    const exchangeBtn = document.getElementById("toolsExchangeBtn");
    const statusEl = document.getElementById("toolsStatusMsg");
    const responseArea = document.getElementById("toolsResponseArea");
    const copyBtn = document.getElementById("toolsCopyResponseBtn");
    const importBtn = document.getElementById("toolsImportSessionBtn");

    exchangeBtn.disabled = true;
    exchangeBtn.textContent = "⏳ Exchanging...";
    statusEl.textContent = "";
    statusEl.style.color = "";
    responseArea.value = "";
    copyBtn.disabled = true;
    importBtn.disabled = true;
    _toolsLastTokenData = null;

    const params = {
        client_id: clientId,
        redirect_uri: redirectUri,
        code: code,
        grant_type: "authorization_code",
    };

    if (scope) params.scope = scope;
    if (codeVerifier) params.code_verifier = codeVerifier;
    if (clientSecret) params.client_secret = clientSecret;

    try {
        const response = await chrome.runtime.sendMessage({
            action: "fetchWithoutOrigin",
            url: tokenUrl,
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: new URLSearchParams(params).toString(),
        });

        if (!response) {
            throw new Error("No response received from background");
        }

        responseArea.value = toolsPrettyPrint(response.body || "");
        copyBtn.disabled = false;

        if (response.ok) {
            statusEl.textContent = "✅ Success (" + response.status + ")";
            statusEl.style.color = "var(--pc-green)";

            try {
                const data = JSON.parse(response.body);
                if (data.access_token || data.refresh_token) {
                    _toolsLastTokenData = data;
                    importBtn.disabled = false;
                }
            } catch (e) {
                // Response is not JSON or has no tokens
            }
        } else {
            statusEl.textContent = "❌ Error (" + response.status + " " + response.statusText + ")";
            statusEl.style.color = "var(--danger-color)";
        }
    } catch (error) {
        statusEl.textContent = "❌ " + error.message;
        statusEl.style.color = "var(--danger-color)";
        responseArea.value = error.message;
        copyBtn.disabled = false;
    } finally {
        exchangeBtn.disabled = false;
        exchangeBtn.textContent = "🔑 Exchange Code";
    }
}

function toolsPrettyPrint(text) {
    try {
        return JSON.stringify(JSON.parse(text), null, 2);
    } catch (e) {
        return text;
    }
}

async function importToolsTokenAsSession() {
    if (!_toolsLastTokenData) return;

    const clientId = document.getElementById("toolsClientIdInput").value.trim();
    const redirectUri = document.getElementById("toolsRedirectUriInput").value.trim();
    const scope = document.getElementById("toolsScopeInput").value.trim();
    const tokenUrl = document.getElementById("toolsTokenUrlInput").value.trim();

    const tokenData = _toolsLastTokenData;

    let userEmail = "Unknown";
    if (tokenData.access_token) {
        try {
            const parts = tokenData.access_token.split(".");
            if (parts.length === 3) {
                const payload = JSON.parse(atob(parts[1]));
                userEmail = payload.upn || payload.unique_name || payload.email || "Unknown";
            }
        } catch (e) {
            // Could not decode token claims
        }
    }

    const session = {
        name: userEmail,
        user: userEmail,
        access_token: tokenData.access_token || "",
        refresh_token: tokenData.refresh_token || "",
        expires_at: Date.now() + (tokenData.expires_in || 3600) * 1000,
        created_at: Date.now(),
        client_id: clientId,
        redirect_uri: redirectUri,
        scope: scope || tokenData.scope || "",
        token_url: tokenUrl,
        auth_url: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    };

    try {
        await saveM365SessionToList(session);
        document.getElementById("toolsImportSessionBtn").disabled = true;
        showToast("✅ Session imported successfully");
    } catch (error) {
        showToast("Failed to import session: " + error.message, "error");
    }
}

function setupToolsListeners() {
    const exchangeBtn = document.getElementById("toolsExchangeBtn");
    if (exchangeBtn) {
        exchangeBtn.addEventListener("click", exchangeAuthCode);
    }

    const copyBtn = document.getElementById("toolsCopyResponseBtn");
    if (copyBtn) {
        copyBtn.addEventListener("click", async () => {
            const responseArea = document.getElementById("toolsResponseArea");
            if (responseArea && responseArea.value) {
                await copyToClipboard(responseArea.value);
            }
        });
    }

    const importBtn = document.getElementById("toolsImportSessionBtn");
    if (importBtn) {
        importBtn.addEventListener("click", importToolsTokenAsSession);
    }
}
