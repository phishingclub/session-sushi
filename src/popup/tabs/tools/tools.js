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

    let resolvedScope = scope;
    if (!resolvedScope && tokenData.access_token) {
        try {
            const payload = JSON.parse(atob(tokenData.access_token.split(".")[1]));
            const aud = payload.aud;
            if (aud && aud !== "00000003-0000-0000-c000-000000000000") {
                resolvedScope = `${aud}/.default offline_access`;
            } else {
                resolvedScope = "https://graph.microsoft.com/.default offline_access";
            }
        } catch (e) {
            resolvedScope = "https://graph.microsoft.com/.default offline_access";
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
        scope: resolvedScope || "https://graph.microsoft.com/.default offline_access",
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

function decodeToken() {
    const input = document.getElementById("toolsDecodeInput").value.trim();
    const statusEl = document.getElementById("toolsDecodeStatus");
    const outputEl = document.getElementById("toolsDecodeOutput");
    const copyBtn = document.getElementById("toolsDecodeCopyBtn");

    statusEl.textContent = "";
    outputEl.value = "";
    copyBtn.disabled = true;

    if (!input) {
        statusEl.textContent = "❌ No token provided";
        statusEl.style.color = "var(--danger-color)";
        return;
    }

    const parts = input.split(".");

    if (parts.length !== 3) {
        statusEl.textContent = "⚠️ Opaque token — cannot be decoded (refresh tokens are not JWTs)";
        statusEl.style.color = "var(--pc-yellow, #f0a500)";
        return;
    }

    try {
        const decode = (part) => JSON.parse(atob(part.replace(/-/g, "+").replace(/_/g, "/")));
        const header = decode(parts[0]);
        const payload = decode(parts[1]);

        const exp = payload.exp ? new Date(payload.exp * 1000).toISOString() : null;
        const iat = payload.iat ? new Date(payload.iat * 1000).toISOString() : null;
        const nbf = payload.nbf ? new Date(payload.nbf * 1000).toISOString() : null;
        const now = Math.floor(Date.now() / 1000);
        const expired = payload.exp && payload.exp < now;

        const summary = [];
        if (payload.upn || payload.unique_name) summary.push(`upn:      ${payload.upn || payload.unique_name}`);
        if (payload.oid)   summary.push(`oid:      ${payload.oid}`);
        if (payload.tid)   summary.push(`tid:      ${payload.tid}`);
        if (payload.aud)   summary.push(`aud:      ${payload.aud}`);
        if (payload.appid) summary.push(`appid:    ${payload.appid}`);
        if (payload.scp)   summary.push(`scp:      ${payload.scp}`);
        if (payload.roles) summary.push(`roles:    ${Array.isArray(payload.roles) ? payload.roles.join(" ") : payload.roles}`);
        if (payload.foci)  summary.push(`foci:     ${payload.foci}`);
        if (exp)           summary.push(`exp:      ${exp}${expired ? " ⚠️ EXPIRED" : ""}`);
        if (iat)           summary.push(`iat:      ${iat}`);
        if (nbf)           summary.push(`nbf:      ${nbf}`);

        const output = [
            "=== Summary ===",
            ...summary,
            "",
            "=== Header ===",
            JSON.stringify(header, null, 2),
            "",
            "=== Payload ===",
            JSON.stringify(payload, null, 2),
        ].join("\n");

        outputEl.value = output;
        copyBtn.disabled = false;
        statusEl.textContent = expired ? "⚠️ Token decoded (expired)" : "✅ Token decoded";
        statusEl.style.color = expired ? "var(--pc-yellow, #f0a500)" : "var(--pc-green)";
    } catch (e) {
        statusEl.textContent = "⚠️ Opaque token — cannot be decoded (not a valid JWT)";
        statusEl.style.color = "var(--pc-yellow, #f0a500)";
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

    const decodeBtn = document.getElementById("toolsDecodeBtn");
    if (decodeBtn) {
        decodeBtn.addEventListener("click", decodeToken);
    }

    const decodeInput = document.getElementById("toolsDecodeInput");
    if (decodeInput) {
        decodeInput.addEventListener("keydown", (e) => {
            if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                decodeToken();
            }
        });
    }

    const loadAccessBtn = document.getElementById("toolsDecodeLoadAccessBtn");
    if (loadAccessBtn) {
        loadAccessBtn.addEventListener("click", async () => {
            const result = await chrome.storage.local.get(["m365_tokens"]);
            const session = result["m365_tokens"];
            if (session && session.access_token) {
                document.getElementById("toolsDecodeInput").value = session.access_token;
                decodeToken();
            } else {
                showToast("No active session access token found", "error");
            }
        });
    }


    const decodeCopyBtn = document.getElementById("toolsDecodeCopyBtn");
    if (decodeCopyBtn) {
        decodeCopyBtn.addEventListener("click", async () => {
            const outputEl = document.getElementById("toolsDecodeOutput");
            if (outputEl && outputEl.value) {
                await copyToClipboard(outputEl.value);
            }
        });
    }

    const clearDecodeBtn = document.getElementById("toolsClearDecodeBtn");
    if (clearDecodeBtn) {
        clearDecodeBtn.addEventListener("click", () => {
            document.getElementById("toolsDecodeInput").value = "";
            document.getElementById("toolsDecodeOutput").value = "";
            document.getElementById("toolsDecodeStatus").textContent = "";
            document.getElementById("toolsDecodeCopyBtn").disabled = true;
        });
    }
}
