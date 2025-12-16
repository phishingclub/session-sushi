let pendingAuthFlow = null;
let extensionWindowId = null;

chrome.windows.onRemoved.addListener((windowId) => {
  if (windowId === extensionWindowId) {
    extensionWindowId = null;
  }
});

chrome.action.onClicked.addListener(async () => {
  if (extensionWindowId) {
    try {
      await chrome.windows.update(extensionWindowId, { focused: true });
      return;
    } catch (error) {
      extensionWindowId = null;
    }
  }

  const window = await chrome.windows.create({
    url: chrome.runtime.getURL("src/popup.html"),
    type: "normal",
    state: "maximized",
  });

  extensionWindowId = window.id;
});

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (!sender.id || sender.id !== chrome.runtime.id) {
    console.warn("Rejected message from unauthorized sender:", sender);
    sendResponse({ success: false, error: "Unauthorized" });
    return false;
  }

  switch (request.action) {
    case "getAllCookies":
      chrome.cookies.getAll({}, (cookies) => {
        sendResponse({ cookies });
      });
      return true;

    case "importCookies":
      importCookies(request.cookies).then((result) => {
        sendResponse(result);
      });
      return true;

    case "clearAllCookies":
      clearAllCookies().then((result) => {
        sendResponse(result);
      });
      return true;

    case "getGraphToken":
      pendingAuthFlow = {
        clientId: request.clientId,
        redirectUri: request.redirectUri,
        scope: request.scope,
        authUrl: request.authUrl,
        tokenUrl: request.tokenUrl,
        timestamp: Date.now(),
      };

      handleGraphTokenFlow(
        request.clientId,
        request.redirectUri,
        request.scope,
        request.authUrl,
        request.tokenUrl,
      )
        .then((result) => {
          pendingAuthFlow = null;
          sendResponse(result);
        })
        .catch((error) => {
          pendingAuthFlow = null;
          sendResponse({ success: false, error: error.message });
        });
      return true;

    case "checkPendingAuth":
      sendResponse({
        pending: pendingAuthFlow !== null,
        info: pendingAuthFlow,
      });
      return true;

    default:
      return true;
  }
});

async function importCookies(cookies) {
  let imported = 0;
  let failed = 0;

  for (const cookie of cookies) {
    try {
      // Normalize sameSite value for Chrome API
      const validSameSiteValues = [
        "no_restriction",
        "lax",
        "strict",
        "unspecified",
      ];
      const sameSite = validSameSiteValues.includes(cookie.sameSite)
        ? cookie.sameSite
        : "lax";

      // SameSite=None requires Secure flag to be true
      let secure = cookie.secure || false;
      if (sameSite === "no_restriction" && !secure) {
        secure = true;
      }

      const cookieData = {
        name: cookie.name,
        value: cookie.value,
        domain: cookie.domain,
        path: cookie.path || "/",
        secure: secure,
        httpOnly: cookie.httpOnly || false,
        sameSite: sameSite,
      };

      const protocol = secure ? "https" : "http";
      const domain = cookie.domain.startsWith(".")
        ? cookie.domain.substring(1)
        : cookie.domain;
      cookieData.url = `${protocol}://${domain}${cookie.path || "/"}`;

      if (cookie.expirationDate) {
        cookieData.expirationDate = cookie.expirationDate;
      }

      await chrome.cookies.set(cookieData);
      imported++;
    } catch (error) {
      console.error("Failed to import cookie:", cookie.name, error);
      failed++;
    }
  }

  return { imported, failed, total: cookies.length };
}

async function clearAllCookies() {
  const cookies = await chrome.cookies.getAll({});
  let removed = 0;

  for (const cookie of cookies) {
    try {
      const protocol = cookie.secure ? "https" : "http";
      const url = `${protocol}://${cookie.domain}${cookie.path}`;
      await chrome.cookies.remove({ url, name: cookie.name });
      removed++;
    } catch (error) {
      console.error("Failed to remove cookie:", cookie.name, error);
    }
  }

  return { removed, total: cookies.length };
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
  const base64 = btoa(String.fromCharCode.apply(null, array));
  return base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
}

async function handleGraphTokenFlow(
  clientId,
  redirectUri,
  scope,
  authUrl,
  tokenUrl,
) {
  return new Promise(async (resolve, reject) => {
    let authWindowId = null;
    let authTabId = null;
    let cleanedUp = false;
    let windowClosedListener = null;
    let webRequestRedirectListener = null;

    // Generate PKCE parameters
    const codeVerifier = generateCodeVerifier();
    const codeChallenge = await generateCodeChallenge(codeVerifier);

    const cleanup = (windowId) => {
      if (cleanedUp) return;
      cleanedUp = true;

      if (webRequestRedirectListener) {
        chrome.webRequest.onBeforeRedirect.removeListener(
          webRequestRedirectListener,
        );
      }
      if (windowClosedListener) {
        chrome.windows.onRemoved.removeListener(windowClosedListener);
      }

      if (windowId) {
        chrome.windows.get(windowId, (win) => {
          if (!chrome.runtime.lastError && win) {
            chrome.windows.remove(windowId);
          }
        });
      }
    };

    const authUrlObj = new URL(
      authUrl ||
        "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    );
    authUrlObj.searchParams.set("client_id", clientId);
    authUrlObj.searchParams.set("response_type", "code");
    authUrlObj.searchParams.set("redirect_uri", redirectUri);
    authUrlObj.searchParams.set("scope", scope);
    authUrlObj.searchParams.set("response_mode", "query");
    authUrlObj.searchParams.set("code_challenge", codeChallenge);
    authUrlObj.searchParams.set("code_challenge_method", "S256");

    chrome.windows.create(
      {
        url: authUrlObj.toString(),
        type: "popup",
        width: 500,
        height: 600,
      },
      (window) => {
        if (!window) {
          reject(new Error("Failed to create auth window"));
          return;
        }

        authWindowId = window.id;
        if (window.tabs && window.tabs.length > 0) {
          authTabId = window.tabs[0].id;
        }

        webRequestRedirectListener = (details) => {
          if (authTabId && details.tabId !== authTabId) return;

          if (
            details.redirectUrl &&
            details.redirectUrl.startsWith(redirectUri)
          ) {
            let code, error, errorDesc;

            try {
              const urlObj = new URL(details.redirectUrl);
              code = urlObj.searchParams.get("code");
              error = urlObj.searchParams.get("error");
              errorDesc = urlObj.searchParams.get("error_description");
            } catch (urlError) {
              const codeMatch = details.redirectUrl.match(/[?&]code=([^&]+)/);
              if (codeMatch) {
                code = decodeURIComponent(codeMatch[1]);
              }
              const errorMatch = details.redirectUrl.match(/[?&]error=([^&]+)/);
              if (errorMatch) {
                error = decodeURIComponent(errorMatch[1]);
              }
            }

            if (code) {
              cleanup(authWindowId);

              exchangeCodeForToken(
                clientId,
                redirectUri,
                scope,
                code,
                codeVerifier,
                tokenUrl,
              )
                .then((tokenData) => {
                  chrome.storage.local.set({
                    lastTokenResult: {
                      success: true,
                      tokenData,
                      clientId,
                      scope,
                      timestamp: Date.now(),
                    },
                  });
                  resolve({ success: true, tokenData, clientId, scope });
                })
                .catch((error) => {
                  console.error("Token exchange failed:", error);
                  chrome.storage.local.set({
                    lastTokenResult: {
                      success: false,
                      error: error.message,
                      timestamp: Date.now(),
                    },
                  });
                  reject(error);
                });
            } else if (error) {
              cleanup(authWindowId);
              reject(
                new Error(
                  `OAuth error: ${error} - ${errorDesc || "No description"}`,
                ),
              );
            }
          }
        };

        chrome.webRequest.onBeforeRedirect.addListener(
          webRequestRedirectListener,
          { urls: ["<all_urls>"] },
        );

        windowClosedListener = (windowId) => {
          if (windowId === authWindowId) {
            cleanup(authWindowId);
            reject(new Error("Auth window closed by user"));
          }
        };

        chrome.windows.onRemoved.addListener(windowClosedListener);
      },
    );
  });
}

async function exchangeCodeForToken(
  clientId,
  redirectUri,
  scope,
  code,
  codeVerifier,
  tokenUrl,
) {
  const params = {
    client_id: clientId,
    scope: scope,
    code: code,
    redirect_uri: redirectUri,
    grant_type: "authorization_code",
    code_verifier: codeVerifier,
  };

  const tokenResponse = await fetch(
    tokenUrl || "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: new URLSearchParams(params),
    },
  );

  if (!tokenResponse.ok) {
    const errorData = await tokenResponse.json();
    throw new Error(errorData.error_description || "Failed to get token");
  }

  return await tokenResponse.json();
}
