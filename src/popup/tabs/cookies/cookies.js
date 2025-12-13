async function loadAllCookies() {
  try {
    showLoadingIndicator(true);
    const response = await chrome.runtime.sendMessage({
      action: "getAllCookies",
    });
    allCookies = response.cookies || [];
    filteredCookies = [...allCookies];

    requestAnimationFrame(() => {
      renderCookies();
      updateStats();
      showLoadingIndicator(false);
    });
  } catch (error) {
    console.error("Loading cookies:", error);
    showToast("Failed to load cookies", "error");
    showLoadingIndicator(false);
  }
}

async function updateStats() {
  try {
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    const isIncognito = tabs[0] && tabs[0].incognito;
    const count = allCookies.length;
    const filtered = filteredCookies.length;
    const displayed = Math.min(currentPage * ROWS_PER_PAGE, filtered);

    let text;
    if (filtered < count) {
      text = `Showing ${displayed} of ${filtered} matching cookies (${count} total)`;
    } else if (displayed < filtered) {
      text = `Showing ${displayed} of ${filtered} cookies (scroll for more)`;
    } else if (isIncognito) {
      text = `Total: ${count} cookies in incognito session`;
    } else {
      text = `Total: ${count} cookies across all sites`;
    }

    document.getElementById("cookieStats").textContent = text;
  } catch (error) {
    console.error("Error updating stats:", error);
  }
}

function filterCookies(query) {
  showLoadingIndicator(true);

  setTimeout(() => {
    const lowerQuery = query.toLowerCase();
    if (!query.trim()) {
      filteredCookies = [...allCookies];
    } else {
      filteredCookies = allCookies.filter(
        (cookie) =>
          cookie.name.toLowerCase().includes(lowerQuery) ||
          cookie.domain.toLowerCase().includes(lowerQuery) ||
          cookie.value.toLowerCase().includes(lowerQuery) ||
          cookie.path.toLowerCase().includes(lowerQuery),
      );
    }
    currentPage = 0;
    renderCookies();
    updateStats();
    showLoadingIndicator(false);
  }, 0);
}

function renderCookies() {
  const tbody = document.getElementById("cookiesTableBody");

  if (filteredCookies.length === 0) {
    tbody.innerHTML = "";
    const emptyRow = document.createElement("tr");
    const emptyCell = document.createElement("td");
    emptyCell.setAttribute("colspan", "7");
    emptyCell.className = "empty-state";
    emptyCell.textContent = "No cookies found";
    emptyRow.appendChild(emptyCell);
    tbody.appendChild(emptyRow);
    setupScrollListener(false);
    return;
  }

  const start = 0;
  const end = Math.min(ROWS_PER_PAGE, filteredCookies.length);
  currentPage = 1;

  const rows = renderCookieRows(start, end);
  tbody.innerHTML = "";
  tbody.appendChild(rows);

  tbody.querySelectorAll("button[data-action]").forEach((btn) => {
    btn.addEventListener("click", handleCookieAction);
  });

  setupScrollListener();
}

function renderCookieRows(start, end) {
  const fragment = document.createDocumentFragment();

  for (let i = start; i < end; i++) {
    const cookie = filteredCookies[i];
    const expires = cookie.expirationDate
      ? new Date(cookie.expirationDate * 1000).toLocaleString()
      : "Session";

    const sameSiteMap = {
      no_restriction: "None",
      unspecified: "Unspecified",
      lax: "Lax",
      strict: "Strict",
    };

    const row = document.createElement("tr");

    const nameCell = document.createElement("td");
    nameCell.title = cookie.name;
    nameCell.textContent = cookie.name;
    row.appendChild(nameCell);

    const domainCell = document.createElement("td");
    domainCell.title = cookie.domain;
    domainCell.textContent = cookie.domain;
    row.appendChild(domainCell);

    const valueCell = document.createElement("td");
    valueCell.title = cookie.value;
    valueCell.textContent = cookie.value;
    row.appendChild(valueCell);

    const pathCell = document.createElement("td");
    pathCell.textContent = cookie.path;
    row.appendChild(pathCell);

    const expiresCell = document.createElement("td");
    expiresCell.textContent = expires;
    row.appendChild(expiresCell);

    const flagsCell = document.createElement("td");
    if (cookie.secure) {
      const secureBadge = document.createElement("span");
      secureBadge.className = "badge badge-success";
      secureBadge.textContent = "Secure";
      flagsCell.appendChild(secureBadge);
      flagsCell.appendChild(document.createTextNode(" "));
    }
    if (cookie.httpOnly) {
      const httpOnlyBadge = document.createElement("span");
      httpOnlyBadge.className = "badge badge-warning";
      httpOnlyBadge.textContent = "HttpOnly";
      flagsCell.appendChild(httpOnlyBadge);
      flagsCell.appendChild(document.createTextNode(" "));
    }
    if (cookie.sameSite && cookie.sameSite !== "unspecified") {
      const sameSiteDisplay = sameSiteMap[cookie.sameSite] || cookie.sameSite;
      const sameSiteBadge = document.createElement("span");
      sameSiteBadge.className = "badge badge-info";
      sameSiteBadge.textContent = sameSiteDisplay;
      flagsCell.appendChild(sameSiteBadge);
    }
    row.appendChild(flagsCell);

    const actionsCell = document.createElement("td");

    const editBtn = document.createElement("button");
    editBtn.className = "btn btn-small btn-secondary";
    editBtn.setAttribute("data-action", "edit");
    editBtn.setAttribute("data-index", i);
    editBtn.className = "btn btn-small btn-secondary btn-compact";
    editBtn.textContent = "✏️ Edit";
    actionsCell.appendChild(editBtn);

    actionsCell.appendChild(document.createTextNode(" "));

    const deleteBtn = document.createElement("button");
    deleteBtn.className = "btn btn-small btn-danger-outline";
    deleteBtn.setAttribute("data-action", "delete");
    deleteBtn.setAttribute("data-index", i);
    deleteBtn.className = "btn btn-small btn-danger btn-compact";
    deleteBtn.textContent = "❌ Delete";
    actionsCell.appendChild(deleteBtn);

    row.appendChild(actionsCell);

    fragment.appendChild(row);
  }

  return fragment;
}

function setupScrollListener(enable = true) {
  const container = document.querySelector(".table-container");

  if (scrollListener) {
    container.removeEventListener("scroll", scrollListener);
    scrollListener = null;
  }

  if (!enable) return;

  scrollListener = () => {
    const scrollTop = container.scrollTop;
    const scrollHeight = container.scrollHeight;
    const clientHeight = container.clientHeight;

    if (scrollTop + clientHeight >= scrollHeight * 0.8) {
      loadMoreCookies();
    }
  };

  container.addEventListener("scroll", scrollListener);
}

function loadMoreCookies() {
  if (isLoadingMore) return;

  const totalPages = Math.ceil(filteredCookies.length / ROWS_PER_PAGE);
  if (currentPage >= totalPages) return;

  isLoadingMore = true;
  showLoadingIndicator(true);

  requestAnimationFrame(() => {
    const start = currentPage * ROWS_PER_PAGE;
    const end = Math.min(start + ROWS_PER_PAGE, filteredCookies.length);

    const tbody = document.getElementById("cookiesTableBody");
    const rows = renderCookieRows(start, end);

    tbody.appendChild(rows);

    const newButtons = tbody.querySelectorAll(
      `button[data-action][data-index]`,
    );
    newButtons.forEach((btn) => {
      const index = parseInt(btn.dataset.index);
      if (index >= start && index < end) {
        btn.addEventListener("click", handleCookieAction);
      }
    });

    currentPage++;
    isLoadingMore = false;
    showLoadingIndicator(false);
    updateStats();
  });
}

async function handleCookieAction(e) {
  const action = e.target.dataset.action;
  const index = parseInt(e.target.dataset.index);

  if (action === "edit") {
    editCookie(index);
  } else if (action === "delete") {
    deleteCookie(index);
  }
}

async function exportCookies() {
  try {
    if (allCookies.length === 0) {
      showToast("No cookies to export", "warning");
      return;
    }

    const exportData = allCookies.map((cookie) => ({
      domain: cookie.domain,
      expirationDate: cookie.expirationDate,
      hostOnly: cookie.hostOnly,
      httpOnly: cookie.httpOnly,
      name: cookie.name,
      path: cookie.path,
      sameSite: cookie.sameSite,
      secure: cookie.secure,
      session: cookie.session,
      storeId: cookie.storeId,
      value: cookie.value,
    }));

    const json = JSON.stringify(exportData, null, 2);
    const filename = `sushi_cookies_${new Date().toISOString().split("T")[0]}.json`;

    downloadFile(json, filename, "application/json");

    showToast(`Exported ${allCookies.length} cookies`, "success");
  } catch (error) {
    console.error("Error exporting cookies:", error);
    showToast("Failed to export cookies", "error");
  }
}

function showImportModal() {
  document.getElementById("importModal").classList.add("modal-show");
  document.getElementById("importData").value = "";
  document.getElementById("importFile").value = "";
}

function closeImportModal() {
  document.getElementById("importModal").classList.remove("modal-show");
}

function handleFileSelect(event) {
  event.stopPropagation();
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    document.getElementById("importData").value = e.target.result;
  };
  reader.onerror = () => {
    showToast("Failed to read file", "error");
  };
  reader.readAsText(file);
}

async function importCookies() {
  try {
    const data = document.getElementById("importData").value.trim();

    if (!data) {
      showToast("Please provide JSON data to import", "warning");
      return;
    }

    let cookies;
    try {
      cookies = JSON.parse(data);
    } catch (error) {
      showToast("Invalid JSON format", "error");
      return;
    }

    if (!Array.isArray(cookies)) {
      showToast("JSON must be an array of cookies", "error");
      return;
    }

    if (cookies.length === 0) {
      showToast("No cookies found in the data", "warning");
      return;
    }

    if (
      !confirm(
        `Import ${cookies.length} cookies? This will add them to your browser.`,
      )
    ) {
      return;
    }

    closeImportModal();
    showToast("Importing cookies...", "info");

    const response = await chrome.runtime.sendMessage({
      action: "importCookies",
      cookies: cookies,
    });

    if (response.imported > 0) {
      showToast(
        `Successfully imported ${response.imported} of ${response.total} cookies`,
        "success",
      );
    } else {
      showToast(
        `Failed to import cookies. ${response.failed} failed.`,
        "error",
      );
    }

    await loadAllCookies();
  } catch (error) {
    console.error("Error importing cookies:", error);
    showToast("Failed to import cookies", "error");
  }
}

function showClearAllModal() {
  if (allCookies.length === 0) {
    showToast("No cookies to clear", "warning");
    return;
  }

  const modal = document.getElementById("clearAllModal");
  const confirmText = document.getElementById("clearAllConfirmText");
  const confirmBtn = document.getElementById("confirmClearAll");

  confirmText.value = "";
  confirmBtn.disabled = true;

  const warningText = modal.querySelector("p");
  warningText.innerHTML = `This will <strong>permanently delete all ${allCookies.length} cookies</strong> from all sites. This action cannot be undone.`;

  modal.classList.add("modal-show");

  setTimeout(() => confirmText.focus(), 100);
}

function closeClearAllModal() {
  const modal = document.getElementById("clearAllModal");
  const confirmText = document.getElementById("clearAllConfirmText");
  const confirmBtn = document.getElementById("confirmClearAll");

  modal.classList.remove("modal-show");
  confirmText.value = "";
  confirmBtn.disabled = true;
}

async function clearAllCookies() {
  try {
    showToast("Clearing cookies...", "info");
    closeClearAllModal();

    const result = await chrome.runtime.sendMessage({
      action: "clearAllCookies",
    });

    if (result.removed > 0) {
      showToast(`Cleared ${result.removed} cookies`, "success");
    } else {
      showToast("No cookies were cleared", "warning");
    }

    await loadAllCookies();
  } catch (error) {
    console.error("Error clearing cookies:", error);
    showToast("Failed to clear cookies", "error");
  }
}

function editCookie(index) {
  currentEditingCookie = filteredCookies[index];
  if (!currentEditingCookie) return;

  document.getElementById("editName").value = currentEditingCookie.name;
  document.getElementById("editValue").value = currentEditingCookie.value;
  document.getElementById("editDomain").value = currentEditingCookie.domain;
  document.getElementById("editPath").value = currentEditingCookie.path;

  if (currentEditingCookie.expirationDate) {
    const date = new Date(currentEditingCookie.expirationDate * 1000);
    const localDateTime = new Date(
      date.getTime() - date.getTimezoneOffset() * 60000,
    )
      .toISOString()
      .slice(0, 16);
    document.getElementById("editExpiration").value = localDateTime;
  } else {
    document.getElementById("editExpiration").value = "";
  }

  document.getElementById("editSecure").checked = currentEditingCookie.secure;
  document.getElementById("editHttpOnly").checked =
    currentEditingCookie.httpOnly;
  document.getElementById("editSameSite").value =
    currentEditingCookie.sameSite || "lax";

  document.getElementById("editModalTitle").textContent =
    `Edit Cookie: ${currentEditingCookie.name}`;
  document.getElementById("editModal").classList.add("modal-show");
}

function closeEditModal() {
  document.getElementById("editModal").classList.remove("modal-show");
  currentEditingCookie = null;
}

async function saveCookieEdit() {
  if (!currentEditingCookie) return;

  try {
    const name = document.getElementById("editName").value;
    const value = document.getElementById("editValue").value;
    const domain = document.getElementById("editDomain").value;
    const path = document.getElementById("editPath").value;
    const expirationInput = document.getElementById("editExpiration").value;
    const secure = document.getElementById("editSecure").checked;
    const httpOnly = document.getElementById("editHttpOnly").checked;
    const sameSite = document.getElementById("editSameSite").value;

    if (!name || !value || !domain || !path) {
      showToast("Name, value, domain, and path are required", "error");
      return;
    }

    // Normalize sameSite value for Chrome API
    const validSameSiteValues = [
      "no_restriction",
      "lax",
      "strict",
      "unspecified",
    ];
    const normalizedSameSite = validSameSiteValues.includes(sameSite)
      ? sameSite
      : "lax";

    // SameSite=None requires Secure flag to be true
    let finalSecure = secure;
    if (normalizedSameSite === "no_restriction" && !secure) {
      finalSecure = true;
      showToast(
        "SameSite=None requires Secure flag. Secure has been enabled automatically.",
        "info",
      );
    }

    const protocol = currentEditingCookie.secure ? "https" : "http";
    const deleteUrl = `${protocol}://${currentEditingCookie.domain}${currentEditingCookie.path}`;

    await chrome.cookies.remove({
      url: deleteUrl,
      name: currentEditingCookie.name,
    });

    const newProtocol = finalSecure ? "https" : "http";
    const cleanDomain = domain.startsWith(".") ? domain.substring(1) : domain;
    const url = `${newProtocol}://${cleanDomain}${path}`;

    const cookieData = {
      name: name,
      value: value,
      domain: domain,
      path: path,
      secure: finalSecure,
      httpOnly: httpOnly,
      sameSite: normalizedSameSite,
      url: url,
    };

    if (expirationInput) {
      cookieData.expirationDate = new Date(expirationInput).getTime() / 1000;
    }

    await chrome.cookies.set(cookieData);

    showToast("Cookie updated successfully", "success");
    closeEditModal();
    await loadAllCookies();
  } catch (error) {
    console.error("Error saving cookie:", error);
    showToast("Failed to save cookie: " + error.message, "error");
  }
}

async function deleteCookie(index) {
  const cookie = filteredCookies[index];
  if (!cookie) return;

  if (!confirm(`Delete cookie "${cookie.name}" from ${cookie.domain}?`)) {
    return;
  }

  try {
    const protocol = cookie.secure ? "https" : "http";
    const url = `${protocol}://${cookie.domain}${cookie.path}`;

    await chrome.cookies.remove({
      url: url,
      name: cookie.name,
    });

    showToast("Cookie deleted successfully", "success");
    await loadAllCookies();
  } catch (error) {
    console.error("Error deleting cookie:", error);
    showToast("Failed to delete cookie", "error");
  }
}
