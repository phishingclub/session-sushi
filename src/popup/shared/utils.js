async function copyToClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    showToast("Copied to clipboard", "success");
  } catch (error) {
    console.error("Error copying to clipboard:", error);
    showToast("Failed to copy to clipboard", "error");
  }
}

async function openInWindow() {
  try {
    const newWindow = await chrome.windows.create({
      url: chrome.runtime.getURL("popup.html"),
      type: "popup",
    });

    await chrome.windows.update(newWindow.id, { state: "maximized" });

    window.close();
  } catch (error) {
    console.error("Error opening window:", error);
    showToast("Failed to open window", "error");
  }
}

function downloadFile(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);

  chrome.downloads.download(
    {
      url: url,
      filename: filename,
      saveAs: true,
    },
    () => {
      URL.revokeObjectURL(url);
    },
  );
}

function showToast(message, type = "info") {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.className = `toast toast-${type} show`;

  setTimeout(() => {
    toast.classList.remove("show");
  }, 3000);
}

function truncate(text, length) {
  if (text.length > length) {
    return text.substring(0, length) + "...";
  }
  return text;
}

function showLoadingIndicator(show) {
  const indicator = document.getElementById("loadingIndicator");
  if (indicator) {
    if (show) {
      indicator.classList.remove("hidden");
    } else {
      indicator.classList.add("hidden");
    }
  }
}

function handleContextImport() {
  const activeTab = document.querySelector(".tab-btn.active");
  const currentTab = activeTab ? activeTab.dataset.tab : "cookies";

  if (currentTab === "sessions") {
    showImportM365SessionModal();
  } else {
    showImportModal();
  }
}

async function handleContextExport() {
  const activeTab = document.querySelector(".tab-btn.active");
  const currentTab = activeTab ? activeTab.dataset.tab : "cookies";

  if (currentTab === "sessions") {
    await exportM365Sessions();
  } else {
    await exportCookies();
  }
}

function handleContextClear() {
  const activeTab = document.querySelector(".tab-btn.active");
  const currentTab = activeTab ? activeTab.dataset.tab : "cookies";

  if (currentTab === "sessions") {
    showClearAllM365Modal();
  } else {
    showClearAllModal();
  }
}

/**
 * Creates a consistent error display element
 * @param {string} message - The error message to display
 * @param {Object} options - Optional configuration
 * @param {string} options.title - Error title (default: "Error:")
 * @param {boolean} options.isPermissionError - Whether this is a permission-related error
 * @returns {HTMLElement} The error box element
 */
function createErrorBox(message, options = {}) {
  const { title = "Error:", isPermissionError = false } = options;

  const errorBox = document.createElement("div");
  errorBox.className = isPermissionError
    ? "error-box permission-error visible"
    : "error-box visible";

  // Add title
  if (title) {
    const strongEl = document.createElement("strong");
    strongEl.textContent = title;
    errorBox.appendChild(strongEl);
  }

  // Add message
  const messageP = document.createElement("p");
  messageP.textContent = message;
  errorBox.appendChild(messageP);

  return errorBox;
}

/**
 * Creates a permission denied error box with consistent styling
 * @param {string} permission - The required permission
 * @param {string} customMessage - Optional custom message
 * @returns {HTMLElement} The permission error box element
 */
function createPermissionError(permission, customMessage = null) {
  const message = customMessage || `Missing required permission: ${permission}`;

  return createErrorBox(message, {
    title: "Permission Denied (403)",
    isPermissionError: true,
  });
}

/**
 * Shows an error in a container, replacing its content
 * @param {HTMLElement} container - The container to show the error in
 * @param {string} message - The error message
 * @param {Object} options - Optional configuration (same as createErrorBox)
 */
function showErrorInContainer(container, message, options = {}) {
  if (!container) return;
  container.innerHTML = "";
  container.appendChild(createErrorBox(message, options));
}
