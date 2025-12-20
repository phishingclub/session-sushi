async function initialize() {
  await initializeTheme();

  const views = chrome.extension.getViews({ type: "popup" });
  // global
  isPopupMode = views.length > 0 && views.includes(window);

  window.addEventListener("beforeunload", () => {
    stopAutoRefreshTimer();
  });
}

// Called by load-tabs.js after all tab HTML is loaded
async function initializeApp() {
  setupEventListeners();

  await loadAllCookies();
  await loadSavedTokens();
  await loadM365Sessions();
  await restoreUIState();

  populatePresetDropdown();
  const presetSelector = document.getElementById("presetSelector");
  if (presetSelector && presetSelector.value) {
    applyPreset(presetSelector.value);
  }

  await checkCompletedAuth();

  setTimeout(() => {
    const footer = document.querySelector(".footer");
    if (footer) {
      footer.classList.add("hidden");
    }
  }, 5000);

  const sessionData = await chrome.storage.session.get([
    "pendingAction",
    "graphTokenConfig",
  ]);
  if (sessionData.pendingAction === "getGraphToken") {
    await chrome.storage.session.remove(["pendingAction", "graphTokenConfig"]);

    if (sessionData.graphTokenConfig) {
      if (sessionData.graphTokenConfig.clientId) {
        document.getElementById("clientIdInput").value =
          sessionData.graphTokenConfig.clientId;
      }
      if (sessionData.graphTokenConfig.redirectUri) {
        document.getElementById("redirectUriInput").value =
          sessionData.graphTokenConfig.redirectUri;
      }
      if (sessionData.graphTokenConfig.scope) {
        document.getElementById("scopeInput").value =
          sessionData.graphTokenConfig.scope;
      }
    }

    switchTab("sessions");

    setTimeout(() => getGraphToken(), 100);
  }

  await loadAutoRefreshSetting();
}

function setupEventListeners() {
  const themeToggleBtn = document.getElementById("themeToggle");
  if (themeToggleBtn) {
    themeToggleBtn.addEventListener("click", toggleTheme);
  }

  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => switchTab(btn.dataset.tab));
  });

  document.addEventListener("click", (e) => {
    const td = e.target.closest("td");
    if (td && td.closest("#cookiesTable")) {
      if (e.target.closest("button")) return;

      const text = td.getAttribute("title") || td.textContent.trim();
      if (text) {
        copyToClipboard(text);
        showToast("Copied to clipboard!");
      }
    }
  });

  document.addEventListener("keydown", (e) => {
    if (e.altKey && e.key === "1") {
      e.preventDefault();
      switchTab("cookies");
    }
    if (e.altKey && e.key === "2") {
      e.preventDefault();
      switchTab("sessions");
    }
    if (e.altKey && e.key === "3") {
      e.preventDefault();
      switchTab("graph");
    }
    if (e.altKey && e.key === "4") {
      e.preventDefault();
      switchTab("user");
    }
    if (e.altKey && e.key === "5") {
      e.preventDefault();
      switchTab("directory");
    }
    if (e.altKey && e.key === "6") {
      e.preventDefault();
      switchTab("mailbox");
    }
    if (e.altKey && e.key === "7") {
      e.preventDefault();
      switchTab("calendar");
    }
    if (e.altKey && e.key === "8") {
      e.preventDefault();
      switchTab("onedrive");
    }
    if (e.altKey && e.key === "9") {
      e.preventDefault();
      switchTab("sharepoint");
    }
    if (e.key === "Escape") {
      // Define modals with their close handlers
      const modals = [
        { id: "editModal", close: closeEditModal },
        { id: "clearAllModal", close: closeClearAllModal },
        { id: "clearAllM365Modal", close: closeClearAllM365Modal },
        { id: "importModal", close: closeImportModal },
        { id: "composeEmailModal", close: closeComposeEmailModal },
        { id: "uploadFileModal", close: closeUploadModal },
        { id: "createFolderModal", close: closeCreateFolderModal },
        { id: "itemDetailsModal", close: closeItemDetailsModal },
        {
          id: "directoryDetailsModal",
          close: closeDirectoryDetailsModal,
        },
        {
          id: "groupMembersModal",
          close: closeGroupMembersModal,
        },
        {
          id: "viewContactsModal",
          close: closeContactsModalHandler,
        },
        {
          id: "appointmentDetailsModal",
          close: () => {
            const modal = document.getElementById("appointmentDetailsModal");
            if (modal) modal.classList.remove("modal-show");
          },
        },
      ];

      // Check each modal and close the first visible one
      for (const { id, close } of modals) {
        const modal = document.getElementById(id);
        if (modal && window.getComputedStyle(modal).display !== "none") {
          e.preventDefault();
          e.stopPropagation();
          close();
          break;
        }
      }
    }
  });

  const headerExportBtn = document.getElementById("headerExport");
  if (headerExportBtn) {
    headerExportBtn.addEventListener("click", handleContextExport);
  }

  const headerImportBtn = document.getElementById("headerImport");
  if (headerImportBtn) {
    headerImportBtn.addEventListener("click", handleContextImport);
  }

  const clearAllCookiesBtn = document.getElementById("clearAllCookies");
  if (clearAllCookiesBtn) {
    clearAllCookiesBtn.addEventListener("click", handleContextClear);
  }

  const cookieSearchInput = document.getElementById("cookieSearch");
  if (cookieSearchInput) {
    cookieSearchInput.addEventListener("input", (e) => {
      clearTimeout(searchDebounceTimer);
      searchDebounceTimer = setTimeout(() => {
        filterCookies(e.target.value);
      }, 150);
    });
  }

  const closeImportModalBtn = document.getElementById("closeImportModal");
  if (closeImportModalBtn) {
    closeImportModalBtn.addEventListener("click", closeImportModal);
  }

  const cancelImportBtn = document.getElementById("cancelImport");
  if (cancelImportBtn) {
    cancelImportBtn.addEventListener("click", closeImportModal);
  }

  const confirmImportBtn = document.getElementById("confirmImport");
  if (confirmImportBtn) {
    confirmImportBtn.addEventListener("click", importCookies);
  }

  const importFileInput = document.getElementById("importFile");
  if (importFileInput) {
    importFileInput.addEventListener("change", handleFileSelect);
  }

  const closeEditModalBtn = document.getElementById("closeEditModal");
  if (closeEditModalBtn) {
    closeEditModalBtn.addEventListener("click", closeEditModal);
  }

  const cancelEditBtn = document.getElementById("cancelEdit");
  if (cancelEditBtn) {
    cancelEditBtn.addEventListener("click", closeEditModal);
  }

  const confirmEditBtn = document.getElementById("confirmEdit");
  if (confirmEditBtn) {
    confirmEditBtn.addEventListener("click", saveCookieEdit);
  }

  // Auto-enable Secure checkbox when SameSite is set to None
  const editSameSiteSelect = document.getElementById("editSameSite");
  const editSecureCheckbox = document.getElementById("editSecure");
  if (editSameSiteSelect && editSecureCheckbox) {
    editSameSiteSelect.addEventListener("change", (e) => {
      if (e.target.value === "no_restriction") {
        editSecureCheckbox.checked = true;
      }
    });
  }

  const closeClearAllModalBtn = document.getElementById("closeClearAllModal");
  if (closeClearAllModalBtn) {
    closeClearAllModalBtn.addEventListener("click", closeClearAllModal);
  }

  const cancelClearAllBtn = document.getElementById("cancelClearAll");
  if (cancelClearAllBtn) {
    cancelClearAllBtn.addEventListener("click", closeClearAllModal);
  }

  const confirmClearAllBtn = document.getElementById("confirmClearAll");
  if (confirmClearAllBtn) {
    confirmClearAllBtn.addEventListener("click", clearAllCookies);
  }

  const clearAllConfirmTextInput = document.getElementById(
    "clearAllConfirmText",
  );
  if (clearAllConfirmTextInput) {
    clearAllConfirmTextInput.addEventListener("input", (e) => {
      const confirmBtn = document.getElementById("confirmClearAll");
      if (confirmBtn) {
        confirmBtn.disabled = e.target.value !== "DELETE";
      }
    });
  }

  const importModal = document.getElementById("importModal");
  if (importModal) {
    importModal.addEventListener("click", (e) => {
      if (e.target.id === "importModal") {
        closeImportModal();
      }
    });
  }

  const editModal = document.getElementById("editModal");
  if (editModal) {
    editModal.addEventListener("click", (e) => {
      if (e.target.id === "editModal") {
        closeEditModal();
      }
    });
  }

  const clearAllModal = document.getElementById("clearAllModal");
  if (clearAllModal) {
    clearAllModal.addEventListener("click", (e) => {
      if (e.target.id === "clearAllModal") {
        closeClearAllModal();
      }
    });
  }

  const closeClearAllM365ModalBtn = document.getElementById(
    "closeClearAllM365Modal",
  );
  if (closeClearAllM365ModalBtn) {
    closeClearAllM365ModalBtn.addEventListener("click", closeClearAllM365Modal);
  }

  const cancelClearAllM365Btn = document.getElementById("cancelClearAllM365");
  if (cancelClearAllM365Btn) {
    cancelClearAllM365Btn.addEventListener("click", closeClearAllM365Modal);
  }

  const confirmClearAllM365Btn = document.getElementById("confirmClearAllM365");
  if (confirmClearAllM365Btn) {
    confirmClearAllM365Btn.addEventListener("click", clearAllM365Sessions);
  }

  const clearAllM365ConfirmTextInput = document.getElementById(
    "clearAllM365ConfirmText",
  );
  if (clearAllM365ConfirmTextInput) {
    clearAllM365ConfirmTextInput.addEventListener("input", (e) => {
      const confirmBtn = document.getElementById("confirmClearAllM365");
      if (confirmBtn) {
        confirmBtn.disabled = e.target.value !== "DELETE";
      }
    });
  }

  const clearAllM365Modal = document.getElementById("clearAllM365Modal");
  if (clearAllM365Modal) {
    clearAllM365Modal.addEventListener("click", (e) => {
      if (e.target.id === "clearAllM365Modal") {
        closeClearAllM365Modal();
      }
    });
  }

  const getGraphTokenBtn = document.getElementById("getGraphToken");
  const refreshActiveSessionBtn = document.getElementById(
    "refreshActiveSession",
  );
  const editActiveSessionBtn = document.getElementById("editActiveSession");
  const clearActiveSessionBtn = document.getElementById("clearActiveSession");

  const presetSelector = document.getElementById("presetSelector");
  if (presetSelector) {
    presetSelector.addEventListener("change", (e) => {
      applyPreset(e.target.value);
    });
  }

  if (getGraphTokenBtn) {
    getGraphTokenBtn.addEventListener("click", async () => {
      await getGraphToken();
    });
  }

  const copyAuthUrlBtn = document.getElementById("copyAuthUrl");
  if (copyAuthUrlBtn) {
    copyAuthUrlBtn.addEventListener("click", async () => {
      await copyAuthUrl();
    });
  }

  if (refreshActiveSessionBtn) {
    refreshActiveSessionBtn.addEventListener("click", async () => {
      await refreshActiveM365Session();
    });
  }

  if (editActiveSessionBtn) {
    editActiveSessionBtn.addEventListener("click", () => {
      showEditM365SessionModal();
    });
  }

  if (clearActiveSessionBtn) {
    clearActiveSessionBtn.addEventListener("click", async () => {
      await clearActiveM365Session();
    });
  }

  const autoRefreshCheckbox = document.getElementById("autoRefreshCheckbox");
  if (autoRefreshCheckbox) {
    autoRefreshCheckbox.addEventListener("change", async () => {
      await toggleAutoRefresh();
    });
  }

  setupQuickActions();

  const executeCustomQueryBtn = document.getElementById("executeGraphBtn");
  const customGraphQueryInput = document.getElementById("customGraphQuery");

  if (executeCustomQueryBtn) {
    executeCustomQueryBtn.addEventListener("click", async () => {
      const query = document.getElementById("customGraphQuery").value;
      await executeGraphQuery(query);
    });
  }

  if (customGraphQueryInput) {
    customGraphQueryInput.addEventListener("keypress", async (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        const query = e.target.value;
        await executeGraphQuery(query);
      }
    });
  }

  const copyGraphResponseBtn = document.getElementById("copyGraphResponse");
  if (copyGraphResponseBtn) {
    copyGraphResponseBtn.addEventListener("click", async () => {
      const response = document.getElementById("graphQueryResponse").value;
      await copyToClipboard(response);
      showToast("Response copied!");
    });
  }

  setupMailboxSearch();
  setupComposeEmailListeners();
  setupMessageDetailListeners();
  setupOneDriveListeners();
  setupUserListeners();
  setupDirectoryListeners();
  setupCalendarListeners();
  setupContactsListeners();
  setupSharePointListeners();
  setupTeamsListeners();

  const refreshMailboxBtn = document.getElementById("refreshMailbox");
  if (refreshMailboxBtn) {
    refreshMailboxBtn.addEventListener("click", async () => {
      await refreshMailbox();
    });
  }

  const viewMailRulesBtn = document.getElementById("viewMailRules");
  if (viewMailRulesBtn) {
    viewMailRulesBtn.addEventListener("click", async () => {
      await viewMailRules();
    });
  }

  const viewMailboxSettingsBtn = document.getElementById("viewMailboxSettings");
  if (viewMailboxSettingsBtn) {
    viewMailboxSettingsBtn.addEventListener("click", async () => {
      await viewMailboxSettings();
    });
  }

  const viewAutoReplyBtn = document.getElementById("viewAutoReply");
  if (viewAutoReplyBtn) {
    viewAutoReplyBtn.addEventListener("click", async () => {
      await viewAutoReply();
    });
  }

  setupFolderManagementListeners();

  document
    .getElementById("closeEditM365SessionModal")
    ?.addEventListener("click", closeEditM365SessionModal);
  document
    .getElementById("cancelEditM365Session")
    ?.addEventListener("click", closeEditM365SessionModal);
  document
    .getElementById("confirmEditM365Session")
    ?.addEventListener("click", confirmEditM365Session);

  document
    .getElementById("closeImportM365SessionModal")
    ?.addEventListener("click", closeImportM365SessionModal);
  document
    .getElementById("cancelImportM365Session")
    ?.addEventListener("click", closeImportM365SessionModal);
  document
    .getElementById("confirmImportM365Session")
    ?.addEventListener("click", confirmImportM365Session);
  document
    .getElementById("importM365SessionFile")
    ?.addEventListener("change", handleM365SessionFileSelect);

  document
    .getElementById("saveM365SessionModal")
    ?.addEventListener("click", (e) => {
      if (e.target.id === "saveM365SessionModal") {
        closeSaveM365SessionModal();
      }
    });
  document
    .getElementById("importM365SessionModal")
    ?.addEventListener("click", (e) => {
      if (e.target.id === "importM365SessionModal") {
        closeImportM365SessionModal();
      }
    });

  document
    .getElementById("composeEmailModal")
    ?.addEventListener("click", (e) => {
      if (e.target.id === "composeEmailModal") {
        closeComposeEmailModal();
      }
    });

  // Global Escape key handler to close modals
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      // Find any open modal and close it
      const allModals = document.querySelectorAll(".modal");

      for (const modal of allModals) {
        if (
          modal.style.display === "flex" ||
          modal.classList.contains("active")
        ) {
          // Find the close button in this modal
          const closeButton = modal.querySelector(".modal-close");
          if (closeButton) {
            closeButton.click();
            break; // Only close the first open modal
          }
        }
      }
    }
  });
}

document.addEventListener("DOMContentLoaded", initialize);
