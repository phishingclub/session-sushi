let currentDirectoryType = null;
let directoryItems = [];
let directorySearchResults = [];
let isSearching = false;

async function initializeDirectory() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showDirectoryNoSession();
    return;
  }

  await loadDirectoryType("users");
}

function showDirectoryNoSession() {
  const container = document.getElementById("directoryContainer");
  if (!container) return;

  container.textContent = "";
  const emptyDiv = document.createElement("div");
  emptyDiv.className = "mailbox-empty";
  emptyDiv.textContent = "No active session";
  container.appendChild(emptyDiv);
}

function clearDirectoryDisplay() {
  const container = document.getElementById("directoryContainer");
  if (!container) return;

  container.textContent = "";
  const emptyDiv = document.createElement("div");
  emptyDiv.className = "mailbox-empty";
  emptyDiv.textContent = "Choose a directory type from the sidebar to browse";
  container.appendChild(emptyDiv);

  const typeLabel = document.getElementById("directoryTypeLabel");
  if (typeLabel) {
    typeLabel.textContent = "";
  }

  const statsBar = document.getElementById("directoryStatsBar");
  if (statsBar) {
    statsBar.classList.add("hidden");
  }
}

async function loadDirectoryType(type) {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  currentDirectoryType = type;
  isSearching = false;

  document.querySelectorAll(".directory-type-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.type === type);
  });

  const typeLabel = document.getElementById("directoryTypeLabel");
  if (typeLabel) {
    const labels = {
      users: "👤 Users",
      groups: "👥 Groups",
      devices: "💻 Devices",
      applications: "📱 Applications",
      servicePrincipals: "🔑 Service Principals",
    };
    typeLabel.textContent = labels[type] || "";
  }

  const inviteGuestBtn = document.getElementById("inviteGuestBtn");
  if (inviteGuestBtn) {
    inviteGuestBtn.style.display = type === "users" ? "block" : "none";
  }

  const searchInput = document.getElementById("directorySearch");
  if (searchInput) {
    searchInput.value = "";
  }

  const container = document.getElementById("directoryContainer");
  if (!container) return;

  container.textContent = "";
  const loadingDiv = document.createElement("div");
  loadingDiv.className = "loading-indicator";
  loadingDiv.textContent = "Loading...";
  container.appendChild(loadingDiv);

  try {
    let url;
    if (type === "users") {
      url =
        "https://graph.microsoft.com/v1.0/users?$top=100&$select=id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,accountEnabled&$orderby=displayName";
    } else if (type === "groups") {
      url =
        "https://graph.microsoft.com/v1.0/groups?$top=100&$select=id,displayName,mail,description,groupTypes,securityEnabled&$orderby=displayName";
    } else if (type === "devices") {
      url =
        "https://graph.microsoft.com/v1.0/devices?$top=100&$select=id,displayName,operatingSystem,operatingSystemVersion,trustType,accountEnabled";
    } else if (type === "applications") {
      url =
        "https://graph.microsoft.com/v1.0/applications?$top=100&$select=id,appId,displayName,createdDateTime,signInAudience";
    } else if (type === "servicePrincipals") {
      url =
        "https://graph.microsoft.com/v1.0/servicePrincipals?$top=100&$select=id,appId,displayName,servicePrincipalType,accountEnabled";
    }

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(
        errorData.error?.message ||
          `HTTP ${response.status}: ${response.statusText}`,
      );
    }

    const data = await response.json();
    directoryItems = data.value || [];

    renderDirectoryItems();
    updateDirectoryStats();
  } catch (error) {
    console.error("Error loading directory:", error);
    showToast(`Failed to load ${type}: ${error.message}`, "error");

    showErrorInContainer(container, error.message, {
      title: "Error loading directory:",
    });
  }
}

function renderDirectoryItems() {
  const container = document.getElementById("directoryContainer");
  if (!container) return;

  const items = isSearching ? directorySearchResults : directoryItems;

  if (items.length === 0) {
    container.textContent = "";
    const emptyDiv = document.createElement("div");
    emptyDiv.className = "mailbox-empty";
    emptyDiv.textContent = "No items found";
    container.appendChild(emptyDiv);
    return;
  }

  container.textContent = "";
  const itemsContainer = document.createElement("div");
  itemsContainer.className = "onedrive-items-container";

  items.forEach((item) => {
    const itemDiv = createDirectoryItemElement(item);
    itemsContainer.appendChild(itemDiv);
  });

  container.appendChild(itemsContainer);
  setupDirectoryItemListeners();
}

function createDirectoryItemElement(item) {
  const itemDiv = document.createElement("div");
  itemDiv.className = "onedrive-item";
  itemDiv.setAttribute("data-item-id", item.id);

  // Icon
  const iconDiv = document.createElement("div");
  iconDiv.className = "onedrive-item-icon";

  if (currentDirectoryType === "users") {
    iconDiv.textContent = item.accountEnabled ? "👤" : "🚫";
    iconDiv.title = item.accountEnabled ? "Enabled" : "Disabled";
  } else if (currentDirectoryType === "groups") {
    const isM365Group = item.groupTypes && item.groupTypes.includes("Unified");
    iconDiv.textContent = isM365Group
      ? "📧"
      : item.securityEnabled
        ? "🔒"
        : "👥";
    iconDiv.title = isM365Group
      ? "Microsoft 365 Group"
      : item.securityEnabled
        ? "Security Group"
        : "Distribution Group";
  } else if (currentDirectoryType === "devices") {
    iconDiv.textContent = item.accountEnabled ? "💻" : "🚫";
    iconDiv.title = item.accountEnabled ? "Enabled" : "Disabled";
  } else if (currentDirectoryType === "applications") {
    iconDiv.textContent = "📱";
  } else if (currentDirectoryType === "servicePrincipals") {
    iconDiv.textContent = item.accountEnabled ? "🔑" : "🚫";
    iconDiv.title = item.accountEnabled ? "Enabled" : "Disabled";
  }

  // Details
  const detailsDiv = document.createElement("div");
  detailsDiv.className = "onedrive-item-details";

  const nameDiv = document.createElement("div");
  nameDiv.className = "onedrive-item-name";
  nameDiv.textContent = item.displayName || "N/A";

  const metaDiv = document.createElement("div");
  metaDiv.className = "onedrive-item-meta";
  metaDiv.textContent = getItemMetaText(item);

  detailsDiv.appendChild(nameDiv);
  detailsDiv.appendChild(metaDiv);

  // Actions
  const actionsDiv = document.createElement("div");
  actionsDiv.className = "onedrive-item-actions";

  // Add Members button for groups
  if (currentDirectoryType === "groups") {
    const membersBtn = document.createElement("button");
    membersBtn.className = "btn btn-small btn-secondary";
    membersBtn.setAttribute("data-action", "members");
    membersBtn.setAttribute("data-item-id", item.id);
    membersBtn.setAttribute("data-item-name", item.displayName || "");
    membersBtn.className = "btn btn-small btn-secondary btn-compact";
    membersBtn.textContent = "👥 Members";
    actionsDiv.appendChild(membersBtn);
  }

  // Details button
  const detailsBtn = document.createElement("button");
  detailsBtn.className = "btn btn-small btn-secondary btn-compact";
  detailsBtn.setAttribute("data-action", "details");
  detailsBtn.setAttribute("data-item-id", item.id);
  detailsBtn.textContent = "ℹ️ Details";
  actionsDiv.appendChild(detailsBtn);

  itemDiv.appendChild(iconDiv);
  itemDiv.appendChild(detailsDiv);
  itemDiv.appendChild(actionsDiv);

  return itemDiv;
}

function getItemMetaText(item) {
  const metaParts = [];

  if (currentDirectoryType === "users") {
    if (item.mail || item.userPrincipalName) {
      metaParts.push(item.mail || item.userPrincipalName);
    }
    if (item.jobTitle) {
      metaParts.push(item.jobTitle);
    }
    if (item.department) {
      metaParts.push(item.department);
    }
    if (item.officeLocation) {
      metaParts.push(item.officeLocation);
    }
  } else if (currentDirectoryType === "groups") {
    if (item.mail) {
      metaParts.push(item.mail);
    }
    if (item.description) {
      metaParts.push(item.description);
    }
  } else if (currentDirectoryType === "devices") {
    if (item.operatingSystem) {
      metaParts.push(item.operatingSystem);
    }
    if (item.operatingSystemVersion) {
      metaParts.push(item.operatingSystemVersion);
    }
    if (item.trustType) {
      metaParts.push(`Trust: ${item.trustType}`);
    }
  } else if (currentDirectoryType === "applications") {
    if (item.appId) {
      metaParts.push(`App ID: ${item.appId}`);
    }
    if (item.signInAudience) {
      metaParts.push(item.signInAudience);
    }
    if (item.createdDateTime) {
      const date = new Date(item.createdDateTime);
      metaParts.push(`Created: ${date.toLocaleDateString()}`);
    }
  } else if (currentDirectoryType === "servicePrincipals") {
    if (item.appId) {
      metaParts.push(`App ID: ${item.appId}`);
    }
    if (item.servicePrincipalType) {
      metaParts.push(item.servicePrincipalType);
    }
  }

  return metaParts.join(" • ") || "No additional info";
}

function updateDirectoryStats() {
  const statsBar = document.getElementById("directoryStatsBar");
  const stats = document.getElementById("directoryStats");

  if (!statsBar || !stats) return;

  const items = isSearching ? directorySearchResults : directoryItems;
  const count = items.length;

  if (statsBar) {
    if (count > 0) {
      statsBar.classList.remove("hidden");
      stats.textContent = `${count} item${count !== 1 ? "s" : ""}`;
    } else {
      statsBar.classList.add("hidden");
    }
  }
}

// Search directory
async function searchDirectory(query) {
  if (!query || !query.trim()) {
    // Clear search
    isSearching = false;
    directorySearchResults = [];
    renderDirectoryItems();
    updateDirectoryStats();
    return;
  }

  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  if (!currentDirectoryType) {
    showToast("Please select a directory type first");
    return;
  }

  const container = document.getElementById("directoryContainer");
  if (!container) return;

  container.textContent = "";
  const loadingDiv = document.createElement("div");
  loadingDiv.className = "loading-indicator";
  loadingDiv.textContent = "Searching...";
  container.appendChild(loadingDiv);

  try {
    const searchTerm = encodeURIComponent(query.trim());
    let url;

    if (currentDirectoryType === "users") {
      url = `https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'${searchTerm}') or startswith(userPrincipalName,'${searchTerm}') or startswith(mail,'${searchTerm}')&$top=100&$select=id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,accountEnabled`;
    } else if (currentDirectoryType === "groups") {
      url = `https://graph.microsoft.com/v1.0/groups?$filter=startswith(displayName,'${searchTerm}') or startswith(mail,'${searchTerm}')&$top=100&$select=id,displayName,mail,description,groupTypes,securityEnabled`;
    } else if (currentDirectoryType === "devices") {
      url = `https://graph.microsoft.com/v1.0/devices?$filter=startswith(displayName,'${searchTerm}')&$top=100&$select=id,displayName,operatingSystem,operatingSystemVersion,trustType,accountEnabled`;
    } else if (currentDirectoryType === "applications") {
      url = `https://graph.microsoft.com/v1.0/applications?$filter=startswith(displayName,'${searchTerm}')&$top=100&$select=id,appId,displayName,createdDateTime,signInAudience`;
    } else if (currentDirectoryType === "servicePrincipals") {
      url = `https://graph.microsoft.com/v1.0/servicePrincipals?$filter=startswith(displayName,'${searchTerm}')&$top=100&$select=id,appId,displayName,servicePrincipalType,accountEnabled`;
    }

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(
        errorData.error?.message ||
          `HTTP ${response.status}: ${response.statusText}`,
      );
    }

    const data = await response.json();
    directorySearchResults = data.value || [];
    isSearching = true;

    renderDirectoryItems();
    updateDirectoryStats();
    showToast(
      `Found ${directorySearchResults.length} result${directorySearchResults.length !== 1 ? "s" : ""}`,
    );
  } catch (error) {
    console.error("Error searching directory:", error);
    showToast(`Search failed: ${error.message}`, "error");

    showErrorInContainer(container, error.message, {
      title: "Search failed:",
    });
  }
}

// Show item details
async function showDirectoryItemDetails(itemId) {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  try {
    let url;
    if (currentDirectoryType === "users") {
      url = `https://graph.microsoft.com/v1.0/users/${itemId}`;
    } else if (currentDirectoryType === "groups") {
      url = `https://graph.microsoft.com/v1.0/groups/${itemId}`;
    } else if (currentDirectoryType === "devices") {
      url = `https://graph.microsoft.com/v1.0/devices/${itemId}`;
    } else if (currentDirectoryType === "applications") {
      url = `https://graph.microsoft.com/v1.0/applications/${itemId}`;
    } else if (currentDirectoryType === "servicePrincipals") {
      url = `https://graph.microsoft.com/v1.0/servicePrincipals/${itemId}`;
    }

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const item = await response.json();
    displayDirectoryItemDetailsModal(item);
  } catch (error) {
    console.error("Error loading item details:", error);
    showToast(`Failed to load details: ${error.message}`);
  }
}

function displayDirectoryItemDetailsModal(item) {
  const modal = document.getElementById("directoryDetailsModal");
  if (!modal) return;

  const content = document.getElementById("directoryDetailsContent");
  if (!content) return;

  content.textContent = "";

  Object.keys(item).forEach((key) => {
    // Skip @odata metadata fields
    if (key.startsWith("@odata")) {
      return;
    }

    const detailRow = document.createElement("div");
    detailRow.className = "detail-row";

    const label = document.createElement("div");
    label.className = "detail-label";
    label.textContent = key;

    const valueDiv = document.createElement("div");
    valueDiv.className = "detail-value";

    const value = item[key];
    if (value === null || value === undefined) {
      valueDiv.textContent = "N/A";
    } else if (Array.isArray(value)) {
      if (value.length === 0) {
        valueDiv.textContent = "[]";
      } else if (value.every((v) => typeof v !== "object")) {
        valueDiv.textContent = value.join(", ");
      } else {
        valueDiv.textContent = JSON.stringify(value, null, 2);
      }
    } else if (typeof value === "object") {
      valueDiv.textContent = JSON.stringify(value, null, 2);
    } else if (typeof value === "boolean") {
      valueDiv.textContent = value ? "Yes" : "No";
    } else {
      valueDiv.textContent = String(value);
    }

    detailRow.appendChild(label);
    detailRow.appendChild(valueDiv);
    content.appendChild(detailRow);
  });

  modal.classList.add("modal-show");
}

function closeDirectoryDetailsModal() {
  const modal = document.getElementById("directoryDetailsModal");
  if (modal) {
    modal.classList.remove("modal-show");
  }
}

// Show group members
async function showGroupMembers(groupId, groupName) {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const existingModal = document.getElementById("groupMembersModal");
  if (existingModal) {
    existingModal.remove();
  }

  const modal = document.createElement("div");
  modal.id = "groupMembersModal";
  modal.className = "modal";
  modal.classList.add("modal-show");

  const content = document.createElement("div");
  content.className = "modal-content";

  const header = document.createElement("div");
  header.className = "modal-header";

  const title = document.createElement("h2");
  title.textContent = `Members of ${groupName}`;

  const closeBtn = document.createElement("button");
  closeBtn.className = "modal-close";
  closeBtn.textContent = "✕";
  closeBtn.addEventListener("click", closeGroupMembersModal);

  header.appendChild(title);
  header.appendChild(closeBtn);

  const body = document.createElement("div");
  body.className = "modal-body";
  body.className = "modal-body overflow-y-auto";
  body.style.maxHeight = "60vh";

  const loadingDiv = document.createElement("div");
  loadingDiv.className = "loading-indicator";
  loadingDiv.textContent = "Loading members...";
  body.appendChild(loadingDiv);

  const footer = document.createElement("div");
  footer.className = "modal-footer";

  const closeFooterBtn = document.createElement("button");
  closeFooterBtn.className = "btn";
  closeFooterBtn.textContent = "Close";
  closeFooterBtn.addEventListener("click", closeGroupMembersModal);

  footer.appendChild(closeFooterBtn);

  content.appendChild(header);
  content.appendChild(body);
  content.appendChild(footer);
  modal.appendChild(content);

  document.body.appendChild(modal);

  modal.addEventListener("click", (e) => {
    if (e.target === modal) {
      closeGroupMembersModal();
    }
  });

  try {
    const url = `https://graph.microsoft.com/v1.0/groups/${groupId}/members?$select=id,displayName,mail,userPrincipalName,jobTitle&$top=100`;
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const data = await response.json();
    const members = data.value || [];

    body.textContent = "";

    if (members.length === 0) {
      const emptyDiv = document.createElement("div");
      emptyDiv.className = "mailbox-empty";
      emptyDiv.textContent = "No members found";
      body.appendChild(emptyDiv);
      return;
    }

    const itemsContainer = document.createElement("div");
    itemsContainer.className = "onedrive-items-container";

    members.forEach((member) => {
      const itemDiv = document.createElement("div");
      itemDiv.className = "onedrive-item";

      const iconDiv = document.createElement("div");
      iconDiv.className = "onedrive-item-icon";
      iconDiv.textContent = "👤";

      const detailsDiv = document.createElement("div");
      detailsDiv.className = "onedrive-item-details";

      const nameDiv = document.createElement("div");
      nameDiv.className = "onedrive-item-name";
      nameDiv.textContent = member.displayName || "N/A";

      const metaDiv = document.createElement("div");
      metaDiv.className = "onedrive-item-meta";
      const metaParts = [];
      if (member.mail || member.userPrincipalName) {
        metaParts.push(member.mail || member.userPrincipalName);
      }
      if (member.jobTitle) {
        metaParts.push(member.jobTitle);
      }
      metaDiv.textContent = metaParts.join(" • ") || "No additional info";

      detailsDiv.appendChild(nameDiv);
      detailsDiv.appendChild(metaDiv);

      itemDiv.appendChild(iconDiv);
      itemDiv.appendChild(detailsDiv);

      itemsContainer.appendChild(itemDiv);
    });

    body.appendChild(itemsContainer);
  } catch (error) {
    console.error("Error loading item details:", error);
    showToast(`Failed to load details: ${error.message}`, "error");

    showErrorInContainer(body, error.message, {
      title: "Error loading details:",
    });
  }
}

function closeGroupMembersModal() {
  const modal = document.getElementById("groupMembersModal");
  if (modal) {
    modal.remove();
  }
}

// Setup event listeners
function setupDirectoryListeners() {
  const refreshBtn = document.getElementById("refreshDirectoryBtn");
  const searchInput = document.getElementById("directorySearch");
  const typesList = document.getElementById("directoryTypesList");
  const inviteGuestBtn = document.getElementById("inviteGuestBtn");

  if (refreshBtn) {
    refreshBtn.addEventListener("click", async () => {
      if (currentDirectoryType) {
        await loadDirectoryType(currentDirectoryType);
      }
    });
  }

  if (inviteGuestBtn) {
    inviteGuestBtn.addEventListener("click", () => {
      showInviteGuestModal();
    });
  }

  if (searchInput) {
    let searchTimeout;
    searchInput.addEventListener("input", (e) => {
      clearTimeout(searchTimeout);
      searchTimeout = setTimeout(() => {
        searchDirectory(e.target.value);
      }, 500);
    });
  }

  // Use event delegation on the parent container
  if (typesList) {
    typesList.addEventListener("click", async (e) => {
      const btn = e.target.closest(".directory-type-btn");
      if (btn) {
        const type = btn.dataset.type;
        await loadDirectoryType(type);
      }
    });
  }

  // Close directory details modal
  const directoryDetailsModal = document.getElementById(
    "directoryDetailsModal",
  );
  if (directoryDetailsModal) {
    const closeBtn = directoryDetailsModal.querySelector(".modal-close");
    if (closeBtn) {
      closeBtn.addEventListener("click", closeDirectoryDetailsModal);
    }

    directoryDetailsModal.addEventListener("click", (e) => {
      if (e.target === directoryDetailsModal) {
        closeDirectoryDetailsModal();
      }
    });
  }
}

function showInviteGuestModal() {
  const modal = document.createElement("div");
  modal.className = "modal modal-show";

  const modalContent = document.createElement("div");
  modalContent.className = "modal-content max-width-600";

  const modalHeader = document.createElement("div");
  modalHeader.className = "modal-header";

  const title = document.createElement("h2");
  title.textContent = "Invite Guest User";
  modalHeader.appendChild(title);

  const closeBtn = document.createElement("button");
  closeBtn.className = "modal-close";
  closeBtn.textContent = "×";
  closeBtn.addEventListener("click", () => modal.remove());
  modalHeader.appendChild(closeBtn);

  modalContent.appendChild(modalHeader);

  const modalBody = document.createElement("div");
  modalBody.className = "modal-body max-height-500 overflow-y-auto";

  const form = document.createElement("div");
  form.className = "form-container";

  const emailGroup = document.createElement("div");
  emailGroup.className = "form-group";
  const emailLabel = document.createElement("label");
  emailLabel.textContent = "Email Address:";
  emailLabel.className = "form-label";
  const emailInput = document.createElement("input");
  emailInput.type = "email";
  emailInput.id = "guestEmail";
  emailInput.placeholder = "guest@external.com";
  emailInput.className = "form-input";
  emailInput.style.background = "var(--input-bg)";
  emailInput.style.color = "var(--text-color)";
  emailGroup.appendChild(emailLabel);
  emailGroup.appendChild(emailInput);
  form.appendChild(emailGroup);

  const nameGroup = document.createElement("div");
  nameGroup.className = "form-group";
  const nameLabel = document.createElement("label");
  nameLabel.textContent = "Display Name:";
  nameLabel.className = "form-label";
  const nameInput = document.createElement("input");
  nameInput.type = "text";
  nameInput.id = "guestDisplayName";
  nameInput.placeholder = "Guest User Name";
  nameInput.className = "form-input";
  nameInput.style.background = "var(--input-bg)";
  nameInput.style.color = "var(--text-color)";
  nameGroup.appendChild(nameLabel);
  nameGroup.appendChild(nameInput);
  form.appendChild(nameGroup);

  const redirectGroup = document.createElement("div");
  redirectGroup.className = "form-group";
  const redirectLabel = document.createElement("label");
  redirectLabel.textContent = "Redirect URL:";
  redirectLabel.className = "form-label";
  const redirectInput = document.createElement("input");
  redirectInput.type = "url";
  redirectInput.id = "guestRedirectUrl";
  redirectInput.placeholder = "https://portal.office.com";
  redirectInput.value = "https://portal.office.com";
  redirectInput.className = "form-input";
  redirectInput.style.background = "var(--input-bg)";
  redirectInput.style.color = "var(--text-color)";
  redirectGroup.appendChild(redirectLabel);
  redirectGroup.appendChild(redirectInput);
  form.appendChild(redirectGroup);

  const messageGroup = document.createElement("div");
  messageGroup.className = "form-group";
  const messageLabel = document.createElement("label");
  messageLabel.textContent = "Custom Message:";
  messageLabel.className = "form-label";
  const messageTextarea = document.createElement("textarea");
  messageTextarea.id = "guestMessage";
  messageTextarea.className = "textarea";
  messageTextarea.rows = 3;
  messageTextarea.placeholder =
    "You have been invited to access our organization...";
  messageTextarea.style.width = "100%";
  messageTextarea.style.padding = "8px 12px";
  messageTextarea.style.border = "1px solid var(--border-color)";
  messageTextarea.style.borderRadius = "6px";
  messageTextarea.style.fontSize = "13px";
  messageTextarea.style.fontFamily = "inherit";
  messageTextarea.style.resize = "vertical";
  messageTextarea.style.background = "var(--input-bg)";
  messageTextarea.style.color = "var(--text-color)";
  messageGroup.appendChild(messageLabel);
  messageGroup.appendChild(messageTextarea);
  form.appendChild(messageGroup);

  const sendEmailGroup = document.createElement("div");
  sendEmailGroup.className = "form-group";
  sendEmailGroup.style.marginBottom = "0";
  const sendEmailLabel = document.createElement("label");
  sendEmailLabel.style.display = "flex";
  sendEmailLabel.style.alignItems = "center";
  sendEmailLabel.style.gap = "8px";
  const sendEmailCheckbox = document.createElement("input");
  sendEmailCheckbox.type = "checkbox";
  sendEmailCheckbox.id = "sendInvitationEmail";
  sendEmailCheckbox.checked = true;
  const sendEmailText = document.createElement("span");
  sendEmailText.textContent = "Send email notification to guest";
  sendEmailLabel.appendChild(sendEmailCheckbox);
  sendEmailLabel.appendChild(sendEmailText);
  sendEmailGroup.appendChild(sendEmailLabel);
  form.appendChild(sendEmailGroup);

  modalBody.appendChild(form);
  modalContent.appendChild(modalBody);

  const modalFooter = document.createElement("div");
  modalFooter.className = "modal-footer";
  modalFooter.style.padding = "16px 20px";
  modalFooter.style.borderTop = "1px solid var(--border-color)";
  modalFooter.style.display = "flex";
  modalFooter.style.justifyContent = "flex-end";
  modalFooter.style.gap = "10px";
  modalFooter.style.background = "var(--bg-secondary)";

  const cancelBtn = document.createElement("button");
  cancelBtn.className = "btn btn-danger-outline";
  cancelBtn.textContent = "Cancel";
  cancelBtn.addEventListener("click", () => modal.remove());
  modalFooter.appendChild(cancelBtn);

  const inviteBtn = document.createElement("button");
  inviteBtn.className = "btn btn-primary";
  inviteBtn.textContent = "Create Invitation";
  inviteBtn.addEventListener("click", () => sendGuestInvitation(modal));
  modalFooter.appendChild(inviteBtn);

  modalContent.appendChild(modalFooter);
  modal.appendChild(modalContent);

  const handleEscape = (e) => {
    if (e.key === "Escape") {
      modal.remove();
      document.removeEventListener("keydown", handleEscape);
    }
  };
  document.addEventListener("keydown", handleEscape);

  modal.addEventListener("click", (e) => {
    if (e.target === modal) {
      modal.remove();
    }
  });

  document.body.appendChild(modal);
}

async function sendGuestInvitation(modal) {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const email = document.getElementById("guestEmail").value.trim();
  const displayName = document.getElementById("guestDisplayName").value.trim();
  const redirectUrl = document.getElementById("guestRedirectUrl").value.trim();
  const message = document.getElementById("guestMessage").value.trim();
  const sendEmail = document.getElementById("sendInvitationEmail").checked;

  if (!email) {
    showToast("Please enter an email address");
    return;
  }

  try {
    showToast("Sending invitation...");

    const invitationData = {
      invitedUserEmailAddress: email,
      inviteRedirectUrl: redirectUrl || "https://portal.office.com",
      sendInvitationMessage: sendEmail,
    };

    if (displayName) {
      invitationData.invitedUserDisplayName = displayName;
    }

    if (message && sendEmail) {
      invitationData.invitedUserMessageInfo = {
        customizedMessageBody: message,
      };
    }

    const response = await fetch(
      "https://graph.microsoft.com/v1.0/invitations",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(invitationData),
      },
    );

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error?.message || "Failed to send invitation");
    }

    modal.remove();
    showToast("✅ Invitation sent successfully");
    showInvitationSuccessModal(data);

    if (currentDirectoryType === "users") {
      await loadDirectoryType("users");
    }
  } catch (error) {
    console.error("Error sending invitation:", error);
    showToast(error.message || "Failed to send invitation", "error");
  }
}

function showInvitationSuccessModal(invitationData) {
  const modal = document.createElement("div");
  modal.className = "modal modal-show";

  const modalContent = document.createElement("div");
  modalContent.className = "modal-content max-width-600";

  const modalHeader = document.createElement("div");
  modalHeader.className = "modal-header";

  const title = document.createElement("h2");
  title.textContent = "Invitation Details";
  modalHeader.appendChild(title);

  const closeBtn = document.createElement("button");
  closeBtn.className = "modal-close";
  closeBtn.textContent = "×";
  closeBtn.addEventListener("click", () => modal.remove());
  modalHeader.appendChild(closeBtn);

  modalContent.appendChild(modalHeader);

  const modalBody = document.createElement("div");
  modalBody.className = "modal-body max-height-500 overflow-y-auto";

  const successMessage = document.createElement("div");
  successMessage.style.cssText =
    "margin-bottom: 20px; padding: 12px; background: var(--bg-secondary); border-radius: 6px; font-size: 14px;";
  successMessage.textContent = `✅ Guest invitation sent to ${invitationData.invitedUserEmailAddress}`;
  modalBody.appendChild(successMessage);

  const detailsContainer = document.createElement("div");
  detailsContainer.style.cssText =
    "display: flex; flex-direction: column; gap: 15px;";

  if (invitationData.inviteRedeemUrl) {
    const urlField = document.createElement("div");
    urlField.style.cssText =
      "background: var(--bg-secondary); padding: 12px; border-radius: 6px;";

    const urlLabel = document.createElement("div");
    urlLabel.style.cssText =
      "font-size: 12px; color: var(--text-secondary); margin-bottom: 4px; font-weight: 600;";
    urlLabel.textContent = "Invitation Redeem URL";

    const urlValue = document.createElement("div");
    urlValue.style.cssText =
      "font-size: 13px; word-break: break-all; margin-bottom: 8px;";
    urlValue.textContent = invitationData.inviteRedeemUrl;

    const copyUrlBtn = document.createElement("button");
    copyUrlBtn.className = "btn btn-small btn-secondary btn-compact";
    copyUrlBtn.textContent = "📋 Copy URL";
    copyUrlBtn.onclick = async () => {
      await copyToClipboard(invitationData.inviteRedeemUrl);
      showToast("Invitation URL copied to clipboard");
    };

    urlField.appendChild(urlLabel);
    urlField.appendChild(urlValue);
    urlField.appendChild(copyUrlBtn);
    detailsContainer.appendChild(urlField);
  }

  if (invitationData.invitedUser?.id) {
    const idField = document.createElement("div");
    idField.style.cssText =
      "background: var(--bg-secondary); padding: 12px; border-radius: 6px;";

    const idLabel = document.createElement("div");
    idLabel.style.cssText =
      "font-size: 12px; color: var(--text-secondary); margin-bottom: 4px; font-weight: 600;";
    idLabel.textContent = "Guest User ID";

    const idValue = document.createElement("div");
    idValue.style.cssText = "font-size: 13px; word-break: break-all;";
    idValue.textContent = invitationData.invitedUser.id;

    idField.appendChild(idLabel);
    idField.appendChild(idValue);
    detailsContainer.appendChild(idField);
  }

  if (invitationData.status) {
    const statusField = document.createElement("div");
    statusField.style.cssText =
      "background: var(--bg-secondary); padding: 12px; border-radius: 6px;";

    const statusLabel = document.createElement("div");
    statusLabel.style.cssText =
      "font-size: 12px; color: var(--text-secondary); margin-bottom: 4px; font-weight: 600;";
    statusLabel.textContent = "Status";

    const statusValue = document.createElement("div");
    statusValue.style.cssText = "font-size: 13px;";
    statusValue.textContent = invitationData.status;

    statusField.appendChild(statusLabel);
    statusField.appendChild(statusValue);
    detailsContainer.appendChild(statusField);
  }

  modalBody.appendChild(detailsContainer);
  modalContent.appendChild(modalBody);

  const modalFooter = document.createElement("div");
  modalFooter.className = "modal-footer";
  modalFooter.style.padding = "16px 20px";
  modalFooter.style.borderTop = "1px solid var(--border-color)";
  modalFooter.style.display = "flex";
  modalFooter.style.justifyContent = "flex-end";
  modalFooter.style.gap = "10px";
  modalFooter.style.background = "var(--bg-secondary)";

  const closeFooterBtn = document.createElement("button");
  closeFooterBtn.className = "btn btn-primary";
  closeFooterBtn.textContent = "Close";
  closeFooterBtn.addEventListener("click", () => modal.remove());
  modalFooter.appendChild(closeFooterBtn);

  modalContent.appendChild(modalFooter);
  modal.appendChild(modalContent);

  const handleEscape = (e) => {
    if (e.key === "Escape") {
      modal.remove();
      document.removeEventListener("keydown", handleEscape);
    }
  };
  document.addEventListener("keydown", handleEscape);

  modal.addEventListener("click", (e) => {
    if (e.target === modal) {
      modal.remove();
    }
  });

  document.body.appendChild(modal);
}

function setupDirectoryItemListeners() {
  const container = document.getElementById("directoryContainer");
  if (!container) return;

  container.addEventListener("click", async (e) => {
    const btn = e.target.closest("button[data-action]");
    if (!btn) return;

    const action = btn.getAttribute("data-action");
    const itemId = btn.getAttribute("data-item-id");

    if (action === "details" && itemId) {
      await showDirectoryItemDetails(itemId);
    } else if (action === "members" && itemId) {
      const itemName = btn.getAttribute("data-item-name") || "Group";
      await showGroupMembers(itemId, itemName);
    }
  });
}
