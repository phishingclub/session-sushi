async function viewMailRules() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  try {
    showToast("Loading mail rules...");

    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules",
      {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      throw new Error(`Failed to load mail rules: ${response.statusText}`);
    }

    const data = await response.json();
    const rules = data.value || [];

    showMailRulesModal(rules);
  } catch (error) {
    console.error("Failed to load mail rules:", error);
    showToast("Failed to load mail rules: " + error.message);
  }
}

function showMailRulesModal(rules) {
  const modal = document.createElement("div");
  modal.className = "modal modal-show modal-align-center";
  modal.id = "mailRulesModal";

  const modalContent = document.createElement("div");
  modalContent.className = "modal-content max-width-700 margin-0";

  const modalHeader = document.createElement("div");
  modalHeader.className = "modal-header";

  const title = document.createElement("h2");
  title.textContent = `Mail Rules (${rules.length})`;
  modalHeader.appendChild(title);

  const closeBtn = document.createElement("button");
  closeBtn.className = "modal-close";
  closeBtn.textContent = "×";
  closeBtn.addEventListener("click", () => modal.remove());
  modalHeader.appendChild(closeBtn);

  const modalBody = document.createElement("div");
  modalBody.className = "modal-body max-height-500 overflow-y-auto";

  const createRuleBtn = document.createElement("button");
  createRuleBtn.className = "btn btn-primary margin-bottom-16";
  createRuleBtn.textContent = "Create Mail Rule";
  createRuleBtn.addEventListener("click", () => {
    modal.remove();
    showCreateMailRuleModal();
  });
  modalBody.appendChild(createRuleBtn);

  if (rules.length === 0) {
    const emptyMsg = document.createElement("p");
    emptyMsg.textContent = "No mail rules configured";
    emptyMsg.className = "empty-message";
    modalBody.appendChild(emptyMsg);
  } else {
    rules.forEach((rule, index) => {
      const ruleCard = document.createElement("div");
      ruleCard.className = "rule-card";

      const ruleHeader = document.createElement("div");
      ruleHeader.className = "rule-header";

      const ruleInfoDiv = document.createElement("div");
      ruleInfoDiv.className = "rule-info";

      const ruleName = document.createElement("div");
      ruleName.className = "rule-name";
      ruleName.textContent = rule.displayName;
      ruleInfoDiv.appendChild(ruleName);

      const statusContainer = document.createElement("div");
      statusContainer.className = "rule-status-container";

      const statusLabel = document.createElement("span");
      statusLabel.textContent = "Status: ";
      statusLabel.className = "rule-status-label";
      statusContainer.appendChild(statusLabel);

      const ruleEnabled = document.createElement("span");
      ruleEnabled.textContent = rule.isEnabled ? "Enabled" : "Disabled";
      ruleEnabled.className = rule.isEnabled
        ? "rule-status-badge enabled"
        : "rule-status-badge disabled";
      statusContainer.appendChild(ruleEnabled);
      ruleInfoDiv.appendChild(statusContainer);

      ruleHeader.appendChild(ruleInfoDiv);

      const deleteBtn = document.createElement("button");
      deleteBtn.className = "btn btn-small btn-danger-outline btn-compact";
      deleteBtn.textContent = "❌ Delete";
      deleteBtn.addEventListener("click", async () => {
        if (confirm(`Delete rule "${rule.displayName}"?`)) {
          await deleteMailRule(rule.id, modal);
        }
      });
      ruleHeader.appendChild(deleteBtn);

      ruleCard.appendChild(ruleHeader);

      // Conditions
      if (rule.conditions) {
        const conditionsDiv = document.createElement("div");
        conditionsDiv.className = "rule-conditions";

        const conditionsTitle = document.createElement("div");
        conditionsTitle.textContent = "When:";
        conditionsTitle.className = "rule-section-title";
        conditionsDiv.appendChild(conditionsTitle);

        const conditionsList = document.createElement("ul");
        conditionsList.className = "rule-list";

        if (rule.conditions.subjectContains?.length > 0) {
          const li = document.createElement("li");
          li.textContent = `Subject contains: ${rule.conditions.subjectContains.join(", ")}`;
          conditionsList.appendChild(li);
        }

        if (rule.conditions.fromAddresses?.length > 0) {
          const li = document.createElement("li");
          const addresses = rule.conditions.fromAddresses
            .map((a) => a.emailAddress?.address || "")
            .join(", ");
          li.textContent = `From: ${addresses}`;
          conditionsList.appendChild(li);
        }

        if (rule.conditions.bodyContains?.length > 0) {
          const li = document.createElement("li");
          li.textContent = `Body contains: ${rule.conditions.bodyContains.join(", ")}`;
          conditionsList.appendChild(li);
        }

        conditionsDiv.appendChild(conditionsList);
        ruleCard.appendChild(conditionsDiv);
      }

      // Actions
      if (rule.actions) {
        const actionsDiv = document.createElement("div");
        actionsDiv.className = "rule-actions";

        const actionsTitle = document.createElement("div");
        actionsTitle.textContent = "Then:";
        actionsTitle.className = "rule-section-title";
        actionsDiv.appendChild(actionsTitle);

        const actionsList = document.createElement("ul");
        actionsList.className = "rule-list";

        if (rule.actions.moveToFolder) {
          const li = document.createElement("li");
          li.textContent = `Move to folder`;
          actionsList.appendChild(li);
        }

        if (rule.actions.delete) {
          const li = document.createElement("li");
          li.textContent = "Delete";
          actionsList.appendChild(li);
        }

        if (rule.actions.markAsRead) {
          const li = document.createElement("li");
          li.textContent = "Mark as read";
          actionsList.appendChild(li);
        }

        if (rule.actions.forwardTo?.length > 0) {
          const li = document.createElement("li");
          li.textContent = "Forward to recipients";
          actionsList.appendChild(li);
        }

        actionsDiv.appendChild(actionsList);
        ruleCard.appendChild(actionsDiv);
      }

      modalBody.appendChild(ruleCard);
    });
  }

  modalContent.appendChild(modalHeader);
  modalContent.appendChild(modalBody);
  modal.appendChild(modalContent);

  modal.addEventListener("click", (e) => {
    if (e.target === modal) {
      modal.remove();
    }
  });

  const handleEscape = (e) => {
    if (e.key === "Escape") {
      modal.remove();
      document.removeEventListener("keydown", handleEscape);
    }
  };
  document.addEventListener("keydown", handleEscape);

  modal._cleanupEscape = () => {
    document.removeEventListener("keydown", handleEscape);
  };

  document.body.appendChild(modal);
}

async function showCreateMailRuleModal() {
  let mailFolders = [];
  try {
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/mailFolders",
      {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );
    if (response.ok) {
      const data = await response.json();
      mailFolders = data.value || [];
    }
  } catch (error) {
    console.error("Failed to fetch folders:", error);
  }

  const modal = document.createElement("div");
  modal.className = "modal modal-show modal-align-center";
  modal.id = "createMailRuleModal";

  const modalContent = document.createElement("div");
  modalContent.className = "modal-content max-width-600";

  const modalHeader = document.createElement("div");
  modalHeader.className = "modal-header";

  const title = document.createElement("h2");
  title.textContent = "Create Mail Rule";
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

  const nameGroup = document.createElement("div");
  nameGroup.className = "form-group";
  const nameLabel = document.createElement("label");
  nameLabel.textContent = "Rule Name:";
  nameLabel.className = "form-label";
  const nameInput = document.createElement("input");
  nameInput.type = "text";
  nameInput.id = "ruleNameInput";
  nameInput.placeholder = "e.g., Forward to external email";
  nameInput.className = "form-input";
  nameGroup.appendChild(nameLabel);
  nameGroup.appendChild(nameInput);
  form.appendChild(nameGroup);

  const conditionGroup = document.createElement("div");
  conditionGroup.className = "form-group";
  const conditionLabel = document.createElement("label");
  conditionLabel.textContent = "Condition:";
  conditionLabel.className = "form-label";
  const conditionSelect = document.createElement("select");
  conditionSelect.id = "conditionTypeSelect";
  conditionSelect.className = "form-select";
  conditionSelect.innerHTML = `
    <option value="subjectContains">Subject contains</option>
    <option value="fromAddresses">From addresses</option>
    <option value="bodyContains">Body contains</option>
    <option value="senderContains">Sender contains</option>
  `;
  conditionGroup.appendChild(conditionLabel);
  conditionGroup.appendChild(conditionSelect);
  form.appendChild(conditionGroup);

  const valueGroup = document.createElement("div");
  valueGroup.className = "form-group";
  const valueLabel = document.createElement("label");
  valueLabel.textContent = "Value:";
  valueLabel.className = "form-label";
  const valueInput = document.createElement("input");
  valueInput.type = "text";
  valueInput.id = "conditionValueInput";
  valueInput.placeholder = "Enter value or email address";
  valueInput.className = "form-input";
  valueGroup.appendChild(valueLabel);
  valueGroup.appendChild(valueInput);
  form.appendChild(valueGroup);

  const actionGroup = document.createElement("div");
  actionGroup.className = "form-group";
  const actionLabel = document.createElement("label");
  actionLabel.textContent = "Action:";
  actionLabel.className = "form-label";
  const actionSelect = document.createElement("select");
  actionSelect.id = "actionTypeSelect";
  actionSelect.className = "form-select";
  actionSelect.innerHTML = `
    <option value="forwardTo">Forward to</option>
    <option value="moveToFolder">Move to folder</option>
    <option value="delete">Delete</option>
    <option value="markAsRead">Mark as read</option>
    <option value="copyToFolder">Copy to folder</option>
  `;
  actionGroup.appendChild(actionLabel);
  actionGroup.appendChild(actionSelect);
  form.appendChild(actionGroup);

  const actionValueGroup = document.createElement("div");
  actionValueGroup.className = "form-group";
  actionValueGroup.id = "actionValueGroup";
  const actionValueLabel = document.createElement("label");
  actionValueLabel.textContent = "Forward to email:";
  actionValueLabel.className = "form-label";
  const actionValueInput = document.createElement("input");
  actionValueInput.type = "text";
  actionValueInput.id = "actionValueInput";
  actionValueInput.placeholder = "recipient@example.com";
  actionValueInput.className = "form-input";
  actionValueGroup.appendChild(actionValueLabel);
  actionValueGroup.appendChild(actionValueInput);
  form.appendChild(actionValueGroup);

  const folderSelectGroup = document.createElement("div");
  folderSelectGroup.className = "form-group";
  folderSelectGroup.id = "folderSelectGroup";
  folderSelectGroup.style.display = "none";
  const folderSelectLabel = document.createElement("label");
  folderSelectLabel.textContent = "Select folder:";
  folderSelectLabel.style.fontWeight = "600";
  folderSelectLabel.style.marginBottom = "6px";
  folderSelectLabel.style.display = "block";
  const folderSelect = document.createElement("select");
  folderSelect.id = "folderSelect";
  folderSelect.style.width = "100%";
  folderSelect.style.padding = "8px 12px";
  folderSelect.style.border = "1px solid var(--border-color)";
  folderSelect.style.borderRadius = "6px";
  folderSelect.style.fontSize = "13px";

  mailFolders.forEach((folder) => {
    const option = document.createElement("option");
    option.value = folder.id;
    option.textContent = folder.displayName;
    folderSelect.appendChild(option);
  });

  folderSelectGroup.appendChild(folderSelectLabel);
  folderSelectGroup.appendChild(folderSelect);
  form.appendChild(folderSelectGroup);

  actionSelect.addEventListener("change", () => {
    const actionType = actionSelect.value;
    if (actionType === "forwardTo") {
      actionValueGroup.style.display = "block";
      folderSelectGroup.style.display = "none";
      actionValueLabel.textContent = "Forward to email:";
      actionValueInput.placeholder = "email@example.com";
    } else if (actionType === "moveToFolder" || actionType === "copyToFolder") {
      actionValueGroup.style.display = "none";
      folderSelectGroup.style.display = "block";
    } else {
      actionValueGroup.style.display = "none";
      folderSelectGroup.style.display = "none";
    }
  });

  const enableGroup = document.createElement("div");
  enableGroup.className = "form-group";
  enableGroup.style.marginBottom = "0";
  const enableLabel = document.createElement("label");
  enableLabel.style.display = "flex";
  enableLabel.style.alignItems = "center";
  enableLabel.style.gap = "8px";
  const enableCheckbox = document.createElement("input");
  enableCheckbox.type = "checkbox";
  enableCheckbox.id = "ruleEnabledCheckbox";
  enableCheckbox.checked = true;
  const enableText = document.createElement("span");
  enableText.textContent = "Enable this rule";
  enableLabel.appendChild(enableCheckbox);
  enableLabel.appendChild(enableText);
  enableGroup.appendChild(enableLabel);
  form.appendChild(enableGroup);

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

  const createBtn = document.createElement("button");
  createBtn.className = "btn btn-primary";
  createBtn.textContent = "Create Rule";
  createBtn.addEventListener("click", () => createMailRule(modal));
  modalFooter.appendChild(createBtn);

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

async function createMailRule(modal) {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const ruleName = document.getElementById("ruleNameInput").value.trim();
  const conditionType = document.getElementById("conditionTypeSelect").value;
  const conditionValue = document
    .getElementById("conditionValueInput")
    .value.trim();
  const actionType = document.getElementById("actionTypeSelect").value;
  const actionValue = document.getElementById("actionValueInput").value.trim();
  const folderId = document.getElementById("folderSelect").value;
  const isEnabled = document.getElementById("ruleEnabledCheckbox").checked;

  if (!ruleName) {
    showToast("Please enter a rule name");
    return;
  }

  if (!conditionValue) {
    showToast("Please enter a condition value");
    return;
  }

  if (actionType === "forwardTo" && !actionValue) {
    showToast("Please enter an email address");
    return;
  }

  if (
    (actionType === "moveToFolder" || actionType === "copyToFolder") &&
    !folderId
  ) {
    showToast("Please select a folder");
    return;
  }

  try {
    showToast("Creating mail rule...");

    const rule = {
      displayName: ruleName,
      sequence: 1,
      isEnabled: isEnabled,
      conditions: {},
      actions: {},
    };

    switch (conditionType) {
      case "subjectContains":
        rule.conditions.subjectContains = [conditionValue];
        break;
      case "fromAddresses":
        rule.conditions.fromAddresses = [
          {
            emailAddress: {
              address: conditionValue,
            },
          },
        ];
        break;
      case "bodyContains":
        rule.conditions.bodyContains = [conditionValue];
        break;
      case "senderContains":
        rule.conditions.senderContains = [conditionValue];
        break;
    }

    switch (actionType) {
      case "forwardTo":
        rule.actions.forwardTo = [
          {
            emailAddress: {
              address: actionValue,
            },
          },
        ];
        break;
      case "moveToFolder":
        rule.actions.moveToFolder = folderId;
        break;
      case "delete":
        rule.actions.delete = true;
        break;
      case "markAsRead":
        rule.actions.markAsRead = true;
        break;
      case "copyToFolder":
        rule.actions.copyToFolder = folderId;
        break;
    }

    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(rule),
      },
    );

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error?.message || response.statusText);
    }

    showToast("✅ Mail rule created successfully");
    modal.remove();

    await viewMailRules();
  } catch (error) {
    console.error("Failed to create mail rule:", error);
    showToast("Failed to create mail rule: " + error.message);
  }
}

async function deleteMailRule(ruleId, modal) {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  try {
    showToast("Deleting mail rule...");

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules/${ruleId}`,
      {
        method: "DELETE",
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
        },
      },
    );

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error?.message || response.statusText);
    }

    showToast("✅ Mail rule deleted successfully");
    modal.remove();

    await viewMailRules();
  } catch (error) {
    console.error("Failed to delete mail rule:", error);
    showToast("Failed to delete mail rule: " + error.message);
  }
}

async function viewMailboxSettings() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  try {
    showToast("Loading mailbox settings...");

    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/mailboxSettings",
      {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      throw new Error(
        `Failed to load mailbox settings: ${response.statusText}`,
      );
    }

    const settings = await response.json();
    showMailboxSettingsModal(settings);
  } catch (error) {
    console.error("Failed to load mailbox settings:", error);
    showToast("Failed to load mailbox settings: " + error.message);
  }
}

function showMailboxSettingsModal(settings) {
  const modal = document.createElement("div");
  modal.className = "modal";
  modal.style.display = "flex";
  modal.style.alignItems = "center";
  modal.style.justifyContent = "center";
  modal.id = "mailboxSettingsModal";

  const modalContent = document.createElement("div");
  modalContent.className = "modal-content";
  modalContent.style.maxWidth = "700px";
  modalContent.style.margin = "0";

  const modalHeader = document.createElement("div");
  modalHeader.className = "modal-header";

  const title = document.createElement("h2");
  title.textContent = "Mailbox Settings";
  modalHeader.appendChild(title);

  const closeBtn = document.createElement("button");
  closeBtn.className = "modal-close";
  closeBtn.textContent = "×";
  closeBtn.addEventListener("click", () => modal.remove());
  modalHeader.appendChild(closeBtn);

  const modalBody = document.createElement("div");
  modalBody.className = "modal-body";
  modalBody.style.maxHeight = "500px";
  modalBody.style.overflowY = "auto";

  if (settings.language) {
    const langCard = createSettingCard(
      "Language",
      `Locale: ${settings.language.locale || "Not set"}\nDisplay Name: ${settings.language.displayName || "Not set"}`,
    );
    modalBody.appendChild(langCard);
  }

  // Time zone
  if (settings.timeZone) {
    const tzCard = createSettingCard("Time Zone", settings.timeZone);
    modalBody.appendChild(tzCard);
  }

  // Date format
  if (settings.dateFormat) {
    const dateCard = createSettingCard("Date Format", settings.dateFormat);
    modalBody.appendChild(dateCard);
  }

  // Time format
  if (settings.timeFormat) {
    const timeCard = createSettingCard("Time Format", settings.timeFormat);
    modalBody.appendChild(timeCard);
  }

  // Working hours
  if (settings.workingHours) {
    const workingHoursText = `Days: ${settings.workingHours.daysOfWeek?.join(", ") || "Not set"}\nStart: ${settings.workingHours.startTime || "Not set"}\nEnd: ${settings.workingHours.endTime || "Not set"}\nTime Zone: ${settings.workingHours.timeZone?.name || "Not set"}`;
    const workCard = createSettingCard("Working Hours", workingHoursText);
    modalBody.appendChild(workCard);
  }

  // Delegated access
  if (settings.delegateMeetingMessageDeliveryOptions) {
    const delegateCard = createSettingCard(
      "Delegate Meeting Messages",
      settings.delegateMeetingMessageDeliveryOptions,
    );
    modalBody.appendChild(delegateCard);
  }

  // User purpose
  if (settings.userPurpose) {
    const purposeCard = createSettingCard("User Purpose", settings.userPurpose);
    modalBody.appendChild(purposeCard);
  }

  modalContent.appendChild(modalHeader);
  modalContent.appendChild(modalBody);
  modal.appendChild(modalContent);

  modal.addEventListener("click", (e) => {
    if (e.target === modal) {
      modal.remove();
    }
  });

  const handleEscape = (e) => {
    if (e.key === "Escape") {
      modal.remove();
      document.removeEventListener("keydown", handleEscape);
    }
  };
  document.addEventListener("keydown", handleEscape);

  modal._cleanupEscape = () => {
    document.removeEventListener("keydown", handleEscape);
  };

  document.body.appendChild(modal);
}

function createSettingCard(title, content) {
  const card = document.createElement("div");
  card.style.padding = "12px 16px";
  card.style.marginBottom = "12px";
  card.style.background = "var(--bg-secondary)";
  card.style.borderRadius = "6px";
  card.style.border = "1px solid var(--border-color)";

  const titleEl = document.createElement("strong");
  titleEl.textContent = title;
  titleEl.style.display = "block";
  titleEl.style.marginBottom = "8px";
  titleEl.style.fontSize = "14px";
  titleEl.style.color = "var(--text-color)";
  card.appendChild(titleEl);

  const contentEl = document.createElement("div");
  contentEl.style.fontSize = "13px";
  contentEl.style.color = "var(--text-color)";
  contentEl.style.whiteSpace = "pre-wrap";
  contentEl.textContent = content;
  card.appendChild(contentEl);

  return card;
}

async function viewAutoReply() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  try {
    showToast("Loading auto-reply settings...");

    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/mailboxSettings",
      {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      throw new Error(
        `Failed to load auto-reply settings: ${response.statusText}`,
      );
    }

    const data = await response.json();
    const autoReply = data.automaticRepliesSetting || {};

    showAutoReplyModal(autoReply);
  } catch (error) {
    console.error("Failed to load auto-reply settings:", error);
    showToast("Failed to load auto-reply settings: " + error.message);
  }
}

function showAutoReplyModal(autoReply) {
  const modal = document.createElement("div");
  modal.className = "modal";
  modal.style.display = "flex";
  modal.style.alignItems = "center";
  modal.style.justifyContent = "center";
  modal.id = "autoReplyModal";

  const modalContent = document.createElement("div");
  modalContent.className = "modal-content";
  modalContent.style.maxWidth = "600px";
  modalContent.style.margin = "0";

  const modalHeader = document.createElement("div");
  modalHeader.className = "modal-header";

  const title = document.createElement("h2");
  title.textContent = "Auto Reply (Out of Office)";
  modalHeader.appendChild(title);

  const closeBtn = document.createElement("button");
  closeBtn.className = "modal-close";
  closeBtn.textContent = "×";
  closeBtn.addEventListener("click", () => modal.remove());
  modalHeader.appendChild(closeBtn);

  const modalBody = document.createElement("div");
  modalBody.className = "modal-body";

  // Status
  const statusDiv = document.createElement("div");
  statusDiv.style.padding = "12px";
  statusDiv.style.marginBottom = "16px";
  statusDiv.style.background = "var(--bg-secondary)";
  statusDiv.style.borderRadius = "6px";
  statusDiv.style.border = "1px solid var(--border-color)";

  const statusLabel = document.createElement("strong");
  statusLabel.textContent = "Status: ";
  statusDiv.appendChild(statusLabel);

  const statusText = document.createElement("span");
  const status = autoReply.status || "disabled";
  statusText.textContent = status.charAt(0).toUpperCase() + status.slice(1);
  statusText.style.fontWeight = "700";
  statusText.style.padding = "4px 8px";
  statusText.style.borderRadius = "4px";
  statusText.style.display = "inline-block";
  if (status === "enabled" || status === "scheduled") {
    statusText.style.background = "rgba(93, 216, 196, 0.15)";
    statusText.style.color = "#047857";
  } else {
    statusText.style.background = "var(--bg-tertiary)";
    statusText.style.color = "var(--text-color)";
  }
  statusDiv.appendChild(statusText);

  modalBody.appendChild(statusDiv);

  // Date range if scheduled
  if (
    autoReply.status === "scheduled" &&
    (autoReply.scheduledStartDateTime || autoReply.scheduledEndDateTime)
  ) {
    const dateDiv = document.createElement("div");
    dateDiv.style.padding = "12px";
    dateDiv.style.marginBottom = "16px";
    dateDiv.style.background = "var(--bg-secondary)";
    dateDiv.style.borderRadius = "6px";
    dateDiv.style.border = "1px solid var(--border-color)";

    if (autoReply.scheduledStartDateTime) {
      const startP = document.createElement("p");
      startP.style.marginBottom = "8px";
      startP.innerHTML = `<strong>Start:</strong> ${new Date(autoReply.scheduledStartDateTime.dateTime).toLocaleString()}`;
      dateDiv.appendChild(startP);
    }

    if (autoReply.scheduledEndDateTime) {
      const endP = document.createElement("p");
      endP.style.marginBottom = "0";
      endP.innerHTML = `<strong>End:</strong> ${new Date(autoReply.scheduledEndDateTime.dateTime).toLocaleString()}`;
      dateDiv.appendChild(endP);
    }

    modalBody.appendChild(dateDiv);
  }

  // Internal message
  if (autoReply.internalReplyMessage) {
    const internalDiv = document.createElement("div");
    internalDiv.style.marginBottom = "16px";

    const internalTitle = document.createElement("h3");
    internalTitle.textContent = "Internal Message";
    internalTitle.style.fontSize = "14px";
    internalTitle.style.marginBottom = "8px";
    internalDiv.appendChild(internalTitle);

    const internalMsg = document.createElement("div");
    internalMsg.style.padding = "12px";
    internalMsg.style.background = "var(--bg-secondary)";
    internalMsg.style.borderRadius = "6px";
    internalMsg.style.border = "1px solid var(--border-color)";
    internalMsg.style.maxHeight = "200px";
    internalMsg.style.overflowY = "auto";

    const internalIframe = document.createElement("iframe");
    internalIframe.sandbox = "";
    internalIframe.style.width = "100%";
    internalIframe.style.border = "none";
    internalIframe.style.background = "white";
    internalIframe.srcdoc = autoReply.internalReplyMessage;

    internalIframe.onload = () => {
      try {
        const iframeDoc =
          internalIframe.contentDocument ||
          internalIframe.contentWindow.document;
        const height = iframeDoc.documentElement.scrollHeight;
        internalIframe.style.height = height + "px";
      } catch (e) {
        internalIframe.style.height = "150px";
      }
    };

    internalMsg.appendChild(internalIframe);
    internalDiv.appendChild(internalMsg);

    modalBody.appendChild(internalDiv);
  }

  // External message
  if (autoReply.externalReplyMessage) {
    const externalDiv = document.createElement("div");

    const externalTitle = document.createElement("h3");
    externalTitle.textContent = "External Message";
    externalTitle.style.fontSize = "14px";
    externalTitle.style.marginBottom = "8px";
    externalDiv.appendChild(externalTitle);

    const externalMsg = document.createElement("div");
    externalMsg.style.padding = "12px";
    externalMsg.style.background = "var(--bg-secondary)";
    externalMsg.style.borderRadius = "6px";
    externalMsg.style.border = "1px solid var(--border-color)";
    externalMsg.style.maxHeight = "200px";
    externalMsg.style.overflowY = "auto";

    const externalIframe = document.createElement("iframe");
    externalIframe.sandbox = "";
    externalIframe.style.width = "100%";
    externalIframe.style.border = "none";
    externalIframe.style.background = "white";
    externalIframe.srcdoc = autoReply.externalReplyMessage;

    externalIframe.onload = () => {
      try {
        const iframeDoc =
          externalIframe.contentDocument ||
          externalIframe.contentWindow.document;
        const height = iframeDoc.documentElement.scrollHeight;
        externalIframe.style.height = height + "px";
      } catch (e) {
        externalIframe.style.height = "150px";
      }
    };

    externalMsg.appendChild(externalIframe);
    externalDiv.appendChild(externalMsg);

    modalBody.appendChild(externalDiv);
  }

  // If no messages
  if (!autoReply.internalReplyMessage && !autoReply.externalReplyMessage) {
    const emptyMsg = document.createElement("p");
    emptyMsg.textContent = "No auto-reply messages configured";
    emptyMsg.style.color = "var(--text-color)";
    emptyMsg.style.textAlign = "center";
    emptyMsg.style.padding = "20px";
    modalBody.appendChild(emptyMsg);
  }

  modalContent.appendChild(modalHeader);
  modalContent.appendChild(modalBody);
  modal.appendChild(modalContent);

  modal.addEventListener("click", (e) => {
    if (e.target === modal) {
      modal.remove();
    }
  });

  const handleEscape = (e) => {
    if (e.key === "Escape") {
      modal.remove();
      document.removeEventListener("keydown", handleEscape);
    }
  };
  document.addEventListener("keydown", handleEscape);

  modal._cleanupEscape = () => {
    document.removeEventListener("keydown", handleEscape);
  };

  document.body.appendChild(modal);
}
