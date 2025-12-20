let currentUserSection = "profile";
let currentUserProfile = null;
let mfaMethods = [];
let userTasks = {
  planner: [],
  todo: [],
};
let currentTaskFilter = "all";

async function initializeUser() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showUserNoSession();
    return;
  }

  await loadUserProfile();
  await loadUserSection("profile");
}

function showUserNoSession() {
  const sections = [
    "userProfileContainer",
    "mfaMethodsContainer",
    "authenticationContainer",
    "permissionsContainer",
    "tasksContainer",
  ];
  sections.forEach((sectionId) => {
    const container = document.getElementById(sectionId);
    if (container) {
      container.textContent = "";
      const emptyDiv = document.createElement("div");
      emptyDiv.className = "mailbox-empty";
      emptyDiv.textContent = "No active session";
      container.appendChild(emptyDiv);
    }
  });
}

async function loadUserProfile() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  try {
    const url =
      "https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones,accountEnabled,createdDateTime,usageLocation,employeeId,employeeType,companyName,preferredLanguage,ageGroup,consentProvidedForMinor,legalAgeGroupClassification,city,country,postalCode,state,streetAddress,surname,givenName,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesDomainName,onPremisesSamAccountName,onPremisesSecurityIdentifier,proxyAddresses,assignedLicenses,licenseAssignmentStates,lastPasswordChangeDateTime,passwordPolicies,authorizationInfo,externalUserState,externalUserStateChangeDateTime,userType,employeeHireDate,birthday,aboutMe,mySite,interests,skills,responsibilities,schools,mailNickname,otherMails,faxNumber,imAddresses";

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(
        errorData.error?.message || "Failed to load user profile",
      );
    }

    currentUserProfile = await response.json();
    return currentUserProfile;
  } catch (error) {
    console.error("Error loading user profile:", error);
    showToast(error.message || "Failed to load user profile");
    throw error;
  }
}

async function loadUserSection(section) {
  if (!activeM365Session || !activeM365Session.access_token) {
    showUserNoSession();
    return;
  }

  currentUserSection = section;

  document.querySelectorAll(".user-section-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.section === section);
  });

  document.querySelectorAll(".user-section").forEach((sec) => {
    sec.style.display = "none";
  });

  const sectionElement = document.getElementById(
    `user${section.charAt(0).toUpperCase() + section.slice(1)}Section`,
  );
  if (sectionElement) {
    sectionElement.style.display = "block";
  }

  switch (section) {
    case "profile":
      await renderUserProfile();
      break;
    case "authentication":
      await loadAuthenticationSettings();
      break;
    case "permissions":
      await loadUserPermissions();
      break;
    case "tasks":
      await loadUserTasks();
      break;
  }
}

async function renderUserProfile() {
  const container = document.getElementById("userProfileContainer");
  if (!container) return;

  container.textContent = "";

  if (!currentUserProfile) {
    try {
      await loadUserProfile();
    } catch (error) {
      const errorDiv = document.createElement("div");
      errorDiv.className = "mailbox-empty";
      errorDiv.textContent = "Failed to load user profile";
      container.appendChild(errorDiv);
      return;
    }
  }

  // Load profile photo
  try {
    const photoUrl = "https://graph.microsoft.com/v1.0/me/photo/$value";
    const photoResponse = await fetch(photoUrl, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
      },
    });

    if (photoResponse.ok) {
      const blob = await photoResponse.blob();
      const imageUrl = URL.createObjectURL(blob);

      const photoCard = document.createElement("div");
      photoCard.style.cssText =
        "background: var(--bg-secondary); padding: 20px; border-radius: 8px; margin-bottom: 20px; display: flex; flex-direction: column; align-items: center; gap: 15px";

      const img = document.createElement("img");
      img.src = imageUrl;
      img.style.cssText =
        "max-width: 100px; max-height: 100px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1)";
      img.alt = "Profile Photo";

      photoCard.appendChild(img);
      container.appendChild(photoCard);
    }
  } catch (error) {
    console.log("Profile photo not available:", error.message);
  }

  const profileCard = document.createElement("div");
  profileCard.style.cssText =
    "background: var(--bg-secondary); padding: 20px; border-radius: 8px; margin-bottom: 20px";

  const fields = [
    { label: "Display Name", value: currentUserProfile.displayName },
    { label: "Given Name", value: currentUserProfile.givenName },
    { label: "Surname", value: currentUserProfile.surname },
    {
      label: "Email",
      value: currentUserProfile.mail || currentUserProfile.userPrincipalName,
    },
    {
      label: "User Principal Name",
      value: currentUserProfile.userPrincipalName,
    },
    { label: "Mail Nickname", value: currentUserProfile.mailNickname },
    {
      label: "Other Emails",
      value: currentUserProfile.otherMails?.join(", "),
    },
    {
      label: "Proxy Addresses",
      value: currentUserProfile.proxyAddresses?.join(", "),
    },
    { label: "User Type", value: currentUserProfile.userType },
    { label: "Job Title", value: currentUserProfile.jobTitle },
    { label: "Department", value: currentUserProfile.department },
    { label: "Company Name", value: currentUserProfile.companyName },
    { label: "Employee ID", value: currentUserProfile.employeeId },
    { label: "Employee Type", value: currentUserProfile.employeeType },
    {
      label: "Employee Hire Date",
      value: currentUserProfile.employeeHireDate
        ? new Date(currentUserProfile.employeeHireDate).toLocaleDateString()
        : null,
    },
    {
      label: "Birthday",
      value: currentUserProfile.birthday
        ? new Date(currentUserProfile.birthday).toLocaleDateString()
        : null,
    },
    { label: "Office Location", value: currentUserProfile.officeLocation },
    { label: "City", value: currentUserProfile.city },
    { label: "State", value: currentUserProfile.state },
    { label: "Country", value: currentUserProfile.country },
    { label: "Postal Code", value: currentUserProfile.postalCode },
    { label: "Street Address", value: currentUserProfile.streetAddress },
    { label: "Mobile Phone", value: currentUserProfile.mobilePhone },
    {
      label: "Business Phones",
      value: currentUserProfile.businessPhones?.join(", "),
    },
    { label: "Fax Number", value: currentUserProfile.faxNumber },
    {
      label: "IM Addresses",
      value: currentUserProfile.imAddresses?.join(", "),
    },
    {
      label: "Preferred Language",
      value: currentUserProfile.preferredLanguage,
    },
    { label: "Usage Location", value: currentUserProfile.usageLocation },
    {
      label: "Account Enabled",
      value: currentUserProfile.accountEnabled ? "Yes" : "No",
    },
    {
      label: "Account Created",
      value: currentUserProfile.createdDateTime
        ? new Date(currentUserProfile.createdDateTime).toLocaleString()
        : null,
    },
    {
      label: "Last Password Change",
      value: currentUserProfile.lastPasswordChangeDateTime
        ? new Date(
            currentUserProfile.lastPasswordChangeDateTime,
          ).toLocaleString()
        : null,
    },
    { label: "Password Policies", value: currentUserProfile.passwordPolicies },
    {
      label: "On-Premises Sync Enabled",
      value:
        currentUserProfile.onPremisesSyncEnabled != null
          ? currentUserProfile.onPremisesSyncEnabled
            ? "Yes"
            : "No"
          : null,
    },
    {
      label: "On-Premises Last Sync",
      value: currentUserProfile.onPremisesLastSyncDateTime
        ? new Date(
            currentUserProfile.onPremisesLastSyncDateTime,
          ).toLocaleString()
        : null,
    },
    {
      label: "On-Premises Domain",
      value: currentUserProfile.onPremisesDomainName,
    },
    {
      label: "On-Premises SAM Account",
      value: currentUserProfile.onPremisesSamAccountName,
    },
    {
      label: "On-Premises SID",
      value: currentUserProfile.onPremisesSecurityIdentifier,
    },
    {
      label: "External User State",
      value: currentUserProfile.externalUserState,
    },
    {
      label: "External User State Changed",
      value: currentUserProfile.externalUserStateChangeDateTime
        ? new Date(
            currentUserProfile.externalUserStateChangeDateTime,
          ).toLocaleString()
        : null,
    },
    { label: "Age Group", value: currentUserProfile.ageGroup },
    {
      label: "Consent Provided For Minor",
      value: currentUserProfile.consentProvidedForMinor,
    },
    {
      label: "Legal Age Group Classification",
      value: currentUserProfile.legalAgeGroupClassification,
    },
    { label: "About Me", value: currentUserProfile.aboutMe },
    { label: "My Site", value: currentUserProfile.mySite },
    {
      label: "Interests",
      value: currentUserProfile.interests?.join(", "),
    },
    {
      label: "Skills",
      value: currentUserProfile.skills?.join(", "),
    },
    {
      label: "Responsibilities",
      value: currentUserProfile.responsibilities?.join(", "),
    },
    {
      label: "Schools",
      value: currentUserProfile.schools?.join(", "),
    },
    {
      label: "Assigned Licenses",
      value: currentUserProfile.assignedLicenses
        ?.map((license) => license.skuId)
        .join(", "),
    },
    {
      label: "Certificate User IDs",
      value:
        currentUserProfile.authorizationInfo?.certificateUserIds?.join(", "),
    },
    { label: "User ID", value: currentUserProfile.id },
  ];

  fields.forEach((field) => {
    if (field.value) {
      const fieldDiv = document.createElement("div");
      fieldDiv.style.cssText =
        "margin-bottom: 12px; padding-bottom: 12px; border-bottom: 1px solid var(--border-color)";

      const labelDiv = document.createElement("div");
      labelDiv.style.cssText =
        "font-size: 12px; color: var(--text-secondary); margin-bottom: 4px; font-weight: 600";
      labelDiv.textContent = field.label;

      const valueDiv = document.createElement("div");
      valueDiv.style.cssText = "font-size: 14px; word-break: break-all";
      valueDiv.textContent = field.value;

      fieldDiv.appendChild(labelDiv);
      fieldDiv.appendChild(valueDiv);
      profileCard.appendChild(fieldDiv);
    }
  });

  container.appendChild(profileCard);
}

async function loadAuthenticationSettings() {
  const container = document.getElementById("authenticationContainer");
  if (!container) return;

  container.textContent = "";
  const loadingDiv = document.createElement("div");
  loadingDiv.className = "loading-indicator";
  loadingDiv.textContent = "Loading authentication settings...";
  container.appendChild(loadingDiv);

  try {
    const url =
      "https://graph.microsoft.com/v1.0/me?$select=passwordProfile,passwordPolicies,lastPasswordChangeDateTime";

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(
        errorData.error?.message || "Failed to load authentication settings",
      );
    }

    const data = await response.json();

    container.textContent = "";

    const settingsCard = document.createElement("div");
    settingsCard.style.cssText =
      "background: var(--bg-secondary); padding: 20px; border-radius: 8px";

    const fields = [
      { label: "Password Policies", value: data.passwordPolicies || "Default" },
      {
        label: "Last Password Change",
        value: data.lastPasswordChangeDateTime
          ? new Date(data.lastPasswordChangeDateTime).toLocaleString()
          : "N/A",
      },
    ];

    fields.forEach((field) => {
      const fieldDiv = document.createElement("div");
      fieldDiv.style.cssText =
        "margin-bottom: 12px; padding-bottom: 12px; border-bottom: 1px solid var(--border-color)";

      const labelDiv = document.createElement("div");
      labelDiv.style.cssText =
        "font-size: 12px; color: var(--text-secondary); margin-bottom: 4px; font-weight: 600";
      labelDiv.textContent = field.label;

      const valueDiv = document.createElement("div");
      valueDiv.style.cssText = "font-size: 14px";
      valueDiv.textContent = field.value;

      fieldDiv.appendChild(labelDiv);
      fieldDiv.appendChild(valueDiv);
      settingsCard.appendChild(fieldDiv);
    });

    container.appendChild(settingsCard);
  } catch (error) {
    console.error("Error loading authentication settings:", error);
    container.textContent = "";
    const errorDiv = document.createElement("div");
    errorDiv.className = "mailbox-empty";
    errorDiv.textContent =
      error.message || "Failed to load authentication settings";
    container.appendChild(errorDiv);
  }
}

async function loadUserPermissions() {
  const container = document.getElementById("permissionsContainer");
  if (!container) return;

  container.textContent = "";
  const loadingDiv = document.createElement("div");
  loadingDiv.className = "loading-indicator";
  loadingDiv.textContent = "Loading permissions...";
  container.appendChild(loadingDiv);

  try {
    const url =
      "https://graph.microsoft.com/v1.0/me/memberOf?$select=id,displayName,description";

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error?.message || "Failed to load permissions");
    }

    const data = await response.json();
    const groups = data.value || [];

    container.textContent = "";

    if (groups.length === 0) {
      const emptyDiv = document.createElement("div");
      emptyDiv.className = "mailbox-empty";
      emptyDiv.textContent = "No group memberships found";
      container.appendChild(emptyDiv);
      return;
    }

    const groupsContainer = document.createElement("div");
    groupsContainer.style.cssText =
      "display: flex; flex-direction: column; gap: 12px";

    groups.forEach((group) => {
      const groupCard = document.createElement("div");
      groupCard.style.cssText =
        "background: var(--bg-secondary); padding: 15px; border-radius: 8px; border: 1px solid var(--border-color)";

      const nameDiv = document.createElement("div");
      nameDiv.style.cssText = "font-weight: 600; margin-bottom: 8px";
      nameDiv.textContent = group.displayName;

      const descDiv = document.createElement("div");
      descDiv.style.cssText =
        "font-size: 13px; color: var(--text-secondary); margin-bottom: 8px";
      descDiv.textContent = group.description || "No description";

      const idDiv = document.createElement("div");
      idDiv.style.cssText =
        "font-size: 12px; color: var(--text-secondary); font-family: monospace";
      idDiv.textContent = `ID: ${group.id}`;

      groupCard.appendChild(nameDiv);
      groupCard.appendChild(descDiv);
      groupCard.appendChild(idDiv);
      groupsContainer.appendChild(groupCard);
    });

    container.appendChild(groupsContainer);
  } catch (error) {
    console.error("Error loading permissions:", error);
    container.textContent = "";
    const errorDiv = document.createElement("div");
    errorDiv.className = "mailbox-empty";
    errorDiv.textContent = error.message || "Failed to load permissions";
    container.appendChild(errorDiv);
  }
}

async function loadUserTasks() {
  const container = document.getElementById("tasksContainer");
  if (!container) return;

  container.textContent = "";
  const loadingDiv = document.createElement("div");
  loadingDiv.className = "loading-indicator";
  loadingDiv.textContent = "Loading tasks...";
  container.appendChild(loadingDiv);

  userTasks = {
    planner: [],
    todo: [],
  };

  await Promise.allSettled([loadPlannerTasks(), loadTodoTasks()]);

  renderUserTasks();
}

async function loadPlannerTasks() {
  try {
    const url = "https://graph.microsoft.com/v1.0/me/planner/tasks";

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (response.ok) {
      const data = await response.json();
      userTasks.planner = data.value || [];
    }
  } catch (error) {
    console.error("Error loading Planner tasks:", error);
  }
}

async function loadTodoTasks() {
  try {
    const listsUrl = "https://graph.microsoft.com/v1.0/me/todo/lists";
    const listsResponse = await fetch(listsUrl, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!listsResponse.ok) return;

    const listsData = await listsResponse.json();
    const lists = listsData.value || [];

    for (const list of lists) {
      try {
        const tasksUrl = `https://graph.microsoft.com/v1.0/me/todo/lists/${list.id}/tasks`;
        const tasksResponse = await fetch(tasksUrl, {
          headers: {
            Authorization: `Bearer ${activeM365Session.access_token}`,
            "Content-Type": "application/json",
          },
        });

        if (tasksResponse.ok) {
          const tasksData = await tasksResponse.json();
          const tasks = tasksData.value || [];
          tasks.forEach((task) => {
            task.listName = list.displayName;
            task.listId = list.id;
          });
          userTasks.todo.push(...tasks);
        }
      } catch (error) {
        console.error(`Error loading tasks for list ${list.id}:`, error);
      }
    }
  } catch (error) {
    console.error("Error loading To-Do tasks:", error);
  }
}

function renderUserTasks() {
  const container = document.getElementById("tasksContainer");
  if (!container) return;

  container.textContent = "";

  let allTasks = [];

  if (currentTaskFilter === "all" || currentTaskFilter === "planner") {
    userTasks.planner.forEach((task) => {
      allTasks.push({
        type: "planner",
        task: task,
        date: task.dueDateTime || task.createdDateTime,
      });
    });
  }

  if (currentTaskFilter === "all" || currentTaskFilter === "todo") {
    userTasks.todo.forEach((task) => {
      allTasks.push({
        type: "todo",
        task: task,
        date: task.dueDateTime || task.createdDateTime,
      });
    });
  }

  const statsBar = document.getElementById("tasksStatsBar");
  const statsSpan = document.getElementById("tasksStats");

  if (statsBar && statsSpan) {
    if (allTasks.length > 0) {
      let statsText = `${allTasks.length} task${allTasks.length !== 1 ? "s" : ""}`;

      if (currentTaskFilter === "all") {
        statsText += ` (${userTasks.planner.length} Planner, ${userTasks.todo.length} To-Do)`;
      }

      statsSpan.textContent = statsText;
      statsBar.style.display = "block";
    } else {
      statsBar.style.display = "none";
    }
  }

  if (allTasks.length === 0) {
    const emptyDiv = document.createElement("div");
    emptyDiv.className = "mailbox-empty";
    emptyDiv.textContent = "No tasks found";
    container.appendChild(emptyDiv);
    return;
  }

  allTasks.sort((a, b) => {
    if (!a.date) return 1;
    if (!b.date) return -1;
    return new Date(a.date) - new Date(b.date);
  });

  const itemsContainer = document.createElement("div");
  itemsContainer.className = "onedrive-items-container";

  allTasks.forEach((item) => {
    const taskElement = createTaskElement(item.task, item.type);
    itemsContainer.appendChild(taskElement);
  });

  container.appendChild(itemsContainer);
}

function createTaskElement(task, type) {
  const taskDiv = document.createElement("div");
  taskDiv.className = "onedrive-item";

  const iconDiv = document.createElement("div");
  iconDiv.className = "onedrive-item-icon";
  const typeIcon = type === "planner" ? "📋" : "✓";
  iconDiv.textContent = typeIcon;

  const detailsDiv = document.createElement("div");
  detailsDiv.className = "onedrive-item-details";

  const nameDiv = document.createElement("div");
  nameDiv.className = "onedrive-item-name";
  nameDiv.textContent = task.title || task.subject || "Untitled";

  const metaParts = [];

  let statusText = "";
  if (type === "planner") {
    const percent = task.percentComplete || 0;
    statusText = `${percent}% complete`;
  } else if (type === "todo") {
    statusText = task.status === "completed" ? "✓ Completed" : "○ Not started";
  }
  metaParts.push(statusText);

  if (task.importance === "high" || task.priority === 1) {
    metaParts.push("🔴 High priority");
  }

  if (task.dueDateTime) {
    const dueDate = new Date(task.dueDateTime.dateTime || task.dueDateTime);
    if (!isNaN(dueDate.getTime())) {
      const now = new Date();
      const isOverdue =
        dueDate < now &&
        (task.percentComplete || 0) < 100 &&
        task.status !== "completed";
      metaParts.push(
        `Due: ${dueDate.toLocaleDateString()}${isOverdue ? " (overdue)" : ""}`,
      );
    }
  }

  if (type === "todo" && task.listName) {
    metaParts.push(`List: ${task.listName}`);
  }

  const metaDiv = document.createElement("div");
  metaDiv.className = "onedrive-item-meta";
  metaDiv.textContent = metaParts.join(" • ");

  detailsDiv.appendChild(nameDiv);
  detailsDiv.appendChild(metaDiv);

  const actionsDiv = document.createElement("div");
  actionsDiv.className = "onedrive-item-actions";

  const detailsBtn = document.createElement("button");
  detailsBtn.className = "btn btn-small btn-secondary btn-compact";
  detailsBtn.textContent = "ℹ️ Details";
  detailsBtn.onclick = () => showTaskDetails(task, type);
  actionsDiv.appendChild(detailsBtn);

  taskDiv.appendChild(iconDiv);
  taskDiv.appendChild(detailsDiv);
  taskDiv.appendChild(actionsDiv);

  return taskDiv;
}

function copyTaskDetails(task, type) {
  const title = task.title || task.subject || "Untitled";
  let details = `Task: ${title}\n`;
  details += `Type: ${type}\n`;

  if (task.percentComplete !== undefined) {
    details += `Progress: ${task.percentComplete}%\n`;
  }

  if (task.status) {
    details += `Status: ${task.status}\n`;
  }

  if (task.importance) {
    details += `Importance: ${task.importance}\n`;
  }

  if (task.priority) {
    details += `Priority: ${task.priority}\n`;
  }

  if (task.dueDateTime) {
    const dueDate = new Date(task.dueDateTime.dateTime || task.dueDateTime);
    if (!isNaN(dueDate.getTime())) {
      details += `Due: ${dueDate.toLocaleString()}\n`;
    }
  }

  if (task.startDateTime) {
    const startDate = new Date(
      task.startDateTime.dateTime || task.startDateTime,
    );
    if (!isNaN(startDate.getTime())) {
      details += `Start: ${startDate.toLocaleString()}\n`;
    }
  }

  if (task.createdDateTime) {
    const createdDate = new Date(task.createdDateTime);
    if (!isNaN(createdDate.getTime())) {
      details += `Created: ${createdDate.toLocaleString()}\n`;
    }
  }

  if (task.listName) {
    details += `List: ${task.listName}\n`;
  }

  if (task.categories && task.categories.length > 0) {
    details += `Categories: ${task.categories.join(", ")}\n`;
  }

  copyToClipboard(details);
  showToast("Task details copied");
}

function showTaskDetails(task, type) {
  const modal = document.createElement("div");
  modal.className = "modal modal-show";

  const modalContent = document.createElement("div");
  modalContent.className = "modal-content max-width-700";

  const modalHeader = document.createElement("div");
  modalHeader.className = "modal-header";

  const title = document.createElement("h2");
  title.textContent = task.title || task.subject || "Task Details";
  modalHeader.appendChild(title);

  const closeBtn = document.createElement("button");
  closeBtn.className = "modal-close";
  closeBtn.innerHTML = "&times;";
  closeBtn.onclick = () => {
    modal.classList.remove("modal-show");
    setTimeout(() => document.body.removeChild(modal), 300);
  };
  modalHeader.appendChild(closeBtn);

  const modalBody = document.createElement("div");
  modalBody.className = "modal-body";

  const detailsContainer = document.createElement("div");
  detailsContainer.style.cssText =
    "display: flex; flex-direction: column; gap: 15px;";

  const fields = [];

  fields.push({ label: "Type", value: type });

  if (task.id) {
    fields.push({ label: "ID", value: task.id });
  }

  if (task.percentComplete !== undefined) {
    fields.push({ label: "Progress", value: `${task.percentComplete}%` });
  }

  if (task.status) {
    fields.push({ label: "Status", value: task.status });
  }

  if (task.importance) {
    fields.push({ label: "Importance", value: task.importance });
  }

  if (task.priority !== undefined) {
    fields.push({ label: "Priority", value: task.priority });
  }

  if (task.dueDateTime) {
    const dueDate = new Date(task.dueDateTime.dateTime || task.dueDateTime);
    if (!isNaN(dueDate.getTime())) {
      fields.push({
        label: "Due Date",
        value: dueDate.toLocaleString(),
      });
    }
  }

  if (task.startDateTime) {
    const startDate = new Date(
      task.startDateTime.dateTime || task.startDateTime,
    );
    if (!isNaN(startDate.getTime())) {
      fields.push({
        label: "Start Date",
        value: startDate.toLocaleString(),
      });
    }
  }

  if (task.createdDateTime) {
    const createdDate = new Date(task.createdDateTime);
    if (!isNaN(createdDate.getTime())) {
      fields.push({
        label: "Created",
        value: createdDate.toLocaleString(),
      });
    }
  }

  if (task.lastModifiedDateTime) {
    const modifiedDate = new Date(task.lastModifiedDateTime);
    if (!isNaN(modifiedDate.getTime())) {
      fields.push({
        label: "Last Modified",
        value: modifiedDate.toLocaleString(),
      });
    }
  }

  if (task.listName) {
    fields.push({ label: "List", value: task.listName });
  }

  if (task.listId) {
    fields.push({ label: "List ID", value: task.listId });
  }

  if (task.bucketId) {
    fields.push({ label: "Bucket ID", value: task.bucketId });
  }

  if (task.planId) {
    fields.push({ label: "Plan ID", value: task.planId });
  }

  if (task.categories && task.categories.length > 0) {
    fields.push({ label: "Categories", value: task.categories.join(", ") });
  }

  if (task.assignments && Object.keys(task.assignments).length > 0) {
    fields.push({
      label: "Assignments",
      value: `${Object.keys(task.assignments).length} user(s)`,
    });
  }

  fields.forEach((field) => {
    const fieldDiv = document.createElement("div");
    fieldDiv.style.cssText =
      "background: var(--bg-secondary); padding: 12px; border-radius: 6px;";

    const labelDiv = document.createElement("div");
    labelDiv.style.cssText =
      "font-size: 12px; color: var(--text-secondary); margin-bottom: 4px; font-weight: 600;";
    labelDiv.textContent = field.label;

    const valueDiv = document.createElement("div");
    valueDiv.style.cssText = "font-size: 14px; word-break: break-all;";
    valueDiv.textContent = field.value;

    fieldDiv.appendChild(labelDiv);
    fieldDiv.appendChild(valueDiv);
    detailsContainer.appendChild(fieldDiv);
  });

  modalBody.appendChild(detailsContainer);
  modalContent.appendChild(modalHeader);
  modalContent.appendChild(modalBody);
  modal.appendChild(modalContent);
  document.body.appendChild(modal);

  modal.addEventListener("click", (e) => {
    if (e.target === modal) {
      modal.classList.remove("modal-show");
      setTimeout(() => document.body.removeChild(modal), 300);
    }
  });
}

function setupUserListeners() {
  const refreshBtn = document.getElementById("refreshUserBtn");
  if (refreshBtn) {
    refreshBtn.addEventListener("click", async () => {
      currentUserProfile = null;
      await loadUserProfile();
      await loadUserSection(currentUserSection);
    });
  }

  const sectionBtns = document.querySelectorAll(".user-section-btn");
  sectionBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      const section = btn.dataset.section;
      loadUserSection(section);
    });
  });

  const refreshTasksBtn = document.getElementById("refreshTasksBtn");
  if (refreshTasksBtn) {
    refreshTasksBtn.addEventListener("click", async () => {
      await loadUserTasks();
    });
  }

  const taskTypeBtns = document.querySelectorAll(".task-type-btn");
  taskTypeBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      currentTaskFilter = btn.dataset.type;
      taskTypeBtns.forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      renderUserTasks();
    });
  });
}
