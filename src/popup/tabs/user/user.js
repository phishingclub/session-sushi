let currentUserSection = "profile";
let currentUserProfile = null;
let mfaMethods = [];

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
        "max-width: 200px; max-height: 200px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1)";
      img.alt = "Profile Photo";

      const downloadBtn = document.createElement("button");
      downloadBtn.className = "btn btn-sm btn-secondary";
      downloadBtn.textContent = "Download Photo";
      downloadBtn.addEventListener("click", () => {
        const a = document.createElement("a");
        a.href = imageUrl;
        a.download = "profile-photo.jpg";
        a.click();
      });

      photoCard.appendChild(img);
      photoCard.appendChild(downloadBtn);
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
}
