function setupQuickActions() {
  const categories = document.querySelectorAll(".qa-category");

  categories.forEach((category) => {
    const label = category.querySelector(".qa-label");
    const buttonsContainer = category.querySelector(".qa-buttons");
    const buttons = buttonsContainer
      ? buttonsContainer.querySelectorAll("button")
      : [];

    if (label) {
      label.className = "quick-action-label";
      label.innerHTML =
        label.textContent + " <span class='arrow-indicator'>▶</span>";

      label.addEventListener("click", () => {
        if (buttonsContainer) {
          const isVisible =
            buttonsContainer.classList.contains("hidden") === false &&
            window.getComputedStyle(buttonsContainer).display !== "none";
          if (isVisible) {
            buttonsContainer.classList.add("hidden");
          } else {
            buttonsContainer.classList.remove("hidden");
          }
          const arrow = isVisible ? "▶" : "◀";
          const baseText = label.textContent.replace(/\s*[▶◀]\s*$/, "");
          label.innerHTML =
            baseText + " <span class='arrow-indicator'>" + arrow + "</span>";
        }
      });
    }

    buttons.forEach((btn) => {
      btn.classList.add("margin-right-4", "margin-bottom-4");
      btn.addEventListener("click", () => {
        const queryType = btn.getAttribute("data-query");
        if (queryType) {
          fillQueryField(queryType);
        }
      });
    });
  });
}

function fillQueryField(queryType) {
  const queries = {
    // Credential Hunt
    searchPasswordEmails:
      '/me/messages?$search="password OR credentials OR username OR login OR reset"&$top=100&$select=subject,from,receivedDateTime,bodyPreview,hasAttachments',
    searchCredFiles:
      "/me/drive/root/search(q='password OR credentials OR .env OR secrets OR auth')?$top=100&$select=name,webUrl,createdDateTime,lastModifiedDateTime,size",
    searchAPIKeys:
      '/me/messages?$search="api_key OR apikey OR access_token OR bearer OR secret_key"&$top=100&$select=subject,from,receivedDateTime,bodyPreview',
    searchSSHKeys:
      "/me/drive/root/search(q='.pem OR .key OR .ppk OR id_rsa OR ssh')?$top=100&$select=name,webUrl,size,createdDateTime",
    searchVPNConfigs:
      "/me/drive/root/search(q='.ovpn OR vpn OR wireguard OR openvpn')?$top=100&$select=name,webUrl,size,lastModifiedDateTime",
    searchCertificates:
      "/me/drive/root/search(q='.pfx OR .p12 OR .cer OR .crt OR certificate')?$top=100&$select=name,webUrl,size,createdDateTime",

    // Sensitive Files
    searchConfidential:
      "/me/drive/root/search(q='confidential OR sensitive OR secret OR restricted OR private')?$top=100&$select=name,webUrl,size,lastModifiedDateTime",
    searchDatabaseFiles:
      "/me/drive/root/search(q='.db OR .sqlite OR .mdb OR .sql OR database')?$top=100&$select=name,webUrl,size,lastModifiedDateTime",
    searchBackupFiles:
      "/me/drive/root/search(q='.bak OR .backup OR backup OR dump OR export')?$top=100&$select=name,webUrl,size,createdDateTime",
    searchConfigFiles:
      "/me/drive/root/search(q='config OR .conf OR .ini OR .yaml OR .yml OR .json')?$top=100&$select=name,webUrl,size,lastModifiedDateTime",
    searchScripts:
      "/me/drive/root/search(q='.ps1 OR .sh OR .bat OR .cmd OR script')?$top=100&$select=name,webUrl,size,lastModifiedDateTime",
    searchCompressedFiles:
      "/me/drive/root/search(q='.zip OR .rar OR .7z OR .tar OR .gz')?$top=100&$orderby=size desc&$select=name,webUrl,size,lastModifiedDateTime",

    // Financial/PII
    searchFinancial:
      "/me/drive/root/search(q='financial OR budget OR invoice OR billing OR payment')?$top=100&$select=name,webUrl,size,lastModifiedDateTime",
    searchPayroll:
      "/me/drive/root/search(q='payroll OR salary OR compensation OR wages')?$top=100&$select=name,webUrl,size,lastModifiedDateTime",
    searchContracts:
      "/me/drive/root/search(q='contract OR agreement OR NDA OR terms')?$top=100&$select=name,webUrl,size,createdDateTime",
    searchPII:
      '/me/messages?$search="SSN OR social security OR credit card OR passport OR driver license"&$top=100&$select=subject,from,receivedDateTime,bodyPreview',

    // Recon
    directoryRoles:
      "/me/memberOf/microsoft.graph.directoryRole?$select=displayName,description,roleTemplateId",
    privilegedRoleMembers:
      "/directoryRoles?$filter=displayName eq 'Global Administrator' or displayName eq 'Privileged Role Administrator'&$expand=members",
    conditionalAccess: "/identity/conditionalAccess/policies",
    domainInfo: "/domains?$select=id,isDefault,isVerified,authenticationType",
    tenantDetails:
      "/organization?$select=displayName,verifiedDomains,technicalNotificationMails,securityComplianceNotificationMails",

    // Persistence
    oauth2Grants:
      "/me/oauth2PermissionGrants?$select=clientId,consentType,principalId,resourceId,scope",
    appPermissions:
      "/servicePrincipals?$select=appRoles,oauth2PermissionScopes,appId,displayName",
    delegatedPermissions:
      "/oauth2PermissionGrants?$top=999&$select=clientId,consentType,principalId,resourceId,scope",
    mailboxDelegates:
      "/me/mailFolders/inbox/messageRules?$select=displayName,sequence,conditions,actions,isEnabled",
    forwardingRules: "/me/mailFolders/inbox/messageRules",

    // Intelligence Gathering
    searchMeetingNotes:
      "/me/drive/root/search(q='meeting notes OR minutes OR agenda')?$top=100&$select=name,webUrl,lastModifiedDateTime",
    searchOneNoteSecrets:
      "/me/onenote/pages?$top=100&$select=title,createdDateTime,contentUrl",
    searchPlannerTasks:
      "/me/planner/tasks?$select=title,assignments,percentComplete,dueDateTime,createdDateTime",
    searchIncidentEmails:
      '/me/messages?$search="incident OR breach OR vulnerability OR security alert"&$top=100&$select=subject,from,receivedDateTime,importance,bodyPreview',
    searchProjectFiles:
      "/me/drive/root/search(q='project OR roadmap OR strategy OR architecture')?$top=100&$select=name,webUrl,lastModifiedDateTime",

    // Audit/Logs
    signInLogs: "/auditLogs/signIns?$top=100&$orderby=createdDateTime desc",
    auditLogs:
      "/auditLogs/directoryAudits?$top=100&$orderby=activityDateTime desc",
    riskDetections: "/identityProtection/riskDetections?$top=100",
    riskyUsers: "/identityProtection/riskyUsers?$top=100",
  };

  const query = queries[queryType];
  if (query) {
    const queryInput = document.getElementById("customGraphQuery");
    if (queryInput) {
      queryInput.value = query;
      queryInput.focus();
    }
  }
}

async function executeGraphQuery(query) {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session. Please get a token first.");
    return;
  }

  const responseEl = document.getElementById("graphQueryResponse");
  responseEl.value = "Loading...";

  try {
    if (!query.startsWith("/")) {
      query = "/" + query;
    }

    const url = `https://graph.microsoft.com/v1.0${query}`;

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    const data = await response.json();

    if (!response.ok) {
      responseEl.value = `Error ${response.status}:\n${JSON.stringify(data, null, 2)}`;
      showToast(`Query failed: ${data.error?.message || response.statusText}`);
    } else {
      responseEl.value = JSON.stringify(data, null, 2);
      showToast("✅ Query successful");
    }
  } catch (error) {
    console.error("Query error:", error);
    responseEl.value = `Error: ${error.message}`;
    showToast("Query failed: " + error.message, "error");
  }
}
