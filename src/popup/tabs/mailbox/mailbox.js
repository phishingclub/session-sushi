let mailboxFolders = [];
let currentFolder = null;
let mailboxMessages = [];
let currentMessage = null;
let mailboxSearchQuery = "";
let isLoadingMailboxMessages = false;
let mailboxNextLink = null;

async function initializeMailbox() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showMailboxNoSession();
    return;
  }

  setupFolderManagementListeners();
  await loadMailboxFolders();
}

function showMailboxNoSession() {
  const foldersContainer = document.getElementById("mailboxFolders");
  const messagesContainer = document.getElementById("mailboxMessages");
  const detailContainer = document.getElementById("mailboxDetail");

  if (foldersContainer) {
    foldersContainer.innerHTML = `
            <div class="mailbox-folder-loading">
                Select an active session to view mailbox
            </div>
        `;
  }

  if (messagesContainer) {
    messagesContainer.innerHTML = `
            <div class="mailbox-empty">
                No active session
            </div>
        `;
  }

  if (detailContainer) {
    detailContainer.innerHTML = `
            <div class="mailbox-empty">
                No active session
            </div>
        `;
  }
}

async function loadMailboxFolders() {
  const foldersContainer = document.getElementById("mailboxFolders");

  if (!foldersContainer) {
    return;
  }

  foldersContainer.innerHTML = `
        <div class="loading-indicator">
            Loading...
        </div>
    `;

  try {
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/mailFolders?$top=100",
      {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      throw new Error(`Failed to load folders: ${response.statusText}`);
    }

    const data = await response.json();
    mailboxFolders = data.value || [];

    const allFolders = [...mailboxFolders];
    for (const folder of mailboxFolders) {
      if (folder.childFolderCount > 0) {
        try {
          const childResponse = await fetch(
            `https://graph.microsoft.com/v1.0/me/mailFolders/${folder.id}/childFolders?$top=100`,
            {
              headers: {
                Authorization: `Bearer ${activeM365Session.access_token}`,
                "Content-Type": "application/json",
              },
            },
          );
          if (childResponse.ok) {
            const childData = await childResponse.json();
            allFolders.push(...(childData.value || []));
          }
        } catch (error) {
          console.error(
            `Failed to fetch children for ${folder.displayName}:`,
            error,
          );
        }
      }
    }

    const folderMap = {};
    allFolders.forEach((folder) => {
      folderMap[folder.id] = folder;
      folder._children = [];
    });

    const rootFolders = [];
    allFolders.forEach((folder) => {
      const parent = folderMap[folder.parentFolderId];
      if (parent) {
        parent._children.push(folder);
      } else {
        rootFolders.push(folder);
      }
    });

    renderMailboxFolders(rootFolders);

    if (rootFolders.length > 0) {
      await selectFolder(rootFolders[0]);
    }
  } catch (error) {
    console.error("Failed to load mailbox folders:", error);
    showErrorInContainer(foldersContainer, error.message, {
      title: "Failed to load folders:",
    });
    showToast("Failed to load mailbox folders", "error");
  }
}

function renderMailboxFolders(rootFolders) {
  const foldersContainer = document.getElementById("mailboxFolders");
  if (!foldersContainer) {
    return;
  }

  foldersContainer.innerHTML = "";

  const sortedRootFolders = rootFolders.sort((a, b) => {
    return a.displayName.localeCompare(b.displayName);
  });

  sortedRootFolders.forEach((folder) => {
    renderFolder(folder, foldersContainer, 0);
  });
}

function renderFolder(folder, container, level) {
  const hasChildren = folder._children && folder._children.length > 0;
  const folderEl = document.createElement("div");
  folderEl.className = `mailbox-folder ${hasChildren ? "expandable" : ""}`;
  folderEl.dataset.folderId = folder.id;
  folderEl.style.paddingLeft = `${8 + level * 16}px`; // Dynamic padding based on nesting level

  const icon = getFolderIcon(folder.displayName);

  const expandIcon = hasChildren ? `<span class="folder-expand">▶</span>` : "";

  folderEl.innerHTML = `
      ${expandIcon}
      <span class="folder-icon">${icon}</span>
      <span class="folder-name"></span>
      <span class="folder-count display-none"></span>
  `;

  const folderNameEl = folderEl.querySelector(".folder-name");
  if (folderNameEl) {
    folderNameEl.textContent = folder.displayName;
  }

  if (folder.unreadItemCount > 0) {
    const countEl = folderEl.querySelector(".folder-count");
    if (countEl) {
      countEl.textContent = folder.unreadItemCount;
      countEl.classList.remove("display-none");
    }
  }

  folderEl.addEventListener("click", async (e) => {
    e.stopPropagation();

    if (hasChildren && e.target.closest(".folder-expand")) {
      toggleFolderExpansion(folderEl, folder._children, level);
    } else {
      await selectFolder(folder);
    }
  });

  folderEl.addEventListener("contextmenu", (e) => {
    e.preventDefault();
    e.stopPropagation();
    showFolderContextMenu(e, folder);
  });

  container.appendChild(folderEl);

  if (hasChildren) {
    folderEl._childFolders = folder._children;
    folderEl._level = level;
  }
}

function toggleFolderExpansion(folderEl, childFolders, level) {
  const isExpanded = folderEl.classList.contains("expanded");

  if (isExpanded) {
    folderEl.classList.remove("expanded");
    let nextSibling = folderEl.nextElementSibling;
    while (
      nextSibling &&
      nextSibling.classList.contains("mailbox-folder-child")
    ) {
      const toRemove = nextSibling;
      nextSibling = nextSibling.nextElementSibling;
      toRemove.remove();
    }
  } else {
    folderEl.classList.add("expanded");
    let insertAfter = folderEl;
    childFolders.forEach((child) => {
      const childEl = document.createElement("div");
      childEl.className = "mailbox-folder mailbox-folder-child";
      childEl.dataset.folderId = child.id;
      childEl.style.paddingLeft = `${8 + (level + 1) * 16}px`; // Dynamic padding

      const hasChildren = child._children && child._children.length > 0;
      if (hasChildren) {
        childEl.classList.add("expandable");
      }

      const icon = getFolderIcon(child.displayName);
      const count =
        child.unreadItemCount > 0
          ? `<span class="folder-count">${child.unreadItemCount}</span>`
          : "";

      const expandIcon = hasChildren
        ? `<span class="folder-expand">▶</span>`
        : "";

      childEl.innerHTML = `
                ${expandIcon}
                <span class="folder-icon">${icon}</span>
                <span class="folder-name"></span>
                <span class="folder-count display-none"></span>
            `;

      const childNameEl = childEl.querySelector(".folder-name");
      if (childNameEl) {
        childNameEl.textContent = child.displayName;
      }

      if (child.unreadItemCount > 0) {
        const childCountEl = childEl.querySelector(".folder-count");
        if (childCountEl && childFolder.unreadItemCount > 0) {
          childCountEl.textContent = childFolder.unreadItemCount;
          childCountEl.classList.remove("display-none");
        }
      }

      childEl.addEventListener("click", async (e) => {
        e.stopPropagation();

        if (hasChildren && e.target.closest(".folder-expand")) {
          toggleFolderExpansion(childEl, child._children, level + 1);
        } else {
          await selectFolder(child);
        }
      });

      insertAfter.parentNode.insertBefore(childEl, insertAfter.nextSibling);
      insertAfter = childEl;
    });
  }
}

function getFolderIcon(folderName) {
  const icons = {
    Inbox: "📥",
    "Sent Items": "📤",
    Drafts: "📝",
    "Deleted Items": "🗑️",
    "Junk Email": "🚫",
    Archive: "📦",
    Outbox: "📮",
  };
  return icons[folderName] || "📁";
}

async function selectFolder(folder) {
  currentFolder = folder;
  mailboxMessages = [];
  currentMessage = null;
  mailboxNextLink = null;

  document.querySelectorAll(".mailbox-folder").forEach((el) => {
    el.classList.remove("active");
  });

  const folderEl = document.querySelector(
    `.mailbox-folder[data-folder-id="${folder.id}"]`,
  );
  if (folderEl) {
    folderEl.classList.add("active");
  }

  const listTitle = document.querySelector(".mailbox-list-title");
  if (listTitle) {
    listTitle.textContent = folder.displayName;
  }

  const detailContainer = document.getElementById("mailboxDetail");
  if (detailContainer) {
    detailContainer.innerHTML = `
            <div class="mailbox-empty">
                Select a message to view
            </div>
        `;
  }

  // Update delete button state
  updateDeleteButtonState();

  // Load messages
  await loadFolderMessages(folder.id);
}

// ============================================
//  MESSAGE LOADING
// ============================================

async function loadFolderMessages(folderId, isSearch = false) {
  if (isLoadingMailboxMessages) {
    return;
  }
  isLoadingMailboxMessages = true;

  const messagesContainer = document.getElementById("mailboxMessages");

  if (!messagesContainer) {
    isLoadingMailboxMessages = false;
    return;
  }

  messagesContainer.innerHTML = `
        <div class="loading-indicator">
            Loading...
        </div>
    `;

  try {
    let url;
    if (isSearch && mailboxSearchQuery) {
      // $orderby is not supported with $search
      url = `https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}/messages?$search="${encodeURIComponent(mailboxSearchQuery)}"&$top=50&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments,importance`;
    } else {
      url = `https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}/messages?$top=50&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments,importance&$orderby=receivedDateTime desc`;
    }

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      throw new Error(`Failed to load messages: ${response.statusText}`);
    }

    const data = await response.json();
    mailboxMessages = data.value || [];
    mailboxNextLink = data["@odata.nextLink"] || null;

    renderMailboxMessages();
    updateMessageCount();

    // Setup scroll listener for infinite scroll
    setupMailboxScrollListener(folderId);
  } catch (error) {
    console.error("Failed to load messages:", error);
    showErrorInContainer(messagesContainer, error.message, {
      title: "Failed to load messages:",
    });
    showToast("Failed to load messages", "error");
  } finally {
    isLoadingMailboxMessages = false;
  }
}

async function loadMoreMessages() {
  if (!mailboxNextLink || isLoadingMailboxMessages) return;
  isLoadingMailboxMessages = true;

  const messagesContainer = document.getElementById("mailboxMessages");
  const loadingEl = document.createElement("div");
  loadingEl.className = "mailbox-loading-more";
  loadingEl.textContent = "Loading more messages...";
  messagesContainer.appendChild(loadingEl);

  try {
    const response = await fetch(mailboxNextLink, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      throw new Error(`Failed to load more messages: ${response.statusText}`);
    }

    const data = await response.json();
    const newMessages = data.value || [];
    mailboxMessages.push(...newMessages);
    mailboxNextLink = data["@odata.nextLink"] || null;

    loadingEl.remove();
    renderMailboxMessages();
    updateMessageCount();
  } catch (error) {
    console.error("Failed to load more messages:", error);
    loadingEl.textContent = "Failed to load more messages";
    showToast("Failed to load more messages");
  } finally {
    isLoadingMailboxMessages = false;
  }
}

function setupMailboxScrollListener(folderId) {
  const messagesContainer = document.getElementById("mailboxMessages");
  if (!messagesContainer) return;

  // Remove existing listener
  messagesContainer.removeEventListener("scroll", handleMailboxScroll);

  // Add new listener
  messagesContainer.addEventListener("scroll", handleMailboxScroll);
}

function handleMailboxScroll(e) {
  const container = e.target;
  const scrollPosition = container.scrollTop + container.clientHeight;
  const scrollHeight = container.scrollHeight;

  // Load more when scrolled to 80% of the way down
  if (scrollPosition >= scrollHeight * 0.8 && mailboxNextLink) {
    loadMoreMessages();
  }
}

function renderMailboxMessages() {
  const messagesContainer = document.getElementById("mailboxMessages");
  if (!messagesContainer) {
    return;
  }

  if (mailboxMessages.length === 0) {
    messagesContainer.innerHTML = `
            <div class="mailbox-empty">
                No messages in this folder
            </div>
        `;
    return;
  }

  messagesContainer.innerHTML = "";

  mailboxMessages.forEach((message) => {
    const messageEl = document.createElement("div");
    messageEl.className = `mailbox-message ${!message.isRead ? "unread" : ""}`;
    messageEl.dataset.messageId = message.id;

    const senderName = message.from?.emailAddress?.name || "Unknown";
    const senderEmail = message.from?.emailAddress?.address || "";
    const subject = message.subject || "(No Subject)";
    const preview = message.bodyPreview || "";
    const time = formatMessageTime(message.receivedDateTime);

    const flags = [];
    if (message.hasAttachments) {
      flags.push('<span class="message-flag attachment">📎</span>');
    }
    if (message.importance === "high") {
      flags.push('<span class="message-flag important">!</span>');
    }

    const flagsHtml =
      flags.length > 0
        ? `<div class="message-flags">${flags.join("")}</div>`
        : "";

    messageEl.innerHTML = `
            <div class="mailbox-message-header">
                <span class="message-sender"></span>
                <span class="message-time"></span>
            </div>
            <div class="message-subject"></div>
            <div class="message-preview"></div>
            ${flagsHtml}
        `;

    // Set untrusted content using textContent to prevent XSS
    const senderEl = messageEl.querySelector(".message-sender");
    if (senderEl) {
      senderEl.textContent = senderName;
      senderEl.title = `${senderName} <${senderEmail}>`;
    }

    const timeEl = messageEl.querySelector(".message-time");
    if (timeEl) {
      timeEl.textContent = time;
    }

    const subjectEl = messageEl.querySelector(".message-subject");
    if (subjectEl) {
      subjectEl.textContent = subject;
    }

    const previewEl = messageEl.querySelector(".message-preview");
    if (previewEl) {
      previewEl.textContent = preview;
    }

    messageEl.onclick = () => {
      selectMessage(message);
    };

    messagesContainer.appendChild(messageEl);
  });
}

function updateMessageCount() {
  const countEl = document.getElementById("mailboxListCount");
  if (countEl) {
    countEl.textContent = `${mailboxMessages.length} message${mailboxMessages.length !== 1 ? "s" : ""}`;
  }
}

function formatMessageTime(dateString) {
  if (!dateString) return "";

  const date = new Date(dateString);
  const now = new Date();
  const diffMs = now - date;
  const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

  if (diffDays === 0) {
    // Today - show time
    return date.toLocaleTimeString("en-US", {
      hour: "numeric",
      minute: "2-digit",
      hour12: true,
    });
  } else if (diffDays === 1) {
    return "Yesterday";
  } else if (diffDays < 7) {
    return date.toLocaleDateString("en-US", { weekday: "short" });
  } else if (date.getFullYear() === now.getFullYear()) {
    return date.toLocaleDateString("en-US", {
      month: "short",
      day: "numeric",
    });
  } else {
    return date.toLocaleDateString("en-US", {
      month: "short",
      day: "numeric",
      year: "numeric",
    });
  }
}

// ============================================
//  MESSAGE DETAIL
// ============================================

async function selectMessage(message) {
  currentMessage = message;

  // Update UI
  document.querySelectorAll(".mailbox-message").forEach((el) => {
    el.classList.remove("active");
  });

  // Find message element by iterating to avoid selector injection
  const messageEls = document.querySelectorAll(".mailbox-message");
  let messageEl = null;
  for (const el of messageEls) {
    if (el.dataset.messageId === message.id) {
      messageEl = el;
      break;
    }
  }
  if (messageEl) {
    messageEl.classList.add("active");
  }

  // Load full message
  await loadMessageDetail(message.id);

  // Don't automatically mark as read - user must do it manually
}

async function loadMessageDetail(messageId) {
  const detailContainer = document.getElementById("mailboxDetail");
  if (!detailContainer) return;

  detailContainer.innerHTML = `
        <div class="loading-indicator">
            Loading...
        </div>
    `;

  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${messageId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,importance&$expand=attachments`,
      {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      throw new Error(`Failed to load message: ${response.statusText}`);
    }

    const message = await response.json();
    renderMessageDetail(message);
  } catch (error) {
    console.error("Failed to load message detail:", error);
    showErrorInContainer(detailContainer, error.message, {
      title: "Failed to load message:",
    });
    showToast("Failed to load message details", "error");
  }
}

function renderMessageDetail(message) {
  const detailContainer = document.getElementById("mailboxDetail");
  if (!detailContainer) return;

  const subject = message.subject || "(No Subject)";
  const senderName = message.from?.emailAddress?.name || "Unknown";
  const senderEmail = message.from?.emailAddress?.address || "";
  const receivedDate = new Date(message.receivedDateTime);
  const formattedDate = receivedDate.toLocaleString("en-US", {
    weekday: "short",
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  });

  const toRecipients = message.toRecipients
    ?.map((r) => `${r.emailAddress.name} <${r.emailAddress.address}>`)
    .join(", ");

  const ccRecipients = message.ccRecipients
    ?.map((r) => `${r.emailAddress.name} <${r.emailAddress.address}>`)
    .join(", ");

  // Extract text content from body (strip HTML)
  let bodyText = "";
  let bodyHtml = "";
  if (message.body) {
    if (message.body.contentType === "text") {
      bodyText = message.body.content;
      bodyHtml = "";
    } else {
      // Strip HTML tags for security in text mode
      bodyText = stripHtml(message.body.content);
      bodyHtml = message.body.content;
    }
  }

  // Store attachments for inline image resolution
  detailContainer._attachments = message.attachments || [];

  const showCcRow = ccRecipients && ccRecipients.length > 0;

  const htmlViewToggle = bodyHtml
    ? `
        <div class="info-box-compact">
            <span class="info-box-label">
                <span id="viewModeLabel">Viewing as: Text</span>
            </span>
            <button class="btn btn-sm btn-primary" id="toggleHtmlView">
                View as HTML
            </button>
        </div>
    `
    : "";

  // Build structure with placeholders
  detailContainer.innerHTML = `
        <div class="mailbox-detail-header">
            <div class="detail-subject"></div>
            <div class="detail-meta">
                <div class="detail-meta-row">
                    <span class="detail-meta-label">From:</span>
                    <span class="detail-meta-value" id="detailFromValue"></span>
                </div>
                <div class="detail-meta-row">
                    <span class="detail-meta-label">To:</span>
                    <span class="detail-meta-value" id="detailToValue"></span>
                </div>
                <div class="detail-meta-row ${showCcRow ? "" : "display-none"}" id="detailCcRow">
                    <span class="detail-meta-label">Cc:</span>
                    <span class="detail-meta-value" id="detailCcValue"></span>
                </div>
                <div class="detail-meta-row">
                    <span class="detail-meta-label">Date:</span>
                    <span class="detail-meta-value" id="detailDateValue"></span>
                </div>
            </div>
            <div class="detail-actions">
                <button class="btn btn-primary btn-sm" data-action="reply">↩️ Reply</button>
                <button class="btn btn-primary btn-sm" data-action="forward">➡️ Forward</button>
                <button class="btn btn-secondary btn-sm" data-action="download">📥 Download EML</button>
                <button class="btn btn-danger-outline btn-sm" data-action="delete">🗑️ Delete</button>
            </div>
        </div>
        <div class="mailbox-detail-body">
            ${htmlViewToggle}
            <div class="detail-body-content" id="messageBodyContent"></div>
            <div id="attachmentsContainer"></div>
        </div>
    `;

  // Set untrusted content using textContent to prevent XSS
  const subjectEl = detailContainer.querySelector(".detail-subject");
  if (subjectEl) {
    subjectEl.textContent = subject;
  }

  const fromEl = detailContainer.querySelector("#detailFromValue");
  if (fromEl) {
    fromEl.textContent = `${senderName} <${senderEmail}>`;
  }

  const toEl = detailContainer.querySelector("#detailToValue");
  if (toEl) {
    toEl.textContent = toRecipients || "";
  }

  if (showCcRow) {
    const ccEl = detailContainer.querySelector("#detailCcValue");
    if (ccEl) {
      ccEl.textContent = ccRecipients;
    }
  }

  const dateEl = detailContainer.querySelector("#detailDateValue");
  if (dateEl) {
    dateEl.textContent = formattedDate;
  }

  // Set body text content (HTML stripped)
  const bodyContentEl = detailContainer.querySelector("#messageBodyContent");
  if (bodyContentEl) {
    bodyContentEl.textContent = bodyText;
  }

  // Render attachments
  if (message.hasAttachments && message.attachments?.length > 0) {
    renderAttachments(message.attachments, message.id);
  }

  // Store the body content and attachments for toggling
  detailContainer._bodyText = bodyText;
  detailContainer._bodyHtml = bodyHtml;
  detailContainer._attachments = message.attachments || [];
  detailContainer._isHtmlView = false;
  detailContainer._subject = subject;
  detailContainer._message = message;
  detailContainer._messageId = message.id;
}

function renderAttachments(attachments, messageId) {
  if (!attachments || attachments.length === 0) return;

  const container = document.getElementById("attachmentsContainer");
  if (!container) return;

  const wrapper = document.createElement("div");
  wrapper.className = "detail-attachments";

  const title = document.createElement("div");
  title.className = "detail-attachments-title";
  title.textContent = `Attachments (${attachments.length})`;
  wrapper.appendChild(title);

  attachments.forEach((attachment) => {
    const attachmentEl = document.createElement("div");
    attachmentEl.className = "detail-attachment";
    attachmentEl.dataset.messageId = messageId;
    attachmentEl.dataset.attachmentId = attachment.id;
    attachmentEl.dataset.attachmentName = attachment.name;

    const size = formatFileSize(attachment.size);
    const icon = getFileIcon(attachment.name);

    attachmentEl.innerHTML = `
      <span class="attachment-icon">${icon}</span>
      <div class="attachment-info">
          <div class="attachment-name"></div>
          <div class="attachment-size"></div>
      </div>
      <span class="attachment-download">⬇️</span>
    `;

    // Set untrusted content using textContent
    const nameEl = attachmentEl.querySelector(".attachment-name");
    if (nameEl) {
      nameEl.textContent = attachment.name;
    }

    const sizeEl = attachmentEl.querySelector(".attachment-size");
    if (sizeEl) {
      sizeEl.textContent = size;
    }

    wrapper.appendChild(attachmentEl);
  });

  container.innerHTML = "";
  container.appendChild(wrapper);
}

function formatFileSize(bytes) {
  if (!bytes) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + " " + sizes[i];
}

function getFileIcon(filename) {
  const ext = filename.split(".").pop().toLowerCase();
  const icons = {
    pdf: "📄",
    doc: "📝",
    docx: "📝",
    xls: "📊",
    xlsx: "📊",
    ppt: "📽️",
    pptx: "📽️",
    zip: "🗜️",
    rar: "🗜️",
    jpg: "🖼️",
    jpeg: "🖼️",
    png: "🖼️",
    gif: "🖼️",
    txt: "📃",
    csv: "📋",
  };
  return icons[ext] || "📎";
}

async function downloadAttachment(messageId, attachmentId, filename) {
  try {
    showToast("Downloading attachment...");

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments/${attachmentId}`,
      {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      throw new Error(`Failed to download attachment: ${response.statusText}`);
    }

    const attachment = await response.json();

    // Convert base64 to blob
    const contentBytes = attachment.contentBytes;
    const byteCharacters = atob(contentBytes);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], {
      type: attachment.contentType || "application/octet-stream",
    });

    // Download the file
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast("✅ Attachment downloaded");
  } catch (error) {
    console.error("Failed to download attachment:", error);
    showToast("Failed to download attachment: " + error.message);
  }
}

async function markMessageAsRead(messageId) {
  try {
    await fetch(`https://graph.microsoft.com/v1.0/me/messages/${messageId}`, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        isRead: true,
      }),
    });
  } catch (error) {
    console.error("Failed to mark message as read:", error);
  }
}

// ============================================
//  MESSAGE ACTIONS
// ============================================

async function deleteMessage(messageId) {
  if (!confirm("Are you sure you want to delete this message?")) {
    return;
  }

  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${messageId}`,
      {
        method: "DELETE",
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
        },
      },
    );

    if (!response.ok) {
      throw new Error(`Failed to delete message: ${response.statusText}`);
    }

    showToast("✅ Message deleted");

    // Remove from list
    mailboxMessages = mailboxMessages.filter((m) => m.id !== messageId);
    renderMailboxMessages();
    updateMessageCount();

    // Clear detail view
    const detailContainer = document.getElementById("mailboxDetail");
    if (detailContainer) {
      detailContainer.innerHTML = `
                <div class="mailbox-empty">
                    Select a message to view
                </div>
            `;
    }

    // Reload folder to update counts
    if (currentFolder) {
      await loadMailboxFolders();
    }
  } catch (error) {
    console.error("Failed to delete message:", error);
    showToast("Failed to delete message: " + error.message);
  }
}

// ============================================
//  SEARCH
// ============================================

function setupMailboxSearch() {
  const searchInput = document.getElementById("mailboxSearch");
  if (!searchInput) return;

  let searchTimeout;
  searchInput.addEventListener("input", (e) => {
    clearTimeout(searchTimeout);
    mailboxSearchQuery = e.target.value.trim();

    searchTimeout = setTimeout(async () => {
      if (currentFolder) {
        await loadFolderMessages(currentFolder.id, true);
      }
    }, 500);
  });
}

// ============================================
//  UTILITY FUNCTIONS
// ============================================

function stripHtml(html) {
  if (!html) return "";

  // Remove script tags and their content
  html = html.replace(
    /<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi,
    "",
  );

  // Remove style tags and their content
  html = html.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, "");

  // Remove all HTML tags
  html = html.replace(/<[^>]+>/g, "");

  // Decode HTML entities
  const textarea = document.createElement("textarea");
  textarea.innerHTML = html;
  html = textarea.value;

  // Remove excessive whitespace
  html = html.replace(/\n\s*\n/g, "\n\n");
  html = html.trim();

  return html;
}

// Note: This function is kept for backwards compatibility but should not be used
// Use textContent directly instead to prevent XSS

// ============================================
//  REFRESH MAILBOX
// ============================================

async function refreshMailbox() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const refreshBtn = document.getElementById("refreshMailbox");
  if (refreshBtn) {
    refreshBtn.disabled = true;
    refreshBtn.textContent = "⟳";
  }

  try {
    // Reload folders
    await loadMailboxFolders();
    showToast("✅ Mailbox refreshed");
  } catch (error) {
    showToast("Failed to refresh mailbox");
  } finally {
    if (refreshBtn) {
      refreshBtn.disabled = false;
      refreshBtn.textContent = "🔄";
    }
  }
}

// ============================================
//  MESSAGE DETAIL EVENT LISTENERS
// ============================================

function setupMessageDetailListeners() {
  const detailContainer = document.getElementById("mailboxDetail");
  if (!detailContainer) return;

  detailContainer.addEventListener("click", async (e) => {
    if (e.target.id === "toggleHtmlView") {
      toggleMessageView();
      return;
    }

    const btn = e.target.closest("button[data-action]");
    if (btn) {
      const action = btn.dataset.action;
      const messageId = detailContainer._messageId;

      switch (action) {
        case "reply":
          if (currentMessage) {
            replyToMessage(currentMessage);
          }
          break;
        case "forward":
          if (currentMessage) {
            forwardMessage(currentMessage);
          }
          break;
        case "download":
          const subject = detailContainer._subject || "message";
          await downloadMessageAsEml(messageId, subject);
          break;
        case "delete":
          await deleteMessage(messageId);
          break;
      }
      return;
    }

    const attachment = e.target.closest(".detail-attachment");
    if (attachment) {
      const messageId = attachment.dataset.messageId;
      const attachmentId = attachment.dataset.attachmentId;
      const attachmentName = attachment.dataset.attachmentName;
      await downloadAttachment(messageId, attachmentId, attachmentName);
    }
  });
}
function setupRemoveAttachmentListeners() {
  const modal = document.getElementById("composeEmailModal");
  if (!modal) return;

  modal.addEventListener("click", (e) => {
    const btn = e.target.closest(".attachment-remove");
    if (btn) {
      const index = parseInt(btn.getAttribute("data-index"), 10);
      removeAttachment(index);
    }
  });
}

// ============================================
//  HTML/TEXT VIEW TOGGLE
// ============================================

function toggleMessageView() {
  if (!confirm("View HTML content? Scripts and external resources may load.")) {
    return;
  }

  const detailContainer = document.getElementById("mailboxDetail");
  if (!detailContainer) return;

  const bodyContentEl = document.getElementById("messageBodyContent");
  const toggleBtn = document.getElementById("toggleHtmlView");
  const viewLabel = document.getElementById("viewModeLabel");

  if (!bodyContentEl || !toggleBtn) return;

  detailContainer._isHtmlView = !detailContainer._isHtmlView;

  if (detailContainer._isHtmlView) {
    // Switch to HTML view
    let htmlContent = detailContainer._bodyHtml;
    const attachments = detailContainer._attachments || [];

    // Convert cid: URLs to data URIs if we have attachments
    if (attachments.length > 0) {
      attachments.forEach((attachment) => {
        if (attachment.contentId && attachment.contentBytes) {
          // Remove < > brackets from contentId if present
          const cid = attachment.contentId.replace(/^<|>$/g, "");
          const dataUri = `data:${attachment.contentType};base64,${attachment.contentBytes}`;

          // Replace cid: references with data URIs
          const cidPattern = new RegExp(`cid:${cid}`, "gi");
          htmlContent = htmlContent.replace(cidPattern, dataUri);
        }
      });
    }

    bodyContentEl.innerHTML = htmlContent;
    toggleBtn.textContent = "View as Text";
    viewLabel.textContent = "Viewing as: HTML";
  } else {
    // Switch to text view
    bodyContentEl.textContent = detailContainer._bodyText;
    toggleBtn.textContent = "View as HTML";
    viewLabel.textContent = "Viewing as: Text";
  }
}

// ============================================
//  DOWNLOAD MESSAGE AS EML
// ============================================

async function downloadMessageAsEml(messageId, subject) {
  try {
    showToast("Downloading message...");

    // Fetch the message in MIME format
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${messageId}/$value`,
      {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
        },
      },
    );

    if (!response.ok) {
      throw new Error(`Failed to download message: ${response.statusText}`);
    }

    const mimeContent = await response.text();

    // Create blob and download
    const blob = new Blob([mimeContent], { type: "message/rfc822" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;

    // Sanitize filename
    const safeSubject = subject.replace(/[^a-z0-9]/gi, "_").substring(0, 50);
    a.download = `${safeSubject || "message"}_${messageId.substring(0, 8)}.eml`;

    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast("✅ Message downloaded");
  } catch (error) {
    console.error("Failed to download message:", error);
    showToast("Failed to download message: " + error.message);
  }
}

// ============================================
// COMPOSE EMAIL FUNCTIONALITY
// ============================================

let selectedAttachments = [];

/**
 * Open the compose email modal
 */
function openComposeEmailModal() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("Please select an active M365 session first");
    return;
  }

  const modal = document.getElementById("composeEmailModal");
  if (modal) {
    modal.classList.add("modal-show");
    clearComposeForm();
  }
}

/**
 * Close the compose email modal
 */
function closeComposeEmailModal() {
  const modal = document.getElementById("composeEmailModal");
  if (modal) {
    modal.classList.remove("modal-show");
    clearComposeForm();
  }
}

/**
 * Clear the compose form
 */
function clearComposeForm() {
  document.getElementById("composeTo").value = "";
  document.getElementById("composeCc").value = "";
  document.getElementById("composeBcc").value = "";
  document.getElementById("composeSubject").value = "";
  document.getElementById("composeBody").value = "";
  document.getElementById("composeUseHtml").checked = true;
  document.getElementById("composeSaveToSent").checked = false;
  selectedAttachments = [];
  renderAttachmentsList();
}

/**
 * Setup compose email listeners
 */
function setupComposeEmailListeners() {
  const composeBtn = document.getElementById("composeEmailBtn");
  if (composeBtn) {
    composeBtn.addEventListener("click", openComposeEmailModal);
  }

  const attachmentsInput = document.getElementById("composeAttachments");
  if (attachmentsInput) {
    attachmentsInput.addEventListener("change", handleAttachmentsSelected);
  }

  // Add attachments button
  const addAttachmentsBtn = document.getElementById("addAttachmentsBtn");
  if (addAttachmentsBtn) {
    addAttachmentsBtn.addEventListener("click", () => {
      document.getElementById("composeAttachments").click();
    });
  }

  // Close button (X)
  const closeBtn = document.getElementById("closeComposeEmailModalBtn");
  if (closeBtn) {
    closeBtn.addEventListener("click", closeComposeEmailModal);
  }

  // Cancel button
  const cancelBtn = document.getElementById("cancelComposeEmailBtn");
  if (cancelBtn) {
    cancelBtn.addEventListener("click", closeComposeEmailModal);
  }

  // Send button
  const sendBtn = document.getElementById("sendEmailBtn");
  if (sendBtn) {
    sendBtn.addEventListener("click", sendEmail);
  }

  setupRemoveAttachmentListeners();
}

/**
 * Handle file attachments selection
 */
function handleAttachmentsSelected(event) {
  const files = Array.from(event.target.files);

  // Check total size (Microsoft Graph has limits)
  const maxTotalSize = 3 * 1024 * 1024; // 3MB for inline attachments
  let currentSize = selectedAttachments.reduce((sum, att) => sum + att.size, 0);

  for (const file of files) {
    if (currentSize + file.size > maxTotalSize) {
      showToast(
        `⚠️ Total attachments size cannot exceed 3MB. File "${file.name}" was not added.`,
      );
      continue;
    }

    selectedAttachments.push(file);
    currentSize += file.size;
  }

  renderAttachmentsList();
  event.target.value = ""; // Reset input
}

/**
 * Render attachments list
 */
function renderAttachmentsList() {
  const container = document.getElementById("attachmentsList");
  if (!container) return;

  if (selectedAttachments.length === 0) {
    container.innerHTML = "";
    return;
  }

  container.innerHTML = "";

  selectedAttachments.forEach((file, index) => {
    const icon = getFileIcon(file.name);
    const size = formatFileSize(file.size);

    const attachmentItem = document.createElement("div");
    attachmentItem.className = "attachment-item";

    const attachmentInfo = document.createElement("div");
    attachmentInfo.className = "attachment-info";

    const iconSpan = document.createElement("span");
    iconSpan.className = "attachment-icon";
    iconSpan.textContent = icon;

    const attachmentDetails = document.createElement("div");
    attachmentDetails.className = "attachment-details";

    const nameSpan = document.createElement("span");
    nameSpan.className = "attachment-name";
    nameSpan.textContent = file.name;

    const sizeSpan = document.createElement("span");
    sizeSpan.className = "attachment-size";
    sizeSpan.textContent = size;

    attachmentDetails.appendChild(nameSpan);
    attachmentDetails.appendChild(sizeSpan);

    attachmentInfo.appendChild(iconSpan);
    attachmentInfo.appendChild(attachmentDetails);

    const removeBtn = document.createElement("button");
    removeBtn.className = "attachment-remove";
    removeBtn.setAttribute("data-index", index);
    removeBtn.title = "Remove";
    removeBtn.textContent = "✕";

    attachmentItem.appendChild(attachmentInfo);
    attachmentItem.appendChild(removeBtn);

    container.appendChild(attachmentItem);
  });
}

/**
 * Remove an attachment from the list
 */
function removeAttachment(index) {
  selectedAttachments.splice(index, 1);
  renderAttachmentsList();
}

/**
 * Convert file to base64
 */
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const base64 = reader.result.split(",")[1];
      resolve(base64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

/**
 * Send email
 */
async function sendEmail() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const to = document.getElementById("composeTo").value.trim();
  const cc = document.getElementById("composeCc").value.trim();
  const bcc = document.getElementById("composeBcc").value.trim();
  const subject = document.getElementById("composeSubject").value.trim();
  const body = document.getElementById("composeBody").value;
  const useHtml = document.getElementById("composeUseHtml").checked;
  const saveToSent = document.getElementById("composeSaveToSent").checked;

  if (!to) {
    showToast("Please enter at least one recipient");
    return;
  }

  if (!subject) {
    showToast("Please enter a subject");
    return;
  }

  try {
    showToast("Sending email...");

    // Parse recipients
    const toRecipients = to
      .split(/[;,]/)
      .map((email) => ({
        emailAddress: { address: email.trim() },
      }))
      .filter((r) => r.emailAddress.address);

    const ccRecipients = cc
      ? cc
          .split(/[;,]/)
          .map((email) => ({
            emailAddress: { address: email.trim() },
          }))
          .filter((r) => r.emailAddress.address)
      : [];

    const bccRecipients = bcc
      ? bcc
          .split(/[;,]/)
          .map((email) => ({
            emailAddress: { address: email.trim() },
          }))
          .filter((r) => r.emailAddress.address)
      : [];

    // Build message object
    const message = {
      subject: subject,
      body: {
        contentType: useHtml ? "HTML" : "Text",
        content: body,
      },
      toRecipients: toRecipients,
      ccRecipients: ccRecipients.length > 0 ? ccRecipients : undefined,
      bccRecipients: bccRecipients.length > 0 ? bccRecipients : undefined,
    };

    // Add attachments if any
    if (selectedAttachments.length > 0) {
      message.attachments = [];

      for (const file of selectedAttachments) {
        const base64Content = await fileToBase64(file);
        message.attachments.push({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: file.name,
          contentType: file.type || "application/octet-stream",
          contentBytes: base64Content,
        });
      }
    }

    // Send email using /sendMail endpoint
    const sendUrl = "https://graph.microsoft.com/v1.0/me/sendMail";
    const sendResponse = await fetch(sendUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        message: message,
        saveToSentItems: saveToSent,
      }),
    });

    if (!sendResponse.ok) {
      const error = await sendResponse.json();
      throw new Error(error.error?.message || "Failed to send");
    }

    showToast("✅ Email sent successfully");
    closeComposeEmailModal();
  } catch (error) {
    console.error("Failed to send email:", error);
    showToast("Failed to send email: " + error.message);
  }
}

/**
 * Reply to a message
 */
function replyToMessage(message) {
  if (!message) return;

  openComposeEmailModal();

  // Populate the compose form with reply data
  const replyTo = message.from?.emailAddress?.address || "";
  const subject = message.subject || "";
  const replySubject = subject.startsWith("Re:") ? subject : `Re: ${subject}`;

  document.getElementById("composeTo").value = replyTo;
  document.getElementById("composeSubject").value = replySubject;

  // Add original message as quoted text
  const originalBody = message.body?.content || "";
  const originalText = stripHtml(originalBody);
  const senderName = message.from?.emailAddress?.name || replyTo;
  const receivedDate = new Date(message.receivedDateTime).toLocaleString();

  // Check if HTML mode is enabled
  const useHtml = document.getElementById("composeUseHtml").checked;

  let quotedMessage;
  if (useHtml) {
    quotedMessage = `<br/><br/><hr/><p><strong>From:</strong> ${senderName} &lt;${replyTo}&gt;<br/><strong>Date:</strong> ${receivedDate}<br/><strong>Subject:</strong> ${subject}</p><br/>${originalText.replace(/\n/g, "<br/>")}`;
  } else {
    quotedMessage = `\n\n────────────────────────────────\n\nFrom: ${senderName} <${replyTo}>\nDate: ${receivedDate}\nSubject: ${subject}\n\n${originalText}`;
  }

  document.getElementById("composeBody").value = quotedMessage;

  showToast("Replying to message");
}

/**
 * Forward a message
 */
function forwardMessage(message) {
  if (!message) return;

  openComposeEmailModal();

  // Populate the compose form with forward data
  const subject = message.subject || "";
  const forwardSubject = subject.startsWith("Fwd:")
    ? subject
    : `Fwd: ${subject}`;

  document.getElementById("composeSubject").value = forwardSubject;

  // Add original message content
  const originalBody = message.body?.content || "";
  const originalText = stripHtml(originalBody);
  const senderName = message.from?.emailAddress?.name || "";
  const senderEmail = message.from?.emailAddress?.address || "";
  const receivedDate = new Date(message.receivedDateTime).toLocaleString();

  const toRecipients = message.toRecipients
    ?.map((r) => `${r.emailAddress.name} <${r.emailAddress.address}>`)
    .join(", ");

  // Check if HTML mode is enabled
  const useHtml = document.getElementById("composeUseHtml").checked;

  let forwardedMessage;
  if (useHtml) {
    forwardedMessage = `<br/><br/><hr/><p><strong>Begin forwarded message:</strong></p><br/><p><strong>From:</strong> ${senderName} &lt;${senderEmail}&gt;<br/><strong>Date:</strong> ${receivedDate}<br/><strong>To:</strong> ${toRecipients || ""}<br/><strong>Subject:</strong> ${subject}</p><br/>${originalText.replace(/\n/g, "<br/>")}`;
  } else {
    forwardedMessage = `\n\n────────────────────────────────\nBegin forwarded message:\n\nFrom: ${senderName} <${senderEmail}>\nDate: ${receivedDate}\nTo: ${toRecipients || ""}\nSubject: ${subject}\n\n${originalText}`;
  }

  document.getElementById("composeBody").value = forwardedMessage;

  // Note: Attachments from the original message are not automatically included
  // Users would need to manually add them if needed
  if (message.hasAttachments) {
    showToast(
      "Note: Original attachments are not automatically included. Please add them manually if needed.",
    );
  } else {
    showToast("Forwarding message");
  }
}

// ============================================
//  FOLDER MANAGEMENT
// ============================================

function setupFolderManagementListeners() {
  const addFolderBtn = document.getElementById("addFolderBtn");
  const deleteFolderBtn = document.getElementById("deleteFolderBtn");
  const closeAddFolderModal = document.getElementById("closeAddFolderModal");
  const cancelAddFolderBtn = document.getElementById("cancelAddFolderBtn");
  const confirmAddFolderBtn = document.getElementById("confirmAddFolderBtn");

  if (addFolderBtn) {
    addFolderBtn.addEventListener("click", openAddFolderModal);
  }

  if (deleteFolderBtn) {
    deleteFolderBtn.addEventListener("click", deleteSelectedFolder);
  }

  if (closeAddFolderModal) {
    closeAddFolderModal.addEventListener("click", closeAddFolderModalHandler);
  }

  if (cancelAddFolderBtn) {
    cancelAddFolderBtn.addEventListener("click", closeAddFolderModalHandler);
  }

  if (confirmAddFolderBtn) {
    confirmAddFolderBtn.addEventListener("click", createNewFolder);
  }

  // Setup keyboard shortcuts for the modal
  const modal = document.getElementById("addMailFolderModal");
  const folderNameInput = document.getElementById("newFolderName");

  if (modal) {
    modal.addEventListener("click", (e) => {
      // Close modal when clicking outside the modal content
      if (e.target === modal) {
        closeAddFolderModalHandler();
      }
    });
  }

  if (folderNameInput) {
    folderNameInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        createNewFolder();
      } else if (e.key === "Escape") {
        e.preventDefault();
        closeAddFolderModalHandler();
      }
    });
  }
}

function openAddFolderModal() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const modal = document.getElementById("addMailFolderModal");
  const parentSelect = document.getElementById("parentFolderSelect");
  const folderNameInput = document.getElementById("newFolderName");

  if (!modal || !parentSelect) {
    return;
  }

  // Clear the input
  if (folderNameInput) {
    folderNameInput.value = "";
  }

  // Populate parent folder options
  parentSelect.innerHTML = '<option value="">Root (No parent)</option>';

  if (mailboxFolders && mailboxFolders.length > 0) {
    mailboxFolders.forEach((folder) => {
      const option = document.createElement("option");
      option.value = folder.id;
      option.textContent = folder.displayName;
      parentSelect.appendChild(option);
    });
  }

  modal.classList.add("modal-show");

  // Focus on the folder name input
  setTimeout(() => {
    if (folderNameInput) {
      folderNameInput.focus();
    }
  }, 100);
}

function closeAddFolderModalHandler() {
  const modal = document.getElementById("addMailFolderModal");
  if (modal) {
    modal.classList.remove("modal-show");
  }
}

async function createNewFolder() {
  const folderNameInput = document.getElementById("newFolderName");
  const parentSelect = document.getElementById("parentFolderSelect");

  if (!folderNameInput) {
    return;
  }

  const folderName = folderNameInput.value.trim();
  const parentFolderId = parentSelect ? parentSelect.value : "";

  if (!folderName) {
    showToast("Please enter a folder name");
    return;
  }

  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  try {
    const confirmBtn = document.getElementById("confirmAddFolderBtn");
    if (confirmBtn) {
      confirmBtn.disabled = true;
      confirmBtn.textContent = "Creating...";
    }

    let url = "https://graph.microsoft.com/v1.0/me/mailFolders";

    // If parent folder is selected, create as child
    if (parentFolderId) {
      url = `https://graph.microsoft.com/v1.0/me/mailFolders/${parentFolderId}/childFolders`;
    }

    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        displayName: folderName,
      }),
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error?.message || "Failed to create folder");
    }

    showToast(`Folder "${folderName}" created successfully`);
    closeAddFolderModalHandler();

    // Reload folders to show the new one
    await loadMailboxFolders();
  } catch (error) {
    console.error("Failed to create folder:", error);
    showToast(`Failed to create folder: ${error.message}`);
  } finally {
    const confirmBtn = document.getElementById("confirmAddFolderBtn");
    if (confirmBtn) {
      confirmBtn.disabled = false;
      confirmBtn.textContent = "Create Folder";
    }
  }
}

async function deleteSelectedFolder() {
  if (!currentFolder) {
    showToast("Please select a folder to delete");
    return;
  }

  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  // Prevent deletion of well-known system folders
  if (currentFolder.wellKnownName) {
    showToast("Cannot delete system folders");
    return;
  }

  const confirmed = confirm(
    `Are you sure you want to delete the folder "${currentFolder.displayName}"? This will also delete all messages in the folder.`,
  );

  if (!confirmed) {
    return;
  }

  try {
    const deleteBtn = document.getElementById("deleteFolderBtn");
    if (deleteBtn) {
      deleteBtn.disabled = true;
      deleteBtn.textContent = "Deleting...";
    }

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/mailFolders/${currentFolder.id}`,
      {
        method: "DELETE",
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
        },
      },
    );

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error?.message || "Failed to delete folder");
    }

    showToast(`Folder "${currentFolder.displayName}" deleted successfully`);

    // Clear current folder and reload
    currentFolder = null;
    await loadMailboxFolders();
  } catch (error) {
    console.error("Failed to delete folder:", error);
    showToast(`Failed to delete folder: ${error.message}`);
  } finally {
    const deleteBtn = document.getElementById("deleteFolderBtn");
    if (deleteBtn) {
      deleteBtn.disabled = true;
      deleteBtn.textContent = "Delete";
    }
  }
}

function updateDeleteButtonState() {
  const deleteBtn = document.getElementById("deleteFolderBtn");
  if (!deleteBtn) {
    return;
  }

  // Enable delete button only if a folder is selected and it's not a system folder
  if (currentFolder && !currentFolder.wellKnownName) {
    deleteBtn.disabled = false;
  } else {
    deleteBtn.disabled = true;
  }
}

function showFolderContextMenu(event, folder) {
  // Remove any existing context menu
  const existingMenu = document.querySelector(".folder-context-menu");
  if (existingMenu) {
    existingMenu.remove();
  }

  // Check if folder can be deleted (system folders have wellKnownName property)
  const canDelete = !folder.wellKnownName;

  // Create context menu
  const menu = document.createElement("div");
  menu.className = "folder-context-menu";
  menu.className = "context-menu";
  menu.style.left = `${event.clientX}px`;
  menu.style.top = `${event.clientY}px`;

  // Add menu items
  const menuItems = [];

  menuItems.push({
    label: "Create Subfolder",
    icon: "➕",
    action: () => {
      // Pre-select this folder as parent
      openAddFolderModal();
      setTimeout(() => {
        const parentSelect = document.getElementById("parentFolderSelect");
        if (parentSelect) {
          parentSelect.value = folder.id;
        }
      }, 150);
    },
  });

  if (canDelete) {
    menuItems.push({
      label: "Delete Folder",
      icon: "🗑️",
      action: async () => {
        currentFolder = folder;
        await deleteSelectedFolder();
      },
      danger: true,
    });
  }

  menuItems.forEach((item) => {
    const menuItem = document.createElement("div");
    menuItem.className = "folder-context-menu-item";
    menuItem.className = item.danger
      ? "context-menu-item danger"
      : "context-menu-item";

    menuItem.innerHTML = `
      <span>${item.icon}</span>
      <span>${item.label}</span>
    `;

    menuItem.addEventListener("mouseenter", () => {
      menuItem.classList.add("hover");
    });

    menuItem.addEventListener("mouseleave", () => {
      menuItem.classList.remove("hover");
    });

    menuItem.addEventListener("click", () => {
      item.action();
      menu.remove();
    });

    menu.appendChild(menuItem);
  });

  document.body.appendChild(menu);

  // Close menu on click outside
  const closeMenu = (e) => {
    if (!menu.contains(e.target)) {
      menu.remove();
      document.removeEventListener("click", closeMenu);
    }
  };

  setTimeout(() => {
    document.addEventListener("click", closeMenu);
  }, 0);
}

// Contacts modal functionality
let allContacts = [];

function setupContactsListeners() {
  const viewContactsBtn = document.getElementById("viewContactsBtn");
  const closeContactsModalBtn = document.getElementById(
    "closeViewContactsModalBtn",
  );
  const copyAllContactsBtn = document.getElementById("copyAllContactsBtn");

  if (viewContactsBtn) {
    viewContactsBtn.addEventListener("click", openContactsModal);
  }

  if (closeContactsModalBtn) {
    closeContactsModalBtn.addEventListener("click", closeContactsModalHandler);
  }

  if (copyAllContactsBtn) {
    copyAllContactsBtn.addEventListener("click", copyAllContacts);
  }

  // Close on outside click
  const modal = document.getElementById("viewContactsModal");
  if (modal) {
    modal.addEventListener("click", (e) => {
      if (e.target === modal) {
        closeContactsModalHandler();
      }
    });
  }
}

async function openContactsModal() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const modal = document.getElementById("viewContactsModal");
  if (!modal) return;

  modal.classList.add("modal-show");

  // Load contacts
  await loadAllContacts();
}

function closeContactsModalHandler() {
  const modal = document.getElementById("viewContactsModal");
  if (modal) {
    modal.classList.remove("modal-show");
  }
}

async function loadAllContacts() {
  const contactsList = document.getElementById("contactsList");
  const contactsCount = document.getElementById("contactsCount");

  if (!contactsList) return;

  // Show loading
  contactsList.innerHTML = '<div class="loading-indicator">Loading...</div>';

  try {
    // Get contacts from user's contacts
    const contactsUrl =
      "https://graph.microsoft.com/v1.0/me/contacts?$top=500&$select=id,displayName,emailAddresses,givenName,surname,jobTitle,companyName,businessPhones,mobilePhone";

    const response = await fetch(contactsUrl, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const data = await response.json();
    allContacts = data.value || [];

    // Also get directory users for additional contacts
    try {
      const usersUrl =
        "https://graph.microsoft.com/v1.0/users?$top=500&$select=id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation";

      const usersResponse = await fetch(usersUrl, {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      });

      if (usersResponse.ok) {
        const usersData = await usersResponse.json();
        const directoryUsers = usersData.value || [];

        // Merge with contacts, avoiding duplicates
        directoryUsers.forEach((user) => {
          if (user.mail || user.userPrincipalName) {
            const exists = allContacts.some((contact) => {
              if (contact.emailAddresses && contact.emailAddresses.length > 0) {
                return contact.emailAddresses.some(
                  (email) =>
                    email.address === user.mail ||
                    email.address === user.userPrincipalName,
                );
              }
              return false;
            });

            if (!exists) {
              // Convert user to contact format
              allContacts.push({
                id: user.id,
                displayName: user.displayName,
                emailAddresses: [
                  {
                    address: user.mail || user.userPrincipalName,
                    name: user.displayName,
                  },
                ],
                jobTitle: user.jobTitle,
                companyName: user.department,
                _fromDirectory: true,
              });
            }
          }
        });
      }
    } catch (dirError) {
      // Could not load directory users
    }

    renderContactsList();

    if (contactsCount) {
      contactsCount.textContent = `${allContacts.length} contact${allContacts.length !== 1 ? "s" : ""}`;
    }
  } catch (error) {
    console.error("Error loading contacts:", error);
    showErrorInContainer(contactsList, error.message, {
      title: "Error loading contacts:",
    });
  }
}

function renderContactsList() {
  const contactsList = document.getElementById("contactsList");
  if (!contactsList) return;

  contactsList.textContent = "";

  if (allContacts.length === 0) {
    const emptyDiv = document.createElement("div");
    emptyDiv.className = "mailbox-empty";
    emptyDiv.textContent = "No contacts found";
    contactsList.appendChild(emptyDiv);
    return;
  }

  // Sort contacts by display name
  const sortedContacts = [...allContacts].sort((a, b) => {
    const nameA = (a.displayName || "").toLowerCase();
    const nameB = (b.displayName || "").toLowerCase();
    return nameA.localeCompare(nameB);
  });

  sortedContacts.forEach((contact) => {
    const contactEl = createContactElement(contact);
    contactsList.appendChild(contactEl);
  });
}

function createContactElement(contact) {
  const contactDiv = document.createElement("div");
  contactDiv.className = "directory-item";
  contactDiv.style.marginBottom = "10px";

  // Name
  const nameDiv = document.createElement("div");
  nameDiv.className = "onedrive-item-name";
  nameDiv.textContent = contact.displayName || "No Name";

  if (contact._fromDirectory) {
    const badge = document.createElement("span");
    badge.style.fontSize = "10px";
    badge.style.marginLeft = "8px";
    badge.style.padding = "2px 6px";
    badge.style.background = "var(--bg-secondary)";
    badge.style.borderRadius = "3px";
    badge.style.color = "var(--text-secondary)";
    badge.textContent = "Directory";
    nameDiv.appendChild(badge);
  }

  contactDiv.appendChild(nameDiv);

  // Email addresses
  if (contact.emailAddresses && contact.emailAddresses.length > 0) {
    contact.emailAddresses.forEach((email) => {
      if (email.address) {
        const emailDiv = document.createElement("div");
        emailDiv.className = "onedrive-item-meta";
        emailDiv.textContent = `📧 ${email.address}`;
        contactDiv.appendChild(emailDiv);
      }
    });
  }

  // Additional info
  const metaParts = [];

  if (contact.jobTitle) {
    metaParts.push(`💼 ${contact.jobTitle}`);
  }

  if (contact.companyName) {
    metaParts.push(`🏢 ${contact.companyName}`);
  }

  if (contact.businessPhones && contact.businessPhones.length > 0) {
    metaParts.push(`📞 ${contact.businessPhones[0]}`);
  } else if (contact.mobilePhone) {
    metaParts.push(`📱 ${contact.mobilePhone}`);
  }

  if (metaParts.length > 0) {
    const metaDiv = document.createElement("div");
    metaDiv.className = "onedrive-item-meta";
    metaDiv.textContent = metaParts.join(" • ");
    contactDiv.appendChild(metaDiv);
  }

  // Copy button
  const actionsDiv = document.createElement("div");
  actionsDiv.style.display = "flex";
  actionsDiv.style.gap = "8px";
  actionsDiv.style.marginTop = "8px";

  const copyBtn = document.createElement("button");
  copyBtn.className = "btn btn-small btn-secondary";
  copyBtn.style.fontSize = "11px";
  copyBtn.style.padding = "6px 10px";
  copyBtn.textContent = "📋 Copy Email";
  copyBtn.onclick = () => {
    if (contact.emailAddresses && contact.emailAddresses.length > 0) {
      copyToClipboard(contact.emailAddresses[0].address);
      showToast("Email copied to clipboard");
    }
  };
  actionsDiv.appendChild(copyBtn);

  contactDiv.appendChild(actionsDiv);

  return contactDiv;
}

async function copyAllContacts() {
  if (allContacts.length === 0) {
    showToast("No contacts to copy");
    return;
  }

  // Create a semicolon-separated list of email addresses
  const contactsText = allContacts
    .map((contact) => {
      const emails =
        contact.emailAddresses && contact.emailAddresses.length > 0
          ? contact.emailAddresses
              .map((e) => e.address)
              .filter(Boolean)
              .join(", ")
          : "";
      return emails;
    })
    .filter(Boolean)
    .join("; ");

  try {
    await copyToClipboard(contactsText);
    showToast(`Copied ${allContacts.length} contacts to clipboard`);
  } catch (error) {
    console.error("Error copying contacts:", error);
    showToast("Error copying contacts");
  }
}

function copyToClipboard(text) {
  return navigator.clipboard.writeText(text);
}
