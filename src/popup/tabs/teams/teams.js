let currentTeam = null;
let teamsChannels = [];
let teamMessages = [];
let chats = [];
let chatMessages = [];
let isLoadingTeams = false;
let currentView = "teams";
let currentChannel = null;
let currentChat = null;
let messagesNextLink = null;
let isLoadingMoreMessages = false;
let teamsEscHandler = null;

async function initializeTeams() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showTeamsNoSession();
    return;
  }
  currentView = "teams";
  updateViewButtons();
  await loadJoinedTeams();
}

function showTeamsNoSession() {
  const teamsContainer = document.getElementById("teamsContainer");
  if (!teamsContainer) return;

  teamsContainer.innerHTML = "";
  const emptyDiv = document.createElement("div");
  emptyDiv.className = "mailbox-empty";
  emptyDiv.textContent = "No active session";

  teamsContainer.appendChild(emptyDiv);
}

function updateViewButtons() {
  const teamsBtn = document.getElementById("viewTeamsBtn");
  const chatsBtn = document.getElementById("viewChatsBtn");

  if (teamsBtn && chatsBtn) {
    if (currentView === "teams") {
      teamsBtn.classList.add("active");
      chatsBtn.classList.remove("active");
    } else {
      teamsBtn.classList.remove("active");
      chatsBtn.classList.add("active");
    }
  }
}

// Load teams the user has joined
async function loadJoinedTeams() {
  if (isLoadingTeams) {
    return;
  }

  const teamsContainer = document.getElementById("teamsContainer");
  if (!teamsContainer) {
    console.error("[Teams] teamsContainer not found in DOM!");
    return;
  }
  isLoadingTeams = true;
  teamsContainer.innerHTML = "";

  const loadingDiv = document.createElement("div");
  loadingDiv.className = "loading-indicator";
  loadingDiv.textContent = "Loading...";
  teamsContainer.appendChild(loadingDiv);

  try {
    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/joinedTeams",
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      const errorText = await response.text();
      console.error("[Teams] API error response:", errorText);
      throw new Error(
        `Failed to load teams: ${response.status} ${response.statusText}`,
      );
    }

    const data = await response.json();

    const teams = data.value || [];

    renderTeams(teams);
  } catch (error) {
    console.error("[Teams] Error loading teams:", error);
    showErrorInContainer(teamsContainer, error.message, {
      title: "Error loading teams:",
    });
  } finally {
    isLoadingTeams = false;
  }
}

function renderTeams(teams) {
  const teamsContainer = document.getElementById("teamsContainer");
  if (!teamsContainer) {
    console.error("[Teams] teamsContainer not found during render!");
    return;
  }

  teamsContainer.innerHTML = "";

  if (teams.length === 0) {
    const emptyDiv = document.createElement("div");
    emptyDiv.className = "mailbox-empty";
    emptyDiv.textContent = "No teams found";
    teamsContainer.appendChild(emptyDiv);
    return;
  }
  const itemsContainer = document.createElement("div");
  itemsContainer.className = "onedrive-items-container";

  teams.forEach((team) => {
    const teamCard = createTeamCard(team);
    itemsContainer.appendChild(teamCard);
  });

  teamsContainer.appendChild(itemsContainer);
}

function createTeamCard(team) {
  const itemDiv = document.createElement("div");
  itemDiv.className = "onedrive-item";
  itemDiv.setAttribute("data-team-id", team.id);

  // Icon
  const iconDiv = document.createElement("div");
  iconDiv.className = "onedrive-item-icon";
  iconDiv.textContent = "👥";
  iconDiv.title = "Team";

  // Details
  const detailsDiv = document.createElement("div");
  detailsDiv.className = "onedrive-item-details";

  const nameDiv = document.createElement("div");
  nameDiv.className = "onedrive-item-name";
  nameDiv.textContent = team.displayName || "Untitled Team";

  const metaDiv = document.createElement("div");
  metaDiv.className = "onedrive-item-meta";
  metaDiv.textContent = team.description || "No description";

  detailsDiv.appendChild(nameDiv);
  detailsDiv.appendChild(metaDiv);

  // Actions
  const actionsDiv = document.createElement("div");
  actionsDiv.className = "onedrive-item-actions";

  const channelsBtn = document.createElement("button");
  channelsBtn.className = "btn btn-small btn-secondary btn-compact";
  channelsBtn.textContent = "📂 View Channels";
  channelsBtn.textContent = "📋 Channels";
  channelsBtn.onclick = (e) => {
    e.stopPropagation();
    loadTeamChannels(team);
  };

  const detailsBtn = document.createElement("button");
  detailsBtn.className = "btn btn-small btn-secondary btn-compact";
  detailsBtn.textContent = "ℹ️ Details";
  detailsBtn.textContent = "ℹ️ Details";
  detailsBtn.onclick = (e) => {
    e.stopPropagation();
    showTeamDetails(team);
  };

  actionsDiv.appendChild(channelsBtn);
  actionsDiv.appendChild(detailsBtn);

  itemDiv.appendChild(iconDiv);
  itemDiv.appendChild(detailsDiv);
  itemDiv.appendChild(actionsDiv);

  return itemDiv;
}

// Load channels for a team
async function loadTeamChannels(team) {
  currentTeam = team;

  const modal = document.getElementById("teamChannelsModal");
  if (!modal) return;

  const teamNameEl = document.getElementById("teamChannelsModalTeamName");
  if (teamNameEl) {
    teamNameEl.textContent = team.displayName || "Team Channels";
  }

  const channelsList = document.getElementById("teamChannelsList");
  if (!channelsList) return;

  channelsList.innerHTML = '<div class="loading-indicator">Loading...</div>';
  modal.classList.add("active");

  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/teams/${team.id}/channels`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      const errorText = await response.text();
      console.error("Channel load failed:", response.status, errorText);
      throw new Error(
        `Failed to load channels: ${response.status} ${response.statusText}`,
      );
    }

    const data = await response.json();
    teamsChannels = data.value || [];

    renderTeamChannels();
  } catch (error) {
    console.error("Error loading team channels:", error);
    showErrorInContainer(channelsList, error.message, {
      title: "Error loading channels:",
    });
  }
}

function renderTeamChannels() {
  const channelsList = document.getElementById("teamChannelsList");
  if (!channelsList) return;

  channelsList.innerHTML = "";

  if (teamsChannels.length === 0) {
    channelsList.innerHTML =
      '<div class="mailbox-empty">No channels found</div>';
    return;
  }

  const itemsContainer = document.createElement("div");
  itemsContainer.className = "onedrive-items-container";

  teamsChannels.forEach((channel) => {
    const channelDiv = document.createElement("div");
    channelDiv.className = "onedrive-item channel-item";

    // Icon
    const iconDiv = document.createElement("div");
    iconDiv.className = "onedrive-item-icon";
    iconDiv.textContent = channel.membershipType === "private" ? "🔒" : "#";
    iconDiv.title =
      channel.membershipType === "private"
        ? "Private Channel"
        : "Standard Channel";

    // Details
    const detailsDiv = document.createElement("div");
    detailsDiv.className = "onedrive-item-details";

    const nameDiv = document.createElement("div");
    nameDiv.className = "onedrive-item-name";
    nameDiv.textContent = channel.displayName || "Untitled Channel";

    const metaDiv = document.createElement("div");
    metaDiv.className = "onedrive-item-meta";
    metaDiv.textContent = channel.description || "No description";

    detailsDiv.appendChild(nameDiv);
    detailsDiv.appendChild(metaDiv);

    // Actions
    const actionsDiv = document.createElement("div");
    actionsDiv.className = "onedrive-item-actions";

    const viewMessagesBtn = document.createElement("button");
    viewMessagesBtn.className = "btn btn-small btn-primary btn-compact";
    viewMessagesBtn.textContent = "💬 View Messages";
    viewMessagesBtn.textContent = "💬 Messages";
    viewMessagesBtn.onclick = (e) => {
      e.stopPropagation();
      loadChannelMessages(channel);
    };

    actionsDiv.appendChild(viewMessagesBtn);

    channelDiv.appendChild(iconDiv);
    channelDiv.appendChild(detailsDiv);
    channelDiv.appendChild(actionsDiv);

    channelDiv.onclick = () => {
      loadChannelMessages(channel);
    };

    itemsContainer.appendChild(channelDiv);
  });

  channelsList.appendChild(itemsContainer);
}

// Load messages from a channel
async function loadChannelMessages(channel) {
  currentChannel = channel;
  currentChat = null;

  const modal = document.getElementById("channelMessagesModal");
  if (!modal) return;

  const channelNameEl = document.getElementById(
    "channelMessagesModalChannelName",
  );
  if (channelNameEl) {
    channelNameEl.textContent = channel.displayName || "Channel Messages";
  }

  const messagesList = document.getElementById("channelMessagesList");
  if (!messagesList) return;

  messagesList.innerHTML = '<div class="loading-indicator">Loading...</div>';
  modal.classList.add("active");

  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/teams/${currentTeam.id}/channels/${channel.id}/messages?$top=50`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      throw new Error(
        `Failed to load messages: ${response.status} ${response.statusText}`,
      );
    }

    const data = await response.json();
    teamMessages = data.value || [];
    messagesNextLink = data["@odata.nextLink"] || null;

    renderChannelMessages();
    setupMessageScrollListener();
  } catch (error) {
    console.error("Error loading channel messages:", error);
    showErrorInContainer(messagesList, error.message, {
      title: "Error loading messages:",
    });
  }
}

function renderChannelMessages() {
  const messagesList = document.getElementById("channelMessagesList");
  if (!messagesList) return;

  messagesList.innerHTML = "";

  if (teamMessages.length === 0) {
    messagesList.innerHTML =
      '<div class="mailbox-empty">No messages found</div>';
    return;
  }

  // Sort messages by creation date (oldest first for chat-like display)
  const sortedMessages = [...teamMessages].sort((a, b) => {
    return new Date(a.createdDateTime) - new Date(b.createdDateTime);
  });

  const messagesContainer = document.createElement("div");
  messagesContainer.className = "messages-container";

  sortedMessages.forEach((message) => {
    const messageDiv = document.createElement("div");
    messageDiv.className = "message-card";

    const header = document.createElement("div");
    header.className = "message-header";

    const authorName = message.from?.user?.displayName || "Unknown";
    const createdDate = new Date(message.createdDateTime);

    const author = document.createElement("div");
    author.className = "message-author";
    author.textContent = authorName;

    const timestamp = document.createElement("div");
    timestamp.className = "message-timestamp";
    timestamp.textContent = formatMessageDate(createdDate);

    header.appendChild(author);
    header.appendChild(timestamp);

    const bodyContent = message.body?.content || "";
    const bodyType = message.body?.contentType || "text";

    const body = document.createElement("div");
    body.className = "message-body";

    if (bodyType === "html") {
      body.textContent = stripHtmlTags(bodyContent);
    } else {
      body.textContent = bodyContent;
    }

    messageDiv.appendChild(header);
    messageDiv.appendChild(body);

    // Show attachments if any
    if (message.attachments && message.attachments.length > 0) {
      const attachmentsDiv = document.createElement("div");
      attachmentsDiv.className = "message-attachments";

      message.attachments.forEach((attachment) => {
        const attachmentItem = document.createElement("div");
        attachmentItem.className = "message-attachment-item";
        attachmentItem.textContent = `📎 ${attachment.name || "Attachment"}`;
        attachmentsDiv.appendChild(attachmentItem);
      });

      messageDiv.appendChild(attachmentsDiv);
    }

    messagesContainer.appendChild(messageDiv);
  });

  messagesList.appendChild(messagesContainer);

  // Scroll to bottom to show latest messages
  messagesList.scrollTop = messagesList.scrollHeight;

  // Add load more indicator at top if there are more messages
  if (messagesNextLink) {
    const loadMoreIndicator = document.createElement("div");
    loadMoreIndicator.id = "loadMoreIndicator";
    loadMoreIndicator.className = "load-more-indicator";
    loadMoreIndicator.textContent = "Scroll up to load older messages...";
    messagesList.insertBefore(loadMoreIndicator, messagesList.firstChild);
  }
}

// Load more messages when scrolling up
async function loadMoreMessages() {
  if (!messagesNextLink || isLoadingMoreMessages) return;

  const messagesList = document.getElementById("channelMessagesList");
  if (!messagesList) return;

  isLoadingMoreMessages = true;
  const previousScrollHeight = messagesList.scrollHeight;

  // Show loading indicator
  const loadMoreIndicator = document.getElementById("loadMoreIndicator");
  if (loadMoreIndicator) {
    loadMoreIndicator.textContent = "Loading older messages...";
  }

  try {
    const response = await fetch(messagesNextLink, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      throw new Error(`Failed to load more messages: ${response.status}`);
    }

    const data = await response.json();
    const newMessages = data.value || [];
    messagesNextLink = data["@odata.nextLink"] || null;

    if (currentChannel) {
      teamMessages = [...newMessages, ...teamMessages];
      renderChannelMessages();
    } else if (currentChat) {
      chatMessages = [...newMessages, ...chatMessages];
      renderChatMessages();
    }

    // Maintain scroll position
    messagesList.scrollTop = messagesList.scrollHeight - previousScrollHeight;

    // Update or remove load more indicator
    const loadMoreIndicator = document.getElementById("loadMoreIndicator");
    if (loadMoreIndicator) {
      if (messagesNextLink) {
        loadMoreIndicator.textContent = "Scroll up to load older messages...";
      } else {
        loadMoreIndicator.remove();
      }
    }
  } catch (error) {
    console.error("Error loading more messages:", error);
    const loadMoreIndicator = document.getElementById("loadMoreIndicator");
    if (loadMoreIndicator) {
      loadMoreIndicator.textContent = "Scroll up to load older messages...";
    }
  } finally {
    isLoadingMoreMessages = false;
  }
}

// Setup scroll listener for infinite scroll
function setupMessageScrollListener() {
  const messagesList = document.getElementById("channelMessagesList");
  if (!messagesList) return;

  // Remove existing listener if any
  messagesList.removeEventListener("scroll", handleMessageScroll);
  messagesList.addEventListener("scroll", handleMessageScroll);
}

function handleMessageScroll() {
  const messagesList = document.getElementById("channelMessagesList");
  if (!messagesList) return;

  // If scrolled near the top, load more messages
  if (
    messagesList.scrollTop < 100 &&
    messagesNextLink &&
    !isLoadingMoreMessages
  ) {
    loadMoreMessages();
  }
}

function formatMessageDate(date) {
  const now = new Date();
  const diffMs = now - date;
  const diffMins = Math.floor(diffMs / 60000);
  const diffHours = Math.floor(diffMs / 3600000);
  const diffDays = Math.floor(diffMs / 86400000);

  if (diffMins < 1) return "Just now";
  if (diffMins < 60) return `${diffMins}m ago`;
  if (diffHours < 24) return `${diffHours}h ago`;
  if (diffDays < 7) return `${diffDays}d ago`;

  return date.toLocaleDateString();
}

function stripHtmlTags(html) {
  const tmp = document.createElement("div");
  tmp.innerHTML = html;
  return tmp.textContent || tmp.innerText || "";
}

async function loadTeamMembers(team) {
  const modal = document.getElementById("teamMembersModal");
  if (!modal) return;

  const teamNameEl = document.getElementById("teamMembersModalTeamName");
  if (teamNameEl) {
    teamNameEl.textContent = team.displayName || "Team Members";
  }

  const membersList = document.getElementById("teamMembersList");
  if (!membersList) return;

  membersList.innerHTML = '<div class="loading-indicator">Loading...</div>';
  modal.classList.add("active");

  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/teams/${team.id}/members`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      throw new Error(
        `Failed to load members: ${response.status} ${response.statusText}`,
      );
    }

    const data = await response.json();
    const members = data.value || [];

    renderTeamMembers(members);
  } catch (error) {
    console.error("Error loading team members:", error);
    showErrorInContainer(membersList, error.message, {
      title: "Error loading team members:",
    });
  }
}

function renderTeamMembers(members) {
  const membersList = document.getElementById("teamMembersList");
  if (!membersList) return;

  membersList.innerHTML = "";

  if (members.length === 0) {
    membersList.innerHTML = '<div class="mailbox-empty">No members found</div>';
    return;
  }

  const itemsContainer = document.createElement("div");
  itemsContainer.className = "onedrive-items-container";

  members.forEach((member) => {
    const memberDiv = document.createElement("div");
    memberDiv.className = "onedrive-item";

    // Icon
    const iconDiv = document.createElement("div");
    iconDiv.className = "onedrive-item-icon";
    iconDiv.textContent = "👤";
    iconDiv.title = "Member";

    // Details
    const detailsDiv = document.createElement("div");
    detailsDiv.className = "onedrive-item-details";

    const nameDiv = document.createElement("div");
    nameDiv.className = "onedrive-item-name";
    nameDiv.textContent = member.displayName || "Unknown";

    const metaDiv = document.createElement("div");
    metaDiv.className = "onedrive-item-meta";
    metaDiv.textContent = member.email || member.userId || "";

    detailsDiv.appendChild(nameDiv);
    detailsDiv.appendChild(metaDiv);

    // Role badge
    const actionsDiv = document.createElement("div");
    actionsDiv.className = "onedrive-item-actions";

    const roleSpan = document.createElement("span");
    roleSpan.className = "role-badge";
    roleSpan.textContent =
      member.roles && member.roles.includes("owner") ? "Owner" : "Member";

    actionsDiv.appendChild(roleSpan);

    memberDiv.appendChild(iconDiv);
    memberDiv.appendChild(detailsDiv);
    memberDiv.appendChild(actionsDiv);

    itemsContainer.appendChild(memberDiv);
  });

  membersList.appendChild(itemsContainer);
}

// Show team details
function showTeamDetails(team) {
  const modal = document.getElementById("teamDetailsModal");
  if (!modal) return;

  const detailsContent = document.getElementById("teamDetailsContent");
  if (!detailsContent) return;

  detailsContent.innerHTML = "";

  const details = [
    { label: "Name", value: team.displayName || "N/A" },
    { label: "Description", value: team.description || "No description" },
    { label: "Team ID", value: team.id || "N/A" },
    { label: "Visibility", value: team.visibility || "N/A" },
    {
      label: "Web URL",
      value: team.webUrl || "N/A",
      link: team.webUrl || null,
    },
  ];

  details.forEach((detail) => {
    const row = document.createElement("div");
    row.className = "detail-row";

    const label = document.createElement("div");
    label.className = "detail-label";
    label.textContent = detail.label;

    const value = document.createElement("div");
    value.className = "detail-value";

    if (detail.link) {
      const link = document.createElement("a");
      try {
        const url = new URL(detail.link);
        if (url.protocol === "http:" || url.protocol === "https:") {
          link.href = detail.link;
          link.target = "_blank";
          link.rel = "noopener noreferrer";
          link.textContent = detail.value;
          link.style.color = "var(--primary-color)";
          value.appendChild(link);
        } else {
          // unexpected protocol, display as text
          value.textContent = detail.value;
        }
      } catch (e) {
        // Invalid URL, display as text
        value.textContent = detail.value;
      }
    } else {
      value.textContent = detail.value;
    }

    row.appendChild(label);
    row.appendChild(value);
    detailsContent.appendChild(row);
  });

  modal.classList.add("active");
}

async function loadChats() {
  if (isLoadingTeams) return;

  const teamsContainer = document.getElementById("teamsContainer");
  if (!teamsContainer) return;

  isLoadingTeams = true;
  teamsContainer.innerHTML = "";

  const loadingDiv = document.createElement("div");
  loadingDiv.className = "loading-indicator";
  loadingDiv.textContent = "Loading...";
  teamsContainer.appendChild(loadingDiv);

  try {
    const response = await fetch("https://graph.microsoft.com/v1.0/me/chats", {
      method: "GET",
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      throw new Error(
        `Failed to load chats: ${response.status} ${response.statusText}`,
      );
    }

    const data = await response.json();
    chats = data.value || [];

    renderChats(chats);
  } catch (error) {
    console.error("Error loading chats:", error);
    showErrorInContainer(teamsContainer, error.message, {
      title: "Error loading chats:",
    });
  } finally {
    isLoadingTeams = false;
  }
}

async function renderChats(chatsData) {
  const teamsContainer = document.getElementById("teamsContainer");
  if (!teamsContainer) return;

  teamsContainer.innerHTML = "";

  if (chatsData.length === 0) {
    const emptyDiv = document.createElement("div");
    emptyDiv.className = "mailbox-empty";
    emptyDiv.textContent = "No chats found";
    teamsContainer.appendChild(emptyDiv);
    return;
  }

  const itemsContainer = document.createElement("div");
  itemsContainer.className = "onedrive-items-container";

  // Load chat cards sequentially to get member names
  for (const chat of chatsData) {
    const chatCard = await createChatCard(chat);
    itemsContainer.appendChild(chatCard);
  }

  teamsContainer.appendChild(itemsContainer);
}

async function createChatCard(chat) {
  const itemDiv = document.createElement("div");
  itemDiv.className = "onedrive-item";
  itemDiv.setAttribute("data-chat-id", chat.id);

  // Icon
  const iconDiv = document.createElement("div");
  iconDiv.className = "onedrive-item-icon";
  iconDiv.textContent = chat.chatType === "group" ? "👥" : "💬";
  iconDiv.title = chat.chatType === "group" ? "Group Chat" : "One-on-One Chat";

  // Details
  const detailsDiv = document.createElement("div");
  detailsDiv.className = "onedrive-item-details";

  const nameDiv = document.createElement("div");
  nameDiv.className = "onedrive-item-name";

  // If no topic, fetch members and build a name
  let chatName = chat.topic || "";
  if (!chatName) {
    try {
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/chats/${chat.id}/members`,
        {
          method: "GET",
          headers: {
            Authorization: `Bearer ${activeM365Session.access_token}`,
            "Content-Type": "application/json",
          },
        },
      );
      if (response.ok) {
        const data = await response.json();
        const members = data.value || [];
        const otherMembers = members.filter((m) => m.userId && m.displayName);
        if (otherMembers.length > 0) {
          chatName = otherMembers.map((m) => m.displayName).join(", ");
        } else {
          chatName = "Chat";
        }
      } else {
        chatName = "Chat";
      }
    } catch (error) {
      console.error("Error fetching chat members:", error);
      chatName = "Chat";
    }
  }

  nameDiv.textContent = chatName;

  const metaDiv = document.createElement("div");
  metaDiv.className = "onedrive-item-meta";
  const lastUpdated = chat.lastUpdatedDateTime
    ? new Date(chat.lastUpdatedDateTime)
    : null;
  metaDiv.textContent = lastUpdated
    ? `Last updated: ${formatMessageDate(lastUpdated)}`
    : "No recent activity";

  detailsDiv.appendChild(nameDiv);
  detailsDiv.appendChild(metaDiv);

  // Actions
  const actionsDiv = document.createElement("div");
  actionsDiv.className = "onedrive-item-actions";

  const messagesBtn = document.createElement("button");
  messagesBtn.className = "btn btn-small btn-secondary";
  messagesBtn.style.fontSize = "11px";
  messagesBtn.style.padding = "6px 10px";
  messagesBtn.textContent = "💬 Messages";
  messagesBtn.onclick = (e) => {
    e.stopPropagation();
    loadChatMessages(chat);
  };

  actionsDiv.appendChild(messagesBtn);

  itemDiv.appendChild(iconDiv);
  itemDiv.appendChild(detailsDiv);
  itemDiv.appendChild(actionsDiv);

  return itemDiv;
}

// Load messages from a chat
async function loadChatMessages(chat) {
  currentChat = chat;
  currentChannel = null;

  const modal = document.getElementById("channelMessagesModal");
  if (!modal) return;

  const chatNameEl = document.getElementById("channelMessagesModalChannelName");
  if (chatNameEl) {
    chatNameEl.textContent = chat.topic || "Chat Messages";
  }

  const messagesList = document.getElementById("channelMessagesList");
  if (!messagesList) return;

  messagesList.innerHTML = '<div class="loading-indicator">Loading...</div>';
  modal.classList.add("active");

  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/chats/${chat.id}/messages?$top=50`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      throw new Error(
        `Failed to load messages: ${response.status} ${response.statusText}`,
      );
    }

    const data = await response.json();
    chatMessages = data.value || [];
    messagesNextLink = data["@odata.nextLink"] || null;

    renderChatMessages();
    setupMessageScrollListener();
  } catch (error) {
    console.error("Error loading chat messages:", error);
    showErrorInContainer(messagesList, error.message, {
      title: "Error loading chat messages:",
    });
  }
}

function renderChatMessages() {
  const messagesList = document.getElementById("channelMessagesList");
  if (!messagesList) return;

  messagesList.innerHTML = "";

  if (chatMessages.length === 0) {
    messagesList.innerHTML =
      '<div class="mailbox-empty">No messages found</div>';
    return;
  }

  // Sort messages by creation date (oldest first for chat-like display)
  const sortedMessages = [...chatMessages].sort((a, b) => {
    return new Date(a.createdDateTime) - new Date(b.createdDateTime);
  });

  const messagesContainer = document.createElement("div");
  messagesContainer.style.cssText =
    "display: flex; flex-direction: column; gap: 8px;";

  sortedMessages.forEach((message) => {
    const messageDiv = document.createElement("div");
    messageDiv.style.cssText = `
      background: var(--surface);
      border: 1px solid var(--border-color);
      border-radius: 8px;
      padding: 10px 12px;
    `;

    const header = document.createElement("div");
    header.style.cssText =
      "display: flex; align-items: center; gap: 8px; margin-bottom: 6px;";

    const authorName = message.from?.user?.displayName || "Unknown";
    const createdDate = new Date(message.createdDateTime);

    const author = document.createElement("div");
    author.style.cssText =
      "font-weight: 600; color: var(--text-primary); font-size: 13px;";
    author.textContent = authorName;

    const timestamp = document.createElement("div");
    timestamp.style.cssText = "font-size: 11px; color: var(--text-secondary);";
    timestamp.textContent = formatMessageDate(createdDate);

    header.appendChild(author);
    header.appendChild(timestamp);

    const bodyContent = message.body?.content || "";
    const bodyType = message.body?.contentType || "text";

    const body = document.createElement("div");
    body.style.cssText =
      "color: var(--text-primary); font-size: 14px; line-height: 1.4; word-wrap: break-word;";

    if (bodyType === "html") {
      body.textContent = stripHtmlTags(bodyContent);
    } else {
      body.textContent = bodyContent;
    }

    messageDiv.appendChild(header);
    messageDiv.appendChild(body);

    // Show attachments if any
    if (message.attachments && message.attachments.length > 0) {
      const attachmentsDiv = document.createElement("div");
      attachmentsDiv.style.cssText =
        "margin-top: 6px; padding-top: 6px; border-top: 1px solid var(--border-color);";

      message.attachments.forEach((attachment) => {
        const attachmentItem = document.createElement("div");
        attachmentItem.style.cssText =
          "font-size: 12px; color: var(--primary); margin-top: 4px;";
        attachmentItem.textContent = `📎 ${attachment.name || "Attachment"}`;
        attachmentsDiv.appendChild(attachmentItem);
      });

      messageDiv.appendChild(attachmentsDiv);
    }

    messagesContainer.appendChild(messageDiv);
  });

  messagesList.appendChild(messagesContainer);

  // Scroll to bottom to show latest messages
  messagesList.scrollTop = messagesList.scrollHeight;

  // Add load more indicator at top if there are more messages
  if (messagesNextLink) {
    const loadMoreIndicator = document.createElement("div");
    loadMoreIndicator.id = "loadMoreIndicator";
    loadMoreIndicator.style.cssText = `
      text-align: center;
      padding: 10px;
      color: var(--text-secondary);
      font-size: 12px;
      font-style: italic;
    `;
    loadMoreIndicator.textContent = "Scroll up to load older messages...";
    messagesList.insertBefore(loadMoreIndicator, messagesList.firstChild);
  }
}

// Close message modal and clean up
function closeMessageModal() {
  const modal = document.getElementById("channelMessagesModal");
  if (modal) {
    modal.classList.remove("active");
    currentChannel = null;
    currentChat = null;
    messagesNextLink = null;

    // Clear compose input
    const composeInput = document.getElementById("composeMessageInput");
    if (composeInput) {
      composeInput.value = "";
    }
  }
}

async function loadChatMembers(chat) {
  const modal = document.getElementById("teamMembersModal");
  if (!modal) return;

  const chatNameEl = document.getElementById("teamMembersModalTeamName");
  if (chatNameEl) {
    chatNameEl.textContent = chat.topic || "Chat Members";
  }

  const membersList = document.getElementById("teamMembersList");
  if (!membersList) return;

  membersList.innerHTML = '<div class="loading-indicator">Loading...</div>';
  modal.classList.add("active");

  try {
    const response = await fetch(
      `https://graph.microsoft.com/me/chats/${chat.id}/members`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      throw new Error(
        `Failed to load members: ${response.status} ${response.statusText}`,
      );
    }

    const data = await response.json();
    const members = data.value || [];

    renderChatMembers(members);
  } catch (error) {
    console.error("Error loading chat members:", error);
    showErrorInContainer(membersList, error.message, {
      title: "Error loading chat members:",
    });
  }
}

function renderChatMembers(members) {
  const membersList = document.getElementById("teamMembersList");
  if (!membersList) return;

  membersList.innerHTML = "";

  if (members.length === 0) {
    membersList.innerHTML = '<div class="mailbox-empty">No members found</div>';
    return;
  }

  const itemsContainer = document.createElement("div");
  itemsContainer.className = "onedrive-items-container";

  members.forEach((member) => {
    const memberDiv = document.createElement("div");
    memberDiv.className = "onedrive-item";

    // Icon
    const iconDiv = document.createElement("div");
    iconDiv.className = "onedrive-item-icon";
    iconDiv.textContent = "👤";
    iconDiv.title = "Member";

    // Details
    const detailsDiv = document.createElement("div");
    detailsDiv.className = "onedrive-item-details";

    const nameDiv = document.createElement("div");
    nameDiv.className = "onedrive-item-name";
    nameDiv.textContent = member.displayName || "Unknown";

    const metaDiv = document.createElement("div");
    metaDiv.className = "onedrive-item-meta";
    metaDiv.textContent = member.email || member.userId || "";

    detailsDiv.appendChild(nameDiv);
    detailsDiv.appendChild(metaDiv);

    memberDiv.appendChild(iconDiv);
    memberDiv.appendChild(detailsDiv);

    itemsContainer.appendChild(memberDiv);
  });

  membersList.appendChild(itemsContainer);
}

// Send a message to a channel or chat
async function sendMessage() {
  const messageInput = document.getElementById("composeMessageInput");
  const sendBtn = document.getElementById("sendMessageBtn");

  if (!messageInput || !sendBtn) return;

  const messageText = messageInput.value.trim();
  if (!messageText) return;

  // Disable input while sending
  sendBtn.disabled = true;
  messageInput.disabled = true;
  sendBtn.textContent = "Sending...";

  try {
    let url, body;

    if (currentChannel && currentTeam) {
      // Send to channel
      url = `https://graph.microsoft.com/v1.0/teams/${currentTeam.id}/channels/${currentChannel.id}/messages`;
      body = {
        body: {
          content: messageText,
        },
      };
    } else if (currentChat) {
      // Send to chat
      url = `https://graph.microsoft.com/v1.0/me/chats/${currentChat.id}/messages`;
      body = {
        body: {
          content: messageText,
        },
      };
    } else {
      throw new Error("No active channel or chat");
    }

    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      throw new Error(`Failed to send message: ${response.status}`);
    }

    // Clear input
    messageInput.value = "";

    // Reload messages to show the new one
    if (currentChannel && currentTeam) {
      await loadChannelMessages(currentChannel);
    } else if (currentChat) {
      await loadChatMessages(currentChat);
    }
  } catch (error) {
    console.error("Error sending message:", error);
    showToast(`Failed to send message: ${error.message}`, "error");
  } finally {
    sendBtn.disabled = false;
    messageInput.disabled = false;
    sendBtn.textContent = "Send";
    messageInput.focus();
  }
}

function setupTeamsListeners() {
  // Close buttons for modals
  const closeChannelsModal = document.getElementById("closeTeamChannelsModal");
  if (closeChannelsModal) {
    closeChannelsModal.onclick = () => {
      const modal = document.getElementById("teamChannelsModal");
      if (modal) modal.classList.remove("active");
    };
  }

  const closeMessagesModal = document.getElementById(
    "closeChannelMessagesModal",
  );
  if (closeMessagesModal) {
    closeMessagesModal.onclick = closeMessageModal;
  }

  const closeMembersModal = document.getElementById("closeTeamMembersModal");
  if (closeMembersModal) {
    closeMembersModal.onclick = () => {
      const modal = document.getElementById("teamMembersModal");
      if (modal) modal.classList.remove("active");
    };
  }

  const closeDetailsModal = document.getElementById("closeTeamDetailsModal");
  if (closeDetailsModal) {
    closeDetailsModal.onclick = () => {
      const modal = document.getElementById("teamDetailsModal");
      if (modal) modal.classList.remove("active");
    };
  }

  // Close modals when clicking outside
  const modals = [
    "teamChannelsModal",
    "channelMessagesModal",
    "teamMembersModal",
    "teamDetailsModal",
  ];

  modals.forEach((modalId) => {
    const modal = document.getElementById(modalId);
    if (modal) {
      modal.addEventListener("click", (e) => {
        if (e.target === modal) {
          if (modalId === "channelMessagesModal") {
            closeMessageModal();
          } else {
            modal.classList.remove("active");
          }
        }
      });
    }
  });

  // ESC key support for closing modals
  // Remove existing listener if any
  if (teamsEscHandler) {
    document.removeEventListener("keydown", teamsEscHandler);
  }

  teamsEscHandler = (e) => {
    if (e.key === "Escape") {
      // Check which modal is open and close it
      const channelMessagesModal = document.getElementById(
        "channelMessagesModal",
      );
      if (
        channelMessagesModal &&
        channelMessagesModal.classList.contains("active")
      ) {
        closeMessageModal();
        return;
      }

      const teamChannelsModal = document.getElementById("teamChannelsModal");
      if (teamChannelsModal && teamChannelsModal.classList.contains("active")) {
        teamChannelsModal.classList.remove("active");
        return;
      }

      const teamMembersModal = document.getElementById("teamMembersModal");
      if (teamMembersModal && teamMembersModal.classList.contains("active")) {
        teamMembersModal.classList.remove("active");
        return;
      }

      const teamDetailsModal = document.getElementById("teamDetailsModal");
      if (teamDetailsModal && teamDetailsModal.classList.contains("active")) {
        teamDetailsModal.classList.remove("active");
        return;
      }
    }
  };

  document.addEventListener("keydown", teamsEscHandler);

  // Refresh button
  const refreshTeamsBtn = document.getElementById("refreshTeamsBtn");
  if (refreshTeamsBtn) {
    refreshTeamsBtn.onclick = async () => {
      refreshTeamsBtn.disabled = true;
      refreshTeamsBtn.textContent = "⏳";
      if (currentView === "teams") {
        await loadJoinedTeams();
      } else {
        await loadChats();
      }
      refreshTeamsBtn.disabled = false;
      refreshTeamsBtn.textContent = "🔄";
    };
  }

  // View Teams button
  const viewTeamsBtn = document.getElementById("viewTeamsBtn");
  if (viewTeamsBtn) {
    viewTeamsBtn.addEventListener("click", async () => {
      currentView = "teams";
      updateViewButtons();
      await loadJoinedTeams();
    });
  }

  // View Chats button
  const viewChatsBtn = document.getElementById("viewChatsBtn");
  if (viewChatsBtn) {
    viewChatsBtn.addEventListener("click", async () => {
      currentView = "chats";
      updateViewButtons();
      await loadChats();
    });
  }

  // Refresh messages button in modal
  const refreshMessagesBtn = document.getElementById("refreshMessagesBtn");
  if (refreshMessagesBtn) {
    refreshMessagesBtn.addEventListener("click", async () => {
      refreshMessagesBtn.disabled = true;
      const originalText = refreshMessagesBtn.textContent;
      refreshMessagesBtn.textContent = "⏳";

      try {
        if (currentChannel && currentTeam) {
          await loadChannelMessages(currentChannel);
        } else if (currentChat) {
          await loadChatMessages(currentChat);
        }
      } finally {
        refreshMessagesBtn.disabled = false;
        refreshMessagesBtn.textContent = originalText;
      }
    });
  }

  // Send message button
  const sendMessageBtn = document.getElementById("sendMessageBtn");
  if (sendMessageBtn) {
    sendMessageBtn.addEventListener("click", sendMessage);
  }

  // Message input keyboard shortcuts
  const composeMessageInput = document.getElementById("composeMessageInput");
  if (composeMessageInput) {
    composeMessageInput.addEventListener("keydown", (e) => {
      // Enter without Shift sends the message
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        sendMessage();
      }
      // Shift+Enter adds a new line (default behavior)
    });
  }
}
