let sharePointSites = [];
let currentSPSite = null;
let currentSPSiteId = null;
let currentSPDriveId = null;
let sharePointItems = [];
let sharePointFolderStack = [];
let sharePointSearchQuery = "";
let sharePointSearchResults = [];
let isSharePointSearching = false;

async function initializeSharePoint() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showSharePointNoSession();
    return;
  }

  await loadSharePointSites();

  if (sharePointSites.length > 0) {
    const rootSite = sharePointSites.find(
      (site) => site.id && site.id.includes("root"),
    );
    if (rootSite) {
      await selectSite(rootSite);
    } else {
      await selectSite(sharePointSites[0]);
    }
  }
}

function showSharePointNoSession() {
  const sitesList = document.getElementById("sharePointSitesList");
  const container = document.getElementById("sharePointContainer");

  if (sitesList) {
    sitesList.textContent = "";
    const emptyDiv = document.createElement("div");
    emptyDiv.className = "mailbox-empty";
    emptyDiv.textContent = "No active session";
    sitesList.appendChild(emptyDiv);
  }

  if (container) {
    container.textContent = "";
    const emptyDiv = document.createElement("div");
    emptyDiv.className = "mailbox-empty";
    emptyDiv.textContent = "No active session";
    container.appendChild(emptyDiv);
  }
}

async function loadSharePointSites() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const sitesList = document.getElementById("sharePointSitesList");
  if (!sitesList) return;

  // Show loading
  sitesList.textContent = "";
  const loadingDiv = document.createElement("div");
  loadingDiv.className = "loading-indicator";
  loadingDiv.textContent = "Loading...";
  sitesList.appendChild(loadingDiv);

  try {
    const allSites = new Map();

    // 1. Get followed sites
    try {
      const followedUrl =
        "https://graph.microsoft.com/v1.0/me/followedSites?$select=id,name,displayName,webUrl,createdDateTime";

      const followedResponse = await fetch(followedUrl, {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      });

      if (followedResponse.ok) {
        const followedData = await followedResponse.json();
        (followedData.value || []).forEach((site) => {
          allSites.set(site.id, site);
        });
      }
    } catch (e) {
      // Could not load followed sites
    }

    // 2. Get all sites via search
    try {
      const searchUrl =
        "https://graph.microsoft.com/v1.0/sites?search=*&$select=id,name,displayName,webUrl,createdDateTime&$top=100";

      const searchResponse = await fetch(searchUrl, {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      });

      if (searchResponse.ok) {
        const searchData = await searchResponse.json();
        (searchData.value || []).forEach((site) => {
          allSites.set(site.id, site);
        });
      }
    } catch (e) {
      // Could not search sites
    }

    // 3. Get root site
    try {
      const rootUrl =
        "https://graph.microsoft.com/v1.0/sites/root?$select=id,name,displayName,webUrl,createdDateTime";

      const rootResponse = await fetch(rootUrl, {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      });

      if (rootResponse.ok) {
        const rootSite = await rootResponse.json();
        allSites.set(rootSite.id, rootSite);
      }
    } catch (e) {
      // Could not load root site
    }

    sharePointSites = Array.from(allSites.values());

    if (sharePointSites.length === 0) {
      throw new Error("No sites found");
    }

    renderSharePointSites();
  } catch (error) {
    console.error("Error loading sites:", error);
    showToast(`Failed to load sites: ${error.message}`, "error");
    if (sitesList) {
      showErrorInContainer(sitesList, error.message, {
        title: "Error loading sites:",
      });
    }
  }
}

function renderSharePointSites() {
  const sitesList = document.getElementById("sharePointSitesList");
  if (!sitesList) return;

  sitesList.innerHTML = "";

  if (sharePointSites.length === 0) {
    const emptyDiv = document.createElement("div");
    emptyDiv.className = "mailbox-empty";
    emptyDiv.textContent = "No sites found";
    sitesList.appendChild(emptyDiv);
    return;
  }

  const sortedSites = [...sharePointSites].sort((a, b) => {
    const nameA = a.displayName || a.name || "";
    const nameB = b.displayName || b.name || "";
    return nameA.localeCompare(nameB);
  });

  sortedSites.forEach((site) => {
    const siteEl = createSiteElement(site);
    sitesList.appendChild(siteEl);
  });
}

function createSiteElement(site) {
  const siteDiv = document.createElement("div");
  siteDiv.className = "sharepoint-site-item";
  siteDiv.textContent = site.displayName || site.name || "Unnamed Site";
  siteDiv.setAttribute("data-site-id", site.id);

  siteDiv.onclick = () => selectSite(site);

  return siteDiv;
}

async function selectSite(site) {
  console.log("Selecting site:", site);

  currentSPSite = site;
  currentSPSiteId = site.id;
  currentSPDriveId = null;
  sharePointItems = [];
  sharePointFolderStack = [];

  const container = document.getElementById("sharePointContainer");
  const breadcrumb = document.getElementById("sharePointBreadcrumb");

  if (breadcrumb) {
    breadcrumb.style.display = "none";
  }

  const siteElements = document.querySelectorAll(
    ".sharepoint-site-item.active",
  );
  siteElements.forEach((el) => el.classList.remove("active"));

  const selectedElement = document.querySelector(`[data-site-id="${site.id}"]`);
  if (selectedElement) {
    selectedElement.classList.add("active");
  }

  const label = document.getElementById("sharePointSiteLabel");
  if (label) {
    label.textContent = `📁 ${site.displayName || site.name}`;
  }

  await loadSiteDocumentLibraries(site.id);
}

async function loadSiteDocumentLibraries(siteId) {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const container = document.getElementById("sharePointContainer");
  if (!container) return;

  // Show loading
  container.textContent = "";
  const loadingDiv = document.createElement("div");
  loadingDiv.className = "loading-indicator";
  loadingDiv.textContent = "Loading...";
  container.appendChild(loadingDiv);

  // Hide breadcrumb
  const breadcrumb = document.getElementById("sharePointBreadcrumb");
  if (breadcrumb) {
    breadcrumb.classList.add("hidden");
  }

  try {
    // Get drives (document libraries) for the site
    const drivesUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=id,name,description,webUrl,driveType`;

    const response = await fetch(drivesUrl, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const data = await response.json();
    const drives = data.value || [];

    container.textContent = "";

    if (drives.length === 0) {
      const emptyDiv = document.createElement("div");
      emptyDiv.className = "mailbox-empty";
      emptyDiv.textContent = "No document libraries found in this site";
      container.appendChild(emptyDiv);
      return;
    }

    // Display drives as clickable cards
    drives.forEach((drive) => {
      const driveEl = createDriveElement(drive);
      container.appendChild(driveEl);
    });
  } catch (error) {
    console.error("Error loading document libraries:", error);
    showToast(`Failed to load libraries: ${error.message}`, "error");
    if (container) {
      showErrorInContainer(container, error.message, {
        title: "Error loading document libraries:",
      });
    }
  }
}

function createDriveElement(drive) {
  const driveDiv = document.createElement("div");
  driveDiv.className = "sharepoint-drive-item";

  const nameDiv = document.createElement("div");
  nameDiv.style.fontSize = "14px";
  nameDiv.style.fontWeight = "600";
  nameDiv.style.color = "var(--text-primary)";
  nameDiv.textContent = `📚 ${drive.name}`;

  const descDiv = document.createElement("div");
  descDiv.style.fontSize = "12px";
  descDiv.style.color = "var(--text-secondary)";
  descDiv.style.marginTop = "4px";

  let driveTypeDisplay = "Document Library";
  if (drive.driveType === "documentLibrary") {
    driveTypeDisplay = "Document Library";
  } else if (drive.driveType) {
    driveTypeDisplay =
      drive.driveType.charAt(0).toUpperCase() + drive.driveType.slice(1);
  }

  descDiv.textContent = drive.description || driveTypeDisplay;

  driveDiv.appendChild(nameDiv);
  driveDiv.appendChild(descDiv);

  driveDiv.onclick = () => {
    currentSPDriveId = drive.id;
    sharePointFolderStack = [];
    loadSharePointItems(drive.id);
  };

  return driveDiv;
}

async function loadSharePointItems(driveId, folderId = "root") {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const container = document.getElementById("sharePointContainer");
  if (!container) return;

  // Show loading
  container.textContent = "";
  const loadingDiv = document.createElement("div");
  loadingDiv.className = "loading-indicator";
  loadingDiv.textContent = "Loading...";
  container.appendChild(loadingDiv);

  try {
    let url;
    if (folderId === "root") {
      url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$select=id,name,folder,file,size,webUrl,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy&$top=200&$orderby=name`;
    } else {
      url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderId}/children?$select=id,name,folder,file,size,webUrl,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy&$top=200&$orderby=name`;
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

    const data = await response.json();
    sharePointItems = data.value || [];

    // Update breadcrumb
    updateSharePointBreadcrumb();

    renderSharePointItems();
  } catch (error) {
    console.error("Error loading folder:", error);
    showToast(`Failed to load folder: ${error.message}`, "error");
    if (container) {
      showErrorInContainer(container, error.message, {
        title: "Error loading folder:",
      });
    }
  }
}

function updateSharePointBreadcrumb() {
  const breadcrumb = document.getElementById("sharePointBreadcrumb");
  if (!breadcrumb) return;

  breadcrumb.innerHTML = "";

  const siteName = currentSPSite
    ? currentSPSite.displayName || currentSPSite.name
    : "SharePoint";

  const homeSpan = document.createElement("span");
  homeSpan.className = "breadcrumb-item";
  homeSpan.textContent = siteName;
  homeSpan.style.cursor = "pointer";
  homeSpan.onclick = () => {
    clearSharePointSearch();
    sharePointFolderStack = [];
    loadSiteDocumentLibraries(currentSPSiteId);
  };
  breadcrumb.appendChild(homeSpan);

  sharePointFolderStack.forEach((folder, index) => {
    const separator = document.createElement("span");
    separator.textContent = " / ";
    separator.style.color = "var(--text-secondary)";
    breadcrumb.appendChild(separator);

    const folderSpan = document.createElement("span");
    folderSpan.className = "breadcrumb-item";
    folderSpan.textContent = folder.name;
    folderSpan.style.cursor = "pointer";
    folderSpan.onclick = () => {
      clearSharePointSearch();
      sharePointFolderStack = sharePointFolderStack.slice(0, index + 1);
      loadSharePointItems(currentSPDriveId, folder.id);
    };
    breadcrumb.appendChild(folderSpan);
  });
}

function renderSharePointItems() {
  const container = document.getElementById("sharePointContainer");
  if (!container) return;

  const itemsToRender = isSharePointSearching
    ? sharePointSearchResults
    : sharePointItems;

  if (itemsToRender.length === 0) {
    const emptyDiv = document.createElement("div");
    emptyDiv.className = "mailbox-empty";
    emptyDiv.textContent = isSharePointSearching
      ? "No items match your search"
      : "This folder is empty";
    container.innerHTML = "";
    container.appendChild(emptyDiv);
    return;
  }

  const folders = itemsToRender.filter((item) => item.folder);
  const files = itemsToRender.filter((item) => !item.folder);

  container.innerHTML = "";

  if (isSharePointSearching) {
    const headerDiv = document.createElement("div");
    headerDiv.style.marginBottom = "15px";
    headerDiv.style.padding = "8px 12px";
    headerDiv.style.background = "var(--bg-secondary)";
    headerDiv.style.borderRadius = "4px";
    headerDiv.style.display = "flex";
    headerDiv.style.justifyContent = "space-between";
    headerDiv.style.alignItems = "center";

    const resultsSpan = document.createElement("span");
    resultsSpan.style.fontSize = "13px";
    resultsSpan.style.color = "var(--text-secondary)";
    resultsSpan.textContent = `Found ${itemsToRender.length} result${itemsToRender.length !== 1 ? "s" : ""}`;

    const backBtn = document.createElement("button");
    backBtn.className = "btn btn-small btn-secondary btn-compact";
    backBtn.textContent = "← Back to folder";
    backBtn.onclick = clearSharePointSearch;

    headerDiv.appendChild(resultsSpan);
    headerDiv.appendChild(backBtn);
    container.appendChild(headerDiv);
  }

  folders.forEach((item) => {
    const itemEl = createSharePointItemElement(item);
    container.appendChild(itemEl);
  });

  files.forEach((item) => {
    const itemEl = createSharePointItemElement(item);
    container.appendChild(itemEl);
  });
}

function createSharePointItemElement(item) {
  const itemDiv = document.createElement("div");
  const isFolder = !!item.folder;

  itemDiv.className = `onedrive-item ${isFolder ? "onedrive-item-folder" : ""}`;
  itemDiv.setAttribute("data-item-id", item.id);
  itemDiv.setAttribute("data-item-name", item.name);
  itemDiv.setAttribute("data-is-folder", isFolder);

  if (isFolder) {
    itemDiv.onclick = () => {
      clearSharePointSearch();
      sharePointFolderStack.push({ id: item.id, name: item.name });
      loadSharePointItems(currentSPDriveId, item.id);
    };
  }

  const iconDiv = document.createElement("div");
  iconDiv.className = "onedrive-item-icon";
  iconDiv.textContent = isFolder ? "📁" : getFileIcon(item.name);

  const detailsDiv = document.createElement("div");
  detailsDiv.className = "onedrive-item-details";

  const nameDiv = document.createElement("div");
  nameDiv.className = "onedrive-item-name";
  nameDiv.title = item.name;
  nameDiv.textContent = item.name;

  const metaDiv = document.createElement("div");
  metaDiv.className = "onedrive-item-meta";

  const metaParts = [];

  if (!isFolder && item.size) {
    metaParts.push(formatFileSize(item.size));
  } else if (isFolder) {
    metaParts.push("Folder");
  }

  if (item.lastModifiedDateTime) {
    const date = new Date(item.lastModifiedDateTime);
    metaParts.push(`Modified ${formatDate(date)}`);
  }

  metaDiv.textContent = metaParts.join(" • ");

  detailsDiv.appendChild(nameDiv);
  detailsDiv.appendChild(metaDiv);

  const actionsDiv = document.createElement("div");
  actionsDiv.className = "onedrive-item-actions";

  if (!isFolder) {
    const copyLinkBtn = document.createElement("button");
    copyLinkBtn.className = "btn btn-small btn-secondary btn-compact";
    copyLinkBtn.textContent = "🔗 Copy Link";
    copyLinkBtn.onclick = async (e) => {
      e.stopPropagation();
      if (item.webUrl) {
        try {
          await navigator.clipboard.writeText(item.webUrl);
          showToast("Link copied to clipboard!");
        } catch (error) {
          showToast("Failed to copy link");
        }
      }
    };
    actionsDiv.appendChild(copyLinkBtn);

    const downloadBtn = document.createElement("button");
    downloadBtn.className = "btn btn-small btn-primary btn-compact";
    downloadBtn.textContent = "⬇️ Download";
    downloadBtn.onclick = (e) => {
      e.stopPropagation();
      downloadSharePointItem(item);
    };
    actionsDiv.appendChild(downloadBtn);
  }

  itemDiv.appendChild(iconDiv);
  itemDiv.appendChild(detailsDiv);
  itemDiv.appendChild(actionsDiv);

  return itemDiv;
}

async function downloadSharePointItem(item) {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  try {
    showToast("Downloading...");

    const url = `https://graph.microsoft.com/v1.0/drives/${currentSPDriveId}/items/${item.id}/content`;

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const blob = await response.blob();
    const downloadUrl = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = downloadUrl;
    a.download = item.name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(downloadUrl);

    showToast("✅ Downloaded");
  } catch (error) {
    console.error("Download error:", error);
    showToast(`Failed to download: ${error.message}`, "error");
  }
}

function getFileIcon(fileName) {
  const ext = fileName.split(".").pop().toLowerCase();
  const icons = {
    pdf: "📄",
    doc: "📝",
    docx: "📝",
    xls: "📊",
    xlsx: "📊",
    ppt: "📊",
    pptx: "📊",
    txt: "📄",
    zip: "🗜️",
    rar: "🗜️",
    jpg: "🖼️",
    jpeg: "🖼️",
    png: "🖼️",
    gif: "🖼️",
    mp4: "🎥",
    mp3: "🎵",
    wav: "🎵",
  };

  return icons[ext] || "📄";
}

function formatFileSize(bytes) {
  if (bytes === 0) return "0 Bytes";
  const k = 1024;
  const sizes = ["Bytes", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + " " + sizes[i];
}

function formatDate(date) {
  const now = new Date();
  const diffMs = now - date;
  const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

  if (diffDays === 0) return "Today";
  if (diffDays === 1) return "Yesterday";
  if (diffDays < 7) return `${diffDays} days ago`;
  if (diffDays < 30) return `${Math.floor(diffDays / 7)} weeks ago`;
  if (diffDays < 365) return `${Math.floor(diffDays / 30)} months ago`;
  return date.toLocaleDateString();
}

async function searchSharePointSite(query) {
  if (!query || !query.trim()) {
    clearSharePointSearch();
    return;
  }

  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  if (!currentSPSiteId || !currentSPDriveId) {
    showToast("Please select a site and document library first");
    return;
  }

  try {
    isSharePointSearching = true;
    sharePointSearchQuery = query.trim();

    const container = document.getElementById("sharePointContainer");
    if (container) {
      container.innerHTML = "";
      const loadingDiv = document.createElement("div");
      loadingDiv.className = "loading-indicator";
      loadingDiv.textContent = "Searching...";
      container.appendChild(loadingDiv);
    }

    const searchUrl = `https://graph.microsoft.com/v1.0/drives/${currentSPDriveId}/root/search(q='${encodeURIComponent(query)}')?$select=id,name,folder,file,size,webUrl,createdDateTime,lastModifiedDateTime&$top=200`;

    const response = await fetch(searchUrl, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const data = await response.json();
    sharePointSearchResults = data.value || [];

    renderSharePointItems();
  } catch (error) {
    console.error("Search error:", error);
    showToast(`Failed to search: ${error.message}`, "error");
    isSharePointSearching = false;
    sharePointSearchResults = [];
    renderSharePointItems();
  }
}

function clearSharePointSearch() {
  isSharePointSearching = false;
  sharePointSearchQuery = "";
  sharePointSearchResults = [];

  const searchInput = document.getElementById("sharePointSearch");
  if (searchInput) {
    searchInput.value = "";
  }

  renderSharePointItems();
}

function setupSharePointSearch() {
  const searchInput = document.getElementById("sharePointSearch");
  if (!searchInput) return;

  let searchTimeout;

  searchInput.addEventListener("input", (e) => {
    clearTimeout(searchTimeout);
    const query = e.target.value.trim();

    if (!query) {
      clearSharePointSearch();
      return;
    }

    searchTimeout = setTimeout(() => {
      searchSharePointSite(query);
    }, 500);
  });

  searchInput.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      clearSharePointSearch();
    }
  });
}

function setupSharePointListeners() {
  const refreshBtn = document.getElementById("refreshSharePointBtn");
  if (refreshBtn) {
    refreshBtn.onclick = async () => {
      if (currentSPDriveId) {
        const currentFolderId =
          sharePointFolderStack.length > 0
            ? sharePointFolderStack[sharePointFolderStack.length - 1].id
            : null;
        await loadSharePointItems(currentSPDriveId, currentFolderId);
      } else if (currentSPSiteId) {
        await loadSiteDocumentLibraries(currentSPSiteId);
      } else {
        await loadSharePointSites();
      }
    };
  }

  setupSharePointSearch();
}
