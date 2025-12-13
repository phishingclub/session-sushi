let currentDriveItems = [];
let currentDrivePath = [];
let currentDriveId = "root";
let oneDriveSearchResults = [];

async function initializeOneDrive() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showOneDriveNoSession();
    return;
  }

  updateOneDriveSessionInfo();
  await loadDriveFolder("root");
}

function showOneDriveNoSession() {
  const container = document.getElementById("oneDriveContainer");
  if (container) {
    container.innerHTML = `
      <div class="mailbox-empty">
        <p class="mb-15">No active session</p>
      </div>
    `;
  }
}

function updateOneDriveSessionInfo() {
  // Info bar updated via banner, no need to show user info here
}

async function loadDriveFolder(folderId, folderName = "OneDrive") {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const container = document.getElementById("oneDriveContainer");
  if (container) {
    container.innerHTML = '<div class="loading-indicator">Loading...</div>';
  }

  try {
    let url;
    if (folderId === "root") {
      url = "https://graph.microsoft.com/v1.0/me/drive/root/children?$top=200";
    } else {
      url = `https://graph.microsoft.com/v1.0/me/drive/items/${folderId}/children?$top=200`;
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
    currentDriveItems = data.value || [];
    currentDriveId = folderId;

    if (folderId === "root") {
      currentDrivePath = [{ id: "root", name: "OneDrive" }];
    } else {
      const existingIndex = currentDrivePath.findIndex(
        (p) => p.id === folderId,
      );
      if (existingIndex >= 0) {
        currentDrivePath = currentDrivePath.slice(0, existingIndex + 1);
      } else {
        currentDrivePath.push({ id: folderId, name: folderName });
      }
    }

    renderDriveItems();
    renderBreadcrumb();
    updateOneDriveStats();
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

function renderBreadcrumb() {
  const breadcrumbContainer = document.getElementById("oneDriveBreadcrumb");
  if (!breadcrumbContainer) return;

  breadcrumbContainer.innerHTML = "";

  currentDrivePath.forEach((item, index) => {
    const isLast = index === currentDrivePath.length - 1;

    const breadcrumbItem = document.createElement("span");
    breadcrumbItem.className = `breadcrumb-item ${isLast ? "active" : ""}`;
    breadcrumbItem.setAttribute("data-folder-id", item.id);
    breadcrumbItem.setAttribute("data-folder-name", item.name);
    // Cursor is handled by CSS classes
    breadcrumbItem.textContent = item.name;

    if (!isLast) {
      breadcrumbItem.addEventListener("click", () => {
        const folderId = breadcrumbItem.getAttribute("data-folder-id");
        const folderName = breadcrumbItem.getAttribute("data-folder-name");
        loadDriveFolder(folderId, folderName);
      });
    }

    breadcrumbContainer.appendChild(breadcrumbItem);

    if (!isLast) {
      const separator = document.createElement("span");
      separator.className = "breadcrumb-separator";
      separator.textContent = "›";
      breadcrumbContainer.appendChild(separator);
    }
  });
}

function renderDriveItems() {
  const container = document.getElementById("oneDriveContainer");
  if (!container) return;

  if (currentDriveItems.length === 0) {
    container.innerHTML =
      '<div class="mailbox-empty">This folder is empty</div>';
    return;
  }

  // Sort: folders first, then files
  const sortedItems = [...currentDriveItems].sort((a, b) => {
    const aIsFolder = !!a.folder;
    const bIsFolder = !!b.folder;
    if (aIsFolder && !bIsFolder) return -1;
    if (!aIsFolder && bIsFolder) return 1;
    return a.name.localeCompare(b.name);
  });

  container.innerHTML = "";
  const itemsContainer = document.createElement("div");
  itemsContainer.className = "onedrive-items-container";

  sortedItems.forEach((item) => {
    const isFolder = !!item.folder;
    const icon = isFolder ? "📁" : getFileIconForOneDrive(item.name);
    const size = isFolder
      ? `${item.folder.childCount} items`
      : formatFileSize(item.size || 0);
    const modified = item.lastModifiedDateTime
      ? new Date(item.lastModifiedDateTime).toLocaleString()
      : "Unknown";

    const itemDiv = document.createElement("div");
    itemDiv.className = `onedrive-item ${isFolder ? "onedrive-item-folder" : ""}`;
    itemDiv.setAttribute("data-item-id", item.id);
    itemDiv.setAttribute("data-item-name", item.name);
    itemDiv.setAttribute("data-is-folder", isFolder);

    // Icon
    const iconDiv = document.createElement("div");
    iconDiv.className = "onedrive-item-icon";
    iconDiv.textContent = icon;

    // Details
    const detailsDiv = document.createElement("div");
    detailsDiv.className = "onedrive-item-details";

    const nameDiv = document.createElement("div");
    nameDiv.className = "onedrive-item-name";
    nameDiv.title = item.name;
    nameDiv.textContent = item.name;

    const metaDiv = document.createElement("div");
    metaDiv.className = "onedrive-item-meta";
    metaDiv.textContent = `${size} • Modified ${modified}`;

    detailsDiv.appendChild(nameDiv);
    detailsDiv.appendChild(metaDiv);

    // Actions
    const actionsDiv = document.createElement("div");
    actionsDiv.className = "onedrive-item-actions";

    const downloadBtn = document.createElement("button");
    downloadBtn.className = "btn btn-small btn-secondary btn-compact";
    downloadBtn.setAttribute("data-action", "download");
    downloadBtn.setAttribute("data-item-id", item.id);
    downloadBtn.setAttribute("data-is-folder", isFolder);
    downloadBtn.textContent = "⬇️ Download";
    actionsDiv.appendChild(downloadBtn);

    const detailsBtn = document.createElement("button");
    detailsBtn.className = "btn btn-small btn-secondary btn-compact";
    detailsBtn.setAttribute("data-action", "details");
    detailsBtn.setAttribute("data-item-id", item.id);
    detailsBtn.textContent = "ℹ️ Details";

    const deleteBtn = document.createElement("button");
    deleteBtn.className = "btn btn-small btn-danger btn-compact";
    deleteBtn.setAttribute("data-action", "delete");
    deleteBtn.setAttribute("data-item-id", item.id);
    deleteBtn.textContent = "🗑️ Delete";

    actionsDiv.appendChild(detailsBtn);
    actionsDiv.appendChild(deleteBtn);

    itemDiv.appendChild(iconDiv);
    itemDiv.appendChild(detailsDiv);
    itemDiv.appendChild(actionsDiv);

    itemsContainer.appendChild(itemDiv);
  });

  container.appendChild(itemsContainer);
}

function updateOneDriveStats() {
  const statsEl = document.getElementById("oneDriveStats");
  if (!statsEl) return;

  const folderCount = currentDriveItems.filter((item) => item.folder).length;
  const fileCount = currentDriveItems.filter((item) => item.file).length;
  const totalSize = currentDriveItems
    .filter((item) => item.size)
    .reduce((sum, item) => sum + item.size, 0);

  statsEl.textContent = `${folderCount} folders, ${fileCount} files (${formatFileSize(totalSize)})`;
}

// Download file or folder from OneDrive
async function downloadDriveItem(itemId, itemName, isFolder) {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  if (isFolder) {
    const confirmed = confirm(
      `Download folder "${itemName}"?\n\n` +
        "Note: The Microsoft Graph API does not support downloading folders as zip files directly. " +
        "We will recursively download each file one by one and package them into a zip file.\n\n" +
        "This might take some time depending on the folder size. Continue?",
    );

    if (!confirmed) return;

    await downloadFolderRecursive(itemId, itemName);
  } else {
    await downloadSingleFile(itemId, itemName);
  }
}

async function downloadSingleFile(itemId, itemName, showToastMessage = true) {
  if (showToastMessage) {
    showToast("Downloading...");
  }

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/content`,
    {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
      },
    },
  );

  if (response.status === 429) {
    const retryAfter = parseInt(response.headers.get("Retry-After") || "5");
    throw { isThrottled: true, retryAfter };
  }

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  }

  const blob = await response.blob();

  if (showToastMessage) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = itemName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    showToast("✅ Downloaded");
  }

  return { blob, name: itemName };
}

async function downloadFolderRecursive(folderId, folderName) {
  try {
    showToast("Preparing folder download...");

    const allFiles = [];
    const statusDiv = createDownloadStatusDiv();

    await collectFilesRecursive(folderId, "", allFiles, statusDiv);

    if (allFiles.length === 0) {
      showToast("Folder is empty", "error");
      document.body.removeChild(statusDiv);
      return;
    }

    updateDownloadStatus(
      statusDiv,
      `Downloading ${allFiles.length} files...`,
      0,
      allFiles.length,
    );

    const downloadedFiles = [];
    for (let i = 0; i < allFiles.length; i++) {
      const fileInfo = allFiles[i];
      updateDownloadStatus(
        statusDiv,
        `Downloading: ${fileInfo.path}`,
        i,
        allFiles.length,
      );

      try {
        const { blob, name } = await downloadSingleFile(
          fileInfo.id,
          fileInfo.name,
          false,
        );
        downloadedFiles.push({ blob, path: fileInfo.path });
      } catch (error) {
        if (error.isThrottled) {
          updateDownloadStatus(
            statusDiv,
            `Rate limited. Waiting ${error.retryAfter} seconds...`,
            i,
            allFiles.length,
          );
          await sleep(error.retryAfter * 1000);
          i--;
          continue;
        }
        console.error(`Failed to download ${fileInfo.path}:`, error);
      }
    }

    updateDownloadStatus(
      statusDiv,
      "Creating zip file...",
      allFiles.length,
      allFiles.length,
    );

    const zipBlob = await createZipFile(downloadedFiles, folderName);

    const url = window.URL.createObjectURL(zipBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${folderName}.zip`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);

    document.body.removeChild(statusDiv);
    showToast(`✅ Downloaded folder with ${downloadedFiles.length} files`);
  } catch (error) {
    console.error("Folder download error:", error);
    showToast(`Failed to download folder: ${error.message}`, "error");
  }
}

async function collectFilesRecursive(
  folderId,
  currentPath,
  allFiles,
  statusDiv,
) {
  updateDownloadStatus(
    statusDiv,
    `Scanning: ${currentPath || "root"}...`,
    0,
    1,
  );

  let url = `https://graph.microsoft.com/v1.0/me/drive/items/${folderId}/children?$top=200`;

  while (url) {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (response.status === 429) {
      const retryAfter = parseInt(response.headers.get("Retry-After") || "5");
      updateDownloadStatus(
        statusDiv,
        `Rate limited. Waiting ${retryAfter} seconds...`,
        0,
        1,
      );
      await sleep(retryAfter * 1000);
      continue;
    }

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const data = await response.json();

    for (const item of data.value) {
      const itemPath = currentPath ? `${currentPath}/${item.name}` : item.name;

      if (item.folder) {
        await collectFilesRecursive(item.id, itemPath, allFiles, statusDiv);
      } else {
        allFiles.push({
          id: item.id,
          name: item.name,
          path: itemPath,
        });
      }
    }

    url = data["@odata.nextLink"];
  }
}

function createDownloadStatusDiv() {
  const statusDiv = document.createElement("div");
  statusDiv.className = "modal modal-show";
  statusDiv.style.zIndex = "10000";

  statusDiv.innerHTML = `
    <div class="modal-content" style="max-width: 400px;">
      <div class="modal-header">
        <h2>📦 Downloading Folder</h2>
      </div>
      <div class="modal-body" style="text-align: center;">
        <div id="downloadStatusText" style="margin-bottom: 15px; font-weight: 500; color: var(--text-primary);">Preparing...</div>
        <div id="downloadProgressBar" style="width: 100%; height: 10px; background: var(--bg-secondary); border-radius: 5px; overflow: hidden; margin-bottom: 10px;">
          <div id="downloadProgressFill" style="width: 0%; height: 100%; background: var(--primary-color); transition: width 0.3s;"></div>
        </div>
        <div id="downloadStatusCount" style="font-size: 13px; color: var(--text-secondary);"></div>
      </div>
    </div>
  `;

  document.body.appendChild(statusDiv);
  return statusDiv;
}

function updateDownloadStatus(statusDiv, text, current, total) {
  const textEl = statusDiv.querySelector("#downloadStatusText");
  const fillEl = statusDiv.querySelector("#downloadProgressFill");
  const countEl = statusDiv.querySelector("#downloadStatusCount");

  if (textEl) textEl.textContent = text;

  if (total > 0 && fillEl) {
    const percent = (current / total) * 100;
    fillEl.style.width = `${percent}%`;
  }

  if (countEl && total > 0) {
    countEl.textContent = `${current} / ${total}`;
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function createZipFile(files, folderName) {
  const zip = new ZipWriter();

  for (const file of files) {
    await zip.addFile(file.path, file.blob);
  }

  return await zip.generate();
}

class ZipWriter {
  constructor() {
    this.files = [];
  }

  async addFile(path, blob) {
    const arrayBuffer = await blob.arrayBuffer();
    this.files.push({
      path: path,
      data: new Uint8Array(arrayBuffer),
    });
  }

  async generate() {
    const centralDirectory = [];
    const fileDataParts = [];
    let offset = 0;

    // Process each file to create ZIP entries
    for (const file of this.files) {
      const fileName = new TextEncoder().encode(file.path);
      const fileData = file.data;
      const crc32 = this.calculateCRC32(fileData);
      const now = new Date();
      const dosTime = this.toDOSTime(now);
      const dosDate = this.toDOSDate(now);

      // Create Local File Header (30 bytes + filename length)
      // This header precedes each file's data in the ZIP archive
      const localHeader = new Uint8Array(30 + fileName.length);
      const view = new DataView(localHeader.buffer);

      // Local file header signature: 0x04034b50
      view.setUint32(0, 0x04034b50, true);
      // Version needed to extract (2.0)
      view.setUint16(4, 20, true);
      // General purpose bit flag (0 = no flags)
      view.setUint16(6, 0, true);
      // Compression method (0 = stored/no compression)
      view.setUint16(8, 0, true);
      // File last modification time (DOS format)
      view.setUint16(10, dosTime, true);
      // File last modification date (DOS format)
      view.setUint16(12, dosDate, true);
      // CRC-32 checksum of uncompressed data
      view.setUint32(14, crc32, true);
      // Compressed size (same as uncompressed since no compression)
      view.setUint32(18, fileData.length, true);
      // Uncompressed size
      view.setUint32(22, fileData.length, true);
      // File name length
      view.setUint16(26, fileName.length, true);
      // Extra field length (0 = no extra field)
      view.setUint16(28, 0, true);
      // Append the actual filename
      localHeader.set(fileName, 30);

      // Add local header and file data to the ZIP structure
      fileDataParts.push(localHeader);
      fileDataParts.push(fileData);

      // Create Central Directory Header (46 bytes + filename length)
      // The central directory contains metadata about all files in the archive
      const cdHeader = new Uint8Array(46 + fileName.length);
      const cdView = new DataView(cdHeader.buffer);

      // Central directory file header signature: 0x02014b50
      cdView.setUint32(0, 0x02014b50, true);
      // Version made by (2.0)
      cdView.setUint16(4, 20, true);
      // Version needed to extract (2.0)
      cdView.setUint16(6, 20, true);
      // General purpose bit flag
      cdView.setUint16(8, 0, true);
      // Compression method (0 = no compression)
      cdView.setUint16(10, 0, true);
      // File last modification time
      cdView.setUint16(12, dosTime, true);
      // File last modification date
      cdView.setUint16(14, dosDate, true);
      // CRC-32 checksum
      cdView.setUint32(16, crc32, true);
      // Compressed size
      cdView.setUint32(20, fileData.length, true);
      // Uncompressed size
      cdView.setUint32(24, fileData.length, true);
      // File name length
      cdView.setUint16(28, fileName.length, true);
      // Extra field length
      cdView.setUint16(30, 0, true);
      // File comment length
      cdView.setUint16(32, 0, true);
      // Disk number where file starts
      cdView.setUint16(34, 0, true);
      // Internal file attributes
      cdView.setUint16(36, 0, true);
      // External file attributes
      cdView.setUint32(38, 0, true);
      // Relative offset of local file header (byte offset from start of ZIP)
      cdView.setUint32(42, offset, true);
      // Append filename
      cdHeader.set(fileName, 46);

      centralDirectory.push(cdHeader);

      // Track the current offset for the next file
      offset += localHeader.length + fileData.length;
    }

    // Calculate total size of central directory
    const cdSize = centralDirectory.reduce((sum, cd) => sum + cd.length, 0);

    // Create End of Central Directory Record (22 bytes)
    // This marks the end of the ZIP file and contains archive metadata
    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer);
    // End of central directory signature: 0x06054b50
    eocdView.setUint32(0, 0x06054b50, true);
    // Number of this disk
    eocdView.setUint16(4, 0, true);
    // Disk where central directory starts
    eocdView.setUint16(6, 0, true);
    // Number of central directory records on this disk
    eocdView.setUint16(8, this.files.length, true);
    // Total number of central directory records
    eocdView.setUint16(10, this.files.length, true);
    // Size of central directory (bytes)
    eocdView.setUint32(12, cdSize, true);
    // Offset of start of central directory
    eocdView.setUint32(16, offset, true);
    // ZIP file comment length
    eocdView.setUint16(20, 0, true);

    // Assemble complete ZIP file: [file data][central directory][end of central directory]
    const allParts = [...fileDataParts, ...centralDirectory, eocd];
    const totalLength = allParts.reduce((sum, part) => sum + part.length, 0);
    const zipData = new Uint8Array(totalLength);

    // Copy all parts into final ZIP data
    let position = 0;
    for (const part of allParts) {
      zipData.set(part, position);
      position += part.length;
    }

    return new Blob([zipData], { type: "application/zip" });
  }

  calculateCRC32(data) {
    // CRC32 is a checksum algorithm used to detect data corruption
    const crcTable = this.makeCRCTable();
    let crc = 0xffffffff; // Start with all bits set

    // Process each byte of data
    for (let i = 0; i < data.length; i++) {
      crc = (crc >>> 8) ^ crcTable[(crc ^ data[i]) & 0xff];
    }

    // Final XOR and convert to unsigned 32-bit integer
    return (crc ^ 0xffffffff) >>> 0;
  }

  makeCRCTable() {
    // Cache the CRC table to avoid recalculating it
    if (this.crcTable) return this.crcTable;

    // Generate lookup table for CRC32 calculation (256 entries)
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      // Calculate CRC for this byte value
      for (let j = 0; j < 8; j++) {
        // Polynomial: 0xEDB88320 (reversed 0x04C11DB7)
        c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      }
      table[i] = c;
    }

    this.crcTable = table;
    return table;
  }

  toDOSTime(date) {
    // Convert JavaScript Date to DOS time format (16-bit)
    // Bits 15-11: hours (0-23), Bits 10-5: minutes (0-59), Bits 4-0: seconds/2 (0-29)
    return (
      ((date.getHours() << 11) |
        (date.getMinutes() << 5) |
        (date.getSeconds() >> 1)) &
      0xffff
    );
  }

  toDOSDate(date) {
    // Convert JavaScript Date to DOS date format (16-bit)
    // Bits 15-9: year-1980 (0-127), Bits 8-5: month (1-12), Bits 4-0: day (1-31)
    return (
      (((date.getFullYear() - 1980) << 9) |
        ((date.getMonth() + 1) << 5) |
        date.getDate()) &
      0xffff
    );
  }
}

// Delete file or folder
async function deleteDriveItem(itemId, itemName) {
  if (
    !confirm(
      `Are you sure you want to delete "${itemName}"? This action cannot be undone.`,
    )
  ) {
    return;
  }

  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}`,
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

    showToast(`✅ Deleted ${itemName}`);
    await loadDriveFolder(currentDriveId);
  } catch (error) {
    console.error("Delete error:", error);
    showToast(`Failed to delete: ${error.message}`, "error");
  }
}

// Upload file to OneDrive
function openUploadModal() {
  const modal = document.getElementById("uploadFileModal");
  if (modal) {
    modal.classList.add("modal-show");
    document.getElementById("uploadFileInput").value = "";
  }
}

function closeUploadModal() {
  const modal = document.getElementById("uploadFileModal");
  if (modal) {
    modal.classList.remove("modal-show");
  }
}

async function uploadFileToDrive() {
  const fileInput = document.getElementById("uploadFileInput");
  const file = fileInput.files[0];

  if (!file) {
    showToast("Please select a file");
    return;
  }

  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const uploadBtn = document.getElementById("confirmUpload");
  if (uploadBtn) uploadBtn.disabled = true;

  try {
    showToast(`Uploading ${file.name}...`);

    // For files under 4MB, use simple upload
    if (file.size < 4 * 1024 * 1024) {
      const url =
        currentDriveId === "root"
          ? `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(file.name)}:/content`
          : `https://graph.microsoft.com/v1.0/me/drive/items/${currentDriveId}:/${encodeURIComponent(file.name)}:/content`;

      const response = await fetch(url, {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/octet-stream",
        },
        body: file,
      });

      if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || response.statusText);
      }

      showToast(`✅ Uploaded ${file.name}`);
    } else {
      // For larger files, use upload session
      await uploadLargeFile(file);
    }

    closeUploadModal();
    await loadDriveFolder(currentDriveId);
  } catch (error) {
    console.error("Upload error:", error);
    showToast(`Upload failed: ${error.message}`, "error");
  } finally {
    if (uploadBtn) uploadBtn.disabled = false;
  }
}

async function uploadLargeFile(file) {
  // Create upload session
  const url =
    currentDriveId === "root"
      ? `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(file.name)}:/createUploadSession`
      : `https://graph.microsoft.com/v1.0/me/drive/items/${currentDriveId}:/${encodeURIComponent(file.name)}:/createUploadSession`;

  const sessionResponse = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${activeM365Session.access_token}`,
      "Content-Type": "application/json",
    },
  });

  if (!sessionResponse.ok) {
    throw new Error("Failed to create upload session");
  }

  const session = await sessionResponse.json();
  const uploadUrl = session.uploadUrl;

  // Upload in chunks (10MB chunks)
  const chunkSize = 10 * 1024 * 1024;
  let offset = 0;

  while (offset < file.size) {
    const chunk = file.slice(offset, offset + chunkSize);
    const endByte = Math.min(offset + chunkSize, file.size);

    const chunkResponse = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": chunk.size,
        "Content-Range": `bytes ${offset}-${endByte - 1}/${file.size}`,
      },
      body: chunk,
    });

    if (!chunkResponse.ok && chunkResponse.status !== 202) {
      throw new Error(`Upload chunk failed: ${chunkResponse.statusText}`);
    }

    offset = endByte;
    const progress = Math.round((offset / file.size) * 100);
    showToast(`Uploading ${file.name}: ${progress}%`);
  }

  showToast(`✅ Uploaded ${file.name}`);
}

// Create new folder
function openCreateFolderModal() {
  const modal = document.getElementById("createFolderModal");
  const input = document.getElementById("newFolderName");
  if (modal && input) {
    input.value = "";
    modal.classList.add("modal-show");
    setTimeout(() => input.focus(), 100);
  }
}

function closeCreateFolderModal() {
  const modal = document.getElementById("createFolderModal");
  const input = document.getElementById("newOneDriveFolderName");
  if (modal) {
    modal.classList.remove("modal-show");
  }
  if (input) {
    input.value = "";
  }
}

async function createNewDriveFolder() {
  const folderInput = document.getElementById("newOneDriveFolderName");

  if (!folderInput) {
    console.error("Folder input element not found");
    showToast("Error: Folder input not found");
    return;
  }

  const folderName = folderInput.value?.trim() || "";

  if (!folderName || folderName.length === 0) {
    showToast("Please enter a folder name");
    folderInput.focus();
    return;
  }

  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const createBtn = document.getElementById("confirmCreateFolder");
  if (createBtn) createBtn.disabled = true;

  try {
    const url =
      currentDriveId === "root"
        ? "https://graph.microsoft.com/v1.0/me/drive/root/children"
        : `https://graph.microsoft.com/v1.0/me/drive/items/${currentDriveId}/children`;

    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        name: folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename",
      }),
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error?.message || response.statusText);
    }

    showToast(`✅ Created folder "${folderName}"`);
    closeCreateFolderModal();
    await loadDriveFolder(currentDriveId);
  } catch (error) {
    console.error("Create folder error:", error);
    showToast(`Failed to create folder: ${error.message}`, "error");
  } finally {
    if (createBtn) createBtn.disabled = false;
  }
}

// Search OneDrive
async function searchOneDrive() {
  const query = document.getElementById("oneDriveSearch").value.trim();

  if (!query) {
    // Empty search - return to normal folder view
    clearSearch();
    return;
  }

  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const container = document.getElementById("oneDriveContainer");
  if (container) {
    container.innerHTML = '<div class="loading-indicator">Searching...</div>';
  }

  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/root/search(q='${encodeURIComponent(query)}')?$top=100`,
      {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const data = await response.json();
    oneDriveSearchResults = data.value || [];

    renderSearchResults();
  } catch (error) {
    console.error("Search error:", error);
    showToast(`Search failed: ${error.message}`, "error");
    if (container) {
      showErrorInContainer(container, error.message, {
        title: "Search failed:",
      });
    }
  }
}

function renderSearchResults() {
  const container = document.getElementById("oneDriveContainer");
  if (!container) return;

  if (oneDriveSearchResults.length === 0) {
    container.innerHTML = '<div class="mailbox-empty">No results found</div>';
    return;
  }

  container.innerHTML = "";

  // Header with back button
  const headerDiv = document.createElement("div");
  headerDiv.className = "search-results-header";

  const backBtn = document.createElement("button");
  backBtn.className = "btn btn-secondary";
  backBtn.id = "clearSearchBtn";
  backBtn.textContent = "← Back to folder";

  const resultsSpan = document.createElement("span");
  resultsSpan.className = "search-results-count";
  resultsSpan.textContent = `Found ${oneDriveSearchResults.length} results`;

  headerDiv.appendChild(backBtn);
  headerDiv.appendChild(resultsSpan);

  // Items container
  const itemsContainer = document.createElement("div");
  itemsContainer.className = "onedrive-items-container";

  oneDriveSearchResults.forEach((item) => {
    const isFolder = !!item.folder;
    const icon = isFolder ? "📁" : getFileIconForOneDrive(item.name);
    const size = isFolder
      ? `${item.folder?.childCount || 0} items`
      : formatFileSize(item.size || 0);
    const modified = item.lastModifiedDateTime
      ? new Date(item.lastModifiedDateTime).toLocaleString()
      : "Unknown";
    const path = item.parentReference?.path || "";
    const pathDisplay = path.split("/").pop() || "Root";

    const itemDiv = document.createElement("div");
    itemDiv.className = "onedrive-item";
    itemDiv.setAttribute("data-item-id", item.id);
    itemDiv.setAttribute("data-item-name", item.name);

    // Icon
    const iconDiv = document.createElement("div");
    iconDiv.className = "onedrive-item-icon";
    iconDiv.textContent = icon;

    // Details
    const detailsDiv = document.createElement("div");
    detailsDiv.className = "onedrive-item-details";

    const nameDiv = document.createElement("div");
    nameDiv.className = "onedrive-item-name";
    nameDiv.title = item.name;
    nameDiv.textContent = item.name;

    const metaDiv = document.createElement("div");
    metaDiv.className = "onedrive-item-meta";
    metaDiv.textContent = `${size} • ${pathDisplay} • Modified ${modified}`;

    detailsDiv.appendChild(nameDiv);
    detailsDiv.appendChild(metaDiv);

    // Actions
    const actionsDiv = document.createElement("div");
    actionsDiv.className = "onedrive-item-actions";

    const downloadBtn = document.createElement("button");
    downloadBtn.className = "btn btn-small btn-secondary btn-compact";
    downloadBtn.setAttribute("data-action", "download");
    downloadBtn.setAttribute("data-item-id", item.id);
    downloadBtn.setAttribute("data-is-folder", isFolder);
    downloadBtn.textContent = "⬇️ Download";
    actionsDiv.appendChild(downloadBtn);

    const detailsBtn = document.createElement("button");
    detailsBtn.className = "btn btn-small btn-secondary btn-compact";
    detailsBtn.setAttribute("data-action", "details");
    detailsBtn.setAttribute("data-item-id", item.id);
    detailsBtn.textContent = "ℹ️ Details";

    actionsDiv.appendChild(detailsBtn);

    itemDiv.appendChild(iconDiv);
    itemDiv.appendChild(detailsDiv);
    itemDiv.appendChild(actionsDiv);

    itemsContainer.appendChild(itemDiv);
  });

  container.appendChild(headerDiv);
  container.appendChild(itemsContainer);

  // Add event listener for clear search button
  document
    .getElementById("clearSearchBtn")
    ?.addEventListener("click", clearSearch);

  // Setup event delegation for search result items
  setupSearchResultListeners();
}

function clearSearch() {
  document.getElementById("oneDriveSearch").value = "";
  oneDriveSearchResults = [];
  loadDriveFolder(currentDriveId);
}

// Show item details
async function showItemDetails(itemId) {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}`,
      {
        headers: {
          Authorization: `Bearer ${activeM365Session.access_token}`,
          "Content-Type": "application/json",
        },
      },
    );

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const item = await response.json();
    displayItemDetailsModal(item);
  } catch (error) {
    console.error("Get item details error:", error);
    showToast(`Failed to get item details: ${error.message}`, "error");
  }
}

function displayItemDetailsModal(item) {
  const modal = document.getElementById("itemDetailsModal");
  const content = document.getElementById("itemDetailsContent");

  if (!modal || !content) return;

  const isFolder = !!item.folder;
  const details = {
    Name: item.name,
    Type: isFolder ? "Folder" : "File",
    Size: isFolder
      ? `${item.folder.childCount} items`
      : formatFileSize(item.size || 0),
    Created: new Date(item.createdDateTime).toLocaleString(),
    Modified: new Date(item.lastModifiedDateTime).toLocaleString(),
    "Created By": item.createdBy?.user?.displayName || "Unknown",
    "Modified By": item.lastModifiedBy?.user?.displayName || "Unknown",
    ID: item.id,
    "Web URL": item.webUrl || "N/A",
  };

  if (item.file) {
    details["MIME Type"] = item.file.mimeType || "Unknown";
    if (item.file.hashes) {
      details["SHA1"] = item.file.hashes.sha1Hash || "N/A";
      details["QuickXor"] = item.file.hashes.quickXorHash || "N/A";
    }
  }

  // Build details using DOM manipulation to avoid XSS
  content.innerHTML = "";

  Object.entries(details).forEach(([key, value]) => {
    const detailRow = document.createElement("div");
    detailRow.className = "detail-row";

    const label = document.createElement("div");
    label.className = "detail-label";
    label.textContent = key + ":";

    const valueDiv = document.createElement("div");
    valueDiv.className = "detail-value";
    if (key === "Web URL" || key === "ID") {
      valueDiv.classList.add("word-break-all");
    }
    valueDiv.textContent = String(value);

    detailRow.appendChild(label);
    detailRow.appendChild(valueDiv);
    content.appendChild(detailRow);
  });

  modal.classList.add("modal-show");
}

function closeItemDetailsModal() {
  const modal = document.getElementById("itemDetailsModal");
  if (modal) {
    modal.classList.remove("modal-show");
  }
}

// Get drive quota/storage info
async function getDriveQuota() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  try {
    const response = await fetch("https://graph.microsoft.com/v1.0/me/drive", {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const data = await response.json();
    displayQuotaInfo(data.quota);
  } catch (error) {
    console.error("Get quota error:", error);
    showToast(`Failed to get quota: ${error.message}`, "error");
  }
}

function displayQuotaInfo(quota) {
  if (!quota) return;

  const total = formatFileSize(quota.total || 0);
  const used = formatFileSize(quota.used || 0);
  const remaining = formatFileSize(quota.remaining || 0);
  const percentage = quota.total
    ? Math.round((quota.used / quota.total) * 100)
    : 0;

  const message = `Storage: ${used} used of ${total} (${percentage}% used) • Remaining: ${remaining}`;

  showToast(message, "info");
}

// Utility function
function getFileIconForOneDrive(filename) {
  const ext = filename.split(".").pop()?.toLowerCase() || "";

  const icons = {
    pdf: "📄",
    doc: "📘",
    docx: "📘",
    xls: "📗",
    xlsx: "📗",
    csv: "📗",
    ppt: "📙",
    pptx: "📙",
    txt: "📝",
    zip: "📦",
    rar: "📦",
    "7z": "📦",
    jpg: "🖼️",
    jpeg: "🖼️",
    png: "🖼️",
    gif: "🖼️",
    bmp: "🖼️",
    svg: "🖼️",
    mp4: "🎬",
    avi: "🎬",
    mov: "🎬",
    wmv: "🎬",
    mp3: "🎵",
    wav: "🎵",
    flac: "🎵",
    html: "🌐",
    htm: "🌐",
    xml: "🌐",
    json: "🌐",
    js: "📜",
    py: "📜",
    java: "📜",
    cpp: "📜",
    c: "📜",
    cs: "📜",
    exe: "⚙️",
    dll: "⚙️",
    msi: "⚙️",
    iso: "💿",
  };

  return icons[ext] || "📄";
}

// Setup event delegation for item actions
function setupItemEventListeners() {
  const container = document.getElementById("oneDriveContainer");
  if (!container) return;

  // Use event delegation for entire container
  container.addEventListener(
    "click",
    (e) => {
      // Check if clicking on action button
      const btn = e.target.closest("[data-action]");
      if (btn) {
        e.stopPropagation();

        const action = btn.getAttribute("data-action");
        const itemId = btn.getAttribute("data-item-id");
        const itemEl = btn.closest(".onedrive-item");
        const itemName = itemEl ? itemEl.getAttribute("data-item-name") : "";
        const isFolder = btn.getAttribute("data-is-folder") === "true";

        switch (action) {
          case "download":
            downloadDriveItem(itemId, itemName, isFolder);
            break;
          case "details":
            showItemDetails(itemId);
            break;
          case "delete":
            deleteDriveItem(itemId, itemName);
            break;
        }
        return;
      }

      // Check if clicking on a folder row (not on actions)
      const itemEl = e.target.closest(".onedrive-item-folder");
      if (itemEl) {
        const itemId = itemEl.getAttribute("data-item-id");
        const itemName = itemEl.getAttribute("data-item-name");
        loadDriveFolder(itemId, itemName);
      }
    },
    false,
  );
}

// Setup event listeners for search results - uses same delegation as main container
function setupSearchResultListeners() {
  // Event delegation is already set up in setupItemEventListeners
  // which handles all clicks in the container including search results
}

// Setup OneDrive event listeners
function setupOneDriveListeners() {
  setupItemEventListeners();

  // Upload modal
  document
    .getElementById("closeUploadModal")
    ?.addEventListener("click", closeUploadModal);
  document
    .getElementById("cancelUpload")
    ?.addEventListener("click", closeUploadModal);
  document
    .getElementById("confirmUpload")
    ?.addEventListener("click", uploadFileToDrive);
  document.getElementById("uploadFileModal")?.addEventListener("click", (e) => {
    if (e.target.id === "uploadFileModal") closeUploadModal();
  });

  // Create folder modal
  document
    .getElementById("closeCreateFolderModal")
    ?.addEventListener("click", closeCreateFolderModal);
  document
    .getElementById("cancelCreateFolder")
    ?.addEventListener("click", closeCreateFolderModal);
  document
    .getElementById("confirmCreateFolder")
    ?.addEventListener("click", createNewDriveFolder);
  document
    .getElementById("createFolderModal")
    ?.addEventListener("click", (e) => {
      if (e.target.id === "createFolderModal") closeCreateFolderModal();
    });

  // Item details modal
  document
    .getElementById("closeItemDetailsModal")
    ?.addEventListener("click", closeItemDetailsModal);
  document
    .getElementById("itemDetailsModal")
    ?.addEventListener("click", (e) => {
      if (e.target.id === "itemDetailsModal") closeItemDetailsModal();
    });

  // Search
  const searchInput = document.getElementById("oneDriveSearch");
  if (searchInput) {
    searchInput.addEventListener("keypress", (e) => {
      if (e.key === "Enter") {
        searchOneDrive();
      }
    });
  }

  // Toolbar buttons
  document
    .getElementById("uploadFileBtn")
    ?.addEventListener("click", openUploadModal);
  document
    .getElementById("createFolderBtn")
    ?.addEventListener("click", openCreateFolderModal);
  document
    .getElementById("refreshOneDriveBtn")
    ?.addEventListener("click", () => {
      loadDriveFolder(currentDriveId);
    });
  // Enter key for folder name
  document
    .getElementById("newOneDriveFolderName")
    ?.addEventListener("keypress", (e) => {
      if (e.key === "Enter") {
        createNewDriveFolder();
      }
    });
}
