let currentCalendarDate = new Date();
let calendarAppointments = [];
let calendarSearchQuery = "";

async function initializeCalendar() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showCalendarNoSession();
    return;
  }

  currentCalendarDate = new Date();
  await loadCalendarAppointments();
}

function showCalendarNoSession() {
  const container = document.getElementById("calendarContainer");
  if (!container) return;

  container.textContent = "";
  const emptyDiv = document.createElement("div");
  emptyDiv.className = "mailbox-empty";
  emptyDiv.textContent = "No active session";
  container.appendChild(emptyDiv);

  updateCalendarDateLabel();
}

function updateCalendarDateLabel() {
  const label = document.getElementById("calendarDateLabel");
  if (!label) return;

  const options = {
    weekday: "long",
    year: "numeric",
    month: "long",
    day: "numeric",
  };
  label.textContent = currentCalendarDate.toLocaleDateString(
    undefined,
    options,
  );
}

async function loadCalendarAppointments() {
  if (!activeM365Session || !activeM365Session.access_token) {
    showToast("No active session");
    return;
  }

  const container = document.getElementById("calendarContainer");
  if (!container) return;

  container.textContent = "";
  const loadingDiv = document.createElement("div");
  loadingDiv.className = "loading-indicator";
  loadingDiv.textContent = "Loading...";
  container.appendChild(loadingDiv);

  updateCalendarDateLabel();

  try {
    // Load 30 days before and 30 days after for better search functionality
    const startDate = new Date(currentCalendarDate);
    startDate.setDate(startDate.getDate() - 30);
    startDate.setHours(0, 0, 0, 0);

    const endDate = new Date(currentCalendarDate);
    endDate.setDate(endDate.getDate() + 30);
    endDate.setHours(23, 59, 59, 999);

    const startISO = startDate.toISOString();
    const endISO = endDate.toISOString();

    const url = `https://graph.microsoft.com/v1.0/me/calendar/calendarView?startDateTime=${startISO}&endDateTime=${endISO}&$orderby=start/dateTime&$top=500`;

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${activeM365Session.access_token}`,
        "Content-Type": "application/json",
        Prefer: 'outlook.timezone="UTC"',
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

    const data = await response.json();
    calendarAppointments = data.value || [];

    renderCalendarAppointments();
  } catch (error) {
    console.error("Error loading calendar:", error);
    showErrorInContainer(container, error.message, {
      title: "Error loading calendar:",
    });
    showToast(`Error loading calendar: ${error.message}`, "error");
  }
}

function renderCalendarAppointments() {
  const container = document.getElementById("calendarContainer");
  if (!container) return;

  container.textContent = "";

  // Filter appointments based on search query and current date
  let filteredAppointments = calendarAppointments;

  if (calendarSearchQuery) {
    // Search across all loaded appointments
    const query = calendarSearchQuery.toLowerCase();
    filteredAppointments = calendarAppointments.filter(
      (apt) =>
        (apt.subject && apt.subject.toLowerCase().includes(query)) ||
        (apt.location &&
          apt.location.displayName &&
          apt.location.displayName.toLowerCase().includes(query)) ||
        (apt.organizer &&
          apt.organizer.emailAddress &&
          apt.organizer.emailAddress.name &&
          apt.organizer.emailAddress.name.toLowerCase().includes(query)) ||
        (apt.bodyPreview && apt.bodyPreview.toLowerCase().includes(query)),
    );
  } else {
    // Filter to current day only when not searching
    const dayStart = new Date(currentCalendarDate);
    dayStart.setHours(0, 0, 0, 0);

    const dayEnd = new Date(currentCalendarDate);
    dayEnd.setHours(23, 59, 59, 999);

    filteredAppointments = calendarAppointments.filter((apt) => {
      const aptStart = new Date(apt.start.dateTime);
      return aptStart >= dayStart && aptStart <= dayEnd;
    });
  }

  if (filteredAppointments.length === 0) {
    const emptyDiv = document.createElement("div");
    emptyDiv.className = "mailbox-empty";
    emptyDiv.textContent = calendarSearchQuery
      ? "No appointments found matching your search"
      : "No appointments scheduled for this day";
    container.appendChild(emptyDiv);
    return;
  }

  filteredAppointments.forEach((appointment) => {
    const aptElement = createAppointmentElement(appointment);
    container.appendChild(aptElement);
  });
}

function createAppointmentElement(appointment) {
  const aptDiv = document.createElement("div");
  aptDiv.className = "appointment-card";

  // Hover effect

  // Time range
  const startTime = new Date(appointment.start.dateTime);
  const endTime = new Date(appointment.end.dateTime);

  const timeDiv = document.createElement("div");
  timeDiv.className = "appointment-time";
  timeDiv.textContent = `${formatTime(startTime)} - ${formatTime(endTime)}`;

  const subjectDiv = document.createElement("div");
  subjectDiv.className = "appointment-subject";
  subjectDiv.textContent = appointment.subject || "(No subject)";

  // Meta information container
  const metaContainer = document.createElement("div");
  metaContainer.className = "appointment-meta-container";

  // Location
  if (
    appointment.location &&
    appointment.location.displayName &&
    appointment.location.displayName.trim()
  ) {
    const locationDiv = document.createElement("div");
    locationDiv.className = "appointment-meta-item";
    locationDiv.textContent = `📍 ${appointment.location.displayName}`;
    metaContainer.appendChild(locationDiv);
  }

  // Online meeting indicator
  if (appointment.isOnlineMeeting) {
    const meetingDiv = document.createElement("div");
    meetingDiv.className = "appointment-meta-item";
    meetingDiv.textContent = "💻 Online Meeting";
    metaContainer.appendChild(meetingDiv);
  }

  if (appointment.organizer && appointment.organizer.emailAddress) {
    const organizerDiv = document.createElement("div");
    organizerDiv.className = "appointment-meta-item";
    organizerDiv.textContent = `👤 ${appointment.organizer.emailAddress.name || appointment.organizer.emailAddress.address}`;
    metaContainer.appendChild(organizerDiv);
  }

  // Attendees count
  if (appointment.attendees && appointment.attendees.length > 0) {
    const attendeesDiv = document.createElement("div");
    attendeesDiv.className = "appointment-meta-item";
    attendeesDiv.textContent = `👥 ${appointment.attendees.length} attendee${appointment.attendees.length !== 1 ? "s" : ""}`;
    metaContainer.appendChild(attendeesDiv);
  }

  aptDiv.appendChild(timeDiv);
  aptDiv.appendChild(subjectDiv);
  if (metaContainer.children.length > 0) {
    aptDiv.appendChild(metaContainer);
  }

  // Body preview if available
  if (appointment.bodyPreview && appointment.bodyPreview.trim()) {
    const bodyDiv = document.createElement("div");
    bodyDiv.className = "appointment-body-preview";
    bodyDiv.textContent = appointment.bodyPreview.substring(0, 150);
    if (appointment.bodyPreview.length > 150) {
      bodyDiv.textContent += "...";
    }
    aptDiv.appendChild(bodyDiv);
  }

  const actionsDiv = document.createElement("div");
  actionsDiv.className = "appointment-actions";

  // View details button
  const detailsBtn = document.createElement("button");
  detailsBtn.className = "btn btn-small btn-secondary btn-medium";
  detailsBtn.textContent = "📋 View Details";
  detailsBtn.onclick = (e) => {
    e.stopPropagation();
    showAppointmentDetails(appointment);
  };
  actionsDiv.appendChild(detailsBtn);

  // Copy meeting link button
  if (appointment.onlineMeeting && appointment.onlineMeeting.joinUrl) {
    const copyLinkBtn = document.createElement("button");
    copyLinkBtn.className = "btn btn-small btn-primary btn-medium";
    copyLinkBtn.textContent = "🔗 Copy Meeting Link";
    copyLinkBtn.onclick = async (e) => {
      e.stopPropagation();
      try {
        await navigator.clipboard.writeText(appointment.onlineMeeting.joinUrl);
        if (typeof showToast === "function") {
          showToast("Meeting link copied to clipboard!");
        }
      } catch (error) {
        console.error("Failed to copy link:", error);
        if (typeof showToast === "function") {
          showToast("Failed to copy link");
        }
      }
    };
    actionsDiv.appendChild(copyLinkBtn);
  }

  aptDiv.appendChild(actionsDiv);

  return aptDiv;
}

function formatTime(date) {
  const hours = date.getHours();
  const minutes = date.getMinutes();
  const ampm = hours >= 12 ? "PM" : "AM";
  const hours12 = hours % 12 || 12;
  const minutesStr = minutes.toString().padStart(2, "0");
  return `${hours12}:${minutesStr} ${ampm}`;
}

function showAppointmentDetails(appointment) {
  let modal = document.getElementById("appointmentDetailsModal");
  if (!modal) {
    // Create modal if it doesn't exist
    createAppointmentDetailsModal();
    modal = document.getElementById("appointmentDetailsModal");
  }

  const detailsContent = document.getElementById("appointmentDetailsContent");
  if (!detailsContent) return;

  detailsContent.textContent = "";

  // Subject
  const subjectH3 = document.createElement("h3");
  subjectH3.textContent = appointment.subject || "(No subject)";
  subjectH3.className = "margin-top-0 mb-15";
  detailsContent.appendChild(subjectH3);

  // Time
  const startTime = new Date(appointment.start.dateTime);
  const endTime = new Date(appointment.end.dateTime);

  addDetailRow(
    detailsContent,
    "⏰ Time",
    `${formatDateTime(startTime)} - ${formatDateTime(endTime)}`,
  );

  // Location
  if (
    appointment.location &&
    appointment.location.displayName &&
    appointment.location.displayName.trim()
  ) {
    addDetailRow(
      detailsContent,
      "📍 Location",
      appointment.location.displayName,
    );
  }

  if (appointment.isOnlineMeeting && appointment.onlineMeeting) {
    const meetingDiv = document.createElement("div");
    meetingDiv.className = "detail-info-section";

    const label = document.createElement("strong");
    label.textContent = "💻 Online Meeting: ";
    meetingDiv.appendChild(label);

    if (appointment.onlineMeeting.joinUrl) {
      const copyBtn = document.createElement("button");
      copyBtn.className = "btn btn-small btn-primary btn-inline-action";
      copyBtn.textContent = "🔗 Copy Join Link";
      copyBtn.onclick = async () => {
        try {
          await navigator.clipboard.writeText(
            appointment.onlineMeeting.joinUrl,
          );
          if (typeof showToast === "function") {
            showToast("Meeting link copied to clipboard!");
          }
        } catch (error) {
          console.error("Failed to copy link:", error);
          if (typeof showToast === "function") {
            showToast("Failed to copy link");
          }
        }
      };
      meetingDiv.appendChild(copyBtn);
    } else {
      meetingDiv.appendChild(document.createTextNode("Yes"));
    }

    detailsContent.appendChild(meetingDiv);
  }

  // Organizer
  if (appointment.organizer && appointment.organizer.emailAddress) {
    addDetailRow(
      detailsContent,
      "👤 Organizer",
      `${appointment.organizer.emailAddress.name || appointment.organizer.emailAddress.address} <${appointment.organizer.emailAddress.address}>`,
    );
  }

  if (appointment.attendees && appointment.attendees.length > 0) {
    const attendeesDiv = document.createElement("div");
    attendeesDiv.className = "detail-info-section";

    const label = document.createElement("strong");
    label.textContent = `👥 Attendees (${appointment.attendees.length}): `;
    attendeesDiv.appendChild(label);

    const attendeesList = document.createElement("div");
    attendeesList.className = "attendee-list";

    appointment.attendees.forEach((attendee) => {
      const attendeeDiv = document.createElement("div");
      attendeeDiv.className = "attendee-item";

      const name =
        attendee.emailAddress.name || attendee.emailAddress.address || "N/A";
      const email = attendee.emailAddress.address || "";
      const responseStatus = attendee.status
        ? attendee.status.response || "none"
        : "none";

      const statusEmoji = {
        accepted: "✅",
        declined: "❌",
        tentative: "❓",
        none: "⚪",
      };

      attendeeDiv.textContent = `${statusEmoji[responseStatus] || "⚪"} ${name}`;
      if (email && email !== name) {
        attendeeDiv.textContent += ` <${email}>`;
      }

      attendeesList.appendChild(attendeeDiv);
    });

    attendeesDiv.appendChild(attendeesList);
    detailsContent.appendChild(attendeesDiv);
  }

  if (appointment.body && appointment.body.content) {
    const bodyDiv = document.createElement("div");
    bodyDiv.className = "detail-body-section";

    const headerRow = document.createElement("div");
    headerRow.className = "detail-header-row";

    const bodyLabel = document.createElement("strong");
    bodyLabel.textContent = "Description:";
    headerRow.appendChild(bodyLabel);

    // HTML toggle button (only show if content is HTML)
    if (appointment.body.contentType === "html") {
      const toggleBtn = document.createElement("button");
      toggleBtn.className = "btn btn-small btn-secondary btn-xs";
      toggleBtn.textContent = "View as HTML";
      toggleBtn.setAttribute("data-view-mode", "text");
      headerRow.appendChild(toggleBtn);
    }

    bodyDiv.appendChild(headerRow);

    const bodyContent = document.createElement("div");
    bodyContent.id = "appointmentBodyContent";
    bodyContent.className = "detail-body-content-box";

    // Always show as text initially
    if (appointment.body.contentType === "html") {
      // Strip HTML tags for text view
      const tempDiv = document.createElement("div");
      tempDiv.innerHTML = appointment.body.content;
      bodyContent.textContent = tempDiv.textContent || tempDiv.innerText || "";
    } else {
      bodyContent.textContent = appointment.body.content;
    }

    // Store the original content for toggling
    bodyContent.setAttribute("data-html-content", appointment.body.content);
    bodyContent.setAttribute("data-content-type", appointment.body.contentType);

    bodyDiv.appendChild(bodyContent);
    detailsContent.appendChild(bodyDiv);

    // Setup toggle functionality
    if (appointment.body.contentType === "html") {
      const toggleBtn = headerRow.querySelector("button");
      toggleBtn.onclick = () => {
        toggleAppointmentBodyView();
      };
    }
  }

  // Store appointment data for toggling
  if (modal) {
    modal.setAttribute("data-appointment-id", appointment.id);
    modal.classList.add("modal-show");
  }
}

function createAppointmentDetailsModal() {
  const modal = document.createElement("div");
  modal.id = "appointmentDetailsModal";
  modal.className = "modal";

  const modalContent = document.createElement("div");
  modalContent.className = "modal-content";
  modalContent.className = "modal-content max-width-700";

  const modalHeader = document.createElement("div");
  modalHeader.className = "modal-header";

  const title = document.createElement("h2");
  title.textContent = "Appointment Details";
  modalHeader.appendChild(title);

  const closeBtn = document.createElement("button");
  closeBtn.className = "modal-close";
  closeBtn.id = "closeAppointmentDetailsBtn";
  closeBtn.innerHTML = "&times;";
  closeBtn.onclick = () => {
    modal.classList.remove("modal-show");
  };
  modalHeader.appendChild(closeBtn);

  const modalBody = document.createElement("div");
  modalBody.className = "modal-body";

  const detailsContent = document.createElement("div");
  detailsContent.id = "appointmentDetailsContent";
  modalBody.appendChild(detailsContent);

  modalContent.appendChild(modalHeader);
  modalContent.appendChild(modalBody);
  modal.appendChild(modalContent);

  document.body.appendChild(modal);

  // Close on outside click
  modal.addEventListener("click", (e) => {
    if (e.target === modal) {
      modal.classList.remove("modal-show");
    }
  });
}

function addDetailRow(container, label, value) {
  const div = document.createElement("div");
  div.className = "detail-info-section";

  const labelEl = document.createElement("strong");
  labelEl.textContent = label + ": ";
  div.appendChild(labelEl);

  const valueEl = document.createElement("span");
  valueEl.textContent = value;
  div.appendChild(valueEl);

  container.appendChild(div);
}

function formatDateTime(date) {
  const options = {
    weekday: "short",
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  };
  return date.toLocaleString(undefined, options);
}

// Calendar navigation
function goToPreviousDay() {
  currentCalendarDate.setDate(currentCalendarDate.getDate() - 1);
  loadCalendarAppointments();
}

function goToNextDay() {
  currentCalendarDate.setDate(currentCalendarDate.getDate() + 1);
  loadCalendarAppointments();
}

function goToToday() {
  // Open date picker instead of going to today
  const datePicker = document.getElementById("calendarDatePicker");
  if (datePicker) {
    // Set the date picker to current date
    const year = currentCalendarDate.getFullYear();
    const month = String(currentCalendarDate.getMonth() + 1).padStart(2, "0");
    const day = String(currentCalendarDate.getDate()).padStart(2, "0");
    datePicker.value = `${year}-${month}-${day}`;

    // Reset to absolute positioning within parent container
    datePicker.style.position = "absolute";
    datePicker.style.left = "0";
    datePicker.style.top = "100%";
    datePicker.style.opacity = "0";
    datePicker.style.pointerEvents = "auto";
    datePicker.style.width = "1px";
    datePicker.style.height = "1px";

    // Trigger the date picker
    datePicker.focus();
    if (datePicker.showPicker) {
      try {
        datePicker.showPicker();
      } catch (e) {
        // Fallback for browsers that don't support showPicker
        datePicker.click();
      }
    } else {
      datePicker.click();
    }
  }
}

function toggleAppointmentBodyView() {
  const bodyContent = document.getElementById("appointmentBodyContent");
  const modal = document.getElementById("appointmentDetailsModal");
  const toggleBtn = modal.querySelector(".modal-body button[data-view-mode]");

  if (!bodyContent || !toggleBtn) return;

  const currentMode = toggleBtn.getAttribute("data-view-mode");
  const htmlContent = bodyContent.getAttribute("data-html-content");
  const contentType = bodyContent.getAttribute("data-content-type");

  if (currentMode === "text") {
    // Switch to HTML view
    toggleBtn.setAttribute("data-view-mode", "html");
    toggleBtn.textContent = "View as Text";
    bodyContent.innerHTML = htmlContent;
  } else {
    // Switch to text view
    toggleBtn.setAttribute("data-view-mode", "text");
    toggleBtn.textContent = "View as HTML";
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = htmlContent;
    bodyContent.textContent = tempDiv.textContent || tempDiv.innerText || "";
  }
}

function handleDatePickerChange(e) {
  const selectedDate = e.target.value;
  if (selectedDate) {
    // Parse the date (YYYY-MM-DD format)
    const [year, month, day] = selectedDate.split("-").map(Number);
    currentCalendarDate = new Date(year, month - 1, day);
    loadCalendarAppointments();
  }

  // Reset date picker styling after selection
  const datePicker = e.target;
  datePicker.style.position = "absolute";
  datePicker.style.opacity = "0";
  datePicker.style.pointerEvents = "none";
  datePicker.style.left = "";
  datePicker.style.top = "";
}

// Calendar search
function setupCalendarSearch() {
  const searchInput = document.getElementById("calendarSearch");
  if (!searchInput) return;

  let searchTimeout;

  searchInput.addEventListener("input", (e) => {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(() => {
      calendarSearchQuery = e.target.value.trim();
      renderCalendarAppointments();
    }, 300);
  });
}

// Setup calendar event listeners
function setupCalendarListeners() {
  const prevBtn = document.getElementById("calendarPrevDay");
  const nextBtn = document.getElementById("calendarNextDay");
  const todayBtn = document.getElementById("calendarToday");
  const refreshBtn = document.getElementById("refreshCalendar");

  if (prevBtn) {
    prevBtn.addEventListener("click", goToPreviousDay);
  }

  if (nextBtn) {
    nextBtn.addEventListener("click", goToNextDay);
  }

  if (todayBtn) {
    todayBtn.addEventListener("click", goToToday);
  }

  if (refreshBtn) {
    refreshBtn.addEventListener("click", () => {
      loadCalendarAppointments();
      showToast("Calendar refreshed");
    });
  }

  // Setup date picker
  const datePicker = document.getElementById("calendarDatePicker");
  if (datePicker) {
    datePicker.addEventListener("change", handleDatePickerChange);
  }

  setupCalendarSearch();
}
