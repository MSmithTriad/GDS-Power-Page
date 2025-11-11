// Helper function to validate login session based on API response
(function (webapi: any, $: any) {
  // Adds/sets PowerApps request header and fails request if session is invalid
  function safeAjax(ajaxOptions: any) {
    // Deferred is like JS's Promise
    var deferredAjax = $.Deferred();

    // shell.getTokenDeferred() is a JavaScript method used in PowerApps Portals to generate an authentication token for making Web API calls
    (window as any).shell
      .getTokenDeferred()
      .done(function (token: string) {
        // add headers for ajax
        if (!ajaxOptions.headers) {
          $.extend(ajaxOptions, {
            headers: {
              __RequestVerificationToken: token,
            },
          });
        } else {
          ajaxOptions.headers["__RequestVerificationToken"] = token;
        }
        $.ajax(ajaxOptions)
          .done(function (
            data: any,
            textStatus: string,
            jqXHR: XMLHttpRequest
          ) {
            validateLoginSession(data, textStatus, jqXHR, deferredAjax.resolve);
          })
          .fail(deferredAjax.reject); //ajax
      })
      .fail(function () {
        deferredAjax.rejectWith(this, arguments); // on token failure, pass the token ajax and args
      });

    return deferredAjax.promise();
  }
  webapi.safeAjax = safeAjax;
})(
  ((window as any).webapi = (window as any).webapi || {}),
  (window as any).jQuery
);

// Global variables to track current delete operation
let currentDeleteId: string | null = null;
let currentDeleteTitle: string | null = null;

/**
 * Shows the delete confirmation modal
 */
function showDeleteConfirmation(
  appointmentId: string,
  appointmentTitle: string
): void {
  currentDeleteId = appointmentId;
  currentDeleteTitle = appointmentTitle;

  // Update modal content
  const titleElement = document.getElementById("delete-appointment-title");
  if (titleElement) {
    titleElement.textContent = appointmentTitle;
  }

  // Show modal
  const modal = document.getElementById("delete-confirmation-modal");
  if (modal) {
    modal.style.display = "flex";

    // Focus on the modal for accessibility
    const confirmButton = document.getElementById("confirm-delete-btn");
    if (confirmButton) {
      confirmButton.focus();
    }

    // Prevent body scroll
    document.body.style.overflow = "hidden";

    // Add escape key listener
    document.addEventListener("keydown", handleModalEscape);
  }
}

/**
 * Hides the delete confirmation modal
 */
function hideDeleteConfirmation(): void {
  const modal = document.getElementById("delete-confirmation-modal");
  if (modal) {
    modal.style.display = "none";

    // Restore body scroll
    document.body.style.overflow = "auto";

    // Remove escape key listener
    document.removeEventListener("keydown", handleModalEscape);

    // Reset loading state if any
    hideDeleteLoadingState();
  }

  // Clear current operation
  currentDeleteId = null;
  currentDeleteTitle = null;
}

/**
 * Handles escape key press to close modal
 */
function handleModalEscape(event: KeyboardEvent): void {
  if (event.key === "Escape") {
    hideDeleteConfirmation();
  }
}

/**
 * Shows loading state on the delete confirmation button
 */
function showDeleteLoadingState(): void {
  const confirmButton = document.getElementById(
    "confirm-delete-btn"
  ) as HTMLButtonElement;
  if (confirmButton) {
    confirmButton.disabled = true;
    confirmButton.classList.add("govuk-button--loading");
    confirmButton.setAttribute("aria-describedby", "delete-loading-text");
    // Store original text
    confirmButton.dataset.originalText = confirmButton.textContent || "";
    confirmButton.textContent = "Deleting...";
  }
}

/**
 * Hides loading state on the delete confirmation button
 */
function hideDeleteLoadingState(): void {
  const confirmButton = document.getElementById(
    "confirm-delete-btn"
  ) as HTMLButtonElement;
  if (confirmButton) {
    confirmButton.disabled = false;
    confirmButton.classList.remove("govuk-button--loading");
    confirmButton.removeAttribute("aria-describedby");
    // Restore original text
    confirmButton.textContent =
      confirmButton.dataset.originalText || "Delete appointment";
  }
}

/**
 * Shows success message for delete operation
 */
function showDeleteSuccessMessage(): void {
  const successMsgBox = document.getElementById("delete-success-msgbox");
  if (successMsgBox) {
    successMsgBox.style.display = "block";
    // Focus on success message for screen readers
    successMsgBox.focus();
    // Scroll to top to ensure message is visible
    successMsgBox.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }
}

/**
 * Shows error message for delete operation
 */
function showDeleteErrorMessage(errorMessage?: string): void {
  const errorMsgBox = document.getElementById("delete-error-msgbox");
  const errorMessageElement = document.getElementById("delete-error-message");

  if (errorMessageElement && errorMessage) {
    errorMessageElement.textContent = errorMessage;
  }

  if (errorMsgBox) {
    errorMsgBox.style.display = "block";
    // Focus on error message for screen readers
    errorMsgBox.focus();
    // Scroll to top to ensure message is visible
    errorMsgBox.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }
}

/**
 * Hides all notification messages
 */
function hideAllDeleteMessages(): void {
  const successMsgBox = document.getElementById("delete-success-msgbox");
  const errorMsgBox = document.getElementById("delete-error-msgbox");

  if (successMsgBox) successMsgBox.style.display = "none";
  if (errorMsgBox) errorMsgBox.style.display = "none";
}

/**
 * Removes the appointment row from the table
 */
function removeAppointmentRow(appointmentId: string): void {
  console.log("Attempting to remove appointment row with ID:", appointmentId);

  // Try different selectors in case the ID format varies
  const selectors = [
    `tr[data-appointment-id="${appointmentId}"]`,
    `tr[data-appointment-id="{${appointmentId}}"]`, // In case it has braces
    `tr[data-appointment-id="${appointmentId.toLowerCase()}"]`, // In case of case differences
    `tr[data-appointment-id="${appointmentId.toUpperCase()}"]`,
  ];

  let row: HTMLTableRowElement | null = null;

  // Try each selector until we find the row
  for (const selector of selectors) {
    row = document.querySelector(selector) as HTMLTableRowElement;
    if (row) {
      console.log("Found row using selector:", selector);
      break;
    }
  }

  if (!row) {
    console.error("Could not find row with appointment ID:", appointmentId);
    console.log(
      "Available rows:",
      Array.from(document.querySelectorAll("[data-appointment-id]")).map((el) =>
        el.getAttribute("data-appointment-id")
      )
    );

    // Force a page reload to ensure the data is current
    window.location.reload();
    return;
  }

  console.log("Removing row for appointment:", appointmentId);

  // Add fade out effect
  row.style.transition = "opacity 0.3s ease";
  row.style.opacity = "0";

  // Remove after animation
  setTimeout(() => {
    row!.remove();
    console.log("Row removed successfully");

    // Check if table is now empty
    const tbody = document.querySelector(
      ".govuk-table__body"
    ) as HTMLTableSectionElement;
    if (tbody && tbody.children.length === 0) {
      console.log("Table is now empty, adding empty state message");
      // Add empty state message
      const emptyRow = document.createElement("tr");
      emptyRow.className = "govuk-table__row";
      emptyRow.innerHTML = `
          <td colspan="4" class="govuk-table__cell govuk-table__cell--empty-state">
            <p class="govuk-body">No appointments found.</p>
            <a href="/" class="govuk-link">Register your first appointment</a>
          </td>
        `;
      tbody.appendChild(emptyRow);
    }
  }, 300);
}

/**
 * Confirms and executes the delete operation
 */
function confirmDelete(): void {
  if (!currentDeleteId) {
    console.error("No appointment ID set for deletion");
    return;
  }

  // Hide any existing messages
  hideAllDeleteMessages();

  // Show loading state
  showDeleteLoadingState();

  // Make the delete API call using the webapi helper
  (window as any).webapi.safeAjax({
    type: "DELETE",
    url: `/_api/appointments(${currentDeleteId})`,
    contentType: "application/json",
    success: function (data: any, textStatus: string, xhr: XMLHttpRequest) {
      // Store the ID locally before clearing the global state
      const appointmentIdToRemove = currentDeleteId;

      console.log(
        "Delete API call successful for appointment:",
        appointmentIdToRemove
      );

      // Hide loading state
      hideDeleteLoadingState();

      // Hide modal
      hideDeleteConfirmation();

      // Remove the row from the table using the stored ID
      if (appointmentIdToRemove) {
        removeAppointmentRow(appointmentIdToRemove);
      }

      // Show success message
      showDeleteSuccessMessage();
    },
    error: function (xhr: XMLHttpRequest, status: string, error: string) {
      console.error("Delete API call failed:", {
        status: xhr.status,
        statusText: xhr.statusText,
        responseText: xhr.responseText,
        appointmentId: currentDeleteId,
      });

      // Hide loading state
      hideDeleteLoadingState();

      // Hide modal
      hideDeleteConfirmation();

      // Determine error message
      let errorMessage =
        "Please try again. If the problem persists, contact support.";

      try {
        const responseText = xhr.responseText;
        if (responseText) {
          const errorData = JSON.parse(responseText);
          if (errorData.error && errorData.error.message) {
            errorMessage = errorData.error.message;
          }
        }
      } catch (e) {
        // Use default error message if parsing fails
      }

      // Handle specific HTTP status codes
      if (xhr.status === 401 || xhr.status === 403) {
        errorMessage =
          "You don't have permission to delete this appointment. Please contact support.";
      } else if (xhr.status === 404) {
        errorMessage =
          "This appointment no longer exists. It may have already been deleted.";
        // Remove the row anyway since it doesn't exist
        removeAppointmentRow(currentDeleteId!);
      } else if (xhr.status === 500) {
        errorMessage = "A server error occurred. Please try again later.";
      } else if (xhr.status === 0) {
        errorMessage =
          "Network error. Please check your connection and try again.";
      }

      // Show error message
      showDeleteErrorMessage(errorMessage);
    },
  });
}

/**
 * Initialize delete functionality when DOM is ready
 */
document.addEventListener("DOMContentLoaded", function () {
  // Hide any messages on page load
  hideAllDeleteMessages();

  console.log("Delete functionality initialized for appointments table");
});
