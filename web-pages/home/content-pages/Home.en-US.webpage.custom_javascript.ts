/**
 * Validates date input and returns formatted date or null if invalid
 */
function validateAndFormatDate(
  day: string,
  month: string,
  year: string
): { isValid: boolean; formattedDate?: string; error?: string } {
  // Clear any existing errors first
  clearDateErrors();

  // Check if all fields are provided
  if (!day.trim() || !month.trim() || !year.trim()) {
    return { isValid: false, error: "Please enter a complete date" };
  }

  // Parse values
  const dayNum = parseInt(day, 10);
  const monthNum = parseInt(month, 10);
  const yearNum = parseInt(year, 10);

  // Basic range validation
  if (dayNum < 1 || dayNum > 31) {
    return { isValid: false, error: "Day must be between 1 and 31" };
  }

  if (monthNum < 1 || monthNum > 12) {
    return { isValid: false, error: "Month must be between 1 and 12" };
  }

  if (yearNum < 1900 || yearNum > 2100) {
    return { isValid: false, error: "Year must be between 1900 and 2100" };
  }

  // Create date object to validate the combination
  const dateObj = new Date(yearNum, monthNum - 1, dayNum); // month is 0-indexed in Date constructor

  // Check if the date is valid (e.g., not Feb 30th)
  if (
    dateObj.getFullYear() !== yearNum ||
    dateObj.getMonth() !== monthNum - 1 ||
    dateObj.getDate() !== dayNum
  ) {
    return { isValid: false, error: "Please enter a valid date" };
  }

  // Check if date is in the future
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Reset time to start of day

  if (dateObj < today) {
    return { isValid: false, error: "Appointment date must be in the future" };
  }

  // Format with leading zeros
  const formattedDay = dayNum.toString().padStart(2, "0");
  const formattedMonth = monthNum.toString().padStart(2, "0");
  const formattedDate = `${yearNum}-${formattedMonth}-${formattedDay}`;

  return { isValid: true, formattedDate };
}

/**
 * Shows date validation errors in the UI
 */
function showDateError(errorMessage: string): void {
  // Get form group element
  const formGroup = document.querySelector(
    ".govuk-form-group:has(#appointment-date)"
  ) as HTMLElement;
  if (!formGroup) return;

  // Add error class to form group
  formGroup.classList.add("govuk-form-group--error");

  // Create error message element if it doesn't exist
  let errorElement = document.getElementById("appointment-date-error");
  if (!errorElement) {
    errorElement = document.createElement("p");
    errorElement.id = "appointment-date-error";
    errorElement.className = "govuk-error-message";
    errorElement.innerHTML =
      '<span class="govuk-visually-hidden">Error:</span> ' + errorMessage;

    // Insert after the hint element
    const hintElement = document.getElementById("appointment-date-hint");
    if (hintElement) {
      hintElement.insertAdjacentElement("afterend", errorElement);
    }
  } else {
    errorElement.innerHTML =
      '<span class="govuk-visually-hidden">Error:</span> ' + errorMessage;
  }

  // Add error classes to input fields
  const dayInput = document.getElementById(
    "appointment-date-day"
  ) as HTMLInputElement;
  const monthInput = document.getElementById(
    "appointment-date-month"
  ) as HTMLInputElement;
  const yearInput = document.getElementById(
    "appointment-date-year"
  ) as HTMLInputElement;

  if (dayInput) dayInput.classList.add("govuk-input--error");
  if (monthInput) monthInput.classList.add("govuk-input--error");
  if (yearInput) yearInput.classList.add("govuk-input--error");

  // Update fieldset aria-describedby to include error
  const fieldset = document.querySelector(
    'fieldset[aria-describedby*="appointment-date-hint"]'
  ) as HTMLElement;
  if (fieldset) {
    const currentDescribedBy = fieldset.getAttribute("aria-describedby") || "";
    if (!currentDescribedBy.includes("appointment-date-error")) {
      fieldset.setAttribute(
        "aria-describedby",
        currentDescribedBy + " appointment-date-error"
      );
    }
  }
}

/**
 * Clears date validation errors from the UI
 */
function clearDateErrors(): void {
  // Remove error class from form group
  const formGroup = document.querySelector(
    ".govuk-form-group:has(#appointment-date)"
  ) as HTMLElement;
  if (formGroup) {
    formGroup.classList.remove("govuk-form-group--error");
  }

  // Remove error message element
  const errorElement = document.getElementById("appointment-date-error");
  if (errorElement) {
    errorElement.remove();
  }

  // Remove error classes from input fields
  const dayInput = document.getElementById(
    "appointment-date-day"
  ) as HTMLInputElement;
  const monthInput = document.getElementById(
    "appointment-date-month"
  ) as HTMLInputElement;
  const yearInput = document.getElementById(
    "appointment-date-year"
  ) as HTMLInputElement;

  if (dayInput) dayInput.classList.remove("govuk-input--error");
  if (monthInput) monthInput.classList.remove("govuk-input--error");
  if (yearInput) yearInput.classList.remove("govuk-input--error");

  // Update fieldset aria-describedby to remove error
  const fieldset = document.querySelector(
    'fieldset[aria-describedby*="appointment-date-error"]'
  ) as HTMLElement;
  if (fieldset) {
    const currentDescribedBy = fieldset.getAttribute("aria-describedby") || "";
    const updatedDescribedBy = currentDescribedBy
      .replace(" appointment-date-error", "")
      .replace("appointment-date-error", "");
    fieldset.setAttribute("aria-describedby", updatedDescribedBy);
  }
}

/**
 * Adds an appointment by calling the Web API
 */
function addAppointment(): void {
  document.getElementById("appointment-success-msgbox")!.style.display = "none";

  const subject = (
    document.getElementById("appointment-title") as HTMLInputElement
  ).value;
  const description = (
    document.getElementById("appointment-details") as HTMLInputElement
  ).value;

  const year = (
    document.getElementById("appointment-date-year") as HTMLInputElement
  ).value;
  const month = (
    document.getElementById("appointment-date-month") as HTMLInputElement
  ).value;
  const day = (
    document.getElementById("appointment-date-day") as HTMLInputElement
  ).value;

  // Validate the date
  const dateValidation = validateAndFormatDate(day, month, year);

  if (!dateValidation.isValid) {
    showDateError(dateValidation.error!);
    return; // Stop execution if date is invalid
  }

  const formattedDate = dateValidation.formattedDate!;
  const startDate = `${formattedDate}T00:00:00Z`;
  const endDate = `${formattedDate}T00:30:00Z`;

  webapi.safeAjax({
    type: "POST",
    url: "/_api/appointments",
    contentType: "application/json",
    data: JSON.stringify({
      description: description,
      subject: subject,
      scheduledstart: startDate,
      scheduledend: endDate,
    }),
    success: function (res: any, status: string, xhr: XMLHttpRequest) {
      const appointmentRef = document.getElementById("appointment-ref");
      const successMsgBox = document.getElementById(
        "appointment-success-msgbox"
      );

      if (appointmentRef) {
        appointmentRef.innerHTML = xhr.getResponseHeader("entityid") || "";
      }
      if (successMsgBox) {
        successMsgBox.style.display = "block";
      }
    },
  });
}
