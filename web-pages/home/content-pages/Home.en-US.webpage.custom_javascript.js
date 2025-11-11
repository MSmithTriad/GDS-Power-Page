/**
 * Validates appointment title input
 */
function validateAppointmentTitle(title) {
    if (!title.trim()) {
        return { isValid: false, error: "Enter an appointment title" };
    }
    if (title.trim().length < 3) {
        return {
            isValid: false,
            error: "Appointment title must be at least 3 characters",
        };
    }
    if (title.trim().length > 100) {
        return {
            isValid: false,
            error: "Appointment title must be 100 characters or less",
        };
    }
    return { isValid: true };
}
/**
 * Validates appointment details input (optional field)
 */
function validateAppointmentDetails(details) {
    // Details are optional, so empty is valid
    if (!details.trim()) {
        return { isValid: true };
    }
    if (details.trim().length > 500) {
        return { isValid: false, error: "Details must be 500 characters or less" };
    }
    return { isValid: true };
}
/**
 * Shows validation error for a specific field
 */
function showFieldError(fieldId, errorMessage) {
    const field = document.getElementById(fieldId);
    if (!field)
        return;
    // Get the form group containing this field
    const formGroup = field.closest(".govuk-form-group");
    if (!formGroup)
        return;
    // Add error class to form group
    formGroup.classList.add("govuk-form-group--error");
    // Create error message element if it doesn't exist
    const errorId = `${fieldId}-error`;
    let errorElement = document.getElementById(errorId);
    if (!errorElement) {
        errorElement = document.createElement("p");
        errorElement.id = errorId;
        errorElement.className = "govuk-error-message";
        errorElement.innerHTML =
            '<span class="govuk-visually-hidden">Error:</span> ' + errorMessage;
        // Insert after the label (or hint if it exists)
        const label = formGroup.querySelector("label");
        if (label) {
            label.insertAdjacentElement("afterend", errorElement);
        }
    }
    else {
        errorElement.innerHTML =
            '<span class="govuk-visually-hidden">Error:</span> ' + errorMessage;
    }
    // Add error class to input field
    field.classList.add("govuk-input--error");
    // Update aria-describedby to include error
    const currentDescribedBy = field.getAttribute("aria-describedby") || "";
    if (!currentDescribedBy.includes(errorId)) {
        field.setAttribute("aria-describedby", (currentDescribedBy + " " + errorId).trim());
    }
}
/**
 * Clears validation error for a specific field
 */
function clearFieldError(fieldId) {
    const field = document.getElementById(fieldId);
    if (!field)
        return;
    // Get the form group containing this field
    const formGroup = field.closest(".govuk-form-group");
    if (formGroup) {
        formGroup.classList.remove("govuk-form-group--error");
    }
    // Remove error message element
    const errorId = `${fieldId}-error`;
    const errorElement = document.getElementById(errorId);
    if (errorElement) {
        errorElement.remove();
    }
    // Remove error class from input field
    field.classList.remove("govuk-input--error");
    // Update aria-describedby to remove error
    const currentDescribedBy = field.getAttribute("aria-describedby") || "";
    const updatedDescribedBy = currentDescribedBy
        .replace(" " + errorId, "")
        .replace(errorId, "")
        .trim();
    if (updatedDescribedBy) {
        field.setAttribute("aria-describedby", updatedDescribedBy);
    }
    else {
        field.removeAttribute("aria-describedby");
    }
}
/**
 * Validates date input and returns formatted date or null if invalid
 */
function validateAndFormatDate(day, month, year) {
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
    if (dateObj.getFullYear() !== yearNum ||
        dateObj.getMonth() !== monthNum - 1 ||
        dateObj.getDate() !== dayNum) {
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
function showDateError(errorMessage) {
    // Get form group element
    const formGroup = document.querySelector(".govuk-form-group:has(#appointment-date)");
    if (!formGroup)
        return;
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
    }
    else {
        errorElement.innerHTML =
            '<span class="govuk-visually-hidden">Error:</span> ' + errorMessage;
    }
    // Add error classes to input fields
    const dayInput = document.getElementById("appointment-date-day");
    const monthInput = document.getElementById("appointment-date-month");
    const yearInput = document.getElementById("appointment-date-year");
    if (dayInput)
        dayInput.classList.add("govuk-input--error");
    if (monthInput)
        monthInput.classList.add("govuk-input--error");
    if (yearInput)
        yearInput.classList.add("govuk-input--error");
    // Update fieldset aria-describedby to include error
    const fieldset = document.querySelector('fieldset[aria-describedby*="appointment-date-hint"]');
    if (fieldset) {
        const currentDescribedBy = fieldset.getAttribute("aria-describedby") || "";
        if (!currentDescribedBy.includes("appointment-date-error")) {
            fieldset.setAttribute("aria-describedby", currentDescribedBy + " appointment-date-error");
        }
    }
}
/**
 * Clears date validation errors from the UI
 */
function clearDateErrors() {
    // Remove error class from form group
    const formGroup = document.querySelector(".govuk-form-group:has(#appointment-date)");
    if (formGroup) {
        formGroup.classList.remove("govuk-form-group--error");
    }
    // Remove error message element
    const errorElement = document.getElementById("appointment-date-error");
    if (errorElement) {
        errorElement.remove();
    }
    // Remove error classes from input fields
    const dayInput = document.getElementById("appointment-date-day");
    const monthInput = document.getElementById("appointment-date-month");
    const yearInput = document.getElementById("appointment-date-year");
    if (dayInput)
        dayInput.classList.remove("govuk-input--error");
    if (monthInput)
        monthInput.classList.remove("govuk-input--error");
    if (yearInput)
        yearInput.classList.remove("govuk-input--error");
    // Update fieldset aria-describedby to remove error
    const fieldset = document.querySelector('fieldset[aria-describedby*="appointment-date-error"]');
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
function addAppointment() {
    document.getElementById("appointment-success-msgbox").style.display = "none";
    // Clear all previous errors
    clearFieldError("appointment-title");
    clearFieldError("appointment-details");
    const subject = document.getElementById("appointment-title").value;
    const description = document.getElementById("appointment-details").value;
    const year = document.getElementById("appointment-date-year").value;
    const month = document.getElementById("appointment-date-month").value;
    const day = document.getElementById("appointment-date-day").value;
    // Validate all fields
    let hasErrors = false;
    // Validate appointment title
    const titleValidation = validateAppointmentTitle(subject);
    if (!titleValidation.isValid) {
        showFieldError("appointment-title", titleValidation.error);
        hasErrors = true;
    }
    // Validate appointment details
    const detailsValidation = validateAppointmentDetails(description);
    if (!detailsValidation.isValid) {
        showFieldError("appointment-details", detailsValidation.error);
        hasErrors = true;
    }
    // Validate the date
    const dateValidation = validateAndFormatDate(day, month, year);
    if (!dateValidation.isValid) {
        showDateError(dateValidation.error);
        hasErrors = true;
    }
    // Stop if there are any validation errors
    if (hasErrors) {
        return;
    }
    const formattedDate = dateValidation.formattedDate;
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
        success: function (res, status, xhr) {
            const appointmentRef = document.getElementById("appointment-ref");
            const successMsgBox = document.getElementById("appointment-success-msgbox");
            if (appointmentRef) {
                appointmentRef.innerHTML = xhr.getResponseHeader("entityid") || "";
            }
            if (successMsgBox) {
                successMsgBox.style.display = "block";
            }
        },
    });
}
