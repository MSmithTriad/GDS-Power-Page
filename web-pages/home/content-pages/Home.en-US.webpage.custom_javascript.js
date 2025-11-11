/* Global functions */
// Placeholder function for validateLoginSession - typically provided by Power Pages
function validateLoginSession(data, textStatus, jqXHR, callback) {
    // In a real Power Pages environment, this function validates the session
    // For now, we'll just call the callback to continue
    callback(data, textStatus, jqXHR);
}
/* End Global functions */
// Helper function to validate login session based on API response
(function (webapi, $) {
    // Adds/sets PowerApps request header and fails request if session is invalid
    function safeAjax(ajaxOptions) {
        // Deferred is like JS's Promise
        var deferredAjax = $.Deferred();
        // shell.getTokenDeferred() is a JavaScript method used in PowerApps Portals to generate an authentication token for making Web API calls
        window.shell
            .getTokenDeferred()
            .done(function (token) {
            // add headers for ajax
            if (!ajaxOptions.headers) {
                $.extend(ajaxOptions, {
                    headers: {
                        __RequestVerificationToken: token,
                    },
                });
            }
            else {
                ajaxOptions.headers["__RequestVerificationToken"] = token;
            }
            $.ajax(ajaxOptions)
                .done(function (data, textStatus, jqXHR) {
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
})((window.webapi = window.webapi || {}), window.jQuery);
/**
 * Validates a field based on its data attributes
 */
function validateField(fieldId) {
    var _a, _b, _c;
    const field = document.getElementById(fieldId);
    if (!field)
        return { isValid: true };
    const value = field.value;
    const required = field.dataset.validationRequired === "true";
    const minLength = parseInt(field.dataset.validationMinLength || "0", 10);
    const maxLength = parseInt(field.dataset.validationMaxLength || "255", 10);
    const label = ((_c = (_b = (_a = field
        .closest(".govuk-form-group")) === null || _a === void 0 ? void 0 : _a.querySelector("label")) === null || _b === void 0 ? void 0 : _b.textContent) === null || _c === void 0 ? void 0 : _c.trim().replace("*", "")) || "This field";
    // Required validation
    if (required && !value.trim()) {
        return { isValid: false, error: `Enter ${label.toLowerCase()}` };
    }
    // Skip other validations if field is empty and not required
    if (!value.trim() && !required) {
        return { isValid: true };
    }
    // Min length validation
    if (value.trim().length < minLength && minLength > 0) {
        return {
            isValid: false,
            error: `${label} must be at least ${minLength} character${minLength === 1 ? "" : "s"}`,
        };
    }
    // Max length validation
    if (value.trim().length > maxLength) {
        return {
            isValid: false,
            error: `${label} must be ${maxLength} characters or less`,
        };
    }
    return { isValid: true };
}
/**
 * Shows validation error for a field
 */
function showFieldError(fieldId, errorMessage) {
    const field = document.getElementById(fieldId);
    if (!field)
        return;
    const formGroup = field.closest(".govuk-form-group");
    if (!formGroup)
        return;
    // Add error class to form group
    formGroup.classList.add("govuk-form-group--error");
    // Create or update error message
    const errorId = `${fieldId}-error`;
    let errorElement = document.getElementById(errorId);
    if (!errorElement) {
        errorElement = document.createElement("p");
        errorElement.id = errorId;
        errorElement.className = "govuk-error-message";
        errorElement.innerHTML =
            '<span class="govuk-visually-hidden">Error:</span> ' + errorMessage;
        // Insert after label or hint
        const label = formGroup.querySelector("label");
        const hint = formGroup.querySelector(".govuk-hint");
        const insertAfter = hint || label;
        if (insertAfter) {
            insertAfter.insertAdjacentElement("afterend", errorElement);
        }
    }
    else {
        errorElement.innerHTML =
            '<span class="govuk-visually-hidden">Error:</span> ' + errorMessage;
    }
    // Add error class to input
    field.classList.add("govuk-input--error");
    // Update aria-describedby
    const currentDescribedBy = field.getAttribute("aria-describedby") || "";
    if (!currentDescribedBy.includes(errorId)) {
        field.setAttribute("aria-describedby", (currentDescribedBy + " " + errorId).trim());
    }
}
/**
 * Clears validation error for a field
 */
function clearFieldError(fieldId) {
    const field = document.getElementById(fieldId);
    if (!field)
        return;
    const formGroup = field.closest(".govuk-form-group");
    if (formGroup) {
        formGroup.classList.remove("govuk-form-group--error");
    }
    // Remove error message
    const errorId = `${fieldId}-error`;
    const errorElement = document.getElementById(errorId);
    if (errorElement) {
        errorElement.remove();
    }
    // Remove error class from input
    field.classList.remove("govuk-input--error");
    // Update aria-describedby
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
 * Validates multiple fields at once
 */
function validateFields(fieldIds) {
    let hasErrors = false;
    fieldIds.forEach((fieldId) => {
        clearFieldError(fieldId);
        const validation = validateField(fieldId);
        if (!validation.isValid) {
            showFieldError(fieldId, validation.error);
            hasErrors = true;
        }
    });
    return !hasErrors;
}
/**
 * Auto-initialize validation for all fields with data-validation attributes
 */
function initializeAutoValidation() {
    const fieldsWithValidation = document.querySelectorAll("[data-validation-required], [data-validation-min-length], [data-validation-max-length]");
    fieldsWithValidation.forEach((field) => {
        const fieldElement = field;
        // Add blur event for real-time validation
        fieldElement.addEventListener("blur", () => {
            clearFieldError(fieldElement.id);
            const validation = validateField(fieldElement.id);
            if (!validation.isValid) {
                showFieldError(fieldElement.id, validation.error);
            }
        });
        // Clear errors on input
        fieldElement.addEventListener("input", () => {
            if (fieldElement.classList.contains("govuk-input--error")) {
                clearFieldError(fieldElement.id);
            }
        });
    });
}
// Auto-initialize when DOM is ready
document.addEventListener("DOMContentLoaded", () => {
    initializeAutoValidation();
});
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
 * Shows the loading state for the form
 */
function showLoadingState() {
    const submitButton = document.querySelector('button[type="submit"]');
    if (submitButton) {
        submitButton.disabled = true;
        submitButton.setAttribute("aria-describedby", "button-loading-text");
        // Store original text
        submitButton.dataset.originalText = submitButton.textContent || "";
        submitButton.innerHTML = `
      <span aria-hidden="true" role="status">
        <span class="govuk-visually-hidden">Loading</span>
        Saving...
      </span>
    `;
    }
}
/**
 * Hides the loading state for the form
 */
function hideLoadingState() {
    const submitButton = document.querySelector('button[type="submit"]');
    if (submitButton) {
        submitButton.disabled = false;
        submitButton.removeAttribute("aria-describedby");
        // Restore original text
        submitButton.textContent = submitButton.dataset.originalText || "Save";
    }
}
/**
 * Shows the success message
 */
function showSuccessMessage(appointmentRef) {
    const appointmentRefElement = document.getElementById("appointment-ref");
    const successMsgBox = document.getElementById("appointment-success-msgbox");
    if (appointmentRefElement) {
        appointmentRefElement.textContent = appointmentRef;
    }
    if (successMsgBox) {
        successMsgBox.style.display = "block";
        // Focus on success message for screen readers
        successMsgBox.focus();
        // Scroll to top to ensure message is visible
        successMsgBox.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }
}
/**
 * Shows the error message
 */
function showErrorMessage(errorMessage) {
    const errorMsgBox = document.getElementById("appointment-error-msgbox");
    const errorMessageElement = document.getElementById("appointment-error-message");
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
function hideAllMessages() {
    const successMsgBox = document.getElementById("appointment-success-msgbox");
    const errorMsgBox = document.getElementById("appointment-error-msgbox");
    if (successMsgBox)
        successMsgBox.style.display = "none";
    if (errorMsgBox)
        errorMsgBox.style.display = "none";
}
/**
 * Adds an appointment by calling the Web API
 */
function addAppointment() {
    // Hide any existing messages
    hideAllMessages();
    // Show loading state
    showLoadingState();
    const subject = document.getElementById("appointment-title").value;
    const description = document.getElementById("appointment-details").value;
    const year = document.getElementById("appointment-date-year").value;
    const month = document.getElementById("appointment-date-month").value;
    const day = document.getElementById("appointment-date-day").value;
    // Validate all fields using the functional validation system
    let hasErrors = false;
    // Validate text inputs using functional validation
    const textInputsValid = validateFields([
        "appointment-title",
        "appointment-details",
    ]);
    if (!textInputsValid) {
        hasErrors = true;
    }
    // Validate the date using existing date validation
    const dateValidation = validateAndFormatDate(day, month, year);
    if (!dateValidation.isValid) {
        showDateError(dateValidation.error);
        hasErrors = true;
    }
    // Stop if there are any validation errors
    if (hasErrors) {
        hideLoadingState();
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
            hideLoadingState();
            const appointmentRef = xhr.getResponseHeader("entityid") || "Unknown";
            showSuccessMessage(appointmentRef);
        },
        error: function (xhr, status, error) {
            hideLoadingState();
            // Try to get a meaningful error message
            let errorMessage = "Please try again. If the problem persists, contact support.";
            try {
                const responseText = xhr.responseText;
                if (responseText) {
                    const errorData = JSON.parse(responseText);
                    if (errorData.error && errorData.error.message) {
                        errorMessage = errorData.error.message;
                    }
                }
            }
            catch (e) {
                // Use default error message if parsing fails
            }
            // Handle specific HTTP status codes
            if (xhr.status === 401 || xhr.status === 403) {
                errorMessage =
                    "You don't have permission to create appointments. Please contact support.";
            }
            else if (xhr.status === 500) {
                errorMessage = "A server error occurred. Please try again later.";
            }
            else if (xhr.status === 0) {
                errorMessage =
                    "Network error. Please check your connection and try again.";
            }
            showErrorMessage(errorMessage);
        },
    });
}
