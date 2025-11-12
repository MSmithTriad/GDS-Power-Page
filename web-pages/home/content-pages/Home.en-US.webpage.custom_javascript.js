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
 * Validates a field using embedded component validation
 */
function validateField(fieldId) {
    // Check if it's a text input with embedded validation
    if (window.GDSTextInput && window.GDSTextInput[fieldId]) {
        return window.GDSTextInput[fieldId].validate();
    }
    // Check if it's a date input with embedded validation
    if (window.GDSDateInput && window.GDSDateInput[fieldId]) {
        const validation = window.GDSDateInput[fieldId].validate();
        if (validation.isValid && validation.date) {
            return { isValid: true, date: validation.date };
        }
        return validation;
    }
    // No component validation found - assume field is valid
    console.warn(`No component validation found for field: ${fieldId}`);
    return { isValid: true };
}
/**
 * Shows validation error for a field using embedded component methods
 */
function showFieldError(fieldId, errorMessage) {
    // Use embedded component error display if available
    if (window.GDSTextInput && window.GDSTextInput[fieldId]) {
        window.GDSTextInput[fieldId].showError(errorMessage);
        return;
    }
    if (window.GDSDateInput && window.GDSDateInput[fieldId]) {
        window.GDSDateInput[fieldId].showError(errorMessage);
        return;
    }
    // No component validation found
    console.warn(`No component error display found for field: ${fieldId}`);
}
/**
 * Clears validation error for a field using embedded component methods
 */
function clearFieldError(fieldId) {
    // Use embedded component error clearing if available
    if (window.GDSTextInput && window.GDSTextInput[fieldId]) {
        window.GDSTextInput[fieldId].clearError();
        return;
    }
    if (window.GDSDateInput && window.GDSDateInput[fieldId]) {
        window.GDSDateInput[fieldId].clearError();
        return;
    }
    // No component validation found
    console.warn(`No component error clearing found for field: ${fieldId}`);
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
// Components handle their own initialization, no global initialization needed
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
    // Validate all fields using the modular validation system
    let hasErrors = false;
    // Validate text inputs using embedded component validation
    const textInputsValid = validateFields([
        "appointment-title",
        "appointment-details",
    ]);
    if (!textInputsValid) {
        hasErrors = true;
    }
    // Validate date using embedded component validation
    let formattedDate = "";
    const dateValidation = validateField("appointment-date");
    if (!dateValidation.isValid) {
        if (dateValidation.error) {
            showFieldError("appointment-date", dateValidation.error);
        }
        hasErrors = true;
    }
    else if (dateValidation.date) {
        formattedDate = dateValidation.date;
    }
    else {
        // No date component validation available - this shouldn't happen in normal usage
        console.error("No date validation available for appointment-date field");
        showFieldError("appointment-date", "Date validation is not available");
        hasErrors = true;
    }
    // Stop if there are any validation errors
    if (hasErrors) {
        hideLoadingState();
        return;
    }
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
