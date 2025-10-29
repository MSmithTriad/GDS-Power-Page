/**
 * Adds an appointment by calling the Web API
 */
function addAppointment() {
  document.getElementById("appointment-success-msgbox").style.display = "none";
  var subject = document.getElementById("appointment-title").value;
  var description = document.getElementById("appointment-details").value;
  var startDate =
    document.getElementById("appointment-date-year").value +
    "-" +
    document.getElementById("appointment-date-month").value +
    "-" +
    document.getElementById("appointment-date-day").value +
    "T00:00:00Z"; //"2025-02-10T13:00:00Z"
  var endDate =
    document.getElementById("appointment-date-year").value +
    "-" +
    document.getElementById("appointment-date-month").value +
    "-" +
    document.getElementById("appointment-date-day").value +
    "T00:30:00Z";

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
      document.getElementById("appointment-ref").innerHTML =
        xhr.getResponseHeader("entityid");
      document.getElementById("appointment-success-msgbox").style.display =
        "block";
      //console.log("entityID: "+ xhr.getResponseHeader("entityid"))
    },
  });
}
