/**
 * DailyOutlookContextService.gs - Daily outlook context helpers.
 */

/**
 * Returns the configured location for daily outlook weather.
 *
 * @returns {string} Location string
 */
function getDailyOutlookLocation() {
  return getUserLocation();
}

/**
 * Builds a context string for daily outlook logging.
 *
 * @param {Date} triggerDate - Date when the trigger fires
 * @returns {string} Log-friendly context string
 */
function getDailyOutlookLogContext(triggerDate) {
  const schedule = getDailyOutlookSchedule();
  const reportDate = getDailyOutlookReportDate(triggerDate);
  const location = getDailyOutlookLocation() || 'unset';

  return `schedule=${schedule}, report_date=${formatDate(reportDate)}, location=${location}`;
}
