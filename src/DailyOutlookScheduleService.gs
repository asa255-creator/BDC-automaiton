/**
 * DailyOutlookScheduleService.gs - Daily outlook schedule helpers.
 */

const DAILY_OUTLOOK_SCHEDULES = {
  DAY_OF: 'day_of',
  NIGHT_BEFORE: 'night_before'
};

/**
 * Returns the configured schedule for daily outlook generation.
 *
 * @returns {string} Schedule value
 */
function getDailyOutlookSchedule() {
  const schedule = PropertiesService.getScriptProperties().getProperty('DAILY_OUTLOOK_SCHEDULE');
  return schedule || DAILY_OUTLOOK_SCHEDULES.DAY_OF;
}

/**
 * Returns the report date based on the configured schedule.
 *
 * @param {Date} baseDate - Date when the trigger runs
 * @returns {Date} Report date
 */
function getDailyOutlookReportDate(baseDate) {
  const schedule = getDailyOutlookSchedule();
  const reportDate = new Date(baseDate);

  if (schedule === DAILY_OUTLOOK_SCHEDULES.NIGHT_BEFORE) {
    reportDate.setDate(reportDate.getDate() + 1);
  }

  return reportDate;
}
