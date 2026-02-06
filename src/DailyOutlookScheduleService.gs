/**
 * DailyOutlookScheduleService.gs - Daily outlook schedule helpers.
 */

const DAILY_OUTLOOK_SCHEDULES = {
  DAY_OF: 'day_of',
  DAY_BEFORE: 'day_before',
  NIGHT_BEFORE: 'night_before'
};

const DAILY_OUTLOOK_DEFAULT_TIME = '07:00';
const DAILY_OUTLOOK_LEGACY_NIGHT_TIME = '20:00';

/**
 * Returns the configured schedule for daily outlook generation.
 *
 * @returns {string} Schedule value
 */
function getDailyOutlookSchedule() {
  const props = PropertiesService.getScriptProperties();
  const daySetting = props.getProperty('DAILY_OUTLOOK_DAY');
  if (daySetting) {
    return daySetting;
  }

  const legacy = props.getProperty('DAILY_OUTLOOK_SCHEDULE');
  if (legacy === DAILY_OUTLOOK_SCHEDULES.NIGHT_BEFORE) {
    return DAILY_OUTLOOK_SCHEDULES.DAY_BEFORE;
  }

  return legacy || DAILY_OUTLOOK_SCHEDULES.DAY_OF;
}

/**
 * Returns the configured time for daily outlook generation.
 *
 * @returns {string} Time value (HH:mm)
 */
function getDailyOutlookTime() {
  const props = PropertiesService.getScriptProperties();
  const timeSetting = props.getProperty('DAILY_OUTLOOK_TIME');
  if (timeSetting) {
    return timeSetting;
  }

  const legacy = props.getProperty('DAILY_OUTLOOK_SCHEDULE');
  if (legacy === DAILY_OUTLOOK_SCHEDULES.NIGHT_BEFORE) {
    return DAILY_OUTLOOK_LEGACY_NIGHT_TIME;
  }

  return DAILY_OUTLOOK_DEFAULT_TIME;
}

/**
 * Returns the trigger hour/minute based on settings.
 *
 * @returns {{hour: number, minute: number}} Trigger time parts
 */
function getDailyOutlookTriggerTime() {
  const timeSetting = getDailyOutlookTime();
  const match = /^(\d{1,2}):(\d{2})$/.exec(timeSetting);
  if (!match) {
    return { hour: 7, minute: 0 };
  }

  const hour = parseInt(match[1], 10);
  const minute = parseInt(match[2], 10);

  if (Number.isNaN(hour) || Number.isNaN(minute) || hour < 0 || hour > 23) {
    return { hour: 7, minute: 0 };
  }

  const validMinutes = [0, 15, 30, 45];
  const normalizedMinute = validMinutes.includes(minute)
    ? minute
    : validMinutes.reduce((closest, candidate) => (
      Math.abs(candidate - minute) < Math.abs(closest - minute) ? candidate : closest
    ), validMinutes[0]);

  return { hour, minute: normalizedMinute };
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

  if (schedule === DAILY_OUTLOOK_SCHEDULES.DAY_BEFORE || schedule === DAILY_OUTLOOK_SCHEDULES.NIGHT_BEFORE) {
    reportDate.setDate(reportDate.getDate() + 1);
  }

  return reportDate;
}
