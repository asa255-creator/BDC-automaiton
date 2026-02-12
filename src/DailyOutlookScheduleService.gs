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
 * Parses a time setting from 24-hour (HH:mm) or 12-hour (h:mm AM/PM) input.
 *
 * @param {string} rawTime - Raw time setting
 * @returns {{hour: number, minute: number}|null} Parsed time parts
 */
function parseDailyOutlookTime(rawTime) {
  if (!rawTime) {
    return null;
  }

  const value = String(rawTime).trim();
  const militaryMatch = /^(\d{1,2}):(\d{2})$/.exec(value);
  if (militaryMatch) {
    const hour = parseInt(militaryMatch[1], 10);
    const minute = parseInt(militaryMatch[2], 10);
    if (!Number.isNaN(hour) && !Number.isNaN(minute) && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
      return { hour, minute };
    }
  }

  const meridiemMatch = /^(\d{1,2}):(\d{2})\s*([AaPp][Mm])$/.exec(value);
  if (meridiemMatch) {
    let hour = parseInt(meridiemMatch[1], 10);
    const minute = parseInt(meridiemMatch[2], 10);
    const meridiem = meridiemMatch[3].toUpperCase();

    if (!Number.isNaN(hour) && !Number.isNaN(minute) && hour >= 1 && hour <= 12 && minute >= 0 && minute <= 59) {
      if (meridiem === 'AM' && hour === 12) {
        hour = 0;
      } else if (meridiem === 'PM' && hour !== 12) {
        hour += 12;
      }
      return { hour, minute };
    }
  }

  return null;
}

/**
 * Converts hour/minute to normalized HH:mm format.
 *
 * @param {number} hour - 24-hour value
 * @param {number} minute - Minute value
 * @returns {string} Normalized time
 */
function formatDailyOutlookTime(hour, minute) {
  return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
}

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
  const parsedTime = parseDailyOutlookTime(timeSetting);
  if (parsedTime) {
    return formatDailyOutlookTime(parsedTime.hour, parsedTime.minute);
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
  const parsedTime = parseDailyOutlookTime(getDailyOutlookTime());
  if (!parsedTime) {
    return { hour: 7, minute: 0 };
  }

  return parsedTime;
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
