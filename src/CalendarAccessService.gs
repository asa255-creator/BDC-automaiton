/**
 * CalendarAccessService.gs - Calendar access helpers for primary + optional secondary calendar.
 */

/**
 * Returns configured calendars in priority order:
 * 1) default calendar
 * 2) optional secondary calendar from settings (if accessible)
 *
 * @returns {Object[]} Array of calendar descriptors
 */
function getConfiguredCalendars() {
  const calendars = [];
  const defaultCalendar = CalendarApp.getDefaultCalendar();

  calendars.push({
    id: defaultCalendar.getId(),
    label: 'primary',
    calendar: defaultCalendar
  });

  const secondaryCalendarId = (PropertiesService.getScriptProperties().getProperty('SECONDARY_CALENDAR_ID') || '').trim();
  if (!secondaryCalendarId) {
    return calendars;
  }

  try {
    const secondaryCalendar = CalendarApp.getCalendarById(secondaryCalendarId);
    if (!secondaryCalendar) {
      Logger.log(`Secondary calendar not found or inaccessible: ${secondaryCalendarId}`);
      return calendars;
    }

    calendars.push({
      id: secondaryCalendarId,
      label: 'secondary',
      calendar: secondaryCalendar
    });
  } catch (error) {
    Logger.log(`Secondary calendar access failed (${secondaryCalendarId}): ${error.message}`);
  }

  return calendars;
}

/**
 * Gets events from all configured calendars in priority order.
 *
 * @param {Date} startTime - Start time
 * @param {Date} endTime - End time
 * @returns {CalendarEvent[]} Combined events
 */
function getEventsFromConfiguredCalendars(startTime, endTime) {
  const events = [];
  const calendars = getConfiguredCalendars();

  for (const calendarInfo of calendars) {
    try {
      const calendarEvents = calendarInfo.calendar.getEvents(startTime, endTime);
      events.push(...calendarEvents);
    } catch (error) {
      Logger.log(`Failed reading ${calendarInfo.label} calendar (${calendarInfo.id}): ${error.message}`);
    }
  }

  return events;
}
