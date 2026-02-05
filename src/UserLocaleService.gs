/**
 * UserLocaleService.gs - User location and timezone settings.
 */

/**
 * Returns the configured user location.
 * Falls back to legacy daily outlook location if present.
 *
 * @returns {string} Location string
 */
function getUserLocation() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('USER_LOCATION')
    || props.getProperty('DAILY_OUTLOOK_LOCATION')
    || '';
}

/**
 * Returns the configured user timezone.
 *
 * @returns {string} IANA timezone
 */
function getUserTimezone() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('USER_TIMEZONE') || Session.getScriptTimeZone();
}
