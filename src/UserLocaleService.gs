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
  const location = getPrimaryWeatherLocationDetails();
  return location.label
    || PropertiesService.getScriptProperties().getProperty('USER_LOCATION')
    || PropertiesService.getScriptProperties().getProperty('DAILY_OUTLOOK_LOCATION')
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

/**
 * Returns structured details for the primary weather location.
 *
 * @returns {Object} Location details
 */
function getPrimaryWeatherLocationDetails() {
  const props = PropertiesService.getScriptProperties();
  const state = props.getProperty('NWS_WEATHER_STATE') || '';
  const city = props.getProperty('NWS_WEATHER_CITY') || '';
  const latValue = props.getProperty('NWS_WEATHER_LAT');
  const lonValue = props.getProperty('NWS_WEATHER_LON');
  const latitude = latValue ? Number(latValue) : null;
  const longitude = lonValue ? Number(lonValue) : null;
  const label = state && city ? `${city}, ${state}` : (props.getProperty('USER_LOCATION') || '');

  return {
    state: state,
    city: city,
    latitude: latitude,
    longitude: longitude,
    label: label
  };
}

/**
 * Returns whether the National Weather Service integration is enabled.
 *
 * @returns {boolean} True if enabled
 */
function isNationalWeatherServiceEnabled() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('NWS_WEATHER_ENABLED') === 'true';
}

/**
 * Returns supported weather location options (state/city + coordinates).
 *
 * @returns {Object[]} Array of state objects with city data
 */
function getWeatherLocationOptions() {
  return [
    {
      state: 'CA',
      cities: [
        { city: 'Los Angeles', latitude: 34.0522, longitude: -118.2437 },
        { city: 'San Diego', latitude: 32.7157, longitude: -117.1611 },
        { city: 'San Francisco', latitude: 37.7749, longitude: -122.4194 }
      ]
    },
    {
      state: 'CO',
      cities: [
        { city: 'Boulder', latitude: 40.01499, longitude: -105.2705 },
        { city: 'Colorado Springs', latitude: 38.8339, longitude: -104.8214 },
        { city: 'Denver', latitude: 39.7392, longitude: -104.9903 }
      ]
    },
    {
      state: 'FL',
      cities: [
        { city: 'Miami', latitude: 25.7617, longitude: -80.1918 },
        { city: 'Orlando', latitude: 28.5383, longitude: -81.3792 },
        { city: 'Tampa', latitude: 27.9506, longitude: -82.4572 }
      ]
    },
    {
      state: 'IL',
      cities: [
        { city: 'Chicago', latitude: 41.8781, longitude: -87.6298 },
        { city: 'Naperville', latitude: 41.7508, longitude: -88.1535 },
        { city: 'Peoria', latitude: 40.6936, longitude: -89.5889 }
      ]
    },
    {
      state: 'MA',
      cities: [
        { city: 'Boston', latitude: 42.3601, longitude: -71.0589 },
        { city: 'Cambridge', latitude: 42.3736, longitude: -71.1097 },
        { city: 'Worcester', latitude: 42.2626, longitude: -71.8023 }
      ]
    },
    {
      state: 'NY',
      cities: [
        { city: 'Albany', latitude: 42.6526, longitude: -73.7562 },
        { city: 'Buffalo', latitude: 42.8864, longitude: -78.8784 },
        { city: 'New York', latitude: 40.7128, longitude: -74.006 }
      ]
    },
    {
      state: 'TX',
      cities: [
        { city: 'Austin', latitude: 30.2672, longitude: -97.7431 },
        { city: 'Dallas', latitude: 32.7767, longitude: -96.797 },
        { city: 'Houston', latitude: 29.7604, longitude: -95.3698 }
      ]
    },
    {
      state: 'WA',
      cities: [
        { city: 'Bellevue', latitude: 47.6101, longitude: -122.2015 },
        { city: 'Seattle', latitude: 47.6062, longitude: -122.3321 },
        { city: 'Spokane', latitude: 47.6588, longitude: -117.426 }
      ]
    }
  ];
}

/**
 * Finds a weather location entry by state/city.
 *
 * @param {string} state - State abbreviation
 * @param {string} city - City name
 * @returns {Object|null} City entry or null
 */
function findWeatherLocationEntry(state, city) {
  if (!state || !city) {
    return null;
  }

  const options = getWeatherLocationOptions();
  const stateEntry = options.find(entry => entry.state === state);
  if (!stateEntry) {
    return null;
  }

  return stateEntry.cities.find(entry => entry.city.toLowerCase() === city.toLowerCase()) || null;
}
