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
      state: 'AL',
      cities: ['Birmingham', 'Montgomery', 'Mobile', 'Huntsville', 'Tuscaloosa']
    },
    {
      state: 'AK',
      cities: ['Anchorage', 'Juneau', 'Fairbanks', 'Sitka', 'Ketchikan']
    },
    {
      state: 'AZ',
      cities: ['Phoenix', 'Tucson', 'Mesa', 'Chandler', 'Scottsdale']
    },
    {
      state: 'AR',
      cities: ['Little Rock', 'Fort Smith', 'Fayetteville', 'Springdale', 'Jonesboro']
    },
    {
      state: 'CA',
      cities: ['Los Angeles', 'San Diego', 'San Jose', 'San Francisco', 'Sacramento']
    },
    {
      state: 'CO',
      cities: ['Denver', 'Colorado Springs', 'Aurora', 'Fort Collins', 'Boulder']
    },
    {
      state: 'CT',
      cities: ['Bridgeport', 'New Haven', 'Stamford', 'Hartford', 'Waterbury']
    },
    {
      state: 'DE',
      cities: ['Wilmington', 'Dover', 'Newark', 'Middletown', 'Smyrna']
    },
    {
      state: 'FL',
      cities: ['Jacksonville', 'Miami', 'Tampa', 'Orlando', 'Tallahassee']
    },
    {
      state: 'GA',
      cities: ['Atlanta', 'Augusta', 'Columbus', 'Macon', 'Savannah']
    },
    {
      state: 'HI',
      cities: ['Honolulu', 'Hilo', 'Kailua', 'Kaneohe', 'Waipahu']
    },
    {
      state: 'ID',
      cities: ['Boise', 'Meridian', 'Nampa', 'Idaho Falls', 'Pocatello']
    },
    {
      state: 'IL',
      cities: ['Chicago', 'Aurora', 'Naperville', 'Joliet', 'Springfield']
    },
    {
      state: 'IN',
      cities: ['Indianapolis', 'Fort Wayne', 'Evansville', 'South Bend', 'Carmel']
    },
    {
      state: 'IA',
      cities: ['Des Moines', 'Cedar Rapids', 'Davenport', 'Sioux City', 'Iowa City']
    },
    {
      state: 'KS',
      cities: ['Wichita', 'Overland Park', 'Kansas City', 'Topeka', 'Olathe']
    },
    {
      state: 'KY',
      cities: ['Louisville', 'Lexington', 'Bowling Green', 'Owensboro', 'Frankfort']
    },
    {
      state: 'LA',
      cities: ['New Orleans', 'Baton Rouge', 'Shreveport', 'Lafayette', 'Lake Charles']
    },
    {
      state: 'ME',
      cities: ['Portland', 'Lewiston', 'Bangor', 'South Portland', 'Augusta']
    },
    {
      state: 'MD',
      cities: ['Baltimore', 'Frederick', 'Rockville', 'Gaithersburg', 'Annapolis']
    },
    {
      state: 'MA',
      cities: ['Boston', 'Worcester', 'Springfield', 'Cambridge', 'Lowell']
    },
    {
      state: 'MI',
      cities: ['Detroit', 'Grand Rapids', 'Warren', 'Sterling Heights', 'Lansing']
    },
    {
      state: 'MN',
      cities: ['Minneapolis', 'Saint Paul', 'Rochester', 'Duluth', 'Bloomington']
    },
    {
      state: 'MS',
      cities: ['Jackson', 'Gulfport', 'Southaven', 'Hattiesburg', 'Biloxi']
    },
    {
      state: 'MO',
      cities: ['Kansas City', 'St. Louis', 'Springfield', 'Columbia', 'Jefferson City']
    },
    {
      state: 'MT',
      cities: ['Billings', 'Missoula', 'Great Falls', 'Bozeman', 'Helena']
    },
    {
      state: 'NE',
      cities: ['Omaha', 'Lincoln', 'Bellevue', 'Grand Island', 'Kearney']
    },
    {
      state: 'NV',
      cities: ['Las Vegas', 'Henderson', 'Reno', 'North Las Vegas', 'Carson City']
    },
    {
      state: 'NH',
      cities: ['Manchester', 'Nashua', 'Concord', 'Derry', 'Dover']
    },
    {
      state: 'NJ',
      cities: ['Newark', 'Jersey City', 'Paterson', 'Elizabeth', 'Trenton']
    },
    {
      state: 'NM',
      cities: ['Albuquerque', 'Las Cruces', 'Rio Rancho', 'Santa Fe', 'Roswell']
    },
    {
      state: 'NY',
      cities: ['New York', 'Buffalo', 'Rochester', 'Yonkers', 'Albany']
    },
    {
      state: 'NC',
      cities: ['Charlotte', 'Raleigh', 'Greensboro', 'Durham', 'Winston-Salem']
    },
    {
      state: 'ND',
      cities: ['Fargo', 'Bismarck', 'Grand Forks', 'Minot', 'West Fargo']
    },
    {
      state: 'OH',
      cities: ['Columbus', 'Cleveland', 'Cincinnati', 'Toledo', 'Akron']
    },
    {
      state: 'OK',
      cities: ['Oklahoma City', 'Tulsa', 'Norman', 'Broken Arrow', 'Edmond']
    },
    {
      state: 'OR',
      cities: ['Portland', 'Salem', 'Eugene', 'Gresham', 'Hillsboro']
    },
    {
      state: 'PA',
      cities: ['Philadelphia', 'Pittsburgh', 'Allentown', 'Erie', 'Harrisburg']
    },
    {
      state: 'RI',
      cities: ['Providence', 'Warwick', 'Cranston', 'Pawtucket', 'East Providence']
    },
    {
      state: 'SC',
      cities: ['Charleston', 'Columbia', 'North Charleston', 'Mount Pleasant', 'Rock Hill']
    },
    {
      state: 'SD',
      cities: ['Sioux Falls', 'Rapid City', 'Aberdeen', 'Brookings', 'Pierre']
    },
    {
      state: 'TN',
      cities: ['Nashville', 'Memphis', 'Knoxville', 'Chattanooga', 'Clarksville']
    },
    {
      state: 'TX',
      cities: ['Houston', 'San Antonio', 'Dallas', 'Austin', 'Fort Worth']
    },
    {
      state: 'UT',
      cities: ['Salt Lake City', 'West Valley City', 'Provo', 'West Jordan', 'Orem']
    },
    {
      state: 'VT',
      cities: ['Burlington', 'South Burlington', 'Rutland', 'Barre', 'Montpelier']
    },
    {
      state: 'VA',
      cities: ['Virginia Beach', 'Norfolk', 'Chesapeake', 'Richmond', 'Newport News']
    },
    {
      state: 'WA',
      cities: ['Seattle', 'Spokane', 'Tacoma', 'Vancouver', 'Olympia']
    },
    {
      state: 'WV',
      cities: ['Charleston', 'Huntington', 'Morgantown', 'Parkersburg', 'Wheeling']
    },
    {
      state: 'WI',
      cities: ['Milwaukee', 'Madison', 'Green Bay', 'Kenosha', 'Racine']
    },
    {
      state: 'WY',
      cities: ['Cheyenne', 'Casper', 'Laramie', 'Gillette', 'Rock Springs']
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

  const normalizedState = normalizeStateInput(state);
  const normalizedCity = normalizeCityName(city);
  if (!normalizedState || !normalizedCity) {
    return null;
  }

  const options = getWeatherLocationOptions();
  const stateEntry = options.find(entry => entry.state === normalizedState);
  if (!stateEntry) {
    return null;
  }

  const exactMatch = stateEntry.cities.find(entry => normalizeCityName(entry.city || entry) === normalizedCity);
  if (exactMatch) {
    return exactMatch;
  }

  return stateEntry.cities.find(entry => {
    const entryCity = normalizeCityName(entry.city || entry);
    return entryCity.startsWith(normalizedCity) || normalizedCity.startsWith(entryCity);
  }) || null;
}

/**
 * Normalizes state input to a two-letter abbreviation.
 *
 * @param {string} stateInput - State abbreviation or name
 * @returns {string} Two-letter abbreviation or empty string
 */
function normalizeStateInput(stateInput) {
  if (!stateInput) {
    return '';
  }

  const normalized = stateInput.toString().trim().toLowerCase();
  if (normalized.length === 2) {
    return normalized.toUpperCase();
  }

  const stateMap = {
    alabama: 'AL',
    alaska: 'AK',
    arizona: 'AZ',
    arkansas: 'AR',
    california: 'CA',
    colorado: 'CO',
    connecticut: 'CT',
    delaware: 'DE',
    florida: 'FL',
    georgia: 'GA',
    hawaii: 'HI',
    idaho: 'ID',
    illinois: 'IL',
    indiana: 'IN',
    iowa: 'IA',
    kansas: 'KS',
    kentucky: 'KY',
    louisiana: 'LA',
    maine: 'ME',
    maryland: 'MD',
    massachusetts: 'MA',
    michigan: 'MI',
    minnesota: 'MN',
    mississippi: 'MS',
    missouri: 'MO',
    montana: 'MT',
    nebraska: 'NE',
    nevada: 'NV',
    newhampshire: 'NH',
    newjersey: 'NJ',
    newmexico: 'NM',
    newyork: 'NY',
    northcarolina: 'NC',
    northdakota: 'ND',
    ohio: 'OH',
    oklahoma: 'OK',
    oregon: 'OR',
    pennsylvania: 'PA',
    rhodeisland: 'RI',
    southcarolina: 'SC',
    southdakota: 'SD',
    tennessee: 'TN',
    texas: 'TX',
    utah: 'UT',
    vermont: 'VT',
    virginia: 'VA',
    washington: 'WA',
    westvirginia: 'WV',
    wisconsin: 'WI',
    wyoming: 'WY'
  };

  const compact = normalized.replace(/\s+/g, '');
  return stateMap[compact] || '';
}

/**
 * Normalizes city input for comparison.
 *
 * @param {string} cityInput - City name
 * @returns {string} Normalized city
 */
function normalizeCityName(cityInput) {
  if (!cityInput) {
    return '';
  }

  return cityInput
    .toString()
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

/**
 * Geocodes a city/state to latitude/longitude using the US Census geocoder.
 *
 * @param {string} state - State abbreviation or name
 * @param {string} city - City name
 * @returns {Object|null} Geocode result with latitude/longitude
 */
function geocodeCityState(state, city) {
  if (!state || !city) {
    return null;
  }

  try {
    const normalizedState = normalizeStateInput(state) || state;
    const address = encodeURIComponent(`${city}, ${normalizedState}`);
    const url = `https://geocoding.geo.census.gov/geocoder/locations/onelineaddress?address=${address}&benchmark=2020&format=json`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log(`Census geocoder error (${response.getResponseCode()}): ${response.getContentText()}`);
      return null;
    }

    const payload = JSON.parse(response.getContentText());
    const matches = payload && payload.result ? payload.result.addressMatches : null;
    if (!matches || matches.length === 0) {
      return null;
    }

    const match = matches[0];
    const coordinates = match.coordinates || {};
    return {
      city: match.addressComponents ? match.addressComponents.city : city,
      state: match.addressComponents ? match.addressComponents.state : normalizedState,
      latitude: coordinates.y,
      longitude: coordinates.x
    };
  } catch (error) {
    Logger.log(`Census geocoder failed: ${error.message}`);
    return null;
  }
}
