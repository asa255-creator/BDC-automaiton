/**
 * WeatherService.gs - Weather lookup and clothing recommendations.
 */

const WEATHER_UNITS = {
  IMPERIAL: 'imperial',
  METRIC: 'metric'
};

/**
 * Fetches weather summary and clothing recommendation for a date/location.
 *
 * @param {Date} date - Date to retrieve weather for
 * @param {Object|string} locationDetails - Location details object or string
 * @returns {Object|null} Weather summary object or null if unavailable
 */
function getDailyOutlookWeather(date, locationDetails) {
  const props = PropertiesService.getScriptProperties();
  const nwsEnabled = props.getProperty('NWS_WEATHER_ENABLED') === 'true';
  const normalizedLocation = normalizeWeatherLocation(locationDetails);

  if (nwsEnabled && normalizedLocation.latitude !== null && normalizedLocation.longitude !== null) {
    const nwsForecast = fetchNationalWeatherServiceForecast(normalizedLocation, date);
    if (nwsForecast) {
      return nwsForecast;
    }
  }

  const apiKey = props.getProperty('WEATHER_API_KEY');
  if (!apiKey || !normalizedLocation.label) {
    return null;
  }

  const forecast = fetchWeatherForecast(normalizedLocation.label, apiKey);
  if (!forecast || !forecast.list) {
    return null;
  }

  const targetDate = new Date(date);
  targetDate.setHours(0, 0, 0, 0);

  const dailyEntries = forecast.list.filter(entry => {
    const entryDate = new Date(entry.dt * 1000);
    entryDate.setHours(0, 0, 0, 0);
    return entryDate.getTime() === targetDate.getTime();
  });

  if (dailyEntries.length === 0) {
    return null;
  }

  const temps = dailyEntries.map(entry => entry.main.temp);
  const minTemp = Math.min.apply(null, temps);
  const maxTemp = Math.max.apply(null, temps);
  const avgTemp = temps.reduce((sum, value) => sum + value, 0) / temps.length;
  const condition = dailyEntries[0].weather && dailyEntries[0].weather[0]
    ? dailyEntries[0].weather[0].description
    : 'unknown conditions';
  const rainEntry = dailyEntries.find(entry => entry.rain && (entry.rain['3h'] || entry.rain['1h']));
  const rainChance = rainEntry ? 'possible' : 'low';

  return {
    location: forecast.city && forecast.city.name ? forecast.city.name : normalizedLocation.label,
    condition: condition,
    minTemp: Math.round(minTemp),
    maxTemp: Math.round(maxTemp),
    avgTemp: Math.round(avgTemp),
    rainChance: rainChance,
    clothingRecommendation: buildClothingRecommendation(avgTemp, rainChance)
  };
}

/**
 * Fetches weather forecast data for a location.
 *
 * @param {string} location - Location string
 * @param {string} apiKey - API key
 * @returns {Object|null} Forecast response
 */
function fetchWeatherForecast(location, apiKey) {
  try {
    const encodedLocation = encodeURIComponent(location);
    const url = `https://api.openweathermap.org/data/2.5/forecast?q=${encodedLocation}&appid=${apiKey}&units=${WEATHER_UNITS.IMPERIAL}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log(`Weather API error (${responseCode}): ${response.getContentText()}`);
      return null;
    }

    return JSON.parse(response.getContentText());
  } catch (error) {
    Logger.log(`Weather API fetch failed: ${error.message}`);
    return null;
  }
}

/**
 * Normalizes location details from settings or plain text.
 *
 * @param {Object|string} locationDetails - Location details
 * @returns {Object} Normalized details
 */
function normalizeWeatherLocation(locationDetails) {
  if (typeof locationDetails === 'string') {
    return {
      label: locationDetails,
      latitude: null,
      longitude: null,
      city: '',
      state: ''
    };
  }

  if (!locationDetails) {
    return {
      label: '',
      latitude: null,
      longitude: null,
      city: '',
      state: ''
    };
  }

  return {
    label: locationDetails.label || '',
    latitude: locationDetails.latitude !== undefined ? locationDetails.latitude : null,
    longitude: locationDetails.longitude !== undefined ? locationDetails.longitude : null,
    city: locationDetails.city || '',
    state: locationDetails.state || ''
  };
}

/**
 * Fetches forecast from the National Weather Service API.
 *
 * @param {Object} location - Normalized location details
 * @param {Date} date - Target date
 * @returns {Object|null} Weather summary
 */
function fetchNationalWeatherServiceForecast(location, date) {
  try {
    const pointUrl = `https://api.weather.gov/points/${location.latitude},${location.longitude}`;
    const headers = {
      'User-Agent': 'BDC Automation (weather@bdc-automation)',
      'Accept': 'application/geo+json'
    };

    const pointResponse = UrlFetchApp.fetch(pointUrl, {
      headers: headers,
      muteHttpExceptions: true
    });

    if (pointResponse.getResponseCode() !== 200) {
      Logger.log(`NWS points request failed (${pointResponse.getResponseCode()}): ${pointResponse.getContentText()}`);
      return null;
    }

    const pointData = JSON.parse(pointResponse.getContentText());
    const forecastUrl = pointData && pointData.properties ? pointData.properties.forecast : null;

    if (!forecastUrl) {
      Logger.log('NWS forecast URL missing from points response.');
      return null;
    }

    const forecastResponse = UrlFetchApp.fetch(forecastUrl, {
      headers: headers,
      muteHttpExceptions: true
    });

    if (forecastResponse.getResponseCode() !== 200) {
      Logger.log(`NWS forecast request failed (${forecastResponse.getResponseCode()}): ${forecastResponse.getContentText()}`);
      return null;
    }

    const forecastData = JSON.parse(forecastResponse.getContentText());
    const periods = forecastData && forecastData.properties ? forecastData.properties.periods : null;

    if (!periods || periods.length === 0) {
      return null;
    }

    return buildNwsWeatherSummary(periods, date, location, pointData);
  } catch (error) {
    Logger.log(`NWS API fetch failed: ${error.message}`);
    return null;
  }
}

/**
 * Builds a summary from NWS forecast periods.
 *
 * @param {Object[]} periods - Forecast periods
 * @param {Date} date - Target date
 * @param {Object} location - Normalized location details
 * @param {Object} pointData - NWS points response
 * @returns {Object|null} Summary
 */
function buildNwsWeatherSummary(periods, date, location, pointData) {
  const targetDate = new Date(date);
  targetDate.setHours(0, 0, 0, 0);

  const matchingPeriods = periods.filter(period => {
    if (!period.startTime) {
      return false;
    }
    const periodDate = new Date(period.startTime);
    periodDate.setHours(0, 0, 0, 0);
    return periodDate.getTime() === targetDate.getTime();
  });

  if (matchingPeriods.length === 0) {
    return null;
  }

  const temps = matchingPeriods
    .map(period => period.temperature)
    .filter(value => typeof value === 'number');
  const minTemp = temps.length ? Math.min.apply(null, temps) : null;
  const maxTemp = temps.length ? Math.max.apply(null, temps) : null;
  const avgTemp = temps.length ? temps.reduce((sum, value) => sum + value, 0) / temps.length : null;

  const condition = matchingPeriods[0].shortForecast || matchingPeriods[0].detailedForecast || 'unknown conditions';
  const rainProbability = matchingPeriods
    .map(period => period.probabilityOfPrecipitation && period.probabilityOfPrecipitation.value)
    .find(value => typeof value === 'number');
  const rainChance = rainProbability !== undefined && rainProbability !== null && rainProbability >= 30 ? 'possible' : 'low';

  const locationLabel = buildNwsLocationLabel(location, pointData);

  if (avgTemp === null) {
    return null;
  }

  return {
    location: locationLabel,
    condition: condition,
    minTemp: minTemp !== null ? Math.round(minTemp) : Math.round(avgTemp),
    maxTemp: maxTemp !== null ? Math.round(maxTemp) : Math.round(avgTemp),
    avgTemp: Math.round(avgTemp),
    rainChance: rainChance,
    clothingRecommendation: buildClothingRecommendation(avgTemp, rainChance)
  };
}

/**
 * Builds a human-friendly label for NWS locations.
 *
 * @param {Object} location - Normalized location details
 * @param {Object} pointData - NWS points response
 * @returns {string} Location label
 */
function buildNwsLocationLabel(location, pointData) {
  if (location.label) {
    return location.label;
  }

  if (pointData && pointData.properties && pointData.properties.relativeLocation) {
    const relative = pointData.properties.relativeLocation.properties;
    if (relative && relative.city && relative.state) {
      return `${relative.city}, ${relative.state}`;
    }
  }

  return 'Selected location';
}

/**
 * Placeholder for detecting secondary weather locations based on travel.
 *
 * @param {Object[]} events - Calendar events
 * @returns {Object|null} Secondary location details
 */
function getSecondaryWeatherLocationFromCalendar(events) {
  return null;
}

/**
 * Builds a clothing recommendation based on temperature and rain chance.
 *
 * @param {number} avgTemp - Average temperature (F)
 * @param {string} rainChance - Rain chance string
 * @returns {string} Recommendation string
 */
function buildClothingRecommendation(avgTemp, rainChance) {
  let recommendation = '';

  if (avgTemp <= 40) {
    recommendation = 'Bundle up with a warm coat.';
  } else if (avgTemp <= 55) {
    recommendation = 'Bring a jacket or layered top.';
  } else if (avgTemp <= 70) {
    recommendation = 'Light layers should be comfortable.';
  } else if (avgTemp <= 85) {
    recommendation = 'Short sleeves are fine; stay hydrated.';
  } else {
    recommendation = 'Dress for heat and prioritize breathable clothing.';
  }

  if (rainChance === 'possible') {
    recommendation += ' Pack a light rain layer or umbrella.';
  }

  return recommendation;
}
