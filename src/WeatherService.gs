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
 * @param {string} location - Location string (city, region)
 * @returns {Object|null} Weather summary object or null if unavailable
 */
function getDailyOutlookWeather(date, location) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('WEATHER_API_KEY');
  if (!apiKey || !location) {
    return null;
  }

  const forecast = fetchWeatherForecast(location, apiKey);
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
    location: forecast.city && forecast.city.name ? forecast.city.name : location,
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
