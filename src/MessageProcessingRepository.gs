/**
 * MessageProcessingRepository.gs - Cache-backed message processing flags.
 */

/**
 * Extracts metadata from HTML comment tags in email body.
 *
 * @param {string} body - The email body HTML
 * @param {string} key - The metadata key to extract
 * @returns {string|null} The extracted value or null
 */
function extractMetadata(body, key) {
  const regex = new RegExp(`<!--${key}:(.+?)-->`);
  const match = body.match(regex);
  return match ? match[1] : null;
}

/**
 * Checks if a message has already been processed.
 *
 * @param {string} messageId - The Gmail message ID
 * @returns {boolean} True if already processed
 */
function isMessageProcessed(messageId) {
  const cache = CacheService.getScriptCache();
  return cache.get(`processed_${messageId}`) !== null;
}

/**
 * Marks a message as processed.
 *
 * @param {string} messageId - The Gmail message ID
 */
function markMessageProcessed(messageId) {
  const cache = CacheService.getScriptCache();
  // Cache for 7 days (604800 seconds)
  cache.put(`processed_${messageId}`, 'true', 604800);
}
