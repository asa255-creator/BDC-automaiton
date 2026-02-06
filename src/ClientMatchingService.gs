/**
 * ClientMatchingService.gs - Client matching helpers.
 */

/**
 * Identifies a client from a comma-separated list of email addresses.
 *
 * @param {string} emailAddresses - Comma-separated email addresses
 * @returns {Object|null} The matched client or null
 */
function identifyClientFromEmailAddresses(emailAddresses) {
  if (!emailAddresses) return null;

  // Parse email addresses (can be "Name <email>" format)
  const emails = emailAddresses.split(',').map(addr => {
    const match = addr.match(/<([^>]+)>/);
    return match ? match[1].trim().toLowerCase() : addr.trim().toLowerCase();
  }).filter(e => e);

  // Try to match against client registry
  const clients = getClientRegistry();

  for (const client of clients) {
    // Check contact emails
    const contactEmails = parseCommaSeparatedList(client.contact_emails)
      .map(e => e.toLowerCase());

    for (const email of emails) {
      if (contactEmails.includes(email)) {
        return client;
      }
    }

    // Check email domains
    const domains = parseCommaSeparatedList(client.email_domains)
      .map(d => d.toLowerCase());

    for (const email of emails) {
      const emailDomain = email.split('@')[1];
      if (emailDomain && domains.includes(emailDomain)) {
        return client;
      }
    }
  }

  return null;
}
