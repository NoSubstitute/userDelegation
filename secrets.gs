/**
 * @OnlyCurrentDoc
 */

// This file is supposed to keep service account secrets, and is excluded when syncing to/from github

/**
 * Most recent key used here, so it only has access to Gmail.
 * private_key_id: "..."
 * 
 * These credentials are now only in this place. Loaded as global constants.
 */
const PRIVATE_KEY = '-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n';
const CLIENT_EMAIL = '...';
