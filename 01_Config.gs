/**
 * 01_Config.gs
 * Centralized constants, cell coordinates, and system timezone helpers.
 */

const CONFIG = {
  SHEETS: {
    DASHBOARD: 'Dashboard',
    CONTENT: 'Content',
    SETTINGS: 'Settings'
  },

  SETTINGS_CELLS: {
    THREADS_ACCOUNT: 'B5',
    THREADS_USER_ID: 'B6',
    CONNECTION_STATUS: 'B7',
    TIMEZONE: 'B8',
    PRIMARY_COLOR: 'B9',
    SECONDARY_COLOR: 'B10'
  },

  CONTENT_COLS: {
    ID: 1,              // A
    CONTENT: 2,         // B
    SCHEDULE_DATE: 3,   // C
    SCHEDULE_TIME: 4,   // D
    STATUS: 5,          // E
    PUBLISHED_AT: 6,    // F
    THREADS_POST_ID: 7, // G
    ERROR_MESSAGE: 8    // H
  },

  CONTENT_START_ROW: 5,
  MAX_POSTS_PER_RUN: 5,
  MAX_CHAR_LIMIT: 500,

  STATUS: {
    DRAFT: 'Draft',
    PENDING: 'Pending',
    PUBLISHING: 'Publishing',
    PUBLISHED: 'Published',
    FAILED: 'Failed',
    CANCELLED: 'Cancelled'
  },

  PROPS: {
    APP_ID: 'THREADS_APP_ID',
    APP_SECRET: 'THREADS_APP_SECRET',
    ACCESS_TOKEN: 'THREADS_ACCESS_TOKEN',
    TOKEN_EXPIRES_AT: 'THREADS_TOKEN_EXPIRES_AT',
    USER_ID: 'THREADS_USER_ID',
    USERNAME: 'THREADS_USERNAME',
    TIMEZONE_LOCKED: 'USER_CUSTOM_TIMEZONE_SET'
  },

  THREADS_API: {
    AUTH_URL: 'https://threads.net/oauth/authorize',
    TOKEN_URL: 'https://graph.threads.net/oauth/access_token',
    EXCHANGE_TOKEN_URL: 'https://graph.threads.net/access_token',
    REFRESH_TOKEN_URL: 'https://graph.threads.net/refresh_access_token',
    BASE_URL: 'https://graph.threads.net/v1.0',
    SCOPES: 'threads_basic,threads_content_publish'
  },

  DEFAULT_THEME: {
    PRIMARY: '#18181B',
    SECONDARY: '#F5F1E8'
  },

  TRIGGER_HANDLER: 'runScheduler'
};

function getSystemTimezone_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'UTC';
}

function getConfiguredTimezone_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  const tz = settingsSheet
    ? String(settingsSheet.getRange(CONFIG.SETTINGS_CELLS.TIMEZONE).getValue() || '').trim()
    : '';
  return tz || getSystemTimezone_();
}

function getSavedBrandColors_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  let primary = CONFIG.DEFAULT_THEME.PRIMARY;
  let secondary = CONFIG.DEFAULT_THEME.SECONDARY;

  if (settingsSheet) {
    const p = String(settingsSheet.getRange(CONFIG.SETTINGS_CELLS.PRIMARY_COLOR).getValue() || '').trim();
    const s = String(settingsSheet.getRange(CONFIG.SETTINGS_CELLS.SECONDARY_COLOR).getValue() || '').trim();
    if (isValidHex_(p)) primary = p.toUpperCase();
    if (isValidHex_(s)) secondary = s.toUpperCase();
  }

  return { primary: primary, secondary: secondary };
}

function isValidHex_(hex) {
  return /^#[0-9A-Fa-f]{6}$/.test(String(hex || '').trim());
}
