/****************************************************
 * Gmail → ZIP (GASunzip) → Email TXT  [V8 runtime]
 *
 * Requires the GASunzip library (ID below) and the following
 * Script Properties (File → Project properties → Script properties):
 *
 *   ALLOWED_SENDER             (optional)  e.g. sender@example.com
 *   GMAIL_QUERY                (optional)  default below
 *   ZIP_PASSWORD               (optional)  leave blank if no password
 *   PROCESSED_LABEL            (optional)  default: "Unzip/processed"
 *   SEARCH_TIMEZONE            (optional)  default: "Asia/Tokyo"
 *   TARGET_DATE_OVERRIDE       (optional)  e.g. "2025-10-24" to reprocess that day
 *
 * Salesforce auth (password grant; password must include security token):
 *   SALESFORCE_LOGIN_URL       (optional)  default: https://login.salesforce.com
 *   SALESFORCE_INSTANCE_URL    (optional)  e.g. https://your-domain.my.salesforce.com
 *   SALESFORCE_CLIENT_ID       (required)
 *   SALESFORCE_CLIENT_SECRET   (required)
 *   SALESFORCE_USERNAME        (required)
 *   SALESFORCE_PASSWORD        (required)  (password + security token)
 *   SALESFORCE_API_VERSION     (optional)  default: "65.0"
 *   SALESFORCE_SOBJECT         (optional)  default: "Mail2X__c"
 *   SALESFORCE_API_PATH        (optional)  override path; if set, used verbatim instead of sObject insert path
 ****************************************************/

const PROPS = PropertiesService.getScriptProperties();

const DEFAULTS = {
  gmailQuery: 'in:anywhere has:attachment filename:zip newer_than:2d',
  processedLabel: 'Unzip/processed',
  searchTimezone: 'Asia/Tokyo',
  salesforceApiVersion: '65.0',
  salesforceLoginUrl: 'https://login.salesforce.com',
  salesforceSobject: 'Mail2X__c',
  salesforceApiPath: ''
};

const STATE_KEYS = {
  lastTs: 'LAST_PROCESSED_INTERNAL_TS',
  lastId: 'LAST_PROCESSED_MESSAGE_ID',
  processedLabelCache: 'PROCESSED_LABEL_ID_CACHE'
};

const ONE_DAY_MS = 24 * 60 * 60 * 1000;

/** ────────────────────────────────────────────────
 *  High-level entrypoints
 * ────────────────────────────────────────────────*/

function unzipNewestZipFromGmail_V8(targetDate) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    console.warn('⚠️ Another execution in progress; skipping this run.');
    return;
  }

  try {
  const config = loadRuntimeConfig_();
  const windowInfo = resolveProcessingWindow_(targetDate, config);
  const searchQuery = buildSearchQuery_(config.gmailQuery, config.processedLabel, windowInfo);
  const lastState = getLastProcessedState_();

    console.info(`🛠 Using Script Property date override: ${windowInfo.sourceProperty || 'none'}`);
    console.info(`📅 Processing date: ${windowInfo.label}${windowInfo.overrideApplied ? ' (override)' : ''}`);
    console.info(`🔎 Searching: ${searchQuery}`);

    const message = findNewestZipMessage_(searchQuery, lastState, config.processedLabelId);
    if (!message) {
      console.info('⏹️ No new ZIP attachments matched the query.');
      return;
    }

    const subjectHeader = getHeaderValue_(message.payload.headers, 'Subject') || '';
    const fromHeaderRaw = getHeaderValue_(message.payload.headers, 'From') || '';
    const fromLower = fromHeaderRaw.toLowerCase();
    console.info(`✉️ Incoming message → from=${fromHeaderRaw} subject=${subjectHeader}`);

    if (config.allowedSender && fromLower.indexOf(config.allowedSender) === -1) {
      console.warn(`Skipped: sender "${fromHeaderRaw}" does not contain required "${config.allowedSender}".`);
      console.info('⏹️ Sender not allowed; aborting run without processing.');
      return;
    }
    console.info(`✔ Sender OK: ${fromHeaderRaw}`);

    const zipAttachments = listZipAttachments_(message.payload);
    if (!zipAttachments.length) {
      console.warn('Message had no ZIP attachments after filtering; marking processed and exiting.');
      markThreadProcessed_(message.threadId, config.processedLabelId);
      updateLastProcessedState_(message.internalDate, message.id);
      return;
    }

    const primaryZipMeta = zipAttachments[0];
    const zipBlob = fetchAttachmentBlob_(message.id, primaryZipMeta);
    const zipName = primaryZipMeta.filename || 'attachment.zip';
    console.info(`📦 ZIP selected → name=${zipName} type=${zipBlob.getContentType()} size=${zipBlob.getBytes().length}`);

    const zipEntries = inspectZip_(zipBlob);
    if (!zipEntries.length) {
      console.warn('ZIP appears to be empty; labeling thread and skipping.');
      markThreadProcessed_(message.threadId, config.processedLabelId);
      updateLastProcessedState_(message.internalDate, message.id);
      return;
    }

    const deflate64 = zipEntries.some(e => e.method === 9);
    if (deflate64) {
      throw new Error('Unsupported ZIP: contains Deflate64 (method 9). Recreate with standard Deflate (method 8).');
    }

    const interestingEntryExists = zipEntries.some(e => /\.txt$/i.test(e.name || ''));
    if (!interestingEntryExists) {
      console.warn(`No CSV/TXT entries inside ${zipName}; labeling and skipping.`);
      markThreadProcessed_(message.threadId, config.processedLabelId);
      updateLastProcessedState_(message.internalDate, message.id);
      return;
    }

    const unzippedBlobs = unzipWithGASunzip_(zipBlob, config.zipPassword);
    const { bodies: txtBodies, names: txtNames } = extractTxtBodies_(unzippedBlobs, 'Shift_JIS');
    const extractedCount = unzippedBlobs.length;

    if (!txtBodies.length) {
      console.warn('No TXT files decoded; labeling thread and skipping Salesforce POST.');
      markThreadProcessed_(message.threadId, config.processedLabelId);
      updateLastProcessedState_(message.internalDate, message.id);
      return;
    }

    console.info(`📝 TXT(s) extracted: ${txtNames.join(', ')}`);

    const sfResult = postTxtBodiesToSalesforce_(txtBodies, txtNames, config);
    markThreadProcessed_(message.threadId, config.processedLabelId);
    updateLastProcessedState_(message.internalDate, message.id);
    console.info(`✅ Done. Extracted=${extractedCount}, TXTs=${txtBodies.length} → Salesforce POST success=${sfResult.successCount}, errors=${sfResult.errorCount}`);
  } finally {
    lock.releaseLock();
  }
}

function cron_unzip_every_15min() {
  unzipNewestZipFromGmail_V8();
}

/** ────────────────────────────────────────────────
 *  Configuration & state utilities
 * ────────────────────────────────────────────────*/

function loadRuntimeConfig_() {
  const allowedSender = (PROPS.getProperty('ALLOWED_SENDER') || '').trim().toLowerCase();
  const gmailQuery = (PROPS.getProperty('GMAIL_QUERY') || DEFAULTS.gmailQuery).trim();
  const zipPassword = (PROPS.getProperty('ZIP_PASSWORD') || '').trim();
  const processedLabel = (PROPS.getProperty('PROCESSED_LABEL') || DEFAULTS.processedLabel).trim();
  const searchTimezone = (PROPS.getProperty('SEARCH_TIMEZONE') || DEFAULTS.searchTimezone).trim() || DEFAULTS.searchTimezone;
  const salesforceLoginUrl = (PROPS.getProperty('SALESFORCE_LOGIN_URL') || DEFAULTS.salesforceLoginUrl).trim() || DEFAULTS.salesforceLoginUrl;
  const salesforceInstanceUrl = (PROPS.getProperty('SALESFORCE_INSTANCE_URL') || '').trim();
  const salesforceClientId = (PROPS.getProperty('SALESFORCE_CLIENT_ID') || '').trim();
  const salesforceClientSecret = (PROPS.getProperty('SALESFORCE_CLIENT_SECRET') || '').trim();
  const salesforceUsername = (PROPS.getProperty('SALESFORCE_USERNAME') || '').trim();
  const salesforcePassword = (PROPS.getProperty('SALESFORCE_PASSWORD') || '').trim(); // include security token
  const salesforceApiPath = (PROPS.getProperty('SALESFORCE_API_PATH') || DEFAULTS.salesforceApiPath).trim() || DEFAULTS.salesforceApiPath;
  const salesforceApiVersion = (PROPS.getProperty('SALESFORCE_API_VERSION') || DEFAULTS.salesforceApiVersion).trim() || DEFAULTS.salesforceApiVersion;
  const salesforceSobject = (PROPS.getProperty('SALESFORCE_SOBJECT') || DEFAULTS.salesforceSobject).trim() || DEFAULTS.salesforceSobject;

  const processedLabelId = processedLabel ? ensureLabelId_(processedLabel) : null;

  return {
    allowedSender,
    gmailQuery,
    zipPassword,
    processedLabel,
    processedLabelId,
    searchTimezone,
    salesforceLoginUrl,
    salesforceInstanceUrl,
    salesforceClientId,
    salesforceClientSecret,
    salesforceUsername,
    salesforcePassword,
    salesforceApiPath,
    salesforceApiVersion,
    salesforceSobject
  };
}

function getLastProcessedState_() {
  return {
    ts: Number(PROPS.getProperty(STATE_KEYS.lastTs) || 0),
    id: PROPS.getProperty(STATE_KEYS.lastId) || ''
  };
}

function updateLastProcessedState_(internalDate, messageId) {
  if (internalDate) PROPS.setProperty(STATE_KEYS.lastTs, String(internalDate));
  if (messageId) PROPS.setProperty(STATE_KEYS.lastId, messageId);
}

function ensureLabelId_(labelName) {
  if (!labelName) return null;

  const cached = PROPS.getProperty(STATE_KEYS.processedLabelCache);
  if (cached) {
    try {
      const lbl = Gmail.Users.Labels.get('me', cached);
      if (lbl && lbl.name === labelName) return cached;
    } catch (err) {
      console.warn(`Processed label cache invalid (${cached}): ${err && err.message ? err.message : err}`);
    }
  }

  const labels = Gmail.Users.Labels.list('me').labels || [];
  const existing = labels.find(l => l.name === labelName);
  if (existing) {
    PROPS.setProperty(STATE_KEYS.processedLabelCache, existing.id);
    return existing.id;
  }

  const created = Gmail.Users.Labels.create({ name: labelName }, 'me');
  PROPS.setProperty(STATE_KEYS.processedLabelCache, created.id);
  return created.id;
}

function markThreadProcessed_(threadId, processedLabelId) {
  if (!threadId || !processedLabelId) return;
  try {
    Gmail.Users.Threads.modify({ addLabelIds: [processedLabelId] }, 'me', threadId);
  } catch (err) {
    console.error(`Failed to label thread ${threadId}: ${err && err.message ? err.message : err}`);
  }
}

/** ────────────────────────────────────────────────
 *  Date window & query helpers
 * ────────────────────────────────────────────────*/

function resolveProcessingWindow_(targetDate, config) {
  const tz = config.searchTimezone;
  const overrideProp = (PROPS.getProperty('TARGET_DATE_OVERRIDE') || '').trim();

  let resolved = '';
  let sourceProperty = null;

  if (typeof targetDate !== 'undefined' && targetDate !== null) {
    if (targetDate instanceof Date) {
      resolved = Utilities.formatDate(targetDate, tz, 'yyyy-MM-dd');
    } else if (typeof targetDate === 'string' || typeof targetDate === 'number') {
      resolved = String(targetDate).trim();
    } else if (typeof targetDate === 'object') {
      const candidate =
        (targetDate.parameter && (targetDate.parameter.targetDate || targetDate.parameter.date || targetDate.parameter.TARGET_DATE_OVERRIDE)) ||
        targetDate.targetDate ||
        targetDate.date;
      if (candidate) {
        resolved = String(candidate).trim();
      } else {
        console.info('resolveProcessingWindow_: ignoring non-date argument (assumed trigger event object).');
      }
    } else {
      console.info(`resolveProcessingWindow_: unsupported targetDate type (${typeof targetDate}); ignoring argument.`);
    }
  }

  if (!resolved && overrideProp && overrideProp.toLowerCase() !== 'today' && overrideProp.toLowerCase() !== 'current') {
    resolved = overrideProp;
    sourceProperty = overrideProp;
  }

  const { start, end, label, overrideApplied } = createDateWindow_(resolved, tz);
  return {
    startMs: start,
    endMs: end,
    label,
    overrideApplied,
    sourceProperty
  };
}

function createDateWindow_(candidate, tz) {
  const zone = (tz && tz.trim()) || DEFAULTS.searchTimezone;
  let start = 0;
  let overrideApplied = false;

  if (candidate) {
    let normalized = String(candidate).trim();
    if (/^(today|current)$/i.test(normalized)) {
      normalized = Utilities.formatDate(new Date(), zone, 'yyyy-MM-dd');
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
      throw new Error(`Invalid date override "${candidate}". Expected format YYYY-MM-DD.`);
    }
    start = getStartOfDateMs_(normalized, zone);
    overrideApplied = true;
  } else {
    start = getStartOfTodayMs_(zone);
  }

  const label = Utilities.formatDate(new Date(start), zone, 'yyyy-MM-dd');
  return { start, end: start + ONE_DAY_MS, label, overrideApplied };
}

function buildSearchQuery_(baseQuery, processedLabelName, windowInfo) {
  let query = (baseQuery || '').trim();

  if (processedLabelName) {
    const escaped = processedLabelName.replace(/\\/g, '\\\\').replace(/"/g, '\"');
    const labelFilter = `-label:"${escaped}"`;
    if (!query.includes(labelFilter)) {
      query = query ? `${query} ${labelFilter}` : labelFilter;
    }
  }

  query = appendAfterClause_(query, windowInfo.startMs);
  query = appendBeforeClause_(query, windowInfo.endMs);
  return query;
}

function appendAfterClause_(query, timestampMs) {
  if (!timestampMs) return query;
  const base = query || '';
  if (base.toLowerCase().includes('after:')) return base;
  const clause = `after:${Math.floor(timestampMs / 1000)}`;
  return base ? `${base} ${clause}` : clause;
}

function appendBeforeClause_(query, timestampMs) {
  if (!timestampMs) return query;
  const base = query || '';
  if (base.toLowerCase().includes('before:')) return base;
  const clause = `before:${Math.floor(timestampMs / 1000)}`;
  return base ? `${base} ${clause}` : clause;
}

function getStartOfTodayMs_(tz) {
  const zone = (tz && tz.trim()) || DEFAULTS.searchTimezone;
  const now = new Date();
  const day = Utilities.formatDate(now, zone, 'yyyy-MM-dd');
  const offsetRaw = Utilities.formatDate(now, zone, 'Z');
  const offset = (offsetRaw && offsetRaw.length === 5)
    ? `${offsetRaw.slice(0, 3)}:${offsetRaw.slice(3)}`
    : '+00:00';
  return new Date(`${day}T00:00:00${offset}`).getTime();
}

function getStartOfDateMs_(dateString, tz) {
  const zone = (tz && tz.trim()) || DEFAULTS.searchTimezone;
  const value = (dateString || '').trim();
  const probe = new Date(`${value}T00:00:00Z`);
  if (isNaN(probe.getTime())) {
    throw new Error(`Invalid date override "${dateString}".`);
  }
  const offsetRaw = Utilities.formatDate(probe, zone, 'Z');
  const offset = (offsetRaw && offsetRaw.length === 5)
    ? `${offsetRaw.slice(0, 3)}:${offsetRaw.slice(3)}`
    : '+00:00';
  return new Date(`${value}T00:00:00${offset}`).getTime();
}

/** ────────────────────────────────────────────────
 *  Gmail fetchers
 * ────────────────────────────────────────────────*/

function findNewestZipMessage_(query, lastState, processedLabelId) {
  const options = {
    q: query,
    maxResults: 50,
    includeSpamTrash: false
  };

  let pageToken = null;
  do {
    const resp = Gmail.Users.Messages.list('me', pageToken ? Object.assign({}, options, { pageToken }) : options);
    const messages = resp.messages || [];
    for (const meta of messages) {
      const message = Gmail.Users.Messages.get('me', meta.id, { format: 'full' });
      const internalDate = Number(message.internalDate || 0);
      if (lastState.id && message.id === lastState.id) continue;

      const labels = message.labelIds || [];
      if (processedLabelId && labels.indexOf(processedLabelId) !== -1) continue;

      const zips = listZipAttachments_(message.payload);
      if (!zips.length) continue;

      return {
        id: message.id,
        threadId: message.threadId,
        internalDate,
        payload: message.payload,
        labelIds: labels
      };
    }
    pageToken = resp.nextPageToken || null;
  } while (pageToken);
  return null;
}

function listZipAttachments_(payload) {
  const results = [];
  collectZipAttachments_(payload, results);
  return results;
}

function collectZipAttachments_(payload, out) {
  if (!payload) return;
  const filename = payload.filename || '';
  const body = payload.body || {};
  if (filename && /\.zip$/i.test(filename) && body.attachmentId) {
    out.push({
      attachmentId: body.attachmentId,
      filename,
      mimeType: payload.mimeType || 'application/zip',
      size: body.size || 0
    });
  }
  const parts = payload.parts || [];
  for (const part of parts) collectZipAttachments_(part, out);
}

function fetchAttachmentBlob_(messageId, meta) {
  const attachment = Gmail.Users.Messages.Attachments.get('me', messageId, meta.attachmentId);
  if (!attachment || !attachment.data) {
    throw new Error('downloadAttachmentBlob_: attachment payload missing.');
  }

  const primaryData = String(attachment.data || '').trim();
  const fallback = buildDataUriFallback_(messageId, meta.attachmentId, meta.mimeType);

  console.info(`Attachment meta: declaredSize=${meta.size || 0}, primaryDataLength=${primaryData.length}, hasFallback=${fallback ? 'yes' : 'no'}`);
  console.info(`Attachment primary head: "${primaryData.slice(0, 128)}"`);

  const errors = [];
  let bytes = null;

  try {
    bytes = decodeAttachmentData_(primaryData);
  } catch (err) {
    errors.push(`primary decode failed: ${err && err.message ? err.message : err}`);
  }

  if (!bytes && fallback) {
    try {
      console.info(`Fallback data head: "${fallback.slice(0, 128)}"`);
      bytes = decodeAttachmentData_(fallback);
      console.info(`Attachment decoded via fallback data URI (length=${fallback.length}).`);
    } catch (err) {
      errors.push(`fallback decode failed: ${err && err.message ? err.message : err}`);
    }
  }

  if (!bytes) {
    errors.forEach(msg => console.warn(`downloadAttachmentBlob_: ${msg}`));
    throw new Error('downloadAttachmentBlob_: could not decode attachment payload.');
  }

  const mime = (meta.mimeType || '').toLowerCase().includes('zip') ? meta.mimeType : 'application/zip';
  const name = meta.filename || 'attachment.zip';
  return Utilities.newBlob(bytes, mime || 'application/zip', name);
}

function buildDataUriFallback_(messageId, attachmentId, mimeType) {
  if (!messageId || !attachmentId) return null;
  try {
    const detail = Gmail.Users.Messages.Attachments.get('me', messageId, attachmentId);
    const data = String(detail && detail.data ? detail.data : '').trim();
    if (!data) return null;
    const mime = (mimeType && mimeType.trim()) ? mimeType : 'application/zip';
    return `data:${mime};base64,${data}`;
  } catch (err) {
    console.error(`buildDataUriFallback_: failed to read attachment detail (${attachmentId}): ${err && err.message ? err.message : err}`);
    return null;
  }
}

/** ────────────────────────────────────────────────
 *  Attachment decoding
 * ────────────────────────────────────────────────*/

function decodeAttachmentData_(rawData, allowSplit) {
  const original = (rawData || '').trim();
  if (!original) throw new Error('decodeAttachmentData_: empty payload.');

  // Detect decimal CSV (e.g., "80,75,3,4,..." possibly with negative values)
  const numericCandidate = original.replace(/\s+/g, '');
  if (/^-?\d+(,-?\d+)*$/.test(numericCandidate)) {
    const nums = numericCandidate.split(',').filter(Boolean).map(s => Number(s));
    if (nums.some(n => isNaN(n))) {
      throw new Error('decodeAttachmentData_: numeric CSV contained invalid byte.');
    }
    return nums.map(n => ((n % 256) + 256) % 256);
  }

  const allowSplitFlag = allowSplit !== false;
  const attemptOrder = buildBase64Variants_(original);
  const failures = [];

  for (const variant of attemptOrder) {
    try {
      if (variant.mode === 'webSafe') {
        return Utilities.base64DecodeWebSafe(variant.data);
      }
      return Utilities.base64Decode(variant.data);
    } catch (err) {
      failures.push(`${variant.mode} (${variant.data.length} chars): ${err && err.message ? err.message : err}`);
    }
  }

  const sanitized = original.replace(/\s+/g, '');
  if (allowSplitFlag && sanitized.indexOf(',') !== -1) {
    const parts = sanitized.split(',').map(part => part.trim()).filter(Boolean);
    if (parts.length > 1) {
      try {
        console.info(`decodeAttachmentData_: attempting segmented decode (${parts.length} parts).`);
        const decodedParts = parts.map(part => decodeAttachmentData_(part, false));
        return concatByteArrays_(decodedParts);
      } catch (err) {
        failures.push(`segmented decode failed: ${err && err.message ? err.message : err}`);
      }
    }
  }

  const invalidChars = sanitized.replace(/[A-Za-z0-9+/=_-]/g, '');
  const invalidInfo = invalidChars ? ` invalidChars=${[...new Set(invalidChars.split(''))].join('')}` : '';

  failures.forEach(msg => console.warn(`decodeAttachmentData_: ${msg}`));
  throw new Error(`downloadAttachmentBlob_: could not decode attachment payload.${invalidInfo}`);
}

function buildBase64Variants_(input) {
  const variants = [];
  const seen = new Set();

  const push = (mode, str) => {
    if (!str) return;
    let data = str.replace(/\s+/g, '');
    if (!data) return;
    const needsPad = data.length % 4;
    if (needsPad) data += '='.repeat(4 - needsPad);
    const key = `${mode}:${data}`;
    if (!seen.has(key)) {
      seen.add(key);
      variants.push({ mode, data });
    }
  };

  const raw = input.replace(/\s+/g, '');
  const stripped = raw.replace(/[^A-Za-z0-9+/=_-]/g, '');

  push('webSafe', raw);
  push('standard', raw.replace(/-/g, '+').replace(/_/g, '/'));
  if (stripped !== raw) {
    push('webSafe', stripped);
    push('standard', stripped.replace(/-/g, '+').replace(/_/g, '/'));
  }

  return variants;
}

/** ────────────────────────────────────────────────
 *  ZIP helpers & TXT extraction
 * ────────────────────────────────────────────────*/

function inspectZip_(blob) {
  const bytes = blob.getBytes();
  const a = new Uint8Array(bytes);
  const le16 = (arr, offset) => arr[offset] | (arr[offset + 1] << 8);
  const le32 = (arr, offset) => (arr[offset] | (arr[offset + 1] << 8) | (arr[offset + 2] << 16) | (arr[offset + 3] << 24)) >>> 0;

  let eocd = -1;
  for (let i = a.length - 22; i >= 0; i--) {
    if (a[i] === 0x50 && a[i + 1] === 0x4b && a[i + 2] === 0x05 && a[i + 3] === 0x06) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return [];

  const cdSize = le32(a, eocd + 12);
  const cdOffs = le32(a, eocd + 16);
  const result = [];
  let p = cdOffs;
  const end = cdOffs + cdSize;

  while (p + 46 <= end) {
    if (!(a[p] === 0x50 && a[p + 1] === 0x4b && a[p + 2] === 0x01 && a[p + 3] === 0x02)) break;
    const gpFlag = le16(a, p + 8);
    const method = le16(a, p + 10);
    const nameLen = le16(a, p + 28);
    const extraLen = le16(a, p + 30);
    const commLen = le16(a, p + 32);
    const nameStart = p + 46;
    let name = '';
    for (let i = 0; i < nameLen; i++) name += String.fromCharCode(a[nameStart + i]);
    const encrypted = (gpFlag & 0x1) !== 0;
    result.push({ name, method, encrypted });
    p = nameStart + nameLen + extraLen + commLen;
  }

  if (result.length) {
    const preview = result.slice(0, 5).map(e => `${e.name} [${zipMethodName_(e.method)}${e.encrypted ? ', enc' : ''}]`).join('; ');
    console.info('🧭 ZIP entries (first few): ' + preview);
  }

  return result;
}

function zipMethodName_(method) {
  const map = { 0: 'Store', 8: 'Deflate', 9: 'Deflate64', 12: 'BZIP2', 14: 'LZMA', 98: 'PPMd', 99: 'AES' };
  return map[method] || `method ${method}`;
}

function extractTxtBodies_(blobs, charset) {
  const txtBlobs = (blobs || []).filter(b => /\.txt$/i.test(b.getName?.() || ''));
  const bodies = [];
  const names = [];

  for (const blob of txtBlobs) {
    const text = decodeTxtBlob_(blob, charset);
    if (text !== null) {
      const clean = text.replace(/^\uFEFF/, '').trim();
      bodies.push(`----- ${blob.getName()} -----\n` + clean);
      names.push(blob.getName());
    } else {
      console.error(`extractTxtBodies_: could not decode ${blob.getName()} with any charset; skipping.`);
    }
  }
  return { bodies, names };
}

/** ────────────────────────────────────────────────
 *  Salesforce REST helpers
 * ────────────────────────────────────────────────*/

function postTxtBodiesToSalesforce_(txtBodies, txtNames, config) {
  if (!config.salesforceClientId || !config.salesforceClientSecret || !config.salesforceUsername || !config.salesforcePassword) {
    throw new Error('Salesforce credentials missing. Please set SALESFORCE_CLIENT_ID/SECRET/USERNAME/PASSWORD in Script Properties (password should include security token).');
  }

  const auth = fetchSalesforceAccessToken_(config);
  const baseUrl = (config.salesforceInstanceUrl || auth.instanceUrl || '').replace(/\/$/, '');
  if (!baseUrl) {
    throw new Error('Salesforce instance URL missing. Set SALESFORCE_INSTANCE_URL or ensure the token response includes instance_url.');
  }
  console.info(`🔐 Salesforce auth OK. instance=${baseUrl}`);

  const pathOverride = (config.salesforceApiPath || DEFAULTS.salesforceApiPath || '').trim();
  let url;
  if (pathOverride) {
    url = `${baseUrl}${pathOverride.startsWith('/') ? '' : '/'}${pathOverride}`;
  } else {
    const apiVersion = config.salesforceApiVersion || DEFAULTS.salesforceApiVersion;
    const sobject = config.salesforceSobject || DEFAULTS.salesforceSobject;
    url = `${baseUrl}/services/data/v${apiVersion}/sobjects/${sobject}`;
  }

  let successCount = 0;
  let errorCount = 0;

  txtBodies.forEach((body, idx) => {
    const name = txtNames[idx] || `file-${idx + 1}.txt`;
    const payload = mapTxtToSalesforcePayload_(name, body, config.salesforceSobject || DEFAULTS.salesforceSobject, config.searchTimezone);
    try {
      const resp = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        headers: {
          Authorization: `Bearer ${auth.accessToken}`
        },
        muteHttpExceptions: true
      });
      const status = resp.getResponseCode();
      if (status >= 200 && status < 300) {
        successCount++;
        let idInfo = '';
        try {
          const parsed = JSON.parse(resp.getContentText() || '{}');
          if (parsed.id) idInfo = ` id=${parsed.id}`;
        } catch (_) {
          // ignore parse errors for success responses
        }
        console.info(`✅ Salesforce insert OK for ${name} (status=${status}${idInfo})`);
      } else {
        errorCount++;
        console.error(`Salesforce POST failed for ${name}: status=${status} body=${resp.getContentText()}`);
      }
    } catch (err) {
      errorCount++;
      console.error(`Salesforce POST threw for ${name}: ${err && err.message ? err.message : err}`);
    }
  });

  return { successCount, errorCount };
}

function fetchSalesforceAccessToken_(config) {
  const tokenUrl = (config.salesforceLoginUrl || DEFAULTS.salesforceLoginUrl || '').replace(/\/$/, '') + '/services/oauth2/token';
  console.info(`🔑 Salesforce login: requesting token at ${tokenUrl}`);
  const payload = {
    grant_type: 'password',
    client_id: config.salesforceClientId,
    client_secret: config.salesforceClientSecret,
    username: config.salesforceUsername,
    password: config.salesforcePassword
  };

  const urlEncodedPayload = Object.keys(payload)
    .map((k) => `${encodeURIComponent(k)}=${encodeURIComponent(payload[k])}`)
    .join('&');

  const resp = UrlFetchApp.fetch(tokenUrl, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: urlEncodedPayload,
    muteHttpExceptions: true
  });

  const status = resp.getResponseCode();
  const text = resp.getContentText();
  console.info(`🔑 Salesforce login response status=${status}`);
  if (status < 200 || status >= 300) {
    const errHint = buildSalesforceGrantHint_(text);
    throw new Error(`Salesforce token request failed: status=${status} body=${text}${errHint}`);
  }

  let parsed = null;
  try {
    parsed = JSON.parse(text);
  } catch (err) {
    throw new Error(`Salesforce token response was not JSON: ${err && err.message ? err.message : err}`);
  }

  if (!parsed.access_token) {
    throw new Error('Salesforce token response missing access_token.');
  }

  const resolvedInstance = parsed.instance_url || config.salesforceInstanceUrl;
  console.info(`🔑 Salesforce login success (instance=${resolvedInstance || 'unknown'})`);

  return {
    accessToken: parsed.access_token,
    instanceUrl: resolvedInstance
  };
}

function buildSalesforceGrantHint_(rawBody) {
  if (!rawBody) return '';
  try {
    const parsed = JSON.parse(rawBody);
    if (parsed.error === 'invalid_grant') {
      return ' (invalid_grant: check username/password/security token combination, IP restrictions, and login URL vs instance type)';
    }
  } catch (_) {
    // ignore parse failure
  }
  return '';
}

function mapTxtToSalesforcePayload_(filename, rawText, sobjectName, timezone) {
  const parsed = parseTxtRequest_(rawText, timezone);
  const comments = [
    parsed.product ? `商品: ${parsed.product}` : '',
    parsed.bodyType ? `ボディタイプ: ${parsed.bodyType}` : '',
    parsed.carCondition ? `クルマの状態: ${parsed.carCondition}` : ''
  ].filter(Boolean).join(' / ');

  return {
    attributes: { type: sobjectName || 'Mail2X__c' },
    RequestDate__c: parsed.requestDateIso || parsed.requestDate,
    AssessmentNumber__c: parsed.assessmentNumber,
    comment__c: comments || undefined,
    maker__c: parsed.brand,
    car_model__c: parsed.carModel,
    model_year__c: parsed.modelYear,
    Grade__c: parsed.grade,
    body_color__c: parsed.bodyColor,
    DoorNumber__c: parsed.doorCount,
    Handle__c: parsed.handle,
    Fuel__c: parsed.fuel,
    Transmission__c: parsed.transmission,
    DriveType__c: parsed.driveType,
    Displacement__c: parsed.displacement,
    mileage__c: parsed.mileage,
    accident_history__c: parsed.accidentHistory,
    desired_time_to_sell__c: parsed.desiredSellTiming,
    Model__c: parsed.modelCode,
    EquipmentInfo__c: parsed.equipmentInfo,
    appeal_point__c: parsed.otherOptions,
    name__c: parsed.customerName,
    name_kana__c: parsed.customerKana,
    PostalCode__c: parsed.postalCode,
    State__c: parsed.state,
    City__c: parsed.city,
    Address__c: parsed.addressLine,
    mail__c: parsed.email,
    Phone__c: parsed.phone,
    Phone2__c: parsed.phone2,
    preferred_contact_time__c: parsed.contactTime
  };
}

function parseTxtRequest_(rawText, timezone) {
  const normalized = (rawText || '').replace(/\r/g, '');
  const valueFor = (labels) => extractLabelValue_(normalized, labels);
  const labelSet = (label, extras) => labelAliases_(label, extras);

  const reqMatch = normalized.match(/査定依頼日時・査定依頼番号][\s\S]*?([0-9]{4}年\d{1,2}月\d{1,2}日[^(\n（]*)[（(]([0-9]+)[)）]/);
  const requestDate = reqMatch ? reqMatch[1].trim() : '';
  const assessmentNumber = reqMatch ? reqMatch[2].trim() : '';
  const requestDateIso = requestDate ? parseJapaneseDateTimeToIsoUtc_(requestDate, timezone) : '';
  const product = valueFor(labelSet('商品'));

  const brand = valueFor(labelSet('ブランド名', ['メーカー', 'メーカー名']));
  const carModel = valueFor(labelSet('車種名', ['車名']));
  const modelYear = valueFor(labelSet('年式'));
  const grade = valueFor(labelSet('グレード'));
  const bodyType = valueFor(labelSet('ボディタイプ', ['ボディタイプ・カテゴリ', 'ボディタイプカテゴリ']));
  const compositeBodyAndDoorRaw = valueFor(['車体色・ドア数', '車体色/ドア数', '車体色･ドア数']);
  const standaloneBodyColor = valueFor(labelSet('色', ['車体色', 'ボディカラー', 'カラー']));
  const standaloneDoorCount = valueFor(labelSet('ドア数'));
  const compositeBodyAndDoor = splitBodyColorAndDoor_(compositeBodyAndDoorRaw);
  const bodyColor = standaloneBodyColor || compositeBodyAndDoor.bodyColor;
  const doorCount = standaloneDoorCount || compositeBodyAndDoor.doorCount;
  const handle = valueFor(labelSet('ハンドル'));
  const fuel = valueFor(labelSet('燃料'));
  const transmission = valueFor(labelSet('ミッション', ['トランスミッション']));
  const driveType = valueFor(labelSet('駆動方式', ['駆動']));
  const displacement = valueFor(labelSet('排気量'));
  const mileage = valueFor(labelSet('走行距離'));
  const inspectionDeadline = valueFor(labelSet('車検時期', ['車検満了日']));
  const accidentHistory = valueFor(labelSet('事故歴'));
  const carCondition = valueFor(labelSet('クルマの状態', ['車の状態']));
  const desiredSellTiming = valueFor(labelSet('売却希望時期', ['売却希望時期目安', '売却希望時期・目安']));
  const compositeModelEquipRaw = valueFor(['型式・装備', '型式/装備', '型式･装備']);
  const compositeModelEquip = splitModelAndEquipment_(compositeModelEquipRaw);
  const modelCode = valueFor(labelSet('型式')) || compositeModelEquip.modelCode;
  const equipmentInfo = valueFor(labelSet('装備')) || compositeModelEquip.equipmentInfo;
  const otherOptions = valueFor(labelSet('その他オプション等', ['その他オプション', 'その他装備']));

  const customerName = valueFor(labelSet('ご依頼者名', ['氏名'])).replace(/様$/, '');
  const customerKana = valueFor(labelSet('ご依頼者カナ名')).replace(/様$/, '');
  const postalCode = valueFor(labelSet('郵便番号'));
  const addressFull = valueFor(labelSet('ご住所', ['住所']));
  const addressParts = splitJapaneseAddress_(addressFull);
  const email = valueFor(labelSet('メールアドレス'));
  const phone = valueFor(labelSet('電話番号'));
  const phone2 = valueFor(labelSet('その他の連絡先', ['サブ連絡先']));
  const contactTime = valueFor(labelSet('連絡可能時間帯', ['連絡希望時間帯']));

  return {
    requestDate,
    requestDateIso,
    assessmentNumber,
    product,
    brand,
    carModel,
    modelYear,
    grade,
    bodyType,
    bodyColor,
    doorCount,
    handle,
    fuel,
    transmission,
    driveType,
    displacement,
    mileage,
    inspectionDeadline,
    accidentHistory,
    carCondition,
    desiredSellTiming,
    modelCode,
    equipmentInfo,
    otherOptions,
    customerName,
    customerKana,
    postalCode,
    state: addressParts.state,
    city: addressParts.city,
    addressLine: addressParts.address,
    email,
    phone,
    phone2,
    contactTime
  };
}

function extractLabelValue_(text, labels) {
  if (!text || !labels) return '';
  const labelList = Array.isArray(labels) ? labels : [labels];
  for (const label of labelList) {
    if (!label) continue;
    const escaped = escapeRegex_(label);
    const re = new RegExp('(?:^|[\\n\\r])[\\u30fb・\\u2022\\s　]*' + escaped + '\\s*[:：]\\s*([^\\n\\r]+)', 'i');
    const m = text.match(re);
    if (m) {
      return cleanupLabelValue_(m[1]);
    }
  }
  return '';
}

function cleanupLabelValue_(value) {
  return (value || '').replace(/・?ラベル$/i, '').trim();
}

function labelAliases_(label, extras) {
  const addon = extras || [];
  return [label, `${label}・ラベル`, `${label}ラベル`, ...addon];
}

function splitBodyColorAndDoor_(raw) {
  const cleaned = cleanupLabelValue_(raw);
  if (!cleaned) return { bodyColor: '', doorCount: '' };

  const normalizedDigits = normalizeZenkakuDigits_(cleaned);
  const doorMatch = normalizedDigits.match(/([0-9]+)\s*(?:ドア|Ｄ|D(?:oor)?|DOOR)?/i);
  const separatorParts = cleaned.split(/[／/|｜・･、,]/).map(p => cleanupLabelValue_(p)).filter(Boolean);

  const bodyColor = doorMatch ? trimTrailingSeparators_(cleanupLabelValue_(cleaned.slice(0, doorMatch.index))) || cleanupLabelValue_(cleaned.slice(doorMatch.index + doorMatch[0].length)) : '';
  const doorCount = doorMatch ? extractDoorDigits_(doorMatch[1]) : '';

  const fallbackBodyColor = separatorParts.length ? separatorParts[0] : '';
  const fallbackDoorSegment = separatorParts.length > 1 ? separatorParts.slice(1).find(p => /ドア|door/i.test(p)) || separatorParts[1] : '';
  const fallbackDoorCount = extractDoorDigits_(fallbackDoorSegment);

  return {
    bodyColor: bodyColor || fallbackBodyColor,
    doorCount: doorCount || fallbackDoorCount
  };
}

function splitModelAndEquipment_(raw) {
  const cleaned = cleanupLabelValue_(raw);
  if (!cleaned) return { modelCode: '', equipmentInfo: '' };
  const parts = cleaned.split(/[／/|｜・･、,，]/).map(p => cleanupLabelValue_(p)).filter(Boolean);
  if (parts.length >= 2) {
    return { modelCode: parts[0], equipmentInfo: parts.slice(1).join('/') };
  }
  return { modelCode: parts[0] || cleaned, equipmentInfo: '' };
}

function normalizeZenkakuDigits_(value) {
  return (value || '').replace(/[０-９]/g, (d) => String('０１２３４５６７８９'.indexOf(d)));
}

function extractDoorDigits_(value) {
  const normalized = normalizeZenkakuDigits_(value);
  const m = normalized.match(/([0-9]+)/);
  return m ? m[1] : '';
}

function escapeRegex_(raw) {
  return (raw || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function trimTrailingSeparators_(value) {
  return (value || '').replace(/[／/|｜・･、,]+$/, '').trim();
}

function splitJapaneseAddress_(full) {
  const value = (full || '').trim();
  if (!value) return { state: '', city: '', address: '' };

  const m = value.match(/^(.+?[都道府県])(.+?[市区町村郡])?(.*)$/);
  if (m) {
    return {
      state: (m[1] || '').trim(),
      city: (m[2] || '').trim(),
      address: (m[3] || '').trim()
    };
  }

  return { state: '', city: '', address: value };
}

function parseJapaneseDateTimeToIsoUtc_(raw, timezone) {
  const tz = (timezone || DEFAULTS.searchTimezone || 'Asia/Tokyo').trim();
  const m = raw.match(/(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日\s*(\d{1,2})時\s*(\d{1,2})分/);
  if (!m) return '';
  const [_, y, mo, d, h, mi] = m;
  const yyyy = Number(y);
  const month = Number(mo);
  const day = Number(d);
  const hour = Number(h);
  const minute = Number(mi);
  const pad = (n) => (n < 10 ? `0${n}` : String(n));

  const probe = new Date(`${yyyy}-${pad(month)}-${pad(day)}T00:00:00Z`);
  const offsetRaw = Utilities.formatDate(probe, tz, 'Z'); // e.g. +0900
  const offset = (offsetRaw && offsetRaw.length === 5) ? `${offsetRaw.slice(0, 3)}:${offsetRaw.slice(3)}` : '+00:00';
  const isoLocal = `${yyyy}-${pad(month)}-${pad(day)}T${pad(hour)}:${pad(minute)}:00${offset}`;
  const asDate = new Date(isoLocal);
  if (isNaN(asDate.getTime())) return '';
  return Utilities.formatDate(asDate, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

/** ────────────────────────────────────────────────
 *  Misc helpers
 * ────────────────────────────────────────────────*/

function getHeaderValue_(headers, name) {
  if (!headers) return '';
  const needle = (name || '').toLowerCase();
  for (const header of headers) {
    if ((header.name || '').toLowerCase() === needle) return header.value || '';
  }
  return '';
}

/** ────────────────────────────────────────────────
 *  GASunzip + TXT helpers
 * ────────────────────────────────────────────────*/

function unzipWithGASunzip_(zipBlob, password) {
  if (typeof GASunzip === 'undefined' || !GASunzip.unzip) {
    throw new Error('GASunzip library not found or identifier not set to \"GASunzip\".');
  }
  if (!zipBlob || !zipBlob.getBytes) {
    throw new Error('Please set a file blob of zip file. (zipBlob missing/invalid)');
  }

  let castZip;
  try {
    castZip = zipBlob.getAs('application/zip');
  } catch (err) {
    castZip = zipBlob;
  }

  const opts = password ? { password } : {};
  try {
    return GASunzip.unzip(castZip, opts);
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    throw new Error(`GASunzip.unzip failed: ${msg}`);
  }
}

function decodeTxtBlob_(blob, preferredCharset) {
  const name = blob.getName?.() || 'file.txt';

  const candidates = [
    preferredCharset || 'Shift_JIS',
    (Utilities.Charset && Utilities.Charset.SHIFT_JIS) ? Utilities.Charset.SHIFT_JIS : null,
    'MS932',
    'Windows-31J',
    'CP932',
    'SJIS'
  ].filter(Boolean);

  const attempt = (candidateBlob) => {
    if (!candidateBlob) return null;
    for (const cs of candidates) {
      try {
        const text = candidateBlob.getDataAsString(cs);
        if (typeof text === 'string' && text.length) return text;
      } catch (err) {
        console.warn(`decodeTxtBlob_: ${name} failed with charset \"${cs}\": ${err && err.message ? err.message : err}`);
      }
    }
    try {
      const fallback = candidateBlob.getDataAsString();
      if (typeof fallback === 'string' && fallback.length) return fallback;
    } catch (err) {
      console.warn(`decodeTxtBlob_: ${name} final-guess failed: ${err && err.message ? err.message : err}`);
    }
    return null;
  };

  try {
    const bytes = blob.getBytes();
    const plain = Utilities.newBlob(bytes, 'text/plain', name);
    const text = attempt(plain);
    if (text !== null) return text;
  } catch (err) {
    console.warn(`decodeTxtBlob_: getBytes/newBlob path failed for ${name}: ${err && err.message ? err.message : err}`);
  }

  try {
    const asPlain = blob.getAs && blob.getAs('text/plain');
    const text = attempt(asPlain);
    if (text !== null) return text;
  } catch (err) {
    console.warn(`decodeTxtBlob_: getAs('text/plain') failed for ${name}: ${err && err.message ? err.message : err}`);
  }

  try {
    const copy = blob.copyBlob().setContentType('text/plain');
    const text = attempt(copy);
    if (text !== null) return text;
  } catch (err) {
    console.warn(`decodeTxtBlob_: copyBlob().setContentType('text/plain') failed for ${name}: ${err && err.message ? err.message : err}`);
  }

  console.error(`decodeTxtBlob_: could not decode ${name} with any method.`);
  return null;
}
