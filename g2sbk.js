/****************************************************
 * Gmail → ZIP (GASunzip) → Email TXT  [V8 runtime]
 *
 * Requires the GASunzip library (ID below) and the following
 * Script Properties (File → Project properties → Script properties):
 *
 *   TARGET_EMAIL              (required)  e.g. you@example.com
 *   ALLOWED_SENDER           (optional)  e.g. sender@example.com
 *   GMAIL_QUERY              (optional)  default below
 *   ZIP_PASSWORD             (optional)  leave blank if no password
 *   EMAIL_SUBJECT_TEMPLATE   (optional)  default: "CSV(s) from {zip}"
 *   PROCESSED_LABEL          (optional)  default: "Unzip/processed"
 *   SEARCH_TIMEZONE          (optional)  default: "Asia/Tokyo"
 *   TARGET_DATE_OVERRIDE     (optional)  e.g. "2025-10-24" to reprocess that day
 ****************************************************/

const PROPS = PropertiesService.getScriptProperties();

const DEFAULTS = {
  gmailQuery: 'in:anywhere has:attachment filename:zip newer_than:2d',
  processedLabel: 'Unzip/processed',
  emailSubjectTemplate: 'CSV(s) from {zip}',
  searchTimezone: 'Asia/Tokyo'
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
      console.warn('No TXT files decoded; using fallback summary for email body.');
    } else {
      console.info(`📝 TXT(s) extracted: ${txtNames.join(', ')}`);
    }

    const emailBody = txtBodies.length ? txtBodies.join('\n\n') : buildFallbackSummary_(zipName, fromHeaderRaw, searchQuery, windowInfo, extractedCount);
    const subject = buildSubject_(config.emailSubjectTemplate, zipName);

    MailApp.sendEmail({
      to: config.targetEmail,
      subject,
      body: emailBody
    });

    markThreadProcessed_(message.threadId, config.processedLabelId);
    updateLastProcessedState_(message.internalDate, message.id);
    console.info(`✅ Done. Extracted=${extractedCount}, TXTs=${txtBodies.length} → emailed to ${config.targetEmail}`);
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
  const targetEmail = requireTargetEmail_();
  const allowedSender = (PROPS.getProperty('ALLOWED_SENDER') || '').trim().toLowerCase();
  const gmailQuery = (PROPS.getProperty('GMAIL_QUERY') || DEFAULTS.gmailQuery).trim();
  const zipPassword = (PROPS.getProperty('ZIP_PASSWORD') || '').trim();
  const processedLabel = (PROPS.getProperty('PROCESSED_LABEL') || DEFAULTS.processedLabel).trim();
  const emailSubjectTemplate = (PROPS.getProperty('EMAIL_SUBJECT_TEMPLATE') || DEFAULTS.emailSubjectTemplate);
  const searchTimezone = (PROPS.getProperty('SEARCH_TIMEZONE') || DEFAULTS.searchTimezone).trim() || DEFAULTS.searchTimezone;

  const processedLabelId = processedLabel ? ensureLabelId_(processedLabel) : null;

  return {
    targetEmail,
    allowedSender,
    gmailQuery,
    zipPassword,
    processedLabel,
    processedLabelId,
    emailSubjectTemplate,
    searchTimezone
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
 *  Misc helpers
 * ────────────────────────────────────────────────*/

function isValidEmail_(s) {
  if (!s || typeof s !== 'string') return false;
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(s.trim());
}

function requireTargetEmail_() {
  const value = (PROPS.getProperty('TARGET_EMAIL') || '').trim();
  if (!isValidEmail_(value)) {
    throw new Error('TARGET_EMAIL in Script Properties is missing or invalid.');
  }
  return value;
}

function buildSubject_(template, zipName) {
  try {
    return (template || '').replace(/\{zip\}/g, zipName || '');
  } catch (err) {
    return `CSV(s) from ${zipName || ''}`;
  }
}

function buildFallbackSummary_(zipName, fromHeader, query, windowInfo, count) {
  return [
    `Extracted ${count} file(s) from ${zipName}.`,
    '',
    `Date: ${windowInfo.label}`,
    `From: ${fromHeader}`,
    `Query: ${query}`
  ].join('\n');
}

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
