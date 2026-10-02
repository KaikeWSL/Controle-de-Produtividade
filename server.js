const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const zlib = require('zlib');
const cors = require('cors');
const compression = require('compression');

const app = express();
const port = process.env.PORT || 3000;
const cacheTtlMs = Math.max(30, Number(process.env.CACHE_TTL_SECONDS) || 300) * 1000;
const cacheMaxItems = Math.max(1000, Number(process.env.CACHE_MAX_ITEMS) || 100000);
const cacheCleanupIntervalMs = Math.max(30, Number(process.env.CACHE_CLEANUP_INTERVAL) || 60) * 1000;
const cacheWarmupEnabled = String(process.env.CACHE_WARMUP_ENABLED || 'true').toLowerCase() !== 'false';
const cacheDirectory = path.join(__dirname, process.env.CACHE_DIRECTORY || '.cache');
const rateLimitWindowMs = 60 * 1000;
const rateLimitMaxRequests = Math.max(10, Number(process.env.RATE_LIMIT_MAX_REQUESTS) || 120);

const configPath = path.join(__dirname, 'config.json');

function loadConfig() {
  try {
    if (process.env.APP_CONFIG) {
      return JSON.parse(process.env.APP_CONFIG);
    }
  } catch (error) {
    console.warn('APP_CONFIG inválido. Usando config.json local.');
  }

  try {
    const rawConfig = fs.readFileSync(configPath, 'utf8');
    const parsed = JSON.parse(rawConfig);

    return {
      tenantId: process.env.TENANT_ID || parsed.tenantId,
      clientId: process.env.CLIENT_ID || parsed.clientId,
      clientSecret: process.env.CLIENT_SECRET || parsed.clientSecret,
      sites: process.env.SHAREPOINT_SITES ? JSON.parse(process.env.SHAREPOINT_SITES) : parsed.sites
    };
  } catch (error) {
    console.error('Falha ao carregar a configuração do SharePoint.', error.message);
    return {
      tenantId: process.env.TENANT_ID || '',
      clientId: process.env.CLIENT_ID || '',
      clientSecret: process.env.CLIENT_SECRET || '',
      sites: process.env.SHAREPOINT_SITES ? JSON.parse(process.env.SHAREPOINT_SITES) : []
    };
  }
}

const config = loadConfig();
const dashboardCache = new Map();
const weightSourceCache = new Map();
const sourceListCache = new Map();
const sourceProfileCache = new Map();
const analyticsCache = new Map();
const sourceLoadPromises = new Map();
const sourceProfilePromises = new Map();
const weightLoadPromises = new Map();
const requestBuckets = new Map();
const dashboardCacheMaxEntries = Math.max(10, Number(process.env.DASHBOARD_CACHE_MAX_ENTRIES) || 30);
const cacheMetrics = {
  hits: 0,
  misses: 0,
  graphFetches: 0,
  tokenRefreshes: 0,
  sourceLoadDeduplications: 0,
  profileLoadDeduplications: 0,
  weightLoadDeduplications: 0
};
let graphTokenCache = null;
let graphTokenPromise = null;

fs.mkdirSync(cacheDirectory, { recursive: true });

function getPublicConfig(source) {
  return { sites: Array.isArray(source.sites) ? source.sites : [] };
}

function applySecurityHeaders(req, res, next) {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'SAMEORIGIN');
  res.setHeader('Referrer-Policy', 'strict-origin-when-cross-origin');
  res.setHeader('Permissions-Policy', 'geolocation=(), microphone=(), camera=()');
  res.setHeader('Content-Security-Policy', "default-src 'self'; script-src 'self' https://cdn.jsdelivr.net; style-src 'self' https://fonts.googleapis.com 'unsafe-inline'; font-src 'self' https://fonts.gstatic.com; img-src 'self' data:; connect-src 'self' https://rak-5ph3.onrender.com");
  next();
}

function rateLimit(req, res, next) {
  const key = req.ip || req.socket.remoteAddress || 'unknown';
  const now = Date.now();
  const bucket = requestBuckets.get(key);
  const activeBucket = !bucket || now - bucket.startedAt >= rateLimitWindowMs ? { startedAt: now, requests: 0 } : bucket;
  activeBucket.requests += 1;
  requestBuckets.set(key, activeBucket);
  res.setHeader('RateLimit-Limit', rateLimitMaxRequests);
  res.setHeader('RateLimit-Remaining', Math.max(0, rateLimitMaxRequests - activeBucket.requests));

  if (activeBucket.requests > rateLimitMaxRequests) {
    res.setHeader('Retry-After', Math.ceil((rateLimitWindowMs - (now - activeBucket.startedAt)) / 1000));
    return res.status(429).json({ message: 'Limite de requisições excedido. Tente novamente em instantes.' });
  }

  next();
}

function validateDashboardQuery(req, res, next) {
  const allowedKeys = new Set(['site', 'activity', 'subactivity', 'projetista', 'uf', 'cidade', 'mes', 'ano', 'refresh', 'page', 'pageSize']);
  const invalidKey = Object.keys(req.query).find((key) => !allowedKeys.has(key));
  const invalidText = ['site', 'activity', 'subactivity', 'projetista', 'uf', 'cidade'].find((key) => String(req.query[key] || '').length > 120);
  const month = String(req.query.mes || '');
  const year = String(req.query.ano || '');
  const page = String(req.query.page || '');
  const pageSize = String(req.query.pageSize || '');

  if (invalidKey || invalidText || (month && !/^(?:[1-9]|1[0-2])$/.test(month)) || (year && !/^\d{4}$/.test(year)) || (page && !/^\d+$/.test(page)) || (pageSize && !/^\d+$/.test(pageSize))) {
    return res.status(400).json({ message: 'Parâmetros de filtro inválidos.' });
  }

  next();
}

app.use(applySecurityHeaders);
app.use(compression({ threshold: 1024 }));
app.use(rateLimit);
app.use((req, res, next) => {
  const requestId = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  req.requestId = requestId;
  res.setHeader('X-Request-Id', requestId);
  console.log(JSON.stringify({ event: 'request', requestId, method: req.method, path: req.path }));
  next();
});

function buildDashboardCacheKey(siteName, activityName, subActivityName, filters = {}) {
  return JSON.stringify({
    siteName: siteName || '',
    activityName: activityName || '',
    subActivityName: subActivityName || '',
    filters: {
      projetista: filters.projetista || '',
      uf: filters.uf || '',
      cidade: filters.cidade || '',
      mes: filters.mes || '',
      ano: filters.ano || ''
    }
  });
}

function buildPhysicalSourceCacheKey(site) {
  const sourceUrl = new URL(site.url);
  return JSON.stringify({
    origin: sourceUrl.origin.toLowerCase(),
    path: sourceUrl.pathname.replace(/\/+$/, '').toLowerCase(),
    listName: String(site.listName || '').trim().toLowerCase()
  });
}

function buildSourceProfileCacheKey(site) {
  return JSON.stringify({
    source: buildPhysicalSourceCacheKey(site),
    fields: {
      projetista: site.fields?.projetista || '',
      dataConclusao: site.fields?.dataConclusao || '',
      uploadVisium: site.fields?.uploadVisium || '',
      uf: site.fields?.uf || '',
      cidade: site.fields?.cidade || ''
    },
    weightSource: site.weightSource || null
  });
}

function getCacheFilePath(cacheKey) {
  const hash = crypto.createHash('sha256').update(cacheKey).digest('hex');
  return path.join(cacheDirectory, `source-${hash}.json.gz`);
}

function touchCacheEntry(cache, key, entry) {
  cache.delete(key);
  entry.lastAccessedAt = Date.now();
  cache.set(key, entry);
  return entry;
}

function totalSourceItems() {
  let total = 0;
  for (const entry of sourceListCache.values()) total += entry.items.length;
  return total;
}

function totalProfileItems() {
  let total = 0;
  for (const entry of sourceProfileCache.values()) total += entry.items.length;
  return total;
}

function totalCachedItems() {
  return totalSourceItems() + totalProfileItems();
}

function removeProfilesForSource(sourceCacheKey) {
  for (const [profileCacheKey, profile] of sourceProfileCache) {
    if (profile.sourceCacheKey === sourceCacheKey) sourceProfileCache.delete(profileCacheKey);
  }
}

function removeSourceSnapshot(sourceCacheKey) {
  sourceListCache.delete(sourceCacheKey);
  removeProfilesForSource(sourceCacheKey);
}

function evictSourceCache() {
  const now = Date.now();
  for (const [key, entry] of sourceListCache) {
    if (entry.expiresAt <= now) removeSourceSnapshot(key);
  }
  for (const [key, entry] of sourceProfileCache) {
    if (entry.expiresAt <= now) sourceProfileCache.delete(key);
  }
  while (totalCachedItems() > cacheMaxItems && sourceProfileCache.size) {
    sourceProfileCache.delete(sourceProfileCache.keys().next().value);
  }
  while (totalCachedItems() > cacheMaxItems && sourceListCache.size) {
    removeSourceSnapshot(sourceListCache.keys().next().value);
  }
}

function cleanupCaches() {
  const now = Date.now();
  [dashboardCache, weightSourceCache, analyticsCache].forEach((cache) => {
    for (const [key, entry] of cache) {
      if (entry.expiresAt <= now) cache.delete(key);
    }
  });
  evictSourceCache();
}

setInterval(cleanupCaches, cacheCleanupIntervalMs).unref();

const allowedOrigins = (process.env.ALLOWED_ORIGINS || 'http://localhost:3000,http://localhost:8000')
  .split(',')
  .map((origin) => origin.trim())
  .filter(Boolean);

app.use(cors({
  origin(origin, callback) {
    if (!origin || allowedOrigins.includes(origin)) return callback(null, true);
    return callback(new Error('Origem não permitida pelo CORS.'));
  },
  credentials: true,
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));
app.use(express.json({ limit: '8mb' }));

const monthNames = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
const summaryAttachmentTarget = {
  siteUrl: 'https://corpclarobr.sharepoint.com/sites/USER-USER-EquipeProcisacpia',
  listName: 'Tabela1',
  title: 'Resumo',
  fileName: 'Resumo.png'
};

function getSiteByName(name) {
  return config.sites.find((site) => (site.name || site.siteName) === name) || config.sites[0];
}

function resolveSelectedSite(siteName, activityName, subActivityName) {
  const site = getSiteByName(siteName);
  if (!site) return null;

  if (Array.isArray(site.activities) && site.activities.length) {
    const selectedActivity = site.activities.find((activity) => (activity.name || activity.activityName) === activityName)
      || site.activities[0];

    if (Array.isArray(selectedActivity?.activities) && selectedActivity.activities.length) {
      const selectedSubActivity = selectedActivity.activities.find((activity) => (activity.name || activity.activityName) === subActivityName)
        || selectedActivity.activities[0];

      return normalizeSiteConfig({
        ...site,
        ...selectedActivity,
        ...selectedSubActivity,
        name: site.name,
        activityName: selectedActivity.name || selectedActivity.activityName || '',
        subActivityName: selectedSubActivity.name || selectedSubActivity.activityName || '',
        activities: site.activities
      });
    }

    return normalizeSiteConfig({
      ...site,
      ...selectedActivity,
      name: site.name,
      activityName: selectedActivity.name || selectedActivity.activityName || '',
      activities: site.activities
    });
  }

  return normalizeSiteConfig(site);
}

function normalizeSiteConfig(site) {
  if (!site) return null;

  return {
    ...site,
    name: site.name || site.siteName || 'Site',
    url: site.url || site.siteUrl || '',
    listName: site.listName || site.listTitle || site.title || '',
    activityName: site.activityName || site.name || '',
    fields: {
      ...site.fields,
      projetista: site.fields?.projetista || site.fields?.project || '',
      dataConclusao: site.fields?.dataConclusao || site.fields?.dateConclusion || '',
      uploadVisium: site.fields?.uploadVisium || site.fields?.uploadVisiumField || '',
      uf: site.fields?.uf || site.fields?.state || '',
      cidade: site.fields?.cidade || site.fields?.city || ''
    }
  };
}

function normalizeFieldName(value) {
  return String(value || '').trim();
}

function decodeSharePointFieldName(value) {
  return String(value || '')
    .replace(/_x([0-9A-Fa-f]{2,4})_/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)))
    .replace(/_x([0-9A-Fa-f]{2,4})/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)))
    .replace(/_x([0-9A-Fa-f]{2,4})_/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
}

function normalizeKeyForMatch(value) {
  const decoded = decodeSharePointFieldName(value);

  return decoded
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]/g, '')
    .toLowerCase();
}

function getValue(item, fieldName) {
  if (!fieldName) return '';
  const key = normalizeFieldName(fieldName);
  if (!key) return '';

  const source = item && item.fields ? item.fields : item;
  const rawKeys = Object.keys(source || {});
  const normalizedKey = decodeSharePointFieldName(key);

  const candidates = new Set([
    key,
    normalizedKey,
    key.replace(/_x0020_/gi, ' '),
    key.replace(/_x0020_/gi, ''),
    key.replace(/_/g, ' '),
    key.replace(/x0020/gi, ' '),
    key.replace(/x0020/gi, ''),
    key.toLowerCase(),
    key.toUpperCase(),
    key.replace(/\s+/g, ''),
    key.replace(/_/g, ''),
    key.replace(/_/g, '').toLowerCase(),
    key.replace(/\s+/g, '').toLowerCase(),
    normalizedKey.replace(/\s+/g, ''),
    normalizedKey.replace(/\s+/g, '').toLowerCase(),
    decodeSharePointFieldName(key).replace(/_/g, ' '),
    decodeSharePointFieldName(key).replace(/_/g, '')
  ]);

  for (const candidate of candidates) {
    const directValue = source[candidate];
    if (directValue !== undefined && directValue !== null) return directValue;
  }

  const matchedKey = rawKeys.find((candidate) => {
    const normalizedCandidate = normalizeKeyForMatch(candidate);
    const normalizedKey = normalizeKeyForMatch(key);
    return normalizedCandidate === normalizedKey || normalizedCandidate.includes(normalizedKey) || normalizedKey.includes(normalizedCandidate);
  });

  if (matchedKey) return source[matchedKey];

  return '';
}

function parseSharePointDate(value) {
  if (!value) return null;

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return null;

    const candidates = [
      trimmed.split('T')[0],
      trimmed.replace(/\//g, '-'),
      trimmed.replace(/\./g, '-'),
      trimmed
    ];

    for (const candidate of candidates) {
      const parsed = new Date(candidate);
      if (!Number.isNaN(parsed.getTime())) return parsed;
    }

    const dmyMatch = trimmed.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
    if (dmyMatch) {
      const [, day, month, year] = dmyMatch;
      const parsed = new Date(`${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`);
      if (!Number.isNaN(parsed.getTime())) return parsed;
    }
  }

  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) return parsed;

  return null;
}

function getListNameCandidates(listName) {
  const raw = String(listName || '').replace(/\s+/g, ' ').trim();
  if (!raw) return [];

  const withoutAccents = raw.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const spaced = withoutAccents.replace(/([a-z])([A-Z])/g, '$1 $2');
  const splitVariants = [
    raw,
    withoutAccents,
    spaced,
    spaced.replace(/[-_]/g, ' '),
    spaced.replace(/[-_]/g, ' ').replace(/\s+/g, ''),
    spaced.replace(/\s+/g, '-'),
    spaced.replace(/\s+/g, '_'),
    spaced.replace(/\s+/g, '')
  ];

  const candidates = new Set();

  splitVariants.forEach((value) => {
    if (!value) return;
    candidates.add(value);
    candidates.add(value.toLowerCase());
    candidates.add(value.toUpperCase());
    candidates.add(value.replace(/\s+/g, ''));
    candidates.add(value.replace(/\s+/g, '').toLowerCase());
    candidates.add(value.replace(/\s+/g, '').toUpperCase());
    candidates.add(value.replace(/[-_\s]+/g, ''));
    candidates.add(value.replace(/[-_\s]+/g, '').toLowerCase());
    candidates.add(value.replace(/[-_\s]+/g, ' ').trim());
    candidates.add(value.replace(/[-_\s]+/g, '-').trim());
    candidates.add(value.replace(/[-_\s]+/g, '_').trim());
  });

  return Array.from(candidates).filter(Boolean);
}

function safeString(value) {
  if (value === undefined || value === null) return '';
  return String(value).trim();
}

function isIgnorableProjetistaName(value) {
  const normalized = safeString(value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');

  return !normalized || [
    'nao informado',
    'nao informada',
    'nao cadastrado',
    'nao cadastra',
    'sem projetista',
    'sem informacao',
    'n/a',
    'na'
  ].includes(normalized);
}

function isDateLikeValue(value) {
  if (value === undefined || value === null || value === '') return false;

  if (value instanceof Date) {
    return !Number.isNaN(value.getTime());
  }

  if (typeof value === 'string') {
    const text = value.trim();
    if (!text) return false;
    return /\d{4}-\d{2}-\d{2}/.test(text) || /\d{2}\/\d{2}\/\d{4}/.test(text) || /\d{4}-\d{2}-\d{2}T/.test(text);
  }

  return false;
}

function normalizeUploadValue(value) {
  if (value === true || value === 1 || value === '1') return 1;
  if (value === false || value === 0 || value === '0') return 0;

  if (isDateLikeValue(value)) {
    return 1;
  }

  const numericValue = parseMetricValue(value);
  return numericValue > 0 ? 1 : 0;
}

function parseMetricValue(value) {
  if (value === true || value === 1 || value === '1') return 1;
  if (value === false || value === 0 || value === '0') return 0;

  if (Array.isArray(value)) {
    return value.reduce((sum, item) => sum + parseMetricValue(item), 0);
  }

  if (value && typeof value === 'object') {
    if ('Value' in value) return parseMetricValue(value.Value);
    if ('Label' in value) return parseMetricValue(value.Label);
    if ('label' in value) return parseMetricValue(value.label);
    if ('results' in value && Array.isArray(value.results)) {
      return value.results.reduce((sum, item) => sum + parseMetricValue(item), 0);
    }
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }

  const text = safeString(value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');

  if (!text) return 0;

  if (['sim', 'yes', 'true', 'ok', 'uploaded', 'upload', 'concluido', 'concluida', 's', 'nao', 'no', 'n'].includes(text)) {
    return text === 'nao' || text === 'no' || text === 'n' ? 0 : 1;
  }

  const normalized = text.replace(/[%.,]/g, '').replace(/\s+/g, '');
  const numeric = Number(normalized);
  return Number.isFinite(numeric) ? numeric : 0;
}

async function readJsonResponse(response) {
  const text = await response.text();

  if (!text) {
    return {};
  }

  try {
    return JSON.parse(text);
  } catch {
    return {
      rawText: text,
      status: response.status,
      statusText: response.statusText,
      contentType: response.headers.get('content-type') || ''
    };
  }
}

async function getGraphToken() {
  if (graphTokenCache && graphTokenCache.expiresAt > Date.now()) return graphTokenCache.accessToken;
  if (graphTokenPromise) return graphTokenPromise;

  graphTokenPromise = (async () => {
  const headers = {
    'Content-Type': 'application/x-www-form-urlencoded'
  };

  const body = new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });

  const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
  const response = await fetch(tokenUrl, { method: 'POST', headers, body });
  const data = await response.json();

  if (!response.ok) {
    throw new Error(`Falha ao obter token do Microsoft Graph: ${data.error_description || JSON.stringify(data)}`);
  }

  graphTokenCache = {
    accessToken: data.access_token,
    expiresAt: Date.now() + Math.max(60, Number(data.expires_in || 3600) - 120) * 1000
  };
  cacheMetrics.tokenRefreshes += 1;
  console.log(JSON.stringify({ event: 'graph_token_refreshed', expiresAt: new Date(graphTokenCache.expiresAt).toISOString() }));
  return graphTokenCache.accessToken;
  })();

  try {
    return await graphTokenPromise;
  } finally {
    graphTokenPromise = null;
  }
}

async function getSharePointToken(siteUrl) {
  const siteOrigin = new URL(siteUrl).origin;
  const body = new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    scope: `${siteOrigin}/.default`,
    grant_type: 'client_credentials'
  });
  const response = await fetch(`https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });
  const data = await response.json();
  if (!response.ok) throw new Error(`Falha ao obter token do SharePoint: ${String(data.error_description || data.error || 'sem detalhe').split('\r\n')[0]}`);
  return data.access_token;
}

function escapeODataString(value) {
  return String(value).replace(/'/g, "''");
}

// Monta mensagem de erro com etapa e status HTTP, sem expor tokens.
function describeSharePointError(step, response, payload) {
  const detail = payload?.error?.message?.value || payload?.error?.message || payload?.error_description || payload?.rawText || '';
  const hint = response.status === 401 || response.status === 403
    ? ' (o SharePoint REST costuma recusar tokens app-only baseados em client secret; use certificado ou conceda acesso à lista)'
    : '';
  return `${step}: HTTP ${response.status} ${String(detail).slice(0, 200)}${hint}`.trim();
}

async function publishSummaryAttachment(pngBuffer) {
  const { siteUrl, listName, title, fileName } = summaryAttachmentTarget;
  const token = await getSharePointToken(siteUrl);
  const listPath = `/_api/web/lists/getbytitle('${escapeODataString(listName)}')`;
  const headers = { Accept: 'application/json;odata=nometadata', Authorization: `Bearer ${token}` };
  const itemLookup = await fetch(`${siteUrl}${listPath}/items?$select=Id&$filter=Title eq '${escapeODataString(title)}'&$top=1`, { headers });
  const lookupData = await readJsonResponse(itemLookup);
  if (!itemLookup.ok) throw new Error(describeSharePointError(`Localizar lista ${listName}`, itemLookup, lookupData));

  let itemId = Array.isArray(lookupData.value) && lookupData.value[0]?.Id;
  if (!itemId) {
    const createResponse = await fetch(`${siteUrl}${listPath}/items`, {
      method: 'POST',
      headers: { ...headers, 'Content-Type': 'application/json;odata=nometadata' },
      body: JSON.stringify({ Title: title })
    });
    const createdItem = await readJsonResponse(createResponse);
    if (!createResponse.ok || !createdItem.Id) throw new Error(describeSharePointError(`Criar item ${title}`, createResponse, createdItem));
    itemId = createdItem.Id;
  }

  const attachmentPath = `${listPath}/items(${itemId})/AttachmentFiles/getbyfilename('${escapeODataString(fileName)}')`;
  const deleteResponse = await fetch(`${siteUrl}${attachmentPath}`, {
    method: 'POST',
    headers: { ...headers, 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' }
  });
  if (!deleteResponse.ok && deleteResponse.status !== 404) throw new Error(describeSharePointError('Remover anexo anterior', deleteResponse, await readJsonResponse(deleteResponse)));

  const uploadResponse = await fetch(`${siteUrl}${listPath}/items(${itemId})/AttachmentFiles/add(FileName='${escapeODataString(fileName)}')`, {
    method: 'POST',
    headers: { ...headers, 'Content-Type': 'image/png' },
    body: pngBuffer
  });
  if (!uploadResponse.ok) throw new Error(describeSharePointError('Enviar anexo', uploadResponse, await readJsonResponse(uploadResponse)));
  return itemId;
}

function buildLocationKey(uf, cidade) {
  return `${safeString(uf)}|${safeString(cidade)}`
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9|]/g, '')
    .toUpperCase();
}

async function fetchWeightMap(weightSource, options = {}) {
  if (!weightSource?.url || !weightSource?.listName || !weightSource?.fields) return new Map();

  const cacheKey = JSON.stringify(weightSource);
  const cached = weightSourceCache.get(cacheKey);
  if (!options.forceRefresh && cached && cached.expiresAt > Date.now()) return cached.weights;
  if (weightLoadPromises.has(cacheKey)) {
    cacheMetrics.weightLoadDeduplications += 1;
    return weightLoadPromises.get(cacheKey);
  }

  const loadPromise = (async () => {
    const sourceUrl = new URL(weightSource.url);
    const fields = [weightSource.fields.uf, weightSource.fields.cidade, weightSource.fields.peso].filter(Boolean);
    const token = await getGraphToken();
    const weights = new Map();
    const listCandidates = getListNameCandidates(weightSource.listName);

    for (const listName of listCandidates) {
      const baseUrl = `https://graph.microsoft.com/v1.0/sites/${sourceUrl.hostname}:/${sourceUrl.pathname.replace(/\/$/, '')}:/lists/${encodeURIComponent(listName)}/items?$expand=fields($select=${fields.map((field) => encodeURIComponent(field)).join(',')})&$top=1000`;
      let nextUrl = baseUrl;
      let loadedAny = false;

      while (nextUrl) {
        const response = await fetch(nextUrl, {
          headers: { Accept: 'application/json;odata.metadata=none', Authorization: `Bearer ${token}` }
        });
        const payload = await readJsonResponse(response);

        if (!response.ok) {
          if ([400, 401, 403, 404].includes(response.status)) break;
          throw new Error(`Erro ao consultar a lista de pesos ${listName}.`);
        }

        loadedAny = true;
        for (const item of Array.isArray(payload.value) ? payload.value : []) {
          const fieldData = item.fields || {};
          const key = buildLocationKey(getValue(fieldData, weightSource.fields.uf), getValue(fieldData, weightSource.fields.cidade));
          const peso = parseMetricValue(getValue(fieldData, weightSource.fields.peso));
          if (key !== '|' && Number.isFinite(peso) && peso > 0) weights.set(key, peso);
        }
        nextUrl = payload['@odata.nextLink'] || null;
      }

      if (loadedAny) break;
    }

    weightSourceCache.set(cacheKey, { weights, expiresAt: Date.now() + cacheTtlMs });
    console.log(JSON.stringify({ event: 'weight_source_loaded', list: weightSource.listName, locations: weights.size }));
    return weights;
  })();

  weightLoadPromises.set(cacheKey, loadPromise);
  try {
    return await loadPromise;
  } finally {
    weightLoadPromises.delete(cacheKey);
  }
}

function addToIndex(index, key, itemIndex) {
  if (!key) return;
  const values = index.get(key);
  if (values) values.push(itemIndex);
  else index.set(key, [itemIndex]);
}

function getConfiguredFieldsForSource(site) {
  const sourceCacheKey = buildPhysicalSourceCacheKey(site);
  const fields = new Set();
  const addSelectionFields = (selection) => {
    const normalizedSelection = normalizeSiteConfig(selection);
    if (!normalizedSelection?.url || !normalizedSelection?.listName) return;
    if (buildPhysicalSourceCacheKey(normalizedSelection) !== sourceCacheKey) return;
    for (const fieldName of Object.values(normalizedSelection.fields || {})) {
      if (fieldName) fields.add(fieldName);
    }
  };

  addSelectionFields(site);
  for (const selection of getConfiguredSelections()) addSelectionFields(selection);
  return Array.from(fields).sort((left, right) => left.localeCompare(right));
}

function createRawSourceSnapshot(items, activeListName, cacheKey, fields) {
  return {
    cacheKey,
    activeListName,
    fields,
    items,
    expiresAt: Date.now() + cacheTtlMs,
    createdAt: Date.now(),
    lastAccessedAt: Date.now()
  };
}

function createSourceProfileSnapshot(items, activeListName, cacheKey, sourceCacheKey, expiresAt, sourceCreatedAt) {
  const indexes = {
    projetista: new Map(),
    uf: new Map(),
    cidade: new Map(),
    ano: new Map(),
    mes: new Map(),
    anoMes: new Map()
  };

  for (let index = 0; index < items.length; index += 1) {
    const item = items[index];
    addToIndex(indexes.projetista, item.projetista.toLowerCase(), index);
    addToIndex(indexes.uf, item.uf, index);
    addToIndex(indexes.cidade, item.cidade.toLowerCase(), index);
    if (item.date && !Number.isNaN(item.date.getTime())) {
      const year = item.date.getFullYear();
      const month = item.date.getMonth() + 1;
      addToIndex(indexes.ano, String(year), index);
      addToIndex(indexes.mes, String(month), index);
      addToIndex(indexes.anoMes, `${year}-${month}`, index);
    }
  }

  return {
    cacheKey,
    sourceCacheKey,
    sourceCreatedAt,
    activeListName,
    items,
    indexes,
    expiresAt,
    createdAt: Date.now(),
    lastAccessedAt: Date.now()
  };
}

function serializeSourceSnapshot(snapshot) {
  return {
    version: 2,
    activeListName: snapshot.activeListName,
    fields: snapshot.fields,
    expiresAt: snapshot.expiresAt,
    createdAt: snapshot.createdAt,
    items: snapshot.items
  };
}

function hydrateSourceSnapshot(data, cacheKey) {
  if (!data || data.version !== 2 || !Array.isArray(data.items) || data.expiresAt <= Date.now()) return null;
  const snapshot = createRawSourceSnapshot(data.items, data.activeListName, cacheKey, data.fields || []);
  snapshot.expiresAt = data.expiresAt;
  snapshot.createdAt = data.createdAt || Date.now();
  return snapshot;
}

function loadPersistedSourceSnapshot(cacheKey) {
  try {
    const compressed = fs.readFileSync(getCacheFilePath(cacheKey));
    return hydrateSourceSnapshot(JSON.parse(zlib.gunzipSync(compressed).toString('utf8')), cacheKey);
  } catch {
    return null;
  }
}

function persistSourceSnapshot(snapshot) {
  let serialized;
  try {
    serialized = JSON.stringify(serializeSourceSnapshot(snapshot));
  } catch (error) {
    console.warn(JSON.stringify({ event: 'source_cache_serialize_error', message: error.message }));
    return;
  }

  zlib.gzip(serialized, (compressionError, compressed) => {
    if (compressionError) {
      console.warn(JSON.stringify({ event: 'source_cache_persist_error', message: compressionError.message }));
      return;
    }
    fs.writeFile(getCacheFilePath(snapshot.cacheKey), compressed, (writeError) => {
      if (writeError) console.warn(JSON.stringify({ event: 'source_cache_persist_error', message: writeError.message }));
    });
  });
}

function getIndexedItems(snapshot, filters = {}) {
  const candidates = [];
  const pushIndex = (index, key) => {
    if (!key) return;
    candidates.push(index.get(key) || []);
  };

  pushIndex(snapshot.indexes.projetista, filters.projetista && String(filters.projetista).toLowerCase());
  pushIndex(snapshot.indexes.uf, filters.uf && String(filters.uf).toUpperCase());
  pushIndex(snapshot.indexes.cidade, filters.cidade && String(filters.cidade).toLowerCase());
  if (filters.ano && filters.mes) pushIndex(snapshot.indexes.anoMes, `${filters.ano}-${Number(filters.mes)}`);
  else if (filters.ano) pushIndex(snapshot.indexes.ano, String(filters.ano));
  else if (filters.mes) pushIndex(snapshot.indexes.mes, String(Number(filters.mes)));

  if (!candidates.length) return snapshot.items;
  candidates.sort((left, right) => left.length - right.length);
  const base = candidates[0];
  if (!base.length) return [];
  const candidateSets = candidates.slice(1).map((values) => new Set(values));
  const result = [];
  for (const itemIndex of base) {
    let matches = true;
    for (const candidateSet of candidateSets) {
      if (!candidateSet.has(itemIndex)) {
        matches = false;
        break;
      }
    }
    if (matches) result.push(snapshot.items[itemIndex]);
  }
  return result;
}

function invalidateDerivedCachesForSource(sourceCacheKey) {
  removeProfilesForSource(sourceCacheKey);
  for (const [key, entry] of dashboardCache) {
    if (entry.sourceCacheKey === sourceCacheKey) dashboardCache.delete(key);
  }
  for (const [key, entry] of analyticsCache) {
    if (entry.sourceCacheKey === sourceCacheKey) analyticsCache.delete(key);
  }
}

async function loadRawSourceSnapshot(site, options = {}) {
  const normalizedSite = normalizeSiteConfig(site);
  if (!normalizedSite || !normalizedSite.url || !normalizedSite.listName) {
    console.error('[SharePoint] Lista não configurada para a seleção:', site);
    return createRawSourceSnapshot([], '', 'invalid', []);
  }

  const sourceCacheKey = buildPhysicalSourceCacheKey(normalizedSite);
  const cached = sourceListCache.get(sourceCacheKey);
  if (!options.forceRefresh && cached && cached.expiresAt > Date.now()) {
    cacheMetrics.hits += 1;
    return touchCacheEntry(sourceListCache, sourceCacheKey, cached);
  }

  if (!options.forceRefresh) {
    const persisted = loadPersistedSourceSnapshot(sourceCacheKey);
    if (persisted) {
      cacheMetrics.hits += 1;
      sourceListCache.set(sourceCacheKey, persisted);
      evictSourceCache();
      console.log(JSON.stringify({ event: 'source_cache_disk_hit', list: persisted.activeListName, items: persisted.items.length }));
      return persisted;
    }
  }

  if (sourceLoadPromises.has(sourceCacheKey)) {
    cacheMetrics.sourceLoadDeduplications += 1;
    return sourceLoadPromises.get(sourceCacheKey);
  }

  const fields = getConfiguredFieldsForSource(normalizedSite);
  if (!fields.length) throw new Error(`Nenhum campo configurado para a lista ${normalizedSite.listName}.`);
  cacheMetrics.misses += 1;

  const loadPromise = (async () => {
    const graphStartedAt = Date.now();
    const siteUrl = new URL(normalizedSite.url);
    const siteHostname = siteUrl.hostname;
    const sitePath = siteUrl.pathname.replace(/\/$/, '');
    const listCandidates = getListNameCandidates(normalizedSite.listName);
    const token = await getGraphToken();
    let activeListName = normalizedSite.listName;
    let compactItems = [];
    let loadedAnyCandidate = false;

    for (const listName of listCandidates) {
      const baseUrl = `https://graph.microsoft.com/v1.0/sites/${siteHostname}:/${sitePath}:/lists/${encodeURIComponent(listName)}/items?$expand=fields($select=${fields.map((field) => encodeURIComponent(field)).join(',')})&$top=1000`;
      let nextUrl = baseUrl;
      let pageNumber = 0;
      let candidateLoaded = false;
      const candidateItems = [];

      while (nextUrl) {
        const response = await fetch(nextUrl, {
          headers: {
            Accept: 'application/json;odata.metadata=none',
            Authorization: `Bearer ${token}`
          }
        });
        const payload = await readJsonResponse(response);
        pageNumber += 1;

        if (!response.ok) {
          const rawText = payload?.rawText || payload?.error?.message || JSON.stringify(payload);
          const detail = rawText && String(rawText).length > 0 ? String(rawText).slice(0, 400) : 'sem mensagem';
          if ([400, 401, 403, 404].includes(response.status)) {
            console.warn(`Lista não encontrada com nome alternativo: ${listName}. Detalhe: ${detail}`);
            break;
          }
          throw new Error(`Erro ao consultar lista ${listName}: ${detail}`);
        }

        candidateLoaded = true;
        cacheMetrics.graphFetches += 1;
        for (const graphItem of Array.isArray(payload.value) ? payload.value : []) {
          const sourceFields = graphItem.fields || {};
          const compactItem = {};
          for (const fieldName of fields) {
            const value = getValue(sourceFields, fieldName);
            if (value !== undefined && value !== null) compactItem[fieldName] = value;
          }
          candidateItems.push(compactItem);
        }
        console.log(`[SharePoint] list=${listName} page=${pageNumber} records=${candidateItems.length} status=${response.status}`);
        nextUrl = payload['@odata.nextLink'] || null;
      }

      if (candidateLoaded) {
        activeListName = listName;
        compactItems = candidateItems;
        loadedAnyCandidate = true;
        break;
      }
    }

    if (!loadedAnyCandidate) {
      console.error('[SharePoint] Nenhuma lista foi encontrada para a configuração atual.', {
        site: normalizedSite.name,
        url: normalizedSite.url,
        listName: normalizedSite.listName,
        candidates: listCandidates,
        attemptedFields: fields
      });
    }

    const snapshot = createRawSourceSnapshot(compactItems, activeListName, sourceCacheKey, fields);
    sourceListCache.set(sourceCacheKey, snapshot);
    invalidateDerivedCachesForSource(sourceCacheKey);
    evictSourceCache();
    persistSourceSnapshot(snapshot);
    console.log(JSON.stringify({
      event: 'source_cache_saved',
      list: activeListName,
      items: compactItems.length,
      fields: fields.length,
      graphFetchMs: Date.now() - graphStartedAt,
      ttlMs: cacheTtlMs
    }));
    return snapshot;
  })();

  sourceLoadPromises.set(sourceCacheKey, loadPromise);
  try {
    return await loadPromise;
  } finally {
    sourceLoadPromises.delete(sourceCacheKey);
  }
}

function createProfileItems(rawItems, site, weightMap) {
  const items = new Array(rawItems.length);
  for (let index = 0; index < rawItems.length; index += 1) {
    const sourceFields = rawItems[index];
    const dataConclusao = getValue(sourceFields, site.fields.dataConclusao);
    const date = parseSharePointDate(dataConclusao);
    const hasUploadVisiumField = Boolean(site.fields.uploadVisium);
    const uploadVisiumRaw = hasUploadVisiumField ? getValue(sourceFields, site.fields.uploadVisium) : '';
    const uf = safeString(getValue(sourceFields, site.fields.uf)).toUpperCase();
    const cidade = safeString(getValue(sourceFields, site.fields.cidade));
    items[index] = {
      date,
      uploadVisium: hasUploadVisiumField ? normalizeUploadValue(uploadVisiumRaw) : 0,
      uf,
      cidade,
      projetista: safeString(getValue(sourceFields, site.fields.projetista)),
      peso: weightMap.get(buildLocationKey(uf, cidade)) || 1
    };
  }
  return items;
}

async function loadSourceSnapshot(site, options = {}) {
  const normalizedSite = normalizeSiteConfig(site);
  if (!normalizedSite || !normalizedSite.url || !normalizedSite.listName) {
    return createSourceProfileSnapshot([], '', 'invalid', 'invalid', Date.now() + cacheTtlMs, Date.now());
  }

  const profileCacheKey = buildSourceProfileCacheKey(normalizedSite);
  const cached = sourceProfileCache.get(profileCacheKey);
  if (!options.forceRefresh && cached && cached.expiresAt > Date.now()) {
    cacheMetrics.hits += 1;
    return touchCacheEntry(sourceProfileCache, profileCacheKey, cached);
  }

  if (sourceProfilePromises.has(profileCacheKey)) {
    cacheMetrics.profileLoadDeduplications += 1;
    return sourceProfilePromises.get(profileCacheKey);
  }

  const loadPromise = (async () => {
    const [rawSnapshot, weightMap] = await Promise.all([
      loadRawSourceSnapshot(normalizedSite, options),
      fetchWeightMap(normalizedSite.weightSource, options)
    ]);
    const items = createProfileItems(rawSnapshot.items, normalizedSite, weightMap);
    const snapshot = createSourceProfileSnapshot(
      items,
      rawSnapshot.activeListName,
      profileCacheKey,
      rawSnapshot.cacheKey,
      rawSnapshot.expiresAt,
      rawSnapshot.createdAt
    );
    sourceProfileCache.set(profileCacheKey, snapshot);
    evictSourceCache();
    console.log(JSON.stringify({ event: 'source_profile_saved', list: snapshot.activeListName, items: items.length }));
    return snapshot;
  })();

  sourceProfilePromises.set(profileCacheKey, loadPromise);
  try {
    return await loadPromise;
  } finally {
    sourceProfilePromises.delete(profileCacheKey);
  }
}

async function fetchListItems(site, filters = {}, options = {}) {
  const snapshot = await loadSourceSnapshot(site, options);
  return getIndexedItems(snapshot, filters);
}

function buildResult(items, options = {}) {
  const includeUploadVisium = Boolean(options.includeUploadVisium);
  const byProjetista = {};
  const byProjetistaUpload = {};
  const byMonth = Array.from({ length: 12 }, (_, idx) => ({ month: idx + 1, label: monthNames[idx], total: 0 }));
  const years = new Set();
  const ufs = new Set();
  const cidades = new Set();
  const projetistas = new Set();
  let total = 0;
  let doneCount = 0;
  let uploadVisiumTotal = 0;

  for (const item of items) {
    const nome = safeString(item.projetista).trim();
    const uploadFlag = Number(item.uploadVisium || 0) > 0 ? 1 : 0;
    const peso = Number(item.peso || 1);
    total += peso;
    if (includeUploadVisium && uploadFlag) uploadVisiumTotal += 1;

    if (!isIgnorableProjetistaName(nome)) {
      byProjetista[nome] = (byProjetista[nome] || 0) + peso;

      if (includeUploadVisium) {
        byProjetistaUpload[nome] = (byProjetistaUpload[nome] || 0) + uploadFlag;
      }

      projetistas.add(nome);
    }

    if (item.date) {
      const monthIndex = item.date.getMonth();
      byMonth[monthIndex].total += peso;
      years.add(item.date.getFullYear());
      doneCount += peso;
    }

    if (item.uf) ufs.add(item.uf);
    if (item.cidade) cidades.add(item.cidade);
  }

  const barData = Object.entries(byProjetista)
    .map(([label, total]) => ({ label, total, upload: includeUploadVisium ? (byProjetistaUpload[label] || 0) : 0 }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 15);
  const yearsList = Array.from(years).sort((a, b) => a - b);

  const summary = {
    total,
    done: doneCount,
    uploadVisiumTotal,
    uniqueProjetistas: Object.keys(byProjetista).length,
    anoInicial: yearsList[0] || new Date().getFullYear(),
    anoFinal: yearsList[yearsList.length - 1] || new Date().getFullYear()
  };

  return {
    summary,
    barData,
    monthlyTrend: byMonth.map((m) => ({ label: m.label, total: m.total })),
    years: yearsList,
    ufs: Array.from(ufs).sort(),
    cidades: Array.from(cidades).sort((a, b) => a.localeCompare(b)),
    projetistas: Array.from(projetistas).sort((a, b) => a.localeCompare(b)),
    includeUploadVisium
  };
}

// Data de referência dos períodos: hoje, ou o fim do mês/ano filtrado quando não for o atual.
// A maior data da lista não é usada, pois pode conter datas futuras.
function resolveAnalyticsReferenceDate(filters = {}) {
  const now = new Date();
  const year = Number(filters.ano) || now.getFullYear();
  const month = Number(filters.mes) || (year === now.getFullYear() ? now.getMonth() + 1 : 12);
  if (year === now.getFullYear() && month === now.getMonth() + 1) return now;
  return new Date(year, month, 0);
}

function buildExecutiveAnalytics(items, comparisonItems = items, filters = {}) {
  const referenceDate = resolveAnalyticsReferenceDate(filters);
  const designers = new Map();
  const cities = new Map();
  let topCity = null;

  for (const item of items) {
    if (!item.date || Number.isNaN(item.date.getTime())) continue;

    const peso = Number(item.peso || 1);
    const designer = safeString(item.projetista);
    const city = safeString(item.cidade);
    if (!isIgnorableProjetistaName(designer)) designers.set(designer, (designers.get(designer) || 0) + peso);
    if (city) {
      const cityTotal = (cities.get(city) || 0) + peso;
      cities.set(city, cityTotal);
      if (!topCity || cityTotal > topCity.total) topCity = { name: city, total: cityTotal };
    }
  }

  const latestDate = referenceDate;
  const dayStart = new Date(latestDate.getFullYear(), latestDate.getMonth(), latestDate.getDate());
  const previousDayStart = new Date(dayStart);
  previousDayStart.setDate(previousDayStart.getDate() - 1);
  const weekStart = new Date(dayStart);
  weekStart.setDate(weekStart.getDate() - 6);
  const monthStart = new Date(dayStart.getFullYear(), dayStart.getMonth(), 1);
  const previousMonthStart = new Date(dayStart.getFullYear(), dayStart.getMonth() - 1, 1);
  let dayTotal = 0;
  let previousDayTotal = 0;
  let weekTotal = 0;
  let monthTotal = 0;
  let previousMonthTotal = 0;
  const dayEnd = new Date(dayStart.getTime() + 86400000);

  for (const item of comparisonItems) {
    if (!item.date || Number.isNaN(item.date.getTime())) continue;
    const peso = Number(item.peso || 1);
    if (item.date >= dayStart && item.date < dayEnd) dayTotal += peso;
    if (item.date >= previousDayStart && item.date < dayStart) previousDayTotal += peso;
    if (item.date >= weekStart && item.date < dayEnd) weekTotal += peso;
    if (item.date >= monthStart && item.date < dayEnd) monthTotal += peso;
    if (item.date >= previousMonthStart && item.date < monthStart) previousMonthTotal += peso;
  }

  const elapsedDays = Math.max(1, Math.floor((dayStart - monthStart) / 86400000) + 1);
  const daysInMonth = new Date(dayStart.getFullYear(), dayStart.getMonth() + 1, 0).getDate();

  return {
    generatedAt: new Date().toISOString(),
    latestDate: latestDate.toISOString(),
    period: {
      dayTotal,
      previousDayTotal,
      weekTotal,
      monthTotal,
      previousMonthTotal,
      projectedMonthTotal: Math.round((monthTotal / elapsedDays) * daysInMonth)
    },
    activeDesigners: designers.size,
    ranking: Array.from(designers.entries())
      .map(([name, total]) => ({ name, total }))
      .sort((left, right) => right.total - left.total)
      .slice(0, 10),
    topCity
  };
}

app.use(express.static(path.join(__dirname, 'public')));

app.get('/health', (_, res) => {
  res.json({ ok: true, service: 'sharepoint-dashboard', sites: config.sites?.length || 0 });
});

app.get('/api/config', (_, res) => {
  res.json(getPublicConfig(config));
});

app.get('/api/cache/stats', (_, res) => {
  const totalRequests = cacheMetrics.hits + cacheMetrics.misses;
  res.json({
    dashboardCacheEntries: dashboardCache.size,
    sourceCacheEntries: sourceListCache.size,
    sourceProfileCacheEntries: sourceProfileCache.size,
    weightCacheEntries: weightSourceCache.size,
    analyticsCacheEntries: analyticsCache.size,
    sourceCachedItems: totalSourceItems(),
    sourceProfileCachedItems: totalProfileItems(),
    totalCachedItems: totalCachedItems(),
    cacheHits: cacheMetrics.hits,
    cacheMisses: cacheMetrics.misses,
    hitRate: `${totalRequests ? ((cacheMetrics.hits / totalRequests) * 100).toFixed(2) : '0.00'}%`,
    graphFetches: cacheMetrics.graphFetches,
    tokenRefreshes: cacheMetrics.tokenRefreshes,
    activeSourceLoads: sourceLoadPromises.size,
    activeProfileLoads: sourceProfilePromises.size,
    activeWeightLoads: weightLoadPromises.size,
    sourceLoadDeduplications: cacheMetrics.sourceLoadDeduplications,
    profileLoadDeduplications: cacheMetrics.profileLoadDeduplications,
    weightLoadDeduplications: cacheMetrics.weightLoadDeduplications
  });
});

app.post('/api/cache/clear', (req, res) => {
  dashboardCache.clear();
  sourceListCache.clear();
  sourceProfileCache.clear();
  weightSourceCache.clear();
  analyticsCache.clear();
  if (req.query.persistent === 'true') {
    for (const file of fs.readdirSync(cacheDirectory)) {
      if (file.startsWith('source-') && file.endsWith('.json.gz')) fs.unlinkSync(path.join(cacheDirectory, file));
    }
  }
  res.json({ ok: true, persistentCleared: req.query.persistent === 'true' });
});

app.post('/api/cache/refresh', async (req, res) => {
  try {
    const selectedSite = resolveSelectedSite(req.body?.site || config.sites[0]?.name, req.body?.activity || '', req.body?.subactivity || '');
    const snapshot = await loadSourceSnapshot(selectedSite, { forceRefresh: true });
    res.json({ ok: true, list: snapshot.activeListName, records: snapshot.items.length });
  } catch (error) {
    res.status(502).json({ message: 'Não foi possível atualizar o cache.', detail: error.message });
  }
});

app.post('/api/reports/summary-snapshot', async (req, res) => {
  try {
    const pngDataUrl = String(req.body?.pngDataUrl || '');
    const match = pngDataUrl.match(/^data:image\/png;base64,([A-Za-z0-9+/=]+)$/);
    if (!match) return res.status(400).json({ message: 'O resumo precisa ser enviado como imagem PNG.' });

    const pngBuffer = Buffer.from(match[1], 'base64');
    if (!pngBuffer.length || pngBuffer.length > 6 * 1024 * 1024) {
      return res.status(400).json({ message: 'O PNG do resumo está vazio ou excede 6 MB.' });
    }

    const itemId = await publishSummaryAttachment(pngBuffer);
    console.log(JSON.stringify({ event: 'summary_snapshot_published', requestId: req.requestId, itemId }));
    res.status(201).json({ ok: true, itemId });
  } catch (error) {
    console.error(JSON.stringify({ event: 'summary_snapshot_error', requestId: req.requestId, message: error.message }));
    res.status(502).json({ message: 'Não foi possível publicar o resumo no SharePoint.', detail: error.message, requestId: req.requestId });
  }
});

function paginateDashboardPayload(payload, query) {
  if (!query.page && !query.pageSize) return payload;
  const page = Math.max(1, Number(query.page) || 1);
  const pageSize = Math.min(1000, Math.max(1, Number(query.pageSize) || 100));
  const totalRecords = payload.items.length;
  const totalPages = Math.max(1, Math.ceil(totalRecords / pageSize));
  const start = (Math.min(page, totalPages) - 1) * pageSize;
  return {
    ...payload,
    page,
    pageSize,
    totalRecords,
    totalPages,
    items: payload.items.slice(start, start + pageSize)
  };
}

function getConfiguredSelections() {
  const selections = [];
  for (const site of config.sites || []) {
    const activities = Array.isArray(site.activities) && site.activities.length ? site.activities : [site];
    for (const activity of activities) {
      const subactivities = Array.isArray(activity.activities) && activity.activities.length ? activity.activities : [null];
      for (const subactivity of subactivities) {
        selections.push(resolveSelectedSite(
          site.name || site.siteName,
          activity.name || activity.activityName || '',
          subactivity?.name || subactivity?.activityName || ''
        ));
      }
    }
  }
  return selections.filter(Boolean);
}

async function warmupCaches() {
  const selections = getConfiguredSelections();
  const byPhysicalSource = new Map();
  for (const selection of selections) {
    const sourceCacheKey = buildPhysicalSourceCacheKey(selection);
    const group = byPhysicalSource.get(sourceCacheKey) || [];
    group.push(selection);
    byPhysicalSource.set(sourceCacheKey, group);
  }

  console.log(JSON.stringify({
    event: 'cache_warmup_started',
    selections: selections.length,
    physicalSources: byPhysicalSource.size
  }));

  for (const [sourceCacheKey, sourceSelections] of byPhysicalSource) {
    const primarySelection = sourceSelections[0];
    try {
      await loadRawSourceSnapshot(primarySelection);
      for (const selection of sourceSelections) {
        const snapshot = await loadSourceSnapshot(selection);
        const filters = { projetista: '', uf: '', cidade: '', mes: '', ano: '' };
        const analyticsKey = `analytics:${buildDashboardCacheKey(selection.name, selection.activityName, selection.subActivityName, filters)}`;
        analyticsCache.set(analyticsKey, {
          value: buildExecutiveAnalytics(snapshot.items, snapshot.items),
          sourceCacheKey,
          expiresAt: snapshot.expiresAt,
          lastAccessedAt: Date.now()
        });
      }
    } catch (error) {
      console.warn(JSON.stringify({ event: 'cache_warmup_error', list: primarySelection.listName, message: error.message }));
    }
  }
  console.log(JSON.stringify({
    event: 'cache_warmup_finished',
    physicalSources: sourceListCache.size,
    sourceItems: totalSourceItems(),
    profileItems: totalProfileItems()
  }));
}

app.get('/api/dashboard', validateDashboardQuery, async (req, res) => {
  try {
    const requestStartedAt = Date.now();
    const siteName = req.query.site || config.sites[0]?.name || config.sites[0]?.siteName;
    const activityName = req.query.activity || '';
    const subActivityName = req.query.subactivity || '';
    const selectedSite = resolveSelectedSite(siteName, activityName, subActivityName);
    const filters = {
      projetista: req.query.projetista || '',
      uf: req.query.uf || '',
      cidade: req.query.cidade || '',
      mes: req.query.mes || '',
      ano: req.query.ano || ''
    };

    const cacheKey = buildDashboardCacheKey(siteName, activityName, subActivityName, filters);
    const shouldRefresh = String(req.query.refresh || '').toLowerCase() === 'true' || String(req.query.refresh || '') === '1';

    console.log(`[dashboard] site=${siteName} activity=${activityName} subactivity=${subActivityName} refresh=${shouldRefresh} filters=${JSON.stringify(filters)}`);

    const cached = dashboardCache.get(cacheKey);
    if (!shouldRefresh && cached && cached.expiresAt > Date.now()) {
      cacheMetrics.hits += 1;
      console.log(JSON.stringify({ event: 'cache_hit', requestId: req.requestId, key: cacheKey }));
      const paginationStartedAt = Date.now();
      const responsePayload = paginateDashboardPayload(touchCacheEntry(dashboardCache, cacheKey, cached).payload, req.query);
      console.log(JSON.stringify({
        event: 'performance',
        requestId: req.requestId,
        cacheMs: Date.now() - requestStartedAt,
        graphFetchMs: 0,
        filterMs: 0,
        analyticsMs: 0,
        aggregationMs: 0,
        paginationMs: Date.now() - paginationStartedAt,
        totalMs: Date.now() - requestStartedAt
      }));
      return res.json(responsePayload);
    }

    if (cached) dashboardCache.delete(cacheKey);
    cacheMetrics.misses += 1;

    const sourceStartedAt = Date.now();
    const snapshot = await loadSourceSnapshot(selectedSite, { forceRefresh: shouldRefresh });
    const sourceMs = Date.now() - sourceStartedAt;
    const filteringStartedAt = Date.now();
    const items = getIndexedItems(snapshot, filters);
    const comparisonFilters = { ...filters, mes: '', ano: '' };
    const comparisonItems = getIndexedItems(snapshot, comparisonFilters);
    const filterMs = Date.now() - filteringStartedAt;
    const analyticsKey = `analytics:${cacheKey}`;
    const cachedAnalytics = !shouldRefresh ? analyticsCache.get(analyticsKey) : null;
    let analytics = cachedAnalytics && cachedAnalytics.expiresAt > Date.now()
      ? touchCacheEntry(analyticsCache, analyticsKey, cachedAnalytics).value
      : null;
    const analyticsStartedAt = Date.now();
    if (!analytics) {
      analytics = buildExecutiveAnalytics(items, comparisonItems, filters);
      analyticsCache.set(analyticsKey, {
        value: analytics,
        sourceCacheKey: snapshot.sourceCacheKey,
        expiresAt: Date.now() + cacheTtlMs,
        lastAccessedAt: Date.now()
      });
    }
    const analyticsMs = Date.now() - analyticsStartedAt;
    console.log(`[dashboard] items received=${items.length} for ${siteName}/${activityName}`);
    const aggregationStartedAt = Date.now();
    const result = buildResult(items, { includeUploadVisium: Boolean(selectedSite?.fields?.uploadVisium) });
    const payload = {
      site: selectedSite,
      filters,
      items: items.map((item) => ({
        projetista: item.projetista || '',
        uf: item.uf || '',
        cidade: item.cidade || '',
        uploadVisium: Number(item.uploadVisium || 0),
          peso: Number(item.peso || 1),
        date: item.date ? item.date.toISOString() : null
      })),
      ...result,
      analytics
    };
    const aggregationMs = Date.now() - aggregationStartedAt;

    dashboardCache.set(cacheKey, {
      payload,
      sourceCacheKey: snapshot.sourceCacheKey,
      expiresAt: Date.now() + cacheTtlMs,
      lastAccessedAt: Date.now()
    });
    while (dashboardCache.size > dashboardCacheMaxEntries) dashboardCache.delete(dashboardCache.keys().next().value);
    const paginationStartedAt = Date.now();
    const responsePayload = paginateDashboardPayload(payload, req.query);
    console.log(JSON.stringify({
      event: 'performance', requestId: req.requestId,
      cacheMs: 0,
      graphFetchMs: sourceMs,
      filterMs,
      analyticsMs,
      aggregationMs,
      paginationMs: Date.now() - paginationStartedAt,
      totalMs: Date.now() - requestStartedAt
    }));
    res.json(responsePayload);
  } catch (error) {
    console.error(JSON.stringify({ event: 'dashboard_error', requestId: req.requestId, message: error.message }));
    res.status(500).json({
      message: 'Erro ao carregar dados do SharePoint.',
      requestId: req.requestId
    });
  }
});

app.use((error, req, res, next) => {
  console.error(JSON.stringify({ event: 'unhandled_error', requestId: req.requestId, message: error.message }));
  res.status(500).json({ message: 'Erro interno do servidor.', requestId: req.requestId });
});

app.listen(port, () => {
  console.log(`Dashboard rodando em http://localhost:${port}`);
  if (cacheWarmupEnabled) setImmediate(() => warmupCaches().catch((error) => console.error(error)));
});
