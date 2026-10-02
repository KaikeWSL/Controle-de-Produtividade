const express = require('express');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
const port = process.env.PORT || 3000;
const cacheTtlMs = Math.max(30, Number(process.env.CACHE_TTL_SECONDS) || 300) * 1000;
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
const requestBuckets = new Map();
let graphTokenCache = null;

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
  const allowedKeys = new Set(['site', 'activity', 'subactivity', 'projetista', 'uf', 'cidade', 'mes', 'ano', 'refresh']);
  const invalidKey = Object.keys(req.query).find((key) => !allowedKeys.has(key));
  const invalidText = ['site', 'activity', 'subactivity', 'projetista', 'uf', 'cidade'].find((key) => String(req.query[key] || '').length > 120);
  const month = String(req.query.mes || '');
  const year = String(req.query.ano || '');

  if (invalidKey || invalidText || (month && !/^(?:[1-9]|1[0-2])$/.test(month)) || (year && !/^\d{4}$/.test(year))) {
    return res.status(400).json({ message: 'Parâmetros de filtro inválidos.' });
  }

  next();
}

app.use(applySecurityHeaders);
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
  return graphTokenCache.accessToken;
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
  if (!response.ok) throw new Error(`Falha ao obter token do SharePoint: ${data.error_description || data.error || 'sem detalhe'}`);
  return data.access_token;
}

function escapeODataString(value) {
  return String(value).replace(/'/g, "''");
}

async function publishSummaryAttachment(pngBuffer) {
  const { siteUrl, listName, title, fileName } = summaryAttachmentTarget;
  const token = await getSharePointToken(siteUrl);
  const listPath = `/_api/web/lists/getbytitle('${escapeODataString(listName)}')`;
  const headers = { Accept: 'application/json;odata=nometadata', Authorization: `Bearer ${token}` };
  const itemLookup = await fetch(`${siteUrl}${listPath}/items?$select=Id&$filter=Title eq '${escapeODataString(title)}'&$top=1`, { headers });
  const lookupData = await readJsonResponse(itemLookup);
  if (!itemLookup.ok) throw new Error(`Não foi possível localizar a lista ${listName}.`);

  let itemId = Array.isArray(lookupData.value) && lookupData.value[0]?.Id;
  if (!itemId) {
    const createResponse = await fetch(`${siteUrl}${listPath}/items`, {
      method: 'POST',
      headers: { ...headers, 'Content-Type': 'application/json;odata=nometadata' },
      body: JSON.stringify({ Title: title })
    });
    const createdItem = await readJsonResponse(createResponse);
    if (!createResponse.ok || !createdItem.Id) throw new Error(`Não foi possível criar o item ${title} na lista ${listName}.`);
    itemId = createdItem.Id;
  }

  const attachmentPath = `${listPath}/items(${itemId})/AttachmentFiles/getbyfilename('${escapeODataString(fileName)}')`;
  const deleteResponse = await fetch(`${siteUrl}${attachmentPath}`, {
    method: 'POST',
    headers: { ...headers, 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' }
  });
  if (!deleteResponse.ok && deleteResponse.status !== 404) throw new Error('Não foi possível substituir o anexo anterior do resumo.');

  const uploadResponse = await fetch(`${siteUrl}${listPath}/items(${itemId})/AttachmentFiles/add(FileName='${escapeODataString(fileName)}')`, {
    method: 'POST',
    headers: { ...headers, 'Content-Type': 'image/png' },
    body: pngBuffer
  });
  if (!uploadResponse.ok) throw new Error('Não foi possível enviar o PNG do resumo ao SharePoint.');
  return itemId;
}

function buildLocationKey(uf, cidade) {
  return `${safeString(uf)}|${safeString(cidade)}`
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9|]/g, '')
    .toUpperCase();
}

async function fetchWeightMap(weightSource) {
  if (!weightSource?.url || !weightSource?.listName || !weightSource?.fields) return new Map();

  const cacheKey = JSON.stringify(weightSource);
  const cached = weightSourceCache.get(cacheKey);
  if (cached && cached.expiresAt > Date.now()) return cached.weights;

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
      (Array.isArray(payload.value) ? payload.value : []).forEach((item) => {
        const fieldData = item.fields || {};
        const key = buildLocationKey(getValue(fieldData, weightSource.fields.uf), getValue(fieldData, weightSource.fields.cidade));
        const peso = parseMetricValue(getValue(fieldData, weightSource.fields.peso));
        if (key !== '|' && Number.isFinite(peso) && peso > 0) weights.set(key, peso);
      });
      nextUrl = payload['@odata.nextLink'] || null;
    }

    if (loadedAny) break;
  }

  weightSourceCache.set(cacheKey, { weights, expiresAt: Date.now() + cacheTtlMs });
  console.log(JSON.stringify({ event: 'weight_source_loaded', list: weightSource.listName, locations: weights.size }));
  return weights;
}

async function fetchListItems(site, filters = {}, options = {}) {
  const normalizedSite = normalizeSiteConfig(site);

  if (!normalizedSite || !normalizedSite.url || !normalizedSite.listName) {
    console.error('[SharePoint] Lista não configurada para a seleção:', site);
    return [];
  }

  const siteUrl = normalizedSite.url;

  const fields = [
    normalizedSite.fields.projetista,
    normalizedSite.fields.dataConclusao,
    normalizedSite.fields.uploadVisium,
    normalizedSite.fields.uf,
    normalizedSite.fields.cidade
  ].filter(Boolean);

  const siteHostname = new URL(siteUrl).hostname;
  const sitePath = new URL(siteUrl).pathname.replace(/\/$/, '');
  const listCandidates = getListNameCandidates(normalizedSite.listName);
  let activeListName = normalizedSite.listName;
  let allItems = [];
  const sourceCacheKey = JSON.stringify({ url: siteUrl, listName: normalizedSite.listName, fields });
  const sourceCache = sourceListCache.get(sourceCacheKey);

  if (!options.forceRefresh && sourceCache && sourceCache.expiresAt > Date.now()) {
    allItems = sourceCache.items;
    activeListName = sourceCache.activeListName;
    console.log(JSON.stringify({ event: 'source_cache_hit', list: activeListName, items: allItems.length }));
  } else {
    const token = await getGraphToken();
    for (const listName of listCandidates) {
      const baseUrl = `https://graph.microsoft.com/v1.0/sites/${siteHostname}:/${sitePath}:/lists/${encodeURIComponent(listName)}/items?$expand=fields($select=${fields.map((field) => encodeURIComponent(field)).join(',')})&$top=1000`;
      let nextUrl = baseUrl;
      const pageItems = [];

      while (nextUrl) {
        const response = await fetch(nextUrl, {
          headers: {
            Accept: 'application/json;odata.metadata=none',
            Authorization: `Bearer ${token}`
          }
        });

        const payload = await readJsonResponse(response);
        console.log(`[SharePoint] list=${listName} page=${pageItems.length + 1} status=${response.status}`);

        if (!response.ok) {
          const rawText = payload?.rawText || payload?.error?.message || JSON.stringify(payload);
          const detail = rawText && String(rawText).length > 0 ? String(rawText).slice(0, 400) : 'sem mensagem';

          if ([400, 401, 403, 404].includes(response.status)) {
            console.warn(`Lista não encontrada com nome alternativo: ${listName}. Detalhe: ${detail}`);
            break;
          }

          throw new Error(`Erro ao consultar lista ${listName}: ${detail}`);
        }

        const items = Array.isArray(payload.value) ? payload.value : [];
        pageItems.push(...items);
        nextUrl = payload['@odata.nextLink'] || null;
      }

      if (pageItems.length > 0) {
        allItems = pageItems;
        activeListName = listName;
        break;
      }
    }

    if (allItems.length) {
      sourceListCache.set(sourceCacheKey, { items: allItems, activeListName, expiresAt: Date.now() + cacheTtlMs });
      console.log(JSON.stringify({ event: 'source_cache_saved', list: activeListName, items: allItems.length, ttlMs: cacheTtlMs }));
    }
  }

  if (!allItems.length && listCandidates.length) {
    console.error('[SharePoint] Nenhuma lista foi encontrada para a configuração atual.', {
      site: normalizedSite.name,
      url: normalizedSite.url,
      listName: normalizedSite.listName,
      candidates: listCandidates,
      attemptedFields: fields
    });
  }

  let uploadDebugLogged = false;
  const weightMap = await fetchWeightMap(normalizedSite.weightSource);

  const mappedItems = allItems
    .map((item) => {
      const fieldData = item.fields || {};
      const dataConclusao = getValue(fieldData, normalizedSite.fields.dataConclusao);
      const date = parseSharePointDate(dataConclusao);
      const hasUploadVisiumField = Boolean(normalizedSite.fields.uploadVisium);
      const uploadVisiumRaw = hasUploadVisiumField ? getValue(fieldData, normalizedSite.fields.uploadVisium) : '';
      const uploadVisium = hasUploadVisiumField ? normalizeUploadValue(uploadVisiumRaw) : 0;
      const uf = safeString(getValue(fieldData, normalizedSite.fields.uf)).toUpperCase();
      const cidade = safeString(getValue(fieldData, normalizedSite.fields.cidade));
      const projetista = safeString(getValue(fieldData, normalizedSite.fields.projetista));
      const peso = weightMap.get(buildLocationKey(uf, cidade)) || 1;

      const shouldLogProblemActivity = ['Migração', 'MIGRAO', 'Sar Fo', 'Revisões', 'Revises', 'As-Built', 'AsBuilt'].includes(normalizedSite.activityName || normalizedSite.name || '')
        || ['Migração', 'MIGRAO', 'Sar Fo', 'Revisões', 'Revises', 'As-Built', 'AsBuilt'].includes(normalizedSite.listName || '');

      if (!date && normalizedSite.fields.dataConclusao && shouldLogProblemActivity) {
        console.warn('[SharePoint] Coluna de data não reconhecida ou vazia.', {
          listName: normalizedSite.listName,
          activityName: normalizedSite.activityName || normalizedSite.name,
          fieldName: normalizedSite.fields.dataConclusao,
          rawValue: dataConclusao,
          sampleKeys: Object.keys(fieldData).slice(0, 20),
          sampleItem: Object.fromEntries(Object.entries(fieldData).slice(0, 6))
        });
      }

      if (
        !normalizedSite.fields.dataConclusao &&
        !normalizedSite.fields.projetista &&
        shouldLogProblemActivity
      ) {
        console.error('[SharePoint] Colunas obrigatórias ausentes para a atividade.', {
          listName: normalizedSite.listName,
          activityName: normalizedSite.activityName || normalizedSite.name,
          fields: normalizedSite.fields
        });
      }

      if (
        !uploadDebugLogged &&
        normalizedSite.fields.uploadVisium &&
        (uploadVisiumRaw === '' || uploadVisiumRaw === undefined || uploadVisiumRaw === null)
      ) {
        console.log('[UploadVisium debug]', {
          listName: normalizedSite.listName,
          fieldName: normalizedSite.fields.uploadVisium,
          sampleKeys: Object.keys(fieldData).slice(0, 10),
          sampleItem: Object.fromEntries(Object.entries(fieldData).slice(0, 5)),
          message: 'Campo configurado, porém vazio no payload real do SharePoint.'
        });
        uploadDebugLogged = true;
      }

      return {
        date,
        uploadVisium,
        uf,
        cidade,
        projetista,
        peso,
        raw: item
      };
    })
    .filter((item) => {
      if (filters.projetista && String(item.projetista || '').toLowerCase() !== String(filters.projetista).toLowerCase()) return false;
      if (filters.uf && item.uf !== filters.uf) return false;
      if (filters.cidade && String(item.cidade || '').toLowerCase() !== String(filters.cidade).toLowerCase()) return false;
      if (item.date) {
        if (filters.ano && item.date.getFullYear() !== Number(filters.ano)) return false;
        if (filters.mes && item.date.getMonth() + 1 !== Number(filters.mes)) return false;
      } else if (filters.ano || filters.mes) {
        return false;
      }

      return true;
    });

  console.log(`[SharePoint] list=${normalizedSite.listName} loaded=${mappedItems.length} filtered=${mappedItems.length} raw=${allItems.length}`);
  return mappedItems;
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
  const doneCount = items.filter((item) => item.date).reduce((sum, item) => sum + Number(item.peso || 1), 0);

  items.forEach((item) => {
    const nome = safeString(item.projetista).trim();
    const uploadFlag = Number(item.uploadVisium || 0) > 0 ? 1 : 0;
    const peso = Number(item.peso || 1);

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
    }

    if (item.uf) ufs.add(item.uf);
    if (item.cidade) cidades.add(item.cidade);
  });

  const uploadVisiumTotal = includeUploadVisium
    ? items.filter((item) => Number(item.uploadVisium || 0) > 0).length
    : 0;

  const barData = Object.entries(byProjetista)
    .map(([label, total]) => ({ label, total, upload: includeUploadVisium ? (byProjetistaUpload[label] || 0) : 0 }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 15);

  const summary = {
    total: items.reduce((sum, item) => sum + Number(item.peso || 1), 0),
    done: doneCount,
    uploadVisiumTotal,
    uniqueProjetistas: Object.keys(byProjetista).length,
    anoInicial: Math.min(...Array.from(years), 0) || new Date().getFullYear(),
    anoFinal: Math.max(...Array.from(years), new Date().getFullYear())
  };

  return {
    summary,
    barData,
    monthlyTrend: byMonth.map((m) => ({ label: m.label, total: m.total })),
    years: Array.from(years).sort((a, b) => a - b),
    ufs: Array.from(ufs).sort(),
    cidades: Array.from(cidades).sort((a, b) => a.localeCompare(b)),
    projetistas: Array.from(projetistas).sort((a, b) => a.localeCompare(b)),
    includeUploadVisium
  };
}

function buildExecutiveAnalytics(items, comparisonItems = items) {
  const datedItems = items.filter((item) => item.date);
  const comparisonDates = comparisonItems
    .filter((item) => !Number.isNaN(item.date.getTime()))
    .map((item) => ({ date: item.date, peso: Number(item.peso || 1) }));
  const latestDates = datedItems.filter((item) => !Number.isNaN(item.date.getTime()));
  const latestDate = latestDates.length ? new Date(Math.max(...latestDates.map((item) => item.date.getTime()))) : new Date();
  const dayStart = new Date(latestDate.getFullYear(), latestDate.getMonth(), latestDate.getDate());
  const previousDayStart = new Date(dayStart);
  previousDayStart.setDate(previousDayStart.getDate() - 1);
  const weekStart = new Date(dayStart);
  weekStart.setDate(weekStart.getDate() - 6);
  const monthStart = new Date(dayStart.getFullYear(), dayStart.getMonth(), 1);
  const previousMonthStart = new Date(dayStart.getFullYear(), dayStart.getMonth() - 1, 1);
  const designers = new Map();
  const cities = new Map();
  const countBetween = (start, end) => comparisonDates
    .filter((item) => item.date >= start && item.date < end)
    .reduce((sum, item) => sum + item.peso, 0);

  datedItems.forEach((item) => {
    const designer = safeString(item.projetista);
    const city = safeString(item.cidade);
    const peso = Number(item.peso || 1);

    if (!isIgnorableProjetistaName(designer)) designers.set(designer, (designers.get(designer) || 0) + peso);
    if (city) cities.set(city, (cities.get(city) || 0) + peso);
  });

  const dayTotal = countBetween(dayStart, new Date(dayStart.getTime() + 86400000));
  const previousDayTotal = countBetween(previousDayStart, dayStart);
  const weekTotal = countBetween(weekStart, new Date(dayStart.getTime() + 86400000));
  const monthTotal = countBetween(monthStart, new Date(dayStart.getTime() + 86400000));
  const previousMonthTotal = countBetween(previousMonthStart, monthStart);
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
    topCity: Array.from(cities.entries())
      .map(([name, total]) => ({ name, total }))
      .sort((left, right) => right.total - left.total)[0] || null
  };
}

app.use(express.static(path.join(__dirname, 'public')));

app.get('/health', (_, res) => {
  res.json({ ok: true, service: 'sharepoint-dashboard', sites: config.sites?.length || 0 });
});

app.get('/api/config', (_, res) => {
  res.json(getPublicConfig(config));
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
    res.status(502).json({ message: 'Não foi possível publicar o resumo no SharePoint.', requestId: req.requestId });
  }
});

app.get('/api/dashboard', validateDashboardQuery, async (req, res) => {
  try {
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
      console.log(JSON.stringify({ event: 'cache_hit', requestId: req.requestId, key: cacheKey }));
      return res.json(cached.payload);
    }

    if (cached) dashboardCache.delete(cacheKey);

    const items = await fetchListItems(selectedSite, filters, { forceRefresh: shouldRefresh });
    const comparisonFilters = { ...filters, mes: '', ano: '' };
    const comparisonItems = await fetchListItems(selectedSite, comparisonFilters);
    console.log(`[dashboard] items received=${items.length} for ${siteName}/${activityName}`);
    const result = buildResult(items, { includeUploadVisium: Boolean(selectedSite?.fields?.uploadVisium) });
    const analytics = buildExecutiveAnalytics(items, comparisonItems);
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

    dashboardCache.set(cacheKey, { payload, expiresAt: Date.now() + cacheTtlMs });
    res.json(payload);
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
});
