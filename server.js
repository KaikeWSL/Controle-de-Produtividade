const express = require('express');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
const port = process.env.PORT || 3000;

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

app.use(cors({
  origin: true,
  credentials: true,
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

const monthNames = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

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

function getValue(item, fieldName) {
  if (!fieldName) return '';
  const key = normalizeFieldName(fieldName);
  if (!key) return '';

  const source = item && item.fields ? item.fields : item;
  const candidates = new Set([
    key,
    key.replace(/_x0020_/gi, ' '),
    key.replace(/_x0020_/gi, ''),
    key.replace(/_/g, ' '),
    key.replace(/x0020/gi, ' '),
    key.replace(/x0020/gi, ''),
    key.toLowerCase(),
    key.toUpperCase()
  ]);

  for (const candidate of candidates) {
    const directValue = source[candidate];
    if (directValue !== undefined && directValue !== null) return directValue;
  }

  const aliasKey = Object.keys(source || {}).find((candidate) => {
    const normalizedCandidate = candidate.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    const normalizedKey = key.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    return normalizedCandidate === normalizedKey;
  });

  if (aliasKey) return source[aliasKey];

  return '';
}

function parseSharePointDate(value) {
  if (!value) return null;

  if (typeof value === 'string') {
    const iso = value.split('T')[0];
    const parsed = new Date(iso);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }

  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) return parsed;

  return null;
}

function safeString(value) {
  if (value === undefined || value === null) return '';
  return String(value).trim();
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

  return data.access_token;
}

async function fetchListItems(site, filters = {}) {
  const normalizedSite = normalizeSiteConfig(site);

  if (!normalizedSite || !normalizedSite.url || !normalizedSite.listName) {
    console.warn('Lista de SharePoint não configurada para esta seleção:', site);
    return [];
  }

  const siteUrl = normalizedSite.url;
  const token = await getGraphToken();

  const fields = [
    normalizedSite.fields.projetista,
    normalizedSite.fields.dataConclusao,
    normalizedSite.fields.uploadVisium,
    normalizedSite.fields.uf,
    normalizedSite.fields.cidade
  ].filter(Boolean);

  const siteHostname = new URL(siteUrl).hostname;
  const sitePath = new URL(siteUrl).pathname.replace(/\/$/, '');
  const listName = (normalizedSite.listName || '').replace(/\s+/g, ' ').trim();
  const baseUrl = `https://graph.microsoft.com/v1.0/sites/${siteHostname}:/${sitePath}:/lists/${encodeURIComponent(listName)}/items?$expand=fields($select=${fields.map((field) => encodeURIComponent(field)).join(',')})&$top=1000`;

  let nextUrl = baseUrl;
  const allItems = [];

  while (nextUrl) {
    const response = await fetch(nextUrl, {
      headers: {
        Accept: 'application/json;odata.metadata=none',
        Authorization: `Bearer ${token}`
      }
    });

    const payload = await readJsonResponse(response);

    if (!response.ok) {
      const rawText = payload?.rawText || payload?.error?.message || JSON.stringify(payload);
      const detail = rawText && String(rawText).length > 0 ? String(rawText).slice(0, 400) : 'sem mensagem';

      if ([400, 401, 403, 404].includes(response.status)) {
        console.warn(`Lista ignorada por configuração inválida ou inexistente: ${normalizedSite.listName}. Detalhe: ${detail}`);
        return [];
      }

      throw new Error(`Erro ao consultar lista ${normalizedSite.listName}: ${detail}`);
    }

    const items = Array.isArray(payload.value) ? payload.value : [];
    allItems.push(...items);
    nextUrl = payload['@odata.nextLink'] || null;
  }

  return allItems
    .map((item) => {
      const fieldData = item.fields || {};
      const dataConclusao = getValue(fieldData, normalizedSite.fields.dataConclusao);
      const date = parseSharePointDate(dataConclusao);
      const hasUploadVisiumField = Boolean(normalizedSite.fields.uploadVisium);
      const uploadVisium = hasUploadVisiumField
        ? parseMetricValue(getValue(fieldData, normalizedSite.fields.uploadVisium))
        : 0;
      const uf = safeString(getValue(fieldData, normalizedSite.fields.uf)).toUpperCase();
      const cidade = safeString(getValue(fieldData, normalizedSite.fields.cidade));
      const projetista = safeString(getValue(fieldData, normalizedSite.fields.projetista));

      return {
        date,
        uploadVisium,
        uf,
        cidade,
        projetista,
        raw: item
      };
    })
    .filter((item) => {
      if (!item.projetista || !item.projetista.trim()) return false;
      if (filters.projetista && item.projetista.toLowerCase() !== String(filters.projetista).toLowerCase()) return false;
      if (filters.uf && item.uf !== filters.uf) return false;
      if (filters.cidade && item.cidade.toLowerCase() !== String(filters.cidade).toLowerCase()) return false;
      if (item.date) {
        if (filters.ano && item.date.getFullYear() !== Number(filters.ano)) return false;
        if (filters.mes && item.date.getMonth() + 1 !== Number(filters.mes)) return false;
      } else if (filters.ano || filters.mes) {
        return false;
      }

      return true;
    });
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
  const doneCount = items.filter((item) => item.date).length;

  items.forEach((item) => {
    const nome = safeString(item.projetista).trim();
    const uploadFlag = Number(item.uploadVisium || 0) > 0 ? 1 : 0;

    if (nome) {
      byProjetista[nome] = (byProjetista[nome] || 0) + 1;

      if (includeUploadVisium) {
        byProjetistaUpload[nome] = (byProjetistaUpload[nome] || 0) + uploadFlag;
      }

      projetistas.add(nome);
    }

    if (item.date) {
      const monthIndex = item.date.getMonth();
      byMonth[monthIndex].total += 1;
      years.add(item.date.getFullYear());
    }

    if (item.uf) ufs.add(item.uf);
    if (item.cidade) cidades.add(item.cidade);
  });

  const uploadVisiumTotal = includeUploadVisium
    ? Object.values(byProjetistaUpload).reduce((sum, value) => sum + Number(value || 0), 0)
    : 0;

  const barData = Object.entries(byProjetista)
    .map(([label, total]) => ({ label, total, upload: includeUploadVisium ? (byProjetistaUpload[label] || 0) : 0 }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 15);

  const summary = {
    total: items.length,
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

app.use(express.static(path.join(__dirname, 'public')));

app.get('/health', (_, res) => {
  res.json({ ok: true, service: 'sharepoint-dashboard', sites: config.sites?.length || 0 });
});

app.get('/api/config', (_, res) => {
  res.json(config);
});

app.get('/api/dashboard', async (req, res) => {
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

    if (!shouldRefresh && dashboardCache.has(cacheKey)) {
      return res.json(dashboardCache.get(cacheKey));
    }

    const items = await fetchListItems(selectedSite, filters);
    const result = buildResult(items, { includeUploadVisium: Boolean(selectedSite?.fields?.uploadVisium) });
    const payload = {
      site: selectedSite,
      filters,
      items: items.map((item) => ({
        projetista: item.projetista || '',
        uf: item.uf || '',
        cidade: item.cidade || '',
        uploadVisium: Number(item.uploadVisium || 0),
        date: item.date ? item.date.toISOString() : null
      })),
      ...result
    };

    dashboardCache.set(cacheKey, payload);
    res.json(payload);
  } catch (error) {
    console.error(error);
    res.status(500).json({
      message: 'Erro ao carregar dados do SharePoint.',
      detail: error.message
    });
  }
});

app.listen(port, () => {
  console.log(`Dashboard rodando em http://localhost:${port}`);
});
