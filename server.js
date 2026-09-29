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

function normalizeSiteConfig(site) {
  if (!site) return null;

  return {
    ...site,
    name: site.name || site.siteName || 'Site',
    url: site.url || site.siteUrl || '',
    listName: site.listName || site.listTitle || site.title || '',
    fields: {
      ...site.fields,
      projetista: site.fields?.projetista || site.fields?.project || '',
      dataConclusao: site.fields?.dataConclusao || site.fields?.dateConclusion || '',
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

  const directValue = source[key];
  if (directValue !== undefined && directValue !== null) return directValue;

  const fallback = source[`${key}`.replace(/_/g, 'x0020')];
  if (fallback !== undefined && fallback !== null) return fallback;

  const aliasKey = Object.keys(source || {}).find((candidate) => candidate.toLowerCase() === key.toLowerCase());
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
  const siteUrl = normalizedSite.url;
  const token = await getGraphToken();

  const fields = [
    normalizedSite.fields.projetista,
    normalizedSite.fields.dataConclusao,
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
      const uf = safeString(getValue(fieldData, normalizedSite.fields.uf)).toUpperCase();
      const cidade = safeString(getValue(fieldData, normalizedSite.fields.cidade));
      const projetista = safeString(getValue(fieldData, normalizedSite.fields.projetista));

      return {
        date,
        uf,
        cidade,
        projetista,
        raw: item
      };
    })
    .filter((item) => {
      if (!item.projetista || !item.date) return false;

      if (filters.projetista && item.projetista.toLowerCase() !== String(filters.projetista).toLowerCase()) return false;
      if (filters.uf && item.uf !== filters.uf) return false;
      if (filters.cidade && item.cidade.toLowerCase() !== String(filters.cidade).toLowerCase()) return false;
      if (filters.ano && item.date.getFullYear() !== Number(filters.ano)) return false;
      if (filters.mes && item.date.getMonth() + 1 !== Number(filters.mes)) return false;

      return true;
    });
}

function buildResult(items) {
  const byProjetista = {};
  const byMonth = Array.from({ length: 12 }, (_, idx) => ({ month: idx + 1, label: monthNames[idx], total: 0 }));
  const years = new Set();
  const ufs = new Set();
  const cidades = new Set();
  const projetistas = new Set();

  items.forEach((item) => {
    const nome = item.projetista || 'Não informado';
    byProjetista[nome] = (byProjetista[nome] || 0) + 1;
    projetistas.add(nome);

    if (item.date) {
      const monthIndex = item.date.getMonth();
      byMonth[monthIndex].total += 1;
      years.add(item.date.getFullYear());
    }

    if (item.uf) ufs.add(item.uf);
    if (item.cidade) cidades.add(item.cidade);
  });

  const barData = Object.entries(byProjetista)
    .map(([label, total]) => ({ label, total }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 15);

  const summary = {
    total: items.length,
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
    projetistas: Array.from(projetistas).sort((a, b) => a.localeCompare(b))
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
    const selectedSite = normalizeSiteConfig(getSiteByName(siteName));
    const filters = {
      projetista: req.query.projetista || '',
      uf: req.query.uf || '',
      cidade: req.query.cidade || '',
      mes: req.query.mes || '',
      ano: req.query.ano || ''
    };

    const items = await fetchListItems(selectedSite, filters);
    const result = buildResult(items);

    res.json({
      site: selectedSite,
      filters,
      ...result
    });
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
