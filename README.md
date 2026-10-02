# Dashboard de produtividade do SharePoint

Projeto em HTML + Node para substituir o Power BI com filtros por site, UF, cidade, mês e ano.

## Pré-requisitos

- Node.js 18+
- Credenciais de aplicativo no Microsoft Entra ID
- Permissão para listar itens da lista do SharePoint em App Only

## Passo 1: preencher a configuração

Você pode usar o arquivo `config.json` localmente ou, no Render, preencher as variáveis de ambiente.

### Opção A — arquivo JSON local

> Nunca grave `tenantId`, `clientId` ou `clientSecret` no `config.json`. Mantenha as credenciais exclusivamente no arquivo `.env`, que já é ignorado pelo Git.

Edite o arquivo `config.json` e informe apenas:

- URL do(s) site(s) SharePoint
- Nome da lista (`Visium`)
- Campos internos do SharePoint

Exemplo:

```json
{
  "sites": [
    {
      "name": "Equipe Procisacpia",
      "url": "https://corpclarobr.sharepoint.com/sites/USER-USER-EquipeProcisacpia",
      "listName": "Visium",
      "fields": {
        "projetista": "Projetista",
        "dataConclusao": "Data_x0020_Conclusao",
        "uf": "UF",
        "cidade": "Cidade"
      }
    }
  ]
}
```

### Opção B — variáveis de ambiente no Render

No Render, defina as seguintes variáveis:

```bash
TENANT_ID=...
CLIENT_ID=...
CLIENT_SECRET=...
APP_CONFIG={"tenantId":"...","clientId":"...","clientSecret":"...","sites":[{"name":"Equipe Procisacpia","url":"https://corpclarobr.sharepoint.com/sites/USER-USER-EquipeProcisacpia","listName":"Visium","fields":{"projetista":"Projetista","dataConclusao":"Data_x0020_Conclusao","uf":"UF","cidade":"Cidade"}}]}
```

## Passo 2: instalar dependências

```bash
npm install
```

## Passo 3: iniciar o dashboard

```bash
npm start
```

Em seguida acesse:

```text
http://localhost:3000
```

## Cache e desempenho

O servidor mantém três níveis de cache: resposta por filtro, fonte SharePoint normalizada em memória e snapshot compactado `gzip` em `.cache/`. A fonte armazena somente os campos usados pelo dashboard, cria índices por projetista, UF, cidade, ano e mês e é compartilhada por todos os usuários. Consultas simultâneas para a mesma lista usam uma única chamada ao Microsoft Graph.

Configure no Render conforme a memória disponível do serviço:

```bash
CACHE_TTL_SECONDS=300
CACHE_MAX_ITEMS=100000
CACHE_CLEANUP_INTERVAL=60
CACHE_WARMUP_ENABLED=true
```

`CACHE_MAX_ITEMS` limita a quantidade total de registros normalizados em memória e remove a fonte menos recentemente usada. O disco local do Render pode ser efêmero; para cache persistente entre deploys/restarts, anexe um Persistent Disk e defina `CACHE_DIRECTORY` para um caminho nesse disco.

Endpoints operacionais:

```text
GET  /api/cache/stats
POST /api/cache/clear
POST /api/cache/clear?persistent=true
POST /api/cache/refresh
```

`POST /api/cache/refresh` aceita opcionalmente JSON com `site`, `activity` e `subactivity`. A rota do dashboard aceita `page` e `pageSize` como extensão compatível, por exemplo `/api/dashboard?site=Hfc&page=1&pageSize=100`.

## Deploy em Render + Netlify

### Backend no Render

- Conecte este projeto ao Render.
- Configure a porta com a variável `PORT` (o Render já faz isso)
- Defina `APP_CONFIG` ou `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`
- O backend ficará em algo como:

```text
https://SEU-SERVICO.onrender.com
```

### Frontend no Netlify

No site estático do Netlify, crie um arquivo de configuração ou variável global para apontar para a API do Render. Exemplo:

```html
<script>
  window.API_BASE_URL = 'https://SEU-SERVICO.onrender.com';
</script>
```

Se preferir, você pode colocar isso em um arquivo `public/config.js` e referenciar antes do `app.js`.

## Permissões no SharePoint / Azure AD

No Microsoft Entra ID, adicione a aplicação com permissão de aplicativo para SharePoint:

- `Sites.FullControl.All`
- ou pelo menos leitura nas listas do site

## Observação importante

O campo `Data_x0020_Conclusao` normalmente é o nome interno do SharePoint. Se a sua coluna tiver nome diferente, ajuste no `config.json` ou na variável `APP_CONFIG`.

Se sua lista tiver nomes com espaço, o SharePoint usa nomes internos como:

- `Data_x0020_Conclusao`
- `Custo_x0020_Total`
- `Cidade`
- `UF`

Use estes nomes no campo `fields`.

## Permissões no SharePoint / Azure AD

No Microsoft Entra ID, adicione a aplicação com permissão de aplicativo para SharePoint:

- Sites.FullControl.All
- ou pelo menos leitura nas listas do site

## Observação importante

O campo `Data_x0020_Conclusao` normalmente é o nome interno do SharePoint. Se a sua coluna tiver nome diferente, ajuste no `config.json`.

Se sua lista tiver nomes com espaço, o SharePoint usa nomes internos como:

- `Data_x0020_Conclusao`
- `Custo_x0020_Total`
- `Cidade`
- `UF`

Use estes nomes no campo `fields`.
