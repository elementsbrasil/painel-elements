// ============================================================================
//  PAINEL COMERCIAL — função de dados
//
//  Roda no servidor da Netlify. Consulta as DUAS fontes ao vivo a cada
//  chamada e devolve um JSON no formato que o index.html espera.
//
//    SharePoint (Microsoft Graph)  -> hero, canais, evolução, vendedores
//    BigQuery (Google Cloud)       -> território, produtos, pedidos, ticket
//
//  CREDENCIAIS: lidas de variáveis de ambiente. Nunca no código, nunca no
//  GitHub. Configurar em Netlify > Site configuration > Environment variables:
//
//    MS_CLIENT_ID              (Azure > App registration > Overview)
//    MS_TENANT_ID              (idem)
//    MS_CLIENT_SECRET          (Azure > Certificates & secrets)  [EXPIRA!]
//    SHAREPOINT_SITE_ID        (Graph Explorer)
//    SHAREPOINT_ITEM_ID        (Graph Explorer)
//    BQ_SERVICE_ACCOUNT_JSON   (conteúdo do arquivo JSON, colado inteiro)
//    MES_ABA                   (opcional: força uma aba, ex. "AGOSTO")
//
//  DEGRADAÇÃO: se uma fonte falhar, a outra ainda é entregue. O painel mostra
//  "aguardando" no bloco sem dado em vez de número inventado (regra 7 da spec).
// ============================================================================

const GRAPH = "https://graph.microsoft.com/v1.0";

// Nomes reais das abas em uso na planilha (não são os meses completos)
const ABAS_POR_MES = [
  "JAN", "FEV", "MAR", "ABRIL", "MAIO", "JUNHO",
  "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
];

const CANAIS = [
  { nome: "Revenda",     codcencus: 10102001, linhaCanal: 3,  vendIni: 13, vendFim: 18, histCols: ["G","H"] },
  { nome: "Corporativo", codcencus: 10102002, linhaCanal: 4,  vendIni: 21, vendFim: 28, histCols: ["L","M"] },
  { nome: "Digital",     codcencus: 10102003, linhaCanal: 5,  vendIni: 31, vendFim: 34, histCols: ["Q","R"] }
];

const PROJETO_BQ = "elements-489322";
const DATASET_BQ = "dw_bronze";

// ---------------------------------------------------------------------------
//  Helpers de valor
// ---------------------------------------------------------------------------
function num(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  let s = String(v).replace(/[R$\s%]/g, "");
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
const pctFrac = v => num(v) * 100;   // planilha guarda 0.5049 -> 50.49
function cel(m, linha, col) {
  if (!m || !m[linha]) return null;
  const v = m[linha][col];
  return v === undefined ? null : v;
}

// ---------------------------------------------------------------------------
//  SHAREPOINT
// ---------------------------------------------------------------------------
async function tokenGraph() {
  const url = `https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`;
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.MS_CLIENT_ID,
      client_secret: process.env.MS_CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials"
    }).toString()
  });
  if (!res.ok) throw new Error(`Token Microsoft falhou (${res.status}): ${await res.text()}`);
  return (await res.json()).access_token;
}

// Lê um range específico de uma aba específica — nunca varre o arquivo inteiro,
// então a restrição das 7 abas ocultas legadas não nos afeta.
async function lerRange(token, aba, endereco) {
  const url = `${GRAPH}/sites/${process.env.SHAREPOINT_SITE_ID}` +
              `/drive/items/${process.env.SHAREPOINT_ITEM_ID}` +
              `/workbook/worksheets('${encodeURIComponent(aba)}')` +
              `/range(address='${endereco}')?$select=values`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Leitura ${aba}!${endereco} falhou (${res.status}): ${await res.text()}`);
  return (await res.json()).values;
}

// Vendedores de um bloco. Mantém quem tem meta e vendeu 0 (aparece com R$ 0).
function vendedores(m, ini, fim) {
  const out = [];
  for (let r = ini; r <= fim; r++) {
    const i = r - 1;                       // planilha 1-based -> range 0-based
    const nome = String(cel(m, i, 0) ?? "").trim();
    if (!nome || nome.toUpperCase() === "TOTAL") continue;
    const meta = num(cel(m, i, 1));        // col B
    const vend = num(cel(m, i, 2));        // col C
    if (!meta && !vend) continue;
    out.push({ nome, valor: vend, meta });
  }
  out.sort((a, b) => b.valor - a.valor);
  return out.slice(0, 4);
}

// Histórico mensal da aba Dashboard (range A4:S15 = Jan..Dez)
function historico(mHist, colAting, colMeta, ateIdx) {
  const ABREV = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
  const IDX = { A:0,B:1,C:2,D:3,E:4,F:5,G:6,H:7,I:8,J:9,K:10,L:11,M:12,N:13,O:14,P:15,Q:16,R:17,S:18 };
  const out = [];
  for (let i = 0; i <= ateIdx && i < 12; i++) {
    const a = num(cel(mHist, i, IDX[colAting]));
    const m = num(cel(mHist, i, IDX[colMeta]));
    if (!a && !m) continue;                // mês futuro, sem movimento
    out.push({ mes: ABREV[i], atingido: a, meta: m });
  }
  return out;
}

async function lerPlanilha(aba, mesIdx) {
  const token = await tokenGraph();
  const [mMes, mHist] = await Promise.all([
    lerRange(token, aba, "A1:N43"),
    lerRange(token, "Dashboard", "A4:S15")
  ]);

  const hero = {
    valor: num(cel(mMes, 2, 1)),               // B3  Vendido
    meta: num(cel(mMes, 1, 1)),                // B2  META
    pct_atingido: pctFrac(cel(mMes, 3, 1)),    // B4  % Atingido
    projecao: num(cel(mMes, 5, 1)),            // B6  Projeção
    projecao_pct: pctFrac(cel(mMes, 6, 1))     // B7  Projeção %
  };
  hero.alcance_vs_previsto = hero.projecao_pct;

  const canais = CANAIS.map(cfg => {
    const i = cfg.linhaCanal - 1;
    return {
      nome: cfg.nome,
      realizado:    num(cel(mMes, i, 8)),      // col I
      projetado:    num(cel(mMes, i, 9)),      // col J
      projecao_pct: pctFrac(cel(mMes, i, 10)), // col K
      meta:         num(cel(mMes, i, 11)),     // col L
      pct_meta:     pctFrac(cel(mMes, i, 12)), // col M
      falta:        num(cel(mMes, i, 13)),     // col N
      historico:    historico(mHist, cfg.histCols[0], cfg.histCols[1], mesIdx),
      vendedores:   vendedores(mMes, cfg.vendIni, cfg.vendFim)
    };
  });
  canais.forEach(c => { c.alcance_vs_previsto = c.projecao_pct; });

  // Taxas de conversão do site (planilha guarda como texto "0,54%")
  const conversao = {
    b2c:  { meta: num(cel(mMes, 19, 12)), atual: num(cel(mMes, 20, 12)) },
    corp: { meta: num(cel(mMes, 19, 13)), atual: num(cel(mMes, 20, 13)) }
  };

  return { hero, canais, conversao };
}

// ---------------------------------------------------------------------------
//  BIGQUERY
//
//  ATENÇÃO: os nomes de tabela/coluna abaixo seguem o padrão Sankhya
//  (TGFCAB cabeçalho · TGFITE itens · TGFPRO produto · TGFPAR parceiro) e
//  AINDA NÃO FORAM CONFIRMADOS contra o schema real. Rodar o endpoint com
//  ?schema=1 para listar tabelas/colunas de verdade e ajustar as constantes.
// ---------------------------------------------------------------------------
const T = {
  cab: `\`${PROJETO_BQ}.${DATASET_BQ}.brz_sankhya_tgfcab_full\``,
  ite: `\`${PROJETO_BQ}.${DATASET_BQ}.brz_sankhya_tgfite_full\``,
  pro: `\`${PROJETO_BQ}.${DATASET_BQ}.brz_sankhya_tgfpro_full\``,
  par: `\`${PROJETO_BQ}.${DATASET_BQ}.brz_sankhya_tgfpar_full\``
};

function clienteBQ() {
  const raw = process.env.BQ_SERVICE_ACCOUNT_JSON;
  if (!raw) throw new Error("BQ_SERVICE_ACCOUNT_JSON não configurada");
  const { BigQuery } = require("@google-cloud/bigquery");
  return new BigQuery({
    projectId: PROJETO_BQ,
    credentials: JSON.parse(raw)
  });
}

// Base comum: notas de venda do mês corrente, deduplicadas.
// tipmov='V' exclui devoluções. QUALIFY resolve duplicata de ingestão das _full.
const CTE_CAB = `
  cab AS (
    SELECT nunota, codcencus, codparc, vlrnota, dtneg
    FROM ${T.cab}
    WHERE tipmov = 'V'
      AND DATE(dtneg) BETWEEN DATE_TRUNC(CURRENT_DATE(), MONTH) AND CURRENT_DATE()
    QUALIFY ROW_NUMBER() OVER (PARTITION BY nunota ORDER BY ingested_at_utc DESC) = 1
  )`;

const Q = {
  // Top 5 produtos por RECEITA (não por unidades), por canal
  produtos: `
    WITH ${CTE_CAB},
    ite AS (
      SELECT nunota, sequencia, codprod, qtdneg, vlrtot
      FROM ${T.ite}
      QUALIFY ROW_NUMBER() OVER (PARTITION BY nunota, sequencia ORDER BY ingested_at_utc DESC) = 1
    ),
    pro AS (
      SELECT codprod, descrprod
      FROM ${T.pro}
      QUALIFY ROW_NUMBER() OVER (PARTITION BY codprod ORDER BY ingested_at_utc DESC) = 1
    )
    SELECT cab.codcencus, pro.descrprod AS nome,
           SUM(ite.vlrtot) AS receita, SUM(ite.qtdneg) AS unidades
    FROM cab
    JOIN ite USING (nunota)
    LEFT JOIN pro USING (codprod)
    WHERE cab.codcencus IN (10102001, 10102002, 10102003)
    GROUP BY cab.codcencus, nome
    QUALIFY ROW_NUMBER() OVER (PARTITION BY cab.codcencus ORDER BY receita DESC) <= 5
    ORDER BY cab.codcencus, receita DESC`,

  // Pedidos, ticket médio e maior cliente do mês, por canal
  pedidos: `
    WITH ${CTE_CAB},
    par AS (
      SELECT codparc, nomeparc
      FROM ${T.par}
      QUALIFY ROW_NUMBER() OVER (PARTITION BY codparc ORDER BY ingested_at_utc DESC) = 1
    ),
    agg AS (
      SELECT codcencus, COUNT(DISTINCT nunota) AS pedidos,
             SUM(vlrnota) AS receita,
             SAFE_DIVIDE(SUM(vlrnota), COUNT(DISTINCT nunota)) AS ticket_medio
      FROM cab GROUP BY codcencus
    ),
    top_cli AS (
      SELECT cab.codcencus, par.nomeparc AS cliente, SUM(cab.vlrnota) AS valor
      FROM cab LEFT JOIN par USING (codparc)
      GROUP BY cab.codcencus, cliente
      QUALIFY ROW_NUMBER() OVER (PARTITION BY cab.codcencus ORDER BY valor DESC) = 1
    )
    SELECT agg.*, top_cli.cliente, top_cli.valor AS valor_cliente
    FROM agg LEFT JOIN top_cli USING (codcencus)`,

  // Representatividade por UF — só canal Digital (B2C)
  territorio: `
    WITH ${CTE_CAB},
    par AS (
      SELECT codparc, uf
      FROM ${T.par}
      QUALIFY ROW_NUMBER() OVER (PARTITION BY codparc ORDER BY ingested_at_utc DESC) = 1
    )
    SELECT par.uf AS nome, SUM(cab.vlrnota) AS receita,
           SAFE_DIVIDE(SUM(cab.vlrnota), SUM(SUM(cab.vlrnota)) OVER ()) * 100 AS pct
    FROM cab LEFT JOIN par USING (codparc)
    WHERE cab.codcencus = 10102003
    GROUP BY nome
    ORDER BY receita DESC
    LIMIT 5`,

  // Descobre o schema real do dataset (não cobra: INFORMATION_SCHEMA é grátis)
  schema: `
    SELECT table_name, column_name, data_type
    FROM \`${PROJETO_BQ}.${DATASET_BQ}.INFORMATION_SCHEMA.COLUMNS\`
    ORDER BY table_name, ordinal_position`
};

async function lerBigQuery() {
  const bq = clienteBQ();
  const run = async q => (await bq.query({ query: q, location: "US" }))[0];

  // Rodadas independentes: se uma query falhar, as outras ainda entregam
  const [prods, peds, terr] = await Promise.allSettled([
    run(Q.produtos), run(Q.pedidos), run(Q.territorio)
  ]);

  const porCanal = {};
  CANAIS.forEach(c => {
    porCanal[c.nome] = {
      produtos: [], pedidos: null, ticket_medio: null, maior_cliente: null
    };
  });
  const nomePorCod = {};
  CANAIS.forEach(c => { nomePorCod[c.codcencus] = c.nome; });

  if (prods.status === "fulfilled") {
    prods.value.forEach(r => {
      const alvo = porCanal[nomePorCod[Number(r.codcencus)]];
      if (alvo) alvo.produtos.push({
        nome: r.nome || "(sem descrição)",
        receita: Number(r.receita) || 0,
        unidades: Number(r.unidades) || 0
      });
    });
  }

  if (peds.status === "fulfilled") {
    peds.value.forEach(r => {
      const alvo = porCanal[nomePorCod[Number(r.codcencus)]];
      if (!alvo) return;
      alvo.pedidos = Number(r.pedidos) || 0;
      alvo.ticket_medio = Number(r.ticket_medio) || 0;
      if (r.cliente) alvo.maior_cliente = {
        nome: r.cliente, valor: Number(r.valor_cliente) || 0
      };
    });
  }

  const territorio = { lista: [], sessoes: null, conversao: null };
  if (terr.status === "fulfilled") {
    territorio.lista = terr.value.map(r => ({
      nome: r.nome || "(sem UF)",
      pct: Number(r.pct) || 0,
      receita: Number(r.receita) || 0
    }));
  }

  const erros = [prods, peds, terr]
    .filter(r => r.status === "rejected")
    .map(r => String(r.reason && r.reason.message ? r.reason.message : r.reason));

  return { porCanal, territorio, erros };
}

// ---------------------------------------------------------------------------
//  HANDLER
// ---------------------------------------------------------------------------
exports.handler = async function (event) {
  const qs = (event && event.queryStringParameters) || {};

  // ?schema=1 -> lista tabelas e colunas reais do dataset (para ajustar as queries)
  if (qs.schema === "1") {
    try {
      const bq = clienteBQ();
      const [rows] = await bq.query({ query: Q.schema, location: "US" });
      const porTabela = {};
      rows.forEach(r => {
        porTabela[r.table_name] = porTabela[r.table_name] || [];
        porTabela[r.table_name].push(`${r.column_name} (${r.data_type})`);
      });
      return json(200, { tabelas: Object.keys(porTabela).length, schema: porTabela });
    } catch (e) {
      return json(500, { error: "Falha ao ler schema", detail: String(e.message || e) });
    }
  }

  const aba = qs.mes || process.env.MES_ABA || ABAS_POR_MES[new Date().getMonth()];
  const mesIdx = ABAS_POR_MES.indexOf(aba) >= 0
    ? ABAS_POR_MES.indexOf(aba) : new Date().getMonth();

  // As duas fontes em paralelo e independentes: uma falhar não derruba a outra
  const [sp, bq] = await Promise.allSettled([ lerPlanilha(aba, mesIdx), lerBigQuery() ]);

  const avisos = [];

  if (sp.status === "rejected") {
    // Sem a planilha não há painel — hero e canais vêm todos dela
    return json(500, {
      error: "Falha ao ler a planilha (SharePoint)",
      detail: String(sp.reason && sp.reason.message ? sp.reason.message : sp.reason)
    });
  }

  const { hero, canais, conversao } = sp.value;

  if (bq.status === "fulfilled") {
    canais.forEach(c => {
      const extra = bq.value.porCanal[c.nome];
      if (extra) Object.assign(c, extra);
    });
    if (bq.value.erros.length) avisos.push(...bq.value.erros);
    var territorio = bq.value.territorio;
  } else {
    canais.forEach(c => {
      c.produtos = []; c.pedidos = null; c.ticket_medio = null; c.maior_cliente = null;
    });
    var territorio = { lista: [], sessoes: null, conversao: null };
    avisos.push("BigQuery: " + String(bq.reason && bq.reason.message ? bq.reason.message : bq.reason));
  }

  // Conversão do site: a planilha tem o número; usamos ela como fonte
  if (territorio.conversao == null && conversao.b2c) {
    territorio.conversao = conversao.b2c.atual;
  }

  return json(200, {
    atualizado_em: new Date().toISOString(),
    mes_vigente: aba,
    hero, canais, territorio, conversao,
    fonte_rodape: "Meta/realizado: planilha GERAL 2026 (SharePoint) · " +
                  "Produtos/território: BigQuery (" + PROJETO_BQ + ")",
    avisos: avisos.length ? avisos : undefined
  });
};

function json(status, body) {
  return {
    statusCode: status,
    headers: {
      "Content-Type": "application/json",
      "Cache-Control": "no-cache, no-store, must-revalidate"
    },
    body: JSON.stringify(body)
  };
}
