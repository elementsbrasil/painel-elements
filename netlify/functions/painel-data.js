// ============================================================================
//  PAINEL COMERCIAL — função de dados
//
//  Roda no servidor da Netlify. Consulta a planilha (SharePoint / Microsoft
//  Graph) a cada chamada e devolve um JSON no formato que o index.html e o
//  compacto.html esperam.
//
//    SharePoint (Microsoft Graph)  -> hero, canais, evolução, vendedores,
//                                     conversão do site, ritmo do canal
//
//  Fonte única de dados: a planilha "GERAL 2026". Não há mais consulta a
//  BigQuery — os blocos que antes vinham de lá (território e top produtos)
//  agora usam indicadores calculados a partir da própria planilha.
//
//  CREDENCIAIS: lidas de variáveis de ambiente. Nunca no código, nunca no
//  GitHub. Configurar em Netlify > Site configuration > Environment variables:
//
//    MS_CLIENT_ID              (Azure > App registration > Overview)
//    MS_TENANT_ID              (idem)
//    MS_CLIENT_SECRET          (Azure > Certificates & secrets)  [EXPIRA!]
//    SHAREPOINT_SITE_ID        (Graph Explorer)
//    SHAREPOINT_ITEM_ID        (Graph Explorer)
//    MES_ABA                   (opcional: força uma aba, ex. "AGOSTO")
//
//  DEGRADAÇÃO: se a planilha falhar, não há painel — hero e canais vêm todos
//  dela. Nesse caso a função devolve erro 500 e o front mostra a mensagem de
//  "não foi possível carregar os dados" (regra 7 da spec original).
// ============================================================================

const GRAPH = "https://graph.microsoft.com/v1.0";

// Nomes reais das abas em uso na planilha (não são os meses completos)
const ABAS_POR_MES = [
  "JAN", "FEV", "MAR", "ABRIL", "MAIO", "JUNHO",
  "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
];

// tipoDias define qual contagem de dias trabalhados usar no cálculo de
// "média de vendas por dia" de cada canal: b2b usa a célula E3 (Trabalhados
// uteis), b2c usa a célula E7 (Trabalhados) — ambas já existem na planilha.
const CANAIS = [
  { nome: "Revenda",     codcencus: 10102001, linhaCanal: 3,  vendIni: 13, vendFim: 18, histCols: ["G","H"], tipoDias: "b2b" },
  { nome: "Corporativo", codcencus: 10102002, linhaCanal: 4,  vendIni: 21, vendFim: 28, histCols: ["L","M"], tipoDias: "b2b" },
  { nome: "Digital",     codcencus: 10102003, linhaCanal: 5,  vendIni: 31, vendFim: 34, histCols: ["Q","R"], tipoDias: "b2c" }
];

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

// Vendedores/composição de um bloco. Mantém quem tem meta e vendeu 0
// (aparece com R$ 0). Devolve até 8 linhas: as 4 primeiras alimentam o
// ranking da linha 3, as 4 seguintes (se existirem) ficam disponíveis para
// quem precisar de mais detalhe.
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
  return out.slice(0, 8);
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

  // Dias trabalhados no mês até agora — usados para calcular a "média de
  // vendas por dia" de cada canal. E3 = Trabalhados uteis (b2b);
  // E7 = Trabalhados (b2c). Presentes nas mesmas células em todas as abas.
  const diasTrabalhados = {
    b2b: num(cel(mMes, 2, 4)),   // E3
    b2c: num(cel(mMes, 6, 4))    // E7
  };

  const canais = CANAIS.map(cfg => {
    const i = cfg.linhaCanal - 1;
    const realizado = num(cel(mMes, i, 8));    // col I
    const falta     = num(cel(mMes, i, 13));   // col N
    const dias      = cfg.tipoDias === "b2c" ? diasTrabalhados.b2c : diasTrabalhados.b2b;
    return {
      nome: cfg.nome,
      realizado,
      projetado:    num(cel(mMes, i, 9)),      // col J
      projecao_pct: pctFrac(cel(mMes, i, 10)), // col K
      meta:         num(cel(mMes, i, 11)),     // col L
      pct_meta:     pctFrac(cel(mMes, i, 12)), // col M
      falta,
      media_dia:    dias > 0 ? realizado / dias : 0,
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
//  HANDLER
// ---------------------------------------------------------------------------
exports.handler = async function (event) {
  const qs = (event && event.queryStringParameters) || {};

  const aba = qs.mes || process.env.MES_ABA || ABAS_POR_MES[new Date().getMonth()];
  const mesIdx = ABAS_POR_MES.indexOf(aba) >= 0
    ? ABAS_POR_MES.indexOf(aba) : new Date().getMonth();

  let sp;
  try {
    sp = await lerPlanilha(aba, mesIdx);
  } catch (e) {
    // Sem a planilha não há painel — hero e canais vêm todos dela
    return json(500, {
      error: "Falha ao ler a planilha (SharePoint)",
      detail: String(e && e.message ? e.message : e)
    });
  }

  const { hero, canais, conversao } = sp;

  return json(200, {
    atualizado_em: new Date().toISOString(),
    mes_vigente: aba,
    hero, canais, conversao,
    fonte_rodape: "Fonte única: planilha GERAL 2026 (SharePoint)"
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
