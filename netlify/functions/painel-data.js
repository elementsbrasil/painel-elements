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
//
//  ---------------------------------------------------------------------------
//  REVISÃO 2026-08 — o que mudou e por quê
//  ---------------------------------------------------------------------------
//  1) TAXA DE CONVERSÃO ZERADA (Digital mostrava 0% de conversão e 0% de meta).
//     Causa: as células de taxa (M20/M21 = B2C, N20/N21 = Corporativo) estão
//     gravadas como TEXTO ("0.54%") mas com formato de célula 0%. Basta alguém
//     redigitar com vírgula ("0,54%") que o Excel converte para NÚMERO (0,0054)
//     — e o parser antigo (num()) devolvia 0,0054, que arredondado vira 0,00%.
//     Correção: pctTaxa() aceita os três formatos ("0.54%", "0,54%" e 0,0054) e
//     as células são localizadas pelo CABEÇALHO ("Taxa conv B2c site" / "Taxa
//     conv corp site"), não por endereço fixo — se alguém mexer nas linhas da
//     planilha, o painel continua puxando.
//  2) FALTA POR DIA por canal (novo bloco pedido na linha 4 / linha 3 compacta).
//     Usa a mesma conta da planilha: falta ÷ dias faltantes (E4 no b2b, E8 no
//     b2c). Confere com E19 / E29 / E35 da planilha.
//  3) VENDEDORES: passam a trazer meta, pct_meta e bateu_meta, e o teto subiu
//     de 8 para 24 nomes (o compacto exibe a lista COMPLETA; o index continua
//     cortando no próprio front, para não estourar o quadrante).
// ============================================================================

const GRAPH = "https://graph.microsoft.com/v1.0";

// Nomes reais das abas em uso na planilha (não são os meses completos)
const ABAS_POR_MES = [
  "JAN", "FEV", "MAR", "ABRIL", "MAIO", "JUNHO",
  "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
];

// tipoDias define qual contagem de dias trabalhados usar no cálculo de
// "média de vendas por dia" e de "falta por dia" de cada canal:
//   b2b -> E2 dias úteis · E3 trabalhados úteis · E4 faltantes
//   b2c -> E6 dias       · E7 trabalhados       · E8 faltantes
const CANAIS = [
  { nome: "Revenda",     codcencus: 10102001, linhaCanal: 3,  vendIni: 13, vendFim: 18, histCols: ["G","H"], tipoDias: "b2b" },
  { nome: "Corporativo", codcencus: 10102002, linhaCanal: 4,  vendIni: 21, vendFim: 28, histCols: ["L","M"], tipoDias: "b2b" },
  { nome: "Digital",     codcencus: 10102003, linhaCanal: 5,  vendIni: 31, vendFim: 34, histCols: ["Q","R"], tipoDias: "b2c" }
];

// Teto de nomes por canal. O compacto mostra a lista inteira; quem corta é o
// front do index (LIMITE_RANK lá), não esta função.
const MAX_VENDEDORES = 24;

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

// Taxa de conversão: a planilha é inconsistente nessas células (às vezes texto
// "0.54%", às vezes texto "0,54%", às vezes número 0,0054 com formato 0%).
// Regra: se veio com "%", o número já está em pontos percentuais. Se veio como
// número puro <= 1, é fração e vira %. Acima de 1 já está em %.
function pctTaxa(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v <= 1 ? v * 100 : v;
  const s = String(v).trim();
  const n = num(s);
  if (s.indexOf("%") >= 0) return n;      // "0,54%" -> 0.54
  return n <= 1 ? n * 100 : n;            // "0,0054" -> 0.54
}

function cel(m, linha, col) {
  if (!m || !m[linha]) return null;
  const v = m[linha][col];
  return v === undefined ? null : v;
}

// Uma string que é só número/percentual não serve como rótulo de linha
function ehRotulo(v) {
  if (typeof v !== "string") return false;
  const s = v.trim();
  if (!s) return false;
  return !/^[\d.,\s%R$+-]+$/.test(s);
}

// ---------------------------------------------------------------------------
//  Localização das taxas de conversão pelo cabeçalho (à prova de mudança de
//  linha na planilha). Cai no endereço fixo se o cabeçalho sumir.
// ---------------------------------------------------------------------------
function acharCabecalho(m, regex) {
  for (let r = 0; r < m.length; r++) {
    const linha = m[r] || [];
    for (let c = 0; c < linha.length; c++) {
      const v = linha[c];
      if (typeof v === "string" && regex.test(v.trim())) return { r, c };
    }
  }
  return null;
}

function lerTaxa(m, regexCabecalho, colFallback, linhaFallback) {
  const h = acharCabecalho(m, regexCabecalho);
  const col  = h ? h.c : colFallback;
  const base = h ? h.r : linhaFallback;   // linha do cabeçalho (0-based)

  let meta = null, atual = null;
  for (let r = base + 1; r <= base + 6 && r < m.length; r++) {
    const linha = m[r] || [];
    const val = linha[col];
    if (val === null || val === undefined || val === "") continue;

    // rótulo da linha = primeira string "de verdade" à esquerda da coluna
    let rot = "";
    for (let c = col - 1; c >= 0 && c >= col - 4; c--) {
      if (ehRotulo(linha[c])) { rot = String(linha[c]).trim().toLowerCase(); break; }
    }

    if (/^meta/.test(rot))       { if (meta  === null) meta  = pctTaxa(val); }
    else if (/^atual/.test(rot)) { if (atual === null) atual = pctTaxa(val); }
    else if (meta  === null)     { meta  = pctTaxa(val); }
    else if (atual === null)     { atual = pctTaxa(val); }
  }
  return { meta: meta || 0, atual: atual || 0 };
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
// (aparece com R$ 0). Agora devolve também meta, pct_meta e bateu_meta —
// o painel mostra "vendido / meta · %" e destaca quem já bateu.
function vendedores(m, ini, fim) {
  const out = [];
  for (let r = ini; r <= fim; r++) {
    const i = r - 1;                       // planilha 1-based -> range 0-based
    const nome = String(cel(m, i, 0) ?? "").trim();
    if (!nome || nome.toUpperCase() === "TOTAL") continue;
    const meta = num(cel(m, i, 1));        // col B
    const vend = num(cel(m, i, 2));        // col C
    if (!meta && !vend) continue;
    const pct = meta > 0 ? (vend / meta) * 100 : 0;
    out.push({
      nome,
      valor: vend,
      meta,
      pct_meta: pct,
      bateu_meta: meta > 0 && vend >= meta
    });
  }
  out.sort((a, b) => b.valor - a.valor);
  return out.slice(0, MAX_VENDEDORES);
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
    falta: num(cel(mMes, 4, 1)),               // B5  Faltam
    projecao: num(cel(mMes, 5, 1)),            // B6  Projeção
    projecao_pct: pctFrac(cel(mMes, 6, 1))     // B7  Projeção %
  };
  hero.alcance_vs_previsto = hero.projecao_pct;

  // Calendário do mês — base da "média de vendas por dia" e da "falta por dia".
  //   b2b: E2 úteis · E3 trabalhados úteis · E4 faltantes
  //   b2c: E6 dias   · E7 trabalhados       · E8 faltantes
  const dias = {
    b2b: {
      total:       num(cel(mMes, 1, 4)),   // E2
      trabalhados: num(cel(mMes, 2, 4)),   // E3
      faltantes:   num(cel(mMes, 3, 4))    // E4
    },
    b2c: {
      total:       num(cel(mMes, 5, 4)),   // E6
      trabalhados: num(cel(mMes, 6, 4)),   // E7
      faltantes:   num(cel(mMes, 7, 4))    // E8
    }
  };

  const canais = CANAIS.map(cfg => {
    const i = cfg.linhaCanal - 1;
    const realizado = num(cel(mMes, i, 8));    // col I
    const falta     = num(cel(mMes, i, 13));   // col N
    const d         = cfg.tipoDias === "b2c" ? dias.b2c : dias.b2b;
    return {
      nome: cfg.nome,
      realizado,
      projetado:    num(cel(mMes, i, 9)),      // col J
      projecao_pct: pctFrac(cel(mMes, i, 10)), // col K
      meta:         num(cel(mMes, i, 11)),     // col L
      pct_meta:     pctFrac(cel(mMes, i, 12)), // col M
      falta,
      // média já realizada por dia trabalhado
      media_dia:    d.trabalhados > 0 ? realizado / d.trabalhados : 0,
      // quanto precisa vender por dia no que resta do mês (= E19/E29/E35)
      falta_dia:    d.faltantes   > 0 ? Math.max(0, falta) / d.faltantes : 0,
      dias:         { total: d.total, trabalhados: d.trabalhados, faltantes: d.faltantes },
      historico:    historico(mHist, cfg.histCols[0], cfg.histCols[1], mesIdx),
      vendedores:   vendedores(mMes, cfg.vendIni, cfg.vendFim)
    };
  });
  canais.forEach(c => { c.alcance_vs_previsto = c.projecao_pct; });

  // Taxas de conversão do site — localizadas pelo cabeçalho, com fallback nos
  // endereços históricos (M20/M21 para B2C, N20/N21 para Corporativo).
  const conversao = {
    b2c:  lerTaxa(mMes, /taxa\s*conv.*b2c/i,  12, 18),
    corp: lerTaxa(mMes, /taxa\s*conv.*corp/i, 13, 18)
  };

  return { hero, canais, conversao, dias };
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

  const { hero, canais, conversao, dias } = sp;

  return json(200, {
    atualizado_em: new Date().toISOString(),
    mes_vigente: aba,
    hero, canais, conversao, dias,
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
