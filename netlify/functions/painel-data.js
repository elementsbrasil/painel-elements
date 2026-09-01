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
//  ===========================================================================
//  PRINCÍPIO DESTA VERSÃO: NADA DE ENDEREÇO FIXO DE LINHA
//  ---------------------------------------------------------------------------
//  A planilha é editada por gente todo mês — linha inserida, vendedor novo,
//  bloco que sobe ou desce. Endereço fixo ("Corporativo = linhas 21 a 28")
//  quebra silenciosamente: o painel continua no ar, com números plausíveis e
//  ERRADOS. Foi o que aconteceu quando os blocos subiram 3 linhas — o
//  Corporativo passou a mostrar "B2C | IN" (que é do Digital) e sumiu com a
//  Bianca e o Gabriel; o Digital ficou só com a Alice.
//
//  Agora TUDO é localizado por RÓTULO:
//    · hero          -> textos da coluna A ("META", "Vendido", "% Atingido"...)
//    · dias do mês   -> textos da coluna D ("Dias uteis (b2b)", "Trabalhados"...)
//    · resumo canal  -> nome do canal na coluna H
//    · vendedores    -> o bloco começa na linha de cabeçalho que tem
//                       B="Meta" e C="Vendido", e termina na linha de TOTAL
//                       (nome vazio + valores preenchidos) ou no próximo bloco
//    · conversão     -> cabeçalhos "Taxa conv B2c site" / "Taxa conv corp site"
//  Cada busca tem fallback para o endereço histórico, então se algum rótulo
//  for renomeado o painel degrada para o comportamento antigo em vez de zerar.
//  ===========================================================================
// ============================================================================

const GRAPH = "https://graph.microsoft.com/v1.0";

// Nomes reais das abas em uso na planilha (não são os meses completos)
const ABAS_POR_MES = [
  "JAN", "FEV", "MAR", "ABRIL", "MAIO", "JUNHO",
  "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
];

// Faixa lida da aba do mês. Sobra folga de propósito: o bloco de vendedores
// pode crescer com nomes novos sem precisar mexer aqui.
const RANGE_MES = "A1:N60";

// tipoDias define qual contagem de dias usar no cálculo de "média de vendas
// por dia" e de "falta por dia" de cada canal:
//   b2b -> dias úteis · trabalhados úteis · faltantes
//   b2c -> dias       · trabalhados       · faltantes
// regexBloco/regexResumo localizam o canal pelo texto, não pela linha.
// linhaCanal/vendIni/vendFim ficam só como rede de segurança (fallback).
const CANAIS = [
  { nome: "Revenda",     codcencus: 10102001, regexBloco: /revenda/i,     linhaCanal: 3, vendIni: 13, vendFim: 18, histCols: ["G","H"], tipoDias: "b2b" },
  { nome: "Corporativo", codcencus: 10102002, regexBloco: /corporativ/i,  linhaCanal: 4, vendIni: 21, vendFim: 28, histCols: ["L","M"], tipoDias: "b2b" },
  { nome: "Digital",     codcencus: 10102003, regexBloco: /digital/i, linhaCanal: 5, vendIni: 31, vendFim: 34, histCols: ["Q","R"], tipoDias: "b2c" }
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
function txt(m, linha, col) {
  const v = cel(m, linha, col);
  return typeof v === "string" ? v.trim() : (v === null || v === undefined ? "" : String(v));
}

// Uma string que é só número/percentual não serve como rótulo
function ehRotulo(v) {
  if (typeof v !== "string") return false;
  const s = v.trim();
  if (!s) return false;
  return !/^[\d.,\s%R$+-]+$/.test(s);
}

// ---------------------------------------------------------------------------
//  Busca por rótulo
// ---------------------------------------------------------------------------
// Primeira linha (a partir de `de`) cujo texto na coluna `col` casa com o regex
function linhaDoRotulo(m, col, regex, de) {
  for (let r = (de || 0); r < m.length; r++) {
    const v = cel(m, r, col);
    if (typeof v === "string" && regex.test(v.trim())) return r;
  }
  return -1;
}

// Valor numérico da coluna `colValor` na linha cujo rótulo (coluna `colRot`)
// casa com o regex. `fallbackLinha` é o índice histórico (0-based).
function porRotulo(m, colRot, regex, colValor, fallbackLinha, de) {
  const r = linhaDoRotulo(m, colRot, regex, de);
  if (r >= 0) return num(cel(m, r, colValor));
  return fallbackLinha != null ? num(cel(m, fallbackLinha, colValor)) : 0;
}

// Procura um cabeçalho em qualquer célula da matriz
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

// ---------------------------------------------------------------------------
//  Taxas de conversão (localizadas pelo cabeçalho, com fallback de endereço)
// ---------------------------------------------------------------------------
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
//  Blocos de vendedores
// ---------------------------------------------------------------------------
// Um bloco começa numa linha de cabeçalho com B="Meta" e C="Vendido"
// (ex.: A="REVENDA (B2B)" | B="Meta" | C="Vendido" | ...).
function blocosDeVendedores(m) {
  const blocos = [];
  for (let r = 0; r < m.length; r++) {
    const b = cel(m, r, 1), c = cel(m, r, 2);
    if (typeof b === "string" && typeof c === "string" &&
        b.trim().toLowerCase() === "meta" &&
        c.trim().toLowerCase() === "vendido") {
      blocos.push({ cab: r, titulo: txt(m, r, 0) });
    }
  }
  blocos.forEach((bl, i) => {
    bl.fim = (i + 1 < blocos.length) ? blocos[i + 1].cab : m.length;
  });
  return blocos;
}

// Lê as pessoas/linhas de um bloco. Para na linha de TOTAL do bloco — que é a
// que NÃO tem nome mas TEM meta/vendido preenchidos. Linhas totalmente vazias
// (espaçadores, que existiam no layout antigo) são puladas, não encerram.
function lerVendedores(m, ini, fim) {
  const out = [];
  for (let r = ini; r < fim && r < m.length; r++) {
    const nome = txt(m, r, 0);
    const meta = num(cel(m, r, 1));
    const vend = num(cel(m, r, 2));

    if (!nome) {
      if (meta || vend) break;   // linha de TOTAL -> fim do bloco
      continue;                  // linha vazia de espaçamento -> segue
    }
    if (/^total/i.test(nome)) break;
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

// ---------------------------------------------------------------------------
//  Leitura da aba do mês
// ---------------------------------------------------------------------------
function montar(mMes, mHist, mesIdx) {
  /* ---- hero: rótulos na coluna A, valores na coluna B ---- */
  const hero = {
    valor:        porRotulo(mMes, 0, /^vendido$/i,          1, 2),
    meta:         porRotulo(mMes, 0, /^meta$/i,             1, 1),
    pct_atingido: pctFrac(cel(mMes, linhaDoRotulo(mMes, 0, /^%\s*atingido/i) >= 0
                    ? linhaDoRotulo(mMes, 0, /^%\s*atingido/i) : 3, 1)),
    falta:        porRotulo(mMes, 0, /^faltam/i,            1, 4),
    projecao:     porRotulo(mMes, 0, /^proje[çc][ãa]o$/i,   1, 5)
  };
  const lProjPct = linhaDoRotulo(mMes, 0, /^proje[çc][ãa]o\s*%/i);
  hero.projecao_pct = pctFrac(cel(mMes, lProjPct >= 0 ? lProjPct : 6, 1));
  hero.alcance_vs_previsto = hero.projecao_pct;

  /* ---- calendário do mês: rótulos na coluna D, valores na coluna E ----
     O bloco b2b vem primeiro ("Dias uteis (b2b)" / Trabalhados / Faltantes) e
     o b2c logo abaixo ("Dias (b2c)" / Trabalhados / Faltantes). Como os dois
     repetem "Trabalhados" e "Faltantes", cada um é procurado A PARTIR da sua
     própria linha âncora — nunca do topo. */
  const ancB2b = linhaDoRotulo(mMes, 3, /dias\s*[úu]teis/i);
  const ancB2c = linhaDoRotulo(mMes, 3, /dias\s*\(?\s*b2c/i);
  const dias = {
    b2b: {
      total:       ancB2b >= 0 ? num(cel(mMes, ancB2b, 4)) : num(cel(mMes, 1, 4)),
      trabalhados: porRotulo(mMes, 3, /^trabalhados/i, 4, 2, ancB2b >= 0 ? ancB2b : 0),
      faltantes:   porRotulo(mMes, 3, /^faltantes/i,   4, 3, ancB2b >= 0 ? ancB2b : 0)
    },
    b2c: {
      total:       ancB2c >= 0 ? num(cel(mMes, ancB2c, 4)) : num(cel(mMes, 5, 4)),
      trabalhados: porRotulo(mMes, 3, /^trabalhados/i, 4, 6, ancB2c >= 0 ? ancB2c : 0),
      faltantes:   porRotulo(mMes, 3, /^faltantes/i,   4, 7, ancB2c >= 0 ? ancB2c : 0)
    }
  };

  /* ---- blocos de vendedores (Revenda / Corporativo / Digital / SDR) ---- */
  const blocos = blocosDeVendedores(mMes);

  const canais = CANAIS.map(cfg => {
    /* resumo do canal: linha em que a coluna H tem o nome do canal */
    let lin = linhaDoRotulo(mMes, 7, new RegExp("^" + cfg.nome + "$", "i"));
    if (lin < 0) lin = cfg.linhaCanal - 1;          // fallback histórico

    const realizado = num(cel(mMes, lin, 8));       // col I
    const falta     = num(cel(mMes, lin, 13));      // col N
    const d         = cfg.tipoDias === "b2c" ? dias.b2c : dias.b2b;

    /* vendedores: bloco cujo título casa com o canal */
    const bloco = blocos.find(b => cfg.regexBloco.test(b.titulo));
    const vendedores = bloco
      ? lerVendedores(mMes, bloco.cab + 1, bloco.fim)
      : lerVendedores(mMes, cfg.vendIni - 1, cfg.vendFim);   // fallback

    return {
      nome: cfg.nome,
      realizado,
      projetado:    num(cel(mMes, lin, 9)),         // col J
      projecao_pct: pctFrac(cel(mMes, lin, 10)),    // col K
      meta:         num(cel(mMes, lin, 11)),        // col L
      pct_meta:     pctFrac(cel(mMes, lin, 12)),    // col M
      falta,
      media_dia:    d.trabalhados > 0 ? realizado / d.trabalhados : 0,
      falta_dia:    d.faltantes   > 0 ? Math.max(0, falta) / d.faltantes : 0,
      dias:         { total: d.total, trabalhados: d.trabalhados, faltantes: d.faltantes },
      historico:    historico(mHist, cfg.histCols[0], cfg.histCols[1], mesIdx),
      vendedores,
      // diagnóstico: em que linha da planilha o bloco foi encontrado
      _origem:      bloco ? { bloco: bloco.titulo, linha: bloco.cab + 1 } : { bloco: "fallback", linha: cfg.vendIni }
    };
  });
  canais.forEach(c => { c.alcance_vs_previsto = c.projecao_pct; });

  const conversao = {
    b2c:  lerTaxa(mMes, /taxa\s*conv.*b2c/i,  12, 18),
    corp: lerTaxa(mMes, /taxa\s*conv.*corp/i, 13, 18)
  };

  return { hero, canais, conversao, dias };
}

async function lerPlanilha(aba, mesIdx) {
  const token = await tokenGraph();
  const [mMes, mHist] = await Promise.all([
    lerRange(token, aba, RANGE_MES),
    lerRange(token, "Dashboard", "A4:S15")
  ]);
  return montar(mMes, mHist, mesIdx);
}

// ---------------------------------------------------------------------------
//  HANDLER
// ---------------------------------------------------------------------------
// Mês pelo relógio de São Paulo. O servidor da Netlify roda em UTC, então
// new Date().getMonth() já virava o mês seguinte às 21h daqui.
function mesBR() {
  const f = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Sao_Paulo', month: 'numeric'
  });
  return parseInt(f.format(new Date()), 10) - 1;   // 0 = janeiro
}

exports.handler = async function (event) {
  const qs = (event && event.queryStringParameters) || {};

  const aba = qs.mes || process.env.MES_ABA || ABAS_POR_MES[mesBR()];
  const mesIdx = ABAS_POR_MES.indexOf(aba) >= 0
    ? ABAS_POR_MES.indexOf(aba) : mesBR();
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

// Exportado só para teste local (não usado pela Netlify)
