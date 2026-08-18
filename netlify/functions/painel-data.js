// netlify/functions/painel-data.js  —  VERSÃO MOCK (dados reais congelados)
//
// Estes valores foram extraídos da planilha GERAL 2026.xlsx real (aba AGOSTO),
// mas estão CONGELADOS no código — não atualizam quando a planilha muda.
//
// Serve para: validar o design/layout antes de plugar o SharePoint.
// Quando as credenciais do Azure estiverem prontas, substitua o conteúdo
// deste arquivo pelo de painel-data-REAL.js.

exports.handler = async function (event) {
  const data = {
      "atualizado_em": "2026-08-18T16:14:50.284Z",
      "mes_vigente": "AGOSTO",
      "hero": {
            "titulo": "Total Geral · AGOSTO",
            "valor": 5140009.9399999995,
            "meta": 10180000,
            "pct_atingido": 50.491256777996064,
            "projecao": 9562193.99604278,
            "projecao_pct": 93.93117874305285,
            "fonte": "SharePoint · GERAL 2026.xlsx · aba AGOSTO"
      },
      "canais": [
            {
                  "nome": "Revenda",
                  "realizado": 1806699,
                  "projetado": 3449152.6363636362,
                  "projecao_pct": 88.43981118881119,
                  "meta": 3900000,
                  "pct_meta": 46.32561538461539,
                  "falta": 2093301,
                  "historico": [
                        {
                              "mes": "Jan",
                              "atingido": 2516418.0700000003,
                              "meta": 4000000
                        },
                        {
                              "mes": "Fev",
                              "atingido": 5540309.9,
                              "meta": 5950000
                        },
                        {
                              "mes": "Mar",
                              "atingido": 1980180.8399999999,
                              "meta": 7000000
                        },
                        {
                              "mes": "Abr",
                              "atingido": 4813731.73,
                              "meta": 5300000
                        },
                        {
                              "mes": "Mai",
                              "atingido": 2987584.61,
                              "meta": 4460000
                        },
                        {
                              "mes": "Jun",
                              "atingido": 4748093.709999999,
                              "meta": 4000000
                        },
                        {
                              "mes": "Jul",
                              "atingido": 3588291.3099999996,
                              "meta": 4450000
                        },
                        {
                              "mes": "Ago",
                              "atingido": 1806699,
                              "meta": 3900000
                        }
                  ]
            },
            {
                  "nome": "Corporativo",
                  "realizado": 1470200.09,
                  "projetado": 2731575.062727273,
                  "projecao_pct": 114.77206145912913,
                  "meta": 2380000,
                  "pct_meta": 61.773113025210094,
                  "falta": 909799.9099999999,
                  "historico": [
                        {
                              "mes": "Jan",
                              "atingido": 1089742.43,
                              "meta": 1300000
                        },
                        {
                              "mes": "Fev",
                              "atingido": 1450326.8499999999,
                              "meta": 1400000
                        },
                        {
                              "mes": "Mar",
                              "atingido": 2437123.56,
                              "meta": 2030000
                        },
                        {
                              "mes": "Abr",
                              "atingido": 1356785.45,
                              "meta": 1800000
                        },
                        {
                              "mes": "Mai",
                              "atingido": 1566736.74,
                              "meta": 2100000
                        },
                        {
                              "mes": "Jun",
                              "atingido": 1395788.7999999998,
                              "meta": 2100000
                        },
                        {
                              "mes": "Jul",
                              "atingido": 2770592.37,
                              "meta": 2100000
                        },
                        {
                              "mes": "Ago",
                              "atingido": 1470200.09,
                              "meta": 2380000
                        }
                  ]
            },
            {
                  "nome": "Digital",
                  "realizado": 1863110.85,
                  "projetado": 3381466.296951872,
                  "projecao_pct": 86.70426402440697,
                  "meta": 3900000,
                  "pct_meta": 47.77207307692308,
                  "falta": 2036889.15,
                  "historico": [
                        {
                              "mes": "Jan",
                              "atingido": 3520407.94,
                              "meta": 4000000
                        },
                        {
                              "mes": "Fev",
                              "atingido": 3954318.2499999995,
                              "meta": 5000000
                        },
                        {
                              "mes": "Mar",
                              "atingido": 4744006.5,
                              "meta": 12000000
                        },
                        {
                              "mes": "Abr",
                              "atingido": 3042783.53,
                              "meta": 5038790.4
                        },
                        {
                              "mes": "Mai",
                              "atingido": 2728402.98,
                              "meta": 3450000
                        },
                        {
                              "mes": "Jun",
                              "atingido": 2670531.6100000003,
                              "meta": 3500000
                        },
                        {
                              "mes": "Jul",
                              "atingido": 3163359.7300000004,
                              "meta": 3800000
                        },
                        {
                              "mes": "Ago",
                              "atingido": 1863110.85,
                              "meta": 3900000
                        }
                  ]
            }
      ],
      "ranking_vendedores": {
            "titulo": "Vendedores · AGOSTO",
            "clientes_novos": {
                  "valor": 0,
                  "total": 0
            },
            "lista": [
                  {
                        "nome": "Mariana",
                        "valor": 1723007
                  },
                  {
                        "nome": "Gabriel",
                        "valor": 714371
                  },
                  {
                        "nome": "Kamila TVA",
                        "valor": 308747
                  },
                  {
                        "nome": "Raquel",
                        "valor": 155990
                  }
            ],
            "por_canal": {
                  "revenda": [
                        {
                              "nome": "Mariana",
                              "valor": 1723007
                        },
                        {
                              "nome": "Maria Eduarda",
                              "valor": 73425
                        },
                        {
                              "nome": "Priscila",
                              "valor": 10267
                        }
                  ],
                  "corporativo": [
                        {
                              "nome": "Gabriel",
                              "valor": 714371
                        },
                        {
                              "nome": "Kamila TVA",
                              "valor": 308747
                        },
                        {
                              "nome": "Raquel",
                              "valor": 155990
                        },
                        {
                              "nome": "Bianca",
                              "valor": 135268
                        }
                  ]
            }
      },
      "historico_geral": [
            {
                  "mes": "Jan",
                  "atingido": 7612063.34,
                  "meta": 9771000
            },
            {
                  "mes": "Fev",
                  "atingido": 10944955,
                  "meta": 12350000
            },
            {
                  "mes": "Mar",
                  "atingido": 9161310.9,
                  "meta": 21030000
            },
            {
                  "mes": "Abr",
                  "atingido": 9213300.71,
                  "meta": 12138790.4
            },
            {
                  "mes": "Mai",
                  "atingido": 7282724.33,
                  "meta": 10010000
            },
            {
                  "mes": "Jun",
                  "atingido": 8814414.12,
                  "meta": 9600000
            },
            {
                  "mes": "Jul",
                  "atingido": 9522243.41,
                  "meta": 10350000
            },
            {
                  "mes": "Ago",
                  "atingido": 5140009.9399999995,
                  "meta": 10180000
            }
      ],
      "territorio": {
            "titulo": "Território",
            "fonte": "BigQuery · pendente",
            "lista": [],
            "rodape": {
                  "valor": 0,
                  "rotulo": "Sessões · mês"
            }
      }
};

  data.atualizado_em = new Date().toISOString();
  data.__aviso = "DADOS CONGELADOS - snapshot da planilha, nao ao vivo";

  return {
    statusCode: 200,
    headers: {
      "Content-Type": "application/json",
      "Cache-Control": "no-cache, no-store, must-revalidate"
    },
    body: JSON.stringify(data)
  };
};
