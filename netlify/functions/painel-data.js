// netlify/functions/painel-data.js  —  MOCK (dados reais congelados)
//
// Snapshot extraído da planilha GERAL 2026.xlsx real, aba AGOSTO.
// Os valores NÃO atualizam quando a planilha muda — servem para validar
// o design/layout antes de plugar o SharePoint.
//
// Quando as credenciais do Azure estiverem prontas, substitua o conteúdo
// deste arquivo pelo de painel-data-REAL.js.

exports.handler = async function (event) {
  const data = {
    "atualizado_em": "2026-08-18T16:53:21.951Z",
    "mes_vigente": "AGOSTO",
    "fonte": "SharePoint · GERAL 2026.xlsx · aba AGOSTO",
    "hero": {
        "titulo": "Total Geral · AGOSTO",
        "valor": 5140009.9399999995,
        "meta": 10180000,
        "pct_atingido": 50.491256777996064,
        "projecao": 9562193.99604278,
        "projecao_pct": 93.93117874305285
    },
    "dias": {
        "b2b": {
            "total": 21,
            "trabalhados": 11,
            "faltantes": 10,
            "pct_decorrido": 52.38095238095239
        },
        "b2c": {
            "total": 31,
            "trabalhados": 17,
            "faltantes": 14,
            "pct_decorrido": 54.83870967741935
        }
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
    "rankings": {
        "revenda": [
            {
                "nome": "Mariana",
                "valor": 1723007,
                "meta": 3570000,
                "pct_meta": 48.26350140056022,
                "media_dia": 156637,
                "projecao": 3289377
            },
            {
                "nome": "Maria Eduarda",
                "valor": 73425,
                "meta": 160000,
                "pct_meta": 45.890625,
                "media_dia": 6675,
                "projecao": 140175
            },
            {
                "nome": "Priscila",
                "valor": 10267,
                "meta": 85000,
                "pct_meta": 12.078823529411764,
                "media_dia": 933.3636363636364,
                "projecao": 19600.636363636364
            },
            {
                "nome": "Julio",
                "valor": 0,
                "meta": 85000,
                "pct_meta": 0,
                "media_dia": 0,
                "projecao": 0
            }
        ],
        "corporativo": [
            {
                "nome": "Gabriel",
                "valor": 714371,
                "meta": 420000,
                "pct_meta": 170.08833333333334,
                "media_dia": 64942.818181818184,
                "projecao": 1363799.1818181819
            },
            {
                "nome": "Kamila TVA",
                "valor": 308747,
                "meta": 400000,
                "pct_meta": 77.18675,
                "media_dia": 28067.909090909092,
                "projecao": 589426.0909090909
            },
            {
                "nome": "Raquel",
                "valor": 155990,
                "meta": 530000,
                "pct_meta": 29.432075471698116,
                "media_dia": 14180.90909090909,
                "projecao": 297799.09090909094
            },
            {
                "nome": "Bianca",
                "valor": 135268,
                "meta": 420000,
                "pct_meta": 32.20666666666666,
                "media_dia": 12297.09090909091,
                "projecao": 258238.90909090912
            }
        ],
        "sdrs": [
            {
                "nome": "Isaac",
                "valor": 736968.6,
                "meta": 380000,
                "pct_meta": 193.9391052631579,
                "media_dia": 43351.09411764706,
                "projecao": 1170479.5411764705
            },
            {
                "nome": "Beatriz",
                "valor": 82774.38,
                "meta": 380000,
                "pct_meta": 21.78273157894737,
                "media_dia": 7524.943636363637,
                "projecao": 158023.81636363638
            },
            {
                "nome": "Lhai",
                "valor": 46216.69,
                "meta": 380000,
                "pct_meta": 12.162286842105264,
                "media_dia": 2718.628823529412,
                "projecao": 73402.97823529413
            },
            {
                "nome": "Paola",
                "valor": 33902,
                "meta": 280000,
                "pct_meta": 12.107857142857142,
                "media_dia": 3082,
                "projecao": 64722
            }
        ],
        "sdr_geral": {
            "meta": 1420000,
            "vendido": 899861.67,
            "pct": 75.96137456140352
        }
    },
    "digital_detalhe": [
        {
            "nome": "B2C | IN",
            "valor": 1404595.07,
            "meta": 2930000,
            "pct_meta": 47.93839829351536,
            "media_dia": 82623.2394117647,
            "projecao": 2561320.4217647063
        },
        {
            "nome": "B2C | OUT",
            "valor": 172047.58,
            "meta": 400000,
            "pct_meta": 43.011894999999996,
            "media_dia": 10120.44588235294,
            "projecao": 273252.0388235294
        },
        {
            "nome": "Larissa TVA",
            "valor": 159685.37,
            "meta": 370000,
            "pct_meta": 43.15820810810811,
            "media_dia": 14516.851818181818,
            "projecao": 304853.8881818182
        },
        {
            "nome": "Alice TVA",
            "valor": 126782.83,
            "meta": 200000,
            "pct_meta": 63.391414999999995,
            "media_dia": 11525.711818181819,
            "projecao": 242039.94818181818
        }
    ],
    "conversao": {
        "b2c": {
            "meta": 0.54,
            "atual": 0.28
        },
        "corp": {
            "meta": 0.55,
            "atual": 0.5
        },
        "cac_corp_meta": 10
    }
};

  data.atualizado_em = new Date().toISOString();
  data.__aviso = "snapshot congelado - nao ao vivo";

  return {
    statusCode: 200,
    headers: {
      "Content-Type": "application/json",
      "Cache-Control": "no-cache, no-store, must-revalidate"
    },
    body: JSON.stringify(data)
  };
};
