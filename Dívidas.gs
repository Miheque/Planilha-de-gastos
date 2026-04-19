//================================================================
// §C01 ── 1_config.gs — Configuração global
//================================================================

const CONFIG = {                                                    // §C01.01
  LINHA_INICIAL:   3,                                               // §C01.02
  LINHA_CABECALHO: 1,                                               // §C01.03
  LINHA_DADOS:     4,                                               // §C01.04
  EMAIL_ALERTA:    "?????@gmail.com",                             // §C01.05

  ABAS_PERMITIDAS: new Set([                                        // §C01.06
    "MARIA","PRISCILA","ISAEL","RAYLAN",
    "REYNAN","EU","SHOPEE","ISAEL SHOPEE",
    "CARTÃO M-PAGO","CRÉDITO M-PAGO","CARTÃO PAN"
  ]),

  COL: {                                                            // §C01.07
    DATA:          1,
    NOME:          2,
    DIA_VENC:      3,
    VALOR:         5,
    PARCELAS:      6,
    PRIMEIRA_PARC: 7
  },

  ABAS_AUXILIARES: new Set([                                        // §C01.08
    "SHOPEE", "ISAEL SHOPEE", "CARTÃO M-PAGO", "CRÉDITO M-PAGO", "CARTÃO PAN"
  ]),

  EMOJI_ALERTA: "⚠️",                                               // §C01.09

  MESES_PT: ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"], // §C01.10

  CORES: {                                                          // §C01.11
    "RAYLAN":   { bg: "#ff9900", fg: "#000000" },
    "MARIA":    { bg: "#ffff00", fg: "#000000" },
    "ISAEL":    { bg: "#ff0000", fg: "#000000" },
    "PRISCILA": { bg: "#800080", fg: "#000000" },
    "REYNAN":   { bg: "#00ffff", fg: "#000000" },
    "EU":       { bg: "#0000ff", fg: "#000000" }
  },

  CORES_EMAIL: {                                                    // §C01.12
    "RAYLAN":   "#ff9900",
    "MARIA":    "#ffff00",
    "ISAEL":    "#ff0000",
    "PRISCILA": "#800080",
    "REYNAN":   "#00ffff",
    "EU":       "#0000ff"
  }
};

const EMAILS_POR_ABA = {                                            // §C02
  "MARIA":    "?????@gmail.com",
  "PRISCILA": "?????@gmail.com",
  "ISAEL":    "?????@gmail.com",
  "RAYLAN":   "?????@gmail.com",
  "REYNAN":   "?????@gmail.com",
  "EU":       "?????@gmail.com"
};

const ABAS_NOTIFICACAO = new Set([                                  // §C03
  "MARIA", "PRISCILA", "ISAEL", "RAYLAN", "REYNAN", "EU"
]);

const MENSAGEM_PADRAO = {                                           // §C04
  introducao: "Você tem {textoDivida} com Michael no valor de R${total} que {vence} (dia {dia}):",
  fechamento: "Paga logo para não virar saudades!"
};

const PIX_CHAVE   = "??????????";                                  // §C05.01
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwGPmbXrY2g1TAg_Ep9eu2RuWkoeos5uCvOTgDIxGUvPUHs31q9wnKIno9Rv1Be8Z-h6Q/exec"; // §C05.02
const PIX_NOME    = "Michael";                                      // §C05.03

const MENSAGENS_PERSONALIZADAS = {                                  // §C06
  "MARIA":    { assunto: "⚠️ Maria – {TextoDivida} com Michael vencendo amanhã (dia {dia})",    saudacao: "Olá, Maria!"    },
  "PRISCILA": { assunto: "⚠️ Priscila – {TextoDivida} com Michael vencendo amanhã (dia {dia})", saudacao: "Olá, Priscila!" },
  "ISAEL":    { assunto: "⚠️ Isael – {TextoDivida} com Michael vencendo amanhã (dia {dia})",    saudacao: "Olá, Isael!"    },
  "RAYLAN":   { assunto: "⚠️ Raylan – {TextoDivida} com Michael vencendo amanhã (dia {dia})",   saudacao: "Olá, Raylan!"   },
  "REYNAN":   { assunto: "⚠️ Reynan – {TextoDivida} com Michael vencendo amanhã (dia {dia})",   saudacao: "Olá, Reynan!"   },
  "EU": {
    assunto:    "⚠️ Miheque – {Sua} {TextoDivida} que {vence} (dia {dia})",
    saudacao:   "Olá, Miheque!",
    introducao: "{Sua} {textoDivida} no valor de R${total} que {vence} (dia {dia}):"
  }
};

//================================================================
// §F01–§F19 ── 2_utils.gs
//================================================================

// §F01 ── getMensagem
function getMensagem(aba) {
  const cfg = MENSAGENS_PERSONALIZADAS[aba] || {};                  // §F01.01
  return {                                                          // §F01.02
    assunto:    cfg.assunto    || ("⚠️ " + aba + " – {TextoDivida} vencendo amanhã (dia {dia})"),
    saudacao:   cfg.saudacao   || "",
    introducao: cfg.introducao || MENSAGEM_PADRAO.introducao,
    fechamento: cfg.fechamento || MENSAGEM_PADRAO.fechamento
  };
}

// §F02 ── substituir
function substituir(template, vars) {
  return template.replace(/\{(\w+)\}/g, function(_, chave) {       // §F02.01
    return vars[chave] !== undefined ? vars[chave] : "{" + chave + "}";
  });
}

// §F03 ── getCor
function getCor(nome) {
  const n = nome.toUpperCase();                                     // §F03.01
  for (const chave in CONFIG.CORES) {                              // §F03.02
    if (n.includes(chave)) return CONFIG.CORES[chave];
  }
  return null;                                                      // §F03.03
}

// §F04 ── getCorEmail
function getCorEmail(aba) { return CONFIG.CORES_EMAIL[aba] || "#333333"; } // §F04

// §F05 ── marcarAlerta
function marcarAlerta(sheet, linha) {
  const cel = sheet.getRange(linha, CONFIG.COL.NOME);              // §F05.01
  const v   = (cel.getValue() || "").toString().trim();             // §F05.02
  if (!v.startsWith(CONFIG.EMOJI_ALERTA)) cel.setValue(CONFIG.EMOJI_ALERTA + " " + v); // §F05.03
}

// §F06 ── limparAlerta
function limparAlerta(sheet, linha) {
  const cel    = sheet.getRange(linha, CONFIG.COL.NOME);           // §F06.01
  const v      = (cel.getValue() || "").toString().trim();          // §F06.02
  const prefix = CONFIG.EMOJI_ALERTA + " ";                         // §F06.03
  if (v.startsWith(prefix)) cel.setValue(v.slice(prefix.length));  // §F06.04
}

// §F07 ── normalizarNome
function normalizarNome(sheet, linha, nome) {
  const prefix   = CONFIG.EMOJI_ALERTA + " ";                       // §F07.01
  const temEmoji = nome.startsWith(prefix);                         // §F07.02
  const base     = temEmoji ? nome.slice(prefix.length) : nome;    // §F07.03
  const maiusc   = (temEmoji ? prefix : "") + base.toUpperCase();  // §F07.04
  if (nome !== maiusc) sheet.getRange(linha, CONFIG.COL.NOME).setValue(maiusc); // §F07.05
}

// §F08 ── _extrairId
function _extrairId(valor) {
  const m = (valor || "").toString().match(/ID-\d+-[A-Z0-9]+/);    // §F08.01
  return m ? m[0] : "";                                             // §F08.02
}

// §F09 ── _garantirId
function _garantirId(sheet, linha) {
  const cel      = sheet.getRange(linha, CONFIG.COL.DATA);         // §F09.01
  const conteudo = (cel.getValue() || "").toString().trim();        // §F09.02
  const idExist  = _extrairId(conteudo);                            // §F09.03
  if (idExist) return idExist;                                      // §F09.04
  const agora  = new Date();                                        // §F09.05
  const dia    = String(agora.getDate()).padStart(2, "0");          // §F09.06
  const mes    = CONFIG.MESES_PT[agora.getMonth()];                 // §F09.07 ── MMM abreviado
  const ano    = agora.getFullYear();                               // §F09.08
  const novoId = "ID-" + agora.getTime() + "-" + Math.random().toString(36).substr(2,5).toUpperCase(); // §F09.09
  cel.setValue(conteudo ? conteudo + " | " + novoId : dia + "/" + mes + "/" + ano + " | " + novoId); // §F09.10
  return novoId;                                                    // §F09.11
}

// §F10 ── preencherData
function preencherData(sheet, linha, valor) {
  if (!valor) return;                                               // §F10.01
  const cel      = sheet.getRange(linha, CONFIG.COL.DATA);         // §F10.02
  const conteudo = (cel.getValue() || "").toString().trim();        // §F10.03
  if (conteudo) return;                                             // §F10.04
  const agora  = new Date();                                        // §F10.05
  const dia    = String(agora.getDate()).padStart(2, "0");          // §F10.06
  const mes    = CONFIG.MESES_PT[agora.getMonth()];                 // §F10.07 ── MMM abreviado
  const ano    = agora.getFullYear();                               // §F10.08
  const novoId = "ID-" + agora.getTime() + "-" + Math.random().toString(36).substr(2,5).toUpperCase(); // §F10.09
  cel.setValue(dia + "/" + mes + "/" + ano + " | " + novoId);      // §F10.10 ── dd/MMM/yyyy | ID
  cel.setNumberFormat("@");                                         // §F10.11
}

// §F11 ── formatarParcelas
function formatarParcelas(sheet, linha, valor, parcelasRaw) {
  let qtd = 0;                                                      // §F11.01
  if (typeof parcelasRaw === "string" && parcelasRaw.includes("×")) { // §F11.02
    qtd = parseInt(parcelasRaw.split("×")[0]) || 0;
  } else {
    qtd = parseInt(parcelasRaw) || 0;                               // §F11.03
  }
  if (qtd > 0 && valor) {                                          // §F11.04
    const vpp    = Math.abs(valor) / qtd;                           // §F11.05
    const vppFmt = vpp.toLocaleString("en-US", { minimumFractionDigits: 2 }); // §F11.06
    const texto  = qtd + "×R$" + vppFmt;                           // §F11.07
    if (texto !== parcelasRaw) sheet.getRange(linha, CONFIG.COL.PARCELAS).setValue(texto); // §F11.08
  }
}

// §F12 ── _aplicarFormatacaoGlobal
function _aplicarFormatacaoGlobal(sheet) {
  const rows = sheet.getMaxRows(), cols = sheet.getMaxColumns();    // §F12.01
  if (rows < 1 || cols < 1) return;                                // §F12.02
  const r = sheet.getRange(1, 1, rows, cols);                      // §F12.03
  r.setFontFamily("Arial"); r.setFontSize(12); r.setFontWeight("bold"); // §F12.04
  r.setHorizontalAlignment("left"); r.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // §F12.05
}

// §F13 ── gerarPixPayload
function gerarPixPayload(chave, nome, valor) {
  function emv(id, c) { return id + c.length.toString().padStart(2,"0") + c; } // §F13.01
  const chaveFmt = chave.startsWith("+") ? chave : "+55" + chave;  // §F13.02
  const pixData  = emv("00","BR.GOV.BCB.PIX") + emv("01", chaveFmt); // §F13.03
  let payload    =                                                  // §F13.04
    emv("00","01") + emv("26",pixData) + emv("52","0000") + emv("53","986") +
    emv("54",valor.toFixed(2)) + emv("58","BR") + emv("59",nome.substring(0,25)) +
    emv("60","SAO PAULO") + emv("62",emv("05","***")) + "6304";
  let crc = 0xFFFF;                                                 // §F13.05
  for (let i = 0; i < payload.length; i++) {                       // §F13.06
    crc ^= payload.charCodeAt(i) << 8;
    for (let b = 0; b < 8; b++) crc = (crc & 0x8000) ? (crc<<1)^0x1021 : crc<<1;
  }
  return payload + (crc & 0xFFFF).toString(16).toUpperCase().padStart(4,"0"); // §F13.07
}

// §F14 ── gerarLinkQR
function gerarLinkQR(chave, nome, valor) {
  return "https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=" + // §F14.01
    encodeURIComponent(gerarPixPayload(chave, nome, valor));
}

// §F15 ── _qtdParcelasDoPessoa
function _qtdParcelasDoPessoa(total, indicePessoa, qtdPessoas) {
  return Math.floor(total / qtdPessoas) + (indicePessoa < (total % qtdPessoas) ? 1 : 0); // §F15.01
}

// §F16 ── _fmtData
function _fmtData(val) {
  if (!val) return "";                                              // §F16.01
  if (val instanceof Date) {                                        // §F16.02
    if (isNaN(val.getTime())) return "";
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const s     = val.toString().trim();                              // §F16.03
  const semId = s.split("|")[0].trim();                             // §F16.04
  const m1    = semId.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);    // §F16.05
  if (m1) return m1[3]+"-"+m1[2].padStart(2,"0")+"-"+m1[1].padStart(2,"0"); // §F16.06
  const m2    = semId.match(/^(\d{1,2})\/(\w{3})\/(\d{4})/);      // §F16.07 ── dd/MMM/yyyy
  if (m2) {
    const mi = CONFIG.MESES_PT.indexOf(m2[2]);                     // §F16.08
    if (mi >= 0) return m2[3]+"-"+String(mi+1).padStart(2,"0")+"-"+m2[1].padStart(2,"0");
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(semId)) return semId.substring(0,10); // §F16.09
  return "";                                                        // §F16.10
}

// §F17 ── _keyMapeamento
function _keyMapeamento(abaDestino, nome, valor, data) {
  const id = _extrairId(data);                                      // §F17.01
  if (id) return ("lnk|"+abaDestino+"|"+id).substring(0,499);      // §F17.02
  const vf = typeof valor==="number" ? valor.toFixed(2) : (valor||"").toString().replace(/\|/g,"_"); // §F17.03
  const df = _fmtData(data);                                        // §F17.04
  const nf = nome.replace(/\|/g,"_").substring(0,80);              // §F17.05
  return ("lnk|"+abaDestino+"|"+nf+"|"+vf+"|"+df).substring(0,499); // §F17.06
}

// §F18 ── _salvarMapeamento
function _salvarMapeamento(abaDestino, linhaDestino, nome, valor, data) {
  PropertiesService.getScriptProperties()                           // §F18.01
    .setProperty(_keyMapeamento(abaDestino, nome, valor, data), String(linhaDestino));
}

// §F19 ── _corTextoContraste
function _corTextoContraste(hex) {
  const r   = parseInt(hex.slice(1,3),16);                          // §F19.01
  const g   = parseInt(hex.slice(3,5),16);                          // §F19.02
  const b   = parseInt(hex.slice(5,7),16);                          // §F19.03
  const lum = (0.299*r + 0.587*g + 0.114*b);                       // §F19.04
  return lum > 128 ? "#000000" : "#ffffff";                         // §F19.05
}

//================================================================
// §F19b ── _limparTracos ── remove "-" das colunas D em diante no rowFull
//================================================================
// DEPOIS
function _limparTracos(rowFull) {
  for (let c = 3; c < rowFull.length; c++) {                       // §F19b.03
    const v = (rowFull[c]||"").toString().trim();
    // §F19b.04 ── remove "-", "−", valor numérico 0 vindo de célula com traço
    if (v === "-" || v === "\u2212" || v === "\u2010" || v === "\u2013") {
      rowFull[c] = "";                                              // §F19b.05
    }
  }
  return rowFull;                                                   // §F19b.06
}

//================================================================
// §F20–§F29 ── 3_core.gs
//================================================================

// §F20 ── handleDividas
function handleDividas(e) {
  if (!e || !e.range) return;                                       // §F20.01

  const sheetEditada  = e.range.getSheet();                         // §F20.02
  const nomeAba       = sheetEditada.getName();                     // §F20.03
  if (!CONFIG.ABAS_PERMITIDAS.has(nomeAba)) return;                 // §F20.04

  const linhaEditada  = e.range.getRow();                           // §F20.05
  const colunaEditada = e.range.getColumn();                        // §F20.06
  const colunaFim     = e.range.getLastColumn();                    // §F20.07
  if (linhaEditada < CONFIG.LINHA_INICIAL) return;                  // §F20.08

  const ultimaColuna = sheetEditada.getLastColumn();                // §F20.09

  const editouColRelevante =                                        // §F20.10
    (colunaEditada <= CONFIG.COL.NOME && colunaFim >= CONFIG.COL.NOME) ||
    colunaEditada === CONFIG.COL.DIA_VENC  ||
    colunaEditada === 4                    ||
    colunaEditada === CONFIG.COL.VALOR     ||
    colunaEditada === CONFIG.COL.PARCELAS  ||
    colunaEditada >= CONFIG.COL.PRIMEIRA_PARC;

  if (!editouColRelevante) return;                                   // §F20.11

  const totalLinhas = sheetEditada.getMaxRows();                    // §F20.12
  const linhaInicio = Math.max(CONFIG.LINHA_DADOS, linhaEditada - 2); // §F20.13
  const linhaFim2   = Math.min(totalLinhas, linhaEditada + 2);      // §F20.14
  const quantidade  = linhaFim2 - linhaInicio + 1;                  // §F20.15

  const dados = sheetEditada                                        // §F20.16
    .getRange(linhaInicio, 1, quantidade, ultimaColuna)
    .getValues();

  dados.forEach(function(dl, idx) {                                 // §F20.17
    formatarLinha(sheetEditada, linhaInicio + idx, dl, ultimaColuna);
  });

  if (CONFIG.ABAS_AUXILIARES.has(nomeAba)) {                        // §F20.18
    const dadosEditados = sheetEditada                              // §F20.19
      .getRange(linhaEditada, 1, e.range.getNumRows(), ultimaColuna)
      .getValues();

    let precisaReconciliar = false;                                  // §F20.20
    let idParaReconciliar  = null;                                   // §F20.21

    dadosEditados.forEach(function(dl, idx) {                       // §F20.22
      const linhaAtual  = linhaEditada + idx;                       // §F20.23
      const nomeNaLinha = (dl[CONFIG.COL.NOME-1]||"").toString().trim(); // §F20.24
      if (!nomeNaLinha || sheetEditada.isRowHiddenByUser(linhaAtual)) { // §F20.25
        precisaReconciliar = true;
        if (!idParaReconciliar) {                                    // §F20.26
          idParaReconciliar = _extrairId(
            sheetEditada.getRange(linhaAtual, CONFIG.COL.DATA).getValue().toString()
          );
        }
      } else {
        copiarParaAbaPessoa(sheetEditada, linhaAtual, dl, ultimaColuna); // §F20.27
      }
    });

    // DEPOIS §F20.28-32
    if (precisaReconciliar) {                                       // §F20.28
      if (idParaReconciliar) {                                      // §F20.29
        // §F20.30 ── linha foi esvaziada mas ID ainda estava em col A
        reconciliarRemocoes(sheetEditada, linhaEditada, idParaReconciliar);
      } else {                                                      // §F20.31
        // §F20.32 ── linha apagada incluindo col A (ID sumiu) — varre tudo
        reconciliarRemocoes(sheetEditada, linhaEditada, "FORCAR_TUDO");
      }
    }
  }

  colorirLinha2(sheetEditada);                                      // §F20.31
}

// §F21 ── _ultimaLinhaComDados
function _ultimaLinhaComDados(sheet) {
  const ultima = sheet.getLastRow();                                // §F21.01
  if (ultima < CONFIG.LINHA_DADOS) return CONFIG.LINHA_DADOS - 1;  // §F21.02
  const vals = sheet                                                // §F21.03
    .getRange(CONFIG.LINHA_DADOS, CONFIG.COL.NOME, ultima-CONFIG.LINHA_DADOS+1, 1)
    .getValues();
  for (let i = vals.length-1; i >= 0; i--) {                       // §F21.04
    const v = (vals[i][0]||"").toString().trim().replace(/^⚠️\s*/,"").replace(/^\s+$/,"");
    if (v !== "") return CONFIG.LINHA_DADOS + i;                    // §F21.05
  }
  return CONFIG.LINHA_DADOS - 1;                                    // §F21.06
}

//================================================================
// §F22 ── copiarParaAbaPessoa
//================================================================
function copiarParaAbaPessoa(sheetOrigem, linhaNum, dadosLinha, ultimaCol) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();        // §F22.01

  const nomeRaw     = (dadosLinha[CONFIG.COL.NOME-1]||"").toString().trim(); // §F22.02
  const nomeNominal = nomeRaw.toUpperCase().replace(/^⚠️\s*/,"").trim(); // §F22.03
  if (!nomeNominal) return;                                         // §F22.04

  const textoParcelas = (dadosLinha[CONFIG.COL.PARCELAS-1]||"").toString().trim(); // §F22.05
  const matchQtdParc  = textoParcelas.match(/^(\d+)×/);            // §F22.06
  const qtdEsperada   = matchQtdParc ? parseInt(matchQtdParc[1]) : 0; // §F22.07
  if (qtdEsperada === 0) return;                                    // §F22.08

  const colParcIdx           = CONFIG.COL.PRIMEIRA_PARC - 1;       // §F22.09
  const todasColunasParcelas = dadosLinha.slice(colParcIdx);        // §F22.10

  const qtdPreenchidas = todasColunasParcelas.filter(function(v) { // §F22.11
    const t = (v||"").toString().trim(); return t !== "" && t !== "-";
  }).length;

  if (qtdPreenchidas < qtdEsperada) {                               // §F22.12
    Logger.log("L"+linhaNum+": "+qtdPreenchidas+"/"+qtdEsperada+" parcelas. Aguardando."); return;
  }
  if (qtdPreenchidas > qtdEsperada) {                               // §F22.13
    Logger.log("L"+linhaNum+": excesso "+qtdPreenchidas+"/"+qtdEsperada+". Verifique col F."); return;
  }

  const idLinha      = _garantirId(sheetOrigem, linhaNum);          // §F22.14
  const conteudoColA = sheetOrigem.getRange(linhaNum, CONFIG.COL.DATA).getValue().toString().trim(); // §F22.15

  const matchValorParc  = textoParcelas.match(/R\$([\d.,]+)/i);     // §F22.16
  const valorPorParcela = matchValorParc ? parseFloat(matchValorParc[1].replace(/,/g,"")) : 0; // §F22.17
  const IDX_COL_E       = 4;                                        // §F22.18
  const valorChave      = dadosLinha[CONFIG.COL.VALOR-1];           // §F22.19
  const valorAbs        = Math.abs(Number(valorChave));             // §F22.20

  // §F22.21 ── captura valor numérico de D antes de qualquer limpeza
  const vDoriginal = parseFloat(
  (dadosLinha[3] || "")
    .toString()
    .replace(/[-\u2212\u2010\u2013]/g, "") // 🔥 remove traços
    .replace(/[R$\s]/g,"")
    .replace(/,/g,".")
  );

  const donosBrutos = Object.keys(EMAILS_POR_ABA)                  // §F22.23
    .filter(function(n)  { return nomeNominal.includes(n); })
    .sort(function(a, b) { return nomeNominal.indexOf(a) - nomeNominal.indexOf(b); });

  if (donosBrutos.length === 0) return;                             // §F22.24

  const qtdPessoas = donosBrutos.length;                            // §F22.25
  const modo       = (nomeNominal.includes(" / ") || qtdEsperada < qtdPessoas) ? "dividido" : "alternado"; // §F22.26

  const posicoesReais = [];                                         // §F22.27
  for (let p = 0; p < todasColunasParcelas.length; p++) {          // §F22.28
    const t = (todasColunasParcelas[p]||"").toString().trim();
    if (t !== "" && t !== "-") posicoesReais.push(colParcIdx + p);
  }

  const qtdColsBase    = Math.max(ultimaCol, 1);                    // §F22.29
  const dadosLinhaBase = [];                                        // §F22.30
  for (let c = 0; c < qtdColsBase; c++) {                          // §F22.31
    dadosLinhaBase.push(dadosLinha[c] !== undefined ? dadosLinha[c] : "");
  }
  dadosLinhaBase[CONFIG.COL.DATA-1] = conteudoColA;                 // §F22.32

  const PREFIXOS_ABA = {                                            // §F22.33
    "SHOPEE":         "MSHO",
    "ISAEL SHOPEE":   "ISHO",
    "CARTÃO M-PAGO":  "CARM",
    "CRÉDITO M-PAGO": "CREM",
    "CARTÃO PAN":     "CARP"
  };
  const prefixo = PREFIXOS_ABA[sheetOrigem.getName()];              // §F22.34
  if (prefixo) {                                                    // §F22.35
    const nomeAtual = (dadosLinhaBase[CONFIG.COL.NOME-1]||"").toString().trim(); // §F22.36
    if (!nomeAtual.startsWith(prefixo+"] ")) {                      // §F22.37
      dadosLinhaBase[CONFIG.COL.NOME-1] = prefixo + "] " + nomeAtual; // §F22.38
    }
  }

  for (let indicePessoa = 0; indicePessoa < donosBrutos.length; indicePessoa++) { // §F22.39
    const nomeAbaDestino = donosBrutos[indicePessoa];               // §F22.40
    const sheetDestino   = spreadsheet.getSheetByName(nomeAbaDestino); // §F22.41
    if (!sheetDestino) continue;                                    // §F22.42

    const qtdColsLeitura   = Math.max(qtdColsBase, sheetDestino.getLastColumn()); // §F22.43
    const ultimaLinhaDados = _ultimaLinhaComDados(sheetDestino);    // §F22.44

    let linhaExistente = -1;                                        // §F22.45

    if (ultimaLinhaDados >= CONFIG.LINHA_DADOS) {                   // §F22.46
      const dadosDest = sheetDestino                                // §F22.47
        .getRange(CONFIG.LINHA_DADOS, 1, ultimaLinhaDados-CONFIG.LINHA_DADOS+1, qtdColsLeitura)
        .getValues();

      for (let i = 0; i < dadosDest.length; i++) {                 // §F22.48 — 1ª passagem: ID exato
        const idDest = _extrairId(dadosDest[i][CONFIG.COL.DATA-1]);
        if (idDest === idLinha) {
          linhaExistente = CONFIG.LINHA_DADOS + i;
          Logger.log("✅ Achou na linha "+linhaExistente);
          break;
        }
      }

      if (linhaExistente < 0) {                                     // §F22.49 — 2ª passagem: NOME+VALOR antigo
        for (let i = 0; i < dadosDest.length; i++) {
          const nomeDest  = (dadosDest[i][CONFIG.COL.NOME-1]||"").toString().trim().toUpperCase().replace(/^⚠️\s*/,"");
          const valorDest = Math.abs(Number(dadosDest[i][CONFIG.COL.VALOR-1]));
          const idDest    = _extrairId(dadosDest[i][CONFIG.COL.DATA-1]);
          if (!idDest && nomeDest === nomeNominal && Math.abs(valorDest - valorAbs) < 0.05) {
            linhaExistente = CONFIG.LINHA_DADOS + i;
            const celDestA  = sheetDestino.getRange(linhaExistente, CONFIG.COL.DATA);
            const conteudoA = (celDestA.getValue()||"").toString().trim();
            celDestA.setValue(conteudoA ? conteudoA+" | "+idLinha : conteudoColA);
            Logger.log("✅ Migrou ID para linha antiga "+linhaExistente);
            break;
          }
        }
      }

      if (linhaExistente < 0) {
        Logger.log("⚠️ ID ["+idLinha+"] não encontrado em \""+nomeAbaDestino+"\" — será inserido como novo"); // §F22.50
      }
    }

    // §F22.51 ── monta rowFull a partir de dadosLinhaBase
    const rowFull = dadosLinhaBase.concat(                          // §F22.52
      new Array(Math.max(0, qtdColsLeitura - dadosLinhaBase.length)).fill("")
    );

    if (IDX_COL_E < rowFull.length && typeof rowFull[IDX_COL_E] === "number") { // §F22.53
      rowFull[IDX_COL_E] = Math.abs(rowFull[IDX_COL_E]);
    }

    if (qtdPessoas > 1) {                                           // §F22.54
      if (modo === "dividido") {                                    // §F22.55
        const fator = 1 / qtdPessoas;                               // §F22.56
        posicoesReais.forEach(function(posReal) {                   // §F22.57
          const t  = (dadosLinha[posReal]||"").toString().trim();
          const mv = t.match(/^(✓?)R?\$?([\d.,]+)(Eup)?$/i);
          if (mv) {
            rowFull[posReal] = (mv[1]||"") + "R$" +
              (parseFloat(mv[2].replace(/,/g,".")) * fator).toFixed(2) + (mv[3]||"");
          }
        });
        if (valorPorParcela > 0) {                                  // §F22.58
          const vDiv = parseFloat((valorPorParcela * fator).toFixed(2));
          rowFull[CONFIG.COL.PARCELAS-1] = qtdEsperada + "×R$" + vDiv.toFixed(2);
          rowFull[IDX_COL_E] = parseFloat((vDiv * qtdEsperada).toFixed(2));
        }
        // §F22.59 ── divide col D usando vDoriginal capturado antes de qualquer limpeza
        if (!isNaN(vDoriginal) && vDoriginal !== 0) {               // §F22.60
          const novoD = parseFloat((vDoriginal * fator).toFixed(2));
          rowFull[3] = isNaN(novoD) ? "" : Math.abs(novoD); // §F22.61
        }
      } else {                                                      // §F22.62 — modo alternado
        posicoesReais.forEach(function(posReal, p) {                // §F22.63
          if (p % qtdPessoas !== indicePessoa) rowFull[posReal] = "";
        });
        const qtdP = _qtdParcelasDoPessoa(qtdEsperada, indicePessoa, qtdPessoas); // §F22.64
        if (qtdP > 0 && valorPorParcela > 0) {                      // §F22.65
          rowFull[CONFIG.COL.PARCELAS-1] = qtdP + "×R$" + valorPorParcela.toFixed(2);
          rowFull[IDX_COL_E] = parseFloat((valorPorParcela * qtdP).toFixed(2));
        }
      }
    }

    // §F22.66 ── limpa "-" de D em diante no rowFull final, antes de escrever
    _limparTracos(rowFull);                                         // §F22.67

     rowFull[3] = (() => {
      let v = String(rowFull[3] || "").replace(/[-\u2212\u2010\u2013]/g, "").trim();
      let n = parseFloat(v.replace(/[R$\s]/g,"").replace(/,/g,"."));
      return isNaN(n) ? "" : Math.abs(n);
    })();


    // DEPOIS
    const _escreverDestino = function(linhaAlvo) {                  // §F22.68
      // §F22.69 ── força "" em D e E no rowFull se fonte tinha "-" ou NaN
      [3, 4].forEach(function(idx) {
        let v = (rowFull[idx] || "").toString().trim();

        // remove QUALQUER traço
        v = v.replace(/[-\u2212\u2010\u2013]/g, "");

        const num = parseFloat(v.replace(/[R$\s]/g,"").replace(/,/g,"."));

        rowFull[idx] = isNaN(num) ? "" : Math.abs(num);
      });
      sheetDestino.getRange(linhaAlvo, 1, 1, rowFull.length).setValues([rowFull]); // §F22.74
      // §F22.75 ── lê de volta col D e E e força clearContent se ainda vier "-"
      const rangDE = sheetDestino.getRange(linhaAlvo, 4, 1, 2);     // §F22.76 — colunas D e E
      const valsDE = rangDE.getDisplayValues()[0];                   // §F22.77 ── getDisplayValues pega o que aparece
      if (valsDE[0].trim() === "-") sheetDestino.getRange(linhaAlvo, 4).clearContent(); // §F22.78 — col D
      if (valsDE[1].trim() === "-") sheetDestino.getRange(linhaAlvo, 5).clearContent(); // §F22.79 — col E
      const dadosRelidos = sheetDestino.getRange(linhaAlvo, 1, 1, rowFull.length).getValues()[0]; // §F22.80
      formatarLinha(sheetDestino, linhaAlvo, dadosRelidos, rowFull.length); // §F22.81
    };

    if (linhaExistente > 0) {                                       // §F22.77
      _escreverDestino(linhaExistente);                              // §F22.78
      _salvarMapeamento(nomeAbaDestino, linhaExistente, nomeNominal, valorChave, conteudoColA); // §F22.79
      Logger.log("Sincronizado L"+linhaNum+" → \""+nomeAbaDestino+"\" L"+linhaExistente+" ["+modo+" "+(indicePessoa+1)+"/"+qtdPessoas+"] ID:"+idLinha); // §F22.80
    } else {                                                        // §F22.81
      const novaLinha = _ultimaLinhaComDados(sheetDestino) + 1;     // §F22.82
      sheetDestino.getRange(novaLinha, 1, 1, rowFull.length).clearContent(); // §F22.83
      _escreverDestino(novaLinha);                                   // §F22.84
      _salvarMapeamento(nomeAbaDestino, novaLinha, nomeNominal, valorChave, conteudoColA); // §F22.85
      Logger.log("Copiada L"+linhaNum+" → \""+nomeAbaDestino+"\" L"+novaLinha+" ["+modo+" "+(indicePessoa+1)+"/"+qtdPessoas+"] ID:"+idLinha); // §F22.86
    }
  }
}

//================================================================
// §F23 ── onChangeSync
//================================================================
function onChangeSync(e) {
  Logger.log("onChangeSync chamado | changeType: " + (e ? e.changeType : "null")); // §F23.01
  if (!e || e.changeType !== "REMOVE_ROW") return;                  // §F23.02 ── só REMOVE_ROW, nunca EDIT
  const sheet = e.source.getActiveSheet();                          // §F23.03
  Logger.log("Aba: " + sheet.getName());                            // §F23.04
  if (!CONFIG.ABAS_AUXILIARES.has(sheet.getName())) return;         // §F23.05
  reconciliarRemocoes(sheet, null, "FORCAR_TUDO");                  // §F23.06
}

//================================================================
// §F24 ── reconciliarRemocoes
//================================================================
function reconciliarRemocoes(sheetOrigem, linhaEspecifica, idEspecifico) {
  if (!idEspecifico) return;                                        // §F24.01
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();        // §F24.02

  // §F24.03 ── coleta todos os IDs vivos em TODAS as abas auxiliares
  const idsVivos = new Set();                                       // §F24.04
  CONFIG.ABAS_AUXILIARES.forEach(function(nomeAba) {               // §F24.05
    const sheet = spreadsheet.getSheetByName(nomeAba);
    if (!sheet) return;
    const ultima = _ultimaLinhaComDados(sheet);
    if (ultima < CONFIG.LINHA_DADOS) return;
    sheet.getRange(CONFIG.LINHA_DADOS, CONFIG.COL.DATA, ultima - CONFIG.LINHA_DADOS + 1, 1)
      .getValues().forEach(function(dl) {
        const id = _extrairId(dl[0]);
        if (id) idsVivos.add(id);                                   // §F24.06
      });
  });

  // §F24.07 ── prefixos que identificam linhas vindas de abas auxiliares
  const PREFIXOS_RE = /^(MSHO|ISHO|CARM|CREM|CARP)\]/;            // §F24.08

  // §F24.09 ── percorre abas pessoa e apaga linhas cujo ID sumiu
  const jaProcessados = new Set();                                  // §F24.10
  Object.keys(EMAILS_POR_ABA).forEach(function(nomeAba) {          // §F24.11
    const sheet = spreadsheet.getSheetByName(nomeAba);
    if (!sheet) return;
    const ultima = _ultimaLinhaComDados(sheet);
    if (ultima < CONFIG.LINHA_DADOS) return;
    const dados = sheet
      .getRange(CONFIG.LINHA_DADOS, 1, ultima - CONFIG.LINHA_DADOS + 1, sheet.getLastColumn())
      .getValues();
    dados.forEach(function(dl, idx) {                               // §F24.12
      const nome = (dl[CONFIG.COL.NOME-1]||"").toString().trim();
      if (!nome || !PREFIXOS_RE.test(nome)) return;                 // §F24.13 — só linhas de aux
      const id = _extrairId(dl[CONFIG.COL.DATA-1]);
      if (!id) return;                                              // §F24.14 — sem ID não processa
      if (idEspecifico !== "FORCAR_TUDO" && id !== idEspecifico) return; // §F24.15
      if (idsVivos.has(id)) return;                                 // §F24.16 — ainda existe, mantém
      const linhaNum  = CONFIG.LINHA_DADOS + idx;                   // §F24.17
      const chaveProc = nomeAba + ":" + linhaNum;                   // §F24.18
      if (jaProcessados.has(chaveProc)) return;                     // §F24.19
      jaProcessados.add(chaveProc);                                 // §F24.20
      Logger.log("🚨 Apagando \""+nomeAba+"\" L"+linhaNum+" ID:"+id+" motivo: ID não existe mais nas abas auxiliares"); // §F24.21
      _limparLinhaDestinoCompleta(sheet, linhaNum, "ID " + id + " não existe mais nas abas auxiliares"); // §F24.22
    });
  });
}

//================================================================
// §F25 ── colorirLinha2
//================================================================
function colorirLinha2(sheet) {
  const ultima = sheet.getLastColumn();                             // §F25.01
  if (ultima < 1) return;                                          // §F25.02
  const faixa  = sheet.getRange(2, 1, 1, ultima);                  // §F25.03
  const vals   = faixa.getValues()[0];                              // §F25.04
  const fundos = [], textos = [];                                   // §F25.05
  vals.forEach(function(v) {                                        // §F25.06
    if (typeof v !== "number") { fundos.push(null); textos.push(null); }
    else if (v < 0)            { fundos.push("#ff0000"); textos.push("#000000"); }
    else                       { fundos.push("#00b050"); textos.push("#000000"); }
  });
  faixa.setBackgrounds([fundos]); faixa.setFontColors([textos]);   // §F25.07
}

//================================================================
// §F26 ── _limparLinhaDestinoCompleta
//================================================================
function _limparLinhaDestinoCompleta(sheetDestino, linhaDestino, motivo) {
  if (!sheetDestino || linhaDestino < 1) return;                    // §F26.01

  motivo = motivo || "não informado";                               // §F26.02

  const ultimaCol  = Math.max(sheetDestino.getLastColumn(), CONFIG.COL.PRIMEIRA_PARC+10); // §F26.03
  const dadosLinha = sheetDestino.getRange(linhaDestino, 1, 1, ultimaCol).getValues()[0]; // §F26.04
  const nomeAtual  = (dadosLinha[CONFIG.COL.NOME-1]||"").toString().trim(); // §F26.05
  const dataRaw    = (dadosLinha[CONFIG.COL.DATA-1]||"").toString().trim(); // §F26.06
  const partesDt   = dataRaw.split("|");                            // §F26.07
  const dataAtual  = partesDt[0].trim();                            // §F26.08
  const idAtual    = partesDt.length > 1 ? partesDt[1].trim() : "(sem ID)"; // §F26.09
  const valorAtual = (dadosLinha[CONFIG.COL.VALOR-1]||"").toString().trim(); // §F26.10
  const diaVnc     = (dadosLinha[CONFIG.COL.DIA_VENC-1]||"").toString().trim(); // §F26.11
  const parcelas   = (dadosLinha[CONFIG.COL.PARCELAS-1]||"").toString().trim(); // §F26.12

  const agora   = new Date();                                       // §F26.13
  const tz      = Session.getScriptTimeZone();                      // §F26.14
  const dataStr = Utilities.formatDate(agora, tz, "dd/MM/yyyy");   // §F26.15
  const horaStr = Utilities.formatDate(agora, tz, "HH:mm:ss");     // §F26.16

  Logger.log("🗑️ _limparLinhaDestinoCompleta: \""+sheetDestino.getName()+"\" L"+linhaDestino+" | nome: "+nomeAtual+" | motivo: "+motivo); // §F26.17

  if (nomeAtual) {                                                  // §F26.18
    const props    = PropertiesService.getScriptProperties();       // §F26.19
    const listaRaw = props.getProperty("emailsPendentes") || "[]";  // §F26.20
    const lista    = JSON.parse(listaRaw);                           // §F26.21
    lista.push({                                                    // §F26.22
      aba:     sheetDestino.getName(),
      assunto: "🗑️ Linha apagada: " + nomeAtual,
      corpo:
        "Linha: "             + linhaDestino           + " da aba: " + sheetDestino.getName() + " apagada automaticamente\n\n" +
        "Motivo: "            + motivo                 + "\n\n" +
        "Data da exclusão: "  + dataStr                + "\n" +
        "Hora da exclusão: "  + horaStr                + "\n" +
        "Aba: "               + sheetDestino.getName() + "\n" +
        "Linha: "             + linhaDestino           + "\n" +
        "Data do gasto: "     + dataAtual              + "\n" +
        "ID da linha: "       + idAtual                + "\n" +
        "Gasto: "             + nomeAtual              + "\n" +
        "Dia de vencimento: " + diaVnc                 + "\n" +
        "Valor: R$"           + valorAtual             + "\n" +
        "Parcelas: "          + parcelas               + "\n\n" +
        "Se foi você, desconsidere."
    });
    props.setProperty("emailsPendentes", JSON.stringify(lista));    // §F26.23

    // §F26.24 ── remove chaves lnk| que apontavam para esta linha
    const allProps = props.getProperties();                         // §F26.25
    for (const k of Object.keys(allProps)) {                        // §F26.26
      if (k.startsWith("lnk|" + sheetDestino.getName() + "|") &&
          parseInt(allProps[k]) === linhaDestino) {
        props.deleteProperty(k);                                    // §F26.27
        Logger.log("🔑 Chave removida: " + k);                      // §F26.28
      }
    }
  }

  const iv = sheetDestino.getRange(linhaDestino, 1, 1, ultimaCol); // §F26.29
  iv.clearContent();                                                // §F26.30
  iv.setBackgrounds([new Array(ultimaCol).fill("#000000")]);        // §F26.31
  iv.setFontColors([new Array(ultimaCol).fill("#ffffff")]);         // §F26.32
  iv.setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.DOUBLE); // §F26.33
  sheetDestino.getRange(linhaDestino, CONFIG.COL.NOME).setValue(""); // §F26.34
}

//================================================================
// §F27 ── formatarLinha
//================================================================
function formatarLinha(sheet, linhaNum, dadosLinha, ultimaCol) {
  const range = sheet.getRange(linhaNum, 1, 1, ultimaCol);          // §F27.01

  const nomeNaLinha  = (dadosLinha[CONFIG.COL.NOME - 1] || "").toString().trim(); // §F27.02
  const valorNaLinha = Number(dadosLinha[CONFIG.COL.VALOR - 1]);    // §F27.03

  range.setFontFamily("Arial");                                     // §F27.04
  range.setFontSize(12);                                            // §F27.05
  range.setFontWeight("bold");                                      // §F27.06
  range.setFontStyle("normal");                                     // §F27.07
  range.setHorizontalAlignment("left");                             // §F27.08
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);          // §F27.09
  range.setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.DOUBLE); // §F27.10

  if (!nomeNaLinha) {                                               // §F27.11
    range.setBackgrounds([new Array(ultimaCol).fill("#000000")]);
    range.setFontColors([new Array(ultimaCol).fill("#ffffff")]);
    return;
  }

  preencherData(sheet, linhaNum, valorNaLinha);                     // §F27.12
  formatarParcelas(sheet, linhaNum, valorNaLinha, dadosLinha[CONFIG.COL.PARCELAS - 1]); // §F27.13
  normalizarNome(sheet, linhaNum, nomeNaLinha);                     // §F27.14

  const arrayFundos = new Array(ultimaCol).fill("#000000");         // §F27.15
  const arrayTextos = new Array(ultimaCol).fill("#ffffff");         // §F27.16

  const corPessoa = getCor(nomeNaLinha);                            // §F27.17
  if (corPessoa) {                                                  // §F27.18
    for (let c = 0; c < 3; c++) {                                   // §F27.19
      arrayFundos[c] = corPessoa.bg;
      arrayTextos[c] = corPessoa.fg;
    }
  }

  [3, 4].forEach(function(idx) {                                    // §F27.20 — colunas D e E
    const raw = dadosLinha[idx];
    const v   = typeof raw === "number" ? raw
              : parseFloat((raw||"").toString().replace(/[R$\s]/g,"").replace(/,/g,".")) || null;
    if (v === null || isNaN(v)) {                                   // §F27.21
      arrayFundos[idx] = "#000000";
      arrayTextos[idx] = "#ffffff";
      return;
    }
    arrayFundos[idx] = v < 0 ? "#ff0000" : "#00b050";              // §F27.22
    arrayTextos[idx] = "#000000";
  });

  const colF = dadosLinha[CONFIG.COL.PARCELAS - 1];                 // §F27.23
  if (colF !== null && colF !== undefined && colF.toString().trim() !== "") { // §F27.24
    arrayFundos[CONFIG.COL.PARCELAS - 1] = "#46bdc6";
    arrayTextos[CONFIG.COL.PARCELAS - 1] = "#000000";
  }

  const { coresFundoParcelas, coresTextoParcelas, todasPagas, temParcela, houvePagamento } = // §F27.25
    processarPagamentos(sheet, linhaNum, corPessoa, ultimaCol, dadosLinha);

  const colIni = CONFIG.COL.PRIMEIRA_PARC;                          // §F27.26
  coresFundoParcelas.forEach(function(bg, i) { if (bg !== null) arrayFundos[colIni - 1 + i] = bg; }); // §F27.27
  coresTextoParcelas.forEach(function(fg, i) { if (fg !== null) arrayTextos[colIni - 1 + i] = fg; }); // §F27.28

  if (todasPagas && temParcela) {                                   // §F27.29
    range.setBackgrounds([new Array(ultimaCol).fill("#00ff00")]);
    range.setFontColors([new Array(ultimaCol).fill("#000000")]);
    if (houvePagamento) limparAlerta(sheet, linhaNum);              // §F27.30
    return;
  }

  range.setBackgrounds([arrayFundos]);                              // §F27.31
  range.setFontColors([arrayTextos]);                               // §F27.32
  if (houvePagamento) limparAlerta(sheet, linhaNum);                // §F27.33
}

//================================================================
// §F28 ── processarPagamentos
//================================================================
function processarPagamentos(sheet, linha, corPessoa, ultimaCol, dadosLinha) {
  const colIni  = CONFIG.COL.PRIMEIRA_PARC;                         // §F28.01
  const qtdCols = ultimaCol - colIni + 1;                           // §F28.02
  if (qtdCols <= 0) {                                               // §F28.03
    return {coresFundoParcelas:[],coresTextoParcelas:[],todasPagas:false,temParcela:false,houvePagamento:false};
  }

  const SUFIXOS = {                                                 // §F28.04
    "eup": { cor: "#0000ff", nome: "Eup" },
    "isp": { cor: "#ff0000", nome: "Isp" },
    "map": { cor: "#ffff00", nome: "Map" },
    "dup": { cor: "#800080", nome: "Dup" },
    "rep": { cor: "#00ffff", nome: "Rep" },
    "rap": { cor: "#ff9900", nome: "Rap" }
  };

  const vals     = dadosLinha.slice(colIni-1);                      // §F28.05
  const fundos   = [], textos = [];                                 // §F28.06
  let todasPagas = true, temParcela = false, houvePagamento = false; // §F28.07

  vals.forEach(function(valorCelula, indice) {                      // §F28.08
    const t = (valorCelula||"").toString().trim();
    if (!t) { fundos.push(null); textos.push(null); return; }       // §F28.09
    temParcela = true;


    if (/^\d+([.,]\d+)?\s*[pP]$/.test(t)) {                        // §F28.10
      const v = parseFloat(t.replace(/[pP]/,"").replace(",","."));
      sheet.getRange(linha, colIni+indice).setValue("✓R$"+v.toFixed(2));
      fundos.push("#00ff00"); textos.push("#000000"); houvePagamento = true;

    } else if (/eup|isp|map|dup|rep|rap/i.test(t)) {               // §F28.11
      const match = t.match(/^(✓?)R?\$?([\d.,]+)\s*(eup|isp|map|dup|rep|rap)$/i);
      if (match) {                                                  // §F28.12
        const sufixoKey = match[3].toLowerCase();
        const cfg       = SUFIXOS[sufixoKey];
        const fmt       = (match[1]||"") + "R$" + parseFloat(match[2].replace(/,/g,".")).toFixed(2) + cfg.nome;
        if (t !== fmt) sheet.getRange(linha, colIni+indice).setValue(fmt);
        todasPagas = false;
        fundos.push(cfg.cor);
        textos.push(_corTextoContraste(cfg.cor));
      } else {                                                      // §F28.13
        todasPagas = false;
        fundos.push(corPessoa ? corPessoa.bg : null);
        textos.push(corPessoa ? corPessoa.fg : null);
      }

    } else if (t.startsWith("✓")) {                                 // §F28.14
      fundos.push("#00ff00"); textos.push("#000000"); houvePagamento = true;

    } else {                                                        // §F28.15
      todasPagas = false;
      fundos.push(corPessoa ? corPessoa.bg : null);
      textos.push(corPessoa ? corPessoa.fg : null);
    }
  });

  return {coresFundoParcelas:fundos,coresTextoParcelas:textos,todasPagas,temParcela,houvePagamento}; // §F28.16
}

// §F29 ── processarLinha
function processarLinha(sheet, linhaNum, dadosLinha, ultimaCol) {
  formatarLinha(sheet, linhaNum, dadosLinha, ultimaCol);             // §F29.01
}

//================================================================
// §F30–§F39 ── 4_email_e_lote.gs
//================================================================

// §F30 ── preencherCabecalhoMeses
function preencherCabecalhoMeses() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();        // §F30.01
  spreadsheet.getSheets().forEach(function(sheet) {                 // §F30.02
    const nomeAba = sheet.getName();                                // §F30.03
    if (!CONFIG.ABAS_PERMITIDAS.has(nomeAba)) return;               // §F30.04
    const colInicio = CONFIG.COL.PRIMEIRA_PARC;                     // §F30.05
    if (sheet.getLastColumn() < colInicio || sheet.getLastRow() < CONFIG.LINHA_DADOS) return; // §F30.06
    let corAba = CONFIG.CORES_EMAIL[nomeAba];                       // §F30.07
    if (!corAba) {                                                  // §F30.08
      for (const chave in CONFIG.CORES_EMAIL) {
        if (nomeAba.toUpperCase().includes(chave)) { corAba = CONFIG.CORES_EMAIL[chave]; break; }
      }
    }
    corAba = corAba || "#800080";                                   // §F30.09
    const agora = new Date(new Date().getFullYear(), 0, 1);         // §F30.10
    const qtdMeses = 24;                                            // §F30.11
    const valores = [[]], fundos = [[]], cores = [[]];              // §F30.12
    for (let m = 0; m < qtdMeses; m++) {                            // §F30.13
      const d = new Date(agora.getFullYear(), agora.getMonth() + m, 1);
      valores[0].push(d); fundos[0].push(corAba); cores[0].push("#ffffff");
    }
    const iv = sheet.getRange(CONFIG.LINHA_CABECALHO, colInicio, 1, qtdMeses); // §F30.14
    iv.setNumberFormat("MMM"); iv.setValues(valores);               // §F30.15
    iv.setBackgrounds(fundos); iv.setFontColors(cores);             // §F30.16
    iv.setFontWeight("bold"); iv.setHorizontalAlignment("center");  // §F30.17
  });
  SpreadsheetApp.getUi().alert("✅ Cabeçalhos de mês atualizados em todas as abas!"); // §F30.18
}

//================================================================
// §F31 ── verificarVencimentosAutomatico
//================================================================
function verificarVencimentosAutomatico() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();        // §F31.01
  const hoje        = new Date(); hoje.setHours(0,0,0,0);           // §F31.02
  const amanha      = new Date(hoje); amanha.setDate(hoje.getDate()+1); amanha.setHours(23,59,59,999); // §F31.03
  const vencPorAba  = {};                                           // §F31.04

  spreadsheet.getSheets().forEach(function(sheet) {                 // §F31.05
    const nomeAba = sheet.getName();                                // §F31.06
    if (!ABAS_NOTIFICACAO.has(nomeAba)) return;                     // §F31.07
    const ultimaLinha = sheet.getLastRow(), ultimaColuna = sheet.getLastColumn(); // §F31.08
    if (ultimaLinha < CONFIG.LINHA_DADOS) return;                   // §F31.09

    const colIni = CONFIG.COL.PRIMEIRA_PARC;                        // §F31.10
    const cabec  = sheet.getRange(CONFIG.LINHA_CABECALHO, colIni, 1, ultimaColuna-colIni+1).getValues()[0]; // §F31.11
    const meses  = cabec.map(function(c) {                          // §F31.12
      if (!c) return null; const d=new Date(c); if (isNaN(d)) return null;
      return {mes:d.getMonth(),ano:d.getFullYear()};
    });

    const linhas = sheet.getRange(CONFIG.LINHA_DADOS, 1, ultimaLinha-CONFIG.LINHA_DADOS+1, ultimaColuna).getValues(); // §F31.13
    const itens = []; let soma = 0;                                 // §F31.14

    linhas.forEach(function(dl, idx) {                              // §F31.15
      const numL=CONFIG.LINHA_DADOS+idx, diaVnc=dl[CONFIG.COL.DIA_VENC-1]; // §F31.16
      if (typeof diaVnc!=="number"||diaVnc<=0) return;              // §F31.17
      const parcsL=dl.slice(colIni-1); let total=0, pagas=0;        // §F31.18
      parcsL.forEach(function(v) {                                  // §F31.19
        const t=(v||"").toString().trim().toUpperCase();
        if (t===""||t==="-") return; total++;
        if (t.startsWith("✓")||t.startsWith("√")||t.includes("PAGO")||t.startsWith("OK")) pagas++;
        if (/EUP$/i.test(t)) pagas++;
      });

      for (let m=0; m<meses.length; m++) {                          // §F31.20
        const info=meses[m]; if (!info) continue;
        const dvenc=new Date(info.ano,info.mes,diaVnc); dvenc.setHours(0,0,0,0);
        if (dvenc>amanha) continue;
        const celParc=(dl[colIni-1+m]||"").toString().trim();
        if (!celParc||celParc==="-"||celParc.startsWith("✓")||celParc.startsWith("√")) continue;

        let vParcEup=0;
        const isEup=/EUP$/i.test(celParc);
        if (isEup) {
          const mvEup=celParc.match(/R?\$?([\d.,]+)\s*Eup$/i);
          if (mvEup) vParcEup=parseFloat(mvEup[1].replace(/,/g,"."));
        }

        const status    = dvenc.getTime()===hoje.getTime()?"hoje":(dvenc>hoje?"amanha":"vencida"); // §F31.21
        const dtVencFmt = Utilities.formatDate(dvenc,Session.getScriptTimeZone(),"dd/MM/yyyy");
        const dtRaw     = dl[CONFIG.COL.DATA-1];

        let dtFmt="sem data";
        if (dtRaw instanceof Date && !isNaN(dtRaw.getTime())) {
          dtFmt=Utilities.formatDate(dtRaw,Session.getScriptTimeZone(),"dd/MM/yyyy");
        } else {
          const txt=dtRaw.toString().split("|")[0].trim();
          const mxNum=txt.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
          if (mxNum) {
            const ano=mxNum[3]?(mxNum[3].length===2?"20"+mxNum[3]:mxNum[3]):new Date().getFullYear();
            dtFmt=mxNum[1].padStart(2,"0")+"/"+mxNum[2].padStart(2,"0")+"/"+ano;
          } else {
            const mxMes=txt.match(/^(\d{1,2})\/(\w{3})(?:\/(\d{4}))?$/);
            if (mxMes) {
              const mesIdx=CONFIG.MESES_PT.indexOf(mxMes[2]);
              if (mesIdx>=0) {
                const ano=mxMes[3]||new Date().getFullYear();
                dtFmt=mxMes[1].padStart(2,"0")+"/"+String(mesIdx+1).padStart(2,"0")+"/"+ano;
              }
            }
          }
        }

        let vParc=0;
        const celParcValor=(dl[colIni-1+m]||"").toString().trim();
        const mvParc=celParcValor.replace(/✓/g,"").match(/R?\$?([\d.,]+)/i);
        if (mvParc) vParc=parseFloat(mvParc[1].replace(/,/g,"."));
        if (vParc===0) {
          const tParcs=(dl[CONFIG.COL.PARCELAS-1]||"").toString();
          const mVP=tParcs.match(/R\$([\d.,]+)/i);
          if (mVP) vParc=parseFloat(mVP[1].replace(/,/g,""));
        }
        const vTotal=Math.abs(parseFloat((dl[CONFIG.COL.VALOR-1]||0).toString().replace(/,/g,""))||0);
        if (isEup&&vParcEup>0) vParc=vParcEup;

        soma+=vParc;
        itens.push({linha:numL,nome:(dl[CONFIG.COL.NOME-1]||"(sem nome)").toString().trim(),
          data:dtFmt,dataVenc:dtVencFmt,valorParcela:vParc,totalParcelas:total,pagas:pagas,
          pendentes:total-pagas,valorTotal:vTotal,status:status});
      }
    });
    if (itens.length>0) vencPorAba[nomeAba]={itens:itens,somaTotal:soma}; // §F31.22
  });

  Object.keys(vencPorAba).forEach(function(nomeAba) {               // §F31.23
    const {itens,somaTotal}=vencPorAba[nomeAba];
    const emailDest=EMAILS_POR_ABA[nomeAba]||CONFIG.EMAIL_ALERTA;
    const msgCfg=getMensagem(nomeAba),unica=itens.length===1,corGasto=getCorEmail(nomeAba);

    const qtdV=itens.filter(function(v){return v.status==="vencida";}).length;
    const qtdH=itens.filter(function(v){return v.status==="hoje";}).length;
    const qtdA=itens.filter(function(v){return v.status==="amanha";}).length;
    const partes=[];
    if(qtdV>0)partes.push(qtdV===1?"está vencida":"estão vencidas");
    if(qtdH>0)partes.push(qtdH===1?"está vencendo hoje":"estão vencendo hoje");
    if(qtdA>0)partes.push(qtdA===1?"vence amanhã":"vencem amanhã");
    const txVence=partes.length===1?partes[0]:partes.length===2?partes[0]+" e "+partes[1]:partes[0]+", "+partes[1]+" e "+partes[2];

    const diaRef=(itens[0].dataVenc||"").split("/")[0]||"";
    const vars={TextoDivida:unica?"Dívida":"Dívidas",textoDivida:unica?"dívida":"dívidas",
      vence:txVence,conta:unica?"conta":"contas",Sua:unica?"Sua":"Suas",dia:diaRef,
      nome:itens[0].nome||nomeAba,
      total:unica?itens[0].valorParcela.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2}):somaTotal.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2})};

    let assunto=substituir(msgCfg.assunto,vars);
    if(itens.length>1)assunto+=" ("+itens.length+" "+vars.conta+")";

    function fmtV(v){return v.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});}
    function lblStatus(s,dv){
      if(s==="vencida")return{txt:"VENCIDA em "+dv,cor:"#cc0000"};
      if(s==="hoje")return{txt:"Vence HOJE ("+dv+")",cor:"#e68a00"};
      return{txt:"Vence amanhã ("+dv+")",cor:"#007700"};
    }
    function corVP(s){return s==="vencida"?"#cc0000":s==="hoje"?"#e68a00":"#007700";}
    function montarItem(item,pfx){
      const sp=fmtV(item.valorParcela),st=fmtV(item.valorTotal);
      const lp=item.pagas===1?"Parcela paga":"Parcelas pagas",lpd=item.pendentes===1?"Pendente":"Pendentes";
      const jp=fmtV(item.pagas*item.valorParcela),pend=fmtV(item.pendentes*item.valorParcela);
      const info=lblStatus(item.status,item.dataVenc),cVP=corVP(item.status),nome=item.nome.replace(/^⚠️\s*/,"");
      return{
        txt:pfx+"Gasto: "+nome+"\n"+pfx+"Data da dívida: "+item.data+"\n"+pfx+"Valor da parcela: R$"+sp+"\n"+
            pfx+lp+": "+item.pagas+" de "+item.totalParcelas+" ("+item.pendentes+" "+lpd+")\n"+
            pfx+"Valor já pago: R$"+jp+" (R$"+pend+" Pendente)\n"+
            pfx+"Total da dívida: R$"+st+" ("+item.totalParcelas+"x R$"+sp+")\n"+
            pfx+"Data de vencimento: "+info.txt+"\n\n",
        html:"<b>Gasto:</b> <span style='color:"+corGasto+";font-weight:bold'>"+nome+"</span><br>"+
             "<b>Data da dívida:</b> "+item.data+"<br>"+
             "<b>Valor da parcela:</b> <span style='color:"+cVP+";font-weight:bold'>R$"+sp+"</span><br>"+
             "<b>"+lp+":</b> "+item.pagas+" de "+item.totalParcelas+" <b>("+item.pendentes+" "+lpd+")</b><br>"+
             "<b>Valor já pago:</b> <span style='color:#007700;font-weight:bold'>R$"+jp+"</span> <b style='color:#007700'>(R$"+pend+" Pendente)</b><br>"+
             "<b>Total da dívida:</b> <span style='color:#007700;font-weight:bold'>R$"+st+"</span> <b>("+item.totalParcelas+"x R$"+sp+")</b><br>"+
             "<b>Data de vencimento:</b> <span style='color:"+info.cor+";font-weight:bold'>"+info.txt+"</span><br><br>"
      };
    }

    let corpoTxt=substituir(msgCfg.saudacao,vars)+" "+substituir(msgCfg.introducao,vars)+"\n";
    let itensTxt="",itensHtml="";
    if(unica){const r=montarItem(itens[0],"");itensTxt+=r.txt;itensHtml+=r.html;}
    else{
      itensTxt+="Encontrei "+itens.length+" "+vars.textoDivida+":\n\n";
      itensHtml+="<b>Encontrei "+itens.length+" "+vars.textoDivida+":</b><br><br>";
      itens.forEach(function(item){const r=montarItem(item,"   ");itensTxt+=r.txt;itensHtml+=r.html;});
    }
    corpoTxt+=itensTxt+msgCfg.fechamento+"\n\nAtenciosamente, Seu sistema de alertas";

    const vPix=unica?itens[0].valorParcela:parseFloat(itens.reduce(function(a,v){return a+v.valorParcela;},0).toFixed(2));
    const pPix=gerarPixPayload(PIX_CHAVE,PIX_NOME,vPix);
    const sPix=vPix.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
    const urlPix=WEB_APP_URL+"?pix="+encodeURIComponent(pPix)+"&valor="+encodeURIComponent(sPix);
    const botao="<div style='text-align:center;margin:20px 0'><a href='"+urlPix+"' style='display:inline-block;background:#32BCAD;color:#fff;padding:12px 28px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:15px'>Pagar R$"+sPix+" via Pix</a></div>";
    const htmlBody="<div style='font-family:Arial,sans-serif;font-size:14px;color:#333;max-width:600px'><b>"+substituir(msgCfg.saudacao,vars)+" "+substituir(msgCfg.introducao,vars)+"</b><br>"+botao+itensHtml+"<b>"+msgCfg.fechamento+"</b><br><br><b>Atenciosamente, Seu sistema de alertas</b></div>";

    GmailApp.sendEmail(emailDest,assunto,corpoTxt,{htmlBody:htmlBody});
    if(emailDest!==CONFIG.EMAIL_ALERTA)GmailApp.sendEmail(CONFIG.EMAIL_ALERTA,assunto,corpoTxt,{htmlBody:htmlBody});
    const sheetMarcar=spreadsheet.getSheetByName(nomeAba);
    itens.forEach(function(item){marcarAlerta(sheetMarcar,item.linha);});
  });
}

//================================================================
// §F32 ── aplicarTodasAsAlteracoes
//================================================================
function aplicarTodasAsAlteracoes() {
  const props=PropertiesService.getScriptProperties(); props.deleteAllProperties(); // §F32.01
  const ss=SpreadsheetApp.getActiveSpreadsheet(), fila=[];
  ss.getSheets().forEach(function(sheet){
    const n=sheet.getName(); if(!CONFIG.ABAS_PERMITIDAS.has(n))return;
    const u=sheet.getLastRow(); if(u<CONFIG.LINHA_INICIAL)return;
    fila.push({aba:n,ultimaLinha:u});
  });
  props.setProperty("fila",JSON.stringify(fila));
  props.setProperty("filaIndex","0");
  props.setProperty("linhaAtual",String(CONFIG.LINHA_INICIAL));
  continuarAplicacao();
}

//================================================================
// §F33 ── continuarAplicacao
//////////////////////////////////////////////////////////////////
function continuarAplicacao() {
  const LOTE=50, MAX=5*60*1000, t0=Date.now();
  const props=PropertiesService.getScriptProperties();
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const fila=JSON.parse(props.getProperty("fila")||"[]");
  let idxAba=parseInt(props.getProperty("filaIndex")||"0");
  let linhaAtual=parseInt(props.getProperty("linhaAtual")||String(CONFIG.LINHA_INICIAL));

  if(idxAba>=fila.length){SpreadsheetApp.getUi().alert("✅ Todas as alterações foram aplicadas!");props.deleteAllProperties();return;}

  while(idxAba<fila.length){
    if(Date.now()-t0>MAX)break;
    const{aba,ultimaLinha}=fila[idxAba];
    const sheetAtual=ss.getSheetByName(aba),ultimaCol=sheetAtual.getLastColumn();
    if(linhaAtual===CONFIG.LINHA_INICIAL)colorirLinha2(sheetAtual);
    while(linhaAtual<=ultimaLinha){
      if(Date.now()-t0>MAX){
        props.setProperty("filaIndex",String(idxAba));props.setProperty("linhaAtual",String(linhaAtual));
        SpreadsheetApp.getUi().alert("⏳ Pausa em: "+aba+", linha "+linhaAtual+".\n\nExecute \"continuarAplicacao\" para continuar.");
        return;
      }
      const linhaFim=Math.min(linhaAtual+LOTE-1,ultimaLinha);
      const dados=sheetAtual.getRange(linhaAtual,1,linhaFim-linhaAtual+1,ultimaCol).getValues();
      dados.forEach(function(dl,idx){formatarLinha(sheetAtual,linhaAtual+idx,dl,ultimaCol);});
      linhaAtual=linhaFim+1;
    }
    idxAba++;linhaAtual=CONFIG.LINHA_INICIAL;
  }
  props.deleteAllProperties();
  SpreadsheetApp.getUi().alert("✅ Todas as alterações foram aplicadas com sucesso!");
}

// §F34 ── formatarTodasAsLinhas
function formatarTodasAsLinhas() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet){
    if(!CONFIG.ABAS_PERMITIDAS.has(sheet.getName()))return;
    _aplicarFormatacaoGlobal(sheet);
  });
  SpreadsheetApp.getUi().alert("✅ Formatação aplicada em todas as linhas de todas as abas!");
}

// §F35 ── testarProcessarLinha
function testarProcessarLinha() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet) {
    const nomeAba=sheet.getName();
    if (!CONFIG.ABAS_PERMITIDAS.has(nomeAba)) return;
    const ultimaCol=sheet.getLastColumn(), totalLinhas=sheet.getMaxRows();
    if (totalLinhas<CONFIG.LINHA_DADOS||ultimaCol<1) return;
    const dados=sheet.getRange(CONFIG.LINHA_DADOS,1,totalLinhas-CONFIG.LINHA_DADOS+1,ultimaCol).getValues();
    dados.forEach(function(dl,idx){formatarLinha(sheet,CONFIG.LINHA_DADOS+idx,dl,ultimaCol);});
    Logger.log("✅ "+nomeAba+" — "+dados.length+" linhas formatadas.");
  });
  SpreadsheetApp.getUi().alert("✅ Formatação completa aplicada em todas as abas!");
}

//================================================================
// §F36 ── enviarEmailsPendentes
//================================================================
function enviarEmailsPendentes() {
  const props    = PropertiesService.getScriptProperties();         // §F36.01
  const listaRaw = props.getProperty("emailsPendentes") || "[]";   // §F36.02
  const lista    = JSON.parse(listaRaw);                             // §F36.03
  if (lista.length === 0) return;                                   // §F36.04

  const porAba = {};                                                // §F36.05
  lista.forEach(function(item) {                                    // §F36.06
    if (!porAba[item.aba]) porAba[item.aba] = [];
    porAba[item.aba].push(item);
  });

  Object.keys(porAba).forEach(function(aba) {                       // §F36.07
    const itens   = porAba[aba];
    const assunto = itens.length === 1
      ? itens[0].assunto
      : "🗑️ " + itens.length + " linhas apagadas na aba \"" + aba + "\"";
    const corpo = itens.map(function(i){return i.corpo;}).join("\n"+"─".repeat(40)+"\n");
    GmailApp.sendEmail(CONFIG.EMAIL_ALERTA, assunto, corpo);
    Logger.log("✅ Email pendente enviado ("+itens.length+" item(s)): "+assunto);
  });

  props.deleteProperty("emailsPendentes");                          // §F36.08
  props.deleteProperty("emailPendente");                            // §F36.09
}

//================================================================
// §F37 ── limparTodasAsChaves
//================================================================
function limparTodasAsChaves() {
  const props=PropertiesService.getScriptProperties().getProperties();
  let count=0;
  for (const chave of Object.keys(props)) {
    if (chave.startsWith("lnk|")) {
      PropertiesService.getScriptProperties().deleteProperty(chave);
      count++;
    }
  }
  SpreadsheetApp.getUi().alert("✅ "+count+" chaves removidas!");
}

//================================================================
// §F38 ── apagarLinhaSelecionada
//================================================================
function apagarLinhaSelecionada() {
  const sheet=SpreadsheetApp.getActiveSheet();
  const linha=sheet.getActiveRange().getRow();
  if (linha<CONFIG.LINHA_DADOS) {
    SpreadsheetApp.getUi().alert("❌ Não é possível apagar esta linha."); return;
  }
  const id  =_extrairId(sheet.getRange(linha,CONFIG.COL.DATA).getValue().toString());
  const nome=sheet.getRange(linha,CONFIG.COL.NOME).getValue().toString().trim();
  const ui  =SpreadsheetApp.getUi();
  const resp=ui.alert("Apagar linha?","Deseja apagar: "+nome+"?",ui.ButtonSet.YES_NO);
  if (resp!==ui.Button.YES) return;
  if (id) reconciliarRemocoes(sheet, linha, id);
  sheet.getRange(linha,1,1,sheet.getLastColumn()).clearContent();
  formatarLinha(sheet,linha,new Array(sheet.getLastColumn()).fill(""),sheet.getLastColumn());
}

//================================================================
// §F39 ── doGet — Web App QR Code Pix
//================================================================
function doGet(e) {
  const pix=e.parameter.pix||"",valor=e.parameter.valor||"";
  const qr="https://api.qrserver.com/v1/create-qr-code/?size=340x340&data="+encodeURIComponent(pix);
  const html="<!DOCTYPE html><html lang='pt-BR'><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Pagar Pix</title><style>body{font-family:Arial,sans-serif;background:#000;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;padding:38px;box-sizing:border-box}.card{background:#fff;border-radius:100px;padding:60px 40px;max-width:680px;width:100%;text-align:center;box-shadow:0 6px 32px rgba(0,0,0,0.12)}h2{color:#000;margin:0 0 6px;font-size:30px}.valor{font-size:60px;font-weight:bold;color:#32BCAD;margin:16px 0 36px}img{border:1px solid #eee;border-radius:16px;margin-bottom:28px;width:340px;height:340px}.codigo{background:#fff;border:5px solid #000;border-radius:10px;padding:18px;font-size:14px;font-family:monospace;word-break:break-all;color:#000;margin-bottom:24px;text-align:left}.btn{display:block;background:#32BCAD;color:#000;border:none;border-radius:12px;padding:22px;font-size:20px;font-weight:bold;cursor:pointer;width:100%;margin-bottom:12px}.btn:active{background:#289e90}.ok{color:#32BCAD;font-weight:bold;font-size:18px;min-height:28px;margin-top:8px}.info{font-size:20px;color:#000;margin-top:24px}</style></head><body><div class='card'><h2>Pagar via Pix ❖</h2><div class='valor'>R$"+valor+"</div><img src='"+qr+"' alt='QR Code Pix'/><p style='font-size:20px;color:#000;margin:0 0 8px'>Pix copia e cola:</p><div class='codigo' id='cd'>"+pix+"</div><button class='btn' onclick='cp()'>📋 Copiar código Pix</button><div class='ok' id='ok'></div><div class='info'>Abra seu banco, escolha Pix → copia e cola, e cole o código</div></div><script>function cp(){var t=document.getElementById('cd').innerText;if(navigator.clipboard){navigator.clipboard.writeText(t).then(function(){document.getElementById('ok').innerText='✅ Copiado! Cole no seu banco.';});}else{var x=document.createElement('textarea');x.value=t;document.body.appendChild(x);x.select();document.execCommand('copy');document.body.removeChild(x);document.getElementById('ok').innerText='✅ Copiado! Cole no seu banco.';}}</script></body></html>";
  return HtmlService.createHtmlOutput(html).setTitle("Pagar Pix").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
