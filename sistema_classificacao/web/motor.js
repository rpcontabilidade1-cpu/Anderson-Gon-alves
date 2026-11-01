/* Motor de classificação IBS/CBS/IS — mesma lógica de classificar.py.
   Funciona no navegador (global Motor) e no Node (module.exports) para testes. */
(function (root) {
  "use strict";
  var ST_OK = "OK", ST_CONF = "A CONFIRMAR", ST_REV = "REVISAR", ST_NCM = "REVISAR NCM";
  var CAT_IS = ["Veículos", "Embarcações e aeronaves", "Produtos fumígenos", "Bebidas alcoólicas",
    "Bebidas açucaradas", "Bens minerais", "Concursos de prognósticos e fantasy sport"];
  var REGIMES = {
    simples: "Simples Nacional – Regime Único e Regime Híbrido (duas visões)",
    mei: "MEI", presumido: "Lucro Presumido", real: "Lucro Real", "": "não informado"
  };
  var OBS_SIMPLES = "Simples puro: segregar a receita deste item no PGDAS-D para refletir o benefício. " +
    "Simples híbrido: benefício aplicado diretamente na alíquota do item.";

  // ------------------------------------------------------------------ utilidades
  function txt(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "object") {
      if (v instanceof Date) return v.toISOString().slice(0, 10);
      if ("result" in v) return txt(v.result);
      if ("text" in v) return txt(v.text);
      if (v.richText) return txt(v.richText.map(function (t) { return t.text; }).join(""));
    }
    if (typeof v === "number" && Number.isInteger(v)) return String(v);
    return String(v).split(/\s+/).filter(Boolean).join(" ");
  }
  function norm(s) {
    s = txt(s).normalize("NFKD").replace(/[̀-ͯ]/g, "").toLowerCase();
    return s.replace(/[^a-z0-9/ -]/g, " ").split(/\s+/).filter(Boolean).join(" ");
  }
  function digitos(v) { return txt(v).replace(/\D/g, ""); }
  function lista(v) { return txt(v).split(";").map(function (x) { return x.trim(); }).filter(Boolean); }
  function perc(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "string") v = parseFloat(v.replace("%", "").replace(",", ".").trim() || "0");
    v = Number(v) || 0;
    return v > 1 ? v / 100 : v;
  }
  function cod(v, n) { var d = digitos(v); return d ? d.padStart(n, "0") : ""; }
  function upper(v) { return txt(v).toUpperCase(); }

  function regimeChave(v) {
    var n = norm(v);
    if (n.split(" ").indexOf("mei") >= 0 || n.indexOf("microempreendedor") >= 0) return "mei";
    if (n.indexOf("simples") >= 0 || n.indexOf("unico") >= 0 || n.indexOf("hibrid") >= 0) return "simples";
    if (n.indexOf("presumido") >= 0) return "presumido";
    if (n.indexOf("real") >= 0) return "real";
    return "";
  }

  // ------------------------------------------------------------------ base
  // bruto = {Regras_NCM:[{...}], Regras_IS:[...], Tabela_cClassTrib:[...], Tabela_CST:[...], Sinonimos_Colunas:[...], Config:[...]}
  function prepararBase(bruto) {
    function pre(list, f) { return (list || []).map(f); }
    var b = {};
    b.regras = pre(bruto.Regras_NCM, function (r) {
      return {
        ID: txt(r.ID), _ncm: lista(r.NCM).map(digitos).filter(Boolean), _exc: lista(r.EXCETO_NCM).map(digitos).filter(Boolean),
        _nbs: lista(r.NBS).map(digitos).filter(Boolean), _chave: lista(r.PALAVRAS_CHAVE).map(norm),
        _excl: lista(r.PALAVRAS_EXCLUIR).map(norm), DESCRICAO_REGRA: txt(r.DESCRICAO_REGRA), ANEXO: txt(r.ANEXO),
        TRATAMENTO: txt(r.TRATAMENTO), CST: cod(r.CST, 3), CCLASSTRIB: cod(r.CCLASSTRIB, 6), BASE_LEGAL: txt(r.BASE_LEGAL),
        CONDICAO: txt(r.CONDICAO), OBSERVACAO: txt(r.OBSERVACAO), _prio: parseInt(r.PRIORIDADE, 10) || 9,
        _conf: upper(r.CONFERIDO) === "S", _serv: norm(r.TIPO_ITEM).indexOf("serv") === 0
      };
    });
    b.regras_is = pre(bruto.Regras_IS, function (r) {
      return { ID: txt(r.ID), _ncm: lista(r.NCM).map(digitos).filter(Boolean), _exc: lista(r.EXCETO_NCM).map(digitos).filter(Boolean),
        CATEGORIA_IS: txt(r.CATEGORIA_IS), _conf: upper(r.CONFERIDO) === "S" };
    });
    b.cct = {};
    (bruto.Tabela_cClassTrib || []).forEach(function (r) {
      var c = cod(r.CCLASSTRIB, 6);
      if (c.length === 6) b.cct[c] = { desc: txt(r.DESCRICAO), ri: perc(r.PRED_IBS), rc: perc(r.PRED_CBS), conf: upper(r.CONFERIDO) !== "N" };
    });
    b.cst = {};
    (bruto.Tabela_CST || []).forEach(function (r) { if (txt(r.CST)) b.cst[cod(r.CST, 3)] = txt(r.DESCRICAO); });
    b.sin = {};
    (bruto.Sinonimos_Colunas || []).forEach(function (r) {
      lista(r["CABECALHOS_ACEITOS (separados por ;)"]).forEach(function (s) { b.sin[norm(s)] = txt(r.CAMPO); });
    });
    b.cfg = {};
    (bruto.Config || []).forEach(function (r) { b.cfg[txt(r.PARAMETRO)] = txt(r.VALOR); });
    b.totalRegras = b.regras.length;
    b.regrasConferidas = b.regras.filter(function (r) { return r._conf; }).length;
    return b;
  }

  // ------------------------------------------------------------------ leitura
  function detectarCabecalho(linhas, sin) {
    var melhor = {}, idx = -1;
    for (var i = 0; i < Math.min(40, linhas.length); i++) {
      var mapa = {}, row = linhas[i] || [];
      for (var j = 0; j < row.length; j++) {
        var campo = sin[norm(row[j])];
        if (campo && !(campo in mapa)) mapa[campo] = j;
      }
      if (("DESCRICAO" in mapa || "NCM" in mapa) && Object.keys(mapa).length > Object.keys(melhor).length) { melhor = mapa; idx = i; }
    }
    return { idx: idx, mapa: melhor };
  }

  function novoItem(o) {
    var it = { codigo: "", descricao: "", ncm_orig: "", nbs: "", tipo_cli: "", cest: "", ean: "", unidade: "", cfop: "",
      cst_cli: "", cct_cli: "", origem: "" };
    for (var k in o) it[k] = o[k];
    return it;
  }

  // folhas = {nomeAba: [[...linha], ...]}
  function lerFolhas(folhas, base) {
    var itens = [], avisos = [];
    Object.keys(folhas).forEach(function (nome) {
      var linhas = folhas[nome], d = detectarCabecalho(linhas, base.sin);
      if (d.idx < 0) { avisos.push("Aba '" + nome + "': cabeçalho não reconhecido (ignorada)."); return; }
      var m = d.mapa;
      for (var k = d.idx + 1; k < linhas.length; k++) {
        var row = linhas[k] || [];
        var g = function (c) { return (c in m && m[c] < row.length) ? row[m[c]] : null; };
        if (!(txt(g("CODIGO")) || txt(g("DESCRICAO")) || txt(g("NCM")))) continue;
        itens.push(novoItem({
          codigo: txt(g("CODIGO")), descricao: txt(g("DESCRICAO")), ncm_orig: txt(g("NCM")), nbs: digitos(g("NBS")),
          tipo_cli: txt(g("TIPO")), cest: digitos(g("CEST")), ean: digitos(g("EAN")), unidade: txt(g("UNIDADE")),
          cfop: digitos(g("CFOP")), cst_cli: cod(g("CST_CLIENTE"), 3), cct_cli: cod(g("CCLASSTRIB_CLIENTE"), 6),
          origem: nome + "!" + (k + 1)
        }));
      }
      avisos.push("Aba '" + nome + "': cabeçalho na linha " + (d.idx + 1) + "; colunas reconhecidas: " + Object.keys(m).sort().join(", ") + ".");
    });
    return { itens: itens, avisos: avisos };
  }

  // doc = Document (DOMParser) de uma NF-e
  function lerXmlNfe(doc, nomeArquivo) {
    var NS = "http://www.portalfiscal.inf.br/nfe", itens = [];
    var dets = doc.getElementsByTagNameNS(NS, "det");
    function filho(el, tag) { var e = el && el.getElementsByTagNameNS(NS, tag)[0]; return e ? e.textContent : ""; }
    for (var i = 0; i < dets.length; i++) {
      var det = dets[i], p = det.getElementsByTagNameNS(NS, "prod")[0];
      if (!p) continue;
      var ibs = det.getElementsByTagNameNS(NS, "IBSCBS")[0];
      itens.push(novoItem({
        codigo: txt(filho(p, "cProd")), descricao: txt(filho(p, "xProd")), ncm_orig: txt(filho(p, "NCM")),
        cest: digitos(filho(p, "CEST")), ean: digitos(filho(p, "cEAN")), unidade: txt(filho(p, "uCom")), cfop: txt(filho(p, "CFOP")),
        cst_cli: ibs ? cod(filho(ibs, "CST"), 3) : "", cct_cli: ibs ? cod(filho(ibs, "cClassTrib"), 6) : "",
        origem: nomeArquivo + " item " + det.getAttribute("nItem")
      }));
    }
    return itens;
  }

  // ------------------------------------------------------------------ classificação
  function casa(c, prefixos, excecoes) {
    if (!c) return 0;
    for (var i = 0; i < excecoes.length; i++) if (c.indexOf(excecoes[i]) === 0) return 0;
    var n = 0;
    prefixos.forEach(function (p) { if (c.indexOf(p) === 0 && p.length > n) n = p.length; });
    return n;
  }
  function contem(desc, lista) { return lista.some(function (p) { return desc.indexOf(p) >= 0; }); }
  function descrOk(r, desc) {
    if (r._excl.length && contem(desc, r._excl)) return false;
    return !r._chave.length || contem(desc, r._chave);
  }

  function classificar(item, base, regime) {
    var ncm = item.ncm, nbs = item.nbs, desc = norm(item.descricao);
    var servicoCli = norm(item.tipo_cli).indexOf("serv") === 0 || nbs.length === 9;
    var res = { cst: cod(base.cfg.REGRA_PADRAO_CST, 3) || "000", cct: cod(base.cfg.REGRA_PADRAO_CCLASSTRIB, 6) || "000001",
      trat: "Tributação integral", lei: "LC 214/2025, arts. 4º e 11 (regra geral)", cond: "", regra: "PADRÃO",
      status: ST_OK, obs: [], tipo: "Bem", is: "N", cat_is: "", obs_regime: "" };
    var ncmOk = ncm.length === 8, cands = [], aplicada = null;
    if (item.ncm_ajustado) res.obs.push("NCM informado com 7 dígitos (" + item.ncm_orig + "); completado com zero à esquerda: " + ncm + ".");
    base.regras.forEach(function (r) {
      if (ncmOk && !r._serv) {
        var n = casa(ncm, r._ncm, r._exc);
        if (n) cands.push([-n, r._prio, r]);
      } else if (!ncmOk && r._serv) {
        var m = r._nbs.length ? casa(nbs, r._nbs, []) : 0;
        if (m || (!r._nbs.length && r._chave.length && descrOk(r, desc))) cands.push([-m, r._prio, r]);
      }
    });
    cands.sort(function (a, b) { return a[0] - b[0] || a[1] - b[1]; });
    for (var i = 0; i < cands.length; i++) if (descrOk(cands[i][2], desc)) { aplicada = cands[i][2]; break; }

    if (!ncmOk && (aplicada || servicoCli)) res.tipo = "Serviço";
    else if (!ncmOk) {
      res.status = ST_NCM;
      res.obs.push("NCM ausente ou inválido (8 dígitos). Se for serviço, informe a NBS ou o tipo 'Serviço'. Classificação padrão aplicada.");
    }
    if (aplicada) {
      var r = aplicada;
      res.cst = r.CST; res.cct = r.CCLASSTRIB; res.trat = r.TRATAMENTO; res.lei = r.BASE_LEGAL; res.cond = r.CONDICAO; res.regra = r.ID;
      if (r.OBSERVACAO) res.obs.push(r.OBSERVACAO);
      if (res.status !== ST_NCM) {
        if (!r._conf) { res.status = ST_CONF; res.obs.push("Regra " + r.ID + " ainda não conferida na base de regras."); }
        if (!r.CCLASSTRIB) { res.status = ST_CONF; res.obs.push("cClassTrib a definir pela tabela oficial vigente."); }
        else if (!base.cct[r.CCLASSTRIB]) { res.status = ST_CONF; res.obs.push("cClassTrib " + r.CCLASSTRIB + " não consta da tabela de referência."); }
        else if (!base.cct[r.CCLASSTRIB].conf) { res.status = ST_CONF; res.obs.push("cClassTrib " + r.CCLASSTRIB + " pendente de conferência na tabela oficial."); }
      }
    } else if (cands.length) {
      var c = cands[0][2], anexo = (c.ANEXO && c.ANEXO !== "-") ? "Anexo " + c.ANEXO + ", " : "";
      res.status = ST_REV; res.regra = c.ID + " (não aplicada)";
      res.obs.push("Possível benefício: " + c.DESCRICAO_REGRA + " (" + anexo + "regra " + c.ID + "). A descrição não confirma o enquadramento; aplicada tributação integral (critério conservador).");
    }
    if (ncmOk) {
      for (var j = 0; j < base.regras_is.length; j++) {
        var ri = base.regras_is[j];
        if (casa(ncm, ri._ncm, ri._exc)) {
          res.is = "S"; res.cat_is = ri.CATEGORIA_IS;
          res.obs.push("Sujeito ao IS (" + res.cat_is + "); CST-IS/cClassTribIS conforme tabela própria." + (ri._conf ? "" : " Regra IS não conferida."));
          if (res.status === ST_OK && !ri._conf) res.status = ST_CONF;
          break;
        }
      }
    }
    if ((item.cst_cli || item.cct_cli) && (item.cst_cli !== res.cst || item.cct_cli !== res.cct))
      res.obs.push("Cliente usa CST " + (item.cst_cli || "-") + " / cClassTrib " + (item.cct_cli || "-") + ": diverge da proposta.");
    if (res.cct && res.cct.slice(0, 3) !== res.cst) { res.status = ST_REV; res.obs.push("cClassTrib incompatível com o CST (verificar regra)."); }
    if (regime === "simples" && res.trat !== "Tributação integral") res.obs_regime = OBS_SIMPLES;
    else if (regime === "mei") res.obs_regime = "MEI: valores fixos; conferir obrigatoriedade de destaque de IBS/CBS.";
    return res;
  }

  function processar(itens, base, regime) {
    var vistos = {}, unicos = [];
    itens.forEach(function (it) {
      it.ncm = digitos(it.ncm_orig);
      it.ncm_ajustado = it.ncm.length === 7; // NCM gravado como número perde o zero à esquerda
      if (it.ncm_ajustado) it.ncm = "0" + it.ncm;
      var k = (it.codigo || it.descricao) + "\u0001" + it.ncm;
      if (vistos[k]) return;
      vistos[k] = 1; unicos.push(it);
    });
    unicos.forEach(function (it) { it.res = classificar(it, base, regime); });
    return unicos;
  }

  function divergencias(itens) {
    var out = [];
    var risco = {};
    risco[ST_NCM] = ["Rejeição do documento fiscal e enquadramento incorreto em toda a cadeia", "Cliente deve informar/corrigir o NCM (bens) ou a NBS (serviços); reprocessar o item"];
    risco[ST_REV] = ["Recolhimento a maior (benefício não aproveitado) ou glosa/autuação se o benefício for aplicado indevidamente", "Analisar a descrição completa/ficha técnica e confirmar o enquadramento no anexo"];
    risco[ST_CONF] = ["Uso de regra ou código ainda não validado com a tabela oficial vigente", "Escritório deve conferir a regra/código na base de regras antes da importação"];
    itens.forEach(function (it) {
      var st = it.res.status;
      if (st !== ST_OK) out.push([it.codigo, it.descricao, it.ncm || it.ncm_orig, st, it.res.obs.join(" | ")].concat(risco[st]));
    });
    var porCod = {};
    itens.forEach(function (it) { if (it.codigo) (porCod[it.codigo] = porCod[it.codigo] || {})[it.ncm || it.ncm_orig] = 1; });
    Object.keys(porCod).forEach(function (c) {
      var n = Object.keys(porCod[c]).sort();
      if (n.length > 1) out.push([c, "", n.join("; "), ST_REV, "Mesmo código de produto com NCMs diferentes no arquivo do cliente",
        "Cadastro inconsistente; risco de classificação divergente entre documentos", "Cliente deve unificar o NCM do produto"]);
    });
    var porNcm = {};
    itens.forEach(function (it) { if (it.ncm.length === 8) (porNcm[it.ncm] = porNcm[it.ncm] || {})[it.res.cst + "/" + it.res.cct] = 1; });
    Object.keys(porNcm).sort().forEach(function (n) {
      var cl = Object.keys(porNcm[n]).sort();
      if (cl.length > 1) out.push(["(vários)", "", n, ST_REV, "Mesmo NCM com classificações diferentes: " + cl.join("; "),
        "Pode ser correto (descrições distintas), mas exige justificativa item a item", "Conferir a descrição de cada item frente ao texto do anexo"]);
    });
    itens.forEach(function (it) {
      if ((it.cst_cli || it.cct_cli) && (it.cst_cli !== it.res.cst || it.cct_cli !== it.res.cct))
        out.push([it.codigo, it.descricao, it.ncm, it.res.status, "Cliente já usa " + it.cst_cli + "/" + it.cct_cli + "; proposta " + it.res.cst + "/" + it.res.cct,
          "Documentos já emitidos podem estar com classificação incorreta", "Validar e, se for o caso, corrigir o cadastro e avaliar documentos emitidos"]);
    });
    return out;
  }

  var api = { ST_OK: ST_OK, ST_CONF: ST_CONF, ST_REV: ST_REV, ST_NCM: ST_NCM, CAT_IS: CAT_IS, REGIMES: REGIMES, OBS_SIMPLES: OBS_SIMPLES,
    txt: txt, norm: norm, digitos: digitos, cod: cod, regimeChave: regimeChave, prepararBase: prepararBase, lerFolhas: lerFolhas,
    lerXmlNfe: lerXmlNfe, classificar: classificar, processar: processar, divergencias: divergencias };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.Motor = api;
})(this);
