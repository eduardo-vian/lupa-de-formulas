/* Nomes dos argumentos das funções, como o Excel em português mostra. Chave: nome em inglês. */
(function (raiz) {
  "use strict";

  // fixos: argumentos na ordem ("[x]" = opcional). rep: grupo que se repete ("#" vira 1, 2, 3...).
  // padroes: o que o Excel usa quando o argumento é omitido (índice a partir de 0).
  const NUMEROS = { rep: ["número#"] };
  const VALORES = { rep: ["valor#"] };
  const ARRED = { fixos: ["número", "núm_dígitos"] };
  const SERIE = { fixos: ["núm_série"] };
  const TESTE = { fixos: ["valor"] };
  const CRITERIOS = (primeiro) => ({ fixos: [primeiro], rep: ["intervalo_critérios#", "critérios#"] });

  const PT = {
    SUM: NUMEROS, AVERAGE: NUMEROS, MAX: NUMEROS, MIN: NUMEROS, PRODUCT: NUMEROS,
    COUNT: VALORES, COUNTA: VALORES,
    SUMPRODUCT: { rep: ["matriz#"] },
    AND: { rep: ["lógico#"] }, OR: { rep: ["lógico#"] }, NOT: { fixos: ["lógico"] },
    IF: { fixos: ["teste_lógico", "[valor_se_verdadeiro]", "[valor_se_falso]"], padroes: { 2: "FALSO" } },
    IFS: { rep: ["teste_lógico#", "valor_se_verdadeiro#"] },
    SWITCH: { fixos: ["expressão"], rep: ["valor#", "resultado#"] },
    IFERROR: { fixos: ["valor", "valor_se_erro"] },
    IFNA: { fixos: ["valor", "valor_se_nd"] },
    VLOOKUP: {
      fixos: ["valor_procurado", "matriz_tabela", "núm_índice_coluna", "[procurar_intervalo]"],
      padroes: { 3: "VERDADEIRO" },
    },
    HLOOKUP: {
      fixos: ["valor_procurado", "matriz_tabela", "núm_índice_lin", "[procurar_intervalo]"],
      padroes: { 3: "VERDADEIRO" },
    },
    XLOOKUP: {
      fixos: ["pesquisa_valor", "pesquisa_matriz", "matriz_retorno", "[se_não_encontrada]", "[modo_correspondência]", "[modo_pesquisa]"],
      padroes: { 3: "#N/D", 4: "0", 5: "1" },
    },
    INDEX: { fixos: ["matriz", "núm_linha", "[núm_coluna]"] },
    MATCH: { fixos: ["valor_procurado", "matriz_procurada", "[tipo_correspondência]"], padroes: { 2: "1" } },
    XMATCH: { fixos: ["pesquisa_valor", "pesquisa_matriz", "[modo_correspondência]", "[modo_pesquisa]"], padroes: { 2: "0", 3: "1" } },
    OFFSET: {
      fixos: ["ref", "lins", "cols", "[altura]", "[largura]"],
      padroes: { 2: "0", 3: "O mesmo que ref", 4: "O mesmo que ref" },
    },
    INDIRECT: { fixos: ["texto_ref", "[a1]"], padroes: { 1: "VERDADEIRO" } },
    CHOOSE: { fixos: ["núm_índice"], rep: ["valor#"] },
    SUMIF: { fixos: ["intervalo", "critérios", "[intervalo_soma]"], padroes: { 2: "O mesmo que intervalo" } },
    AVERAGEIF: { fixos: ["intervalo", "critérios", "[intervalo_média]"], padroes: { 2: "O mesmo que intervalo" } },
    COUNTIF: { fixos: ["intervalo", "critérios"] },
    SUMIFS: CRITERIOS("intervalo_soma"),
    AVERAGEIFS: CRITERIOS("intervalo_média"),
    MAXIFS: CRITERIOS("intervalo_máximo"),
    MINIFS: CRITERIOS("intervalo_mínimo"),
    COUNTIFS: { rep: ["intervalo_critérios#", "critérios#"] },
    ROUND: ARRED, ROUNDUP: ARRED, ROUNDDOWN: ARRED,
    INT: { fixos: ["núm"] }, ABS: { fixos: ["núm"] }, MOD: { fixos: ["núm", "divisor"] },
    YEAR: SERIE, MONTH: SERIE, DAY: SERIE,
    DATE: { fixos: ["ano", "mês", "dia"] },
    EOMONTH: { fixos: ["data_inicial", "meses"] }, EDATE: { fixos: ["data_inicial", "meses"] },
    TODAY: {},
    NPV: { fixos: ["taxa"], rep: ["valor#"] },
    XNPV: { fixos: ["taxa", "valores", "datas"] },
    IRR: { fixos: ["valores", "[estimativa]"], padroes: { 1: "10%" } },
    XIRR: { fixos: ["valores", "datas", "[estimativa]"], padroes: { 2: "10%" } },
    PMT: { fixos: ["taxa", "nper", "vp", "[vf]", "[tipo]"], padroes: { 3: "0", 4: "0" } },
    PV: { fixos: ["taxa", "nper", "pgto", "[vf]", "[tipo]"], padroes: { 3: "0", 4: "0" } },
    FV: { fixos: ["taxa", "nper", "pgto", "[vp]", "[tipo]"], padroes: { 3: "0", 4: "0" } },
    TEXT: { fixos: ["valor", "formato_texto"] },
    LEFT: { fixos: ["texto", "[núm_caract]"], padroes: { 1: "1" } },
    RIGHT: { fixos: ["texto", "[núm_caract]"], padroes: { 1: "1" } },
    MID: { fixos: ["texto", "núm_inicial", "núm_caract"] },
    LEN: { fixos: ["texto"] },
    CONCAT: { rep: ["texto#"] }, CONCATENATE: { rep: ["texto#"] },
    SUBTOTAL: { fixos: ["núm_função"], rep: ["ref#"] },
    ISERROR: TESTE, ISERR: TESTE, ISNA: TESTE, ISBLANK: TESTE, ISNUMBER: TESTE, ISTEXT: TESTE,
    ROW: { fixos: ["[ref]"] }, COLUMN: { fixos: ["[ref]"] },
  };

  function especificacao(funcao, idioma) {
    return /^pt/i.test(idioma || "") ? PT[funcao] || null : null;
  }

  function nome(funcao, i, idioma) {
    const spec = especificacao(funcao, idioma);
    if (!spec) return "";
    const fixos = spec.fixos || [];
    if (i < fixos.length) return fixos[i].replace(/^\[|\]$/g, "");
    if (!spec.rep) return "";
    const k = i - fixos.length;
    return spec.rep[k % spec.rep.length].replace("#", String(Math.floor(k / spec.rep.length) + 1));
  }

  function padrao(funcao, i, idioma) {
    const spec = especificacao(funcao, idioma);
    return (spec && spec.padroes && spec.padroes[i]) || "";
  }

  // Índices dos argumentos opcionais que ficaram de fora da fórmula.
  function opcionaisAusentes(funcao, informados, idioma) {
    const spec = especificacao(funcao, idioma);
    const fixos = (spec && spec.fixos) || [];
    const lista = [];
    for (let i = informados; i < fixos.length; i++) if (fixos[i].startsWith("[")) lista.push(i);
    return lista;
  }

  const api = { nome, padrao, opcionaisAusentes };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else raiz.Argumentos = api;
})(this);
