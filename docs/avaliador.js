/* Calcula cada parte da fórmula e monta a árvore exibida na janela.
   Não fala com o Excel diretamente: recebe um "adaptador" com ler(), chamar() e metodo(). */
(function (raiz) {
  "use strict";

  const carregar = (global, caminho) =>
    typeof module !== "undefined" && module.exports ? require(caminho) : raiz[global];
  const E = carregar("Enderecos", "./enderecos.js");
  const F = carregar("Formula", "./formula.js");
  const A = carregar("Argumentos", "./argumentos.js");
  const B = carregar("Blocos", "./blocos.js");

  const MAX_RODADAS = 60;
  const MAX_PROFUNDIDADE_NOMES = 3; // nome dentro de nome dentro de nome...
  const ERROS_PT = {
    "#N/A": "#N/D", "#VALUE!": "#VALOR!", "#NAME?": "#NOME?",
    "#NUM!": "#NÚM!", "#NULL!": "#NULO!", "#SPILL!": "#DESPEJAR!",
  };
  const COMPARACOES = ["=", "<>", "<", ">", "<=", ">="];
  const TIPOS_VALIDACAO = {
    List: "lista", WholeNumber: "número inteiro", Decimal: "decimal", Date: "data",
    Time: "hora", TextLength: "tamanho do texto", Custom: "personalizada",
  };

  // ---------- Valores ----------
  // num, txt, bool, erro, vazio, ref (intervalo lido do Excel) e nd (não deu para calcular).

  const num = (v) => ({ t: "num", v: Object.is(v, -0) ? 0 : v });
  const txt = (v) => ({ t: "txt", v });
  const bool = (v) => ({ t: "bool", v });
  const erro = (v) => ({ t: "erro", v });
  const nd = (motivo) => ({ t: "nd", motivo });
  const VAZIO = { t: "vazio" };

  const ehPt = (idioma) => /^pt/i.test(idioma || "");
  const ruim = (v) => v.t === "erro" || v.t === "nd";

  // Intervalo de uma célula vira o valor da célula; intervalo maior não tem valor único.
  function escalar(val) {
    if (val.t !== "ref") return val;
    if (!val.celulas) return nd("intervalo grande");
    if (val.celulas.length === 1 && val.celulas[0].length === 1) return val.celulas[0][0];
    return nd("intervalo com várias células");
  }

  // Devolve um número, ou o valor de erro/nd que impede o cálculo.
  function numero(val) {
    const e = escalar(val);
    switch (e.t) {
      case "num": return e.v;
      case "bool": return e.v ? 1 : 0;
      case "vazio": return 0;
      case "txt": {
        const limpo = e.v.trim().replace(",", ".");
        const n = Number(limpo);
        return limpo !== "" && Number.isFinite(n) ? n : erro("#VALUE!");
      }
      default: return e;
    }
  }

  function texto(val, idioma) {
    const e = escalar(val);
    switch (e.t) {
      case "txt": return e.v;
      case "num": return formatoGeral(e.v, idioma);
      case "bool": return nomeLogico(e.v, idioma);
      case "vazio": return "";
      default: return e;
    }
  }

  function logico(val) {
    const e = escalar(val);
    switch (e.t) {
      case "bool": return e.v;
      case "num": return e.v !== 0;
      case "vazio": return false;
      case "txt": {
        const s = e.v.trim().toUpperCase();
        if (s === "TRUE" || s === "VERDADEIRO") return true;
        if (s === "FALSE" || s === "FALSO") return false;
        return erro("#VALUE!");
      }
      default: return e;
    }
  }

  // Inteiro de um argumento; argumento ausente ou vazio vira o padrão.
  function inteiroOu(val, padrao) {
    if (val === undefined || val.t === "vazio") return padrao;
    const n = numero(val);
    return typeof n === "number" ? Math.trunc(n) : n;
  }

  function aplicarOperador(op, a, b, idioma) {
    if (op === "&") {
      const x = texto(a, idioma);
      if (typeof x !== "string") return x;
      const y = texto(b, idioma);
      return typeof y === "string" ? txt(x + y) : y;
    }
    if (COMPARACOES.includes(op)) return comparar(op, a, b, idioma);
    const x = numero(a);
    if (typeof x !== "number") return x;
    const y = numero(b);
    if (typeof y !== "number") return y;
    switch (op) {
      case "+": return num(x + y);
      case "-": return num(x - y);
      case "*": return num(x * y);
      case "/": return y === 0 ? erro("#DIV/0!") : num(x / y);
      case "^": {
        const r = Math.pow(x, y);
        return Number.isFinite(r) ? num(r) : erro("#NUM!");
      }
      default: return nd(`operador ${op}`);
    }
  }

  // Ordem do Excel entre tipos diferentes: número < texto < lógico. Texto sem diferenciar maiúsculas.
  function comparar(op, a, b, idioma) {
    let x = escalar(a);
    let y = escalar(b);
    for (const v of [x, y]) if (ruim(v)) return v;
    if (x.t === "vazio") x = vazioComo(y);
    if (y.t === "vazio") y = vazioComo(x);
    const ordem = { num: 0, txt: 1, bool: 2 };
    let c;
    if (x.t !== y.t) c = ordem[x.t] - ordem[y.t];
    else if (x.t === "txt") c = x.v.localeCompare(y.v, idioma || undefined, { sensitivity: "base" });
    else c = Number(x.v) - Number(y.v);
    return bool(comparaNum(op, c, 0));
  }

  function vazioComo(outro) {
    if (outro.t === "txt") return txt("");
    if (outro.t === "bool") return bool(false);
    return num(0);
  }

  function comparaNum(op, a, b) {
    switch (op) {
      case "=": return a === b;
      case "<>": return a !== b;
      case "<": return a < b;
      case ">": return a > b;
      case "<=": return a <= b;
      case ">=": return a >= b;
      default: return false;
    }
  }

  // ---------- Busca (PROCV, PROCH, CORRESP, PROCX, CORRESPX) ----------

  const escapar = (ch) => ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

  // Curingas do Excel: * qualquer sequência, ? um caractere, ~ antes de * ? ~ para o próprio caractere.
  function curinga(padrao) {
    let re = "";
    for (let i = 0; i < padrao.length; i++) {
      const ch = padrao[i];
      if (ch === "~" && i + 1 < padrao.length) re += escapar(padrao[++i]);
      else re += ch === "*" ? "[\\s\\S]*" : ch === "?" ? "[\\s\\S]" : escapar(ch);
    }
    return new RegExp(`^${re}$`, "iu");
  }

  const mesmoTipo = (a, b) => a.t === b.t && (a.t === "num" || a.t === "txt" || a.t === "bool");

  function cmp(a, b) {
    return a.t === "txt" ? a.v.localeCompare(b.v, undefined, { sensitivity: "base" }) : Number(a.v) - Number(b.v);
  }

  function iguaisNaBusca(x, v, comCuringa) {
    if (!mesmoTipo(x, v)) return false;
    if (x.t === "txt") return comCuringa ? curinga(v.v).test(x.v) : cmp(x, v) === 0;
    return x.v === v.v;
  }

  // modo: "exato", "curinga", "crescente" (maior valor <= procurado, dados em ordem crescente)
  // ou "decrescente" (menor valor >= procurado, dados em ordem decrescente). Devolve o índice ou -1.
  function procurar(lista, v, modo) {
    if (modo === "exato" || modo === "curinga") {
      return lista.findIndex((x) => iguaisNaBusca(x, v, modo === "curinga"));
    }
    let candidato = -1;
    for (let i = 0; i < lista.length; i++) {
      if (!mesmoTipo(lista[i], v)) continue;
      const c = cmp(lista[i], v);
      if (modo === "crescente" ? c <= 0 : c >= 0) candidato = i;
      else break;
    }
    return candidato;
  }

  // PROCX/CORRESPX: modo 0 exato, -1 exato ou o próximo menor, 1 exato ou o próximo maior, 2 curinga.
  function procurarX(lista, ordem, v, modo) {
    for (const i of ordem) if (iguaisNaBusca(lista[i], v, modo === 2)) return i;
    if (modo !== -1 && modo !== 1) return -1;
    let melhor = -1;
    for (const i of ordem) {
      if (!mesmoTipo(lista[i], v)) continue;
      const c = cmp(lista[i], v);
      if (modo === -1 ? c >= 0 : c <= 0) continue;
      if (melhor < 0 || (modo === -1 ? cmp(lista[i], lista[melhor]) > 0 : cmp(lista[i], lista[melhor]) < 0)) melhor = i;
    }
    return melhor;
  }

  // i-ésima célula de um vetor (coluna ou linha).
  function celulaEm(ref, i) {
    const vertical = ref.c1 === ref.c2;
    return { folha: ref.folha, r: vertical ? ref.r1 + i : ref.r1, c: vertical ? ref.c1 : ref.c1 + i };
  }

  function lerCelula(folha, r, c) {
    return { ler: { folha, ref: E.formatarRef({ r1: r, c1: c, r2: r, c2: c }) } };
  }

  function procv(args, sentido) {
    const [valor, tabela, indice, aproximado] = args;
    if (!tabela) return erro("#VALUE!");
    if (ruim(tabela)) return tabela;
    if (tabela.t !== "ref" || !tabela.celulas) return { excel: true };
    const v = escalar(valor || VAZIO);
    if (ruim(v)) return v;
    const n = inteiroOu(indice, 0);
    if (typeof n !== "number") return n;
    const tamanho = sentido === "linha" ? tabela.c2 - tabela.c1 + 1 : tabela.r2 - tabela.r1 + 1;
    if (n < 1) return erro("#VALUE!");
    if (n > tamanho) return erro("#REF!");
    // 4º argumento ausente = VERDADEIRO (aproximado); presente e vazio = FALSO (exato).
    const exato = aproximado === undefined ? false : aproximado.t === "vazio" || logico(aproximado) === false;
    const lista = sentido === "linha" ? tabela.celulas.map((linha) => linha[0]) : tabela.celulas[0];
    const i = procurar(lista, v, exato ? "curinga" : "crescente");
    if (i < 0) return erro("#N/A");
    return sentido === "linha"
      ? lerCelula(tabela.folha, tabela.r1 + i, tabela.c1 + n - 1)
      : lerCelula(tabela.folha, tabela.r1 + n - 1, tabela.c1 + i);
  }

  function vetor(ref) {
    if (ref.r1 !== ref.r2 && ref.c1 !== ref.c2) return null;
    return ref.celulas.flat();
  }

  // ---------- Critérios (SOMASE, SOMASES, CONT.SES...) ----------

  function numeroDeTexto(s) {
    const m = /^([+-]?\d+(?:[.,]\d+)?(?:[eE][+-]?\d+)?)(%?)$/.exec(s.trim());
    if (!m) return null;
    const n = Number(m[1].replace(",", "."));
    return m[2] ? n / 100 : n;
  }

  // Critério do Excel ("<>b", ">=10", "a*", 5...) -> função que testa o valor de uma célula.
  function criterioDe(val) {
    const e = escalar(val || VAZIO);
    if (e.t === "num") return (c) => (c.t === "num" ? c.v === e.v : c.t === "txt" && numeroDeTexto(c.v) === e.v);
    if (e.t === "bool") return (c) => c.t === "bool" && c.v === e.v;
    if (e.t === "vazio") return (c) => c.t === "vazio" || (c.t === "txt" && c.v === "");
    if (e.t !== "txt") return null;
    const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(e.v);
    const op = m[1] || "=";
    const resto = m[2];
    const n = numeroDeTexto(resto);
    if (n !== null) return (c) => (c.t === "num" ? comparaNum(op, c.v, n) : op === "<>");
    if (op === "=" || op === "<>") {
      const padrao = resto === "" ? null : curinga(resto);
      const igual = (c) => (padrao ? c.t === "txt" && padrao.test(c.v) : c.t === "vazio" || (c.t === "txt" && c.v === ""));
      return op === "=" ? igual : (c) => !igual(c);
    }
    return (c) => c.t === "txt" && comparaNum(op, c.v.localeCompare(resto, undefined, { sensitivity: "base" }), 0);
  }

  const valorEm = (ref, i, j) => (ref.celulas && ref.celulas[i] && ref.celulas[i][j]) || VAZIO;

  function celulasDe(ref) {
    const saida = [];
    ref.celulas.forEach((linha, i) => linha.forEach((v, j) => saida.push({ folha: ref.folha, r: ref.r1 + i, c: ref.c1 + j, v })));
    return saida;
  }

  // SOMASE(intervalo; critérios; [intervalo_soma]) e parecidas.
  function filtroSe(intervalo, criterio, soma) {
    if (!intervalo || intervalo.t !== "ref" || !intervalo.celulas || !criterio) return null;
    const teste = criterioDe(criterio);
    if (!teste) return null;
    const destino = soma && soma.t === "ref" ? soma : intervalo;
    const saida = [];
    intervalo.celulas.forEach((linha, i) => linha.forEach((v, j) => {
      if (teste(v)) saida.push({ folha: destino.folha, r: destino.r1 + i, c: destino.c1 + j, v: valorEm(destino, i, j) });
    }));
    return saida;
  }

  // SOMASES(destino; intervalo1; critério1; ...): posições que passam em todos os critérios.
  function filtroSes(destino, pares) {
    if (!destino || destino.t !== "ref" || !destino.celulas) return null;
    const testes = [];
    for (let k = 0; k + 1 < pares.length; k += 2) {
      const intervalo = pares[k];
      const teste = criterioDe(pares[k + 1]);
      if (!intervalo || intervalo.t !== "ref" || !intervalo.celulas || !teste) return null;
      testes.push({ intervalo, teste });
    }
    if (!testes.length) return null;
    const saida = [];
    destino.celulas.forEach((linha, i) => linha.forEach((v, j) => {
      if (testes.every(({ intervalo, teste }) => teste(valorEm(intervalo, i, j)))) {
        saida.push({ folha: destino.folha, r: destino.r1 + i, c: destino.c1 + j, v });
      }
    }));
    return saida;
  }

  // Células que têm o maior (ou menor) valor numérico da lista.
  function extremo(celulas, maior) {
    if (!celulas) return null;
    const numeros = celulas.filter((c) => c.v.t === "num");
    if (!numeros.length) return null;
    const alvo = numeros.reduce((m, c) => (maior ? Math.max(m, c.v.v) : Math.min(m, c.v.v)), numeros[0].v.v);
    return numeros.filter((c) => c.v.v === alvo);
  }

  function celulasNumericas(args) {
    if (args.some((a) => a.t === "ref" && !a.celulas)) return null;
    return args.filter((a) => a.t === "ref").flatMap(celulasDe);
  }

  // Funções que vão para o Excel calcular, mas cujas células "escolhidas" a Lupa destaca.
  const FILTROS = {
    SUMIF: (a) => filtroSe(a[0], a[1], a[2]),
    AVERAGEIF: (a) => filtroSe(a[0], a[1], a[2]),
    COUNTIF: (a) => filtroSe(a[0], a[1]),
    SUMIFS: (a) => filtroSes(a[0], a.slice(1)),
    AVERAGEIFS: (a) => filtroSes(a[0], a.slice(1)),
    COUNTIFS: (a) => filtroSes(a[0], a),
    MAXIFS: (a) => extremo(filtroSes(a[0], a.slice(1)), true),
    MINIFS: (a) => extremo(filtroSes(a[0], a.slice(1)), false),
    MAX: (a) => extremo(celulasNumericas(a), true),
    MIN: (a) => extremo(celulasNumericas(a), false),
  };

  // ---------- Funções calculadas pela própria Lupa ----------
  // Devolvem um valor, ou { val, ler, excel, alvos, escolhido }:
  //   ler: outro intervalo a ler (DESLOC, ÍNDICE, PROCV...)     excel: deixa o Excel calcular
  //   alvos: células a destacar                                  escolhido: argumento em negrito

  const PROPRIAS = {
    IF(args) {
      const c = logico(args[0] || VAZIO);
      if (typeof c !== "boolean") return c;
      const i = c ? 1 : 2;
      const escolhido = args[i];
      if (escolhido === undefined) return c ? num(0) : bool(false);
      return { val: escolhido.t === "vazio" ? num(0) : escolhido, escolhido: i };
    },
    IFS(args) {
      for (let i = 0; i + 1 < args.length; i += 2) {
        const c = logico(args[i]);
        if (typeof c !== "boolean") return c;
        if (c) return { val: args[i + 1].t === "vazio" ? num(0) : args[i + 1], escolhido: i + 1 };
      }
      return erro("#N/A");
    },
    SWITCH(args) {
      const alvo = escalar(args[0] || VAZIO);
      if (ruim(alvo)) return alvo;
      for (let i = 1; i + 1 < args.length; i += 2) {
        const igual = comparar("=", alvo, args[i]);
        if (igual.t === "bool" && igual.v) return { val: args[i + 1], escolhido: i + 1 };
      }
      if (args.length > 1 && (args.length - 1) % 2 === 1) return { val: args[args.length - 1], escolhido: args.length - 1 };
      return erro("#N/A");
    },
    CHOOSE(args) {
      const k = inteiroOu(args[0] || VAZIO, 0);
      if (typeof k !== "number") return k;
      if (k < 1 || k >= args.length) return erro("#VALUE!");
      return { val: args[k].t === "vazio" ? num(0) : args[k], escolhido: k };
    },
    IFERROR: (args) => seErro(args, () => true),
    IFNA: (args) => seErro(args, (v) => v === "#N/A"),
    NOT(args) {
      const v = logico(args[0] || VAZIO);
      return typeof v === "boolean" ? bool(!v) : v;
    },
    AND: (args) => juntarLogicos(args, (a, b) => a && b, true),
    OR: (args) => juntarLogicos(args, (a, b) => a || b, false),
    ISERROR: (args) => testarValor(args, (v) => v.t === "erro"),
    ISERR: (args) => testarValor(args, (v) => v.t === "erro" && v.v !== "#N/A"),
    ISNA: (args) => testarValor(args, (v) => v.t === "erro" && v.v === "#N/A"),
    ISBLANK: (args) => testarValor(args, (v) => v.t === "vazio" && args[0].t === "ref"),
    ISNUMBER: (args) => testarValor(args, (v) => v.t === "num"),
    ISTEXT: (args) => testarValor(args, (v) => v.t === "txt"),
    OFFSET: deslocar,
    INDIRECT: indireto,
    INDEX: indice,
    VLOOKUP: (args) => procv(args, "linha"),
    HLOOKUP: (args) => procv(args, "coluna"),
    MATCH(args) {
      const [valor, matriz, tipo] = args;
      if (!matriz) return erro("#VALUE!");
      if (ruim(matriz)) return matriz;
      if (matriz.t !== "ref" || !matriz.celulas) return { excel: true };
      const lista = vetor(matriz);
      if (!lista) return erro("#N/A");
      const v = escalar(valor || VAZIO);
      if (ruim(v)) return v;
      // 3º argumento ausente = 1; presente e vazio = 0.
      const t = tipo === undefined ? 1 : inteiroOu(tipo, 0);
      if (typeof t !== "number") return t;
      const i = procurar(lista, v, t === 0 ? "curinga" : t > 0 ? "crescente" : "decrescente");
      return i < 0 ? erro("#N/A") : { val: num(i + 1), alvos: [celulaEm(matriz, i)] };
    },
    XLOOKUP(args) {
      const [valor, procura, retorno, seNao, modoArg, buscaArg] = args;
      if (!procura || !retorno) return erro("#VALUE!");
      for (const x of [procura, retorno]) if (ruim(x)) return x;
      if (procura.t !== "ref" || retorno.t !== "ref" || !procura.celulas) return { excel: true };
      const lista = vetor(procura);
      if (!lista) return erro("#VALUE!");
      const v = escalar(valor || VAZIO);
      if (ruim(v)) return v;
      const modo = inteiroOu(modoArg, 0);
      const busca = inteiroOu(buscaArg, 1);
      for (const n of [modo, busca]) if (typeof n !== "number") return n;
      const ordem = lista.map((_, k) => (busca < 0 ? lista.length - 1 - k : k));
      const i = procurarX(lista, ordem, v, modo);
      if (i < 0) return seNao !== undefined && seNao.t !== "vazio" ? seNao : erro("#N/A");
      const vertical = procura.c1 === procura.c2;
      const ret = vertical
        ? { r1: retorno.r1 + i, r2: retorno.r1 + i, c1: retorno.c1, c2: retorno.c2 }
        : { r1: retorno.r1, r2: retorno.r2, c1: retorno.c1 + i, c2: retorno.c1 + i };
      return { ler: { folha: retorno.folha, ref: E.formatarRef(ret) } };
    },
    XMATCH(args) {
      const [valor, procura, modoArg, buscaArg] = args;
      if (!procura) return erro("#VALUE!");
      if (ruim(procura)) return procura;
      if (procura.t !== "ref" || !procura.celulas) return { excel: true };
      const lista = vetor(procura);
      if (!lista) return erro("#VALUE!");
      const v = escalar(valor || VAZIO);
      if (ruim(v)) return v;
      const modo = inteiroOu(modoArg, 0);
      const busca = inteiroOu(buscaArg, 1);
      for (const n of [modo, busca]) if (typeof n !== "number") return n;
      const ordem = lista.map((_, k) => (busca < 0 ? lista.length - 1 - k : k));
      const i = procurarX(lista, ordem, v, modo);
      return i < 0 ? erro("#N/A") : { val: num(i + 1), alvos: [celulaEm(procura, i)] };
    },
  };

  function seErro(args, condicao) {
    const v = args[0] || VAZIO;
    const e = escalar(v);
    if (e.t === "nd") return e;
    if (e.t === "erro" && condicao(e.v)) {
      const alternativa = args[1] || VAZIO;
      return { val: alternativa.t === "vazio" ? num(0) : alternativa, escolhido: 1 };
    }
    return { val: v.t === "vazio" ? num(0) : v, escolhido: 0 };
  }

  function testarValor(args, teste) {
    const e = escalar(args[0] || VAZIO);
    return e.t === "nd" ? e : bool(teste(e));
  }

  // E/OU: dentro de intervalos, textos e vazios são ignorados, como no Excel.
  function juntarLogicos(args, juntar, inicial) {
    let acumulado = inicial;
    let achou = false;
    for (const a of args) {
      const valores = a.t === "ref" ? (a.celulas ? a.celulas.flat() : null) : [a];
      if (!valores) return nd("intervalo grande");
      for (const v of valores) {
        if (ruim(v)) return v;
        if (v.t === "vazio" || (v.t === "txt" && a.t === "ref")) continue;
        const l = logico(v);
        if (typeof l !== "boolean") return l;
        acumulado = juntar(acumulado, l);
        achou = true;
      }
    }
    return achou ? bool(acumulado) : erro("#VALUE!");
  }

  function deslocar(args) {
    const base = args[0] || VAZIO;
    if (ruim(base)) return base;
    if (base.t !== "ref") return erro("#VALUE!");
    const lins = inteiroOu(args[1], 0);
    const cols = inteiroOu(args[2], 0);
    const altura = inteiroOu(args[3], base.r2 - base.r1 + 1);
    const largura = inteiroOu(args[4], base.c2 - base.c1 + 1);
    for (const n of [lins, cols, altura, largura]) if (typeof n !== "number") return n;
    if (altura < 1 || largura < 1) return nd("altura ou largura negativa");
    const r1 = base.r1 + lins;
    const c1 = base.c1 + cols;
    const ret = { r1, c1, r2: r1 + altura - 1, c2: c1 + largura - 1 };
    if (r1 < 1 || c1 < 1 || ret.r2 > E.MAX_LINHAS || ret.c2 > E.MAX_COLUNAS) return erro("#REF!");
    return { ler: { folha: base.folha, ref: E.formatarRef(ret) } };
  }

  function indireto(args, ctx) {
    const t = texto(args[0] || VAZIO, ctx.idioma);
    if (typeof t !== "string") return t;
    if (args[1] && args[1].t !== "vazio" && logico(args[1]) === false) return nd("estilo L1C1");
    const { folha, ref } = E.separar(t.trim());
    const ret = E.analisarRef(ref);
    if (ret) return { ler: { folha: folha || ctx.folhaBase, ref: E.formatarRef(ret) } };
    if (/^[\p{L}_\\][\p{L}\p{N}_.]*$/u.test(ref)) return { ler: { nome: ref, folha } };
    return erro("#REF!");
  }

  // ÍNDICE(matriz; linha; [coluna]): devolve a célula (ou a linha/coluna inteira, com 0).
  function indice(args) {
    const [matriz, linArg, colArg] = args;
    if (!matriz) return erro("#VALUE!");
    if (ruim(matriz)) return matriz;
    if (matriz.t !== "ref") return { excel: true };
    const altura = matriz.r2 - matriz.r1 + 1;
    const largura = matriz.c2 - matriz.c1 + 1;
    let lin = inteiroOu(linArg, 0);
    let col = inteiroOu(colArg, 0);
    for (const n of [lin, col]) if (typeof n !== "number") return n;
    // Vetor horizontal com um número só: o número é a coluna.
    if (colArg === undefined && altura === 1 && largura > 1) {
      col = lin;
      lin = 1;
    }
    if (lin === 0 && altura === 1) lin = 1;
    if (col === 0 && largura === 1) col = 1;
    if (lin < 0 || col < 0 || lin > altura || col > largura) return erro("#REF!");
    const ret = {
      r1: lin ? matriz.r1 + lin - 1 : matriz.r1,
      r2: lin ? matriz.r1 + lin - 1 : matriz.r2,
      c1: col ? matriz.c1 + col - 1 : matriz.c1,
      c2: col ? matriz.c1 + col - 1 : matriz.c2,
    };
    return { ler: { folha: matriz.folha, ref: E.formatarRef(ret) } };
  }

  // ---------- Cálculo em rodadas ----------
  // Cada rodada calcula tudo o que já é possível e junta as leituras e chamadas ao Excel que
  // faltam, para mandar de uma vez só. Repete até não sobrar nada pendente.

  async function avaliar(ast, adaptador, ctx) {
    const cache = new Map();
    for (let rodada = 0; rodada < MAX_RODADAS; rodada++) {
      const est = { cache, leituras: new Map(), chamadas: [], adaptador, ctx, novos: false };
      visitar(ast, est);
      if (!est.leituras.size && !est.chamadas.length && !est.novos) break;
      if (est.leituras.size) {
        const lidos = await adaptador.ler([...est.leituras.values()]);
        for (const chave of est.leituras.keys()) cache.set(chave, lidos.get(chave) || nd("não lido"));
      }
      if (est.chamadas.length) {
        const resultados = await adaptador.chamar(est.chamadas.map(({ metodo, args }) => ({ metodo, args })));
        est.chamadas.forEach((c, i) => {
          c.no.val = resultados[i] || nd("sem resultado");
        });
      }
    }
    completar(ast);
    return ast;
  }

  // Filhos na árvore. O nome definido mostra a definição como filho, mas não depende dela.
  function filhosDe(no) {
    switch (no.tipo) {
      case "operacao":
      case "intervalo": return no.operandos;
      case "neg":
      case "pct": return [no.filho];
      case "func": return no.args;
      case "nome": return no.definicao ? [no.definicao.ast] : [];
      default: return [];
    }
  }

  const dependenciasDe = (no) => (no.tipo === "nome" ? [] : filhosDe(no));

  function visitar(no, est) {
    filhosDe(no).forEach((f) => visitar(f, est));
    if (no.val !== undefined) return;
    if (dependenciasDe(no).some((f) => f.val === undefined)) return;
    const val = calcularNo(no, est);
    if (val !== undefined) no.val = val;
  }

  function completar(no) {
    filhosDe(no).forEach(completar);
    if (no.val === undefined) no.val = nd("não calculado");
  }

  function calcularNo(no, est) {
    const idioma = est.ctx.idioma;
    switch (no.tipo) {
      case "num": return num(Number(no.texto));
      case "txt": return txt(no.texto.slice(1, -1).replace(/""/g, '"'));
      case "logico": return bool(/^(TRUE|VERDADEIRO)$/i.test(no.texto));
      case "erro": return erro(E.separar(no.texto).ref.toUpperCase());
      case "vazio": return VAZIO;
      case "matriz": return nd("matriz constante");
      case "estruturada": return nd("referência de tabela");
      case "ref": return referencia(no, est);
      case "nome": {
        const { folha, ref } = E.separar(no.texto);
        const val = pedir(pedidoDeNome(ref, folha), est);
        if (val && val.definicao) anexarDefinicao(no, val.definicao, est);
        return val;
      }
      case "neg": {
        const x = numero(no.filho.val);
        return typeof x === "number" ? num(no.sinal === "-" ? -x : x) : x;
      }
      case "pct": {
        const x = numero(no.filho.val);
        return typeof x === "number" ? num(x / 100) : x;
      }
      case "operacao":
        return no.ops.reduce((acc, op, i) => aplicarOperador(op, acc, no.operandos[i + 1].val, idioma), no.operandos[0].val);
      case "intervalo": return unir(no, est);
      case "func": return funcao(no, est);
      default: return nd("não suportado");
    }
  }

  function anexarDefinicao(no, definicao, est) {
    const profundidade = (no.profundidade || 0) + 1;
    if (no.definicao || profundidade > MAX_PROFUNDIDADE_NOMES) return;
    try {
      const tokens = F.tokenizar(definicao);
      const ast = F.analisar(tokens);
      const marcar = (n) => {
        n.profundidade = profundidade;
        filhosDe(n).forEach(marcar);
      };
      marcar(ast);
      no.definicao = { ast, tokens, texto: definicao };
      est.novos = true;
    } catch (e) {
      // definição que a Lupa não sabe ler: o nome fica sem o detalhe
    }
  }

  function referencia(no, est) {
    const { folha, ref } = E.separar(no.texto);
    if (folha && folha.startsWith("[")) return nd("referência a outro arquivo");
    if (ref.endsWith("#")) return nd("intervalo despejado (#)");
    const ret = E.analisarRef(ref);
    if (!ret) return erro("#REF!");
    return pedir(pedidoDeRef(folha || est.ctx.folhaBase, E.formatarRef(ret)), est);
  }

  // "A1:DESLOC(...)": junta as duas pontas num intervalo só.
  function unir(no, est) {
    const vals = no.operandos.map((o) => o.val);
    const problema = vals.find(ruim);
    if (problema) return problema;
    if (vals.some((v) => v.t !== "ref" || v.folha !== vals[0].folha)) return erro("#VALUE!");
    return pedir(pedidoDeRef(vals[0].folha, E.formatarRef(E.caixa(vals))), est);
  }

  function funcao(no, est) {
    if (no.aguardandoExcel) return undefined;
    const args = no.args.map((a) => a.val);
    const propria = PROPRIAS[no.nome];
    if (propria) {
      const r = propria(args, est.ctx);
      const res = r && r.t ? { val: r } : r || { val: nd("sem resultado") };
      if (res.escolhido !== undefined) no.escolhido = res.escolhido;
      if (res.alvos) no.alvos = res.alvos;
      if (!res.excel) return res.ler ? pedir(pedidoDe(res.ler), est) : res.val;
    }
    const filtro = FILTROS[no.nome];
    if (filtro && !no.alvos) {
      try {
        no.alvos = filtro(args) || undefined;
      } catch (e) {
        no.alvos = undefined;
      }
    }
    const metodo = est.adaptador.metodo(no.nome);
    if (!metodo) return nd("função que a Lupa ainda não calcula");
    const problema = args.find(ruim);
    if (problema) return problema;
    const lista = args.map(paraArgumento);
    while (lista.length && lista[lista.length - 1] === undefined) lista.pop();
    no.aguardandoExcel = true;
    est.chamadas.push({ no, metodo, args: lista.map((a) => (a === undefined ? null : a)) });
    return undefined;
  }

  function paraArgumento(val) {
    if (val.t === "ref") return { ref: true, folha: val.folha, endereco: val.ref };
    if (val.t === "vazio") return undefined;
    return val.v;
  }

  function pedidoDeRef(folha, ref) {
    return { chave: `${folha}!${ref}`, folha, ref };
  }

  function pedidoDeNome(nome, folha) {
    return { chave: `nome:${folha || ""}!${nome.toUpperCase()}`, nome, folha: folha || null };
  }

  function pedidoDe(ler) {
    return ler.nome !== undefined ? pedidoDeNome(ler.nome, ler.folha) : pedidoDeRef(ler.folha, ler.ref);
  }

  // Devolve o valor já lido ou registra a leitura para o fim da rodada.
  function pedir(pedido, est) {
    if (est.cache.has(pedido.chave)) return est.cache.get(pedido.chave);
    est.leituras.set(pedido.chave, pedido);
    return undefined;
  }

  // ---------- Exibição ----------

  function formatarNumero(x, idioma) {
    if (!Number.isFinite(x)) return String(x);
    if (Number.isInteger(x)) return x.toLocaleString(idioma);
    const abs = Math.abs(x);
    const casas = abs >= 1 ? 2 : Math.min(8, 2 - Math.floor(Math.log10(abs)) + 1);
    return x.toLocaleString(idioma, { minimumFractionDigits: 2, maximumFractionDigits: casas });
  }

  // Número como o Excel escreve ao concatenar: sem separador de milhar.
  function formatoGeral(x, idioma) {
    const s = String(Number(x.toPrecision(15)));
    return ehPt(idioma) ? s.replace(".", ",") : s;
  }

  function nomeLogico(v, idioma) {
    if (ehPt(idioma)) return v ? "VERDADEIRO" : "FALSO";
    return v ? "TRUE" : "FALSE";
  }

  function nomeErro(v, idioma) {
    return ehPt(idioma) ? ERROS_PT[v] || v : v;
  }

  function exibirValor(val, idioma) {
    if (!val) return "—";
    switch (val.t) {
      case "num": return formatarNumero(val.v, idioma);
      case "txt": return val.v;
      case "bool": return nomeLogico(val.v, idioma);
      case "erro": return nomeErro(val.v, idioma);
      case "vazio": return "<vazio>";
      case "ref": return valorDeRef(val, idioma);
      default: return "—";
    }
  }

  // Célula única: o texto como aparece na planilha. Intervalo: a soma dos números, como no Arixcel.
  function valorDeRef(val, idioma) {
    if (!val.celulas) return "—";
    if (val.celulas.length === 1 && val.celulas[0].length === 1) {
      const c = val.celulas[0][0];
      if (c.t === "vazio") return "<vazio>";
      const visivel = val.texto !== undefined && val.texto !== "" && !/^#+$/.test(val.texto);
      return visivel ? val.texto : exibirValor(c, idioma);
    }
    const numeros = val.celulas.flat().filter((c) => c.t === "num");
    return numeros.length ? formatarNumero(numeros.reduce((s, c) => s + c.v, 0), idioma) : "";
  }

  function spanDe(no, ctx) {
    if (!ctx.alinhado || no.t0 == null) return { ini: null, fim: null };
    return { ini: ctx.tokensLocal[no.t0].ini, fim: ctx.tokensLocal[no.t1].fim };
  }

  function textoDe(no, ctx) {
    if (no.t0 == null) return "";
    const tokens = ctx.alinhado ? ctx.tokensLocal : ctx.tokens;
    const texto = ctx.alinhado ? ctx.textoLocal : ctx.texto;
    return texto.slice(tokens[no.t0].ini, tokens[no.t1].fim);
  }

  // Coluna "Local" e o que a janela seleciona no Excel ao clicar na linha: as células destacadas
  // pela função (SOMASES, CORRESP...) ou o retângulo com todas as células usadas por esta parte.
  function localizacao(no, ctx) {
    if (no.alvos && no.alvos.length) return localDeCelulas(no.alvos, ctx);
    const refs = [];
    const coletar = (n) => {
      if (n.val && n.val.t === "ref") refs.push(n.val);
      else dependenciasDe(n).forEach(coletar);
    };
    coletar(no);
    if (!refs.length) return { local: "", alvo: null };
    const folhas = new Set(refs.map((r) => r.folha));
    if (folhas.size > 1) return { local: "várias abas", alvo: null };
    const [folha] = folhas;
    const ref = E.formatarRef(E.caixa(refs));
    const alvo = `${E.prefixoFolha(folha)}!${ref}`;
    return { local: folha === ctx.folhaBase ? ref : alvo, alvo };
  }

  function localDeCelulas(celulas, ctx) {
    const folhas = new Set(celulas.map((c) => c.folha));
    if (folhas.size > 1) return { local: "várias abas", alvo: null };
    const [folha] = folhas;
    const rets = B.retangulos(celulas);
    const prefixo = folha === ctx.folhaBase ? "" : `${E.prefixoFolha(folha)}!`;
    const local = rets.length === 1 ? prefixo + E.formatarRef(rets[0]) : B.contagem(celulas.length);
    return { local, alvo: B.endereco(folha, rets) };
  }

  function noExibicao(no, ctx, info) {
    const exibido = {
      ...B.noVazio(),
      elemento: textoDe(no, ctx),
      info,
      valor: exibirValor(no.val, ctx.idioma),
      ...localizacao(no, ctx),
      ...spanDe(no, ctx),
    };
    switch (no.tipo) {
      case "func": {
        const tokens = ctx.alinhado ? ctx.tokensLocal : ctx.tokens;
        exibido.elemento = F.nomeFuncao(tokens[no.tokenNome].texto);
        exibido.filhos = argumentos(no, ctx);
        break;
      }
      case "operacao":
      case "intervalo":
        exibido.filhos = no.operandos.map((o) => noExibicao(o, ctx, ""));
        break;
      case "neg":
      case "pct":
        if (!["num", "txt", "logico"].includes(no.filho.tipo)) exibido.filhos = [noExibicao(no.filho, ctx, "")];
        break;
      case "nome":
        if (no.definicao) {
          const sub = { ...ctx, tokens: no.definicao.tokens, texto: no.definicao.texto, alinhado: false };
          exibido.filhos = [noExibicao(no.definicao.ast, sub, "definição")];
        }
        break;
      default:
        break;
    }
    return exibido;
  }

  function argumentos(no, ctx) {
    const filhos = no.args.map((arg, i) => {
      const filho = arg.tipo === "vazio" ? omitido(no.nome, i, ctx) : noExibicao(arg, ctx, A.nome(no.nome, i, ctx.idioma));
      filho.escolhido = i === no.escolhido;
      return filho;
    });
    for (const i of A.opcionaisAusentes(no.nome, no.args.length, ctx.idioma)) filhos.push(omitido(no.nome, i, ctx));
    return filhos;
  }

  function omitido(funcao, i, ctx) {
    return {
      ...B.noVazio(),
      elemento: "<omitido>",
      info: A.nome(funcao, i, ctx.idioma),
      valor: A.padrao(funcao, i, ctx.idioma),
      omitido: true,
    };
  }

  // Fórmula avulsa (fonte de uma validação, por exemplo) já calculada e pronta para exibir.
  async function subarvore(formula, adaptador, ctx, info) {
    try {
      const tokens = F.tokenizar(formula);
      const ast = F.analisar(tokens);
      await avaliar(ast, adaptador, ctx);
      return noExibicao(ast, { ...ctx, tokens, texto: formula, alinhado: false }, info);
    } catch (e) {
      return null;
    }
  }

  // validacao: { tipo, regra } como o Excel devolve em Range.dataValidation.
  async function noValidacao(validacao, adaptador, ctx) {
    const no = { ...B.noVazio(), elemento: `Validação: ${TIPOS_VALIDACAO[validacao.tipo] || validacao.tipo}` };
    const regra = validacao.regra || {};
    let formula = null;
    if (validacao.tipo === "List" && regra.list) formula = regra.list.source;
    if (validacao.tipo === "Custom" && regra.custom) formula = regra.custom.formula;
    if (typeof formula === "string" && formula.startsWith("=")) {
      const sub = await subarvore(formula, adaptador, ctx, validacao.tipo === "List" ? "fonte" : "fórmula");
      if (sub) {
        no.filhos = [sub];
        no.local = sub.local;
        no.alvo = sub.alvo;
      } else {
        no.valor = formula;
      }
    } else if (formula) {
      no.valor = String(formula);
    } else {
      const partes = Object.values(regra).find((r) => r && typeof r === "object") || {};
      no.valor = [partes.operator, partes.formula1, partes.formula2].filter((x) => x !== undefined && x !== null && x !== "").join(" ");
    }
    return no;
  }

  // celula: { folha, ref, texto, formula, formulaLocal, separadores, idioma, validacao? }
  // Devolve o que a janela precisa: endereço, fórmula local, árvore e um aviso opcional.
  async function montarDados(celula, adaptador) {
    const { folha, ref, idioma } = celula;
    const endereco = `${E.prefixoFolha(folha)}!${ref}`;
    const arvore = {
      ...B.noVazio(),
      elemento: ref,
      valor: celula.texto === "" || celula.texto == null ? "<vazio>" : String(celula.texto),
      local: ref,
      alvo: endereco,
    };
    const formula = typeof celula.formula === "string" ? celula.formula : "";
    const formulaLocal = celula.formulaLocal == null ? "" : String(celula.formulaLocal);
    const dados = { modo: "formula", editavel: true, titulo: endereco, celula: endereco, formula: formulaLocal, arvore, aviso: "" };
    const ctx = { folhaBase: folha, idioma };
    if (formula.startsWith("=")) {
      let tokens = null;
      let ast = null;
      try {
        tokens = F.tokenizar(formula);
        ast = F.analisar(tokens);
      } catch (e) {
        dados.aviso = `Não consegui ler esta fórmula (${e.message}).`;
      }
      if (ast) {
        const tokensLocal = F.tokenizar(formulaLocal, celula.separadores);
        const alinhado = F.alinhados(tokens, tokensLocal);
        await avaliar(ast, adaptador, ctx);
        arvore.filhos.push(noExibicao(ast, { ...ctx, tokens, tokensLocal, texto: formula, textoLocal: formulaLocal, alinhado }, ""));
        if (!alinhado) dados.aviso = "Mostrando as partes em inglês: não consegui alinhar com a fórmula em português.";
      }
    } else {
      dados.aviso = "Esta célula não tem fórmula. Digite uma abaixo para criar.";
    }
    if (celula.validacao) arvore.filhos.push(await noValidacao(celula.validacao, adaptador, ctx));
    return dados;
  }

  const api = { avaliar, montarDados, exibirValor, formatarNumero };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else raiz.Avaliador = api;
})(this);
