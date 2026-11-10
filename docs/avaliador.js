/* Calcula cada parte da fórmula e monta a árvore exibida na janela.
   Não fala com o Excel diretamente: recebe um "adaptador" com ler(), chamar() e metodo(). */
(function (raiz) {
  "use strict";

  const carregar = (global, caminho) =>
    typeof module !== "undefined" && module.exports ? require(caminho) : raiz[global];
  const E = carregar("Enderecos", "./enderecos.js");
  const F = carregar("Formula", "./formula.js");
  const A = carregar("Argumentos", "./argumentos.js");

  const MAX_RODADAS = 60;
  const ERROS_PT = {
    "#N/A": "#N/D", "#VALUE!": "#VALOR!", "#NAME?": "#NOME?",
    "#NUM!": "#NÚM!", "#NULL!": "#NULO!", "#SPILL!": "#DESPEJAR!",
  };
  const COMPARACOES = ["=", "<>", "<", ">", "<=", ">="];

  // ---------- Valores ----------
  // num, txt, bool, erro, vazio, ref (intervalo lido do Excel) e nd (não deu para calcular).

  const num = (v) => ({ t: "num", v: Object.is(v, -0) ? 0 : v });
  const txt = (v) => ({ t: "txt", v });
  const bool = (v) => ({ t: "bool", v });
  const erro = (v) => ({ t: "erro", v });
  const nd = (motivo) => ({ t: "nd", motivo });
  const VAZIO = { t: "vazio" };

  const ehPt = (idioma) => /^pt/i.test(idioma || "");

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
    for (const v of [x, y]) if (v.t === "erro" || v.t === "nd") return v;
    if (x.t === "vazio") x = vazioComo(y);
    if (y.t === "vazio") y = vazioComo(x);
    const ordem = { num: 0, txt: 1, bool: 2 };
    let c;
    if (x.t !== y.t) c = ordem[x.t] - ordem[y.t];
    else if (x.t === "txt") c = x.v.localeCompare(y.v, idioma || undefined, { sensitivity: "base" });
    else c = Number(x.v) - Number(y.v);
    return bool({ "=": c === 0, "<>": c !== 0, "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0 }[op]);
  }

  function vazioComo(outro) {
    if (outro.t === "txt") return txt("");
    if (outro.t === "bool") return bool(false);
    return num(0);
  }

  // ---------- Funções calculadas pela própria Lupa ----------
  // As de referência (DESLOC, INDIRETO) e as que escolhem um ramo (SE, SEERRO...) ficam aqui;
  // o resto vai para o Excel. Uma função pode devolver { ler } quando precisa de outro intervalo.

  const PROPRIAS = {
    IF(args) {
      const c = logico(args[0] || VAZIO);
      if (typeof c !== "boolean") return c;
      const escolhido = c ? args[1] : args[2];
      if (escolhido === undefined) return c ? num(0) : bool(false);
      return escolhido.t === "vazio" ? num(0) : escolhido;
    },
    IFERROR: (args) => seErro(args, () => true),
    IFNA: (args) => seErro(args, (v) => v === "#N/A"),
    CHOOSE(args) {
      const i = numero(args[0] || VAZIO);
      if (typeof i !== "number") return i;
      const k = Math.trunc(i);
      if (k < 1 || k >= args.length) return erro("#VALUE!");
      return args[k].t === "vazio" ? num(0) : args[k];
    },
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
  };

  function seErro(args, condicao) {
    const v = args[0] || VAZIO;
    const e = escalar(v);
    if (e.t === "nd") return e;
    if (e.t === "erro" && condicao(e.v)) {
      const alternativa = args[1] || VAZIO;
      return alternativa.t === "vazio" ? num(0) : alternativa;
    }
    return v.t === "vazio" ? num(0) : v;
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
        if (v.t === "erro" || v.t === "nd") return v;
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
    if (base.t === "erro" || base.t === "nd") return base;
    if (base.t !== "ref") return erro("#VALUE!");
    const inteiro = (v, padrao) => {
      if (!v || v.t === "vazio") return padrao;
      const n = numero(v);
      return typeof n === "number" ? Math.trunc(n) : n;
    };
    const lins = inteiro(args[1], 0);
    const cols = inteiro(args[2], 0);
    const altura = inteiro(args[3], base.r2 - base.r1 + 1);
    const largura = inteiro(args[4], base.c2 - base.c1 + 1);
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

  // ---------- Cálculo em rodadas ----------
  // Cada rodada calcula tudo o que já é possível e junta as leituras e chamadas ao Excel que
  // faltam, para mandar de uma vez só. Repete até a fórmula inteira ter valor.

  async function avaliar(ast, adaptador, ctx) {
    const cache = new Map();
    for (let rodada = 0; rodada < MAX_RODADAS && ast.val === undefined; rodada++) {
      const est = { cache, leituras: new Map(), chamadas: [], adaptador, ctx };
      visitar(ast, est);
      if (ast.val !== undefined || (!est.leituras.size && !est.chamadas.length)) break;
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

  function filhosDe(no) {
    switch (no.tipo) {
      case "operacao":
      case "intervalo": return no.operandos;
      case "neg":
      case "pct": return [no.filho];
      case "func": return no.args;
      default: return [];
    }
  }

  function visitar(no, est) {
    if (no.val !== undefined) return;
    const filhos = filhosDe(no);
    filhos.forEach((f) => visitar(f, est));
    if (filhos.some((f) => f.val === undefined)) return;
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
        return pedir(pedidoDeNome(ref, folha), est);
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
    const problema = vals.find((v) => v.t === "erro" || v.t === "nd");
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
      return r && r.ler ? pedir(pedidoDe(r.ler), est) : r;
    }
    const metodo = est.adaptador.metodo(no.nome);
    if (!metodo) return nd("função que a Lupa ainda não calcula");
    const problema = args.find((v) => v.t === "nd" || v.t === "erro");
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

  // Coluna "Local": o intervalo que contém todas as células usadas por esta parte da fórmula.
  function localDe(no, ctx) {
    const refs = [];
    const coletar = (n) => {
      if (n.val && n.val.t === "ref") refs.push(n.val);
      else filhosDe(n).forEach(coletar);
    };
    coletar(no);
    if (!refs.length) return "";
    const folhas = new Set(refs.map((r) => r.folha));
    if (folhas.size > 1) return "várias abas";
    const [folha] = folhas;
    const ref = E.formatarRef(E.caixa(refs));
    return folha === ctx.folhaBase ? ref : `${E.prefixoFolha(folha)}!${ref}`;
  }

  function noExibicao(no, ctx, info) {
    const exibido = {
      elemento: textoDe(no, ctx),
      info,
      valor: exibirValor(no.val, ctx.idioma),
      local: localDe(no, ctx),
      ...spanDe(no, ctx),
      navegar: no.val && no.val.t === "ref" ? `${E.prefixoFolha(no.val.folha)}!${no.val.ref}` : null,
      omitido: false,
      filhos: [],
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
      default:
        break;
    }
    return exibido;
  }

  function argumentos(no, ctx) {
    const filhos = no.args.map((arg, i) => {
      const nome = A.nome(no.nome, i, ctx.idioma);
      return arg.tipo === "vazio" ? omitido(no.nome, i, ctx) : noExibicao(arg, ctx, nome);
    });
    for (const i of A.opcionaisAusentes(no.nome, no.args.length, ctx.idioma)) filhos.push(omitido(no.nome, i, ctx));
    return filhos;
  }

  function omitido(funcao, i, ctx) {
    return {
      elemento: "<omitido>",
      info: A.nome(funcao, i, ctx.idioma),
      valor: A.padrao(funcao, i, ctx.idioma),
      local: "",
      ini: null,
      fim: null,
      navegar: null,
      omitido: true,
      filhos: [],
    };
  }

  // celula: { folha, ref, texto, formula, formulaLocal, separadores, idioma }
  // Devolve o que a janela precisa: endereço, fórmula local, árvore e um aviso opcional.
  async function montarDados(celula, adaptador) {
    const { folha, ref, idioma } = celula;
    const arvore = {
      elemento: ref,
      info: "",
      valor: celula.texto === "" || celula.texto == null ? "<vazio>" : String(celula.texto),
      local: ref,
      ini: null,
      fim: null,
      navegar: null,
      omitido: false,
      filhos: [],
    };
    const formula = typeof celula.formula === "string" ? celula.formula : "";
    const formulaLocal = celula.formulaLocal == null ? "" : String(celula.formulaLocal);
    const dados = { celula: `${E.prefixoFolha(folha)}!${ref}`, formula: formulaLocal, arvore, aviso: "" };
    if (!formula.startsWith("=")) {
      dados.aviso = "Esta célula não tem fórmula. Digite uma abaixo para criar.";
      return dados;
    }
    let tokens;
    let ast;
    try {
      tokens = F.tokenizar(formula);
      ast = F.analisar(tokens);
    } catch (e) {
      dados.aviso = `Não consegui ler esta fórmula (${e.message}).`;
      return dados;
    }
    const tokensLocal = F.tokenizar(formulaLocal, celula.separadores);
    const alinhado = F.alinhados(tokens, tokensLocal);
    const ctx = { folhaBase: folha, idioma };
    await avaliar(ast, adaptador, ctx);
    arvore.filhos = [noExibicao(ast, { ...ctx, tokens, tokensLocal, texto: formula, textoLocal: formulaLocal, alinhado }, "")];
    if (!alinhado) dados.aviso = "Mostrando as partes em inglês: não consegui alinhar com a fórmula em português.";
    return dados;
  }

  const api = { avaliar, montarDados, exibirValor, formatarNumero };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else raiz.Avaliador = api;
})(this);
