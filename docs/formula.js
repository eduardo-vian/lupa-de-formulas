/* Leitura de fórmulas do Excel: tokens e árvore sintática. Não depende do Excel; testada em node. */
(function (raiz) {
  "use strict";

  // Range.formulas devolve a fórmula em inglês: ponto decimal e vírgula entre argumentos.
  const SEPARADORES_EN = { decimal: ".", lista: "," };

  const PALAVRA = /^[\p{L}_\\$][\p{L}\p{N}_.$]*/u;
  const CELULA = /^\$?[A-Za-z]{1,3}\$?\d{1,7}$/;
  const COLUNA = /^\$?[A-Za-z]{1,3}$/;
  const LINHAS = /^\$?\d+:\$?\d+/;
  const ERRO = /^#[\p{L}\p{N}_/]+[!?]?/u;
  const LOGICOS = new Set(["TRUE", "FALSE", "VERDADEIRO", "FALSO"]);
  const OPERADORES = ["<>", "<=", ">=", "+", "-", "*", "/", "^", "&", "=", "<", ">", "%", ":"];
  const PREFIXOS_FUNCAO = /^_(xlfn|xlws|xludf)\./i;

  // Precedência do Excel, da mais fraca para a mais forte. Abaixo delas: %, negação e ":".
  const NIVEIS = [["=", "<>", "<", ">", "<=", ">="], ["&"], ["+", "-"], ["*", "/"], ["^"]];
  const FOLHAS = new Set(["num", "txt", "logico", "erro", "ref", "nome", "estruturada", "matriz"]);

  class ErroFormula extends Error {}

  // ---------- Tokens ----------

  function tokenizar(texto, separadores = SEPARADORES_EN) {
    const dec = `\\${separadores.decimal}`;
    const numero = new RegExp(`^(\\d+(${dec}\\d*)?|${dec}\\d+)([eE][+-]?\\d+)?`);
    const tokens = [];
    const adicionar = (tipo, ini, fim) => tokens.push({ tipo, texto: texto.slice(ini, fim), ini, fim });

    let i = texto.startsWith("=") ? 1 : 0;
    while (i < texto.length) {
      const c = texto[i];
      const resto = texto.slice(i);
      let m;
      if (/\s/.test(c)) {
        i++;
        continue;
      }
      if (c === '"') {
        const fim = fimDeAspas(texto, i, '"');
        adicionar("txt", i, fim);
        i = fim;
        continue;
      }
      if (c === "{") {
        const fecha = texto.indexOf("}", i);
        const fim = fecha < 0 ? texto.length : fecha + 1;
        adicionar("matriz", i, fim);
        i = fim;
        continue;
      }
      if ((m = ERRO.exec(resto))) {
        adicionar("erro", i, i + m[0].length);
        i += m[0].length;
        continue;
      }
      if (c === "(" || c === ")" || c === separadores.lista) {
        adicionar(c === "(" ? "abre" : c === ")" ? "fecha" : "sep", i, i + 1);
        i++;
        continue;
      }
      if (c === "'" || c === "[") {
        const alvo = lerComFolha(texto, i);
        if (alvo) {
          adicionar(alvo.tipo, i, alvo.fim);
          i = alvo.fim;
          continue;
        }
        if (c === "[") {
          const fim = fimDeColchetes(texto, i);
          adicionar("estruturada", i, fim);
          i = fim;
          continue;
        }
      }
      if ((m = LINHAS.exec(resto)) || (m = numero.exec(resto))) {
        adicionar(LINHAS.test(m[0]) ? "ref" : "num", i, i + m[0].length);
        i += m[0].length;
        continue;
      }
      if ((m = PALAVRA.exec(resto))) {
        const j = i + m[0].length;
        let alvo;
        if (texto[j] === "(") alvo = { tipo: "func", fim: j };
        else if (texto[j] === "!") alvo = lerAlvo(texto, j + 1);
        else if (texto[j] === "[") alvo = { tipo: "estruturada", fim: fimDeColchetes(texto, j) };
        else alvo = lerAlvo(texto, i);
        adicionar(alvo.tipo, i, alvo.fim);
        i = alvo.fim;
        continue;
      }
      const op = OPERADORES.find((o) => texto.startsWith(o, i));
      adicionar(op ? "op" : "desconhecido", i, i + (op ? op.length : 1));
      i += op ? op.length : 1;
    }
    return tokens;
  }

  // "'Aba x'!A1" ou "[1]Plan1!A1": devolve { tipo, fim }, ou null se o prefixo não terminar em "!".
  function lerComFolha(texto, i) {
    let j;
    if (texto[i] === "'") {
      j = fimDeAspas(texto, i, "'");
    } else {
      const fecha = texto.indexOf("]", i);
      const m = fecha < 0 ? null : PALAVRA.exec(texto.slice(fecha + 1));
      if (!m) return null;
      j = fecha + 1 + m[0].length;
    }
    return texto[j] === "!" ? lerAlvo(texto, j + 1) : null;
  }

  // Lê a partir de k uma referência (A1, A1:B2, A:C, 3:5, A1#), um erro (#REF!) ou um nome.
  function lerAlvo(texto, k) {
    const resto = texto.slice(k);
    let m = LINHAS.exec(resto) || ERRO.exec(resto);
    if (m) return { tipo: m[0].startsWith("#") ? "erro" : "ref", fim: k + m[0].length };
    m = PALAVRA.exec(resto);
    if (!m) return { tipo: "desconhecido", fim: k };
    const palavra = m[0];
    let fim = k + palavra.length;
    if (CELULA.test(palavra)) {
      if (texto[fim] === ":") {
        const m2 = PALAVRA.exec(texto.slice(fim + 1));
        const depois = m2 ? texto[fim + 1 + m2[0].length] : "";
        if (m2 && CELULA.test(m2[0]) && depois !== "(" && depois !== "!") fim += 1 + m2[0].length;
      }
      if (texto[fim] === "#") fim++;
      return { tipo: "ref", fim };
    }
    if (COLUNA.test(palavra) && texto[fim] === ":") {
      const m2 = PALAVRA.exec(texto.slice(fim + 1));
      if (m2 && COLUNA.test(m2[0])) return { tipo: "ref", fim: fim + 1 + m2[0].length };
    }
    return { tipo: LOGICOS.has(palavra.toUpperCase()) ? "logico" : "nome", fim };
  }

  function fimDeAspas(texto, i, aspa) {
    let j = i + 1;
    while (j < texto.length) {
      if (texto[j] === aspa) {
        if (texto[j + 1] === aspa) {
          j += 2;
          continue;
        }
        return j + 1;
      }
      j++;
    }
    return texto.length;
  }

  function fimDeColchetes(texto, i) {
    let nivel = 0;
    for (let j = i; j < texto.length; j++) {
      if (texto[j] === "[") nivel++;
      else if (texto[j] === "]" && --nivel === 0) return j + 1;
    }
    return texto.length;
  }

  function nomeFuncao(texto) {
    return texto.replace(PREFIXOS_FUNCAO, "").toUpperCase();
  }

  // As versões em inglês e local da mesma fórmula têm os mesmos tokens, só com textos diferentes.
  function alinhados(a, b) {
    if (a.length !== b.length) return false;
    const palavra = (tipo) => tipo === "nome" || tipo === "logico";
    return a.every((t, i) => t.tipo === b[i].tipo || (palavra(t.tipo) && palavra(b[i].tipo)));
  }

  // ---------- Árvore sintática ----------
  // Cada nó guarda t0 e t1: o primeiro e o último token que ele cobre (para recortar o texto).

  function analisar(tokens) {
    let p = 0;
    const atual = () => tokens[p];
    const ehOp = (...ops) => atual() !== undefined && atual().tipo === "op" && ops.includes(atual().texto);
    const falhar = (msg) => {
      throw new ErroFormula(msg);
    };
    const esperar = (tipo, descricao) => {
      if (!atual() || atual().tipo !== tipo) {
        falhar(`esperava ${descricao}${atual() ? ` antes de "${atual().texto}"` : " no fim"}`);
      }
      return p++;
    };

    function nivel(n) {
      if (n === NIVEIS.length) return percentual();
      const primeiro = nivel(n + 1);
      if (!ehOp(...NIVEIS[n])) return primeiro;
      const no = { tipo: "operacao", operandos: [primeiro], ops: [], t0: primeiro.t0, t1: null };
      while (ehOp(...NIVEIS[n])) {
        no.ops.push(tokens[p++].texto);
        no.operandos.push(nivel(n + 1));
      }
      no.t1 = no.operandos[no.operandos.length - 1].t1;
      return no;
    }

    function percentual() {
      let no = negacao();
      while (ehOp("%")) no = { tipo: "pct", filho: no, t0: no.t0, t1: p++ };
      return no;
    }

    // No Excel a negação vem antes da potência: =-2^2 dá 4.
    function negacao() {
      if (!ehOp("-", "+")) return intervalo();
      const t0 = p++;
      const filho = negacao();
      return { tipo: "neg", sinal: tokens[t0].texto, filho, t0, t1: filho.t1 };
    }

    function intervalo() {
      const primeiro = primario();
      if (!ehOp(":")) return primeiro;
      const no = { tipo: "intervalo", operandos: [primeiro], t0: primeiro.t0, t1: null };
      while (ehOp(":")) {
        p++;
        no.operandos.push(primario());
      }
      no.t1 = no.operandos[no.operandos.length - 1].t1;
      return no;
    }

    function primario() {
      const tk = atual();
      if (!tk) falhar("a fórmula terminou antes da hora");
      if (FOLHAS.has(tk.tipo)) {
        p++;
        return { tipo: tk.tipo, texto: tk.texto, t0: p - 1, t1: p - 1 };
      }
      if (tk.tipo === "func") return funcao();
      if (tk.tipo === "abre") {
        const t0 = p++;
        const dentro = nivel(0);
        const t1 = esperar("fecha", '")"');
        return { ...dentro, t0, t1, parenteses: true };
      }
      return falhar(`símbolo inesperado "${tk.texto}"`);
    }

    function funcao() {
      const tokenNome = p++;
      const no = { tipo: "func", nome: nomeFuncao(tokens[tokenNome].texto), args: [], tokenNome, t0: tokenNome, t1: null };
      esperar("abre", '"("');
      if (atual() && atual().tipo === "fecha") {
        no.t1 = p++;
        return no;
      }
      for (;;) {
        const vazio = !atual() || atual().tipo === "sep" || atual().tipo === "fecha";
        no.args.push(vazio ? { tipo: "vazio", t0: null, t1: null } : nivel(0));
        if (atual() && atual().tipo === "sep") {
          p++;
          continue;
        }
        no.t1 = esperar("fecha", '")"');
        return no;
      }
    }

    const raizNo = nivel(0);
    if (p < tokens.length) falhar(`símbolo inesperado "${tokens[p].texto}"`);
    return raizNo;
  }

  const api = { SEPARADORES_EN, ErroFormula, tokenizar, analisar, alinhados, nomeFuncao };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else raiz.Formula = api;
})(this);
