// Adaptador falso para testar o avaliador sem o Excel: células num dicionário e algumas funções.
const E = require("../docs/enderecos.js");

const LIMITE_VALORES = 20000;

// celulas: { "Modelo!A1": 10, ... }   nomes: { RECEITA: "Premissas!C2" ou um número }
function criarAdaptador(celulas = {}, nomes = {}) {
  function celula(folha, linha, coluna) {
    const v = celulas[`${folha}!${E.numeroParaColuna(coluna)}${linha}`];
    if (v === undefined || v === null) return { t: "vazio" };
    if (typeof v === "number") return { t: "num", v };
    if (typeof v === "boolean") return { t: "bool", v };
    if (String(v).startsWith("#")) return { t: "erro", v };
    return { t: "txt", v };
  }

  // Imita Range.text com o formato "#.##0,00".
  function textoExibido(c) {
    switch (c.t) {
      case "num": return c.v.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      case "vazio": return "";
      case "bool": return c.v ? "VERDADEIRO" : "FALSO";
      default: return String(c.v);
    }
  }

  function lerRef(folha, ref) {
    const ret = E.analisarRef(ref);
    const val = { t: "ref", folha, ref: E.formatarRef(ret), ...ret, celulas: null };
    if ((ret.r2 - ret.r1 + 1) * (ret.c2 - ret.c1 + 1) > LIMITE_VALORES) return val;
    val.celulas = [];
    for (let r = ret.r1; r <= ret.r2; r++) {
      const linha = [];
      for (let c = ret.c1; c <= ret.c2; c++) linha.push(celula(folha, r, c));
      val.celulas.push(linha);
    }
    if (ret.r1 === ret.r2 && ret.c1 === ret.c2) val.texto = textoExibido(val.celulas[0][0]);
    return val;
  }

  async function ler(pedidos) {
    const resultado = new Map();
    for (const p of pedidos) {
      if (p.nome === undefined) {
        resultado.set(p.chave, lerRef(p.folha, p.ref));
        continue;
      }
      const alvo = nomes[p.nome.toUpperCase()];
      if (alvo === undefined) resultado.set(p.chave, { t: "erro", v: "#NAME?" });
      else if (typeof alvo === "number") resultado.set(p.chave, { t: "num", v: alvo });
      else {
        const { folha, ref } = E.separar(alvo);
        resultado.set(p.chave, lerRef(folha, ref));
      }
    }
    return resultado;
  }

  function numeros(args) {
    const lista = [];
    for (const a of args) {
      if (a && a.t === "ref") a.celulas.flat().forEach((c) => c.t === "num" && lista.push(c.v));
      else if (typeof a === "number") lista.push(a);
    }
    return lista;
  }

  const FUNCOES = {
    sum: (...args) => ({ t: "num", v: numeros(args).reduce((s, n) => s + n, 0) }),
    max: (...args) => ({ t: "num", v: Math.max(...numeros(args)) }),
    round: (n, casas) => ({ t: "num", v: Number(n.toFixed(casas)) }),
  };

  async function chamar(chamadas) {
    return chamadas.map(({ metodo, args }) =>
      FUNCOES[metodo](...args.map((a) => (a && a.ref ? lerRef(a.folha, a.endereco) : a))));
  }

  const metodo = (nome) => (FUNCOES[nome.toLowerCase()] ? nome.toLowerCase() : null);
  return { ler, chamar, metodo };
}

module.exports = { criarAdaptador };
