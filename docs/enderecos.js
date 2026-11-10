/* Funções puras de endereço do Excel. Usadas pelo painel, pelo avaliador e pelos testes em node. */
(function (raiz) {
  "use strict";

  const MAX_LINHAS = 1048576;
  const MAX_COLUNAS = 16384;

  // "'Minha aba'!A1:B2" -> { folha: "Minha aba", ref: "A1:B2" }
  function separar(endereco) {
    const i = endereco.lastIndexOf("!");
    if (i < 0) return { folha: null, ref: endereco };
    let folha = endereco.slice(0, i);
    if (folha.length >= 2 && folha.startsWith("'") && folha.endsWith("'")) {
      folha = folha.slice(1, -1).replace(/''/g, "'");
    }
    return { folha, ref: endereco.slice(i + 1) };
  }

  // "Plan1!A1:B4, 'Aba, 2'!D1" -> ["Plan1!A1:B4", "'Aba, 2'!D1"]
  // Vírgula entre aspas faz parte do nome da aba. Bloco sem aba herda a aba do bloco anterior.
  function dividirAreas(texto) {
    const partes = [];
    let atual = "";
    let entreAspas = false;
    for (let i = 0; i < texto.length; i++) {
      const c = texto[i];
      if (c === "'") {
        if (entreAspas && texto[i + 1] === "'") {
          atual += "''";
          i++;
          continue;
        }
        entreAspas = !entreAspas;
      }
      if (c === "," && !entreAspas) {
        if (atual.trim()) partes.push(atual.trim());
        atual = "";
        continue;
      }
      atual += c;
    }
    if (atual.trim()) partes.push(atual.trim());

    let prefixo = null;
    return partes.map((parte) => {
      const i = parte.lastIndexOf("!");
      if (i >= 0) {
        prefixo = parte.slice(0, i);
        return parte;
      }
      return prefixo ? `${prefixo}!${parte}` : parte;
    });
  }

  // Tira os "$" só da referência, nunca do nome da aba.
  function semCifrao(endereco) {
    const i = endereco.lastIndexOf("!");
    return endereco.slice(0, i + 1) + endereco.slice(i + 1).replace(/\$/g, "");
  }

  function colunaParaNumero(letras) {
    let n = 0;
    for (const letra of letras.toUpperCase()) n = n * 26 + (letra.charCodeAt(0) - 64);
    return n;
  }

  function numeroParaColuna(n) {
    let letras = "";
    while (n > 0) {
      const resto = (n - 1) % 26;
      letras = String.fromCharCode(65 + resto) + letras;
      n = Math.floor((n - 1) / 26);
    }
    return letras;
  }

  // "A1", "$A$1:B2", "A:C", "3:5" -> { r1, c1, r2, c2 } (contando de 1). null se não for referência.
  function analisarRef(ref) {
    const partes = ref.replace(/\$/g, "").toUpperCase().split(":");
    if (partes.length > 2) return null;
    const [a, b = a] = partes;
    const celA = /^([A-Z]{1,3})(\d+)$/.exec(a);
    const celB = /^([A-Z]{1,3})(\d+)$/.exec(b);
    if (celA && celB) {
      return normalizar(Number(celA[2]), colunaParaNumero(celA[1]), Number(celB[2]), colunaParaNumero(celB[1]));
    }
    if (partes.length === 2 && /^[A-Z]{1,3}$/.test(a) && /^[A-Z]{1,3}$/.test(b)) {
      return normalizar(1, colunaParaNumero(a), MAX_LINHAS, colunaParaNumero(b));
    }
    if (partes.length === 2 && /^\d+$/.test(a) && /^\d+$/.test(b)) {
      return normalizar(Number(a), 1, Number(b), MAX_COLUNAS);
    }
    return null;
  }

  function normalizar(r1, c1, r2, c2) {
    return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
  }

  // { r1, c1, r2, c2 } -> "A1", "A1:B2", "A:C" ou "3:5"
  function formatarRef({ r1, c1, r2, c2 }) {
    if (r1 === 1 && r2 === MAX_LINHAS) return `${numeroParaColuna(c1)}:${numeroParaColuna(c2)}`;
    if (c1 === 1 && c2 === MAX_COLUNAS) return `${r1}:${r2}`;
    const inicio = `${numeroParaColuna(c1)}${r1}`;
    return r1 === r2 && c1 === c2 ? inicio : `${inicio}:${numeroParaColuna(c2)}${r2}`;
  }

  // Menor retângulo que contém todos os retângulos da lista.
  function caixa(retangulos) {
    return retangulos.reduce((acc, r) => ({
      r1: Math.min(acc.r1, r.r1),
      c1: Math.min(acc.c1, r.c1),
      r2: Math.max(acc.r2, r.r2),
      c2: Math.max(acc.c2, r.c2),
    }));
  }

  // Nome de aba pronto para ir antes do "!": com aspas quando o Excel exigiria.
  function prefixoFolha(folha) {
    const simples = /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(folha) && !/^[A-Za-z]{1,3}\d+$/.test(folha);
    return simples ? folha : `'${folha.replace(/'/g, "''")}'`;
  }

  const api = {
    MAX_LINHAS, MAX_COLUNAS,
    separar, dividirAreas, semCifrao,
    colunaParaNumero, numeroParaColuna, analisarRef, formatarRef, caixa, prefixoFolha,
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else raiz.Enderecos = api;
})(this);
