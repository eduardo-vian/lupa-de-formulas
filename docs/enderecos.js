/* Funções puras de endereço do Excel. Usadas pelo painel e pelos testes em node. */
(function (raiz) {
  "use strict";

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

  const api = { separar, dividirAreas, semCifrao };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else raiz.Enderecos = api;
})(this);
