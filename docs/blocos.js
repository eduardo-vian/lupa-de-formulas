/* Agrupa células em blocos retangulares e monta a árvore do rastreio.
   Puro: usado pelo avaliador, pelo rastreio, pelo fluxo de cálculo e pelos testes em node. */
(function (raiz) {
  "use strict";

  const carregar = (global, caminho) =>
    typeof module !== "undefined" && module.exports ? require(caminho) : raiz[global];
  const E = carregar("Enderecos", "./enderecos.js");

  const MAX_AREAS_ALVO = 300; // áreas num endereço de seleção; acima disso o resto fica de fora
  const CORES_RASTREIO = { origem: "#F4B6C2", dependentes: "#C6E0B4", precedentes: "#BDD7EE" };

  // Linha da árvore da janela com todos os campos, para os outros só preencherem o que muda.
  function noVazio() {
    return {
      elemento: "", info: "", valor: "", local: "", alvo: null,
      ini: null, fim: null, omitido: false, escolhido: false, filhos: [],
    };
  }

  const contagem = (n) => `${n.toLocaleString("pt-BR")} ${n === 1 ? "célula" : "células"}`;

  // Células { r, c } -> retângulos que as cobrem exatamente: corridas de linhas seguidas em cada
  // coluna, juntadas quando as colunas vizinhas têm a mesma corrida.
  function retangulos(celulas) {
    const porColuna = new Map();
    for (const { r, c } of celulas) {
      if (!porColuna.has(c)) porColuna.set(c, new Set());
      porColuna.get(c).add(r);
    }
    const corridas = [];
    for (const [c, conjunto] of porColuna) {
      const linhas = [...conjunto].sort((a, b) => a - b);
      let inicio = linhas[0];
      let anterior = linhas[0];
      for (let k = 1; k <= linhas.length; k++) {
        if (linhas[k] === anterior + 1) {
          anterior = linhas[k];
          continue;
        }
        corridas.push({ c, r1: inicio, r2: anterior });
        inicio = anterior = linhas[k];
      }
    }
    corridas.sort((a, b) => a.r1 - b.r1 || a.r2 - b.r2 || a.c - b.c);
    const rets = [];
    for (const k of corridas) {
      const ultimo = rets[rets.length - 1];
      if (ultimo && ultimo.r1 === k.r1 && ultimo.r2 === k.r2 && ultimo.c2 === k.c - 1) ultimo.c2 = k.c;
      else rets.push({ r1: k.r1, c1: k.c, r2: k.r2, c2: k.c });
    }
    return rets.sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1);
  }

  // Células { r, c, chave } -> um grupo por chave (por exemplo, a fórmula em R1C1), com seus retângulos.
  function agruparPorChave(celulas) {
    const grupos = new Map();
    for (const cel of celulas) {
      if (!grupos.has(cel.chave)) grupos.set(cel.chave, []);
      grupos.get(cel.chave).push(cel);
    }
    return [...grupos].map(([chave, lista]) => ({ chave, rets: retangulos(lista), total: lista.length }));
  }

  // "Aba!A1:B2,Aba!C3": endereço de seleção para uma lista de retângulos de uma aba.
  function endereco(folha, rets) {
    const prefixo = E.prefixoFolha(folha);
    return rets.slice(0, MAX_AREAS_ALVO).map((r) => `${prefixo}!${E.formatarRef(r)}`).join(",");
  }

  function totalCelulas(rets) {
    return rets.reduce((s, r) => s + (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1), 0);
  }

  function resumir(texto, max = 60) {
    return texto.length > max ? `${texto.slice(0, max - 1)}…` : texto;
  }

  // ---------- Árvore do rastreio ----------
  // areas: [{ folha, ret, celulas }]; celulas: [{ r, c, chave, formula, texto }], ou null quando
  // a área é grande demais para ler célula por célula (vira um bloco só).

  function arvoreDeRastreio({ direcao, origens, areas }) {
    const porFolha = new Map();
    const vistas = new Set();
    for (const area of areas) {
      if (!porFolha.has(area.folha)) porFolha.set(area.folha, { celulas: [], grandes: [] });
      const grupo = porFolha.get(area.folha);
      if (!area.celulas) {
        grupo.grandes.push(area.ret);
        continue;
      }
      for (const cel of area.celulas) {
        const chave = `${area.folha}!${cel.r},${cel.c}`;
        if (vistas.has(chave)) continue;
        vistas.add(chave);
        grupo.celulas.push(cel);
      }
    }

    const abas = [];
    const cores = coresDaOrigem(origens);
    let totalBlocos = 0;
    for (const [folha, { celulas, grandes }] of porFolha) {
      const primeiras = new Map(celulas.map((c) => [`${c.r},${c.c}`, c]));
      const blocos = [];
      for (const grupo of agruparPorChave(celulas)) {
        for (const ret of grupo.rets) blocos.push({ ret, no: noDeBloco(folha, ret, primeiras.get(`${ret.r1},${ret.c1}`)) });
      }
      for (const ret of grandes) blocos.push({ ret, no: noDeBloco(folha, ret, null) });
      if (!blocos.length) continue;
      blocos.sort((a, b) => a.ret.r1 - b.ret.r1 || a.ret.c1 - b.ret.c1);
      totalBlocos += blocos.length;
      const rets = blocos.map((b) => b.ret);
      abas.push({
        ...noVazio(),
        elemento: folha,
        info: `${blocos.length} ${blocos.length === 1 ? "bloco" : "blocos"}`,
        alvo: endereco(folha, rets),
        filhos: blocos.map((b) => b.no),
      });
      cores.push({ folha, rets, cor: CORES_RASTREIO[direcao] });
    }

    const rotulo = direcao === "dependentes" ? "Dependentes diretos" : "Precedentes diretos";
    const origemTexto = resumir(origens.map(E.semCifrao).join(", "));
    const arvore = {
      ...noVazio(),
      elemento: `${rotulo} de ${origemTexto}`,
      info: `${totalBlocos} ${totalBlocos === 1 ? "bloco" : "blocos"}`,
      local: origemTexto,
      alvo: origens.join(","),
      filhos: abas,
    };
    const aviso = totalBlocos
      ? ""
      : `Nenhum ${direcao === "dependentes" ? "dependente" : "precedente"} direto encontrado. `
        + "Referências feitas por DESLOC, INDIRETO e macros não são detectadas.";
    return { modo: "rastreio", titulo: rotulo, editavel: false, formula: "", arvore, aviso, cores };
  }

  function noDeBloco(folha, ret, primeira) {
    const n = (ret.r2 - ret.r1 + 1) * (ret.c2 - ret.c1 + 1);
    const local = `${E.prefixoFolha(folha)}!${E.formatarRef(ret)}`;
    return {
      ...noVazio(),
      elemento: E.formatarRef(ret),
      info: primeira ? contagem(n) : "intervalo grande",
      valor: primeira && n === 1 ? primeira.texto || "<vazio>" : "",
      local,
      alvo: local,
      formula: primeira && primeira.formula ? primeira.formula : "",
    };
  }

  function coresDaOrigem(origens) {
    const porFolha = new Map();
    for (const origem of origens) {
      const { folha, ref } = E.separar(origem);
      const ret = E.analisarRef(ref);
      if (!folha || !ret) continue;
      if (!porFolha.has(folha)) porFolha.set(folha, []);
      porFolha.get(folha).push(ret);
    }
    return [...porFolha].map(([folha, rets]) => ({ folha, rets, cor: CORES_RASTREIO.origem }));
  }

  const api = {
    MAX_AREAS_ALVO, CORES_RASTREIO,
    noVazio, contagem, retangulos, agruparPorChave, endereco, totalCelulas, resumir, arvoreDeRastreio,
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else raiz.Blocos = api;
})(this);
