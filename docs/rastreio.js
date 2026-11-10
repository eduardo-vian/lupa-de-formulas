/* global Excel, Enderecos, Blocos */
"use strict";

// Rastreio de precedentes ou dependentes diretos de várias células ao mesmo tempo.
// Lê do Excel o que foi encontrado; a árvore em blocos é montada por Blocos.arvoreDeRastreio.
const LupaRastreio = (() => {
  const LIMITE_AREA = 5000; // área maior que isso vira um bloco só, sem ler fórmula por fórmula

  async function montar(origens, direcao) {
    const metodo = direcao === "dependentes" ? "getDirectDependents" : "getDirectPrecedents";
    const vistas = new Set();
    const encontradas = [];
    for (const origem of origens) {
      let enderecos = [];
      try {
        enderecos = await Excel.run(async (context) => {
          const { folha, ref } = Enderecos.separar(origem);
          const abas = context.workbook.worksheets;
          const aba = folha ? abas.getItem(folha) : abas.getActiveWorksheet();
          const areas = aba.getRange(ref)[metodo]();
          areas.load("addresses");
          await context.sync();
          return areas.addresses;
        });
      } catch (e) {
        if (e.code !== "ItemNotFound") throw e; // ItemNotFound: esta origem não tem ligações
      }
      for (const texto of enderecos) {
        for (const area of Enderecos.dividirAreas(texto)) {
          const { folha, ref } = Enderecos.separar(area);
          const ret = Enderecos.analisarRef(ref);
          if (!ret) continue;
          const chave = `${folha}!${Enderecos.formatarRef(ret)}`.toUpperCase();
          if (vistas.has(chave)) continue;
          vistas.add(chave);
          encontradas.push({ folha, ret });
        }
      }
    }
    const areas = await lerAreas(encontradas);
    return Blocos.arvoreDeRastreio({ direcao, origens, areas });
  }

  // Fórmulas (em R1C1, para agrupar as iguais), fórmula local e texto de cada célula encontrada.
  async function lerAreas(encontradas) {
    if (!encontradas.length) return [];
    return Excel.run(async (context) => {
      const itens = encontradas.map(({ folha, ret }) => {
        let intervalo = null;
        if (Blocos.totalCelulas([ret]) <= LIMITE_AREA) {
          intervalo = context.workbook.worksheets.getItem(folha).getRange(Enderecos.formatarRef(ret));
          intervalo.load("formulasR1C1,formulasLocal,text");
        }
        return { folha, ret, intervalo };
      });
      await context.sync();
      return itens.map(({ folha, ret, intervalo }) => ({ folha, ret, celulas: intervalo ? celulasDe(ret, intervalo) : null }));
    });
  }

  function celulasDe(ret, intervalo) {
    const celulas = [];
    intervalo.formulasR1C1.forEach((linha, i) => linha.forEach((r1c1, j) => {
      const ehFormula = typeof r1c1 === "string" && r1c1.startsWith("=");
      celulas.push({
        r: ret.r1 + i,
        c: ret.c1 + j,
        chave: ehFormula ? r1c1 : "(valor)",
        formula: ehFormula ? intervalo.formulasLocal[i][j] : "",
        texto: intervalo.text[i][j],
      });
    }));
    return celulas;
  }

  return { montar };
})();
