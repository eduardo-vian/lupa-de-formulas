/* global Excel, Enderecos, Fluxo */
"use strict";

// Fluxo de cálculo: lê do Excel as fórmulas do âmbito e os nomes definidos, e entrega ao Fluxo.
const LupaFluxo = (() => {
  const LIMITE_CELULAS = 2000000; // células usadas somando as abas do âmbito
  const CELULAS_POR_LEITURA = 50000; // tamanho de cada pedaço lido, para respostas razoáveis

  // ambito: {} | { folha } | { folha, ref }; foco: null | { folha, ref }
  async function analisar(ambito, foco, rotulo) {
    const dados = await ler(ambito);
    const resultado = Fluxo.analisar({ ...dados, ambito, foco });
    const { visoes, grupos } = Fluxo.montarVisoes(resultado, rotulo);
    const aviso = resultado.totais.formulas ? "" : "Nenhuma fórmula no âmbito escolhido.";
    return { visoes, grupos, aviso, totais: resultado.totais };
  }

  async function ler(ambito) {
    return Excel.run(async (context) => {
      const wb = context.workbook;
      let nomesAbas = ambito.folha ? [ambito.folha] : null;
      if (!nomesAbas) {
        wb.worksheets.load("items/name");
        await context.sync();
        nomesAbas = wb.worksheets.items.map((a) => a.name);
      }
      const usados = nomesAbas.map((nome) => {
        const u = wb.worksheets.getItem(nome).getUsedRangeOrNullObject(true);
        u.load("address,rowCount,columnCount");
        return { nome, u };
      });
      await context.sync();

      const folhas = [];
      let total = 0;
      for (const { nome, u } of usados) {
        if (u.isNullObject) continue;
        total += u.rowCount * u.columnCount;
        if (total > LIMITE_CELULAS) {
          throw new Error(`o âmbito passa de ${LIMITE_CELULAS.toLocaleString("pt-BR")} células usadas; escolha uma aba ou um bloco`);
        }
        const ret = Enderecos.analisarRef(Enderecos.separar(u.address).ref);
        folhas.push({ nome, r0: ret.r1, c0: ret.c1, linhas: u.rowCount, colunas: u.columnCount, formulas: [] });
      }

      for (const f of folhas) {
        const aba = wb.worksheets.getItem(f.nome);
        const passo = Math.max(1, Math.floor(CELULAS_POR_LEITURA / f.colunas));
        for (let i = 0; i < f.linhas; i += passo) {
          const faixa = aba.getRangeByIndexes(f.r0 - 1 + i, f.c0 - 1, Math.min(passo, f.linhas - i), f.colunas);
          faixa.load("formulas");
          await context.sync();
          for (const linha of faixa.formulas) f.formulas.push(linha);
        }
      }
      return { folhas, nomes: await lerNomes(context, nomesAbas) };
    });
  }

  // Nomes definidos que apontam para intervalos: { "NOME" ou "ABA!NOME": { folha, ref } }
  async function lerNomes(context, nomesAbas) {
    const wb = context.workbook;
    const colecoes = [{ escopo: null, colecao: wb.names }]
      .concat(nomesAbas.map((aba) => ({ escopo: aba, colecao: wb.worksheets.getItem(aba).names })));
    colecoes.forEach(({ colecao }) => colecao.load("items/name,items/type"));
    await context.sync();
    const intervalos = [];
    for (const { escopo, colecao } of colecoes) {
      for (const nome of colecao.items) {
        if (nome.type !== "Range") continue;
        const r = nome.getRangeOrNullObject();
        r.load("address");
        intervalos.push({ escopo, nome: nome.name, r });
      }
    }
    await context.sync();
    const mapa = {};
    for (const { escopo, nome, r } of intervalos) {
      if (r.isNullObject) continue;
      const chave = escopo ? `${escopo.toUpperCase()}!${nome.toUpperCase()}` : nome.toUpperCase();
      mapa[chave] = Enderecos.separar(r.address);
    }
    return mapa;
  }

  return { analisar };
})();
