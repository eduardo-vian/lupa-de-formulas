/* global Excel, Enderecos */
"use strict";

// Cores na planilha. São regras de formatação condicional marcadas com o nome da Lupa: não mexem
// na formatação da pessoa e saem inteiras ao limpar. Temporárias somem ao fechar a janela;
// "mantidas" (Alt+X) ficam até "Remover cores"; "Fixar cores" vira preenchimento de verdade.
const LupaCores = (() => {
  const MARCA = "lupa-de-formulas";
  const TEMPORARIA = `=N("${MARCA}")=0`; // sempre VERDADEIRO; o texto identifica a regra
  const MANTIDA = `=N("${MARCA}:mantida")=0`;
  const AREAS_POR_REGRA = 100;
  const folhasPintadas = new Set();

  // grupos: [{ folha, rets: [{ r1, c1, r2, c2 }], cor }]
  async function pintar(grupos) {
    const comCelulas = grupos.filter((g) => g.rets.length);
    if (!comCelulas.length) return;
    await Excel.run(async (context) => {
      for (const { folha, rets, cor } of comCelulas) {
        const aba = context.workbook.worksheets.getItem(folha);
        for (let i = 0; i < rets.length; i += AREAS_POR_REGRA) {
          const endereco = rets.slice(i, i + AREAS_POR_REGRA).map(Enderecos.formatarRef).join(",");
          const regra = aba.getRanges(endereco).conditionalFormats.add(Excel.ConditionalFormatType.custom);
          regra.custom.rule.formula = TEMPORARIA;
          regra.custom.format.fill.color = cor;
        }
        folhasPintadas.add(folha);
      }
      await context.sync();
    });
  }

  // Regras da Lupa nas abas pedidas (todas, ou só as pintadas nesta sessão).
  async function regrasDaLupa(context, todas) {
    const abas = context.workbook.worksheets;
    abas.load("items/name");
    await context.sync();
    const escolhidas = abas.items.filter((a) => todas || folhasPintadas.has(a.name));
    const colecoes = escolhidas.map((a) => {
      const c = a.getRange().conditionalFormats;
      c.load("items/type");
      return c;
    });
    await context.sync();
    const personalizadas = [];
    for (const c of colecoes) {
      for (const regra of c.items) {
        if (regra.type !== "Custom") continue;
        regra.custom.rule.load("formula");
        personalizadas.push(regra);
      }
    }
    await context.sync();
    return personalizadas.filter((r) => (r.custom.rule.formula || "").includes(MARCA));
  }

  const ehMantida = (regra) => regra.custom.rule.formula.includes(`${MARCA}:mantida`);

  // Sem "todas": tira só as temporárias das abas pintadas nesta sessão.
  async function limpar({ todas = false } = {}) {
    await Excel.run(async (context) => {
      const regras = await regrasDaLupa(context, todas);
      for (const regra of regras) if (todas || !ehMantida(regra)) regra.delete();
      await context.sync();
    });
    if (todas) folhasPintadas.clear();
  }

  async function manter() {
    await Excel.run(async (context) => {
      const regras = await regrasDaLupa(context, false);
      for (const regra of regras) if (!ehMantida(regra)) regra.custom.rule.formula = MANTIDA;
      await context.sync();
    });
  }

  // Troca as regras da Lupa por preenchimento normal das células (passa a fazer parte do arquivo).
  async function fixar() {
    await Excel.run(async (context) => {
      const regras = await regrasDaLupa(context, true);
      for (const regra of regras) regra.custom.format.fill.load("color");
      await context.sync();
      for (const regra of regras) {
        regra.getRanges().format.fill.color = regra.custom.format.fill.color;
        regra.delete();
      }
      await context.sync();
    });
    folhasPintadas.clear();
  }

  return { pintar, limpar, manter, fixar };
})();
