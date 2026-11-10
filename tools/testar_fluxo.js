// Testes dos blocos, da árvore do rastreio e do fluxo de cálculo. Uso: node tools/testar_fluxo.js
const assert = require("node:assert/strict");
const E = require("../docs/enderecos.js");
const B = require("../docs/blocos.js");
const Fluxo = require("../docs/fluxo.js");
const { FOLHAS_FLUXO } = require("./casos.js");

const enderecos = (rets) => rets.map(E.formatarRef);

// ---------- Blocos ----------
assert.deepEqual(
  enderecos(B.retangulos([{ r: 1, c: 1 }, { r: 2, c: 1 }, { r: 1, c: 2 }, { r: 2, c: 2 }, { r: 5, c: 1 }])),
  ["A1:B2", "A5"],
);
assert.deepEqual(enderecos(B.retangulos([{ r: 1, c: 1 }, { r: 3, c: 1 }, { r: 3, c: 1 }])), ["A1", "A3"]);
assert.deepEqual(
  B.agruparPorChave([
    { r: 2, c: 2, chave: "=R[-1]C*2" }, { r: 3, c: 2, chave: "=R[-1]C*2" }, { r: 5, c: 2, chave: "=1" },
  ]).map((g) => [g.chave, enderecos(g.rets), g.total]),
  [["=R[-1]C*2", ["B2:B3"], 2], ["=1", ["B5"], 1]],
);
assert.equal(B.endereco("Aba com espaço", [E.analisarRef("A1"), E.analisarRef("C2:C3")]), "'Aba com espaço'!A1,'Aba com espaço'!C2:C3");

// ---------- Árvore do rastreio ----------
const rastreio = B.arvoreDeRastreio({
  direcao: "dependentes",
  origens: ["Modelo!B3"],
  areas: [
    { folha: "Modelo", ret: E.analisarRef("B4"), celulas: [{ r: 4, c: 2, chave: "=R[-1]C*Premissas!R3C3", formula: "=B3*Premissas!C3", texto: "302,50" }] },
    {
      folha: "Resumo, v2",
      ret: E.analisarRef("B1:B2"),
      celulas: [
        { r: 1, c: 2, chave: "=Modelo!R3C2+1", formula: "=Modelo!B3+1", texto: "1.211,00" },
        { r: 2, c: 2, chave: "=Modelo!R3C2+1", formula: "=Modelo!B3+1", texto: "1.211,00" },
      ],
    },
    { folha: "Grande", ret: E.analisarRef("A:A"), celulas: null },
  ],
});
assert.equal(rastreio.arvore.elemento, "Dependentes diretos de Modelo!B3");
assert.equal(rastreio.arvore.info, "3 blocos");
assert.deepEqual(rastreio.arvore.filhos.map((a) => [a.elemento, a.info]), [["Modelo", "1 bloco"], ["Resumo, v2", "1 bloco"], ["Grande", "1 bloco"]]);
const blocoResumo = rastreio.arvore.filhos[1].filhos[0];
assert.deepEqual(
  [blocoResumo.elemento, blocoResumo.info, blocoResumo.valor, blocoResumo.local, blocoResumo.formula],
  ["B1:B2", "2 células", "", "'Resumo, v2'!B1:B2", "=Modelo!B3+1"],
);
assert.deepEqual([rastreio.arvore.filhos[0].filhos[0].valor, rastreio.arvore.filhos[2].filhos[0].info], ["302,50", "intervalo grande"]);
assert.deepEqual(rastreio.cores.map((c) => [c.folha, c.cor]), [
  ["Modelo", B.CORES_RASTREIO.origem], ["Modelo", B.CORES_RASTREIO.dependentes],
  ["Resumo, v2", B.CORES_RASTREIO.dependentes], ["Grande", B.CORES_RASTREIO.dependentes],
]);
assert.match(B.arvoreDeRastreio({ direcao: "precedentes", origens: ["Modelo!A1"], areas: [] }).aviso, /Nenhum precedente/);

// ---------- Fluxo de cálculo ----------
// Modelo: A1 10 | B1 =Premissas!C2 / A2 =A1*2 / A3 =A2+1 / A4 =5+5 / A5 texto / A6 =A3*2
// Premissas: C2 0,1 | C3 7
const lista = (porCategoria) => Object.fromEntries(Object.entries(porCategoria).map(([cat, grupos]) =>
  [cat, grupos.flatMap((g) => g.rets.map((r) => `${g.folha}!${E.formatarRef(r)}`))]));

const arquivo = Fluxo.analisar({ folhas: FOLHAS_FLUXO, ambito: {} });
assert.deepEqual(lista(arquivo.entradasSaidas), {
  entrada: ["Modelo!A1", "Premissas!C2"],
  calculo: ["Modelo!A2:A3"],
  saida: ["Modelo!B1", "Modelo!A6"],
  orfa: ["Modelo!A4"],
});
assert.equal(arquivo.fluxos, null);
assert.deepEqual(arquivo.totais, { celulas: 9, formulas: 5 });

const aba = Fluxo.analisar({ folhas: FOLHAS_FLUXO, ambito: { folha: "Modelo" } });
assert.deepEqual(lista(aba.entradasSaidas), {
  entrada: ["Modelo!A1"],
  calculo: ["Modelo!A2:A3"],
  saida: ["Modelo!A6"],
  orfa: ["Modelo!B1", "Modelo!A4"],
});

const comFoco = Fluxo.analisar({ folhas: FOLHAS_FLUXO, ambito: {}, foco: { folha: "Modelo", ref: "A2:A3" } });
assert.deepEqual(lista(comFoco.fluxos), {
  precedenteFora: ["Modelo!A1"],
  entraFoco: ["Modelo!A2"],
  saiFoco: ["Modelo!A3"],
  dependenteFora: ["Modelo!A6"],
});
const focoIgual = Fluxo.analisar({ folhas: FOLHAS_FLUXO, ambito: { folha: "Modelo", ref: "A2:A3" }, foco: { folha: "Modelo", ref: "A2:A3" } });
assert.equal(focoIgual.fluxos, null);

// Nome definido conta como referência.
const comNome = Fluxo.analisar({
  folhas: [{ nome: "Plan", r0: 1, c0: 1, linhas: 2, colunas: 1, formulas: [[5], ["=Base*2"]] }],
  nomes: { BASE: { folha: "Plan", ref: "A1" } },
  ambito: {},
});
assert.deepEqual(lista(comNome.entradasSaidas), { entrada: ["Plan!A1"], calculo: [], saida: ["Plan!A2"], orfa: [] });

const { visoes, grupos } = Fluxo.montarVisoes(comFoco, "Arquivo inteiro");
assert.deepEqual(visoes.fluxos.filhos.map((c) => [c.elemento, c.info]), [
  ["Fora: precedentes", "1 célula"], ["Dentro: entradas", "1 célula"], ["Dentro: saídas", "1 célula"], ["Fora: dependentes", "1 célula"],
]);
assert.deepEqual(visoes.entradasSaidas.filhos[1].filhos[0].filhos.map((b) => b.local), ["Modelo!A2:A3"]);
assert.deepEqual(grupos.fluxos.map((g) => g.categoria), ["precedenteFora", "entraFoco", "saiFoco", "dependenteFora"]);

console.log("blocos.js e fluxo.js: todos os testes passaram");
