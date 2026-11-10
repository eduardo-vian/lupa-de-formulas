// Testes das funções de endereço. Uso: node tools/testar_enderecos.js
const assert = require("node:assert/strict");
const E = require("../docs/enderecos.js");

assert.deepEqual(E.separar("Plan1!A1:B2"), { folha: "Plan1", ref: "A1:B2" });
assert.deepEqual(E.separar("'Aba com espaço'!$B$1"), { folha: "Aba com espaço", ref: "$B$1" });
assert.deepEqual(E.separar("'O''Brien'!C3"), { folha: "O'Brien", ref: "C3" });
assert.deepEqual(E.separar("A1"), { folha: null, ref: "A1" });

assert.deepEqual(E.dividirAreas("Modelo!B2:B3, Modelo!D1"), ["Modelo!B2:B3", "Modelo!D1"]);
assert.deepEqual(E.dividirAreas("'Resumo, v2'!B1"), ["'Resumo, v2'!B1"]);
assert.deepEqual(
  E.dividirAreas("'O''Brien, x'!A1, 'O''Brien, x'!C1"),
  ["'O''Brien, x'!A1", "'O''Brien, x'!C1"],
);
assert.deepEqual(E.dividirAreas("Modelo!A1:A2, B5"), ["Modelo!A1:A2", "Modelo!B5"]);
assert.deepEqual(E.dividirAreas("'Aba x'!A1, B5"), ["'Aba x'!A1", "'Aba x'!B5"]);
assert.deepEqual(E.dividirAreas(""), []);

assert.equal(E.semCifrao("Modelo!$B$4"), "Modelo!B4");
assert.equal(E.semCifrao("'Aba $'!$A$1"), "'Aba $'!A1");
assert.equal(E.semCifrao("$C$3"), "C3");

assert.equal(E.numeroParaColuna(28), "AB");
assert.equal(E.numeroParaColuna(16384), "XFD");
assert.equal(E.colunaParaNumero("xfd"), 16384);
assert.deepEqual(E.analisarRef("$B$2:A1"), { r1: 1, c1: 1, r2: 2, c2: 2 });
assert.deepEqual(E.analisarRef("C:A"), { r1: 1, c1: 1, r2: E.MAX_LINHAS, c2: 3 });
assert.equal(E.analisarRef("Receita"), null);
for (const ref of ["A1", "A1:B2", "A:C", "3:5", "XFD1048576"]) assert.equal(E.formatarRef(E.analisarRef(ref)), ref);
assert.equal(E.formatarRef(E.caixa([E.analisarRef("O111"), E.analisarRef("O115"), E.analisarRef("N112")])), "N111:O115");
assert.equal(E.prefixoFolha("Premissas"), "Premissas");
assert.equal(E.prefixoFolha("Ação_2"), "Ação_2");
assert.equal(E.prefixoFolha("Aba com espaço"), "'Aba com espaço'");
assert.equal(E.prefixoFolha("O'Brien"), "'O''Brien'");
assert.equal(E.prefixoFolha("AB12"), "'AB12'");

console.log("enderecos.js: todos os testes passaram");
