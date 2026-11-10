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

console.log("enderecos.js: todos os testes passaram");
