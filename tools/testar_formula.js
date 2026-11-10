// Testes do leitor de fórmulas e do avaliador. Uso: node tools/testar_formula.js
const assert = require("node:assert/strict");
const F = require("../docs/formula.js");
const Av = require("../docs/avaliador.js");
const { criarAdaptador } = require("./adaptador_falso.js");
const { PT, CELULAS, DESLOC, SOMA_SIMPLES } = require("./casos.js");

const tokens = (texto, sep) => F.tokenizar(texto, sep).map((t) => `${t.tipo}:${t.texto}`);

// Uma linha por nó: nível, Elemento, Info, Valor, Local.
function achatar(no, nivel = 0, linhas = []) {
  linhas.push([nivel, no.elemento, no.info, no.valor, no.local]);
  no.filhos.forEach((f) => achatar(f, nivel + 1, linhas));
  return linhas;
}

async function calcular(formula, celulas = {}, nomes = {}) {
  const ast = F.analisar(F.tokenizar(formula));
  await Av.avaliar(ast, criarAdaptador(celulas, nomes), { folhaBase: "Modelo", idioma: "pt-BR" });
  return ast.val;
}

async function principal() {
  // ---------- Tokens ----------
  assert.deepEqual(tokens("=SUM(O111:OFFSET(O116,-1,,,))"), [
    "func:SUM", "abre:(", "ref:O111", "op::", "func:OFFSET", "abre:(", "ref:O116", "sep:,",
    "op:-", "num:1", "sep:,", "sep:,", "sep:,", "fecha:)", "fecha:)",
  ]);
  assert.ok(F.alinhados(F.tokenizar(DESLOC.formula), F.tokenizar(DESLOC.formulaLocal, PT)));
  assert.deepEqual(
    tokens("='Aba com espaço'!B1+Modelo!B4:B5+Receita+Plan1!Nome+A:A+3:5+Tabela1[Col]+A1#+[1]Ext!A1"),
    [
      "ref:'Aba com espaço'!B1", "op:+", "ref:Modelo!B4:B5", "op:+", "nome:Receita", "op:+",
      "nome:Plan1!Nome", "op:+", "ref:A:A", "op:+", "ref:3:5", "op:+", "estruturada:Tabela1[Col]",
      "op:+", "ref:A1#", "op:+", "ref:[1]Ext!A1",
    ],
  );
  assert.deepEqual(tokens('=IF(A1="a""b",1.5,TRUE)'), [
    "func:IF", "abre:(", "ref:A1", "op:=", 'txt:"a""b"', "sep:,", "num:1.5", "sep:,", "logico:TRUE", "fecha:)",
  ]);
  assert.deepEqual(tokens('=SE(A1="x";1,5;VERDADEIRO)', PT), [
    "func:SE", "abre:(", "ref:A1", "op:=", 'txt:"x"', "sep:;", "num:1,5", "sep:;", "logico:VERDADEIRO", "fecha:)",
  ]);
  assert.deepEqual(tokens("=#REF!+Plan1!#REF!"), ["erro:#REF!", "op:+", "erro:Plan1!#REF!"]);
  assert.deepEqual(tokens("=LOG10(A1)").slice(0, 1), ["func:LOG10"]);
  assert.equal(F.analisar(F.tokenizar("=_xlfn.XLOOKUP(A1,B:B,C:C)")).nome, "XLOOKUP");

  // ---------- Erros de leitura ----------
  for (const ruim of ["=SUM(1,", "=1+", "=(1+2", "=1 2", "="]) {
    assert.throws(() => F.analisar(F.tokenizar(ruim)), F.ErroFormula, ruim);
  }

  // ---------- Cálculo ----------
  assert.deepEqual(await calcular("=-2^2"), { t: "num", v: 4 });
  assert.deepEqual(await calcular("=1+2*3"), { t: "num", v: 7 });
  assert.deepEqual(await calcular("=(1+2)*3"), { t: "num", v: 9 });
  assert.deepEqual(await calcular("=2*3%"), { t: "num", v: 0.06 });
  assert.deepEqual(await calcular('="a"&"b"&1.5'), { t: "txt", v: "ab1,5" });
  assert.deepEqual(await calcular("=1/0"), { t: "erro", v: "#DIV/0!" });
  assert.deepEqual(await calcular('=IF(A1>0,"pos","neg")', { "Modelo!A1": 5 }), { t: "txt", v: "pos" });
  assert.deepEqual(await calcular('=A1=""'), { t: "bool", v: true });
  assert.deepEqual(await calcular("=IFERROR(1/0,7)"), { t: "num", v: 7 });
  assert.deepEqual(await calcular("=AND(1>0,2>1)"), { t: "bool", v: true });
  assert.deepEqual(await calcular('=SUM(INDIRECT("B1:B2"))', { "Modelo!B1": 1, "Modelo!B2": 2 }), { t: "num", v: 3 });
  assert.deepEqual(await calcular("=SUM(Receita)", CELULAS, { RECEITA: "Premissas!C2" }), { t: "num", v: 0.1 });
  assert.deepEqual(await calcular("=Taxa*2", {}, { TAXA: 0.5 }), { t: "num", v: 1 });
  assert.deepEqual(await calcular("=Inexistente"), { t: "erro", v: "#NAME?" });
  assert.equal((await calcular("=XLOOKUP(1,A:A,B:B)")).t, "nd");
  assert.equal((await calcular("=SUM(OFFSET(A1,-1,0))")).v, "#REF!");

  // ---------- Árvore: print do DESLOC ----------
  const desloc = await Av.montarDados(DESLOC, criarAdaptador(CELULAS));
  assert.equal(desloc.celula, "Modelo!O110");
  assert.equal(desloc.aviso, "");
  assert.deepEqual(achatar(desloc.arvore), [
    [0, "O110", "", "-2.461.833,08", "O110"],
    [1, "SOMA", "", "-2.461.833,08", "O111:O115"],
    [2, "O111:DESLOC(O116;-1;;;)", "número1", "-2.461.833,08", "O111:O115"],
    [3, "O111", "", "-1.031.327,53", "O111"],
    [3, "DESLOC", "", "-200.000,00", "O115"],
    [4, "O116", "ref", "199.245.096,56", "O116"],
    [4, "-1", "lins", "-1", ""],
    [4, "<omitido>", "cols", "0", ""],
    [4, "<omitido>", "altura", "O mesmo que ref", ""],
    [4, "<omitido>", "largura", "O mesmo que ref", ""],
  ]);
  const soma = desloc.arvore.filhos[0];
  assert.equal(desloc.formula.slice(soma.ini, soma.fim), "SOMA(O111:DESLOC(O116;-1;;;))");
  const intervalo = soma.filhos[0];
  assert.equal(desloc.formula.slice(intervalo.ini, intervalo.fim), "O111:DESLOC(O116;-1;;;)");
  assert.equal(intervalo.filhos[1].navegar, "Modelo!O115");
  assert.equal(intervalo.filhos[0].navegar, "Modelo!O111");
  // "alvo": o que o Excel seleciona quando a linha é clicada na janela.
  assert.equal(desloc.arvore.alvo, "Modelo!O110");
  assert.equal(soma.alvo, "Modelo!O111:O115");
  assert.equal(intervalo.filhos[1].filhos[0].alvo, "Modelo!O116");
  assert.equal(intervalo.filhos[1].filhos[1].alvo, null);
  assert.equal(intervalo.filhos[1].filhos[2].alvo, null);

  // ---------- Árvore: print do E23+E24 ----------
  const simples = await Av.montarDados(SOMA_SIMPLES, criarAdaptador(CELULAS));
  assert.deepEqual(achatar(simples.arvore), [
    [0, "E25", "", "14.530,26", "E25"],
    [1, "E23+E24", "", "14.530,26", "E23:E24"],
    [2, "E23", "", "<vazio>", "E23"],
    [2, "E24", "", "14.530,26", "E24"],
  ]);
  const conta = simples.arvore.filhos[0];
  assert.equal(simples.formula.slice(conta.ini, conta.fim), "E23+E24");

  // ---------- Outras abas, argumento opcional ausente e alinhamento ----------
  const procv = await Av.montarDados({
    ...SOMA_SIMPLES, ref: "B9", texto: "", formula: "=VLOOKUP(Premissas!C2,A1:B3,2)", formulaLocal: "=PROCV(Premissas!C2;A1:B3;2)",
  }, criarAdaptador(CELULAS));
  assert.deepEqual(achatar(procv.arvore).slice(1), [
    [1, "PROCV", "", "—", "várias abas"],
    [2, "Premissas!C2", "valor_procurado", "0,10", "Premissas!C2"],
    [2, "A1:B3", "matriz_tabela", "1.000", "A1:B3"],
    [2, "2", "núm_índice_coluna", "2", ""],
    [2, "<omitido>", "procurar_intervalo", "VERDADEIRO", ""],
  ]);
  assert.equal(procv.arvore.filhos[0].alvo, null);
  assert.equal(procv.arvore.filhos[0].filhos[0].alvo, "Premissas!C2");
  const desalinhado = await Av.montarDados({ ...SOMA_SIMPLES, formulaLocal: "=SOMA(E23)" }, criarAdaptador(CELULAS));
  assert.match(desalinhado.aviso, /inglês/);
  assert.equal(desalinhado.arvore.filhos[0].ini, null);
  const semFormula = await Av.montarDados({ ...SOMA_SIMPLES, formula: 42, formulaLocal: 42 }, criarAdaptador(CELULAS));
  assert.equal(semFormula.formula, "42");
  assert.match(semFormula.aviso, /não tem fórmula/);

  console.log("formula.js e avaliador.js: todos os testes passaram");
}

principal().catch((e) => {
  console.error(e);
  process.exit(1);
});
