// Gera docs/demo.json com o exemplo do print do Arixcel, para ver a janela fora do Excel
// (janela.html?demo). Uso: node tools/gerar_demo.js
const fs = require("node:fs");
const path = require("node:path");
const Av = require("../docs/avaliador.js");
const { criarAdaptador } = require("./adaptador_falso.js");
const { CELULAS, DESLOC } = require("./casos.js");

Av.montarDados(DESLOC, criarAdaptador(CELULAS)).then((dados) => {
  const saida = path.join(__dirname, "..", "docs", "demo.json");
  fs.writeFileSync(saida, `${JSON.stringify({ ...dados, podeVoltar: true }, null, 2)}\n`);
  console.log(`${saida} gerado`);
});
