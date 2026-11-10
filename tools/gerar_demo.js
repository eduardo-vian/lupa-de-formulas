// Gera os exemplos para ver a janela fora do Excel:
//   docs/demo.json (janela.html?demo)               fórmula do print do Arixcel
//   docs/demo-rastreio.json (janela.html?demo=rastreio)   dependentes em blocos
//   docs/demo-fluxo.json (janela.html?demo=fluxo)         fluxo de cálculo com foco
// Uso: node tools/gerar_demo.js
const fs = require("node:fs");
const path = require("node:path");
const E = require("../docs/enderecos.js");
const B = require("../docs/blocos.js");
const Av = require("../docs/avaliador.js");
const Fluxo = require("../docs/fluxo.js");
const { criarAdaptador } = require("./adaptador_falso.js");
const { CELULAS, DESLOC, FOLHAS_FLUXO } = require("./casos.js");

const PASTA = path.join(__dirname, "..", "docs");

function gravar(nome, dados) {
  fs.writeFileSync(path.join(PASTA, nome), `${JSON.stringify(dados, null, 2)}\n`);
  console.log(`docs/${nome} gerado`);
}

async function principal() {
  gravar("demo.json", { ...(await Av.montarDados(DESLOC, criarAdaptador(CELULAS))), podeVoltar: true });

  const { cores, ...rastreio } = B.arvoreDeRastreio({
    direcao: "dependentes",
    origens: ["Modelo!B3"],
    areas: [
      { folha: "Modelo", ret: E.analisarRef("B4"), celulas: [{ r: 4, c: 2, chave: "=R[-1]C*Premissas!R3C3", formula: "=B3*Premissas!C3", texto: "302,50" }] },
      {
        folha: "Resumo, v2",
        ret: E.analisarRef("B1:B3"),
        celulas: [1, 2, 3].map((r) => ({ r, c: 2, chave: "=Modelo!R3C2*RC1", formula: `=Modelo!$B$3*A${r}`, texto: "1.210,00" })),
      },
    ],
  });
  gravar("demo-rastreio.json", { ...rastreio, podeVoltar: true, cores: undefined });

  const resultado = Fluxo.analisar({ folhas: FOLHAS_FLUXO, ambito: {}, foco: { folha: "Modelo", ref: "A2:A3" } });
  const { visoes } = Fluxo.montarVisoes(resultado, "arquivo inteiro");
  gravar("demo-fluxo.json", { modo: "fluxo", etapa: "resultado", titulo: "Fluxo de cálculo: arquivo inteiro", visoes, visao: "fluxos" });
  return cores;
}

principal().catch((e) => {
  console.error(e);
  process.exit(1);
});
