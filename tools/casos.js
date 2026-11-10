// Casos de teste tirados dos prints do Arixcel. Usados nos testes e na demonstração da janela.
const PT = { decimal: ",", lista: ";" };

const CELULAS = {
  "Modelo!O111": -1031327.53,
  "Modelo!O112": -500000,
  "Modelo!O113": -430505.55,
  "Modelo!O114": -300000,
  "Modelo!O115": -200000,
  "Modelo!O116": 199245096.56,
  "Modelo!E24": 14530.26,
  "Modelo!B1": 1000,
  "Premissas!C2": 0.1,
};

const DESLOC = {
  folha: "Modelo",
  ref: "O110",
  texto: "-2.461.833,08",
  formula: "=SUM(O111:OFFSET(O116,-1,,,))",
  formulaLocal: "=SOMA(O111:DESLOC(O116;-1;;;))",
  separadores: PT,
  idioma: "pt-BR",
};

const SOMA_SIMPLES = {
  folha: "Modelo",
  ref: "E25",
  texto: "14.530,26",
  formula: "=E23+E24",
  formulaLocal: "=E23+E24",
  separadores: PT,
  idioma: "pt-BR",
};

module.exports = { PT, CELULAS, DESLOC, SOMA_SIMPLES };
