/* global Office, Excel, Enderecos, LupaExcel, LupaCores, Janela */
"use strict";

const VERSAO = "0.3.0";
const LIMITE_CELULAS = 5000; // seleção maior que isso não busca precedentes nem dependentes
const LIMITE_ITENS = 200; // máximo de blocos listados por tipo

// Conjuntos de requisitos conferidos no diagnóstico.
const CONJUNTOS = [
  ["ExcelApi", "1.7"], ["ExcelApi", "1.9"], ["ExcelApi", "1.12"], ["ExcelApi", "1.13"],
  ["ExcelApi", "1.14"], ["ExcelApi", "1.15"], ["ExcelApi", "1.17"],
  ["SharedRuntime", "1.1"], ["KeyboardShortcuts", "1.1"], ["RibbonApi", "1.1"], ["DialogApi", "1.2"],
];

// Ações do shortcuts.json: tecla padrão, descrição e o que cada uma faz.
const ACOES = {
  MostrarFormula: { tecla: "Ctrl+Q", descricao: "Fórmula ou precedentes", executar: () => Janela.explorarSelecao("precedentes") },
  RastrearDependentes: { tecla: "Ctrl+Shift+Q", descricao: "Dependentes", executar: () => Janela.explorarSelecao("dependentes") },
  FluxoCalculo: { tecla: "Ctrl+Shift+F", descricao: "Fluxo de cálculo", executar: () => Janela.abrirFluxo() },
  VoltarExploracao: { tecla: "Ctrl+Shift+Z", descricao: "Voltar à exploração anterior", executar: () => Janela.voltar() },
  MostrarFormulaReserva: { tecla: "Ctrl+Alt+Shift+F", descricao: "Abrir o painel", executar: () => mostrarPainel() },
};

const estado = { historico: [], eventos: [], diagnostico: [], rodada: 0 };

const $ = (id) => document.getElementById(id);

Office.onReady((info) => {
  ligarBotoes();
  if (info.host !== Office.HostType.Excel) {
    aviso("Este painel só funciona dentro do Excel.");
    return;
  }
  registrarAtalhos();
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => atualizar());
  if (Office.addin && typeof Office.addin.onVisibilityModeChanged === "function") {
    Office.addin.onVisibilityModeChanged((args) => {
      if (args.visibilityMode === Office.VisibilityMode.taskpane) {
        registrarEvento("painel aberto");
        atualizar();
      }
    });
  }
  registrarEvento("complemento carregado");
  atualizar();
  diagnosticar();
  mostrarAtalhos();
});

function ligarBotoes() {
  const acao = (id, executar) => $(id).addEventListener("click", () => executarComAviso(executar));
  acao("btn-explorar", ACOES.MostrarFormula.executar);
  acao("btn-dependentes", ACOES.RastrearDependentes.executar);
  acao("btn-fluxo", ACOES.FluxoCalculo.executar);
  acao("btn-voltar-exploracao", ACOES.VoltarExploracao.executar);
  acao("btn-remover-cores", () => LupaCores.limpar({ todas: true }));
  acao("btn-fixar-cores", () => LupaCores.fixar());
  $("btn-voltar").addEventListener("click", voltar);
  $("btn-atualizar").addEventListener("click", () => {
    atualizar();
    diagnosticar();
  });
  $("btn-copiar").addEventListener("click", copiarDiagnostico);
  $("btn-salvar-atalhos").addEventListener("click", salvarAtalhos);
  $("btn-restaurar-atalhos").addEventListener("click", restaurarAtalhos);
}

async function executarComAviso(executar) {
  try {
    await executar();
  } catch (e) {
    aviso(`Não deu certo: ${descreverErro(e)}`);
    registrarEvento(`erro: ${descreverErro(e)}`);
  }
}

async function mostrarPainel() {
  try {
    await Office.addin.showAsTaskpane();
  } catch (e) {
    registrarEvento(`não consegui abrir o painel: ${descreverErro(e)}`);
  }
  await atualizar();
}

// ---------- Atalhos ----------

function registrarAtalhos() {
  if (!Office.actions || typeof Office.actions.associate !== "function") {
    registrarEvento("Office.actions indisponível: os atalhos não vão funcionar nesta versão");
    return;
  }
  for (const [acao, { tecla, executar }] of Object.entries(ACOES)) {
    try {
      Office.actions.associate(acao, () => {
        registrarEvento(`atalho ${tecla} recebido`);
        $("ultimo-atalho").textContent = `Último atalho: ${tecla} às ${hora()}`;
        return executarComAviso(executar);
      });
    } catch (e) {
      registrarEvento(`atalho ${acao} não registrado: ${descreverErro(e)}`);
    }
  }
}

// Só lista as ações que o Office conhece: uma versão antiga do shortcuts.json em cache não tem as novas.
async function mostrarAtalhos() {
  const caixa = $("atalhos");
  caixa.replaceChildren();
  if (!Office.context.requirements.isSetSupported("KeyboardShortcuts", "1.1")) {
    caixa.textContent = "Esta versão do Excel não deixa trocar os atalhos.";
    return;
  }
  let atuais = {};
  try {
    atuais = await Office.actions.getShortcuts();
  } catch (e) {
    caixa.textContent = `Não consegui ler os atalhos: ${descreverErro(e)}`;
    return;
  }
  for (const [acao, { descricao, tecla }] of Object.entries(ACOES)) {
    if (!(acao in atuais)) continue;
    const linha = elemento("label");
    const campo = document.createElement("input");
    campo.value = atuais[acao] || "";
    campo.placeholder = tecla;
    campo.dataset.acao = acao;
    linha.append(elemento("span", null, descricao), campo);
    caixa.append(linha);
  }
}

async function salvarAtalhos() {
  const novos = {};
  for (const campo of $("atalhos").querySelectorAll("input")) novos[campo.dataset.acao] = campo.value.trim() || null;
  try {
    await Office.actions.replaceShortcuts(novos);
    $("atalhos-msg").textContent = "Atalhos salvos.";
  } catch (e) {
    $("atalhos-msg").textContent = `O Office não aceitou: ${descreverErro(e)}`;
  }
  mostrarAtalhos();
}

async function restaurarAtalhos() {
  const padroes = {};
  for (const campo of $("atalhos").querySelectorAll("input")) padroes[campo.dataset.acao] = null;
  try {
    await Office.actions.replaceShortcuts(padroes);
    $("atalhos-msg").textContent = "Atalhos de volta ao padrão.";
  } catch (e) {
    $("atalhos-msg").textContent = `O Office não aceitou: ${descreverErro(e)}`;
  }
  mostrarAtalhos();
}

// ---------- Leitura da célula, precedentes e dependentes (painel) ----------

async function atualizar() {
  const rodada = ++estado.rodada;
  const valeAinda = () => rodada === estado.rodada; // descarta resultado de seleção antiga
  try {
    await Excel.run(async (context) => {
      const ativa = context.workbook.getActiveCell();
      ativa.load("address,cellCount,formulas,formulasLocal,text");
      await context.sync();
      if (!valeAinda()) return;
      mostrarCelula(ativa);

      const { alvo, nota } = await lerSelecao(context, ativa);
      if (!valeAinda()) return;
      $("sel-info").textContent = nota;
      if (alvo.cellCount > LIMITE_CELULAS) {
        const msg = `Seleção grande demais (${numero(alvo.cellCount)} células). Selecione até ${numero(LIMITE_CELULAS)}.`;
        mostrarLista("prec", { erro: msg });
        mostrarLista("dep", { erro: msg });
        return;
      }
      const prec = await lerLigacoes(context, alvo, "precedentes");
      if (!valeAinda()) return;
      mostrarLista("prec", prec);
      const dep = await lerLigacoes(context, alvo, "dependentes");
      if (!valeAinda()) return;
      mostrarLista("dep", dep);
    });
    if (valeAinda()) aviso("");
  } catch (e) {
    if (valeAinda()) aviso(`Não consegui ler a célula ativa: ${descreverErro(e)}`);
  }
}

// getSelectedRange falha quando a seleção tem várias áreas; nesse caso usa só a célula ativa.
async function lerSelecao(context, ativa) {
  const sel = context.workbook.getSelectedRange();
  sel.load("address,cellCount");
  try {
    await context.sync();
  } catch (e) {
    return { alvo: ativa, nota: "Seleção com várias áreas: aqui uso só a célula ativa (a janela usa todas)." };
  }
  const nota = sel.cellCount > 1
    ? `Precedentes e dependentes das ${numero(sel.cellCount)} células selecionadas (${Enderecos.semCifrao(sel.address)}).`
    : "";
  return { alvo: sel, nota };
}

async function lerLigacoes(context, alvo, tipo) {
  const metodo = tipo === "precedentes" ? "getDirectPrecedents" : "getDirectDependents";
  if (typeof alvo[metodo] !== "function") {
    return { erro: `Esta versão do Excel não oferece ${metodo}.` };
  }
  const areas = alvo[metodo]();
  areas.load("addresses");
  try {
    await context.sync();
  } catch (e) {
    if (e.code === "ItemNotFound") return { itens: [], total: 0 };
    return { erro: descreverErro(e) };
  }
  const blocos = areas.addresses.flatMap((a) => Enderecos.dividirAreas(a));
  const itens = blocos.slice(0, LIMITE_ITENS).map((endereco) => ({ endereco }));
  await detalharCelulas(context, itens);
  return { itens, total: blocos.length };
}

// Para blocos de uma célula só, busca o valor exibido e a fórmula. Falha aqui não impede a lista.
async function detalharCelulas(context, itens) {
  const pendentes = [];
  for (const item of itens) {
    const { folha, ref } = Enderecos.separar(item.endereco);
    if (!folha || ref.includes(":")) continue;
    const r = context.workbook.worksheets.getItem(folha).getRange(ref);
    r.load("text,formulasLocal");
    pendentes.push([item, r]);
  }
  if (!pendentes.length) return;
  try {
    await context.sync();
    for (const [item, r] of pendentes) {
      item.valor = r.text[0][0];
      const f = r.formulasLocal[0][0];
      if (typeof f === "string" && f.startsWith("=")) item.formula = f;
    }
  } catch (e) {
    registrarEvento(`detalhes das células indisponíveis: ${descreverErro(e)}`);
  }
}

// ---------- Navegação pelas listas do painel ----------

async function irPara(endereco) {
  const origem = await enderecoAtual();
  if (await LupaExcel.selecionar(endereco)) {
    if (origem) {
      estado.historico.push(origem);
      $("btn-voltar").disabled = false;
    }
  } else {
    aviso(`Não consegui ir até ${Enderecos.semCifrao(endereco)}.`);
  }
}

async function voltar() {
  const destino = estado.historico.pop();
  $("btn-voltar").disabled = estado.historico.length === 0;
  if (destino) await LupaExcel.selecionar(destino);
}

async function enderecoAtual() {
  try {
    return (await LupaExcel.lerSelecao()).enderecos.join(",");
  } catch (e) {
    return null;
  }
}

// ---------- Exibição ----------

function mostrarCelula(ativa) {
  const f = ativa.formulas[0][0];
  const temFormula = typeof f === "string" && f.startsWith("=");
  $("cel-endereco").textContent = Enderecos.semCifrao(ativa.address);
  $("cel-valor").textContent = ativa.text[0][0] || "(vazia)";
  $("cel-formula").textContent = temFormula ? ativa.formulasLocal[0][0] : "(sem fórmula)";
  $("cel-formula-en").textContent = temFormula ? f : "—";
}

function mostrarLista(tipo, resultado) {
  const lista = $(`lista-${tipo}`);
  const conta = $(`conta-${tipo}`);
  lista.replaceChildren();
  if (resultado.erro) {
    conta.textContent = "";
    lista.append(elemento("li", "erro", resultado.erro));
    return;
  }
  conta.textContent = resultado.total > resultado.itens.length
    ? `${resultado.itens.length} de ${resultado.total}`
    : String(resultado.total);
  if (!resultado.itens.length) {
    lista.append(elemento("li", "vazio", "Nenhum."));
    return;
  }
  for (const item of resultado.itens) {
    const li = elemento("li", "item");
    const botao = elemento("button", "ir", Enderecos.semCifrao(item.endereco));
    botao.title = "Ir até esta célula";
    botao.addEventListener("click", () => irPara(item.endereco));
    li.append(botao);
    if (item.valor !== undefined) li.append(elemento("span", "valor", item.valor));
    if (item.formula) li.append(elemento("div", "formula", item.formula));
    lista.append(li);
  }
}

function elemento(tag, classe, texto) {
  const el = document.createElement(tag);
  if (classe) el.className = classe;
  if (texto !== undefined) el.textContent = texto;
  return el;
}

function aviso(msg) {
  const el = $("aviso");
  el.textContent = msg;
  el.hidden = !msg;
}

// ---------- Diagnóstico ----------

async function diagnosticar() {
  const linhas = [];
  const d = Office.context.diagnostics || {};
  linhas.push(["Protótipo", VERSAO]);
  linhas.push(["Excel", `${d.platform || "?"} • versão ${d.version || "?"}`]);
  linhas.push(["Idioma", Office.context.displayLanguage || "?"]);
  for (const [nome, versao] of CONJUNTOS) {
    const ok = Office.context.requirements.isSetSupported(nome, versao);
    linhas.push([`${nome} ${versao}`, ok ? "sim" : "não", ok]);
  }
  await diagnosticarAtalhos(linhas);
  await diagnosticarFuncoes(linhas);
  estado.diagnostico = linhas;
  mostrarDiagnostico();
}

async function diagnosticarAtalhos(linhas) {
  if (!Office.context.requirements.isSetSupported("KeyboardShortcuts", "1.1")) return;
  try {
    const atuais = await Office.actions.getShortcuts();
    for (const acao of Object.keys(ACOES)) {
      if (!(acao in atuais)) {
        linhas.push([`Atalho ${acao}`, "o Office ainda não carregou esta ação", false]);
        continue;
      }
      const tecla = atuais[acao];
      linhas.push([`Atalho ${acao}`, tecla === null ? "desativado (conflito resolvido para outra ação)" : tecla, tecla !== null]);
    }
  } catch (e) {
    linhas.push(["Leitura dos atalhos", `erro: ${descreverErro(e)}`, false]);
  }
}

// Confirma que o Excel calcula funções sem gravar em células (base da janela de fórmula).
async function diagnosticarFuncoes(linhas) {
  try {
    await Excel.run(async (context) => {
      const f = context.workbook.functions;
      const soma = f.sum(1, 2);
      soma.load("value");
      await context.sync();
      linhas.push(["Excel calcula funções sem gravar (SOMA 1+2)", String(soma.value), soma.value === 3]);
    });
  } catch (e) {
    linhas.push(["Cálculo de funções pela API", `erro: ${descreverErro(e)}`, false]);
  }
}

function mostrarDiagnostico() {
  const lista = $("diag-lista");
  lista.replaceChildren();
  for (const [nome, valor, ok] of estado.diagnostico) {
    const li = elemento("li", ok === undefined ? "info" : ok ? "ok" : "falha");
    li.append(elemento("span", "nome", nome), elemento("span", "valor", valor));
    lista.append(li);
  }
  $("diag-texto").value = [
    ...estado.diagnostico.map(([nome, valor]) => `${nome}: ${valor}`),
    "",
    "Eventos:",
    ...estado.eventos,
  ].join("\n");
}

function registrarEvento(msg) {
  estado.eventos.push(`${hora()} ${msg}`);
  if (estado.eventos.length > 30) estado.eventos.shift();
  mostrarDiagnostico();
}

async function copiarDiagnostico() {
  const caixa = $("diag-texto");
  try {
    await navigator.clipboard.writeText(caixa.value);
  } catch (e) {
    caixa.select();
    document.execCommand("copy");
  }
  $("copiado").textContent = "Copiado.";
  setTimeout(() => ($("copiado").textContent = ""), 2000);
}

// ---------- Utilidades ----------

function descreverErro(e) {
  if (!e) return "erro desconhecido";
  return [e.code, e.message].filter(Boolean).join(" — ") || String(e);
}

function hora() {
  return new Date().toLocaleTimeString("pt-BR");
}

function numero(n) {
  return n.toLocaleString("pt-BR");
}
