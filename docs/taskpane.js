/* global Office, Excel, Enderecos, LupaExcel */
"use strict";

const VERSAO = "0.2.0";
const LIMITE_CELULAS = 5000; // seleção maior que isso não busca precedentes nem dependentes
const LIMITE_ITENS = 200; // máximo de blocos listados por tipo
const TAMANHO_JANELA = { width: 42, height: 55 }; // em % da tela

// Conjuntos de requisitos conferidos no diagnóstico.
const CONJUNTOS = [
  ["ExcelApi", "1.7"], ["ExcelApi", "1.9"], ["ExcelApi", "1.12"], ["ExcelApi", "1.13"],
  ["ExcelApi", "1.14"], ["ExcelApi", "1.15"], ["ExcelApi", "1.17"],
  ["SharedRuntime", "1.1"], ["KeyboardShortcuts", "1.1"], ["RibbonApi", "1.1"], ["DialogApi", "1.2"],
];

// Ações do shortcuts.json: tecla padrão e o que cada uma faz.
const ACOES = {
  MostrarFormula: { tecla: "Ctrl+Q", executar: () => abrirJanela() },
  MostrarFormulaReserva: { tecla: "Ctrl+Alt+Shift+F", executar: () => mostrarPainel() },
};

const estado = { historico: [], eventos: [], diagnostico: [], rodada: 0 };

// Janela da fórmula (diálogo do Office). Os cálculos ficam aqui; a janela só exibe.
const janela = { dialogo: null, abrindo: false, pronta: false, ultima: null, dados: null, historico: [] };

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
});

function ligarBotoes() {
  $("btn-janela").addEventListener("click", abrirJanela);
  $("btn-voltar").addEventListener("click", voltar);
  $("btn-atualizar").addEventListener("click", () => {
    atualizar();
    diagnosticar();
  });
  $("btn-copiar").addEventListener("click", copiarDiagnostico);
}

// ---------- Atalhos ----------

function registrarAtalhos() {
  if (!Office.actions || typeof Office.actions.associate !== "function") {
    registrarEvento("Office.actions indisponível: os atalhos não vão funcionar nesta versão");
    return;
  }
  for (const [acao, { tecla, executar }] of Object.entries(ACOES)) {
    Office.actions.associate(acao, () => {
      registrarEvento(`atalho ${tecla} recebido`);
      $("ultimo-atalho").textContent = `Último atalho: ${tecla} às ${hora()}`;
      return executar();
    });
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

// ---------- Janela da fórmula ----------

async function abrirJanela() {
  janela.historico = [];
  if (!janela.dialogo && !janela.abrindo) abrirDialogo();
  await recarregarJanela(null);
}

function abrirDialogo() {
  janela.abrindo = true;
  janela.pronta = false;
  const url = new URL(`janela.html?v=${VERSAO}`, location.href).href;
  Office.context.ui.displayDialogAsync(url, { ...TAMANHO_JANELA, displayInIframe: false }, (resultado) => {
    janela.abrindo = false;
    if (resultado.status !== Office.AsyncResultStatus.Succeeded) {
      registrarEvento(`não consegui abrir a janela: ${resultado.error.code} ${resultado.error.message}`);
      aviso("Não consegui abrir a janela da fórmula. Veja o diagnóstico abaixo.");
      mostrarPainel();
      return;
    }
    janela.dialogo = resultado.value;
    janela.dialogo.addEventHandler(Office.EventType.DialogMessageReceived, receberDaJanela);
    janela.dialogo.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
      registrarEvento(`janela fechada (${arg.error})`);
      esquecerJanela();
    });
  });
}

// Lê a célula (ativa, ou o endereço dado), calcula a árvore e manda para a janela.
async function recarregarJanela(endereco) {
  try {
    janela.dados = await LupaExcel.montarDados(endereco);
  } catch (e) {
    enviarParaJanela({ tipo: "erro", mensagem: `Não consegui ler a célula: ${descreverErro(e)}` });
    return;
  }
  janela.dados.podeVoltar = janela.historico.length > 0;
  enviarParaJanela({ tipo: "arvore", dados: janela.dados });
}

// A janela recém-aberta ainda está carregando: guarda a última mensagem e manda quando ela avisar.
function enviarParaJanela(msg) {
  janela.ultima = msg;
  if (!janela.dialogo || !janela.pronta) return;
  try {
    janela.dialogo.messageChild(JSON.stringify(msg));
  } catch (e) {
    registrarEvento(`não consegui falar com a janela: ${descreverErro(e)}`);
  }
}

async function receberDaJanela(arg) {
  let msg;
  try {
    msg = JSON.parse(arg.message);
  } catch (e) {
    return;
  }
  switch (msg.tipo) {
    case "pronta":
      janela.pronta = true;
      if (janela.ultima) enviarParaJanela(janela.ultima);
      break;
    case "ir":
      if (janela.dados) janela.historico.push(janela.dados.celula);
      if (await selecionar(msg.endereco)) {
        await recarregarJanela(msg.endereco);
      } else {
        janela.historico.pop();
        enviarParaJanela({ tipo: "erro", mensagem: `Não consegui ir até ${msg.endereco}.` });
      }
      break;
    case "voltar": {
      const destino = janela.historico.pop();
      if (destino && (await selecionar(destino))) await recarregarJanela(destino);
      break;
    }
    case "gravar":
      await gravarDaJanela(msg);
      break;
    case "fechar":
      fecharJanela();
      break;
    default:
      break;
  }
}

async function gravarDaJanela(msg) {
  if (!janela.dados) return;
  try {
    await LupaExcel.gravarFormula(janela.dados.celula, msg.formula);
  } catch (e) {
    enviarParaJanela({ tipo: "erro", mensagem: `O Excel não aceitou a fórmula: ${descreverErro(e)}` });
    return;
  }
  registrarEvento(`fórmula gravada em ${janela.dados.celula}`);
  if (msg.fechar) fecharJanela();
  else await recarregarJanela(janela.dados.celula);
}

function fecharJanela() {
  if (janela.dialogo) {
    try {
      janela.dialogo.close();
    } catch (e) {
      // a janela já tinha sido fechada
    }
  }
  esquecerJanela();
}

function esquecerJanela() {
  janela.dialogo = null;
  janela.pronta = false;
  janela.ultima = null;
}

// ---------- Leitura da célula, precedentes e dependentes ----------

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
    return { alvo: ativa, nota: "Seleção com várias áreas: por enquanto uso só a célula ativa." };
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

// ---------- Navegação ----------

async function irPara(endereco) {
  const origem = await enderecoAtual();
  if (await selecionar(endereco) && origem) {
    estado.historico.push(origem);
    $("btn-voltar").disabled = false;
  }
}

async function voltar() {
  const destino = estado.historico.pop();
  $("btn-voltar").disabled = estado.historico.length === 0;
  if (destino) await selecionar(destino);
}

async function selecionar(endereco) {
  try {
    await Excel.run(async (context) => {
      const { folha, ref } = Enderecos.separar(endereco);
      const aba = folha
        ? context.workbook.worksheets.getItem(folha)
        : context.workbook.worksheets.getActiveWorksheet();
      aba.activate();
      aba.getRange(ref).select();
      await context.sync();
    });
    return true;
  } catch (e) {
    aviso(`Não consegui ir até ${Enderecos.semCifrao(endereco)}: ${descreverErro(e)}`);
    return false;
  }
}

async function enderecoAtual() {
  for (const ler of [(wb) => wb.getSelectedRange(), (wb) => wb.getActiveCell()]) {
    try {
      return await Excel.run(async (context) => {
        const r = ler(context.workbook);
        r.load("address");
        await context.sync();
        return r.address;
      });
    } catch (e) {
      // tenta a próxima forma de ler a posição atual
    }
  }
  return null;
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
    for (const [acao, tecla] of Object.entries(atuais)) {
      const ativo = tecla !== null;
      linhas.push([`Atalho ${acao}`, ativo ? tecla : "desativado (conflito resolvido para outra ação)", ativo]);
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
