/* global Office */
"use strict";

// Janela da Lupa: mostra o que o painel calcula e devolve as ações do usuário.
// Modos: "formula" (árvore da fórmula, editável), "rastreio" (precedentes ou dependentes em blocos)
// e "fluxo" (escolha do âmbito e, depois, os grupos de células).
// Para o painel: pronta, mostrar, explorar, voltar, gravar, fechar, fluxo, visao, cor.
// Do painel: dados, erro, progresso. Fora do Excel: janela.html?demo[=rastreio|fluxo|escolher].

const estado = {
  dados: null, modo: "formula", raiz: null, todos: [], visiveis: [], abertos: new Set(),
  cursor: null, ancora: null, selecionados: new Set(), original: "", demo: false, temporizador: null, visao: null,
};

const $ = (id) => document.getElementById(id);

Office.onReady(() => {
  ligarEventos();
  const ui = Office.context && Office.context.ui;
  const demo = new URLSearchParams(location.search).has("demo");
  if (!demo && ui && typeof ui.messageParent === "function" && typeof ui.addHandlerAsync === "function") {
    ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => receber(JSON.parse(arg.message)));
    enviar({ tipo: "pronta" });
  } else {
    estado.demo = true;
    carregarDemo();
  }
});

// ---------- Conversa com o painel ----------

function enviar(msg) {
  if (estado.demo) {
    if (msg.tipo === "fluxo") carregarDemo("fluxo");
    else mostrarMensagem(`(demonstração) ação enviada ao Excel: ${msg.tipo}`);
    return;
  }
  try {
    Office.context.ui.messageParent(JSON.stringify(msg));
  } catch (e) {
    mostrarMensagem(`Não consegui falar com o Excel: ${e.message}`, true);
  }
}

function receber(msg) {
  if (msg.tipo === "dados" || msg.tipo === "arvore") mostrarDados(msg.dados);
  else if (msg.tipo === "erro") mostrarMensagem(msg.mensagem, true);
  else if (msg.tipo === "progresso") mostrarMensagem(msg.texto);
}

async function carregarDemo(nome = new URLSearchParams(location.search).get("demo")) {
  if (nome === "escolher") {
    mostrarDados({ modo: "fluxo", etapa: "escolher", titulo: "Fluxo de cálculo", opcoes: { aba: "Modelo", selecao: "Modelo!A2:A3" } });
    return;
  }
  try {
    const resposta = await fetch(nome ? `demo-${nome}.json` : "demo.json", { cache: "no-store" });
    mostrarDados(await resposta.json());
  } catch (e) {
    mostrarMensagem("Abra esta janela pelo Excel, com Ctrl+Q.", true);
  }
}

// ---------- Dados recebidos ----------

function mostrarDados(dados) {
  estado.dados = dados;
  estado.modo = dados.modo || "formula";
  document.body.dataset.modo = estado.modo;
  document.body.dataset.etapa = dados.etapa || "";
  $("titulo").textContent = dados.titulo || dados.celula || "";
  $("btn-voltar").disabled = !dados.podeVoltar;
  mostrarMensagem(dados.aviso || "");
  const area = $("formula");
  area.readOnly = !dados.editavel;
  area.value = dados.formula || "";
  estado.original = area.value;
  atualizarBotoes();

  $("fluxo-form").hidden = dados.etapa !== "escolher";
  $("abas").hidden = !dados.visoes;
  if (dados.etapa === "escolher") {
    prepararFormulario(dados.opcoes || {});
    return;
  }
  if (dados.visoes) {
    const disponivel = (k) => Boolean(dados.visoes[k]);
    estado.visao = disponivel(dados.visao) ? dados.visao : Object.keys(dados.visoes).find(disponivel);
    montarAbas();
    prepararArvore(dados.visoes[estado.visao]);
  } else {
    prepararArvore(dados.arvore);
  }
  renderizar();
  atualizarFormulaDaLinha();
  $("arvore").focus();
}

function prepararArvore(raiz) {
  estado.raiz = raiz;
  estado.todos = [];
  estado.abertos.clear();
  if (!raiz) return;
  const preparar = (no, pai, nivel) => {
    no.id = estado.todos.length;
    no.pai = pai;
    no.nivel = nivel;
    estado.todos.push(no);
    no.filhos.forEach((f) => preparar(f, no, nivel + 1));
  };
  preparar(raiz, null, 0);
  // Árvore pequena abre inteira; grande abre só o começo.
  const niveisAbertos = estado.modo === "formula" ? 2 : 1;
  for (const no of estado.todos) {
    if (no.filhos.length && (estado.todos.length <= 60 || no.nivel < niveisAbertos)) estado.abertos.add(no.id);
  }
  estado.cursor = raiz;
  estado.ancora = raiz;
  estado.selecionados = new Set([raiz]);
}

function montarAbas() {
  const caixa = $("abas");
  caixa.replaceChildren();
  for (const [chave, rotulo] of [["fluxos", "Fluxos"], ["entradasSaidas", "Entradas e saídas"]]) {
    if (!estado.dados.visoes[chave]) continue;
    const botao = el("button", chave === estado.visao ? "aba ativa" : "aba", rotulo);
    botao.addEventListener("click", () => trocarVisao(chave));
    caixa.append(botao);
  }
}

function trocarVisao(chave) {
  if (chave === estado.visao) return;
  estado.visao = chave;
  montarAbas();
  prepararArvore(estado.dados.visoes[chave]);
  renderizar();
  enviar({ tipo: "visao", visao: chave });
}

function prepararFormulario(opcoes) {
  $("op-aba").textContent = opcoes.aba ? `"${opcoes.aba}"` : "";
  const semSelecao = "(selecione um bloco antes de abrir)";
  $("op-sel-ambito").textContent = opcoes.selecao || semSelecao;
  $("op-sel-foco").textContent = opcoes.selecao || semSelecao;
  for (const campo of document.querySelectorAll('input[value="selecao"]')) campo.disabled = !opcoes.selecao;
  document.querySelector('input[name="ambito"][value="arquivo"]').checked = true;
  document.querySelector('input[name="foco"][value="ambito"]').checked = true;
  $("btn-analisar").focus();
}

function analisarFluxo() {
  const ambito = document.querySelector('input[name="ambito"]:checked').value;
  const foco = document.querySelector('input[name="foco"]:checked').value;
  mostrarMensagem("Analisando…");
  enviar({ tipo: "fluxo", ambito, foco });
}

// ---------- Árvore ----------

function renderizar() {
  const caixa = $("linhas");
  caixa.replaceChildren();
  estado.visiveis = [];
  const visitar = (no) => {
    estado.visiveis.push(no);
    caixa.append(linhaDe(no));
    if (estado.abertos.has(no.id)) no.filhos.forEach(visitar);
  };
  if (estado.raiz) visitar(estado.raiz);
  const atual = caixa.querySelector(".cursor");
  if (atual) atual.scrollIntoView({ block: "nearest" });
  atualizarRealce();
}

function linhaDe(no) {
  const classes = ["linha"];
  if (estado.selecionados.has(no)) classes.push("sel");
  if (no === estado.cursor) classes.push("cursor");
  if (no.omitido) classes.push("omitido");
  if (no.escolhido) classes.push("escolhido");
  const linha = el("div", classes.join(" "));
  linha.setAttribute("role", "treeitem");
  if (no.dica) linha.title = no.dica;

  const elemento = el("span", "elemento");
  elemento.style.paddingLeft = `${4 + no.nivel * 14}px`;
  const temFilhos = no.filhos.length > 0;
  const exp = el("span", temFilhos ? "exp" : "exp vazio", temFilhos ? (estado.abertos.has(no.id) ? "−" : "+") : "");
  exp.addEventListener("click", (ev) => {
    ev.stopPropagation();
    alternar(no);
  });
  elemento.append(exp);
  if (no.cor) {
    const amostra = el("span", no.apagada ? "amostra apagada" : "amostra");
    amostra.style.background = no.cor;
    amostra.title = "Liga ou desliga a cor deste grupo na planilha";
    amostra.addEventListener("click", (ev) => {
      ev.stopPropagation();
      alternarCor(no);
    });
    elemento.append(amostra);
  }
  elemento.append(el("span", "txt", no.elemento));
  if (!no.dica) elemento.title = no.elemento;

  linha.append(elemento, el("span", "info", no.info), el("span", "valor", no.valor), el("span", "local", no.local));
  linha.addEventListener("click", (ev) => clicar(no, ev));
  linha.addEventListener("dblclick", () => explorar("precedentes"));
  return linha;
}

// Clique: uma linha. Cmd/Ctrl+clique: junta ou tira da seleção. Shift+clique: faixa de linhas.
function clicar(no, ev) {
  if (ev.shiftKey && estado.ancora) {
    selecionarFaixa(no);
    return;
  }
  if (ev.metaKey || ev.ctrlKey) {
    if (estado.selecionados.has(no) && estado.selecionados.size > 1) estado.selecionados.delete(no);
    else estado.selecionados.add(no);
  } else {
    estado.selecionados = new Set([no]);
  }
  estado.cursor = no;
  estado.ancora = no;
  aposSelecionar();
}

function selecionarUnico(no) {
  if (!no) return;
  estado.selecionados = new Set([no]);
  estado.cursor = no;
  estado.ancora = no;
  aposSelecionar();
}

function selecionarFaixa(no) {
  const a = estado.visiveis.indexOf(estado.ancora);
  const b = estado.visiveis.indexOf(no);
  if (a < 0 || b < 0) {
    selecionarUnico(no);
    return;
  }
  estado.selecionados = new Set(estado.visiveis.slice(Math.min(a, b), Math.max(a, b) + 1));
  estado.cursor = no;
  aposSelecionar();
}

function aposSelecionar() {
  renderizar();
  atualizarFormulaDaLinha();
  mostrarNoExcel();
}

// As linhas selecionadas selecionam no Excel as células delas. A pequena espera evita uma
// seleção no Excel a cada seta do teclado.
function mostrarNoExcel() {
  clearTimeout(estado.temporizador);
  const enderecos = [...estado.selecionados].map((n) => n.alvo).filter(Boolean);
  if (enderecos.length) estado.temporizador = setTimeout(() => enviar({ tipo: "mostrar", enderecos }), 120);
}

// Explora (precedentes) ou rastreia (dependentes) as linhas selecionadas.
function explorar(direcao) {
  clearTimeout(estado.temporizador);
  const enderecos = [...estado.selecionados].map((n) => n.alvo).filter(Boolean);
  if (!enderecos.length) {
    mostrarMensagem("Esta linha não aponta para células.");
    return;
  }
  mostrarMensagem(direcao === "dependentes" ? "Rastreando dependentes…" : "Explorando…");
  enviar({ tipo: "explorar", enderecos, direcao });
}

function alternar(no) {
  if (!no.filhos.length) return;
  if (estado.abertos.has(no.id)) {
    estado.abertos.delete(no.id);
    // Se o cursor sumiu dentro do ramo fechado, ele sobe para o ramo.
    for (let p = estado.cursor && estado.cursor.pai; p; p = p.pai) {
      if (p === no) {
        selecionarUnico(no);
        return;
      }
    }
  } else {
    estado.abertos.add(no.id);
  }
  renderizar();
}

function alternarCor(no) {
  no.apagada = !no.apagada;
  renderizar();
  enviar({ tipo: "cor", categoria: no.categoria, ligada: !no.apagada });
}

function teclaArvore(ev) {
  const no = estado.cursor;
  if (!no) return;
  if (ev.ctrlKey && ev.code === "KeyQ") {
    ev.preventDefault();
    explorar(ev.shiftKey ? "dependentes" : "precedentes");
    return;
  }
  const i = estado.visiveis.indexOf(no);
  const mover = (alvo) => {
    if (!alvo) return;
    if (ev.shiftKey) selecionarFaixa(alvo);
    else selecionarUnico(alvo);
  };
  switch (ev.key) {
    case "ArrowDown": mover(estado.visiveis[i + 1]); break;
    case "ArrowUp": mover(estado.visiveis[i - 1]); break;
    case "Home": selecionarUnico(estado.raiz); break;
    case "End": selecionarUnico(estado.visiveis[estado.visiveis.length - 1]); break;
    case "ArrowRight":
      if (no.filhos.length && !estado.abertos.has(no.id)) alternar(no);
      else if (no.filhos.length) selecionarUnico(no.filhos[0]);
      break;
    case "ArrowLeft":
      if (no.filhos.length && estado.abertos.has(no.id)) alternar(no);
      else if (no.pai) selecionarUnico(no.pai);
      break;
    case "Enter": fechar(false); break;
    case "F2": editarParte(); break;
    default: return;
  }
  ev.preventDefault();
}

// ---------- Fórmula ----------

// Destaca na fórmula os trechos das linhas selecionadas. Depois de editar, os trechos não valem mais.
function atualizarRealce() {
  const area = $("formula");
  const realce = $("realce");
  const texto = area.value;
  realce.replaceChildren();
  const trechos = estado.modo === "formula" && texto === estado.original
    ? [...estado.selecionados].filter((n) => n.ini != null && n !== estado.raiz).sort((a, b) => a.ini - b.ini)
    : [];
  let pos = 0;
  let primeira = null;
  for (const n of trechos) {
    if (n.ini < pos) continue;
    realce.append(texto.slice(pos, n.ini));
    const marca = el("mark", null, texto.slice(n.ini, n.fim));
    realce.append(marca);
    primeira = primeira || marca;
    pos = n.fim;
  }
  realce.append(texto.slice(pos) + "\n");
  if (primeira) {
    const topo = primeira.offsetTop - 6;
    if (topo < area.scrollTop || topo > area.scrollTop + area.clientHeight - 24) area.scrollTop = Math.max(0, topo);
  }
  realce.scrollTop = area.scrollTop;
}

// No rastreio, a caixa de baixo mostra a fórmula do bloco selecionado.
function atualizarFormulaDaLinha() {
  if (estado.modo === "formula" || !estado.cursor) return;
  $("formula").value = estado.cursor.formula || "";
  estado.original = $("formula").value;
  atualizarRealce();
}

// Clique na fórmula: seleciona a parte da árvore que está ali (Cmd/Ctrl+clique junta).
function clicarNaFormula(ev) {
  const area = $("formula");
  if (estado.modo !== "formula" || area.value !== estado.original || area.selectionStart !== area.selectionEnd) return;
  const pos = area.selectionStart;
  let melhor = null;
  for (const no of estado.todos) {
    if (no.ini == null || no === estado.raiz || pos < no.ini || pos > no.fim) continue;
    if (!melhor || no.fim - no.ini < melhor.fim - melhor.ini) melhor = no;
  }
  if (!melhor) return;
  for (let p = melhor.pai; p; p = p.pai) estado.abertos.add(p.id);
  if (ev.metaKey || ev.ctrlKey) {
    if (estado.selecionados.has(melhor) && estado.selecionados.size > 1) estado.selecionados.delete(melhor);
    else estado.selecionados.add(melhor);
  } else {
    estado.selecionados = new Set([melhor]);
  }
  estado.cursor = melhor;
  estado.ancora = melhor;
  aposSelecionar();
}

// F2: edita a parte selecionada (o trecho fica selecionado na caixa, pronto para digitar por cima).
function editarParte() {
  if (estado.modo !== "formula" || !estado.dados || !estado.dados.editavel) return;
  const area = $("formula");
  const no = estado.cursor;
  area.focus();
  if (no && no.ini != null && area.value === estado.original) area.setSelectionRange(no.ini, no.fim);
  else area.setSelectionRange(area.value.length, area.value.length);
}

function atualizarBotoes() {
  $("btn-aplicar").disabled = $("formula").value === estado.original;
}

function aplicar(fecharDepois) {
  const formula = $("formula").value;
  if (estado.modo !== "formula" || formula === estado.original) {
    if (fecharDepois) fechar(false);
    return;
  }
  mostrarMensagem("Gravando a fórmula…");
  enviar({ tipo: "gravar", formula, fechar: fecharDepois });
}

// voltar: o Excel volta ao lugar da exploração (Esc). manterCores: as cores ficam na planilha (Alt+X).
function fechar(voltar, manterCores = false) {
  clearTimeout(estado.temporizador);
  enviar({ tipo: "fechar", voltar, manterCores });
}

// ---------- Menu e mensagens ----------

function executarMenu(acao) {
  if (acao === "explorar") return explorar("precedentes");
  if (acao === "dependentes") return explorar("dependentes");
  if (acao === "manter") return fechar(false, true);
  if (!estado.raiz) return undefined;
  if (acao === "copiar") return copiar(textoDaArvore());
  estado.abertos.clear();
  for (const no of estado.todos) {
    if (no.filhos.length && (acao === "expandir" || no.nivel === 0)) estado.abertos.add(no.id);
  }
  if (acao === "recolher") selecionarUnico(estado.raiz);
  else renderizar();
  return undefined;
}

function textoDaArvore() {
  return estado.todos.map((no) => `${"  ".repeat(no.nivel)}${no.elemento}\t${no.info}\t${no.valor}\t${no.local}`).join("\n");
}

async function copiar(texto) {
  try {
    await navigator.clipboard.writeText(texto);
  } catch (e) {
    const area = document.createElement("textarea");
    area.value = texto;
    document.body.append(area);
    area.select();
    document.execCommand("copy");
    area.remove();
  }
  mostrarMensagem("Árvore copiada.");
}

function mostrarMensagem(texto, erro = false) {
  const m = $("mensagem");
  m.textContent = texto;
  m.className = erro ? "mensagem erro" : "mensagem";
  m.hidden = !texto;
}

function el(tag, classe, texto) {
  const e = document.createElement(tag);
  if (classe) e.className = classe;
  if (texto !== undefined) e.textContent = texto;
  return e;
}

function ligarEventos() {
  const area = $("formula");
  $("arvore").addEventListener("keydown", teclaArvore);
  area.addEventListener("click", clicarNaFormula);
  area.addEventListener("input", () => {
    atualizarBotoes();
    atualizarRealce();
  });
  area.addEventListener("scroll", () => {
    $("realce").scrollTop = area.scrollTop;
  });
  area.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") {
      ev.preventDefault();
      if (area.value !== estado.original) aplicar(false);
      else fechar(false);
    } else if (ev.key === "Escape" && area.value !== estado.original) {
      // Primeiro Esc desfaz a edição; o segundo fecha a janela.
      ev.preventDefault();
      ev.stopPropagation();
      area.value = estado.original;
      atualizarBotoes();
      atualizarRealce();
      $("arvore").focus();
    }
  });
  $("btn-voltar").addEventListener("click", () => enviar({ tipo: "voltar" }));
  $("btn-aplicar").addEventListener("click", () => aplicar(false));
  $("btn-ok").addEventListener("click", () => aplicar(true));
  $("btn-analisar").addEventListener("click", analisarFluxo);
  $("btn-mais").addEventListener("click", (ev) => {
    ev.stopPropagation();
    $("menu-mais").hidden = !$("menu-mais").hidden;
  });
  $("menu-mais").addEventListener("click", (ev) => {
    const acao = ev.target.dataset.acao;
    $("menu-mais").hidden = true;
    if (acao) executarMenu(acao);
  });
  document.addEventListener("click", () => {
    $("menu-mais").hidden = true;
  });
  document.addEventListener("keydown", (ev) => {
    if (ev.key === "Escape") {
      ev.preventDefault();
      fechar(true);
    } else if (ev.altKey && ev.code === "KeyX") {
      ev.preventDefault();
      fechar(false, true);
    } else if (ev.altKey && ev.key === "ArrowLeft" && ev.target !== area) {
      ev.preventDefault();
      if (!$("btn-voltar").disabled) enviar({ tipo: "voltar" });
    }
  });
}
