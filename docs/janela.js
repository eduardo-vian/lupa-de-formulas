/* global Office */
"use strict";

// Janela da fórmula: exibe a árvore que o painel calcula e devolve as ações do usuário.
// Mensagens para o painel: pronta, ir, voltar, gravar, fechar. Do painel: arvore, erro.
// Fora do Excel (janela.html?demo) mostra o exemplo de demo.json.

const estado = { dados: null, abertos: new Set(), selecionado: null, visiveis: [], original: "", demo: false };

const $ = (id) => document.getElementById(id);

Office.onReady(() => {
  ligarEventos();
  const ui = Office.context && Office.context.ui;
  const demo = /[?&]demo\b/.test(location.search);
  if (!demo && ui && typeof ui.messageParent === "function" && typeof ui.addHandlerAsync === "function") {
    ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => receber(JSON.parse(arg.message)));
    enviar({ tipo: "pronta" });
  } else {
    estado.demo = true;
    carregarDemo();
  }
});

function enviar(msg) {
  if (estado.demo) {
    mostrarMensagem(`(demonstração) ação enviada ao Excel: ${msg.tipo}`);
    return;
  }
  try {
    Office.context.ui.messageParent(JSON.stringify(msg));
  } catch (e) {
    mostrarMensagem(`Não consegui falar com o Excel: ${e.message}`, true);
  }
}

function receber(msg) {
  if (msg.tipo === "arvore") mostrarArvore(msg.dados);
  else if (msg.tipo === "erro") mostrarMensagem(msg.mensagem, true);
}

async function carregarDemo() {
  try {
    const resposta = await fetch("demo.json", { cache: "no-store" });
    mostrarArvore(await resposta.json());
  } catch (e) {
    mostrarMensagem("Abra esta janela pelo Excel, com Ctrl+Q.", true);
  }
}

// ---------- Árvore ----------

function mostrarArvore(dados) {
  estado.dados = dados;
  estado.abertos.clear();
  let total = 0;
  const preparar = (no, pai, nivel) => {
    no.id = total++;
    no.pai = pai;
    no.nivel = nivel;
    no.filhos.forEach((f) => preparar(f, no, nivel + 1));
  };
  preparar(dados.arvore, null, 0);
  // Árvore pequena abre inteira; grande abre só os dois primeiros níveis.
  const abrir = (no) => {
    if (no.filhos.length && (total <= 60 || no.nivel < 2)) estado.abertos.add(no.id);
    no.filhos.forEach(abrir);
  };
  abrir(dados.arvore);

  estado.selecionado = dados.arvore;
  $("titulo").textContent = dados.celula;
  $("btn-voltar").disabled = !dados.podeVoltar;
  $("formula").value = dados.formula || "";
  estado.original = $("formula").value;
  atualizarBotoes();
  mostrarMensagem(dados.aviso || "");
  renderizar();
  $("arvore").focus();
}

function renderizar() {
  const caixa = $("linhas");
  caixa.replaceChildren();
  estado.visiveis = [];
  const visitar = (no) => {
    estado.visiveis.push(no);
    caixa.append(linhaDe(no));
    if (estado.abertos.has(no.id)) no.filhos.forEach(visitar);
  };
  if (estado.dados) visitar(estado.dados.arvore);
  const selecionada = caixa.querySelector(".sel");
  if (selecionada) selecionada.scrollIntoView({ block: "nearest" });
  atualizarRealce();
}

function linhaDe(no) {
  const classes = ["linha"];
  if (no === estado.selecionado) classes.push("sel");
  if (no.omitido) classes.push("omitido");
  const linha = el("div", classes.join(" "));
  linha.setAttribute("role", "treeitem");

  const elemento = el("span", "elemento");
  elemento.style.paddingLeft = `${4 + no.nivel * 14}px`;
  elemento.title = no.elemento;
  const temFilhos = no.filhos.length > 0;
  const exp = el("span", temFilhos ? "exp" : "exp vazio", temFilhos ? (estado.abertos.has(no.id) ? "−" : "+") : "");
  exp.addEventListener("click", (ev) => {
    ev.stopPropagation();
    alternar(no);
  });
  elemento.append(exp, el("span", "txt", no.elemento));

  linha.append(elemento, el("span", "info", no.info), el("span", "valor", no.valor), el("span", "local", no.local));
  linha.addEventListener("click", () => selecionar(no));
  linha.addEventListener("dblclick", () => navegar(no));
  return linha;
}

function selecionar(no) {
  if (!no) return;
  estado.selecionado = no;
  renderizar();
}

function alternar(no) {
  if (!no.filhos.length) return;
  if (estado.abertos.has(no.id)) {
    estado.abertos.delete(no.id);
    // Se a linha selecionada sumiu dentro do ramo fechado, a seleção sobe para o ramo.
    for (let p = estado.selecionado; p; p = p.pai) {
      if (p.pai === no) {
        estado.selecionado = no;
        break;
      }
    }
  } else {
    estado.abertos.add(no.id);
  }
  renderizar();
}

function navegar(no) {
  if (!no.navegar) {
    mostrarMensagem("Esta parte não aponta para células.");
    return;
  }
  mostrarMensagem(`Indo para ${no.navegar}…`);
  enviar({ tipo: "ir", endereco: no.navegar });
}

function teclaArvore(ev) {
  const no = estado.selecionado;
  if (!no) return;
  const i = estado.visiveis.indexOf(no);
  switch (ev.key) {
    case "ArrowDown": selecionar(estado.visiveis[Math.min(i + 1, estado.visiveis.length - 1)]); break;
    case "ArrowUp": selecionar(estado.visiveis[Math.max(i - 1, 0)]); break;
    case "Home": selecionar(estado.visiveis[0]); break;
    case "End": selecionar(estado.visiveis[estado.visiveis.length - 1]); break;
    case "ArrowRight":
      if (no.filhos.length && !estado.abertos.has(no.id)) alternar(no);
      else if (no.filhos.length) selecionar(no.filhos[0]);
      break;
    case "ArrowLeft":
      if (no.filhos.length && estado.abertos.has(no.id)) alternar(no);
      else if (no.pai) selecionar(no.pai);
      break;
    case "Enter": navegar(no); break;
    default: return;
  }
  ev.preventDefault();
}

// ---------- Fórmula ----------

// Destaca na fórmula o trecho da linha selecionada. Depois de editar, os trechos não valem mais.
function atualizarRealce() {
  const area = $("formula");
  const realce = $("realce");
  const texto = area.value;
  const no = estado.selecionado;
  realce.replaceChildren();
  if (!no || texto !== estado.original || no.ini == null) {
    realce.append(texto + "\n");
  } else {
    const marca = el("mark", null, texto.slice(no.ini, no.fim));
    realce.append(texto.slice(0, no.ini), marca, texto.slice(no.fim) + "\n");
    const topo = marca.offsetTop - 6;
    if (topo < area.scrollTop || topo > area.scrollTop + area.clientHeight - 24) area.scrollTop = Math.max(0, topo);
  }
  realce.scrollTop = area.scrollTop;
}

function atualizarBotoes() {
  $("btn-aplicar").disabled = $("formula").value === estado.original;
}

function aplicar(fechar) {
  const formula = $("formula").value;
  if (formula === estado.original) {
    if (fechar) enviar({ tipo: "fechar" });
    return;
  }
  mostrarMensagem("Gravando a fórmula…");
  enviar({ tipo: "gravar", formula, fechar });
}

// ---------- Menu e mensagens ----------

function executarMenu(acao) {
  if (!estado.dados) return;
  if (acao === "copiar") {
    copiar(textoDaArvore());
    return;
  }
  estado.abertos.clear();
  const visitar = (no) => {
    if (no.filhos.length && (acao === "expandir" || no.nivel === 0)) estado.abertos.add(no.id);
    no.filhos.forEach(visitar);
  };
  visitar(estado.dados.arvore);
  if (acao === "recolher") estado.selecionado = estado.dados.arvore;
  renderizar();
}

function textoDaArvore() {
  const linhas = [];
  const visitar = (no) => {
    linhas.push(`${"  ".repeat(no.nivel)}${no.elemento}\t${no.info}\t${no.valor}\t${no.local}`);
    no.filhos.forEach(visitar);
  };
  visitar(estado.dados.arvore);
  return linhas.join("\n");
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
  area.addEventListener("input", () => {
    atualizarBotoes();
    atualizarRealce();
  });
  area.addEventListener("scroll", () => {
    $("realce").scrollTop = area.scrollTop;
  });
  area.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter" && (ev.ctrlKey || ev.metaKey)) {
      ev.preventDefault();
      aplicar(false);
    }
  });
  $("btn-voltar").addEventListener("click", () => enviar({ tipo: "voltar" }));
  $("btn-aplicar").addEventListener("click", () => aplicar(false));
  $("btn-ok").addEventListener("click", () => aplicar(true));
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
      enviar({ tipo: "fechar" });
    } else if (ev.key === "ArrowLeft" && ev.altKey && ev.target !== area) {
      ev.preventDefault();
      if (!$("btn-voltar").disabled) enviar({ tipo: "voltar" });
    }
  });
}
