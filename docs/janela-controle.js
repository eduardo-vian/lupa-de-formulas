/* global Office, Enderecos, LupaExcel, LupaRastreio, LupaFluxo, LupaCores,
   VERSAO, registrarEvento, aviso, descreverErro, mostrarPainel */
"use strict";

// Janela da Lupa (diálogo do Office): abre, conversa com ela e guarda o histórico de explorações.
// Quem calcula é o painel; a janela só exibe. Uma exploração é { tipo: "formula", celula } ou
// { tipo: "rastreio", enderecos, direcao }. O fluxo de cálculo fica fora do histórico.
const Janela = (() => {
  const TAMANHO = { width: 46, height: 60 }; // em % da tela
  const MAX_HISTORICO = 100;
  const st = {
    dialogo: null, abrindo: false, pronta: false, ultima: null,
    modo: null, atual: null, historico: [], origem: null, fluxo: null, comCores: false,
  };

  const aberta = () => Boolean(st.dialogo) || st.abrindo;
  const lugarDe = (x) => (x.tipo === "formula" ? [x.celula] : x.enderecos);
  const normalizar = (lista) => lista.flatMap((e) => Enderecos.dividirAreas(e))
    .map((e) => Enderecos.semCifrao(e).toUpperCase()).sort().join("|");
  const mesmoLugar = (a, b) => normalizar(a) === normalizar(b);
  const mesma = (x, y) => x.tipo === y.tipo && x.direcao === y.direcao && mesmoLugar(lugarDe(x), lugarDe(y));

  // ---------- Atalhos e botões do painel ----------

  // Ctrl+Q (precedentes): uma célula abre a fórmula; várias células rastreiam os precedentes.
  // Ctrl+Shift+Q (dependentes): sempre o rastreio.
  async function explorarSelecao(direcao) {
    const sel = await LupaExcel.lerSelecao();
    if (!aberta()) st.origem = sel.enderecos;
    const x = direcao === "precedentes" && sel.celulas === 1
      ? { tipo: "formula", celula: sel.ativa }
      : { tipo: "rastreio", enderecos: sel.enderecos, direcao };
    await mostrar(x, true);
  }

  // Ctrl+Shift+Z e o botão Voltar. Com a janela fechada, primeiro volta ao lugar da exploração
  // atual (as linhas clicadas mudam a seleção); se já estiver lá, volta para a anterior.
  async function voltar() {
    if (!aberta() && st.atual) {
      const sel = await LupaExcel.lerSelecao().catch(() => null);
      const lugar = lugarDe(st.atual);
      if (sel && !mesmoLugar(sel.enderecos, lugar)) {
        await LupaExcel.selecionar(lugar);
        return;
      }
    }
    const anterior = st.historico.pop();
    if (!anterior) {
      registrarEvento("não há exploração anterior");
      if (aberta()) enviar({ tipo: "progresso", texto: "Não há exploração anterior." });
      return;
    }
    await LupaExcel.selecionar(lugarDe(anterior));
    if (aberta()) await mostrar(anterior, false);
    else st.atual = anterior;
  }

  // Ctrl+Shift+F: a janela pergunta o âmbito e o foco antes de analisar.
  async function abrirFluxo() {
    const sel = await LupaExcel.lerSelecao();
    if (!aberta()) st.origem = sel.enderecos;
    st.modo = "fluxo";
    st.fluxo = { selecao: sel, resultado: null, visao: null, desligadas: new Set() };
    garantirDialogo();
    await limparCores();
    const bloco = sel.enderecos.length === 1 && sel.celulas > 1 ? sel.enderecos[0] : null;
    enviar({
      tipo: "dados",
      dados: { modo: "fluxo", etapa: "escolher", titulo: "Fluxo de cálculo", opcoes: { aba: sel.folha, selecao: bloco } },
    });
  }

  // ---------- Explorações ----------

  async function mostrar(x, empilhar) {
    if (empilhar && st.atual && !mesma(st.atual, x)) {
      st.historico.push(st.atual);
      if (st.historico.length > MAX_HISTORICO) st.historico.shift();
    }
    st.atual = x;
    st.modo = x.tipo;
    st.fluxo = null;
    garantirDialogo();
    await limparCores();
    let dados;
    try {
      dados = x.tipo === "formula"
        ? await LupaExcel.montarDados(x.celula)
        : await LupaRastreio.montar(x.enderecos, x.direcao);
    } catch (e) {
      enviar({ tipo: "erro", mensagem: `Não consegui montar a exploração: ${descreverErro(e)}` });
      return;
    }
    const { cores, ...paraJanela } = dados;
    paraJanela.podeVoltar = st.historico.length > 0;
    enviar({ tipo: "dados", dados: paraJanela });
    if (cores && cores.length) await pintar(cores);
  }

  // Duplo clique, Ctrl+Q ou Ctrl+Shift+Q dentro da janela, nas linhas selecionadas.
  async function explorarEnderecos(enderecos, direcao) {
    const areas = enderecos.flatMap((e) => Enderecos.dividirAreas(e));
    if (!areas.length) return;
    const umaCelula = areas.length === 1 && !Enderecos.separar(areas[0]).ref.includes(":");
    const x = direcao === "precedentes" && umaCelula
      ? { tipo: "formula", celula: areas[0] }
      : { tipo: "rastreio", enderecos: areas, direcao };
    await LupaExcel.selecionar(areas);
    await mostrar(x, true);
  }

  async function gravar({ formula, fechar: fecharDepois }) {
    if (!st.atual || st.atual.tipo !== "formula") return;
    try {
      await LupaExcel.gravarFormula(st.atual.celula, formula);
    } catch (e) {
      enviar({ tipo: "erro", mensagem: `O Excel não aceitou a fórmula: ${descreverErro(e)}` });
      return;
    }
    registrarEvento(`fórmula gravada em ${st.atual.celula}`);
    if (fecharDepois) await fechar({ voltar: false });
    else await mostrar(st.atual, false);
  }

  // Esc fecha e volta ao lugar da exploração; Enter fecha e fica onde está; Alt+X mantém as cores.
  async function fechar({ voltar: voltarAoLugar = false, manterCores = false } = {}) {
    const lugar = st.modo === "fluxo" || !st.atual ? st.origem : lugarDe(st.atual);
    fecharDialogo();
    if (manterCores && st.comCores) {
      st.comCores = false;
      try {
        await LupaCores.manter();
      } catch (e) {
        registrarEvento(`não consegui manter as cores: ${descreverErro(e)}`);
      }
    } else {
      await limparCores();
    }
    if (voltarAoLugar && lugar) await LupaExcel.selecionar(lugar);
    st.modo = null;
    st.fluxo = null;
  }

  // ---------- Fluxo de cálculo ----------

  async function analisarFluxo({ ambito: tipoAmbito, foco: tipoFoco }) {
    if (!st.fluxo) return;
    const { selecao } = st.fluxo;
    const area = selecao.enderecos[0];
    const { folha, ref } = Enderecos.separar(area);
    const ambito = tipoAmbito === "arquivo" ? {} : tipoAmbito === "aba" ? { folha: selecao.folha } : { folha, ref };
    const foco = tipoFoco === "selecao" ? { folha, ref } : null;
    const rotulo = tipoAmbito === "arquivo" ? "arquivo inteiro" : tipoAmbito === "aba" ? `aba ${selecao.folha}` : area;
    enviar({ tipo: "progresso", texto: "Lendo as fórmulas e montando o mapa de dependências…" });
    try {
      const resultado = await LupaFluxo.analisar(ambito, foco, rotulo);
      st.fluxo.resultado = resultado;
      st.fluxo.visao = resultado.visoes.fluxos ? "fluxos" : "entradasSaidas";
      st.fluxo.desligadas.clear();
      enviar({
        tipo: "dados",
        dados: {
          modo: "fluxo", etapa: "resultado", titulo: `Fluxo de cálculo: ${rotulo}`,
          visoes: resultado.visoes, visao: st.fluxo.visao, aviso: resultado.aviso,
        },
      });
      await repintarFluxo();
    } catch (e) {
      enviar({ tipo: "erro", mensagem: `Não consegui analisar: ${descreverErro(e)}` });
    }
  }

  async function repintarFluxo() {
    const fluxo = st.fluxo;
    if (!fluxo || !fluxo.resultado) return;
    await limparCores();
    await pintar(fluxo.resultado.grupos[fluxo.visao].filter((g) => !fluxo.desligadas.has(g.categoria)));
  }

  // ---------- Cores ----------

  async function pintar(grupos) {
    try {
      await LupaCores.pintar(grupos);
      st.comCores = true;
    } catch (e) {
      registrarEvento(`não consegui colorir: ${descreverErro(e)}`);
    }
  }

  async function limparCores() {
    if (!st.comCores) return;
    st.comCores = false;
    try {
      await LupaCores.limpar();
    } catch (e) {
      registrarEvento(`não consegui tirar as cores: ${descreverErro(e)}`);
    }
  }

  // ---------- Diálogo ----------

  function garantirDialogo() {
    if (st.dialogo || st.abrindo) return;
    st.abrindo = true;
    st.pronta = false;
    const url = new URL(`janela.html?v=${VERSAO}`, location.href).href;
    Office.context.ui.displayDialogAsync(url, { ...TAMANHO, displayInIframe: false }, (resultado) => {
      st.abrindo = false;
      if (resultado.status !== Office.AsyncResultStatus.Succeeded) {
        registrarEvento(`não consegui abrir a janela: ${resultado.error.code} ${resultado.error.message}`);
        aviso("Não consegui abrir a janela da Lupa. Veja o diagnóstico abaixo.");
        mostrarPainel();
        return;
      }
      st.dialogo = resultado.value;
      st.dialogo.addEventHandler(Office.EventType.DialogMessageReceived, receber);
      st.dialogo.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
        registrarEvento(`janela fechada (${arg.error})`);
        esquecerDialogo();
        limparCores();
        st.modo = null;
        st.fluxo = null;
      });
    });
  }

  // A janela recém-aberta ainda está carregando: guarda a última mensagem e manda quando ela avisar.
  function enviar(msg) {
    if (msg.tipo !== "progresso") st.ultima = msg;
    if (!st.dialogo || !st.pronta) return;
    try {
      st.dialogo.messageChild(JSON.stringify(msg));
    } catch (e) {
      registrarEvento(`não consegui falar com a janela: ${descreverErro(e)}`);
    }
  }

  async function receber(arg) {
    let msg;
    try {
      msg = JSON.parse(arg.message);
    } catch (e) {
      return;
    }
    try {
      switch (msg.tipo) {
        case "pronta":
          st.pronta = true;
          if (st.ultima) enviar(st.ultima);
          break;
        case "mostrar": await LupaExcel.selecionar(msg.enderecos); break;
        case "explorar": await explorarEnderecos(msg.enderecos, msg.direcao); break;
        case "voltar": await voltar(); break;
        case "gravar": await gravar(msg); break;
        case "fechar": await fechar(msg); break;
        case "fluxo": await analisarFluxo(msg); break;
        case "visao":
          if (st.fluxo) {
            st.fluxo.visao = msg.visao;
            await repintarFluxo();
          }
          break;
        case "cor":
          if (st.fluxo) {
            if (msg.ligada) st.fluxo.desligadas.delete(msg.categoria);
            else st.fluxo.desligadas.add(msg.categoria);
            await repintarFluxo();
          }
          break;
        default:
          break;
      }
    } catch (e) {
      enviar({ tipo: "erro", mensagem: descreverErro(e) });
    }
  }

  function fecharDialogo() {
    if (st.dialogo) {
      try {
        st.dialogo.close();
      } catch (e) {
        // a janela já tinha sido fechada
      }
    }
    esquecerDialogo();
  }

  function esquecerDialogo() {
    st.dialogo = null;
    st.pronta = false;
    st.ultima = null;
  }

  return { explorarSelecao, abrirFluxo, voltar };
})();
