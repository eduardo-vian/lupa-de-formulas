/* Fluxo de cálculo: classifica as células de um âmbito em entradas, cálculos, saídas e fórmulas
   órfãs, e mostra o que entra e o que sai de uma área de foco.
   Puro: recebe as fórmulas já lidas do Excel. Só conta referências diretas (DESLOC e INDIRETO não). */
(function (raiz) {
  "use strict";

  const carregar = (global, caminho) =>
    typeof module !== "undefined" && module.exports ? require(caminho) : raiz[global];
  const E = carregar("Enderecos", "./enderecos.js");
  const F = carregar("Formula", "./formula.js");
  const B = carregar("Blocos", "./blocos.js");

  const VAZIA = 0;
  const VALOR = 1;
  const FORMULA = 2;
  const MAX_BLOCOS = 500; // blocos listados por aba em cada categoria

  const CATEGORIAS = {
    fluxos: [
      { id: "precedenteFora", nome: "Fora: precedentes", descricao: "Fora do foco, usadas por células do foco", cor: "#F8CBAD" },
      { id: "entraFoco", nome: "Dentro: entradas", descricao: "No foco, usam valores de fora", cor: "#FFE699" },
      { id: "saiFoco", nome: "Dentro: saídas", descricao: "No foco, usadas por células de fora", cor: "#C6E0B4" },
      { id: "dependenteFora", nome: "Fora: dependentes", descricao: "Fora do foco, usam células do foco", cor: "#BDD7EE" },
    ],
    entradasSaidas: [
      { id: "entrada", nome: "Entradas", descricao: "Têm dependentes, mas nenhum precedente no âmbito", cor: "#FFE699" },
      { id: "calculo", nome: "Cálculos", descricao: "Têm precedentes e dependentes no âmbito", cor: "#BDD7EE" },
      { id: "saida", nome: "Saídas", descricao: "Têm precedentes, mas nenhum dependente no âmbito", cor: "#C6E0B4" },
      { id: "orfa", nome: "Fórmulas órfãs", descricao: "Fórmulas sem precedentes nem dependentes no âmbito", cor: "#F8CBAD" },
    ],
  };

  // dados: {
  //   folhas: [{ nome, r0, c0, linhas, colunas, formulas }]  (formulas = Range.formulas do intervalo usado)
  //   nomes:  { "NOME" ou "ABA!NOME": { folha, ref } }       (nomes definidos que apontam para intervalos)
  //   ambito: {} arquivo inteiro | { folha } uma aba | { folha, ref } um bloco
  //   foco:   null | { folha, ref }
  // }
  function analisar({ folhas, nomes = {}, ambito = {}, foco = null }) {
    const mapa = new Map();
    for (const dados of folhas) mapa.set(dados.nome.toUpperCase(), prepararFolha(dados, ambito));

    const focoFolha = foco ? mapa.get(foco.folha.toUpperCase()) : null;
    const focoRet = foco ? E.analisarRef(foco.ref) : null;
    // Foco igual ao âmbito não tem "fora": a visão de fluxos não se aplica.
    const focoIgualAmbito = foco && ambito.folha && ambito.ref
      && foco.folha.toUpperCase() === ambito.folha.toUpperCase()
      && E.formatarRef(focoRet) === E.formatarRef(E.analisarRef(ambito.ref));
    const comFoco = Boolean(focoFolha && focoRet && !focoIgualAmbito);
    const dentroDoFoco = (f, r, c) => comFoco && f === focoFolha
      && r >= focoRet.r1 && r <= focoRet.r2 && c >= focoRet.c1 && c <= focoRet.c2;

    // 1) Fórmulas do âmbito e as referências de cada uma.
    const formulas = [];
    for (const f of mapa.values()) {
      if (!f.temAmbito) continue;
      f.porIndice = new Map();
      for (let i = 0; i < f.linhas; i++) {
        for (let j = 0; j < f.colunas; j++) {
          const k = i * f.colunas + j;
          if (f.tipo[k] !== FORMULA || !f.noAmbito[k]) continue;
          const r = f.r0 + i;
          const c = f.c0 + j;
          const fm = { f, r, c, dentro: dentroDoFoco(f, r, c), refs: referencias(f.formulas[i][j], f, nomes, mapa) };
          f.porIndice.set(k, fm);
          formulas.push(fm);
        }
      }
    }

    // 2) Quem é usado por quem: cobertura das referências (com foco, separada por dentro e fora).
    const usadas = coberturas(formulas, () => true);
    const usadasPorDentro = comFoco ? coberturas(formulas, (fm) => fm.dentro) : null;
    const usadasPorFora = comFoco ? coberturas(formulas, (fm) => !fm.dentro) : null;
    const somaFoco = comFoco
      ? somasPrefixo(focoFolha, (k) => focoFolha.tipo[k] !== VAZIA && focoFolha.noAmbito[k]
        && dentroDoFoco(focoFolha, focoFolha.r0 + Math.floor(k / focoFolha.colunas), focoFolha.c0 + (k % focoFolha.colunas)))
      : null;
    const naoVaziasNoAmbito = (ref) => contar(ref.f, ref.f.somaAmbito, ref.ret);
    const naoVaziasNoFoco = (ref) => (ref.f === focoFolha ? contar(focoFolha, somaFoco, ref.ret) : 0);

    // 3) Classificação célula a célula.
    const es = { entrada: [], calculo: [], saida: [], orfa: [] };
    const fl = comFoco ? { precedenteFora: [], entraFoco: [], saiFoco: [], dependenteFora: [] } : null;
    let celulas = 0;
    for (const f of mapa.values()) {
      if (!f.temAmbito) continue;
      const cob = usadas.get(f);
      const cobDentro = usadasPorDentro && usadasPorDentro.get(f);
      const cobFora = usadasPorFora && usadasPorFora.get(f);
      for (let i = 0; i < f.linhas; i++) {
        for (let j = 0; j < f.colunas; j++) {
          const k = i * f.colunas + j;
          if (!f.noAmbito[k] || f.tipo[k] === VAZIA) continue;
          celulas++;
          const r = f.r0 + i;
          const c = f.c0 + j;
          const cel = { folha: f.nome, r, c };
          const temDependente = Boolean(cob && cob[k] > 0);
          const fm = f.tipo[k] === FORMULA ? f.porIndice.get(k) : null;
          if (fm) {
            // A própria célula dentro do intervalo que ela usa (referência circular) não conta.
            const temPrecedente = fm.refs.some((ref) => naoVaziasNoAmbito(ref) - (contem(ref, f, r, c) ? 1 : 0) > 0);
            if (temPrecedente) (temDependente ? es.calculo : es.saida).push(cel);
            else (temDependente ? es.entrada : es.orfa).push(cel);
          } else if (temDependente) {
            es.entrada.push(cel);
          }
          if (!comFoco) continue;
          const dentro = dentroDoFoco(f, r, c);
          if (!dentro && cobDentro && cobDentro[k] > 0) fl.precedenteFora.push(cel);
          if (dentro && cobFora && cobFora[k] > 0) fl.saiFoco.push(cel);
          if (fm && dentro && fm.refs.some((ref) => naoVaziasNoAmbito(ref) - naoVaziasNoFoco(ref) > 0)) fl.entraFoco.push(cel);
          if (fm && !dentro && fm.refs.some((ref) => naoVaziasNoFoco(ref) > 0)) fl.dependenteFora.push(cel);
        }
      }
    }
    return {
      entradasSaidas: agruparPorFolha(es),
      fluxos: fl ? agruparPorFolha(fl) : null,
      totais: { celulas, formulas: formulas.length },
    };
  }

  function prepararFolha(dados, ambito) {
    const { linhas, colunas } = dados;
    const tipo = new Uint8Array(linhas * colunas);
    const noAmbito = new Uint8Array(linhas * colunas);
    const todaFolha = !ambito.folha;
    const mesma = Boolean(ambito.folha) && ambito.folha.toUpperCase() === dados.nome.toUpperCase();
    const ret = mesma && ambito.ref ? E.analisarRef(ambito.ref) : null;
    for (let i = 0; i < linhas; i++) {
      for (let j = 0; j < colunas; j++) {
        const k = i * colunas + j;
        const v = dados.formulas[i][j];
        if (v === "" || v === null || v === undefined) tipo[k] = VAZIA;
        else tipo[k] = typeof v === "string" && v.startsWith("=") ? FORMULA : VALOR;
        const r = dados.r0 + i;
        const c = dados.c0 + j;
        const noBloco = !ret || (r >= ret.r1 && r <= ret.r2 && c >= ret.c1 && c <= ret.c2);
        noAmbito[k] = todaFolha || (mesma && noBloco) ? 1 : 0;
      }
    }
    const f = { ...dados, tipo, noAmbito, temAmbito: todaFolha || mesma };
    f.somaAmbito = somasPrefixo(f, (k) => tipo[k] !== VAZIA && noAmbito[k]);
    return f;
  }

  // Referências diretas de uma fórmula (em inglês), já ligadas às abas carregadas.
  function referencias(formula, folha, nomes, mapa) {
    let tokens;
    try {
      tokens = F.tokenizar(formula);
    } catch (e) {
      return [];
    }
    const refs = [];
    for (const t of tokens) {
      let alvo = null;
      if (t.tipo === "ref") {
        const { folha: nomeFolha, ref } = E.separar(t.texto);
        if (nomeFolha && nomeFolha.startsWith("[")) continue;
        const ret = E.analisarRef(ref.replace(/#$/, ""));
        if (ret) alvo = { nome: nomeFolha || folha.nome, ret };
      } else if (t.tipo === "nome") {
        const { folha: nomeFolha, ref: nome } = E.separar(t.texto);
        const escopo = (nomeFolha || folha.nome).toUpperCase();
        const def = nomes[`${escopo}!${nome.toUpperCase()}`] || nomes[nome.toUpperCase()];
        if (def) alvo = { nome: def.folha, ret: E.analisarRef(def.ref) };
      }
      const f = alvo && alvo.ret ? mapa.get(alvo.nome.toUpperCase()) : null;
      if (f) refs.push({ f, ret: alvo.ret });
    }
    return refs;
  }

  // Somas de prefixo 2D: contam, em tempo constante, as células de um retângulo que passam no teste.
  function somasPrefixo(f, teste) {
    const largura = f.colunas + 1;
    const s = new Int32Array((f.linhas + 1) * largura);
    for (let i = 0; i < f.linhas; i++) {
      for (let j = 0; j < f.colunas; j++) {
        s[(i + 1) * largura + j + 1] = (teste(i * f.colunas + j) ? 1 : 0)
          + s[i * largura + j + 1] + s[(i + 1) * largura + j] - s[i * largura + j];
      }
    }
    return s;
  }

  function recortar(f, ret) {
    const i1 = Math.max(ret.r1 - f.r0, 0);
    const i2 = Math.min(ret.r2 - f.r0, f.linhas - 1);
    const j1 = Math.max(ret.c1 - f.c0, 0);
    const j2 = Math.min(ret.c2 - f.c0, f.colunas - 1);
    return i1 > i2 || j1 > j2 ? null : { i1, i2, j1, j2 };
  }

  function contar(f, s, ret) {
    const q = recortar(f, ret);
    if (!q) return 0;
    const w = f.colunas + 1;
    return s[(q.i2 + 1) * w + q.j2 + 1] - s[q.i1 * w + q.j2 + 1] - s[(q.i2 + 1) * w + q.j1] + s[q.i1 * w + q.j1];
  }

  function contem(ref, f, r, c) {
    return ref.f === f && r >= ref.ret.r1 && r <= ref.ret.r2 && c >= ref.ret.c1 && c <= ref.ret.c2;
  }

  // Quantas vezes cada célula é usada pelas fórmulas que passam no filtro (soma de diferenças 2D).
  function coberturas(formulas, filtro) {
    const porFolha = new Map();
    for (const fm of formulas) {
      if (!filtro(fm)) continue;
      for (const ref of fm.refs) {
        if (!porFolha.has(ref.f)) porFolha.set(ref.f, []);
        porFolha.get(ref.f).push(ref.ret);
      }
    }
    const resultado = new Map();
    for (const [f, rets] of porFolha) {
      const largura = f.colunas + 1;
      const d = new Int32Array((f.linhas + 1) * largura);
      for (const ret of rets) {
        const q = recortar(f, ret);
        if (!q) continue;
        d[q.i1 * largura + q.j1]++;
        d[q.i1 * largura + q.j2 + 1]--;
        d[(q.i2 + 1) * largura + q.j1]--;
        d[(q.i2 + 1) * largura + q.j2 + 1]++;
      }
      const cob = new Int32Array(f.linhas * f.colunas);
      for (let i = 0; i < f.linhas; i++) {
        for (let j = 0; j < f.colunas; j++) {
          const k = i * f.colunas + j;
          cob[k] = d[i * largura + j]
            + (i > 0 ? cob[k - f.colunas] : 0)
            + (j > 0 ? cob[k - 1] : 0)
            - (i > 0 && j > 0 ? cob[k - f.colunas - 1] : 0);
        }
      }
      resultado.set(f, cob);
    }
    return resultado;
  }

  // { categoria: [células] } -> { categoria: [{ folha, rets, total }] }
  function agruparPorFolha(porCategoria) {
    const saida = {};
    for (const [categoria, celulas] of Object.entries(porCategoria)) {
      const porFolha = new Map();
      for (const cel of celulas) {
        if (!porFolha.has(cel.folha)) porFolha.set(cel.folha, []);
        porFolha.get(cel.folha).push(cel);
      }
      saida[categoria] = [...porFolha].map(([folha, lista]) => ({ folha, rets: B.retangulos(lista), total: lista.length }));
    }
    return saida;
  }

  // ---------- Árvores para a janela e grupos para colorir ----------

  function montarVisoes(resultado, rotulo) {
    const visoes = {};
    const grupos = {};
    for (const [visao, titulo] of [["fluxos", "Fluxos"], ["entradasSaidas", "Entradas e saídas"]]) {
      const porCategoria = resultado[visao];
      if (!porCategoria) {
        visoes[visao] = null;
        grupos[visao] = [];
        continue;
      }
      visoes[visao] = {
        ...B.noVazio(),
        elemento: titulo,
        info: rotulo,
        filhos: CATEGORIAS[visao].map((cat) => noDeCategoria(cat, porCategoria[cat.id])),
      };
      grupos[visao] = CATEGORIAS[visao].flatMap((cat) =>
        porCategoria[cat.id].map((g) => ({ categoria: cat.id, folha: g.folha, rets: g.rets, cor: cat.cor })));
    }
    return { visoes, grupos };
  }

  function noDeCategoria(cat, gruposDaCategoria) {
    const total = gruposDaCategoria.reduce((s, g) => s + g.total, 0);
    const abas = gruposDaCategoria.map((g) => {
      const blocos = g.rets.slice(0, MAX_BLOCOS).map((ret) => {
        const local = `${E.prefixoFolha(g.folha)}!${E.formatarRef(ret)}`;
        return { ...B.noVazio(), elemento: E.formatarRef(ret), info: B.contagem(B.totalCelulas([ret])), local, alvo: local };
      });
      if (g.rets.length > MAX_BLOCOS) {
        blocos.push({ ...B.noVazio(), elemento: `… mais ${g.rets.length - MAX_BLOCOS} blocos`, omitido: true });
      }
      return { ...B.noVazio(), elemento: g.folha, info: B.contagem(g.total), alvo: B.endereco(g.folha, g.rets), filhos: blocos };
    });
    return {
      ...B.noVazio(),
      elemento: cat.nome,
      info: B.contagem(total),
      dica: cat.descricao,
      cor: cat.cor,
      categoria: cat.id,
      alvo: abas.length ? abas[0].alvo : null,
      filhos: abas,
    };
  }

  const api = { CATEGORIAS, analisar, montarVisoes };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else raiz.Fluxo = api;
})(this);
