/* global Excel, Office, Enderecos, Avaliador */
"use strict";

// Ponte entre o Excel e o resto da Lupa: lê intervalos e a seleção, chama funções do Excel,
// seleciona células e grava fórmulas.
const LupaExcel = (() => {
  const LIMITE_VALORES = 20000; // intervalo maior que isso é usado sem carregar os valores
  const MAX_AREAS_SELECAO = 300;
  const NAO_SAO_FUNCOES = ["constructor", "toJSON", "load", "set", "track", "untrack"];
  let metodos = null;

  // Nome em inglês (SUM, NORM.DIST) -> método de context.workbook.functions (sum, norm_Dist).
  function metodo(nomeEn) {
    if (!metodos) {
      metodos = new Map();
      const proto = Excel.Functions && Excel.Functions.prototype;
      for (const nome of proto ? Object.getOwnPropertyNames(proto) : []) {
        if (nome.startsWith("_") || NAO_SAO_FUNCOES.includes(nome)) continue;
        const descritor = Object.getOwnPropertyDescriptor(proto, nome);
        if (descritor && typeof descritor.value === "function") {
          metodos.set(nome.toUpperCase().replace(/_/g, "."), nome);
        }
      }
    }
    return metodos.get(nomeEn) || null;
  }

  // ---------- Leitura para o avaliador ----------

  async function ler(pedidos) {
    try {
      return await Excel.run((context) => lerNoContexto(context, pedidos));
    } catch (e) {
      // Um pedido inválido derruba o lote inteiro; refaz um por um para isolar o problema.
      const resultado = new Map();
      for (const pedido of pedidos) {
        try {
          const um = await Excel.run((context) => lerNoContexto(context, [pedido]));
          um.forEach((val, chave) => resultado.set(chave, val));
        } catch (e2) {
          resultado.set(pedido.chave, { t: "erro", v: "#REF!" });
        }
      }
      return resultado;
    }
  }

  async function lerNoContexto(context, pedidos) {
    const wb = context.workbook;
    const resultado = new Map();

    // 1) Confere se as abas e os nomes existem. Nomes trazem também a definição.
    const conferidos = pedidos.map((pedido) => {
      if (pedido.nome !== undefined) {
        const colecao = pedido.folha ? wb.worksheets.getItem(pedido.folha).names : wb.names;
        const nome = colecao.getItemOrNullObject(pedido.nome);
        nome.load("type,value,formula");
        return { pedido, nome };
      }
      const aba = wb.worksheets.getItemOrNullObject(pedido.folha);
      aba.load("name");
      return { pedido, aba };
    });
    await context.sync();

    // 2) Pega os intervalos e o tamanho de cada um.
    const intervalos = [];
    for (const { pedido, nome, aba } of conferidos) {
      let intervalo;
      let definicao;
      if (nome) {
        if (nome.isNullObject) {
          resultado.set(pedido.chave, { t: "erro", v: "#NAME?" });
          continue;
        }
        definicao = nome.formula;
        if (nome.type !== "Range") {
          resultado.set(pedido.chave, { ...valorDeNome(nome), definicao });
          continue;
        }
        intervalo = nome.getRange();
      } else {
        if (aba.isNullObject) {
          resultado.set(pedido.chave, { t: "erro", v: "#REF!" });
          continue;
        }
        intervalo = aba.getRange(pedido.ref);
      }
      intervalo.load("address,cellCount");
      intervalos.push({ pedido, intervalo, definicao });
    }
    await context.sync();

    // 3) Valores, só dos intervalos que não são grandes demais.
    for (const { intervalo } of intervalos) {
      if (intervalo.cellCount <= LIMITE_VALORES) intervalo.load("values,valueTypes");
      if (intervalo.cellCount === 1) intervalo.load("text");
    }
    if (intervalos.length) await context.sync();
    for (const { pedido, intervalo, definicao } of intervalos) {
      const val = paraRef(intervalo);
      if (definicao) val.definicao = definicao;
      resultado.set(pedido.chave, val);
    }
    return resultado;
  }

  function paraRef(intervalo) {
    const { folha, ref } = Enderecos.separar(intervalo.address);
    const ret = Enderecos.analisarRef(ref);
    const val = { t: "ref", folha, ref: Enderecos.formatarRef(ret), ...ret, celulas: null };
    if (intervalo.cellCount <= LIMITE_VALORES) {
      val.celulas = intervalo.values.map((linha, i) => linha.map((v, j) => celula(v, intervalo.valueTypes[i][j])));
    }
    if (intervalo.cellCount === 1) val.texto = intervalo.text[0][0];
    return val;
  }

  function celula(valor, tipo) {
    switch (tipo) {
      case "Double":
      case "Integer": return { t: "num", v: valor };
      case "String": return { t: "txt", v: valor };
      case "Boolean": return { t: "bool", v: valor };
      case "Error": return { t: "erro", v: String(valor) };
      case "Empty": return { t: "vazio" };
      default: return { t: "nd", motivo: "tipo de valor" };
    }
  }

  function valorDeNome(nome) {
    switch (nome.type) {
      case "Integer":
      case "Double": return { t: "num", v: Number(nome.value) };
      case "String": return { t: "txt", v: String(nome.value) };
      case "Boolean": return { t: "bool", v: Boolean(nome.value) };
      case "Error": return { t: "erro", v: String(nome.value) };
      default: return { t: "nd", motivo: "nome sem intervalo" };
    }
  }

  // ---------- Funções do Excel ----------

  async function chamar(chamadas) {
    try {
      return await Excel.run((context) => executar(context, chamadas));
    } catch (e) {
      const resultados = [];
      for (const chamada of chamadas) {
        try {
          resultados.push((await Excel.run((context) => executar(context, [chamada])))[0]);
        } catch (e2) {
          resultados.push({ t: "nd", motivo: "o Excel não calculou" });
        }
      }
      return resultados;
    }
  }

  async function executar(context, chamadas) {
    const funcoes = context.workbook.functions;
    const resultados = chamadas.map(({ metodo: m, args }) => {
      const r = funcoes[m](...args.map((a) => argumento(context, a)));
      r.load("value,error");
      return r;
    });
    await context.sync();
    return resultados.map(resultadoParaValor);
  }

  function argumento(context, a) {
    if (a && typeof a === "object" && a.ref) return context.workbook.worksheets.getItem(a.folha).getRange(a.endereco);
    return a;
  }

  function resultadoParaValor(r) {
    if (r.error) return { t: "erro", v: r.error };
    let v = r.value;
    if (Array.isArray(v)) {
      if (v.length !== 1 || v[0].length !== 1) return { t: "nd", motivo: "resultado em matriz" };
      v = v[0][0];
    }
    if (typeof v === "number") return { t: "num", v };
    if (typeof v === "boolean") return { t: "bool", v };
    if (typeof v === "string") return { t: "txt", v };
    return { t: "nd", motivo: "resultado desconhecido" };
  }

  // ---------- Seleção ----------

  // { ativa, enderecos (áreas com aba), celulas, folha }
  async function lerSelecao() {
    return Excel.run(async (context) => {
      const ativa = context.workbook.getActiveCell();
      ativa.load("address");
      const areas = context.workbook.getSelectedRanges();
      areas.load("address,cellCount");
      await context.sync();
      const enderecos = Enderecos.dividirAreas(areas.address).map(Enderecos.semCifrao);
      return {
        ativa: Enderecos.semCifrao(ativa.address),
        enderecos,
        celulas: areas.cellCount,
        folha: Enderecos.separar(enderecos[0]).folha,
      };
    });
  }

  // Seleciona uma ou várias áreas. Áreas em abas diferentes: fica na aba ativa, se ela estiver
  // entre elas; senão vai para a aba da primeira área.
  async function selecionar(enderecos) {
    const areas = [].concat(enderecos).flatMap((e) => Enderecos.dividirAreas(e));
    if (!areas.length) return false;
    try {
      await Excel.run(async (context) => {
        const ativa = context.workbook.worksheets.getActiveWorksheet();
        ativa.load("name");
        await context.sync();
        const porFolha = new Map();
        for (const area of areas) {
          const { folha, ref } = Enderecos.separar(area);
          const nome = folha || ativa.name;
          if (!porFolha.has(nome)) porFolha.set(nome, []);
          porFolha.get(nome).push(ref);
        }
        const nome = porFolha.has(ativa.name) ? ativa.name : porFolha.keys().next().value;
        const refs = porFolha.get(nome).slice(0, MAX_AREAS_SELECAO);
        const aba = context.workbook.worksheets.getItem(nome);
        aba.activate();
        const alvo = refs.length === 1 ? aba.getRange(refs[0]) : aba.getRanges(refs.join(","));
        if (typeof alvo.select === "function") alvo.select();
        else aba.getRange(refs[0]).select();
        await context.sync();
      });
      return true;
    } catch (e) {
      return false;
    }
  }

  // ---------- Célula da janela ----------

  async function montarDados(endereco) {
    const celulaInfo = await Excel.run(async (context) => {
      const wb = context.workbook;
      let alvo;
      if (endereco) {
        const { folha, ref } = Enderecos.separar(endereco);
        alvo = wb.worksheets.getItem(folha).getRange(ref).getCell(0, 0);
      } else {
        alvo = wb.getActiveCell();
      }
      alvo.load("address,formulas,formulasLocal,text");
      let formatoNumero = null;
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.11")) {
        formatoNumero = context.application.cultureInfo.numberFormat;
        formatoNumero.load("numberDecimalSeparator");
      }
      await context.sync();
      const { folha, ref } = Enderecos.separar(alvo.address);
      const idioma = Office.context.displayLanguage || "pt-BR";
      const decimal = formatoNumero ? formatoNumero.numberDecimalSeparator : /^pt/i.test(idioma) ? "," : ".";
      return {
        folha,
        ref: Enderecos.formatarRef(Enderecos.analisarRef(ref)),
        texto: alvo.text[0][0],
        formula: alvo.formulas[0][0],
        formulaLocal: alvo.formulasLocal[0][0],
        // O Excel usa ";" entre argumentos quando a vírgula é o separador decimal.
        separadores: { decimal, lista: decimal === "," ? ";" : "," },
        idioma,
        validacao: await lerValidacao(context, alvo),
      };
    });
    return Avaliador.montarDados(celulaInfo, { ler, chamar, metodo });
  }

  // Regra de validação de dados da célula, ou null. Falha aqui não impede a janela.
  async function lerValidacao(context, alvo) {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.8")) return null;
    try {
      const dv = alvo.dataValidation;
      dv.load("type");
      await context.sync();
      if (!dv.type || ["None", "Inconsistent", "MixedCriteria"].includes(dv.type)) return null;
      dv.load("rule");
      await context.sync();
      return { tipo: dv.type, regra: dv.rule };
    } catch (e) {
      return null;
    }
  }

  async function gravarFormula(endereco, texto) {
    await Excel.run(async (context) => {
      const { folha, ref } = Enderecos.separar(endereco);
      context.workbook.worksheets.getItem(folha).getRange(ref).formulasLocal = [[texto]];
      await context.sync();
    });
  }

  return { montarDados, gravarFormula, lerSelecao, selecionar };
})();
