/* global Excel, Office, Enderecos, Avaliador */
"use strict";

// Ponte entre o Excel e o avaliador: lê intervalos, chama funções do Excel e grava fórmulas.
const LupaExcel = (() => {
  const LIMITE_VALORES = 20000; // intervalo maior que isso é usado sem carregar os valores
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

  // ---------- Leitura ----------

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

    // 1) Confere se as abas e os nomes existem.
    const conferidos = pedidos.map((pedido) => {
      if (pedido.nome !== undefined) {
        const colecao = pedido.folha ? wb.worksheets.getItem(pedido.folha).names : wb.names;
        const nome = colecao.getItemOrNullObject(pedido.nome);
        nome.load("type,value");
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
      if (nome) {
        if (nome.isNullObject) {
          resultado.set(pedido.chave, { t: "erro", v: "#NAME?" });
          continue;
        }
        if (nome.type !== "Range") {
          resultado.set(pedido.chave, valorDeNome(nome));
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
      intervalos.push({ pedido, intervalo });
    }
    await context.sync();

    // 3) Valores, só dos intervalos que não são grandes demais.
    for (const { intervalo } of intervalos) {
      if (intervalo.cellCount <= LIMITE_VALORES) intervalo.load("values,valueTypes");
      if (intervalo.cellCount === 1) intervalo.load("text");
    }
    if (intervalos.length) await context.sync();
    for (const { pedido, intervalo } of intervalos) resultado.set(pedido.chave, paraRef(intervalo));
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
      };
    });
    return Avaliador.montarDados(celulaInfo, { ler, chamar, metodo });
  }

  async function gravarFormula(endereco, texto) {
    await Excel.run(async (context) => {
      const { folha, ref } = Enderecos.separar(endereco);
      context.workbook.worksheets.getItem(folha).getRange(ref).formulasLocal = [[texto]];
      await context.sync();
    });
  }

  return { montarDados, gravarFormula };
})();
