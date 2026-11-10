# Lupa de Fórmulas

Complemento do Excel (Office.js) para auditoria de fórmulas, inspirado nas funções do Arixcel Explorer. Funciona no Mac, no Windows e no Excel da web.

**Situação:** versão 0.3.0.

| Atalho (no Mac, Ctrl = Control) | O que faz |
|---|---|
| **Ctrl+Q** numa célula | Janela da fórmula: árvore com Elemento, Info, Valor e Local; ramo escolhido de SE/SES/ESCOLHER/PARÂMETRO em negrito; célula encontrada por PROCV, PROCH, PROCX, ÍNDICE, DESLOC e INDIRETO; células filtradas por SOMASE(S), CONT.SE(S), MÉDIASE(S), MÁXIMOSES, MÍNIMOSES, MÁXIMO e MÍNIMO; nomes definidos e validação de dados na árvore |
| **Ctrl+Q** em várias células | Precedentes diretos de todas, em blocos de mesma fórmula, com cores na planilha |
| **Ctrl+Shift+Q** | Dependentes diretos da seleção, em blocos, com cores |
| **Ctrl+Shift+F** | Fluxo de cálculo: entradas, cálculos, saídas e fórmulas órfãs; com foco, o que entra e sai da área |
| **Ctrl+Shift+Z** | Volta à exploração anterior (até 100), mesmo com a janela fechada |
| **Ctrl+Alt+Shift+F** | Abre o painel |

Dentro da janela: clique seleciona no Excel; Cmd/Ctrl+clique e Shift+clique selecionam vários; clique na fórmula seleciona a parte; duplo clique ou Ctrl+Q explora; Ctrl+Shift+Q rastreia dependentes; F2 edita a parte; Enter aplica a edição ou fecha; Esc fecha e volta; Alt+X fecha mantendo as cores; Home volta ao início.

O painel tem botões para tudo isso, "Remover cores", "Fixar cores" e a troca dos atalhos.

## Estrutura

| Caminho | O que é |
|---|---|
| `docs/` | O que é publicado na hospedagem |
| `docs/taskpane.*` | Painel, atalhos e diagnóstico (roda no runtime compartilhado) |
| `docs/janela-controle.js` | Abre a janela, conversa com ela e guarda o histórico de explorações |
| `docs/janela.*` | Janela (diálogo); só exibe, quem calcula é o painel |
| `docs/formula.js` | Tokens e árvore sintática da fórmula (testado em node) |
| `docs/avaliador.js` | Valor de cada parte, buscas, filtros, nomes e validação; recebe um adaptador (testado em node) |
| `docs/argumentos.js` | Nomes dos argumentos em português |
| `docs/blocos.js` | Células em blocos retangulares e árvore do rastreio (testado em node) |
| `docs/fluxo.js` | Classificação do fluxo de cálculo (testado em node) |
| `docs/excel.js` | Adaptador real: lê intervalos, chama funções do Excel, seleciona e grava |
| `docs/rastreio.js`, `docs/fluxo-excel.js` | Leitura do Excel para o rastreio e para o fluxo |
| `docs/cores.js` | Cores temporárias (formatação condicional marcada) |
| `docs/enderecos.js` | Funções puras de endereço (testadas em node) |
| `docs/demo*.json` | Exemplos para ver a janela fora do Excel: `janela.html?demo`, `?demo=rastreio`, `?demo=fluxo`, `?demo=escolher` |
| `docs/shortcuts.json` | Atalhos de teclado |

## Comandos

```bash
python tools/gerar_icones.py
python tools/gerar_planilha_teste.py
python tools/gerar_manifesto.py https://ENDERECO-DA-HOSPEDAGEM
node tools/testar_enderecos.js
node tools/testar_formula.js
node tools/testar_fluxo.js
node tools/gerar_demo.js
npx office-addin-manifest validate docs/manifest.xml
```

Ao publicar uma versão nova, troque o número em `taskpane.js` (`VERSAO`) e nos `?v=` de `taskpane.html` e `janela.html`, para o Excel não usar a versão guardada em cache.
