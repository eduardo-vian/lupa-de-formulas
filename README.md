# Lupa de Fórmulas

Complemento do Excel (Office.js) para auditoria de fórmulas. Funciona no Mac, no Windows e no Excel da web.

**Situação:** versão 0.2.0.

- **Ctrl+Q** abre a janela da fórmula: árvore com Elemento, Info, Valor e Local, destaque do trecho na fórmula, edição da fórmula e navegação pelas referências (duplo clique e Voltar).
- O painel (botão **Abrir Lupa** ou Ctrl+Alt+Shift+F) mostra precedentes e dependentes diretos e um diagnóstico do Excel em uso.

## Estrutura

| Caminho | O que é |
|---|---|
| `docs/` | O que é publicado na hospedagem: painel, janela, atalhos, ícones, manifesto gerado e planilha de teste |
| `docs/taskpane.js` | Painel e controle da janela (roda no runtime compartilhado) |
| `docs/janela.*` | Janela da fórmula; só exibe, quem calcula é o painel |
| `docs/formula.js` | Tokens e árvore sintática da fórmula (testado em node) |
| `docs/avaliador.js` | Valor de cada parte e árvore exibida; recebe um adaptador (testado em node) |
| `docs/argumentos.js` | Nomes dos argumentos em português |
| `docs/excel.js` | Adaptador real: lê intervalos e chama funções do Excel |
| `docs/enderecos.js` | Funções puras de endereço (testadas em node) |
| `docs/demo.json` | Exemplo para ver a janela fora do Excel: `janela.html?demo` |
| `docs/shortcuts.json` | Atalhos de teclado |
| `manifest.template.xml` | Modelo do manifesto; `{{BASE_URL}}` vira o endereço de hospedagem |
| `tools/` | Scripts: ícones, manifesto, planilha de teste, testes |

## Comandos

```bash
python tools/gerar_icones.py
python tools/gerar_planilha_teste.py
python tools/gerar_manifesto.py https://ENDERECO-DA-HOSPEDAGEM
node tools/testar_enderecos.js
node tools/testar_formula.js
node tools/gerar_demo.js
npx office-addin-manifest validate docs/manifest.xml
```

Ao publicar uma versão nova, troque o número em `taskpane.js` (`VERSAO`) e nos `?v=` de `taskpane.html` e `janela.html`, para o Excel não usar a versão guardada em cache.
