# Lupa de Fórmulas

Complemento do Excel (Office.js) para auditoria de fórmulas. Funciona no Mac, no Windows e no Excel da web.

**Situação:** etapa 0, protótipo de viabilidade (versão 0.1.0). Ele mostra a fórmula da célula ativa, os precedentes e dependentes diretos com navegação, e um diagnóstico do Excel em uso. Abre pelo botão **Abrir Lupa** ou pelo atalho Ctrl+Q.

## Estrutura

| Caminho | O que é |
|---|---|
| `docs/` | O que é publicado na hospedagem: painel, atalhos, ícones, manifesto gerado e planilha de teste |
| `docs/taskpane.js` | Lógica do painel |
| `docs/enderecos.js` | Funções puras de endereço (testadas em node) |
| `docs/shortcuts.json` | Atalhos de teclado |
| `manifest.template.xml` | Modelo do manifesto; `{{BASE_URL}}` vira o endereço de hospedagem |
| `tools/` | Scripts: ícones, manifesto, planilha de teste, testes |

## Comandos

```bash
python tools/gerar_icones.py
python tools/gerar_planilha_teste.py
python tools/gerar_manifesto.py https://ENDERECO-DA-HOSPEDAGEM
node tools/testar_enderecos.js
npx office-addin-manifest validate docs/manifest.xml
```

Ao publicar uma versão nova, troque o número em `taskpane.js` (`VERSAO`) e nos `?v=` do `taskpane.html`, para o Excel não usar a versão guardada em cache.
