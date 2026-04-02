# 📋 Guia de Integração — Inventário de Servidores no SharePoint

## O que foi preparado

| Arquivo | Descrição |
|---|---|
| `index_sharepoint.html` | Página principal com imagens embutidas (base64), pronta para upload |

> **Atenção:** O HTML foi adaptado para ser **autossuficiente** — as imagens (`img_1.png` e `img_2.png`) estão embutidas diretamente no arquivo. As dependências externas de CSS (`css/base.css`) e JavaScript (`js/main.js`) precisam ser tratadas conforme a Opção A ou B abaixo.

---

## ⚠️ Dependências ainda necessárias

O sistema original carrega dois arquivos externos que **não estão incluídos no HTML**:

- `css/base.css` — estilos visuais do sistema
- `js/main.js` — toda a lógica de negócio (inventário, fitas, dashboard, etc.)

Você precisa escolher uma das opções abaixo para integrar corretamente.

---

## Opção A — SharePoint com Biblioteca de Documentos (Recomendado)

### Estrutura de pastas no SharePoint

```
Sites/SuaSite/
└── Documentos/
    └── InventarioServidores/
        ├── index_sharepoint.html   ← arquivo principal
        ├── css/
        │   └── base.css
        ├── js/
        │   └── main.js
        ├── img_1.png
        └── img_2.png
```

### Passo a passo

1. Acesse a **Biblioteca de Documentos** do site SharePoint.
2. Crie a pasta `InventarioServidores` e suba todos os arquivos mantendo a estrutura acima.
3. Abra o `index_sharepoint.html` diretamente pelo navegador via URL da biblioteca.
4. Copie a URL gerada (ex: `https://empresa.sharepoint.com/sites/.../index_sharepoint.html`).

---

## Opção B — Web Part "Incorporar" (Embed)

Use esta opção para exibir o sistema **dentro de uma página do SharePoint**.

1. Edite uma página do SharePoint.
2. Clique em **"+"** para adicionar uma Web Part.
3. Selecione **"Incorporar"** (Embed).
4. Cole a URL do arquivo HTML da biblioteca (conforme Opção A).
5. Ajuste a altura do iframe (recomendado: **800px ou mais**).
6. Salve e publique a página.

---

## Opção C — HTML inline (apenas para testes simples)

> Não recomendado para produção. O SharePoint limita o HTML inline e pode bloquear scripts.

1. Edite uma página SharePoint.
2. Adicione a Web Part **"Inserir código"** ou **"Script Editor"** (disponível no SharePoint clássico).
3. Cole o conteúdo completo do `index_sharepoint.html`.

---

## 🔒 Permissões recomendadas

- Apenas usuários internos da Claro devem ter acesso à biblioteca.
- Configure as permissões na biblioteca → **Gerenciar Acesso** → restrinja ao grupo correto.

---

## 🌐 Fontes externas utilizadas

O sistema carrega recursos externos que precisam estar liberados no firewall/proxy:

| Recurso | URL |
|---|---|
| Fontes Claro (Mondrian) | `https://mondrian.claro.com.br/fonts/` |
| Biblioteca XLSX | `https://cdnjs.cloudflare.com/ajax/libs/xlsx/` |
| Biblioteca jsPDF | `https://cdnjs.cloudflare.com/ajax/libs/jspdf/` |
| Biblioteca Chart.js | `https://cdnjs.cloudflare.com/ajax/libs/Chart.js/` |

---

## 💡 Dica: tornar o HTML 100% independente

Se precisar de um único arquivo sem dependências externas, solicite o **build completo** com os arquivos `base.css` e `main.js` fornecidos. Eles poderão ser injetados inline no HTML, eliminando qualquer dependência de rede.

---

*Preparado para integração SharePoint — Claro Empresas © 2026*
