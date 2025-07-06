# 📘 GCA - Gerenciador de Controle Acadêmico

[![Status](https://img.shields.io/badge/status-em%20desenvolvimento-yellow)](https://github.com)
[![VBA](https://img.shields.io/badge/plataforma-Excel%20VBA-blue)](https://learn.microsoft.com/pt-br/office/vba/api/overview/excel)
[![Compatibilidade](https://img.shields.io/badge/Excel-2016%20ou%20superior-green)](https://www.microsoft.com/)

Sistema desenvolvido em **VBA para Excel**, com Ribbon customizada e estrutura modular, voltado para facilitar o gerenciamento e configuração de rotinas acadêmicas de forma automatizada, intuitiva e organizada.

---

## 🧠 Objetivo

Simplificar e agilizar o trabalho de gestores acadêmicos por meio de uma **interface personalizada** no Excel, com botões organizados por categorias e ações programadas de forma modular.

---

## 🔧 Funcionalidades

✅ Ribbon personalizada com abas:
- Logon / Menu Inicial
- Configurações Administrativas
- Infraestrutura
- Acadêmicas
- Educacionais
- Eventos
- Edição de Cronograma

✅ Mapeamento de botões com:
- `Enum IdentificadoresDeBotoesDaRibbon`
- `Dictionary(Of TEnum, Action)` para execução
- Ícones únicos via recurso `.resx`

✅ Ações implementadas:
- Navegação entre telas com `Invalidate`
- Acesso seguro a `Worksheet`, `Range` e `ListObject`
- Mensagens automáticas de "Em breve..." para futuras features

---

## 🗂️ Estrutura do Projeto

```
📦 GCA___Turma_000._0000._0000
├── 📄 Rbbn_GCA.vb                 ' Lógica da Ribbon do Excel
├── 📄 Md_ExcelHelper.vb          ' Acesso seguro a Planilhas e Ranges
├── 📄 Md_VariaveisGlobais.vb     ' Enums, Dicionários e RibbonUI
├── 📄 Md_FuncoesGlobais.vb       ' Navegação e mensagens gerais
├── 📁 Resources/                 ' Ícones de botões (.resx)
│   ├── icn_*.png
│   └── ...
└── 📄 Rbbn_GCA.xml               ' Definição visual da Ribbon
```

---

## 🚀 Como Executar

1. Abrir o Excel com o projeto do GCA já vinculado.
2. A Ribbon personalizada será carregada automaticamente.
3. Clique nos botões para navegar entre seções ou executar ações.
4. Itens não implementados exibirão `Em breve...`.

---

## ⚙️ Requisitos

- Excel 2016 ou superior (recomendado 64 bits)
- VBA habilitado
- Referências:
  - `Microsoft Office Object Library`
  - `Microsoft Excel Object Library`
- Recursos XML embutidos via `GetManifestResourceStream`

---

## ✍️ Exemplos de Código

```vbnet
' Mapeamento da ação do botão:
MapeamentoDeAcoesDaRibbon("Bttn_IniciarSessao") = IdentificadoresDeBotoesDaRibbon.IniciarSessao

' Execução via delegate:
ExecutorDeAcoesDaRibbon(IdentificadoresDeBotoesDaRibbon.IniciarSessao).Invoke()

' Acesso seguro a intervalo:
Dim rng As Range = ObterIntervalo(PL_MENU, "RNG_MenuPrincipal")
```

---

## 👨‍💻 Autor

**H4rzel**  
Desenvolvedor responsável pelo design, lógica e implementação do GCA.  
📧 *(adicione seu e-mail aqui, se desejar)*

