# 📊 Automação de Consulta e Validação de CNPJ no Excel

![Status do Projeto](https://img.shields.io/badge/Status-Concluído-brightgreen)
![Excel](https://img.shields.io/badge/Tools-Excel%20|%20Power%20Query-green)
![API](https://img.shields.io/badge/API-BrasilAPI-blue)

## 🎯 Objetivo
Este projeto foi desenvolvido para resolver um problema comum em processos de compliance e cadastro de clientes: a divergência entre dados digitados manualmente e os dados oficiais da Receita Federal. 

O objetivo principal é automatizar a consulta de dados via API, formatar endereços e realizar uma **comparação inteligente** que aceita abreviações, garantindo a integridade da base de dados de forma rápida e visual.

## 🚀 Funcionalidades
- **Integração em Tempo Real:** Conecta o Excel à [BrasilAPI](https://brasilapi.com.br/) via Power Query para buscar dados cadastrais de qualquer CNPJ.
- **Tratamento de Dados:** Limpeza automática de CNPJs (remove pontos, traços e espaços) e formatação de endereço (Logradouro + Número).
- **Lógica de Comparação Parcial (Fuzzy Match):** Algoritmo em fórmulas que valida se a Razão Social e o Endereço "batem", mesmo que um esteja abreviado (ex: "Microsoft Enterprises" vs "Microsoft Intrp").
- **Validação de Situação Cadastral:** Verifica automaticamente se a empresa está com status "ATIVA".
- **Painel Visual:** Coluna de veredito final com ícones (✅/❌) para identificação imediata de erros.

## 🛠️ Tecnologias Utilizadas
- **Microsoft Excel:** Interface do usuário e lógica de fórmulas.
- **Power Query (M Language):** Consumo da API JSON, tratamento e transformação dos dados.
- **BrasilAPI:** Fonte de dados pública da Receita Federal.
- **Fórmulas Avançadas:** `XLOOKUP`, `SEARCH`, `ISNUMBER` e lógica `IF` aninhada.

## 📖 Como Funciona (Tutorial)

1. **Preenchimento:** Insira o número do CNPJ na coluna **B** da aba `Sheet1`.
2. **Atualização:** Vá até a guia **Dados** e clique no botão **Atualizar Tudo**.
3. **Processamento:**
   - O Power Query limpa o CNPJ e faz a requisição para a API.
   - Os dados oficiais retornam nas colunas E, F e G.
4. **Validação:**
   - As colunas **Check** comparam os primeiros caracteres dos seus dados com os da API (ignorando maiúsculas/minúsculas).
   - A coluna **Validação Final** exibe ✅ se Razão, Endereço e Situação estiverem corretos.

## 💻 Visual do Código (Power Query)
O coração do projeto está no script M que limpa e consulta a API:

```powerquery
let
    Fonte = Excel.CurrentWorkbook(){[Name="SuperTable"]}[Content],
    LimpezaCNPJ = Table.TransformColumns(Fonte, {{"CNPJ", each Text.Select(Text.From(_), {"0".."9"})}}),
    ConsultaAPI = Table.AddColumn(LimpezaCNPJ, "Dados", each 
        try Json.Document(Web.Contents("[https://brasilapi.com.br/api/cnpj/v1/](https://brasilapi.com.br/api/cnpj/v1/)" & [CNPJ])) otherwise null
    ),
    // ... expansão e formatação de dados ...
in
    Final
