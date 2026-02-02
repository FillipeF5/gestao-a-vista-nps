# Fluxo de Transformação (ETL) – Gestão à Vista NPS

Este documento descreve o fluxo completo de tratamento da base NPS dentro da consulta principal **Gestão à Vista**, incluindo importação, limpeza, tipagem, normalização e classificação automática baseada em padrões.

---

## 1. Importação
- Leitura de arquivos via `Folder.Files`, apontando para a pasta:  
  `...\Gestão a vista - NPS\Relatorios\Centro de Especialidades\Relatório\`
- A consulta busca especificamente o arquivo **RelatórioCE.xlsx**
- O arquivo é acessado por índice usando o nome e caminho corretos
- A planilha utilizada é `"Sheet 1"` (ajustável caso mude no futuro)
- Transformação inicial ocorre no objeto retornado `Excel.Workbook`

---

## 2. Transformações Iniciais
- Remoção das cinco primeiras linhas (cabeçalho técnico da planilha)
- Promoção de cabeçalhos reais para a primeira linha
- Tipagem inicial das colunas relevantes (data, protocolo, texto, números)
- Remoção das últimas linhas (rodapés ou totalizadores)
- Remoção de colunas que não agregam ao modelo NPS:
  - `Fila`, `Agente`, `Classificação`, `Column15`, `Column16`
- Renomeação das perguntas NPS para nomes mais descritivos:
  - Pergunta 1 – Atendimento Médico  
  - Pergunta 2 – Local do Atendimento  
  - Pergunta 3 – Recomendaria  
  - Pergunta 4 – Pergunta Livre

---

## 3. Normalização
- A coluna **"Pergunta 4 - Pergunta Livre"**, que contém texto livre do cliente, nem sempre vem em tipo correto.
- Para evitar erros de função e conversão, a consulta:
  - força o tipo da coluna para **type text**
  - substitui valores **null** por `""` (string vazia)
- Esses passos eliminam erros do tipo:
  - *Expression.Error: Não conseguimos converter o valor X em tipo Text*

---

## 4. Preparar mapa de palavras‑chave
- A consulta referenciada `MapaPalavrasChave` é tipada internamente para garantir:
  - `Padrao` → text  
  - `Categoria` → text  
  - `Prioridade` → number  
- Isso assegura que a função `fxClassificarPorPadroes` receba um mapa limpo, evitando falhas em:
  - comparação de textos
  - ordenação por prioridade
  - conversões implícitas de tipo

---

## 5. Classificação
A etapa mais importante: é aqui que o texto livre do cliente é transformado em uma categoria padronizada.

### ✔ O que é feito
- A consulta adiciona uma nova coluna chamada **`CategoriaFeedback`**
- Cada linha é enviada para a função:  