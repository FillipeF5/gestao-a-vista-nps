# Arquitetura do Projeto – Gestão à Vista NPS

Este projeto segue quatro blocos principais:

---

## 1. Funções M
### 🔹 fxRemoveAcentos
- Remove acentos de qualquer texto
- Utiliza tabela interna de substituições
- Evita uso de `Text.RemoveDiacritics`

### 🔹 fxClassificarPorPadroes
- Classifica texto livre
- Converte qualquer dado para texto (`Text.From`)
- Normaliza com `fxRemoveAcentos`
- Ordena padrões por prioridade
- Retorna a primeira categoria compatível

---

## 2. Tabela de mapeamento – `MapaPalavrasChave`
Contém:
- Padrão (raiz da palavra)
- Categoria final
- Prioridade

Pode ser mantida dentro do Power Query ou carregada via CSV.

---

## 3. Consulta Principal – `Gestão à Vista`
Faz:
- Importação das planilhas NPS
- Tipagem de dados
- Normalização
- Aplicação da função `fxClassificarPorPadroes`
- Criação da coluna `CategoriaFeedback`

---

## 4. Power BI
O resultado é carregado no modelo
- DAX para métricas (NPS, % Elogios, etc.)
- Dashboards