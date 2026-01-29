# Fluxo de Transformação (ETL) – Gestão à Vista NPS

---

## 1. Importação
- Leitura de pasta via `Folder.Files`
- Seleção do arquivo `RelatórioCE.xlsx`

## 2. Transformações Iniciais
- Remoção de primeiras linhas
- Promoção de cabeçalhos
- Tipagem de dados

## 3. Normalização
- Garantir que "Pergunta 4 - Pergunta Livre" está como texto
- Substituir valores nulos por string vazia

## 4. Preparar mapa
- Garantir tipagem de `Padrao`, `Categoria`, `Prioridade`

## 5. Classificação
A consulta aplica:
