# Gestão à Vista – NPS (Power BI / Power Query)

Projeto para classificar automaticamente respostas textuais de clientes no processo de NPS da Unimed Vale do Aço.

## Funcionalidades
- Classificação automática por padrões e raízes de palavras
- Prioridade entre categorias
- Remoção de acentos
- Coluna final “CategoriaFeedback”
- Arquitetura modular e reutilizável

## Estrutura do Projeto
- `fxRemoveAcentos.m` — remove acentos
- `fxClassificarPorPadroes.m` — classifica textos
- `MapaPalavrasChave.xlsx` — padrões, categorias e prioridade
- `gestao_a_vista.m` — consulta principal

## Como Reutilizar
1. Importe as funções no Power Query
2. Crie ou edite o `MapaPalavrasChave`
3. Aplique `fxClassificarPorPadroes` na sua coluna de texto
4. Adicione a coluna `CategoriaFeedback`

## Autor
FILLIPE FREITAS MONTEIRO