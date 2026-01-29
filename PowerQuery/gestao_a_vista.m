let
    // ====== 1) Fonte: pasta e arquivo ======
    Fonte = Folder.Files("C:\Users\uva004714\UNIMED VALE DO ACO COOPERATIVA DE TRABALHO MEDICO\BRENO PATROCINIO SANTOS SILVA - MARCO TULIO GOMES CAMPOS - Pesquisa NPS\Gestão a vista - NPS"),

    ArquivoCE = Fonte{
        [
            #"Folder Path"="C:\Users\uva004714\UNIMED VALE DO ACO COOPERATIVA DE TRABALHO MEDICO\BRENO PATROCINIO SANTOS SILVA - MARCO TULIO GOMES CAMPOS - Pesquisa NPS\Gestão a vista - NPS\Relatorios\Centro de Especialidades\Relatório\",
            Name="RelatórioCE.xlsx"
        ]
    }[Content],

    // ====== 2) Abrir o Excel ======
    WB = Excel.Workbook(ArquivoCE, null, true),
    Sheet1 = WB{[Item="Sheet 1", Kind="Sheet"]}[Data],

    // ====== 3) Limpeza inicial ======
    LinhasSuperioresRemovidas = Table.Skip(Sheet1, 6),
    CabecalhosPromovidos = Table.PromoteHeaders(LinhasSuperioresRemovidas, [PromoteAllScalars=true]),

    // ====== 4) Tipagem ======
    TipoAlterado = Table.TransformColumnTypes(CabecalhosPromovidos, {
        {"Data", type datetime},
        {"Protocolo", Int64.Type},
        {"Nome", type text},
        {"Contato", Int64.Type},
        {"Fila", type text},
        {"Agente", type text},
        {"Classificação", type text},
        {"Direção", type text},
        {"Status", type text},
        {"Situação", type text},
        {"!9.13.0.0 [PRO CA ME] - NPS Pergunta 1 - Projeto Valor (Atendimento Médico)", type text},
        {"!9.13.0.1 [PRO CA ME] - NPS Pergunta 2 - Projeto Valor (Local Do Atendimento)", type text},
        {"!9.13.0.2 [PRO CA ME] - NPS Pergunta 3 - Projeto Valor (Recomendaria)", type text},
        {"!9.13.3 [PRO CA PL] NPS Pergunta 4 - Projeto Valor [ Pergunta Livre ]", type any},
        {"Column17", Int64.Type},
        {"Column16", Int64.Type}
    }),

    LinhasInferioresRemovidas = Table.RemoveLastN(TipoAlterado, 1),
    ColunasRemovidas = Table.RemoveColumns(LinhasInferioresRemovidas, {"Fila", "Agente", "Classificação", "Column17", "Column16"}),

    // ====== 5) Renomear perguntas ======
    ColunasRenomeadas = Table.RenameColumns(ColunasRemovidas, {
        {"!9.13.0.0 [PRO CA ME] - NPS Pergunta 1 - Projeto Valor (Atendimento Médico)", "Pergunta 1 - Atendimento Médico"},
        {"!9.13.0.1 [PRO CA ME] - NPS Pergunta 2 - Projeto Valor (Local Do Atendimento)", "Pergunta 2 - Local Do Atendimento"},
        {"!9.13.0.2 [PRO CA ME] - NPS Pergunta 3 - Projeto Valor (Recomendaria)", "Pergunta 3 - Recomendaria"},
        {"!9.13.3 [PRO CA PL] NPS Pergunta 4 - Projeto Valor [ Pergunta Livre ]", "Pergunta 4 - Pergunta Livre"}
    }),


    // 6) Garante texto e nulos na pergunta livre
    TipoAlterado1 = Table.TransformColumnTypes(ColunasRenomeadas, {{"Pergunta 4 - Pergunta Livre", type text}}),
    NulosComoVazio = Table.ReplaceValue(
        TipoAlterado1, null, "", Replacer.ReplaceValue, {"Pergunta 4 - Pergunta Livre"}
    ),

    #"Texto em Minúscula" = Table.TransformColumns(NulosComoVazio,{{"Pergunta 4 - Pergunta Livre", Text.Lower, type text}}),

    // 7) (Importante) Blinda os tipos do mapa vindo de outra query/tabela
    MapaTipado = Table.TransformColumnTypes(
        MapaPalavrasChave,
        {{"Padrao", type text}, {"Categoria", type text}, {"Prioridade", Int64.Type}}
    ),

    // 8) adiciona a coluna com a função
    CategoriaFeedback = Table.AddColumn(
        #"Texto em Minúscula",
        "CategoriaFeedback",
        each fxClassificarPorPadroes([#"Pergunta 4 - Pergunta Livre"], MapaTipado),
        type text
    ),

    // Divide a coluna Omnitag em uma lista com 3 partes
    ColunasSeparadas = Table.SplitColumn(
        #"CategoriaFeedback",
        "Omnitag",
        Splitter.SplitTextByDelimiter(" - ", QuoteStyle.Csv),
        {"Cidade", "Medico", "Cooperado SN"}
    )

in
    ColunasSeparadas