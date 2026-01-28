let
    Fonte = Folder.Files("C:\...\Gestão a vista - NPS"),

    ArquivoCE = Fonte{
        [
            #"Folder Path"="C:\...\Relatório\",
            Name="RelatórioCE.xlsx"
        ]
    }[Content],

    WB = Excel.Workbook(ArquivoCE, null, true),
    Sheet1 = WB{[Item="Sheet 1", Kind="Sheet"]}[Data],

    LinhasSuperioresRemovidas = Table.Skip(Sheet1, 5),
    CabecalhosPromovidos = Table.PromoteHeaders(LinhasSuperioresRemovidas),

    TipoAlterado = Table.TransformColumnTypes(
        CabecalhosPromovidos,
        {
            {"Data", type datetime},
            {"Protocolo", Int64.Type},
            {"Nome", type text},
            {"Contato", Int64.Type},
            {"Situação", type text},
            {"!9.13.3 [PRO CA PL] NPS Pergunta 4 - Projeto Valor [ Pergunta Livre ]", type any}
        }
    ),

    LinhasInferioresRemovidas     = Table.RemoveLastN(TipoAlterado,1),
    ColunasRenomeadas             = Table.RenameColumns(LinhasInferioresRemovidas,
        {{"!9.13.3 [PRO CA PL] NPS Pergunta 4 - Projeto Valor [ Pergunta Livre ]","Pergunta 4 - Pergunta Livre"}}),

    TipoAlterado1 = Table.TransformColumnTypes(
        ColunasRenomeadas,
        {{"Pergunta 4 - Pergunta Livre", type text}}
    ),

    NulosComoVazio = Table.ReplaceValue(
        TipoAlterado1,
        null,
        "",
        Replacer.ReplaceValue,
        {"Pergunta 4 - Pergunta Livre"}
    ),

    MapaTipado = Table.TransformColumnTypes(
        MapaPalavrasChave,
        {{"Padrao", type text}, {"Categoria", type text}, {"Prioridade", Int64.Type}}
    ),

    CategoriaFeedback = Table.AddColumn(
        NulosComoVazio,
        "CategoriaFeedback",
        each fxClassificarPorPadroes([#"Pergunta 4 - Pergunta Livre"], MapaTipado),
        type text
    )

in
    CategoriaFeedback