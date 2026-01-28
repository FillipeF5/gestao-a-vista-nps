(texto as nullable any, mapa as table) as text =>
let
    // Normaliza texto independentemente do tipo
    TextoSeguro = try Text.From(texto) otherwise "",
    TextoBase   = Text.Lower(TextoSeguro),
    TextoNorm   = fxRemoveAcentos(TextoBase),

    // Prepara o mapa (garante tipos e remove acentos)
    MapaOrdenado = Table.Sort(mapa, {{"Prioridade", Order.Ascending}}),
    MapaNorm = Table.TransformColumns(
        MapaOrdenado,
        {{"Padrao", each fxRemoveAcentos(Text.Lower(try Text.From(_) otherwise "")), type text}}
    ),

    Linhas = Table.ToRecords(MapaNorm),

    Correspondencias =
        List.Select(
            Linhas,
            (r as record) => Text.Contains(TextoNorm, r[Padrao])
        ),

    Resultado =
        if List.Count(Correspondencias) > 0
        then Correspondencias{0}[Categoria]
        else "Outros"
in
    Resultado