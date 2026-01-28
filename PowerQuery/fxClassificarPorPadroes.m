// fxClassificarPorPadroes(texto as nullable any, mapa as table) as nullable text
(texto as nullable any, mapa as table) as nullable text =>
let
    // 1) Normaliza o texto de entrada (converte qualquer tipo para texto)
    TextoSeguro = try Text.From(texto) otherwise "",
    TextoSeguroTrim = Text.Trim(TextoSeguro),

    // *** Early return: se estiver vazio (nulo ou só espaços), retorna null ***
    ResultadoSeVazio = if TextoSeguroTrim = "" then null else null,  // placeholder para clareza

    // Se não estiver vazio, segue normalização
    TextoBase   = Text.Lower(TextoSeguroTrim),
    TextoNorm   = fxRemoveAcentos(TextoBase),

    // 2) Normaliza o mapa e aplica prioridade
    MapaOrdenado = Table.Sort(mapa, {{"Prioridade", Order.Ascending}}),

    // Garantir que 'Padrao' é tratado como texto e normalizado
    MapaNorm = Table.TransformColumns(
        MapaOrdenado,
        {{"Padrao", each fxRemoveAcentos(Text.Lower(try Text.From(_) otherwise "")), type text}}
    ),

    // 3) Procura o primeiro padrão encontrado por ordem de prioridade
    Linhas = Table.ToRecords(MapaNorm),
    Correspondencias = List.Select(Linhas, (r as record) => Text.Contains(TextoNorm, r[Padrao])),

    ResultadoNaoVazio =
        if List.Count(Correspondencias) > 0
        then Correspondencias{0}[Categoria]
        else "Outros",

    // 4) Decide resultado final
    ResultadoFinal = if TextoSeguroTrim = "" then null else ResultadoNaoVazio
in
    ResultadoFinal