let
    Fonte = Excel.Workbook(Parâmetro1, null, true),
    #"Sheet 1_Sheet" = Fonte{[Item="Sheet 1",Kind="Sheet"]}[Data],

    // Limpeza inicial 
    LinhasSuperioresRemovidas = Table.Skip(#"Sheet 1_Sheet", 6),
    CabecalhosPromovidos = Table.PromoteHeaders(LinhasSuperioresRemovidas, [PromoteAllScalars=true]),

    // Tipagem 
    TipoAlterado = Table.TransformColumnTypes(CabecalhosPromovidos, {
        {"Data", type date},
        {"Protocolo", Int64.Type},
        {"Nome", type text},
        {"Contato", type text},
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

    // Renomear perguntas
    ColunasRenomeadas = Table.RenameColumns(ColunasRemovidas, {
        {"!9.13.0.0 [PRO CA ME] - NPS Pergunta 1 - Projeto Valor (Atendimento Médico)", "Pergunta 1 - Atendimento Médico"},
        {"!9.13.0.1 [PRO CA ME] - NPS Pergunta 2 - Projeto Valor (Local Do Atendimento)", "Pergunta 2 - Local Do Atendimento"},
        {"!9.13.0.2 [PRO CA ME] - NPS Pergunta 3 - Projeto Valor (Recomendaria)", "Pergunta 3 - Recomendaria"},
        {"!9.13.3 [PRO CA PL] NPS Pergunta 4 - Projeto Valor [ Pergunta Livre ]", "Pergunta 4 - Pergunta Livre"}
    }),

    // Garante texto e nulos na pergunta livre
    TipoAlterado1 = Table.TransformColumnTypes(ColunasRenomeadas, {{"Pergunta 4 - Pergunta Livre", type text}}),
    NulosComoVazio = Table.ReplaceValue(
        TipoAlterado1, null, "", Replacer.ReplaceValue, {"Pergunta 4 - Pergunta Livre"}
    ),

    #"Texto em Minúscula" = Table.TransformColumns(NulosComoVazio,{{"Pergunta 4 - Pergunta Livre", Text.Lower, type text}}),

    #"CountRespostas" = Table.AddColumn(
        #"Texto em Minúscula", 
        "countRespostas",
        each if [#"Pergunta 4 - Pergunta Livre"] = "" then "Sem resposta" else "Respondido", type text
    ),

    // (Importante) Blinda os tipos do mapa vindo de outra query/tabela
    MapaTipado = Table.TransformColumnTypes(
        MapaPalavrasChave,
        {{"Padrao", type text}, {"Categoria", type text}, {"Prioridade", Int64.Type}}
    ),

    // Adiciona a coluna com a função
    CategoriaFeedback = Table.AddColumn(
        #"CountRespostas",
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
    ),
    #"SubstituicoesEmTodas" =
    let
    //Colunas onde as substituições serão aplicadas
    colunas = {
        "Pergunta 1 - Atendimento Médico",
        "Pergunta 2 - Local Do Atendimento",
        "Pergunta 3 - Recomendaria"
    },

    //Normaliza: garante texto, trim e normaliza hífens em todas as colunas alvo
    Normalizadas = Table.TransformColumns(
        ColunasSeparadas,
        List.Transform(
            colunas,
            each {_, (v) => Text.Trim(Text.Replace(Text.Replace(Text.Replace(Text.From(v & ""), "–","-"), "—","-"), "−","-")), type text}
        )
    ),

    //Substitui nulos e strings vazias por "0" (IMPORTANTE: null sem aspas)
    NulosParaZero = Table.ReplaceValue(Normalizadas, null, "0", Replacer.ReplaceValue, colunas),
    VaziosParaZero = Table.ReplaceValue(NulosParaZero, "", "0", Replacer.ReplaceValue, colunas),

    //Lista de pares {de -> para} (strings corrigidas — veja as quantidades de estrelas)
    substituicoes = {
        {"10 - ⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐", "10"},
        {"10-⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐", "10"},
        {"10 -⭐⭐⭐⭐⭐⭐⭐⭐⭐⭐", "10"},
        {"Nota 10", "10"},
        {"9 - ⭐⭐⭐⭐⭐⭐⭐⭐⭐", "9"},
        {"9-⭐⭐⭐⭐⭐⭐⭐⭐⭐", "9"},
        {"8 - ⭐⭐⭐⭐⭐⭐⭐⭐", "8"},
        {"8-⭐⭐⭐⭐⭐⭐⭐⭐", "8"},
        {"7 - ⭐⭐⭐⭐⭐⭐⭐", "7"},
        {"6 - ⭐⭐⭐⭐⭐⭐", "6"},
        {"5 - ⭐⭐⭐⭐⭐", "5"},
        {"4 - ⭐⭐⭐⭐", "4"},
        {"3 - ⭐⭐⭐", "3"},
        {"2 - ⭐⭐", "2"},
        {"1 - ⭐", "1"}
    },

    //Aplica todas as substituições em todas as colunas (uma iteração para cada par)
    Resultado =
        List.Accumulate(
            substituicoes,
            VaziosParaZero,
            (estado, par) =>
                Table.ReplaceValue(
                    estado,
                    par{0},
                    par{1},
                    Replacer.ReplaceText,
                    colunas
                )
        )
    in
    Resultado,
    #"Tipo Alterado1" = Table.TransformColumnTypes(SubstituicoesEmTodas,{
        {"Pergunta 1 - Atendimento Médico", Int64.Type},
        {"Pergunta 2 - Local Do Atendimento", Int64.Type},
        {"Pergunta 3 - Recomendaria", Int64.Type}
    }),
    
    #"Erros Removidos1" = Table.RemoveRowsWithErrors(#"Tipo Alterado1", {
        "Pergunta 1 - Atendimento Médico", 
        "Pergunta 2 - Local Do Atendimento", 
        "Pergunta 3 - Recomendaria"
    }),

    #"Média de avaliação do cliente" = 
    Table.AddColumn(#"Erros Removidos1", "media_da_avaliacao_do_cliente", each 
        List.Average(
            {[#"Pergunta 1 - Atendimento Médico"],
            [#"Pergunta 2 - Local Do Atendimento"],
            [#"Pergunta 3 - Recomendaria"]}
        )
    ),

    #"Classificação da avaliação" = Table.AddColumn(#"Média de avaliação do cliente", "classificacao", 
    each if [#"Pergunta 3 - Recomendaria"] <= 6 then "Detrator"
    else if [#"Pergunta 3 - Recomendaria"] > 6 and [#"Pergunta 3 - Recomendaria"] < 9 then "Neutro"
    else if [#"Pergunta 3 - Recomendaria"] >= 9 then "Promotor"
    else ""),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Classificação da avaliação",{{"Data", type date}, {"media_da_avaliacao_do_cliente", type number}, {"Pergunta 1 - Atendimento Médico", type number}, {"Pergunta 2 - Local Do Atendimento", type number}, {"Pergunta 3 - Recomendaria", type number}})

in
    #"Tipo Alterado"