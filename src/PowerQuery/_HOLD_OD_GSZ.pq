// Запрос: _HOLD_OD_GSZ
// Excel 2016 (русский), Power Query
// Сопоставление «Холдинг» (_HOLD_OD) с ключами из _base_gsz

let
    // ─── Загрузка источников ───────────────────────────────────────────────
    Источник_HOLD_OD =
        Excel.CurrentWorkbook(){[Name = "_HOLD_OD"]}[Content],

    Источник_base_gsz =
        Excel.CurrentWorkbook(){[Name = "_base_gsz"]}[Content],

    Справочник_буфер =
        Table.Buffer(Источник_base_gsz),

    // ─── Имена колонок ───────────────────────────────────────────────────────
    Имя_колонки_холдинг = "Холдинг",
    Имя_колонки_гсз = "Наименование, регион",

    Колонки_and_full = {"key_and_full_1", "key_and_full_2", "key_and_full_3"},
    Колонки_and_not = {"key_and_not_1", "key_and_not_2", "key_and_not_3"},
    Колонки_or_full = {"key_or_full_1", "key_or_full_2", "key_or_full_3"},
    Колонки_or_not = {"key_or_not_1", "key_or_not_2", "key_or_not_3"},

    Все_колонки_ключей =
        Колонки_and_full
        & Колонки_and_not
        & Колонки_or_full
        & Колонки_or_not,

    // ─── Нормализация текста ─────────────────────────────────────────────────
    fnНормализоватьТекст = (значение as any) as text =>
        let
            какТекст =
                if значение = null then
                    ""
                else
                    Text.From(значение),
            обрезанный = Text.Trim(какТекст)
        in
            Text.Lower(обрезанный),

    // ─── Буква (кириллица + латиница) ────────────────────────────────────────
    fnЭтоБуква = (символ as text) as logical =>
        if Text.Length(символ) <> 1 then
            false
        else
            let
                с = Text.Lower(символ),
                латиница = с >= "a" and с <= "z",
                кириллица = (с >= "а" and с <= "я") or с = "ё"
            in
                латиница or кириллица,

    // ─── Все позиции подстроки (full), 0-based ───────────────────────────────
    fnВсеПозицииFull = (текст as text, слово as text) as list =>
        let
            длСлова = Text.Length(слово),
            длТекста = Text.Length(текст),
            максСтарт =
                if длСлова = 0 or длТекста < длСлова then
                    -1
                else
                    длТекста - длСлова,
            старты = List.Numbers(0, максСтарт + 1, 1),
            совпадения =
                List.Select(
                    старты,
                    (поз) => Text.Range(текст, поз, длСлова) = слово
                ),
            интервалы =
                List.Transform(
                    совпадения,
                    (поз) => [Start = поз, End = поз + длСлова - 1]
                )
        in
            интервалы,

    // ─── Позиции отдельного слова (not) ──────────────────────────────────────
    fnПозицииNot = (текст as text, слово as text) as list =>
        let
            все = fnВсеПозицииFull(текст, слово),
            длТекста = Text.Length(текст),
            сГраницами =
                List.Select(
                    все,
                    (инт) =>
                        let
                            старт = инт[Start],
                            конец = инт[End],
                            слеваОк =
                                старт = 0
                                or not fnЭтоБуква(Text.Middle(текст, старт - 1, 1)),
                            справаОк =
                                конец = длТекста - 1
                                or not fnЭтоБуква(Text.Middle(текст, конец + 1, 1))
                        in
                            слеваОк and справаОк
                )
        in
            сГраницами,

  // ─── Пересечение интервалов ──────────────────────────────────────────────
    fnИнтервалыПересекаются = (а as record, б as record) as logical =>
        not (а[End] < б[Start] or б[End] < а[Start]),

    // ─── AND: все токены без пересечения позиций ─────────────────────────────
    fnПроверитьAndТокены = (текст as text, токены as list) as logical =>
        if List.Count(токены) = 0 then
            true
        else
            let
                сведения =
                    List.Transform(
                        токены,
                        (т) =>
                            let
                                слово = т[Word],
                                позиции =
                                    if т[IsFull] then
                                        fnВсеПозицииFull(текст, слово)
                                    else
                                        fnПозицииNot(текст, слово)
                            in
                                [Word = слово, Positions = позиции]
                    ),
                естьПустой =
                    List.AnyTrue(
                        List.Transform(сведения, each List.Count(_[Positions]) = 0)
                    ),
                подобрать =
                    (индекс as number, выбранные as list) as logical =>
                        if индекс >= List.Count(сведения) then
                            true
                        else
                            let
                                варианты = сведения{индекс}[Positions],
                                проверитьВариант = (инт as record) as logical =>
                                    let
                                        конфликт =
                                            List.AnyTrue(
                                                List.Transform(
                                                    выбранные,
                                                    (в) => fnИнтервалыПересекаются(в, инт)
                                                )
                                            )
                                    in
                                        if конфликт then
                                            false
                                        else
                                            @подобрать(индекс + 1, выбранные & {инт})
                            in
                                List.AnyTrue(List.Transform(варианты, проверитьВариант))
            in
                if естьПустой then
                    false
                else
                    подобрать(0, {}),

    // ─── Одно поле OR ─────────────────────────────────────────────────────────
    fnПолеOrСовпало = (текст as text, значение as any, этоFull as logical) as logical =>
        let
            слово = fnНормализоватьТекст(значение)
        in
            if слово = "" then
                false
            else if этоFull then
                List.Count(fnВсеПозицииFull(текст, слово)) > 0
            else
                List.Count(fnПозицииNot(текст, слово)) > 0,

    // ─── Значение поля записи (безопасно) ────────────────────────────────────
    fnЗначениеПоля = (строка as record, имя as text) as any =>
        if Record.HasFields(строка, имя) then
            Record.Field(строка, имя)
        else
            null,

    // ─── Собрать AND-токены из строки справочника ────────────────────────────
    fnСобратьAndТокены = (строка as record) as list =>
        let
            изFull =
                List.Transform(
                    Колонки_and_full,
                    (имя) =>
                        let
                            слово = fnНормализоватьТекст(fnЗначениеПоля(строка, имя))
                        in
                            if слово = "" then
                                null
                            else
                                [Word = слово, IsFull = true]
                ),
            изNot =
                List.Transform(
                    Колонки_and_not,
                    (имя) =>
                        let
                            слово = fnНормализоватьТекст(fnЗначениеПоля(строка, имя))
                        in
                            if слово = "" then
                                null
                            else
                                [Word = слово, IsFull = false]
                )
        in
            List.RemoveNulls(изFull & изNot),

    // ─── Есть ли хотя бы один непустой ключ ─────────────────────────────────
    fnЕстьКлючи = (строка as record) as logical =>
        List.AnyTrue(
            List.Transform(
                Все_колонки_ключей,
                (имя) => fnНормализоватьТекст(fnЗначениеПоля(строка, имя)) <> ""
            )
        ),

    // ─── OR-блок ─────────────────────────────────────────────────────────────
    fnПроверитьOrБлок = (текст as text, строка as record) as logical =>
        let
            пары =
                List.Combine({
                    List.Transform(Колонки_or_full, (имя) => [Имя = имя, IsFull = true]),
                    List.Transform(Колонки_or_not, (имя) => [Имя = имя, IsFull = false])
                }),
            заполненные =
                List.Select(
                    пары,
                    (п) => fnНормализоватьТекст(fnЗначениеПоля(строка, п[Имя])) <> ""
                ),
            результаты =
                List.Transform(
                    заполненные,
                    (п) => fnПолеOrСовпало(текст, fnЗначениеПоля(строка, п[Имя]), п[IsFull])
                )
        in
            List.Count(заполненные) = 0 or List.AnyTrue(результаты),

    // ─── Совпадение одной строки _base_gsz ──────────────────────────────────
    fnСтрокаGszСовпала = (текстХолдинга as text, строкаGsz as record) as logical =>
        let
            текст = fnНормализоватьТекст(текстХолдинга)
        in
            if текст = "" then
                false
            else if not fnЕстьКлючи(строкаGsz) then
                false
            else
                let
                    andТокены = fnСобратьAndТокены(строкаGsz),
                    andОк = fnПроверитьAndТокены(текст, andТокены),
                    orОк = fnПроверитьOrБлок(текст, строкаGsz)
                in
                    andОк and orОк,

    // ─── Все совпадения для холдинга (порядок как в _base_gsz) ───────────────
    // Справочник_буфер подставляется из замыкания — один аргумент в вызове
    fnНайтиВсеGsz = (текстХолдинга as any) as list =>
        let
            подходящие =
                Table.SelectRows(
                    Справочник_буфер,
                    (строка) => fnСтрокаGszСовпала(текстХолдинга, строка)
                ),
            именаСырые = Table.Column(подходящие, Имя_колонки_гсз),
            имена =
                List.Transform(
                    именаСырые,
                    each if _ = null then "" else Text.From(_)
                ),
            безПустых = List.Select(имена, each Text.Trim(_) <> "")
        in
            безПустых,

    // ─── Поиск совпадений (один проход на строку) ────────────────────────────
    СКолонкой_список_совпадений =
        Table.AddColumn(
            Источник_HOLD_OD,
            "СовпаденияГСЗ",
            each fnНайтиВсеGsz(Record.Field(_, Имя_колонки_холдинг))
        ),

    СКолонкой_условное_ГСЗ =
        Table.AddColumn(
            СКолонкой_список_совпадений,
            "условное ГСЗ",
            each
                let
                    список = Record.Field(_, "СовпаденияГСЗ")
                in
                    if List.Count(список) = 0 then
                        "-"
                    else
                        List.First(список)
        ),

    СКолонкой_отладка =
        Table.AddColumn(
            СКолонкой_условное_ГСЗ,
            "Отладка_совпадения_ГСЗ",
            each
                let
                    список = Record.Field(_, "СовпаденияГСЗ")
                in
                    if List.Count(список) = 0 then
                        "-"
                    else
                        Text.Combine(список, "; ")
        ),

    Без_служебной_колонки =
        Table.RemoveColumns(СКолонкой_отладка, {"СовпаденияГСЗ"}),

    Результат =
        Без_служебной_колонки
in
    Результат
