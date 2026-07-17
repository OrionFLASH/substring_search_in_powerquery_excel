// =============================================================================
// Запрос: _HOLD_OD_GSZ_FAST
// Excel 2016 (русский), Power Query
// =============================================================================
//
// Ускоренное сопоставление «Холдинг» (_HOLD_OD) с ключами из _base_gsz.
// Без поддержки and_non / or_non — для этого используйте _HOLD_OD_GSZ_FAST_V2.
//
// Группы ключей (фиксированные 3 колонки на группу):
//   and_full / and_not — обязательные; or_full / or_not — альтернативные.
//
// Выход: «условное ГСЗ», «Отладка_совпадения_ГСЗ».
// Оптимизации: List.Buffer + предфильтр по всем AND/OR-токенам (без отсечения ключей).
// =============================================================================

let
    // ─── Загрузка источников ───────────────────────────────────────────────
    // Смарт-таблицы текущей книги Excel.
    Источник_HOLD_OD =
        Excel.CurrentWorkbook(){[Name = "_HOLD_OD"]}[Content],

    Источник_base_gsz =
        Excel.CurrentWorkbook(){[Name = "_base_gsz"]}[Content],

    // ─── Имена колонок ─────────────────────────────────────────────────────
    Имя_колонки_холдинг = "Холдинг",
    Имя_колонки_гсз = "Наименование, регион",

    // Фиксированный набор колонок ключей (в отличие от V2 — без автоопределения).
    // or_* / and_* — все непустые full/not токены строки участвуют в предфильтре.
    Колонки_and_full = {"key_and_full_1", "key_and_full_2", "key_and_full_3"},
    Колонки_and_not = {"key_and_not_1", "key_and_not_2", "key_and_not_3"},
    Колонки_or_full = {"key_or_full_1", "key_or_full_2", "key_or_full_3"},
    Колонки_or_not = {"key_or_not_1", "key_or_not_2", "key_or_not_3"},

    Все_колонки_ключей =
        Колонки_and_full
        & Колонки_and_not
        & Колонки_or_full
        & Колонки_or_not,

    // ─── Нормализация текста ───────────────────────────────────────────────
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

    // ─── Буква (кириллица + латиница) ──────────────────────────────────────
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

    // ─── Нормализованные слова строки (для быстрых "not"-проверок) ─────────
    fnСлова = (текст as text) as list =>
        let
            разделители = " ,.;:!?()[]{}<>""'«»/\\|-_+=*&^%$#@~`" & "#(tab)" & "#(cr)" & "#(lf)",
            куски = Text.SplitAny(текст, разделители),
            безПустых = List.Select(куски, each _ <> ""),
            нижний = List.Transform(безПустых, each Text.Lower(_)),
            уникальные = List.Distinct(нижний)
        in
            уникальные,

    // ─── Все позиции подстроки (full), 0-based ─────────────────────────────
    fnВсеПозицииFull = (текст as text, слово as text) as list =>
        let
            длСлова = Text.Length(слово),
            сырыеПозиции =
                if длСлова = 0 then
                    {}
                else
                    Text.PositionOf(текст, слово, Occurrence.All),
            позиции =
                if сырыеПозиции is list then
                    сырыеПозиции
                else if сырыеПозиции = -1 then
                    {}
                else
                    {сырыеПозиции},
            интервалы =
                List.Transform(
                    позиции,
                    (поз) => [Start = поз, End = поз + длСлова - 1]
                )
        in
            интервалы,

    // ─── Позиции отдельного слова (not) ────────────────────────────────────
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

    // ─── Пересечение интервалов ─────────────────────────────────────────────
    fnИнтервалыПересекаются = (а as record, б as record) as logical =>
        not (а[End] < б[Start] or б[End] < а[Start]),

    // ─── Одно поле OR ───────────────────────────────────────────────────────
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

    // ─── Значение поля записи (безопасно) ──────────────────────────────────
    fnЗначениеПоля = (строка as record, имя as text) as any =>
        if Record.HasFields(строка, имя) then
            Record.Field(строка, имя)
        else
            null,

    // ─── Проверка AND-токенов ───────────────────────────────────────────────
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

    // ─── Проверка OR-блока ──────────────────────────────────────────────────
    fnПроверитьOrБлок = (текст as text, orПары as list) as logical =>
        let
            результаты =
                List.Transform(
                    orПары,
                    (п) => fnПолеOrСовпало(текст, п[Word], п[IsFull])
                )
        in
            List.Count(orПары) = 0 or List.AnyTrue(результаты),

    // ─── Токены предфильтра ────────────────────────────────────────────────
    // Все AND и OR (full/not) — без отсечения до одного «якоря».
    // Кандидат, если найден хотя бы один; полная логика AND→OR→non — в fnСтрокаGszСовпала.
    fnВыбратьЯкоря = (andТокены as list, orПары as list) as list =>
        andТокены & orПары,

    // ─── Подготовка справочника один раз ────────────────────────────────────
    fnПодготовитьСтрокуGsz = (строка as record) as record =>
        let
            andFullСырые =
                List.Transform(
                    Колонки_and_full,
                    (имя) => fnНормализоватьТекст(fnЗначениеПоля(строка, имя))
                ),
            andNotСырые =
                List.Transform(
                    Колонки_and_not,
                    (имя) => fnНормализоватьТекст(fnЗначениеПоля(строка, имя))
                ),
            orFullСырые =
                List.Transform(
                    Колонки_or_full,
                    (имя) => fnНормализоватьТекст(fnЗначениеПоля(строка, имя))
                ),
            orNotСырые =
                List.Transform(
                    Колонки_or_not,
                    (имя) => fnНормализоватьТекст(fnЗначениеПоля(строка, имя))
                ),

            andТокены =
                List.RemoveNulls(
                    List.Combine({
                        List.Transform(andFullСырые, (w) => if w = "" then null else [Word = w, IsFull = true]),
                        List.Transform(andNotСырые, (w) => if w = "" then null else [Word = w, IsFull = false])
                    })
                ),
            orПары =
                List.RemoveNulls(
                    List.Combine({
                        List.Transform(orFullСырые, (w) => if w = "" then null else [Word = w, IsFull = true]),
                        List.Transform(orNotСырые, (w) => if w = "" then null else [Word = w, IsFull = false])
                    })
                ),
            естьКлючи =
                List.AnyTrue(
                    List.Transform(
                        Все_колонки_ключей,
                        (имя) => fnНормализоватьТекст(fnЗначениеПоля(строка, имя)) <> ""
                    )
                ),
            якоря = fnВыбратьЯкоря(andТокены, orПары)
        in
            [
                Исходная = строка,
                ЕстьКлючи = естьКлючи,
                AndTokens = andТокены,
                OrPairs = orПары,
                Anchors = якоря
            ],

    Справочник_подготовленный_таблица =
        Table.AddColumn(
            Источник_base_gsz,
            "Meta",
            each fnПодготовитьСтрокуGsz(_)
        ),

    Справочник_подготовленный =
        List.Buffer(Table.Column(Справочник_подготовленный_таблица, "Meta")),

    // ─── Совпала ли строка справочника (точная бизнес-логика) ───────────────
    fnСтрокаGszСовпала = (текст as text, строкаMeta as record) as logical =>
        if текст = "" then
            false
        else if not строкаMeta[ЕстьКлючи] then
            false
        else
            let
                andОк = fnПроверитьAndТокены(текст, строкаMeta[AndTokens]),
                orОк = fnПроверитьOrБлок(текст, строкаMeta[OrPairs])
            in
                andОк and orОк,

    // ─── Быстрый предфильтр по всем AND/OR-токенам (достаточно одного) ─────
    fnКандидатПоОдномуЯкорю = (текст as text, слова as list, якорь as any) as logical =>
        if якорь = null then
            true
        else if якорь[IsFull] then
            Text.Contains(текст, якорь[Word])
        else
            List.Contains(слова, якорь[Word]),

    fnКандидатПоЯкорям = (текст as text, слова as list, якоря as list) as logical =>
        if List.Count(якоря) = 0 then
            true
        else
            List.AnyTrue(
                List.Transform(якоря, each fnКандидатПоОдномуЯкорю(текст, слова, _))
            ),

    // ─── Все совпадения для холдинга ────────────────────────────────────────
    fnНайтиВсеGsz = (текстХолдинга as any) as list =>
        let
            текст = fnНормализоватьТекст(текстХолдинга),
            слова = fnСлова(текст),

            кандидаты =
                List.Select(
                    Справочник_подготовленный,
                    (строкаMeta) => fnКандидатПоЯкорям(текст, слова, строкаMeta[Anchors])
                ),

            подходящие =
                List.Select(
                    кандидаты,
                    (строкаMeta) => fnСтрокаGszСовпала(текст, строкаMeta)
                ),
            имена =
                List.Transform(
                    подходящие,
                    (строкаMeta) =>
                        let
                            имя = fnЗначениеПоля(строкаMeta[Исходная], Имя_колонки_гсз)
                        in
                            if имя = null then "" else Text.From(имя)
                ),
            безПустых = List.Select(имена, each Text.Trim(_) <> "")
        in
            безПустых,

    // ─── Поиск совпадений (один проход на строку) ───────────────────────────
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
