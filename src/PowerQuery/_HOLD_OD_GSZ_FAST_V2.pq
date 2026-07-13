// =============================================================================
// Запрос: _HOLD_OD_GSZ_FAST_V2
// Excel 2016 (русский), Power Query
// =============================================================================
//
// Назначение:
//   Сопоставление текста холдинга (лист _HOLD_OD) с условными ГСЗ из справочника
//   _base_gsz по шести группам ключевых колонок.
//
// Группы ключей (колонки key_<группа>_<N>):
//   and_full / and_not — обязательные совпадения (подстрока / отдельное слово);
//   or_full  / or_not  — альтернативные совпадения (достаточно одного);
//   and_non  / or_non  — исключения по тексту холдинга БЕЗ пробелов:
//     and_non — отклонить, если найдены ВСЕ non-ключи;
//     or_non  — отклонить, если найден ХОТЯ БЫ ОДИН non-ключ.
//
// Отличия V2 от _HOLD_OD_GSZ_FAST:
//   1) Поддержка and_non / or_non.
//   2) Автоопределение числа колонок key_* по заголовкам _base_gsz
//      (или явный override через Колонки_*_явно в начале запроса).
//
// Выходные колонки на _HOLD_OD:
//   «условное ГСЗ» — первое совпадение или «есть пересечения по ключам»;
//   «Отладка_совпадения_ГСЗ» — все найденные ГСЗ через «; ».
//
// Оптимизации (как в Python-матчере):
//   - справочник подготавливается один раз (List.Buffer);
//   - якорный предфильтр кандидатов по самому длинному ключу строки.
// =============================================================================

let
    // ─── Загрузка источников ───────────────────────────────────────────────
    // Смарт-таблицы текущей книги Excel (имена должны совпадать с config).
    Источник_HOLD_OD =
        Excel.CurrentWorkbook(){[Name = "_HOLD_OD"]}[Content],

    Источник_base_gsz =
        Excel.CurrentWorkbook(){[Name = "_base_gsz"]}[Content],

    // ─── Имена колонок ─────────────────────────────────────────────────────
    Имя_колонки_холдинг = "Холдинг",
    Имя_колонки_гсз = "Наименование, регион",

    // Явные списки колонок. Если список НЕ пустой — используется вместо автоопределения.
    // Пример override: Колонки_and_full_явно = {"key_and_full_1", "key_and_full_2", "key_and_full_4"}.
    Колонки_and_full_явно = {},
    Колонки_and_not_явно = {},
    Колонки_and_non_явно = {},
    Колонки_or_full_явно = {},
    Колонки_or_not_явно = {},
    Колонки_or_non_явно = {},

    ИменаКолонок_base = Table.ColumnNames(Источник_base_gsz),

    // ─── Универсальный подбор колонок key_<группа>_<N> ─────────────────────
    // Извлекает числовой суффикс из имени колонки (key_and_full_3 → 3).
    fnЧислоСуффикса = (имя as text) as nullable number =>
        let
            части = Text.Split(имя, "_"),
            последняя = List.Last(части)
        in
            try Number.From(последняя) otherwise null,

    // Находит все колонки с заданным префиксом и сортирует по числовому суффиксу.
    fnКолонкиПоПрефиксу = (все as list, префикс as text) as list =>
        let
            matched = List.Select(все, each Text.StartsWith(_, префикс)),
            asRecords =
                List.Transform(
                    matched,
                    (имя) => [Name = имя, N = fnЧислоСуффикса(имя)]
                ),
            withNum = List.Select(asRecords, each _[N] <> null),
            sorted =
                List.Sort(
                    withNum,
                    (a, b) =>
                        if a[N] < b[N] then
                            -1
                        else if a[N] > b[N] then
                            1
                        else
                            0
                ),
            names = List.Transform(sorted, each _[Name])
        in
            names,

    // Явный список приоритетнее; иначе — авто по префиксу (key_and_full_, key_and_non_ и т.д.).
    fnВыбратьКолонки = (явно as list, префикс as text) as list =>
        if List.Count(явно) > 0 then
            явно
        else
            fnКолонкиПоПрефиксу(ИменаКолонок_base, префикс),

    Колонки_and_full = fnВыбратьКолонки(Колонки_and_full_явно, "key_and_full_"),
    Колонки_and_not = fnВыбратьКолонки(Колонки_and_not_явно, "key_and_not_"),
    Колонки_and_non = fnВыбратьКолонки(Колонки_and_non_явно, "key_and_non_"),
    Колонки_or_full = fnВыбратьКолонки(Колонки_or_full_явно, "key_or_full_"),
    Колонки_or_not = fnВыбратьКолонки(Колонки_or_not_явно, "key_or_not_"),
    Колонки_or_non = fnВыбратьКолонки(Колонки_or_non_явно, "key_or_non_"),

    // Объединение всех групп — для проверки «есть ли хоть один ключ» в строке справочника.
    Все_колонки_ключей =
        Колонки_and_full
        & Колонки_and_not
        & Колонки_and_non
        & Колонки_or_full
        & Колонки_or_not
        & Колонки_or_non,

    // ─── Нормализация текста ───────────────────────────────────────────────
    // Обычные ключи: trim + нижний регистр.
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

    // Non-ключи: тот же trim/lower, но без пробелов, табов и переводов строк.
    // Позволяет находить «ММБ» и в «ну пикммб», и в «ну пик ММБ».
    fnНормализоватьNonТекст = (значение as any) as text =>
        let
            норм = fnНормализоватьТекст(значение),
            символы = Text.ToList(норм),
            безПробелов =
                List.Select(
                    символы,
                    each
                        _ <> " "
                        and _ <> "#(tab)"
                        and _ <> "#(cr)"
                        and _ <> "#(lf)"
                )
        in
            Text.Combine(безПробелов, ""),

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
    // Все токены должны встретиться; интервалы вхождений не должны пересекаться
    // (backtracking по вариантам позиций).
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
    // Пустой OR-блок считается пройденным; иначе нужен хотя бы один совпавший токен.
    fnПроверитьOrБлок = (текст as text, orПары as list) as logical =>
        let
            результаты =
                List.Transform(
                    orПары,
                    (п) => fnПолеOrСовпало(текст, п[Word], п[IsFull])
                )
        in
            List.Count(orПары) = 0 or List.AnyTrue(результаты),

    // ─── Исключения non: AND — все найдены, OR — хотя бы один ───────────────
    // Поиск в compact-тексте холдинга (без пробелов). AND non использует тот же
    // backtracking по непересекающимся интервалам, что и обычный AND-блок.
    fnПроверитьAndNonИсключение = (компактТекст as text, токены as list) as logical =>
        if List.Count(токены) = 0 then
            false
        else
            let
                сведения =
                    List.Transform(
                        токены,
                        (слово) => [Word = слово, Positions = fnВсеПозицииFull(компактТекст, слово)]
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

    fnПроверитьOrNonИсключение = (компактТекст as text, токены as list) as logical =>
        if List.Count(токены) = 0 then
            false
        else
            List.AnyTrue(
                List.Transform(
                    токены,
                    (слово) => List.Count(fnВсеПозицииFull(компактТекст, слово)) > 0
                )
            ),

    // ─── Выбрать "якорь" для быстрого предфильтра ──────────────────────────
    // Приоритет: and_not → and_full → or_not → or_full; внутри группы — самый длинный ключ.
    fnЛучшийЯкорь = (кандидаты as list) as any =>
        if List.Count(кандидаты) = 0 then
            null
        else
            let
                сДлиной =
                    List.Transform(
                        кандидаты,
                        (к) => [Word = к[Word], IsFull = к[IsFull], Len = Text.Length(к[Word])]
                    ),
                отсортированные =
                    List.Sort(
                        сДлиной,
                        (a, b) =>
                            if a[Len] > b[Len] then
                                -1
                            else if a[Len] < b[Len] then
                                1
                            else
                                0
                    )
            in
                if List.Count(отсортированные) = 0 then
                    null
                else
                    [Word = отсортированные{0}[Word], IsFull = отсортированные{0}[IsFull]],

    fnВыбратьЯкорь = (andТокены as list, orПары as list) as any =>
        let
            andNot = List.Select(andТокены, each _[IsFull] = false),
            andFull = List.Select(andТокены, each _[IsFull] = true),
            orNot = List.Select(orПары, each _[IsFull] = false),
            orFull = List.Select(orПары, each _[IsFull] = true),
            якорь =
                if List.Count(andNot) > 0 then
                    fnЛучшийЯкорь(andNot)
                else if List.Count(andFull) > 0 then
                    fnЛучшийЯкорь(andFull)
                else if List.Count(orNot) > 0 then
                    fnЛучшийЯкорь(orNot)
                else
                    fnЛучшийЯкорь(orFull)
        in
            якорь,

    // ─── Подготовка справочника один раз ────────────────────────────────────
    // Для каждой строки _base_gsz собираем токены, non-токены, флаг «есть ключи» и якорь.
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
            andNonСырые =
                List.Transform(
                    Колонки_and_non,
                    (имя) => fnНормализоватьNonТекст(fnЗначениеПоля(строка, имя))
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
            orNonСырые =
                List.Transform(
                    Колонки_or_non,
                    (имя) => fnНормализоватьNonТекст(fnЗначениеПоля(строка, имя))
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
            andNonТокены = List.Select(andNonСырые, each _ <> ""),
            orNonТокены = List.Select(orNonСырые, each _ <> ""),
            // Строка участвует в сопоставлении, только если заполнен хотя бы один ключ
            естьКлючи =
                List.AnyTrue(
                    List.Transform(
                        Все_колонки_ключей,
                        (имя) =>
                            let
                                значение = fnЗначениеПоля(строка, имя),
                                этоNon =
                                    List.Contains(Колонки_and_non, имя)
                                    or List.Contains(Колонки_or_non, имя)
                            in
                                if этоNon then
                                    fnНормализоватьNonТекст(значение) <> ""
                                else
                                    fnНормализоватьТекст(значение) <> ""
                    )
                ),
            якорь = fnВыбратьЯкорь(andТокены, orПары)
        in
            [
                Исходная = строка,
                ЕстьКлючи = естьКлючи,
                AndTokens = andТокены,
                OrPairs = orПары,
                AndNonTokens = andNonТокены,
                OrNonTokens = orNonТокены,
                Anchor = якорь
            ],

    Справочник_подготовленный_таблица =
        Table.AddColumn(
            Источник_base_gsz,
            "Meta",
            each fnПодготовитьСтрокуGsz(_)
        ),

    // List.Buffer — справочник читается из памяти при каждом холдинге (ускорение).
    Справочник_подготовленный =
        List.Buffer(Table.Column(Справочник_подготовленный_таблица, "Meta")),

    // ─── Совпала ли строка справочника ─────────────────────────────────────
    // Порядок проверок: AND → OR → not and_non → not or_non (как в Python row_matches).
    fnСтрокаGszСовпала = (текст as text, строкаMeta as record) as logical =>
        if текст = "" then
            false
        else if not строкаMeta[ЕстьКлючи] then
            false
        else
            let
                компактТекст = fnНормализоватьNonТекст(текст),
                andОк = fnПроверитьAndТокены(текст, строкаMeta[AndTokens]),
                orОк = fnПроверитьOrБлок(текст, строкаMeta[OrPairs]),
                andNonОтклонить = fnПроверитьAndNonИсключение(компактТекст, строкаMeta[AndNonTokens]),
                orNonОтклонить = fnПроверитьOrNonИсключение(компактТекст, строкаMeta[OrNonTokens])
            in
                andОк and orОк and not andNonОтклонить and not orNonОтклонить,

    // ─── Быстрый предфильтр по якорю ────────────────────────────────────────
    // full-якорь: Text.Contains; not-якорь: слово есть в списке токенов холдинга.
    fnКандидатПоЯкорю = (текст as text, слова as list, якорь as any) as logical =>
        if якорь = null then
            true
        else if якорь[IsFull] then
            Text.Contains(текст, якорь[Word])
        else
            List.Contains(слова, якорь[Word]),

    // ─── Все совпадения для холдинга ────────────────────────────────────────
    // 1) предфильтр по якорю; 2) полная проверка fnСтрокаGszСовпала; 3) имена ГСЗ.
    fnНайтиВсеGsz = (текстХолдинга as any) as list =>
        let
            текст = fnНормализоватьТекст(текстХолдинга),
            слова = fnСлова(текст),

            кандидаты =
                List.Select(
                    Справочник_подготовленный,
                    (строкаMeta) => fnКандидатПоЯкорю(текст, слова, строкаMeta[Anchor])
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

    // ─── Итоговые колонки на листе _HOLD_OD ─────────────────────────────────
    // Служебная «СовпаденияГСЗ» (список) удаляется перед выдачей результата.
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
                    // 0 совпадений → «-»; 1 → имя ГСЗ; несколько → фиксированная фраза
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
