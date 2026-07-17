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
//   3) key_fix_id — фиксированное соответствие ГСЗ ↔ холдинг по ID.
//
// Выходные колонки на _HOLD_OD:
//   «условное ГСЗ» / «Бренд» — первое фактическое значение;
//   «статус» — текстовый статус (найдено / фикс / пересечения / «-»);
//   «Отладка_совпадения_ГСЗ» — все найденные ГСЗ через «; ».
//
// Оптимизации (как в Python-матчере):
//   - справочник подготавливается один раз (List.Buffer);
//   - предфильтр по всем AND/OR full|not-токенам (ключи не отбрасываются);
//     кандидат, если сработал хотя бы один — далее полная логика AND→OR→non.
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
    Имя_колонки_id = "ID холдинга",
    Имя_колонки_fix = "key_fix_id",

    // Тексты колонки «статус» (как в Python-матчере).
    Статус_нет = "-",
    Статус_одно = "найдено соответствие",
    Статус_несколько = "есть пересечения по ключам",
    Статус_фикс = "зафиксированное значение",
    Статус_фикс_часть = "найдена часть фиксированных значений",
    Статус_фикс_не_найден = "фикс значение не найдено",

    // Все ID холдингов из _HOLD_OD — для режима key_fix_id.
    fnНормализоватьId = (значение as any) as text =>
        Text.Trim(if значение = null then "" else Text.From(значение)),

    МножествоIdХолдингов =
        List.Buffer(
            List.Distinct(
                List.Select(
                    List.Transform(
                        Table.Column(Источник_HOLD_OD, Имя_колонки_id),
                        each fnНормализоватьId(_)
                    ),
                    each _ <> ""
                )
            )
        ),

    fnРазобратьFixIds = (значение as any) as list =>
        let
            текст = Text.Trim(if значение = null then "" else Text.From(значение)),
            части = if текст = "" then {} else Text.Split(текст, ";"),
            обрезанные = List.Transform(части, each Text.Trim(_)),
            безПустых = List.Select(обрезанные, each _ <> "")
        in
            безПустых,

    fnРежимFix = (fixIds as list) as text =>
        if List.Count(fixIds) = 0 then
            "none"
        else
            let
                найдено = List.Select(fixIds, each List.Contains(МножествоIdХолдингов, _)),
                nНайдено = List.Count(найдено),
                nВсего = List.Count(fixIds)
            in
                if nНайдено = 0 then
                    "fallback"
                else if nНайдено = nВсего then
                    "resolved"
                else
                    "partial",

    fnСтатусХолдинга = (список as list, числоФикс as number) as text =>
        let
            nВсего = List.Count(список),
            nКлюч = nВсего - числоФикс,
            строки = List.RemoveNulls({
                if числоФикс > 0 then Статус_фикс else null,
                if nКлюч > 1 then Статус_несколько else null,
                if nКлюч = 1 and числоФикс = 0 then Статус_одно else null,
                if nКлюч = 1 and числоФикс > 0 then Статус_одно else null
            })
        in
            if nВсего = 0 then
                Статус_нет
            else
                Text.Combine(строки, "#(lf)"),

    // Явные списки колонок. Если список НЕ пустой — используется вместо автоопределения.
    // Пример override: Колонки_and_full_явно = {"key_and_full_1", "key_and_full_2", "key_and_full_4"}.
    // or_* / and_* — все непустые full/not токены строки участвуют в предфильтре.
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

    // ─── Токены предфильтра ────────────────────────────────────────────────
    // Все AND и OR (full/not) — без отсечения до одного «якоря».
    // Кандидат, если найден хотя бы один; полная логика AND→OR→non — в fnСтрокаGszСовпала.
    fnВыбратьЯкоря = (andТокены as list, orПары as list) as list =>
        andТокены & orПары,

    // ─── Подготовка справочника один раз ────────────────────────────────────
    // Для каждой строки _base_gsz собираем токены, non-токены, флаг «есть ключи» и якоря.
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
            fixIds = fnРазобратьFixIds(fnЗначениеПоля(строка, Имя_колонки_fix)),
            fixMode = fnРежимFix(fixIds),
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
            якоря = fnВыбратьЯкоря(andТокены, orПары)
        in
            [
                Исходная = строка,
                ЕстьКлючи = естьКлючи,
                AndTokens = andТокены,
                OrPairs = orПары,
                AndNonTokens = andNonТокены,
                OrNonTokens = orNonТокены,
                FixIds = fixIds,
                FixMode = fixMode,
                Anchors = якоря
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

    // ─── Быстрый предфильтр по всем AND/OR-токенам (достаточно одного) ─────
    // full: Text.Contains; not: слово в списке токенов холдинга.
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

    // ─── Все совпадения для холдинга (fix-ID + ключи) ───────────────────────
    fnИмяGszИзMeta = (строкаMeta as record) as text =>
        let
            имя = fnЗначениеПоля(строкаMeta[Исходная], Имя_колонки_гсз)
        in
            if имя = null then "" else Text.From(имя),

    fnНайтиСовпадения = (строкаХолдинга as record) as record =>
        let
            idХолдинга = fnНормализоватьId(Record.Field(строкаХолдинга, Имя_колонки_id)),
            текст = fnНормализоватьТекст(Record.Field(строкаХолдинга, Имя_колонки_холдинг)),
            слова = fnСлова(текст),

            фиксСтроки =
                if idХолдинга = "" then
                    {}
                else
                    List.Select(
                        Справочник_подготовленный,
                        (строкаMeta) =>
                            (строкаMeta[FixMode] = "resolved" or строкаMeta[FixMode] = "partial")
                            and List.Contains(строкаMeta[FixIds], idХолдинга)
                    ),

            фиксИмена =
                List.Select(
                    List.Transform(фиксСтроки, each fnИмяGszИзMeta(_)),
                    each Text.Trim(_) <> ""
                ),

            кандидатыПоКлючам =
                List.Select(
                    Справочник_подготовленный,
                    (строкаMeta) =>
                        (строкаMeta[FixMode] = "none" or строкаMeta[FixMode] = "fallback" or строкаMeta[FixMode] = "partial")
                        and fnКандидатПоЯкорям(текст, слова, строкаMeta[Anchors])
                ),

            ключСтроки =
                List.Select(
                    кандидатыПоКлючам,
                    (строкаMeta) => fnСтрокаGszСовпала(текст, строкаMeta)
                ),

            ключИмена =
                List.Select(
                    List.Transform(ключСтроки, each fnИмяGszИзMeta(_)),
                    each Text.Trim(_) <> ""
                ),

            всеИмена = фиксИмена & ключИмена,
            статус = fnСтатусХолдинга(всеИмена, List.Count(фиксИмена))
        in
            [Matches = всеИмена, Status = статус, Count = List.Count(всеИмена)],

    // ─── Итоговые колонки на листе _HOLD_OD ─────────────────────────────────
    СКолонкой_результат =
        Table.AddColumn(
            Источник_HOLD_OD,
            "РезультатСопоставления",
            each fnНайтиСовпадения(_)
        ),

    СКолонкой_условное_ГСЗ =
        Table.AddColumn(
            СКолонкой_результат,
            "условное ГСЗ",
            each
                let
                    список = Record.Field(Record.Field(_, "РезультатСопоставления"), "Matches")
                in
                    if List.Count(список) = 0 then
                        Статус_нет
                    else
                        List.First(список)
        ),

    СКолонкой_статус =
        Table.AddColumn(
            СКолонкой_условное_ГСЗ,
            "статус",
            each Record.Field(Record.Field(_, "РезультатСопоставления"), "Status")
        ),

    СКолонкой_отладка =
        Table.AddColumn(
            СКолонкой_статус,
            "Отладка_совпадения_ГСЗ",
            each
                let
                    список = Record.Field(Record.Field(_, "РезультатСопоставления"), "Matches")
                in
                    if List.Count(список) = 0 then
                        Статус_нет
                    else
                        Text.Combine(список, "; ")
        ),

    Без_служебной_колонки =
        Table.RemoveColumns(СКолонкой_отладка, {"РезультатСопоставления"}),

    Результат =
        Без_служебной_колонки
in
    Результат
