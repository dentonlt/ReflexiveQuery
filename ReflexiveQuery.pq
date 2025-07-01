//ReflexiveQuery.pq
// CHANGES
//   20250701 - v1; dentonlt

let

    // Set name of this query. Since it's self-referencing, it needs to know its own name.
    ThisQueryName = "OutputFinal",

    // Set source data query/location. This wrapper will append to that query.
    Source = SourceAndChanges,

    // Set column name for unique keys.
    KeyCol = "ID",

    // Set list of required 'additional' columns. Add/remove whenever. When removing a column from this list, its data will stay in the table until manually removed (Table -> Delete Column).
    ColumnsAdditionalRequired = {"Comment1", "Comment2"}, 

    /* =====
    From this point on, the code just manages the incoming data.
    ==== */
    #"===Processing Follows===" = null,

    // 1. Save a snapshot of the "current" output from this query (if that output table exists). We'll merge
    // some of that data back later.
    PreviousOutput =
        if List.Contains(Table.Column(Excel.CurrentWorkbook(), "Name"), ThisQueryName)
        then Excel.CurrentWorkbook(){[Name=ThisQueryName]}[Content]
        else null,

    // 2. Get 'additional' column names. We'll need these to decide what columns to expand or pitch after merging.
    // We can't just rely up on the list of names incoming (from ColumnsRequired, above) because we don't know if we have added
    // those columns in a previous step or not ... or if one or more has been intentionally removed.
    ColumnsAdditional = List.Difference(
        if PreviousOutput is null then {}
            else Table.ColumnNames(PreviousOutput),
        Table.ColumnNames(Source)),


    // 3. optional/conditional merge
    Merged = if PreviousOutput is null then Source
        else
            // Merge, then expand needed values. The 'Historical' label is arbitrary.
            Table.ExpandTableColumn(
            Table.NestedJoin(Source, KeyCol, PreviousOutput, KeyCol, "Historical", JoinKind.LeftOuter),
            "Historical", ColumnsAdditional, ColumnsAdditional),

    // 4. if we are missing any of the 'additional' columns ... add them (multiple steps ...)
    ColumnsMissing = List.Select(ColumnsAdditionalRequired, each not Table.HasColumns(Merged, _)), // only those missing
    FinalTable = List.Accumulate(ColumnsMissing, Merged,
        (state, current) => Table.AddColumn(state, current, each null)
    ),

    // bonus/fluff: sort by ID
    FinalTableSorted = Table.Sort(FinalTable, KeyCol)
in
    FinalTableSorted
