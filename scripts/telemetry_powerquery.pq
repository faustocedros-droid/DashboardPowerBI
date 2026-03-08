// =============================================================================
// telemetry_powerquery.pq
// -----------------------------------------------------------------------------
// Complete Power Query M script for importing and transforming
// data/sample_telemetry.xlsx into a clean table ready for Power BI charts.
//
// HOW TO USE
// ----------
// 1. Open Power BI Desktop.
// 2. Home → Get Data → Blank Query.
// 3. In the Power Query Editor: Home → Advanced Editor.
// 4. Replace all existing text with the contents of this file.
// 5. Replace the FilePath parameter value with the actual path to your
//    sample_telemetry.xlsx (absolute path or use a Power BI parameter).
// 6. Click Done → Close & Apply.
// =============================================================================

let
    // -------------------------------------------------------------------------
    // PARAMETER – update this path to point at your local copy of the file.
    // -------------------------------------------------------------------------
    FilePath = "C:/path/to/data/sample_telemetry.xlsx",

    // -------------------------------------------------------------------------
    // 1. Load the Excel workbook
    // -------------------------------------------------------------------------
    Source = Excel.Workbook(
        File.Contents(FilePath),
        null,
        true
    ),

    // -------------------------------------------------------------------------
    // 2. Select the "Telemetry" worksheet
    // -------------------------------------------------------------------------
    Telemetry_Sheet = Source{[Item = "Telemetry", Kind = "Sheet"]}[Data],

    // -------------------------------------------------------------------------
    // 3. Promote row 1 as column headers
    //    (Row 1 of the sheet contains the column names)
    // -------------------------------------------------------------------------
    PromotedHeaders = Table.PromoteHeaders(
        Telemetry_Sheet,
        [PromoteAllScalars = true]
    ),

    // -------------------------------------------------------------------------
    // 4. Remove the units / sub-header row (original row 2 of the sheet)
    // -------------------------------------------------------------------------
    RemovedUnitsRow = Table.Skip(PromotedHeaders, 1),

    // -------------------------------------------------------------------------
    // 5. Keep only the columns needed for the eight charts
    // -------------------------------------------------------------------------
    SelectedColumns = Table.SelectColumns(
        RemovedUnitsRow,
        {
            "LAP",
            "LAPTIME",
            "FUEL/LAP",
            "T wat, Max", "T wat, avg",
            "T oil, Max", "T oil, avg",
            "Poil, Min",  "Poil, avg",
            "Alt V, Min", "Alt V, avg",
            "Pfuel, Min", "Pfuel, avg",
            "LiftPump1 current, avg",
            "LiftPump2 current, avg"
        }
    ),

    // -------------------------------------------------------------------------
    // 6. Remove rows where LAP is null (empty trailing rows)
    // -------------------------------------------------------------------------
    RemovedNullLaps = Table.SelectRows(
        SelectedColumns,
        each [LAP] <> null and [LAP] <> ""
    ),

    // -------------------------------------------------------------------------
    // 7. Convert LAPTIME from MM.SS.mmm text to numeric seconds (SSS.mmm)
    //    Example: "1.23.456"  →  1×60 + 23 + 456/1000  =  83.456
    // -------------------------------------------------------------------------
    AddedLaptimeSeconds = Table.AddColumn(
        RemovedNullLaps,
        "LAPTIME_s",
        each
            let
                raw   = Text.Trim([LAPTIME]),
                parts = Text.Split(raw, ".")
            in
                if List.Count(parts) = 3 then
                    Number.FromText(parts{0}) * 60
                        + Number.FromText(parts{1})
                        + Number.FromText(parts{2}) / 1000
                else
                    error Error.Record(
                        "FormatError",
                        "LAPTIME value """ & raw & """ is not in MM.SS.mmm format"
                    ),
        type number
    ),

    // -------------------------------------------------------------------------
    // 8. Drop the original text LAPTIME column (replaced by LAPTIME_s)
    // -------------------------------------------------------------------------
    RemovedRawLaptime = Table.RemoveColumns(
        AddedLaptimeSeconds,
        {"LAPTIME"}
    ),

    // -------------------------------------------------------------------------
    // 9. Reorder columns so LAPTIME_s appears right after LAP
    // -------------------------------------------------------------------------
    ReorderedColumns = Table.ReorderColumns(
        RemovedRawLaptime,
        {
            "LAP",
            "LAPTIME_s",
            "FUEL/LAP",
            "T wat, Max", "T wat, avg",
            "T oil, Max", "T oil, avg",
            "Poil, Min",  "Poil, avg",
            "Alt V, Min", "Alt V, avg",
            "Pfuel, Min", "Pfuel, avg",
            "LiftPump1 current, avg",
            "LiftPump2 current, avg"
        }
    ),

    // -------------------------------------------------------------------------
    // 10. Set correct data types for every column
    // -------------------------------------------------------------------------
    TypedTable = Table.TransformColumnTypes(
        ReorderedColumns,
        {
            {"LAP",                    Int64.Type },
            {"LAPTIME_s",              type number},
            {"FUEL/LAP",               type number},
            {"T wat, Max",             type number},
            {"T wat, avg",             type number},
            {"T oil, Max",             type number},
            {"T oil, avg",             type number},
            {"Poil, Min",              type number},
            {"Poil, avg",              type number},
            {"Alt V, Min",             type number},
            {"Alt V, avg",             type number},
            {"Pfuel, Min",             type number},
            {"Pfuel, avg",             type number},
            {"LiftPump1 current, avg", type number},
            {"LiftPump2 current, avg", type number}
        }
    )

in
    TypedTable
