/*
  power_query.m
  -------------
  Complete Power Query M script for importing data/sample_telemetry.xlsx
  into Power BI Desktop.

  HOW TO USE
  ----------
  1. Open Power BI Desktop.
  2. Home → Get Data → Blank Query.
  3. In the Query Editor: Home → Advanced Editor.
  4. Replace all existing text with the contents of this file.
  5. Update the file path on the "Source" line to match the location of
     sample_telemetry.xlsx on your machine (use forward slashes or double
     backslashes on Windows, e.g. "C:/Projects/DashboardPowerBI/data/sample_telemetry.xlsx").
  6. Click Done, then Close & Apply.

  RESULT
  ------
  A table called "Telemetry" with 15 columns and one data row per lap:
    LAP, LAPTIME (raw text), LAPTIME_s (converted to seconds), FUEL/LAP,
    T wat Max/avg, T oil Max/avg, Poil Min/avg, Alt V Min/avg,
    Pfuel Min/avg, LiftPump1 current avg, LiftPump2 current avg.
*/

let
    // ── 1. Connect to the Excel workbook ──────────────────────────────────
    Source = Excel.Workbook(
        File.Contents("C:/REPLACE_THIS_PATH/data/sample_telemetry.xlsx"),
        null,
        true
    ),

    // ── 2. Navigate to the "Telemetry" sheet ──────────────────────────────
    Telemetry_Sheet = Source{[Item = "Telemetry", Kind = "Sheet"]}[Data],

    // ── 3. Promote row 1 as column headers ────────────────────────────────
    PromotedHeaders = Table.PromoteHeaders(
        Telemetry_Sheet,
        [PromoteAllScalars = true]
    ),

    // ── 4. Remove the units/sub-header row (original row 2) ───────────────
    RemovedUnitsRow = Table.Skip(PromotedHeaders, 1),

    // ── 5. Keep only the 15 meaningful telemetry columns ──────────────────
    //    (drops the two empty "reserved" columns B and C)
    SelectedColumns = Table.SelectColumns(
        RemovedUnitsRow,
        {
            "LAP",
            "LAPTIME",
            "FUEL/LAP",
            "T wat, Max",
            "T wat, avg",
            "T oil, Max",
            "T oil, avg",
            "Poil, Min",
            "Poil, avg",
            "Alt V, Min",
            "Alt V, avg",
            "Pfuel, Min",
            "Pfuel, avg",
            "LiftPump1 current, avg",
            "LiftPump2 current, avg"
        }
    ),

    // ── 6. Assign correct data types ──────────────────────────────────────
    TypedTable = Table.TransformColumnTypes(
        SelectedColumns,
        {
            {"LAP",                    Int64.Type},
            {"LAPTIME",                type text},
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
    ),

    // ── 7. Add LAPTIME_s: convert MM.SS.mmm → total seconds (SSS.mmm) ─────
    //    e.g. "1.23.456"  →  1×60 + 23 + 0.456  =  83.456 s
    //    Returns null for any value that does not match the expected format.
    AddedLaptimeS = Table.AddColumn(
        TypedTable,
        "LAPTIME_s",
        each
            let
                raw   = Text.Trim([LAPTIME]),
                parts = Text.Split(raw, ".")
            in
                if List.Count(parts) = 3 then
                    let
                        mm  = try Number.FromText(parts{0}) otherwise null,
                        ss  = try Number.FromText(parts{1}) otherwise null,
                        mmm = try Number.FromText(parts{2}) otherwise null
                    in
                        if mm <> null and ss <> null and mmm <> null
                        then mm * 60 + ss + mmm / 1000
                        else null
                else
                    null,
        type number
    )

in
    AddedLaptimeS
