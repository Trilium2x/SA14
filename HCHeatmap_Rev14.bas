Attribute VB_Name = "HCHeatmap"
Option Explicit

'====================================================================
' HC HEATMAP MODULE - Rev14
'
' NON-DESTRUCTIVE toggle between Gantt phase view and HC intensity view.
'
' ARCHITECTURE:
'   - CF rules are never saved to volatile memory. Instead, a persistent
'     "HC CF Definitions" sheet stores all Gantt CF definitions.
'   - When toggling TO HC view: CF is deleted, HC colors painted.
'   - When toggling BACK: HC colors cleared, CF rebuilt from the
'     definitions sheet (survives VBA reset, workbook save/reopen).
'   - Users can customize CF colors by editing the definitions sheet.
'
' USAGE:
'   1. Run CreateSegmentedToggle from the VBA editor.
'      Places a [Gantt | HC] segmented toggle on the ACTIVE sheet.
'
'   2. Click the HC segment to switch to HC view.
'      - Deletes Gantt CF rules
'      - Paints Gantt cell backgrounds based on NIF assignment dates:
'        * Cool-to-warm spectrum (1-5+) = minimum daily HC that week
'        * Magenta = at least one person assigned that week but
'          coverage drops to 0 on some day(s)
'      - Invalid NIF dates get yellow border (tracked for cleanup)
'      - Cell TEXT is never changed
'      - HC Legend displayed above NIF headers at row 13
'
'   3. Click the Gantt segment to restore original Gantt view.
'      - Clears HC painted colors
'      - Rebuilds CF from "HC CF Definitions" sheet
'      - Clears only the yellow borders that were added
'      - Removes HC legend
'
' Rev11 changes from Rev10:
'   - CRASH FIX: Chunked Union batching (flush every UNION_CHUNK_LIMIT
'     cells) prevents Excel crash from massive non-contiguous ranges
'   - CRASH FIX: lastDataRow detection uses ListObject boundary to
'     exclude HC tables below the main data (NIF_Builder puts 8 HC
'     tables below data; End(xlUp) would include them)
'   - PERF: Row-by-row border restoration replaced with bulk
'     xlInsideHorizontal on the entire Gantt range
'   - PERF: Hidden row state pre-read into Boolean array (eliminates
'     per-row COM calls in the hot loop)
'   - PERF: Single-pass CF rule application in RebuildCFFromDefinitions
'     (was 4 separate loops over the same array)
'   - FIX: ClearHCLegend accepts nifStartCol parameter instead of
'     redundantly re-detecting layout
'   - FIX: DetectNIFEmployeeCount bounded by lastCol to prevent
'     runaway scan on corrupted headers
'   - FIX: StoreYellowBorderCells guards against address string overflow
'   - FIX: ApplyTodayMarker cross-module call guarded with targeted
'     error handler (survives missing GanttBuilder module)
'   - CLEANUP: All Dim statements hoisted to procedure top per convention
'   - CLEANUP: Removed dead code (CLR_PURPLE, HCGreenColor)
'   - CLEANUP: Updated OnAction to HCHeatmap_Rev11
'
' PERFORMANCE:
'   - All NIF data and Gantt dates read into arrays (single bulk read)
'   - Hidden row state pre-read into Boolean array
'   - HC computed entirely in memory (zero cell reads in inner loop)
'   - Colors applied via chunked Union batching (flush every 400 cells)
'   - Border restoration via single bulk xlInsideHorizontal
'
' GUARANTEES:
'   - No cell content is ever changed
'   - No cells outside the Gantt data range are touched
'     (except yellow border on NIF name cells for invalid dates,
'      and HC legend at row 13 in the NIF column area)
'   - HC tables, NIF columns, slicers, formulas are never affected
'   - State survives VBA reset, workbook save/close/reopen
'   - Multiple Working Sheets operate independently (no shared state)
'====================================================================

' ---- Layout ----
Private Const DATA_START_ROW As Long = TIS_DATA_START_ROW
Private Const HC_LEGEND_ROW As Long = 13

' ---- CF Definitions sheet ----
Private Const HC_CF_SHEET As String = "HC CF Definitions"
Private Const CFD_FIRST_DATA_ROW As Long = 4
Private Const CFD_COL_ABBREV As Long = 1
Private Const CFD_COL_RULETYPE As Long = 2
Private Const CFD_COL_BG_R As Long = 3
Private Const CFD_COL_BG_G As Long = 4
Private Const CFD_COL_BG_B As Long = 5
Private Const CFD_COL_STOPIFTRUE As Long = 6
Private Const CFD_COL_PRIORITY As Long = 7
Private Const CFD_COL_PREVIEW As Long = 8

' ---- Toggle button shape names ----
Private Const TOGGLE_GANTT_NAME As String = "HCToggle_Gantt"
Private Const TOGGLE_HC_NAME As String = "HCToggle_HC"
Private Const OLD_TOGGLE_NAME As String = "HCHeatmapToggle"

' ---- Toggle colors ----
Private Const ACTIVE_GANTT_BG As Long = 16346203      ' RGB(91, 108, 249) = THEME_ACCENT
Private Const ACTIVE_HC_BG As Long = 10081076          ' RGB(52, 211, 153) = THEME_SUCCESS
Private Const INACTIVE_BG As Long = 4074026            ' RGB(42, 42, 62) = THEME_SURFACE
Private Const INACTIVE_FG As Long = 12100500           ' RGB(148, 163, 184) = THEME_TEXT_SEC

' ---- Union chunking (Rev11 crash fix) ----
Private Const UNION_CHUNK_LIMIT As Long = 400

' ---- Named range prefixes ----
Private Const NR_PREFIX_MODE As String = "HC_MODE_"
Private Const NR_PREFIX_YELLOW As String = "HC_YELLOW_"

' ---- Yellow border address limit ----
Private Const MAX_NAMED_RANGE_ADDR_LEN As Long = 200

'====================================================================
' PUBLIC ENTRY POINT 1: Create segmented toggle on ActiveSheet
'====================================================================

Public Sub CreateSegmentedToggle(Optional targetSheet As Worksheet = Nothing)
    Dim ws As Worksheet
    Dim ganttStartCol As Long
    Dim btnLeft As Double, btnTop As Double
    Dim segWidth As Double, segHeight As Double
    Dim shpLeft As Shape, shpRight As Shape
    Dim existLeft As Shape, existRight As Shape

    If Not targetSheet Is Nothing Then
        Set ws = targetSheet
    Else
        Set ws = ActiveSheet
    End If

    If Not ValidateWorkingSheet(ws) Then
        MsgBox "This sheet is not a valid Working Sheet with Gantt and NIF sections." & vbCrLf & _
               "Please run this on a Working Sheet.", vbExclamation, "HC Heatmap"
        Exit Sub
    End If

    ' Clean up old-style single button if present
    On Error Resume Next
    ws.Shapes(OLD_TOGGLE_NAME).Delete
    On Error GoTo 0

    ' Check if new toggle already exists
    On Error Resume Next
    Set existLeft = ws.Shapes(TOGGLE_GANTT_NAME)
    Set existRight = ws.Shapes(TOGGLE_HC_NAME)
    On Error GoTo 0

    If Not existLeft Is Nothing Or Not existRight Is Nothing Then
        ' Toggle already exists. When called programmatically (targetSheet provided),
        ' exit silently -- no MsgBox interrupting automated workflows like TIS load.
        ' When called interactively (no targetSheet), inform the user.
        If targetSheet Is Nothing Then
            MsgBox "Toggle already exists on '" & ws.Name & "'.", vbInformation, "HC Heatmap"
        End If
        Exit Sub
    End If

    ' Locate position
    ganttStartCol = FindGanttStartCol(ws)
    btnLeft = ws.Cells(9, ganttStartCol).Left
    btnTop = ws.Cells(9, ganttStartCol).Top
    segWidth = 65
    segHeight = 22

    ' ---- Left segment: "Gantt" ----
    Set shpLeft = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        btnLeft, btnTop, segWidth, segHeight)
    shpLeft.Name = TOGGLE_GANTT_NAME
    shpLeft.OnAction = "HCHeatmap.ToggleHCHeatmap"
    shpLeft.Placement = xlFreeFloating
    shpLeft.Shadow.Visible = msoFalse
    shpLeft.Line.Visible = msoFalse

    On Error Resume Next
    shpLeft.Adjustments(1) = 0.35
    On Error GoTo 0

    With shpLeft.TextFrame
        .Characters.Text = "Gantt"
        .Characters.Font.Name = "Segoe UI"
        .Characters.Font.Size = 9
        .Characters.Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .MarginLeft = 2: .MarginRight = 2
        .MarginTop = 1: .MarginBottom = 1
    End With

    ' ---- Right segment: "HC" ----
    Set shpRight = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        btnLeft + segWidth - 1, btnTop, segWidth, segHeight)
    shpRight.Name = TOGGLE_HC_NAME
    shpRight.OnAction = "HCHeatmap.ToggleHCHeatmap"
    shpRight.Placement = xlFreeFloating
    shpRight.Shadow.Visible = msoFalse
    shpRight.Line.Visible = msoFalse

    On Error Resume Next
    shpRight.Adjustments(1) = 0.35
    On Error GoTo 0

    With shpRight.TextFrame
        .Characters.Text = "HC"
        .Characters.Font.Name = "Segoe UI"
        .Characters.Font.Size = 9
        .Characters.Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .MarginLeft = 2: .MarginRight = 2
        .MarginTop = 1: .MarginBottom = 1
    End With

    ' ---- WhatIf toggle button (Rev14) ----
    Dim shpWIF As Shape
    Dim existWIF As Shape
    Set existWIF = Nothing
    On Error Resume Next
    Set existWIF = ws.Shapes("btn_WhatIf")
    On Error GoTo 0
    If existWIF Is Nothing Then
        Set shpWIF = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
            btnLeft + (segWidth * 2) + 20, btnTop, segWidth + 10, segHeight)
        shpWIF.Name = "btn_WhatIf"
        shpWIF.OnAction = "TIS_Launcher.ToggleWhatIf"
        shpWIF.Placement = xlFreeFloating
        shpWIF.Shadow.Visible = msoFalse
        shpWIF.Line.Visible = msoFalse
        On Error Resume Next
        shpWIF.Adjustments(1) = 0.35
        On Error GoTo 0
        With shpWIF.TextFrame
            .Characters.Text = "WhatIf"
            .Characters.Font.Name = THEME_FONT
            .Characters.Font.Size = 9
            .Characters.Font.Bold = True
            .Characters.Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .MarginLeft = 2: .MarginRight = 2
            .MarginTop = 1: .MarginBottom = 1
        End With
        shpWIF.Fill.ForeColor.RGB = RGB(255, 183, 77)  ' Amber
    End If

    ' Set initial state
    SetHCModeFlag ws, False
    UpdateSegmentedToggle ws, False

    ' Ensure CF definitions sheet exists
    Dim wsBeforeCF As Worksheet
    Set wsBeforeCF = ActiveSheet
    EnsureCFDefinitionsSheet
    ' Restore focus if CF sheet creation navigated away
    If Not ActiveSheet Is wsBeforeCF Then
        On Error Resume Next
        wsBeforeCF.Activate
        On Error GoTo 0
    End If

    DebugLog "HCHeatmap: segmented toggle + WhatIf created on '" & ws.Name & "'"
End Sub

'====================================================================
' PUBLIC ENTRY POINT 2: Toggle button click handler
'====================================================================

Public Sub ToggleHCHeatmap()
    Dim ws As Worksheet
    Dim appSt As AppState
    Dim startTime As Double

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    ' Validate sheet (Issue #18)
    If Not ValidateWorkingSheet(ws) Then
        MsgBox "This sheet is not a valid Working Sheet with Gantt and NIF sections." & vbCrLf & _
               "Please run this on a Working Sheet.", vbExclamation, "HC Heatmap"
        Exit Sub
    End If

    appSt = SaveAppState()
    SetPerformanceMode

    startTime = Timer

    On Error GoTo ErrorHandler

    If IsHCModeActive(ws) Then
        RestoreGanttView ws
        SetHCModeFlag ws, False
        UpdateSegmentedToggle ws, False
        DebugLog "HCHeatmap: -> GANTT view (" & Format(Timer - startTime, "0.00") & "s)"
    Else
        PaintHCHeatmap ws
        SetHCModeFlag ws, True
        UpdateSegmentedToggle ws, True
        DebugLog "HCHeatmap: -> HC view (" & Format(Timer - startTime, "0.00") & "s)"
    End If

    GoTo Cleanup

ErrorHandler:
    MsgBox "HC Heatmap error: " & Err.Description & vbCrLf & _
           "Error #" & Err.Number, vbCritical, "HC Heatmap"
    DebugLog "HCHeatmap ERROR: " & Err.Description & " (#" & Err.Number & ")"

Cleanup:
    RestoreAppState appSt
End Sub

'====================================================================
' PUBLIC ENTRY POINT 3: Ensure CF Definitions sheet exists
'====================================================================

Public Sub EnsureCFDefinitionsSheet()
    Dim wsCF As Worksheet
    Dim headers As Variant
    Dim c As Long
    Dim defaults As Variant
    Dim rowIdx As Long, d As Variant
    Dim bgR As Long, bgG As Long, bgB As Long, brightness As Long
    Dim ruleTypeRange As Range, stopRange As Range

    If SheetExists(ThisWorkbook, HC_CF_SHEET) Then Exit Sub

    Set wsCF = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsCF.Name = HC_CF_SHEET
    wsCF.Tab.Color = 16346203    ' THEME_ACCENT RGB(91, 108, 249)

    ' ---- Dark background ----
    wsCF.Cells.Interior.Color = 3022366   ' THEME_BG RGB(30, 30, 46)
    wsCF.Cells.Font.Color = 15788258      ' THEME_TEXT RGB(226, 232, 240)

    ' ---- Title ----
    With wsCF.Cells(1, 1)
        .Value = "Gantt CF Definitions"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = 15788258            ' THEME_TEXT
    End With

    ' ---- Instructions ----
    With wsCF.Cells(2, 1)
        .Value = "Customize phase colors below. Rows are applied as CF rules when toggling back from HC view."
        .Font.Size = 9
        .Font.Color = 12100500            ' THEME_TEXT_SEC
    End With
    wsCF.Range("A2:H2").Merge

    ' ---- Headers ----
    headers = Array("Abbreviation", "Rule Type", "BG Red", "BG Green", "BG Blue", _
                    "StopIfTrue", "Priority", "Preview")
    For c = 1 To 8
        With wsCF.Cells(3, c)
            .Value = headers(c - 1)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Font.Size = 9
            .Interior.Color = 4074026         ' THEME_SURFACE RGB(42, 42, 62)
            .HorizontalAlignment = xlCenter
        End With
    Next c

    ' ---- Default phase data (matches GanttBuilder exactly) ----
    ' Format: Abbreviation, RuleType, R, G, B, StopIfTrue, Priority
    defaults = Array( _
        Array("Reused", "reused", 192, 0, 0, False, 1), _
        Array("BOD", "bod", 0, 0, 0, True, 2), _
        Array("SDD", "sdd", 255, 255, 0, True, 900), _
        Array("SET", "phase", 56, 189, 248, False, 499), _
        Array("SL1", "phase", 74, 222, 128, False, 498), _
        Array("SL2", "phase", 250, 204, 21, False, 497), _
        Array("SQ", "phase", 99, 179, 237, False, 496), _
        Array("CV", "phase", 251, 146, 60, False, 899), _
        Array("PF", "phase", 148, 163, 184, False, 52), _
        Array("DC", "phase", 251, 191, 36, False, 51), _
        Array("DM", "phase", 239, 68, 68, False, 50), _
        Array("MRCL", "phase", 168, 130, 255, False, 40) _
    )

    rowIdx = CFD_FIRST_DATA_ROW

    For Each d In defaults
        wsCF.Cells(rowIdx, CFD_COL_ABBREV).Value = d(0)
        wsCF.Cells(rowIdx, CFD_COL_RULETYPE).Value = d(1)
        wsCF.Cells(rowIdx, CFD_COL_BG_R).Value = d(2)
        wsCF.Cells(rowIdx, CFD_COL_BG_G).Value = d(3)
        wsCF.Cells(rowIdx, CFD_COL_BG_B).Value = d(4)
        wsCF.Cells(rowIdx, CFD_COL_STOPIFTRUE).Value = d(5)
        wsCF.Cells(rowIdx, CFD_COL_PRIORITY).Value = d(6)

        ' Preview cell
        bgR = CLng(d(2)): bgG = CLng(d(3)): bgB = CLng(d(4))
        With wsCF.Cells(rowIdx, CFD_COL_PREVIEW)
            .Value = d(0)
            .Interior.Color = RGB(bgR, bgG, bgB)
            .Font.Bold = True
            .Font.Size = 9
            .HorizontalAlignment = xlCenter
            brightness = (bgR + bgG + bgB) \ 3
            If brightness < 160 Then
                .Font.Color = RGB(255, 255, 255)
            Else
                .Font.Color = RGB(40, 40, 40)
            End If
        End With

        rowIdx = rowIdx + 1
    Next d

    ' ---- Data validation: Rule Type ----
    Set ruleTypeRange = wsCF.Range(wsCF.Cells(CFD_FIRST_DATA_ROW, CFD_COL_RULETYPE), _
                                    wsCF.Cells(rowIdx + 10, CFD_COL_RULETYPE))
    With ruleTypeRange.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="phase,sdd,bod,reused"
        .ShowError = True
        .ErrorTitle = "Invalid Rule Type"
        .ErrorMessage = "Must be: phase, sdd, bod, or reused"
    End With

    ' ---- Data validation: StopIfTrue ----
    Set stopRange = wsCF.Range(wsCF.Cells(CFD_FIRST_DATA_ROW, CFD_COL_STOPIFTRUE), _
                                wsCF.Cells(rowIdx + 10, CFD_COL_STOPIFTRUE))
    With stopRange.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="TRUE,FALSE"
        .ShowError = True
    End With

    ' ---- Column widths ----
    wsCF.Columns(1).ColumnWidth = 14
    wsCF.Columns(2).ColumnWidth = 12
    wsCF.Columns(3).ColumnWidth = 8
    wsCF.Columns(4).ColumnWidth = 8
    wsCF.Columns(5).ColumnWidth = 8
    wsCF.Columns(6).ColumnWidth = 11
    wsCF.Columns(7).ColumnWidth = 9
    wsCF.Columns(8).ColumnWidth = 12

    DebugLog "HCHeatmap: CF Definitions sheet created with " & UBound(defaults) - LBound(defaults) + 1 & " defaults"

    ' Auto-sync any milestones from Definitions that aren't in defaults
    SyncMilestonesFromDefinitions
End Sub

'====================================================================
' SYNC MILESTONES FROM DEFINITIONS SHEET
' Scans Definitions Column F/G for milestone tokens and appends
' any missing abbreviations to the CF Definitions sheet.
'====================================================================

Private Sub SyncMilestonesFromDefinitions()
    Dim wsCF As Worksheet, wsDef As Worksheet
    Dim existingAbbrevs As Object
    Dim cfRow As Long
    Dim defLastRow As Long
    Dim defData As Variant
    Dim milNames As Object
    Dim i As Long, fText As String, gText As String
    Dim tokens As Variant, tokenVal As Variant
    Dim letter As String, num As Long
    Dim dispAbbrev As String
    Dim syncPalette As Variant
    Dim palIdx As Long
    Dim nextRow As Long
    Dim key As Variant
    Dim added As Long
    Dim bgR As Long, bgG As Long, bgB As Long, brightness As Long

    If Not SheetExists(ThisWorkbook, HC_CF_SHEET) Then Exit Sub
    If Not SheetExists(ThisWorkbook, "Definitions") Then Exit Sub

    Set wsCF = ThisWorkbook.Sheets(HC_CF_SHEET)
    Set wsDef = ThisWorkbook.Sheets("Definitions")

    ' Collect existing abbreviations from CF sheet
    Set existingAbbrevs = CreateObject("Scripting.Dictionary")
    cfRow = CFD_FIRST_DATA_ROW
    Do While Trim(CStr(wsCF.Cells(cfRow, CFD_COL_ABBREV).Value)) <> ""
        existingAbbrevs(UCase(Trim(CStr(wsCF.Cells(cfRow, CFD_COL_ABBREV).Value)))) = True
        cfRow = cfRow + 1
    Loop

    ' Parse Definitions Column F/G for milestone tokens
    defLastRow = wsDef.Cells(wsDef.Rows.Count, 1).End(xlUp).row
    If defLastRow < 2 Then Exit Sub

    defData = wsDef.Range(wsDef.Cells(1, 1), wsDef.Cells(defLastRow, 7)).Value

    Set milNames = CreateObject("Scripting.Dictionary")

    For i = 2 To UBound(defData, 1)
        fText = Trim(CStr(defData(i, 6)))
        gText = Trim(CStr(defData(i, 7)))
        If fText <> "" Then
            tokens = Split(fText, "|")
            For Each tokenVal In tokens
                tokenVal = UCase(Trim(CStr(tokenVal)))
                If Len(CStr(tokenVal)) >= 2 And IsNumeric(Mid(CStr(tokenVal), 2)) Then
                    letter = Left(CStr(tokenVal), 1)
                    num = CLng(Mid(CStr(tokenVal), 2))
                    If num = 1 And gText <> "" Then
                        dispAbbrev = UCase(GetLastWord(gText))
                        If dispAbbrev = "CONVERSION" Then dispAbbrev = "CV"
                        milNames(dispAbbrev) = True
                    End If
                End If
            Next tokenVal
        End If
    Next i

    ' Fallback colors for auto-synced milestones
    syncPalette = Array( _
        Array(0, 200, 180), _
        Array(200, 100, 220), _
        Array(255, 120, 150), _
        Array(100, 180, 100), _
        Array(220, 160, 80), _
        Array(120, 140, 220) _
    )
    palIdx = 0

    ' Append missing milestones
    nextRow = cfRow
    added = 0

    For Each key In milNames.keys
        If Not existingAbbrevs.exists(CStr(key)) Then
            bgR = syncPalette(palIdx Mod (UBound(syncPalette) + 1))(0)
            bgG = syncPalette(palIdx Mod (UBound(syncPalette) + 1))(1)
            bgB = syncPalette(palIdx Mod (UBound(syncPalette) + 1))(2)

            wsCF.Cells(nextRow, CFD_COL_ABBREV).Value = CStr(key)
            wsCF.Cells(nextRow, CFD_COL_RULETYPE).Value = "phase"
            wsCF.Cells(nextRow, CFD_COL_BG_R).Value = bgR
            wsCF.Cells(nextRow, CFD_COL_BG_G).Value = bgG
            wsCF.Cells(nextRow, CFD_COL_BG_B).Value = bgB
            wsCF.Cells(nextRow, CFD_COL_STOPIFTRUE).Value = False
            wsCF.Cells(nextRow, CFD_COL_PRIORITY).Value = 30 - added

            ' Preview cell
            With wsCF.Cells(nextRow, CFD_COL_PREVIEW)
                .Value = CStr(key)
                .Interior.Color = RGB(bgR, bgG, bgB)
                .Font.Bold = True
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                brightness = (bgR + bgG + bgB) \ 3
                If brightness < 160 Then
                    .Font.Color = RGB(255, 255, 255)
                Else
                    .Font.Color = RGB(40, 40, 40)
                End If
            End With

            nextRow = nextRow + 1
            palIdx = palIdx + 1
            added = added + 1
        End If
    Next key

    If added > 0 Then
        DebugLog "HCHeatmap: Auto-synced " & added & " milestone(s) from Definitions to CF sheet"
    End If
End Sub

'====================================================================
' PAINT HC HEATMAP (array-based, chunked Union batching)
'====================================================================

Private Sub PaintHCHeatmap(ws As Worksheet)
    On Error GoTo PaintError
    ' ---- All declarations at procedure top (Rev11 convention) ----
    Dim ganttStartCol As Long, ganttWeeks As Long, ganttEndCol As Long
    Dim nifStartCol As Long, nifEmployeeCount As Long
    Dim firstDataRow As Long, lastDataRow As Long
    Dim headerVals As Variant
    Dim weekDates() As Date
    Dim w As Long
    Dim nifEndCol As Long
    Dim nifData As Variant
    Dim rowCount As Long
    Dim ganttRange As Range
    Dim rngGreen1 As Range, rngGreen2 As Range, rngGreen3 As Range
    Dim rngGreen4 As Range, rngGreen5 As Range, rngPurple As Range
    Dim rngYellowBorder As Range
    Dim visibleCount As Long, assignCount As Long, invalidCount As Long
    Dim paintedGreen As Long, paintedPurple As Long
    Dim r As Long, i As Long, d As Long
    Dim nifStarts() As Date, nifEnds() As Date, nifValid() As Boolean
    Dim slotCount As Long
    Dim nameIdx As Long, startIdx As Long, endIdx As Long
    Dim nifNameVal As String
    Dim weekStart As Date, dayDate As Date
    Dim dailyCount As Long, minDaily As Long, maxDaily As Long
    Dim targetCell As Range
    Dim rowHidden() As Boolean
    Dim unionCount As Long

    ' ---- Auto-detect layout ----
    ganttStartCol = FindGanttStartCol(ws)
    If ganttStartCol = 0 Then Exit Sub

    ganttWeeks = DetectGanttWeeks(ws, ganttStartCol)
    ganttEndCol = ganttStartCol + ganttWeeks - 1

    nifStartCol = FindNIFStartCol(ws, ganttEndCol)
    If nifStartCol = 0 Then
        DebugLog "HCHeatmap: NIF columns not found"
        Exit Sub
    End If

    nifEmployeeCount = DetectNIFEmployeeCount(ws, nifStartCol)

    firstDataRow = DATA_START_ROW + 1

    ' ---- Rev11 FIX: Use ListObject boundary to exclude HC tables ----
    lastDataRow = FindDataLastRow(ws)
    If lastDataRow < firstDataRow Then Exit Sub

    ' ---- Read week dates from header row into array (Issue #15 fix) ----
    headerVals = ws.Range(ws.Cells(DATA_START_ROW, ganttStartCol), _
                           ws.Cells(DATA_START_ROW, ganttEndCol)).Value

    ReDim weekDates(1 To ganttWeeks)
    For w = 1 To ganttWeeks
        If IsDate(headerVals(1, w)) Then
            weekDates(w) = CDate(headerVals(1, w))
        ElseIf w > 1 Then
            weekDates(w) = weekDates(w - 1) + 7
        Else
            DebugLog "HCHeatmap: Gantt start date not found"
            Exit Sub
        End If
    Next w

    ' ---- Read all NIF data into array (single bulk read) ----
    nifEndCol = nifStartCol + nifEmployeeCount * 3 - 1
    nifData = ws.Range(ws.Cells(firstDataRow, nifStartCol), _
                        ws.Cells(lastDataRow, nifEndCol)).Value

    rowCount = lastDataRow - firstDataRow + 1

    DebugLog "HCHeatmap: ganttCol=" & ganttStartCol & "-" & ganttEndCol & _
                " (" & ganttWeeks & "wks) nifCol=" & nifStartCol & _
                " (" & nifEmployeeCount & " slots) rows=" & firstDataRow & "-" & lastDataRow

    ' ---- Rev11 PERF: Pre-read hidden row state into Boolean array ----
    ReDim rowHidden(1 To rowCount)
    For r = 1 To rowCount
        rowHidden(r) = ws.Rows(firstDataRow + r - 1).Hidden
    Next r

    ' ---- Delete CF and reset to dark Gantt background ----
    Set ganttRange = ws.Range(ws.Cells(firstDataRow, ganttStartCol), _
                               ws.Cells(lastDataRow, ganttEndCol))

    On Error GoTo PaintError

    ' Delete CF on the full gantt area INCLUDING header rows (CF may span headers+data)
    On Error Resume Next
    Dim fullGanttRange As Range
    Set fullGanttRange = ws.Range(ws.Cells(1, ganttStartCol), ws.Cells(lastDataRow, ganttEndCol))
    fullGanttRange.FormatConditions.Delete
    If Err.Number <> 0 Then
        DebugLog "HCHeatmap: Full CF delete failed: " & Err.Description & " - trying cells.delete"
        Err.Clear
        ' Fallback: clear CF from entire sheet columns in gantt area
        ws.Range(ws.Columns(ganttStartCol), ws.Columns(ganttEndCol)).FormatConditions.Delete
        If Err.Number <> 0 Then
            DebugLog "HCHeatmap: Column CF delete also failed: " & Err.Description
            Err.Clear
        End If
    End If
    On Error GoTo PaintError

    ganttRange.Interior.Color = THEME_BG
    ganttRange.Font.Color = THEME_TEXT

    ' ---- Init NIF slot arrays ----
    ReDim nifStarts(1 To nifEmployeeCount)
    ReDim nifEnds(1 To nifEmployeeCount)
    ReDim nifValid(1 To nifEmployeeCount)

    ' ---- Track Union accumulator sizes for chunked flushing (Rev11) ----
    unionCount = 0

    ' ---- Process each row ----
    For r = 1 To rowCount
        ' Skip hidden rows (Rev11: uses pre-read array, no COM call)
        If rowHidden(r) Then GoTo NextRow
        visibleCount = visibleCount + 1

        ' ---- Read NIF slots for this row from array ----
        slotCount = 0
        For i = 1 To nifEmployeeCount
            nifValid(i) = False
            nameIdx = (i - 1) * 3 + 1
            startIdx = nameIdx + 1
            endIdx = nameIdx + 2

            nifNameVal = Trim(CStr(nifData(r, nameIdx)))
            If nifNameVal = "" Then GoTo NextNIFSlot

            If Not IsDate(nifData(r, startIdx)) Or Not IsDate(nifData(r, endIdx)) Then
                ' Track yellow border cell
                BatchUnion rngYellowBorder, ws.Cells(firstDataRow + r - 1, nifStartCol + (i - 1) * 3)
                invalidCount = invalidCount + 1
                GoTo NextNIFSlot
            End If

            nifStarts(i) = CDate(nifData(r, startIdx))
            nifEnds(i) = CDate(nifData(r, endIdx))
            nifValid(i) = True
            slotCount = slotCount + 1
            assignCount = assignCount + 1
NextNIFSlot:
        Next i

        If slotCount = 0 Then GoTo NextRow

        ' ---- Compute min/max daily HC for each Gantt week ----
        For w = 1 To ganttWeeks
            weekStart = weekDates(w)
            maxDaily = 0
            minDaily = 9999

            For d = 0 To 6
                dayDate = weekStart + d
                dailyCount = 0
                For i = 1 To nifEmployeeCount
                    If nifValid(i) Then
                        If nifStarts(i) <= dayDate And nifEnds(i) >= dayDate Then
                            dailyCount = dailyCount + 1
                        End If
                    End If
                Next i
                If dailyCount > maxDaily Then maxDaily = dailyCount
                If dailyCount < minDaily Then minDaily = dailyCount
            Next d

            ' No one at all this week -> skip
            If maxDaily = 0 Then GoTo NextWeek

            Set targetCell = ws.Cells(firstDataRow + r - 1, ganttStartCol + w - 1)

            If minDaily = 0 Then
                ' Coverage gap: someone assigned but drops to 0 some day
                BatchUnion rngPurple, targetCell
                paintedPurple = paintedPurple + 1
            Else
                ' Full coverage: color by minimum HC count
                Select Case minDaily
                    Case 1: BatchUnion rngGreen1, targetCell
                    Case 2: BatchUnion rngGreen2, targetCell
                    Case 3: BatchUnion rngGreen3, targetCell
                    Case 4: BatchUnion rngGreen4, targetCell
                    Case Else: BatchUnion rngGreen5, targetCell
                End Select
                paintedGreen = paintedGreen + 1
            End If

            ' ---- Rev11 CRASH FIX: Flush Union accumulators periodically ----
            unionCount = unionCount + 1
            If unionCount >= UNION_CHUNK_LIMIT Then
                FlushHCColors rngGreen1, rngGreen2, rngGreen3, rngGreen4, rngGreen5, rngPurple
                unionCount = 0
            End If
NextWeek:
        Next w
NextRow:
    Next r

    ' ---- Final flush: apply remaining colors ----
    FlushHCColors rngGreen1, rngGreen2, rngGreen3, rngGreen4, rngGreen5, rngPurple

    ' ---- Apply yellow borders and store for targeted cleanup (Issue #8) ----
    If Not rngYellowBorder Is Nothing Then
        On Error Resume Next
        With rngYellowBorder.Borders
            .LineStyle = xlContinuous
            .Color = RGB(255, 200, 0)
            .Weight = xlMedium
        End With
        On Error GoTo 0
    End If
    StoreYellowBorderCells ws, rngYellowBorder

    ' ---- Create HC Legend ----
    CreateHCLegend ws, nifStartCol

    DebugLog "HCHeatmap: " & visibleCount & " visible, " & assignCount & " assignments, " & _
                invalidCount & " invalid, " & paintedGreen & " green, " & paintedPurple & " purple"
    Exit Sub
PaintError:
    ' Show detailed error info to help diagnose the 1004
    MsgBox "PaintHCHeatmap failed:" & vbCrLf & _
           "Error: " & Err.Description & " (#" & Err.Number & ")" & vbCrLf & _
           "ganttStartCol=" & ganttStartCol & " ganttEndCol=" & ganttEndCol & vbCrLf & _
           "firstDataRow=" & firstDataRow & " lastDataRow=" & lastDataRow & vbCrLf & _
           "nifStartCol=" & nifStartCol & " nifEmployees=" & nifEmployeeCount, _
           vbExclamation, "HCHeatmap Debug"
    DebugLog "HCHeatmap PaintHCHeatmap ERROR: " & Err.Description & " (#" & Err.Number & ")"
End Sub

'====================================================================
' FLUSH HC COLORS (Rev11 - chunked Union application)
' Applies accumulated colors and resets all accumulators to Nothing.
'====================================================================

Private Sub FlushHCColors(ByRef rng1 As Range, ByRef rng2 As Range, ByRef rng3 As Range, _
                           ByRef rng4 As Range, ByRef rng5 As Range, ByRef rngP As Range)
    ' Cool-to-warm spectrum on dark background
    ApplyHCColor rng1, RGB(25, 60, 120), RGB(120, 180, 255)     ' deep blue (HC=1)
    ApplyHCColor rng2, RGB(15, 95, 90), RGB(80, 220, 210)       ' teal (HC=2)
    ApplyHCColor rng3, RGB(20, 110, 40), RGB(130, 235, 150)     ' green (HC=3)
    ApplyHCColor rng4, RGB(140, 100, 10), RGB(255, 215, 80)     ' amber (HC=4)
    ApplyHCColor rng5, RGB(150, 45, 20), RGB(255, 170, 130)     ' warm red (HC=5+)
    ApplyHCColor rngP, RGB(120, 20, 80), RGB(255, 140, 200)     ' magenta (gap)

    ' Reset accumulators
    Set rng1 = Nothing
    Set rng2 = Nothing
    Set rng3 = Nothing
    Set rng4 = Nothing
    Set rng5 = Nothing
    Set rngP = Nothing
End Sub

'====================================================================
' RESTORE GANTT VIEW
'====================================================================

Private Sub RestoreGanttView(ws As Worksheet)
    On Error GoTo RestoreError
    Dim ganttStartCol As Long, ganttWeeks As Long, ganttEndCol As Long
    Dim firstDataRow As Long, lastDataRow As Long
    Dim ganttRange As Range
    Dim nifStartCol As Long

    ' ---- Auto-detect layout ----
    ganttStartCol = FindGanttStartCol(ws)
    If ganttStartCol = 0 Then Exit Sub

    ganttWeeks = DetectGanttWeeks(ws, ganttStartCol)
    ganttEndCol = ganttStartCol + ganttWeeks - 1

    firstDataRow = DATA_START_ROW + 1
    ' Rev11: Use ListObject boundary
    lastDataRow = FindDataLastRow(ws)

    Set ganttRange = ws.Range(ws.Cells(firstDataRow, ganttStartCol), _
                               ws.Cells(lastDataRow, ganttEndCol))

    ' ---- Reset to dark background (clearing HC paint) ----
    On Error Resume Next
    ganttRange.Interior.Color = THEME_BG
    ganttRange.Font.Color = THEME_TEXT
    On Error GoTo RestoreError

    ' Restore header rows: quarter (12), month (13), week dates (DATA_START_ROW)
    ws.Range(ws.Cells(12, ganttStartCol), ws.Cells(12, ganttEndCol)).Interior.Color = THEME_ACCENT
    ws.Range(ws.Cells(12, ganttStartCol), ws.Cells(12, ganttEndCol)).Font.Color = THEME_TEXT
    ws.Range(ws.Cells(13, ganttStartCol), ws.Cells(13, ganttEndCol)).Interior.Color = THEME_SURFACE
    ws.Range(ws.Cells(13, ganttStartCol), ws.Cells(13, ganttEndCol)).Font.Color = THEME_TEXT
    ws.Range(ws.Cells(DATA_START_ROW, ganttStartCol), ws.Cells(DATA_START_ROW, ganttEndCol)).Interior.Color = THEME_SURFACE
    ws.Range(ws.Cells(DATA_START_ROW, ganttStartCol), ws.Cells(DATA_START_ROW, ganttEndCol)).Font.Color = THEME_TEXT_SEC

    ' ---- Rev11 PERF: Bulk border restoration (single COM call) ----
    On Error Resume Next
    With ganttRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = THEME_BORDER
    End With
    On Error GoTo 0

    ' ---- Clear only tracked yellow borders (Issue #8 fix) ----
    ClearYellowBorderCells ws

    ' ---- Clear HC legend (Rev11: pass nifStartCol to avoid re-detection) ----
    nifStartCol = FindNIFStartCol(ws, ganttEndCol)
    ClearHCLegend ws, nifStartCol

    ' ---- Rebuild CF from definitions (Issues #1-7 fix) ----
    On Error Resume Next
    RebuildCFFromDefinitions ws, ganttRange
    If Err.Number <> 0 Then
        DebugLog "HCHeatmap: RebuildCFFromDefinitions failed: " & Err.Description & " (#" & Err.Number & ")"
        Err.Clear
    End If
    On Error GoTo 0

    ' ---- Re-apply today marker (guarded cross-module call) ----
    ' Use Application.Run to defer resolution to runtime (avoids compile error
    ' if GanttBuilder module is not loaded)
    On Error Resume Next
    Application.Run "ApplyTodayMarker", ws, ganttStartCol, lastDataRow
    If Err.Number <> 0 Then
        DebugLog "HCHeatmap: ApplyTodayMarker skipped (GanttBuilder not loaded?): " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    DebugLog "HCHeatmap: Gantt view restored"
    Exit Sub
RestoreError:
    DebugLog "HCHeatmap RestoreGanttView ERROR: " & Err.Description & " (#" & Err.Number & ")"
End Sub

'====================================================================
' REBUILD CF FROM DEFINITIONS SHEET (Rev11: single-pass application)
'====================================================================

Private Sub RebuildCFFromDefinitions(ws As Worksheet, ganttRange As Range)
    On Error GoTo RebuildCFError
    ' ---- All declarations at procedure top ----
    Dim wsCF As Worksheet
    Dim lastRow As Long
    Dim cfData As Variant
    Dim ruleCount As Long
    Dim firstDataRow As Long
    Dim ganttStartCol As Long
    Dim firstCellAddr As String
    Dim ganttColGap As Long, lastTableCol As Long
    Dim newReusedColIdx As Long, sddColIdx As Long
    Dim a As Long, b As Long, col As Long
    Dim tempVal As Variant
    Dim priA As Long, priB As Long
    Dim i As Long
    Dim abbrev As String, ruleType As String
    Dim bgR As Long, bgG As Long, bgB As Long
    Dim stopIfTrue As Boolean
    Dim brightness As Long
    Dim fontR As Long, fontG As Long, fontB As Long
    Dim fc As FormatCondition
    Dim formula As String
    Dim nrColLetter As String
    Dim sddColLetter As String, weekRefAddr As String, sddRowRef As String

    ' Ensure sheet exists (lazy creation)
    EnsureCFDefinitionsSheet
    ' Sync any new milestones from Definitions
    SyncMilestonesFromDefinitions

    Set wsCF = ThisWorkbook.Sheets(HC_CF_SHEET)

    ' Read definitions into array (single bulk read)
    lastRow = wsCF.Cells(wsCF.Rows.Count, CFD_COL_ABBREV).End(xlUp).row
    If lastRow < CFD_FIRST_DATA_ROW Then
        DebugLog "HCHeatmap: CF Definitions sheet is empty"
        Exit Sub
    End If

    cfData = wsCF.Range(wsCF.Cells(CFD_FIRST_DATA_ROW, 1), _
                         wsCF.Cells(lastRow, CFD_COL_PRIORITY)).Value

    ruleCount = UBound(cfData, 1)

    ' Delete existing CF
    ganttRange.FormatConditions.Delete

    ' First cell address for formula anchoring
    firstDataRow = ganttRange.row
    ganttStartCol = ganttRange.Column
    firstCellAddr = ganttRange.Cells(1, 1).Address(False, False)

    ' Find support columns (New/Reused, SDD)
    ganttColGap = 10
    lastTableCol = ganttStartCol - ganttColGap
    newReusedColIdx = FindColumnByHeader(ws, lastTableCol, "new/reused")
    sddColIdx = FindColumnByHeaderExact(ws, lastTableCol, "sdd")

    ' Sort rules by Priority ascending (bubble sort on in-memory array)
    For a = 1 To ruleCount - 1
        For b = a + 1 To ruleCount
            If IsNumeric(cfData(a, CFD_COL_PRIORITY)) Then priA = CLng(cfData(a, CFD_COL_PRIORITY)) Else priA = 999
            If IsNumeric(cfData(b, CFD_COL_PRIORITY)) Then priB = CLng(cfData(b, CFD_COL_PRIORITY)) Else priB = 999
            If priA > priB Then
                For col = 1 To CFD_COL_PRIORITY
                    tempVal = cfData(a, col)
                    cfData(a, col) = cfData(b, col)
                    cfData(b, col) = tempVal
                Next col
            End If
        Next b
    Next a

    ' ---- Rev11: Single-pass rule application (was 4 separate loops) ----
    ' Process "reused" first, then all others in priority order.
    ' Reused must be added first to get Priority=1 in the CF stack.

    For i = 1 To ruleCount
        ruleType = LCase(Trim(CStr(cfData(i, CFD_COL_RULETYPE))))
        abbrev = Trim(CStr(cfData(i, CFD_COL_ABBREV)))
        If abbrev = "" Then GoTo NextCFRule
        If Not SafeReadRGB(cfData, i, bgR, bgG, bgB) Then GoTo NextCFRule
        stopIfTrue = SafeReadBool(cfData(i, CFD_COL_STOPIFTRUE))

        ' Compute auto font color
        brightness = (bgR + bgG + bgB) \ 3
        If brightness < 160 Then
            fontR = 255: fontG = 255: fontB = 255
        Else
            fontR = 40: fontG = 40: fontB = 40
        End If

        formula = ""

        Select Case ruleType
            Case "reused"
                If newReusedColIdx > 0 Then
                    nrColLetter = "$" & ColLetter(newReusedColIdx)
                    formula = "=OR(" & nrColLetter & firstDataRow & "=""Reused""," & _
                             nrColLetter & firstDataRow & "=""Demo"")"
                End If

            Case "bod"
                formula = "=" & firstCellAddr & "=""BOD"""

            Case "sdd"
                If sddColIdx > 0 Then
                    sddColLetter = "$" & ColLetter(sddColIdx)
                    weekRefAddr = ws.Cells(DATA_START_ROW, ganttStartCol).Address(True, False)
                    sddRowRef = sddColLetter & firstDataRow
                    formula = "=AND(ISNUMBER(" & sddRowRef & ")," & _
                             sddRowRef & ">=" & weekRefAddr & "," & _
                             sddRowRef & "<=" & weekRefAddr & "+6)"
                End If

            Case "phase"
                formula = "=" & firstCellAddr & "=""" & abbrev & """"

            Case Else
                GoTo NextCFRule
        End Select

        If formula = "" Then GoTo NextCFRule

        On Error Resume Next
        Set fc = ganttRange.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
        If Err.Number = 0 Then
            ' Reused rule: font color only (no bg fill)
            If ruleType = "reused" Then
                fc.Font.Color = RGB(bgR, bgG, bgB)
                fc.Font.Bold = True
                fc.StopIfTrue = False
                fc.Priority = 1
            Else
                fc.Interior.Color = RGB(bgR, bgG, bgB)
                fc.Font.Color = RGB(fontR, fontG, fontB)
                fc.Font.Bold = True
                fc.StopIfTrue = stopIfTrue
                ' SDD gets highest priority
                If ruleType = "sdd" Then
                    fc.Priority = 1
                End If
            End If
        Else
            DebugLog "HCHeatmap: CF rule '" & abbrev & "' (" & ruleType & ") failed: " & Err.Description
        End If
        Err.Clear
        On Error GoTo 0
NextCFRule:
    Next i

    DebugLog "HCHeatmap: rebuilt CF from definitions (" & ruleCount & " rules processed)"
    Exit Sub
RebuildCFError:
    DebugLog "HCHeatmap RebuildCFFromDefinitions ERROR: " & Err.Description & " (#" & Err.Number & ")"
End Sub

'====================================================================
' SEGMENTED TOGGLE HELPERS
'====================================================================

Private Sub UpdateSegmentedToggle(ws As Worksheet, isHCMode As Boolean)
    Dim shpGantt As Shape, shpHC As Shape

    On Error GoTo ShapeNotFound
    Set shpGantt = ws.Shapes(TOGGLE_GANTT_NAME)
    Set shpHC = ws.Shapes(TOGGLE_HC_NAME)
    On Error GoTo 0

    If isHCMode Then
        ' HC active, Gantt inactive
        StyleToggleSegment shpGantt, INACTIVE_BG, INACTIVE_FG, False
        StyleToggleSegment shpHC, ACTIVE_HC_BG, vbWhite, True
    Else
        ' Gantt active, HC inactive
        StyleToggleSegment shpGantt, ACTIVE_GANTT_BG, vbWhite, True
        StyleToggleSegment shpHC, INACTIVE_BG, INACTIVE_FG, False
    End If
    Exit Sub

ShapeNotFound:
    DebugLog "HCHeatmap: toggle shapes not found on '" & ws.Name & "'"
End Sub

Private Sub StyleToggleSegment(shp As Shape, bgColor As Long, fgColor As Long, isBold As Boolean)
    shp.Fill.ForeColor.RGB = bgColor
    shp.Fill.Transparency = 0
    With shp.TextFrame
        .Characters.Font.Color = fgColor
        .Characters.Font.Bold = isBold
        .Characters.Font.Name = "Segoe UI"
        .Characters.Font.Size = 9
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
End Sub

'====================================================================
' HC LEGEND (row 13, above NIF headers)
'====================================================================

Private Sub CreateHCLegend(ws As Worksheet, nifStartCol As Long)
    Dim col As Long
    Dim greenColors As Variant, greenLabels As Variant, greenFonts As Variant
    Dim idx As Long

    col = nifStartCol

    ' Label
    With ws.Cells(HC_LEGEND_ROW, col)
        .Value = "HC Legend:"
        .Font.Bold = True
        .Font.Size = 7
        .Font.Color = 12100500            ' THEME_TEXT_SEC
        .HorizontalAlignment = xlLeft
    End With
    col = col + 1

    ' HC intensity cells (1-5+) -- cool-to-warm spectrum
    greenColors = Array(RGB(25, 60, 120), RGB(15, 95, 90), _
                        RGB(20, 110, 40), RGB(140, 100, 10), RGB(150, 45, 20))
    greenLabels = Array("1", "2", "3", "4", "5+")
    greenFonts = Array(RGB(120, 180, 255), RGB(80, 220, 210), RGB(130, 235, 150), _
                       RGB(255, 215, 80), RGB(255, 170, 130))

    For idx = 0 To 4
        With ws.Cells(HC_LEGEND_ROW, col)
            .Value = greenLabels(idx)
            .Interior.Color = greenColors(idx)
            .Font.Size = 6
            .Font.Bold = True
            .Font.Color = greenFonts(idx)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        col = col + 1
    Next idx

    ' Gap (magenta) cell -- dark theme
    With ws.Cells(HC_LEGEND_ROW, col)
        .Value = "Gap"
        .Interior.Color = RGB(120, 20, 80)
        .Font.Size = 6
        .Font.Bold = True
        .Font.Color = RGB(255, 140, 200)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub ClearHCLegend(ws As Worksheet, nifStartCol As Long)
    ' Rev11: accepts nifStartCol parameter (avoids redundant layout detection)
    If nifStartCol = 0 Then Exit Sub

    ' Clear 8 cells: label + 5 HC levels + 1 gap + 1 buffer
    ws.Range(ws.Cells(HC_LEGEND_ROW, nifStartCol), _
             ws.Cells(HC_LEGEND_ROW, nifStartCol + 7)).Clear
End Sub

'====================================================================
' AUTO-DETECTION FUNCTIONS
'====================================================================

Private Function DetectGanttWeeks(ws As Worksheet, ganttStartCol As Long) As Long
    ' Scan row 15 rightward from ganttStartCol until non-date (Issue #17 fix)
    Dim col As Long, cnt As Long
    Dim lastUsedCol As Long

    lastUsedCol = ws.Cells(DATA_START_ROW, ws.Columns.Count).End(xlToLeft).Column
    cnt = 0
    For col = ganttStartCol To lastUsedCol
        If IsDate(ws.Cells(DATA_START_ROW, col).Value) Then
            cnt = cnt + 1
        Else
            Exit For
        End If
    Next col

    If cnt = 0 Then cnt = 104  ' fallback
    DetectGanttWeeks = cnt
End Function

Private Function DetectNIFEmployeeCount(ws As Worksheet, nifStartCol As Long) As Long
    ' Scan headers for "NIF" + number pattern, each NIF = 3 columns (Issue #16 fix)
    ' Rev11: bounded by lastCol to prevent runaway scan
    Dim col As Long, cnt As Long, hv As String
    Dim lastCol As Long

    lastCol = ws.Cells(DATA_START_ROW, ws.Columns.Count).End(xlToLeft).Column
    cnt = 0
    col = nifStartCol

    Do While col <= lastCol
        hv = LCase(Trim(CStr(ws.Cells(DATA_START_ROW, col).Value)))
        If Len(hv) >= 4 And Left(hv, 3) = "nif" Then
            If IsNumeric(Mid(hv, 4)) Then
                cnt = cnt + 1
                col = col + 3  ' skip Start and End columns
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop

    If cnt = 0 Then cnt = 5  ' fallback
    DetectNIFEmployeeCount = cnt
End Function

'====================================================================
' DATA BOUNDARY DETECTION (Rev11 - excludes HC tables)
'====================================================================

Private Function FindDataLastRow(ws As Worksheet) As Long
    ' Uses the ListObject (table) on the Working Sheet to determine
    ' the actual data boundary, excluding HC tables below the data.
    ' Falls back to scanning column 1 for the last non-empty row
    ' within the data section (stops at first empty cell after data start).
    Dim lo As ListObject
    Dim r As Long

    ' Try ListObject first (most reliable)
    On Error Resume Next
    For Each lo In ws.ListObjects
        If lo.Range.row <= DATA_START_ROW + 1 Then
            FindDataLastRow = lo.Range.row + lo.Range.Rows.Count - 1
            On Error GoTo 0
            Exit Function
        End If
    Next lo
    On Error GoTo 0

    ' Fallback: scan down column 1 from first data row until empty
    ' This avoids picking up HC table content far below the data
    r = DATA_START_ROW + 1
    Do While r <= ws.Rows.Count
        If IsEmpty(ws.Cells(r, 1).Value) Then Exit Do
        r = r + 1
    Loop
    FindDataLastRow = r - 1

    ' Safety: if we got nothing, try End(xlUp) as last resort
    If FindDataLastRow < DATA_START_ROW + 1 Then
        FindDataLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    End If
End Function

'====================================================================
' COLUMN FINDER HELPERS
'====================================================================

Private Function FindGanttStartCol(ws As Worksheet) As Long
    ' Searches row 6 for GANTT_START marker
    Dim j As Long, lastCol As Long

    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    For j = 1 To lastCol
        If CStr(ws.Cells(6, j).Value) = "GANTT_START" Then
            FindGanttStartCol = j
            Exit Function
        End If
    Next j
    FindGanttStartCol = 0
End Function

Private Function FindNIFStartCol(ws As Worksheet, ganttEndCol As Long) As Long
    ' Searches row 15 after Gantt for "nif1" header
    Dim j As Long, hv As String
    Dim lastCol As Long

    lastCol = ws.Cells(DATA_START_ROW, ws.Columns.Count).End(xlToLeft).Column
    For j = ganttEndCol + 1 To lastCol
        hv = LCase(Trim(CStr(ws.Cells(DATA_START_ROW, j).Value)))
        If hv = "nif1" Then
            FindNIFStartCol = j
            Exit Function
        End If
    Next j
    FindNIFStartCol = 0
End Function

Private Function FindColumnByHeader(ws As Worksheet, maxCol As Long, searchText As String) As Long
    ' Fuzzy search in row 15 for a column header containing searchText
    Dim j As Long, hv As String

    FindColumnByHeader = 0
    For j = 1 To maxCol
        hv = LCase(Trim(Replace(Replace(CStr(ws.Cells(DATA_START_ROW, j).Value), vbLf, ""), vbCr, "")))
        If InStr(1, hv, LCase(searchText), vbTextCompare) > 0 Then
            FindColumnByHeader = j
            Exit Function
        End If
    Next j
End Function

Private Function FindColumnByHeaderExact(ws As Worksheet, maxCol As Long, searchText As String) As Long
    ' Exact match in row 15 for a column header
    Dim j As Long, hv As String

    FindColumnByHeaderExact = 0
    For j = 1 To maxCol
        hv = LCase(Trim(Replace(Replace(CStr(ws.Cells(DATA_START_ROW, j).Value), vbLf, ""), vbCr, "")))
        If hv = LCase(Trim(searchText)) Then
            FindColumnByHeaderExact = j
            Exit Function
        End If
    Next j
End Function

'====================================================================
' STATE MANAGEMENT (Named Ranges - persist across saves/reopen)
'====================================================================

Private Function IsHCModeActive(ws As Worksheet) As Boolean
    Dim nmName As String
    Dim nm As Name

    nmName = NR_PREFIX_MODE & SanitizeForNamedRange(ws.Name)

    On Error Resume Next
    Set nm = ThisWorkbook.Names(nmName)
    If Not nm Is Nothing Then
        IsHCModeActive = (nm.RefersTo = "=TRUE")
    Else
        IsHCModeActive = False
    End If
    On Error GoTo 0
End Function

Private Sub SetHCModeFlag(ws As Worksheet, active As Boolean)
    Dim nmName As String

    nmName = NR_PREFIX_MODE & SanitizeForNamedRange(ws.Name)

    On Error Resume Next
    ThisWorkbook.Names(nmName).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=nmName, RefersTo:="=" & UCase(CStr(active))
End Sub

'====================================================================
' YELLOW BORDER TRACKING (Issue #8 - targeted cleanup)
'====================================================================

Private Sub StoreYellowBorderCells(ws As Worksheet, rng As Range)
    Dim nmName As String
    Dim addrStr As String

    nmName = NR_PREFIX_YELLOW & SanitizeForNamedRange(ws.Name)

    ' Clean up any existing entry
    On Error Resume Next
    ThisWorkbook.Names(nmName).Delete
    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    ' Rev11 FIX: Guard against address string overflow
    addrStr = rng.Address(External:=True)
    If Len(addrStr) > MAX_NAMED_RANGE_ADDR_LEN Then
        DebugLog "HCHeatmap: Yellow border address too long (" & Len(addrStr) & _
                 " chars), truncating to first " & MAX_NAMED_RANGE_ADDR_LEN & " chars worth of cells"
        ' Store only the first area to avoid named range corruption
        ' The rest will be cleared by the full Gantt range reset on toggle-back
        On Error Resume Next
        addrStr = rng.Areas(1).Address(External:=True)
        On Error GoTo 0
    End If

    On Error Resume Next
    ThisWorkbook.Names.Add Name:=nmName, RefersTo:="=" & addrStr
    If Err.Number <> 0 Then
        DebugLog "HCHeatmap: Failed to store yellow borders: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub ClearYellowBorderCells(ws As Worksheet)
    Dim nmName As String
    Dim nm As Name
    Dim rng As Range

    nmName = NR_PREFIX_YELLOW & SanitizeForNamedRange(ws.Name)

    On Error Resume Next
    Set nm = ThisWorkbook.Names(nmName)
    If nm Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If

    Set rng = nm.RefersToRange
    On Error GoTo 0

    If Not rng Is Nothing Then
        rng.Borders.LineStyle = xlNone
    End If

    ' Remove the tracking named range
    On Error Resume Next
    nm.Delete
    On Error GoTo 0
End Sub

'====================================================================
' SHEET VALIDATION (Issue #18)
'====================================================================

Private Function ValidateWorkingSheet(ws As Worksheet) As Boolean
    Dim ganttStartCol As Long
    Dim ganttWeeks As Long, ganttEndCol As Long

    ValidateWorkingSheet = False

    ' Check 1: sheet name matches pattern
    If Not (ws.Name Like "Working Sheet*") Then Exit Function

    ' Check 2: GANTT_START marker exists
    ganttStartCol = FindGanttStartCol(ws)
    If ganttStartCol = 0 Then Exit Function

    ' Check 3: NIF columns exist
    ganttWeeks = DetectGanttWeeks(ws, ganttStartCol)
    ganttEndCol = ganttStartCol + ganttWeeks - 1
    If FindNIFStartCol(ws, ganttEndCol) = 0 Then Exit Function

    ValidateWorkingSheet = True
End Function

'====================================================================
' PAINTING HELPERS
'====================================================================

Private Sub BatchUnion(ByRef accumulator As Range, cell As Range)
    If accumulator Is Nothing Then
        Set accumulator = cell
    Else
        Set accumulator = Union(accumulator, cell)
    End If
End Sub

Private Sub ApplyHCColor(ByRef rng As Range, bgColor As Long, fontColor As Long)
    If rng Is Nothing Then Exit Sub
    rng.Interior.Color = bgColor
    rng.Font.Color = fontColor
End Sub

'====================================================================
' SAFE RGB READER (validates user-editable CF Definitions data)
'====================================================================

Private Function SafeReadBool(val As Variant) As Boolean
    Dim s As String

    On Error Resume Next
    If IsEmpty(val) Then SafeReadBool = False: Exit Function
    s = UCase(Trim(CStr(val)))
    If s = "TRUE" Or s = "YES" Or s = "1" Then
        SafeReadBool = True
    Else
        SafeReadBool = False
    End If
    On Error GoTo 0
End Function

Private Function SafeReadRGB(cfData As Variant, rowIdx As Long, _
                              ByRef outR As Long, ByRef outG As Long, ByRef outB As Long) As Boolean
    ' Returns False if any RGB value is invalid (non-numeric or out of range)
    SafeReadRGB = False

    If Not IsNumeric(cfData(rowIdx, CFD_COL_BG_R)) Then Exit Function
    If Not IsNumeric(cfData(rowIdx, CFD_COL_BG_G)) Then Exit Function
    If Not IsNumeric(cfData(rowIdx, CFD_COL_BG_B)) Then Exit Function

    outR = CLng(cfData(rowIdx, CFD_COL_BG_R))
    outG = CLng(cfData(rowIdx, CFD_COL_BG_G))
    outB = CLng(cfData(rowIdx, CFD_COL_BG_B))

    ' Clamp to valid RGB range
    If outR < 0 Then outR = 0
    If outR > 255 Then outR = 255
    If outG < 0 Then outG = 0
    If outG > 255 Then outG = 255
    If outB < 0 Then outB = 0
    If outB > 255 Then outB = 255

    SafeReadRGB = True
End Function
