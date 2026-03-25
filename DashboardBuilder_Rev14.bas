Attribute VB_Name = "DashboardBuilder"
'====================================================================
' Dashboard Builder Module - Rev14
'
' Creates a management-level Dashboard sheet with native Excel charts,
' KPI cards, drill-through navigation, and slicer-connected filtering.
'
' Theme: Applied Materials Brand - Deep navy base with AMAT Silver Lake Blue accents.
'   Primary brand color: RGB(86,156,190) #569CBE - AMAT Silver Lake Blue
'   Palette: Deep Navy bg, Steel Blue surfaces, Emerald success, Teal secondary,
'            Coral Red danger, Amber warning. Tables use ice-white/frost-blue rows.
'   Charts and KPI cards use THEME_SURFACE/THEME_BG dark styling.
'
' Rev11 changes from Rev10:
'   - KPI cards: Added Group dropdown filter, Conversions + Completed cards (9 total)
'   - System Counters: Reordered columns (New/Reused/Total | Demo/Conversions/Completed),
'     collapsible CEID rows via row grouping
'   - HC Gap Analysis: Adapted for NIF Rev11 structure (7 sub-rows per group),
'     removed System Type dropdown, 3 charts per group (Total/New/Reused),
'     deterministic chart naming (HC_*) for idempotent rebuild,
'     refactored into DiscoverHCTables/WriteHCGroupBlock/WriteHCTotalBlock helpers,
'     removed fixed 50-row scan cap, unicode markers in row labels
'   - Charts: Added chart start date selector
'   - Helper Table: Added Conversion column (11 cols), updated Completed logic
'   Rev14: Install Base (Total Tools) section with separate PivotCache,
'     GETPIVOTDATA data table, stacked column chart, 5 slicers, IB_BASELINE.
'     Helper table expanded to 15 cols (+MRCLFinish, InstallQtr, InstallDelta, HasSetStart)
'
' Sections:
'   1. Title Bar
'   2. KPI Cards (Total, New, Reused, Demo, CT Miss, Escalated, Watched, Conversions, Completed)
'   3. System Counters - Grouped by Group -> CEID drill-down (collapsible)
'   4. HC Gap Analysis - Per Group (7 rows: New/Reuse Need/Avail/Gap + Total Gap)
'      3 charts per group: Total | New | Reused Need vs Available
'      TOTAL section with same layout + executive-sized charts
'   5. Slicers (Group, CEID, Entity Type) - filter charts
'   6. Activity Graphs (Monthly stacked bar + Active Systems area)
'   7. Group Breakdown stacked bar
'   8. Escalation Tracker
'   9. Install Base (Total Tools) - Stacked column chart with IB_BASELINE,
'      GETPIVOTDATA data table, 5 independent slicers on separate PivotCache
'
' Public Subs:
'   BuildDashboard - Main entry point (called from Launcher or auto)
'   BuildInstallBaseSection - Install Base chart section (called from BuildDashboard)
'====================================================================

Option Explicit

' Layout constants
Private Const DASH_SHEET As String = "Dashboard"
Private Const DATA_START_ROW As Long = TIS_DATA_START_ROW

' Row anchors for each section
Private Const ROW_TITLE As Long = 1
Private Const ROW_KPI_LABEL As Long = 4
Private Const ROW_KPI_VALUE As Long = 5
Private Const ROW_DIVIDER As Long = 8

' Column layout
Private Const COL_START As Long = 2        ' Column B
Private Const COL_END As Long = 28         ' Column AB (expanded for 9 KPI cards)
Private Const CHART_HEIGHT_ROWS As Long = 12

' Applied Materials table colors
Private Const TABLE_HEADER_BG As Long = 3349260     ' Deep Navy RGB(12, 27, 51)
Private Const TABLE_ROW_BG As Long = 16578284       ' Ice White RGB(236, 246, 252)
Private Const TABLE_ALT_ROW_BG As Long = 16314076   ' Frost Blue RGB(220, 238, 248)
Private Const TABLE_TEXT As Long = 5587251           ' Slate Gray RGB(51, 65, 85)
Private Const TABLE_HEADER_TEXT As Long = 16777215   ' White

' Group row color (dark blue accent for group headers)
Private Const TABLE_GROUP_BG As Long = 6240018       ' Deep Blue RGB(18, 55, 95)

' HC Gap Analysis chart sizing
Private Const HC_CHART_WIDTH As Double = 205       ' Width per mini chart (px)
Private Const HC_CHART_HEIGHT As Double = 118      ' Height per mini chart (px)
Private Const HC_CHART_GAP As Double = 12          ' Horizontal gap between charts (px)
Private Const HC_TOT_CHART_WIDTH As Double = 220   ' TOTAL section chart width (px)
Private Const HC_TOT_CHART_HEIGHT As Double = 140  ' TOTAL section chart height (px)
Private Const HC_MONTHS As Long = 12               ' Months in HC rolling view

' HC table reference structure (populated by DiscoverHCTables)
Private Type HCTableRefs
    NewTitle As Long
    ReusedTitle As Long
    CombTitle As Long
    NewAvailTitle As Long
    ReuAvailTitle As Long
    NewGapTitle As Long
    ReuGapTitle As Long
    TotalGapTitle As Long
    gcNewNeed As Long
    gcReuNeed As Long
    gcComb As Long
    gcNewAvail As Long
    gcReuAvail As Long
    gcNewGap As Long
    gcReuGap As Long
    gcTotalGap As Long
    refTitle As Long
    refHdr As Long
    refGC As Long
    dsc As Long
    dec As Long
End Type

' Module-level state
Private m_ws As Worksheet               ' Dashboard sheet
Private m_workSheet As Worksheet         ' Working Sheet
Private m_tbl As ListObject              ' Working Sheet table
Private m_tblName As String              ' Table name
Private m_firstDataRow As Long           ' First data row on Working Sheet
Private m_lastDataRow As Long            ' Last data row
Private m_nrCol As Long                  ' New/Reused column index
Private m_groupCol As Long               ' Group column index
Private m_setStartCol As Long            ' Set Start column index
Private m_escCol As Long                 ' Escalated column index
Private m_ctCol As Long                  ' Est Cycle Time column index
Private m_entityCodeCol As Long          ' Entity Code column index
Private m_wsName As String               ' Working Sheet name (for formulas)
Private m_entityTypeCol As Long          ' Entity Type column index
Private m_ceidCol As Long                ' CEID column index
Private m_sqFinishCol As Long            ' Supplier Qual Finish column index
Private m_cvStartCol As Long             ' CV Start column index
Private m_statusCol As Long              ' Status column index
Private m_conversionCol As Long          ' Conversion column index (BP)

' Original header names (for PivotField references)
Private m_nrHeader As String
Private m_groupHeader As String
Private m_ceidHeader As String
Private m_entityTypeHeader As String
Private m_setStartHeader As String
Private m_entityCodeHeader As String

' Pivot infrastructure for slicer-connected charts
Private m_helperSheet As Worksheet
Private m_helperTable As ListObject       ' Helper table on DashHelper sheet
Private m_pivotCache As PivotCache
Private m_ptMonthly As PivotTable
Private m_ptActive As PivotTable       ' Active Systems PivotTable
Private m_dropdownCells As Collection  ' Cells with dropdowns (to unlock during protection)
Private m_ibPivotCache As PivotCache    ' Separate PivotCache for Install Base section
Private m_ptInstallBase As PivotTable   ' Install Base PivotTable

' Chart start date (default Jan 2026)
Private m_chartStartDate As Date

' Cached groups (avoid repeated full-column reads)
Private m_cachedGroups As Collection

' Cached group system counts: group name -> total (New + Reused) count (for sort order)
Private m_groupSysCounts As Object

' Has-date filter string (reused by KPIs, System Counters, Group Breakdown)
Private m_hasDateFilter As String

' Helper table column names (fixed)
Private Const HLP_COL_GROUP As String = "Group"
Private Const HLP_COL_CEID As String = "CEID"
Private Const HLP_COL_ENTITY_TYPE As String = "EntityType"
Private Const HLP_COL_NR As String = "NewReused"
Private Const HLP_COL_PROJSTART As String = "ProjectStart"
Private Const HLP_COL_PROJEND As String = "ProjectEnd"
Private Const HLP_COL_PROJMONTH As String = "ProjectMonth"
Private Const HLP_COL_STATUS As String = "Status"
Private Const HLP_COL_ROWTYPE As String = "RowType"
Private Const HLP_COL_PRESMONTH As String = "PresenceMonth"
Private Const HLP_COL_CONVERSION As String = "Conversion"
Private Const HLP_COL_MRCLFINISH As String = "MRCLFinish"
Private Const HLP_COL_INSTALLQTR As String = "InstallQtr"
Private Const HLP_COL_INSTALLDELTA As String = "InstallDelta"
Private Const HLP_COL_HASSETSTART As String = "HasSetStart"

' Dynamic row trackers (set during build)
Private m_nextRow As Long                ' Next available row during layout
Private m_chartSectionRow As Long        ' Row where chart section begins (for slicer placement)

'====================================================================
' MAIN ENTRY POINT
'====================================================================

Public Sub BuildDashboard(Optional silent As Boolean = False)
    On Error GoTo ErrorHandler

    Dim appSt As AppState
    Dim startTime As Double
    Dim currentSection As String   ' tracks which section we are in for error reporting

    appSt = SaveAppState()
    SetPerformanceMode
    Application.DisplayAlerts = False
    startTime = Timer

    ' Validate prerequisites
    Set m_workSheet = FindWorkingSheet()
    If m_workSheet Is Nothing Then
        If Not silent Then MsgBox "No Working Sheet found. Run Build Working Sheet first.", vbExclamation
        GoTo Cleanup
    End If

    Set m_tbl = Nothing
    If m_workSheet.ListObjects.Count > 0 Then
        Set m_tbl = m_workSheet.ListObjects(1)
    End If
    If m_tbl Is Nothing Then
        If Not silent Then MsgBox "No table found on Working Sheet.", vbExclamation
        GoTo Cleanup
    End If

    m_tblName = m_tbl.Name
    m_wsName = m_workSheet.Name
    m_firstDataRow = DATA_START_ROW + 1
    m_lastDataRow = m_tbl.Range.row + m_tbl.Range.Rows.Count - 1

    ' Discover column positions
    currentSection = "DiscoverColumns"
    DiscoverColumns

    ' Build hasDateFilter once (reused by KPIs, System Counters, Group Breakdown)
    currentSection = "BuildHasDateFilter"
    BuildHasDateFilter

    ' Create or clear Dashboard sheet
    currentSection = "CreateOrClearDashboardSheet"
    CreateOrClearDashboardSheet

    ' Set default chart start date
    m_chartStartDate = DateSerial(Year(Date), Month(Date), 1)

    ' Create pivot infrastructure (hidden helper sheet + PivotCache)
    currentSection = "Pivot Infrastructure"
    Application.StatusBar = "Building Dashboard... 10% - Pivot Infrastructure"
    CreatePivotInfrastructure

    ' Initialize layout tracker and dropdown cell tracker
    m_nextRow = ROW_TITLE
    Set m_dropdownCells = New Collection

    ' === Section 1: Title Bar ===
    currentSection = "Title Bar"
    Application.StatusBar = "Building Dashboard... 5% - Title Bar"
    BuildTitleBar

    ' === Section 2: KPI Cards ===
    currentSection = "KPI Cards"
    Application.StatusBar = "Building Dashboard... 15% - KPI Cards"
    BuildKPICards

    ' === Section 3: System Counters Per Group -> CEID ===
    currentSection = "System Counters"
    Application.StatusBar = "Building Dashboard... 30% - System Counters"
    BuildSystemCountersTable

    ' === Section 4: HC Gap Analysis ===
    currentSection = "HC Gap Analysis"
    Application.StatusBar = "Building Dashboard... 40% - HC Gap Analysis"
    BuildHCGapAnalysis

    ' === Section 5: Reserve space for Chart Filters (slicers placed last) ===
    currentSection = "Chart Filters"
    Application.StatusBar = "Building Dashboard... 45% - Chart Filters"
    m_chartSectionRow = m_nextRow + 1
    WriteSectionTitle m_chartSectionRow, "Chart Filters", "Use slicers below to filter chart data"
    m_nextRow = m_chartSectionRow + 12

    ' === Chart Date Selector ===
    currentSection = "Chart Date Selector"
    Dim chartDateRow As Long: chartDateRow = m_nextRow + 1
    m_ws.Range(m_ws.Cells(chartDateRow, COL_START), m_ws.Cells(chartDateRow, COL_START + 1)).Merge
    m_ws.Cells(chartDateRow, COL_START).Value = "Chart Start:"
    m_ws.Cells(chartDateRow, COL_START).Font.Bold = True
    m_ws.Cells(chartDateRow, COL_START).Font.Size = 11
    m_ws.Cells(chartDateRow, COL_START).Font.Color = RGB(100, 116, 139)

    ' Merge date cell for wider visibility
    Dim chartDateMerge As Range
    Set chartDateMerge = m_ws.Range(m_ws.Cells(chartDateRow, COL_START + 2), _
        m_ws.Cells(chartDateRow, COL_START + 3))
    chartDateMerge.Merge

    Dim chartDateCell As Range
    Set chartDateCell = m_ws.Cells(chartDateRow, COL_START + 2)
    chartDateCell.Value = m_chartStartDate
    chartDateCell.NumberFormat = "mmm yyyy"
    chartDateCell.Font.Bold = True
    chartDateCell.Font.Size = 11
    chartDateCell.Font.Color = RGB(30, 41, 59)
    chartDateCell.Interior.Color = RGB(255, 255, 255)
    chartDateCell.HorizontalAlignment = xlCenter
    With chartDateMerge.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(86, 156, 190)
    End With
    With chartDateMerge.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225)
    End With
    With chartDateMerge.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225)
    End With
    With chartDateMerge.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225)
    End With
    With chartDateCell.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
             Operator:=xlGreater, Formula1:="1/1/2020"
    End With
    If Not m_dropdownCells Is Nothing Then m_dropdownCells.Add chartDateCell

    ' Store chart date cell address in a named range for Worksheet_Change handler
    On Error Resume Next
    ThisWorkbook.Names("DASH_CHART_START").Delete
    ThisWorkbook.Names.Add Name:="DASH_CHART_START", RefersTo:=chartDateCell
    On Error GoTo 0

    m_nextRow = chartDateRow + 1

    ' === Section 6: Activity Graphs (PivotCharts) ===
    currentSection = "Monthly Activity Chart"
    Application.StatusBar = "Building Dashboard... 55% - Monthly Activity"
    BuildMonthlyActivityChart

    currentSection = "Active Systems Chart"
    Application.StatusBar = "Building Dashboard... 65% - Active Systems"
    BuildActiveSystemsChart

    ' === Section 7: Group Breakdown (Horizontal Bar Chart) ===
    currentSection = "Group Breakdown"
    Application.StatusBar = "Building Dashboard... 75% - Group Breakdown"
    BuildGroupBreakdown

    ' === Section 8: Escalation Tracker ===
    currentSection = "Escalation Tracker"
    Application.StatusBar = "Building Dashboard... 85% - Escalation Tracker"
    BuildEscalationTracker

    ' === Section 9: Install Base (Total Tools) ===
    currentSection = "Install Base"
    Application.StatusBar = "Building Dashboard... 85% - Install Base"
    BuildInstallBaseSection m_ws, m_workSheet, m_helperSheet, m_helperTable

    ' === Section 10: Slicers (created after PivotTables exist) ===
    currentSection = "Slicers"
    Application.StatusBar = "Building Dashboard... 90% - Adding Slicers"
    BuildDashboardSlicers

    ' Final formatting
    currentSection = "Final Formatting"
    Application.StatusBar = "Building Dashboard... 95% - Formatting"
    ApplyDashboardFormatting
    AddNavigationLinks
    ProtectDashboardSheet

    ' Inject Worksheet_Change event handler for chart start date reactivity
    currentSection = "Change Handler Injection"
    InjectDashboardChangeHandler

    ' Recalc
    Application.Calculation = xlCalculationAutomatic
    m_ws.Calculate

    If Not silent Then
        Application.ScreenUpdating = True
        MsgBox "Dashboard built successfully!" & vbCrLf & _
               "Time: " & Format(Timer - startTime, "0.00") & "s", vbInformation
        Application.ScreenUpdating = False
    End If

    GoTo Cleanup

ErrorHandler:
    If Not silent Then
        MsgBox "Error in BuildDashboard [" & currentSection & "]: " & Err.Description & vbCrLf & _
               "Error #: " & Err.Number, vbCritical
    End If
    DebugLog "DashboardBuilder ERROR [" & currentSection & "]: " & Err.Description & " (#" & Err.Number & ")"

Cleanup:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    RestoreAppState appSt
    ' Force automatic calculation â€” Dashboard formulas require it
    ' (RestoreAppState may restore manual mode from user's prior setting)
    Application.Calculation = xlCalculationAutomatic
    Set m_ws = Nothing
    Set m_workSheet = Nothing
    Set m_tbl = Nothing
    Set m_helperSheet = Nothing
    Set m_helperTable = Nothing
    Set m_pivotCache = Nothing
    Set m_ptMonthly = Nothing
    Set m_ptActive = Nothing
    Set m_dropdownCells = Nothing
    Set m_ibPivotCache = Nothing
    Set m_ptInstallBase = Nothing
    Set m_cachedGroups = Nothing
    Set m_groupSysCounts = Nothing
    m_hasDateFilter = ""
End Sub

'====================================================================
' DISCOVER COLUMNS on Working Sheet
'====================================================================

Private Sub DiscoverColumns()
    Dim j As Long, hv As String
    Dim lastCol As Long

    m_nrCol = 0: m_groupCol = 0: m_setStartCol = 0: m_escCol = 0
    m_ctCol = 0: m_entityCodeCol = 0
    m_entityTypeCol = 0: m_ceidCol = 0: m_sqFinishCol = 0
    m_cvStartCol = 0: m_statusCol = 0: m_conversionCol = 0

    lastCol = m_workSheet.Cells(DATA_START_ROW, m_workSheet.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Sub

    ' Read header row into array (single Range.Value read)
    Dim hdrArr() As Variant
    hdrArr = m_workSheet.Range(m_workSheet.Cells(DATA_START_ROW, 1), _
        m_workSheet.Cells(DATA_START_ROW, lastCol)).Value

    Dim rawH As String
    For j = 1 To lastCol
        rawH = Trim(Replace(Replace(CStr(hdrArr(1, j)), vbLf, ""), vbCr, ""))
        hv = LCase(rawH)
        Select Case hv
            Case "new/reused", "new/reused/demo", "new-reused": m_nrCol = j: m_nrHeader = rawH
            Case "group": m_groupCol = j: m_groupHeader = rawH
            Case "set start": m_setStartCol = j: m_setStartHeader = rawH
            Case "escalated": m_escCol = j
            Case "est cycle time", "estcycle time": m_ctCol = j
            Case "entity code", "entitycode": m_entityCodeCol = j: m_entityCodeHeader = rawH
            Case "entity type", "entitytype": m_entityTypeCol = j: m_entityTypeHeader = rawH
            Case "ceid": m_ceidCol = j: m_ceidHeader = rawH
            Case "supplier qual finish", "supplier qualfinish": m_sqFinishCol = j
            Case "cv start", "convert start", "conversion start", "cvstart", "convertstart", "conversionstart": m_cvStartCol = j
            Case LCase(TIS_COL_STATUS): m_statusCol = j
            Case "conversion": m_conversionCol = j
        End Select
    Next j
End Sub

'====================================================================
' BUILD HAS-DATE FILTER
' Constructs a SUMPRODUCT-compatible filter string that excludes
' projects with zero milestone dates. Uses Definitions milestone
' columns (same source as Working Sheet tfFilter).
' Stored in m_hasDateFilter for reuse across sections.
'====================================================================

Private Sub BuildHasDateFilter()
    m_hasDateFilter = ""
    If m_tbl Is Nothing Then Exit Sub

    Dim msHeaders As Collection
    Set msHeaders = GetMilestoneStartHeaders()

    Dim lastCol As Long
    lastCol = m_workSheet.Cells(DATA_START_ROW, m_workSheet.Columns.Count).End(xlToLeft).Column
    Dim hdrArr() As Variant
    hdrArr = m_workSheet.Range(m_workSheet.Cells(DATA_START_ROW, 1), _
        m_workSheet.Cells(DATA_START_ROW, lastCol)).Value

    Dim hasDateParts As String: hasDateParts = ""
    Dim msItem As Variant, msName As String, mj As Long, whv As String
    For Each msItem In msHeaders
        msName = CStr(msItem)
        If LCase(msName) = "sdd" Then GoTo NextMsBDF
        If InStr(1, LCase(msName), "prefac", vbTextCompare) > 0 Then GoTo NextMsBDF
        If InStr(1, LCase(msName), "pre-fac", vbTextCompare) > 0 Then GoTo NextMsBDF
        For mj = 1 To lastCol
            whv = Trim(Replace(Replace(CStr(hdrArr(1, mj)), vbLf, ""), vbCr, ""))
            If StrComp(whv, msName, vbTextCompare) = 0 Then
                Dim colRef As String
                colRef = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, mj).Value & "]"
                If hasDateParts <> "" Then hasDateParts = hasDateParts & "+"
                hasDateParts = hasDateParts & "ISNUMBER(" & colRef & ")"
                Exit For
            End If
        Next mj
NextMsBDF:
    Next msItem

    ' Ensure Set Start is included
    If m_setStartCol > 0 Then
        Dim ssRef As String
        ssRef = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_setStartCol).Value & "]"
        If InStr(1, hasDateParts, m_workSheet.Cells(DATA_START_ROW, m_setStartCol).Value) = 0 Then
            If hasDateParts <> "" Then hasDateParts = hasDateParts & "+"
            hasDateParts = hasDateParts & "ISNUMBER(" & ssRef & ")"
        End If
    End If

    If hasDateParts <> "" Then
        m_hasDateFilter = "*((" & hasDateParts & ")>0)"
    End If
End Sub

'====================================================================
' CREATE OR CLEAR DASHBOARD SHEET
'====================================================================

Private Sub CreateOrClearDashboardSheet()
    Application.DisplayAlerts = False
    On Error Resume Next

    ' 1) Remove only slicer caches whose PivotTable lives on DashHelper
    '    or whose slicers are placed on Dashboard (preserves Working Sheet slicers)
    Dim scIdx As Long
    For scIdx = ThisWorkbook.SlicerCaches.Count To 1 Step -1
        Dim shouldDel As Boolean
        shouldDel = False
        On Error Resume Next
        ' Check if source PivotTable is on DashHelper
        Dim scPtSheet As String
        scPtSheet = ThisWorkbook.SlicerCaches(scIdx).PivotTable.Parent.Name
        If Err.Number = 0 Then
            If scPtSheet = "DashHelper" Then shouldDel = True
        End If
        Err.Clear
        ' Check if any slicer in this cache is placed on the Dashboard
        Dim sl As Slicer
        For Each sl In ThisWorkbook.SlicerCaches(scIdx).Slicers
            If sl.Parent.Name = DASH_SHEET Then shouldDel = True
        Next sl
        On Error GoTo 0
        If shouldDel Then
            On Error Resume Next
            ThisWorkbook.SlicerCaches(scIdx).Delete
            On Error GoTo 0
        End If
    Next scIdx

    ' 2) Remove PivotTables on old DashHelper (they hold PivotCache refs)
    If SheetExists(ThisWorkbook, "DashHelper") Then
        Dim oldHelper As Worksheet
        Set oldHelper = ThisWorkbook.Sheets("DashHelper")
        oldHelper.Visible = xlSheetVisible  ' must be visible to manipulate
        Dim ptIdx As Long
        For ptIdx = oldHelper.PivotTables.Count To 1 Step -1
            On Error Resume Next
            oldHelper.PivotTables(ptIdx).TableRange2.Clear
            On Error GoTo 0
        Next ptIdx
        ' Remove ListObjects (breaks source links for PivotCaches)
        Dim lo As ListObject
        For Each lo In oldHelper.ListObjects
            On Error Resume Next
            lo.Delete
            On Error GoTo 0
        Next lo
    End If

    ' 3) Clear orphaned PivotCaches
    On Error Resume Next
    Dim pc As PivotCache
    For Each pc In ThisWorkbook.PivotCaches
        pc.MissingItemsLimit = xlMissingItemsNone
    Next pc
    On Error GoTo 0

    ' 4) Delete old Dashboard sheet (PivotCharts live here)
    '    Ensure at least one other visible sheet exists before deleting
    If SheetExists(ThisWorkbook, DASH_SHEET) Then
        ' Make sure Working Sheet is visible (safety net so we never delete the only visible sheet)
        On Error Resume Next
        m_workSheet.Visible = xlSheetVisible
        On Error GoTo 0
        Application.DisplayAlerts = False
        On Error Resume Next
        ThisWorkbook.Sheets(DASH_SHEET).Delete
        On Error GoTo 0
    End If

    ' 5) Delete old DashHelper sheet (now safe â€” no refs remain)
    If SheetExists(ThisWorkbook, "DashHelper") Then
        Application.DisplayAlerts = False
        On Error Resume Next
        ThisWorkbook.Sheets("DashHelper").Delete
        On Error GoTo 0
    End If

    Application.DisplayAlerts = True

    ' 6) Create new Dashboard sheet
    '    If old Dashboard still exists (deletion failed), clear it instead
    If SheetExists(ThisWorkbook, DASH_SHEET) Then
        Set m_ws = ThisWorkbook.Sheets(DASH_SHEET)
        m_ws.Cells.Clear
        ' Remove all chart objects
        Dim co As ChartObject
        For Each co In m_ws.ChartObjects
            On Error Resume Next
            co.Delete
            On Error GoTo 0
        Next co
        ' Remove all shapes
        Dim shp As Shape
        For Each shp In m_ws.Shapes
            On Error Resume Next
            shp.Delete
            On Error GoTo 0
        Next shp
        ' Ungroup all rows/columns
        On Error Resume Next
        m_ws.Cells.EntireRow.Ungroup
        m_ws.Cells.EntireColumn.Ungroup
        On Error GoTo 0
    Else
        Set m_ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        m_ws.Name = DASH_SHEET
    End If

    m_ws.Tab.Color = THEME_ACCENT

    ' Set column widths (wider default to accommodate "Conversions", "Completed", etc.)
    m_ws.Columns("A").ColumnWidth = 2
    Dim c As Long
    For c = COL_START To COL_END
        m_ws.Columns(c).ColumnWidth = 11
    Next c
    ' Widen label column for HC Gap labels ("New Available", "Reuse Available", etc.)
    m_ws.Columns(COL_START).ColumnWidth = 18

    ' Light theme base: use Excel's default white background
    ' Only set font — no blanket Interior.Color (saves formatting 17B cells)
    m_ws.Cells.Font.Name = THEME_FONT
    m_ws.Cells.Font.Color = RGB(30, 41, 59)   ' Slate-900: near-black for readability
End Sub

'====================================================================
' SECTION 1: TITLE BAR
'====================================================================

Private Sub BuildTitleBar()
    With m_ws.Range(m_ws.Cells(ROW_TITLE, COL_START), m_ws.Cells(ROW_TITLE, COL_END))
        .Merge
        .Value = "TIS Tracker Dashboard"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = THEME_WHITE
        .Interior.Color = RGB(12, 27, 51)    ' Deep Navy
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    m_ws.Rows(ROW_TITLE).RowHeight = 42

    ' Subtitle with timestamp (light gray bar below dark hero)
    With m_ws.Range(m_ws.Cells(ROW_TITLE + 1, COL_START), m_ws.Cells(ROW_TITLE + 1, COL_END))
        .Merge
        .Value = TIS_VERSION & "  |  " & Format(Now, "mm/dd/yyyy hh:mm") & "  |  Source: " & m_wsName
        .Font.Size = 9
        .Font.Color = RGB(100, 116, 139)
        .Interior.Color = RGB(241, 245, 249)
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    m_ws.Rows(ROW_TITLE + 1).RowHeight = 22

    ' Accent line below subtitle
    With m_ws.Range(m_ws.Cells(ROW_TITLE + 1, COL_START), m_ws.Cells(ROW_TITLE + 1, COL_END))
        .Borders(xlEdgeBottom).Color = RGB(86, 156, 190)  ' Brand blue accent line
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

    ' Spacer row
    m_ws.Rows(ROW_TITLE + 2).RowHeight = 6
End Sub

'====================================================================
' SECTION 2: KPI CARDS (9 cards with Group dropdown filter)
'====================================================================

Private Sub BuildKPICards()
    Dim currentCol As Long

    ' --- Build structured table references (auto-adjust to table size) ---
    ' Raw header values include vbLf; Excel handles them in structured refs
    Dim nrRef As String, grpRef As String, ctRef As String, escRef As String
    Dim cvRef As String, statusRef As String

    If m_nrCol > 0 Then nrRef = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_nrCol).Value & "]"
    If m_groupCol > 0 Then grpRef = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_groupCol).Value & "]"
    If m_ctCol > 0 Then ctRef = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_ctCol).Value & "]"
    If m_escCol > 0 Then escRef = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_escCol).Value & "]"
    If m_cvStartCol > 0 Then cvRef = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_cvStartCol).Value & "]"
    If m_statusCol > 0 Then statusRef = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_statusCol).Value & "]"

    ' Use pre-computed hasDateFilter from BuildHasDateFilter
    Dim hasDateFilter As String
    hasDateFilter = m_hasDateFilter

    ' --- Group Filter Dropdown for KPIs (placed at B7:C7) ---
    Dim kpiGrpDropRow As Long: kpiGrpDropRow = ROW_DIVIDER - 1

    m_ws.Cells(kpiGrpDropRow, COL_START).Value = "Group Filter:"
    m_ws.Cells(kpiGrpDropRow, COL_START).Font.Bold = True
    m_ws.Cells(kpiGrpDropRow, COL_START).Font.Size = 11
    m_ws.Cells(kpiGrpDropRow, COL_START).Font.Color = RGB(100, 116, 139)

    Dim kpiGroups As Collection
    Set kpiGroups = CollectGroups()
    Dim grpList As String: grpList = "All"
    Dim kgi As Long
    For kgi = 1 To kpiGroups.Count
        grpList = grpList & "," & CStr(kpiGroups(kgi))
    Next kgi

    Dim kpiGrpMerge As Range
    Set kpiGrpMerge = m_ws.Range(m_ws.Cells(kpiGrpDropRow, COL_START + 1), _
        m_ws.Cells(kpiGrpDropRow, COL_START + 2))
    kpiGrpMerge.Merge

    Dim kpiGrpCell As Range
    Set kpiGrpCell = m_ws.Cells(kpiGrpDropRow, COL_START + 1)
    With kpiGrpCell
        .Value = "All"
        .Font.Bold = True: .Font.Size = 11
        .Font.Color = RGB(30, 41, 59)
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        With kpiGrpMerge.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(86, 156, 190)
        End With
        With kpiGrpMerge.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225)
        End With
        With kpiGrpMerge.Borders(xlEdgeRight)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225)
        End With
        With kpiGrpMerge.Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225)
        End With
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=grpList
            .IgnoreBlank = True: .InCellDropdown = True
        End With
    End With
    If Not m_dropdownCells Is Nothing Then m_dropdownCells.Add kpiGrpCell

    On Error Resume Next
    ThisWorkbook.Names("DASH_KPI_GROUP").Delete
    ThisWorkbook.Names.Add Name:="DASH_KPI_GROUP", RefersTo:=kpiGrpCell
    On Error GoTo 0

    Dim kpiGrpRef As String
    kpiGrpRef = "$" & ColLetter(COL_START + 1) & "$" & kpiGrpDropRow

    ' --- Helper: group filter expression for SUMPRODUCT ---
    ' When group = "All", this evaluates to 1 (no filter). Otherwise filters to group.
    Dim gfAll As String, gfGrp As String
    If m_groupCol > 0 Then
        gfAll = ""
        gfGrp = "*(" & grpRef & "=" & kpiGrpRef & ")"
    Else
        gfAll = "": gfGrp = ""
    End If

    currentCol = COL_START

    ' Card 1: Total Systems (New + Reused + Demo with dates; exclude Completed/Cancelled/Non IQ)
    Dim compExcl As String
    If m_statusCol > 0 Then
        compExcl = "*(" & statusRef & "<>""Completed"")*(" & statusRef & "<>""Cancelled"")*(" & statusRef & "<>""Non IQ"")"
    Else
        compExcl = ""
    End If

    If m_nrCol > 0 And m_groupCol > 0 Then
        BuildSingleKPICard currentCol, "Total Systems", _
            "=IF(" & kpiGrpRef & "=""All""," & _
            "SUMPRODUCT((" & nrRef & "<>"""")" & compExcl & hasDateFilter & ")," & _
            "SUMPRODUCT((" & grpRef & "=" & kpiGrpRef & ")*(" & nrRef & "<>"""")" & compExcl & hasDateFilter & "))", RGB(86, 156, 190)
    ElseIf m_nrCol > 0 Then
        BuildSingleKPICard currentCol, "Total Systems", _
            "=SUMPRODUCT((" & nrRef & "<>"""")" & compExcl & hasDateFilter & ")", RGB(86, 156, 190)
    Else
        BuildSingleKPICard currentCol, "Total Systems", "0", RGB(86, 156, 190)
    End If
    currentCol = currentCol + 3

    ' Card 2: New
    If m_nrCol > 0 And m_groupCol > 0 Then
        BuildSingleKPICard currentCol, "New", _
            "=IF(" & kpiGrpRef & "=""All""," & _
            "SUMPRODUCT((" & nrRef & "=""New"")" & compExcl & hasDateFilter & ")," & _
            "SUMPRODUCT((" & grpRef & "=" & kpiGrpRef & ")*(" & nrRef & "=""New"")" & compExcl & hasDateFilter & "))", THEME_SUCCESS
    ElseIf m_nrCol > 0 Then
        BuildSingleKPICard currentCol, "New", _
            "=SUMPRODUCT((" & nrRef & "=""New"")" & compExcl & hasDateFilter & ")", THEME_SUCCESS
    Else
        BuildSingleKPICard currentCol, "New", "0", THEME_SUCCESS
    End If
    currentCol = currentCol + 3

    ' Card 3: Reused
    If m_nrCol > 0 And m_groupCol > 0 Then
        BuildSingleKPICard currentCol, "Reused", _
            "=IF(" & kpiGrpRef & "=""All""," & _
            "SUMPRODUCT((" & nrRef & "=""Reused"")" & compExcl & hasDateFilter & ")," & _
            "SUMPRODUCT((" & grpRef & "=" & kpiGrpRef & ")*(" & nrRef & "=""Reused"")" & compExcl & hasDateFilter & "))", THEME_ACCENT2
    ElseIf m_nrCol > 0 Then
        BuildSingleKPICard currentCol, "Reused", _
            "=SUMPRODUCT((" & nrRef & "=""Reused"")" & compExcl & hasDateFilter & ")", THEME_ACCENT2
    Else
        BuildSingleKPICard currentCol, "Reused", "0", THEME_ACCENT2
    End If
    currentCol = currentCol + 3

    ' Card 4: Demo â€” no hasDateFilter (Demo uses Decon/Demo milestones, matches Working Sheet)
    If m_nrCol > 0 And m_groupCol > 0 Then
        BuildSingleKPICard currentCol, “Demo”, _
            “=IF(“ & kpiGrpRef & “=””All””,” & _
            “SUMPRODUCT((“ & nrRef & “=””Demo””)” & compExcl & “*1),” & _
            “SUMPRODUCT((“ & grpRef & “=” & kpiGrpRef & “)*(“ & nrRef & “=””Demo””)” & compExcl & “))”, THEME_DANGER
    ElseIf m_nrCol > 0 Then
        BuildSingleKPICard currentCol, “Demo”, _
            “=SUMPRODUCT((“ & nrRef & “=””Demo””)” & compExcl & “*1)”, THEME_DANGER
    Else
        BuildSingleKPICard currentCol, “Demo”, “0”, THEME_DANGER
    End If
    currentCol = currentCol + 3

    ' Card 5: CT Est Miss (New only, CT > threshold, with hasDateFilter)
    If m_ctCol > 0 And m_nrCol > 0 Then
        Dim ctThreshold As Long
        ctThreshold = 85
        On Error Resume Next
        If SheetExists(ThisWorkbook, TIS_SHEET_DEFINITIONS) Then
            If IsNumeric(ThisWorkbook.Sheets(TIS_SHEET_DEFINITIONS).Range("S1").Value) Then
                ctThreshold = CLng(ThisWorkbook.Sheets(TIS_SHEET_DEFINITIONS).Range("S1").Value)
            End If
        End If
        On Error GoTo 0
        If m_groupCol > 0 Then
            BuildSingleKPICard currentCol, "CT Miss", _
                "=IF(" & kpiGrpRef & "=""All""," & _
                "SUMPRODUCT((" & nrRef & "=""New"")*(" & ctRef & ">" & ctThreshold & ")" & compExcl & hasDateFilter & ")," & _
                "SUMPRODUCT((" & grpRef & "=" & kpiGrpRef & ")*(" & nrRef & "=""New"")*(" & ctRef & ">" & ctThreshold & ")" & compExcl & hasDateFilter & "))", THEME_DANGER
        Else
            BuildSingleKPICard currentCol, "CT Miss", _
                "=SUMPRODUCT((" & nrRef & "=""New"")*(" & ctRef & ">" & ctThreshold & ")" & compExcl & hasDateFilter & ")", THEME_DANGER
        End If
    Else
        BuildSingleKPICard currentCol, "CT Miss", "0", THEME_DANGER
    End If
    currentCol = currentCol + 3

    ' Card 6: Escalated (no date filter â€” it's a manual status flag; exclude Completed/Cancelled/Non IQ)
    If m_escCol > 0 And m_groupCol > 0 Then
        BuildSingleKPICard currentCol, “Escalated”, _
            “=IF(“ & kpiGrpRef & “=””All””,” & _
            “SUMPRODUCT((“ & escRef & “=””Escalated””)” & compExcl & “*1),” & _
            “SUMPRODUCT((“ & grpRef & “=” & kpiGrpRef & “)*(“ & escRef & “=””Escalated””)” & compExcl & “))”, THEME_WARNING
    ElseIf m_escCol > 0 Then
        BuildSingleKPICard currentCol, “Escalated”, _
            “=SUMPRODUCT((“ & escRef & “=””Escalated””)” & compExcl & “*1)”, THEME_WARNING
    Else
        BuildSingleKPICard currentCol, “Escalated”, “0”, THEME_WARNING
    End If
    currentCol = currentCol + 3

    ' Card 7: Watched (no date filter â€” it's a manual status flag; exclude Completed/Cancelled/Non IQ)
    If m_escCol > 0 And m_groupCol > 0 Then
        BuildSingleKPICard currentCol, “Watched”, _
            “=IF(“ & kpiGrpRef & “=””All””,” & _
            “SUMPRODUCT((“ & escRef & “=””Watched””)” & compExcl & “*1),” & _
            “SUMPRODUCT((“ & grpRef & “=” & kpiGrpRef & “)*(“ & escRef & “=””Watched””)” & compExcl & “))”, THEME_ACCENT2
    ElseIf m_escCol > 0 Then
        BuildSingleKPICard currentCol, “Watched”, _
            “=SUMPRODUCT((“ & escRef & “=””Watched””)” & compExcl & “*1)”, THEME_ACCENT2
    Else
        BuildSingleKPICard currentCol, “Watched”, “0”, THEME_ACCENT2
    End If
    currentCol = currentCol + 3

    ' Card 8: Conversions (uses Conversion column - TRUE = count; exclude Completed/Cancelled/Non IQ)
    If m_conversionCol > 0 And m_groupCol > 0 Then
        Dim convRef As String
        convRef = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_conversionCol).Value & "]"
        BuildSingleKPICard currentCol, "Conversions", _
            "=IF(" & kpiGrpRef & "=""All""," & _
            "SUMPRODUCT((" & convRef & "=TRUE)" & compExcl & "*1)," & _
            "SUMPRODUCT((" & grpRef & "=" & kpiGrpRef & ")*(" & convRef & "=TRUE)" & compExcl & "))", THEME_DANGER
    ElseIf m_conversionCol > 0 Then
        Dim convRef2 As String
        convRef2 = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_conversionCol).Value & "]"
        BuildSingleKPICard currentCol, "Conversions", _
            "=SUMPRODUCT((" & convRef2 & "=TRUE)" & compExcl & "*1)", THEME_DANGER
    Else
        BuildSingleKPICard currentCol, "Conversions", "0", THEME_DANGER
    End If
    currentCol = currentCol + 3

    ' Card 9: Completed
    If m_statusCol > 0 And m_groupCol > 0 Then
        BuildSingleKPICard currentCol, "Completed", _
            "=IF(" & kpiGrpRef & "=""All""," & _
            "SUMPRODUCT((" & statusRef & "=""Completed"")*1)," & _
            "SUMPRODUCT((" & grpRef & "=" & kpiGrpRef & ")*(" & statusRef & "=""Completed"")))", THEME_SUCCESS
    ElseIf m_statusCol > 0 Then
        BuildSingleKPICard currentCol, "Completed", _
            "=SUMPRODUCT((" & statusRef & "=""Completed"")*1)", THEME_SUCCESS
    Else
        BuildSingleKPICard currentCol, "Completed", "0", THEME_SUCCESS
    End If

    ' Divider line
    With m_ws.Range(m_ws.Cells(ROW_DIVIDER, COL_START), m_ws.Cells(ROW_DIVIDER, COL_END))
        .Borders(xlEdgeBottom).Color = THEME_ACCENT
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    m_ws.Rows(ROW_DIVIDER).RowHeight = 8

    m_nextRow = ROW_DIVIDER + 1
End Sub

Private Sub BuildSingleKPICard(col As Long, label As String, formula As String, accentColor As Long)
    Dim cardRange As Range
    Set cardRange = m_ws.Range(m_ws.Cells(ROW_KPI_LABEL, col), m_ws.Cells(ROW_KPI_VALUE, col + 1))

    ' Card background (white card with colored left accent)
    FormatCardStyle cardRange, RGB(255, 255, 255), accentColor

    ' Bottom accent border (colored by KPI type)
    With m_ws.Range(m_ws.Cells(ROW_KPI_VALUE, col), m_ws.Cells(ROW_KPI_VALUE, col + 1))
        .Borders(xlEdgeBottom).Color = accentColor
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

    ' Top accent border (thin, same color)
    With m_ws.Range(m_ws.Cells(ROW_KPI_LABEL, col), m_ws.Cells(ROW_KPI_LABEL, col + 1))
        .Borders(xlEdgeTop).Color = accentColor
        .Borders(xlEdgeTop).Weight = xlThin
    End With

    ' Label (slate text on white card)
    With m_ws.Range(m_ws.Cells(ROW_KPI_LABEL, col), m_ws.Cells(ROW_KPI_LABEL, col + 1))
        .Merge
        .Value = label
        .Font.Size = 9
        .Font.Color = RGB(100, 116, 139)
        .Font.Name = THEME_FONT
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With

    ' Value (accent-colored number)
    With m_ws.Range(m_ws.Cells(ROW_KPI_VALUE, col), m_ws.Cells(ROW_KPI_VALUE, col + 1))
        .Merge
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = accentColor
        .Font.Name = THEME_FONT
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "#,##0"
    End With

    ' Set formula or value
    If Left(formula, 1) = "=" Then
        On Error Resume Next
        m_ws.Cells(ROW_KPI_VALUE, col).formula = formula
        If Err.Number <> 0 Then
            Err.Clear
            m_ws.Cells(ROW_KPI_VALUE, col).Formula2 = formula
            If Err.Number <> 0 Then
                Err.Clear
                m_ws.Cells(ROW_KPI_VALUE, col).Value = 0
            End If
        End If
        On Error GoTo 0
    Else
        m_ws.Cells(ROW_KPI_VALUE, col).Value = CLng(formula)
    End If

    m_ws.Rows(ROW_KPI_LABEL).RowHeight = 22
    m_ws.Rows(ROW_KPI_VALUE).RowHeight = 38
End Sub

'====================================================================
' SECTION 3: SYSTEM COUNTERS - Grouped by Group -> CEID breakdown
' Uses CEID column (falls back to Entity Code if no CEID column)
' Columns: Group/CEID | (spacer) | New | Reused | Total | (spacer) | Demo | Conversions | Completed
' CEID rows are collapsible via row grouping.
'====================================================================

Private Sub BuildSystemCountersTable()
    Dim startRow As Long
    startRow = m_nextRow + 1

    WriteSectionTitle startRow, "System Counters by Group", "Per-group summary with CEID breakdown"

    ' Determine which column to use for sub-row drill-down
    Dim drillCol As Long
    If m_ceidCol > 0 Then
        drillCol = m_ceidCol
    ElseIf m_entityCodeCol > 0 Then
        drillCol = m_entityCodeCol
    Else
        drillCol = 0
    End If

    If drillCol = 0 Or m_nrCol = 0 Then
        m_ws.Cells(startRow + 2, COL_START).Value = "CEID/Entity Code or New/Reused column not found."
        m_ws.Cells(startRow + 2, COL_START).Font.Size = 9
        m_ws.Cells(startRow + 2, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 4
        Exit Sub
    End If

    ' Build structured table references (auto-adjust to table size)
    Dim drillRange As String, nrRange As String, grpRange As String
    drillRange = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, drillCol).Value & "]"
    nrRange = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_nrCol).Value & "]"
    If m_groupCol > 0 Then
        grpRange = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_groupCol).Value & "]"
    End If

    ' Build CV Start, Status and Conversion structured refs
    Dim cvStartRange As String, statusRange As String, convRange As String
    If m_cvStartCol > 0 Then
        cvStartRange = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_cvStartCol).Value & "]"
    End If
    If m_statusCol > 0 Then
        statusRange = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_statusCol).Value & "]"
    End If
    If m_conversionCol > 0 Then
        convRange = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_conversionCol).Value & "]"
    End If

    ' Collect groups and CEIDs per group
    Dim groups As Collection
    Set groups = CollectGroups()

    Dim ceidsByGroup As Object  ' Dictionary: groupName -> Dictionary of CEIDs
    Set ceidsByGroup = CreateObject("Scripting.Dictionary")
    Dim allDict As Object
    Set allDict = CreateObject("Scripting.Dictionary")

    ' Read drill and group columns into arrays (avoid per-cell reads)
    Dim drillData() As Variant
    drillData = m_workSheet.Range(m_workSheet.Cells(m_firstDataRow, drillCol), _
        m_workSheet.Cells(m_lastDataRow, drillCol)).Value
    Dim grpData() As Variant
    If m_groupCol > 0 Then
        grpData = m_workSheet.Range(m_workSheet.Cells(m_firstDataRow, m_groupCol), _
            m_workSheet.Cells(m_lastDataRow, m_groupCol)).Value
    End If

    Dim r As Long, gVal As String, cVal As String
    For r = 1 To UBound(drillData, 1)
        cVal = Trim(CStr(drillData(r, 1)))
        If cVal = "" Then GoTo NextCeidRow
        If m_groupCol > 0 Then
            gVal = Trim(CStr(grpData(r, 1)))
        Else
            gVal = "(No Group)"
        End If
        If gVal = "" Then gVal = "(No Group)"

        If Not ceidsByGroup.exists(gVal) Then
            Set ceidsByGroup(gVal) = CreateObject("Scripting.Dictionary")
        End If
        ceidsByGroup(gVal)(cVal) = True
        allDict(cVal) = True
NextCeidRow:
    Next r

    If allDict.Count = 0 Then
        m_ws.Cells(startRow + 2, COL_START).Value = "No CEIDs found in data."
        m_ws.Cells(startRow + 2, COL_START).Font.Size = 9
        m_ws.Cells(startRow + 2, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 4
        Exit Sub
    End If

    ' Table header row
    Dim hdrRow As Long
    hdrRow = startRow + 2

    Dim headers As Variant
    headers = Array("Group / CEID", "", "New", "Reused", "Total", "", "Demo", "Conversions", "Completed")
    Dim h As Long
    For h = 0 To UBound(headers)
        With m_ws.Cells(hdrRow, COL_START + h)
            .Value = headers(h)
            .Font.Bold = True
            .Font.Size = 10
            .Font.Color = TABLE_HEADER_TEXT
            .Font.Name = THEME_FONT
            .Interior.Color = TABLE_HEADER_BG
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders(xlEdgeBottom).Color = RGB(86, 156, 190)
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
    Next h
    ' Merge first two header cells for wider Group/CEID column
    m_ws.Range(m_ws.Cells(hdrRow, COL_START), m_ws.Cells(hdrRow, COL_START + 1)).Merge
    m_ws.Rows(hdrRow).RowHeight = 24
    ' Spacer column (index 5) - use formatting only, NOT narrow width
    ' (narrow width causes ## in HC Gap section which reuses the same physical column)
    m_ws.Columns(COL_START + 5).ColumnWidth = 5
    m_ws.Cells(hdrRow, COL_START + 5).Interior.Color = TABLE_HEADER_BG  ' match header
    ' Add right border on Total column for visual separation
    m_ws.Cells(hdrRow, COL_START + 4).Borders(xlEdgeRight).Color = RGB(203, 213, 225)
    m_ws.Cells(hdrRow, COL_START + 4).Borders(xlEdgeRight).Weight = xlThin

    ' Write grouped data rows
    Dim dataRow As Long
    dataRow = hdrRow + 1

    Dim gKey As Variant, cKey As Variant
    Dim groupCeids As Object
    Dim rowRng As Range
    Dim firstCeidRow As Long, lastCeidRow As Long, ceidCount As Long

    ' Iterate groups in sorted order (most systems first)
    Dim sortedGi As Long
    For sortedGi = 1 To groups.Count
        gKey = groups(sortedGi)
        If Not ceidsByGroup.exists(CStr(gKey)) Then GoTo NextSortedGroup
        ' === Group header row ===
        m_ws.Range(m_ws.Cells(dataRow, COL_START), m_ws.Cells(dataRow, COL_START + 1)).Merge
        m_ws.Cells(dataRow, COL_START).Value = CStr(gKey)
        m_ws.Cells(dataRow, COL_START).Font.Bold = True
        m_ws.Cells(dataRow, COL_START).Font.Size = 10
        m_ws.Cells(dataRow, COL_START).Font.Color = THEME_WHITE
        m_ws.Cells(dataRow, COL_START).Font.Name = THEME_FONT
        m_ws.Cells(dataRow, COL_START).IndentLevel = 0

        ' Group-level formulas (SUMPRODUCT with hasDateFilter for New/Reused)
        Dim grpCompFilter As String
        If m_statusCol > 0 Then
            grpCompFilter = "*(" & statusRange & "<>""Completed"")*(" & statusRange & "<>""Cancelled"")*(" & statusRange & "<>""Non IQ"")"
        Else
            grpCompFilter = ""
        End If

        If m_groupCol > 0 Then
            On Error Resume Next
            ' New (with hasDateFilter, exclude completed)
            SafeFormulaWrite m_ws, dataRow, COL_START + 2, _
                "=SUMPRODUCT((" & grpRange & "=""" & CStr(gKey) & """)*(" & nrRange & "=""New"")" & grpCompFilter & m_hasDateFilter & ")"
            ' Reused (with hasDateFilter, exclude completed)
            SafeFormulaWrite m_ws, dataRow, COL_START + 3, _
                "=SUMPRODUCT((" & grpRange & "=""" & CStr(gKey) & """)*(" & nrRange & "=""Reused"")" & grpCompFilter & m_hasDateFilter & ")"
            ' Total = New + Reused (NOT Demo)
            m_ws.Cells(dataRow, COL_START + 4).formula = _
                "=" & ColLetter(COL_START + 2) & dataRow & "+" & ColLetter(COL_START + 3) & dataRow
            ' Spacer col 5 = empty
            ' Demo (no hasDateFilter, exclude Completed/Cancelled/Non IQ)
            SafeFormulaWrite m_ws, dataRow, COL_START + 6, _
                "=SUMPRODUCT((" & grpRange & "=""" & CStr(gKey) & """)*(" & nrRange & "=""Demo"")" & grpCompFilter & "*1)"
            ' Conversions (use Conversion column if available, fallback to CV Start; exclude Completed/Cancelled/Non IQ)
            If m_conversionCol > 0 Then
                SafeFormulaWrite m_ws, dataRow, COL_START + 7, _
                    "=SUMPRODUCT((" & grpRange & "=""" & CStr(gKey) & """)*(" & convRange & "=TRUE)" & grpCompFilter & ")"
            ElseIf m_cvStartCol > 0 Then
                SafeFormulaWrite m_ws, dataRow, COL_START + 7, _
                    "=SUMPRODUCT((" & grpRange & "=""" & CStr(gKey) & """)*(ISNUMBER(" & cvStartRange & "))" & grpCompFilter & ")"
            End If
            ' Completed
            If m_statusCol > 0 Then
                SafeFormulaWrite m_ws, dataRow, COL_START + 8, _
                    "=SUMPRODUCT((" & grpRange & "=""" & CStr(gKey) & """)*(" & statusRange & "=""Completed""))"
            End If
            On Error GoTo 0
        End If

        Dim gc As Long
        For gc = 2 To 8
            If gc <> 5 Then  ' skip spacer
                m_ws.Cells(dataRow, COL_START + gc).Font.Bold = True
                m_ws.Cells(dataRow, COL_START + gc).Font.Size = 10
                m_ws.Cells(dataRow, COL_START + gc).Font.Color = THEME_WHITE
                m_ws.Cells(dataRow, COL_START + gc).HorizontalAlignment = xlCenter
            End If
        Next gc

        ' Group row background (dark blue accent with bottom border)
        Set rowRng = m_ws.Range(m_ws.Cells(dataRow, COL_START), m_ws.Cells(dataRow, COL_START + 8))
        rowRng.Interior.Color = TABLE_GROUP_BG
        rowRng.Borders(xlEdgeBottom).Color = RGB(203, 213, 225)
        rowRng.Borders(xlEdgeBottom).Weight = xlThin
        rowRng.Borders(xlEdgeTop).Color = RGB(203, 213, 225)
        rowRng.Borders(xlEdgeTop).Weight = xlThin
        ' Spacer cell + Total right border
        m_ws.Cells(dataRow, COL_START + 5).Interior.Color = TABLE_GROUP_BG
        m_ws.Cells(dataRow, COL_START + 4).Borders(xlEdgeRight).Color = RGB(203, 213, 225)
        m_ws.Cells(dataRow, COL_START + 4).Borders(xlEdgeRight).Weight = xlThin

        dataRow = dataRow + 1
        firstCeidRow = dataRow  ' track first CEID row for grouping

        ' === CEID sub-rows ===
        Set groupCeids = ceidsByGroup(gKey)
        Dim ceidIdx As Long
        ceidIdx = 0
        ceidCount = 0
        For Each cKey In groupCeids.keys
            m_ws.Range(m_ws.Cells(dataRow, COL_START), m_ws.Cells(dataRow, COL_START + 1)).Merge
            m_ws.Cells(dataRow, COL_START).Value = "   " & CStr(cKey)
            m_ws.Cells(dataRow, COL_START).Font.Size = 9
            m_ws.Cells(dataRow, COL_START).Font.Color = RGB(100, 116, 139)
            m_ws.Cells(dataRow, COL_START).Font.Name = THEME_FONT

            ' CEID-level formulas (SUMPRODUCT with hasDateFilter for New/Reused)
            Dim ceidCompFilter As String
            If m_statusCol > 0 Then
                ceidCompFilter = "*(" & statusRange & "<>""Completed"")*(" & statusRange & "<>""Cancelled"")*(" & statusRange & "<>""Non IQ"")"
            Else
                ceidCompFilter = ""
            End If

            On Error Resume Next
            ' New (with hasDateFilter, exclude completed)
            SafeFormulaWrite m_ws, dataRow, COL_START + 2, _
                "=SUMPRODUCT((" & drillRange & "=""" & CStr(cKey) & """)*(" & nrRange & "=""New"")" & ceidCompFilter & m_hasDateFilter & ")"
            ' Reused (with hasDateFilter, exclude completed)
            SafeFormulaWrite m_ws, dataRow, COL_START + 3, _
                "=SUMPRODUCT((" & drillRange & "=""" & CStr(cKey) & """)*(" & nrRange & "=""Reused"")" & ceidCompFilter & m_hasDateFilter & ")"
            ' Total = New + Reused
            m_ws.Cells(dataRow, COL_START + 4).formula = _
                "=" & ColLetter(COL_START + 2) & dataRow & "+" & ColLetter(COL_START + 3) & dataRow
            ' Demo (no hasDateFilter, exclude Completed/Cancelled/Non IQ)
            SafeFormulaWrite m_ws, dataRow, COL_START + 6, _
                "=SUMPRODUCT((" & drillRange & "=""" & CStr(cKey) & """)*(" & nrRange & "=""Demo"")" & ceidCompFilter & "*1)"
            ' Conversions (use Conversion column if available, fallback to CV Start; exclude Completed/Cancelled/Non IQ)
            If m_conversionCol > 0 Then
                SafeFormulaWrite m_ws, dataRow, COL_START + 7, _
                    "=SUMPRODUCT((" & drillRange & "=""" & CStr(cKey) & """)*(" & convRange & "=TRUE)" & ceidCompFilter & ")"
            ElseIf m_cvStartCol > 0 Then
                SafeFormulaWrite m_ws, dataRow, COL_START + 7, _
                    "=SUMPRODUCT((" & drillRange & "=""" & CStr(cKey) & """)*(ISNUMBER(" & cvStartRange & "))" & ceidCompFilter & ")"
            End If
            ' Completed
            If m_statusCol > 0 Then
                SafeFormulaWrite m_ws, dataRow, COL_START + 8, _
                    "=SUMPRODUCT((" & drillRange & "=""" & CStr(cKey) & """)*(" & statusRange & "=""Completed""))"
            End If
            On Error GoTo 0

            Dim cc As Long
            For cc = 2 To 8
                If cc <> 5 Then
                    m_ws.Cells(dataRow, COL_START + cc).Font.Size = 9
                    m_ws.Cells(dataRow, COL_START + cc).Font.Color = TABLE_TEXT
                    m_ws.Cells(dataRow, COL_START + cc).HorizontalAlignment = xlCenter
                End If
            Next cc

            ' Zebra striping for sub-rows (brand light blue tones)
            Set rowRng = m_ws.Range(m_ws.Cells(dataRow, COL_START), m_ws.Cells(dataRow, COL_START + 8))
            If ceidIdx Mod 2 = 0 Then
                rowRng.Interior.Color = TABLE_ROW_BG
            Else
                rowRng.Interior.Color = TABLE_ALT_ROW_BG
            End If
            rowRng.Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
            rowRng.Borders(xlEdgeBottom).Weight = xlHairline
            ' Spacer cell + Total right border
            m_ws.Cells(dataRow, COL_START + 5).Interior.Color = rowRng.Cells(1, 1).Interior.Color
            m_ws.Cells(dataRow, COL_START + 4).Borders(xlEdgeRight).Color = RGB(22, 54, 92)
            m_ws.Cells(dataRow, COL_START + 4).Borders(xlEdgeRight).Weight = xlThin

            dataRow = dataRow + 1
            ceidIdx = ceidIdx + 1
            ceidCount = ceidCount + 1
        Next cKey

        lastCeidRow = dataRow - 1

        ' Group CEID rows for collapsibility
        If ceidCount > 0 Then
            On Error Resume Next
            m_ws.Rows(firstCeidRow & ":" & lastCeidRow).Group
            On Error GoTo 0
        End If
NextSortedGroup:
    Next sortedGi

    ' Set outline to collapsed by default
    On Error Resume Next
    m_ws.Outline.ShowLevels RowLevels:=1
    On Error GoTo 0

    ' Totals row
    Dim totalsRow As Long
    totalsRow = dataRow
    m_ws.Range(m_ws.Cells(totalsRow, COL_START), m_ws.Cells(totalsRow, COL_START + 1)).Merge
    m_ws.Cells(totalsRow, COL_START).Value = "TOTAL"
    m_ws.Cells(totalsRow, COL_START).Font.Bold = True
    m_ws.Cells(totalsRow, COL_START).Font.Size = 10
    m_ws.Cells(totalsRow, COL_START).Font.Color = TABLE_HEADER_TEXT

    Dim sumCol As Long
    Dim compFilter As String
    If m_statusCol > 0 Then
        compFilter = "*(" & statusRange & "<>""Completed"")*(" & statusRange & "<>""Cancelled"")*(" & statusRange & "<>""Non IQ"")"
    Else
        compFilter = ""
    End If

    For sumCol = 2 To 8
        If sumCol = 5 Then GoTo NextSumCol  ' skip spacer
        Select Case sumCol
            Case 2  ' New (with hasDateFilter, exclude completed)
                SafeFormulaWrite m_ws, totalsRow, COL_START + sumCol, _
                    "=SUMPRODUCT((" & nrRange & "=""New"")" & compFilter & m_hasDateFilter & ")"
            Case 3  ' Reused (with hasDateFilter, exclude completed)
                SafeFormulaWrite m_ws, totalsRow, COL_START + sumCol, _
                    "=SUMPRODUCT((" & nrRange & "=""Reused"")" & compFilter & m_hasDateFilter & ")"
            Case 4  ' Total = New + Reused
                m_ws.Cells(totalsRow, COL_START + sumCol).formula = _
                    "=" & ColLetter(COL_START + 2) & totalsRow & "+" & ColLetter(COL_START + 3) & totalsRow
            Case 6  ' Demo (no hasDateFilter, exclude Completed/Cancelled/Non IQ)
                SafeFormulaWrite m_ws, totalsRow, COL_START + sumCol, _
                    "=SUMPRODUCT((" & nrRange & "=""Demo"")" & compFilter & "*1)"
            Case 7  ' Conversions (use Conversion column, TRUE = count; exclude Completed/Cancelled/Non IQ)
                If m_conversionCol > 0 Then
                    SafeFormulaWrite m_ws, totalsRow, COL_START + sumCol, _
                        "=SUMPRODUCT((" & convRange & "=TRUE)" & compFilter & "*1)"
                ElseIf m_cvStartCol > 0 Then
                    SafeFormulaWrite m_ws, totalsRow, COL_START + sumCol, _
                        "=SUMPRODUCT((ISNUMBER(" & cvStartRange & "))" & compFilter & "*1)"
                End If
            Case 8  ' Completed
                If m_statusCol > 0 Then
                    SafeFormulaWrite m_ws, totalsRow, COL_START + sumCol, _
                        "=SUMPRODUCT((" & statusRange & "=""Completed"")*1)"
                End If
        End Select
        m_ws.Cells(totalsRow, COL_START + sumCol).Font.Bold = True
        m_ws.Cells(totalsRow, COL_START + sumCol).Font.Size = 10
        m_ws.Cells(totalsRow, COL_START + sumCol).Font.Color = TABLE_TEXT
        m_ws.Cells(totalsRow, COL_START + sumCol).HorizontalAlignment = xlCenter
NextSumCol:
    Next sumCol

    ' Totals row formatting
    Set rowRng = m_ws.Range(m_ws.Cells(totalsRow, COL_START), m_ws.Cells(totalsRow, COL_START + 8))
    rowRng.Interior.Color = TABLE_HEADER_BG
    rowRng.Font.Color = TABLE_HEADER_TEXT
    rowRng.Borders(xlEdgeTop).Color = RGB(203, 213, 225)
    rowRng.Borders(xlEdgeTop).Weight = xlMedium
    rowRng.Borders(xlEdgeBottom).Color = RGB(203, 213, 225)
    rowRng.Borders(xlEdgeBottom).Weight = xlMedium
    ' Spacer cell + Total right border
    m_ws.Cells(totalsRow, COL_START + 5).Interior.Color = TABLE_HEADER_BG
    m_ws.Cells(totalsRow, COL_START + 4).Borders(xlEdgeRight).Color = RGB(203, 213, 225)
    m_ws.Cells(totalsRow, COL_START + 4).Borders(xlEdgeRight).Weight = xlThin

    m_nextRow = totalsRow + 2
End Sub

'====================================================================
' SECTION 4: HC GAP ANALYSIS - PER-GROUP MONTHLY VIEW
' Averages weekly NIF HC Analyzer data to monthly.
' Shows 12-month dynamic timeframe with 3 charts per group:
'   Total Need vs Available | New Need vs Available | Reused Need vs Available
' All groups displayed upfront (no group dropdown).
' Rev11: Adapted for NIF Rev11 structure with separate New/Reuse tables.
'        Charts use deterministic names (HC_*) for idempotent rebuild.
'====================================================================

Private Sub BuildHCGapAnalysis()
    Dim startRow As Long
    startRow = m_nextRow + 1

    WriteSectionTitle startRow, "HC Gap Analysis", _
        ChrW(&H25B2) & " Need  vs  " & ChrW(&H25CF) & " Available   |   " & _
        "Monthly avg of weekly HC data   |   12-month rolling window"

    If m_groupCol = 0 Or m_nrCol = 0 Then
        WriteHCPlaceholder startRow + 2, "Group or New/Reused column not found."
        m_nextRow = startRow + 4: Exit Sub
    End If

    ' --- Step 1: Discover HC source tables ---
    Dim refs As HCTableRefs
    If Not DiscoverHCTables(refs) Then
        WriteHCPlaceholder startRow + 2, "NIF HC Analyzer tables not found. Run NIF Builder first."
        m_nextRow = startRow + 4: Exit Sub
    End If

    ' --- Step 2: Build month buckets ---
    Dim mcs(1 To HC_MONTHS) As Long, mce(1 To HC_MONTHS) As Long
    Dim hcStart As Date, mi As Long, jj As Long
    hcStart = DateSerial(Year(Date), Month(Date), 1)

    For mi = 1 To HC_MONTHS
        Dim ms As Date, me2 As Date
        ms = DateAdd("m", mi - 1, hcStart)
        me2 = DateAdd("m", mi, hcStart)
        mcs(mi) = 0: mce(mi) = 0
        For jj = refs.dsc To refs.dec
            If IsDate(m_workSheet.Cells(refs.refHdr, jj).Value) Then
                Dim dt As Date: dt = CDate(m_workSheet.Cells(refs.refHdr, jj).Value)
                If dt >= ms And dt < me2 Then
                    If mcs(mi) = 0 Then mcs(mi) = jj
                    mce(mi) = jj
                End If
            End If
        Next jj
    Next mi

    ' --- Step 3: Collect groups (sorted by total system count, most first) ---
    Dim hcGroupsRaw As New Collection
    Dim refDsr As Long: refDsr = refs.refTitle + 2
    Dim refLastRow As Long: refLastRow = GetHCTableDataLastRow(refs.refTitle, refs.refGC)
    Dim dr As Long
    For dr = refDsr To refLastRow
        Dim cv As String: cv = Trim(CStr(m_workSheet.Cells(dr, refs.refGC).Value))
        If cv <> "" And LCase(cv) <> "total" Then hcGroupsRaw.Add cv
    Next dr

    ' Sort by system count (matches System Counters / Group Breakdown order)
    Dim hcGroups As Collection
    Set hcGroups = SortGroupsBySystemCount(hcGroupsRaw)

    If hcGroups.Count = 0 Then
        WriteHCPlaceholder startRow + 2, "No groups found in HC tables."
        m_nextRow = startRow + 4: Exit Sub
    End If

    ' --- Step 4: Build group-to-row maps ---
    Dim newGrpRows As Object, reuGrpRows As Object
    Dim newAvailGrpRows As Object, reuAvailGrpRows As Object
    Dim newGapGrpRows As Object, reuGapGrpRows As Object, totalGapGrpRows As Object

    Set newGrpRows = BuildHCGroupRowMap(refs.NewTitle, refs.gcNewNeed)
    Set reuGrpRows = BuildHCGroupRowMap(refs.ReusedTitle, refs.gcReuNeed)
    Set newAvailGrpRows = BuildHCGroupRowMap(refs.NewAvailTitle, refs.gcNewAvail)
    Set reuAvailGrpRows = BuildHCGroupRowMap(refs.ReuAvailTitle, refs.gcReuAvail)
    Set newGapGrpRows = BuildHCGroupRowMap(refs.NewGapTitle, refs.gcNewGap)
    Set reuGapGrpRows = BuildHCGroupRowMap(refs.ReuGapTitle, refs.gcReuGap)
    Set totalGapGrpRows = BuildHCGroupRowMap(refs.TotalGapTitle, refs.gcTotalGap)

    ' Total rows per source table (indexed 0-7)
    Dim totRows(0 To 7) As Long
    totRows(0) = FindHCTotalRow(refs.NewTitle, refs.gcNewNeed)
    totRows(1) = FindHCTotalRow(refs.ReusedTitle, refs.gcReuNeed)
    totRows(2) = 0   ' Combined (not used directly)
    totRows(3) = FindHCTotalRow(refs.NewAvailTitle, refs.gcNewAvail)
    totRows(4) = FindHCTotalRow(refs.ReuAvailTitle, refs.gcReuAvail)
    totRows(5) = FindHCTotalRow(refs.NewGapTitle, refs.gcNewGap)
    totRows(6) = FindHCTotalRow(refs.ReuGapTitle, refs.gcReuGap)
    totRows(7) = FindHCTotalRow(refs.TotalGapTitle, refs.gcTotalGap)

    ' --- Step 5: Clean up old HC charts ---
    CleanupHCCharts

    ' --- Step 6: Write month header ---
    Dim hdrRow As Long: hdrRow = startRow + 2
    Dim lastMCol As Long
    lastMCol = Application.WorksheetFunction.Min(COL_START + HC_MONTHS, COL_END)
    WriteHCMonthHeaderRow hdrRow, hcStart, lastMCol

    ' --- Step 7: Per-group blocks ---
    Dim curRow As Long: curRow = hdrRow + 1
    Dim chartCol As Long: chartCol = lastMCol + 2
    Dim gi As Long

    For gi = 1 To hcGroups.Count
        curRow = WriteHCGroupBlock(curRow, CStr(hcGroups(gi)), hdrRow, lastMCol, chartCol, _
            refs, newGrpRows, reuGrpRows, newAvailGrpRows, reuAvailGrpRows, _
            newGapGrpRows, reuGapGrpRows, totalGapGrpRows, mcs, mce)
    Next gi

    ' --- Step 8: TOTAL block ---
    curRow = WriteHCTotalBlock(curRow, hdrRow, lastMCol, chartCol, refs, totRows, mcs, mce)

    m_nextRow = curRow + 1
End Sub

'====================================================================
' HC HELPER: DISCOVER HC TABLES
' Locates all HC analysis tables on the Working Sheet and populates
' the HCTableRefs structure. Returns False if no tables found.
'====================================================================

Private Function DiscoverHCTables(ByRef refs As HCTableRefs) As Boolean
    DiscoverHCTables = False

    refs.NewTitle = FindHCTableRow("New Systems - HC Need")
    refs.ReusedTitle = FindHCTableRow("Reused Systems - HC Need")
    refs.CombTitle = FindHCTableRow("Combined - HC Need")
    refs.NewAvailTitle = FindHCTableRow("New Available HC")
    refs.ReuAvailTitle = FindHCTableRow("Reused Available HC")
    refs.NewGapTitle = FindHCTableRow("New HC Gap")
    refs.ReuGapTitle = FindHCTableRow("Reused HC Gap")
    refs.TotalGapTitle = FindHCTableRow("Total HC Gap")

    If refs.NewTitle = 0 And refs.CombTitle = 0 Then Exit Function

    ' Pick reference table for date columns / group discovery
    If refs.NewAvailTitle > 0 Then
        refs.refTitle = refs.NewAvailTitle
    ElseIf refs.ReuAvailTitle > 0 Then
        refs.refTitle = refs.ReuAvailTitle
    ElseIf refs.NewTitle > 0 Then
        refs.refTitle = refs.NewTitle
    Else
        refs.refTitle = refs.CombTitle
    End If
    refs.refHdr = refs.refTitle + 1
    refs.refGC = FindHCGroupCol(refs.refHdr)
    If refs.refGC = 0 Then refs.refGC = 1

    ' Per-table Group column positions
    If refs.NewTitle > 0 Then refs.gcNewNeed = FindHCGroupCol(refs.NewTitle + 1) Else refs.gcNewNeed = refs.refGC
    If refs.ReusedTitle > 0 Then refs.gcReuNeed = FindHCGroupCol(refs.ReusedTitle + 1) Else refs.gcReuNeed = refs.refGC
    If refs.CombTitle > 0 Then refs.gcComb = FindHCGroupCol(refs.CombTitle + 1) Else refs.gcComb = refs.refGC
    If refs.NewAvailTitle > 0 Then refs.gcNewAvail = FindHCGroupCol(refs.NewAvailTitle + 1) Else refs.gcNewAvail = refs.refGC
    If refs.ReuAvailTitle > 0 Then refs.gcReuAvail = FindHCGroupCol(refs.ReuAvailTitle + 1) Else refs.gcReuAvail = refs.refGC
    If refs.NewGapTitle > 0 Then refs.gcNewGap = FindHCGroupCol(refs.NewGapTitle + 1) Else refs.gcNewGap = refs.refGC
    If refs.ReuGapTitle > 0 Then refs.gcReuGap = FindHCGroupCol(refs.ReuGapTitle + 1) Else refs.gcReuGap = refs.refGC
    If refs.TotalGapTitle > 0 Then refs.gcTotalGap = FindHCGroupCol(refs.TotalGapTitle + 1) Else refs.gcTotalGap = refs.refGC

    ' Date column bounds
    refs.dsc = 0: refs.dec = 0
    Dim jj As Long
    For jj = refs.refGC + 1 To m_workSheet.Cells(refs.refHdr, m_workSheet.Columns.Count).End(xlToLeft).Column
        If IsDate(m_workSheet.Cells(refs.refHdr, jj).Value) Then
            If refs.dsc = 0 Then refs.dsc = jj
            refs.dec = jj
        End If
    Next jj

    DiscoverHCTables = (refs.dsc > 0)
End Function

'====================================================================
' HC HELPER: WRITE MONTH HEADER ROW
'====================================================================

Private Sub WriteHCMonthHeaderRow(hdrRow As Long, hcStart As Date, lastMCol As Long)
    m_ws.Cells(hdrRow, COL_START).Value = ""
    m_ws.Cells(hdrRow, COL_START).Font.Bold = True
    m_ws.Cells(hdrRow, COL_START).Font.Size = 10
    m_ws.Cells(hdrRow, COL_START).Font.Color = TABLE_HEADER_TEXT
    m_ws.Cells(hdrRow, COL_START).Interior.Color = TABLE_HEADER_BG

    Dim mi As Long
    For mi = 1 To HC_MONTHS
        Dim mCol As Long: mCol = COL_START + mi
        If mCol > COL_END Then Exit For
        Dim mDate As Date: mDate = DateAdd("m", mi - 1, hcStart)
        With m_ws.Cells(hdrRow, mCol)
            .Value = mDate
            .NumberFormat = "yy-mmm"
            .Font.Bold = True
            .Font.Size = 8
            .Font.Color = TABLE_HEADER_TEXT
            .HorizontalAlignment = xlCenter
            If Year(mDate) = Year(Date) And Month(mDate) = Month(Date) Then
                .Interior.Color = RGB(86, 156, 190)   ' Brand blue highlight for current month
            Else
                .Interior.Color = TABLE_HEADER_BG
            End If
        End With
    Next mi
    m_ws.Rows(hdrRow).RowHeight = 22
End Sub

'====================================================================
' HC HELPER: CLEANUP OLD HC CHARTS
' Removes chart objects prefixed with "HC_" to prevent duplicates.
'====================================================================

Private Sub CleanupHCCharts()
    If m_ws Is Nothing Then Exit Sub
    Dim co As ChartObject
    Dim i As Long
    For i = m_ws.ChartObjects.Count To 1 Step -1
        Set co = m_ws.ChartObjects(i)
        If Left(co.Name, 3) = "HC_" Then co.Delete
    Next i
End Sub

'====================================================================
' HC HELPER: WRITE PLACEHOLDER MESSAGE
'====================================================================

Private Sub WriteHCPlaceholder(row As Long, msg As String)
    m_ws.Cells(row, COL_START).Value = msg
    m_ws.Cells(row, COL_START).Font.Size = 9
    m_ws.Cells(row, COL_START).Font.Color = RGB(100, 116, 139)
End Sub

'====================================================================
' HC HELPER: CLEAN GROUP NAME FOR CHART OBJECT NAME
' Strips non-alphanumeric characters for deterministic naming.
'====================================================================

Private Function CleanChartName(grpName As String) As String
    Dim result As String, i As Long, c As String
    result = ""
    For i = 1 To Len(grpName)
        c = Mid(grpName, i, 1)
        If c Like "[A-Za-z0-9]" Then result = result & c
    Next i
    If result = "" Then result = "Grp"
    CleanChartName = result
End Function

'====================================================================
' HC HELPER: WRITE ONE GROUP BLOCK
' Writes group header + 7 data rows + 2 summary rows + 3 charts.
' Returns the next available row.
'====================================================================

Private Function WriteHCGroupBlock(curRow As Long, grpName As String, _
    hdrRow As Long, lastMCol As Long, chartCol As Long, _
    refs As HCTableRefs, _
    newGrpRows As Object, reuGrpRows As Object, _
    newAvailGrpRows As Object, reuAvailGrpRows As Object, _
    newGapGrpRows As Object, reuGapGrpRows As Object, _
    totalGapGrpRows As Object, _
    mcs() As Long, mce() As Long) As Long

    Dim r As Long: r = curRow
    Dim grpStartRow As Long: grpStartRow = r

    ' --- Group header card ---
    Dim grpHdrRng As Range
    Set grpHdrRng = m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol))
    grpHdrRng.Interior.Color = RGB(230, 242, 250)
    With m_ws.Cells(r, COL_START).Font
        .Bold = True: .Size = 10
        .Color = RGB(12, 27, 51): .Name = THEME_FONT
    End With
    m_ws.Cells(r, COL_START).Value = grpName
    grpHdrRng.Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
    grpHdrRng.Borders(xlEdgeBottom).Weight = xlHairline
    m_ws.Rows(r).RowHeight = 26
    r = r + 1

    ' --- Data rows ---
    ' Visual banding: faint blue = Need, faint green = Available, white = Gap
    Dim newNeedRow As Long, reuNeedRow As Long
    Dim newAvailRow As Long, reuAvailRow As Long
    Dim HC_NEED_BG As Long: HC_NEED_BG = RGB(237, 243, 252)        ' Faint blue tint
    Dim HC_AVAIL_BG As Long: HC_AVAIL_BG = RGB(235, 251, 242)      ' Faint green tint
    Dim HC_SUMMARY_NEED As Long: HC_SUMMARY_NEED = RGB(222, 234, 250) ' Medium blue band
    Dim HC_SUMMARY_AVAIL As Long: HC_SUMMARY_AVAIL = RGB(222, 250, 234) ' Medium green band

    ' Row 1: New HC Need
    newNeedRow = r
    WriteHCGroupRowDirect r, "  " & ChrW(&H25B2) & " New HC Need", grpName, refs.NewTitle, refs.gcNewNeed, _
        newGrpRows, mcs, mce, lastMCol, RGB(86, 156, 190)
    m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol)).Interior.Color = HC_NEED_BG
    r = r + 1

    ' Row 2: Reuse HC Need
    reuNeedRow = r
    WriteHCGroupRowDirect r, "  " & ChrW(&H25B2) & " Reuse HC Need", grpName, refs.ReusedTitle, refs.gcReuNeed, _
        reuGrpRows, mcs, mce, lastMCol, RGB(86, 156, 190)
    m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol)).Interior.Color = HC_NEED_BG
    r = r + 1

    ' Row 3: New Available
    newAvailRow = r
    WriteHCGroupRowDirect r, "  " & ChrW(&H25CF) & " New Available", grpName, refs.NewAvailTitle, refs.gcNewAvail, _
        newAvailGrpRows, mcs, mce, lastMCol, RGB(46, 184, 92)
    m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol)).Interior.Color = HC_AVAIL_BG
    r = r + 1

    ' Row 4: Reuse Available
    reuAvailRow = r
    WriteHCGroupRowDirect r, "  " & ChrW(&H25CF) & " Reuse Available", grpName, refs.ReuAvailTitle, refs.gcReuAvail, _
        reuAvailGrpRows, mcs, mce, lastMCol, RGB(46, 184, 92)
    m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol)).Interior.Color = HC_AVAIL_BG
    r = r + 1

    ' Row 5: New Gap (white base for CF red/green overlay)
    WriteHCGroupRowDirect r, "  " & ChrW(&H394) & " New Gap", grpName, refs.NewGapTitle, refs.gcNewGap, _
        newGapGrpRows, mcs, mce, lastMCol, TABLE_TEXT
    ApplyGapConditionalFormatting r, lastMCol
    r = r + 1

    ' Row 6: Reuse Gap
    WriteHCGroupRowDirect r, "  " & ChrW(&H394) & " Reuse Gap", grpName, refs.ReuGapTitle, refs.gcReuGap, _
        reuGapGrpRows, mcs, mce, lastMCol, TABLE_TEXT
    ApplyGapConditionalFormatting r, lastMCol
    r = r + 1

    ' Row 7: Total Gap (bold, white base for CF)
    Dim totGapRow As Long: totGapRow = r
    WriteHCGroupRowDirect r, "  " & ChrW(&H394) & " Total Gap", grpName, refs.TotalGapTitle, refs.gcTotalGap, _
        totalGapGrpRows, mcs, mce, lastMCol, TABLE_TEXT
    m_ws.Cells(r, COL_START).Font.Bold = True
    m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol)).Interior.Color = THEME_WHITE
    ApplyGapConditionalFormatting r, lastMCol
    r = r + 1

    ' Row 8: Total Need summary (medium blue band — visually distinct from group header)
    Dim totNeedRow As Long: totNeedRow = r
    WriteHCSummaryFormulaRow r, "  " & ChrW(&H25B2) & " Total Need", lastMCol, _
        newNeedRow, reuNeedRow, RGB(239, 83, 80), HC_SUMMARY_NEED
    r = r + 1

    ' Row 9: Total Available summary (medium green band)
    Dim totAvailRow As Long: totAvailRow = r
    WriteHCSummaryFormulaRow r, "  " & ChrW(&H25CF) & " Total Available", lastMCol, _
        newAvailRow, reuAvailRow, RGB(46, 184, 92), HC_SUMMARY_AVAIL

    ' Bottom separator
    m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol)).Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
    m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol)).Borders(xlEdgeBottom).Weight = xlThin
    r = r + 1

    ' Group detail rows for collapsibility
    Dim detailFirst As Long: detailFirst = grpStartRow + 1
    Dim detailLast As Long: detailLast = totGapRow - 1
    If detailLast >= detailFirst Then
        On Error Resume Next
        m_ws.Rows(detailFirst & ":" & detailLast).Group
        On Error GoTo 0
    End If

    ' --- 3 mini charts: Total / New / Reused ---
    Dim cl As Double, ct As Double
    cl = m_ws.Cells(grpStartRow, chartCol).Left
    ct = m_ws.Cells(grpStartRow, chartCol).Top

    AddHCMiniChart cl, ct, HC_CHART_WIDTH, HC_CHART_HEIGHT, _
        grpName & " - Total", "HC_" & CleanChartName(grpName) & "_Total", _
        hdrRow, lastMCol, totNeedRow, totAvailRow, _
        RGB(239, 83, 80), RGB(46, 184, 92), 2, 4, True

    AddHCMiniChart cl + HC_CHART_WIDTH + HC_CHART_GAP, ct, HC_CHART_WIDTH, HC_CHART_HEIGHT, _
        grpName & " - New", "HC_" & CleanChartName(grpName) & "_New", _
        hdrRow, lastMCol, newNeedRow, newAvailRow, _
        RGB(86, 156, 190), RGB(46, 184, 92), 1.75, 4, False

    AddHCMiniChart cl + (HC_CHART_WIDTH + HC_CHART_GAP) * 2, ct, HC_CHART_WIDTH, HC_CHART_HEIGHT, _
        grpName & " - Reused", "HC_" & CleanChartName(grpName) & "_Reused", _
        hdrRow, lastMCol, reuNeedRow, reuAvailRow, _
        RGB(0, 172, 193), RGB(46, 184, 92), 1.75, 4, False

    WriteHCGroupBlock = r
End Function

'====================================================================
' HC HELPER: WRITE SUMMARY FORMULA ROW (Total Need or Total Available)
' row1 + row2 formula per month column.
'====================================================================

Private Sub WriteHCSummaryFormulaRow(curRow As Long, label As String, _
    lastMCol As Long, row1 As Long, row2 As Long, _
    labelColor As Long, bgColor As Long)

    m_ws.Cells(curRow, COL_START).Value = label
    With m_ws.Cells(curRow, COL_START).Font
        .Size = 9: .Bold = True: .Color = labelColor
    End With
    m_ws.Cells(curRow, COL_START).IndentLevel = 1

    Dim sm As Long, smCol As Long
    For sm = 1 To HC_MONTHS
        smCol = COL_START + sm
        If smCol <= COL_END Then
            SafeFormulaWrite m_ws, curRow, smCol, _
                "=" & ColLetter(smCol) & row1 & "+" & ColLetter(smCol) & row2
            With m_ws.Cells(curRow, smCol)
                .NumberFormat = "0.0"
                .Font.Size = 9
                .Font.Color = TABLE_TEXT
                .HorizontalAlignment = xlCenter
            End With
        End If
    Next sm

    m_ws.Range(m_ws.Cells(curRow, COL_START), m_ws.Cells(curRow, lastMCol)).Interior.Color = bgColor
End Sub

'====================================================================
' HC HELPER: WRITE TOTAL SECTION BLOCK
' Writes TOTAL header + 7 data rows + 2 summary rows + 3 charts.
' Returns the next available row.
'====================================================================

Private Function WriteHCTotalBlock(curRow As Long, hdrRow As Long, _
    lastMCol As Long, chartCol As Long, refs As HCTableRefs, _
    totRows() As Long, mcs() As Long, mce() As Long) As Long

    Dim r As Long: r = curRow + 1  ' extra spacing before TOTAL

    ' --- TOTAL header (executive emphasis) ---
    Dim totSectionStart As Long: totSectionStart = r
    Dim totHdrRng As Range
    Set totHdrRng = m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol))
    totHdrRng.Interior.Color = RGB(218, 237, 248)
    m_ws.Cells(r, COL_START).Value = ChrW(&H25A0) & "  TOTAL  " & ChrW(&H2014) & "  All Groups"
    With m_ws.Cells(r, COL_START).Font
        .Bold = True: .Size = 11
        .Color = RGB(12, 27, 51): .Name = THEME_FONT
    End With
    totHdrRng.Borders(xlEdgeTop).Color = RGB(226, 232, 240)
    totHdrRng.Borders(xlEdgeTop).Weight = xlThin
    totHdrRng.Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
    totHdrRng.Borders(xlEdgeBottom).Weight = xlHairline
    m_ws.Rows(r).RowHeight = 28
    r = r + 1

    ' Track dashboard row positions for summary formulas + charts
    Dim totSecNewNeedRow As Long, totSecReuNeedRow As Long
    Dim totSecNewAvailRow As Long, totSecReuAvailRow As Long

    ' Semantic banding colors (same as group blocks for visual consistency)
    Dim HC_NEED_BG As Long: HC_NEED_BG = RGB(237, 243, 252)
    Dim HC_AVAIL_BG As Long: HC_AVAIL_BG = RGB(235, 251, 242)
    Dim HC_SUMMARY_NEED As Long: HC_SUMMARY_NEED = RGB(222, 234, 250)
    Dim HC_SUMMARY_AVAIL As Long: HC_SUMMARY_AVAIL = RGB(222, 250, 234)

    ' Total New HC Need
    totSecNewNeedRow = r
    WriteHCTotalRowDirect r, "  " & ChrW(&H25B2) & " New HC Need", refs.NewTitle, totRows(0), mcs, mce, lastMCol, RGB(86, 156, 190)
    m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol)).Interior.Color = HC_NEED_BG
    r = r + 1

    ' Total Reuse HC Need
    totSecReuNeedRow = r
    WriteHCTotalRowDirect r, "  " & ChrW(&H25B2) & " Reuse HC Need", refs.ReusedTitle, totRows(1), mcs, mce, lastMCol, RGB(86, 156, 190)
    m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol)).Interior.Color = HC_NEED_BG
    r = r + 1

    ' Total New Available
    totSecNewAvailRow = r
    WriteHCTotalRowDirect r, "  " & ChrW(&H25CF) & " New Available", refs.NewAvailTitle, totRows(3), mcs, mce, lastMCol, RGB(46, 184, 92)
    m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol)).Interior.Color = HC_AVAIL_BG
    r = r + 1

    ' Total Reuse Available
    totSecReuAvailRow = r
    WriteHCTotalRowDirect r, "  " & ChrW(&H25CF) & " Reuse Available", refs.ReuAvailTitle, totRows(4), mcs, mce, lastMCol, RGB(46, 184, 92)
    m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol)).Interior.Color = HC_AVAIL_BG
    r = r + 1

    ' Total New Gap
    WriteHCTotalRowDirect r, "  " & ChrW(&H394) & " New Gap", refs.NewGapTitle, totRows(5), mcs, mce, lastMCol, TABLE_TEXT
    ApplyGapConditionalFormatting r, lastMCol
    r = r + 1

    ' Total Reuse Gap
    WriteHCTotalRowDirect r, "  " & ChrW(&H394) & " Reuse Gap", refs.ReuGapTitle, totRows(6), mcs, mce, lastMCol, TABLE_TEXT
    ApplyGapConditionalFormatting r, lastMCol
    r = r + 1

    ' Total HC Gap (bold, white base for CF)
    WriteHCTotalRowDirect r, "  " & ChrW(&H394) & " Total Gap", refs.TotalGapTitle, totRows(7), mcs, mce, lastMCol, TABLE_TEXT
    m_ws.Cells(r, COL_START).Font.Bold = True
    m_ws.Range(m_ws.Cells(r, COL_START), m_ws.Cells(r, lastMCol)).Interior.Color = THEME_WHITE
    ApplyGapConditionalFormatting r, lastMCol
    r = r + 1

    ' Total Need summary (medium blue band)
    Dim totSecNeedRow As Long: totSecNeedRow = r
    WriteHCSummaryFormulaRow r, "  " & ChrW(&H25B2) & " Total Need", lastMCol, _
        totSecNewNeedRow, totSecReuNeedRow, RGB(239, 83, 80), HC_SUMMARY_NEED
    m_ws.Cells(r, COL_START).Font.Bold = True
    Dim sm As Long, smCol As Long
    For sm = 1 To HC_MONTHS
        smCol = COL_START + sm
        If smCol <= COL_END Then m_ws.Cells(r, smCol).Font.Bold = True
    Next sm
    r = r + 1

    ' Total Available summary (medium green band)
    Dim totSecAvailRow As Long: totSecAvailRow = r
    WriteHCSummaryFormulaRow r, "  " & ChrW(&H25CF) & " Total Available", lastMCol, _
        totSecNewAvailRow, totSecReuAvailRow, RGB(46, 184, 92), HC_SUMMARY_AVAIL
    m_ws.Cells(r, COL_START).Font.Bold = True
    For sm = 1 To HC_MONTHS
        smCol = COL_START + sm
        If smCol <= COL_END Then m_ws.Cells(r, smCol).Font.Bold = True
    Next sm
    r = r + 1

    ' Bottom border
    Dim bottomRng As Range
    Set bottomRng = m_ws.Range(m_ws.Cells(r - 1, COL_START), m_ws.Cells(r - 1, lastMCol))
    bottomRng.Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
    bottomRng.Borders(xlEdgeBottom).Weight = xlThin

    ' --- 3 TOTAL charts (slightly larger for executive view) ---
    Dim cl As Double, ct As Double
    cl = m_ws.Cells(totSectionStart, chartCol).Left
    ct = m_ws.Cells(totSectionStart, chartCol).Top

    AddHCMiniChart cl, ct, HC_TOT_CHART_WIDTH, HC_TOT_CHART_HEIGHT, _
        "ALL - Total", "HC_ALL_Total", _
        hdrRow, lastMCol, totSecNeedRow, totSecAvailRow, _
        RGB(239, 83, 80), RGB(46, 184, 92), 2.5, 5, True

    AddHCMiniChart cl + HC_TOT_CHART_WIDTH + HC_CHART_GAP, ct, HC_TOT_CHART_WIDTH, HC_TOT_CHART_HEIGHT, _
        "ALL - New", "HC_ALL_New", _
        hdrRow, lastMCol, totSecNewNeedRow, totSecNewAvailRow, _
        RGB(86, 156, 190), RGB(46, 184, 92), 2.25, 5, False

    AddHCMiniChart cl + (HC_TOT_CHART_WIDTH + HC_CHART_GAP) * 2, ct, HC_TOT_CHART_WIDTH, HC_TOT_CHART_HEIGHT, _
        "ALL - Reused", "HC_ALL_Reused", _
        hdrRow, lastMCol, totSecReuNeedRow, totSecReuAvailRow, _
        RGB(0, 172, 193), RGB(46, 184, 92), 2.25, 5, False

    WriteHCTotalBlock = r
End Function

'--------------------------------------------------------------------
' HC Gap Helpers: Write rows with direct AVERAGE references (Rev11)
'--------------------------------------------------------------------

Private Sub WriteHCGroupRowDirect(curRow As Long, label As String, _
    grpName As String, titleRow As Long, hcGC As Long, grpRows As Object, _
    mcs() As Long, mce() As Long, lastMCol As Long, labelColor As Long)

    m_ws.Cells(curRow, COL_START).Value = label
    m_ws.Cells(curRow, COL_START).Font.Size = 9
    m_ws.Cells(curRow, COL_START).Font.Color = labelColor
    m_ws.Cells(curRow, COL_START).IndentLevel = 1

    Dim mi As Long, mCol As Long
    For mi = 1 To 12
        mCol = COL_START + mi
        If mCol > COL_END Or mcs(mi) = 0 Then
            m_ws.Cells(curRow, mCol).Value = ""
        Else
            Dim f As String
            f = BuildHCDirectAvg(titleRow, hcGC, grpRows, grpName, mcs(mi), mce(mi))
            If f = "" Then f = "0"
            SafeFormulaWrite m_ws, curRow, mCol, "=" & f
        End If
        m_ws.Cells(curRow, mCol).NumberFormat = "0.0"
        m_ws.Cells(curRow, mCol).Font.Size = 9
        m_ws.Cells(curRow, mCol).Font.Color = TABLE_TEXT
        m_ws.Cells(curRow, mCol).HorizontalAlignment = xlCenter
    Next mi
End Sub

Private Sub WriteHCTotalRowDirect(curRow As Long, label As String, _
    titleRow As Long, totalRow As Long, _
    mcs() As Long, mce() As Long, lastMCol As Long, labelColor As Long)

    m_ws.Cells(curRow, COL_START).Value = label
    m_ws.Cells(curRow, COL_START).Font.Bold = True
    m_ws.Cells(curRow, COL_START).Font.Size = 9
    m_ws.Cells(curRow, COL_START).Font.Color = labelColor

    Dim mi As Long, mCol As Long
    For mi = 1 To 12
        mCol = COL_START + mi
        If mCol > COL_END Or mcs(mi) = 0 Then
            m_ws.Cells(curRow, mCol).Value = ""
        Else
            Dim af As String
            af = BuildHCRowAvg(titleRow, totalRow, mcs(mi), mce(mi))
            If af = "" Then af = "0"
            SafeFormulaWrite m_ws, curRow, mCol, "=" & af
        End If
        m_ws.Cells(curRow, mCol).NumberFormat = "0.0"
        m_ws.Cells(curRow, mCol).Font.Size = 9
        m_ws.Cells(curRow, mCol).Font.Bold = True
        m_ws.Cells(curRow, mCol).Font.Color = TABLE_TEXT
        m_ws.Cells(curRow, mCol).HorizontalAlignment = xlCenter
    Next mi
End Sub

Private Sub ApplyGapConditionalFormatting(curRow As Long, lastMCol As Long)
    ' Conditional formatting on Gap row (green positive, red negative)
    Dim gapRange As Range
    Set gapRange = m_ws.Range(m_ws.Cells(curRow, COL_START + 1), m_ws.Cells(curRow, lastMCol))
    Dim fc As FormatCondition
    Set fc = gapRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    fc.Font.Color = RGB(185, 28, 28)
    fc.Interior.Color = RGB(254, 226, 226)
    fc.Font.Bold = True
    Set fc = gapRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
    fc.Font.Color = RGB(46, 184, 92)
    fc.Interior.Color = RGB(220, 252, 231)
    fc.Font.Bold = True
End Sub

Private Sub AddHCMiniChart(chartLeft As Double, chartTop As Double, _
    chartWidth As Double, chartHeight As Double, _
    chartTitle As String, chartName As String, _
    hdrRow As Long, lastMCol As Long, needRow As Long, availRow As Long, _
    needColor As Long, availColor As Long, lineWeight As Double, _
    markerSize As Long, showLegend As Boolean)

    On Error GoTo CleanFail

    Dim chartObj As ChartObject
    Set chartObj = m_ws.ChartObjects.Add(Left:=chartLeft, Top:=chartTop, _
        Width:=chartWidth, Height:=chartHeight)
    If chartObj Is Nothing Then Exit Sub

    ' Assign deterministic name for cleanup on rebuild
    On Error Resume Next
    chartObj.Name = chartName
    On Error GoTo CleanFail

    Dim catRng As Range
    Set catRng = m_ws.Range(m_ws.Cells(hdrRow, COL_START + 1), m_ws.Cells(hdrRow, lastMCol))

    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = chartTitle
        .ChartTitle.Font.Size = 8
        .ChartTitle.Font.Color = RGB(226, 232, 240)
        .ChartTitle.Font.Name = THEME_FONT

        ' Series 1: Need
        Dim sNeed As Series
        Set sNeed = .SeriesCollection.NewSeries
        sNeed.Name = ChrW(&H25B2) & " Need"
        sNeed.Values = m_ws.Range(m_ws.Cells(needRow, COL_START + 1), m_ws.Cells(needRow, lastMCol))
        sNeed.XValues = catRng
        sNeed.Format.Line.ForeColor.RGB = needColor
        sNeed.Format.Line.Weight = lineWeight
        sNeed.MarkerStyle = xlMarkerStyleCircle
        sNeed.MarkerSize = markerSize

        ' Series 2: Available
        Dim sAvail As Series
        Set sAvail = .SeriesCollection.NewSeries
        sAvail.Name = ChrW(&H25CF) & " Available"
        sAvail.Values = m_ws.Range(m_ws.Cells(availRow, COL_START + 1), m_ws.Cells(availRow, lastMCol))
        sAvail.XValues = catRng
        sAvail.Format.Line.ForeColor.RGB = availColor
        sAvail.Format.Line.Weight = lineWeight
        sAvail.MarkerStyle = xlMarkerStyleCircle
        sAvail.MarkerSize = markerSize

        ' Chart area styling
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(15, 35, 62)
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(12, 27, 51)
        .ChartArea.Format.Line.ForeColor.RGB = RGB(50, 75, 110)
        .ChartArea.Format.Line.Weight = 0.5

        ' Axis styling
        .Axes(xlCategory).TickLabels.Font.Size = 6
        .Axes(xlCategory).TickLabels.Font.Color = RGB(180, 190, 200)
        .Axes(xlCategory).TickLabels.NumberFormat = "mmm"
        .Axes(xlCategory).TickLabelSpacing = 1
        .Axes(xlValue).TickLabels.Font.Size = 6
        .Axes(xlValue).TickLabels.Font.Color = RGB(180, 190, 200)
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(40, 65, 100)

        ' Legend: show only on first chart per group for cleaner layout
        .HasLegend = showLegend
        If showLegend Then
            .Legend.Position = xlLegendPositionBottom
            .Legend.Font.Size = 7
            .Legend.Font.Color = RGB(226, 232, 240)
        End If
    End With

    Exit Sub

CleanFail:
    Err.Clear
End Sub

'====================================================================
' HC HELPER: BUILD GROUP-TO-ROW MAP
' Returns a Dictionary mapping group name (LCase) -> row index in the
' given HC table. If titleRow=0, returns empty dictionary.
'====================================================================

Private Function BuildHCGroupRowMap(titleRow As Long, grpCol As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set BuildHCGroupRowMap = dict
    If titleRow = 0 Then Exit Function

    Dim dsr As Long: dsr = titleRow + 2
    Dim lastRow As Long
    lastRow = GetHCTableDataLastRow(titleRow, grpCol)
    Dim dr As Long
    For dr = dsr To lastRow
        Dim cv As String
        cv = Trim(CStr(m_workSheet.Cells(dr, grpCol).Value))
        If cv = "" Or LCase(cv) = "total" Then Exit For
        dict(LCase(cv)) = dr
    Next dr
End Function

'====================================================================
' HC HELPER: FIND TOTAL ROW in an HC table
' Returns the row containing "Total" in the group column. 0 if not found.
'====================================================================

Private Function FindHCTotalRow(titleRow As Long, grpCol As Long) As Long
    FindHCTotalRow = 0
    If titleRow = 0 Then Exit Function

    Dim dsr As Long: dsr = titleRow + 2
    Dim lastRow As Long
    lastRow = GetHCTableDataLastRow(titleRow, grpCol)
    Dim dr As Long
    For dr = dsr To lastRow
        Dim cv As String
        cv = LCase(Trim(CStr(m_workSheet.Cells(dr, grpCol).Value)))
        If cv = "total" Then
            FindHCTotalRow = dr
            Exit Function
        End If
        If cv = "" Then Exit For
    Next dr
End Function

'====================================================================
' HC HELPER: FIND LAST USED DATA ROW IN AN HC TABLE
' Returns the last non-empty group row, including the Total row when
' present. Removes the previous fixed 50-row cap.
'====================================================================

Private Function GetHCTableDataLastRow(titleRow As Long, grpCol As Long) As Long
    GetHCTableDataLastRow = 0
    If titleRow = 0 Or grpCol = 0 Then Exit Function

    Dim dsr As Long
    dsr = titleRow + 2

    Dim maxRow As Long
    maxRow = m_workSheet.UsedRange.row + m_workSheet.UsedRange.Rows.Count - 1
    If maxRow < dsr Then Exit Function

    Dim dr As Long
    For dr = dsr To maxRow
        Dim cv As String
        cv = Trim(CStr(m_workSheet.Cells(dr, grpCol).Value))
        If cv = "" Then Exit For

        GetHCTableDataLastRow = dr
        If LCase(cv) = "total" Then Exit Function
    Next dr
End Function

'====================================================================
' HC HELPER: BUILD DIRECT AVERAGE FOR A SPECIFIC GROUP ROW
' Returns AVERAGE formula body (no leading =) for a group's row
' in a specific HC table. Looks up group name in the row map.
' Returns "" if table or group not found.
'====================================================================

Private Function BuildHCDirectAvg(titleRow As Long, grpCol As Long, _
    grpRows As Object, grpName As String, _
    colStart As Long, colEnd As Long) As String

    BuildHCDirectAvg = ""
    If titleRow = 0 Then Exit Function
    If Not grpRows.exists(LCase(grpName)) Then Exit Function

    Dim dataRow As Long
    dataRow = CLng(grpRows(LCase(grpName)))

    Dim csL As String, ceL As String
    csL = ColLetter(colStart)
    ceL = ColLetter(colEnd)

    BuildHCDirectAvg = "AVERAGE('" & m_wsName & "'!$" & csL & "$" & dataRow & _
        ":$" & ceL & "$" & dataRow & ")"
End Function

'====================================================================
' HC HELPER: BUILD AVERAGE FOR A SPECIFIC ROW (for totals)
' Returns AVERAGE formula body (no leading =) for a specific row
' in an HC table. Returns "" if titleRow=0 or rowIdx=0.
'====================================================================

Private Function BuildHCRowAvg(titleRow As Long, rowIdx As Long, _
    colStart As Long, colEnd As Long) As String

    BuildHCRowAvg = ""
    If titleRow = 0 Or rowIdx = 0 Then Exit Function

    Dim csL As String, ceL As String
    csL = ColLetter(colStart)
    ceL = ColLetter(colEnd)

    BuildHCRowAvg = "AVERAGE('" & m_wsName & "'!$" & csL & "$" & rowIdx & _
        ":$" & ceL & "$" & rowIdx & ")"
End Function

'====================================================================
' PIVOT INFRASTRUCTURE
' Creates a hidden helper sheet with a derived table containing
' Group, CEID, EntityType, NewReused, ProjectStart per row.
' ProjectStart is computed from milestone start headers (not Set Start).
' PivotCache is created from this helper table.
'====================================================================

Private Sub CreatePivotInfrastructure()
    If m_tbl Is Nothing Then Exit Sub
    If m_nrCol = 0 Then Exit Sub

    ' Old DashHelper + slicer caches already cleaned up in CreateOrClearDashboardSheet

    ' Create hidden helper sheet
    Set m_helperSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    m_helperSheet.Name = "DashHelper"
    m_helperSheet.Visible = xlSheetVeryHidden

    ' Build the helper table with computed ProjectStart
    BuildHelperTable

    ' Create PivotCache from the helper table
    If Not m_helperTable Is Nothing Then
        On Error Resume Next
        Set m_pivotCache = ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=m_helperTable.Range)
        On Error GoTo 0
    End If
End Sub

'====================================================================
' BUILD HELPER TABLE (array-optimized)
' Reads Working Sheet into VBA arrays, computes all derived data in
' memory, then writes the entire table in a single Range.Value call.
' Generates three row types:
'   Monthly     - one per system (for Monthly Activity PivotChart)
'   Active      - one per (system x month) where in-flight
'   Cumulative  - one per (system x month) where started
' Gap-filler rows ensure all months appear in charts.
' Rev11: 11 columns (added Conversion), updated Completed logic.
'====================================================================

Private Sub BuildHelperTable()
    If m_helperSheet Is Nothing Then Exit Sub

    ' --- Write header row (15 columns) ---
    m_helperSheet.Range("A1:O1").Value = Array(HLP_COL_GROUP, HLP_COL_CEID, _
        HLP_COL_ENTITY_TYPE, HLP_COL_NR, HLP_COL_PROJSTART, HLP_COL_PROJEND, _
        HLP_COL_PROJMONTH, HLP_COL_STATUS, HLP_COL_ROWTYPE, HLP_COL_PRESMONTH, _
        HLP_COL_CONVERSION, HLP_COL_MRCLFINISH, HLP_COL_INSTALLQTR, _
        HLP_COL_INSTALLDELTA, HLP_COL_HASSETSTART)

    ' --- Read Working Sheet data into arrays (fast bulk read) ---
    Dim lastCol As Long
    lastCol = m_workSheet.Cells(DATA_START_ROW, m_workSheet.Columns.Count).End(xlToLeft).Column
    Dim nDataRows As Long
    nDataRows = m_lastDataRow - m_firstDataRow + 1
    If nDataRows < 1 Then Exit Sub

    ' Header array for column discovery
    Dim wsHdrs() As Variant
    wsHdrs = m_workSheet.Range(m_workSheet.Cells(DATA_START_ROW, 1), _
        m_workSheet.Cells(DATA_START_ROW, lastCol)).Value

    ' Full data array (all columns x all rows)
    Dim wsData() As Variant
    wsData = m_workSheet.Range(m_workSheet.Cells(m_firstDataRow, 1), _
        m_workSheet.Cells(m_lastDataRow, lastCol)).Value

    ' --- Discover CV Start, Completed, MRCL Finish, Set Start column indices in data array ---
    Dim cvStartColIdx As Long: cvStartColIdx = 0
    Dim statusColIdx As Long: statusColIdx = 0
    Dim ourMrclFIdx As Long: ourMrclFIdx = 0     ' Our Date MRCL.F column
    Dim tisMrclFIdx As Long: tisMrclFIdx = 0     ' TIS MRCL Finish column
    Dim setStartIdx As Long: setStartIdx = 0     ' Set Start column
    Dim dj As Long, dhv As String
    For dj = 1 To lastCol
        dhv = LCase(Trim(Replace(Replace(CStr(wsHdrs(1, dj)), vbLf, ""), vbCr, "")))
        If dhv = "cv start" Or dhv = "convert start" Or dhv = "conversion start" Or dhv = "cvstart" Or dhv = "convertstart" Or dhv = "conversionstart" Then cvStartColIdx = dj
        If dhv = LCase(TIS_COL_STATUS) Then statusColIdx = dj
        If dhv = LCase(TIS_COL_OUR_MRCLF) Then ourMrclFIdx = dj
        If dhv = LCase(TIS_SRC_MRCLF) Then tisMrclFIdx = dj
        If dhv = "set start" Then setStartIdx = dj
    Next dj

    ' --- Discover milestone start columns from header array ---
    Dim msHeaders As Collection
    Set msHeaders = GetMilestoneStartHeaders()

    Dim msColIndices As Collection
    Set msColIndices = New Collection
    Dim msItem As Variant, msName As String, mj As Long, whv As String
    For Each msItem In msHeaders
        msName = CStr(msItem)
        If LCase(msName) = "sdd" Then GoTo NextMsHeader
        If InStr(1, LCase(msName), "prefac", vbTextCompare) > 0 Then GoTo NextMsHeader
        If InStr(1, LCase(msName), "pre-fac", vbTextCompare) > 0 Then GoTo NextMsHeader
        For mj = 1 To lastCol
            whv = Trim(Replace(Replace(CStr(wsHdrs(1, mj)), vbLf, ""), vbCr, ""))
            If StrComp(whv, msName, vbTextCompare) = 0 Then
                msColIndices.Add mj
                Exit For
            End If
        Next mj
NextMsHeader:
    Next msItem

    ' --- Discover Demo/Decon start & finish columns from header array ---
    Dim deconStartCols As Collection, deconFinishCols As Collection
    Set deconStartCols = New Collection
    Set deconFinishCols = New Collection
    For dj = 1 To lastCol
        dhv = LCase(Trim(Replace(Replace(CStr(wsHdrs(1, dj)), vbLf, ""), vbCr, "")))
        If (InStr(1, dhv, "decon", vbTextCompare) > 0 Or _
            InStr(1, dhv, "demo", vbTextCompare) > 0) Then
            If InStr(1, dhv, "start", vbTextCompare) > 0 Then deconStartCols.Add dj
            If InStr(1, dhv, "finish", vbTextCompare) > 0 Or _
               InStr(1, dhv, "end", vbTextCompare) > 0 Then deconFinishCols.Add dj
        End If
    Next dj

    ' --- Supplier Qual Finish column (project end boundary; excludes MRCL duration) ---
    Dim sqFinCol As Long
    sqFinCol = m_sqFinishCol
    If sqFinCol = 0 Then
        For dj = 1 To lastCol
            dhv = LCase(Trim(Replace(Replace(CStr(wsHdrs(1, dj)), vbLf, ""), vbCr, "")))
            If dhv = "supplier qual finish" Or dhv = "supplier qualfinish" Then
                sqFinCol = dj: Exit For
            End If
        Next dj
    End If

    ' --- Pre-allocate output array (15 columns) ---
    Dim maxOut As Long
    maxOut = CLng(nDataRows) * 160 + 2000
    If maxOut < 5000 Then maxOut = 5000
    If maxOut > 600000 Then maxOut = 600000
    Dim outArr() As Variant
    ReDim outArr(1 To maxOut, 1 To 15)
    Dim outIdx As Long: outIdx = 0

    ' Track dates for expanded rows and months for gap fillers
    Dim xMinDate As Date, xMaxDate As Date, xHasDate As Boolean
    xHasDate = False
    Dim allMonths As Object
    Set allMonths = CreateObject("Scripting.Dictionary")
    Dim sysRowEnd As Long  ' last system row index (before gap fillers)

    ' --- Process each Working Sheet row into Monthly output rows ---
    Dim ri As Long
    Dim grpVal As String, ceidVal As String, etVal As String, nrVal As String
    Dim projStart As Variant, projEnd As Variant
    Dim projMonth As String, projStatus As String
    Dim demoMin As Variant, demoMax As Variant
    Dim dcItem As Variant, dcIdx As Long, dcVal As Variant
    Dim mk As Long, msCI As Long, msVal As Variant, sqVal As Variant

    For ri = 1 To nDataRows
        grpVal = "": ceidVal = "": etVal = "": nrVal = ""
        If m_groupCol > 0 Then grpVal = Trim(CStr(wsData(ri, m_groupCol)))
        If m_ceidCol > 0 Then ceidVal = Trim(CStr(wsData(ri, m_ceidCol)))
        If m_entityTypeCol > 0 Then etVal = Trim(CStr(wsData(ri, m_entityTypeCol)))
        If m_nrCol > 0 Then nrVal = Trim(CStr(wsData(ri, m_nrCol)))

        If nrVal = "" And grpVal = "" Then GoTo NextHelperRow

        ' --- Compute ProjectStart ---
        projStart = Empty
        If LCase(nrVal) = "demo" Then
            demoMin = Empty
            For Each dcItem In deconStartCols
                dcIdx = CLng(dcItem): dcVal = wsData(ri, dcIdx)
                If IsDate(dcVal) Then
                    If IsEmpty(demoMin) Then
                        demoMin = CDate(dcVal)
                    ElseIf CDate(dcVal) < CDate(demoMin) Then
                        demoMin = CDate(dcVal)
                    End If
                End If
            Next dcItem
            If Not IsEmpty(demoMin) Then projStart = demoMin
        End If
        If IsEmpty(projStart) Then
            For mk = 1 To msColIndices.Count
                msCI = CLng(msColIndices(mk)): msVal = wsData(ri, msCI)
                If IsDate(msVal) Then projStart = CDate(msVal): Exit For
            Next mk
        End If

        ' --- Compute ProjectEnd ---
        projEnd = Empty
        If LCase(nrVal) = "demo" Then
            demoMax = Empty
            For Each dcItem In deconFinishCols
                dcIdx = CLng(dcItem): dcVal = wsData(ri, dcIdx)
                If IsDate(dcVal) Then
                    If IsEmpty(demoMax) Then
                        demoMax = CDate(dcVal)
                    ElseIf CDate(dcVal) > CDate(demoMax) Then
                        demoMax = CDate(dcVal)
                    End If
                End If
            Next dcItem
            If Not IsEmpty(demoMax) Then projEnd = demoMax
        End If
        ' Use Supplier Qual FINISH as project end (includes SQ, excludes MRCL from Active chart)
        If IsEmpty(projEnd) And sqFinCol > 0 Then
            sqVal = wsData(ri, sqFinCol)
            If IsDate(sqVal) Then projEnd = CDate(sqVal)
        End If

        ' --- Compute ProjectMonth & Status ---
        projMonth = ""
        If Not IsEmpty(projStart) Then projMonth = Format(CDate(projStart), "YYYY-MM")

        ' Status logic: use Status column if available
        projStatus = ""
        If statusColIdx > 0 Then
            Dim statVal As String
            statVal = LCase(Trim(CStr(wsData(ri, statusColIdx))))
            If statVal = "completed" Then
                projStatus = "Completed"
            ElseIf statVal = "on hold" Then
                projStatus = "On Hold"
            ElseIf statVal = "cancelled" Then
                projStatus = "Cancelled"
            ElseIf statVal = "non iq" Then
                projStatus = "Non IQ"
            ElseIf IsEmpty(projStart) Or CDate(projStart) > Date Then
                projStatus = "Not Started"
            Else
                projStatus = "Active"
            End If
        Else
            ' Fallback to date-based logic
            If IsEmpty(projStart) Then
                projStatus = "Not Started"
            ElseIf CDate(projStart) > Date Then
                projStatus = "Not Started"
            ElseIf Not IsEmpty(projEnd) And CDate(projEnd) <= Date Then
                projStatus = "Completed"
            Else
                projStatus = "Active"
            End If
        End If

        ' Track date extremes
        If Not IsEmpty(projStart) Then
            If Not xHasDate Then
                xMinDate = CDate(projStart): xMaxDate = CDate(projStart): xHasDate = True
            End If
            If CDate(projStart) < xMinDate Then xMinDate = CDate(projStart)
            If CDate(projStart) > xMaxDate Then xMaxDate = CDate(projStart)
        End If
        If Not IsEmpty(projEnd) And xHasDate Then
            If CDate(projEnd) > xMaxDate Then xMaxDate = CDate(projEnd)
        End If
        If projMonth <> "" Then allMonths(projMonth) = True

        ' Skip "On Hold", "Cancelled", and "Non IQ" systems entirely from all chart series
        ' Skip "Completed" systems only if they have no end date (would appear active forever)
        If projStatus = "On Hold" Then GoTo NextHelperRow
        If projStatus = "Cancelled" Then GoTo NextHelperRow
        If projStatus = "Non IQ" Then GoTo NextHelperRow
        If projStatus = "Completed" And IsEmpty(projEnd) Then GoTo NextHelperRow

        ' Add Monthly row to output array (15 columns)
        outIdx = outIdx + 1
        outArr(outIdx, 1) = grpVal
        outArr(outIdx, 2) = ceidVal
        outArr(outIdx, 3) = etVal
        outArr(outIdx, 4) = nrVal
        If Not IsEmpty(projStart) Then outArr(outIdx, 5) = CDate(projStart)
        If Not IsEmpty(projEnd) Then outArr(outIdx, 6) = CDate(projEnd)
        outArr(outIdx, 7) = projMonth
        outArr(outIdx, 8) = projStatus
        outArr(outIdx, 9) = "Monthly"
        outArr(outIdx, 10) = projMonth
        ' Column 11: Conversion
        If cvStartColIdx > 0 Then
            If IsDate(wsData(ri, cvStartColIdx)) Then
                outArr(outIdx, 11) = "Yes"
            Else
                outArr(outIdx, 11) = ""
            End If
        Else
            outArr(outIdx, 11) = ""
        End If

        ' Column 12: MRCLFinish — Our Date MRCL.F, fallback to TIS MRCL Finish
        Dim mrclFVal As Variant: mrclFVal = Empty
        If ourMrclFIdx > 0 Then mrclFVal = wsData(ri, ourMrclFIdx)
        If Not IsDate(mrclFVal) Then
            If tisMrclFIdx > 0 Then mrclFVal = wsData(ri, tisMrclFIdx)
        End If
        If IsDate(mrclFVal) Then
            outArr(outIdx, 12) = CDate(mrclFVal)
        End If

        ' Column 13: InstallQtr — YYYY-Q# from MRCLFinish (for New/Reused);
        '   for Demo, use earliest Decon Start date, fallback to MRCLFinish
        Dim ibDate As Variant: ibDate = Empty
        If LCase(nrVal) = "demo" Then
            ' Use earliest Decon Start for Demo
            If Not IsEmpty(demoMin) Then
                ibDate = demoMin
            ElseIf IsDate(mrclFVal) Then
                ibDate = mrclFVal
            End If
        Else
            ibDate = mrclFVal
        End If
        If IsDate(ibDate) Then
            outArr(outIdx, 13) = CStr(Year(CDate(ibDate))) & "-Q" & _
                CStr(Int((Month(CDate(ibDate)) - 1) / 3) + 1)
        End If

        ' Column 14: InstallDelta — +1 for New/Reused, -1 for Demo
        If IsDate(ibDate) Then
            If LCase(nrVal) = "demo" Then
                outArr(outIdx, 14) = -1
            Else
                outArr(outIdx, 14) = 1
            End If
        Else
            outArr(outIdx, 14) = 0
        End If

        ' Column 15: HasSetStart — "Yes" if Set Start date exists, Demos always "Yes"
        If LCase(nrVal) = "demo" Then
            outArr(outIdx, 15) = "Yes"
        ElseIf setStartIdx > 0 Then
            If IsDate(wsData(ri, setStartIdx)) Then
                outArr(outIdx, 15) = "Yes"
            Else
                outArr(outIdx, 15) = "No"
            End If
        Else
            outArr(outIdx, 15) = ""
        End If
NextHelperRow:
    Next ri

    sysRowEnd = outIdx  ' mark end of system rows

    ' --- Monthly gap fillers (ensure zero-count months appear) ---
    If allMonths.Count > 1 Then
        Dim minMo As String, maxMo As String
        minMo = "": maxMo = ""
        Dim mKey As Variant
        For Each mKey In allMonths.keys
            If minMo = "" Or CStr(mKey) < minMo Then minMo = CStr(mKey)
            If maxMo = "" Or CStr(mKey) > maxMo Then maxMo = CStr(mKey)
        Next mKey

        Dim yyG As Long, mmG As Long, yyEnd As Long, mmEnd As Long
        yyG = CLng(Left(minMo, 4)): mmG = CLng(Mid(minMo, 6, 2))
        yyEnd = CLng(Left(maxMo, 4)): mmEnd = CLng(Mid(maxMo, 6, 2))
        Dim nrFill As Variant, nfi As Long, testMo As String

        Do While yyG < yyEnd Or (yyG = yyEnd And mmG <= mmEnd)
            testMo = CStr(yyG) & "-" & Format(mmG, "00")
            If Not allMonths.exists(testMo) Then
                nrFill = Array("New", "Reused", "Demo")
                For nfi = 0 To 2
                    outIdx = outIdx + 1
                    If outIdx > maxOut Then Exit Do
                    outArr(outIdx, 4) = nrFill(nfi)
                    outArr(outIdx, 7) = testMo
                    outArr(outIdx, 9) = "Monthly"
                    outArr(outIdx, 10) = testMo
                    outArr(outIdx, 11) = ""
                Next nfi
            End If
            mmG = mmG + 1
            If mmG > 12 Then mmG = 1: yyG = yyG + 1
        Loop
    End If

    ' --- Expanded Active/Cumulative rows (all in-memory from outArr) ---
    If xHasDate Then
        xMinDate = DateSerial(Year(xMinDate), Month(xMinDate), 1)
        xMaxDate = DateSerial(Year(xMaxDate), Month(xMaxDate) + 1, 1)
        Dim xMonthCount As Long
        xMonthCount = DateDiff("m", xMinDate, xMaxDate)
        If xMonthCount < 1 Then xMonthCount = 1
        If xMonthCount > 72 Then xMonthCount = 72

        Dim activeMonths As Object, cumMonths As Object
        Set activeMonths = CreateObject("Scripting.Dictionary")
        Set cumMonths = CreateObject("Scripting.Dictionary")

        Dim xr As Long, xmi As Long
        Dim mStart As Date, mEndD As Date, xPM As String
        Dim xIsActive As Boolean

        For xr = 1 To sysRowEnd
            ' Skip rows with empty CEID or no start date
            If IsEmpty(outArr(xr, 2)) Or CStr(outArr(xr, 2)) = "" Then GoTo NextExpandRow
            If IsEmpty(outArr(xr, 5)) Then GoTo NextExpandRow
            ' Skip "On Hold" and "Cancelled" systems entirely from Active/Cumulative expansion
            If LCase(Trim(CStr(outArr(xr, 8)))) = "on hold" Then GoTo NextExpandRow
            If LCase(Trim(CStr(outArr(xr, 8)))) = "cancelled" Then GoTo NextExpandRow
            ' Skip Completed systems with no end date (they would never drop off the chart)
            If CStr(outArr(xr, 8)) = "Completed" And IsEmpty(outArr(xr, 6)) Then GoTo NextExpandRow

            For xmi = 0 To xMonthCount - 1
                mStart = DateAdd("m", xmi, xMinDate)
                mEndD = DateAdd("m", xmi + 1, xMinDate)
                xPM = Format(mStart, "YYYY-MM")

                If CDate(outArr(xr, 5)) < mEndD Then
                    ' Cumulative row
                    outIdx = outIdx + 1
                    If outIdx > maxOut Then GoTo OverflowExit
                    outArr(outIdx, 1) = outArr(xr, 1)
                    outArr(outIdx, 2) = outArr(xr, 2)
                    outArr(outIdx, 3) = outArr(xr, 3)
                    outArr(outIdx, 4) = outArr(xr, 4)
                    outArr(outIdx, 9) = "Cumulative"
                    outArr(outIdx, 10) = xPM
                    outArr(outIdx, 11) = ""
                    cumMonths(xPM) = True

                    ' Active row (if not ended before this month)
                    xIsActive = True
                    If Not IsEmpty(outArr(xr, 6)) Then
                        If IsDate(outArr(xr, 6)) Then
                            If CDate(outArr(xr, 6)) < mStart Then xIsActive = False
                        End If
                    End If
                    If xIsActive Then
                        outIdx = outIdx + 1
                        If outIdx > maxOut Then GoTo OverflowExit
                        outArr(outIdx, 1) = outArr(xr, 1)
                        outArr(outIdx, 2) = outArr(xr, 2)
                        outArr(outIdx, 3) = outArr(xr, 3)
                        outArr(outIdx, 4) = outArr(xr, 4)
                        outArr(outIdx, 9) = "Active"
                        outArr(outIdx, 10) = xPM
                        outArr(outIdx, 11) = ""
                        activeMonths(xPM) = True
                    End If
                End If
            Next xmi
NextExpandRow:
        Next xr

        ' Gap fillers for Active/Cumulative
        Dim nrFill2 As Variant, nfi2 As Long
        nrFill2 = Array("New", "Reused", "Demo")
        For xmi = 0 To xMonthCount - 1
            mStart = DateAdd("m", xmi, xMinDate)
            xPM = Format(mStart, "YYYY-MM")
            If Not cumMonths.exists(xPM) Then
                For nfi2 = 0 To 2
                    outIdx = outIdx + 1
                    outArr(outIdx, 4) = nrFill2(nfi2)
                    outArr(outIdx, 9) = "Cumulative"
                    outArr(outIdx, 10) = xPM
                    outArr(outIdx, 11) = ""
                Next nfi2
            End If
            If Not activeMonths.exists(xPM) Then
                For nfi2 = 0 To 2
                    outIdx = outIdx + 1
                    outArr(outIdx, 4) = nrFill2(nfi2)
                    outArr(outIdx, 9) = "Active"
                    outArr(outIdx, 10) = xPM
                    outArr(outIdx, 11) = ""
                Next nfi2
            End If
        Next xmi
    End If

    ' --- Write output array to helper sheet in a single operation ---
OverflowExit:
    If outIdx > 0 Then
        Dim finalArr() As Variant
        ReDim finalArr(1 To outIdx, 1 To 15)
        Dim wi As Long, wj As Long
        For wi = 1 To outIdx
            For wj = 1 To 15
                finalArr(wi, wj) = outArr(wi, wj)
            Next wj
        Next wi
        m_helperSheet.Range(m_helperSheet.Cells(2, 1), _
            m_helperSheet.Cells(outIdx + 1, 15)).Value = finalArr

        ' Format date columns
        m_helperSheet.Range(m_helperSheet.Cells(2, 5), _
            m_helperSheet.Cells(outIdx + 1, 5)).NumberFormat = "yyyy-mm-dd"
        m_helperSheet.Range(m_helperSheet.Cells(2, 6), _
            m_helperSheet.Cells(outIdx + 1, 6)).NumberFormat = "yyyy-mm-dd"
        m_helperSheet.Range(m_helperSheet.Cells(2, 12), _
            m_helperSheet.Cells(outIdx + 1, 12)).NumberFormat = "yyyy-mm-dd"

        ' Create ListObject (15 columns A:O)
        Dim tblRange As Range
        Set tblRange = m_helperSheet.Range(m_helperSheet.Cells(1, 1), _
            m_helperSheet.Cells(outIdx + 1, 15))
        Set m_helperTable = m_helperSheet.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        m_helperTable.Name = "DashHelperTable"
    End If
End Sub

'====================================================================
' SECTION 5: DASHBOARD SLICERS
' Creates slicers from Monthly Activity PivotTable for Group, CEID,
' Entity Type. Group Breakdown is no longer a PivotChart, so slicers
' only connect to PT_Monthly.
' Must run AFTER chart subs so PivotTables exist.
'====================================================================

Private Sub BuildDashboardSlicers()
    ' Need the Monthly PivotTable to create slicers
    If m_ptMonthly Is Nothing Then Exit Sub

    Dim slicerRow As Long
    slicerRow = m_chartSectionRow + 2
    Dim slicerLeft As Double
    Dim slicerWidth As Double, slicerHeight As Double
    slicerWidth = 150
    slicerHeight = 160
    slicerLeft = m_ws.Cells(slicerRow, COL_START).Left

    On Error Resume Next

    ' Group slicer
    If m_groupCol > 0 Then
        Dim scGrp As SlicerCache
        Set scGrp = ThisWorkbook.SlicerCaches.Add2(m_ptMonthly, HLP_COL_GROUP)
        If Not scGrp Is Nothing Then
            scGrp.Slicers.Add m_ws, , "Slicer_Group", "Group", _
                m_ws.Cells(slicerRow, COL_START).Top, slicerLeft, slicerWidth, slicerHeight
            ' Connect to Active Systems PivotTable too (same PivotCache)
            If Not m_ptActive Is Nothing Then scGrp.PivotTables.AddPivotTable m_ptActive
            slicerLeft = slicerLeft + slicerWidth + 10
        End If
    End If

    ' New/Reused/Demo slicer
    If m_nrCol > 0 Then
        Dim scNR As SlicerCache
        Set scNR = ThisWorkbook.SlicerCaches.Add2(m_ptMonthly, HLP_COL_NR)
        If Not scNR Is Nothing Then
            scNR.Slicers.Add m_ws, , "Slicer_NewReused", "New / Reused / Demo", _
                m_ws.Cells(slicerRow, COL_START).Top, slicerLeft, slicerWidth, slicerHeight
            If Not m_ptActive Is Nothing Then scNR.PivotTables.AddPivotTable m_ptActive
            slicerLeft = slicerLeft + slicerWidth + 10
        End If
    End If

    ' CEID slicer
    If m_ceidCol > 0 Then
        Dim scCeid As SlicerCache
        Set scCeid = ThisWorkbook.SlicerCaches.Add2(m_ptMonthly, HLP_COL_CEID)
        If Not scCeid Is Nothing Then
            scCeid.Slicers.Add m_ws, , "Slicer_CEID", "CEID", _
                m_ws.Cells(slicerRow, COL_START).Top, slicerLeft, slicerWidth, slicerHeight
            If Not m_ptActive Is Nothing Then scCeid.PivotTables.AddPivotTable m_ptActive
            slicerLeft = slicerLeft + slicerWidth + 10
        End If
    End If

    ' Entity Type slicer
    If m_entityTypeCol > 0 Then
        Dim scET As SlicerCache
        Set scET = ThisWorkbook.SlicerCaches.Add2(m_ptMonthly, HLP_COL_ENTITY_TYPE)
        If Not scET Is Nothing Then
            scET.Slicers.Add m_ws, , "Slicer_EntityType", "Entity Type", _
                m_ws.Cells(slicerRow, COL_START).Top, slicerLeft, slicerWidth, slicerHeight
            If Not m_ptActive Is Nothing Then scET.PivotTables.AddPivotTable m_ptActive
        End If
    End If

    On Error GoTo 0
End Sub

'====================================================================
' SECTION 6A: MONTHLY ACTIVITY CHART (PivotChart)
' Creates a PivotTable on hidden DashHelper sheet using the helper
' table. Uses text ProjectMonth (YYYY-MM) as row field - avoids
' PivotTable date grouping issues. Stacked bar with New/Reused/Demo.
' Slicer-connected through shared PivotCache.
' Rev11: Filters months before m_chartStartDate.
'====================================================================

Private Sub BuildMonthlyActivityChart()
    Dim startRow As Long
    startRow = m_nextRow + 1

    WriteSectionTitle startRow, "Monthly Activity", _
        "New/Reused/Demo systems added per month (by computed Project Start)"

    If m_nrCol = 0 Then
        m_ws.Cells(startRow + 2, COL_START).Value = "New/Reused column not found."
        m_ws.Cells(startRow + 2, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 4
        Exit Sub
    End If

    If m_pivotCache Is Nothing Or m_helperSheet Is Nothing Then
        m_ws.Cells(startRow + 2, COL_START).Value = "Pivot infrastructure not available."
        m_ws.Cells(startRow + 2, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 4
        Exit Sub
    End If

    ' Create PivotTable on hidden helper sheet (place after helper table range)
    Dim ptDest As Range
    Dim ptStartRow As Long
    ptStartRow = m_helperTable.Range.row + m_helperTable.Range.Rows.Count + 2
    Set ptDest = m_helperSheet.Cells(ptStartRow, 10)  ' Column J, away from helper table

    On Error Resume Next
    Set m_ptMonthly = m_pivotCache.CreatePivotTable( _
        TableDestination:=ptDest, _
        TableName:="PT_Monthly")
    On Error GoTo 0

    If m_ptMonthly Is Nothing Then
        m_ws.Cells(startRow + 2, COL_START).Value = "Could not create Monthly PivotTable."
        m_ws.Cells(startRow + 2, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 4
        Exit Sub
    End If

    ' Configure PivotTable fields
    On Error Resume Next
    With m_ptMonthly
        .ManualUpdate = True

        ' Row field: ProjectMonth (text YYYY-MM - no date grouping needed)
        .PivotFields(HLP_COL_PROJMONTH).Orientation = xlRowField
        .PivotFields(HLP_COL_PROJMONTH).Position = 1
        .PivotFields(HLP_COL_PROJMONTH).ShowAllItems = True

        ' Column field: NewReused (for stacked colors)
        .PivotFields(HLP_COL_NR).Orientation = xlColumnField
        .PivotFields(HLP_COL_NR).Position = 1
        .PivotFields(HLP_COL_NR).ShowAllItems = True

        ' Data field: Count of CEID
        .AddDataField .PivotFields(HLP_COL_CEID), "Count", xlCount

        ' Page field: RowType = "Monthly" (excludes Active/Cumulative rows)
        .PivotFields(HLP_COL_ROWTYPE).Orientation = xlPageField
        .PivotFields(HLP_COL_ROWTYPE).CurrentPage = "Monthly"

        .ManualUpdate = False
    End With

    ' Hide (blank) items in ProjectMonth if present (two-phase: show all, then hide)
    Dim pi As PivotItem
    m_ptMonthly.ManualUpdate = True
    For Each pi In m_ptMonthly.PivotFields(HLP_COL_PROJMONTH).PivotItems
        pi.Visible = True
    Next pi
    For Each pi In m_ptMonthly.PivotFields(HLP_COL_PROJMONTH).PivotItems
        If pi.Name = "" Or LCase(pi.Name) = "(blank)" Then
            On Error Resume Next
            pi.Visible = False
            On Error GoTo 0
        End If
    Next pi
    m_ptMonthly.ManualUpdate = False

    ' Hide (blank) items in NewReused (two-phase)
    m_ptMonthly.ManualUpdate = True
    For Each pi In m_ptMonthly.PivotFields(HLP_COL_NR).PivotItems
        pi.Visible = True
    Next pi
    For Each pi In m_ptMonthly.PivotFields(HLP_COL_NR).PivotItems
        If pi.Name = "" Or LCase(pi.Name) = "(blank)" Then
            On Error Resume Next
            pi.Visible = False
            On Error GoTo 0
        End If
    Next pi
    m_ptMonthly.ManualUpdate = False

    ' Filter months before chart start date (two-phase)
    Dim startMonthStr As String
    startMonthStr = Format(m_chartStartDate, "YYYY-MM")
    Dim pf As PivotField
    Set pf = m_ptMonthly.PivotFields(HLP_COL_PROJMONTH)
    m_ptMonthly.ManualUpdate = True
    For Each pi In pf.PivotItems
        pi.Visible = True
    Next pi
    For Each pi In pf.PivotItems
        If pi.Name = "" Or LCase(pi.Name) = "(blank)" Then
            On Error Resume Next
            pi.Visible = False
            On Error GoTo 0
        ElseIf pi.Name < startMonthStr Then
            On Error Resume Next
            pi.Visible = False
            On Error GoTo 0
        End If
    Next pi
    m_ptMonthly.ManualUpdate = False
    On Error GoTo 0

    ' Create PivotChart on Dashboard (ScreenUpdating must be True for PivotChart)
    Application.ScreenUpdating = True
    On Error GoTo MonthlyChartErr

    Dim chartRow As Long
    chartRow = startRow + 2
    Dim chartWidth As Double
    chartWidth = m_ws.Cells(1, COL_END + 1).Left - m_ws.Cells(1, COL_START).Left
    Dim chartHeight As Double
    chartHeight = CHART_HEIGHT_ROWS * 18

    Dim cht As ChartObject
    Set cht = m_ws.ChartObjects.Add( _
        Left:=m_ws.Cells(chartRow, COL_START).Left, _
        Top:=m_ws.Cells(chartRow, COL_START).Top, _
        Width:=chartWidth, Height:=chartHeight)

    ' Link chart to PivotTable (makes it a PivotChart)
    cht.Chart.SetSourceData m_ptMonthly.TableRange2
    cht.Chart.ChartType = xlColumnStacked

    ' Style the PivotChart
    With cht.Chart
        .HasTitle = True
        .ChartTitle.Text = "Monthly Systems Added"
        .ChartTitle.Font.Color = RGB(226, 232, 240)
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Name = THEME_FONT

        .PlotArea.Format.Fill.ForeColor.RGB = RGB(15, 35, 62)
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(12, 27, 51)

        ' Color series by New/Reused/Demo type (brand palette)
        Dim srs As Series
        For Each srs In .SeriesCollection
            Select Case LCase(srs.Name)
                Case "new": srs.Format.Fill.ForeColor.RGB = THEME_SUCCESS      ' Emerald
                Case "reused": srs.Format.Fill.ForeColor.RGB = THEME_ACCENT2     ' Teal
                Case "demo": srs.Format.Fill.ForeColor.RGB = THEME_DANGER      ' Coral Red
            End Select
        Next srs

        On Error Resume Next
        .Axes(xlCategory).TickLabels.Font.Color = RGB(180, 190, 200)
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabelSpacing = 1
        .Axes(xlValue).TickLabels.Font.Color = RGB(180, 190, 200)
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(50, 75, 110)
        .Axes(xlValue).Format.Line.ForeColor.RGB = RGB(50, 75, 110)
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(40, 65, 100)
        On Error GoTo MonthlyChartErr

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Color = RGB(226, 232, 240)
        .Legend.Font.Size = 9

        ' Hide PivotChart field buttons (cleaner look)
        On Error Resume Next
        .ShowAllFieldButtons = False
        On Error GoTo 0
    End With

    m_nextRow = chartRow + CHART_HEIGHT_ROWS + 2
    Application.ScreenUpdating = False
    Exit Sub

MonthlyChartErr:
    Application.ScreenUpdating = False
    DebugLog "DashboardBuilder: Monthly Activity chart error: " & Err.Description
    m_nextRow = startRow + 4
End Sub

'====================================================================
' SECTION 6B: ACTIVE SYSTEMS CHART (PivotChart)
' Stacked area chart - slicer-connected through shared PivotCache.
' Uses expanded rows in helper table (RowType = Active / Cumulative).
' PresenceMonth as row field, NewReused as column field.
' RowType page field toggles Active vs Cumulative view.
' Rev11: Filters months before m_chartStartDate.
'====================================================================

Private Sub BuildActiveSystemsChart()
    Dim startRow As Long
    startRow = m_nextRow + 1

    WriteSectionTitle startRow, "Active Systems", _
        "Slicer-filtered  |  Use RowType dropdown on chart to toggle Active / Cumulative"

    If m_nrCol = 0 Then
        m_ws.Cells(startRow + 2, COL_START).Value = "New/Reused column not found."
        m_ws.Cells(startRow + 2, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 4
        Exit Sub
    End If

    If m_pivotCache Is Nothing Or m_helperSheet Is Nothing Then
        m_ws.Cells(startRow + 2, COL_START).Value = "Pivot infrastructure not available."
        m_ws.Cells(startRow + 2, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 4
        Exit Sub
    End If

    ' Create PivotTable on hidden helper sheet (place below existing content)
    Dim ptRow As Long
    ptRow = m_helperSheet.Cells(m_helperSheet.Rows.Count, 10).End(xlUp).row + 3
    Dim ptRow2 As Long
    ptRow2 = m_helperSheet.Cells(m_helperSheet.Rows.Count, 15).End(xlUp).row + 3
    If ptRow2 > ptRow Then ptRow = ptRow2
    Dim ptDest As Range
    Set ptDest = m_helperSheet.Cells(ptRow, 15)  ' Column O, away from helper table

    On Error Resume Next
    Set m_ptActive = m_pivotCache.CreatePivotTable( _
        TableDestination:=ptDest, _
        TableName:="PT_Active")
    On Error GoTo 0

    If m_ptActive Is Nothing Then
        m_ws.Cells(startRow + 2, COL_START).Value = "Could not create Active Systems PivotTable."
        m_ws.Cells(startRow + 2, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 4
        Exit Sub
    End If

    ' Configure PivotTable fields
    On Error Resume Next
    With m_ptActive
        .ManualUpdate = True

        ' Page/filter field: RowType - default to "Active"
        .PivotFields(HLP_COL_ROWTYPE).Orientation = xlPageField
        .PivotFields(HLP_COL_ROWTYPE).CurrentPage = "Active"

        ' Row field: PresenceMonth (YYYY-MM text, sorts chronologically)
        .PivotFields(HLP_COL_PRESMONTH).Orientation = xlRowField
        .PivotFields(HLP_COL_PRESMONTH).Position = 1
        .PivotFields(HLP_COL_PRESMONTH).ShowAllItems = True

        ' Column field: NewReused (for stacked colors)
        .PivotFields(HLP_COL_NR).Orientation = xlColumnField
        .PivotFields(HLP_COL_NR).Position = 1
        .PivotFields(HLP_COL_NR).ShowAllItems = True

        ' Data field: Count of CEID
        .AddDataField .PivotFields(HLP_COL_CEID), "Count", xlCount

        .ManualUpdate = False
    End With

    ' Hide (blank) items in PresenceMonth (two-phase: show all, then hide)
    Dim pi As PivotItem
    m_ptActive.ManualUpdate = True
    For Each pi In m_ptActive.PivotFields(HLP_COL_PRESMONTH).PivotItems
        pi.Visible = True
    Next pi
    For Each pi In m_ptActive.PivotFields(HLP_COL_PRESMONTH).PivotItems
        If pi.Name = "" Or LCase(pi.Name) = "(blank)" Then
            On Error Resume Next
            pi.Visible = False
            On Error GoTo 0
        End If
    Next pi
    m_ptActive.ManualUpdate = False

    ' Hide (blank) items in NewReused (two-phase)
    m_ptActive.ManualUpdate = True
    For Each pi In m_ptActive.PivotFields(HLP_COL_NR).PivotItems
        pi.Visible = True
    Next pi
    For Each pi In m_ptActive.PivotFields(HLP_COL_NR).PivotItems
        If pi.Name = "" Or LCase(pi.Name) = "(blank)" Then
            On Error Resume Next
            pi.Visible = False
            On Error GoTo 0
        End If
    Next pi
    m_ptActive.ManualUpdate = False

    ' Filter months before chart start date (two-phase)
    Dim startMonthStr As String
    startMonthStr = Format(m_chartStartDate, "YYYY-MM")
    Dim pf As PivotField
    Set pf = m_ptActive.PivotFields(HLP_COL_PRESMONTH)
    m_ptActive.ManualUpdate = True
    For Each pi In pf.PivotItems
        pi.Visible = True
    Next pi
    For Each pi In pf.PivotItems
        If pi.Name = "" Or LCase(pi.Name) = "(blank)" Then
            On Error Resume Next
            pi.Visible = False
            On Error GoTo 0
        ElseIf pi.Name < startMonthStr Then
            On Error Resume Next
            pi.Visible = False
            On Error GoTo 0
        End If
    Next pi
    m_ptActive.ManualUpdate = False
    On Error GoTo 0

    ' Create PivotChart on Dashboard (ScreenUpdating must be True for PivotChart)
    Application.ScreenUpdating = True
    On Error GoTo ActiveChartErr

    Dim chartRow As Long
    chartRow = startRow + 2
    Dim chartWidth As Double
    chartWidth = m_ws.Cells(1, COL_END + 1).Left - m_ws.Cells(1, COL_START).Left
    Dim chartHeight As Double
    chartHeight = CHART_HEIGHT_ROWS * 18

    Dim chtAS As ChartObject
    Set chtAS = m_ws.ChartObjects.Add( _
        Left:=m_ws.Cells(chartRow, COL_START).Left, _
        Top:=m_ws.Cells(chartRow, COL_START).Top, _
        Width:=chartWidth, Height:=chartHeight)

    ' Link chart to PivotTable (makes it a PivotChart)
    chtAS.Chart.SetSourceData m_ptActive.TableRange2
    chtAS.Chart.ChartType = xlAreaStacked

    ' Style the PivotChart
    With chtAS.Chart
        .HasTitle = True
        .ChartTitle.Text = "Systems Over Time"
        .ChartTitle.Font.Color = RGB(226, 232, 240)
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Name = THEME_FONT

        .PlotArea.Format.Fill.ForeColor.RGB = RGB(15, 35, 62)
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(12, 27, 51)

        ' Color series by New/Reused/Demo type (brand palette)
        Dim srs As Series
        For Each srs In .SeriesCollection
            Select Case LCase(srs.Name)
                Case "new": srs.Format.Fill.ForeColor.RGB = THEME_SUCCESS      ' Emerald
                Case "reused": srs.Format.Fill.ForeColor.RGB = THEME_ACCENT2     ' Teal
                Case "demo": srs.Format.Fill.ForeColor.RGB = THEME_DANGER      ' Coral Red
            End Select
        Next srs

        On Error Resume Next
        .Axes(xlCategory).TickLabels.Font.Color = RGB(180, 190, 200)
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabelSpacing = 1
        .Axes(xlValue).TickLabels.Font.Color = RGB(180, 190, 200)
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(50, 75, 110)
        .Axes(xlValue).Format.Line.ForeColor.RGB = RGB(50, 75, 110)
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(40, 65, 100)
        On Error GoTo ActiveChartErr

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Color = RGB(226, 232, 240)
        .Legend.Font.Size = 9

        ' Show only report filter button (Active/Cumulative toggle)
        On Error Resume Next
        .ShowReportFilterFieldButtons = True
        .ShowAxisFieldButtons = False
        .ShowLegendFieldButtons = False
        .ShowValueFieldButtons = False
        On Error GoTo 0
    End With

    m_nextRow = chartRow + CHART_HEIGHT_ROWS + 2
    Application.ScreenUpdating = False
    Exit Sub

ActiveChartErr:
    Application.ScreenUpdating = False
    DebugLog "DashboardBuilder: Active Systems chart error: " & Err.Description
    m_nextRow = startRow + 4
End Sub

'====================================================================
' SECTION 7: GROUP BREAKDOWN (Horizontal Bar Chart)
' Regular chart (not PivotChart) showing system count per group,
' stacked by New/Reused/Demo. Horizontal bars (xlBarStacked).
' All/Currently Active dropdown toggles between showing all systems
' or only active ones. Uses SUMPRODUCT on Working Sheet table with
' hasDateFilter (matches KPI + System Counter logic).
' Group order matches System Counters section (cached CollectGroups).
'====================================================================

Private Sub BuildGroupBreakdown()
    Dim startRow As Long
    startRow = m_nextRow + 1

    WriteSectionTitle startRow, "Group Breakdown", "System count per group split by New / Reused / Demo"

    If m_groupCol = 0 Or m_nrCol = 0 Then
        m_ws.Cells(startRow + 2, COL_START).Value = "Group or New/Reused column not found."
        m_ws.Cells(startRow + 2, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 4
        Exit Sub
    End If

    ' --- Active toggle dropdown ---
    Dim dropRow As Long: dropRow = startRow + 2
    m_ws.Cells(dropRow, COL_START).Value = "Show:"
    m_ws.Cells(dropRow, COL_START).Font.Bold = True
    m_ws.Cells(dropRow, COL_START).Font.Size = 10
    m_ws.Cells(dropRow, COL_START).Font.Color = RGB(100, 116, 139)
    With m_ws.Cells(dropRow, COL_START + 1)
        .Value = "All"
        .Font.Bold = True: .Font.Size = 10
        .Font.Color = RGB(30, 41, 59)
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        With .Borders
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225)
        End With
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="All,Currently Active"
            .IgnoreBlank = True: .InCellDropdown = True
        End With
    End With
    If Not m_dropdownCells Is Nothing Then m_dropdownCells.Add m_ws.Cells(dropRow, COL_START + 1)
    Dim ddRef As String
    ddRef = "$" & ColLetter(COL_START + 1) & "$" & dropRow

    ' Use cached groups (same order as System Counters)
    Dim groups As Collection
    Set groups = CollectGroups()
    If groups.Count = 0 Then
        m_ws.Cells(startRow + 4, COL_START).Value = "No groups found."
        m_ws.Cells(startRow + 4, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 6
        Exit Sub
    End If

    ' Build Working Sheet structured table references
    Dim grpRange As String, nrRange As String
    grpRange = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_groupCol).Value & "]"
    nrRange = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_nrCol).Value & "]"

    ' Completed/Cancelled/Non IQ exclusion filter
    Dim compExcl As String: compExcl = ""
    If m_statusCol > 0 Then
        Dim statusRef2 As String
        statusRef2 = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_statusCol).Value & "]"
        compExcl = "*(" & statusRef2 & "<>""Completed"")*(" & statusRef2 & "<>""Cancelled"")*(" & statusRef2 & "<>""Non IQ"")"
    End If

    ' Active filter: Set Start <= TODAY AND SQ Finish >= TODAY (matches Working Sheet)
    Dim activeFilter As String: activeFilter = ""
    If m_setStartCol > 0 And m_sqFinishCol > 0 Then
        Dim ssRange As String, sqfRange As String
        ssRange = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_setStartCol).Value & "]"
        sqfRange = m_tblName & "[" & m_workSheet.Cells(DATA_START_ROW, m_sqFinishCol).Value & "]"
        activeFilter = "*(ISNUMBER(" & ssRange & "))*(" & ssRange & "<=TODAY())*(ISNUMBER(" & sqfRange & "))*(" & sqfRange & ">=TODAY())"
    End If

    ' Build data table for chart
    Dim tblRow As Long
    tblRow = dropRow + 2
    m_ws.Cells(tblRow, COL_START).Value = "Group"
    m_ws.Cells(tblRow + 1, COL_START).Value = "New"
    m_ws.Cells(tblRow + 2, COL_START).Value = "Reused"
    m_ws.Cells(tblRow + 3, COL_START).Value = "Demo"
    Dim lb As Long
    For lb = 0 To 3
        m_ws.Cells(tblRow + lb, COL_START).Font.Bold = True
        m_ws.Cells(tblRow + lb, COL_START).Font.Size = 9
        m_ws.Cells(tblRow + lb, COL_START).Font.Color = RGB(100, 116, 139)
    Next lb

    ' For each group: SUMPRODUCT per type with active toggle
    Dim gi As Long, gName As String
    For gi = 1 To groups.Count
        Dim gCol As Long
        gCol = COL_START + gi
        If gCol > COL_END Then Exit For

        gName = CStr(groups(gi))
        m_ws.Cells(tblRow, gCol).Value = gName
        m_ws.Cells(tblRow, gCol).Font.Bold = True
        m_ws.Cells(tblRow, gCol).Font.Size = 8
        m_ws.Cells(tblRow, gCol).Font.Color = RGB(100, 116, 139)
        m_ws.Cells(tblRow, gCol).HorizontalAlignment = xlCenter

        Dim types As Variant
        types = Array("New", "Reused", "Demo")
        Dim ti2 As Long
        For ti2 = 0 To 2
            Dim gFormula As String
            Dim baseExpr As String
            baseExpr = "(" & grpRange & "=""" & gName & """)*(" & nrRange & "=""" & types(ti2) & """)"

            ' "All" mode: with hasDateFilter for New/Reused, no filter for Demo
            Dim allExpr As String
            If types(ti2) = "Demo" Then
                allExpr = "SUMPRODUCT(" & baseExpr & compExcl & "*1)"
            Else
                allExpr = "SUMPRODUCT(" & baseExpr & compExcl & m_hasDateFilter & ")"
            End If

            ' "Currently Active" mode: with activeFilter
            Dim actExpr As String
            If activeFilter <> "" Then
                actExpr = "SUMPRODUCT(" & baseExpr & compExcl & activeFilter & ")"
            Else
                actExpr = allExpr  ' fallback if no date columns
            End If

            gFormula = "=IF(" & ddRef & "=""All""," & allExpr & "," & actExpr & ")"
            SafeFormulaWrite m_ws, tblRow + 1 + ti2, gCol, gFormula
            m_ws.Cells(tblRow + 1 + ti2, gCol).Font.Size = 9
            m_ws.Cells(tblRow + 1 + ti2, gCol).Font.Color = RGB(30, 41, 59)
            m_ws.Cells(tblRow + 1 + ti2, gCol).HorizontalAlignment = xlCenter
        Next ti2
    Next gi

    Dim lastGrpCol As Long
    lastGrpCol = COL_START + Application.WorksheetFunction.Min(CLng(groups.Count), COL_END - COL_START)
    If lastGrpCol > COL_END Then lastGrpCol = COL_END

    ' Create horizontal stacked bar chart (ScreenUpdating must be True for chart)
    Application.ScreenUpdating = True
    On Error GoTo GrpChartErr

    Dim chartRow As Long
    chartRow = tblRow + 5
    Dim chartWidth As Double
    chartWidth = m_ws.Cells(1, COL_END + 1).Left - m_ws.Cells(1, COL_START).Left
    Dim chartHeight As Double
    chartHeight = CHART_HEIGHT_ROWS * 18

    Dim chtGrp As ChartObject
    Set chtGrp = m_ws.ChartObjects.Add( _
        Left:=m_ws.Cells(chartRow, COL_START).Left, _
        Top:=m_ws.Cells(chartRow, COL_START).Top, _
        Width:=chartWidth, Height:=chartHeight)

    With chtGrp.Chart
        .ChartType = xlBarStacked  ' Horizontal stacked bars

        ' Category axis = group names
        Dim catRange As Range
        Set catRange = m_ws.Range(m_ws.Cells(tblRow, COL_START + 1), m_ws.Cells(tblRow, lastGrpCol))

        Dim s1 As Series
        Set s1 = .SeriesCollection.NewSeries
        s1.Name = "New"
        s1.Values = m_ws.Range(m_ws.Cells(tblRow + 1, COL_START + 1), m_ws.Cells(tblRow + 1, lastGrpCol))
        s1.XValues = catRange
        s1.Format.Fill.ForeColor.RGB = THEME_SUCCESS       ' Emerald

        Dim s2 As Series
        Set s2 = .SeriesCollection.NewSeries
        s2.Name = "Reused"
        s2.Values = m_ws.Range(m_ws.Cells(tblRow + 2, COL_START + 1), m_ws.Cells(tblRow + 2, lastGrpCol))
        s2.XValues = catRange
        s2.Format.Fill.ForeColor.RGB = THEME_ACCENT2        ' Teal

        Dim s3 As Series
        Set s3 = .SeriesCollection.NewSeries
        s3.Name = "Demo"
        s3.Values = m_ws.Range(m_ws.Cells(tblRow + 3, COL_START + 1), m_ws.Cells(tblRow + 3, lastGrpCol))
        s3.XValues = catRange
        s3.Format.Fill.ForeColor.RGB = THEME_DANGER         ' Coral Red

        .HasTitle = True
        .ChartTitle.Text = "Systems by Group"
        .ChartTitle.Font.Color = RGB(226, 232, 240)
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Name = THEME_FONT

        .PlotArea.Format.Fill.ForeColor.RGB = RGB(15, 35, 62)
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(12, 27, 51)

        On Error Resume Next
        .ChartGroups(1).GapWidth = 80  ' Thinner bars
        ' Reverse category axis so highest count is at top (bar chart default is bottom-up)
        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlCategory).TickLabels.Font.Color = RGB(180, 190, 200)
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabelSpacing = 1
        .Axes(xlValue).TickLabels.Font.Color = RGB(180, 190, 200)
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(50, 75, 110)
        .Axes(xlValue).Format.Line.ForeColor.RGB = RGB(50, 75, 110)
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(40, 65, 100)
        On Error GoTo GrpChartErr

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Color = RGB(226, 232, 240)
        .Legend.Font.Size = 9
    End With

    m_nextRow = chartRow + CHART_HEIGHT_ROWS + 2
    Application.ScreenUpdating = False
    Exit Sub

GrpChartErr:
    Application.ScreenUpdating = False
    DebugLog "DashboardBuilder: Group Breakdown chart error: " & Err.Description
    m_nextRow = startRow + 4
End Sub

'====================================================================
' SECTION 8: ESCALATION TRACKER
' Lists escalated and watched systems from Working Sheet.
'====================================================================

Private Sub BuildEscalationTracker()
    Dim startRow As Long
    startRow = m_nextRow + 1

    WriteSectionTitle startRow, "Escalation Tracker", "Systems flagged as Escalated or Watched"

    If m_escCol = 0 Then
        m_ws.Cells(startRow + 2, COL_START).Value = "Escalated column not found."
        m_ws.Cells(startRow + 2, COL_START).Font.Size = 9
        m_ws.Cells(startRow + 2, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 4
        Exit Sub
    End If

    ' Collect escalated/watched systems (use array read for performance)
    Dim escData As Collection
    Set escData = New Collection
    Dim nEscRows As Long
    nEscRows = m_lastDataRow - m_firstDataRow + 1
    If nEscRows < 1 Then
        m_ws.Cells(startRow + 2, COL_START).Value = "No data rows."
        m_nextRow = startRow + 4
        Exit Sub
    End If

    ' Read relevant columns into arrays (single Range.Value read per column)
    Dim escArr() As Variant
    escArr = m_workSheet.Range(m_workSheet.Cells(m_firstDataRow, m_escCol), _
        m_workSheet.Cells(m_lastDataRow, m_escCol)).Value
    Dim ecArr() As Variant, grpArrE() As Variant, nrArrE() As Variant
    If m_entityCodeCol > 0 Then _
        ecArr = m_workSheet.Range(m_workSheet.Cells(m_firstDataRow, m_entityCodeCol), _
            m_workSheet.Cells(m_lastDataRow, m_entityCodeCol)).Value
    If m_groupCol > 0 Then _
        grpArrE = m_workSheet.Range(m_workSheet.Cells(m_firstDataRow, m_groupCol), _
            m_workSheet.Cells(m_lastDataRow, m_groupCol)).Value
    If m_nrCol > 0 Then _
        nrArrE = m_workSheet.Range(m_workSheet.Cells(m_firstDataRow, m_nrCol), _
            m_workSheet.Cells(m_lastDataRow, m_nrCol)).Value

    Dim ri2 As Long, escVal As String
    Dim eCode As String, eGrp As String, eNR As String
    Dim statusArrE() As Variant
    If m_statusCol > 0 Then _
        statusArrE = m_workSheet.Range(m_workSheet.Cells(m_firstDataRow, m_statusCol), _
            m_workSheet.Cells(m_lastDataRow, m_statusCol)).Value

    For ri2 = 1 To nEscRows
        ' Skip cancelled, completed, and non-IQ systems
        If m_statusCol > 0 Then
            Dim escStat As String
            escStat = LCase(Trim(CStr(statusArrE(ri2, 1) & "")))
            If escStat = "cancelled" Or escStat = "completed" Or escStat = "non iq" Then GoTo NextEscRow
        End If
        escVal = Trim(CStr(escArr(ri2, 1)))
        If LCase(escVal) = "escalated" Or LCase(escVal) = "watched" Then
            eCode = "": eGrp = "": eNR = ""
            If m_entityCodeCol > 0 Then eCode = CStr(ecArr(ri2, 1))
            If m_groupCol > 0 Then eGrp = CStr(grpArrE(ri2, 1))
            If m_nrCol > 0 Then eNR = CStr(nrArrE(ri2, 1))
            ' Array() creates a fresh Variant array each call (no shared-reference bug)
            escData.Add Array(eCode, eGrp, escVal, eNR)
        End If
NextEscRow:
    Next ri2

    If escData.Count = 0 Then
        m_ws.Cells(startRow + 2, COL_START).Value = "No escalated or watched systems."
        m_ws.Cells(startRow + 2, COL_START).Font.Size = 9
        m_ws.Cells(startRow + 2, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 4
        Exit Sub
    End If

    ' Sort: Escalated items first, then Watched (higher severity on top)
    Dim sortedEsc As New Collection
    Dim ePassItem As Variant
    For Each ePassItem In escData
        If LCase(CStr(ePassItem(2))) = "escalated" Then sortedEsc.Add ePassItem
    Next ePassItem
    For Each ePassItem In escData
        If LCase(CStr(ePassItem(2))) = "watched" Then sortedEsc.Add ePassItem
    Next ePassItem
    Set escData = sortedEsc

    ' Table header
    Dim hdrRow As Long
    hdrRow = startRow + 2
    Dim eHeaders As Variant
    eHeaders = Array("Entity Code", "", "Group", "Status", "Type")
    Dim eh As Long
    For eh = 0 To UBound(eHeaders)
        With m_ws.Cells(hdrRow, COL_START + eh)
            .Value = eHeaders(eh)
            .Font.Bold = True
            .Font.Size = 10
            .Font.Color = TABLE_HEADER_TEXT
            .Interior.Color = TABLE_HEADER_BG
            .HorizontalAlignment = xlCenter
        End With
    Next eh
    m_ws.Range(m_ws.Cells(hdrRow, COL_START), m_ws.Cells(hdrRow, COL_START + 1)).Merge
    m_ws.Rows(hdrRow).RowHeight = 24

    ' Data rows
    Dim dataRow As Long
    dataRow = hdrRow + 1
    Dim eItem As Variant
    Dim escIdx As Long: escIdx = 0
    For Each eItem In escData
        Dim escInfo As Variant
        escInfo = eItem

        m_ws.Range(m_ws.Cells(dataRow, COL_START), m_ws.Cells(dataRow, COL_START + 1)).Merge
        m_ws.Cells(dataRow, COL_START).Value = escInfo(0)
        m_ws.Cells(dataRow, COL_START + 2).Value = escInfo(1)
        m_ws.Cells(dataRow, COL_START + 3).Value = escInfo(2)
        m_ws.Cells(dataRow, COL_START + 4).Value = escInfo(3)

        Dim eci As Long
        For eci = 0 To 4
            m_ws.Cells(dataRow, COL_START + eci).Font.Size = 9
            m_ws.Cells(dataRow, COL_START + eci).Font.Color = TABLE_TEXT
            m_ws.Cells(dataRow, COL_START + eci).HorizontalAlignment = xlCenter
        Next eci

        ' Status color coding
        If LCase(CStr(escInfo(2))) = "escalated" Then
            m_ws.Cells(dataRow, COL_START + 3).Font.Color = RGB(185, 28, 28)
            m_ws.Cells(dataRow, COL_START + 3).Font.Bold = True
        ElseIf LCase(CStr(escInfo(2))) = "watched" Then
            m_ws.Cells(dataRow, COL_START + 3).Font.Color = RGB(180, 83, 9)
            m_ws.Cells(dataRow, COL_START + 3).Font.Bold = True
        End If

        ' Zebra striping
        Dim eRowRng As Range
        Set eRowRng = m_ws.Range(m_ws.Cells(dataRow, COL_START), m_ws.Cells(dataRow, COL_START + 4))
        If escIdx Mod 2 = 0 Then
            eRowRng.Interior.Color = TABLE_ROW_BG
        Else
            eRowRng.Interior.Color = TABLE_ALT_ROW_BG
        End If
        eRowRng.Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
        eRowRng.Borders(xlEdgeBottom).Weight = xlHairline

        dataRow = dataRow + 1
        escIdx = escIdx + 1
    Next eItem

    m_nextRow = dataRow + 2
End Sub

'====================================================================
' APPLY DASHBOARD FORMATTING
'====================================================================

Private Sub ApplyDashboardFormatting()
    On Error Resume Next

    ' Hide gridlines
    m_ws.Activate
    ActiveWindow.DisplayGridlines = False

    ' Freeze panes at row 3 (below title)
    m_ws.Cells(ROW_TITLE + 2, 1).Select
    ActiveWindow.FreezePanes = True

    ' Set print area
    m_ws.PageSetup.PrintArea = m_ws.Range(m_ws.Cells(1, 1), _
        m_ws.Cells(m_nextRow, COL_END)).Address

    ' Zoom
    ActiveWindow.Zoom = 90

    On Error GoTo 0
End Sub

'====================================================================
' ADD NAVIGATION LINKS
'====================================================================

Private Sub AddNavigationLinks()
    On Error Resume Next

    ' Add hyperlink back to Working Sheet in the subtitle area
    Dim linkRow As Long
    linkRow = ROW_TITLE + 1

    If SheetExists(ThisWorkbook, m_wsName) Then
        m_ws.Hyperlinks.Add Anchor:=m_ws.Cells(linkRow, COL_END), _
            Address:="", SubAddress:="'" & m_wsName & "'!A1", _
            TextToDisplay:=">> Working Sheet"
        m_ws.Cells(linkRow, COL_END).Font.Color = RGB(86, 156, 190)
        m_ws.Cells(linkRow, COL_END).Font.Size = 9
        m_ws.Cells(linkRow, COL_END).Font.Underline = xlUnderlineStyleSingle
    End If

    On Error GoTo 0
End Sub

'====================================================================
' PROTECT DASHBOARD SHEET
' Allows slicer interaction and editable HC Available cells.
'====================================================================

Private Sub ProtectDashboardSheet()
    ' Sheet protection removed — no-op
End Sub

'====================================================================
' HELPER: WRITE SECTION TITLE
' Writes a bold section title and italic subtitle on light background
' using TABLE_TEXT (dark) for visibility.
'====================================================================

Private Sub WriteSectionTitle(row As Long, title As String, subtitle As String)
    ' Section title with left accent bar on white background
    With m_ws.Range(m_ws.Cells(row, COL_START), m_ws.Cells(row, COL_END))
        .Merge
        .Value = title
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(12, 27, 51)
        .Font.Name = THEME_FONT
        .Interior.Color = RGB(248, 250, 252)    ' Near-white (subtle lift)
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .Borders(xlEdgeLeft).Color = RGB(86, 156, 190)   ' Brand blue accent bar
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeBottom).LineStyle = xlNone
    End With
    m_ws.Rows(row).RowHeight = 28

    If subtitle <> "" Then
        With m_ws.Range(m_ws.Cells(row + 1, COL_START), m_ws.Cells(row + 1, COL_END))
            .Merge
            .Value = subtitle
            .Font.Size = 9
            .Font.Italic = True
            .Font.Color = RGB(100, 116, 139)   ' Slate subtitle
            .Font.Name = THEME_FONT
            .Interior.Color = RGB(248, 250, 252)
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
        End With
        m_ws.Rows(row + 1).RowHeight = 18
    End If
End Sub

'====================================================================
' HELPER: COLLECT GROUPS
' Returns a Collection of unique group names from Working Sheet.
'====================================================================

Private Function CollectGroups() As Collection
    ' Return cached result if available
    If Not m_cachedGroups Is Nothing Then
        Set CollectGroups = m_cachedGroups
        Exit Function
    End If

    Dim result As New Collection
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' Always initialise counts dictionary (even on early exit)
    If m_groupSysCounts Is Nothing Then
        Set m_groupSysCounts = CreateObject("Scripting.Dictionary")
    End If

    If m_groupCol = 0 Then
        Set CollectGroups = result
        Set m_cachedGroups = result
        Exit Function
    End If

    ' Read group column into array (single Range.Value read)
    Dim grpArr As Variant
    grpArr = m_workSheet.Range(m_workSheet.Cells(m_firstDataRow, m_groupCol), _
        m_workSheet.Cells(m_lastDataRow, m_groupCol)).Value

    ' Handle single-cell read (returns scalar instead of 2-D array)
    If Not IsArray(grpArr) Then
        Dim tmp As Variant: tmp = grpArr
        ReDim grpArr(1 To 1, 1 To 1)
        grpArr(1, 1) = tmp
    End If

    ' Also read NR column for system count sorting
    Dim nrArr As Variant
    Dim hasNR As Boolean: hasNR = (m_nrCol > 0)
    If hasNR Then
        nrArr = m_workSheet.Range(m_workSheet.Cells(m_firstDataRow, m_nrCol), _
            m_workSheet.Cells(m_lastDataRow, m_nrCol)).Value
        If Not IsArray(nrArr) Then
            Dim tmp2 As Variant: tmp2 = nrArr
            ReDim nrArr(1 To 1, 1 To 1)
            nrArr(1, 1) = tmp2
        End If
    End If

    ' Build unique groups and count New + Reused per group
    Dim countDict As Object
    Set countDict = CreateObject("Scripting.Dictionary")

    Dim ri3 As Long, gv As String, nrv As String
    For ri3 = 1 To UBound(grpArr, 1)
        If IsError(grpArr(ri3, 1)) Then GoTo NextGrpRow
        gv = Trim(CStr(grpArr(ri3, 1)))
        If gv <> "" Then
            If Not dict.exists(gv) Then
                dict.Add gv, True
                countDict.Add gv, CLng(0)
            End If
            If hasNR Then
                If Not IsError(nrArr(ri3, 1)) Then
                    nrv = Trim(CStr(nrArr(ri3, 1)))
                    If nrv = "New" Or nrv = "Reused" Then
                        countDict.Item(gv) = CLng(countDict.Item(gv)) + 1
                    End If
                End If
            End If
        End If
NextGrpRow:
    Next ri3

    ' Copy to typed arrays for sorting (avoids Variant-array edge cases)
    Dim gCount As Long: gCount = countDict.Count
    If gCount > 0 Then
        Dim grpNames() As String, grpCnts() As Long
        ReDim grpNames(1 To gCount): ReDim grpCnts(1 To gCount)
        Dim dk As Variant, ix As Long: ix = 0
        For Each dk In countDict.keys
            ix = ix + 1
            grpNames(ix) = CStr(dk)
            grpCnts(ix) = CLng(countDict.Item(dk))
        Next dk

        ' Bubble-sort descending by count
        Dim si As Long, sj As Long
        Dim swpN As String, swpC As Long
        For si = 1 To gCount - 1
            For sj = si + 1 To gCount
                If grpCnts(sj) > grpCnts(si) Then
                    swpN = grpNames(si): grpNames(si) = grpNames(sj): grpNames(sj) = swpN
                    swpC = grpCnts(si): grpCnts(si) = grpCnts(sj): grpCnts(sj) = swpC
                End If
            Next sj
        Next si
        For si = 1 To gCount
            result.Add grpNames(si)
        Next si
    End If

    ' Cache counts for reuse (HC Gap sorting)
    Set m_groupSysCounts = countDict

    Set CollectGroups = result
    Set m_cachedGroups = result
End Function


'====================================================================
' HELPER: SORT A GROUP COLLECTION BY SYSTEM COUNT
' Re-orders any Collection of group names using the cached
' m_groupSysCounts dictionary. Groups not in the dictionary are
' placed at the end (count = 0). Returns a new Collection.
'====================================================================

Private Function SortGroupsBySystemCount(src As Collection) As Collection
    Dim sorted As New Collection
    If src.Count = 0 Then
        Set SortGroupsBySystemCount = sorted
        Exit Function
    End If

    ' Ensure counts are available
    If m_groupSysCounts Is Nothing Then
        Dim dummy As Collection
        Set dummy = CollectGroups()    ' populates m_groupSysCounts
    End If

    ' Safety: if still Nothing (e.g. m_groupCol=0), return unsorted
    If m_groupSysCounts Is Nothing Then
        Set SortGroupsBySystemCount = src
        Exit Function
    End If

    ' Copy to typed arrays for sorting
    Dim n As Long: n = src.Count
    Dim names() As String, cnts() As Long
    ReDim names(1 To n): ReDim cnts(1 To n)
    Dim idx As Long
    For idx = 1 To n
        names(idx) = CStr(src(idx))
        If m_groupSysCounts.exists(names(idx)) Then
            cnts(idx) = CLng(m_groupSysCounts.Item(names(idx)))
        Else
            cnts(idx) = 0
        End If
    Next idx

    ' Swap sort descending
    Dim a As Long, b As Long
    Dim tN As String, tC As Long
    For a = 1 To n - 1
        For b = a + 1 To n
            If cnts(b) > cnts(a) Then
                tN = names(a): names(a) = names(b): names(b) = tN
                tC = cnts(a): cnts(a) = cnts(b): cnts(b) = tC
            End If
        Next b
    Next a

    For idx = 1 To n
        sorted.Add names(idx)
    Next idx
    Set SortGroupsBySystemCount = sorted
End Function


'====================================================================
' HELPER: FIND HC TABLE ROW
' Searches Working Sheet below the main data for a section title
' containing the given text. Uses Excel Find for reliability
' (HC tables may start at high column numbers near the Gantt area).
' Returns 0 if not found.
'====================================================================

Private Function FindHCTableRow(searchText As String) As Long
    FindHCTableRow = 0

    ' Determine the last used row across all columns
    Dim ur As Range
    Set ur = m_workSheet.UsedRange
    Dim maxRow As Long
    maxRow = ur.row + ur.Rows.Count - 1
    If maxRow <= m_lastDataRow Then Exit Function

    ' Define search area: everything below the main data table
    Dim maxCol As Long
    maxCol = ur.Column + ur.Columns.Count - 1
    If maxCol < 1 Then maxCol = 1

    Dim searchArea As Range
    Set searchArea = m_workSheet.Range( _
        m_workSheet.Cells(m_lastDataRow + 1, 1), _
        m_workSheet.Cells(maxRow, maxCol))

    ' Use Excel Find (searches all columns regardless of position)
    On Error Resume Next
    Dim found As Range
    Set found = searchArea.Find(What:=searchText, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False)
    On Error GoTo 0

    If Not found Is Nothing Then
        FindHCTableRow = found.row
    End If
End Function

'====================================================================
' HELPER: FIND HC GROUP COLUMN
' Returns the column index containing "Group" in the HC table header row.
'====================================================================

Private Function FindHCGroupCol(headerRow As Long) As Long
    FindHCGroupCol = 0
    Dim lastCol As Long
    lastCol = m_workSheet.Cells(headerRow, m_workSheet.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1

    Dim jc As Long
    For jc = 1 To lastCol
        If LCase(Trim(CStr(m_workSheet.Cells(headerRow, jc).Value))) = "group" Then
            FindHCGroupCol = jc
            Exit Function
        End If
    Next jc
    ' If not found by name, assume it's the leftmost non-empty column
    For jc = 1 To lastCol
        If Trim(CStr(m_workSheet.Cells(headerRow, jc).Value)) <> "" Then
            FindHCGroupCol = jc
            Exit Function
        End If
    Next jc
End Function

'====================================================================
' INJECT WORKSHEET_CHANGE HANDLER
' Programmatically adds a Worksheet_Change event to the Dashboard
' sheet module so the chart start date cell is reactive.
' Requires: Trust access to the VBA project object model.
'====================================================================

Private Sub InjectDashboardChangeHandler()
    On Error Resume Next

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    If vbProj Is Nothing Then Exit Sub

    ' Find the Dashboard sheet's code module
    Dim wsCodeMod As Object
    Set wsCodeMod = vbProj.VBComponents(m_ws.CodeName).CodeModule
    If wsCodeMod Is Nothing Then Exit Sub

    ' Clear any existing code in the sheet module
    If wsCodeMod.CountOfLines > 0 Then
        wsCodeMod.DeleteLines 1, wsCodeMod.CountOfLines
    End If

    ' Build the event handler code
    ' Handles both chart start date changes (re-filters PivotItems)
    ' and KPI group dropdown changes (forces recalculation).
    ' Uses Application.Range for reliable workbook-scoped name lookup.
    ' EnableEvents guard prevents recursive triggering.
    Dim code As String
    code = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
           "    On Error Resume Next" & vbCrLf & _
           "    Dim dateRng As Range, grpRng As Range" & vbCrLf & _
           "    Set dateRng = Application.Range(""DASH_CHART_START"")" & vbCrLf & _
           "    Set grpRng = Application.Range(""DASH_KPI_GROUP"")" & vbCrLf & _
           "    Dim needCalc As Boolean" & vbCrLf & _
           "    needCalc = False" & vbCrLf & _
           "    If Not dateRng Is Nothing Then" & vbCrLf & _
           "        If Not Intersect(Target, dateRng) Is Nothing Then" & vbCrLf & _
           "            Application.EnableEvents = False" & vbCrLf & _
           "            DashboardBuilder.RefreshChartStartDate" & vbCrLf & _
           "            Application.EnableEvents = True" & vbCrLf & _
           "            needCalc = True" & vbCrLf & _
           "        End If" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    If Not grpRng Is Nothing Then" & vbCrLf & _
           "        If Not Intersect(Target, grpRng) Is Nothing Then" & vbCrLf & _
           "            needCalc = True" & vbCrLf & _
           "        End If" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    If needCalc Then" & vbCrLf & _
           "        Application.Calculation = xlCalculationAutomatic" & vbCrLf & _
           "        Me.Calculate" & vbCrLf & _
           "    End If" & vbCrLf & _
           "End Sub"

    wsCodeMod.AddFromString code

    On Error GoTo 0
End Sub

'====================================================================
' REFRESH CHART START DATE
' Public entry point called by Dashboard Worksheet_Change event.
' Re-filters PivotItems on Monthly and Active Systems PivotTables
' based on the current value of the DASH_CHART_START named range.
'====================================================================

Public Sub RefreshChartStartDate()
    On Error Resume Next

    ' Read the start date from the named range
    Dim dateCell As Range
    Set dateCell = ThisWorkbook.Names("DASH_CHART_START").RefersToRange
    If dateCell Is Nothing Then Exit Sub
    If Not IsDate(dateCell.Value) Then Exit Sub

    Dim startDate As Date
    startDate = CDate(dateCell.Value)
    Dim startMonthStr As String
    startMonthStr = Format(startDate, "YYYY-MM")

    ' Get DashHelper sheet
    Dim hlpSheet As Worksheet
    Set hlpSheet = ThisWorkbook.Sheets("DashHelper")
    If hlpSheet Is Nothing Then Exit Sub

    Dim pt As PivotTable
    Dim pf As PivotField
    Dim pi As PivotItem

    ' Filter PT_Monthly (ProjectMonth field)
    Set pt = hlpSheet.PivotTables("PT_Monthly")
    If Not pt Is Nothing Then
        pt.ManualUpdate = True
        Set pf = pt.PivotFields("ProjectMonth")
        If Not pf Is Nothing Then
            ' Phase 1: Show ALL items first (prevents "cannot hide last item" error)
            For Each pi In pf.PivotItems
                pi.Visible = True
            Next pi
            ' Phase 2: Now hide items before start date + blanks
            For Each pi In pf.PivotItems
                If pi.Name = "" Or LCase(pi.Name) = "(blank)" Then
                    On Error Resume Next
                    pi.Visible = False
                    On Error GoTo 0
                ElseIf pi.Name < startMonthStr Then
                    On Error Resume Next
                    pi.Visible = False
                    On Error GoTo 0
                End If
            Next pi
        End If
        pt.ManualUpdate = False
    End If

    ' Filter PT_Active (PresenceMonth field)
    Set pt = hlpSheet.PivotTables("PT_Active")
    If Not pt Is Nothing Then
        pt.ManualUpdate = True
        Set pf = pt.PivotFields("PresenceMonth")
        If Not pf Is Nothing Then
            ' Phase 1: Show ALL items first
            For Each pi In pf.PivotItems
                pi.Visible = True
            Next pi
            ' Phase 2: Now hide items before start date + blanks
            For Each pi In pf.PivotItems
                If pi.Name = "" Or LCase(pi.Name) = "(blank)" Then
                    On Error Resume Next
                    pi.Visible = False
                    On Error GoTo 0
                ElseIf pi.Name < startMonthStr Then
                    On Error Resume Next
                    pi.Visible = False
                    On Error GoTo 0
                End If
            Next pi
        End If
        pt.ManualUpdate = False
    End If

    On Error GoTo 0
End Sub

'====================================================================
' SECTION 10: INSTALL BASE (TOTAL TOOLS)
' Stacked column chart showing cumulative tool count over quarters.
' Uses a SEPARATE PivotCache from the main dashboard charts.
' PivotTable PT_InstallBase on DashHelper: Row=InstallQtr, Col=NewReused,
' Value=Sum(InstallDelta), filtered to Monthly rows.
' Data table on Dashboard: per-quarter deltas, cumulative sums, total.
' IB_BASELINE named range for user-entered starting tool count.
'====================================================================

Public Sub BuildInstallBaseSection(ws As Worksheet, workSheet As Worksheet, _
        helperSheet As Worksheet, helperTable As ListObject)
    On Error GoTo IBErrorHandler

    Dim startRow As Long
    startRow = m_nextRow + 1

    WriteSectionTitle startRow, "Install Base (Total Tools)", _
        "Cumulative tool count by quarter — enter baseline in cell below"

    ' --- Validate prerequisites ---
    If helperTable Is Nothing Then
        ws.Cells(startRow + 2, COL_START).Value = "Helper table not available."
        ws.Cells(startRow + 2, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 4
        Exit Sub
    End If

    ' --- IB_BASELINE cell (user-editable) ---
    Dim baseRow As Long: baseRow = startRow + 2
    ws.Cells(baseRow, COL_START).Value = "Existing Tools (Baseline):"
    ws.Cells(baseRow, COL_START).Font.Bold = True
    ws.Cells(baseRow, COL_START).Font.Size = 10
    ws.Cells(baseRow, COL_START).Font.Color = RGB(100, 116, 139)

    Dim baseMerge As Range
    Set baseMerge = ws.Range(ws.Cells(baseRow, COL_START + 3), _
        ws.Cells(baseRow, COL_START + 4))
    baseMerge.Merge

    Dim baseCell As Range
    Set baseCell = ws.Cells(baseRow, COL_START + 3)
    baseCell.Value = 0
    baseCell.NumberFormat = "#,##0"
    baseCell.Font.Bold = True
    baseCell.Font.Size = 14
    baseCell.Font.Color = RGB(30, 41, 59)
    baseCell.Interior.Color = RGB(255, 255, 255)
    baseCell.HorizontalAlignment = xlCenter
    With baseMerge.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(86, 156, 190)
    End With
    With baseMerge.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225)
    End With
    With baseMerge.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225)
    End With
    With baseMerge.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(203, 213, 225)
    End With
    If Not m_dropdownCells Is Nothing Then m_dropdownCells.Add baseCell

    ' Create IB_BASELINE named range
    On Error Resume Next
    ThisWorkbook.Names("IB_BASELINE").Delete
    ThisWorkbook.Names.Add Name:="IB_BASELINE", RefersTo:=baseCell
    On Error GoTo IBErrorHandler

    ' --- Create SEPARATE PivotCache for Install Base ---
    On Error Resume Next
    Set m_ibPivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=helperTable.Range)
    On Error GoTo IBErrorHandler

    If m_ibPivotCache Is Nothing Then
        ws.Cells(startRow + 4, COL_START).Value = "Could not create Install Base PivotCache."
        ws.Cells(startRow + 4, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 6
        Exit Sub
    End If

    ' --- Create PivotTable PT_InstallBase on helper sheet ---
    Dim ptDest As Range
    Dim ptStartRow As Long
    ptStartRow = helperSheet.Cells(helperSheet.Rows.Count, 1).End(xlUp).row + 5
    Set ptDest = helperSheet.Cells(ptStartRow, 18)  ' Column R, away from other PTs

    On Error Resume Next
    Set m_ptInstallBase = m_ibPivotCache.CreatePivotTable( _
        TableDestination:=ptDest, _
        TableName:="PT_InstallBase")
    On Error GoTo IBErrorHandler

    If m_ptInstallBase Is Nothing Then
        ws.Cells(startRow + 4, COL_START).Value = "Could not create Install Base PivotTable."
        ws.Cells(startRow + 4, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 6
        Exit Sub
    End If

    ' Configure PT_InstallBase
    On Error Resume Next
    With m_ptInstallBase
        .ManualUpdate = True

        ' Row field: InstallQtr
        .PivotFields(HLP_COL_INSTALLQTR).Orientation = xlRowField
        .PivotFields(HLP_COL_INSTALLQTR).Position = 1

        ' Column field: NewReused
        .PivotFields(HLP_COL_NR).Orientation = xlColumnField
        .PivotFields(HLP_COL_NR).Position = 1

        ' Data field: Sum of InstallDelta
        .AddDataField .PivotFields(HLP_COL_INSTALLDELTA), "Sum of InstallDelta", xlSum

        ' Page field: RowType = "Monthly"
        .PivotFields(HLP_COL_ROWTYPE).Orientation = xlPageField
        .PivotFields(HLP_COL_ROWTYPE).CurrentPage = "Monthly"

        .ManualUpdate = False
    End With

    ' Hide (blank) items in InstallQtr (two-phase)
    Dim pi As PivotItem
    m_ptInstallBase.ManualUpdate = True
    For Each pi In m_ptInstallBase.PivotFields(HLP_COL_INSTALLQTR).PivotItems
        pi.Visible = True
    Next pi
    For Each pi In m_ptInstallBase.PivotFields(HLP_COL_INSTALLQTR).PivotItems
        If pi.Name = "" Or LCase(pi.Name) = "(blank)" Then
            On Error Resume Next
            pi.Visible = False
            On Error GoTo 0
        End If
    Next pi
    m_ptInstallBase.ManualUpdate = False

    ' Hide (blank) items in NewReused (two-phase)
    m_ptInstallBase.ManualUpdate = True
    For Each pi In m_ptInstallBase.PivotFields(HLP_COL_NR).PivotItems
        pi.Visible = True
    Next pi
    For Each pi In m_ptInstallBase.PivotFields(HLP_COL_NR).PivotItems
        If pi.Name = "" Or LCase(pi.Name) = "(blank)" Then
            On Error Resume Next
            pi.Visible = False
            On Error GoTo 0
        End If
    Next pi
    m_ptInstallBase.ManualUpdate = False
    On Error GoTo IBErrorHandler

    ' --- Create IB_PT_ANCHOR named range for GETPIVOTDATA ---
    ' Anchor is the top-left cell of the PivotTable data body
    Dim ptAnchorCell As Range
    Set ptAnchorCell = m_ptInstallBase.TableRange2.Cells(1, 1)
    On Error Resume Next
    ThisWorkbook.Names("IB_PT_ANCHOR").Delete
    ThisWorkbook.Names.Add Name:="IB_PT_ANCHOR", RefersTo:=ptAnchorCell
    On Error GoTo IBErrorHandler

    ' --- Collect quarters from PivotTable ---
    Dim quarters As Collection
    Set quarters = New Collection
    On Error Resume Next
    Dim pf As PivotField
    Set pf = m_ptInstallBase.PivotFields(HLP_COL_INSTALLQTR)
    For Each pi In pf.PivotItems
        If pi.Visible And pi.Name <> "" And LCase(pi.Name) <> "(blank)" Then
            quarters.Add pi.Name
        End If
    Next pi
    On Error GoTo IBErrorHandler

    If quarters.Count = 0 Then
        ws.Cells(startRow + 4, COL_START).Value = "No quarter data available for Install Base."
        ws.Cells(startRow + 4, COL_START).Font.Color = RGB(100, 116, 139)
        m_nextRow = startRow + 6
        Exit Sub
    End If

    ' --- Build 9-row data table on Dashboard ---
    Dim tblRow As Long: tblRow = baseRow + 2
    Dim nQtrs As Long: nQtrs = quarters.Count
    If nQtrs > COL_END - COL_START - 1 Then nQtrs = COL_END - COL_START - 1

    ' Row labels (column B)
    Dim ibLabels As Variant
    ibLabels = Array("New", "Reused", "Demo", "Cum. New", "Cum. Reused", _
        "Cum. Demo", "Existing", "Total", "Quarter")
    Dim lbi As Long
    For lbi = 0 To 8
        ws.Cells(tblRow + lbi, COL_START).Value = ibLabels(lbi)
        ws.Cells(tblRow + lbi, COL_START).Font.Bold = True
        ws.Cells(tblRow + lbi, COL_START).Font.Size = 9
        ws.Cells(tblRow + lbi, COL_START).Font.Color = RGB(100, 116, 139)
    Next lbi

    ' Quarter columns
    Dim baseRef As String
    baseRef = "IB_BASELINE"
    Dim anchorRef As String
    anchorRef = "IB_PT_ANCHOR"

    Dim qi As Long, qCol As Long, qName As String
    For qi = 1 To nQtrs
        qCol = COL_START + qi
        qName = quarters(qi)

        ' Row 9 (index 8): Quarter label
        ws.Cells(tblRow + 8, qCol).Value = qName
        ws.Cells(tblRow + 8, qCol).Font.Bold = True
        ws.Cells(tblRow + 8, qCol).Font.Size = 8
        ws.Cells(tblRow + 8, qCol).Font.Color = RGB(100, 116, 139)
        ws.Cells(tblRow + 8, qCol).HorizontalAlignment = xlCenter

        ' Row 1 (index 0): New per quarter — GETPIVOTDATA
        SafeFormulaWrite ws, tblRow, qCol, _
            "=IFERROR(GETPIVOTDATA(""Sum of InstallDelta""," & anchorRef & ",""" & _
            HLP_COL_INSTALLQTR & """,""" & qName & """,""" & HLP_COL_NR & """,""New""),0)"

        ' Row 2 (index 1): Reused per quarter
        SafeFormulaWrite ws, tblRow + 1, qCol, _
            "=IFERROR(GETPIVOTDATA(""Sum of InstallDelta""," & anchorRef & ",""" & _
            HLP_COL_INSTALLQTR & """,""" & qName & """,""" & HLP_COL_NR & """,""Reused""),0)"

        ' Row 3 (index 2): Demo per quarter (absolute value for display)
        SafeFormulaWrite ws, tblRow + 2, qCol, _
            "=IFERROR(ABS(GETPIVOTDATA(""Sum of InstallDelta""," & anchorRef & ",""" & _
            HLP_COL_INSTALLQTR & """,""" & qName & """,""" & HLP_COL_NR & """,""Demo"")),0)"

        ' Row 4 (index 3): Cumulative New (running sum)
        If qi = 1 Then
            ws.Cells(tblRow + 3, qCol).formula = "=" & ColLetter(qCol) & tblRow
        Else
            ws.Cells(tblRow + 3, qCol).formula = "=" & ColLetter(qCol - 1) & (tblRow + 3) & _
                "+" & ColLetter(qCol) & tblRow
        End If

        ' Row 5 (index 4): Cumulative Reused (running sum)
        If qi = 1 Then
            ws.Cells(tblRow + 4, qCol).formula = "=" & ColLetter(qCol) & (tblRow + 1)
        Else
            ws.Cells(tblRow + 4, qCol).formula = "=" & ColLetter(qCol - 1) & (tblRow + 4) & _
                "+" & ColLetter(qCol) & (tblRow + 1)
        End If

        ' Row 6 (index 5): Cumulative Demo (running sum)
        If qi = 1 Then
            ws.Cells(tblRow + 5, qCol).formula = "=" & ColLetter(qCol) & (tblRow + 2)
        Else
            ws.Cells(tblRow + 5, qCol).formula = "=" & ColLetter(qCol - 1) & (tblRow + 5) & _
                "+" & ColLetter(qCol) & (tblRow + 2)
        End If

        ' Row 7 (index 6): Existing = IB_BASELINE - Cumulative Demo
        ws.Cells(tblRow + 6, qCol).formula = "=" & baseRef & "-" & ColLetter(qCol) & (tblRow + 5)

        ' Row 8 (index 7): Total = Existing + Cum.New + Cum.Reused
        ws.Cells(tblRow + 7, qCol).formula = "=" & ColLetter(qCol) & (tblRow + 6) & _
            "+" & ColLetter(qCol) & (tblRow + 3) & "+" & ColLetter(qCol) & (tblRow + 4)

        ' Formatting for data cells
        Dim dr As Long
        For dr = 0 To 7
            ws.Cells(tblRow + dr, qCol).Font.Size = 9
            ws.Cells(tblRow + dr, qCol).Font.Color = RGB(30, 41, 59)
            ws.Cells(tblRow + dr, qCol).HorizontalAlignment = xlCenter
            ws.Cells(tblRow + dr, qCol).NumberFormat = "#,##0"
        Next dr
    Next qi

    Dim lastQCol As Long
    lastQCol = COL_START + nQtrs

    ' --- Create stacked column chart ---
    Application.ScreenUpdating = True
    On Error GoTo IBChartErr

    Dim chartRow As Long: chartRow = tblRow + 10
    Dim chartWidth As Double
    chartWidth = ws.Cells(1, COL_END + 1).Left - ws.Cells(1, COL_START).Left
    Dim chartHeight As Double
    chartHeight = CHART_HEIGHT_ROWS * 18

    Dim chtIB As ChartObject
    Set chtIB = ws.ChartObjects.Add( _
        Left:=ws.Cells(chartRow, COL_START).Left, _
        Top:=ws.Cells(chartRow, COL_START).Top, _
        Width:=chartWidth, Height:=chartHeight)

    With chtIB.Chart
        .ChartType = xlColumnStacked

        ' Category axis = quarter labels (row 9)
        Dim catRange As Range
        Set catRange = ws.Range(ws.Cells(tblRow + 8, COL_START + 1), _
            ws.Cells(tblRow + 8, lastQCol))

        ' Series 1: Existing (bottom of stack)
        Dim sExist As Series
        Set sExist = .SeriesCollection.NewSeries
        sExist.Name = "Existing"
        sExist.Values = ws.Range(ws.Cells(tblRow + 6, COL_START + 1), _
            ws.Cells(tblRow + 6, lastQCol))
        sExist.XValues = catRange
        sExist.Format.Fill.ForeColor.RGB = THEME_ACCENT     ' AMAT Blue

        ' Series 2: New (cumulative)
        Dim sNew As Series
        Set sNew = .SeriesCollection.NewSeries
        sNew.Name = "New"
        sNew.Values = ws.Range(ws.Cells(tblRow + 3, COL_START + 1), _
            ws.Cells(tblRow + 3, lastQCol))
        sNew.XValues = catRange
        sNew.Format.Fill.ForeColor.RGB = THEME_SUCCESS       ' Emerald

        ' Series 3: Reused (cumulative)
        Dim sReused As Series
        Set sReused = .SeriesCollection.NewSeries
        sReused.Name = "Reused"
        sReused.Values = ws.Range(ws.Cells(tblRow + 4, COL_START + 1), _
            ws.Cells(tblRow + 4, lastQCol))
        sReused.XValues = catRange
        sReused.Format.Fill.ForeColor.RGB = THEME_ACCENT2    ' Teal

        ' Chart styling (dark theme matching other dashboard charts)
        .HasTitle = True
        .ChartTitle.Text = "Install Base (Total Tools)"
        .ChartTitle.Font.Color = RGB(226, 232, 240)
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Name = THEME_FONT

        .PlotArea.Format.Fill.ForeColor.RGB = RGB(15, 35, 62)
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(12, 27, 51)

        On Error Resume Next
        .ChartGroups(1).GapWidth = 80
        .Axes(xlCategory).TickLabels.Font.Color = RGB(180, 190, 200)
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabelSpacing = 1
        .Axes(xlValue).TickLabels.Font.Color = RGB(180, 190, 200)
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(50, 75, 110)
        .Axes(xlValue).Format.Line.ForeColor.RGB = RGB(50, 75, 110)
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(40, 65, 100)
        On Error GoTo IBChartErr

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Color = RGB(226, 232, 240)
        .Legend.Font.Size = 9

        ' Data labels from Total row (row 8 = tblRow + 7)
        ' Apply to the topmost visible series (Reused is last added = top of stack)
        On Error Resume Next
        Dim topSrs As Series
        Set topSrs = .SeriesCollection(.SeriesCollection.Count)
        If Not topSrs Is Nothing Then
            topSrs.HasDataLabels = True
            topSrs.DataLabels.ShowValue = False
            topSrs.DataLabels.ShowCategoryName = False
            ' Link each data label to the Total row cell
            Dim dli As Long
            For dli = 1 To nQtrs
                topSrs.Points(dli).DataLabel.Text = _
                    "=" & "'" & ws.Name & "'!" & _
                    ColLetter(COL_START + dli) & CStr(tblRow + 7)
                topSrs.Points(dli).DataLabel.Font.Color = RGB(255, 255, 255)
                topSrs.Points(dli).DataLabel.Font.Size = 9
                topSrs.Points(dli).DataLabel.Font.Bold = True
                topSrs.Points(dli).DataLabel.NumberFormat = "#,##0"
                topSrs.Points(dli).DataLabel.Position = xlLabelPositionOutsideEnd
            Next dli
        End If
        On Error GoTo IBChartErr
    End With

    m_nextRow = chartRow + CHART_HEIGHT_ROWS + 2

    ' --- Create 5 independent slicers on the IB PivotCache ---
    BuildInstallBaseSlicers ws

    Application.ScreenUpdating = False
    Exit Sub

IBChartErr:
    Application.ScreenUpdating = False
    DebugLog "DashboardBuilder: Install Base chart error: " & Err.Description
    m_nextRow = startRow + 4
    Exit Sub

IBErrorHandler:
    DebugLog "DashboardBuilder: Install Base section error: " & Err.Description
    m_nextRow = startRow + 4
End Sub

'====================================================================
' INSTALL BASE SLICERS
' Creates 5 independent slicers on the Install Base PivotCache.
' Uses Err.Clear before each SlicerCaches.Add2 to prevent cascading.
'====================================================================

Private Sub BuildInstallBaseSlicers(ws As Worksheet)
    If m_ptInstallBase Is Nothing Then Exit Sub

    Dim slicerRow As Long
    slicerRow = m_nextRow + 1

    WriteSectionTitle slicerRow, "Install Base Filters", _
        "Use slicers to filter the Install Base chart"

    Dim slicerPlaceRow As Long: slicerPlaceRow = slicerRow + 2
    Dim slicerLeft As Double
    Dim slicerWidth As Double, slicerHeight As Double
    slicerWidth = 140
    slicerHeight = 150
    slicerLeft = ws.Cells(slicerPlaceRow, COL_START).Left

    On Error Resume Next

    ' Slicer 1: NewReused
    Err.Clear
    Dim scIBNR As SlicerCache
    Set scIBNR = ThisWorkbook.SlicerCaches.Add2(m_ptInstallBase, HLP_COL_NR)
    If Not scIBNR Is Nothing Then
        scIBNR.Slicers.Add ws, , "IB_Slicer_NewReused", "New / Reused / Demo", _
            ws.Cells(slicerPlaceRow, COL_START).Top, slicerLeft, slicerWidth, slicerHeight
        slicerLeft = slicerLeft + slicerWidth + 8
    End If

    ' Slicer 2: Group
    Err.Clear
    Dim scIBGrp As SlicerCache
    Set scIBGrp = ThisWorkbook.SlicerCaches.Add2(m_ptInstallBase, HLP_COL_GROUP)
    If Not scIBGrp Is Nothing Then
        scIBGrp.Slicers.Add ws, , "IB_Slicer_Group", "Group", _
            ws.Cells(slicerPlaceRow, COL_START).Top, slicerLeft, slicerWidth, slicerHeight
        slicerLeft = slicerLeft + slicerWidth + 8
    End If

    ' Slicer 3: CEID
    Err.Clear
    Dim scIBCeid As SlicerCache
    Set scIBCeid = ThisWorkbook.SlicerCaches.Add2(m_ptInstallBase, HLP_COL_CEID)
    If Not scIBCeid Is Nothing Then
        scIBCeid.Slicers.Add ws, , "IB_Slicer_CEID", "CEID", _
            ws.Cells(slicerPlaceRow, COL_START).Top, slicerLeft, slicerWidth, slicerHeight
        slicerLeft = slicerLeft + slicerWidth + 8
    End If

    ' Slicer 4: EntityType
    Err.Clear
    Dim scIBET As SlicerCache
    Set scIBET = ThisWorkbook.SlicerCaches.Add2(m_ptInstallBase, HLP_COL_ENTITY_TYPE)
    If Not scIBET Is Nothing Then
        scIBET.Slicers.Add ws, , "IB_Slicer_EntityType", "Entity Type", _
            ws.Cells(slicerPlaceRow, COL_START).Top, slicerLeft, slicerWidth, slicerHeight
        slicerLeft = slicerLeft + slicerWidth + 8
    End If

    ' Slicer 5: HasSetStart
    Err.Clear
    Dim scIBSS As SlicerCache
    Set scIBSS = ThisWorkbook.SlicerCaches.Add2(m_ptInstallBase, HLP_COL_HASSETSTART)
    If Not scIBSS Is Nothing Then
        scIBSS.Slicers.Add ws, , "IB_Slicer_HasSetStart", "Has Set Start", _
            ws.Cells(slicerPlaceRow, COL_START).Top, slicerLeft, slicerWidth, slicerHeight
    End If

    On Error GoTo 0

    m_nextRow = slicerPlaceRow + 10
End Sub

'====================================================================
' HELPER: SAFE FORMULA WRITE
' Writes a formula to a cell with error handling.
'====================================================================

Private Sub SafeFormulaWrite(ws As Worksheet, row As Long, col As Long, formula As String)
    On Error Resume Next
    ws.Cells(row, col).formula = formula
    If Err.Number <> 0 Then
        Err.Clear
        ws.Cells(row, col).Formula2 = formula
        If Err.Number <> 0 Then
            Err.Clear
            ws.Cells(row, col).Value = 0
        End If
    End If
    On Error GoTo 0
End Sub
