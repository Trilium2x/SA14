Attribute VB_Name = "RampAlignment"
'====================================================================
' Ramp Alignment Module - Rev14
'
' Generates per-group ramp alignment reports for Intel customer
' alignment. Standalone module — no impact on other modules.
'
' Two-phase UX:
'   Phase 1: BuildRampAlignment() — creates control panel UI
'            (group dropdown + Generate button) on the active sheet.
'   Phase 2: RampAlignment_Generate() — reads selected group,
'            creates "Ramp - {GroupName}" report sheet.
'
' Report sections:
'   1. Title Bar
'   2. Counters (Total, New, Reused, Demo, Conversion)
'   3. Shipping Schedule (Entity Code, SDD, Ship Date)
'   4. Standard Durations (per CEID from Milestones sheet)
'   5. Current Schedule (correct date column per milestone type)
'   6. Actual Duration & Gaps
'   7. Intel Requirements per Milestone (template)
'   8. Sign-off Criteria per CEID (template)
'   9. Conversion Scope (template)
'
' Rev11 changes from Rev10:
'   - WriteStandardDurations: filters to Definitions milestones only
'   - WriteCurrentSchedule: shows correct date column per milestone
'     (Finish for SL1/SL2/SQ, Start for Set/Decon/Demo, Both for CV,
'      SDD with red if after Set Start)
'   - New helper: MilestoneHeaderToAbbrev, GetDefinedMilestoneAbbrevs
'   - New helper: DiscoverScheduleColumns (m_schedCols)
'
' Public Subs:
'   BuildRampAlignment      — Phase 1 entry point
'   RampAlignment_Generate  — Phase 2 entry point (called by button)
'====================================================================

Option Explicit

' Layout constants
Private Const DATA_START_ROW As Long = TIS_DATA_START_ROW
Private Const RAMP_SHEET_PREFIX As String = "Ramp - "
Private Const COL_START As Long = 2       ' Column B
Private Const COL_END As Long = 30        ' Wide enough for milestone columns

' Light table colors (consistent with DashboardBuilder)
Private Const TABLE_HEADER_BG As Long = 3022366      ' THEME_BG (dark header)
Private Const TABLE_ROW_BG As Long = 16777215         ' White
Private Const TABLE_ALT_ROW_BG As Long = 15790320     ' Light gray RGB(240, 240, 240)
Private Const TABLE_TEXT As Long = 2105376             ' Dark gray RGB(32, 32, 32)
Private Const TABLE_HEADER_TEXT As Long = 16777215     ' White
Private Const TABLE_GROUP_BG As Long = 15984610        ' RGB(226, 232, 243)
Private Const TABLE_SUBTITLE As Long = 7500402         ' RGB(100, 110, 114)

' SSPS highlight color
Private Const SSPS_BG As Long = 14024703               ' RGB(255, 255, 213) light yellow
Private Const GAP_RED_BG As Long = 14803198            ' RGB(254, 226, 226)
Private Const GAP_RED_TEXT As Long = 1792665            ' RGB(153, 27, 27)
Private Const GAP_GREEN_BG As Long = 15138012           ' RGB(220, 252, 231)
Private Const GAP_GREEN_TEXT As Long = 1467670           ' RGB(22, 101, 52)
Private Const GAP_YELLOW_BG As Long = 10092543          ' RGB(255, 251, 153)
Private Const GAP_YELLOW_TEXT As Long = 2442394          ' RGB(154, 52, 37)

' Module-level state
Private m_workSheet As Worksheet
Private m_wsName As String
Private m_tbl As ListObject
Private m_firstDataRow As Long
Private m_lastDataRow As Long

' Column indices on Working Sheet
Private m_groupCol As Long
Private m_nrCol As Long
Private m_ceidCol As Long
Private m_entityCodeCol As Long
Private m_entityTypeCol As Long
Private m_sddCol As Long
Private m_shipDateCol As Long
Private m_setStartCol As Long
Private m_sqFinishCol As Long

' Milestone column info (dynamic)
Private m_msStartCols As Collection     ' Collection of (colIndex, headerName)
Private m_msActualCols As Collection    ' Collection of (colIndex, milestoneName)
Private m_msStdCols As Collection       ' Collection of (colIndex, milestoneName)
Private m_msGapCols As Collection       ' Collection of (colIndex, milestoneName)
Private m_cvStartCol As Long
Private m_cvFinishCol As Long
Private m_statusCol As Long

' Schedule columns for Current Schedule section (Rev11)
Private m_schedCols As Collection       ' Collection of Array(colIndex, displayName, colType)
Private m_setStartColIdx As Long        ' Column index for Set Start (for SDD red check)

' Control panel references
Private m_controlSheet As Worksheet
Private m_dropdownCell As Range

'====================================================================
' PHASE 1: BUILD RAMP ALIGNMENT CONTROL PANEL
' Creates group dropdown + Generate button on the active sheet.
'====================================================================

Public Sub BuildRampAlignment()
    On Error GoTo ErrorHandler

    ' Validate prerequisites
    Set m_workSheet = FindWorkingSheet()
    If m_workSheet Is Nothing Then
        MsgBox "No Working Sheet found. Run Build Working Sheet first.", vbExclamation
        Exit Sub
    End If

    Set m_tbl = Nothing
    If m_workSheet.ListObjects.Count > 0 Then Set m_tbl = m_workSheet.ListObjects(1)
    If m_tbl Is Nothing Then
        MsgBox "No table found on Working Sheet.", vbExclamation
        Exit Sub
    End If

    m_firstDataRow = DATA_START_ROW + 1
    m_lastDataRow = m_tbl.Range.row + m_tbl.Range.Rows.Count - 1

    ' Discover columns
    DiscoverColumns

    If m_groupCol = 0 Then
        MsgBox "Group column not found on Working Sheet.", vbExclamation
        Exit Sub
    End If

    ' Collect unique groups
    Dim groups As Collection
    Set groups = CollectGroups()
    If groups.Count = 0 Then
        MsgBox "No groups found in data.", vbExclamation
        Exit Sub
    End If

    ' Build group list string for data validation
    Dim groupList As String
    Dim gi As Long
    groupList = ""
    For gi = 1 To groups.Count
        If gi > 1 Then groupList = groupList & ","
        groupList = groupList & CStr(groups(gi))
    Next gi

    ' Create control panel on the active sheet
    Set m_controlSheet = ActiveSheet
    Dim panelRow As Long
    panelRow = 2

    ' Title
    With m_controlSheet.Range(m_controlSheet.Cells(panelRow, COL_START), _
            m_controlSheet.Cells(panelRow, COL_START + 4))
        .Merge
        .Value = "Ramp Alignment Report Generator"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = THEME_ACCENT
        .Font.Name = THEME_FONT
    End With

    ' Group label
    panelRow = panelRow + 2
    With m_controlSheet.Cells(panelRow, COL_START)
        .Value = "Select Group:"
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = TABLE_TEXT
        .Font.Name = THEME_FONT
    End With

    ' Group dropdown
    Set m_dropdownCell = m_controlSheet.Cells(panelRow, COL_START + 2)
    With m_dropdownCell
        .Value = CStr(groups(1))   ' Default to first group
        .Font.Bold = True
        .Font.Size = 11
        .Font.Name = THEME_FONT
        .Interior.Color = RGB(255, 248, 220)
        .HorizontalAlignment = xlCenter
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(180, 180, 180)
        End With
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:=groupList
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    End With

    ' Store dropdown cell address in a named range so Generate can find it
    On Error Resume Next
    ThisWorkbook.Names("RampAlign_DropdownCell").Delete
    ThisWorkbook.Names("RampAlign_ControlSheet").Delete
    On Error GoTo ErrorHandler
    ThisWorkbook.Names.Add Name:="RampAlign_DropdownCell", _
        RefersTo:=m_dropdownCell
    ThisWorkbook.Names.Add Name:="RampAlign_ControlSheet", _
        RefersTo:=m_controlSheet.Range("A1")

    ' Generate button (shape)
    panelRow = panelRow + 2
    Dim btnLeft As Double, btnTop As Double
    btnLeft = m_controlSheet.Cells(panelRow, COL_START).Left
    btnTop = m_controlSheet.Cells(panelRow, COL_START).Top

    ' Remove old button if exists
    Dim shp As Shape
    For Each shp In m_controlSheet.Shapes
        If shp.Name = "RampAlign_GenerateBtn" Then
            shp.Delete
            Exit For
        End If
    Next shp

    Dim btn As Shape
    Set btn = m_controlSheet.Shapes.AddShape(msoShapeRoundedRectangle, _
        btnLeft, btnTop, 160, 32)
    With btn
        .Name = "RampAlign_GenerateBtn"
        .TextFrame2.TextRange.Text = "Generate Report"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Fill.ForeColor.RGB = RGB(91, 108, 249)  ' THEME_ACCENT
        .Line.Visible = msoFalse
        .OnAction = "RampAlignment_Generate"
    End With

    MsgBox "Control panel created." & vbCrLf & _
           "Select a group from the dropdown, then click Generate Report.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error in BuildRampAlignment: " & Err.Description, vbCritical
    DebugLog "RampAlignment ERROR (Phase1): " & Err.Description & " (#" & Err.Number & ")"
End Sub

'====================================================================
' PHASE 2: GENERATE RAMP ALIGNMENT REPORT
' Reads selected group, creates report sheet with all sections.
'====================================================================

Public Sub RampAlignment_Generate()
    On Error GoTo ErrorHandler

    Dim appSt As AppState
    Dim startTime As Double
    appSt = SaveAppState()
    SetPerformanceMode
    Application.DisplayAlerts = False
    startTime = Timer

    ' Retrieve dropdown cell from named range
    Dim ddRange As Range
    On Error Resume Next
    Set ddRange = ThisWorkbook.Names("RampAlign_DropdownCell").RefersToRange
    On Error GoTo ErrorHandler

    If ddRange Is Nothing Then
        MsgBox "Control panel not found. Run BuildRampAlignment first.", vbExclamation
        GoTo Cleanup
    End If

    Dim selectedGroup As String
    selectedGroup = Trim(CStr(ddRange.Value))
    If selectedGroup = "" Then
        MsgBox "Please select a group from the dropdown.", vbExclamation
        GoTo Cleanup
    End If

    ' Validate prerequisites
    Set m_workSheet = FindWorkingSheet()
    If m_workSheet Is Nothing Then
        MsgBox "No Working Sheet found.", vbExclamation
        GoTo Cleanup
    End If

    Set m_tbl = Nothing
    If m_workSheet.ListObjects.Count > 0 Then Set m_tbl = m_workSheet.ListObjects(1)
    If m_tbl Is Nothing Then
        MsgBox "No table found on Working Sheet.", vbExclamation
        GoTo Cleanup
    End If

    m_wsName = m_workSheet.Name
    m_firstDataRow = DATA_START_ROW + 1
    m_lastDataRow = m_tbl.Range.row + m_tbl.Range.Rows.Count - 1

    DiscoverColumns
    DiscoverMilestoneColumns
    DiscoverScheduleColumns

    If m_groupCol = 0 Then
        MsgBox "Group column not found.", vbExclamation
        GoTo Cleanup
    End If

    Application.StatusBar = "Generating Ramp Alignment for " & selectedGroup & "..."

    ' Read Working Sheet data into arrays (bulk read)
    Dim lastCol As Long
    lastCol = m_workSheet.Cells(DATA_START_ROW, m_workSheet.Columns.Count).End(xlToLeft).Column
    Dim nDataRows As Long
    nDataRows = m_lastDataRow - m_firstDataRow + 1
    If nDataRows < 1 Then
        MsgBox "No data rows found.", vbExclamation
        GoTo Cleanup
    End If

    Dim wsData() As Variant
    wsData = m_workSheet.Range(m_workSheet.Cells(m_firstDataRow, 1), _
        m_workSheet.Cells(m_lastDataRow, lastCol)).Value

    ' Filter rows for selected group
    Dim groupRows As Collection   ' Collection of row indices (1-based in wsData)
    Set groupRows = New Collection
    Dim ri As Long, gv As String
    For ri = 1 To nDataRows
        ' Only include Active systems in customer-facing Ramp reports
        If m_statusCol > 0 Then
            Dim rowStatus As String
            rowStatus = LCase(Trim(CStr(wsData(ri, m_statusCol) & "")))
            If rowStatus <> "active" Then GoTo NextRampRow
        End If
        If m_groupCol > 0 Then
            gv = Trim(CStr(wsData(ri, m_groupCol)))
            If gv = selectedGroup Then groupRows.Add ri
        End If
NextRampRow:
    Next ri

    If groupRows.Count = 0 Then
        MsgBox "No systems found for group: " & selectedGroup, vbExclamation
        GoTo Cleanup
    End If

    ' Sort groupRows by Entity Code
    Dim sortedRows() As Variant
    ReDim sortedRows(1 To groupRows.Count, 1 To 2) ' (entityCode, rowIdx)
    Dim si As Long
    For si = 1 To groupRows.Count
        Dim rowIdx As Long
        rowIdx = CLng(groupRows(si))
        If m_entityCodeCol > 0 Then
            sortedRows(si, 1) = CStr(wsData(rowIdx, m_entityCodeCol))
        Else
            sortedRows(si, 1) = CStr(si)
        End If
        sortedRows(si, 2) = rowIdx
    Next si
    SortRowsByColumn sortedRows, 1

    ' Collect unique CEIDs (sorted)
    Dim uniqueCeids As Collection
    Set uniqueCeids = New Collection
    Dim ceidDict As Object
    Set ceidDict = CreateObject("Scripting.Dictionary")
    If m_ceidCol > 0 Then
        For si = 1 To UBound(sortedRows, 1)
            Dim cv As String
            cv = Trim(CStr(wsData(CLng(sortedRows(si, 2)), m_ceidCol)))
            If cv <> "" And Not ceidDict.exists(cv) Then
                ceidDict(cv) = True
                uniqueCeids.Add cv
            End If
        Next si
    End If

    ' Read Milestones sheet for standard durations
    Dim milData() As Variant, milHdrs() As Variant
    Dim milRows As Long, milCols As Long
    Dim milCeidCol As Long
    Dim hasMilestones As Boolean
    hasMilestones = False
    If SheetExists(ThisWorkbook, TIS_SHEET_MILESTONES) Then
        Dim milSheet As Worksheet
        Set milSheet = ThisWorkbook.Sheets(TIS_SHEET_MILESTONES)
        Dim milHdrRow As Long
        milHdrRow = FindMilestoneHeaderRow(milSheet)
        If milHdrRow > 0 Then
            milCols = milSheet.Cells(milHdrRow, milSheet.Columns.Count).End(xlToLeft).Column
            milRows = milSheet.Cells(milSheet.Rows.Count, 1).End(xlUp).row
            If milRows > milHdrRow And milCols > 0 Then
                milHdrs = milSheet.Range(milSheet.Cells(milHdrRow, 1), _
                    milSheet.Cells(milHdrRow, milCols)).Value
                milData = milSheet.Range(milSheet.Cells(milHdrRow + 1, 1), _
                    milSheet.Cells(milRows, milCols)).Value
                ' Find CEID column in Milestones sheet
                milCeidCol = 0
                Dim mhi As Long
                For mhi = 1 To milCols
                    If LCase(Trim(CStr(milHdrs(1, mhi)))) = "ceid" Then
                        milCeidCol = mhi: Exit For
                    End If
                Next mhi
                If milCeidCol = 0 Then milCeidCol = 1  ' fallback to first col
                hasMilestones = True
            End If
        End If
    End If

    ' Create or replace report sheet
    Dim sheetName As String
    sheetName = RAMP_SHEET_PREFIX & selectedGroup
    If SheetExists(ThisWorkbook, sheetName) Then
        ThisWorkbook.Sheets(sheetName).Delete
    End If

    Dim rptWs As Worksheet
    Set rptWs = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    rptWs.Name = sheetName
    rptWs.Tab.Color = THEME_ACCENT

    ' Pre-fill light background
    rptWs.Cells.Interior.Color = RGB(245, 245, 250)
    rptWs.Cells.Font.Name = THEME_FONT
    rptWs.Cells.Font.Color = TABLE_TEXT
    rptWs.Columns("A").ColumnWidth = 2

    ' Track cells to unlock for user editing
    Dim editableCells As Collection
    Set editableCells = New Collection

    ' Build report sections
    Dim curRow As Long
    curRow = 1

    Application.StatusBar = "Ramp Alignment: Title..."
    WriteReportTitle rptWs, selectedGroup, curRow

    Application.StatusBar = "Ramp Alignment: Counters..."
    WriteCounters rptWs, wsData, sortedRows, curRow

    Application.StatusBar = "Ramp Alignment: Shipping Schedule..."
    WriteShippingSchedule rptWs, wsData, sortedRows, curRow

    Application.StatusBar = "Ramp Alignment: Standard Durations..."
    WriteStandardDurations rptWs, uniqueCeids, milData, milHdrs, milCeidCol, hasMilestones, curRow

    Application.StatusBar = "Ramp Alignment: Current Schedule..."
    WriteCurrentSchedule rptWs, wsData, sortedRows, curRow

    Application.StatusBar = "Ramp Alignment: Actual Duration & Gaps..."
    WriteActualDurationsAndGaps rptWs, wsData, sortedRows, curRow

    Application.StatusBar = "Ramp Alignment: Intel Requirements..."
    WriteIntelRequirements rptWs, editableCells, curRow

    Application.StatusBar = "Ramp Alignment: Sign-off Criteria..."
    WriteSignOffCriteria rptWs, uniqueCeids, editableCells, curRow

    Application.StatusBar = "Ramp Alignment: Conversion Scope..."
    WriteConversionScope rptWs, wsData, sortedRows, editableCells, curRow

    ' Protect sheet with editable cells unlocked
    Application.StatusBar = "Ramp Alignment: Protecting..."
    ' Hide gridlines
    rptWs.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 90

    ' Autofit column widths
    rptWs.Columns.AutoFit

    Application.Calculation = xlCalculationAutomatic
    rptWs.Calculate

    Application.ScreenUpdating = True
    MsgBox "Ramp Alignment report created: " & sheetName & vbCrLf & _
           "Time: " & Format(Timer - startTime, "0.00") & "s", vbInformation
    Application.ScreenUpdating = False

    GoTo Cleanup

ErrorHandler:
    MsgBox "Error in RampAlignment_Generate: " & Err.Description & vbCrLf & _
           "Error #: " & Err.Number, vbCritical
    DebugLog "RampAlignment ERROR (Phase2): " & Err.Description & " (#" & Err.Number & ")"

Cleanup:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    RestoreAppState appSt
    Set m_workSheet = Nothing
    Set m_tbl = Nothing
End Sub

'====================================================================
' COLUMN DISCOVERY
'====================================================================

Private Sub DiscoverColumns()
    Dim lastCol As Long
    lastCol = m_workSheet.Cells(DATA_START_ROW, m_workSheet.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Sub

    m_groupCol = 0: m_nrCol = 0: m_ceidCol = 0: m_entityCodeCol = 0
    m_entityTypeCol = 0: m_sddCol = 0: m_shipDateCol = 0
    m_setStartCol = 0: m_sqFinishCol = 0
    m_cvStartCol = 0: m_cvFinishCol = 0: m_statusCol = 0

    Dim hdrArr() As Variant
    hdrArr = m_workSheet.Range(m_workSheet.Cells(DATA_START_ROW, 1), _
        m_workSheet.Cells(DATA_START_ROW, lastCol)).Value

    Dim j As Long, rawH As String, hv As String
    For j = 1 To lastCol
        rawH = Trim(Replace(Replace(CStr(hdrArr(1, j)), vbLf, ""), vbCr, ""))
        hv = LCase(rawH)
        Select Case hv
            Case "group": m_groupCol = j
            Case "new/reused", "new/reused/demo", "new-reused": m_nrCol = j
            Case "ceid": m_ceidCol = j
            Case "entity code", "entitycode": m_entityCodeCol = j
            Case "entity type", "entitytype": m_entityTypeCol = j
            Case "sdd": m_sddCol = j
            Case "set start": m_setStartCol = j
            Case "supplier qual finish", "supplier qualfinish": m_sqFinishCol = j
            Case LCase(TIS_COL_STATUS): m_statusCol = j
        End Select

        ' Ship Date (multi-line header: "Ship" & vbLf & "Date")
        If hv = "shipdate" Or hv = "ship date" Then m_shipDateCol = j

        ' Conversion start/finish
        If InStr(1, hv, "convert", vbTextCompare) > 0 Or _
           InStr(1, hv, "cv ", vbTextCompare) > 0 Then
            If InStr(1, hv, "start", vbTextCompare) > 0 Then
                If m_cvStartCol = 0 Then m_cvStartCol = j
            End If
            If InStr(1, hv, "finish", vbTextCompare) > 0 Or _
               InStr(1, hv, "end", vbTextCompare) > 0 Then
                If m_cvFinishCol = 0 Then m_cvFinishCol = j
            End If
        End If
    Next j
End Sub

'====================================================================
' DISCOVER MILESTONE COLUMNS
' Finds Actual Duration, STD Duration, and Gap columns on Working Sheet.
' Also finds milestone start columns via GetMilestoneStartHeaders().
'====================================================================

Private Sub DiscoverMilestoneColumns()
    Set m_msStartCols = New Collection
    Set m_msActualCols = New Collection
    Set m_msStdCols = New Collection
    Set m_msGapCols = New Collection

    Dim lastCol As Long
    lastCol = m_workSheet.Cells(DATA_START_ROW, m_workSheet.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Sub

    Dim hdrArr() As Variant
    hdrArr = m_workSheet.Range(m_workSheet.Cells(DATA_START_ROW, 1), _
        m_workSheet.Cells(DATA_START_ROW, lastCol)).Value

    ' Get milestone start headers from Definitions
    Dim msHeaders As Collection
    Set msHeaders = GetMilestoneStartHeaders()

    ' Find milestone start columns
    Dim msItem As Variant, msName As String, j As Long, rawH As String
    For Each msItem In msHeaders
        msName = CStr(msItem)
        For j = 1 To lastCol
            rawH = Trim(Replace(Replace(CStr(hdrArr(1, j)), vbLf, ""), vbCr, ""))
            If StrComp(rawH, msName, vbTextCompare) = 0 Then
                m_msStartCols.Add Array(j, msName)
                Exit For
            End If
        Next j
    Next msItem

    ' Find Actual Duration, STD Duration, and Gap columns
    Dim hv As String
    For j = 1 To lastCol
        rawH = Trim(Replace(Replace(CStr(hdrArr(1, j)), vbLf, ""), vbCr, ""))
        hv = LCase(rawH)

        If Left(hv, Len("actualduration")) = "actualduration" Then
            Dim actName As String
            actName = Trim(Mid(rawH, Len("ActualDuration") + 1))
            If actName = "" Then actName = Mid(rawH, InStr(1, rawH, "n") + 1)
            ' Extract milestone name from multi-line header
            actName = ExtractMilestoneName(rawH, "Actual Duration")
            If actName <> "" Then m_msActualCols.Add Array(j, actName)

        ElseIf Left(hv, Len("stdduration")) = "stdduration" Then
            Dim stdName As String
            stdName = ExtractMilestoneName(rawH, "STD Duration")
            If stdName <> "" Then m_msStdCols.Add Array(j, stdName)

        ElseIf Left(hv, Len("gap")) = "gap" And Len(hv) > 3 Then
            Dim gapName As String
            gapName = ExtractMilestoneName(rawH, "Gap")
            If gapName <> "" Then m_msGapCols.Add Array(j, gapName)
        End If
    Next j
End Sub

'====================================================================
' DISCOVER SCHEDULE COLUMNS (Rev11)
' Finds the correct date column per milestone type for Current Schedule:
'   Set -> Set Start (start date)
'   SL1 -> SL1 Signoff Finish
'   SL2 -> SL2 Signoff Finish
'   SQ  -> Supplier Qual Finish
'   Decon -> Decon Start
'   Demo -> Demo Start
'   CV  -> Convert Start AND Convert Finish (both)
'   SDD -> SDD (with red if > Set Start)
'====================================================================

Private Sub DiscoverScheduleColumns()
    Set m_schedCols = New Collection
    m_setStartColIdx = 0

    Dim ws As Worksheet
    Set ws = m_workSheet
    Dim lastCol As Long
    lastCol = ws.Cells(DATA_START_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Sub

    Dim hdrArr() As Variant
    hdrArr = ws.Range(ws.Cells(DATA_START_ROW, 1), ws.Cells(DATA_START_ROW, lastCol)).Value

    Dim j As Long
    Dim rawH As String, hv As String

    ' Discover specific columns
    Dim setStartCol As Long, sl1FinCol As Long, sl2FinCol As Long, sqFinCol As Long
    Dim deconStartCol As Long, demoStartCol As Long
    Dim cvSchedStartCol As Long, cvSchedFinishCol As Long, sddSchedCol As Long

    setStartCol = 0: sl1FinCol = 0: sl2FinCol = 0: sqFinCol = 0
    deconStartCol = 0: demoStartCol = 0: cvSchedStartCol = 0: cvSchedFinishCol = 0: sddSchedCol = 0

    For j = 1 To UBound(hdrArr, 2)
        rawH = Trim(Replace(Replace(CStr(hdrArr(1, j)), vbLf, ""), vbCr, ""))
        hv = LCase(rawH)

        Select Case True
            Case hv = "set start" Or hv = "setstart"
                setStartCol = j
            Case hv = "sl1 signoff finish" Or hv = "sl1 finish" Or hv = "sl1signofffinish"
                sl1FinCol = j
            Case hv = "sl2 signoff finish" Or hv = "sl2 finish" Or hv = "sl2signofffinish"
                sl2FinCol = j
            Case hv = "supplier qual finish" Or hv = "sq finish" Or hv = "supplierqualfinish"
                sqFinCol = j
            Case hv = "decon start" Or hv = "deconstart"
                deconStartCol = j
            Case hv = "demo start" Or hv = "demostart"
                demoStartCol = j
            Case hv = "convert start" Or hv = "conversion start" Or hv = "convertstart" Or hv = "cv start"
                cvSchedStartCol = j
            Case hv = "convert finish" Or hv = "conversion finish" Or hv = "convertfinish" Or hv = "cv finish"
                cvSchedFinishCol = j
            Case hv = "sdd"
                sddSchedCol = j
        End Select
    Next j

    ' Store Set Start column index for SDD red comparison
    m_setStartColIdx = setStartCol

    ' Build schedule columns in display order
    If deconStartCol > 0 Then m_schedCols.Add Array(deconStartCol, "Decon Start", "start")
    If demoStartCol > 0 Then m_schedCols.Add Array(demoStartCol, "Demo Start", "start")
    If setStartCol > 0 Then m_schedCols.Add Array(setStartCol, "Set Start", "start")
    If sl1FinCol > 0 Then m_schedCols.Add Array(sl1FinCol, "SL1 Finish", "finish")
    If sl2FinCol > 0 Then m_schedCols.Add Array(sl2FinCol, "SL2 Finish", "finish")
    If cvSchedStartCol > 0 Then m_schedCols.Add Array(cvSchedStartCol, "CV Start", "start")
    If cvSchedFinishCol > 0 Then m_schedCols.Add Array(cvSchedFinishCol, "CV Finish", "finish")
    If sqFinCol > 0 Then m_schedCols.Add Array(sqFinCol, "SQ Finish", "finish")
    If sddSchedCol > 0 Then m_schedCols.Add Array(sddSchedCol, "SDD", "sdd")
End Sub

'====================================================================
' HELPER: EXTRACT MILESTONE NAME from multi-line header
' e.g., "ActualDurationSET" -> "SET", "STDDurationSL1" -> "SL1"
'====================================================================

Private Function ExtractMilestoneName(rawHeader As String, prefix As String) As String
    ExtractMilestoneName = ""
    ' Remove prefix (case-insensitive, ignoring spaces from vbLf removal)
    Dim cleaned As String
    cleaned = Trim(Replace(Replace(rawHeader, vbLf, ""), vbCr, ""))
    Dim prefixClean As String
    prefixClean = Replace(prefix, " ", "")

    If Len(cleaned) > Len(prefixClean) Then
        Dim remainder As String
        remainder = Mid(cleaned, Len(prefixClean) + 1)
        ExtractMilestoneName = Trim(remainder)
    End If

    ' Also handle "Total" — skip it
    If LCase(ExtractMilestoneName) = "total" Then ExtractMilestoneName = ""
End Function

'====================================================================
' HELPER: MAP DEFINITIONS HEADER TO MILESTONE ABBREVIATION (Rev11)
' Maps full Definitions milestone header names to the short abbreviations
' used as column headers on the Milestones sheet.
' e.g., "Set Start" -> "SET", "SL1 Signoff Start" -> "SL1",
'       "Convert Start" -> "CV", "Supplier Qual Start" -> "SQ"
'====================================================================

Private Function MilestoneHeaderToAbbrev(header As String) As String
    Dim h As String
    h = UCase(Trim(header))
    ' Extract first word
    Dim spacePos As Long
    spacePos = InStr(h, " ")
    Dim firstWord As String
    If spacePos > 0 Then firstWord = Left(h, spacePos - 1) Else firstWord = h

    Select Case firstWord
        Case "SET": MilestoneHeaderToAbbrev = "SET"
        Case "SL1": MilestoneHeaderToAbbrev = "SL1"
        Case "SL2": MilestoneHeaderToAbbrev = "SL2"
        Case "CONVERT", "CONVERSION": MilestoneHeaderToAbbrev = "CV"
        Case "SUPPLIER": MilestoneHeaderToAbbrev = "SQ"
        Case "DECON", "DECONTAMINATION": MilestoneHeaderToAbbrev = "DC"
        Case "DEMO", "DEMOLITION": MilestoneHeaderToAbbrev = "DM"
        Case "PRE-FAC", "PREFAC", "PRE": MilestoneHeaderToAbbrev = "PF"
        Case "MRCL": MilestoneHeaderToAbbrev = "MRCL"
        Case "SDD": MilestoneHeaderToAbbrev = "SDD"
        Case Else: MilestoneHeaderToAbbrev = firstWord
    End Select
End Function

'====================================================================
' HELPER: GET DEFINED MILESTONE ABBREVIATIONS (Rev11)
' Returns a Dictionary of milestone abbreviations defined on the
' Definitions sheet. Used to filter Milestones sheet columns in
' WriteStandardDurations so only relevant milestones are shown.
'====================================================================

Private Function GetDefinedMilestoneAbbrevs() As Object
    Dim definedAbbrevs As Object
    Set definedAbbrevs = CreateObject("Scripting.Dictionary")
    definedAbbrevs.CompareMode = vbTextCompare

    Dim msHeaders As Collection
    Set msHeaders = GetMilestoneStartHeaders()

    Dim msH As Variant
    For Each msH In msHeaders
        Dim abbr As String
        abbr = MilestoneHeaderToAbbrev(CStr(msH))
        If abbr <> "" And Not definedAbbrevs.exists(abbr) Then
            definedAbbrevs(abbr) = True
        End If
    Next msH

    Set GetDefinedMilestoneAbbrevs = definedAbbrevs
End Function

'====================================================================
' COLLECT GROUPS
'====================================================================

Private Function CollectGroups() As Collection
    Dim result As New Collection
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If m_groupCol = 0 Then
        Set CollectGroups = result
        Exit Function
    End If

    Dim grpArr() As Variant
    grpArr = m_workSheet.Range(m_workSheet.Cells(m_firstDataRow, m_groupCol), _
        m_workSheet.Cells(m_lastDataRow, m_groupCol)).Value

    Dim ri As Long, gv As String
    For ri = 1 To UBound(grpArr, 1)
        gv = Trim(CStr(grpArr(ri, 1)))
        If gv <> "" And Not dict.exists(gv) Then
            dict(gv) = True
            result.Add gv
        End If
    Next ri

    Set CollectGroups = result
End Function

'====================================================================
' SORT ROWS BY COLUMN — simple bubble sort for 2D variant array
'====================================================================

Private Sub SortRowsByColumn(arr() As Variant, sortCol As Long)
    Dim i As Long, j As Long, n As Long
    n = UBound(arr, 1)
    Dim cols As Long
    cols = UBound(arr, 2)
    Dim tmp As Variant, k As Long

    For i = 1 To n - 1
        For j = i + 1 To n
            If CStr(arr(j, sortCol)) < CStr(arr(i, sortCol)) Then
                For k = 1 To cols
                    tmp = arr(i, k)
                    arr(i, k) = arr(j, k)
                    arr(j, k) = tmp
                Next k
            End If
        Next j
    Next i
End Sub

'====================================================================
' FIND MILESTONE HEADER ROW on Milestones sheet
'====================================================================

Private Function FindMilestoneHeaderRow(milSheet As Worksheet) As Long
    FindMilestoneHeaderRow = 0
    Dim r As Long
    For r = 1 To 5
        Dim cVal As String
        Dim c As Long
        For c = 1 To 10
            cVal = LCase(Trim(CStr(milSheet.Cells(r, c).Value)))
            If cVal = "ceid" Or cVal = "entity code" Then
                FindMilestoneHeaderRow = r
                Exit Function
            End If
        Next c
    Next r
    ' Default: assume row 1
    If milSheet.Cells(1, 1).Value <> "" Then FindMilestoneHeaderRow = 1
End Function

'====================================================================
' SECTION 1: TITLE BAR
'====================================================================

Private Sub WriteReportTitle(rptWs As Worksheet, groupName As String, ByRef curRow As Long)
    With rptWs.Range(rptWs.Cells(curRow, COL_START), rptWs.Cells(curRow, COL_START + 10))
        .Merge
        .Value = "Ramp Alignment Report  —  " & groupName
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = THEME_WHITE
        .Interior.Color = THEME_ACCENT
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    rptWs.Rows(curRow).RowHeight = 40

    curRow = curRow + 1
    With rptWs.Range(rptWs.Cells(curRow, COL_START), rptWs.Cells(curRow, COL_START + 10))
        .Merge
        .Value = "Generated: " & Format(Now, "mm/dd/yyyy hh:mm") & "  |  Source: " & m_wsName
        .Font.Size = 9
        .Font.Color = THEME_TEXT_SEC
        .Interior.Color = THEME_SURFACE
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    rptWs.Rows(curRow).RowHeight = 22

    curRow = curRow + 2
End Sub

'====================================================================
' SECTION 2: COUNTERS
'====================================================================

Private Sub WriteCounters(rptWs As Worksheet, wsData() As Variant, _
        sortedRows() As Variant, ByRef curRow As Long)

    WriteSectionHeader rptWs, curRow, "System Summary", ""

    Dim total As Long, nNew As Long, nReused As Long, nDemo As Long, nConversion As Long
    total = UBound(sortedRows, 1)
    nNew = 0: nReused = 0: nDemo = 0: nConversion = 0

    Dim si As Long, rowIdx As Long, nrVal As String
    For si = 1 To total
        rowIdx = CLng(sortedRows(si, 2))
        If m_nrCol > 0 Then
            nrVal = LCase(Trim(CStr(wsData(rowIdx, m_nrCol))))
            Select Case nrVal
                Case "new": nNew = nNew + 1
                Case "reused": nReused = nReused + 1
                Case "demo": nDemo = nDemo + 1
            End Select
        End If
        ' Conversion: has CV start date
        If m_cvStartCol > 0 Then
            If IsDate(wsData(rowIdx, m_cvStartCol)) Then nConversion = nConversion + 1
        End If
    Next si

    ' Write counter cards
    Dim labels As Variant, values As Variant, colors As Variant
    labels = Array("Total", "New", "Reused", "Demo", "Conversion")
    values = Array(total, nNew, nReused, nDemo, nConversion)
    colors = Array(THEME_TEXT, THEME_SUCCESS, THEME_ACCENT, THEME_ACCENT2, THEME_WARNING)

    Dim ci As Long, cardCol As Long
    cardCol = COL_START
    For ci = 0 To 4
        Dim cardRng As Range
        Set cardRng = rptWs.Range(rptWs.Cells(curRow, cardCol), rptWs.Cells(curRow + 1, cardCol + 1))
        FormatCardStyle cardRng, THEME_SURFACE, CLng(colors(ci))

        ' Label
        With rptWs.Range(rptWs.Cells(curRow, cardCol), rptWs.Cells(curRow, cardCol + 1))
            .Merge
            .Value = CStr(labels(ci))
            .Font.Size = 9
            .Font.Color = THEME_TEXT_SEC
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With

        ' Value
        With rptWs.Range(rptWs.Cells(curRow + 1, cardCol), rptWs.Cells(curRow + 1, cardCol + 1))
            .Merge
            .Value = CLng(values(ci))
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = CLng(colors(ci))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormat = "#,##0"
        End With

        cardCol = cardCol + 3
    Next ci

    rptWs.Rows(curRow).RowHeight = 22
    rptWs.Rows(curRow + 1).RowHeight = 36

    curRow = curRow + 4
End Sub

'====================================================================
' SECTION 3: SHIPPING SCHEDULE
'====================================================================

Private Sub WriteShippingSchedule(rptWs As Worksheet, wsData() As Variant, _
        sortedRows() As Variant, ByRef curRow As Long)

    WriteSectionHeader rptWs, curRow, "Shipping Schedule", _
        "Entity codes sorted by name with SDD and Ship Date"

    If m_sddCol = 0 And m_shipDateCol = 0 Then
        rptWs.Cells(curRow, COL_START).Value = "SDD and Ship Date columns not found."
        rptWs.Cells(curRow, COL_START).Font.Size = 9
        curRow = curRow + 2
        Exit Sub
    End If

    ' Header row
    Dim hdrRow As Long: hdrRow = curRow
    WriteTableHeaderCell rptWs, hdrRow, COL_START, "Entity Code", 20
    WriteTableHeaderCell rptWs, hdrRow, COL_START + 1, "SDD", 14
    WriteTableHeaderCell rptWs, hdrRow, COL_START + 2, "Ship Date", 14
    rptWs.Rows(hdrRow).RowHeight = 24
    curRow = curRow + 1

    ' Data rows
    Dim si As Long, rowIdx As Long
    For si = 1 To UBound(sortedRows, 1)
        rowIdx = CLng(sortedRows(si, 2))
        Dim ecVal As String
        ecVal = ""
        If m_entityCodeCol > 0 Then ecVal = CStr(wsData(rowIdx, m_entityCodeCol))

        rptWs.Cells(curRow, COL_START).Value = ecVal
        If m_sddCol > 0 Then
            If IsDate(wsData(rowIdx, m_sddCol)) Then
                rptWs.Cells(curRow, COL_START + 1).Value = CDate(wsData(rowIdx, m_sddCol))
                rptWs.Cells(curRow, COL_START + 1).NumberFormat = "mm/dd/yyyy"
            End If
        End If
        If m_shipDateCol > 0 Then
            If IsDate(wsData(rowIdx, m_shipDateCol)) Then
                rptWs.Cells(curRow, COL_START + 2).Value = CDate(wsData(rowIdx, m_shipDateCol))
                rptWs.Cells(curRow, COL_START + 2).NumberFormat = "mm/dd/yyyy"
            End If
        End If

        FormatDataRow rptWs, curRow, COL_START, COL_START + 2, (si Mod 2 = 0)
        curRow = curRow + 1
    Next si

    curRow = curRow + 1
End Sub

'====================================================================
' SECTION 4: STANDARD DURATIONS (Rev11 — filtered to Definitions)
' Only shows milestone columns that correspond to milestones defined
' on the Definitions sheet, using GetDefinedMilestoneAbbrevs().
'====================================================================

Private Sub WriteStandardDurations(rptWs As Worksheet, uniqueCeids As Collection, _
        milData() As Variant, milHdrs() As Variant, milCeidCol As Long, _
        hasMilestones As Boolean, ByRef curRow As Long)

    WriteSectionHeader rptWs, curRow, "Standard Durations", _
        "Per CEID from Milestones sheet (filtered to Definitions milestones)"

    If uniqueCeids.Count = 0 Or Not hasMilestones Then
        rptWs.Cells(curRow, COL_START).Value = "Milestones data or CEIDs not available."
        rptWs.Cells(curRow, COL_START).Font.Size = 9
        curRow = curRow + 2
        Exit Sub
    End If

    ' Get defined milestone abbreviations from Definitions sheet (Rev11)
    Dim definedAbbrevs As Object
    Set definedAbbrevs = GetDefinedMilestoneAbbrevs()

    ' Determine milestone columns in Milestones sheet (skip CEID col)
    ' Only include columns whose header matches a defined abbreviation (Rev11)
    Dim milMsCols As Collection  ' Collection of Array(colIdx, headerName)
    Set milMsCols = New Collection
    Dim mc As Long
    For mc = 1 To UBound(milHdrs, 2)
        If mc <> milCeidCol Then
            Dim mhVal As String
            mhVal = Trim(CStr(milHdrs(1, mc)))
            If mhVal <> "" Then
                ' Rev11: Only include milestones defined on Definitions sheet
                ' Convert Milestones sheet header (e.g. "Set - SL1") to abbreviation
                ' (e.g. "SET") before matching against Definitions abbreviations
                Dim mhAbbrev As String
                mhAbbrev = MilestoneHeaderToAbbrev(mhVal)
                If definedAbbrevs.exists(mhAbbrev) Then
                    milMsCols.Add Array(mc, mhVal)
                End If
            End If
        End If
    Next mc

    If milMsCols.Count = 0 Then
        rptWs.Cells(curRow, COL_START).Value = "No matching milestones found between Milestones sheet and Definitions."
        rptWs.Cells(curRow, COL_START).Font.Size = 9
        curRow = curRow + 2
        Exit Sub
    End If

    ' Header row
    Dim hdrRow As Long: hdrRow = curRow
    WriteTableHeaderCell rptWs, hdrRow, COL_START, "CEID", 16
    Dim hci As Long, mcItem As Variant
    hci = 1
    For Each mcItem In milMsCols
        WriteTableHeaderCell rptWs, hdrRow, COL_START + hci, CStr(mcItem(1)), 10
        hci = hci + 1
    Next mcItem
    WriteTableHeaderCell rptWs, hdrRow, COL_START + hci, "Total", 10
    rptWs.Rows(hdrRow).RowHeight = 24
    curRow = curRow + 1

    ' Build CEID->row map for Milestones sheet
    Dim milCeidMap As Object
    Set milCeidMap = CreateObject("Scripting.Dictionary")
    If hasMilestones Then
        Dim mr As Long
        For mr = 1 To UBound(milData, 1)
            Dim milCeidVal As String
            milCeidVal = Trim(CStr(milData(mr, milCeidCol)))
            If milCeidVal <> "" And Not milCeidMap.exists(milCeidVal) Then
                milCeidMap(milCeidVal) = mr
            End If
        Next mr
    End If

    ' Data rows (one per unique CEID)
    Dim ceidItem As Variant, ceidIdx As Long: ceidIdx = 0
    For Each ceidItem In uniqueCeids
        Dim ceidName As String
        ceidName = CStr(ceidItem)
        rptWs.Cells(curRow, COL_START).Value = ceidName

        Dim totalDur As Double: totalDur = 0
        hci = 1
        If milCeidMap.exists(ceidName) Then
            Dim dataRowIdx As Long
            dataRowIdx = CLng(milCeidMap(ceidName))
            For Each mcItem In milMsCols
                Dim durVal As Variant
                durVal = milData(dataRowIdx, CLng(mcItem(0)))
                If IsNumeric(durVal) And Not IsEmpty(durVal) Then
                    rptWs.Cells(curRow, COL_START + hci).Value = CDbl(durVal)
                    totalDur = totalDur + CDbl(durVal)
                End If
                rptWs.Cells(curRow, COL_START + hci).HorizontalAlignment = xlCenter
                hci = hci + 1
            Next mcItem
        Else
            For Each mcItem In milMsCols
                hci = hci + 1
            Next mcItem
        End If

        ' Total column
        rptWs.Cells(curRow, COL_START + hci).Value = totalDur
        rptWs.Cells(curRow, COL_START + hci).Font.Bold = True
        rptWs.Cells(curRow, COL_START + hci).HorizontalAlignment = xlCenter

        FormatDataRow rptWs, curRow, COL_START, COL_START + hci, (ceidIdx Mod 2 = 0)
        ceidIdx = ceidIdx + 1
        curRow = curRow + 1
    Next ceidItem

    curRow = curRow + 1
End Sub

'====================================================================
' SECTION 5: CURRENT SCHEDULE (Rev11 — correct date columns)
' Shows the correct date column per milestone type:
'   Set -> Set Start, SL1/SL2/SQ -> Finish dates
'   Decon/Demo -> Start dates, CV -> Both Start and Finish
'   SDD -> SDD date (red if after Set Start)
'====================================================================

Private Sub WriteCurrentSchedule(rptWs As Worksheet, wsData() As Variant, _
        sortedRows() As Variant, ByRef curRow As Long)

    WriteSectionHeader rptWs, curRow, "Current Schedule", _
        "Milestone dates per definitions configuration"

    If m_schedCols.Count = 0 Then
        rptWs.Cells(curRow, COL_START).Value = "No schedule columns found."
        rptWs.Cells(curRow, COL_START).Font.Size = 9
        curRow = curRow + 2
        Exit Sub
    End If

    ' Header row
    Dim hdrRow As Long: hdrRow = curRow
    WriteTableHeaderCell rptWs, hdrRow, COL_START, "Entity Code", 20

    Dim hci As Long: hci = 1
    Dim msItem As Variant
    For Each msItem In m_schedCols
        WriteTableHeaderCell rptWs, hdrRow, COL_START + hci, CStr(msItem(1)), 14
        hci = hci + 1
    Next msItem
    rptWs.Rows(hdrRow).RowHeight = 24
    Dim lastDataCol As Long: lastDataCol = COL_START + hci - 1
    curRow = curRow + 1

    ' Data rows
    Dim si As Long, rowIdx As Long
    For si = 1 To UBound(sortedRows, 1)
        rowIdx = CLng(sortedRows(si, 2))

        ' Entity Code
        If m_entityCodeCol > 0 Then _
            rptWs.Cells(curRow, COL_START).Value = CStr(wsData(rowIdx, m_entityCodeCol))

        ' Get Set Start date for SDD comparison
        Dim setStartDate As Variant
        setStartDate = Empty
        If m_setStartColIdx > 0 And m_setStartColIdx <= UBound(wsData, 2) Then
            If IsDate(wsData(rowIdx, m_setStartColIdx)) Then
                setStartDate = CDate(wsData(rowIdx, m_setStartColIdx))
            End If
        End If

        hci = 1
        For Each msItem In m_schedCols
            Dim colIdx As Long: colIdx = CLng(msItem(0))
            Dim colType As String: colType = CStr(msItem(2))

            If colIdx <= UBound(wsData, 2) Then
                If IsDate(wsData(rowIdx, colIdx)) Then
                    rptWs.Cells(curRow, COL_START + hci).Value = CDate(wsData(rowIdx, colIdx))
                    rptWs.Cells(curRow, COL_START + hci).NumberFormat = "mm/dd/yy"

                    ' SDD: mark red if after Set Start
                    If colType = "sdd" And Not IsEmpty(setStartDate) Then
                        If CDate(wsData(rowIdx, colIdx)) > CDate(setStartDate) Then
                            rptWs.Cells(curRow, COL_START + hci).Font.Color = GAP_RED_TEXT
                            rptWs.Cells(curRow, COL_START + hci).Font.Bold = True
                        End If
                    End If
                End If
            End If
            rptWs.Cells(curRow, COL_START + hci).HorizontalAlignment = xlCenter
            hci = hci + 1
        Next msItem

        FormatDataRow rptWs, curRow, COL_START, lastDataCol, (si Mod 2 = 0)
        curRow = curRow + 1
    Next si

    curRow = curRow + 1
End Sub

'====================================================================
' SECTION 6: ACTUAL DURATION & GAPS
'====================================================================

Private Sub WriteActualDurationsAndGaps(rptWs As Worksheet, wsData() As Variant, _
        sortedRows() As Variant, ByRef curRow As Long)

    WriteSectionHeader rptWs, curRow, "Actual Duration & Gaps to Standard", _
        "Gap = Actual Duration - Standard Duration  (negative = behind schedule)"

    If m_msActualCols.Count = 0 And m_msGapCols.Count = 0 Then
        rptWs.Cells(curRow, COL_START).Value = "No duration/gap columns found on Working Sheet."
        rptWs.Cells(curRow, COL_START).Font.Size = 9
        curRow = curRow + 2
        Exit Sub
    End If

    ' Build paired columns: Actual | Gap for each milestone
    Dim pairCols As Collection  ' Collection of Array(actualColIdx, gapColIdx, milestoneName)
    Set pairCols = New Collection

    Dim actItem As Variant, gItem As Variant
    Dim actName As String, gapName As String
    Dim gapIdx As Long

    For Each actItem In m_msActualCols
        actName = CStr(actItem(1))
        gapIdx = 0
        ' Find matching gap column
        For Each gItem In m_msGapCols
            gapName = CStr(gItem(1))
            If StrComp(actName, gapName, vbTextCompare) = 0 Then
                gapIdx = CLng(gItem(0))
                Exit For
            End If
        Next gItem
        pairCols.Add Array(CLng(actItem(0)), gapIdx, actName)
    Next actItem

    ' Header row
    Dim hdrRow As Long: hdrRow = curRow
    WriteTableHeaderCell rptWs, hdrRow, COL_START, "Entity Code", 20
    Dim hci As Long: hci = 1
    Dim pItem As Variant
    For Each pItem In pairCols
        WriteTableHeaderCell rptWs, hdrRow, COL_START + hci, CStr(pItem(2)) & " Actual", 10
        hci = hci + 1
        If CLng(pItem(1)) > 0 Then
            WriteTableHeaderCell rptWs, hdrRow, COL_START + hci, CStr(pItem(2)) & " Gap", 10
            hci = hci + 1
        End If
    Next pItem
    rptWs.Rows(hdrRow).RowHeight = 24
    Dim lastDataCol As Long: lastDataCol = COL_START + hci - 1
    curRow = curRow + 1

    ' Data rows
    Dim si As Long, rowIdx As Long
    For si = 1 To UBound(sortedRows, 1)
        rowIdx = CLng(sortedRows(si, 2))

        If m_entityCodeCol > 0 Then _
            rptWs.Cells(curRow, COL_START).Value = CStr(wsData(rowIdx, m_entityCodeCol))

        hci = 1
        For Each pItem In pairCols
            ' Actual duration
            Dim actVal As Variant
            actVal = wsData(rowIdx, CLng(pItem(0)))
            If IsNumeric(actVal) And Not IsEmpty(actVal) Then
                rptWs.Cells(curRow, COL_START + hci).Value = actVal
            End If
            rptWs.Cells(curRow, COL_START + hci).HorizontalAlignment = xlCenter
            hci = hci + 1

            ' Gap
            If CLng(pItem(1)) > 0 Then
                Dim gapVal As Variant
                gapVal = wsData(rowIdx, CLng(pItem(1)))
                If IsNumeric(gapVal) And Not IsEmpty(gapVal) Then
                    rptWs.Cells(curRow, COL_START + hci).Value = gapVal
                    ' Conditional formatting inline
                    If CDbl(gapVal) < 0 Then
                        rptWs.Cells(curRow, COL_START + hci).Interior.Color = GAP_RED_BG
                        rptWs.Cells(curRow, COL_START + hci).Font.Color = GAP_RED_TEXT
                        rptWs.Cells(curRow, COL_START + hci).Font.Bold = True
                    ElseIf CDbl(gapVal) >= 0 And CDbl(gapVal) <= 4 Then
                        rptWs.Cells(curRow, COL_START + hci).Interior.Color = GAP_GREEN_BG
                        rptWs.Cells(curRow, COL_START + hci).Font.Color = GAP_GREEN_TEXT
                    Else
                        rptWs.Cells(curRow, COL_START + hci).Interior.Color = GAP_YELLOW_BG
                        rptWs.Cells(curRow, COL_START + hci).Font.Color = GAP_YELLOW_TEXT
                    End If
                End If
                rptWs.Cells(curRow, COL_START + hci).HorizontalAlignment = xlCenter
                hci = hci + 1
            End If
        Next pItem

        ' Base row formatting (only for non-gap cells)
        rptWs.Cells(curRow, COL_START).Font.Size = 9
        rptWs.Cells(curRow, COL_START).Font.Color = TABLE_TEXT
        If si Mod 2 = 0 Then
            If rptWs.Cells(curRow, COL_START).Interior.Color = RGB(245, 245, 250) Then
                rptWs.Cells(curRow, COL_START).Interior.Color = TABLE_ALT_ROW_BG
            End If
        End If
        curRow = curRow + 1
    Next si

    curRow = curRow + 1
End Sub

'====================================================================
' SECTION 7: INTEL REQUIREMENTS PER MILESTONE
'====================================================================

Private Sub WriteIntelRequirements(rptWs As Worksheet, editableCells As Collection, _
        ByRef curRow As Long)

    WriteSectionHeader rptWs, curRow, "Intel Requirements per Milestone", _
        "Fill in requirements and SSPS (System/Site Preparation Specification) needs"

    ' Get milestone names from Definitions sheet
    Dim milestoneNames As Collection
    Set milestoneNames = New Collection

    If SheetExists(ThisWorkbook, TIS_SHEET_DEFINITIONS) Then
        Dim defWs As Worksheet
        Set defWs = ThisWorkbook.Sheets(TIS_SHEET_DEFINITIONS)
        Dim lastRow As Long
        lastRow = defWs.Cells(defWs.Rows.Count, 1).End(xlUp).row
        Dim di As Long
        For di = 2 To lastRow
            Dim hn As String
            hn = Trim(CStr(defWs.Cells(di, 1).Value))
            If hn <> "" Then
                ' Include milestone-related headers (start or finish)
                Dim hvLower As String
                hvLower = LCase(hn)
                If InStr(1, hvLower, "start", vbTextCompare) > 0 Then
                    ' Extract milestone name (before "Start")
                    Dim msName As String
                    msName = Trim(Replace(hn, "Start", "", , , vbTextCompare))
                    msName = Trim(Replace(msName, " - ", " "))
                    If msName <> "" Then
                        ' Avoid duplicates
                        Dim isDup As Boolean: isDup = False
                        Dim chk As Variant
                        For Each chk In milestoneNames
                            If StrComp(CStr(chk), msName, vbTextCompare) = 0 Then isDup = True: Exit For
                        Next chk
                        If Not isDup Then milestoneNames.Add msName
                    End If
                End If
            End If
        Next di
    End If

    ' Also add standard phases that might not be in Definitions
    Dim fixedPhases As Variant
    fixedPhases = Array("Pre-Fac", "Decon", "Demo", "SDD")
    Dim fi As Long
    For fi = 0 To UBound(fixedPhases)
        Dim isDup2 As Boolean: isDup2 = False
        Dim chk2 As Variant
        For Each chk2 In milestoneNames
            If StrComp(CStr(chk2), CStr(fixedPhases(fi)), vbTextCompare) = 0 Then isDup2 = True: Exit For
        Next chk2
        If Not isDup2 Then milestoneNames.Add CStr(fixedPhases(fi))
    Next fi

    If milestoneNames.Count = 0 Then
        rptWs.Cells(curRow, COL_START).Value = "No milestones found."
        rptWs.Cells(curRow, COL_START).Font.Size = 9
        curRow = curRow + 2
        Exit Sub
    End If

    ' Header row
    Dim hdrRow As Long: hdrRow = curRow
    WriteTableHeaderCell rptWs, hdrRow, COL_START, "Milestone", 20
    WriteTableHeaderCell rptWs, hdrRow, COL_START + 1, "Intel Requirements", 40
    WriteTableHeaderCell rptWs, hdrRow, COL_START + 2, "SSPS Requirements", 40
    rptWs.Rows(hdrRow).RowHeight = 24
    curRow = curRow + 1

    ' Data rows (one per milestone)
    Dim msItem As Variant, msIdx As Long: msIdx = 0
    For Each msItem In milestoneNames
        Dim milName As String
        milName = CStr(msItem)
        rptWs.Cells(curRow, COL_START).Value = milName
        rptWs.Cells(curRow, COL_START).Font.Bold = True
        rptWs.Cells(curRow, COL_START).Font.Size = 10
        rptWs.Cells(curRow, COL_START).Font.Color = TABLE_TEXT

        ' Editable cells
        editableCells.Add rptWs.Cells(curRow, COL_START + 1)
        editableCells.Add rptWs.Cells(curRow, COL_START + 2)

        ' Check if SSPS-related (highlight row)
        Dim isSsps As Boolean
        isSsps = (InStr(1, LCase(milName), "ssps", vbTextCompare) > 0) Or _
                 (InStr(1, LCase(milName), "site prep", vbTextCompare) > 0) Or _
                 (InStr(1, LCase(milName), "system prep", vbTextCompare) > 0)

        If isSsps Then
            rptWs.Range(rptWs.Cells(curRow, COL_START), rptWs.Cells(curRow, COL_START + 2)).Interior.Color = SSPS_BG
            rptWs.Cells(curRow, COL_START).Font.Color = RGB(146, 64, 14)
        Else
            FormatDataRow rptWs, curRow, COL_START, COL_START + 2, (msIdx Mod 2 = 0)
        End If

        rptWs.Rows(curRow).RowHeight = 28  ' Taller for writing space
        msIdx = msIdx + 1
        curRow = curRow + 1
    Next msItem

    curRow = curRow + 1
End Sub

'====================================================================
' SECTION 8: SIGN-OFF CRITERIA
'====================================================================

Private Sub WriteSignOffCriteria(rptWs As Worksheet, uniqueCeids As Collection, _
        editableCells As Collection, ByRef curRow As Long)

    WriteSectionHeader rptWs, curRow, "Sign-off Criteria", _
        "Per CEID — fill in sign-off criteria for each system"

    If uniqueCeids.Count = 0 Then
        rptWs.Cells(curRow, COL_START).Value = "No CEIDs found."
        rptWs.Cells(curRow, COL_START).Font.Size = 9
        curRow = curRow + 2
        Exit Sub
    End If

    ' Header row
    Dim hdrRow As Long: hdrRow = curRow
    WriteTableHeaderCell rptWs, hdrRow, COL_START, "CEID", 20
    WriteTableHeaderCell rptWs, hdrRow, COL_START + 1, "Sign-off Criteria", 60
    rptWs.Rows(hdrRow).RowHeight = 24
    curRow = curRow + 1

    ' Data rows
    Dim ceidItem As Variant, ceidIdx As Long: ceidIdx = 0
    For Each ceidItem In uniqueCeids
        rptWs.Cells(curRow, COL_START).Value = CStr(ceidItem)
        rptWs.Cells(curRow, COL_START).Font.Bold = True
        rptWs.Cells(curRow, COL_START).Font.Size = 10
        rptWs.Cells(curRow, COL_START).Font.Color = TABLE_TEXT

        editableCells.Add rptWs.Cells(curRow, COL_START + 1)

        FormatDataRow rptWs, curRow, COL_START, COL_START + 1, (ceidIdx Mod 2 = 0)
        rptWs.Rows(curRow).RowHeight = 28
        ceidIdx = ceidIdx + 1
        curRow = curRow + 1
    Next ceidItem

    curRow = curRow + 1
End Sub

'====================================================================
' SECTION 9: CONVERSION SCOPE
'====================================================================

Private Sub WriteConversionScope(rptWs As Worksheet, wsData() As Variant, _
        sortedRows() As Variant, editableCells As Collection, ByRef curRow As Long)

    WriteSectionHeader rptWs, curRow, "Conversion Scope", _
        "Entity codes with conversion (CV) milestones — fill in scope details"

    If m_cvStartCol = 0 Or m_cvFinishCol = 0 Then
        rptWs.Cells(curRow, COL_START).Value = "CV Start/Finish columns not found on Working Sheet."
        rptWs.Cells(curRow, COL_START).Font.Size = 9
        curRow = curRow + 2
        Exit Sub
    End If

    ' Find entities with conversion dates
    Dim cvRows As Collection
    Set cvRows = New Collection
    Dim si As Long, rowIdx As Long
    For si = 1 To UBound(sortedRows, 1)
        rowIdx = CLng(sortedRows(si, 2))
        If IsDate(wsData(rowIdx, m_cvStartCol)) And IsDate(wsData(rowIdx, m_cvFinishCol)) Then
            cvRows.Add rowIdx
        End If
    Next si

    If cvRows.Count = 0 Then
        rptWs.Cells(curRow, COL_START).Value = "No conversion projects in this group."
        rptWs.Cells(curRow, COL_START).Font.Size = 9
        rptWs.Cells(curRow, COL_START).Font.Italic = True
        curRow = curRow + 2
        Exit Sub
    End If

    ' Header row
    Dim hdrRow As Long: hdrRow = curRow
    WriteTableHeaderCell rptWs, hdrRow, COL_START, "Entity Code", 20
    WriteTableHeaderCell rptWs, hdrRow, COL_START + 1, "CV Duration (days)", 16
    WriteTableHeaderCell rptWs, hdrRow, COL_START + 2, "Conversion Scope", 50
    rptWs.Rows(hdrRow).RowHeight = 24
    curRow = curRow + 1

    ' Data rows
    Dim cvItem As Variant, cvIdx As Long: cvIdx = 0
    For Each cvItem In cvRows
        rowIdx = CLng(cvItem)

        If m_entityCodeCol > 0 Then _
            rptWs.Cells(curRow, COL_START).Value = CStr(wsData(rowIdx, m_entityCodeCol))
        rptWs.Cells(curRow, COL_START).Font.Size = 10
        rptWs.Cells(curRow, COL_START).Font.Color = TABLE_TEXT

        ' CV Duration = CV Finish - CV Start
        Dim cvDur As Long
        cvDur = CLng(CDate(wsData(rowIdx, m_cvFinishCol)) - CDate(wsData(rowIdx, m_cvStartCol)))
        rptWs.Cells(curRow, COL_START + 1).Value = cvDur
        rptWs.Cells(curRow, COL_START + 1).HorizontalAlignment = xlCenter
        rptWs.Cells(curRow, COL_START + 1).Font.Bold = True

        editableCells.Add rptWs.Cells(curRow, COL_START + 2)

        FormatDataRow rptWs, curRow, COL_START, COL_START + 2, (cvIdx Mod 2 = 0)
        rptWs.Rows(curRow).RowHeight = 28
        cvIdx = cvIdx + 1
        curRow = curRow + 1
    Next cvItem

    curRow = curRow + 1
End Sub

'====================================================================
' FORMATTING HELPERS
'====================================================================

Private Sub WriteSectionHeader(rptWs As Worksheet, ByRef curRow As Long, _
        title As String, subtitle As String)

    With rptWs.Range(rptWs.Cells(curRow, COL_START), rptWs.Cells(curRow, COL_START + 10))
        .Merge
        .Value = title
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = TABLE_TEXT
        .Font.Name = THEME_FONT
        .HorizontalAlignment = xlLeft
    End With
    rptWs.Rows(curRow).RowHeight = 24
    curRow = curRow + 1

    If subtitle <> "" Then
        With rptWs.Range(rptWs.Cells(curRow, COL_START), rptWs.Cells(curRow, COL_START + 10))
            .Merge
            .Value = subtitle
            .Font.Size = 9
            .Font.Italic = True
            .Font.Color = TABLE_SUBTITLE
            .Font.Name = THEME_FONT
            .HorizontalAlignment = xlLeft
        End With
        rptWs.Rows(curRow).RowHeight = 18
        curRow = curRow + 1
    End If
End Sub

Private Sub WriteTableHeaderCell(rptWs As Worksheet, row As Long, col As Long, _
        headerText As String, colWidth As Double)
    With rptWs.Cells(row, col)
        .Value = headerText
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = TABLE_HEADER_TEXT
        .Font.Name = THEME_FONT
        .Interior.Color = TABLE_HEADER_BG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).Color = THEME_ACCENT
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    rptWs.Columns(col).ColumnWidth = colWidth
End Sub

Private Sub FormatDataRow(rptWs As Worksheet, row As Long, _
        colStart As Long, colEnd As Long, isAlt As Boolean)
    Dim rng As Range
    Set rng = rptWs.Range(rptWs.Cells(row, colStart), rptWs.Cells(row, colEnd))

    Dim c As Long
    For c = colStart To colEnd
        rptWs.Cells(row, c).Font.Size = 9
        rptWs.Cells(row, c).Font.Color = TABLE_TEXT
        ' Only set background if not already set (e.g., by gap formatting)
        If rptWs.Cells(row, c).Interior.Color = RGB(245, 245, 250) Then
            If isAlt Then
                rptWs.Cells(row, c).Interior.Color = TABLE_ALT_ROW_BG
            Else
                rptWs.Cells(row, c).Interior.Color = TABLE_ROW_BG
            End If
        End If
    Next c

    rng.Borders(xlEdgeBottom).Color = RGB(230, 230, 230)
    rng.Borders(xlEdgeBottom).Weight = xlHairline
End Sub
