Attribute VB_Name = "TISLoader"
'====================================================================
' TIS Loader Module - Rev14 - Simplified TIS Loading Flow
'
' Rev14 CHANGES (from Rev11):
'   - NEW FLOW: LoadNewTIS backs up TIS to TISold (values only),
'     loads new TIS file into TIS sheet, then generates TIScompare
'     by reading from TIS and TISold sheets (no reopening files)
'   - REMOVED: ApplyChangesToWorkingSheet and all related subs
'     (ArchiveRemovedRows, FillDownFormulaColumns, ResortWorkingSheet,
'     DetectFormulaColumns, SaveNewTISSheet, ShellSortDescending).
'     User runs Build Working Sheet separately.
'   - REMOVED: "Apply Changes" button from TIScompare sheet
'   - REMOVED: Named range storing new TIS file path
'   - TISold sheet preserves old TIS data as VALUES ONLY so
'     formulas referencing =TIS!... continue to work after new data
'     is loaded into the TIS sheet object (which stays in place)
'   - Comparison reads from TIS and TISold sheets, not from files
'   - Summary dialog updated to guide user to run Build Working Sheet
'
' Kept from Rev11:
'   - All comparison logic (IdentifyChanges, NormalizeValue, etc.)
'   - CreateTISCompareSheet (minus Apply button)
'   - Milestone header loading, date section, change tracking
'   - LoadFirstTIS for initial TIS load
'====================================================================

Option Explicit

' -- Constants --
Private Const SHEET_TIS As String = "TIS"
Private Const SHEET_TISOLD As String = "TISold"
Private Const SHEET_DEFINITIONS As String = "Definitions"
Private Const SHEET_TISCOMPARE_BASE As String = "TIScompare"
Private Const SHEET_CHANGETRACKING As String = "TIS change tracking"
Private Const SHEET_CEIDS As String = "CEIDs"

Private Const COL_SIRFISHEADERS As String = "A"
Private Const COL_FILTER As String = "B"
Private Const COL_SORT As String = "C"

Private Const DATA_START_ROW As Long = TIS_DATA_START_ROW

Private Const KEY_SITE As String = "Site"
Private Const KEY_ENTITY_CODE As String = "Entity Code"
Private Const KEY_EVENT_TYPE As String = "Event Type"

Private Const IGNORE_COLS_FOR_CHANGES As String = "published|group"
Private Const TRACKING_COLS As String = "ID|Site|Entity Code|Entity Type|CEID"

' Module-level cache
Private m_GroupCache As Object

' Milestone header cache (populated per comparison workflow from Definitions F/G)
Private m_MilestoneStartHeaders As Collection
Private m_MilestoneEndHeaders As Collection

'====================================================================
' MAIN ENTRY POINT -- LoadNewTIS (Rev14: new 3-step flow)
'
' Step 1: Backup existing TIS to TISold (values only)
' Step 2: Load new TIS file into TIS sheet
' Step 3: Generate TIScompare from TIS vs TISold
'====================================================================

Public Sub LoadNewTIS()
    On Error GoTo ErrorHandler

    Dim appSt As AppState
    appSt = SaveAppState()
    SetPerformanceMode

    Set m_GroupCache = Nothing
    Set m_MilestoneStartHeaders = Nothing
    Set m_MilestoneEndHeaders = Nothing

    If CheckTISExists() Then
        ' Step 1: Backup current TIS to TISold (values only)
        BackupTISToTISold

        ' Step 2: Load new TIS file into TIS sheet
        Dim loaded As Boolean
        loaded = LoadNewTISIntoSheet()
        If Not loaded Then GoTo Cleanup

        ' Load Definitions filters once — used by both TIScompare and Working Sheet update
        Dim filters As Object
        Set filters = GetDefinitionsFilters()

        ' Step 3: Generate TIScompare if TISold exists
        If SheetExists(ThisWorkbook, SHEET_TISOLD) Then
            CompareTISWorkflow filters
        End If

        ' Step 4: Update Working Sheet TIS columns in-place
        Application.StatusBar = "Step 4: Updating Working Sheet..."
        UpdateWorkingSheetFromTIS filters
    Else
        Application.ScreenUpdating = True
        MsgBox "No TIS currently loaded. Loading TIS...", vbInformation
        Application.ScreenUpdating = False
        LoadFirstTIS
    End If

    GoTo Cleanup

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    DebugLog "TISLoader ERROR: " & Err.Description

Cleanup:
    Set m_GroupCache = Nothing
    Set m_MilestoneStartHeaders = Nothing
    Set m_MilestoneEndHeaders = Nothing
    RestoreAppState appSt
End Sub

'====================================================================
' BACKUP TIS TO TISold (values only)
' Preserves old TIS data for comparison. TIS sheet object stays
' in place so formulas referencing =TIS!... remain valid.
'====================================================================

Private Sub BackupTISToTISold()
    Dim wsTIS As Worksheet
    Dim wsTISold As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range
    Dim dataValues As Variant

    Set wsTIS = ThisWorkbook.Sheets(SHEET_TIS)
    lastRow = wsTIS.Cells(wsTIS.Rows.Count, 1).End(xlUp).Row
    lastCol = wsTIS.Cells(1, wsTIS.Columns.Count).End(xlToLeft).Column

    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    ' Read all TIS data as values
    dataValues = wsTIS.Range(wsTIS.Cells(1, 1), wsTIS.Cells(lastRow, lastCol)).Value

    If SheetExists(ThisWorkbook, SHEET_TISOLD) Then
        ' Clear existing TISold and paste values
        Set wsTISold = ThisWorkbook.Sheets(SHEET_TISOLD)
        wsTISold.Cells.Clear
    Else
        ' Create TISold sheet
        Set wsTISold = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsTISold.Name = SHEET_TISOLD
    End If

    ' Write values only (no formulas, no formatting)
    wsTISold.Range(wsTISold.Cells(1, 1), wsTISold.Cells(lastRow, lastCol)).Value = dataValues
End Sub

'====================================================================
' LOAD NEW TIS FILE INTO TIS SHEET
' Opens file picker, reads new TIS file, replaces TIS sheet contents.
' Returns True if successful, False if cancelled or failed.
'====================================================================

Private Function LoadNewTISIntoSheet() As Boolean
    Dim fd As FileDialog
    Dim newWB As Workbook
    Dim filePath As String
    Dim wsSource As Worksheet
    Dim wsTIS As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim sourceData As Variant

    LoadNewTISIntoSheet = False

    Application.ScreenUpdating = True
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .filters.Clear
        .filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        .Title = "Select new TIS file"
        If .Show <> -1 Then Exit Function
        filePath = .SelectedItems(1)
    End With
    Application.ScreenUpdating = False

    If Not FileExists(filePath) Then
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Function
    End If

    On Error Resume Next
    Set newWB = Application.Workbooks.Open(filePath, ReadOnly:=True)
    On Error GoTo 0

    If newWB Is Nothing Then
        MsgBox "Could not open file: " & filePath, vbExclamation
        Exit Function
    End If

    If Not SheetExists(newWB, SHEET_TIS) Then
        MsgBox "Selected file does not have a 'TIS' sheet.", vbExclamation
        newWB.Close False
        Exit Function
    End If

    ' Read new TIS data
    Set wsSource = newWB.Sheets(SHEET_TIS)
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

    If lastRow < 1 Or lastCol < 1 Then
        MsgBox "New TIS file appears to be empty.", vbExclamation
        newWB.Close False
        Exit Function
    End If

    sourceData = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol)).Value
    newWB.Close False

    ' Clear existing TIS sheet and write new data
    Set wsTIS = ThisWorkbook.Sheets(SHEET_TIS)
    wsTIS.Cells.Clear
    wsTIS.Range(wsTIS.Cells(1, 1), wsTIS.Cells(lastRow, lastCol)).Value = sourceData

    LoadNewTISIntoSheet = True
End Function

'====================================================================
' CHECK IF TIS SHEET EXISTS
'====================================================================

Private Function CheckTISExists() As Boolean
    Dim ws As Worksheet
    CheckTISExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_TIS Then
            CheckTISExists = True
            Exit Function
        End If
    Next ws
End Function

'====================================================================
' LOAD FIRST TIS FILE (unchanged from v2.0)
'====================================================================

Private Sub LoadFirstTIS()
    Dim fd As FileDialog
    Dim newWB As Workbook
    Dim filePath As String

    Application.ScreenUpdating = True
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .filters.Clear
        .filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        .Title = "Select TIS File to Load"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    Application.ScreenUpdating = False

    If Not FileExists(filePath) Then
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set newWB = Application.Workbooks.Open(filePath, ReadOnly:=True)
    On Error GoTo 0

    If newWB Is Nothing Then
        MsgBox "Could not open file: " & filePath, vbExclamation
        Exit Sub
    End If

    If Not SheetExists(newWB, SHEET_TIS) Then
        MsgBox "Selected file does not have a 'TIS' sheet.", vbExclamation
        newWB.Close False
        Exit Sub
    End If

    On Error Resume Next
    newWB.Sheets(SHEET_TIS).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    On Error GoTo 0

    newWB.Close False

    If CheckTISExists() Then
        Application.ScreenUpdating = True
        MsgBox "New TIS loaded successfully!", vbInformation
        Application.ScreenUpdating = False
    Else
        Application.ScreenUpdating = True
        MsgBox "Failed to copy TIS sheet from file." & vbCrLf & _
               "The file may be protected or corrupted.", vbExclamation
        Application.ScreenUpdating = False
    End If
End Sub

'====================================================================
' COMPARISON WORKFLOW (Rev14: reads from TIS and TISold sheets)
'====================================================================

Private Sub CompareTISWorkflow(filters As Object)
    Dim oldData As Variant, newData As Variant
    Dim oldHeaders As Object, newHeaders As Object
    Dim oldLastRow As Long, newLastRow As Long, oldLastCol As Long, newLastCol As Long
    Dim requiredCols As Object
    Dim siteIdxOld As Long, siteIdxNew As Long
    Dim entityCodeIdxOld As Long, entityCodeIdxNew As Long
    Dim eventTypeIdxOld As Long, eventTypeIdxNew As Long
    Dim changes As Collection
    Dim wsNew As Worksheet, wsOld As Worksheet
    Dim compareSheetName As String
    Dim hasCEIDsSheet As Boolean

    If Not SheetExists(ThisWorkbook, SHEET_DEFINITIONS) Then
        MsgBox "Definitions sheet not found.", vbExclamation
        Exit Sub
    End If

    If Not SheetExists(ThisWorkbook, SHEET_TISOLD) Then
        MsgBox "TISold sheet not found. Cannot compare.", vbExclamation
        Exit Sub
    End If

    hasCEIDsSheet = SheetExists(ThisWorkbook, SHEET_CEIDS)
    If Not hasCEIDsSheet Then
        MsgBox "CEIDs sheet not found. Group column will be empty.", vbInformation
    End If

    Set m_GroupCache = CreateObject("Scripting.Dictionary")

    Set requiredCols = GetSirfisHeaders()
    If requiredCols.Count = 0 Then
        MsgBox "No sirfisheaders defined in Definitions sheet.", vbExclamation
        Exit Sub
    End If

    ' Initialize milestone headers from Definitions F/G for project start/end detection
    LoadMilestoneHeaders

    ' Read old data from TISold sheet
    Set wsOld = ThisWorkbook.Sheets(SHEET_TISOLD)
    oldLastRow = wsOld.Cells(wsOld.Rows.Count, 1).End(xlUp).Row
    oldLastCol = wsOld.Cells(1, wsOld.Columns.Count).End(xlToLeft).Column
    oldData = wsOld.Range(wsOld.Cells(1, 1), wsOld.Cells(oldLastRow, oldLastCol)).Value

    ' Read new data from TIS sheet (already loaded in Step 2)
    Set wsNew = ThisWorkbook.Sheets(SHEET_TIS)
    newLastRow = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
    newLastCol = wsNew.Cells(1, wsNew.Columns.Count).End(xlToLeft).Column
    newData = wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(newLastRow, newLastCol)).Value

    Set oldHeaders = MapHeaders(oldData)
    Set newHeaders = MapHeaders(newData)

    If Not ValidateRequiredHeaders(oldHeaders, newHeaders, requiredCols, _
                                   siteIdxOld, entityCodeIdxOld, eventTypeIdxOld, _
                                   siteIdxNew, entityCodeIdxNew, eventTypeIdxNew) Then
        Exit Sub
    End If

    ValidateDateFormats newData, newHeaders, requiredCols, "New TIS"
    ValidateDateFormats oldData, oldHeaders, requiredCols, "Old TIS"

    Set changes = IdentifyChanges(oldData, newData, oldHeaders, newHeaders, requiredCols, _
                                   siteIdxOld, entityCodeIdxOld, eventTypeIdxOld, _
                                   siteIdxNew, entityCodeIdxNew, eventTypeIdxNew, _
                                   filters)

    ' Always replace the single TIScompare sheet rather than accumulating TIScompare1, TIScompare2...
    compareSheetName = SHEET_TISCOMPARE_BASE
    Application.DisplayAlerts = False
    If SheetExists(ThisWorkbook, compareSheetName) Then
        ThisWorkbook.Sheets(compareSheetName).Delete
    End If
    Application.DisplayAlerts = True

    CreateTISCompareSheet compareSheetName, newData, oldData, newHeaders, oldHeaders, _
                          changes, requiredCols, hasCEIDsSheet

    ' CreateChangeTrackingLog removed in Rev14: UpdateWorkingSheetFromTIS now writes
    ' orange fills + appended comments directly in the Working Sheet, making the
    ' separate "TIS change tracking" sheet redundant.

End Sub

'====================================================================
' Helper functions carried forward from Rev11
'====================================================================

Private Sub ValidateDateFormats(dataArray As Variant, headers As Object, _
                                 requiredCols As Object, sourceName As String)
    Dim keyVar As Variant
    Dim colName As String
    Dim colIdx As Long
    Dim i As Long, lastRow As Long
    Dim textDateCount As Long
    Dim realDateCount As Long
    Dim totalTextDates As Long
    Dim problemCols As String
    Dim cellVal As Variant

    lastRow = UBound(dataArray, 1)
    totalTextDates = 0
    problemCols = ""

    For Each keyVar In requiredCols.Keys
        colName = requiredCols(keyVar)
        If Not headers.Exists(LCase(colName)) Then GoTo NextDateCol
        colIdx = headers(LCase(colName))
        If colIdx < 1 Or colIdx > UBound(dataArray, 2) Then GoTo NextDateCol
        textDateCount = 0
        realDateCount = 0
        For i = 2 To Application.WorksheetFunction.Min(lastRow, 51)
            cellVal = dataArray(i, colIdx)
            If Not IsEmpty(cellVal) And CStr(cellVal) <> "" Then
                If IsDate(cellVal) And VarType(cellVal) = vbDate Then
                    realDateCount = realDateCount + 1
                ElseIf IsDate(CStr(cellVal)) And VarType(cellVal) = vbString Then
                    textDateCount = textDateCount + 1
                End If
            End If
        Next i
        If textDateCount > 0 And (realDateCount > 0 Or textDateCount >= 3) Then
            totalTextDates = totalTextDates + textDateCount
            If problemCols <> "" Then problemCols = problemCols & ", "
            problemCols = problemCols & colName & " (" & textDateCount & " text"
            If realDateCount > 0 Then problemCols = problemCols & " / " & realDateCount & " real"
            problemCols = problemCols & ")"
        End If
NextDateCol:
    Next keyVar
    If totalTextDates > 0 Then
        MsgBox "Date Format Warning (" & sourceName & "):" & vbCrLf & vbCrLf & _
               "Some date columns contain text strings instead of real dates. " & _
               "This may cause incorrect date comparisons." & vbCrLf & vbCrLf & _
               "Affected columns:" & vbCrLf & problemCols & vbCrLf & vbCrLf & _
               "Tip: Select the column(s) in Excel and format as Date, or use " & _
               "Text to Columns to convert.", vbExclamation, "Date Format Warning"
    End If
End Sub

Private Function GetSirfisHeaders() As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim headerName As String
    Set dict = CreateObject("Scripting.Dictionary")
    If Not SheetExists(ThisWorkbook, SHEET_DEFINITIONS) Then
        Set GetSirfisHeaders = dict: Exit Function
    End If
    Set ws = ThisWorkbook.Sheets(SHEET_DEFINITIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        headerName = Trim(CStr(ws.Cells(i, 1).Value))
        If headerName <> "" Then dict(LCase(headerName)) = headerName
    Next i
    Set GetSirfisHeaders = dict
End Function

Private Function GetDefinitionsFilters() As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim headerName As String
    Dim filterValue As String
    Dim filterValues() As String
    Set dict = CreateObject("Scripting.Dictionary")
    If Not SheetExists(ThisWorkbook, SHEET_DEFINITIONS) Then
        Set GetDefinitionsFilters = dict: Exit Function
    End If
    Set ws = ThisWorkbook.Sheets(SHEET_DEFINITIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        headerName = Trim(CStr(ws.Cells(i, 1).Value))
        filterValue = Trim(CStr(ws.Cells(i, 2).Value))
        If headerName <> "" And filterValue <> "" Then
            filterValues = Split(filterValue, " ")
            dict(LCase(headerName)) = filterValues
        End If
    Next i
    Set GetDefinitionsFilters = dict
End Function

Private Function RowMatchesFilters(dataArray As Variant, rowIndex As Long, _
                                    headers As Object, filters As Object) As Boolean
    Dim filterKey As Variant
    Dim filterValues As Variant
    Dim cellValue As String
    Dim colIdx As Long
    Dim matchFound As Boolean
    Dim j As Long
    If filters.Count = 0 Then RowMatchesFilters = True: Exit Function
    For Each filterKey In filters.Keys
        If Not headers.Exists(filterKey) Then GoTo NextFilter
        colIdx = headers(filterKey)
        If colIdx < 1 Or colIdx > UBound(dataArray, 2) Then GoTo NextFilter
        cellValue = LCase(Trim(CStr(dataArray(rowIndex, colIdx))))
        filterValues = filters(filterKey)
        matchFound = False
        For j = LBound(filterValues) To UBound(filterValues)
            If cellValue = LCase(Trim(CStr(filterValues(j)))) Then matchFound = True: Exit For
        Next j
        If Not matchFound Then RowMatchesFilters = False: Exit Function
NextFilter:
    Next filterKey
    RowMatchesFilters = True
End Function

Private Function BuildFilterDescription(filters As Object) As String
    Dim filterKey As Variant
    Dim filterValues As Variant
    Dim desc As String
    Dim j As Long
    Dim headerName As String
    If filters.Count = 0 Then
        BuildFilterDescription = "No filters applied - all projects will be included.": Exit Function
    End If
    desc = "Filters to apply (AND logic):" & vbCrLf & vbCrLf
    For Each filterKey In filters.Keys
        filterValues = filters(filterKey)
        headerName = UCase(Left(filterKey, 1)) & Mid(filterKey, 2)
        desc = desc & "  " & Chr(149) & " " & headerName & ": "
        For j = LBound(filterValues) To UBound(filterValues)
            If j > LBound(filterValues) Then desc = desc & " OR "
            desc = desc & Chr(34) & filterValues(j) & Chr(34)
        Next j
        desc = desc & vbCrLf
    Next filterKey
    BuildFilterDescription = desc
End Function

Private Function ValidateRequiredHeaders(oldHeaders As Object, newHeaders As Object, _
                                         requiredCols As Object, _
                                         ByRef siteIdxOld As Long, ByRef entityCodeIdxOld As Long, ByRef eventTypeIdxOld As Long, _
                                         ByRef siteIdxNew As Long, ByRef entityCodeIdxNew As Long, ByRef eventTypeIdxNew As Long) As Boolean
    Dim availableCols() As String
    Dim colCount As Long
    Dim key As Variant
    Dim selectedCol As String
    ValidateRequiredHeaders = True
    If Not newHeaders.Exists(LCase(KEY_SITE)) Then
        ReDim availableCols(1 To newHeaders.Count): colCount = 0
        For Each key In newHeaders.Keys: colCount = colCount + 1: availableCols(colCount) = CStr(key): Next key
        ReDim Preserve availableCols(1 To colCount)
        selectedCol = ShowColumnSelectionDialog("Site column not found in new TIS. Select Site column:", availableCols)
        If selectedCol = "" Then ValidateRequiredHeaders = False: Exit Function
        newHeaders.Add LCase(KEY_SITE), newHeaders(LCase(selectedCol))
    End If
    If Not newHeaders.Exists(LCase(KEY_ENTITY_CODE)) Then
        ReDim availableCols(1 To newHeaders.Count): colCount = 0
        For Each key In newHeaders.Keys: colCount = colCount + 1: availableCols(colCount) = CStr(key): Next key
        ReDim Preserve availableCols(1 To colCount)
        selectedCol = ShowColumnSelectionDialog("Entity Code column not found in new TIS. Select Entity Code column:", availableCols)
        If selectedCol = "" Then ValidateRequiredHeaders = False: Exit Function
        newHeaders.Add LCase(KEY_ENTITY_CODE), newHeaders(LCase(selectedCol))
    End If
    If Not newHeaders.Exists(LCase(KEY_EVENT_TYPE)) Then
        ReDim availableCols(1 To newHeaders.Count): colCount = 0
        For Each key In newHeaders.Keys: colCount = colCount + 1: availableCols(colCount) = CStr(key): Next key
        ReDim Preserve availableCols(1 To colCount)
        selectedCol = ShowColumnSelectionDialog("Event Type column not found in new TIS. Select Event Type column:", availableCols)
        If selectedCol = "" Then ValidateRequiredHeaders = False: Exit Function
        newHeaders.Add LCase(KEY_EVENT_TYPE), newHeaders(LCase(selectedCol))
    End If
    If Not oldHeaders.Exists(LCase(KEY_SITE)) Or Not oldHeaders.Exists(LCase(KEY_ENTITY_CODE)) Or Not oldHeaders.Exists(LCase(KEY_EVENT_TYPE)) Then
        MsgBox "Required columns (Site, Entity Code, Event Type) not found in old TIS data.", vbExclamation
        ValidateRequiredHeaders = False: Exit Function
    End If
    siteIdxOld = oldHeaders(LCase(KEY_SITE)): entityCodeIdxOld = oldHeaders(LCase(KEY_ENTITY_CODE)): eventTypeIdxOld = oldHeaders(LCase(KEY_EVENT_TYPE))
    siteIdxNew = newHeaders(LCase(KEY_SITE)): entityCodeIdxNew = newHeaders(LCase(KEY_ENTITY_CODE)): eventTypeIdxNew = newHeaders(LCase(KEY_EVENT_TYPE))
End Function

Private Function ShowColumnSelectionDialog(prompt As String, availableCols() As String) As String
    Dim i As Long
    Dim response As String
    On Error Resume Next
    response = InputBox(prompt & vbCrLf & vbCrLf & "Available columns: " & Join(availableCols, ", "), "Select Column")
    On Error GoTo 0
    If response = "" Then ShowColumnSelectionDialog = "": Exit Function
    For i = LBound(availableCols) To UBound(availableCols)
        If LCase(availableCols(i)) = LCase(response) Then ShowColumnSelectionDialog = response: Exit Function
    Next i
    MsgBox "Column not found in list. Please try again.", vbExclamation
    ShowColumnSelectionDialog = ""
End Function

Private Function NormalizeValue(val As Variant) As String
    Dim strVal As String
    Dim testDate As Date
    ' Rev11 FIX: for real date values, use serial number to avoid locale issues
    If IsDate(val) And Not IsEmpty(val) Then
        If VarType(val) = vbDate Then
            NormalizeValue = CStr(CLng(CDate(val))): Exit Function
        End If
        NormalizeValue = CStr(CLng(CDate(val))): Exit Function
    End If
    strVal = Trim(CStr(val))
    ' For string dates, still try to parse -- but use serial for comparison
    If IsDate(strVal) And Len(strVal) > 4 Then
        On Error Resume Next
        testDate = CDate(strVal)
        If Err.Number = 0 Then NormalizeValue = CStr(CLng(testDate)): On Error GoTo 0: Exit Function
        On Error GoTo 0
    End If
    NormalizeValue = LCase(strVal)
End Function

Private Function ShouldIgnoreColumnForChanges(colName As String) As Boolean
    Dim ignoreCols() As String
    Dim i As Long
    Dim lowerColName As String
    lowerColName = LCase(Trim(colName))
    ignoreCols = Split(IGNORE_COLS_FOR_CHANGES, "|")
    For i = LBound(ignoreCols) To UBound(ignoreCols)
        If lowerColName = LCase(Trim(ignoreCols(i))) Then ShouldIgnoreColumnForChanges = True: Exit Function
    Next i
    ShouldIgnoreColumnForChanges = False
End Function

'====================================================================
' LOAD MILESTONE HEADERS from Definitions columns F/G
'====================================================================

Private Sub LoadMilestoneHeaders()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim fText As String, headerName As String
    Dim tokens As Variant, tokenVal As Variant
    Dim letter As String, num As Long
    Dim milGroups As Object
    Dim groupKey As Variant

    Set m_MilestoneStartHeaders = New Collection
    Set m_MilestoneEndHeaders = New Collection

    If Not SheetExists(ThisWorkbook, SHEET_DEFINITIONS) Then Exit Sub
    Set ws = ThisWorkbook.Sheets(SHEET_DEFINITIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Set milGroups = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        fText = Trim(CStr(ws.Cells(i, 6).Value))
        headerName = Trim(CStr(ws.Cells(i, 1).Value))
        If fText <> "" And headerName <> "" Then
            tokens = Split(fText, "|")
            For Each tokenVal In tokens
                tokenVal = UCase(Trim(CStr(tokenVal)))
                If Len(CStr(tokenVal)) >= 2 And IsNumeric(Mid(CStr(tokenVal), 2)) Then
                    letter = Left(CStr(tokenVal), 1)
                    num = CLng(Mid(CStr(tokenVal), 2))
                    If Not milGroups.Exists(letter) Then
                        Set milGroups(letter) = CreateObject("Scripting.Dictionary")
                    End If
                    milGroups(letter)(num) = headerName
                End If
            Next tokenVal
        End If
    Next i

    ' Collect position-1 headers as starts, position-2 as ends
    For Each groupKey In milGroups.Keys
        If milGroups(groupKey).Exists(1) Then
            m_MilestoneStartHeaders.Add milGroups(groupKey)(1)
        End If
        If milGroups(groupKey).Exists(2) Then
            m_MilestoneEndHeaders.Add milGroups(groupKey)(2)
        End If
    Next groupKey
End Sub

'====================================================================
' GET PROJECT START DATE
'====================================================================

Private Function GetProjectStartDate(dataArray As Variant, rowIndex As Long, _
                                      headers As Object) As Variant
    Dim colName As String
    Dim colIdx As Long
    Dim cellValue As Variant
    Dim i As Long
    Dim minDate As Date
    Dim foundDate As Boolean
    Dim testDate As Date

    If m_MilestoneStartHeaders Is Nothing Then LoadMilestoneHeaders
    If m_MilestoneStartHeaders.Count = 0 Then GetProjectStartDate = Empty: Exit Function

    foundDate = False
    minDate = DateSerial(2099, 12, 31)

    For i = 1 To m_MilestoneStartHeaders.Count
        colName = m_MilestoneStartHeaders(i)
        If headers.Exists(LCase(colName)) Then
            colIdx = headers(LCase(colName))
            If colIdx > 0 And colIdx <= UBound(dataArray, 2) Then
                cellValue = dataArray(rowIndex, colIdx)
                If IsDate(cellValue) Then
                    testDate = CDate(cellValue)
                    If testDate < minDate Then minDate = testDate: foundDate = True
                ElseIf VarType(cellValue) = vbString And IsDate(CStr(cellValue)) Then
                    testDate = CDate(CStr(cellValue))
                    If testDate < minDate Then minDate = testDate: foundDate = True
                End If
            End If
        End If
    Next i

    If foundDate Then GetProjectStartDate = minDate Else GetProjectStartDate = Empty
End Function

Private Function CalcDateChangeWeeks(oldDate As Variant, newDate As Variant) As Variant
    If IsEmpty(oldDate) Or IsEmpty(newDate) Then CalcDateChangeWeeks = "": Exit Function
    If Not IsDate(oldDate) Or Not IsDate(newDate) Then CalcDateChangeWeeks = "": Exit Function
    CalcDateChangeWeeks = Round((CDate(newDate) - CDate(oldDate)) / 7, 1)
End Function

'====================================================================
' GET PROJECT END DATE
'====================================================================

Private Function GetProjectEndDate(dataArray As Variant, rowIndex As Long, headers As Object) As Variant
    Dim colName As String
    Dim colIdx As Long
    Dim cellValue As Variant
    Dim i As Long
    Dim maxDate As Date
    Dim foundDate As Boolean
    Dim testDate As Date

    If m_MilestoneEndHeaders Is Nothing Then LoadMilestoneHeaders
    If m_MilestoneEndHeaders.Count = 0 Then GetProjectEndDate = Empty: Exit Function

    foundDate = False
    maxDate = DateSerial(1900, 1, 1)

    For i = 1 To m_MilestoneEndHeaders.Count
        colName = m_MilestoneEndHeaders(i)
        If headers.Exists(LCase(colName)) Then
            colIdx = headers(LCase(colName))
            If colIdx > 0 And colIdx <= UBound(dataArray, 2) Then
                cellValue = dataArray(rowIndex, colIdx)
                If IsDate(cellValue) Then
                    testDate = CDate(cellValue)
                    If testDate > maxDate Then maxDate = testDate: foundDate = True
                ElseIf VarType(cellValue) = vbString And IsDate(CStr(cellValue)) Then
                    testDate = CDate(CStr(cellValue))
                    If testDate > maxDate Then maxDate = testDate: foundDate = True
                End If
            End If
        End If
    Next i

    If foundDate Then GetProjectEndDate = maxDate Else GetProjectEndDate = Empty
End Function

Private Function LookupGroup(entityType As String, hasCEIDsSheet As Boolean) As String
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim lookupVal As String
    LookupGroup = ""
    If Not hasCEIDsSheet Then Exit Function
    If Trim(entityType) = "" Then Exit Function
    lookupVal = LCase(Trim(entityType))
    If Not m_GroupCache Is Nothing Then
        If m_GroupCache.Exists(lookupVal) Then LookupGroup = m_GroupCache(lookupVal): Exit Function
    End If
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CEIDS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastRow
        If LCase(Trim(CStr(ws.Cells(i, 1).Value))) = lookupVal Then
            LookupGroup = CStr(ws.Cells(i, 2).Value)
            If Not m_GroupCache Is Nothing Then m_GroupCache(lookupVal) = LookupGroup
            Exit Function
        End If
    Next i
    If Not m_GroupCache Is Nothing Then m_GroupCache(lookupVal) = ""
End Function

Private Function MapHeaders(dataArray As Variant) As Object
    Dim dict As Object
    Dim j As Long, lastCol As Long
    Dim header As String
    Set dict = CreateObject("Scripting.Dictionary")
    lastCol = UBound(dataArray, 2)
    For j = 1 To lastCol
        header = LCase(Trim(CStr(dataArray(1, j))))
        If header <> "" Then dict(header) = j
    Next j
    Set MapHeaders = dict
End Function

Private Function IdentifyChanges(oldData As Variant, newData As Variant, _
                                  oldHeaders As Object, newHeaders As Object, _
                                  requiredCols As Object, _
                                  siteIdxOld As Long, entityCodeIdxOld As Long, eventTypeIdxOld As Long, _
                                  siteIdxNew As Long, entityCodeIdxNew As Long, eventTypeIdxNew As Long, _
                                  filters As Object) As Collection
    Dim changes As Collection
    Dim oldDict As Object, newDict As Object, oldFilteredDict As Object
    Dim oldLastRow As Long, newLastRow As Long
    Dim i As Long
    Dim key As String, keyVar As Variant
    Dim changeRecord As Object
    Dim isDuplicate2 As Boolean
    Dim changedCols As Collection
    Dim colName As String
    Dim newVal As String, oldVal As String
    Dim oldStartDate As Variant, newStartDate As Variant
    Dim oldEndDate As Variant, newEndDate As Variant
    Dim startChangeWeeks As Variant, endChangeWeeks As Variant

    Set changes = New Collection
    Set oldDict = CreateObject("Scripting.Dictionary")
    Set newDict = CreateObject("Scripting.Dictionary")
    Set oldFilteredDict = CreateObject("Scripting.Dictionary")
    oldLastRow = UBound(oldData, 1)
    newLastRow = UBound(newData, 1)

    For i = 2 To oldLastRow
        If Not RowMatchesFilters(oldData, i, oldHeaders, filters) Then GoTo NextOldRow
        If oldData(i, siteIdxOld) <> "" And oldData(i, entityCodeIdxOld) <> "" And oldData(i, eventTypeIdxOld) <> "" Then
            key = LCase(CStr(oldData(i, siteIdxOld))) & "|" & LCase(CStr(oldData(i, entityCodeIdxOld))) & "|" & LCase(CStr(oldData(i, eventTypeIdxOld)))
            If Not oldDict.Exists(key) Then
                oldDict(key) = i
                oldFilteredDict(key) = i
            End If
        End If
NextOldRow:
    Next i

    For i = 2 To newLastRow
        If Not RowMatchesFilters(newData, i, newHeaders, filters) Then GoTo NextNewRow
        If newData(i, siteIdxNew) <> "" And newData(i, entityCodeIdxNew) <> "" And newData(i, eventTypeIdxNew) <> "" Then
            key = LCase(CStr(newData(i, siteIdxNew))) & "|" & LCase(CStr(newData(i, entityCodeIdxNew))) & "|" & LCase(CStr(newData(i, eventTypeIdxNew)))
            isDuplicate2 = newDict.Exists(key)
            If Not newDict.Exists(key) Then newDict(key) = i
            Set changeRecord = CreateObject("Scripting.Dictionary")
            changeRecord("Key") = key
            changeRecord("NewRow") = i
            changeRecord("IsDuplicate") = isDuplicate2
            changeRecord("StartChange") = ""
            changeRecord("EndChange") = ""
            changeRecord("OldRow") = 0
            If Not oldDict.Exists(key) Then
                changeRecord("Status") = "Added"
                Set changeRecord("ChangedCols") = Nothing
            Else
                changeRecord("OldRow") = oldDict(key)
                Set changedCols = New Collection
                For Each keyVar In requiredCols.Keys
                    colName = requiredCols(keyVar)
                    If ShouldIgnoreColumnForChanges(colName) Then GoTo NextCol
                    If newHeaders.Exists(LCase(colName)) And oldHeaders.Exists(LCase(colName)) Then
                        newVal = NormalizeValue(newData(i, newHeaders(LCase(colName))))
                        oldVal = NormalizeValue(oldData(oldDict(key), oldHeaders(LCase(colName))))
                        If newVal <> oldVal Then changedCols.Add newHeaders(LCase(colName))
                    End If
NextCol:
                Next keyVar
                If changedCols.Count > 0 Then
                    changeRecord("Status") = "Modified"
                    Set changeRecord("ChangedCols") = changedCols
                    oldStartDate = GetProjectStartDate(oldData, oldDict(key), oldHeaders)
                    newStartDate = GetProjectStartDate(newData, i, newHeaders)
                    changeRecord("StartChange") = CalcDateChangeWeeks(oldStartDate, newStartDate)
                    oldEndDate = GetProjectEndDate(oldData, oldDict(key), oldHeaders)
                    newEndDate = GetProjectEndDate(newData, i, newHeaders)
                    changeRecord("EndChange") = CalcDateChangeWeeks(oldEndDate, newEndDate)
                Else
                    changeRecord("Status") = "Unchanged"
                    Set changeRecord("ChangedCols") = Nothing
                End If
            End If
            changes.Add changeRecord
        End If
NextNewRow:
    Next i

    For Each keyVar In oldFilteredDict.Keys
        key = CStr(keyVar)
        If Not newDict.Exists(key) Then
            Set changeRecord = CreateObject("Scripting.Dictionary")
            changeRecord("Key") = key
            changeRecord("OldRow") = oldFilteredDict(key)
            changeRecord("NewRow") = 0
            changeRecord("Status") = "Removed"
            changeRecord("IsDuplicate") = False
            changeRecord("StartChange") = ""
            changeRecord("EndChange") = ""
            Set changeRecord("ChangedCols") = Nothing
            changes.Add changeRecord
        End If
    Next keyVar

    Set IdentifyChanges = changes
End Function

Private Sub SortChangesByDate(changes As Collection, newColIndices() As Long, oldColIndices() As Long, _
                               colCount As Long, newData As Variant, oldData As Variant)
    Dim i As Long, j As Long
    Dim sortIndices() As Long
    Dim earliestDates() As Date
    Dim gap As Long
    Dim tempIdx As Long
    Dim n As Long
    Dim tempCollection As Collection
    Dim k As Long

    n = changes.Count
    If n <= 1 Then Exit Sub
    ReDim sortIndices(1 To n)
    ReDim earliestDates(1 To n)
    For i = 1 To n
        sortIndices(i) = i
        earliestDates(i) = GetEarliestDate(changes(i), newColIndices, oldColIndices, colCount, newData, oldData)
    Next i
    gap = n \ 2
    Do While gap > 0
        For i = gap + 1 To n
            tempIdx = sortIndices(i)
            j = i
            Do While j > gap
                If earliestDates(tempIdx) < earliestDates(sortIndices(j - gap)) Then
                    sortIndices(j) = sortIndices(j - gap)
                    j = j - gap
                Else
                    Exit Do
                End If
            Loop
            sortIndices(j) = tempIdx
        Next i
        gap = gap \ 2
    Loop
    Set tempCollection = New Collection
    For k = 1 To n: tempCollection.Add changes(sortIndices(k)): Next k
    For k = n To 1 Step -1: changes.Remove k: Next k
    For k = 1 To tempCollection.Count: changes.Add tempCollection(k): Next k
End Sub

Private Function GetEarliestDate(changeRecord As Object, newColIndices() As Long, oldColIndices() As Long, _
                                  colCount As Long, newData As Variant, oldData As Variant) As Date
    Dim sourceRow As Long
    Dim minDate As Date
    Dim j As Long
    Dim cellValue As Variant
    Dim cellDate As Date
    Dim colIdx As Long
    minDate = DateSerial(2099, 12, 31)
    If changeRecord("Status") = "Removed" Then
        sourceRow = changeRecord("OldRow")
        For j = 1 To colCount
            colIdx = oldColIndices(j)
            If colIdx > 0 And colIdx <= UBound(oldData, 2) Then
                cellValue = oldData(sourceRow, colIdx)
                If IsDate(cellValue) Then cellDate = CDate(cellValue): If cellDate < minDate Then minDate = cellDate
            End If
        Next j
    Else
        sourceRow = changeRecord("NewRow")
        For j = 1 To colCount
            colIdx = newColIndices(j)
            If colIdx > 0 And colIdx <= UBound(newData, 2) Then
                cellValue = newData(sourceRow, colIdx)
                If IsDate(cellValue) Then cellDate = CDate(cellValue): If cellDate < minDate Then minDate = cellDate
            End If
        Next j
    End If
    GetEarliestDate = minDate
End Function

Private Function ColumnInChangedList(colIndex As Long, changedCols As Collection) As Boolean
    Dim i As Long
    ColumnInChangedList = False
    For i = 1 To changedCols.Count
        If CLng(changedCols(i)) = colIndex Then ColumnInChangedList = True: Exit Function
    Next i
End Function

'====================================================================
' CREATE TIS COMPARE SHEET
' Rev14: Removed Apply/Skip column, Apply button, and TIS path named range.
' Layout: Status(1), StartChg(2), EndChg(3), data(4+)
'====================================================================

Private Sub CreateTISCompareSheet(sheetName As String, newData As Variant, oldData As Variant, _
                                   newHeaders As Object, oldHeaders As Object, _
                                   changes As Collection, requiredCols As Object, _
                                   hasCEIDsSheet As Boolean)
    Dim ws As Worksheet
    Dim i As Long, j As Long, outCol As Long
    Dim changeRecord As Object
    Dim colNames() As String, newColIndices() As Long, oldColIndices() As Long
    Dim colCount As Long, totalCols As Long
    Dim keyVar As Variant
    Dim colName As String
    Dim sourceRow As Long, colIdx As Long
    Dim ceidPos As Long, groupInsertPos As Long
    Dim entityTypePos As Long
    Dim entityTypeOutCol As Long
    Dim outputArray() As Variant
    Dim formulaCells As Collection
    Dim formulaInfo As Object
    Dim outputRow As Long
    Dim fi As Variant
    Dim statusVal As String
    Dim commentOutCol As Long, oldValue As Variant, oldColIdx As Long
    Dim dataStartCol As Long

    If changes.Count = 0 Then
        MsgBox "No changes detected between TIS files.", vbInformation
        Exit Sub
    End If

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName

    ' Dark SaaS theme base setup
    ws.Cells.Interior.Color = 3022366       ' THEME_BG
    ws.Cells.Font.Name = "Segoe UI"         ' THEME_FONT
    ws.Cells.Font.Color = 15788258           ' THEME_TEXT
    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ' Build column mapping
    colCount = requiredCols.Count
    ReDim colNames(1 To colCount)
    ReDim newColIndices(1 To colCount)
    ReDim oldColIndices(1 To colCount)

    i = 0
    ceidPos = 0
    entityTypePos = 0
    For Each keyVar In requiredCols.Keys
        colName = requiredCols(keyVar)
        i = i + 1
        colNames(i) = colName
        If LCase(colName) = "ceid" Then ceidPos = i
        If LCase(colName) = "entity type" Then entityTypePos = i
        If newHeaders.Exists(LCase(colName)) Then newColIndices(i) = newHeaders(LCase(colName)) Else newColIndices(i) = 0
        If oldHeaders.Exists(LCase(colName)) Then oldColIndices(i) = oldHeaders(LCase(colName)) Else oldColIndices(i) = 0
    Next keyVar

    If ceidPos > 0 Then groupInsertPos = ceidPos + 1 Else groupInsertPos = 11
    If groupInsertPos > colCount Then groupInsertPos = colCount + 1

    ' Rev14: totalCols = Status + StartChange + EndChange + colCount + Group (no Action column)
    totalCols = colCount + 4

    ' Calculate Entity Type output column
    If entityTypePos > 0 Then
        entityTypeOutCol = 3 + entityTypePos
        If groupInsertPos <= entityTypePos Then entityTypeOutCol = entityTypeOutCol + 1
    Else
        entityTypeOutCol = 0
    End If

    SortChangesByDate changes, newColIndices, oldColIndices, colCount, newData, oldData

    ' Build output array
    ReDim outputArray(1 To changes.Count + 1, 1 To totalCols)
    Set formulaCells = New Collection

    ' Header row
    outCol = 1
    outputArray(1, outCol) = "Status": outCol = outCol + 1
    outputArray(1, outCol) = "Project Start Change (weeks)": outCol = outCol + 1
    outputArray(1, outCol) = "Project End Change (weeks)": outCol = outCol + 1

    For j = 1 To colCount
        If j = groupInsertPos Then outputArray(1, outCol) = "Group": outCol = outCol + 1
        outputArray(1, outCol) = colNames(j): outCol = outCol + 1
    Next j
    If groupInsertPos > colCount Then outputArray(1, outCol) = "Group": outCol = outCol + 1

    ' Data rows
    For i = 1 To changes.Count
        Set changeRecord = changes(i)
        outputRow = i + 1
        outCol = 1

        ' Status
        statusVal = changeRecord("Status")
        outputArray(outputRow, outCol) = statusVal: outCol = outCol + 1

        ' Start/End change
        If statusVal = "Modified" Then
            outputArray(outputRow, outCol) = changeRecord("StartChange")
        Else
            outputArray(outputRow, outCol) = ""
        End If
        outCol = outCol + 1

        If statusVal = "Modified" Then
            outputArray(outputRow, outCol) = changeRecord("EndChange")
        Else
            outputArray(outputRow, outCol) = ""
        End If
        outCol = outCol + 1

        ' Source data
        If statusVal = "Removed" Then sourceRow = changeRecord("OldRow") Else sourceRow = changeRecord("NewRow")

        For j = 1 To colCount
            If j = groupInsertPos Then
                If hasCEIDsSheet And entityTypeOutCol > 0 Then
                    outputArray(outputRow, outCol) = ""
                    Set formulaInfo = CreateObject("Scripting.Dictionary")
                    formulaInfo("Row") = outputRow
                    formulaInfo("Col") = outCol
                    formulaInfo("ETCol") = entityTypeOutCol
                    formulaCells.Add formulaInfo
                Else
                    outputArray(outputRow, outCol) = ""
                End If
                outCol = outCol + 1
            End If

            If statusVal = "Removed" Then
                colIdx = oldColIndices(j)
                If colIdx > 0 And colIdx <= UBound(oldData, 2) Then outputArray(outputRow, outCol) = oldData(sourceRow, colIdx) Else outputArray(outputRow, outCol) = ""
            Else
                colIdx = newColIndices(j)
                If colIdx > 0 And colIdx <= UBound(newData, 2) Then outputArray(outputRow, outCol) = newData(sourceRow, colIdx) Else outputArray(outputRow, outCol) = ""
            End If
            outCol = outCol + 1
        Next j

        If groupInsertPos > colCount Then
            If hasCEIDsSheet And entityTypeOutCol > 0 Then
                outputArray(outputRow, outCol) = ""
                Set formulaInfo = CreateObject("Scripting.Dictionary")
                formulaInfo("Row") = outputRow
                formulaInfo("Col") = outCol
                formulaInfo("ETCol") = entityTypeOutCol
                formulaCells.Add formulaInfo
            Else
                outputArray(outputRow, outCol) = ""
            End If
            outCol = outCol + 1
        End If
    Next i

    ' Batch write
    ws.Range(ws.Cells(1, 1), ws.Cells(changes.Count + 1, totalCols)).Value = outputArray

    ' -- Add cell comments showing old values for ALL changed cells --
    dataStartCol = 4
    For i = 1 To changes.Count
        Set changeRecord = changes(i)
        If changeRecord("Status") = "Modified" Then
            If Not changeRecord("ChangedCols") Is Nothing Then
                commentOutCol = dataStartCol
                For j = 1 To colCount
                    If j = groupInsertPos Then commentOutCol = commentOutCol + 1

                    If ColumnInChangedList(newColIndices(j), changeRecord("ChangedCols")) Then
                        oldColIdx = oldColIndices(j)
                        If oldColIdx > 0 And changeRecord("OldRow") > 0 Then
                            If oldColIdx <= UBound(oldData, 2) Then
                                oldValue = oldData(changeRecord("OldRow"), oldColIdx)
                                ' Format dates for readability
                                If IsDate(oldValue) And Not IsEmpty(oldValue) Then
                                    oldValue = Format(CDate(oldValue), "mm/dd/yyyy")
                                End If
                                On Error Resume Next
                                ws.Cells(i + 1, commentOutCol).ClearComments
                                If IsEmpty(oldValue) Then oldValue = ""
                                ws.Cells(i + 1, commentOutCol).AddComment "Was: " & CStr(oldValue)
                                If Not ws.Cells(i + 1, commentOutCol).Comment Is Nothing Then
                                    ws.Cells(i + 1, commentOutCol).Comment.Shape.TextFrame.AutoSize = True
                                End If
                                On Error GoTo 0
                            End If
                        End If
                    End If

                    commentOutCol = commentOutCol + 1
                Next j
            End If
        End If
    Next i

    ' Group VLOOKUP formulas
    For Each fi In formulaCells
        Set formulaInfo = fi
        ws.Cells(formulaInfo("Row"), formulaInfo("Col")).Formula = "=IFERROR(VLOOKUP(" & _
            ws.Cells(formulaInfo("Row"), formulaInfo("ETCol")).Address(False, False) & _
            ",CEIDs!$A:$B,2,FALSE),"""")"
    Next fi
    ws.Calculate

    ' Apply formatting (Rev14: status is now col 1, data starts at col 4)
    ApplyTISCompareFormatting ws, changes, newColIndices, colCount, changes.Count + 1, groupInsertPos

    ' -- Milestone date comparison section (Old/New/delta-weeks per milestone) --
    WriteMilestoneDateSection ws, changes, oldData, newData, oldHeaders, newHeaders, totalCols
End Sub

'====================================================================
' APPLY TIS COMPARE FORMATTING
' Rev14: No Action column. Status=1, StartChg=2, EndChg=3, data=4+
'====================================================================
Private Sub ApplyTISCompareFormatting(ws As Worksheet, changes As Collection, _
                                       newColIndices() As Long, colCount As Long, lastRow As Long, _
                                       Optional groupInsertPos As Long = 0)
    Dim i As Long, j As Long, outCol As Long
    Dim changeRecord As Object
    Dim statusVal As String
    Dim totalCols As Long
    Dim startChangeVal As Variant
    Dim endChangeVal As Variant
    Dim headerRange As Range
    Dim statusCol As Long, startChgCol As Long, endChgCol As Long, dataStartCol As Long

    If changes.Count = 0 Then Exit Sub

    ' Rev14: Status=1, StartChg=2, EndChg=3, data=4+
    statusCol = 1
    startChgCol = 2
    endChgCol = 3
    dataStartCol = 4
    totalCols = colCount + 4

    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, totalCols))
    With headerRange
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(91, 108, 249)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 30
    End With
    With headerRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(61, 61, 85)
    End With

    For i = 2 To lastRow
        If i - 1 > changes.Count Then Exit For
        Set changeRecord = changes(i - 1)
        statusVal = CStr(changeRecord("Status"))

        If i Mod 2 = 0 Then
            ws.Range(ws.Cells(i, statusCol), ws.Cells(i, totalCols)).Interior.Color = RGB(42, 42, 62)
        End If

        Select Case statusVal
            Case "Added"
                ws.Cells(i, statusCol).Interior.Color = RGB(60, 20, 20)
                ws.Cells(i, statusCol).Font.Color = RGB(248, 113, 113)
                ws.Cells(i, statusCol).Font.Bold = True
            Case "Modified"
                ws.Cells(i, statusCol).Interior.Color = RGB(60, 50, 15)
                ws.Cells(i, statusCol).Font.Color = RGB(251, 191, 36)
                ws.Cells(i, statusCol).Font.Bold = True
            Case "Unchanged"
                ws.Cells(i, statusCol).Interior.Color = RGB(16, 60, 40)
                ws.Cells(i, statusCol).Font.Color = RGB(52, 211, 153)
                ws.Cells(i, statusCol).Font.Bold = True
            Case "Removed"
                ws.Cells(i, statusCol).Interior.Color = RGB(60, 20, 20)
                ws.Cells(i, statusCol).Font.Color = RGB(248, 113, 113)
                ws.Cells(i, statusCol).Font.Bold = True
                ws.Range(ws.Cells(i, statusCol), ws.Cells(i, totalCols)).Font.Italic = True
        End Select

        startChangeVal = ws.Cells(i, startChgCol).Value
        If IsNumeric(startChangeVal) And startChangeVal <> "" Then
            If CDbl(startChangeVal) > 0 Then
                ws.Cells(i, startChgCol).Interior.Color = RGB(60, 20, 20)
                ws.Cells(i, startChgCol).Font.Color = RGB(248, 113, 113)
            ElseIf CDbl(startChangeVal) < 0 Then
                ws.Cells(i, startChgCol).Interior.Color = RGB(16, 60, 40)
                ws.Cells(i, startChgCol).Font.Color = RGB(52, 211, 153)
            End If
            ws.Cells(i, startChgCol).Font.Bold = True
        End If

        endChangeVal = ws.Cells(i, endChgCol).Value
        If IsNumeric(endChangeVal) And endChangeVal <> "" Then
            If CDbl(endChangeVal) > 0 Then
                ws.Cells(i, endChgCol).Interior.Color = RGB(60, 20, 20)
                ws.Cells(i, endChgCol).Font.Color = RGB(248, 113, 113)
            ElseIf CDbl(endChangeVal) < 0 Then
                ws.Cells(i, endChgCol).Interior.Color = RGB(16, 60, 40)
                ws.Cells(i, endChgCol).Font.Color = RGB(52, 211, 153)
            End If
            ws.Cells(i, endChgCol).Font.Bold = True
        End If

        If statusVal = "Modified" Then
            If Not changeRecord("ChangedCols") Is Nothing Then
                outCol = dataStartCol
                For j = 1 To colCount
                    If j = groupInsertPos Then outCol = outCol + 1
                    If ColumnInChangedList(newColIndices(j), changeRecord("ChangedCols")) Then
                        ws.Cells(i, outCol).Interior.Color = RGB(60, 50, 15)
                        ws.Cells(i, outCol).Font.Color = RGB(251, 191, 36)
                        ws.Cells(i, outCol).Font.Bold = True
                    End If
                    outCol = outCol + 1
                Next j
            End If
        End If

        If changeRecord("IsDuplicate") Then
            ws.Range(ws.Cells(i, statusCol), ws.Cells(i, totalCols)).Font.Color = RGB(167, 139, 250)
            ws.Range(ws.Cells(i, statusCol), ws.Cells(i, totalCols)).Font.Bold = True
        End If
    Next i

    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, totalCols)).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(61, 61, 85)
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, totalCols)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(61, 61, 85)
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, totalCols)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(61, 61, 85)
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, totalCols)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(61, 61, 85)
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, totalCols)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(61, 61, 85)
    End With

    ws.Columns.AutoFit
    If ws.Columns(1).ColumnWidth < 10 Then ws.Columns(1).ColumnWidth = 10
    If ws.Columns(2).ColumnWidth < 12 Then ws.Columns(2).ColumnWidth = 12
    If ws.Columns(3).ColumnWidth < 12 Then ws.Columns(3).ColumnWidth = 12

    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select
End Sub

'====================================================================
' WRITE MILESTONE DATE COMPARISON SECTION
'====================================================================

Private Sub WriteMilestoneDateSection( _
        ws As Worksheet, _
        changes As Collection, _
        oldData As Variant, newData As Variant, _
        oldHeaders As Object, newHeaders As Object, _
        totalCols As Long)

    Dim milHeaders As Collection
    Dim milCount As Long
    Dim milStartCol As Long, milColCount As Long
    Dim milArray() As Variant
    Dim rowCount As Long
    Dim i As Long, m As Long, mCol As Long, outRow As Long
    Dim changeRecord As Object
    Dim statusVal As String
    Dim milHeaderName As String
    Dim oIdx As Long, nIdx As Long
    Dim oldMilVal As Variant, newMilVal As Variant
    Dim milHeaderRange As Range, milFullRange As Range
    Dim chgVal As Variant
    Dim oldCellVal As Variant, newCellVal As Variant
    Dim milOldCol As Long, milNewCol As Long, milDeltaCol As Long

    ' -- Collect all milestone date headers --
    Set milHeaders = New Collection

    If Not m_MilestoneStartHeaders Is Nothing Then
        For i = 1 To m_MilestoneStartHeaders.Count
            milHeaders.Add CStr(m_MilestoneStartHeaders(i))
        Next i
    End If
    If Not m_MilestoneEndHeaders Is Nothing Then
        For i = 1 To m_MilestoneEndHeaders.Count
            milHeaders.Add CStr(m_MilestoneEndHeaders(i))
        Next i
    End If

    milCount = milHeaders.Count
    If milCount = 0 Then Exit Sub

    ' -- Layout calculation --
    milStartCol = totalCols + 2
    milColCount = milCount * 3
    rowCount = changes.Count + 1

    ' -- Build output array --
    ReDim milArray(1 To rowCount, 1 To milColCount)

    ' Headers
    mCol = 0
    For m = 1 To milCount
        mCol = mCol + 1
        milArray(1, mCol) = "Old: " & milHeaders(m)
        mCol = mCol + 1
        milArray(1, mCol) = "New: " & milHeaders(m)
        mCol = mCol + 1
        milArray(1, mCol) = ChrW(916) & " " & milHeaders(m) & " (weeks)"
    Next m

    ' Data rows
    For i = 1 To changes.Count
        Set changeRecord = changes(i)
        outRow = i + 1
        statusVal = changeRecord("Status")

        mCol = 0
        For m = 1 To milCount
            milHeaderName = milHeaders(m)
            oldMilVal = Empty
            newMilVal = Empty

            If changeRecord("OldRow") > 0 Then
                If oldHeaders.Exists(LCase(milHeaderName)) Then
                    oIdx = oldHeaders(LCase(milHeaderName))
                    If oIdx > 0 And oIdx <= UBound(oldData, 2) Then
                        oldMilVal = oldData(changeRecord("OldRow"), oIdx)
                    End If
                End If
            End If

            If changeRecord("NewRow") > 0 Then
                If newHeaders.Exists(LCase(milHeaderName)) Then
                    nIdx = newHeaders(LCase(milHeaderName))
                    If nIdx > 0 And nIdx <= UBound(newData, 2) Then
                        newMilVal = newData(changeRecord("NewRow"), nIdx)
                    End If
                End If
            End If

            ' Old column
            mCol = mCol + 1
            Select Case statusVal
                Case "Added"
                    milArray(outRow, mCol) = ""
                Case "Removed", "Modified", "Unchanged"
                    milArray(outRow, mCol) = oldMilVal
            End Select

            ' New column
            mCol = mCol + 1
            Select Case statusVal
                Case "Removed"
                    milArray(outRow, mCol) = ""
                Case "Added", "Modified", "Unchanged"
                    milArray(outRow, mCol) = newMilVal
            End Select

            ' Delta weeks column (Modified only)
            mCol = mCol + 1
            If statusVal = "Modified" Then
                milArray(outRow, mCol) = CalcDateChangeWeeks(oldMilVal, newMilVal)
            Else
                milArray(outRow, mCol) = ""
            End If
        Next m
    Next i

    ' -- Batch write --
    ws.Range(ws.Cells(1, milStartCol), _
             ws.Cells(rowCount, milStartCol + milColCount - 1)).Value = milArray

    ' -- Format gap column --
    With ws.Cells(1, totalCols + 1)
        .Value = ""
        .Interior.Color = RGB(61, 61, 85)
    End With
    ws.Columns(totalCols + 1).ColumnWidth = 2
    ws.Range(ws.Cells(2, totalCols + 1), ws.Cells(rowCount, totalCols + 1)).Interior.Color = RGB(42, 42, 62)

    ' -- Format milestone headers --
    Set milHeaderRange = ws.Range(ws.Cells(1, milStartCol), _
                                  ws.Cells(1, milStartCol + milColCount - 1))
    With milHeaderRange
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(167, 139, 250)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 30
    End With

    ' -- Date format for Old/New columns, number format for Delta columns --
    For m = 0 To milCount - 1
        ws.Range(ws.Cells(2, milStartCol + m * 3), _
                 ws.Cells(rowCount, milStartCol + m * 3)).NumberFormat = "mm/dd/yyyy"
        ws.Range(ws.Cells(2, milStartCol + m * 3 + 1), _
                 ws.Cells(rowCount, milStartCol + m * 3 + 1)).NumberFormat = "mm/dd/yyyy"
        ws.Range(ws.Cells(2, milStartCol + m * 3 + 2), _
                 ws.Cells(rowCount, milStartCol + m * 3 + 2)).NumberFormat = "0.0"
    Next m

    ' -- Highlight changed dates (Modified rows where Old <> New) --
    For i = 2 To rowCount
        If i - 1 <= changes.Count Then
            Set changeRecord = changes(i - 1)
            If changeRecord("Status") = "Modified" Then
                For m = 0 To milCount - 1
                    milOldCol = milStartCol + m * 3
                    milNewCol = milStartCol + m * 3 + 1
                    milDeltaCol = milStartCol + m * 3 + 2

                    oldCellVal = ws.Cells(i, milOldCol).Value
                    newCellVal = ws.Cells(i, milNewCol).Value

                    If NormalizeValue(oldCellVal) <> NormalizeValue(newCellVal) Then
                        ws.Cells(i, milOldCol).Interior.Color = RGB(60, 50, 15)
                        ws.Cells(i, milOldCol).Font.Color = RGB(251, 191, 36)
                        ws.Cells(i, milOldCol).Font.Bold = True
                        ws.Cells(i, milNewCol).Interior.Color = RGB(60, 50, 15)
                        ws.Cells(i, milNewCol).Font.Color = RGB(251, 191, 36)
                        ws.Cells(i, milNewCol).Font.Bold = True

                        chgVal = ws.Cells(i, milDeltaCol).Value
                        If IsNumeric(chgVal) And CStr(chgVal) <> "" Then
                            If CDbl(chgVal) > 0 Then
                                ws.Cells(i, milDeltaCol).Interior.Color = RGB(60, 20, 20)
                                ws.Cells(i, milDeltaCol).Font.Color = RGB(248, 113, 113)
                            ElseIf CDbl(chgVal) < 0 Then
                                ws.Cells(i, milDeltaCol).Interior.Color = RGB(16, 60, 40)
                                ws.Cells(i, milDeltaCol).Font.Color = RGB(52, 211, 153)
                            End If
                            ws.Cells(i, milDeltaCol).Font.Bold = True
                        End If
                    End If
                Next m
            End If

            ' Alternating row shading (skip already-highlighted cells)
            If i Mod 2 = 0 Then
                For m = milStartCol To milStartCol + milColCount - 1
                    If ws.Cells(i, m).Interior.Color <> RGB(60, 50, 15) And _
                       ws.Cells(i, m).Interior.Color <> RGB(16, 60, 40) And _
                       ws.Cells(i, m).Interior.Color <> RGB(60, 20, 20) Then
                        ws.Cells(i, m).Interior.Color = RGB(42, 42, 62)
                    End If
                Next m
            End If
        End If
    Next i

    ' -- Borders --
    Set milFullRange = ws.Range(ws.Cells(1, milStartCol), _
                                ws.Cells(rowCount, milStartCol + milColCount - 1))
    With milFullRange.Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(61, 61, 85)
    End With
    With milFullRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(61, 61, 85)
    End With
    With milFullRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(61, 61, 85)
    End With
    With milFullRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(61, 61, 85)
    End With
    With milFullRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(61, 61, 85)
    End With

    ' -- AutoFit milestone columns --
    ws.Range(ws.Columns(milStartCol), ws.Columns(milStartCol + milColCount - 1)).AutoFit

End Sub

' === Change Tracking (unchanged from v2.0) ===
Private Sub CreateChangeTrackingLog(changes As Collection, newData As Variant, oldData As Variant, _
                                     newHeaders As Object, oldHeaders As Object, hasCEIDsSheet As Boolean)
    Dim ws As Worksheet
    Dim wsExists As Boolean
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim changeRecord As Object
    Dim startRow As Long
    Dim sourceRow As Long
    Dim colIdx As Long
    Dim logArray() As Variant
    Dim logRow As Long
    Dim changeCount As Long
    Dim nowStamp As String
    Dim trackCols() As String
    Dim trackNewIdx() As Long, trackOldIdx() As Long
    Dim trackCount As Long
    Dim entityTypeVal As String
    Dim entityTypeNewIdx As Long, entityTypeOldIdx As Long
    Dim eventTypeNewIdx As Long, eventTypeOldIdx As Long
    Dim publishedNewIdx As Long, publishedOldIdx As Long
    Const TOTAL_LOG_COLS As Long = 12
    If changes.Count = 0 Then Exit Sub
    trackCols = Split(TRACKING_COLS, "|")
    trackCount = UBound(trackCols) - LBound(trackCols) + 1
    ReDim trackNewIdx(0 To trackCount - 1): ReDim trackOldIdx(0 To trackCount - 1)
    For j = 0 To trackCount - 1
        If newHeaders.Exists(LCase(trackCols(j))) Then trackNewIdx(j) = newHeaders(LCase(trackCols(j))) Else trackNewIdx(j) = 0
        If oldHeaders.Exists(LCase(trackCols(j))) Then trackOldIdx(j) = oldHeaders(LCase(trackCols(j))) Else trackOldIdx(j) = 0
    Next j
    entityTypeNewIdx = 0: entityTypeOldIdx = 0
    If newHeaders.Exists("entity type") Then entityTypeNewIdx = newHeaders("entity type")
    If oldHeaders.Exists("entity type") Then entityTypeOldIdx = oldHeaders("entity type")
    eventTypeNewIdx = 0: eventTypeOldIdx = 0
    If newHeaders.Exists("event type") Then eventTypeNewIdx = newHeaders("event type")
    If oldHeaders.Exists("event type") Then eventTypeOldIdx = oldHeaders("event type")
    publishedNewIdx = 0: publishedOldIdx = 0
    If newHeaders.Exists("published") Then publishedNewIdx = newHeaders("published")
    If oldHeaders.Exists("published") Then publishedOldIdx = oldHeaders("published")
    changeCount = 0
    For i = 1 To changes.Count
        If changes(i)("Status") <> "Unchanged" Then changeCount = changeCount + 1
    Next i
    If changeCount = 0 Then Exit Sub
    wsExists = SheetExists(ThisWorkbook, SHEET_CHANGETRACKING)
    If Not wsExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_CHANGETRACKING
        ws.Cells.Interior.Color = 3022366       ' THEME_BG
        ws.Cells.Font.Name = "Segoe UI"         ' THEME_FONT
        ws.Cells.Font.Color = 15788258           ' THEME_TEXT
        ws.Activate
        ActiveWindow.DisplayGridlines = False
        ws.Cells(1, 1).Value = "Status": ws.Cells(1, 2).Value = "Project Start Change (weeks)"
        ws.Cells(1, 3).Value = "Project End Change (weeks)": ws.Cells(1, 4).Value = "Date Logged"
        ws.Cells(1, 5).Value = "ID": ws.Cells(1, 6).Value = "Site"
        ws.Cells(1, 7).Value = "Entity Code": ws.Cells(1, 8).Value = "Entity Type"
        ws.Cells(1, 9).Value = "CEID": ws.Cells(1, 10).Value = "Group"
        ws.Cells(1, 11).Value = "Event Type": ws.Cells(1, 12).Value = "Published"
        ApplyTrackingHeaderFormat ws
        startRow = 2
    Else
        Set ws = ThisWorkbook.Sheets(SHEET_CHANGETRACKING)
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        startRow = lastRow + 1
    End If
    ReDim logArray(1 To changeCount, 1 To TOTAL_LOG_COLS)
    nowStamp = Format(Now(), "mm/dd/yyyy hh:nn:ss")
    logRow = 0
    For i = 1 To changes.Count
        Set changeRecord = changes(i)
        If changeRecord("Status") <> "Unchanged" Then
            logRow = logRow + 1
            logArray(logRow, 1) = changeRecord("Status")
            If changeRecord("Status") = "Modified" Then logArray(logRow, 2) = changeRecord("StartChange"): logArray(logRow, 3) = changeRecord("EndChange") Else logArray(logRow, 2) = "": logArray(logRow, 3) = ""
            logArray(logRow, 4) = nowStamp
            If changeRecord("Status") = "Removed" Then
                sourceRow = changeRecord("OldRow")
                For j = 0 To trackCount - 1
                    colIdx = trackOldIdx(j)
                    If colIdx > 0 And colIdx <= UBound(oldData, 2) Then logArray(logRow, j + 5) = oldData(sourceRow, colIdx) Else logArray(logRow, j + 5) = ""
                Next j
                If entityTypeOldIdx > 0 And entityTypeOldIdx <= UBound(oldData, 2) Then entityTypeVal = CStr(oldData(sourceRow, entityTypeOldIdx)) Else entityTypeVal = ""
                If eventTypeOldIdx > 0 And eventTypeOldIdx <= UBound(oldData, 2) Then logArray(logRow, 11) = oldData(sourceRow, eventTypeOldIdx) Else logArray(logRow, 11) = ""
                If publishedOldIdx > 0 And publishedOldIdx <= UBound(oldData, 2) Then logArray(logRow, 12) = oldData(sourceRow, publishedOldIdx) Else logArray(logRow, 12) = ""
            Else
                sourceRow = changeRecord("NewRow")
                For j = 0 To trackCount - 1
                    colIdx = trackNewIdx(j)
                    If colIdx > 0 And colIdx <= UBound(newData, 2) Then logArray(logRow, j + 5) = newData(sourceRow, colIdx) Else logArray(logRow, j + 5) = ""
                Next j
                If entityTypeNewIdx > 0 And entityTypeNewIdx <= UBound(newData, 2) Then entityTypeVal = CStr(newData(sourceRow, entityTypeNewIdx)) Else entityTypeVal = ""
                If eventTypeNewIdx > 0 And eventTypeNewIdx <= UBound(newData, 2) Then logArray(logRow, 11) = newData(sourceRow, eventTypeNewIdx) Else logArray(logRow, 11) = ""
                If publishedNewIdx > 0 And publishedNewIdx <= UBound(newData, 2) Then logArray(logRow, 12) = newData(sourceRow, publishedNewIdx) Else logArray(logRow, 12) = ""
            End If
            logArray(logRow, 10) = LookupGroup(entityTypeVal, hasCEIDsSheet)
        End If
    Next i
    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + changeCount - 1, TOTAL_LOG_COLS)).Value = logArray
    ApplyTrackingDataFormat ws, startRow, startRow + changeCount - 1, TOTAL_LOG_COLS
    ws.Columns.AutoFit
    If ws.Columns(1).ColumnWidth < 10 Then ws.Columns(1).ColumnWidth = 10
    If ws.Columns(2).ColumnWidth < 12 Then ws.Columns(2).ColumnWidth = 12
    If ws.Columns(3).ColumnWidth < 12 Then ws.Columns(3).ColumnWidth = 12
End Sub

Private Sub ApplyTrackingHeaderFormat(ws As Worksheet)
    Dim headerRange As Range
    Const TOTAL_LOG_COLS As Long = 12
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, TOTAL_LOG_COLS))
    With headerRange
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(91, 108, 249)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True: .RowHeight = 30
    End With
    With headerRange.Borders: .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(61, 61, 85): End With
    ws.Rows(2).Select: ActiveWindow.FreezePanes = True: ws.Cells(1, 1).Select
End Sub

Private Sub ApplyTrackingDataFormat(ws As Worksheet, startRow As Long, endRow As Long, totalCols As Long)
    Dim i As Long
    Dim c As Long
    Dim statusVal As String
    Dim changeVal As Variant
    For i = startRow To endRow
        If i Mod 2 = 0 Then ws.Range(ws.Cells(i, 1), ws.Cells(i, totalCols)).Interior.Color = RGB(42, 42, 62)
        statusVal = CStr(ws.Cells(i, 1).Value)
        Select Case statusVal
            Case "Added": ws.Cells(i, 1).Interior.Color = RGB(60, 20, 20): ws.Cells(i, 1).Font.Color = RGB(248, 113, 113): ws.Cells(i, 1).Font.Bold = True
            Case "Modified": ws.Cells(i, 1).Interior.Color = RGB(60, 50, 15): ws.Cells(i, 1).Font.Color = RGB(251, 191, 36): ws.Cells(i, 1).Font.Bold = True
            Case "Removed": ws.Cells(i, 1).Interior.Color = RGB(60, 20, 20): ws.Cells(i, 1).Font.Color = RGB(248, 113, 113): ws.Cells(i, 1).Font.Bold = True
        End Select
        For c = 2 To 3
            changeVal = ws.Cells(i, c).Value
            If IsNumeric(changeVal) And changeVal <> "" Then
                If CDbl(changeVal) > 0 Then ws.Cells(i, c).Interior.Color = RGB(60, 20, 20): ws.Cells(i, c).Font.Color = RGB(248, 113, 113)
                If CDbl(changeVal) < 0 Then ws.Cells(i, c).Interior.Color = RGB(16, 60, 40): ws.Cells(i, c).Font.Color = RGB(52, 211, 153)
                ws.Cells(i, c).Font.Bold = True
            End If
        Next c
    Next i
    With ws.Range(ws.Cells(1, 1), ws.Cells(endRow, totalCols)).Borders: .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(61, 61, 85): End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(endRow, totalCols)).Borders(xlEdgeTop): .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(61, 61, 85): End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(endRow, totalCols)).Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(61, 61, 85): End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(endRow, totalCols)).Borders(xlEdgeLeft): .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(61, 61, 85): End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(endRow, totalCols)).Borders(xlEdgeRight): .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(61, 61, 85): End With
End Sub

'====================================================================
' SHOW SUMMARY REPORT
' Rev14: Updated to guide user to run Build Working Sheet
'====================================================================

Private Sub ShowSummaryReport(changes As Collection, compareSheetName As String)
    Dim countAdded As Long, countModified As Long, countRemoved As Long, countDuplicates As Long
    Dim countDateChanges As Long, countDateChangeSystems As Long
    Dim i As Long
    Dim changeRecord As Object
    Dim message As String
    Dim systemHasDateChange As Boolean

    For i = 1 To changes.Count
        Set changeRecord = changes(i)
        Select Case changeRecord("Status")
            Case "Added": countAdded = countAdded + 1
            Case "Modified":
                countModified = countModified + 1
                ' Count date changes
                systemHasDateChange = False
                If changeRecord("StartChange") <> "" Then
                    If IsNumeric(changeRecord("StartChange")) Then
                        If CDbl(changeRecord("StartChange")) <> 0 Then
                            countDateChanges = countDateChanges + 1
                            systemHasDateChange = True
                        End If
                    End If
                End If
                If changeRecord("EndChange") <> "" Then
                    If IsNumeric(changeRecord("EndChange")) Then
                        If CDbl(changeRecord("EndChange")) <> 0 Then
                            countDateChanges = countDateChanges + 1
                            If Not systemHasDateChange Then systemHasDateChange = True
                        End If
                    End If
                End If
                If systemHasDateChange Then countDateChangeSystems = countDateChangeSystems + 1
            Case "Removed": countRemoved = countRemoved + 1
        End Select
        If changeRecord("IsDuplicate") Then countDuplicates = countDuplicates + 1
    Next i

    message = "TIS Changes Found:" & vbCrLf & vbCrLf & _
              "  " & Chr(149) & " " & countAdded & " new systems" & vbCrLf & _
              "  " & Chr(149) & " " & countRemoved & " systems removed" & vbCrLf & _
              "  " & Chr(149) & " " & countDateChanges & " date changes across " & countDateChangeSystems & " systems" & vbCrLf

    If countDuplicates > 0 Then
        message = message & "  " & Chr(149) & " " & countDuplicates & " duplicates" & vbCrLf
    End If

    message = message & vbCrLf & _
              "TIScompare sheet updated for detailed review." & vbCrLf & _
              "Working Sheet will be updated automatically."

    MsgBox message, vbInformation, "TIS Comparison Complete"
End Sub

' === Utility functions ===
Private Function GetUniqueTISCompareSheetName() As String
    Dim counter As Long
    Dim testName As String
    testName = SHEET_TISCOMPARE_BASE
    counter = 0
    Do While SheetExists(ThisWorkbook, testName)
        counter = counter + 1
        testName = SHEET_TISCOMPARE_BASE & CStr(counter)
    Loop
    GetUniqueTISCompareSheetName = testName
End Function

'====================================================================
' UPDATE WORKING SHEET FROM TIS
' Updates Working Sheet TIS date columns in-place from the TIS sheet.
' Non-destructive: only TIS columns change. User data untouched.
' Also: adds new systems, cancels removed systems, reactivates returning.
'====================================================================

Private Sub UpdateWorkingSheetFromTIS(Optional filters As Object = Nothing)
    On Error GoTo UpdateErrorHandler
    ' PERFORMANCE: disable screen updates, events, and auto-calc during entire update
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Ensure filters is a valid (possibly empty) Dictionary
    If filters Is Nothing Then Set filters = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Set ws = TISCommon.FindWorkingSheet()
    If ws Is Nothing Then GoTo UpdateCleanup

    Dim wsTIS As Worksheet
    If Not TISCommon.SheetExists(ThisWorkbook, SHEET_TIS) Then GoTo UpdateCleanup
    Set wsTIS = ThisWorkbook.Sheets(SHEET_TIS)

    ' 1. Bulk-read Working Sheet into array
    Dim wsHdr As Long
    wsHdr = TISCommon.FindHeaderRow(ws)
    If wsHdr = 0 Then GoTo UpdateCleanup
    Dim wsLastRow As Long
    wsLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If wsLastRow <= wsHdr Then wsLastRow = wsHdr  ' empty sheet
    Dim wsMaxCol As Long
    wsMaxCol = ws.Cells(wsHdr, ws.Columns.Count).End(xlToLeft).Column

    ' Build WS header map (LCase normalized -> column index)
    Dim wsHeaderMap As Object
    Set wsHeaderMap = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To wsMaxCol
        Dim hv As String
        hv = LCase(Trim(Replace(Replace(CStr(ws.Cells(wsHdr, c).Value), vbLf, ""), vbCr, "")))
        If hv <> "" And Not wsHeaderMap.exists(hv) Then wsHeaderMap(hv) = c
    Next c

    ' Schema check: Working Sheet must have Our Date columns (Rev14 schema)
    Dim hasSchema As Boolean
    hasSchema = wsHeaderMap.exists(LCase(TIS_COL_OUR_CONVS))
    If Not hasSchema Then
        Application.ScreenUpdating = True
        MsgBox "Working Sheet does not have Our Date columns." & vbCrLf & _
               "Run 'Build Working Sheet' first to create the Rev14 schema.", _
               vbExclamation, "Schema Mismatch"
        GoTo UpdateCleanup
    End If

    ' Find key columns in Working Sheet
    Dim wsSiteCol As Long, wsECCol As Long, wsETCol As Long, wsStatusCol As Long
    wsSiteCol = 0: wsECCol = 0: wsETCol = 0: wsStatusCol = 0
    If wsHeaderMap.exists("site") Then wsSiteCol = wsHeaderMap("site")
    If wsHeaderMap.exists("entity code") Then wsECCol = wsHeaderMap("entity code")
    If wsHeaderMap.exists("event type") Then wsETCol = wsHeaderMap("event type")
    If wsHeaderMap.exists(LCase(TIS_COL_STATUS)) Then wsStatusCol = wsHeaderMap(LCase(TIS_COL_STATUS))
    If wsSiteCol = 0 Or wsECCol = 0 Or wsETCol = 0 Then GoTo UpdateCleanup

    ' Build WS key->row dictionary (bulk array read for performance)
    Dim wsKeyMap As Object
    Set wsKeyMap = CreateObject("Scripting.Dictionary")
    Dim r As Long
    Dim wsAllData As Variant
    If wsLastRow > wsHdr Then
        wsAllData = ws.Range(ws.Cells(wsHdr + 1, 1), ws.Cells(wsLastRow, wsMaxCol)).Value
        For r = 1 To UBound(wsAllData, 1)
            Dim wsKey As String
            wsKey = LCase(Trim(CStr(wsAllData(r, wsSiteCol) & ""))) & "|" & _
                    LCase(Trim(CStr(wsAllData(r, wsECCol) & ""))) & "|" & _
                    LCase(Trim(CStr(wsAllData(r, wsETCol) & "")))
            If wsKey <> "||" And Not wsKeyMap.exists(wsKey) Then wsKeyMap(wsKey) = wsHdr + r
        Next r
    End If

    ' 2. Bulk-read TIS sheet into array
    Dim tisHdr As Long
    tisHdr = TISCommon.FindHeaderRow(wsTIS)
    If tisHdr = 0 Then GoTo UpdateCleanup
    Dim tisLastRow As Long
    tisLastRow = wsTIS.Cells(wsTIS.Rows.Count, 1).End(xlUp).Row
    If tisLastRow <= tisHdr Then GoTo UpdateCleanup
    Dim tisMaxCol As Long
    tisMaxCol = wsTIS.Cells(tisHdr, wsTIS.Columns.Count).End(xlToLeft).Column

    ' Build TIS header map
    Dim tisHeaderMap As Object
    Set tisHeaderMap = CreateObject("Scripting.Dictionary")
    For c = 1 To tisMaxCol
        hv = LCase(Trim(Replace(Replace(CStr(wsTIS.Cells(tisHdr, c).Value), vbLf, ""), vbCr, "")))
        If hv <> "" And Not tisHeaderMap.exists(hv) Then tisHeaderMap(hv) = c
    Next c

    ' Find key columns in TIS
    Dim tisSiteCol As Long, tisECCol As Long, tisETCol As Long
    tisSiteCol = 0: tisECCol = 0: tisETCol = 0
    If tisHeaderMap.exists("site") Then tisSiteCol = tisHeaderMap("site")
    If tisHeaderMap.exists("entity code") Then tisECCol = tisHeaderMap("entity code")
    If tisHeaderMap.exists("event type") Then tisETCol = tisHeaderMap("event type")
    If tisSiteCol = 0 Or tisECCol = 0 Or tisETCol = 0 Then GoTo UpdateCleanup

    ' Build TIS key->row dictionary with duplicate detection (two-pass).
    ' Pass 1: count occurrences of each key.
    ' Pass 2: build tisKeyMap (first occurrence wins); collect extra occurrences into
    '         tisDupExtraRows so ALL rows are appended. Duplicates get red borders on Entity Code.
    Dim tisKeyMap As Object
    Set tisKeyMap = CreateObject("Scripting.Dictionary")
    Dim tisAllData As Variant
    tisAllData = wsTIS.Range(wsTIS.Cells(tisHdr + 1, 1), wsTIS.Cells(tisLastRow, tisMaxCol)).Value

    Dim tisCounts As Object
    Set tisCounts = CreateObject("Scripting.Dictionary")
    Dim tisKey As String
    For r = 1 To UBound(tisAllData, 1)
        ' Apply Definitions filters — skip TIS rows that don't match
        If Not RowMatchesFilters(tisAllData, r, tisHeaderMap, filters) Then GoTo NextTisCount
        tisKey = LCase(Trim(CStr(tisAllData(r, tisSiteCol) & ""))) & "|" & _
                 LCase(Trim(CStr(tisAllData(r, tisECCol) & ""))) & "|" & _
                 LCase(Trim(CStr(tisAllData(r, tisETCol) & "")))
        If tisKey <> "||" Then
            If tisCounts.exists(tisKey) Then
                tisCounts(tisKey) = tisCounts(tisKey) + 1
            Else
                tisCounts(tisKey) = 1
            End If
        End If
NextTisCount:
    Next r

    ' Build tisDupKeys set (keys with count > 1) and warn user
    Dim tisDupKeys As Object
    Set tisDupKeys = CreateObject("Scripting.Dictionary")
    Dim tisDupExtraRows As Collection
    Set tisDupExtraRows = New Collection
    Dim dupMsg As String
    Dim dupCount As Long
    dupMsg = ""
    dupCount = 0
    Dim tkv As Variant
    For Each tkv In tisCounts.keys
        If tisCounts(tkv) > 1 Then
            dupCount = dupCount + 1
            tisDupKeys(CStr(tkv)) = True
            Dim dupParts() As String
            dupParts = Split(CStr(tkv), "|")
            dupMsg = dupMsg & "  " & Chr(149) & " " & dupParts(1) & " / " & dupParts(2) & _
                     " (" & tisCounts(tkv) & " rows)" & vbCrLf
        End If
    Next tkv

    If dupCount > 0 Then
        Application.ScreenUpdating = True
        MsgBox "Duplicate systems found in TIS (" & dupCount & " key(s) appear more than once):" & _
               vbCrLf & vbCrLf & dupMsg & vbCrLf & _
               "All rows will be kept. Duplicate rows will be appended with red borders on Entity Code.", _
               vbExclamation, "TIS Duplicate Keys"
        Application.ScreenUpdating = False
    End If

    ' Pass 2: first occurrence goes into tisKeyMap; extra occurrences collected into tisDupExtraRows
    Dim tisKeyFirstSeen As Object
    Set tisKeyFirstSeen = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(tisAllData, 1)
        ' Apply same Definitions filter as Pass 1
        If Not RowMatchesFilters(tisAllData, r, tisHeaderMap, filters) Then GoTo NextTisMap
        tisKey = LCase(Trim(CStr(tisAllData(r, tisSiteCol) & ""))) & "|" & _
                 LCase(Trim(CStr(tisAllData(r, tisECCol) & ""))) & "|" & _
                 LCase(Trim(CStr(tisAllData(r, tisETCol) & "")))
        If tisKey <> "||" Then
            If Not tisKeyFirstSeen.exists(tisKey) Then
                tisKeyMap(tisKey) = tisHdr + r
                tisKeyFirstSeen(tisKey) = True
            ElseIf tisDupKeys.exists(tisKey) Then
                tisDupExtraRows.Add tisHdr + r
            End If
        End If
NextTisMap:
    Next r
    Set tisCounts = Nothing
    Set tisKeyFirstSeen = Nothing

    ' 3. Identify which WS columns are TIS-sourced (not Our Dates, not user fields)
    ' TIS columns = columns whose headers exist in the TIS header map
    ' Exclude: Our Date columns, Status, Lock?, Health, WhatIf, user fields
    Dim excludeHeaders As Object
    Set excludeHeaders = CreateObject("Scripting.Dictionary")
    ' Our Date columns
    excludeHeaders(LCase(TIS_COL_OUR_SET)) = True
    excludeHeaders(LCase(TIS_COL_OUR_SL1)) = True
    excludeHeaders(LCase(TIS_COL_OUR_SL2)) = True
    excludeHeaders(LCase(TIS_COL_OUR_SQ)) = True
    excludeHeaders(LCase(TIS_COL_OUR_CONVS)) = True
    excludeHeaders(LCase(TIS_COL_OUR_CONVF)) = True
    excludeHeaders(LCase(TIS_COL_OUR_MRCLS)) = True
    excludeHeaders(LCase(TIS_COL_OUR_MRCLF)) = True
    ' Operational columns
    excludeHeaders(LCase(TIS_COL_STATUS)) = True
    excludeHeaders(LCase(TIS_COL_LOCK)) = True
    excludeHeaders(LCase(TIS_COL_HEALTH)) = True
    excludeHeaders(LCase(TIS_COL_WHATIF)) = True
    ' User fields
    excludeHeaders("escalated") = True
    excludeHeaders("ship date") = True
    excludeHeaders("soc available") = True
    excludeHeaders("soc uploaded?") = True
    excludeHeaders("staffed?") = True
    excludeHeaders("comments") = True
    excludeHeaders("watch") = True
    excludeHeaders("completed") = True
    excludeHeaders("bod1") = True
    excludeHeaders("bod2") = True
    ' Key fields (never update these)
    excludeHeaders("site") = True
    excludeHeaders("entity code") = True
    excludeHeaders("event type") = True

    ' Build list of TIS-updatable columns: WS col -> TIS col pairs
    Dim updateCols As Object  ' wsColIdx -> tisColIdx
    Set updateCols = CreateObject("Scripting.Dictionary")
    Dim wsHdrKey As Variant
    For Each wsHdrKey In wsHeaderMap.keys
        If Not excludeHeaders.exists(CStr(wsHdrKey)) Then
            If tisHeaderMap.exists(CStr(wsHdrKey)) Then
                updateCols(wsHeaderMap(CStr(wsHdrKey))) = tisHeaderMap(CStr(wsHdrKey))
            End If
        End If
    Next wsHdrKey

    ' Also handle the SDD rename: TIS has "SDD", WS has "TIS SDD"
    If wsHeaderMap.exists("tis sdd") And tisHeaderMap.exists("sdd") Then
        updateCols(wsHeaderMap("tis sdd")) = tisHeaderMap("sdd")
    End If

    ' 4. Process: update existing, cancel missing, add new
    Dim changedCells As New Collection  ' Array(wsRow, wsCol, oldValStr)
    Dim newSystemRows As New Collection  ' TIS row numbers for new systems

    ' wsAllData and tisAllData were already bulk-read above for key-map construction.
    ' Reuse those arrays here — no second COM read needed.

    ' Collect writes for batch application after loop
    Dim writeQueue As New Collection

    ' 4a. Update existing projects + detect changes (IN MEMORY)
    Dim tisKeyVar As Variant
    Dim processedWSKeys As Object
    Set processedWSKeys = CreateObject("Scripting.Dictionary")

    For Each tisKeyVar In tisKeyMap.keys
        Dim tKey As String
        tKey = CStr(tisKeyVar)
        Dim tisRow As Long
        tisRow = tisKeyMap(tKey)

        If wsKeyMap.exists(tKey) Then
            ' Project exists in both - update TIS columns
            Dim wsRow As Long
            wsRow = wsKeyMap(tKey)
            processedWSKeys(tKey) = True

            Dim ucKey As Variant
            For Each ucKey In updateCols.keys
                Dim wsC As Long, tisC As Long
                wsC = CLng(ucKey)
                tisC = updateCols(ucKey)

                Dim oldVal As Variant, newVal As Variant
                ' Read from arrays, not cells (PERFORMANCE)
                Dim wsAI As Long, tisAI As Long
                wsAI = wsRow - wsHdr: tisAI = tisRow - tisHdr
                If wsAI >= 1 And wsAI <= UBound(wsAllData, 1) And wsC <= UBound(wsAllData, 2) Then
                    oldVal = wsAllData(wsAI, wsC)
                Else: oldVal = Empty
                End If
                If tisAI >= 1 And tisAI <= UBound(tisAllData, 1) And tisC <= UBound(tisAllData, 2) Then
                    newVal = tisAllData(tisAI, tisC)
                Else: newVal = Empty
                End If

                ' Compare (date-aware)
                Dim changed As Boolean
                changed = False
                If IsDate(oldVal) And IsDate(newVal) Then
                    changed = (CLng(CDate(oldVal)) <> CLng(CDate(newVal)))
                ElseIf IsEmpty(oldVal) And IsEmpty(newVal) Then
                    changed = False
                Else
                    changed = (LCase(Trim(CStr(oldVal & ""))) <> LCase(Trim(CStr(newVal & ""))))
                End If

                If changed Then
                    writeQueue.Add Array(wsRow, wsC, newVal)
                    ' Collect for batch orange fill + comment
                    Dim oldStr As String
                    If IsDate(oldVal) Then
                        oldStr = Format(oldVal, "mm/dd/yyyy")
                    ElseIf IsEmpty(oldVal) Then
                        oldStr = "(empty)"
                    Else
                        oldStr = CStr(oldVal)
                    End If
                    changedCells.Add Array(wsRow, wsC, oldStr)
                End If
            Next ucKey

            ' Reactivation: if was Cancelled but now back in TIS
            If wsStatusCol > 0 Then
                Dim reactAI As Long
                reactAI = wsRow - wsHdr
                If reactAI >= 1 And reactAI <= UBound(wsAllData, 1) Then
                    If LCase(Trim(CStr(wsAllData(reactAI, wsStatusCol) & ""))) = "cancelled" Then
                        writeQueue.Add Array(wsRow, wsStatusCol, "Active")
                    End If
                End If
            End If
        Else
            ' New system - collect for append later
            newSystemRows.Add tisRow
        End If
    Next tisKeyVar

    ' 4b. Cancel projects in WS but not in TIS
    Dim wsKeyVar As Variant
    For Each wsKeyVar In wsKeyMap.keys
        If Not processedWSKeys.exists(CStr(wsKeyVar)) Then
            Dim cancelRow As Long
            cancelRow = wsKeyMap(CStr(wsKeyVar))
            If wsStatusCol > 0 Then
                Dim currentStatus As String
                Dim canAI As Long
                canAI = cancelRow - wsHdr
                If canAI >= 1 And canAI <= UBound(wsAllData, 1) Then
                    currentStatus = LCase(Trim(CStr(wsAllData(canAI, wsStatusCol) & "")))
                    ' Cancel Active and On Hold rows removed from TIS.
                    ' Completed and Non IQ rows are left unchanged — they are terminal states.
                    If currentStatus = "active" Or currentStatus = "on hold" Then
                        writeQueue.Add Array(cancelRow, wsStatusCol, "Cancelled")
                    End If
                End If
            End If
        End If
    Next wsKeyVar

    ' Add duplicate extra rows (2nd+ occurrences of the same key) to newSystemRows for appending.
    ' These will be treated exactly like new systems and will receive red borders below.
    Dim dupExtraRow As Variant
    For Each dupExtraRow In tisDupExtraRows
        newSystemRows.Add dupExtraRow
    Next dupExtraRow
    Set tisDupExtraRows = Nothing

    ' Show summary before applying changes
    Dim addCount As Long, changeCount As Long, cancelCount As Long, reactivateCount As Long
    addCount = newSystemRows.Count
    changeCount = changedCells.Count

    ' Count cancellations and reactivations from writeQueue
    cancelCount = 0: reactivateCount = 0
    Dim qi As Long
    For qi = 1 To writeQueue.Count
        Dim qInfo As Variant
        qInfo = writeQueue(qi)
        If CLng(qInfo(1)) = wsStatusCol Then
            If CStr(qInfo(2)) = "Cancelled" Then cancelCount = cancelCount + 1
            If CStr(qInfo(2)) = "Active" Then reactivateCount = reactivateCount + 1
        End If
    Next qi

    Application.ScreenUpdating = True
    Dim summaryMsg As String
    summaryMsg = "TIS Update Summary:" & vbCrLf & vbCrLf
    If addCount > 0 Then summaryMsg = summaryMsg & "  " & Chr(149) & " " & addCount & " new system(s) added" & vbCrLf
    If changeCount > 0 Then summaryMsg = summaryMsg & "  " & Chr(149) & " " & changeCount & " date change(s) detected" & vbCrLf
    If cancelCount > 0 Then summaryMsg = summaryMsg & "  " & Chr(149) & " " & cancelCount & " system(s) removed from TIS" & vbCrLf
    If reactivateCount > 0 Then summaryMsg = summaryMsg & "  " & Chr(149) & " " & reactivateCount & " system(s) reactivated" & vbCrLf
    If addCount = 0 And changeCount = 0 And cancelCount = 0 And reactivateCount = 0 Then
        summaryMsg = summaryMsg & "  No changes detected." & vbCrLf
    End If
    summaryMsg = summaryMsg & vbCrLf & "Changed cells will be highlighted in orange."
    MsgBox summaryMsg, vbInformation, "TIS Update"
    Application.ScreenUpdating = False

    ' BATCH WRITE all queued changes (updates + reactivations + cancellations)
    Dim wi As Long
    For wi = 1 To writeQueue.Count
        Dim wInfo As Variant
        wInfo = writeQueue(wi)
        ws.Cells(wInfo(0), wInfo(1)).Value = wInfo(2)
        If IsDate(wInfo(2)) Then ws.Cells(wInfo(0), wInfo(1)).NumberFormat = "mm/dd/yyyy"
    Next wi

    ' 4c. Append new systems (BULK: build rows in memory, write all at once)
    If newSystemRows.Count > 0 Then
        Dim appendRow As Long
        appendRow = ws.Cells(ws.Rows.Count, wsSiteCol).End(xlUp).Row + 1

        ' Build output array for all new rows at once
        Dim newRowCount As Long
        newRowCount = newSystemRows.Count
        Dim newRowArr() As Variant
        ReDim newRowArr(1 To newRowCount, 1 To wsMaxCol)

        ' Our Date TIS mapping
        Dim ourTISPairs As Variant
        ourTISPairs = Array( _
            Array(LCase(TIS_COL_OUR_SET), LCase(TIS_SRC_SET)), _
            Array(LCase(TIS_COL_OUR_SL1), LCase(TIS_SRC_SL1)), _
            Array(LCase(TIS_COL_OUR_SL2), LCase(TIS_SRC_SL2)), _
            Array(LCase(TIS_COL_OUR_SQ), LCase(TIS_SRC_SQ)), _
            Array(LCase(TIS_COL_OUR_CONVS), LCase(TIS_SRC_CONVS)), _
            Array(LCase(TIS_COL_OUR_CONVF), LCase(TIS_SRC_CONVF)), _
            Array(LCase(TIS_COL_OUR_MRCLS), LCase(TIS_SRC_MRCLS)), _
            Array(LCase(TIS_COL_OUR_MRCLF), LCase(TIS_SRC_MRCLF)))

        Dim idHeaders As Variant
        idHeaders = Array("ceid", "entity type", "group")

        Dim nri As Long
        nri = 0
        Dim newSysRow As Variant
        For Each newSysRow In newSystemRows
            nri = nri + 1
            Dim tisR As Long
            tisR = CLng(newSysRow)
            Dim tisArrR As Long
            tisArrR = tisR - tisHdr

            ' Copy TIS columns from tisAllData array
            Dim ucKey2 As Variant
            For Each ucKey2 In updateCols.keys
                wsC = CLng(ucKey2)
                tisC = updateCols(ucKey2)
                If tisArrR >= 1 And tisArrR <= UBound(tisAllData, 1) And _
                   tisC >= 1 And tisC <= UBound(tisAllData, 2) And _
                   wsC >= 1 And wsC <= wsMaxCol Then
                    newRowArr(nri, wsC) = tisAllData(tisArrR, tisC)
                End If
            Next ucKey2

            ' Key fields from tisAllData
            If tisArrR >= 1 And tisArrR <= UBound(tisAllData, 1) Then
                newRowArr(nri, wsSiteCol) = tisAllData(tisArrR, tisSiteCol)
                newRowArr(nri, wsECCol) = tisAllData(tisArrR, tisECCol)
                newRowArr(nri, wsETCol) = tisAllData(tisArrR, tisETCol)
            End If

            ' Identity fields from tisAllData
            Dim ih As Long
            For ih = LBound(idHeaders) To UBound(idHeaders)
                If wsHeaderMap.exists(idHeaders(ih)) And tisHeaderMap.exists(idHeaders(ih)) Then
                    Dim idWsC As Long, idTisC As Long
                    idWsC = wsHeaderMap(idHeaders(ih))
                    idTisC = tisHeaderMap(idHeaders(ih))
                    If idWsC <= wsMaxCol And tisArrR >= 1 And tisArrR <= UBound(tisAllData, 1) And _
                       idTisC <= UBound(tisAllData, 2) Then
                        newRowArr(nri, idWsC) = tisAllData(tisArrR, idTisC)
                    End If
                End If
            Next ih

            ' Status = Active
            If wsStatusCol > 0 And wsStatusCol <= wsMaxCol Then newRowArr(nri, wsStatusCol) = "Active"

            ' Our Dates from TIS
            Dim pi As Long
            For pi = LBound(ourTISPairs) To UBound(ourTISPairs)
                Dim ourKey2 As String, tisKey2 As String
                ourKey2 = ourTISPairs(pi)(0)
                tisKey2 = ourTISPairs(pi)(1)
                If wsHeaderMap.exists(ourKey2) And tisHeaderMap.exists(tisKey2) Then
                    Dim ourWsC As Long, ourTisC As Long
                    ourWsC = wsHeaderMap(ourKey2)
                    ourTisC = tisHeaderMap(tisKey2)
                    If ourWsC <= wsMaxCol And tisArrR >= 1 And tisArrR <= UBound(tisAllData, 1) And _
                       ourTisC <= UBound(tisAllData, 2) Then
                        Dim tisDateVal As Variant
                        tisDateVal = tisAllData(tisArrR, ourTisC)
                        If IsDate(tisDateVal) Then newRowArr(nri, ourWsC) = tisDateVal
                    End If
                End If
            Next pi
        Next newSysRow

        ' Extend the ListObject table to cover the new rows before writing.
        ' Without this, appended rows sit outside the table: no AutoFilter, no slicer coverage.
        Dim lo As ListObject
        If ws.ListObjects.Count > 0 Then
            Set lo = ws.ListObjects(1)
            Dim loEndRow As Long
            loEndRow = lo.Range.Row + lo.Range.Rows.Count - 1
            If appendRow + newRowCount - 1 > loEndRow Then
                On Error Resume Next
                lo.Resize ws.Range(lo.Range.Cells(1, 1), _
                    ws.Cells(appendRow + newRowCount - 1, _
                             lo.Range.Column + lo.Range.Columns.Count - 1))
                On Error GoTo 0
            End If
        End If

        ' SINGLE bulk write for ALL new rows
        ws.Range(ws.Cells(appendRow, 1), ws.Cells(appendRow + newRowCount - 1, wsMaxCol)).Value = newRowArr

        ' Apply date format and blue border to Our Date columns for new rows.
        ' Blue border = CLAUDE.md visual signal: "auto-filled from TIS — please review and confirm."
        Dim pi2 As Long
        For pi2 = LBound(ourTISPairs) To UBound(ourTISPairs)
            If wsHeaderMap.exists(ourTISPairs(pi2)(0)) Then
                Dim fmtCol As Long
                fmtCol = wsHeaderMap(ourTISPairs(pi2)(0))
                Dim fmtRng As Range
                Set fmtRng = ws.Range(ws.Cells(appendRow, fmtCol), ws.Cells(appendRow + newRowCount - 1, fmtCol))
                fmtRng.NumberFormat = "mm/dd/yyyy"
                With fmtRng.Borders
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = CLR_NEW_DATE_BORDER
                End With
            End If
        Next pi2
        ' Apply red border to Entity Code cells for newly appended duplicate rows
        If tisDupKeys.Count > 0 And wsECCol > 0 Then
            Dim dupNri As Long
            dupNri = 0
            Dim dupSysRow As Variant
            For Each dupSysRow In newSystemRows
                dupNri = dupNri + 1
                Dim dupTisArrR As Long
                dupTisArrR = CLng(dupSysRow) - tisHdr
                If dupTisArrR >= 1 And dupTisArrR <= UBound(tisAllData, 1) Then
                    Dim dupChkKey As String
                    dupChkKey = LCase(Trim(CStr(tisAllData(dupTisArrR, tisSiteCol) & ""))) & "|" & _
                                LCase(Trim(CStr(tisAllData(dupTisArrR, tisECCol) & ""))) & "|" & _
                                LCase(Trim(CStr(tisAllData(dupTisArrR, tisETCol) & "")))
                    If tisDupKeys.exists(dupChkKey) Then
                        With ws.Cells(appendRow + dupNri - 1, wsECCol).Borders
                            .LineStyle = xlContinuous
                            .Weight = xlMedium
                            .Color = vbRed
                        End With
                    End If
                End If
            Next dupSysRow
        End If

        ' Extend Health formula to new rows
        Dim healthCol As Long
        healthCol = 0
        If wsHeaderMap.exists(LCase(TIS_COL_HEALTH)) Then healthCol = wsHeaderMap(LCase(TIS_COL_HEALTH))
        If healthCol > 0 Then
            Dim firstDataRow As Long
            firstDataRow = wsHdr + 1
            If ws.Cells(firstDataRow, healthCol).HasFormula Then
                ws.Cells(firstDataRow, healthCol).Copy
                ws.Range(ws.Cells(appendRow, healthCol), _
                         ws.Cells(appendRow + newRowCount - 1, healthCol)).PasteSpecial xlPasteFormulas
                Application.CutCopyMode = False
            End If
        End If

        ' Extend Data Validation for Lock? to new rows
        Dim lockCol As Long
        lockCol = 0
        If wsHeaderMap.exists(LCase(TIS_COL_LOCK)) Then lockCol = wsHeaderMap(LCase(TIS_COL_LOCK))
        If lockCol > 0 Then
            Dim firstDR As Long
            firstDR = wsHdr + 1
            ws.Cells(firstDR, lockCol).Copy
            ws.Range(ws.Cells(appendRow, lockCol), _
                     ws.Cells(appendRow + newRowCount - 1, lockCol)).PasteSpecial xlPasteValidation
            Application.CutCopyMode = False
        End If

        ' Copy Status validation to new rows
        If wsStatusCol > 0 Then
            Dim firstDR2 As Long
            firstDR2 = wsHdr + 1
            ws.Cells(firstDR2, wsStatusCol).Copy
            ws.Range(ws.Cells(appendRow, wsStatusCol), _
                     ws.Cells(appendRow + newRowCount - 1, wsStatusCol)).PasteSpecial xlPasteValidation
            Application.CutCopyMode = False
        End If
    End If ' newSystemRows.Count > 0

    ' 5. Batch apply orange fill + comments for changed cells
    If changedCells.Count > 0 Then
        Dim orangeRange As Range
        Dim ci As Long
        For ci = 1 To changedCells.Count
            Dim cellInfo As Variant
            cellInfo = changedCells(ci)
            If orangeRange Is Nothing Then
                Set orangeRange = ws.Cells(cellInfo(0), cellInfo(1))
            Else
                Set orangeRange = Union(orangeRange, ws.Cells(cellInfo(0), cellInfo(1)))
            End If
        Next ci
        If Not orangeRange Is Nothing Then
            orangeRange.Interior.Color = CLR_CHANGE_FILL
        End If

        ' Write comments
        For ci = 1 To changedCells.Count
            cellInfo = changedCells(ci)
            Dim cmtText As String
            cmtText = "[" & Format(Date, "YYYY-MM-DD") & "] Changed from: " & CStr(cellInfo(2))
            Dim tgtCell As Range
            Set tgtCell = ws.Cells(cellInfo(0), cellInfo(1))
            On Error Resume Next
            If tgtCell.Comment Is Nothing Then
                tgtCell.AddComment cmtText
            Else
                Dim existingCmt As String
                existingCmt = tgtCell.Comment.Text
                If Len(existingCmt) + Len(cmtText) < 1024 Then
                    tgtCell.Comment.Text existingCmt & vbLf & cmtText
                End If
            End If
            On Error GoTo 0
        Next ci
    End If

    ' Apply red border to Entity Code cells in existing WS rows that share a key with TIS duplicates
    If tisDupKeys.Count > 0 And wsECCol > 0 Then
        Dim wsDupR As Long
        For wsDupR = 1 To UBound(wsAllData, 1)
            Dim wsDupKey As String
            wsDupKey = LCase(Trim(CStr(wsAllData(wsDupR, wsSiteCol) & ""))) & "|" & _
                       LCase(Trim(CStr(wsAllData(wsDupR, wsECCol) & ""))) & "|" & _
                       LCase(Trim(CStr(wsAllData(wsDupR, wsETCol) & "")))
            If tisDupKeys.exists(wsDupKey) Then
                With ws.Cells(wsHdr + wsDupR, wsECCol).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = vbRed
                End With
            End If
        Next wsDupR
    End If
    Set tisDupKeys = Nothing

    ' 6. Sort the Working Sheet
    SortWorkingSheet ws

    ' 7. Rebuild Gantt and NIF so charts reflect the updated dates.
    '    Cross-module calls are guarded: failure is logged but does not abort the update.
    On Error Resume Next
    GanttBuilder.BuildGantt silent:=True, targetSheet:=ws
    If Err.Number <> 0 Then
        DebugLog "WARNING: GanttBuilder failed after TIS update: " & Err.Description
        Err.Clear
    End If
    NIF_Builder.BuildNIF silent:=True, targetSheet:=ws
    If Err.Number <> 0 Then
        DebugLog "WARNING: NIF_Builder failed after TIS update: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    GoTo UpdateCleanup

UpdateErrorHandler:
    DebugLog "UpdateWorkingSheetFromTIS ERROR: " & Err.Description & " (#" & Err.Number & ")"
    Application.ScreenUpdating = True
    MsgBox "Error updating Working Sheet: " & Err.Description, vbCritical, "TIS Update"

UpdateCleanup:
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'====================================================================
' SORT WORKING SHEET
' Sorts by Status (custom order) then project start date (earliest first).
' Active systems first, Cancelled last.
'====================================================================

Private Sub SortWorkingSheet(ws As Worksheet)
    Dim wsHdr As Long
    wsHdr = TISCommon.FindHeaderRow(ws)
    If wsHdr = 0 Then Exit Sub
    Dim wsLastRow As Long
    wsLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If wsLastRow <= wsHdr + 1 Then Exit Sub  ' Need at least 2 data rows to sort

    Dim wsMaxCol As Long
    wsMaxCol = ws.Cells(wsHdr, ws.Columns.Count).End(xlToLeft).Column

    ' Find Status column
    Dim statusCol As Long
    statusCol = TISCommon.FindHeaderCol(ws, wsHdr, TIS_COL_STATUS, wsMaxCol)
    If statusCol = 0 Then Exit Sub

    ' Find Our Date columns for MIN calculation
    Dim ourDateHeaders As Variant
    ourDateHeaders = Array(TIS_COL_OUR_SET, TIS_COL_OUR_SL1, TIS_COL_OUR_SL2, _
                           TIS_COL_OUR_SQ, TIS_COL_OUR_CONVS, TIS_COL_OUR_CONVF, _
                           TIS_COL_OUR_MRCLS, TIS_COL_OUR_MRCLF)

    ' Build MIN formula using a contiguous range reference.
    ' MIN(range) ignores blanks, whereas MIN(cell1,cell2,...) treats blanks as 0.
    Dim minFirstCol As Long, minLastCol As Long
    minFirstCol = 0: minLastCol = 0
    Dim oi As Long
    For oi = LBound(ourDateHeaders) To UBound(ourDateHeaders)
        Dim ourCol As Long
        ourCol = TISCommon.FindHeaderCol(ws, wsHdr, CStr(ourDateHeaders(oi)), wsMaxCol)
        If ourCol > 0 Then
            If minFirstCol = 0 Or ourCol < minFirstCol Then minFirstCol = ourCol
            If ourCol > minLastCol Then minLastCol = ourCol
        End If
    Next oi

    If minFirstCol = 0 Then Exit Sub

    ' Add temporary helper column for project start date
    Dim helperCol As Long
    helperCol = wsMaxCol + 1
    ws.Cells(wsHdr, helperCol).Value = "_SortHelper"

    ' Write MIN formula to first data row — range ref so blanks are ignored
    Dim firstDataRow As Long
    firstDataRow = wsHdr + 1
    ws.Cells(firstDataRow, helperCol).Formula = "=MIN(" & _
        TISCommon.ColLetter(minFirstCol) & firstDataRow & ":" & _
        TISCommon.ColLetter(minLastCol) & firstDataRow & ")"

    ' FillDown to all data rows
    If wsLastRow > firstDataRow Then
        ws.Cells(firstDataRow, helperCol).Copy
        ws.Range(ws.Cells(firstDataRow + 1, helperCol), ws.Cells(wsLastRow, helperCol)).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False
    End If

    ' Convert formulas to values for sorting (formulas can be slow to sort)
    Dim helperRange As Range
    Set helperRange = ws.Range(ws.Cells(firstDataRow, helperCol), ws.Cells(wsLastRow, helperCol))
    helperRange.Value = helperRange.Value

    ' Add temporary status sort helper (numeric: Active=1, Completed=2, On Hold=3, Non IQ=4, Cancelled=5)
    Dim statusHelperCol As Long
    statusHelperCol = helperCol + 1
    ws.Cells(wsHdr, statusHelperCol).Value = "_StatusSort"

    Dim sr As Long
    Dim statusArr As Variant
    statusArr = ws.Range(ws.Cells(firstDataRow, statusCol), ws.Cells(wsLastRow, statusCol)).Value
    Dim statusSortArr() As Variant
    ReDim statusSortArr(1 To wsLastRow - wsHdr, 1 To 1)
    For sr = 1 To UBound(statusArr, 1)
        Select Case LCase(Trim(CStr(statusArr(sr, 1) & "")))
            Case "active": statusSortArr(sr, 1) = 1
            Case "completed": statusSortArr(sr, 1) = 2
            Case "on hold": statusSortArr(sr, 1) = 3
            Case "non iq": statusSortArr(sr, 1) = 4
            Case "cancelled": statusSortArr(sr, 1) = 5
            Case Else: statusSortArr(sr, 1) = 6
        End Select
    Next sr
    ws.Range(ws.Cells(firstDataRow, statusHelperCol), ws.Cells(wsLastRow, statusHelperCol)).Value = statusSortArr

    ' Perform the sort
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(firstDataRow, statusHelperCol), ws.Cells(wsLastRow, statusHelperCol)), _
            SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=ws.Range(ws.Cells(firstDataRow, helperCol), ws.Cells(wsLastRow, helperCol)), _
            SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange ws.Range(ws.Cells(wsHdr, 1), ws.Cells(wsLastRow, statusHelperCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' Delete helper columns
    ws.Range(ws.Columns(helperCol), ws.Columns(statusHelperCol)).Delete
End Sub
