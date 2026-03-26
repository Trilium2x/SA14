Attribute VB_Name = "NIF_Builder"
'====================================================================
' NIF Builder Module - Rev14
'
' Builds NIF Assignment columns and HC Analyzer tables on the
' Working Sheet.  Designed to work with WorkfileBuilder_Rev10.
'
' Rev11 Changes (from Rev10.1):
'   - Removed SaveFilterSelection function and its call in BuildNIF
'   - Removed System Type dropdown from AddHCNeedSections
'   - Replaced single BuildAvailable/BuildGap with separate
'     New/Reused Available HC tables and New/Reused/Total Gap tables
'   - BuildAvailable now takes systemType string parameter and
'     hardcodes the NR filter instead of referencing a dropdown cell
'   - BuildGap replaced by BuildGapDirect (per-type) and
'     BuildTotalGap (sums New + Reused gaps)
'   - filterVal removed from BuildNIF and AddHCNeedSections
'
' Rev10.1 Changes (from Rev10):
'   - Fixed HC Need formula mismatch: GetHCMilestoneCodes now applies
'     GetLastWord() after GetShortAbbrev() so milestone codes match
'     the Gantt's short display values (e.g. "SL1" instead of "SET - SL1").
'
' Rev10 Changes (from Rev9):
'   - Removed duplicate CL(), GetShortAbbrev(), GetLastWord() that
'     now live in TISCommon
'   - Renamed terse helpers: SF->SafeFormulaWrite, FH->FormatSectionHeader,
'     FDH->FormatDateHeader, FTitle->FormatSectionTitle, FTotal->FormatTotalRow,
'     FHCData->FormatHCDataCell
'   - BuildRowKey static cache replaced with module-level m_RowKeyCache
'   - Debug.Print replaced with DebugLog throughout
'   - Cross-module references updated to Rev10
'
' Rev9 Changes (from Rev8):
'   - Shared utility functions (FindWorkingSheet, FindHeaderRow,
'     ShellSortVariantArray) consolidated into TISCommon module
'   - Variable naming cleanup in CollectUniqueGroups and GetHCMilestoneCodes
'
' Rev8 Changes (from Rev7):
'   - BuildNIF accepts optional targetSheet and sourceSheet parameters
'     so WorkfileBuilder can direct NIF to the new sheet and read
'     saved data from the old sheet during versioned rebuilds
'   - When sourceSheet is provided, Save operations read from sourceSheet
'     and Restore operations write to targetSheet (cross-sheet migration)
'   FUNCTIONAL
'   - Group names written as static VBA text (no UNIQUE helper row).
'   - Milestone codes inlined as string literals in formulas
'     (no hidden helper row).
'   - Layout simplified: title -> header -> data (no pre-rows).
'   - ClearExistingNIF walks down from first data row to find true
'     main-data end, then scans below for ALL content and clears it.
'     Handles both old (v3.2 helper rows) and new formats.
'
'   VISUAL
'   - Unified 3-point color scale for all HC tables: white (0) ->
'     yellow -> orange-red.  Same logic everywhere for consistency.
'   - Gap table keeps diverging red-white-green scale.
'   - HC data font size reduced to 8 for narrow Gantt columns.
'   - Date headers use compact "m/d" format, rotated 90 degrees.
'   - Section titles use consistent modern palette.
'   - Total rows get subtle gray background.
'====================================================================

Option Explicit

' ---- CONSTANTS ----
Private Const NIF_COL_GAP As Long = 10
Private Const NIF_EMPLOYEE_COUNT As Long = 5
Private Const HC_ROW_GAP As Long = 6
Private Const HC_SECTION_GAP As Long = 3       ' blank rows between HC sections
Private Const HC_FONT_SIZE As Long = 8          ' fits narrow Gantt columns

' AMAT brand-aligned colors (matches TISCommon THEME_ constants)
Private Const CLR_DARK As Long = 3349260        ' RGB(12, 27, 51)   - Deep Navy (THEME_BG)
Private Const CLR_BLUE As Long = 12491862       ' RGB(86, 156, 190) - AMAT Silver Lake Blue (THEME_ACCENT)
Private Const CLR_TEAL As Long = 6076462        ' RGB(46, 184, 92)  - Emerald (THEME_SUCCESS)
Private Const CLR_RED_TITLE As Long = 5264367   ' RGB(239, 83, 80)  - Coral Red (THEME_DANGER)
Private m_HeaderRow As Long
Private m_RowKeyCache As Object

Private Type NIFAssignment
    RowKey As String
    SlotIndex As Long
    NifName As String
    StartDate As Variant
    EndDate As Variant
End Type

Private Type HCRequirement
    GroupName As String
    MilestoneCode As String
    HCValue As Double
End Type

'====================================================================
' MAIN ENTRY POINT
'====================================================================

Public Sub BuildNIF(Optional silent As Boolean = False, _
                    Optional targetSheet As Worksheet = Nothing, _
                    Optional sourceSheet As Worksheet = Nothing)
    On Error GoTo ErrorHandler
    Dim startTime As Double
    Dim prevSU As Boolean, prevEE As Boolean, prevCalc As XlCalculation
    Dim ws As Worksheet, wsSave As Worksheet
    Dim ganttEndCol As Long, nifStartCol As Long
    Dim savedNIF() As NIFAssignment
    Dim savedHCNew() As HCRequirement, savedHCReused() As HCRequirement
    Dim nifCount As Long, hcNewCount As Long, hcReusedCount As Long

    ' Capture current state BEFORE any changes (safe defaults if error fires early)
    prevSU = Application.screenUpdating
    prevEE = Application.enableEvents
    prevCalc = Application.Calculation

    ' Reset row key cache for this build
    Set m_RowKeyCache = Nothing

    ' Use provided target sheet or find latest Working Sheet
    If Not targetSheet Is Nothing Then
        Set ws = targetSheet
    Else
        Set ws = FindWorkingSheet()
    End If
    If ws Is Nothing Then
        If Not silent Then MsgBox "No Working Sheet found.", vbExclamation
        Exit Sub
    End If

    ' Use sourceSheet for saving data (cross-sheet migration), or same sheet
    If Not sourceSheet Is Nothing Then
        Set wsSave = sourceSheet
    Else
        Set wsSave = ws
    End If

    m_HeaderRow = FindHeaderRow(ws)
    If m_HeaderRow = 0 Then
        If Not silent Then MsgBox "Could not find header row.", vbExclamation
        Exit Sub
    End If
    ganttEndCol = FindGanttEndColumn(ws)

    ' If target sheet has no Gantt yet, try to get ganttEndCol from source
    If ganttEndCol = 0 And Not sourceSheet Is Nothing Then
        Dim srcHeaderRow As Long
        srcHeaderRow = FindHeaderRow(sourceSheet)
        If srcHeaderRow > 0 Then ganttEndCol = FindGanttEndColumn(sourceSheet)
    End If

    If ganttEndCol = 0 Then
        If Not silent Then MsgBox "No Gantt chart found.", vbExclamation
        Exit Sub
    End If

    startTime = Timer
    Application.screenUpdating = False
    Application.enableEvents = False
    Application.Calculation = xlCalculationManual

    nifStartCol = ganttEndCol + NIF_COL_GAP

    ' Save from source sheet (old sheet during migration, or same sheet for in-place rebuild)
    Dim saveHeaderRow As Long
    saveHeaderRow = FindHeaderRow(wsSave)
    DebugLog "BuildNIF: wsSave=" & wsSave.name & " saveHeaderRow=" & saveHeaderRow
    If saveHeaderRow > 0 Then
        Dim saveGanttEnd As Long
        saveGanttEnd = FindGanttEndColumn(wsSave)
        DebugLog "BuildNIF: saveGanttEnd=" & saveGanttEnd
        If saveGanttEnd > 0 Then
            Dim saveNifStart As Long
            saveNifStart = saveGanttEnd + NIF_COL_GAP
            DebugLog "BuildNIF: saveNifStart=" & saveNifStart & _
                        " cell at nifStart=" & CStr(wsSave.Cells(m_HeaderRow, saveNifStart).Value)
            SaveNIFData wsSave, saveNifStart, savedNIF, nifCount
            SaveHCRequirements wsSave, saveGanttEnd, "New", savedHCNew, hcNewCount
            SaveHCRequirements wsSave, saveGanttEnd, "Reused", savedHCReused, hcReusedCount
        Else
            DebugLog "BuildNIF: EXIT save - no Gantt on source sheet"
        End If
    Else
        DebugLog "BuildNIF: EXIT save - no header row on source sheet"
    End If

    ' Clear and rebuild on target sheet
    ClearExistingNIF ws, ganttEndCol
    nifStartCol = ganttEndCol + NIF_COL_GAP
    AddNIFColumns ws, nifStartCol
    RestoreNIFData ws, nifStartCol, savedNIF, nifCount
    AddHCNeedSections ws, ganttEndCol, nifStartCol, _
                       savedHCNew, hcNewCount, savedHCReused, hcReusedCount

    Application.Calculation = xlCalculationAutomatic
    ws.Calculate

    ' Auto-create Gantt/HC toggle button
    On Error Resume Next
    HCHeatmap.CreateSegmentedToggle targetSheet:=ws
    On Error GoTo 0

    If Not silent Then
        Application.screenUpdating = True
        MsgBox "NIF Builder " & TIS_VERSION & " completed!" & vbCrLf & _
               "Time: " & Format(Timer - startTime, "0.00") & "s", vbInformation
        Application.screenUpdating = False
    End If
    GoTo Cleanup

ErrorHandler:
    If Not silent Then
        MsgBox "Error in BuildNIF: " & Err.Description & vbCrLf & _
               "Error #: " & Err.Number, vbCritical
    End If
Cleanup:
    Application.Calculation = prevCalc
    Application.enableEvents = prevEE
    Application.screenUpdating = prevSU
End Sub

'====================================================================
' SHEET / COLUMN FINDERS
'====================================================================

Private Function FindGanttEndColumn(ws As Worksheet) As Long
    Dim j As Long, lc As Long
    lc = ws.Cells(m_HeaderRow, ws.Columns.Count).End(xlToLeft).Column
    For j = lc To 1 Step -1
        If IsDate(ws.Cells(m_HeaderRow, j).Value) Then
            FindGanttEndColumn = j: Exit Function
        End If
    Next j
    FindGanttEndColumn = 0
End Function

Private Function FindGanttStartColumn(ws As Worksheet) As Long
    Dim j As Long, lc As Long
    lc = ws.Cells(m_HeaderRow, ws.Columns.Count).End(xlToLeft).Column
    For j = 1 To lc
        If IsDate(ws.Cells(m_HeaderRow, j).Value) Then
            FindGanttStartColumn = j: Exit Function
        End If
    Next j
    FindGanttStartColumn = 0
End Function

Private Function FindColumnByHeader(ws As Worksheet, hn As String) As Long
    Dim j As Long, v As String, lc As Long
    lc = ws.Cells(m_HeaderRow, ws.Columns.Count).End(xlToLeft).Column
    For j = 1 To lc
        v = LCase(Trim(Replace(Replace(CStr(ws.Cells(m_HeaderRow, j).Value), vbLf, ""), vbCr, "")))
        If v = LCase(hn) Then FindColumnByHeader = j: Exit Function
    Next j
    For j = 1 To lc
        v = LCase(Trim(Replace(Replace(CStr(ws.Cells(m_HeaderRow, j).Value), vbLf, ""), vbCr, "")))
        If InStr(1, v, LCase(hn), vbTextCompare) > 0 Then
            FindColumnByHeader = j: Exit Function
        End If
    Next j
    FindColumnByHeader = 0
End Function

'====================================================================
' UTILITY FUNCTIONS
'====================================================================

Private Function BuildRowKey(ws As Worksheet, row As Long) As String
    ' Stable composite key: Site|Entity Code|Event Type
    ' Must match across rebuilds regardless of sort order or volatile columns
    Dim siteCol As Long, ecCol As Long, etCol As Long
    Dim j As Long, hv As String
    Dim lastCol As Long

    ' Initialize cache dictionary on first use
    If m_RowKeyCache Is Nothing Then Set m_RowKeyCache = CreateObject("Scripting.Dictionary")

    ' Cache header positions per sheet (avoid rescanning every row)
    If Not m_RowKeyCache.exists(ws.Name) Then
        siteCol = 0: ecCol = 0: etCol = 0
        lastCol = ws.Cells(m_HeaderRow, ws.Columns.Count).End(xlToLeft).Column
        If lastCol > 50 Then lastCol = 50

        For j = 1 To lastCol
            hv = LCase(Trim(Replace(Replace(CStr(ws.Cells(m_HeaderRow, j).Value), vbLf, ""), vbCr, "")))
            If hv = "site" And siteCol = 0 Then siteCol = j
            If hv = "entity code" And ecCol = 0 Then ecCol = j
            If hv = "event type" And etCol = 0 Then etCol = j
        Next j

        ' Fallback: use CEID if Entity Code not found
        If ecCol = 0 Then
            For j = 1 To lastCol
                hv = LCase(Trim(Replace(Replace(CStr(ws.Cells(m_HeaderRow, j).Value), vbLf, ""), vbCr, "")))
                If hv = "ceid" Then ecCol = j: Exit For
            Next j
        End If

        m_RowKeyCache(ws.Name) = Array(siteCol, ecCol, etCol)
        DebugLog "BuildRowKey: cached for [" & ws.Name & "] siteCol=" & siteCol & " ecCol=" & ecCol & " etCol=" & etCol
    End If

    Dim cached As Variant
    cached = m_RowKeyCache(ws.Name)
    siteCol = cached(0): ecCol = cached(1): etCol = cached(2)

    Dim k As String
    k = ""
    If siteCol > 0 Then k = k & Trim(CStr(ws.Cells(row, siteCol).Value))
    k = k & "|"
    If ecCol > 0 Then k = k & Trim(CStr(ws.Cells(row, ecCol).Value))
    k = k & "|"
    If etCol > 0 Then k = k & Trim(CStr(ws.Cells(row, etCol).Value))
    BuildRowKey = k
End Function

Private Function FindMainDataEnd(ws As Worksheet) As Long
    ' Use ListObject boundary when available (single property access, zero loops)
    If ws.ListObjects.Count > 0 Then
        Dim tbl As ListObject: Set tbl = ws.ListObjects(1)
        FindMainDataEnd = tbl.Range.row + tbl.Range.Rows.Count - 1
        Exit Function
    End If
    ' Fallback: walk contiguous non-empty cells in column 1 from first data row.
    Dim r As Long
    FindMainDataEnd = m_HeaderRow
    For r = m_HeaderRow + 1 To m_HeaderRow + 10000
        If IsEmpty(ws.Cells(r, 1).Value) Or _
           Trim(CStr(ws.Cells(r, 1).Value)) = "" Then
            FindMainDataEnd = r - 1
            Exit Function
        End If
    Next r
    FindMainDataEnd = r - 1
End Function

' Safe formula write: tries .Formula, then .Formula2 for dynamic arrays
Private Sub SafeFormulaWrite(ws As Worksheet, row As Long, col As Long, f As String)
    On Error Resume Next
    ws.Cells(row, col).formula = f
    If Err.Number <> 0 Then
        Err.Clear
        ws.Cells(row, col).Formula2 = f
        If Err.Number <> 0 Then
            Err.Clear
            ws.Cells(row, col).Value = "ERR"
        End If
    End If
    On Error GoTo 0
End Sub

'====================================================================
' FORMATTING HELPERS
'====================================================================

' Section header cell (white text on colored background)
Private Sub FormatSectionHeader(ws As Worksheet, row As Long, col As Long, t As String, bg As Long)
    With ws.Cells(row, col)
        .Value = t: .Font.Bold = True: .Font.Size = 8
        .Font.Color = RGB(255, 255, 255): .Interior.Color = bg
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
End Sub

' Date header (rotated, compact)
Private Sub FormatDateHeader(ws As Worksheet, row As Long, col As Long)
    With ws.Cells(row, col)
        .Font.Size = 6: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(55, 75, 95)
        .HorizontalAlignment = xlCenter: .Orientation = 90
    End With
End Sub

' Section title (left-aligned, large, colored)
Private Sub FormatSectionTitle(ws As Worksheet, row As Long, col As Long, t As String, clr As Long)
    With ws.Cells(row, col)
        .Value = t: .Font.Bold = True: .Font.Size = 11: .Font.Color = clr
    End With
End Sub

' Total row formatting
Private Sub FormatTotalRow(ws As Worksheet, row As Long, sc As Long, ec As Long)
    With ws.Range(ws.Cells(row, sc), ws.Cells(row, ec))
        .Font.Bold = True: .Font.Size = HC_FONT_SIZE
        .Interior.Color = RGB(242, 242, 242)   ' subtle gray
        .HorizontalAlignment = xlCenter
    End With
    ws.Cells(row, sc).Value = "TOTAL"
    ws.Cells(row, sc).Font.Color = CLR_DARK
End Sub

' HC data cell formatting (applies to a single cell - used for edge cases)
Private Sub FormatHCDataCell(ws As Worksheet, row As Long, col As Long, nf As String)
    With ws.Cells(row, col)
        .NumberFormat = nf
        .HorizontalAlignment = xlCenter
        .Font.Size = HC_FONT_SIZE
    End With
End Sub

' BATCH: Format entire HC data range at once (replaces per-cell FormatHCDataCell calls)
Private Sub BatchFormatHCData(ws As Worksheet, startRow As Long, endRow As Long, _
    startCol As Long, endCol As Long, nf As String)
    If endRow < startRow Or endCol < startCol Then Exit Sub
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(endRow, endCol))
        .NumberFormat = nf
        .HorizontalAlignment = xlCenter
        .Font.Size = HC_FONT_SIZE
    End With
End Sub

' BATCH: Copy Gantt date headers to an HC table header row and format in bulk
Private Sub BatchWriteDateHeaders(ws As Worksheet, hdrRow As Long, _
    gsc As Long, gec As Long)
    ' Bulk copy header values (single Range.Value read + write)
    Dim nCols As Long: nCols = gec - gsc + 1
    If nCols < 1 Then Exit Sub
    Dim hdrVals As Variant
    hdrVals = ws.Range(ws.Cells(m_HeaderRow, gsc), ws.Cells(m_HeaderRow, gec)).Value
    ws.Range(ws.Cells(hdrRow, gsc), ws.Cells(hdrRow, gec)).Value = hdrVals
    ' Bulk format entire date header range
    With ws.Range(ws.Cells(hdrRow, gsc), ws.Cells(hdrRow, gec))
        .NumberFormat = "m/d"
        .Font.Size = 6: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(55, 75, 95)
        .HorizontalAlignment = xlCenter: .Orientation = 90
    End With
End Sub

'====================================================================
' COLLECT UNIQUE GROUPS (sorted, via VBA dictionary)
'====================================================================

Private Function CollectUniqueGroups(ws As Worksheet, gcol As Long, _
                                      lastDataRow As Long) As Variant
    Dim groupDict As Object, cellValue As String
    Dim keys As Variant, result() As String, i As Long

    Set groupDict = CreateObject("Scripting.Dictionary")

    ' Single bulk read of group column into array
    Dim nRows As Long
    nRows = lastDataRow - m_HeaderRow
    If nRows < 1 Then
        CollectUniqueGroups = Array()
        Exit Function
    End If
    Dim grpArr As Variant
    grpArr = ws.Range(ws.Cells(m_HeaderRow + 1, gcol), ws.Cells(lastDataRow, gcol)).Value
    If Not IsArray(grpArr) Then
        ' Single-cell read returns scalar
        Dim tmp As Variant: tmp = grpArr
        ReDim grpArr(1 To 1, 1 To 1): grpArr(1, 1) = tmp
    End If

    Dim ri As Long
    For ri = 1 To UBound(grpArr, 1)
        If Not IsError(grpArr(ri, 1)) Then
            cellValue = Trim(CStr(grpArr(ri, 1)))
            If cellValue <> "" And Not groupDict.exists(cellValue) Then groupDict.Add cellValue, groupDict.Count + 1
        End If
    Next ri

    If groupDict.Count = 0 Then
        CollectUniqueGroups = Array()
        Exit Function
    End If

    keys = groupDict.keys
    ShellSortVariantArray keys

    ReDim result(1 To groupDict.Count)
    For i = 0 To UBound(keys)
        result(i + 1) = CStr(keys(i))
    Next i

    CollectUniqueGroups = result
End Function

'====================================================================
' GET HC MILESTONE CODES (from Definitions sheet)
'====================================================================

Private Function GetHCMilestoneCodes() As Collection
    Dim codes As New Collection, wsDef As Worksheet, defLastRow As Long, defData As Variant
    Dim i As Long, flagText As String, groupText As String, tokens As Variant, tokenValue As Variant
    Dim letterCode As String, milestoneNum As Long, milestoneGroups As Object, milestoneNames As Object
    Dim sortedKeys As Variant, idx As Long, abbrev As String, shortCode As String
    Dim isDuplicate As Boolean, item As Variant, excludeList As String

    excludeList = "|PF|DC|DM|IQ|SDD|"

    On Error Resume Next
    Set wsDef = ThisWorkbook.Sheets("Definitions")
    On Error GoTo 0

    If wsDef Is Nothing Then GoTo DefaultCodes
    defLastRow = wsDef.Cells(wsDef.Rows.Count, 1).End(xlUp).row
    If defLastRow < 2 Then GoTo DefaultCodes

    defData = wsDef.Range(wsDef.Cells(1, 1), wsDef.Cells(defLastRow, 7)).Value
    Set milestoneGroups = CreateObject("Scripting.Dictionary")
    Set milestoneNames = CreateObject("Scripting.Dictionary")

    For i = 2 To UBound(defData, 1)
        flagText = Trim(CStr(defData(i, 6))): groupText = Trim(CStr(defData(i, 7)))
        If flagText <> "" Then
            tokens = Split(flagText, "|")
            For Each tokenValue In tokens
                tokenValue = UCase(Trim(tokenValue))
                If Len(tokenValue) >= 2 And IsNumeric(Mid(tokenValue, 2)) Then
                    letterCode = Left(tokenValue, 1): milestoneNum = CLng(Mid(tokenValue, 2))
                    If Not milestoneGroups.exists(letterCode) Then Set milestoneGroups(letterCode) = CreateObject("Scripting.Dictionary")
                    milestoneGroups(letterCode)(milestoneNum) = True
                    If milestoneNum = 1 And groupText <> "" Then milestoneNames(letterCode) = groupText
                End If
            Next tokenValue
        End If
    Next i

    If milestoneGroups.Count > 0 Then
        sortedKeys = milestoneGroups.keys: ShellSortVariantArray sortedKeys
        For idx = LBound(sortedKeys) To UBound(sortedKeys)
            letterCode = CStr(sortedKeys(idx))
            If milestoneGroups(letterCode).exists(1) And milestoneGroups(letterCode).exists(2) Then
                abbrev = ""
                If milestoneNames.exists(letterCode) Then abbrev = UCase(milestoneNames(letterCode))
                If abbrev = "" Then abbrev = UCase(letterCode)
                shortCode = GetLastWord(GetShortAbbrev(abbrev))
                If InStr(1, excludeList, "|" & shortCode & "|", vbTextCompare) = 0 Then
                    isDuplicate = False
                    For Each item In codes
                        If item = shortCode Then isDuplicate = True: Exit For
                    Next item
                    If Not isDuplicate Then codes.Add shortCode
                End If
            End If
        Next idx
    End If

    If codes.Count = 0 Then GoTo DefaultCodes

    ' Append MRCL as an HC milestone (after dynamic milestones)
    Dim hasMRCL As Boolean
    hasMRCL = False
    For Each item In codes
        If CStr(item) = "MRCL" Then hasMRCL = True: Exit For
    Next item
    If Not hasMRCL Then codes.Add "MRCL"

    Set GetHCMilestoneCodes = codes
    Exit Function

DefaultCodes:
    Set codes = New Collection
    codes.Add "CV": codes.Add "SET": codes.Add "SL1": codes.Add "SL2": codes.Add "SQ"
    Set GetHCMilestoneCodes = codes
End Function

'====================================================================
' CLEAR / SAVE / RESTORE
'====================================================================

Private Sub ClearExistingNIF(ws As Worksheet, gec As Long)
    Dim mainDataEnd As Long, clearEnd As Long
    Dim r As Long, sheetLastCol As Long

    ' Clear NIF columns (right of Gantt)
    On Error Resume Next
    ws.Range(ws.Cells(1, gec + 1), _
             ws.Cells(ws.Rows.Count, gec + NIF_COL_GAP + NIF_EMPLOYEE_COUNT * 3 + 5)).Clear
    On Error GoTo 0

    ' Find true main data end (contiguous data in column 1)
    mainDataEnd = FindMainDataEnd(ws)

    ' Find last row of ANY content below main data using UsedRange boundary
    clearEnd = mainDataEnd
    sheetLastCol = ws.Cells(m_HeaderRow, ws.Columns.Count).End(xlToLeft).Column
    If sheetLastCol < gec Then sheetLastCol = gec
    Dim usedEnd As Long
    usedEnd = ws.UsedRange.row + ws.UsedRange.Rows.Count - 1
    If usedEnd < mainDataEnd Then usedEnd = mainDataEnd
    For r = mainDataEnd + 1 To usedEnd
        If Application.WorksheetFunction.CountA( _
           ws.Range(ws.Cells(r, 1), ws.Cells(r, sheetLastCol))) > 0 Then
            clearEnd = r
        ElseIf r > clearEnd + 10 Then
            Exit For
        End If
    Next r

    ' Clear everything below main data
    If clearEnd > mainDataEnd Then
        On Error Resume Next
        ws.Range(ws.Cells(mainDataEnd + 1, 1), _
                 ws.Cells(clearEnd + 5, sheetLastCol)).Clear
        On Error GoTo 0
    End If
End Sub

Private Sub SaveNIFData(ws As Worksheet, nsc As Long, s() As NIFAssignment, cnt As Long)
    Dim fd As Long, ld As Long, r As Long, i As Long
    Dim nc As Long, sc As Long, ec As Long, nv As String
    Dim actualNsc As Long, j As Long, hv As String
    cnt = 0
    fd = m_HeaderRow + 1
    ld = FindMainDataEnd(ws)
    DebugLog "SaveNIFData: ws=" & ws.name & " m_HeaderRow=" & m_HeaderRow & _
                " fd=" & fd & " ld=" & ld & " nsc(calc)=" & nsc
    If ld < fd Then
        DebugLog "SaveNIFData: EXIT - ld < fd"
        Exit Sub
    End If

    ' Find actual NIF start column by scanning for "NIF1" header
    actualNsc = 0
    On Error Resume Next
    ' Try calculated position first
    If LCase(Left(CStr(ws.Cells(m_HeaderRow, nsc).Value), 3)) = "nif" Then
        actualNsc = nsc
        DebugLog "SaveNIFData: found NIF at calculated pos " & nsc
    Else
        DebugLog "SaveNIFData: NOT at calc pos, cell=" & CStr(ws.Cells(m_HeaderRow, nsc).Value) & " - scanning..."
        ' Scan from Gantt area rightward for NIF1 header
        Dim searchStart As Long
        searchStart = Application.Max(nsc - 20, 1)
        For j = searchStart To ws.Cells(m_HeaderRow, ws.Columns.Count).End(xlToLeft).Column
            hv = LCase(Trim(CStr(ws.Cells(m_HeaderRow, j).Value)))
            If hv = "nif1" Then
                actualNsc = j
                DebugLog "SaveNIFData: found NIF1 at col " & j & " (by scan)"
                Exit For
            End If
        Next j
    End If
    On Error GoTo 0

    If actualNsc = 0 Then
        DebugLog "SaveNIFData: EXIT - NIF1 header not found anywhere"
        Exit Sub
    End If

    ReDim s(1 To (ld - fd + 1) * NIF_EMPLOYEE_COUNT)
    ' Debug: print first 3 row keys for comparison
    Dim dbgR As Long
    For dbgR = fd To Application.Min(fd + 2, ld)
        DebugLog "  SaveKey row" & dbgR & ": [" & BuildRowKey(ws, dbgR) & "]"
    Next dbgR
    For r = fd To ld
        For i = 1 To NIF_EMPLOYEE_COUNT
            nc = actualNsc + (i - 1) * 3: sc = nc + 1: ec = nc + 2
            nv = Trim(CStr(ws.Cells(r, nc).Value))
            If nv <> "" Or Not IsEmpty(ws.Cells(r, sc).Value) Or _
               Not IsEmpty(ws.Cells(r, ec).Value) Then
                cnt = cnt + 1
                s(cnt).RowKey = BuildRowKey(ws, r)
                s(cnt).SlotIndex = i
                s(cnt).NifName = nv
                s(cnt).StartDate = ws.Cells(r, sc).Value
                s(cnt).EndDate = ws.Cells(r, ec).Value
            End If
        Next i
    Next r
    DebugLog "SaveNIFData: saved " & cnt & " NIF assignments"
End Sub

Private Sub RestoreNIFData(ws As Worksheet, nsc As Long, s() As NIFAssignment, cnt As Long)
    Dim fd As Long, ld As Long, r As Long, k As Long
    Dim nc As Long, sc As Long, ec As Long, rk As String
    Dim matchCount As Long
    matchCount = 0
    DebugLog "RestoreNIFData: ws=" & ws.name & " nsc=" & nsc & " cnt=" & cnt
    If cnt = 0 Then
        DebugLog "RestoreNIFData: EXIT - cnt=0 (nothing to restore)"
        Exit Sub
    End If
    fd = m_HeaderRow + 1
    ld = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' Build dictionary: RowKey -> Collection of indices into s()
    Dim keyMap As Object
    Set keyMap = CreateObject("Scripting.Dictionary")
    For k = 1 To cnt
        If Not keyMap.exists(s(k).RowKey) Then
            Set keyMap(s(k).RowKey) = New Collection
        End If
        keyMap(s(k).RowKey).Add k
    Next k

    ' Debug: print first 3 row keys for comparison
    Dim dbgR2 As Long
    For dbgR2 = fd To Application.Min(fd + 2, ld)
        DebugLog "  RestoreKey row" & dbgR2 & ": [" & BuildRowKey(ws, dbgR2) & "]"
    Next dbgR2

    ' O(n) restore: single pass over rows with O(1) key lookup
    Dim items As Collection
    Dim idx As Variant
    For r = fd To ld
        rk = BuildRowKey(ws, r)
        If keyMap.exists(rk) Then
            Set items = keyMap(rk)
            For Each idx In items
                k = CLng(idx)
                nc = nsc + (s(k).SlotIndex - 1) * 3: sc = nc + 1: ec = nc + 2
                If s(k).NifName <> "" Then ws.Cells(r, nc).Value = s(k).NifName
                If IsDate(s(k).StartDate) Then ws.Cells(r, sc).Value = s(k).StartDate
                If IsDate(s(k).EndDate) Then ws.Cells(r, ec).Value = s(k).EndDate
                matchCount = matchCount + 1
            Next idx
        End If
    Next r
    DebugLog "RestoreNIFData: matched " & matchCount & " of " & cnt & " assignments"
End Sub

Private Sub SaveHCRequirements(ws As Worksheet, gec As Long, st As String, _
                                sv() As HCRequirement, hcc As Long)
    Dim r As Long, c As Long, fh As Long, gs As Long, hr As Long
    Dim gc As Long, lu As Long
    Dim mCols As Collection, gn As String, hv As String, mc As Variant

    hcc = 0
    gs = FindGanttStartColumn(ws)
    If gs = 0 Then Exit Sub
    lu = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lu <= m_HeaderRow Then Exit Sub

    ' Use UsedRange for reliable scan boundary
    Dim scanEnd As Long
    scanEnd = ws.UsedRange.row + ws.UsedRange.Rows.Count - 1
    If scanEnd < lu Then scanEnd = lu

    fh = 0
    Dim foundTitle As Boolean: foundTitle = False
    For r = m_HeaderRow + 1 To scanEnd
        For c = 1 To gs
            If InStr(1, CStr(ws.Cells(r, c).Value), st & " Systems - HC Need", vbTextCompare) > 0 Then
                fh = r: foundTitle = True: Exit For
            End If
        Next c
        If foundTitle Then Exit For
    Next r
    If fh = 0 Then Exit Sub

    ' Header is right after title (works for both v3.2 and Rev6 layouts)
    hr = fh + 1
    gc = 0
    Set mCols = New Collection
    For c = 1 To gs - 1
        hv = Trim(CStr(ws.Cells(hr, c).Value))
        If LCase(hv) = "group" Then gc = c
        If Right(LCase(hv), 3) = " hc" Then _
            mCols.Add Array(c, UCase(Trim(Left(hv, Len(hv) - 3))))
    Next c
    If gc = 0 Or mCols.Count = 0 Then Exit Sub

    ReDim sv(1 To 200)
    r = hr + 1
    Do While r <= lu + 100
        gn = Trim(CStr(ws.Cells(r, gc).Value))
        If gn = "" Or UCase(gn) = "TOTAL" Then Exit Do
        For Each mc In mCols
            hcc = hcc + 1
            If hcc > UBound(sv) Then ReDim Preserve sv(1 To hcc + 100)
            sv(hcc).GroupName = gn
            sv(hcc).MilestoneCode = CStr(mc(1))
            sv(hcc).HCValue = Val(CStr(ws.Cells(r, CLng(mc(0))).Value))
        Next mc
        r = r + 1
    Loop
End Sub

Private Sub RestoreHCValues(ws As Worksheet, dsr As Long, gdc As Long, _
                             msc As Long, hcCodes As Collection, _
                             sv() As HCRequirement, hcc As Long)
    Dim r As Long, k As Long, mo As Long, gn As String, c As Variant
    If hcc = 0 Then Exit Sub
    Application.Calculation = xlCalculationAutomatic
    ws.Calculate
    Application.Calculation = xlCalculationManual
    r = dsr
    Do While Trim(CStr(ws.Cells(r, gdc).Value)) <> "" And _
             UCase(Trim(CStr(ws.Cells(r, gdc).Value))) <> "TOTAL"
        gn = Trim(CStr(ws.Cells(r, gdc).Value))
        mo = 0
        For Each c In hcCodes
            For k = 1 To hcc
                If UCase(sv(k).GroupName) = UCase(gn) And _
                   UCase(sv(k).MilestoneCode) = UCase(CStr(c)) Then
                    ws.Cells(r, msc + mo).Value = sv(k).HCValue
                    Exit For
                End If
            Next k
            mo = mo + 1
        Next c
        r = r + 1
    Loop
End Sub

'====================================================================
' ADD NIF COLUMNS (NIF1..5, Start1..5, End1..5, Staffed)
'====================================================================

Private Sub AddNIFColumns(ws As Worksheet, startCol As Long)
    Dim cc As Long, i As Long, ld As Long, fd As Long, sc As Long, gc As Long
    Dim hcCodes As Collection

    ld = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    fd = m_HeaderRow + 1

    For gc = startCol - NIF_COL_GAP To startCol - 1
        If gc > 0 Then ws.Columns(gc).ColumnWidth = 0.8
    Next gc

    ws.Cells(m_HeaderRow - 1, startCol).Value = "NIF Assignment Plan"
    ws.Cells(m_HeaderRow - 1, startCol).Font.Bold = True
    ws.Cells(m_HeaderRow - 1, startCol).Font.Size = 11
    ws.Cells(m_HeaderRow - 1, startCol).Font.Color = CLR_DARK

    cc = startCol
    For i = 1 To NIF_EMPLOYEE_COUNT
        FormatSectionHeader ws, m_HeaderRow, cc, "NIF" & i, CLR_DARK
        ws.Columns(cc).ColumnWidth = 12: cc = cc + 1
        FormatSectionHeader ws, m_HeaderRow, cc, "Start" & i, CLR_DARK
        ws.Columns(cc).NumberFormat = "mm/dd/yy"
        ws.Columns(cc).ColumnWidth = 10: cc = cc + 1
        FormatSectionHeader ws, m_HeaderRow, cc, "End" & i, CLR_DARK
        ws.Columns(cc).NumberFormat = "mm/dd/yy"
        ws.Columns(cc).ColumnWidth = 10: cc = cc + 1
    Next i

    sc = cc
    FormatSectionHeader ws, m_HeaderRow, sc, "Staffed", CLR_TEAL
    ws.Columns(sc).ColumnWidth = 9

    Set hcCodes = GetHCMilestoneCodes()
    WriteStaffedFormulas ws, startCol, sc, fd, ld, hcCodes
    ApplyNIFOverlapCF ws, startCol, fd, ld
    ApplyStaffedCF ws, sc, fd, ld
End Sub

'====================================================================
' NIF OVERLAP CF
'====================================================================

Private Sub ApplyNIFOverlapCF(ws As Worksheet, nsc As Long, fd As Long, ld As Long)
    Dim i As Long, j As Long, rng As Range, fc As FormatCondition
    Dim f As String, op As String, sch As String
    Dim nc As Long, sc As Long, ec As Long
    Dim onc As Long, osc As Long, oec As Long
    Dim rr As String, nr As String, sr As String, er As String

    If ld < fd Then Exit Sub  ' No data rows

    rr = CStr(fd)
    For i = 1 To NIF_EMPLOYEE_COUNT
        nc = nsc + (i - 1) * 3: sc = nc + 1: ec = nc + 2
        Set rng = ws.Range(ws.Cells(fd, nc), ws.Cells(ld, ec))
        On Error Resume Next
        rng.FormatConditions.Delete
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        op = ""
        For j = 1 To NIF_EMPLOYEE_COUNT
            If j <> i Then
                onc = nsc + (j - 1) * 3: osc = onc + 1: oec = onc + 2
                nr = "$" & ColLetter(onc) & "$" & fd & ":$" & ColLetter(onc) & "$" & ld
                sr = "$" & ColLetter(osc) & "$" & fd & ":$" & ColLetter(osc) & "$" & ld
                er = "$" & ColLetter(oec) & "$" & fd & ":$" & ColLetter(oec) & "$" & ld
                If op <> "" Then op = op & "+"
                op = op & "SUMPRODUCT(--(" & nr & "=$" & ColLetter(nc) & rr & ")," & _
                     "--(" & nr & "<>""""),--(" & sr & "<=$" & ColLetter(ec) & rr & ")," & _
                     "--(" & er & ">=$" & ColLetter(sc) & rr & "))"
            End If
        Next j
        nr = "$" & ColLetter(nc) & "$" & fd & ":$" & ColLetter(nc) & "$" & ld
        sr = "$" & ColLetter(sc) & "$" & fd & ":$" & ColLetter(sc) & "$" & ld
        er = "$" & ColLetter(ec) & "$" & fd & ":$" & ColLetter(ec) & "$" & ld
        sch = "SUMPRODUCT(--(" & nr & "=$" & ColLetter(nc) & rr & ")," & _
              "--(" & nr & "<>""""),--(" & sr & "<=$" & ColLetter(ec) & rr & ")," & _
              "--(" & er & ">=$" & ColLetter(sc) & rr & "))"
        f = "=AND($" & ColLetter(nc) & rr & "<>"""",OR((" & op & ")>0,(" & sch & ")>1))"
        On Error Resume Next
        Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=f)
        If Err.Number = 0 And Not fc Is Nothing Then
            fc.Interior.Color = RGB(254, 232, 232)
            fc.Font.Color = RGB(185, 28, 28)
            fc.Font.Bold = True
            fc.StopIfTrue = False
        Else
            DebugLog "NIF Overlap CF failed for slot " & i & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next i
End Sub

'====================================================================
' STAFFED FORMULA + CF
'====================================================================

Private Sub WriteStaffedFormulas(ws As Worksheet, nsc As Long, stc As Long, _
                                  fd As Long, ld As Long, hcCodes As Collection)
    Dim gsc As Long, gec As Long, wr As String, gr As String
    Dim nc As String, rc As String, f As String, ir As String
    Dim i As Long, nCol As Long, sCol As Long, eCol As Long
    Dim rs As String, c As Variant

    gsc = FindGanttStartColumn(ws)
    gec = FindGanttEndColumn(ws)
    If gsc = 0 Or gec = 0 Then Exit Sub

    rs = CStr(fd)
    wr = ColLetter(gsc) & "$" & m_HeaderRow & ":" & ColLetter(gec) & "$" & m_HeaderRow
    gr = "$" & ColLetter(gsc) & rs & ":$" & ColLetter(gec) & rs

    rc = ""
    For Each c In hcCodes
        If rc <> "" Then rc = rc & "+"
        rc = rc & "(" & gr & "=""" & CStr(c) & """)"
    Next c
    ir = "--((" & rc & ")>0)"

    nc = ""
    For i = 1 To NIF_EMPLOYEE_COUNT
        nCol = nsc + (i - 1) * 3: sCol = nCol + 1: eCol = nCol + 2
        If nc <> "" Then nc = nc & "+"
        nc = nc & "($" & ColLetter(nCol) & rs & "<>"""")*" & _
             "($" & ColLetter(sCol) & rs & "<=" & wr & "+6)*" & _
             "($" & ColLetter(eCol) & rs & ">=" & wr & ")"
    Next i

    f = "=IF(SUMPRODUCT(" & ir & ")=0,""N/A""," & _
        "IF(SUMPRODUCT(" & ir & ",--(" & nc & "=0))=0,""Yes"",""No""))"

    On Error Resume Next
    ws.Cells(fd, stc).formula = f
    If Err.Number <> 0 Then
        Err.Clear
        ws.Cells(fd, stc).Value = "N/A"
    Else
        If ld > fd Then ws.Range(ws.Cells(fd, stc), ws.Cells(ld, stc)).FillDown
    End If
    On Error GoTo 0

    With ws.Range(ws.Cells(fd, stc), ws.Cells(ld, stc))
        .HorizontalAlignment = xlCenter: .Font.Bold = True: .Font.Size = 8
    End With
End Sub

Private Sub ApplyStaffedCF(ws As Worksheet, sc As Long, fd As Long, ld As Long)
    Dim rng As Range, fc As FormatCondition, ca As String
    If ld < fd Then Exit Sub  ' No data rows
    Set rng = ws.Range(ws.Cells(fd, sc), ws.Cells(ld, sc))
    On Error Resume Next
    rng.FormatConditions.Delete
    If Err.Number <> 0 Then Err.Clear
    ca = ws.Cells(fd, sc).Address(False, False)
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & ca & "=""Yes""")
    If Err.Number = 0 And Not fc Is Nothing Then
        fc.Interior.Color = RGB(220, 252, 231): fc.Font.Color = RGB(22, 101, 52): fc.Font.Bold = True
    Else: Err.Clear
    End If
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & ca & "=""No""")
    If Err.Number = 0 And Not fc Is Nothing Then
        fc.Interior.Color = RGB(254, 232, 232): fc.Font.Color = RGB(185, 28, 28): fc.Font.Bold = True
    Else: Err.Clear
    End If
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & ca & "=""N/A""")
    If Err.Number = 0 And Not fc Is Nothing Then
        fc.Interior.Color = RGB(241, 243, 245): fc.Font.Color = RGB(100, 116, 139)
    Else: Err.Clear
    End If
    On Error GoTo 0
End Sub

'====================================================================
' HC NEED SECTIONS -- MAIN ORCHESTRATOR
'====================================================================

Private Sub AddHCNeedSections(ws As Worksheet, gec As Long, nsc As Long, _
                               sHN() As HCRequirement, hcnc As Long, _
                               sHR() As HCRequirement, hcrc As Long)
    Dim ld As Long, hsr As Long, gsc As Long, gcol As Long, nrc As Long
    Dim hcC As Collection, icc As Long
    Dim groups As Variant, gc As Long
    Dim gdc As Long, msc As Long
    Dim nER As Long, uER As Long, cER As Long
    Dim usr As Long

    ld = FindMainDataEnd(ws)
    ' Formula-safe last data row: use ListObject boundary (auto-extends with table)
    ' CRITICAL: Must NOT extend into HC tables (which start at hsr below) or formulas
    ' in Gantt columns will create circular references.
    Dim fld As Long
    If ws.ListObjects.Count > 0 Then
        Dim tbl As ListObject: Set tbl = ws.ListObjects(1)
        fld = tbl.Range.row + tbl.Range.Rows.Count - 1
    Else
        fld = ld
    End If
    hsr = ld + HC_ROW_GAP
    gsc = FindGanttStartColumn(ws)
    If gsc = 0 Then gsc = gec - 103
    gcol = FindColumnByHeader(ws, "Group")
    If gcol = 0 Then MsgBox "Group column not found.", vbExclamation: Exit Sub
    nrc = FindColumnByHeader(ws, "New/Reused")
    If nrc = 0 Then nrc = FindColumnByHeader(ws, "New/Used")

    ' Find Status column for filtering (exclude Cancelled/Completed/Non IQ)
    Dim stcol As Long
    stcol = FindColumnByHeader(ws, "Status")

    Set hcC = GetHCMilestoneCodes()
    icc = 1 + hcC.Count + 1

    groups = CollectUniqueGroups(ws, gcol, ld)
    If Not IsArray(groups) Then
        MsgBox "No groups found.", vbExclamation: Exit Sub
    End If
    gc = UBound(groups) - LBound(groups) + 1
    If gc = 0 Then MsgBox "No groups found.", vbExclamation: Exit Sub

    gdc = gsc - icc
    If gdc < 1 Then gdc = 1
    msc = gdc + 1

    ' New Systems HC Need
    nER = BuildHCNeed(ws, hsr, gsc, gec, gcol, nrc, hcC, "New", 1.5, fld, icc, groups, gc, stcol)
    RestoreHCValues ws, hsr + 2, gdc, msc, hcC, sHN, hcnc

    ' Reused Systems HC Need
    usr = nER + HC_SECTION_GAP
    uER = BuildHCNeed(ws, usr, gsc, gec, gcol, nrc, hcC, "Reused", 2#, fld, icc, groups, gc, stcol)
    RestoreHCValues ws, usr + 2, gdc, msc, hcC, sHR, hcrc

    ' Combined HC Need
    cER = BuildCombined(ws, uER + HC_SECTION_GAP, gsc, gec, gc, nER, uER, groups)

    ' New Available HC
    Dim naER As Long
    naER = BuildAvailable(ws, cER + HC_SECTION_GAP, gsc, gec, gcol, nsc, fld, gc, nrc, "New", groups, stcol)

    ' Reused Available HC
    Dim raER As Long
    raER = BuildAvailable(ws, naER + HC_SECTION_GAP, gsc, gec, gcol, nsc, fld, gc, nrc, "Reused", groups, stcol)

    ' New HC Gap
    Dim ngER As Long
    ngER = BuildGapDirect(ws, raER + HC_SECTION_GAP, gsc, gec, gc, naER, nER, "New", groups)

    ' Reused HC Gap
    Dim rgER As Long
    rgER = BuildGapDirect(ws, ngER + HC_SECTION_GAP, gsc, gec, gc, raER, uER, "Reused", groups)

    ' Total HC Gap
    Dim tgER As Long
    tgER = BuildTotalGap(ws, rgER + HC_SECTION_GAP, gsc, gec, gc, ngER, rgER, groups)
End Sub

'====================================================================
' BUILD HC NEED TABLE
' Layout: title(sr) -> header(sr+1) -> data(sr+2)
'====================================================================

Private Function BuildHCNeed(ws As Worksheet, sr As Long, gsc As Long, gec As Long, _
    gcol As Long, nrc As Long, hcC As Collection, st As String, dHC As Double, _
    lmd As Long, icc As Long, groups As Variant, gc As Long, Optional stcol As Long = 0) As Long

    Dim tr As Long, hdr As Long, dsr As Long
    Dim cc As Long, c As Variant, gdc As Long, msc As Long, qc As Long
    Dim gl As String, nrl As String, gdl As String
    Dim gi As Long, totalR As Long, gd As Long, cr As Long

    tr = sr: hdr = sr + 1: dsr = sr + 2
    gl = ColLetter(gcol)
    If nrc > 0 Then nrl = ColLetter(nrc)
    gdc = gsc - icc
    If gdc < 1 Then gdc = 1
    msc = gdc + 1
    qc = gsc - 1
    gdl = ColLetter(gdc)

    ' Build Status filter string for COUNTIFS/SUMPRODUCT (exclude Cancelled/Completed/Non IQ)
    Dim stl As String, stFilter As String
    stFilter = ""
    If stcol > 0 Then
        stl = ColLetter(stcol)
        stFilter = ",$" & stl & "$" & (m_HeaderRow + 1) & ":$" & stl & "$"
    End If

    FormatSectionTitle ws, tr, gdc, st & " Systems - HC Need", CLR_DARK

    FormatSectionHeader ws, hdr, gdc, "Group", CLR_DARK
    ws.Columns(gdc).ColumnWidth = 15
    cc = msc
    For Each c In hcC
        FormatSectionHeader ws, hdr, cc, CStr(c) & " HC", CLR_BLUE
        ws.Columns(cc).ColumnWidth = 8
        cc = cc + 1
    Next c
    FormatSectionHeader ws, hdr, qc, "Proj QTY", CLR_TEAL
    ws.Columns(qc).ColumnWidth = 10

    ' Batch copy + format date headers (replaces per-cell loop)
    BatchWriteDateHeaders ws, hdr, gsc, gec

    cr = dsr
    For gi = 1 To gc
        ws.Cells(cr, gdc).Value = groups(gi)
        ws.Cells(cr, gdc).Font.Bold = True
        cc = msc
        For Each c In hcC
            If CStr(c) = "MRCL" Then
                ws.Cells(cr, cc).Value = 0
            Else
                ws.Cells(cr, cc).Value = dHC
            End If
            ws.Cells(cr, cc).HorizontalAlignment = xlCenter
            ws.Cells(cr, cc).NumberFormat = "0.0"
            cc = cc + 1
        Next c
        If nrc > 0 Then
            Dim cfFormula As String
            cfFormula = "=COUNTIFS($" & gl & "$" & (m_HeaderRow + 1) & _
                ":$" & gl & "$" & lmd & "," & gdl & cr & _
                ",$" & nrl & "$" & (m_HeaderRow + 1) & _
                ":$" & nrl & "$" & lmd & ",""" & st & """"
            If stcol > 0 Then
                cfFormula = cfFormula & stFilter & lmd & ",""<>Cancelled""" & _
                    stFilter & lmd & ",""<>Completed""" & _
                    stFilter & lmd & ",""<>Non IQ"""
            End If
            cfFormula = cfFormula & ")"
            SafeFormulaWrite ws, cr, qc, cfFormula
        Else
            ws.Cells(cr, qc).Value = 0
        End If
        ws.Cells(cr, qc).HorizontalAlignment = xlCenter
        For gd = gsc To gec
            SafeFormulaWrite ws, cr, gd, BuildHCF(gl, gdl, cr, gd, nrc, nrl, st, hcC, msc, m_HeaderRow + 1, lmd, stcol)
        Next gd
        cr = cr + 1
    Next gi

    ' Batch format all Gantt data cells at once
    If cr > dsr Then BatchFormatHCData ws, dsr, cr - 1, gsc, gec, "0.0"

    totalR = cr
    FormatTotalRow ws, totalR, gdc, gec
    For gd = gsc To gec
        ws.Cells(totalR, gd).formula = "=SUM(" & ColLetter(gd) & dsr & ":" & ColLetter(gd) & (totalR - 1) & ")"
        ws.Cells(totalR, gd).NumberFormat = "0.0"
    Next gd

    If totalR > dsr Then ApplyHeatCS ws, gsc, gec, dsr, totalR - 1
    BuildHCNeed = totalR + 1
End Function

'====================================================================
' BUILD HC FORMULA (milestone codes inlined)
'====================================================================

Private Function BuildHCF(gl As String, gdl As String, cr As Long, gd As Long, _
    nrc As Long, nrl As String, st As String, hcC As Collection, _
    msc As Long, fd As Long, ld As Long, Optional stcol As Long = 0) As String

    Dim f As String, gcl As String
    Dim mo As Long, fp As Boolean, c As Variant, mcl As String

    ' Build status filter fragment for SUMPRODUCT (exclude Cancelled/Completed/Non IQ)
    Dim stFrag As String
    stFrag = ""
    If stcol > 0 Then
        Dim stl As String
        stl = ColLetter(stcol)
        stFrag = ",--($" & stl & "$" & fd & ":$" & stl & "$" & ld & "<>""Cancelled"")" & _
                 ",--($" & stl & "$" & fd & ":$" & stl & "$" & ld & "<>""Completed"")" & _
                 ",--($" & stl & "$" & fd & ":$" & stl & "$" & ld & "<>""Non IQ"")"
    End If

    gcl = ColLetter(gd)
    f = "": mo = 0: fp = True

    For Each c In hcC
        mcl = ColLetter(msc + mo)
        If Not fp Then f = f & "+"
        fp = False
        f = f & "SUMPRODUCT(" & _
            "--($" & gl & "$" & fd & ":$" & gl & "$" & ld & "=" & gdl & cr & ")," & _
            "--($" & gcl & "$" & fd & ":$" & gcl & "$" & ld & "=""" & CStr(c) & """)"
        If nrc > 0 Then
            f = f & ",--($" & nrl & "$" & fd & ":$" & nrl & "$" & ld & "=""" & st & """)"
        End If
        f = f & stFrag
        f = f & ")*$" & mcl & "$" & cr
        mo = mo + 1
    Next c

    BuildHCF = "=" & f
End Function

'====================================================================
' BUILD COMBINED TABLE
'====================================================================

Private Function BuildCombined(ws As Worksheet, sr As Long, gsc As Long, gec As Long, _
    gc As Long, nER As Long, uER As Long, groups As Variant) As Long

    Dim tr As Long, hdr As Long, dsr As Long, cr As Long, gd As Long
    Dim gi As Long, nds As Long, uds As Long, totalR As Long, gdc As Long

    tr = sr: hdr = sr + 1: dsr = sr + 2
    gdc = gsc - 1
    If gdc < 1 Then gdc = 1
    nds = nER - 1 - gc: uds = uER - 1 - gc

    FormatSectionTitle ws, tr, gdc, "Combined - HC Need", CLR_DARK

    FormatSectionHeader ws, hdr, gdc, "Group", CLR_DARK
    ws.Columns(gdc).ColumnWidth = 15
    BatchWriteDateHeaders ws, hdr, gsc, gec

    cr = dsr
    For gi = 1 To gc
        ws.Cells(cr, gdc).Value = groups(gi)
        ws.Cells(cr, gdc).Font.Bold = True
        For gd = gsc To gec
            ws.Cells(cr, gd).formula = "=" & ColLetter(gd) & (nds + gi - 1) & _
                "+" & ColLetter(gd) & (uds + gi - 1)
        Next gd
        cr = cr + 1
    Next gi

    ' Batch format all Gantt data cells
    If cr > dsr Then BatchFormatHCData ws, dsr, cr - 1, gsc, gec, "0.0"

    totalR = cr
    FormatTotalRow ws, totalR, gdc, gec
    For gd = gsc To gec
        ws.Cells(totalR, gd).formula = "=SUM(" & ColLetter(gd) & dsr & ":" & ColLetter(gd) & (totalR - 1) & ")"
        ws.Cells(totalR, gd).NumberFormat = "0.0"
    Next gd

    If totalR > dsr Then ApplyHeatCS ws, gsc, gec, dsr, totalR - 1
    BuildCombined = totalR + 1
End Function

'====================================================================
' BUILD AVAILABLE HC TABLE
'====================================================================

Private Function BuildAvailable(ws As Worksheet, sr As Long, gsc As Long, gec As Long, _
    gcol As Long, nsc As Long, lmd As Long, gc As Long, _
    nrc As Long, systemType As String, groups As Variant, Optional stcol As Long = 0) As Long

    Dim tr As Long, hdr As Long, dsr As Long, cr As Long, gd As Long
    Dim gl As String, wr As String, f As String
    Dim i As Long, nc As Long, sc As Long, ec As Long, fd As Long
    Dim totalR As Long, gdc As Long
    Dim nrl As String, gdl As String, gi As Long

    tr = sr: hdr = sr + 1: dsr = sr + 2
    fd = m_HeaderRow + 1
    gl = ColLetter(gcol)
    gdc = gsc - 1
    If gdc < 1 Then gdc = 1
    gdl = ColLetter(gdc)
    If nrc > 0 Then nrl = ColLetter(nrc)

    ' Build status filter fragment for SUMPRODUCT (exclude Cancelled/Completed/Non IQ)
    Dim stFrag As String
    stFrag = ""
    If stcol > 0 Then
        Dim stl As String
        stl = ColLetter(stcol)
        stFrag = ",--($" & stl & "$" & fd & ":$" & stl & "$" & lmd & "<>""Cancelled"")" & _
                 ",--($" & stl & "$" & fd & ":$" & stl & "$" & lmd & "<>""Completed"")" & _
                 ",--($" & stl & "$" & fd & ":$" & stl & "$" & lmd & "<>""Non IQ"")"
    End If

    FormatSectionTitle ws, tr, gdc, systemType & " Available HC (from NIF)", CLR_TEAL

    FormatSectionHeader ws, hdr, gdc, "Group", CLR_DARK
    BatchWriteDateHeaders ws, hdr, gsc, gec

    cr = dsr
    For gi = 1 To gc
        ws.Cells(cr, gdc).Value = groups(gi)
        ws.Cells(cr, gdc).Font.Bold = True
        For gd = gsc To gec
            wr = ColLetter(gd) & "$" & m_HeaderRow
            f = ""
            For i = 1 To NIF_EMPLOYEE_COUNT
                nc = nsc + (i - 1) * 3: sc = nc + 1: ec = nc + 2
                If f <> "" Then f = f & "+"
                f = f & "SUMPRODUCT(" & _
                    "--($" & gl & "$" & fd & ":$" & gl & "$" & lmd & "=" & gdl & cr & ")," & _
                    "--($" & ColLetter(nc) & "$" & fd & ":$" & ColLetter(nc) & "$" & lmd & "<>"""")," & _
                    "--($" & ColLetter(sc) & "$" & fd & ":$" & ColLetter(sc) & "$" & lmd & "<=" & wr & "+6)," & _
                    "--($" & ColLetter(ec) & "$" & fd & ":$" & ColLetter(ec) & "$" & lmd & ">=" & wr & ")"
                If nrc > 0 Then
                    f = f & ",--($" & nrl & "$" & fd & ":$" & nrl & "$" & lmd & "=""" & systemType & """)"
                End If
                f = f & stFrag & ")"
            Next i
            SafeFormulaWrite ws, cr, gd, "=" & f
        Next gd
        cr = cr + 1
    Next gi

    ' Batch format all Gantt data cells
    If cr > dsr Then BatchFormatHCData ws, dsr, cr - 1, gsc, gec, "0"

    totalR = cr
    FormatTotalRow ws, totalR, gdc, gec
    For gd = gsc To gec
        ws.Cells(totalR, gd).formula = "=SUM(" & ColLetter(gd) & dsr & ":" & ColLetter(gd) & (totalR - 1) & ")"
        ws.Cells(totalR, gd).NumberFormat = "0"
    Next gd

    If totalR > dsr Then ApplyHeatCS ws, gsc, gec, dsr, totalR - 1
    BuildAvailable = totalR + 1
End Function

'====================================================================
' BUILD GAP DIRECT (Available - Need for a specific system type)
'====================================================================

Private Function BuildGapDirect(ws As Worksheet, sr As Long, gsc As Long, gec As Long, _
    gc As Long, availER As Long, needER As Long, _
    systemType As String, groups As Variant) As Long

    Dim tr As Long, hdr As Long, dsr As Long, cr As Long, gd As Long
    Dim gi As Long, ads As Long, nds As Long
    Dim totalR As Long, gdc As Long
    Dim colL As String

    tr = sr: hdr = sr + 1: dsr = sr + 2
    gdc = gsc - 1
    If gdc < 1 Then gdc = 1
    ads = availER - 1 - gc
    nds = needER - 1 - gc

    FormatSectionTitle ws, tr, gdc, systemType & " HC Gap (Available - Need)", CLR_RED_TITLE

    FormatSectionHeader ws, hdr, gdc, "Group", CLR_DARK
    BatchWriteDateHeaders ws, hdr, gsc, gec

    cr = dsr
    For gi = 1 To gc
        ws.Cells(cr, gdc).Value = groups(gi)
        ws.Cells(cr, gdc).Font.Bold = True
        For gd = gsc To gec
            colL = ColLetter(gd)
            ws.Cells(cr, gd).formula = "=" & colL & (ads + gi - 1) & _
                "-" & colL & (nds + gi - 1)
        Next gd
        cr = cr + 1
    Next gi

    ' Batch format all Gantt data cells
    If cr > dsr Then BatchFormatHCData ws, dsr, cr - 1, gsc, gec, "0.0"

    totalR = cr
    FormatTotalRow ws, totalR, gdc, gec
    For gd = gsc To gec
        ws.Cells(totalR, gd).formula = "=SUM(" & ColLetter(gd) & dsr & ":" & ColLetter(gd) & (totalR - 1) & ")"
        ws.Cells(totalR, gd).NumberFormat = "0.0"
    Next gd

    If totalR > dsr Then ApplyGapCS ws, gsc, gec, dsr, totalR - 1
    BuildGapDirect = totalR + 1
End Function

'====================================================================
' BUILD TOTAL GAP (New Gap + Reused Gap)
'====================================================================

Private Function BuildTotalGap(ws As Worksheet, sr As Long, gsc As Long, gec As Long, _
    gc As Long, newGapER As Long, reuGapER As Long, groups As Variant) As Long

    Dim tr As Long, hdr As Long, dsr As Long, cr As Long, gd As Long
    Dim gi As Long, ngds As Long, rgds As Long
    Dim totalR As Long, gdc As Long
    Dim colL As String

    tr = sr: hdr = sr + 1: dsr = sr + 2
    gdc = gsc - 1
    If gdc < 1 Then gdc = 1
    ngds = newGapER - 1 - gc
    rgds = reuGapER - 1 - gc

    FormatSectionTitle ws, tr, gdc, "Total HC Gap", CLR_RED_TITLE

    FormatSectionHeader ws, hdr, gdc, "Group", CLR_DARK
    BatchWriteDateHeaders ws, hdr, gsc, gec

    cr = dsr
    For gi = 1 To gc
        ws.Cells(cr, gdc).Value = groups(gi)
        ws.Cells(cr, gdc).Font.Bold = True
        For gd = gsc To gec
            colL = ColLetter(gd)
            ws.Cells(cr, gd).formula = "=" & colL & (ngds + gi - 1) & _
                "+" & colL & (rgds + gi - 1)
        Next gd
        cr = cr + 1
    Next gi

    ' Batch format all Gantt data cells
    If cr > dsr Then BatchFormatHCData ws, dsr, cr - 1, gsc, gec, "0.0"

    totalR = cr
    FormatTotalRow ws, totalR, gdc, gec
    For gd = gsc To gec
        ws.Cells(totalR, gd).formula = "=SUM(" & ColLetter(gd) & dsr & ":" & ColLetter(gd) & (totalR - 1) & ")"
        ws.Cells(totalR, gd).NumberFormat = "0.0"
    Next gd

    If totalR > dsr Then ApplyGapCS ws, gsc, gec, dsr, totalR - 1
    BuildTotalGap = totalR + 1
End Function

'====================================================================
' COLOR SCALES
'====================================================================

' Heat scale: white(0) -> yellow -> orange -> red (for Need/Available)
Private Sub ApplyHeatCS(ws As Worksheet, sc As Long, ec As Long, sr As Long, er As Long)
    Dim rng As Range, cs As ColorScale
    If er < sr Or ec < sc Then Exit Sub
    Set rng = ws.Range(ws.Cells(sr, sc), ws.Cells(er, ec))
    On Error Resume Next
    rng.FormatConditions.Delete
    If Err.Number <> 0 Then Err.Clear
    Set cs = rng.FormatConditions.AddColorScale(ColorScaleType:=3)
    If Err.Number <> 0 Then
        DebugLog "ApplyHeatCS failed: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    cs.ColorScaleCriteria(1).Type = xlConditionValueNumber
    cs.ColorScaleCriteria(1).Value = 0
    cs.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 255, 255)   ' white
    cs.ColorScaleCriteria(2).Type = xlConditionValuePercentile
    cs.ColorScaleCriteria(2).Value = 50
    cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)   ' yellow
    cs.ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    cs.ColorScaleCriteria(3).FormatColor.Color = RGB(230, 100, 80)    ' orange-red
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

' Gap scale: red(negative) -> white(0) -> green(positive)
Private Sub ApplyGapCS(ws As Worksheet, sc As Long, ec As Long, sr As Long, er As Long)
    Dim rng As Range, cs As ColorScale
    If er < sr Or ec < sc Then Exit Sub
    Set rng = ws.Range(ws.Cells(sr, sc), ws.Cells(er, ec))
    On Error Resume Next
    rng.FormatConditions.Delete
    If Err.Number <> 0 Then Err.Clear
    Set cs = rng.FormatConditions.AddColorScale(ColorScaleType:=3)
    If Err.Number <> 0 Then
        DebugLog "ApplyGapCS failed: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    cs.ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    cs.ColorScaleCriteria(1).FormatColor.Color = RGB(230, 100, 80)    ' red
    cs.ColorScaleCriteria(2).Type = xlConditionValueNumber
    cs.ColorScaleCriteria(2).Value = 0
    cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 255)   ' white
    cs.ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    cs.ColorScaleCriteria(3).FormatColor.Color = RGB(99, 190, 123)    ' green
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub
