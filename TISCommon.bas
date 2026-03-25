Attribute VB_Name = "TISCommon"
'====================================================================
' TIS Common Utilities - Rev14
'
' Shared helper functions used by all TIS Tracker modules.
' Extracted to eliminate duplication and ensure consistency.
'
' Rev10 changes from Rev9:
'   - Added AMAT Brand theme color constants
'   - Added DebugLog conditional compilation wrapper
'   - Added FindWorkingSheetTable utility
'   - Added FormatCardStyle utility
'   - Added GetMilestoneStartHeaders for multi-date fallback
'   - DEBUG_MODE constant controls Debug.Print output
'
' Utilities:
'   ColLetter           - Column number to letter(s)
'   SheetExists         - Safe sheet existence check
'   FileExists          - Safe file existence check
'   FindWorkingSheet    - Locate latest "Working Sheet*"
'   FindWorkingSheetTable - Locate ListObject on Working Sheet
'   FindHeaderRow       - Locate header row in a worksheet
'   FindHeaderCol       - Find header column in specific row
'   BuildProjectKey     - Create composite key for matching
'   SanitizeForNamedRange - Sanitize string for Excel named ranges
'   FormatCardStyle     - Apply card-style formatting to range
'
' Sorting:
'   ShellSortVariantArray - Sort variant array ascending
'   ShellSortDescending   - Sort Long array descending
'   IsDateHeader          - Check if header represents a date field
'   GetSortPriorityColumns - Read sort priorities from Definitions col C
'   SortWithHelperColumns  - Sort worksheet using helper columns
'   GetMilestoneStartHeaders - Get milestone start column headers
'
' String:
'   GetLastWord       - Extract last word from string
'   GetShortAbbrev    - Abbreviation normalization
'
' State:
'   SaveAppState / RestoreAppState / SetPerformanceMode
'   GetDefinitionsFilterDescription
'   DebugLog          - Conditional debug output
'====================================================================

Option Explicit

' Conditional compilation: set to True during development for debug output
#Const DEBUG_MODE = False

' Shared constants
Public Const TIS_DATA_START_ROW As Long = 15
Public Const TIS_VERSION As String = "Rev14"
Public Const TIS_SHEET_WORKING As String = "Working Sheet"
Public Const TIS_SHEET_TIS As String = "TIS"
Public Const TIS_SHEET_DEFINITIONS As String = "Definitions"
Public Const TIS_SHEET_CEIDS As String = "CEIDs"
Public Const TIS_SHEET_MILESTONES As String = "Milestones"
Public Const TIS_SHEET_REMOVED As String = "Removed Systems"
Public Const TIS_SHEET_DASHBOARD As String = "Dashboard"

'====================================================================
' REV14 COLUMN HEADERS -- Committed milestone date columns (user-owned)
' These are the exact header strings written to the Working Sheet.
' No "Our" prefix -- these ARE the dates. TIS dates are the reference feed.
'====================================================================
Public Const TIS_COL_OUR_SET    As String = "Set"
Public Const TIS_COL_OUR_SL1    As String = "SL1"
Public Const TIS_COL_OUR_SL2    As String = "SL2"
Public Const TIS_COL_OUR_SQ     As String = "SQ"
Public Const TIS_COL_OUR_CONVS  As String = "Conv.S"
Public Const TIS_COL_OUR_CONVF  As String = "Conv.F"
Public Const TIS_COL_OUR_MRCLS  As String = "MRCL.S"
Public Const TIS_COL_OUR_MRCLF  As String = "MRCL.F"

' Operational columns
Public Const TIS_COL_STATUS     As String = "Status"
Public Const TIS_COL_LOCK       As String = "Lock?"
Public Const TIS_COL_HEALTH     As String = "Health"
Public Const TIS_COL_WHATIF     As String = "WhatIf"          ' User enters hypothetical start date

' TIS source column names (exact headers in TIS.xlsx)
Public Const TIS_SRC_SDD        As String = "SDD"
Public Const TIS_SRC_SET        As String = "Set Start"
Public Const TIS_SRC_SL1        As String = "SL1 Signoff Finish"
Public Const TIS_SRC_SL2        As String = "SL2 Signoff Finish"
Public Const TIS_SRC_SQ         As String = "Supplier Qual Finish"
Public Const TIS_SRC_CONVS      As String = "Convert Start"
Public Const TIS_SRC_CONVF      As String = "Convert Finish"
Public Const TIS_SRC_MRCLS      As String = "MRCL Start"
Public Const TIS_SRC_MRCLF      As String = "MRCL Finish"

' Fill/border colors for change tracking
Public Const CLR_CHANGE_FILL    As Long = 42495   ' RGB(255, 165, 0) Orange -- TIS field changed
Public Const CLR_NEW_DATE_BORDER As Long = 16711680 ' RGB(0, 0, 255) Blue -- auto-populated Our Date

' Health / Status conditional formatting colors
Public Const STATUS_ONTRACK_FG  As Long = 1409045   ' RGB(21, 128, 61)
Public Const STATUS_ONTRACK_BG  As Long = 15204060  ' RGB(220, 252, 231)
Public Const STATUS_ATRISK_FG   As Long = 520097    ' RGB(161, 98, 7)
Public Const STATUS_ATRISK_BG   As Long = 13107198  ' RGB(254, 243, 199)
Public Const STATUS_BEHIND_FG   As Long = 1842617   ' RGB(185, 28, 28)
Public Const STATUS_BEHIND_BG   As Long = 14869246  ' RGB(254, 226, 226)

' Slate scale -- neutral grays for text, borders, backgrounds
Public Const SLATE_50  As Long = 16579320   ' RGB(248, 250, 252) Near-white
Public Const SLATE_100 As Long = 16381425   ' RGB(241, 245, 249) Frost
Public Const SLATE_200 As Long = 15787746   ' RGB(226, 232, 240) Light border
Public Const SLATE_300 As Long = 14800331   ' RGB(203, 213, 225) Border/divider
Public Const SLATE_500 As Long = 9139300    ' RGB(100, 116, 139) Muted text
Public Const SLATE_700 As Long = 5587251    ' RGB(51, 65, 85)    Secondary text
Public Const SLATE_900 As Long = 3875102    ' RGB(30, 41, 59)    Primary text

' Zone header colors (background/foreground pairs for Working Sheet header zones)
Public Const ZONE_IDENTITY_BG   As Long = 7029760   ' RGB(0, 56, 107)    Navy -- identity columns
Public Const ZONE_IDENTITY_FG   As Long = 16777215  ' RGB(255, 255, 255) White text
Public Const ZONE_OUR_BG        As Long = 2241804    ' RGB(12, 51, 34)    Deep green -- Our Dates zone
Public Const ZONE_OUR_FG        As Long = 12706467   ' RGB(163, 228, 193) Light green text
Public Const ZONE_TIS_BG        As Long = 4992527    ' RGB(15, 46, 76)    Deep blue -- TIS Dates zone
Public Const ZONE_TIS_FG        As Long = 16045237   ' RGB(181, 212, 244) Light blue text
Public Const ZONE_USER_BG       As Long = 5930035    ' RGB(51, 122, 90)   Deep green -- user fields
Public Const ZONE_USER_FG       As Long = 16777215  ' RGB(255, 255, 255) White text
Public Const ZONE_CALC_BG       As Long = 4176516    ' RGB(132, 172, 63)  Deep amber -- calculated
Public Const ZONE_CALC_FG       As Long = 16777215  ' RGB(255, 255, 255) White text

'====================================================================
' APPLIED MATERIALS BRAND COLOR PALETTE
' Applied Materials Brand Theme - AMAT Silver Lake Blue (#569CBE) primary
' Deep navy backgrounds, silver lake blue accents, emerald/teal/coral semantics.
' Values are VBA Long format: R + G*256 + B*65536
'====================================================================

' Backgrounds (neutral tones)
Public Const THEME_BG As Long = 3349260                  ' RGB(12, 27, 51)   Deep Navy (dark bg)
Public Const THEME_SURFACE As Long = 6043158              ' RGB(22, 54, 92)   Steel Blue (card/panel bg)
Public Const THEME_WHITE As Long = 16777215              ' RGB(255, 255, 255) contrast text on dark

' Accent colors (brand secondary)
Public Const THEME_ACCENT As Long = 12491862             ' RGB(86, 156, 190) AMAT Silver Lake Blue
Public Const THEME_ACCENT2 As Long = 12692480            ' RGB(0, 172, 193)  Teal

' Semantic colors (brand secondary)
Public Const THEME_SUCCESS As Long = 6076462              ' RGB(46, 184, 92)  Emerald
Public Const THEME_WARNING As Long = 5093375             ' RGB(255, 183, 77) Amber
Public Const THEME_DANGER As Long = 5264367                ' RGB(239, 83, 80)  Coral Red

' Borders and dividers
Public Const THEME_BORDER As Long = 14800331             ' RGB(203, 213, 225) Slate Border

' Text colors
Public Const THEME_TEXT As Long = 16777215               ' RGB(255, 255, 255) White (on dark bg)
Public Const THEME_TEXT_SEC As Long = 9139300             ' RGB(100, 116, 139) Slate Gray

' Typography
Public Const THEME_FONT As String = "Segoe UI"

' Application state holder
Public Type AppState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    Calculation As XlCalculation
End Type

'====================================================================
' DEBUG LOG - conditional debug output
' Only outputs when DEBUG_MODE is True (set at top of module)
'====================================================================

Public Sub DebugLog(msg As String)
    #If DEBUG_MODE Then
        Debug.Print Format(Now, "hh:nn:ss") & " | " & msg
    #End If
End Sub

'====================================================================
' COLUMN LETTER
'====================================================================

Public Function ColLetter(colNum As Long) As String
    Dim n As Long
    Dim result As String
    n = colNum: result = ""
    Do While n > 0
        result = Chr(((n - 1) Mod 26) + 65) & result
        n = (n - 1) \ 26
    Loop
    ColLetter = result
End Function

'====================================================================
' SHEET EXISTS (For Each - safe with chart sheets)
'====================================================================

Public Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    SheetExists = False
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

'====================================================================
' FIND WORKING SHEET - returns latest "Working Sheet*" by index
'====================================================================

Public Function FindWorkingSheet() As Worksheet
    Dim ws As Worksheet
    Dim bestSheet As Worksheet
    Dim bestIdx As Long

    Set bestSheet = Nothing
    bestIdx = 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like TIS_SHEET_WORKING & "*" Then
            If ws.Index > bestIdx Then
                bestIdx = ws.Index
                Set bestSheet = ws
            End If
        End If
    Next ws

    Set FindWorkingSheet = bestSheet
End Function

'====================================================================
' FIND WORKING SHEET TABLE - returns ListObject on latest Working Sheet
'====================================================================

Public Function FindWorkingSheetTable() As ListObject
    Dim ws As Worksheet
    Set ws = FindWorkingSheet()
    If ws Is Nothing Then
        Set FindWorkingSheetTable = Nothing
        Exit Function
    End If
    If ws.ListObjects.Count > 0 Then
        Set FindWorkingSheetTable = ws.ListObjects(1)
    Else
        Set FindWorkingSheetTable = Nothing
    End If
End Function

'====================================================================
' SHELL SORT VARIANT ARRAY (ascending)
'====================================================================

Public Sub ShellSortVariantArray(arr As Variant)
    Dim gap As Long, i As Long, j As Long, n As Long
    Dim temp As Variant
    n = UBound(arr) - LBound(arr) + 1
    gap = n \ 2
    Do While gap > 0
        For i = LBound(arr) + gap To UBound(arr)
            temp = arr(i): j = i
            Do While j >= LBound(arr) + gap
                If arr(j - gap) > temp Then
                    arr(j) = arr(j - gap): j = j - gap
                Else: Exit Do
                End If
            Loop
            arr(j) = temp
        Next i
        gap = gap \ 2
    Loop
End Sub

'====================================================================
' GET LAST WORD from a string (e.g., "Set - SL1" -> "SL1")
'====================================================================

Public Function GetLastWord(s As String) As String
    Dim parts As Variant
    Dim cleaned As String
    cleaned = Trim(s)
    If cleaned = "" Then GetLastWord = "": Exit Function
    parts = Split(cleaned, " ")
    GetLastWord = parts(UBound(parts))
End Function

'====================================================================
' GET SHORT ABBREVIATION (normalize display names)
'====================================================================

Public Function GetShortAbbrev(abbrev As String) As String
    Dim ua As String
    ua = UCase(Trim(abbrev))
    Select Case ua
        Case "CONVERSION": GetShortAbbrev = "CV"
        Case "PREFAC", "PRE-FAC", "PREFAB": GetShortAbbrev = "PF"
        Case "DECON", "DECONTAMINATION": GetShortAbbrev = "DC"
        Case "DEMO", "DEMOLITION": GetShortAbbrev = "DM"
        Case Else: GetShortAbbrev = ua
    End Select
End Function

'====================================================================
' FIND HEADER ROW in a worksheet (searches rows 1-20 for known headers)
'====================================================================

Public Function FindHeaderRow(ws As Worksheet) As Long
    Dim r As Long, c As Long, cellVal As String
    FindHeaderRow = 0
    For r = 1 To 20
        For c = 1 To 50
            cellVal = LCase(Trim(Replace(Replace(CStr(ws.Cells(r, c).Value), vbLf, ""), vbCr, "")))
            If cellVal = "ceid" Or cellVal = "entity code" Or cellVal = "site" Then
                FindHeaderRow = r: Exit Function
            End If
        Next c
    Next r
End Function

'====================================================================
' BUILD PROJECT KEY (Site|Entity Code|Event Type)
'====================================================================

Public Function BuildProjectKey(ws As Worksheet, row As Long, _
                                 siteCol As Long, entityCodeCol As Long, eventTypeCol As Long) As String
    Dim s As String, ec As String, et As String
    s = "": ec = "": et = ""
    If siteCol > 0 Then s = LCase(Trim(CStr(ws.Cells(row, siteCol).Value)))
    If entityCodeCol > 0 Then ec = LCase(Trim(CStr(ws.Cells(row, entityCodeCol).Value)))
    If eventTypeCol > 0 Then et = LCase(Trim(CStr(ws.Cells(row, eventTypeCol).Value)))
    BuildProjectKey = s & "|" & ec & "|" & et
End Function

'====================================================================
' APPLICATION STATE MANAGEMENT
'====================================================================

Public Function SaveAppState() As AppState
    Dim st As AppState
    st.ScreenUpdating = Application.ScreenUpdating
    st.EnableEvents = Application.EnableEvents
    st.Calculation = Application.Calculation
    SaveAppState = st
End Function

Public Sub RestoreAppState(st As AppState)
    Application.Calculation = st.Calculation
    Application.EnableEvents = st.EnableEvents
    Application.ScreenUpdating = st.ScreenUpdating
End Sub

Public Sub SetPerformanceMode()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

'====================================================================
' FIND HEADER COLUMN in a specific row
'====================================================================

Public Function FindHeaderCol(ws As Worksheet, headerRow As Long, _
                               searchText As String, maxCol As Long) As Long
    Dim c As Long, cellVal As String
    Dim searchLower As String
    searchLower = LCase(Trim(Replace(Replace(searchText, vbLf, ""), vbCr, "")))
    FindHeaderCol = 0
    For c = 1 To maxCol
        cellVal = LCase(Trim(Replace(Replace(CStr(ws.Cells(headerRow, c).Value), vbLf, ""), vbCr, "")))
        If cellVal = searchLower Then
            FindHeaderCol = c: Exit Function
        End If
    Next c
End Function

'====================================================================
' GET DEFINITIONS FILTER DESCRIPTION
'====================================================================

Public Function GetDefinitionsFilterDescription() As String
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim headerName As String, filterValue As String
    Dim desc As String

    If Not SheetExists(ThisWorkbook, TIS_SHEET_DEFINITIONS) Then
        GetDefinitionsFilterDescription = "No Definitions sheet found."
        Exit Function
    End If

    Set ws = ThisWorkbook.Sheets(TIS_SHEET_DEFINITIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    desc = ""

    For i = 2 To lastRow
        headerName = Trim(CStr(ws.Cells(i, 1).Value))
        filterValue = Trim(CStr(ws.Cells(i, 2).Value))
        If headerName <> "" And filterValue <> "" Then
            If desc <> "" Then desc = desc & vbCrLf
            desc = desc & "  " & Chr(149) & " " & headerName & ": " & Chr(34) & filterValue & Chr(34)
        End If
    Next i

    If desc = "" Then
        GetDefinitionsFilterDescription = "No filters applied - all projects included."
    Else
        GetDefinitionsFilterDescription = "Active filters:" & vbCrLf & desc
    End If
End Function

'====================================================================
' IS DATE HEADER - checks if a column header represents a date field
'====================================================================

Public Function IsDateHeader(headerName As String) As Boolean
    Dim lowerName As String
    lowerName = LCase(headerName)
    IsDateHeader = (InStr(lowerName, "start") > 0) Or (InStr(lowerName, "finish") > 0) Or _
                   (InStr(lowerName, "date") > 0) Or (InStr(lowerName, "end") > 0)
End Function

'====================================================================
' FILE EXISTS - safe file existence check using Dir()
'====================================================================

Public Function FileExists(filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
End Function

'====================================================================
' SHELL SORT DESCENDING (Long array, 1-based, n elements)
'====================================================================

Public Sub ShellSortDescending(arr() As Long, n As Long)
    Dim gap As Long, i As Long, j As Long
    Dim temp As Long

    gap = n \ 2
    Do While gap > 0
        For i = gap + 1 To n
            temp = arr(i)
            j = i
            Do While j > gap
                If arr(j - gap) < temp Then
                    arr(j) = arr(j - gap)
                    j = j - gap
                Else
                    Exit Do
                End If
            Loop
            arr(j) = temp
        Next i
        gap = gap \ 2
    Loop
End Sub

'====================================================================
' GET SORT PRIORITY COLUMNS from Definitions sheet
'====================================================================

Public Function GetSortPriorityColumns() As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim headerName As String
    Dim sortValue As Variant
    Dim priority As Long

    Set dict = CreateObject("Scripting.Dictionary")
    If Not SheetExists(ThisWorkbook, TIS_SHEET_DEFINITIONS) Then Set GetSortPriorityColumns = dict: Exit Function
    Set ws = ThisWorkbook.Sheets(TIS_SHEET_DEFINITIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    For i = 2 To lastRow
        headerName = Trim(CStr(ws.Cells(i, 1).Value))
        sortValue = ws.Cells(i, 3).Value
        If headerName <> "" And IsNumeric(sortValue) Then
            priority = CLng(sortValue)
            If priority > 0 And IsDateHeader(headerName) Then
                If Not dict.exists(priority) Then
                    dict(priority) = headerName
                Else
                    dict(priority) = dict(priority) & "|" & headerName
                End If
            End If
        End If
    Next i

    Set GetSortPriorityColumns = dict
End Function

'====================================================================
' SORT WITH HELPER COLUMNS
'====================================================================

Public Sub SortWithHelperColumns(ws As Worksheet, sortDict As Object, _
        headerMap As Object, headerRow As Long, lastDataRow As Long)
    Dim sortKeys As Variant
    Dim sortKey As Variant
    Dim i As Long, j As Long, r As Long
    Dim helperColStart As Long, helperColCount As Long, currentHelperCol As Long
    Dim sortHeaderNames() As String
    Dim sortColIndex As Long
    Dim dataRowsCount As Long
    Dim helperArr() As Variant
    Dim srcArr As Variant
    Dim srcArrays() As Variant
    Dim validCols As Long
    Dim dateVal As Variant, minVal As Date
    Dim colIdx As Long
    Dim sortRange As Range
    Dim fullLastCol As Long

    If sortDict.Count = 0 Then Exit Sub
    If lastDataRow <= headerRow Then Exit Sub

    dataRowsCount = lastDataRow - headerRow

    sortKeys = sortDict.keys
    ShellSortVariantArray sortKeys

    fullLastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    helperColStart = fullLastCol + 1
    helperColCount = 0

    With ws.Sort
        .SortFields.Clear

        For i = LBound(sortKeys) To UBound(sortKeys)
            sortKey = sortKeys(i)
            If Not sortDict.exists(sortKey) Then GoTo NextSortKey

            sortHeaderNames = Split(sortDict(sortKey), "|")
            If UBound(sortHeaderNames) < 0 Then GoTo NextSortKey

            currentHelperCol = helperColStart + helperColCount
            helperColCount = helperColCount + 1
            ws.Cells(headerRow, currentHelperCol).Value = "Helper_Sort_" & CStr(sortKey)

            If dataRowsCount > 0 Then
                ReDim helperArr(1 To dataRowsCount, 1 To 1)

                If UBound(sortHeaderNames) = 0 Then
                    sortColIndex = 0
                    If headerMap.exists(LCase(sortHeaderNames(0))) Then
                        sortColIndex = headerMap(LCase(sortHeaderNames(0)))
                    End If
                    If sortColIndex > 0 Then
                        srcArr = ws.Range(ws.Cells(headerRow + 1, sortColIndex), _
                                          ws.Cells(lastDataRow, sortColIndex)).Value
                        For r = 1 To dataRowsCount
                            helperArr(r, 1) = srcArr(r, 1)
                        Next r
                    End If
                Else
                    validCols = 0
                    ReDim srcArrays(0 To UBound(sortHeaderNames))
                    For j = 0 To UBound(sortHeaderNames)
                        colIdx = 0
                        If headerMap.exists(LCase(sortHeaderNames(j))) Then
                            colIdx = headerMap(LCase(sortHeaderNames(j)))
                        End If
                        If colIdx > 0 Then
                            srcArrays(j) = ws.Range(ws.Cells(headerRow + 1, colIdx), _
                                                     ws.Cells(lastDataRow, colIdx)).Value
                            validCols = validCols + 1
                        Else
                            srcArrays(j) = Empty
                        End If
                    Next j

                    If validCols > 0 Then
                        For r = 1 To dataRowsCount
                            minVal = DateSerial(2099, 12, 31)
                            For j = 0 To UBound(sortHeaderNames)
                                If Not IsEmpty(srcArrays(j)) Then
                                    dateVal = srcArrays(j)(r, 1)
                                    If IsDate(dateVal) Then
                                        If CDate(dateVal) < minVal Then minVal = CDate(dateVal)
                                    End If
                                End If
                            Next j
                            If minVal < DateSerial(2099, 12, 31) Then
                                helperArr(r, 1) = minVal
                            End If
                        Next r
                    End If
                End If

                ws.Range(ws.Cells(headerRow + 1, currentHelperCol), _
                         ws.Cells(lastDataRow, currentHelperCol)).Value = helperArr
            End If

            .SortFields.Add key:=ws.Cells(headerRow + 1, currentHelperCol), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
NextSortKey:
        Next i

        If .SortFields.Count > 0 Then
            Set sortRange = ws.Range(ws.Cells(headerRow, 1), _
                                     ws.Cells(lastDataRow, helperColStart + helperColCount - 1))
            .SetRange sortRange
            .header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End If
    End With

    If helperColCount > 0 Then
        ws.Range(ws.Columns(helperColStart), _
                 ws.Columns(helperColStart + helperColCount - 1)).Delete
    End If
End Sub

'====================================================================
' SANITIZE STRING FOR NAMED RANGE
'====================================================================

Public Function SanitizeForNamedRange(rawName As String) As String
    Dim i As Long, ch As String, result As String
    result = ""
    For i = 1 To Len(rawName)
        ch = Mid(rawName, i, 1)
        If ch Like "[A-Za-z0-9_.]" Then
            result = result & ch
        Else
            result = result & "_"
        End If
    Next i
    If Len(result) > 0 Then
        If result Like "[0-9]*" Then result = "_" & result
    End If
    SanitizeForNamedRange = result
End Function

'====================================================================
' APPLY DARK BACKGROUND - sets dark bg + hides gridlines
'====================================================================

Public Sub ApplyDarkBackground(ws As Worksheet)
    ws.Cells.Interior.Color = THEME_BG
    ws.Cells.Font.Name = THEME_FONT
    ws.Cells.Font.Color = THEME_TEXT
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub

'====================================================================
' FORMAT CARD STYLE - applies card-like formatting to a range
' Dark theme: surface bg with thick accent left border.
'====================================================================

Public Sub FormatCardStyle(rng As Range, bgColor As Long, Optional accentColor As Long = 0)
    With rng
        .Interior.Color = bgColor
        .Font.Color = RGB(100, 116, 139)   ' Slate label text
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = RGB(226, 232, 240)
        .Borders(xlEdgeRight).Weight = xlHairline
        .Borders(xlEdgeTop).Color = RGB(226, 232, 240)
        .Borders(xlEdgeTop).Weight = xlHairline
        .Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
        .Borders(xlEdgeBottom).Weight = xlHairline
        If accentColor <> 0 Then
            .Borders(xlEdgeLeft).Color = accentColor
            .Borders(xlEdgeLeft).Weight = xlThick
        Else
            .Borders(xlEdgeLeft).Color = THEME_ACCENT
            .Borders(xlEdgeLeft).Weight = xlThick
        End If
    End With
End Sub

'====================================================================
' GET MILESTONE START HEADERS
' Returns a Collection of all milestone "Start" header names from
' the Definitions sheet (column A) that have sort priority > 0
' in column C and contain "start" in their name.
' Used for multi-date fallback in dashboard formulas where Set Start
' may be empty (e.g., Reused systems without Set Start dates).
'====================================================================

Public Function GetMilestoneStartHeaders() As Collection
    Dim headers As New Collection
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim headerName As String
    Dim sortValue As Variant

    If Not SheetExists(ThisWorkbook, TIS_SHEET_DEFINITIONS) Then
        Set GetMilestoneStartHeaders = headers
        Exit Function
    End If

    Set ws = ThisWorkbook.Sheets(TIS_SHEET_DEFINITIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    For i = 2 To lastRow
        headerName = Trim(CStr(ws.Cells(i, 1).Value))
        sortValue = ws.Cells(i, 3).Value
        If headerName <> "" Then
            ' Include headers that are date-type "Start" columns with sort priority
            If InStr(1, LCase(headerName), "start", vbTextCompare) > 0 Then
                If IsNumeric(sortValue) Then
                    If CLng(sortValue) > 0 Then
                        headers.Add headerName
                    End If
                End If
            End If
        End If
    Next i

    ' Also check for SDD as a fallback date
    For i = 2 To lastRow
        headerName = Trim(CStr(ws.Cells(i, 1).Value))
        If LCase(headerName) = "sdd" Then
            Dim isDup As Boolean
            Dim item As Variant
            isDup = False
            For Each item In headers
                If LCase(CStr(item)) = "sdd" Then isDup = True: Exit For
            Next item
            If Not isDup Then headers.Add headerName
            Exit For
        End If
    Next i

    Set GetMilestoneStartHeaders = headers
End Function

'====================================================================
' ZONE HEADER FORMATTING HELPERS
' Used by WorkfileBuilder to apply zone-colored headers on the Working Sheet.
'====================================================================

Public Sub ApplyZoneHeader(rng As Range, bgColor As Long, fgColor As Long)
    With rng
        .Interior.Color = bgColor
        .Font.Color = fgColor
        .Font.Bold = True
        .Font.name = "Segoe UI"
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub

Public Sub ApplyZoneCategoryLabel(ws As Worksheet, catRow As Long, startCol As Long, endCol As Long, _
                                   labelText As String, bgColor As Long)
    If endCol < startCol Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(catRow, startCol), ws.Cells(catRow, endCol))
    If endCol > startCol Then
        On Error Resume Next
        rng.Merge
        On Error GoTo 0
    End If
    With rng.Cells(1, 1)
        .Value = labelText
        .Interior.Color = bgColor
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.name = "Segoe UI"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ' Apply bg to all cells in case merge failed
    rng.Interior.Color = bgColor
End Sub

Public Sub ApplyTitleBar(ws As Worksheet, lastCol As Long, titleText As String)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    On Error Resume Next
    rng.Merge
    On Error GoTo 0
    With rng.Cells(1, 1)
        .Value = titleText
        .Font.name = "Segoe UI"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    rng.Interior.Color = THEME_BG
    ws.Rows(1).RowHeight = 36
End Sub

Public Sub ApplySubtitleBar(ws As Worksheet, lastCol As Long, subtitleText As String)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol))
    On Error Resume Next
    rng.Merge
    On Error GoTo 0
    With rng.Cells(1, 1)
        .Value = subtitleText
        .Font.name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = THEME_TEXT_SEC
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    rng.Interior.Color = THEME_SURFACE
    ws.Rows(2).RowHeight = 22
End Sub
