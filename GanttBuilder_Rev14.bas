Attribute VB_Name = "GanttBuilder"
'====================================================================
' Gantt Builder Module - Rev14 - Weekly Gantt Chart on Working Sheet
'
' Architecture:
'   TEXT:   Formula-based nested IFs (auto-updates on data change)
'   COLOR:  Conditional Formatting rules (auto-updates on data change)
'           No VBA color painting needed - fully dynamic!
'
' Rev11 changes from Rev10:
'   - Updated NIF rebuild call to NIF_Builder_Rev11 (split HC tables)
'
' Rev10 changes from Rev9:
'   - Removed duplicate GetShortAbbrev/GetLastWord (use TISCommon public versions)
'   - Gantt start date now configurable via Definitions!S2 (falls back to constants)
'   - Replaced Debug.Print with DebugLog (conditional compilation via TISCommon)
'   - Quarter row uses THEME_ACCENT from TISCommon for consistent styling
'   - Dark theme: THEME_SURFACE for month/week headers, THEME_BORDER for borders
'   - MRCL start date = SQ Finish + 1 day (duration_offset type)
'
' Rev9 changes from Rev8:
'   - Shared utility functions (ColLetter, SheetExists, FindWorkingSheet,
'     ShellSortVariantArray) consolidated into TISCommon module
'
' Rev8 changes from Rev7:
'   - BuildGantt accepts optional targetSheet parameter so
'     WorkfileBuilder can direct Gantt to the correct (new) sheet
'     when creating versioned sheets (Working Sheet2, etc.)
'
' Rev7 changes from Rev6:
'   - MRCL changed from single-date to duration (Start+Finish)
'   - Legend, Quarter, Month rows moved to 12-14 (above date headers)
'   - Shared utilities moved to TISCommon module
'
' v2.2 Design:
'   - CF Rule per phase: =cellValue="SET" -> SET color
'   - CF Rule priority order = color priority
'   - SDD special rule: checks actual SDD date vs week date
'     (highest CF priority, paints yellow even when text shows other phase)
'   - Text formula priority: nested IF, highest textPri checked first
'
' Features:
'   - 10-col gap after data, default start 01/19/2026
'   - Week dates via formulas (+7)
'   - Duration phases: SET, SL1, SL2, SQ, CV, PF, DC, DM, MRCL
'   - Single-date: SDD
'   - Quarter (Fiscal Q1=Nov, row 12) + Month (row 13) + Legend (row 14)
'   - Week date headers at row 15 (DATA_START_ROW)
'   - Today marker (green)
'   - RefreshGantt rebuilds, no separate color refresh needed
'   - Reused/Demo red text CF overlay
'
' Color Palette (from reference):
'   SET=Cyan  SL1=LightPink  SL2=HotPink  SQ=LightGreen
'   CV=DarkRed  PF=Cream  DC=LightOrange  DM=LightPurple
'   IQ=Gray  SDD=BrightYellow
'====================================================================

Option Explicit

Private Const SHEET_WORKING_BASE As String = "Working Sheet"
Private Const SHEET_DEFINITIONS As String = "Definitions"
Private Const DATA_START_ROW As Long = TIS_DATA_START_ROW
Private Const GANTT_WEEKS As Long = 104
Private Const GANTT_CELL_WIDTH As Double = 3
Private Const GANTT_COL_GAP As Long = 10
Private Const DEFAULT_START_YEAR As Long = 2026
Private Const DEFAULT_START_MONTH As Long = 1
Private Const DEFAULT_START_DAY As Long = 19
Private Const EXTRA_PALETTE As String = "255,218,185|176,224,230|221,160,221|240,230,140|211,211,211|188,143,143"

'====================================================================
' MAIN ENTRY: BUILD GANTT
'====================================================================

Public Sub BuildGantt(Optional silent As Boolean = False, Optional targetSheet As Worksheet = Nothing)
    On Error GoTo ErrorHandler

    Dim startTime As Double
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalcMode As XlCalculation
    Dim ws As Worksheet
    Dim ganttStartDate As Date
    Dim phaseDict As Object
    Dim ganttStartCol As Long
    Dim ganttExists As Boolean
    Dim userChoice As VbMsgBoxResult

    ' Use provided target sheet or find latest Working Sheet
    If Not targetSheet Is Nothing Then
        Set ws = targetSheet
    Else
        Set ws = FindWorkingSheet()
    End If
    If ws Is Nothing Then
        If Not silent Then MsgBox "No Working Sheet found. Run Create Work File first.", vbExclamation
        Exit Sub
    End If
    
    ganttExists = GanttChartExists(ws)
    
    ' If Gantt exists, ask user for confirmation to recreate (skip if silent)
    If ganttExists Then
        If silent Then
            ' Auto-recreate when called from WorkfileBuilder
        Else
            userChoice = MsgBox("A Gantt chart already exists on this sheet." & vbCrLf & vbCrLf & _
                               "Do you want to delete and recreate it?" & vbCrLf & vbCrLf & _
                               "(Your data will NOT be affected)", _
                               vbYesNo + vbQuestion, "Gantt Chart Exists")
            If userChoice <> vbYes Then
                Exit Sub
            End If
        End If
    End If

    startTime = Timer
    prevScreenUpdating = Application.screenUpdating
    prevEnableEvents = Application.enableEvents
    prevCalcMode = Application.Calculation

    Application.screenUpdating = False
    Application.enableEvents = False
    Application.Calculation = xlCalculationManual

    ganttStartDate = GetAlignedStartDate()
    ClearExistingGantt ws

    Set phaseDict = BuildPhaseDefinitions(ws)
    If phaseDict.Count = 0 Then
        MsgBox "No milestone phases found. Check Definitions sheet.", vbExclamation
        GoTo Cleanup
    End If

    ganttStartCol = GetGanttStartColumn(ws)

    ' Find BOD (Blackout Date) columns
    Dim bod1Col As Long, bod2Col As Long, bj As Long
    Dim bhVal As String
    bod1Col = 0: bod2Col = 0
    For bj = 1 To ganttStartCol - 1
        bhVal = LCase(Trim(CStr(ws.Cells(DATA_START_ROW, bj).Value)))
        If bhVal = "bod1" Then bod1Col = bj
        If bhVal = "bod2" Then bod2Col = bj
    Next bj

    ' Build structure: headers + formulas
    RenderGanttStructure ws, ganttStartDate, phaseDict, ganttStartCol, bod1Col, bod2Col

    ' Force formula evaluation so CF rules can match Gantt cell values.
    ' Without this, Gantt cells are unevaluated under xlCalculationManual
    ' and CF rules see empty cells instead of phase abbreviations.
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    ws.Calculate
    Application.Calculation = prevCalc

    ' Apply conditional formatting rules for colors
    ApplyGanttConditionalFormatting ws, ganttStartCol, phaseDict, bod1Col, bod2Col

    ' Today marker AFTER phase CF so it isn't deleted by ganttRange.FormatConditions.Delete
    ApplyTodayMarker ws, ganttStartCol, GetDataLastRow(ws)

    ' Rebuild NIF section (Gantt clear wipes shared columns)
    ' Only call NIF from here when GanttBuilder is invoked standalone (no targetSheet)
    ' When called from WorkfileBuilder, it handles NIF separately with sourceSheet
    If targetSheet Is Nothing Then
        NIF_Builder.BuildNIF silent:=True, targetSheet:=ws
        On Error GoTo ErrorHandler
    End If

    ' Recalc
    If Not silent Then
        Application.Calculation = xlCalculationAutomatic
        ws.Calculate
    End If

    If Not silent Then
        Application.screenUpdating = True
        MsgBox "Gantt chart built successfully!" & vbCrLf & _
               "Range: " & Format(ganttStartDate, "mm/dd/yyyy") & " + 2 years" & vbCrLf & _
               "Time: " & Format(Timer - startTime, "0.00") & "s", vbInformation
        Application.screenUpdating = False
    End If

    GoTo Cleanup

ErrorHandler:
    MsgBox "Error in BuildGantt: " & Err.Description & vbCrLf & _
           "Error #: " & Err.Number, vbCritical

Cleanup:
    If Application.Calculation <> prevCalcMode Then Application.Calculation = prevCalcMode
    Application.enableEvents = prevEnableEvents
    Application.screenUpdating = prevScreenUpdating
End Sub

'====================================================================
' GET CONFIGURED START DATE
' Reads Definitions!S2 for a user-configured start date.
' Falls back to DEFAULT_START_YEAR/MONTH/DAY constants if S2 is empty
' or not a valid date.
'====================================================================

Private Function GetConfiguredStartDate() As Date
    Dim wsDef As Worksheet
    Dim cellVal As Variant

    On Error Resume Next
    If SheetExists(ThisWorkbook, SHEET_DEFINITIONS) Then
        Set wsDef = ThisWorkbook.Sheets(SHEET_DEFINITIONS)
        cellVal = wsDef.Range("S2").Value
        If IsDate(cellVal) Then
            GetConfiguredStartDate = CDate(cellVal)
            DebugLog "GanttBuilder: Using configured start date from Definitions!S2: " & Format(GetConfiguredStartDate, "mm/dd/yyyy")
            Exit Function
        End If
    End If
    On Error GoTo 0

    ' Fall back to hard-coded defaults
    GetConfiguredStartDate = DateSerial(DEFAULT_START_YEAR, DEFAULT_START_MONTH, DEFAULT_START_DAY)
    DebugLog "GanttBuilder: Using default start date: " & Format(GetConfiguredStartDate, "mm/dd/yyyy")
End Function

Private Function GetAlignedStartDate() As Date
    Dim d As Date
    d = GetConfiguredStartDate()
    d = d - (Weekday(d, vbMonday) - 1)
    GetAlignedStartDate = d
End Function

Private Function GetGanttStartColumn(ws As Worksheet) As Long
    Dim lastCol As Long
    Dim tbl As ListObject
    Dim j As Long
    Dim fullLastCol As Long
    
    ' Use ListObject table boundary if available (excludes Gantt/NIF columns)
    If ws.ListObjects.Count > 0 Then
        Set tbl = ws.ListObjects(1)
        lastCol = tbl.Range.Column + tbl.Range.Columns.Count - 1
    Else
        ' Fallback: scan row for GANTT_START marker first
        fullLastCol = ws.Cells(DATA_START_ROW, ws.Columns.Count).End(xlToLeft).Column
        
        ' Check if GANTT_START exists -- use the column just before it
        For j = 1 To fullLastCol + GANTT_WEEKS + 20
            If j > ws.Columns.Count Then Exit For
            If CStr(ws.Cells(6, j).Value) = "GANTT_START" Then
                lastCol = j - GANTT_COL_GAP
                GetGanttStartColumn = j  ' Reuse same start position
                Exit Function
            End If
        Next j
        
        lastCol = fullLastCol
    End If
    
    GetGanttStartColumn = lastCol + GANTT_COL_GAP
End Function

'====================================================================
' CHECK IF GANTT CHART EXISTS
' Looks for the GANTT_START marker in row 6
'====================================================================

Private Function GanttChartExists(ws As Worksheet) As Boolean
    Dim lastCol As Long
    Dim j As Long

    GanttChartExists = False
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = ws.Cells(DATA_START_ROW, ws.Columns.Count).End(xlToLeft).Column

    For j = 1 To lastCol + GANTT_WEEKS + 20
        If j > ws.Columns.Count Then Exit For
        If CStr(ws.Cells(6, j).Value) = "GANTT_START" Then
            GanttChartExists = True
            Exit For
        End If
    Next j
End Function

'====================================================================
' GET DATA LAST ROW
' Uses ListObject boundary to find true data extent, excluding HC tables
' below the data. Falls back to End(xlUp) on column 1 if no ListObject.
'====================================================================

Private Function GetDataLastRow(ws As Worksheet) As Long
    Dim lo As ListObject
    On Error Resume Next
    For Each lo In ws.ListObjects
        If lo.Range.row <= DATA_START_ROW + 1 Then
            GetDataLastRow = lo.Range.row + lo.Range.Rows.Count - 1
            On Error GoTo 0
            Exit Function
        End If
    Next lo
    On Error GoTo 0
    ' Fallback
    GetDataLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If GetDataLastRow < DATA_START_ROW + 1 Then GetDataLastRow = DATA_START_ROW + 1
End Function

'====================================================================
' CLEAR EXISTING GANTT
'====================================================================

Private Sub ClearExistingGantt(ws As Worksheet)
    Dim lastCol As Long
    Dim j As Long
    Dim clearEnd As Long
    Dim usedLastRow As Long
    Dim ganttClearEnd As Long

    ' Search for GANTT_START marker in row 6
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = ws.Cells(DATA_START_ROW, ws.Columns.Count).End(xlToLeft).Column

    For j = 1 To lastCol + GANTT_WEEKS + 20
        If j > ws.Columns.Count Then Exit For
        If CStr(ws.Cells(6, j).Value) = "GANTT_START" Then
            ' Only clear the Gantt columns (GANTT_START to GANTT_START + GANTT_WEEKS + small pad)
            ' Do NOT clear NIF section after the Gantt
            ganttClearEnd = j + GANTT_WEEKS + 5
            If ganttClearEnd > ws.Columns.Count Then ganttClearEnd = ws.Columns.Count
            
            ' Clear rows 1 to max used row but only Gantt columns
            Dim ur As Range
            Set ur = ws.UsedRange
            usedLastRow = ur.Rows.Count + ur.row - 1
            ws.Range(ws.Cells(1, j), ws.Cells(usedLastRow, ganttClearEnd)).Clear
            Exit For
        End If
    Next j
End Sub

'====================================================================
' BUILD PHASE DEFINITIONS
' Returns dict: key=abbreviation, value=Array(9 elements)
'   0=type, 1=R, 2=G, 3=B, 4=textPri, 5=colorPri,
'   6=startCol/singleCol, 7=endCol/"", 8=shortAbbrev
'====================================================================

Private Function BuildPhaseDefinitions(ws As Worksheet) As Object
    Dim dict As Object
    Dim wsDef As Worksheet
    Dim defLastRow As Long
    Dim defData As Variant
    Dim i As Long
    Dim fText As String
    Dim gText As String
    Dim headerName As String
    Dim tokens As Variant
    Dim tokenVal As Variant
    Dim letter As String
    Dim num As Long
    Dim milGroups As Object
    Dim milNames As Object
    Dim sortedKeys As Variant
    Dim idx As Long
    Dim abbreviation As String
    Dim startHeader As String
    Dim endHeader As String
    Dim colorR As Long
    Dim colorG As Long
    Dim colorB As Long
    Dim extraColors As Variant
    Dim extraIdx As Long
    Dim phaseOrder As Long
    Dim textPri As Long
    Dim colorPri As Long
    Dim displayAbbrev As String

    Set dict = CreateObject("Scripting.Dictionary")

    If Not SheetExists(ThisWorkbook, SHEET_DEFINITIONS) Then
        Set BuildPhaseDefinitions = dict
        Exit Function
    End If
    Set wsDef = ThisWorkbook.Sheets(SHEET_DEFINITIONS)

    defLastRow = wsDef.Cells(wsDef.Rows.Count, 1).End(xlUp).row
    If defLastRow < 2 Then
        Set BuildPhaseDefinitions = dict
        Exit Function
    End If
    defData = wsDef.Range(wsDef.Cells(1, 1), wsDef.Cells(defLastRow, 7)).Value

    Set milGroups = CreateObject("Scripting.Dictionary")
    Set milNames = CreateObject("Scripting.Dictionary")

    ' Parse milestone tokens from Definitions F/G
    For i = 2 To UBound(defData, 1)
        fText = Trim(CStr(defData(i, 6)))
        gText = Trim(CStr(defData(i, 7)))
        headerName = Trim(CStr(defData(i, 1)))

        If fText <> "" Then
            tokens = Split(fText, "|")
            For Each tokenVal In tokens
                tokenVal = UCase(Trim(tokenVal))
                If Len(tokenVal) >= 2 And IsNumeric(Mid(tokenVal, 2)) Then
                    letter = Left(tokenVal, 1)
                    num = CLng(Mid(tokenVal, 2))
                    If Not milGroups.exists(letter) Then
                        Set milGroups(letter) = CreateObject("Scripting.Dictionary")
                    End If
                    milGroups(letter)(num) = headerName
                    If num = 1 And gText <> "" Then milNames(letter) = gText
                End If
            Next tokenVal
        End If
    Next i

    ' Build duration phases from milestone groups
    extraColors = Split(EXTRA_PALETTE, "|")
    extraIdx = 0
    phaseOrder = 0

    If milGroups.Count > 0 Then
        sortedKeys = milGroups.keys
        ShellSortVariantArray sortedKeys

        For idx = LBound(sortedKeys) To UBound(sortedKeys)
            letter = CStr(sortedKeys(idx))
            If milGroups(letter).exists(1) And milGroups(letter).exists(2) Then
                phaseOrder = phaseOrder + 1
                abbreviation = ""
                If milNames.exists(letter) Then abbreviation = UCase(milNames(letter))
                If abbreviation = "" Then abbreviation = UCase(letter)

                startHeader = milGroups(letter)(1)
                endHeader = milGroups(letter)(2)
                
                ' Extract display abbreviation from milestone name (fully dynamic)
                displayAbbrev = UCase(GetLastWord(abbreviation))
                ' Normalize CV/CONVERSION
                If displayAbbrev = "CONVERSION" Then displayAbbrev = "CV"
                
                AssignPhaseColor displayAbbrev, colorR, colorG, colorB, extraColors, extraIdx

                If displayAbbrev = "CV" Then
                    textPri = 999: colorPri = 899  ' CV = highest text priority, 2nd color priority (after SDD)
                Else
                    ' First milestone = highest priority (500-phaseOrder, so SET=499, SL1=498, etc.)
                    textPri = 500 - phaseOrder: colorPri = 500 - phaseOrder
                End If

                ' Use milestone name from Definitions (fully dynamic)
                dict(displayAbbrev) = Array("duration", colorR, colorG, colorB, textPri, colorPri, _
                                           startHeader, endHeader, displayAbbrev)
            End If
        Next idx
    End If

    ' PF, DC, DM as duration phases (lower priority than definition milestones)
    AddDurationPhaseFromHeaders dict, ws, "PF", "Prefac", 191, 191, 191, 52, 52    ' #BFBFBF gray
    AddDurationPhaseFromHeaders dict, ws, "DC", "Decon", 255, 192, 0, 51, 51      ' #FFC000 gold
    AddDurationPhaseFromHeaders dict, ws, "DM", "Demo", 233, 113, 50, 50, 50      ' #E97132 burnt orange

    ' MRCL: Start = SQ Finish + 1 day (duration_offset type)
    ' Scan row 15 for SQ Finish and MRCL Finish headers
    Dim mrclJ As Long, mrclLastCol As Long, mrclHv As String
    Dim sqFinishHeader As String, mrclFinishHeader As String
    sqFinishHeader = "": mrclFinishHeader = ""
    mrclLastCol = ws.Cells(DATA_START_ROW, ws.Columns.Count).End(xlToLeft).Column
    For mrclJ = 1 To mrclLastCol
        mrclHv = LCase(Trim(CStr(ws.Cells(DATA_START_ROW, mrclJ).Value)))
        If sqFinishHeader = "" Then
            If InStr(1, mrclHv, "sq", vbTextCompare) > 0 And InStr(1, mrclHv, "finish", vbTextCompare) > 0 Then
                sqFinishHeader = CStr(ws.Cells(DATA_START_ROW, mrclJ).Value)
            End If
        End If
        If mrclFinishHeader = "" Then
            If InStr(1, mrclHv, "mrcl", vbTextCompare) > 0 And InStr(1, mrclHv, "finish", vbTextCompare) > 0 Then
                mrclFinishHeader = CStr(ws.Cells(DATA_START_ROW, mrclJ).Value)
            End If
        End If
    Next mrclJ

    If sqFinishHeader <> "" And mrclFinishHeader <> "" Then
        ' Store as duration_offset: start from SQ Finish col (+1 day applied in formula), end at MRCL Finish col
        dict("MRCL") = Array("duration_offset", 168, 130, 255, 40, 550, sqFinishHeader, mrclFinishHeader, "MRCL")
    Else
        ' Fallback: use standard duration phase if SQ Finish header not found
        AddDurationPhaseFromHeaders dict, ws, "MRCL", "MRCL", 168, 130, 255, 40, 550  ' #A882FF
    End If

    ' SDD single-date: lowest text(1), highest color(900)
    AddSingleDatePhase dict, ws, "SDD", "SDD", 255, 255, 0, 1, 900

    Set BuildPhaseDefinitions = dict
End Function

'====================================================================
' ADD DURATION PHASE FROM HEADERS
'====================================================================

Private Sub AddDurationPhaseFromHeaders(dict As Object, ws As Worksheet, _
                                         abbrev As String, searchTerm As String, _
                                         cr As Long, cG As Long, cB As Long, _
                                         txtPri As Long, clrPri As Long)
    Dim lastCol As Long
    Dim j As Long
    Dim hv As String
    Dim startName As String
    Dim endName As String

    If dict.exists(UCase(abbrev)) Then Exit Sub
    lastCol = ws.Cells(DATA_START_ROW, ws.Columns.Count).End(xlToLeft).Column
    startName = "": endName = ""

    For j = 1 To lastCol
        hv = CStr(ws.Cells(DATA_START_ROW, j).Value)
        If startName = "" Then
            If InStr(1, LCase(hv), LCase(searchTerm), vbTextCompare) > 0 And _
               InStr(1, LCase(hv), "start", vbTextCompare) > 0 Then startName = hv
        End If
        If endName = "" Then
            If InStr(1, LCase(hv), LCase(searchTerm), vbTextCompare) > 0 And _
               (InStr(1, LCase(hv), "finish", vbTextCompare) > 0 Or _
                InStr(1, LCase(hv), "end", vbTextCompare) > 0) Then endName = hv
        End If
    Next j

    If startName <> "" And endName <> "" Then
        dict(UCase(abbrev)) = Array("duration", cr, cG, cB, txtPri, clrPri, _
                                     startName, endName, GetShortAbbrev(abbrev))
    End If
End Sub

'====================================================================
' ADD SINGLE-DATE PHASE
'====================================================================

Private Sub AddSingleDatePhase(dict As Object, ws As Worksheet, _
                                abbrev As String, searchTerm As String, _
                                cr As Long, cG As Long, cB As Long, _
                                txtPri As Long, clrPri As Long)
    Dim lastCol As Long
    Dim j As Long
    Dim hv As String
    Dim colName As String
    Dim colIdx As Long

    If dict.exists(UCase(abbrev)) Then Exit Sub
    lastCol = ws.Cells(DATA_START_ROW, ws.Columns.Count).End(xlToLeft).Column
    colName = ""
    colIdx = 0

    ' Search for column - EXACT MATCH ONLY for SDD to avoid matching "PK SDD"
    For j = 1 To lastCol
        hv = Trim(CStr(ws.Cells(DATA_START_ROW, j).Value))
        ' Exact match
        If LCase(hv) = LCase(Trim(searchTerm)) Or LCase(hv) = LCase(Trim(abbrev)) Then
            colName = hv
            colIdx = j
            Exit For
        End If
    Next j
    
    ' If not found and NOT SDD, try partial match (skip for SDD to avoid "PK SDD")
    If colIdx = 0 And UCase(abbrev) <> "SDD" Then
        For j = 1 To lastCol
            hv = Trim(CStr(ws.Cells(DATA_START_ROW, j).Value))
            If InStr(1, LCase(hv), LCase(searchTerm), vbTextCompare) > 0 Or _
               InStr(1, LCase(hv), LCase(abbrev), vbTextCompare) > 0 Then
                colName = hv
                colIdx = j
                Exit For
            End If
        Next j
    End If

    If colName <> "" And colIdx > 0 Then
        ' Store column index in element 7 for direct access
        dict(UCase(abbrev)) = Array("single", cr, cG, cB, txtPri, clrPri, _
                                     colName, CStr(colIdx), GetShortAbbrev(abbrev))
    End If
End Sub

'====================================================================
' ASSIGN PHASE COLOR
'====================================================================

Private Sub AssignPhaseColor(abbrev As String, ByRef r As Long, ByRef g As Long, ByRef b As Long, _
                              extraColors As Variant, ByRef extraIdx As Long)
    Dim parts As Variant
    Select Case UCase(abbrev)
        Case "SET": r = 56: g = 189: b = 248        ' #38BDF8 vivid sky blue
        Case "SL1": r = 74: g = 222: b = 128        ' #4ADE80 vivid green
        Case "SL2": r = 250: g = 204: b = 21        ' #FACC15 vivid golden yellow
        Case "SQ": r = 99: g = 179: b = 237         ' #63B3ED medium blue
        Case "CV", "CONVERSION": r = 251: g = 146: b = 60  ' #FB923C vivid orange
        Case "PF": r = 148: g = 163: b = 184        ' #94A3B8 slate gray
        Case "DC": r = 251: g = 191: b = 36         ' #FBBF24 amber
        Case "DM": r = 239: g = 68: b = 68          ' #EF4444 vivid red
        Case "MRCL": r = 168: g = 130: b = 255       ' #A882FF vivid purple
        Case Else
            If extraIdx <= UBound(extraColors) Then
                parts = Split(extraColors(extraIdx), ",")
                r = CLng(parts(0)): g = CLng(parts(1)): b = CLng(parts(2))
                extraIdx = extraIdx + 1
            Else
                r = 180: g = 180: b = 180
            End If
    End Select
End Sub

' GetShortAbbrev and GetLastWord removed in Rev10 - use TISCommon public versions

'====================================================================
' RENDER GANTT STRUCTURE (headers + text formulas)
'====================================================================

Private Sub RenderGanttStructure(ws As Worksheet, ganttStart As Date, phaseDict As Object, _
                                 ganttStartCol As Long, bod1Col As Long, bod2Col As Long)
    Dim lastDataRow As Long
    Dim lastTableCol As Long
    Dim w As Long
    Dim r As Long
    Dim j As Long
    Dim gapCol As Long
    Dim hVal As String
    Dim wsHeaderMap As Object
    Dim phaseColMap As Object
    Dim phaseKey As Variant
    Dim phaseInfo As Variant
    Dim pType As String
    Dim sIdx As Long
    Dim eIdx As Long

    ' Place marker in row 6
    ws.Cells(6, ganttStartCol).Value = "GANTT_START"
    ws.Cells(6, ganttStartCol).Font.Color = THEME_BG
    ws.Cells(6, ganttStartCol).Font.Size = 1

    lastDataRow = GetDataLastRow(ws)
    lastTableCol = ganttStartCol - GANTT_COL_GAP

    ' Build header map from Working Sheet (bulk array read -- single COM call)
    Set wsHeaderMap = CreateObject("Scripting.Dictionary")
    If lastTableCol >= 1 Then
        Dim hdrArr As Variant
        hdrArr = ws.Range(ws.Cells(DATA_START_ROW, 1), ws.Cells(DATA_START_ROW, lastTableCol)).Value
        For j = 1 To lastTableCol
            hVal = CStr(hdrArr(1, j) & "")
            If hVal <> "" Then wsHeaderMap(LCase(Trim(hVal))) = j
        Next j
    End If

    ' Rev14: Redirect Gantt to read Our Date columns instead of TIS date columns.
    ' Phase definitions reference TIS header names (e.g. "Set Start", "SL1 Signoff Finish").
    ' If the corresponding Our Date column exists in the header map, override the TIS header
    ' entry to point at the Our Date column position. Milestones without Our Date equivalents
    ' (PF, DC, DM) fall through unchanged.
    Dim ourDateRedirect As Object
    Set ourDateRedirect = CreateObject("Scripting.Dictionary")
    ourDateRedirect(LCase(TIS_SRC_SET)) = LCase(TIS_COL_OUR_SET)
    ourDateRedirect(LCase(TIS_SRC_SL1)) = LCase(TIS_COL_OUR_SL1)
    ourDateRedirect(LCase(TIS_SRC_SL2)) = LCase(TIS_COL_OUR_SL2)
    ourDateRedirect(LCase(TIS_SRC_SQ)) = LCase(TIS_COL_OUR_SQ)
    ourDateRedirect(LCase(TIS_SRC_CONVS)) = LCase(TIS_COL_OUR_CONVS)
    ourDateRedirect(LCase(TIS_SRC_CONVF)) = LCase(TIS_COL_OUR_CONVF)
    ourDateRedirect(LCase(TIS_SRC_MRCLS)) = LCase(TIS_COL_OUR_MRCLS)
    ourDateRedirect(LCase(TIS_SRC_MRCLF)) = LCase(TIS_COL_OUR_MRCLF)

    Dim rdKey As Variant
    Dim ourKey As String
    For Each rdKey In ourDateRedirect.Keys
        ourKey = ourDateRedirect(rdKey)
        If wsHeaderMap.exists(ourKey) Then
            ' Our Date column found -- redirect TIS header lookup to Our Date column position
            wsHeaderMap(CStr(rdKey)) = wsHeaderMap(ourKey)
        End If
    Next rdKey

    ' Phase column indices
    Set phaseColMap = CreateObject("Scripting.Dictionary")
    For Each phaseKey In phaseDict.keys
        phaseInfo = phaseDict(phaseKey)
        pType = CStr(phaseInfo(0))
        sIdx = 0: eIdx = 0
        If pType = "duration" Or pType = "duration_offset" Then
            If wsHeaderMap.exists(LCase(Trim(CStr(phaseInfo(6))))) Then sIdx = wsHeaderMap(LCase(Trim(CStr(phaseInfo(6)))))
            If wsHeaderMap.exists(LCase(Trim(CStr(phaseInfo(7))))) Then eIdx = wsHeaderMap(LCase(Trim(CStr(phaseInfo(7)))))
        Else
            If CStr(phaseInfo(7)) <> "" And IsNumeric(CStr(phaseInfo(7))) Then
                sIdx = CLng(CStr(phaseInfo(7)))
            ElseIf wsHeaderMap.exists(LCase(Trim(CStr(phaseInfo(6))))) Then
                sIdx = wsHeaderMap(LCase(Trim(CStr(phaseInfo(6)))))
            End If
        End If
        phaseColMap(CStr(phaseKey)) = Array(sIdx, eIdx)
    Next phaseKey

    ' Headers
    WriteQuarterHeaders ws, ganttStartCol, ganttStart
    WriteMonthHeaders ws, ganttStartCol, ganttStart
    WriteWeekHeaders ws, ganttStartCol, ganttStart

    ' Find Status column for active-only Gantt filtering
    Dim statusCol As Long
    statusCol = 0
    If wsHeaderMap.exists(LCase(TIS_COL_STATUS)) Then statusCol = wsHeaderMap(LCase(TIS_COL_STATUS))

    ' Text formulas - write first row, then FillDown
    WriteGanttFormulas ws, ganttStartCol, lastDataRow, phaseDict, phaseColMap, bod1Col, bod2Col, statusCol

    ' Today marker moved to BuildGantt (after ApplyGanttConditionalFormatting)
    ' to avoid being deleted by ganttRange.FormatConditions.Delete

    ' Column widths
    For w = 0 To GANTT_WEEKS - 1
        ws.Columns(ganttStartCol + w).ColumnWidth = GANTT_CELL_WIDTH
    Next w

    ' Dark background on Gantt data cells
    Dim ganttEndCol As Long
    ganttEndCol = ganttStartCol + GANTT_WEEKS - 1
    Dim firstRow As Long
    firstRow = DATA_START_ROW + 1
    ws.Range(ws.Cells(firstRow, ganttStartCol), ws.Cells(lastDataRow, ganttEndCol)).Interior.Color = THEME_BG
    ws.Range(ws.Cells(firstRow, ganttStartCol), ws.Cells(lastDataRow, ganttEndCol)).Font.Color = THEME_TEXT

    ' Gap columns (narrow white spacer between data table and Gantt)
    For gapCol = lastTableCol + 1 To ganttStartCol - 1
        ws.Columns(gapCol).ColumnWidth = 0.8
        ws.Columns(gapCol).Interior.Color = RGB(255, 255, 255)
    Next gapCol

    ' Legend
    WriteColorLegend ws, ganttStartCol, phaseDict
End Sub

'====================================================================
' WRITE WEEK HEADERS WITH FORMULAS
'====================================================================

Private Sub WriteWeekHeaders(ws As Worksheet, startCol As Long, ganttStart As Date)
    Dim w As Long
    Dim col As Long

    For w = 0 To GANTT_WEEKS - 1
        col = startCol + w
        If w = 0 Then
            ws.Cells(DATA_START_ROW, col).Value = ganttStart
        Else
            ws.Cells(DATA_START_ROW, col).formula = "=" & ColLetter(col - 1) & DATA_START_ROW & "+7"
        End If

        With ws.Cells(DATA_START_ROW, col)
            .NumberFormat = "m/d"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = THEME_TEXT_SEC
            .Interior.Color = THEME_SURFACE
            .WrapText = False
            .Orientation = 90
        End With
    Next w
    ws.Rows(DATA_START_ROW).RowHeight = 32
End Sub

'====================================================================
' WRITE GANTT FORMULAS (first data row, then FillDown)
'====================================================================

Private Sub WriteGanttFormulas(ws As Worksheet, ganttStartCol As Long, lastDataRow As Long, _
                                phaseDict As Object, phaseColMap As Object, _
                                bod1Col As Long, bod2Col As Long, _
                                Optional statusCol As Long = 0)
    Dim w As Long
    Dim col As Long
    Dim firstDataRow As Long
    Dim weekDateRef As String
    Dim formula As String
    Dim phaseKey As Variant
    Dim phaseInfo As Variant
    Dim colIndices As Variant
    Dim pType As String
    Dim sIdx As Long
    Dim eIdx As Long
    Dim sLetter As String
    Dim eLetter As String
    Dim abbrev As String
    Dim condition As String
    Dim phaseCount As Long
    Dim pi As Long
    Dim a As Long
    Dim b As Long
    Dim tempStr As String
    Dim tempLng As Long
    Dim phaseKeys() As String
    Dim phaseTxtPris() As Long
    Dim ganttEndCol As Long
    Dim rw As Long
    Dim bodCondition As String
    Dim bod1Letter As String, bod2Letter As String
    Dim bodOverridePhases As String
    Dim innerFormula As String
    Dim bod2Cond As String

    firstDataRow = DATA_START_ROW + 1
    
    ' BOD column letters for formula references
    If bod1Col > 0 Then bod1Letter = "$" & ColLetter(bod1Col)
    If bod2Col > 0 Then bod2Letter = "$" & ColLetter(bod2Col)
    
    ' Phases that take priority OVER blackout dates
    ' (pre-fac, SDD, Demo, Decon, MRCL -- these can't be blacked out)
    bodOverridePhases = """PF"",""SDD"",""DC"",""DM"",""MRCL"""

    ' Sort phases by textPriority descending
    phaseCount = phaseDict.Count
    If phaseCount = 0 Then Exit Sub

    ReDim phaseKeys(1 To phaseCount)
    ReDim phaseTxtPris(1 To phaseCount)

    pi = 0
    For Each phaseKey In phaseDict.keys
        pi = pi + 1
        phaseKeys(pi) = CStr(phaseKey)
        phaseTxtPris(pi) = CLng(phaseDict(phaseKey)(4))
    Next phaseKey

    For a = 1 To phaseCount - 1
        For b = a + 1 To phaseCount
            If phaseTxtPris(a) < phaseTxtPris(b) Then
                tempStr = phaseKeys(a): phaseKeys(a) = phaseKeys(b): phaseKeys(b) = tempStr
                tempLng = phaseTxtPris(a): phaseTxtPris(a) = phaseTxtPris(b): phaseTxtPris(b) = tempLng
            End If
        Next b
    Next a

    ' Write formula for each week column in first data row, then FillDown
    For w = 0 To GANTT_WEEKS - 1
        col = ganttStartCol + w
        weekDateRef = ColLetter(col) & "$" & DATA_START_ROW

        formula = ""

        For pi = phaseCount To 1 Step -1
            phaseInfo = phaseDict(phaseKeys(pi))
            pType = CStr(phaseInfo(0))
            colIndices = phaseColMap(phaseKeys(pi))
            sIdx = CLng(colIndices(0))
            eIdx = CLng(colIndices(1))
            abbrev = CStr(phaseInfo(8))
            condition = ""

            If pType = "duration" And sIdx > 0 And eIdx > 0 Then
                sLetter = "$" & ColLetter(sIdx)
                eLetter = "$" & ColLetter(eIdx)
                condition = "AND(ISNUMBER(" & sLetter & firstDataRow & ")," & _
                           "ISNUMBER(" & eLetter & firstDataRow & ")," & _
                           sLetter & firstDataRow & "<=" & weekDateRef & "+6," & _
                           eLetter & firstDataRow & ">=" & weekDateRef & ")"
            ElseIf pType = "duration_offset" And sIdx > 0 And eIdx > 0 Then
                ' sIdx = SQ Finish column, eIdx = MRCL Finish column
                ' MRCL starts the day AFTER SQ Finish (+1)
                sLetter = "$" & ColLetter(sIdx)
                eLetter = "$" & ColLetter(eIdx)
                condition = "AND(ISNUMBER(" & sLetter & firstDataRow & ")," & _
                           "ISNUMBER(" & eLetter & firstDataRow & ")," & _
                           sLetter & firstDataRow & "+1<=" & weekDateRef & "+6," & _
                           eLetter & firstDataRow & ">=" & weekDateRef & ")"
            ElseIf pType = "single" And sIdx > 0 Then
                sLetter = "$" & ColLetter(sIdx)
                condition = "AND(ISNUMBER(" & sLetter & firstDataRow & ")," & _
                           sLetter & firstDataRow & ">=" & weekDateRef & "," & _
                           sLetter & firstDataRow & "<=" & weekDateRef & "+6)"
            End If

            If condition <> "" Then
                If formula = "" Then
                    formula = "IF(" & condition & ",""" & abbrev & ""","""")"
                Else
                    formula = "IF(" & condition & ",""" & abbrev & """," & formula & ")"
                End If
            End If
        Next pi

        If formula <> "" Then
            ' Wrap with BOD (Blackout Date) logic:
            ' BOD1 = blackout start date, BOD2 = blackout end date (range)
            ' Week overlaps BOD range if: BOD1<=weekEnd AND BOD2>=weekStart
            ' Show "BOD" UNLESS the inner result is a priority phase
            If bod1Col > 0 And bod2Col > 0 Then
                innerFormula = formula
                
                ' Range overlap: BOD1 <= weekDateRef+6 AND BOD2 >= weekDateRef
                bodCondition = "AND(ISNUMBER(" & bod1Letter & firstDataRow & ")," & _
                              "ISNUMBER(" & bod2Letter & firstDataRow & ")," & _
                              bod1Letter & firstDataRow & "<=" & weekDateRef & "+6," & _
                              bod2Letter & firstDataRow & ">=" & weekDateRef & ")"
                
                ' Final formula: LET(phase, innerFormula,
                '   IF(AND(bodCondition, NOT(ISNUMBER(MATCH(phase,{"PF","SDD","DC","DM","MRCL"},0)))), "BOD", phase))
                formula = "LET(phase," & innerFormula & "," & _
                         "IF(AND(" & bodCondition & "," & _
                         "NOT(ISNUMBER(MATCH(phase,{" & bodOverridePhases & "},0))))," & _
                         """BOD"",phase))"
            End If
            
            ' Rev14: Only show Gantt for Active systems
            If statusCol > 0 Then
                Dim statusLetter As String
                statusLetter = "$" & ColLetter(statusCol)
                formula = "IF(" & statusLetter & firstDataRow & "=""Active""," & formula & ","""")"
            End If

            On Error Resume Next
            ws.Cells(firstDataRow, col).formula = "=" & formula
            If Err.Number <> 0 Then
                Err.Clear
                ws.Cells(firstDataRow, col).Formula2 = "=" & formula
                If Err.Number <> 0 Then
                    Err.Clear
                    ws.Cells(firstDataRow, col).Value = ""
                End If
            End If
            On Error GoTo 0

            If lastDataRow > firstDataRow Then
                ws.Range(ws.Cells(firstDataRow, col), ws.Cells(lastDataRow, col)).FillDown
            End If
        End If

        With ws.Range(ws.Cells(firstDataRow, col), ws.Cells(lastDataRow, col))
            .Font.Size = 9
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next w
    
    ' Add horizontal borders between project rows (bulk -- single COM call)
    ganttEndCol = ganttStartCol + GANTT_WEEKS - 1
    If lastDataRow >= firstDataRow Then
        With ws.Range(ws.Cells(firstDataRow, ganttStartCol), ws.Cells(lastDataRow, ganttEndCol)).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = THEME_BORDER
        End With
        ' Bottom edge of last row
        With ws.Range(ws.Cells(lastDataRow, ganttStartCol), ws.Cells(lastDataRow, ganttEndCol)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = THEME_BORDER
        End With
    End If
End Sub

'====================================================================
' APPLY GANTT CONDITIONAL FORMATTING
'====================================================================

Private Sub ApplyGanttConditionalFormatting(ws As Worksheet, ganttStartCol As Long, phaseDict As Object, _
                                            bod1Col As Long, bod2Col As Long)
    Dim lastDataRow As Long
    Dim firstDataRow As Long
    Dim ganttEndCol As Long
    Dim ganttRange As Range
    Dim fc As FormatCondition
    Dim phaseKey As Variant
    Dim phaseInfo As Variant
    Dim abbrev As String
    Dim bgR As Long
    Dim bgG As Long
    Dim bgB As Long
    Dim fontR As Long
    Dim fontG As Long
    Dim fontB As Long
    Dim brightness As Long
    Dim firstCellAddr As String
    Dim phaseCount As Long
    Dim i As Long
    Dim a As Long
    Dim b As Long
    Dim tempStr As String
    Dim tempLng As Long
    Dim lastTableCol As Long
    Dim sddColIdx As Long
    Dim sddColLetter As String
    Dim weekDateAddr As String
    Dim sddFormula As String
    Dim hj As Long
    Dim sddColName As String
    Dim newReusedColIdx As Long
    Dim newReusedColLetter As String
    Dim reusedFormula As String
    Dim weekRefAddr As String
    Dim sddRowRef As String
    Dim rawHdr As String
    Dim phaseKeys() As String
    Dim phaseClrPris() As Long

    firstDataRow = DATA_START_ROW + 1
    ' Use ListObject boundary to avoid including HC tables below data.
    ' End(xlUp) on column 1 picks up HC table content, making ganttRange 100K+ cells
    ' and causing FormatConditions.Add to fail silently on the oversized range.
    lastDataRow = GetDataLastRow(ws)
    ganttEndCol = ganttStartCol + GANTT_WEEKS - 1
    lastTableCol = ganttStartCol - GANTT_COL_GAP

    Set ganttRange = ws.Range(ws.Cells(firstDataRow, ganttStartCol), _
                               ws.Cells(lastDataRow, ganttEndCol))

    On Error Resume Next
    ganttRange.FormatConditions.Delete
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    firstCellAddr = ws.Cells(firstDataRow, ganttStartCol).Address(False, False)

    ' Find New/Reused column for red text on Reused/Demo systems
    newReusedColIdx = 0
    For hj = 1 To lastTableCol
        rawHdr = LCase(Trim(Replace(Replace(CStr(ws.Cells(DATA_START_ROW, hj).Value), vbLf, ""), vbCr, "")))
        If rawHdr = "new/reused" Or InStr(1, rawHdr, "new/reused", vbTextCompare) > 0 Then
            newReusedColIdx = hj
            Exit For
        End If
    Next hj

    ' Sort phases by colorPriority ASCENDING
    phaseCount = phaseDict.Count
    If phaseCount = 0 Then Exit Sub

    ReDim phaseKeys(1 To phaseCount)
    ReDim phaseClrPris(1 To phaseCount)

    i = 0
    For Each phaseKey In phaseDict.keys
        i = i + 1
        phaseKeys(i) = CStr(phaseKey)
        phaseClrPris(i) = CLng(phaseDict(phaseKey)(5))
    Next phaseKey

    For a = 1 To phaseCount - 1
        For b = a + 1 To phaseCount
            If phaseClrPris(a) > phaseClrPris(b) Then
                tempStr = phaseKeys(a): phaseKeys(a) = phaseKeys(b): phaseKeys(b) = tempStr
                tempLng = phaseClrPris(a): phaseClrPris(a) = phaseClrPris(b): phaseClrPris(b) = tempLng
            End If
        Next b
    Next a

    ' Reused/Demo red text: Priority=1 so red font takes precedence over phase backgrounds.
    ' This rule only sets font color (no background), so phase backgrounds still show.
    If newReusedColIdx > 0 Then
        newReusedColLetter = "$" & ColLetter(newReusedColIdx)
        reusedFormula = "=OR(" & newReusedColLetter & firstDataRow & "=""Reused""," & _
                       newReusedColLetter & firstDataRow & "=""Demo"")"

        On Error Resume Next
        Set fc = ganttRange.FormatConditions.Add(Type:=xlExpression, Formula1:=reusedFormula)
        If Err.Number = 0 And Not fc Is Nothing Then
            fc.Font.Color = RGB(192, 0, 0)
            fc.Font.Bold = True
            fc.StopIfTrue = False
            On Error Resume Next
            fc.priority = 1
            On Error GoTo 0
        Else
            DebugLog "CF ADD FAILED for Reused/Demo: " & Err.Description
        End If
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If

    ' BOD (Blackout Date) rule: black cell with white "BOD" text
    ' StopIfTrue = True so black overrides all other phase colors
    On Error Resume Next
    Set fc = ganttRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=" & firstCellAddr & "=""BOD""")
    If Err.Number = 0 And Not fc Is Nothing Then
        fc.Interior.Color = RGB(0, 0, 0)
        fc.Font.Color = RGB(255, 255, 255)
        fc.Font.Bold = True
        fc.StopIfTrue = True
    Else
        DebugLog "CF ADD FAILED for BOD: " & Err.Description
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    For i = 1 To phaseCount
        phaseInfo = phaseDict(phaseKeys(i))
        abbrev = CStr(phaseInfo(8))
        bgR = CLng(phaseInfo(1)): bgG = CLng(phaseInfo(2)): bgB = CLng(phaseInfo(3))

        brightness = (bgR + bgG + bgB) \ 3
        If brightness < 160 Then
            fontR = 255: fontG = 255: fontB = 255
        Else
            fontR = 40: fontG = 40: fontB = 40
        End If

        If UCase(phaseKeys(i)) = "SDD" Then
            sddColIdx = 0
            If CStr(phaseInfo(7)) <> "" And IsNumeric(CStr(phaseInfo(7))) Then
                sddColIdx = CLng(CStr(phaseInfo(7)))
            Else
                sddColName = CStr(phaseInfo(6))
                For hj = 1 To lastTableCol
                    If LCase(Trim(CStr(ws.Cells(DATA_START_ROW, hj).Value))) = LCase(Trim(sddColName)) Or _
                       LCase(Trim(CStr(ws.Cells(DATA_START_ROW, hj).Value))) = "sdd" Then
                        sddColIdx = hj
                        Exit For
                    End If
                Next hj
            End If

            If sddColIdx > 0 Then
                sddColLetter = "$" & ColLetter(sddColIdx)
                weekRefAddr = ws.Cells(DATA_START_ROW, ganttStartCol).Address(True, False)
                sddRowRef = sddColLetter & firstDataRow

                sddFormula = "=AND(ISNUMBER(" & sddRowRef & ")," & _
                            sddRowRef & ">=" & weekRefAddr & "," & _
                            sddRowRef & "<=" & weekRefAddr & "+6)"

                On Error Resume Next
                Set fc = ganttRange.FormatConditions.Add(Type:=xlExpression, Formula1:=sddFormula)
                If Err.Number = 0 And Not fc Is Nothing Then
                    fc.Interior.Color = RGB(bgR, bgG, bgB)
                    fc.Font.Color = RGB(fontR, fontG, fontB)
                    fc.Font.Bold = True
                    fc.StopIfTrue = True
                    On Error Resume Next
                    fc.priority = 1
                    On Error GoTo 0
                Else
                    DebugLog "CF ADD FAILED for SDD: " & Err.Description
                End If
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
            End If
        Else
            On Error Resume Next
            Set fc = ganttRange.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=" & firstCellAddr & "=""" & abbrev & """")
            If Err.Number = 0 And Not fc Is Nothing Then
                fc.Interior.Color = RGB(bgR, bgG, bgB)
                fc.Font.Color = RGB(fontR, fontG, fontB)
                fc.Font.Bold = True
                fc.StopIfTrue = False
            Else
                DebugLog "CF ADD FAILED for " & abbrev & ": " & Err.Description
            End If
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
    Next i
End Sub

'====================================================================
' QUARTER HEADERS (row 12 -- directly above month/date headers)
'====================================================================

Private Sub WriteQuarterHeaders(ws As Worksheet, startCol As Long, ganttStart As Date)
    Dim w As Long
    Dim weekDate As Date
    Dim currentQtr As String
    Dim prevQtr As String
    Dim qtrStartCol As Long

    prevQtr = "": qtrStartCol = startCol

    For w = 0 To GANTT_WEEKS - 1
        weekDate = ganttStart + (w * 7)
        currentQtr = GetFiscalQuarter(weekDate)
        If currentQtr <> prevQtr Then
            If prevQtr <> "" And w > 0 Then
                FormatMergedHeader ws, 12, qtrStartCol, startCol + w - 1, prevQtr, THEME_ACCENT, 9
            End If
            qtrStartCol = startCol + w
            prevQtr = currentQtr
        End If
    Next w
    If prevQtr <> "" Then FormatMergedHeader ws, 12, qtrStartCol, startCol + GANTT_WEEKS - 1, prevQtr, THEME_ACCENT, 9
End Sub

Private Function GetFiscalQuarter(d As Date) As String
    Dim mon As Long
    Dim fiscalYear As Long
    Dim fiscalQtr As Long
    mon = Month(d)
    Select Case mon
        Case 11, 12: fiscalQtr = 1: fiscalYear = Year(d) + 1
        Case 1: fiscalQtr = 1: fiscalYear = Year(d)
        Case 2, 3, 4: fiscalQtr = 2: fiscalYear = Year(d)
        Case 5, 6, 7: fiscalQtr = 3: fiscalYear = Year(d)
        Case 8, 9, 10: fiscalQtr = 4: fiscalYear = Year(d)
    End Select
    GetFiscalQuarter = "Q" & fiscalQtr & "'" & Right(CStr(fiscalYear), 2)
End Function

'====================================================================
' MONTH HEADERS (row 13)
'====================================================================

Private Sub WriteMonthHeaders(ws As Worksheet, startCol As Long, ganttStart As Date)
    Dim w As Long
    Dim weekDate As Date
    Dim currentMonth As Long
    Dim currentYear As Long
    Dim prevMonth As Long
    Dim prevYear As Long
    Dim monthStartCol As Long
    Dim monthLabel As String

    prevMonth = 0: prevYear = 0: monthStartCol = startCol

    For w = 0 To GANTT_WEEKS - 1
        weekDate = ganttStart + (w * 7)
        currentMonth = Month(weekDate): currentYear = Year(weekDate)
        If currentMonth <> prevMonth Or currentYear <> prevYear Then
            If prevMonth > 0 And w > 0 Then
                monthLabel = Format(DateSerial(prevYear, prevMonth, 1), "mmm yy")
                FormatMergedHeader ws, 13, monthStartCol, startCol + w - 1, monthLabel, THEME_SURFACE, 7
            End If
            monthStartCol = startCol + w
            prevMonth = currentMonth: prevYear = currentYear
        End If
    Next w
    If prevMonth > 0 Then
        monthLabel = Format(DateSerial(prevYear, prevMonth, 1), "mmm yy")
        FormatMergedHeader ws, 13, monthStartCol, startCol + GANTT_WEEKS - 1, monthLabel, THEME_SURFACE, 7
    End If
End Sub

'====================================================================
' FORMAT MERGED HEADER
'====================================================================

Private Sub FormatMergedHeader(ws As Worksheet, rowNum As Long, c1 As Long, c2 As Long, _
                                text As String, bgColor As Long, fontSize As Long)
    Dim rng As Range
    If c2 < c1 Then c2 = c1
    Set rng = ws.Range(ws.Cells(rowNum, c1), ws.Cells(rowNum, c2))
    On Error Resume Next
    rng.Merge
    On Error GoTo 0
    rng.Cells(1, 1).Value = text
    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = fontSize
        .Font.Color = THEME_TEXT
        .Interior.Color = bgColor
    End With
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlHairline: .Color = THEME_BORDER
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlHairline: .Color = THEME_BORDER
    End With
End Sub

'====================================================================
' TODAY MARKER
'====================================================================

Public Sub ApplyTodayMarker(ws As Worksheet, ganttStartCol As Long, lastDataRow As Long)
    ' CF-based today marker: dynamically highlights the current week column.
    ' Uses TODAY() so it auto-updates without rebuilding.
    Dim ganttEndCol As Long
    Dim firstDataRow As Long
    Dim dataRange As Range
    Dim headerRange As Range
    Dim fc As FormatCondition
    Dim weekDateRef As String
    Dim todayFormula As String
    Dim w As Long, col As Long

    ganttEndCol = ganttStartCol + GANTT_WEEKS - 1
    firstDataRow = DATA_START_ROW + 1
    If lastDataRow < firstDataRow Then Exit Sub

    ' Formula checks if the week date in row 15 of this column contains TODAY
    ' Using mixed ref: column-relative, row-absolute to row 15
    weekDateRef = ColLetter(ganttStartCol) & "$" & DATA_START_ROW
    todayFormula = "=AND(ISNUMBER(" & weekDateRef & ")," & weekDateRef & "<=TODAY()," & weekDateRef & "+6>=TODAY())"

    ' CF on data rows: subtle green tint on today's column
    On Error Resume Next
    Set dataRange = ws.Range(ws.Cells(firstDataRow, ganttStartCol), _
                              ws.Cells(lastDataRow, ganttEndCol))
    Set fc = dataRange.FormatConditions.Add(Type:=xlExpression, Formula1:=todayFormula)
    If Not fc Is Nothing Then
        fc.Interior.Color = RGB(20, 60, 40)
        fc.StopIfTrue = False
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    ' CF on header rows 12, 13, 15: green background
    On Error Resume Next
    Set headerRange = ws.Range(ws.Cells(12, ganttStartCol), ws.Cells(12, ganttEndCol))
    Set fc = headerRange.FormatConditions.Add(Type:=xlExpression, Formula1:=todayFormula)
    If Not fc Is Nothing Then
        fc.Interior.Color = RGB(0, 128, 0)
        fc.Font.Color = RGB(255, 255, 255)
        fc.StopIfTrue = False
    End If
    If Err.Number <> 0 Then Err.Clear

    Set headerRange = ws.Range(ws.Cells(13, ganttStartCol), ws.Cells(13, ganttEndCol))
    Set fc = headerRange.FormatConditions.Add(Type:=xlExpression, Formula1:=todayFormula)
    If Not fc Is Nothing Then
        fc.Interior.Color = RGB(0, 128, 0)
        fc.Font.Color = RGB(255, 255, 255)
        fc.StopIfTrue = False
    End If
    If Err.Number <> 0 Then Err.Clear

    Set headerRange = ws.Range(ws.Cells(DATA_START_ROW, ganttStartCol), ws.Cells(DATA_START_ROW, ganttEndCol))
    Set fc = headerRange.FormatConditions.Add(Type:=xlExpression, Formula1:=todayFormula)
    If Not fc Is Nothing Then
        fc.Interior.Color = RGB(0, 128, 0)
        fc.Font.Color = RGB(255, 255, 255)
        fc.StopIfTrue = False
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    ' Row 11: formula-driven arrow indicator
    ' Guarded: row 11 may have merged cells from WorkfileBuilder header area
    For w = 0 To GANTT_WEEKS - 1
        col = ganttStartCol + w
        On Error Resume Next
        ws.Cells(11, col).formula = "=IF(AND(ISNUMBER(" & ColLetter(col) & "$" & DATA_START_ROW & ")," & _
                                    ColLetter(col) & "$" & DATA_START_ROW & "<=TODAY()," & _
                                    ColLetter(col) & "$" & DATA_START_ROW & "+6>=TODAY()),""" & ChrW(9660) & ""","""")"
        ws.Cells(11, col).Font.Color = RGB(0, 128, 0)
        ws.Cells(11, col).Font.Size = 10
        ws.Cells(11, col).HorizontalAlignment = xlCenter
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    Next w
End Sub

'====================================================================
' COLOR LEGEND (row 14 -- directly above date headers at row 15)
'====================================================================

Private Sub WriteColorLegend(ws As Worksheet, ganttStartCol As Long, phaseDict As Object)
    Dim col As Long
    Dim phaseKey As Variant
    Dim phaseInfo As Variant
    Dim phaseCount As Long
    Dim i As Long
    Dim a As Long
    Dim b As Long
    Dim tempStr As String
    Dim tempLng As Long
    Dim bgR As Long
    Dim bgG As Long
    Dim bgB As Long
    Dim brightness As Long
    Dim phaseKeys() As String
    Dim phasePris() As Long

    phaseCount = phaseDict.Count
    If phaseCount = 0 Then Exit Sub

    ReDim phaseKeys(1 To phaseCount)
    ReDim phasePris(1 To phaseCount)

    i = 0
    For Each phaseKey In phaseDict.keys
        i = i + 1
        phaseKeys(i) = CStr(phaseKey)
        phasePris(i) = CLng(phaseDict(phaseKey)(5))
    Next phaseKey

    For a = 1 To phaseCount - 1
        For b = a + 1 To phaseCount
            If phasePris(a) > phasePris(b) Then
                tempStr = phaseKeys(a): phaseKeys(a) = phaseKeys(b): phaseKeys(b) = tempStr
                tempLng = phasePris(a): phasePris(a) = phasePris(b): phasePris(b) = tempLng
            End If
        Next b
    Next a

    col = ganttStartCol
    ws.Cells(14, col).Value = "Legend:"
    ws.Cells(14, col).Font.Bold = True
    ws.Cells(14, col).Font.Size = 7
    ws.Cells(14, col).Font.Color = THEME_TEXT_SEC
    col = col + 1

    For i = 1 To phaseCount
        phaseInfo = phaseDict(phaseKeys(i))
        bgR = CLng(phaseInfo(1)): bgG = CLng(phaseInfo(2)): bgB = CLng(phaseInfo(3))

        ws.Cells(14, col).Interior.Color = RGB(bgR, bgG, bgB)
        ws.Cells(14, col).Value = CStr(phaseInfo(8))
        ws.Cells(14, col).Font.Size = 6
        ws.Cells(14, col).Font.Bold = True
        ws.Cells(14, col).HorizontalAlignment = xlCenter
        ws.Cells(14, col).VerticalAlignment = xlCenter

        brightness = (bgR + bgG + bgB) \ 3
        If brightness < 160 Then
            ws.Cells(14, col).Font.Color = RGB(255, 255, 255)
        Else
            ws.Cells(14, col).Font.Color = RGB(40, 40, 40)
        End If

        col = col + 1
    Next i

    ws.Cells(14, col).Interior.Color = RGB(0, 128, 0)
    ws.Cells(14, col).Value = "TODAY"
    ws.Cells(14, col).Font.Size = 6
    ws.Cells(14, col).Font.Bold = True
    ws.Cells(14, col).Font.Color = RGB(255, 255, 255)
    ws.Cells(14, col).HorizontalAlignment = xlCenter
    col = col + 1

    ' BOD legend entry
    ws.Cells(14, col).Interior.Color = RGB(0, 0, 0)
    ws.Cells(14, col).Value = "BOD"
    ws.Cells(14, col).Font.Size = 6
    ws.Cells(14, col).Font.Bold = True
    ws.Cells(14, col).Font.Color = RGB(255, 255, 255)
    ws.Cells(14, col).HorizontalAlignment = xlCenter

    ws.Rows(14).RowHeight = 12
End Sub

