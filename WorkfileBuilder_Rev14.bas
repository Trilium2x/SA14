Attribute VB_Name = "WorkfileBuilder"
'====================================================================
' Work File Builder Module - Rev14
'
' Rev14 changes:
'   - Two-path architecture: FRESH BUILD (no Working Sheet) vs
'     UPDATE IN PLACE (existing Rev14 schema detected)
'   - UpdateWorkingSheet: fast in-place update — compares TIS against
'     existing Working Sheet, updates changed cells (orange fill +
'     comment), cancels removed systems, reintroduces returning
'     systems, appends new systems with Our Dates from TIS
'   - Schema detection via "Conv.S" header sentinel
'   - Full rebuild path (CreateWorkingSheet) preserved for first-time
'     builds and schema upgrades from pre-Rev14 workbooks
'
' Rev11 changes:
'   - In-place rebuild: reuses existing "Working Sheet" instead of
'     creating versioned copies ("Working Sheet 2", etc.)
'   - BackupWorkingSheet creates dated backup ("Old YYYY-MM-DD")
'   - ClearWorkingSheet fully resets the sheet object for reuse
'   - GetUniqueWorkingSheetName removed (no longer needed)
'   - NIF_Builder_Rev11 called instead of Rev10
'   - Error handler no longer deletes sheet on in-place rebuild
'
' Rev10 changes:
'   - Dark SaaS theme UI overhaul (THEME_BG, THEME_SURFACE, THEME_ACCENT palette)
'   - Dashboard counter formulas use multi-date fallback for Reused systems
'   - DebugLog replaced with DebugLog conditional wrapper
'   - Color palette now references TISCommon THEME_* constants
'   - Dashboard auto-call removed (now button-only from TIS Tracker sheet)
'   - Sheet background set to THEME_BG, gridlines hidden
'   - Magic number additionalColCount replaced with named constant
'
' Rev9 changes:
'   - Shared utility functions consolidated into TISCommon module
'   - Sort logic delegated to TISCommon.SortWithHelperColumns
'   - Dead sirfisHeaders parameter removed from ApplySortingWithHelperColumns
'
' Rev8.1 changes:
'   1. Dashboard slicer fix: Worksheet_Calculate event handler installed
'      automatically — forces dashboard recalc when slicers/filters change
'   2. Application.Calculation forced to xlAutomatic at end (prevents
'      stuck Manual mode from prior crashes)
'   3. OFFSET pattern updated to BKM: explicit height=1, width=1 params
'   4. Removed unused helper columns (h_New, h_Reused, h_Demo, h_Visible)
'      — h_Active/h_ActiveDemo removed; active counts use inline SUMPRODUCT
'   5. Color palette centralized as module-level constants (17 colors)
'   6. Progress bar via Application.StatusBar during build
'
' Rev8 changes from Rev7:
'   1. Source sheet passed to GanttBuilder + NIF_Builder so user data
'      (NIF assignments, HC values) migrates to versioned sheets
'   2. Active count bug fix: Reused systems exclude pre-fac events
'      (uses Set Start for lifecycle start, not pre-fac phases)
'   3. Added "Completed" column (TRUE/FALSE) — merged into Status column in Rev14
'   4. Added Event Type slicer alongside Group slicer
'
' Rev7 features preserved:
'   - Clean rebuild: calls GanttBuilder + NIF_Builder after data build
'   - Active count shows Active New/Reused/Demo breakdown
'   - "CT est miss" dashboard card
'   - CT Threshold read from Definitions!S1
'   - Time frame filter with weeks input
'   - Escalated + Watched counter cards
'   - Group Slicer on ListObject table
'   - Import system: user data, change detection, removed systems archival
'====================================================================

Option Explicit

' Sheet constants
Private Const SHEET_TIS As String = "TIS"
Private Const SHEET_DEFINITIONS As String = "Definitions"
Private Const SHEET_CEIDS As String = "CEIDs"
Private Const SHEET_MILESTONES As String = "Milestones"
Private Const SHEET_WORKING As String = "Working Sheet"
Private Const SHEET_NEW_REUSED As String = "New-reused"
Private Const SHEET_SN As String = "SN"
Private Const SHEET_REMOVED As String = "Removed Systems"
Private Const DATA_START_ROW As Long = TIS_DATA_START_ROW
Private Const NIF_EMPLOYEE_COUNT As Long = 5

' Required headers
Private Const HEADER_ENTITY_CODE As String = "Entity Code"
Private Const HEADER_EVENT_TYPE As String = "Event Type"
Private Const HEADER_SDD As String = "SDD"
Private Const HEADER_SUPPLIER_QUAL_FINISH As String = "Supplier Qual Finish"
Private Const HEADER_SET_START As String = "Set Start"

' Defaults
Private Const DEFAULT_CYCLE_TIME_THRESHOLD As Long = 85

' Color palette constants - Light theme (SA9 style)
Private Const CLR_DARK_TEXT As Long = 5258796        ' RGB(44, 62, 80)
Private Const CLR_MUTED_TEXT As Long = 9208440       ' RGB(120, 130, 140)
Private Const CLR_TIMESTAMP As Long = 11184928       ' RGB(160, 160, 170)
Private Const CLR_CARD_BG As Long = 16710908         ' RGB(252, 252, 254)
Private Const CLR_CARD_BORDER As Long = 15458780     ' RGB(220, 225, 235)
Private Const CLR_NEW As Long = 4564776              ' RGB(40, 167, 69)
Private Const CLR_REUSED As Long = 13400576          ' RGB(0, 122, 204)
Private Const CLR_DEMO As Long = 10903957            ' RGB(149, 97, 166)
Private Const CLR_CT_MISS As Long = 4535772          ' RGB(220, 53, 69)
Private Const CLR_ESCALATED As Long = 2250751        ' RGB(255, 87, 34)
Private Const CLR_WATCHED As Long = 39167            ' RGB(255, 152, 0)
Private Const CLR_ACTIVE_LABEL As Long = 39167       ' RGB(255, 152, 0)
Private Const CLR_HEADER_BG As Long = 9910032        ' RGB(16, 55, 151)
Private Const CLR_HEADER_TEXT As Long = 16777215     ' RGB(255, 255, 255)
Private Const CLR_TF_INPUT_BG As Long = 16446965     ' RGB(245, 245, 250)
Private Const CLR_TF_INPUT_BORDER As Long = 13813960 ' RGB(200, 200, 210)

' Additional column count estimate (New/Reused, Escalated, Tool S/N, Ship Date,
' Pre-Install Meeting, Est CAR Date, Est Cycle Time, SOC Available, SOC Uploaded,
' Staffed, Comments, BOD1, BOD2 = 13 visible + 2 hidden helpers = 15)
Private Const ESTIMATED_ADDITIONAL_COLS As Long = 27

' Module-level cache
Private m_GroupCache As Object

'====================================================================
' MAIN ENTRY POINT
'====================================================================

Public Sub CreateWorkFile()
    On Error GoTo ErrorHandler
    
    Dim startTime As Double
    Dim screenUpdating As Boolean, enableEvents As Boolean
    Dim calcMode As XlCalculation
    Dim errOccurred As Boolean
    
    startTime = Timer
    
    screenUpdating = Application.screenUpdating
    enableEvents = Application.enableEvents
    calcMode = Application.Calculation
    
    Application.screenUpdating = False
    Application.enableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set m_GroupCache = CreateObject("Scripting.Dictionary")

    ' Rev14: Unprotect Working Sheet if it exists (protection applied at end in Cleanup)
    Dim wsToUnprotect As Worksheet
    ' Validate prerequisites
    If Not SheetExists(ThisWorkbook, SHEET_TIS) Then
        MsgBox "The '" & SHEET_TIS & "' sheet does not exist.", vbExclamation
        GoTo Cleanup
    End If
    If Not SheetExists(ThisWorkbook, SHEET_DEFINITIONS) Then
        MsgBox "The '" & SHEET_DEFINITIONS & "' sheet does not exist.", vbExclamation
        GoTo Cleanup
    End If
    
    Dim hasCEIDsSheet As Boolean, hasMilestonesSheet As Boolean
    hasCEIDsSheet = SheetExists(ThisWorkbook, SHEET_CEIDS)
    hasMilestonesSheet = SheetExists(ThisWorkbook, SHEET_MILESTONES)
    
    ' Build warnings string for the confirmation dialog
    Dim buildWarnings As String
    buildWarnings = ""
    If Not hasCEIDsSheet Then buildWarnings = buildWarnings & vbCrLf & Chr(9874) & " CEIDs sheet not found. Group column will be empty."
    If Not hasMilestonesSheet Then buildWarnings = buildWarnings & vbCrLf & Chr(9874) & " Milestones sheet not found. STD Duration columns will be skipped."

    ' Pre-check which milestones are missing from Milestones sheet
    If hasMilestonesSheet Then
        Dim msMissing As String
        msMissing = ""
        Dim msSheet As Worksheet
        Set msSheet = ThisWorkbook.Sheets(SHEET_MILESTONES)
        Dim msLastRow As Long
        msLastRow = msSheet.Cells(msSheet.Rows.Count, 1).End(xlUp).Row
        ' Check for MRCL and Conversion milestones (commonly missing)
        Dim msNames As Variant
        msNames = Array("MRCL", "Conversion")
        Dim msIdx As Long
        For msIdx = LBound(msNames) To UBound(msNames)
            Dim msFound As Boolean
            msFound = False
            Dim msR As Long
            For msR = 2 To msLastRow
                If InStr(1, CStr(msSheet.Cells(msR, 1).Value), msNames(msIdx), vbTextCompare) > 0 Then
                    msFound = True: Exit For
                End If
            Next msR
            If Not msFound Then msMissing = msMissing & "  - " & msNames(msIdx) & vbCrLf
        Next msIdx
        If msMissing <> "" Then
            buildWarnings = buildWarnings & vbCrLf & "Milestones not found (STD Duration skipped):" & vbCrLf & msMissing
        End If
    End If

    ' Detect existing Working Sheet
    Dim oldSheet As Worksheet
    Dim doImport As Boolean
    Set oldSheet = FindLatestWorkingSheet()
    doImport = False

    ' Rev14: Check if existing sheet has Rev14 schema (sentinel: "Conv.S" header)
    Dim hasRev14Schema As Boolean
    hasRev14Schema = False
    If Not oldSheet Is Nothing Then
        Dim schemaChkHdr As Long
        schemaChkHdr = TISCommon.FindHeaderRow(oldSheet)
        If schemaChkHdr > 0 Then
            If TISCommon.FindHeaderCol(oldSheet, schemaChkHdr, TIS_COL_OUR_CONVS, _
                    oldSheet.Cells(schemaChkHdr, oldSheet.Columns.Count).End(xlToLeft).Column) > 0 Then
                hasRev14Schema = True
            End If
        End If
    End If

    ' === FULL BUILD PATH (always rebuild — TIS update mechanism will be added later) ===
    If True Then
        If Not oldSheet Is Nothing Then
            Application.ScreenUpdating = True
            If MsgBox("Rebuild Working Sheet with Rev14 schema?" & vbCrLf & vbCrLf & _
                      "This will:" & vbCrLf & _
                      "  " & Chr(149) & " Back up current sheet as 'Old " & Format(Date, "YYYY-MM-DD") & "'" & vbCrLf & _
                      "  " & Chr(149) & " Create new Working Sheet with Our Date columns" & vbCrLf & _
                      "  " & Chr(149) & " Import your data (Escalated, Comments, NIF, BOD)" & vbCrLf & _
                      "  " & Chr(149) & " Build Gantt chart" & _
                      IIf(buildWarnings <> "", vbCrLf & vbCrLf & "Warnings:" & buildWarnings, ""), _
                      vbOKCancel + vbQuestion, "Build Working Sheet") = vbCancel Then
                GoTo Cleanup
            End If
            Application.ScreenUpdating = False
            doImport = True
        End If

        errOccurred = False
        CreateWorkingSheet hasCEIDsSheet, hasMilestonesSheet, errOccurred, oldSheet, doImport

        If errOccurred Then GoTo Cleanup
    End If

    ' Show completion in status bar (no second dialog)
    Application.StatusBar = "Working Sheet built in " & Format(Timer - startTime, "0.0") & "s"
    
    GoTo Cleanup

ErrorHandler:
    MsgBox "Error in CreateWorkFile: " & Err.Description & vbCrLf & _
           "Error number: " & Err.Number, vbCritical
    
Cleanup:
    Set m_GroupCache = Nothing
    Application.StatusBar = False  ' Clear status bar

    ' Rev14: Store named ranges for Lock? Data Validation handler
    Dim wsToCheck As Worksheet
    On Error Resume Next
    Set wsToCheck = ThisWorkbook.Worksheets(TIS_SHEET_WORKING)
    If Not wsToCheck Is Nothing Then
        Dim pLockCol As Long, pOurStart As Long, pOurEnd As Long
        Dim pHdr As Long, pC As Long
        pLockCol = 0: pOurStart = 0: pOurEnd = 0
        pHdr = TISCommon.FindHeaderRow(wsToCheck)
        If pHdr > 0 Then
            Dim pMaxC As Long
            pMaxC = wsToCheck.Cells(pHdr, wsToCheck.Columns.Count).End(xlToLeft).Column
            For pC = 1 To pMaxC
                Dim pHV As String
                pHV = LCase(Trim(Replace(Replace(CStr(wsToCheck.Cells(pHdr, pC).Value), vbLf, ""), vbCr, "")))
                If pHV = LCase(TIS_COL_LOCK) Then pLockCol = pC
                If pHV = LCase(TIS_COL_OUR_SET) And pOurStart = 0 Then pOurStart = pC
                If pHV = LCase(TIS_COL_OUR_MRCLF) Then pOurEnd = pC
            Next pC
            If pLockCol > 0 And pOurStart > 0 And pOurEnd >= pOurStart Then
                ThisWorkbook.Names.Add name:="OUR_DATE_START", RefersTo:="=" & pOurStart
                ThisWorkbook.Names.Add name:="OUR_DATE_END", RefersTo:="=" & pOurEnd
                ThisWorkbook.Names.Add name:="LOCK_COL", RefersTo:="=" & pLockCol
            End If
        End If
    End If
    Set wsToCheck = Nothing
    On Error GoTo 0

    ' ALWAYS restore to Automatic so formulas respond to slicers/filters
    Application.Calculation = xlCalculationAutomatic
    Application.enableEvents = enableEvents
    Application.screenUpdating = screenUpdating
    ' Force full recalculation so dashboard formulas pick up current state
    Application.Calculate
End Sub

'====================================================================
' CLEAR WORKING SHEET - Reset existing sheet for in-place rebuild
'====================================================================

Private Sub ClearWorkingSheet(ws As Worksheet)
    On Error Resume Next

    ' 1. Delete slicer caches that reference tables on this sheet
    Dim sc As SlicerCache
    Dim scIdx As Long
    For scIdx = ThisWorkbook.SlicerCaches.Count To 1 Step -1
        Set sc = ThisWorkbook.SlicerCaches(scIdx)
        Dim shouldDel As Boolean: shouldDel = False
        Dim ptSheet As String: ptSheet = ""
        On Error Resume Next
        ptSheet = sc.PivotTable.Parent.Name
        On Error GoTo 0
        If ptSheet = ws.Name Then shouldDel = True
        Dim sl As Slicer
        On Error Resume Next
        For Each sl In sc.Slicers
            If sl.Parent.Name = ws.Name Then shouldDel = True
        Next sl
        On Error GoTo 0
        If shouldDel Then
            On Error Resume Next
            sc.Delete
            On Error GoTo 0
        End If
    Next scIdx

    ' 2. Delete all ListObjects (Excel Tables)
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        lo.Delete
    Next lo

    ' 3. Clear all conditional formatting
    ws.Cells.FormatConditions.Delete

    ' 4. Delete all shapes (buttons, slicers visual, etc.)
    Dim shp As Shape
    Dim si As Long
    For si = ws.Shapes.Count To 1 Step -1
        ws.Shapes(si).Delete
    Next si

    ' 5. Delete all chart objects
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    ' 6. Unmerge all cells
    ws.Cells.UnMerge

    ' 7. Clear outline/grouping
    ws.Cells.ClearOutline

    ' 8. Delete all named ranges scoped to this sheet
    Dim nm As Name
    Dim ni As Long
    For ni = ws.Names.Count To 1 Step -1
        ws.Names(ni).Delete
    Next ni

    ' 10. Remove freeze panes
    ws.Activate
    ActiveWindow.FreezePanes = False

    ' 11. Clear all content, formats, validations
    ws.Cells.Clear

    ' 12. Reset column widths and row heights
    ws.Cells.ColumnWidth = 8.43
    ws.Cells.RowHeight = 15

    On Error GoTo 0
End Sub

'====================================================================
' BACKUP WORKING SHEET - Create dated backup copy
'====================================================================

Private Sub BackupWorkingSheet(ws As Worksheet)
    Dim backupName As String
    backupName = "Old " & Format(Date, "YYYY-MM-DD")

    ' If name already exists (same-day rebuild), append time
    If SheetExists(ThisWorkbook, backupName) Then
        backupName = backupName & " " & Format(Now, "HHh") & Format(Now, "nn")
    End If
    ' Final collision check
    If SheetExists(ThisWorkbook, backupName) Then
        backupName = backupName & "_" & Format(Now, "ss")
    End If

    Application.DisplayAlerts = False
    ws.Copy After:=ws
    ActiveSheet.Name = backupName
    Application.DisplayAlerts = True

    DebugLog "BackupWorkingSheet: created '" & backupName & "'"
End Sub

'====================================================================
' UPDATE WORKING SHEET - FAST IN-PLACE UPDATE (Rev14)
'
' Called when the Working Sheet already has the Rev14 column schema.
' Compares TIS data against existing rows, updates changed TIS fields,
' cancels removed systems, reintroduces returning systems, appends
' new systems with Our Dates auto-populated from TIS.
' Does NOT clear/rebuild the sheet structure — only touches data cells.
' Rebuilds Gantt + NIF at the end (they clear their own sections).
'====================================================================

Private Sub UpdateWorkingSheet(ws As Worksheet)
    On Error GoTo UpdateError

    Dim stepDesc As String
    stepDesc = "Init update"

    ' --- Load TIS data ---
    Dim tisSheet As Worksheet, wsDef As Worksheet
    Set tisSheet = ThisWorkbook.Sheets(SHEET_TIS)
    Set wsDef = ThisWorkbook.Sheets(SHEET_DEFINITIONS)

    ' Get sirfisHeaders from Definitions
    If wsDef.Range("A2").Value = "" Then
        MsgBox "The 'sirfisheaders' range is empty.", vbExclamation
        Exit Sub
    End If
    Dim sirfisHeaders As Range
    Set sirfisHeaders = wsDef.Range("A2", wsDef.Cells(wsDef.Rows.Count, "A").End(xlUp))

    ' Build headerDict from TIS columns
    Dim headerDict As Object, filterDict As Object
    Set headerDict = CreateObject("Scripting.Dictionary")
    Set filterDict = CreateObject("Scripting.Dictionary")

    Dim tisLastCol As Long, tisLastRow As Long
    tisLastCol = tisSheet.Cells(1, tisSheet.Columns.Count).End(xlToLeft).Column
    tisLastRow = tisSheet.Cells(tisSheet.Rows.Count, 1).End(xlUp).Row
    If tisLastRow < 2 Then
        MsgBox "TIS sheet has no data rows.", vbExclamation
        Exit Sub
    End If
    Dim tisData As Variant
    tisData = tisSheet.Range(tisSheet.Cells(1, 1), tisSheet.Cells(tisLastRow, tisLastCol)).Value

    ' Match TIS headers and build filter/sort dicts
    Dim header As Range, colIndex As Long, filterValue As String, sortOrder As Variant
    Dim siteColIdx As Long, entityCodeColIdx As Long, eventTypeColIdx As Long
    siteColIdx = 0: entityCodeColIdx = 0: eventTypeColIdx = 0
    Dim pastDateFilterDict As Object
    Set pastDateFilterDict = CreateObject("Scripting.Dictionary")

    For Each header In sirfisHeaders
        filterValue = Trim(CStr(header.Offset(0, 1).Value))
        sortOrder = Trim(header.Offset(0, 2).Value)
        For colIndex = 1 To tisLastCol
            If LCase(header.Value) = LCase(tisData(1, colIndex)) Then
                headerDict(header.Value) = colIndex
                If filterValue <> "" Then filterDict(header.Value) = Split(filterValue, " ")
                If UCase(sortOrder) = "X" Then
                    If IsDateHeader(header.Value) Then pastDateFilterDict(header.Value) = True
                End If
                If LCase(header.Value) = "site" Then siteColIdx = colIndex
                If LCase(header.Value) = LCase(HEADER_ENTITY_CODE) Then entityCodeColIdx = colIndex
                If LCase(header.Value) = LCase(HEADER_EVENT_TYPE) Then eventTypeColIdx = colIndex
                Exit For
            End If
        Next colIndex
    Next header

    ' --- Build TIS key map (filtered) ---
    stepDesc = "Build TIS key map"
    Dim tisKeyMap As Object
    Set tisKeyMap = CreateObject("Scripting.Dictionary")
    Dim excludeSystemsList As Collection
    Set excludeSystemsList = ParseExcludeSystems(wsDef)

    Dim ti As Long, tisKey As String
    Dim tSite As String, tEC As String, tET As String
    For ti = 2 To tisLastRow
        If RowPassesFilters(tisData, ti, headerDict, filterDict, pastDateFilterDict, _
                            excludeSystemsList, siteColIdx, entityCodeColIdx, eventTypeColIdx) Then
            tSite = LCase(Trim(CStr(tisData(ti, siteColIdx))))
            tEC = LCase(Trim(CStr(tisData(ti, entityCodeColIdx))))
            tET = LCase(Trim(CStr(tisData(ti, eventTypeColIdx))))
            tisKey = tSite & "|" & tEC & "|" & tET
            If tisKey <> "||" And Not tisKeyMap.exists(tisKey) Then
                tisKeyMap(tisKey) = ti
            End If
        End If
    Next ti

    ' --- Build Working Sheet header map and key map ---
    stepDesc = "Build WS maps"
    Dim wsHdr As Long
    wsHdr = TISCommon.FindHeaderRow(ws)
    If wsHdr = 0 Then
        MsgBox "Could not find header row on Working Sheet.", vbCritical
        Exit Sub
    End If
    Dim wsLastRow As Long, wsMaxCol As Long
    wsLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    wsMaxCol = ws.Cells(wsHdr, ws.Columns.Count).End(xlToLeft).Column

    Dim wsHeaderMap As Object  ' LCase header → column index
    Set wsHeaderMap = CreateObject("Scripting.Dictionary")
    Dim wc As Long, whv As String
    For wc = 1 To wsMaxCol
        whv = LCase(Trim(Replace(Replace(CStr(ws.Cells(wsHdr, wc).Value), vbLf, ""), vbCr, "")))
        If whv <> "" And Not wsHeaderMap.exists(whv) Then wsHeaderMap(whv) = wc
    Next wc

    ' Find key columns in Working Sheet
    Dim wsSiteCol As Long, wsECCol As Long, wsETCol As Long, wsStatusCol As Long
    wsSiteCol = 0: wsECCol = 0: wsETCol = 0: wsStatusCol = 0
    If wsHeaderMap.exists("site") Then wsSiteCol = wsHeaderMap("site")
    If wsHeaderMap.exists(LCase(HEADER_ENTITY_CODE)) Then wsECCol = wsHeaderMap(LCase(HEADER_ENTITY_CODE))
    If wsHeaderMap.exists(LCase(HEADER_EVENT_TYPE)) Then wsETCol = wsHeaderMap(LCase(HEADER_EVENT_TYPE))
    If wsHeaderMap.exists(LCase(TIS_COL_STATUS)) Then wsStatusCol = wsHeaderMap(LCase(TIS_COL_STATUS))

    ' Build excludeFromCompare set (key fields and published — not compared for changes)
    Dim excludeFromCompare As Object
    Set excludeFromCompare = CreateObject("Scripting.Dictionary")
    excludeFromCompare("site") = True
    excludeFromCompare(LCase(HEADER_ENTITY_CODE)) = True
    excludeFromCompare(LCase(HEADER_EVENT_TYPE)) = True
    excludeFromCompare("published") = True

    ' --- SCAN WORKING SHEET ROWS ---
    stepDesc = "Update existing rows"
    Application.StatusBar = "Updating Working Sheet... scanning rows"
    Dim r As Long, wsKey As String
    Dim changedCount As Long, cancelledCount As Long, reintroducedCount As Long
    changedCount = 0: cancelledCount = 0: reintroducedCount = 0

    ' Performance: Bulk-read the entire Working Sheet data block into an array
    ' to avoid per-cell COM calls during comparison (1000 rows x 30+ cols = 30K+ calls saved)
    Dim wsDataArr As Variant
    If wsLastRow > wsHdr And wsMaxCol > 0 Then
        wsDataArr = ws.Range(ws.Cells(wsHdr, 1), ws.Cells(wsLastRow, wsMaxCol)).Value
    End If
    ' wsDataArr row 1 = wsHdr (headers), row 2 = first data row, etc.

    ' Pre-build key column arrays from the bulk-read data for fast key construction
    ' (avoids calling BuildProjectKey which does per-cell reads)
    Dim wsDataRows As Long
    wsDataRows = wsLastRow - wsHdr

    ' Collect changed cells and status updates for batch operations
    Dim updateChangedCells As New Collection  ' Array(sheetRow, colIdx, oldValStr)
    Dim updateWriteCells As New Collection    ' Array(sheetRow, colIdx, newValue, isDate)
    Dim statusWrites As New Collection        ' Array(sheetRow, statusColIdx, newStatusStr)

    For r = wsHdr + 1 To wsLastRow
        ' Build project key from array (row offset: r - wsHdr + 1, but array row 1 = wsHdr header)
        Dim wsArrRow As Long
        wsArrRow = r - wsHdr + 1  ' array index for sheet row r
        Dim kS As String, kEC As String, kET As String
        kS = "": kEC = "": kET = ""
        If wsSiteCol > 0 And wsSiteCol <= wsMaxCol Then kS = LCase(Trim(CStr(wsDataArr(wsArrRow, wsSiteCol))))
        If wsECCol > 0 And wsECCol <= wsMaxCol Then kEC = LCase(Trim(CStr(wsDataArr(wsArrRow, wsECCol))))
        If wsETCol > 0 And wsETCol <= wsMaxCol Then kET = LCase(Trim(CStr(wsDataArr(wsArrRow, wsETCol))))
        wsKey = kS & "|" & kEC & "|" & kET
        If wsKey = "||" Then GoTo NextWSRow

        If tisKeyMap.exists(wsKey) Then
            ' --- Project exists in TIS: update TIS columns ---
            Dim tisRow As Long
            tisRow = tisKeyMap(wsKey)

            ' Compare and update each TIS-sourced column (reads from arrays, writes collected for batch)
            Dim hdr As Range, hdrLower As String, wsColIdx As Long, tisColIdx As Long
            Dim wsMapKey As String
            For Each hdr In sirfisHeaders
                hdrLower = LCase(hdr.Value)
                wsMapKey = hdrLower

                If Not excludeFromCompare.exists(hdrLower) Then
                    If wsHeaderMap.exists(wsMapKey) And headerDict.exists(hdr.Value) Then
                        wsColIdx = wsHeaderMap(wsMapKey)
                        tisColIdx = headerDict(hdr.Value)

                        ' Read from arrays instead of cells
                        Dim oldCellVal As Variant, newTISVal As Variant
                        If wsColIdx <= wsMaxCol Then oldCellVal = wsDataArr(wsArrRow, wsColIdx) Else oldCellVal = ""
                        newTISVal = tisData(tisRow, tisColIdx)

                        Dim valChanged As Boolean
                        valChanged = False
                        If IsDate(oldCellVal) And IsDate(newTISVal) Then
                            valChanged = (CLng(CDate(oldCellVal)) <> CLng(CDate(newTISVal)))
                        Else
                            valChanged = (LCase(Trim(CStr(oldCellVal))) <> LCase(Trim(CStr(newTISVal))))
                        End If

                        ' Always update TIS column to latest value (collect for write)
                        updateWriteCells.Add Array(r, wsColIdx, newTISVal, IsDate(newTISVal))

                        If valChanged Then
                            changedCount = changedCount + 1
                            ' Collect for batch orange fill + comment
                            updateChangedCells.Add Array(r, wsColIdx, CStr(oldCellVal))
                        End If
                    End If
                End If
            Next hdr

            ' Reintroduction: if Status was Cancelled, set back to Active
            If wsStatusCol > 0 And wsStatusCol <= wsMaxCol Then
                If LCase(Trim(CStr(wsDataArr(wsArrRow, wsStatusCol)))) = "cancelled" Then
                    statusWrites.Add Array(r, wsStatusCol, "Active")
                    reintroducedCount = reintroducedCount + 1
                End If
            End If

            ' Remove from TIS dict (processed)
            tisKeyMap.Remove wsKey
        Else
            ' --- Project NOT in TIS: cancel if active ---
            If wsStatusCol > 0 And wsStatusCol <= wsMaxCol Then
                Dim currentStatus As String
                currentStatus = LCase(Trim(CStr(wsDataArr(wsArrRow, wsStatusCol))))
                If currentStatus = "active" Then
                    statusWrites.Add Array(r, wsStatusCol, "Cancelled")
                    cancelledCount = cancelledCount + 1
                End If
            End If
        End If
NextWSRow:
    Next r

    ' --- BATCH WRITE: TIS column updates ---
    ' Write all collected TIS value updates to the sheet
    stepDesc = "Batch write TIS updates"
    Application.StatusBar = "Updating Working Sheet... writing updates"
    Dim uwi As Long, uwInfo As Variant
    For uwi = 1 To updateWriteCells.Count
        uwInfo = updateWriteCells(uwi)
        ws.Cells(uwInfo(0), uwInfo(1)).Value = uwInfo(2)
        If uwInfo(3) Then ws.Cells(uwInfo(0), uwInfo(1)).NumberFormat = "mm/dd/yyyy"
    Next uwi

    ' --- BATCH WRITE: Status updates ---
    For uwi = 1 To statusWrites.Count
        uwInfo = statusWrites(uwi)
        ws.Cells(uwInfo(0), uwInfo(1)).Value = uwInfo(2)
    Next uwi

    ' --- BATCH APPLY: Orange fill for changed cells using Union ---
    If updateChangedCells.Count > 0 Then
        Dim uOrangeRange As Range
        Dim uci As Long, uCellInfo As Variant
        For uci = 1 To updateChangedCells.Count
            uCellInfo = updateChangedCells(uci)
            If uOrangeRange Is Nothing Then
                Set uOrangeRange = ws.Cells(uCellInfo(0), uCellInfo(1))
            Else
                Set uOrangeRange = Union(uOrangeRange, ws.Cells(uCellInfo(0), uCellInfo(1)))
            End If
        Next uci
        If Not uOrangeRange Is Nothing Then
            uOrangeRange.Interior.Color = CLR_CHANGE_FILL
        End If

        ' Write comments (per-cell, acceptable with ScreenUpdating=False)
        For uci = 1 To updateChangedCells.Count
            uCellInfo = updateChangedCells(uci)
            Dim commentText As String
            commentText = "[" & Format(Date, "YYYY-MM-DD") & "] Changed from: " & CStr(uCellInfo(2))
            On Error Resume Next
            Dim uCell As Range
            Set uCell = ws.Cells(uCellInfo(0), uCellInfo(1))
            If uCell.Comment Is Nothing Then
                uCell.AddComment commentText
            Else
                Dim uExisting As String
                uExisting = uCell.Comment.Text
                If Len(uExisting) + Len(commentText) + 1 < 1024 Then
                    uCell.Comment.Text uExisting & vbLf & commentText
                End If
            End If
            On Error GoTo UpdateError
        Next uci
    End If

    ' --- APPEND NEW SYSTEMS from remaining TIS dict ---
    stepDesc = "Append new systems"
    Application.StatusBar = "Updating Working Sheet... adding new systems"
    Dim newCount As Long
    newCount = 0

    ' Get the ListObject table to append rows properly
    Dim tbl As ListObject
    Set tbl = Nothing
    If ws.ListObjects.Count > 0 Then Set tbl = ws.ListObjects(1)

    Dim newTISKey As Variant
    For Each newTISKey In tisKeyMap.Keys
        Dim newTISRow As Long
        newTISRow = tisKeyMap(newTISKey)

        ' Add a new row to the table (or append below data)
        Dim appendRow As Long
        If Not tbl Is Nothing Then
            Dim newListRow As ListRow
            Set newListRow = tbl.ListRows.Add
            appendRow = newListRow.Range.Row
        Else
            wsLastRow = wsLastRow + 1
            appendRow = wsLastRow
        End If

        ' Write TIS data to matching columns
        For Each hdr In sirfisHeaders
            hdrLower = LCase(hdr.Value)
            wsMapKey = hdrLower
            If wsHeaderMap.exists(wsMapKey) And headerDict.exists(hdr.Value) Then
                wsColIdx = wsHeaderMap(wsMapKey)
                tisColIdx = headerDict(hdr.Value)
                ws.Cells(appendRow, wsColIdx).Value = tisData(newTISRow, tisColIdx)
                If IsDate(tisData(newTISRow, tisColIdx)) Then
                    ws.Cells(appendRow, wsColIdx).NumberFormat = "mm/dd/yyyy"
                End If
            End If
        Next hdr

        ' Set Status = "Active"
        If wsStatusCol > 0 Then ws.Cells(appendRow, wsStatusCol).Value = "Active"

        ' Populate Group formula (VLOOKUP to CEIDs sheet)
        Dim wsGroupCol As Long
        wsGroupCol = 0
        If wsHeaderMap.exists("group") Then wsGroupCol = wsHeaderMap("group")
        If wsGroupCol > 0 Then
            Dim wsEntityTypeCol As Long
            wsEntityTypeCol = 0
            If wsHeaderMap.exists("entity type") Then wsEntityTypeCol = wsHeaderMap("entity type")
            If wsEntityTypeCol > 0 And SheetExists(ThisWorkbook, SHEET_CEIDS) Then
                ws.Cells(appendRow, wsGroupCol).Formula = _
                    "=IFERROR(VLOOKUP(" & ColLetter(wsEntityTypeCol) & appendRow & ",CEIDs!A:B,2,FALSE),"""")"
            End If
        End If

        ' Auto-populate Our Dates from TIS dates (blue border marker)
        Dim ourToTIS(0 To 7, 0 To 1) As String
        ourToTIS(0, 0) = LCase(TIS_COL_OUR_SET):   ourToTIS(0, 1) = LCase(TIS_SRC_SET)
        ourToTIS(1, 0) = LCase(TIS_COL_OUR_SL1):   ourToTIS(1, 1) = LCase(TIS_SRC_SL1)
        ourToTIS(2, 0) = LCase(TIS_COL_OUR_SL2):   ourToTIS(2, 1) = LCase(TIS_SRC_SL2)
        ourToTIS(3, 0) = LCase(TIS_COL_OUR_SQ):    ourToTIS(3, 1) = LCase(TIS_SRC_SQ)
        ourToTIS(4, 0) = LCase(TIS_COL_OUR_CONVS): ourToTIS(4, 1) = LCase(TIS_SRC_CONVS)
        ourToTIS(5, 0) = LCase(TIS_COL_OUR_CONVF): ourToTIS(5, 1) = LCase(TIS_SRC_CONVF)
        ourToTIS(6, 0) = LCase(TIS_COL_OUR_MRCLS): ourToTIS(6, 1) = LCase(TIS_SRC_MRCLS)
        ourToTIS(7, 0) = LCase(TIS_COL_OUR_MRCLF): ourToTIS(7, 1) = LCase(TIS_SRC_MRCLF)

        Dim mi As Long, ourColPos As Long, tisColPos As Long
        For mi = 0 To 7
            ourColPos = 0: tisColPos = 0
            If wsHeaderMap.exists(ourToTIS(mi, 0)) Then ourColPos = wsHeaderMap(ourToTIS(mi, 0))
            If wsHeaderMap.exists(ourToTIS(mi, 1)) Then tisColPos = wsHeaderMap(ourToTIS(mi, 1))
            If ourColPos > 0 And tisColPos > 0 Then
                Dim newTISCellVal As Variant
                newTISCellVal = ws.Cells(appendRow, tisColPos).Value
                If IsDate(newTISCellVal) Then
                    ws.Cells(appendRow, ourColPos).Value = newTISCellVal
                    ws.Cells(appendRow, ourColPos).NumberFormat = "mm/dd/yyyy"
                    With ws.Cells(appendRow, ourColPos).Borders
                        .LineStyle = xlContinuous: .Weight = xlThin: .Color = CLR_NEW_DATE_BORDER
                    End With
                End If
            End If
        Next mi

        ' Mark Entity Code with blue border (new system indicator)
        If wsECCol > 0 Then
            ApplyNewProjectBorder ws.Cells(appendRow, wsECCol)
        End If

        ' Populate New/Reused formula for the new row
        Dim wsNRCol As Long, wsEvtCol As Long
        wsNRCol = 0: wsEvtCol = 0
        Dim nrKey As String
        nrKey = "new/" & vbLf & "reused"
        If wsHeaderMap.exists(nrKey) Then wsNRCol = wsHeaderMap(nrKey)
        If wsNRCol = 0 And wsHeaderMap.exists("new/reused") Then wsNRCol = wsHeaderMap("new/reused")
        If wsHeaderMap.exists(LCase(HEADER_EVENT_TYPE)) Then wsEvtCol = wsHeaderMap(LCase(HEADER_EVENT_TYPE))
        If wsNRCol > 0 And wsEvtCol > 0 And wsECCol > 0 Then
            Dim evtLetter As String, ecLetter As String
            evtLetter = ColLetter(wsEvtCol)
            ecLetter = ColLetter(wsECCol)
            If SheetExists(ThisWorkbook, SHEET_NEW_REUSED) Then
                ws.Cells(appendRow, wsNRCol).Formula = _
                    "=IF(LOWER(" & evtLetter & appendRow & ")=""demo"",""Demo""," & _
                    "IF(IFERROR(VLOOKUP(" & ecLetter & appendRow & ",'New-reused'!$D:$G,4,FALSE),"""")=""New"",""New"",""Reused""))"
            Else
                ws.Cells(appendRow, wsNRCol).Formula = _
                    "=IF(LOWER(" & evtLetter & appendRow & ")=""demo"",""Demo"",""Reused"")"
            End If
        End If

        ' Populate Lock? validation for the new row
        Dim wsLockCol As Long
        wsLockCol = 0
        If wsHeaderMap.exists(LCase(TIS_COL_LOCK)) Then wsLockCol = wsHeaderMap(LCase(TIS_COL_LOCK))
        If wsLockCol > 0 Then
            On Error Resume Next
            ws.Cells(appendRow, wsLockCol).Validation.Delete
            ws.Cells(appendRow, wsLockCol).Validation.Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, Formula1:="TRUE,FALSE"
            ws.Cells(appendRow, wsLockCol).Validation.IgnoreBlank = True
            ws.Cells(appendRow, wsLockCol).Validation.InCellDropdown = True
            On Error GoTo UpdateError
        End If

        ' Populate Status validation for the new row
        If wsStatusCol > 0 Then
            On Error Resume Next
            ws.Cells(appendRow, wsStatusCol).Validation.Delete
            ws.Cells(appendRow, wsStatusCol).Validation.Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, Formula1:="Active,Completed,On Hold,Non IQ,Cancelled"
            ws.Cells(appendRow, wsStatusCol).Validation.IgnoreBlank = True
            ws.Cells(appendRow, wsStatusCol).Validation.InCellDropdown = True
            On Error GoTo UpdateError
        End If

        newCount = newCount + 1
    Next newTISKey

    ' --- REBUILD GANTT + NIF ---
    stepDesc = "Rebuild Gantt"
    Application.StatusBar = "Updating Working Sheet... rebuilding Gantt"
    On Error Resume Next
    GanttBuilder.BuildGantt silent:=True, targetSheet:=ws
    If Err.Number <> 0 Then
        DebugLog "WARNING: GanttBuilder failed during update: " & Err.Description
        Err.Clear
    End If
    On Error GoTo UpdateError

    stepDesc = "Rebuild NIF"
    Application.StatusBar = "Updating Working Sheet... rebuilding NIF"
    On Error Resume Next
    NIF_Builder.BuildNIF silent:=True, targetSheet:=ws
    If Err.Number <> 0 Then
        DebugLog "WARNING: NIF_Builder failed during update: " & Err.Description
        Err.Clear
    End If
    On Error GoTo UpdateError

    ' --- Recompute Health column ---
    stepDesc = "Recompute Health"
    Application.StatusBar = "Updating Working Sheet... computing Health"
    Dim wsDataRowCount As Long
    wsDataRowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).row - wsHdr
    If wsDataRowCount > 0 Then
        PopulateHealthColumn ws, wsHeaderMap, wsHdr, wsDataRowCount
    End If

    ' --- Summary ---
    Application.StatusBar = False
    Dim summary As String
    summary = "TIS Update complete:" & vbCrLf & vbCrLf
    If changedCount > 0 Then summary = summary & "  " & changedCount & " TIS field(s) changed (orange)" & vbCrLf
    If newCount > 0 Then summary = summary & "  " & newCount & " new system(s) added" & vbCrLf
    If cancelledCount > 0 Then summary = summary & "  " & cancelledCount & " system(s) cancelled (removed from TIS)" & vbCrLf
    If reintroducedCount > 0 Then summary = summary & "  " & reintroducedCount & " system(s) reintroduced" & vbCrLf
    If changedCount = 0 And newCount = 0 And cancelledCount = 0 And reintroducedCount = 0 Then
        summary = summary & "  No changes detected."
    End If
    MsgBox summary, vbInformation, "TIS Update"

    Exit Sub

UpdateError:
    Application.StatusBar = False
    MsgBox "Error in UpdateWorkingSheet at step: " & stepDesc & vbCrLf & vbCrLf & _
           Err.Description & " (#" & Err.Number & ")", vbCritical
    DebugLog "UpdateWorkingSheet ERROR at " & stepDesc & ": " & Err.Description
End Sub

'====================================================================
' CREATE WORKING SHEET - MAIN ORCHESTRATOR
'====================================================================

Private Sub CreateWorkingSheet(hasCEIDsSheet As Boolean, hasMilestonesSheet As Boolean, _
                                ByRef errOccurred As Boolean, _
                                Optional ByVal existingSheet As Worksheet = Nothing, _
                                Optional ByVal doImport As Boolean = False)
    On Error GoTo WSErrorHandler
    
    Dim stepDesc As String
    stepDesc = "Init"
    Application.StatusBar = "Building Working Sheet...  5% - Initializing"
    
    ' Debug: check slicer state at the very start
    Dim dbgSCInit As SlicerCache
    DebugLog "=== CreateWorkingSheet START: slicer caches ==="
    For Each dbgSCInit In ThisWorkbook.SlicerCaches
        DebugLog "  cache='" & dbgSCInit.Name & "' slicers=" & dbgSCInit.Slicers.Count
    Next dbgSCInit
    DebugLog "=== END slicer check ==="
    
    Dim tisSheet As Worksheet, wsDef As Worksheet, newSheet As Worksheet
    Dim tisData As Variant, outputData() As Variant
    Dim headerDict As Object, filterDict As Object, sortDict As Object
    Dim pastDateFilterDict As Object, gatingDict As Object
    Dim excludeSystemsList As Collection
    Dim siteColIdx As Long
    Dim sirfisHeaders As Range, header As Range
    Dim lastCol As Long, lastRow As Long, colIndex As Long
    Dim filterValue As String, sortOrder As Variant, gatingValue As String
    Dim missingHeaders As String, missingSortHeaders As String
    Dim outputRow As Long, outputCol As Long
    Dim i As Long, j As Long
    Dim entityTypeColIdx As Long, ceidColIdx As Long, eventTypeColIdx As Long
    Dim entityCodeColIdx As Long
    Dim groupColPos As Long, colCounter As Long
    Dim milestoneGroups As Object, milestoneNames As Object, durationExcludes As Object
    Dim baseColCount As Long, totalColCount As Long, milestoneCount As Long
    Dim filteredRowCount As Long, dataRowCount As Long
    Dim actualStartCol As Long, stdStartCol As Long, additionalStartCol As Long
    Dim stdEndCol As Long
    Dim outputColMap As Object
    Dim newReusedColPos As Long, estCycleColPos As Long
    Dim entityTypeOutputCol As Long
    Dim stdSectionEndCol As Long, gapSectionStartCol As Long, gapSectionEndCol As Long
    Dim oldSheet As Worksheet

    stepDesc = "Load TIS sheet"
    Application.StatusBar = "Building Working Sheet... 10% - Loading TIS data"
    Set tisSheet = ThisWorkbook.Sheets(SHEET_TIS)
    Set wsDef = ThisWorkbook.Sheets(SHEET_DEFINITIONS)
    
    ' Get sirfisheaders
    If wsDef.Range("A2").Value = "" Then
        MsgBox "The 'sirfisheaders' range is empty.", vbExclamation
        Exit Sub
    End If
    Set sirfisHeaders = wsDef.Range("A2", wsDef.Cells(wsDef.Rows.Count, "A").End(xlUp))
    
    ' Initialize dictionaries
    Set headerDict = CreateObject("Scripting.Dictionary")
    Set filterDict = CreateObject("Scripting.Dictionary")   ' key=headerName, value=Array of allowed values
    Set sortDict = CreateObject("Scripting.Dictionary")
    Set pastDateFilterDict = CreateObject("Scripting.Dictionary")
    Set gatingDict = CreateObject("Scripting.Dictionary")
    
    Dim groupedDict As Object   ' key=headerName, value=group number
    Set groupedDict = CreateObject("Scripting.Dictionary")
    
    missingHeaders = ""
    missingSortHeaders = ""
    
    ' Load TIS data into array
    lastCol = tisSheet.Cells(1, tisSheet.Columns.Count).End(xlToLeft).Column
    lastRow = tisSheet.Cells(tisSheet.Rows.Count, 1).End(xlUp).row
    tisData = tisSheet.Range(tisSheet.Cells(1, 1), tisSheet.Cells(lastRow, lastCol)).Value
    
    ' Match headers and build dictionaries
    Dim groupedValue As String
    For Each header In sirfisHeaders
        filterValue = Trim(CStr(header.Offset(0, 1).Value))
        sortOrder = Trim(header.Offset(0, 2).Value)
        gatingValue = Trim(header.Offset(0, 3).Value)
        groupedValue = Trim(CStr(header.Offset(0, 9).Value))  ' Column J = Grouped
        
        For colIndex = 1 To lastCol
            If LCase(header.Value) = LCase(tisData(1, colIndex)) Then
                headerDict(header.Value) = colIndex
                
                ' Store filter - support space-separated OR values (TISLoader parity)
                If filterValue <> "" Then
                    filterDict(header.Value) = Split(filterValue, " ")
                End If
                
                ' Sort/date filter
                If UCase(sortOrder) = "X" Then
                    If IsDateHeader(header.Value) Then
                        pastDateFilterDict(header.Value) = True
                    End If
                ElseIf IsNumeric(sortOrder) And sortOrder > 0 Then
                    If IsDateHeader(header.Value) Then
                        If Not sortDict.exists(sortOrder) Then
                            sortDict(sortOrder) = header.Value
                        Else
                            sortDict(sortOrder) = sortDict(sortOrder) & "|" & header.Value
                        End If
                    End If
                End If
                
                ' Gating
                If gatingValue <> "" Then gatingDict(header.Value) = gatingValue
                
                ' Grouped columns (for column grouping/collapsing)
                If groupedValue <> "" And IsNumeric(groupedValue) Then
                    groupedDict(header.Value) = CLng(groupedValue)
                End If
                
                ' Track key columns
                If LCase(header.Value) = "entity type" Then entityTypeColIdx = colIndex
                If LCase(header.Value) = "ceid" Then ceidColIdx = colIndex
                If LCase(header.Value) = LCase(HEADER_EVENT_TYPE) Then eventTypeColIdx = colIndex
                If LCase(header.Value) = LCase(HEADER_ENTITY_CODE) Then entityCodeColIdx = colIndex
                If LCase(header.Value) = "site" Then siteColIdx = colIndex
                
                Exit For
            End If
        Next colIndex
        
        If Not headerDict.exists(header.Value) Then
            missingHeaders = missingHeaders & header.Value & vbNewLine
            If IsNumeric(sortOrder) And sortOrder > 0 Then
                missingSortHeaders = missingSortHeaders & header.Value & " (Sort Priority: " & sortOrder & ")" & vbNewLine
            End If
        End If
    Next header
    
    ' Display missing headers
    If missingHeaders <> "" Then
        Dim msg As String
        msg = "The following headers were not found in the TIS sheet:" & vbNewLine & missingHeaders
        If missingSortHeaders <> "" Then
            msg = msg & vbNewLine & "Sort will not be applied for:" & vbNewLine & missingSortHeaders
        End If
        MsgBox msg, vbExclamation
        Exit Sub
    End If
    
    If headerDict.Count = 0 Then
        MsgBox "No matching headers found. Exiting.", vbExclamation
        Exit Sub
    End If
    
    ' Validate required headers (using lowercase dictionary)
    Dim lcHeaderDict As Object
    Set lcHeaderDict = CreateObject("Scripting.Dictionary")
    Dim hKey As Variant
    For Each hKey In headerDict.Keys
        lcHeaderDict(LCase(CStr(hKey))) = headerDict(hKey)
    Next hKey
    
    Dim missingRequiredHeaders As String
    missingRequiredHeaders = ""
    If Not lcHeaderDict.exists(LCase(HEADER_ENTITY_CODE)) Then missingRequiredHeaders = missingRequiredHeaders & "  - " & HEADER_ENTITY_CODE & vbCrLf
    If Not lcHeaderDict.exists(LCase(HEADER_SDD)) Then missingRequiredHeaders = missingRequiredHeaders & "  - " & HEADER_SDD & vbCrLf
    If Not lcHeaderDict.exists(LCase(HEADER_SUPPLIER_QUAL_FINISH)) Then missingRequiredHeaders = missingRequiredHeaders & "  - " & HEADER_SUPPLIER_QUAL_FINISH & vbCrLf
    If Not lcHeaderDict.exists(LCase(HEADER_SET_START)) Then missingRequiredHeaders = missingRequiredHeaders & "  - " & HEADER_SET_START & vbCrLf
    
    If missingRequiredHeaders <> "" Then
        If MsgBox("The following required headers for additional columns were not found:" & vbCrLf & vbCrLf & _
                  missingRequiredHeaders & vbCrLf & _
                  "Some calculated columns may not work correctly." & vbCrLf & vbCrLf & _
                  "Continue anyway?", vbYesNo + vbExclamation, "Missing Required Headers") = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Parse milestone definitions
    Set milestoneGroups = CreateObject("Scripting.Dictionary")
    Set milestoneNames = CreateObject("Scripting.Dictionary")
    Set durationExcludes = CreateObject("Scripting.Dictionary")
    ParseMilestoneDefinitions wsDef, milestoneGroups, milestoneNames, durationExcludes
    
    ' Parse exclude systems from column E
    Set excludeSystemsList = ParseExcludeSystems(wsDef)
    
    ' Determine Group column position (after CEID)
    groupColPos = 0
    colCounter = 0
    For Each header In sirfisHeaders
        If headerDict.exists(header.Value) Then
            colCounter = colCounter + 1
            If LCase(header.Value) = "ceid" Then
                groupColPos = colCounter + 1
                Exit For
            End If
        End If
    Next header
    If groupColPos = 0 Then groupColPos = headerDict.Count + 1
    
    ' Calculate column counts
    baseColCount = headerDict.Count + 1  ' +1 for Group
    milestoneCount = milestoneGroups.Count - durationExcludes.Count
    ' Count filtered rows (using OR filter logic)
    filteredRowCount = 0
    For i = 2 To lastRow
        If RowPassesFilters(tisData, i, headerDict, filterDict, pastDateFilterDict, _
                            excludeSystemsList, siteColIdx, entityCodeColIdx, eventTypeColIdx) Then
            filteredRowCount = filteredRowCount + 1
        End If
    Next i
    
    If filteredRowCount = 0 Then
        MsgBox "No rows match the current filters.", vbInformation
        Exit Sub
    End If
    
    ' Prepare output array
    ReDim outputData(1 To filteredRowCount + 1, 1 To baseColCount)
    
    ' Build header row with Group insertion
    outputCol = 1
    Dim outHdrVal As String
    For Each header In sirfisHeaders
        If headerDict.exists(header.Value) Then
            If outputCol = groupColPos Then
                outputData(1, outputCol) = "Group"
                outputCol = outputCol + 1
            End If
            outHdrVal = header.Value
            outputData(1, outputCol) = outHdrVal
            outputCol = outputCol + 1
        End If
    Next header
    If outputCol = groupColPos Then outputData(1, outputCol) = "Group"
    
    ' Build output column map (properly handling Group offset)
    Set outputColMap = CreateObject("Scripting.Dictionary")
    outputCol = 1
    Dim mapKey As String
    For Each header In sirfisHeaders
        If headerDict.exists(header.Value) Then
            If outputCol = groupColPos Then
                ' Group column occupies this position, push data column to next
                outputCol = outputCol + 1
            End If
            mapKey = LCase(header.Value)
            outputColMap(mapKey) = outputCol
            outputCol = outputCol + 1
        End If
    Next header
    ' Register Group column in the map (needed for slicer)
    If groupColPos > 0 Then outputColMap("group") = groupColPos
    
    ' Get Entity Type output column for dashboard positioning
    If outputColMap.exists("entity type") Then
        entityTypeOutputCol = outputColMap("entity type")
    Else
        entityTypeOutputCol = 1
    End If
    
    ' Filter and copy rows
    outputRow = 2
    For i = 2 To lastRow
        If RowPassesFilters(tisData, i, headerDict, filterDict, pastDateFilterDict, _
                            excludeSystemsList, siteColIdx, entityCodeColIdx, eventTypeColIdx) Then
            outputCol = 1
            For Each header In sirfisHeaders
                If headerDict.exists(header.Value) Then
                    If outputCol = groupColPos Then
                        ' Group column left empty - formula inserted after data write
                        outputData(outputRow, outputCol) = ""
                        outputCol = outputCol + 1
                    End If
                    colIndex = headerDict(header.Value)
                    outputData(outputRow, outputCol) = tisData(i, colIndex)
                    outputCol = outputCol + 1
                End If
            Next header
            If outputCol = groupColPos Then
                ' Group column left empty - formula inserted after data write
                outputData(outputRow, outputCol) = ""
            End If
            outputRow = outputRow + 1
        End If
    Next i
    
    ' Create or reuse sheet
    stepDesc = "Create/reuse sheet"
    Application.StatusBar = "Building Working Sheet... 30% - Creating sheet"

    If Not existingSheet Is Nothing Then
        ' === IN-PLACE REBUILD ===
        Set oldSheet = Nothing

        ' Create backup copy first
        BackupWorkingSheet existingSheet

        ' The backup is now the ActiveSheet (from ws.Copy)
        If doImport Then
            Set oldSheet = ActiveSheet  ' The backup copy has the data
        End If

        ' Now clear the original
        ClearWorkingSheet existingSheet

        ' Reuse the original sheet object
        Set newSheet = existingSheet
    Else
        ' === FIRST-TIME CREATION ===
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = SHEET_WORKING
    End If

    ' Light theme: set default font
    newSheet.Cells.Font.Name = "Calibri"

    ' Write base data
    dataRowCount = outputRow - 1
    newSheet.Range(newSheet.Cells(DATA_START_ROW, 1), _
                   newSheet.Cells(DATA_START_ROW + dataRowCount - 1, baseColCount)).Value = outputData
    
    ' Insert Group column formula (VLOOKUP to CEIDs sheet)
    If hasCEIDsSheet And groupColPos > 0 And outputColMap.exists("entity type") Then
        InsertGroupColumnFormula newSheet, groupColPos, outputColMap("entity type"), _
                                  DATA_START_ROW, dataRowCount
    End If
    
    ' Apply sorting
    stepDesc = "ApplySorting"
    Application.StatusBar = "Building Working Sheet... 35% - Sorting data"
    If sortDict.Count > 0 Then
        ApplySortingWithHelperColumns newSheet, sortDict, outputColMap, DATA_START_ROW
    End If
    
    ' Rev14: Add Our Dates block after base data, before milestones
    stepDesc = "AddOurDatesBlock"
    Application.StatusBar = "Building Working Sheet... 38% - Adding committed date columns"
    Dim ourBlockEndCol As Long
    AddOurDatesBlock newSheet, outputColMap, DATA_START_ROW, dataRowCount, baseColCount + 1, ourBlockEndCol

    ' Add milestone columns (start after Our Dates block)
    stepDesc = "Add milestone columns"
    Application.StatusBar = "Building Working Sheet... 40% - Adding milestones"
    stdStartCol = 0
    stdEndCol = 0
    actualStartCol = ourBlockEndCol
    
    If milestoneCount > 0 Then
        AddActualDurationColumns newSheet, milestoneGroups, milestoneNames, outputColMap, _
                                  DATA_START_ROW, dataRowCount, actualStartCol, durationExcludes
        
        stdStartCol = actualStartCol + milestoneCount + 1
        
        If hasMilestonesSheet Then
            AddSTDDurationColumns newSheet, milestoneGroups, milestoneNames, outputColMap, _
                                   DATA_START_ROW, dataRowCount, stdStartCol, actualStartCol, stdEndCol, _
                                   stdSectionEndCol, gapSectionStartCol, gapSectionEndCol
            additionalStartCol = stdEndCol
        Else
            additionalStartCol = actualStartCol + milestoneCount + 1
        End If
    Else
        additionalStartCol = baseColCount + 1
    End If
    
    ' Add additional columns (returns positions of key columns)
    stepDesc = "AddAdditionalColumns"
    Application.StatusBar = "Building Working Sheet... 50% - Adding columns"
    AddAdditionalColumns newSheet, outputColMap, DATA_START_ROW, dataRowCount, additionalStartCol, _
                          newReusedColPos, estCycleColPos, milestoneGroups
    
    ' Calculate final totalColCount from actual last used column
    totalColCount = newSheet.Cells(DATA_START_ROW, newSheet.Columns.Count).End(xlToLeft).Column
    
    ' Read cycle time threshold from Definitions!S1 (no longer written to worksheet)
    Dim ctThreshold As Long
    ctThreshold = DEFAULT_CYCLE_TIME_THRESHOLD
    On Error Resume Next
    If IsNumeric(wsDef.Range("S1").Value) And wsDef.Range("S1").Value > 0 Then
        ctThreshold = CLng(wsDef.Range("S1").Value)
    End If
    On Error GoTo WSErrorHandler
    
    ' Apply cycle time CF using S1 reference
    If estCycleColPos > 0 Then
        ApplyCycleTimeConditionalFormatting newSheet, estCycleColPos, DATA_START_ROW, dataRowCount
    End If
    
    ' Apply date formatting
    ApplyDateFormatting newSheet, sirfisHeaders, outputColMap, DATA_START_ROW, dataRowCount
    
    ' Apply gating CF
    stepDesc = "ApplyGatingCF"
    Application.StatusBar = "Building Working Sheet... 55% - Conditional formatting"
    If gatingDict.Count > 0 Then
        ApplyGatingConditionalFormatting newSheet, gatingDict, sirfisHeaders, outputColMap, DATA_START_ROW
    End If
    
    ' Create table FIRST (with no style) then apply custom formatting
    stepDesc = "Create ListObject table"
    Application.StatusBar = "Building Working Sheet... 60% - Creating table"
    Dim tbl As ListObject
    Dim tableRange As Range
    Set tableRange = newSheet.Range(newSheet.Cells(DATA_START_ROW, 1), _
                                     newSheet.Cells(DATA_START_ROW + dataRowCount - 1, totalColCount))
    Set tbl = newSheet.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    
    ' Debug: check if creating a new ListObject killed old slicers
    Dim dbgSCTbl As SlicerCache
    DebugLog "=== AFTER ListObject.Add: slicer caches ==="
    For Each dbgSCTbl In ThisWorkbook.SlicerCaches
        DebugLog "  cache='" & dbgSCTbl.Name & "' slicers=" & dbgSCTbl.Slicers.Count
    Next dbgSCTbl
    DebugLog "=== END post-table check ==="
    
    ' Generate unique table name (avoid conflict with existing tables)
    Dim candidateTableName As String
    candidateTableName = Replace(SHEET_WORKING, " ", "_")
    On Error Resume Next
    tbl.name = candidateTableName
    If Err.Number <> 0 Then
        Err.Clear
        candidateTableName = candidateTableName & "_" & Format(Now, "hhnnss")
        tbl.name = candidateTableName
    End If
    On Error GoTo WSErrorHandler
    
    tbl.TableStyle = ""  ' No style - preserve our custom formatting
    
    ' Apply column grouping from Definitions sheet (collapsed by default)
    stepDesc = "ApplyColumnGrouping"
    Application.StatusBar = "Building Working Sheet... 65% - Grouping columns"
    ApplyColumnGrouping newSheet, groupedDict, outputColMap
    
    ' Group STD duration columns (collapsed by default)
    If stdStartCol > 0 And stdEndCol > stdStartCol Then
        GroupColumnsCollapsed newSheet, stdStartCol, stdEndCol - 1
    End If
    
    ' Apply ultra-polished formatting AFTER table creation
    stepDesc = "ApplyWorkingSheetFormatting"
    Application.StatusBar = "Building Working Sheet... 70% - Formatting"
    ApplyWorkingSheetFormatting newSheet, DATA_START_ROW, dataRowCount, totalColCount, _
                                 actualStartCol, stdStartCol, milestoneCount, additionalStartCol, _
                                 stdSectionEndCol, gapSectionStartCol, gapSectionEndCol, outputColMap
    
    ' Center all calculated number columns
    CenterCalculatedColumns newSheet, DATA_START_ROW, dataRowCount, actualStartCol, totalColCount, baseColCount
    
    ' Add summary dashboard (rows 2-7 with time frame, active row, new counters)
    Dim sddColPos As Long, setStartColPos As Long
    Dim sqFinishColPos As Long, escalatedColPos As Long, watchColPos As Long
    sddColPos = 0: setStartColPos = 0: sqFinishColPos = 0
    escalatedColPos = 0: watchColPos = 0
    If outputColMap.exists("sdd") Then sddColPos = outputColMap("sdd")
    If outputColMap.exists(LCase(HEADER_SET_START)) Then setStartColPos = outputColMap(LCase(HEADER_SET_START))
    If outputColMap.exists(LCase(HEADER_SUPPLIER_QUAL_FINISH)) Then sqFinishColPos = outputColMap(LCase(HEADER_SUPPLIER_QUAL_FINISH))
    
    ' Find escalated column (holds both "Escalated" and "Watched" values)
    Dim ecol As Long
    For ecol = 1 To totalColCount
        Dim hdrVal As String
        hdrVal = LCase(Trim(Replace(Replace(CStr(newSheet.Cells(DATA_START_ROW, ecol).Value), vbLf, ""), vbCr, "")))
        If hdrVal = "escalated" Then escalatedColPos = ecol
    Next ecol
    ' watchColPos = same column (Watched is a value in the Escalated column)
    watchColPos = escalatedColPos
    
    stepDesc = "AddSummaryDashboard"
    Application.StatusBar = "Building Working Sheet... 75% - Building dashboard"
    AddSummaryDashboard newSheet, DATA_START_ROW, dataRowCount, totalColCount, _
                         newReusedColPos, estCycleColPos, entityTypeOutputCol, tbl.Name, _
                         sddColPos, setStartColPos, sqFinishColPos, ctThreshold, _
                         escalatedColPos, watchColPos, outputColMap, milestoneGroups
    
    ' Add Excel native Slicers (Group + New/Reused)
    ' Slicers will be added at the very end after Gantt and NIF
    Dim lastCounterCol As Long
    lastCounterCol = entityTypeOutputCol + 9
    If entityTypeOutputCol < 5 Then lastCounterCol = 5 + 9
    
    ' Set title
    newSheet.Cells(1, 1).Value = "Working Sheet"
    newSheet.Cells(1, 1).Font.Bold = True
    newSheet.Cells(1, 1).Font.Size = 16
    newSheet.Cells(1, 1).Font.Color = CLR_DARK_TEXT
    
    ' Timestamp
    newSheet.Cells(1, 4).Value = Format(Now, "mm/dd/yyyy hh:mm")
    newSheet.Cells(1, 4).Font.Size = 9
    newSheet.Cells(1, 4).Font.Color = CLR_TIMESTAMP
    newSheet.Cells(1, 4).HorizontalAlignment = xlLeft
    
    ' === IMPORT SYSTEM: Compare with old sheet, import user data only ===
    stepDesc = "ImportUserDataFromOldSheet"
    Application.StatusBar = "Building Working Sheet... 85% - Importing data"
    If doImport And Not oldSheet Is Nothing Then
        ImportUserDataFromOldSheet oldSheet, newSheet, outputColMap, sirfisHeaders, headerDict, _
                            DATA_START_ROW, dataRowCount, totalColCount
        ' Resize table to include any appended rows (removed systems carried forward)
        Dim actualLastRow As Long
        actualLastRow = newSheet.Cells(newSheet.Rows.Count, 1).End(xlUp).Row
        If actualLastRow > DATA_START_ROW + dataRowCount - 1 Then
            Dim newDataRowCount As Long
            newDataRowCount = actualLastRow - DATA_START_ROW
            On Error Resume Next
            tbl.Resize newSheet.Range(newSheet.Cells(DATA_START_ROW, 1), _
                                       newSheet.Cells(actualLastRow, totalColCount))
            On Error GoTo WSErrorHandler
            dataRowCount = newDataRowCount + 1  ' Update for downstream steps
        End If
    End If

    ' Rev14: Auto-populate Our Dates for new systems from TIS dates
    stepDesc = "PopulateNewSystemOurDates"
    Application.StatusBar = "Building Working Sheet... 86% - Populating Our Dates"
    Dim rev14SiteCol As Long, rev14EntityCol As Long, rev14EventCol As Long
    rev14SiteCol = 0: rev14EntityCol = 0: rev14EventCol = 0
    If outputColMap.exists("site") Then rev14SiteCol = outputColMap("site")
    If outputColMap.exists(LCase(HEADER_ENTITY_CODE)) Then rev14EntityCol = outputColMap(LCase(HEADER_ENTITY_CODE))
    If outputColMap.exists(LCase(HEADER_EVENT_TYPE)) Then rev14EventCol = outputColMap(LCase(HEADER_EVENT_TYPE))

    ' Check if old sheet had Our Date columns. If not, pass an empty oldKeyMap
    ' so ALL rows are treated as "new" and get TIS dates copied into Our Dates.
    Dim oldHasOurDates As Boolean
    oldHasOurDates = False
    If doImport And Not oldSheet Is Nothing Then
        Dim chkHdr As Long
        chkHdr = TISCommon.FindHeaderRow(oldSheet)
        If chkHdr > 0 Then
            ' Look for "Conv.S" — unique to Our Dates (never in raw TIS)
            If TISCommon.FindHeaderCol(oldSheet, chkHdr, TIS_COL_OUR_CONVS, _
                    oldSheet.Cells(chkHdr, oldSheet.Columns.Count).End(xlToLeft).Column) > 0 Then
                oldHasOurDates = True
            End If
        End If
    End If

    Dim rev14OldKeyMap As Object
    Set rev14OldKeyMap = CreateObject("Scripting.Dictionary")
    If doImport And Not oldSheet Is Nothing And oldHasOurDates Then
        Dim rev14OldHdr As Long
        rev14OldHdr = TISCommon.FindHeaderRow(oldSheet)
        If rev14OldHdr > 0 Then
            Dim rev14OldLastRow As Long, rev14c As Long, rev14r As Long
            Dim rev14hv As String, rev14pk As String
            rev14OldLastRow = oldSheet.Cells(oldSheet.Rows.Count, 1).End(xlUp).row
            Dim rev14OldSiteC As Long, rev14OldECC As Long, rev14OldETC As Long
            rev14OldSiteC = 0: rev14OldECC = 0: rev14OldETC = 0
            For rev14c = 1 To oldSheet.Cells(rev14OldHdr, oldSheet.Columns.Count).End(xlToLeft).Column
                rev14hv = LCase(Trim(Replace(Replace(CStr(oldSheet.Cells(rev14OldHdr, rev14c).Value), vbLf, ""), vbCr, "")))
                If rev14hv = "site" Then rev14OldSiteC = rev14c
                If rev14hv = LCase(HEADER_ENTITY_CODE) Then rev14OldECC = rev14c
                If rev14hv = LCase(HEADER_EVENT_TYPE) Then rev14OldETC = rev14c
            Next rev14c
            For rev14r = rev14OldHdr + 1 To rev14OldLastRow
                rev14pk = TISCommon.BuildProjectKey(oldSheet, rev14r, rev14OldSiteC, rev14OldECC, rev14OldETC)
                If rev14pk <> "||" And Not rev14OldKeyMap.exists(rev14pk) Then
                    rev14OldKeyMap(rev14pk) = rev14r
                End If
            Next rev14r
        End If
    End If
    PopulateNewSystemOurDates newSheet, outputColMap, DATA_START_ROW, dataRowCount, _
                               rev14OldKeyMap, rev14SiteCol, rev14EntityCol, rev14EventCol

    ' Rev14: Compute Health column
    stepDesc = "PopulateHealthColumn"
    Application.StatusBar = "Building Working Sheet... 87% - Computing Health"
    PopulateHealthColumn newSheet, outputColMap, DATA_START_ROW, dataRowCount

    ' Rev14: Sort by Status (Active first, Cancelled last) then project start date
    stepDesc = "StatusSort"
    Application.StatusBar = "Building Working Sheet... 88% - Sorting by status and start date"
    SortWorkingSheetByStatus newSheet, outputColMap, DATA_START_ROW

    ' Auto-fit visible data columns only (not the entire sheet — expensive at 100+ cols)
    stepDesc = "AutoFit and FreezePanes"
    Application.StatusBar = "Building Working Sheet... 90% - Final layout"
    Dim autoFitLastCol As Long
    autoFitLastCol = newSheet.Cells(DATA_START_ROW, newSheet.Columns.Count).End(xlToLeft).Column
    If autoFitLastCol > 0 Then
        newSheet.Range(newSheet.Columns(1), newSheet.Columns(autoFitLastCol)).AutoFit
    End If
    Dim freezeCol As Long
    freezeCol = 1
    If outputColMap.exists(LCase(HEADER_ENTITY_CODE)) Then
        freezeCol = outputColMap(LCase(HEADER_ENTITY_CODE)) + 1
    End If
    newSheet.Activate
    ActiveWindow.FreezePanes = False
    newSheet.Cells(DATA_START_ROW + 1, freezeCol).Select
    ActiveWindow.FreezePanes = True
    newSheet.Cells(1, 1).Select

    ' Note: full recalc deferred to Cleanup block (Application.Calculate) — no per-sheet Calculate needed
    
    ' Install event handler for slicer/filter responsiveness
    stepDesc = "InstallSheetEvents"
    InstallSheetEvents newSheet
    
    ' === CALL GANTT BUILDER (fresh generation on new sheet, silent mode) ===
    ' Pass targetSheet so Gantt builds on the correct (new) sheet
    stepDesc = "Call GanttBuilder"
    Application.StatusBar = "Building Working Sheet... 93% - Building Gantt chart"
    On Error Resume Next
    GanttBuilder.BuildGantt silent:=True, targetSheet:=newSheet
    On Error GoTo WSErrorHandler
    
    ' === CALL NIF BUILDER (with source sheet for data migration) ===
    ' When old sheet exists, NIF_Builder reads saved data from oldSheet
    ' and restores it onto newSheet (cross-sheet migration)
    stepDesc = "Call NIF_Builder"
    Application.StatusBar = "Building Working Sheet... 97% - Building NIF"
    On Error GoTo WSErrorHandler
    If doImport And Not oldSheet Is Nothing Then
        NIF_Builder.BuildNIF silent:=True, targetSheet:=newSheet, sourceSheet:=oldSheet
    Else
        NIF_Builder.BuildNIF silent:=True, targetSheet:=newSheet
    End If
    
    ' === ADD SLICERS (last step — sheet must be fully built and active) ===
    stepDesc = "AddSlicers"
    Application.StatusBar = "Building Working Sheet... 99% - Adding slicers"
    
    ' Debug: check slicer state before AddSlicers
    Dim dbgSCPre As SlicerCache
    DebugLog "=== PRE-AddSlicers: slicer caches ==="
    For Each dbgSCPre In ThisWorkbook.SlicerCaches
        DebugLog "  cache='" & dbgSCPre.Name & "' slicers=" & dbgSCPre.Slicers.Count
    Next dbgSCPre
    DebugLog "=== END pre-slicer check ==="
    Application.ScreenUpdating = True
    newSheet.Activate
    DoEvents
    AddSlicers newSheet, tbl, outputColMap, lastCounterCol
    Application.ScreenUpdating = False
    
    Exit Sub
    
WSErrorHandler:
    errOccurred = True
    MsgBox "Error in CreateWorkingSheet at step: " & stepDesc & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & _
           "Error number: " & Err.Number, vbCritical
    DebugLog "WorkfileBuilder ERROR at " & stepDesc & ": " & Err.Description
    ' Only delete if this was a newly created sheet (not a reuse)
    If Not newSheet Is Nothing And existingSheet Is Nothing Then
        Application.DisplayAlerts = False
        newSheet.Delete
        Application.DisplayAlerts = True
    End If
End Sub

'====================================================================
' PARSE MILESTONE DEFINITIONS
'====================================================================

Private Sub ParseMilestoneDefinitions(wsDef As Worksheet, ByRef milestoneGroups As Object, _
                                       ByRef milestoneNames As Object, _
                                       Optional ByRef durationExcludes As Object = Nothing)
    Dim defLastRow As Long, i As Long
    Dim defData As Variant
    Dim fText As String, tokens As Variant, token As Variant
    Dim letter As String, num As Long
    Dim lastCol As Long

    defLastRow = wsDef.Cells(wsDef.Rows.Count, 1).End(xlUp).row
    If defLastRow < 2 Then Exit Sub

    ' Read up to column K (11) to pick up Actual Duration exclude flag
    lastCol = 11
    If wsDef.Cells(1, 11).Value = "" Then lastCol = 7
    defData = wsDef.Range(wsDef.Cells(1, 1), wsDef.Cells(defLastRow, lastCol)).Value

    ' Initialize excludes dictionary if caller wants it
    If durationExcludes Is Nothing Then Set durationExcludes = CreateObject("Scripting.Dictionary")

    For i = 2 To UBound(defData, 1)
        fText = Trim(CStr(defData(i, 6)))
        If fText <> "" Then
            tokens = Split(fText, "|")
            For Each token In tokens
                token = UCase(Trim(token))
                If Len(token) >= 2 And IsNumeric(Mid(token, 2)) Then
                    letter = Left(token, 1)
                    num = CLng(Mid(token, 2))
                    If Not milestoneGroups.exists(letter) Then
                        Set milestoneGroups(letter) = CreateObject("Scripting.Dictionary")
                    End If
                    milestoneGroups(letter)(num) = Array(token, i, CStr(defData(i, 1)))
                    If num = 1 And Trim(CStr(defData(i, 7))) <> "" Then
                        milestoneNames(letter) = Trim(CStr(defData(i, 7)))
                    End If
                    ' Column K (11) = Actual Duration exclude
                    ' Only apply to num=1 (start token) — Column G names this group
                    If num = 1 And lastCol >= 11 Then
                        If UCase(Trim(CStr(defData(i, 11)))) = "X" Then
                            durationExcludes(letter) = True
                        End If
                    End If
                End If
            Next token
        End If
    Next i
End Sub

'====================================================================
' PARSE EXCLUDE SYSTEMS FROM COLUMN E
' Format: Site|Entity Code|Event Type (e.g., F24|CSB117|Qual)
'====================================================================

Private Function ParseExcludeSystems(wsDef As Worksheet) As Collection
    Dim excludeList As Collection
    Dim defLastRow As Long, i As Long
    Dim cellValue As String
    Dim parts() As String
    
    Set excludeList = New Collection
    
    defLastRow = wsDef.Cells(wsDef.Rows.Count, "A").End(xlUp).row
    If defLastRow < 2 Then
        Set ParseExcludeSystems = excludeList
        Exit Function
    End If
    
    ' Column E = "Exclude systems"
    For i = 2 To defLastRow
        cellValue = Trim(CStr(wsDef.Cells(i, 5).Value))  ' Column E = 5
        If cellValue <> "" Then
            parts = Split(cellValue, "|")
            If UBound(parts) >= 2 Then
                ' Store as array: (0)=site, (1)=entityCode, (2)=eventType (all lowercase)
                excludeList.Add Array(LCase(Trim(parts(0))), _
                                       LCase(Trim(parts(1))), _
                                       LCase(Trim(parts(2))))
            End If
        End If
    Next i
    
    Set ParseExcludeSystems = excludeList
End Function

'====================================================================
' ADD OUR DATES BLOCK — 9 committed milestone date columns
' Inserted after base data, before milestone columns.
' New systems get auto-populated from TIS dates (blue border marker).
'====================================================================

Private Sub AddOurDatesBlock(ws As Worksheet, outputColMap As Object, _
                              dataStartRow As Long, dataRowCount As Long, _
                              startCol As Long, ByRef endCol As Long)
    Dim currentCol As Long, lastDataRow As Long
    Dim oi As Long

    currentCol = startCol
    lastDataRow = dataStartRow + dataRowCount - 1

    Dim ourDateCols(0 To 7) As String
    ourDateCols(0) = TIS_COL_OUR_SET
    ourDateCols(1) = TIS_COL_OUR_SL1
    ourDateCols(2) = TIS_COL_OUR_SL2
    ourDateCols(3) = TIS_COL_OUR_SQ
    ourDateCols(4) = TIS_COL_OUR_CONVS
    ourDateCols(5) = TIS_COL_OUR_CONVF
    ourDateCols(6) = TIS_COL_OUR_MRCLS
    ourDateCols(7) = TIS_COL_OUR_MRCLF

    For oi = 0 To 7
        ws.Cells(dataStartRow, currentCol).Value = ourDateCols(oi)
        ws.Columns(currentCol).ColumnWidth = 11
        ' Apply THEME_ACCENT background to header cell
        ws.Cells(dataStartRow, currentCol).Interior.Color = THEME_ACCENT
        ws.Cells(dataStartRow, currentCol).Font.Color = THEME_WHITE
        ws.Cells(dataStartRow, currentCol).Font.Bold = True
        If lastDataRow > dataStartRow Then
            ws.Range(ws.Cells(dataStartRow + 1, currentCol), _
                     ws.Cells(lastDataRow, currentCol)).NumberFormat = "mm/dd/yyyy"
            ws.Range(ws.Cells(dataStartRow + 1, currentCol), _
                     ws.Cells(lastDataRow, currentCol)).HorizontalAlignment = xlCenter
        End If
        outputColMap(LCase(ourDateCols(oi))) = currentCol
        currentCol = currentCol + 1
    Next oi

    ' === Lock-aware Data Validation on Our Date columns ===
    ' When Lock?=TRUE for a row, validation rejects any edit to Our Date cells.
    ' Zero runtime cost — Excel evaluates only when user presses Enter on a cell.
    ' Lock? column will be at currentCol+1 (after Status at currentCol).
    ' We defer this until after Lock? column is created — see below.

    ' === Status column ===
    ws.Cells(dataStartRow, currentCol).Value = TIS_COL_STATUS
    ws.Columns(currentCol).ColumnWidth = 10
    ws.Cells(dataStartRow, currentCol).Interior.Color = THEME_ACCENT
    ws.Cells(dataStartRow, currentCol).Font.Color = THEME_WHITE
    ws.Cells(dataStartRow, currentCol).Font.Bold = True
    If lastDataRow > dataStartRow Then
        Dim statusRange As Range
        Set statusRange = ws.Range(ws.Cells(dataStartRow + 1, currentCol), ws.Cells(lastDataRow, currentCol))
        statusRange.Value = "Active"
        On Error Resume Next
        statusRange.Validation.Delete
        statusRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="Active,Completed,On Hold,Non IQ,Cancelled"
        statusRange.Validation.IgnoreBlank = True
        statusRange.Validation.InCellDropdown = True
        On Error GoTo 0
        statusRange.HorizontalAlignment = xlCenter
    End If
    outputColMap(LCase(TIS_COL_STATUS)) = currentCol
    currentCol = currentCol + 1

    ' === Lock? column ===
    ws.Cells(dataStartRow, currentCol).Value = TIS_COL_LOCK
    ws.Columns(currentCol).ColumnWidth = 7
    ws.Cells(dataStartRow, currentCol).Interior.Color = THEME_ACCENT
    ws.Cells(dataStartRow, currentCol).Font.Color = THEME_WHITE
    ws.Cells(dataStartRow, currentCol).Font.Bold = True
    If lastDataRow > dataStartRow Then
        Dim lockRange As Range
        Set lockRange = ws.Range(ws.Cells(dataStartRow + 1, currentCol), ws.Cells(lastDataRow, currentCol))
        On Error Resume Next
        lockRange.Validation.Delete
        lockRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="TRUE,FALSE"
        lockRange.Validation.IgnoreBlank = True
        lockRange.Validation.InCellDropdown = True
        On Error GoTo 0
        lockRange.HorizontalAlignment = xlCenter
    End If
    outputColMap(LCase(TIS_COL_LOCK)) = currentCol
    Dim lockColPos As Long
    lockColPos = currentCol
    currentCol = currentCol + 1

    ' === Apply Lock-aware Data Validation to Our Date columns ===
    ' Custom formula: =NOT($LockCol_Row=TRUE)
    ' When Lock?=TRUE, validation rejects edits. When FALSE/empty, edits allowed.
    ' Uses absolute column ref ($) + relative row ref so it works for every row.
    If lastDataRow > dataStartRow Then
        Dim lockColLetter As String
        lockColLetter = ColLetter(lockColPos)
        Dim dvFormula As String
        ' Formula uses first data row reference — Excel adjusts per row automatically
        dvFormula = "=NOT($" & lockColLetter & CStr(dataStartRow + 1) & "=TRUE)"

        Dim dvi As Long
        For dvi = 0 To 7
            Dim dvCol As Long
            dvCol = outputColMap(LCase(ourDateCols(dvi)))
            If dvCol > 0 Then
                Dim dvRange As Range
                Set dvRange = ws.Range(ws.Cells(dataStartRow + 1, dvCol), ws.Cells(lastDataRow, dvCol))
                On Error Resume Next
                dvRange.Validation.Delete
                dvRange.Validation.Add Type:=xlValidateCustom, _
                    AlertStyle:=xlValidAlertStop, _
                    Formula1:=dvFormula
                dvRange.Validation.ErrorTitle = "Row Locked"
                dvRange.Validation.ErrorMessage = "This row's dates are locked. Set Lock? to FALSE to edit."
                dvRange.Validation.ShowError = True
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
            End If
        Next dvi

        ' === Visual indicator: gray out locked Our Date cells via CF ===
        ' When Lock?=TRUE, Our Date cells get a light gray background
        Dim lockCFFormula As String
        lockCFFormula = "=$" & lockColLetter & CStr(dataStartRow + 1) & "=TRUE"
        Dim ourFirstCol As Long, ourLastCol As Long
        ourFirstCol = outputColMap(LCase(ourDateCols(0)))
        ourLastCol = outputColMap(LCase(ourDateCols(7)))
        Dim lockCFRange As Range
        Set lockCFRange = ws.Range(ws.Cells(dataStartRow + 1, ourFirstCol), ws.Cells(lastDataRow, ourLastCol))
        On Error Resume Next
        Dim lockFC As FormatCondition
        Set lockFC = lockCFRange.FormatConditions.Add(Type:=xlExpression, Formula1:=lockCFFormula)
        If Not lockFC Is Nothing Then
            lockFC.Interior.Color = SLATE_200  ' Light gray background
            lockFC.Font.Color = SLATE_500      ' Muted text
            lockFC.StopIfTrue = False
        End If
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If

    ' === Health column (auto-calculated) ===
    ws.Cells(dataStartRow, currentCol).Value = TIS_COL_HEALTH
    ws.Columns(currentCol).ColumnWidth = 10
    ws.Cells(dataStartRow, currentCol).Interior.Color = THEME_ACCENT
    ws.Cells(dataStartRow, currentCol).Font.Color = THEME_WHITE
    ws.Cells(dataStartRow, currentCol).Font.Bold = True
    If lastDataRow > dataStartRow Then
        ws.Range(ws.Cells(dataStartRow + 1, currentCol), ws.Cells(lastDataRow, currentCol)).HorizontalAlignment = xlCenter
    End If
    outputColMap(LCase(TIS_COL_HEALTH)) = currentCol
    currentCol = currentCol + 1

    ' === WhatIf column (user-facing) ===
    ws.Cells(dataStartRow, currentCol).Value = TIS_COL_WHATIF
    ws.Columns(currentCol).ColumnWidth = 11
    ws.Cells(dataStartRow, currentCol).Interior.Color = THEME_ACCENT
    ws.Cells(dataStartRow, currentCol).Font.Color = THEME_WHITE
    ws.Cells(dataStartRow, currentCol).Font.Bold = True
    If lastDataRow > dataStartRow Then
        ws.Range(ws.Cells(dataStartRow + 1, currentCol), ws.Cells(lastDataRow, currentCol)).NumberFormat = "mm/dd/yyyy"
        ws.Range(ws.Cells(dataStartRow + 1, currentCol), ws.Cells(lastDataRow, currentCol)).HorizontalAlignment = xlCenter
    End If
    outputColMap(LCase(TIS_COL_WHATIF)) = currentCol
    currentCol = currentCol + 1

    endCol = currentCol
End Sub

'====================================================================
' POPULATE NEW SYSTEM OUR DATES
' For rows new to TIS or with all Our Dates empty: copy TIS date values
' into Our Date columns and mark with blue border.
'====================================================================

Private Sub PopulateNewSystemOurDates(ws As Worksheet, outputColMap As Object, _
                                       dataStartRow As Long, dataRowCount As Long, _
                                       oldKeyMap As Object, _
                                       newSiteCol As Long, newEntityCodeCol As Long, newEventTypeCol As Long)
    ' Map: Our Date key -> TIS source key in outputColMap
    Dim ourToTIS(0 To 7, 0 To 1) As String
    ourToTIS(0, 0) = LCase(TIS_COL_OUR_SET):   ourToTIS(0, 1) = LCase(TIS_SRC_SET)
    ourToTIS(1, 0) = LCase(TIS_COL_OUR_SL1):   ourToTIS(1, 1) = LCase(TIS_SRC_SL1)
    ourToTIS(2, 0) = LCase(TIS_COL_OUR_SL2):   ourToTIS(2, 1) = LCase(TIS_SRC_SL2)
    ourToTIS(3, 0) = LCase(TIS_COL_OUR_SQ):    ourToTIS(3, 1) = LCase(TIS_SRC_SQ)
    ourToTIS(4, 0) = LCase(TIS_COL_OUR_CONVS): ourToTIS(4, 1) = LCase(TIS_SRC_CONVS)
    ourToTIS(5, 0) = LCase(TIS_COL_OUR_CONVF): ourToTIS(5, 1) = LCase(TIS_SRC_CONVF)
    ourToTIS(6, 0) = LCase(TIS_COL_OUR_MRCLS): ourToTIS(6, 1) = LCase(TIS_SRC_MRCLS)
    ourToTIS(7, 0) = LCase(TIS_COL_OUR_MRCLF): ourToTIS(7, 1) = LCase(TIS_SRC_MRCLF)

    Dim lastDataRow As Long
    lastDataRow = dataStartRow + dataRowCount - 1
    If lastDataRow <= dataStartRow Then Exit Sub

    Dim actualDataRows As Long
    actualDataRows = lastDataRow - dataStartRow  ' number of data rows (excl header)

    ' Resolve column positions for Our Dates and TIS sources
    Dim ourCols(0 To 7) As Long, tisCols(0 To 7) As Long
    Dim mi As Long
    For mi = 0 To 7
        ourCols(mi) = 0: tisCols(mi) = 0
        If outputColMap.exists(ourToTIS(mi, 0)) Then ourCols(mi) = outputColMap(ourToTIS(mi, 0))
        If outputColMap.exists(ourToTIS(mi, 1)) Then tisCols(mi) = outputColMap(ourToTIS(mi, 1))
    Next mi

    ' Rev14: Lock? column position
    Dim lockColPos As Long
    lockColPos = 0
    If outputColMap.exists(LCase(TIS_COL_LOCK)) Then lockColPos = outputColMap(LCase(TIS_COL_LOCK))

    ' Performance: Bulk-read key columns into arrays for key building
    Dim siteArr As Variant, ecArr As Variant, etArr As Variant
    If newSiteCol > 0 Then siteArr = ws.Range(ws.Cells(dataStartRow + 1, newSiteCol), ws.Cells(lastDataRow, newSiteCol)).Value
    If newEntityCodeCol > 0 Then ecArr = ws.Range(ws.Cells(dataStartRow + 1, newEntityCodeCol), ws.Cells(lastDataRow, newEntityCodeCol)).Value
    If newEventTypeCol > 0 Then etArr = ws.Range(ws.Cells(dataStartRow + 1, newEventTypeCol), ws.Cells(lastDataRow, newEventTypeCol)).Value

    ' Bulk-read Lock? column
    Dim lockArr As Variant
    If lockColPos > 0 Then lockArr = ws.Range(ws.Cells(dataStartRow + 1, lockColPos), ws.Cells(lastDataRow, lockColPos)).Value

    ' Bulk-read TIS source columns into arrays (8 columns)
    Dim tisArrays(0 To 7) As Variant
    For mi = 0 To 7
        If tisCols(mi) > 0 Then
            tisArrays(mi) = ws.Range(ws.Cells(dataStartRow + 1, tisCols(mi)), ws.Cells(lastDataRow, tisCols(mi))).Value
        End If
    Next mi

    ' Bulk-read existing Our Date columns into arrays (to check if empty)
    Dim ourArrays(0 To 7) As Variant
    For mi = 0 To 7
        If ourCols(mi) > 0 Then
            ourArrays(mi) = ws.Range(ws.Cells(dataStartRow + 1, ourCols(mi)), ws.Cells(lastDataRow, ourCols(mi))).Value
        End If
    Next mi

    ' Build output arrays for each Our Date column (start with existing values)
    Dim tmpArr() As Variant
    Dim ourOutArrays(0 To 7) As Variant
    For mi = 0 To 7
        If ourCols(mi) > 0 Then
            ReDim tmpArr(1 To actualDataRows, 1 To 1)
            Dim ri As Long
            For ri = 1 To actualDataRows
                If Not IsEmpty(ourArrays(mi)) Then tmpArr(ri, 1) = ourArrays(mi)(ri, 1) Else tmpArr(ri, 1) = Empty
            Next ri
            ourOutArrays(mi) = tmpArr
        End If
    Next mi

    ' Track which cells need blue border (row index in data, milestone index)
    Dim blueBorderCells As New Collection  ' Each item = Array(sheetRow, colPos)

    ' Process each row in memory
    Dim i As Long, pk As String
    For i = 1 To actualDataRows
        Dim pS As String, pEC As String, pET As String
        pS = "": pEC = "": pET = ""
        If newSiteCol > 0 Then pS = LCase(Trim(CStr(siteArr(i, 1))))
        If newEntityCodeCol > 0 Then pEC = LCase(Trim(CStr(ecArr(i, 1))))
        If newEventTypeCol > 0 Then pET = LCase(Trim(CStr(etArr(i, 1))))
        pk = pS & "|" & pEC & "|" & pET
        If pk = "||" Then GoTo NextPopRow

        ' Skip locked rows
        If lockColPos > 0 Then
            If LCase(Trim(CStr(lockArr(i, 1)))) = "true" Then GoTo NextPopRow
        End If

        ' Only populate if: new system (not in old sheet) OR all Our Dates empty
        Dim isNewOrEmpty As Boolean
        isNewOrEmpty = False
        If Not oldKeyMap.exists(pk) Then
            isNewOrEmpty = True
        Else
            ' Check if all 8 Our Dates are empty (in-memory check)
            isNewOrEmpty = True
            For mi = 0 To 7
                If ourCols(mi) > 0 And Not IsEmpty(ourArrays(mi)) Then
                    If Not IsEmpty(ourArrays(mi)(i, 1)) Then
                        If ourArrays(mi)(i, 1) <> "" Then
                            isNewOrEmpty = False
                            Exit For
                        End If
                    End If
                End If
            Next mi
        End If

        If isNewOrEmpty Then
            Dim sheetRow As Long
            sheetRow = dataStartRow + i  ' actual sheet row
            For mi = 0 To 7
                If ourCols(mi) > 0 And tisCols(mi) > 0 Then
                    Dim tisCellVal As Variant
                    If Not IsEmpty(tisArrays(mi)) Then tisCellVal = tisArrays(mi)(i, 1) Else tisCellVal = Empty
                    If IsDate(tisCellVal) Then
                        ' Write to output array
                        Dim oArr As Variant
                        oArr = ourOutArrays(mi)
                        oArr(i, 1) = tisCellVal
                        ourOutArrays(mi) = oArr
                        ' Track for blue border
                        blueBorderCells.Add Array(sheetRow, ourCols(mi))
                    End If
                End If
            Next mi
        End If
NextPopRow:
    Next i

    ' Bulk-write each Our Date column back to the sheet
    For mi = 0 To 7
        If ourCols(mi) > 0 Then
            ws.Range(ws.Cells(dataStartRow + 1, ourCols(mi)), ws.Cells(lastDataRow, ourCols(mi))).Value = ourOutArrays(mi)
            ' Set number format for the entire Our Date column at once
            ws.Range(ws.Cells(dataStartRow + 1, ourCols(mi)), ws.Cells(lastDataRow, ourCols(mi))).NumberFormat = "mm/dd/yyyy"
        End If
    Next mi

    ' Apply blue borders in batch using Union
    If blueBorderCells.Count > 0 Then
        Dim blueRange As Range
        Dim bi As Long
        For bi = 1 To blueBorderCells.Count
            Dim bInfo As Variant
            bInfo = blueBorderCells(bi)
            If blueRange Is Nothing Then
                Set blueRange = ws.Cells(bInfo(0), bInfo(1))
            Else
                Set blueRange = Union(blueRange, ws.Cells(bInfo(0), bInfo(1)))
            End If
        Next bi
        If Not blueRange Is Nothing Then
            With blueRange.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = CLR_NEW_DATE_BORDER
            End With
        End If
    End If
End Sub

'====================================================================
' ALL OUR DATES EMPTY - checks if all 8 Our Date cells are blank
' (Retained for backward compat; new callers use in-memory array check)
'====================================================================

Private Function AllOurDatesEmpty(ws As Worksheet, row As Long, outputColMap As Object) As Boolean
    AllOurDatesEmpty = True
    Dim ourHeaders As Variant
    ourHeaders = Array(LCase(TIS_COL_OUR_SET), LCase(TIS_COL_OUR_SL1), _
                       LCase(TIS_COL_OUR_SL2), LCase(TIS_COL_OUR_SQ), LCase(TIS_COL_OUR_CONVS), _
                       LCase(TIS_COL_OUR_CONVF), LCase(TIS_COL_OUR_MRCLS), LCase(TIS_COL_OUR_MRCLF))
    Dim hi As Long
    For hi = LBound(ourHeaders) To UBound(ourHeaders)
        If outputColMap.exists(ourHeaders(hi)) Then
            If Not IsEmpty(ws.Cells(row, outputColMap(ourHeaders(hi))).Value) Then
                If ws.Cells(row, outputColMap(ourHeaders(hi))).Value <> "" Then
                    AllOurDatesEmpty = False
                    Exit Function
                End If
            End If
        End If
    Next hi
End Function

'====================================================================
' POPULATE HEALTH COLUMN
' Compares Our Dates vs TIS Dates for each milestone.
' Deviation = Our Date - TIS Date (days). Positive = our date later.
' Health = max deviation across all milestones:
'   <= 0 -> "On Track", 1-7 -> "At Risk", > 7 -> "Behind"
' Demo rows and rows with no dates are left blank.
'====================================================================

Private Sub PopulateHealthColumn(ws As Worksheet, outputColMap As Object, _
                                  dataStartRow As Long, dataRowCount As Long)
    ' Health = live formula: Match / Minor / Gap based on max deviation (Our Date - TIS Date)
    ' Match (green): all deviations <= 0 (aligned or ahead of TIS)
    ' Minor (amber): max deviation 1-3 days (small drift)
    ' Gap (red): max deviation > 3 days (needs attention)
    ' Formula auto-recalculates when user edits an Our Date — no rebuild needed.

    Dim healthCol As Long
    healthCol = 0
    If outputColMap.exists(LCase(TIS_COL_HEALTH)) Then healthCol = outputColMap(LCase(TIS_COL_HEALTH))
    If healthCol = 0 Then Exit Sub

    Dim lastDataRow As Long
    lastDataRow = dataStartRow + dataRowCount - 1
    If lastDataRow <= dataStartRow Then Exit Sub

    ' Build MAX formula from Our Date vs TIS Date pairs
    ' MAX(IF(AND(ISNUMBER(our),ISNUMBER(tis)), our-tis, 0), ...)
    Dim ourKeys(0 To 7) As String, tisKeys(0 To 7) As String
    ourKeys(0) = LCase(TIS_COL_OUR_SET):   tisKeys(0) = LCase(TIS_SRC_SET)
    ourKeys(1) = LCase(TIS_COL_OUR_SL1):   tisKeys(1) = LCase(TIS_SRC_SL1)
    ourKeys(2) = LCase(TIS_COL_OUR_SL2):   tisKeys(2) = LCase(TIS_SRC_SL2)
    ourKeys(3) = LCase(TIS_COL_OUR_SQ):    tisKeys(3) = LCase(TIS_SRC_SQ)
    ourKeys(4) = LCase(TIS_COL_OUR_CONVS): tisKeys(4) = LCase(TIS_SRC_CONVS)
    ourKeys(5) = LCase(TIS_COL_OUR_CONVF): tisKeys(5) = LCase(TIS_SRC_CONVF)
    ourKeys(6) = LCase(TIS_COL_OUR_MRCLS): tisKeys(6) = LCase(TIS_SRC_MRCLS)
    ourKeys(7) = LCase(TIS_COL_OUR_MRCLF): tisKeys(7) = LCase(TIS_SRC_MRCLF)

    ' Build the MAX(...) part with each pair as IF(AND(ISNUMBER(our),ISNUMBER(tis)),our-tis,0)
    Dim firstRow As Long
    firstRow = dataStartRow + 1
    Dim maxParts As String
    maxParts = ""
    Dim pi As Long
    Dim ourC As Long, tisC As Long
    Dim ourRef As String, tisRef As String
    For pi = 0 To 7
        ourC = 0: tisC = 0
        If outputColMap.exists(ourKeys(pi)) Then ourC = outputColMap(ourKeys(pi))
        If outputColMap.exists(tisKeys(pi)) Then tisC = outputColMap(tisKeys(pi))
        If ourC > 0 And tisC > 0 Then
            ourRef = ColLetter(ourC) & firstRow
            tisRef = ColLetter(tisC) & firstRow
            If maxParts <> "" Then maxParts = maxParts & ","
            maxParts = maxParts & "IF(AND(ISNUMBER(" & ourRef & "),ISNUMBER(" & tisRef & "))," & ourRef & "-" & tisRef & ",0)"
        End If
    Next pi

    If maxParts = "" Then Exit Sub

    ' Count how many pairs have data (to detect "no dates to compare")
    Dim countParts As String
    countParts = ""
    For pi = 0 To 7
        ourC = 0: tisC = 0
        If outputColMap.exists(ourKeys(pi)) Then ourC = outputColMap(ourKeys(pi))
        If outputColMap.exists(tisKeys(pi)) Then tisC = outputColMap(tisKeys(pi))
        If ourC > 0 And tisC > 0 Then
            ourRef = ColLetter(ourC) & firstRow
            tisRef = ColLetter(tisC) & firstRow
            If countParts <> "" Then countParts = countParts & "+"
            countParts = countParts & "IF(AND(ISNUMBER(" & ourRef & "),ISNUMBER(" & tisRef & ")),1,0)"
        End If
    Next pi

    ' Final formula: =IF(countParts=0, "", IF(MAX(...)>3, "Gap", IF(MAX(...)>0, "Minor", "Match")))
    Dim maxExpr As String
    maxExpr = "MAX(" & maxParts & ")"
    Dim healthFormula As String
    healthFormula = "=IF(" & countParts & "=0,""""," & _
                    "IF(" & maxExpr & ">3,""Gap""," & _
                    "IF(" & maxExpr & ">0,""Minor"",""Match"")))"

    ' Write formula to first data row, then FillDown for performance
    On Error Resume Next
    ws.Cells(firstRow, healthCol).Formula = healthFormula
    If Err.Number <> 0 Then
        Err.Clear
        ws.Cells(firstRow, healthCol).Formula2 = healthFormula
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
    End If
    On Error GoTo 0

    If lastDataRow > firstRow Then
        ws.Cells(firstRow, healthCol).Copy
        ws.Range(ws.Cells(firstRow + 1, healthCol), ws.Cells(lastDataRow, healthCol)).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False
    End If

    ' Apply CF: Match=green, Minor=amber, Gap=red
    Dim healthRng As Range
    Set healthRng = ws.Range(ws.Cells(firstRow, healthCol), ws.Cells(lastDataRow, healthCol))
    Dim healthAddr As String
    healthAddr = ws.Cells(firstRow, healthCol).Address(False, False)

    On Error Resume Next
    healthRng.FormatConditions.Delete
    If Err.Number <> 0 Then Err.Clear

    Dim hfc As FormatCondition
    Set hfc = healthRng.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & healthAddr & "=""Gap""")
    If Not hfc Is Nothing Then
        hfc.Interior.Color = STATUS_BEHIND_BG: hfc.Font.Color = STATUS_BEHIND_FG: hfc.Font.Bold = True: hfc.StopIfTrue = True
    End If

    Set hfc = healthRng.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & healthAddr & "=""Minor""")
    If Not hfc Is Nothing Then
        hfc.Interior.Color = STATUS_ATRISK_BG: hfc.Font.Color = STATUS_ATRISK_FG: hfc.Font.Bold = True: hfc.StopIfTrue = True
    End If

    Set hfc = healthRng.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & healthAddr & "=""Match""")
    If Not hfc Is Nothing Then
        hfc.Interior.Color = STATUS_ONTRACK_BG: hfc.Font.Color = STATUS_ONTRACK_FG: hfc.Font.Bold = True: hfc.StopIfTrue = True
    End If
    On Error GoTo 0
End Sub

'====================================================================
' ADD ACTUAL DURATION COLUMNS (FillDown for performance)
'====================================================================

Private Sub AddActualDurationColumns(ws As Worksheet, milestoneGroups As Object, _
                                      milestoneNames As Object, outputColMap As Object, _
                                      dataStartRow As Long, dataRowCount As Long, startCol As Long, _
                                      Optional durationExcludes As Object = Nothing)
    Dim k As Variant, currentCol As Long, displayName As String
    Dim startColLetter As String, endColLetter As String
    Dim startHeaderName As String, endHeaderName As String
    Dim startColIdx As Long, endColIdx As Long
    Dim r As Long, formula As String
    Dim group As Object
    Dim firstMilCol As Long, lastMilCol As Long
    Dim milestoneKeys As Variant, idx As Long
    Dim lastDataRow As Long

    currentCol = startCol
    firstMilCol = startCol
    lastDataRow = dataStartRow + dataRowCount - 1

    milestoneKeys = milestoneGroups.Keys
    Call ShellSortVariantArray(milestoneKeys)

    For idx = LBound(milestoneKeys) To UBound(milestoneKeys)
        k = milestoneKeys(idx)

        ' Skip milestones marked as excluded in Definitions column K
        If Not durationExcludes Is Nothing Then
            If durationExcludes.exists(CStr(k)) Then GoTo NextMilestone
        End If

        Set group = milestoneGroups(k)
        
        displayName = IIf(milestoneNames.exists(k), milestoneNames(k), k)
        ws.Cells(dataStartRow, currentCol).Value = "Actual" & vbLf & "Duration" & vbLf & displayName
        
        If group.exists(1) And group.exists(2) Then
            startHeaderName = group(1)(2)
            endHeaderName = group(2)(2)
            startColIdx = 0: endColIdx = 0
            If outputColMap.exists(LCase(startHeaderName)) Then startColIdx = outputColMap(LCase(startHeaderName))
            If outputColMap.exists(LCase(endHeaderName)) Then endColIdx = outputColMap(LCase(endHeaderName))
            
            If startColIdx > 0 And endColIdx > 0 Then
                startColLetter = ColLetter(startColIdx)
                endColLetter = ColLetter(endColIdx)
                
                ' Write formula to first cell, then FillDown
                r = dataStartRow + 1
                formula = "=IF(AND(ISNUMBER(" & startColLetter & r & "),ISNUMBER(" & endColLetter & r & "))," & _
                          endColLetter & r & "-" & startColLetter & r & "+1,"""")"
                ws.Cells(r, currentCol).formula = formula
                If dataRowCount > 2 Then
                    ws.Range(ws.Cells(r, currentCol), ws.Cells(lastDataRow, currentCol)).FillDown
                End If
            End If
        End If
        
        ws.Columns(currentCol).NumberFormat = "0"
        currentCol = currentCol + 1
NextMilestone:
    Next idx

    lastMilCol = currentCol - 1
    
    ' Total Actual Duration
    ws.Cells(dataStartRow, currentCol).Value = "Total" & vbLf & "Actual" & vbLf & "Duration"
    If lastMilCol >= firstMilCol Then
        r = dataStartRow + 1
        ws.Cells(r, currentCol).formula = "=SUM(" & ColLetter(firstMilCol) & r & ":" & ColLetter(lastMilCol) & r & ")"
        If dataRowCount > 2 Then
            ws.Range(ws.Cells(r, currentCol), ws.Cells(lastDataRow, currentCol)).FillDown
        End If
    End If
    ws.Columns(currentCol).NumberFormat = "0"
End Sub

'====================================================================
' ADD STD DURATION COLUMNS WITH VLOOKUP + FUZZY MATCHING FALLBACK
'====================================================================

Private Sub AddSTDDurationColumns(ws As Worksheet, milestoneGroups As Object, _
                                   milestoneNames As Object, outputColMap As Object, _
                                   dataStartRow As Long, dataRowCount As Long, startCol As Long, _
                                   actualStartCol As Long, ByRef endCol As Long, _
                                   ByRef outStdEndCol As Long, ByRef outGapStartCol As Long, ByRef outGapEndCol As Long)
    Dim k As Variant, currentCol As Long, displayName As String
    Dim r As Long, formula As String
    Dim wsMil As Worksheet
    Dim milLastCol As Long, milLastRow As Long
    Dim milHeaderMap As Object
    Dim milLookupRange As String
    Dim ceidColInMil As Long, ceidColInWs As Long, ceidColLetter As String
    Dim milColIdx As Long
    Dim firstSTDCol As Long, lastSTDCol As Long
    Dim actualCol As Long
    Dim milestoneKeys As Variant, idx As Long
    Dim j As Long
    Dim missingMilestones As String
    Dim milestoneFound As Boolean
    Dim foundMilestoneFlags As Object
    Dim milHeaderRow As Long
    Dim stdColMap As Object
    Dim lastDataRow As Long
    
    On Error Resume Next
    Set wsMil = ThisWorkbook.Sheets(SHEET_MILESTONES)
    On Error GoTo 0
    If wsMil Is Nothing Then
        endCol = startCol
        Exit Sub
    End If
    
    lastDataRow = dataStartRow + dataRowCount - 1
    
    ' Detect header row
    milHeaderRow = 1
    milLastCol = wsMil.Cells(1, wsMil.Columns.Count).End(xlToLeft).Column
    For j = 1 To milLastCol
        If LCase(Trim(CStr(wsMil.Cells(2, j).Value))) = "ceid" Then
            milHeaderRow = 2
            Exit For
        End If
    Next j
    
    milLastRow = wsMil.Cells(wsMil.Rows.Count, 1).End(xlUp).row
    milLastCol = wsMil.Cells(milHeaderRow, wsMil.Columns.Count).End(xlToLeft).Column
    
    ' Build header map
    Set milHeaderMap = CreateObject("Scripting.Dictionary")
    For j = 1 To milLastCol
        milHeaderMap(LCase(Trim(CStr(wsMil.Cells(milHeaderRow, j).Value)))) = j
    Next j
    
    ' Find CEID column
    ceidColInMil = 0
    If milHeaderMap.exists("ceid") Then ceidColInMil = milHeaderMap("ceid")
    If ceidColInMil = 0 Then ceidColInMil = 3
    
    milLookupRange = SHEET_MILESTONES & "!$" & ColLetter(ceidColInMil) & "$" & (milHeaderRow + 1) & ":$" & ColLetter(milLastCol) & "$" & milLastRow
    
    ceidColInWs = 0
    If outputColMap.exists("ceid") Then ceidColInWs = outputColMap("ceid")
    If ceidColInWs = 0 Then
        MsgBox "CEID column not found in working sheet. STD Duration columns will be empty.", vbExclamation
        endCol = startCol
        Exit Sub
    End If
    ceidColLetter = ColLetter(ceidColInWs)
    
    currentCol = startCol
    firstSTDCol = startCol
    actualCol = actualStartCol
    missingMilestones = ""
    Set foundMilestoneFlags = CreateObject("Scripting.Dictionary")
    
    milestoneKeys = milestoneGroups.Keys
    Call ShellSortVariantArray(milestoneKeys)
    
    ' First pass: check which milestones exist (with fuzzy fallback)
    For idx = LBound(milestoneKeys) To UBound(milestoneKeys)
        k = milestoneKeys(idx)
        displayName = IIf(milestoneNames.exists(k), milestoneNames(k), CStr(k))
        
        milColIdx = 0
        milestoneFound = False
        
        ' Exact match first
        For j = 1 To milLastCol
            If LCase(Trim(CStr(wsMil.Cells(milHeaderRow, j).Value))) = LCase(Trim(displayName)) Then
                milColIdx = j - ceidColInMil + 1
                milestoneFound = True
                Exit For
            End If
        Next j
        
        ' Fuzzy fallback: contains or starts with
        If Not milestoneFound Then
            For j = 1 To milLastCol
                Dim milHeader As String
                milHeader = LCase(Trim(CStr(wsMil.Cells(milHeaderRow, j).Value)))
                If milHeader <> "" Then
                    If InStr(1, milHeader, LCase(Trim(displayName)), vbTextCompare) > 0 Or _
                       InStr(1, LCase(Trim(displayName)), milHeader, vbTextCompare) > 0 Then
                        milColIdx = j - ceidColInMil + 1
                        milestoneFound = True
                        Exit For
                    End If
                End If
            Next j
        End If
        
        foundMilestoneFlags(k) = Array(milestoneFound, milColIdx, displayName, actualCol)
        If Not milestoneFound Then missingMilestones = missingMilestones & "  - " & displayName & vbCrLf
        actualCol = actualCol + 1
    Next idx
    
    If missingMilestones <> "" Then
        ' Warning already shown in the confirmation dialog — just log it
        DebugLog "Missing milestones: " & Replace(missingMilestones, vbCrLf, ", ")
    End If
    
    ' Second pass: create columns for found milestones
    Set stdColMap = CreateObject("Scripting.Dictionary")
    
    For idx = LBound(milestoneKeys) To UBound(milestoneKeys)
        k = milestoneKeys(idx)
        milestoneFound = foundMilestoneFlags(k)(0)
        milColIdx = foundMilestoneFlags(k)(1)
        displayName = foundMilestoneFlags(k)(2)
        
        If milestoneFound And milColIdx > 0 Then
            ws.Cells(dataStartRow, currentCol).Value = "STD" & vbLf & "Duration" & vbLf & displayName
            
            r = dataStartRow + 1
            formula = "=IFERROR(VLOOKUP(" & ceidColLetter & r & "," & milLookupRange & "," & milColIdx & ",FALSE),"""")"
            ws.Cells(r, currentCol).formula = formula
            If dataRowCount > 2 Then
                ws.Range(ws.Cells(r, currentCol), ws.Cells(lastDataRow, currentCol)).FillDown
            End If
            
            ' Apply duration CF
            Dim actualColForThis As Long
            actualColForThis = foundMilestoneFlags(k)(3)
            ApplyDurationConditionalFormatting ws, actualColForThis, currentCol, dataStartRow, dataRowCount
            
            stdColMap(k) = currentCol
            ws.Columns(currentCol).NumberFormat = "0"
            currentCol = currentCol + 1
        End If
    Next idx
    
    lastSTDCol = currentCol - 1
    
    ' Total STD Duration
    ws.Cells(dataStartRow, currentCol).Value = "Total" & vbLf & "STD" & vbLf & "Duration"
    If lastSTDCol >= firstSTDCol Then
        r = dataStartRow + 1
        formula = "=SUM(" & ColLetter(firstSTDCol) & r & ":" & ColLetter(lastSTDCol) & r & ")"
        ws.Cells(r, currentCol).formula = formula
        If dataRowCount > 2 Then
            ws.Range(ws.Cells(r, currentCol), ws.Cells(lastDataRow, currentCol)).FillDown
        End If
    End If
    ws.Columns(currentCol).NumberFormat = "0"
    currentCol = currentCol + 1
    
    ' Track STD section end (includes Total STD column)
    outStdEndCol = currentCol - 1
    
    ' Gap columns
    Dim gapFirstCol As Long
    Dim gapCFRange As Range
    Set gapCFRange = Nothing
    gapFirstCol = currentCol
    For idx = LBound(milestoneKeys) To UBound(milestoneKeys)
        k = milestoneKeys(idx)
        If stdColMap.exists(k) Then
            displayName = IIf(milestoneNames.exists(k), milestoneNames(k), CStr(k))
            ws.Cells(dataStartRow, currentCol).Value = "Gap" & vbLf & displayName
            
            Dim actualColForGap As Long, stdColForGap As Long
            actualColForGap = foundMilestoneFlags(k)(3)
            stdColForGap = stdColMap(k)
            
            r = dataStartRow + 1
            formula = "=IF(AND(ISNUMBER(" & ColLetter(actualColForGap) & r & "),ISNUMBER(" & ColLetter(stdColForGap) & r & "))," & _
                      ColLetter(actualColForGap) & r & "-" & ColLetter(stdColForGap) & r & ","""")"
            ws.Cells(r, currentCol).formula = formula
            If dataRowCount > 2 Then
                ws.Range(ws.Cells(r, currentCol), ws.Cells(lastDataRow, currentCol)).FillDown
            End If
            
            ' Collect gap column into union range
            Dim gapColRng As Range
            Set gapColRng = ws.Range(ws.Cells(dataStartRow + 1, currentCol), _
                                      ws.Cells(dataStartRow + dataRowCount - 1, currentCol))
            If gapCFRange Is Nothing Then
                Set gapCFRange = gapColRng
            Else
                Set gapCFRange = Union(gapCFRange, gapColRng)
            End If
            
            ws.Columns(currentCol).NumberFormat = "0"
            currentCol = currentCol + 1
        End If
    Next idx
    
    ' Apply gap CF once to all gap columns (3 rules instead of 3 x N milestones)
    If Not gapCFRange Is Nothing Then
        Dim gapFC As FormatCondition, gapAddr As String
        gapAddr = gapCFRange.Areas(1).Cells(1, 1).Address(False, False)
        gapCFRange.FormatConditions.Delete
        
        Set gapFC = gapCFRange.FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=AND(ISNUMBER(" & gapAddr & ")," & gapAddr & "<0)")
        gapFC.Interior.Color = RGB(254, 226, 226): gapFC.Font.Color = RGB(153, 27, 27): gapFC.StopIfTrue = True

        Set gapFC = gapCFRange.FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=AND(ISNUMBER(" & gapAddr & ")," & gapAddr & ">=0," & gapAddr & "<=4)")
        gapFC.Interior.Color = RGB(209, 250, 229): gapFC.Font.Color = RGB(6, 95, 70): gapFC.StopIfTrue = True

        Set gapFC = gapCFRange.FormatConditions.Add(Type:=xlExpression, _
            Formula1:="=AND(ISNUMBER(" & gapAddr & ")," & gapAddr & ">4)")
        gapFC.Interior.Color = RGB(254, 243, 199): gapFC.Font.Color = RGB(146, 64, 14): gapFC.StopIfTrue = True
    End If
    
    ' Track gap section boundaries
    If currentCol > gapFirstCol Then
        outGapStartCol = gapFirstCol
        outGapEndCol = currentCol - 1
    Else
        outGapStartCol = 0
        outGapEndCol = 0
    End If
    
    endCol = currentCol
End Sub

'====================================================================
' ADD ADDITIONAL COLUMNS (New/Reused/Demo, Escalated, Tool S/N, etc.)
'====================================================================

Private Sub AddAdditionalColumns(ws As Worksheet, outputColMap As Object, _
                                  dataStartRow As Long, dataRowCount As Long, startCol As Long, _
                                  ByRef newReusedColPos As Long, ByRef estCycleColPos As Long, _
                                  Optional milestoneGroups As Object = Nothing)
    Dim currentCol As Long, r As Long, formula As String
    Dim entityCodeCol As Long, sddCol As Long, sqFinishCol As Long, setStartCol As Long
    Dim eventTypeCol As Long
    Dim entityCodeLetter As String, sddLetter As String, sqFinishLetter As String, setStartLetter As String
    Dim eventTypeLetter As String
    Dim hasNewReusedSheet As Boolean, hasSNSheet As Boolean
    Dim newReusedColLetter As String, shipDateLetter As String
    Dim lastDataRow As Long
    
    currentCol = startCol
    lastDataRow = dataStartRow + dataRowCount - 1
    
    hasNewReusedSheet = SheetExists(ThisWorkbook, SHEET_NEW_REUSED)
    hasSNSheet = SheetExists(ThisWorkbook, SHEET_SN)
    
    ' Get column indices
    entityCodeCol = 0: sddCol = 0: sqFinishCol = 0: setStartCol = 0: eventTypeCol = 0
    If outputColMap.exists(LCase(HEADER_ENTITY_CODE)) Then entityCodeCol = outputColMap(LCase(HEADER_ENTITY_CODE))
    If outputColMap.exists("sdd") Then sddCol = outputColMap("sdd")
    If outputColMap.exists(LCase(HEADER_SUPPLIER_QUAL_FINISH)) Then sqFinishCol = outputColMap(LCase(HEADER_SUPPLIER_QUAL_FINISH))
    If outputColMap.exists(LCase(HEADER_SET_START)) Then setStartCol = outputColMap(LCase(HEADER_SET_START))
    If outputColMap.exists(LCase(HEADER_EVENT_TYPE)) Then eventTypeCol = outputColMap(LCase(HEADER_EVENT_TYPE))
    
    If entityCodeCol > 0 Then entityCodeLetter = ColLetter(entityCodeCol)
    If sddCol > 0 Then sddLetter = ColLetter(sddCol)
    If sqFinishCol > 0 Then sqFinishLetter = ColLetter(sqFinishCol)
    If setStartCol > 0 Then setStartLetter = ColLetter(setStartCol)
    If eventTypeCol > 0 Then eventTypeLetter = ColLetter(eventTypeCol)

    ' Validate critical columns exist
    Dim missingCols As String
    missingCols = ""
    If entityCodeCol = 0 Then missingCols = missingCols & "Entity Code, "
    If eventTypeCol = 0 Then missingCols = missingCols & "Event Type, "
    If setStartCol = 0 Then missingCols = missingCols & "Set Start, "
    If missingCols <> "" Then
        DebugLog "WARNING: Missing columns: " & Left(missingCols, Len(missingCols) - 2)
    End If

    ' === COLUMN 1: New/Reused/Demo ===
    ws.Cells(dataStartRow, currentCol).Value = "New/" & vbLf & "Reused"
    newReusedColPos = currentCol
    outputColMap("new/reused") = currentCol
    newReusedColLetter = ColLetter(currentCol)
    
    r = dataStartRow + 1
    If eventTypeCol > 0 And hasNewReusedSheet And entityCodeCol > 0 Then
        ' Priority: Demo (from Event Type) > New (from sheet lookup) > Reused (default)
        formula = "=IF(LOWER(" & eventTypeLetter & r & ")=""demo"",""Demo""," & _
                 "IF(IFERROR(VLOOKUP(" & entityCodeLetter & r & ",'New-reused'!$D:$G,4,FALSE),"""")=""New"",""New"",""Reused""))"
        ws.Cells(r, currentCol).formula = formula
    ElseIf eventTypeCol > 0 And entityCodeCol > 0 Then
        ' No New-reused sheet, just check Demo
        formula = "=IF(LOWER(" & eventTypeLetter & r & ")=""demo"",""Demo"",""Reused"")"
        ws.Cells(r, currentCol).formula = formula
    ElseIf hasNewReusedSheet And entityCodeCol > 0 Then
        ' No Event Type column
        formula = "=IF(IFERROR(VLOOKUP(" & entityCodeLetter & r & ",'New-reused'!$D:$G,4,FALSE),"""")=""New"",""New"",""Reused"")"
        ws.Cells(r, currentCol).formula = formula
    Else
        ws.Cells(r, currentCol).Value = "Reused"
    End If
    If dataRowCount > 2 Then
        ws.Range(ws.Cells(r, currentCol), ws.Cells(lastDataRow, currentCol)).FillDown
    End If
    
    ' Apply New/Reused/Demo conditional formatting
    ApplyNewReusedConditionalFormatting ws, currentCol, dataStartRow, dataRowCount
    currentCol = currentCol + 1
    
    ' === COLUMN 2: Escalated ===
    ws.Cells(dataStartRow, currentCol).Value = "Escalated"
    Dim escRange As Range
    Set escRange = ws.Range(ws.Cells(dataStartRow + 1, currentCol), ws.Cells(lastDataRow, currentCol))
    With escRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=",Escalated,Watched"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
    End With
    ApplyEscalatedConditionalFormatting ws, currentCol, dataStartRow, dataRowCount
    currentCol = currentCol + 1
    
    ' === COLUMN 3: Tool S/N ===
    ws.Cells(dataStartRow, currentCol).Value = "Tool" & vbLf & "S/N"
    If hasSNSheet And entityCodeCol > 0 Then
        r = dataStartRow + 1
        formula = "=IFERROR(VLOOKUP(" & entityCodeLetter & r & ",SN!$A:$B,2,FALSE),"""")"
        ws.Cells(r, currentCol).formula = formula
        If dataRowCount > 2 Then
            ws.Range(ws.Cells(r, currentCol), ws.Cells(lastDataRow, currentCol)).FillDown
        End If
    End If
    currentCol = currentCol + 1
    
    ' === COLUMN 4: Ship Date ===
    ws.Cells(dataStartRow, currentCol).Value = "Ship" & vbLf & "Date"
    shipDateLetter = ColLetter(currentCol)
    ws.Columns(currentCol).NumberFormat = "mm/dd/yyyy"
    currentCol = currentCol + 1
    
    ' === COLUMN 5: Pre-Install Meeting ===
    ws.Cells(dataStartRow, currentCol).Value = "Pre-Install" & vbLf & "Meeting"
    r = dataStartRow + 1
    If sddCol > 0 Then
        formula = "=IF(" & newReusedColLetter & r & "=""New""," & _
                 "IF(ISNUMBER(" & shipDateLetter & r & ")," & shipDateLetter & r & "-60," & _
                 "IF(ISNUMBER(" & sddLetter & r & ")," & sddLetter & r & "-74,"""")),"""")"
    Else
        formula = "=IF(" & newReusedColLetter & r & "=""New""," & _
                 "IF(ISNUMBER(" & shipDateLetter & r & ")," & shipDateLetter & r & "-60,""""),"""")"
    End If
    ws.Cells(r, currentCol).formula = formula
    If dataRowCount > 2 Then
        ws.Range(ws.Cells(r, currentCol), ws.Cells(lastDataRow, currentCol)).FillDown
    End If
    ApplyPreInstallConditionalFormatting ws, currentCol, newReusedColPos, dataStartRow, dataRowCount
    ws.Columns(currentCol).NumberFormat = "mm/dd/yyyy"
    currentCol = currentCol + 1
    
    ' === COLUMN 6: Est CAR Date ===
    ws.Cells(dataStartRow, currentCol).Value = "Est CAR" & vbLf & "Date"
    If sqFinishCol > 0 Then
        r = dataStartRow + 1
        formula = "=IF(" & newReusedColLetter & r & "=""New""," & _
                 "IF(ISNUMBER(" & sqFinishLetter & r & ")," & sqFinishLetter & r & "+14,""""),"""")"
        ws.Cells(r, currentCol).formula = formula
        If dataRowCount > 2 Then
            ws.Range(ws.Cells(r, currentCol), ws.Cells(lastDataRow, currentCol)).FillDown
        End If
    End If
    ApplyEstCARConditionalFormatting ws, currentCol, dataStartRow, dataRowCount
    ws.Columns(currentCol).NumberFormat = "mm/dd/yyyy"
    currentCol = currentCol + 1
    
    ' === COLUMN 7: Est Cycle Time ===
    ws.Cells(dataStartRow, currentCol).Value = "Est Cycle" & vbLf & "Time"
    estCycleColPos = currentCol
    If sqFinishCol > 0 And setStartCol > 0 Then
        r = dataStartRow + 1
        formula = "=IF(" & newReusedColLetter & r & "=""New""," & _
                 "IF(AND(ISNUMBER(" & sqFinishLetter & r & "),ISNUMBER(" & setStartLetter & r & "))," & _
                 "(" & sqFinishLetter & r & "+14)-" & setStartLetter & r & ",""""),"""")"
        ws.Cells(r, currentCol).formula = formula
        If dataRowCount > 2 Then
            ws.Range(ws.Cells(r, currentCol), ws.Cells(lastDataRow, currentCol)).FillDown
        End If
    End If
    ws.Columns(currentCol).NumberFormat = "0"
    currentCol = currentCol + 1
    
    ' === COLUMNS 8-11: Manual entry ===
    ws.Cells(dataStartRow, currentCol).Value = "SOC" & vbLf & "Available"
    outputColMap("soc available") = currentCol
    currentCol = currentCol + 1
    ws.Cells(dataStartRow, currentCol).Value = "SOC" & vbLf & "Uploaded?": currentCol = currentCol + 1
    ws.Cells(dataStartRow, currentCol).Value = "Staffed?": currentCol = currentCol + 1
    ws.Cells(dataStartRow, currentCol).Value = "Comments"
    ws.Columns(currentCol).ColumnWidth = 30
    currentCol = currentCol + 1
    
    ' === COLUMN 12: BOD1 (Blackout Date 1) ===
    ws.Cells(dataStartRow, currentCol).Value = "BOD1"
    ws.Columns(currentCol).ColumnWidth = 10
    If lastDataRow > dataStartRow Then
        ws.Range(ws.Cells(dataStartRow + 1, currentCol), ws.Cells(lastDataRow, currentCol)).NumberFormat = "mm/dd/yyyy"
        ws.Range(ws.Cells(dataStartRow + 1, currentCol), ws.Cells(lastDataRow, currentCol)).HorizontalAlignment = xlCenter
    End If
    outputColMap("bod1") = currentCol
    currentCol = currentCol + 1
    
    ' === COLUMN 14: BOD2 (Blackout Date 2) ===
    ws.Cells(dataStartRow, currentCol).Value = "BOD2"
    ws.Columns(currentCol).ColumnWidth = 10
    If lastDataRow > dataStartRow Then
        ws.Range(ws.Cells(dataStartRow + 1, currentCol), ws.Cells(lastDataRow, currentCol)).NumberFormat = "mm/dd/yyyy"
        ws.Range(ws.Cells(dataStartRow + 1, currentCol), ws.Cells(lastDataRow, currentCol)).HorizontalAlignment = xlCenter
    End If
    outputColMap("bod2") = currentCol
    currentCol = currentCol + 1
    
End Sub
'====================================================================
' ADD SUMMARY DASHBOARD (Rows 2-5) - Table-based formulas
'====================================================================

Private Sub AddSummaryDashboard(ws As Worksheet, dataStartRow As Long, dataRowCount As Long, _
                                  totalColCount As Long, newReusedCol As Long, estCycleCol As Long, _
                                  entityTypeCol As Long, tableName As String, _
                                  Optional sddCol As Long = 0, Optional setStartCol As Long = 0, _
                                  Optional sqFinishCol As Long = 0, Optional ctThreshold As Long = 85, _
                                  Optional escalatedCol As Long = 0, Optional watchCol As Long = 0, _
                                  Optional colMap As Object = Nothing, _
                                  Optional milestoneGroups As Object = Nothing)
    Dim startCol As Long, currentCol As Long
    Dim nrRange As String, ctRange As String, ssRange As String, sqfRange As String
    Dim escRange As String, watchRange As String
    Dim tfCell As String
    Dim cardBg As Long, cardBorder As Long, labelColor As Long
    Dim cutoffExpr As String, tfFilter As String, visBase As String, activeBase As String
    Dim newCol As Long, reusedCol As Long, demoCol As Long
    Dim tfCol As Long
    
    cardBg = CLR_CARD_BG
    cardBorder = CLR_CARD_BORDER
    labelColor = CLR_MUTED_TEXT

    ' ---- Build structured table references (auto-adjust to table size) ----
    ' Read actual header text from row 15 for each column (includes vbLf)
    If newReusedCol > 0 Then nrRange = tableName & "[" & ws.Cells(dataStartRow, newReusedCol).Value & "]"
    If estCycleCol > 0 Then ctRange = tableName & "[" & ws.Cells(dataStartRow, estCycleCol).Value & "]"
    If setStartCol > 0 Then ssRange = tableName & "[" & ws.Cells(dataStartRow, setStartCol).Value & "]"
    If sqFinishCol > 0 Then sqfRange = tableName & "[" & ws.Cells(dataStartRow, sqFinishCol).Value & "]"
    If escalatedCol > 0 Then escRange = tableName & "[" & ws.Cells(dataStartRow, escalatedCol).Value & "]"
    If watchCol > 0 Then watchRange = tableName & "[" & ws.Cells(dataStartRow, watchCol).Value & "]"

    ' Header cell ref for OFFSET anchor (uses [#Headers] structured ref)
    Dim nrHdrRef As String
    If newReusedCol > 0 Then nrHdrRef = tableName & "[[#Headers],[" & ws.Cells(dataStartRow, newReusedCol).Value & "]]"

    ' Time Frame input
    startCol = entityTypeCol: If startCol < 5 Then startCol = 5
    ws.Cells(2, startCol).Value = "Time Frame (weeks):"
    With ws.Cells(2, startCol): .Font.Size = 8: .Font.Color = labelColor: .Font.Bold = True: .HorizontalAlignment = xlRight: End With
    tfCol = startCol + 1
    ws.Cells(2, tfCol).Value = ""
    With ws.Cells(2, tfCol): .Font.Size = 11: .Font.Bold = True: .Font.Color = CLR_DARK_TEXT: .HorizontalAlignment = xlCenter: .Interior.Color = CLR_TF_INPUT_BG: End With
    With ws.Cells(2, tfCol).Borders: .LineStyle = xlContinuous: .Weight = xlThin: .Color = CLR_TF_INPUT_BORDER: End With
    tfCell = ColLetter(tfCol) & "2"

    ' Time frame filter: checks Definition milestone start dates only (not decon/demo/SDD)
    ' Uses structured table refs so formula auto-adjusts to table size.
    If Not colMap Is Nothing And Not milestoneGroups Is Nothing Then
        Dim tfParts As String
        Dim mgKey As Variant, mgGrp As Object, mgHdr As String, mgTblRef As String
        tfParts = ""
        cutoffExpr = "IF(" & tfCell & "="""",DATE(2099,12,31),TODAY()+" & tfCell & "*7)"
        For Each mgKey In milestoneGroups.keys
            Set mgGrp = milestoneGroups(mgKey)
            If mgGrp.exists(1) Then
                mgHdr = LCase(CStr(mgGrp(1)(2)))
                If colMap.exists(mgHdr) Then
                    mgTblRef = tableName & "[" & ws.Cells(dataStartRow, colMap(mgHdr)).Value & "]"
                    If tfParts <> "" Then tfParts = tfParts & "+"
                    tfParts = tfParts & "(ISNUMBER(" & mgTblRef & ")*(" & mgTblRef & "<=" & cutoffExpr & "))"
                End If
            End If
        Next mgKey
        ' Ensure Set Start is included
        If setStartCol > 0 And InStr(1, tfParts, ws.Cells(dataStartRow, setStartCol).Value) = 0 Then
            If tfParts <> "" Then tfParts = tfParts & "+"
            tfParts = tfParts & "(ISNUMBER(" & ssRange & ")*(" & ssRange & "<=" & cutoffExpr & "))"
        End If
        If tfParts <> "" Then
            tfFilter = "*((" & tfParts & ")>0)"
        Else
            tfFilter = ""
        End If
    Else
        tfFilter = ""
    End If

    ' visBase: per-row SUBTOTAL for slicer-aware counting (structured ref version)
    If newReusedCol > 0 Then
        visBase = "SUBTOTAL(103,OFFSET(" & nrHdrRef & ",ROW(" & nrRange & ")-ROW(" & nrHdrRef & "),0,1,1))"
    Else
        visBase = "1"
    End If
    
    currentCol = tfCol + 2
    
    ' Card 1: Total (exact Rev7)
    ws.Cells(3, currentCol).Value = "Total": ws.Cells(3, currentCol).Font.Size = 8: ws.Cells(3, currentCol).Font.Color = labelColor: ws.Cells(3, currentCol).HorizontalAlignment = xlCenter
    If newReusedCol > 0 And setStartCol > 0 Then
        ws.Cells(4, currentCol).formula = "=SUMPRODUCT(" & visBase & tfFilter & ")"
    ElseIf newReusedCol > 0 Then
        ws.Cells(4, currentCol).formula = "=SUBTOTAL(103," & nrRange & ")"
    Else
        ws.Cells(4, currentCol).Value = 0
    End If
    ws.Cells(4, currentCol).Font.Size = 18: ws.Cells(4, currentCol).Font.Bold = True: ws.Cells(4, currentCol).Font.Color = CLR_DARK_TEXT: ws.Cells(4, currentCol).HorizontalAlignment = xlCenter
    FormatDashboardCard ws, 3, currentCol, 4, currentCol, cardBg, cardBorder
    currentCol = currentCol + 1
    
    ' Card 2: New (exact Rev7)
    ws.Cells(3, currentCol).Value = "New": ws.Cells(3, currentCol).Font.Size = 8: ws.Cells(3, currentCol).Font.Color = labelColor: ws.Cells(3, currentCol).HorizontalAlignment = xlCenter
    newCol = currentCol
    If newReusedCol > 0 Then
        ws.Cells(4, currentCol).formula = "=SUMPRODUCT(" & visBase & "*(" & nrRange & "=""New"")" & tfFilter & ")"
    Else: ws.Cells(4, currentCol).Value = 0
    End If
    ws.Cells(4, currentCol).Font.Size = 18: ws.Cells(4, currentCol).Font.Bold = True: ws.Cells(4, currentCol).Font.Color = CLR_NEW: ws.Cells(4, currentCol).HorizontalAlignment = xlCenter
    FormatDashboardCard ws, 3, currentCol, 4, currentCol, cardBg, cardBorder
    currentCol = currentCol + 1
    
    ' Card 3: Reused (exact Rev7)
    ws.Cells(3, currentCol).Value = "Reused": ws.Cells(3, currentCol).Font.Size = 8: ws.Cells(3, currentCol).Font.Color = labelColor: ws.Cells(3, currentCol).HorizontalAlignment = xlCenter
    reusedCol = currentCol
    If newReusedCol > 0 Then
        ws.Cells(4, currentCol).formula = "=SUMPRODUCT(" & visBase & "*(" & nrRange & "=""Reused"")" & tfFilter & ")"
    Else: ws.Cells(4, currentCol).Value = 0
    End If
    ws.Cells(4, currentCol).Font.Size = 18: ws.Cells(4, currentCol).Font.Bold = True: ws.Cells(4, currentCol).Font.Color = CLR_REUSED: ws.Cells(4, currentCol).HorizontalAlignment = xlCenter
    FormatDashboardCard ws, 3, currentCol, 4, currentCol, cardBg, cardBorder
    currentCol = currentCol + 1
    
    ' Card 4: Demo — no tfFilter (Demo systems use Decon/Demo milestones, not Set Start)
    ws.Cells(3, currentCol).Value = "Demo": ws.Cells(3, currentCol).Font.Size = 8: ws.Cells(3, currentCol).Font.Color = labelColor: ws.Cells(3, currentCol).HorizontalAlignment = xlCenter
    demoCol = currentCol
    If newReusedCol > 0 Then
        ws.Cells(4, currentCol).formula = "=SUMPRODUCT(" & visBase & "*(" & nrRange & "=""Demo""))"
    Else: ws.Cells(4, currentCol).Value = 0
    End If
    ws.Cells(4, currentCol).Font.Size = 18: ws.Cells(4, currentCol).Font.Bold = True: ws.Cells(4, currentCol).Font.Color = CLR_DEMO: ws.Cells(4, currentCol).HorizontalAlignment = xlCenter
    FormatDashboardCard ws, 3, currentCol, 4, currentCol, cardBg, cardBorder
    currentCol = currentCol + 1
    
    ' Card 5: CT est miss (exact Rev7)
    ws.Cells(3, currentCol).Value = "CT est miss": ws.Cells(3, currentCol).Font.Size = 8: ws.Cells(3, currentCol).Font.Color = labelColor: ws.Cells(3, currentCol).HorizontalAlignment = xlCenter
    If estCycleCol > 0 And newReusedCol > 0 Then
        ws.Cells(4, currentCol).formula = "=SUMPRODUCT(" & visBase & "*(" & nrRange & "=""New"")*(" & ctRange & ">" & ctThreshold & ")" & tfFilter & ")"
    Else: ws.Cells(4, currentCol).Value = 0
    End If
    ws.Cells(4, currentCol).Font.Size = 18: ws.Cells(4, currentCol).Font.Bold = True: ws.Cells(4, currentCol).Font.Color = CLR_CT_MISS: ws.Cells(4, currentCol).HorizontalAlignment = xlCenter
    FormatDashboardCard ws, 3, currentCol, 4, currentCol, cardBg, cardBorder
    currentCol = currentCol + 1
    
    ' Card 6: Escalated (exact Rev7)
    ws.Cells(3, currentCol).Value = "Escalated": ws.Cells(3, currentCol).Font.Size = 8: ws.Cells(3, currentCol).Font.Color = labelColor: ws.Cells(3, currentCol).HorizontalAlignment = xlCenter
    If escalatedCol > 0 Then
        ws.Cells(4, currentCol).formula = "=SUMPRODUCT(" & visBase & "*(" & escRange & "=""Escalated""))"
    Else: ws.Cells(4, currentCol).Value = 0
    End If
    ws.Cells(4, currentCol).Font.Size = 18: ws.Cells(4, currentCol).Font.Bold = True: ws.Cells(4, currentCol).Font.Color = CLR_ESCALATED: ws.Cells(4, currentCol).HorizontalAlignment = xlCenter
    ws.Cells(4, currentCol).NumberFormat = "General"
    FormatDashboardCard ws, 3, currentCol, 4, currentCol, cardBg, cardBorder
    currentCol = currentCol + 1
    
    ' Card 7: Watched (exact Rev7)
    ws.Cells(3, currentCol).Value = "Watched": ws.Cells(3, currentCol).Font.Size = 8: ws.Cells(3, currentCol).Font.Color = labelColor: ws.Cells(3, currentCol).HorizontalAlignment = xlCenter
    If watchCol > 0 Then
        ws.Cells(4, currentCol).formula = "=SUMPRODUCT(" & visBase & "*(" & watchRange & "=""Watched""))"
    Else: ws.Cells(4, currentCol).Value = 0
    End If
    ws.Cells(4, currentCol).Font.Size = 18: ws.Cells(4, currentCol).Font.Bold = True: ws.Cells(4, currentCol).Font.Color = CLR_WATCHED: ws.Cells(4, currentCol).HorizontalAlignment = xlCenter
    ws.Cells(4, currentCol).NumberFormat = "General"
    FormatDashboardCard ws, 3, currentCol, 4, currentCol, cardBg, cardBorder
    
    ' Active row: inline SUMPRODUCT (no helper columns)
    ' Active = Set Start <= TODAY AND SQ Finish >= TODAY (started but not finished)
    ws.Cells(5, newCol - 1).Value = "Active:"
    With ws.Cells(5, newCol - 1): .Font.Size = 8: .Font.Bold = True: .Font.Color = CLR_WATCHED: .HorizontalAlignment = xlRight: End With

    If setStartCol > 0 And sqFinishCol > 0 And newReusedCol > 0 Then
        activeBase = visBase & "*(ISNUMBER(" & ssRange & ")*(" & ssRange & "<=TODAY())" & _
                     "*ISNUMBER(" & sqfRange & ")*(" & sqfRange & ">=TODAY()))"
        ws.Cells(5, newCol).formula = "=SUMPRODUCT(" & activeBase & "*(" & nrRange & "=""New""))"
        ws.Cells(5, reusedCol).formula = "=SUMPRODUCT(" & activeBase & "*(" & nrRange & "=""Reused""))"
        ws.Cells(5, demoCol).formula = "=SUMPRODUCT(" & activeBase & "*(" & nrRange & "=""Demo""))"
    Else
        ws.Cells(5, newCol).Value = "N/A"
        ws.Cells(5, reusedCol).Value = "N/A"
        ws.Cells(5, demoCol).Value = "N/A"
    End If
    ws.Cells(5, newCol).Font.Size = 11: ws.Cells(5, newCol).Font.Bold = True: ws.Cells(5, newCol).Font.Color = CLR_NEW: ws.Cells(5, newCol).HorizontalAlignment = xlCenter
    ws.Cells(5, reusedCol).Font.Size = 11: ws.Cells(5, reusedCol).Font.Bold = True: ws.Cells(5, reusedCol).Font.Color = CLR_REUSED: ws.Cells(5, reusedCol).HorizontalAlignment = xlCenter
    ws.Cells(5, demoCol).Font.Size = 11: ws.Cells(5, demoCol).Font.Bold = True: ws.Cells(5, demoCol).Font.Color = CLR_DEMO: ws.Cells(5, demoCol).HorizontalAlignment = xlCenter
    
    ' Row heights
    ws.Rows(2).RowHeight = 18: ws.Rows(3).RowHeight = 16: ws.Rows(4).RowHeight = 28
    ws.Rows(5).RowHeight = 18: ws.Rows(6).RowHeight = 6
End Sub


'====================================================================
' FORMAT DASHBOARD CARD (single cell style, no merge)
'====================================================================

Private Sub FormatDashboardCard(ws As Worksheet, r1 As Long, c1 As Long, r2 As Long, c2 As Long, _
                                  bgColor As Long, borderColor As Long)
    ' Apply background and centering (no merge)
    Dim cardRange As Range
    Set cardRange = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
    cardRange.Interior.Color = bgColor
    cardRange.HorizontalAlignment = xlCenter
    
    ' Borders
    With cardRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = borderColor
    End With
    With cardRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = borderColor
    End With
    With cardRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = borderColor
    End With
    With cardRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = borderColor
    End With
End Sub

'====================================================================
' ADD SLICERS (Group + New/Reused on ListObject table)
'====================================================================

Private Sub AddSlicers(ws As Worksheet, tbl As ListObject, outputColMap As Object, _
                        lastCounterCol As Long)
    On Error GoTo SlicerError
    
    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim scName As String
    Dim slicerLeft As Double, slicerTop As Double
    Dim nrHeaderName As String
    Dim c As Long
    Dim rawH As String
    
    ' No rename, no cleanup — just use unique timestamps for cache names
    ' and let Excel auto-name slicer shapes (omit Name param in Slicers.Add)
    
    ' Position slicers right after the last counter card
    slicerTop = ws.Cells(3, 1).Top
    slicerLeft = ws.Cells(1, lastCounterCol + 2).Left
    
    DebugLog "AddSlicers: table=" & tbl.Name
    
    ' Debug: list all existing slicer caches
    Dim dbgSC As SlicerCache
    DebugLog "AddSlicers: existing caches BEFORE:"
    For Each dbgSC In ThisWorkbook.SlicerCaches
        DebugLog "  cache='" & dbgSC.Name & "' slicers=" & dbgSC.Slicers.Count
    Next dbgSC
    
    ' === SLICER 1: Group ===
    If outputColMap.exists("group") Then
        scName = "Slicer_Group_" & Replace(tbl.Name, " ", "_") & "_" & Format(Now, "hhnnss")
        DebugLog "AddSlicers: creating Group cache: " & scName
        
        On Error Resume Next
        Set sc = Nothing
        Set sc = ThisWorkbook.SlicerCaches.Add2(tbl, "Group", scName)
        If Err.Number <> 0 Then DebugLog "AddSlicers: Group cache FAILED: " & Err.Description & " (#" & Err.Number & ")"
        On Error GoTo SlicerError
        
        DebugLog "AddSlicers: Group sc Is Nothing = " & (sc Is Nothing)
        
        If Not sc Is Nothing Then
            On Error Resume Next
            ' Omit Name parameter — let Excel auto-generate unique name (BKM)
            Set sl = sc.Slicers.Add(ws, , , "Group", _
                                     slicerTop, slicerLeft, 550, 65)
            If Err.Number <> 0 Then DebugLog "AddSlicers: Group slicer add FAILED: " & Err.Description & " (#" & Err.Number & ")"
            DebugLog "AddSlicers: Group sl Is Nothing = " & (sl Is Nothing)
            If Not sl Is Nothing Then
                DebugLog "AddSlicers: Group slicer auto-named: " & sl.Name
                sl.Style = "SlicerStyleLight1"
                sl.NumberOfColumns = 7
            End If
            On Error GoTo SlicerError
            slicerTop = slicerTop + 70
        End If
    End If
    
    ' === SLICER 2: New/Reused ===
    nrHeaderName = ""
    For c = 1 To tbl.Range.Columns.Count
        rawH = LCase(Trim(Replace(Replace(CStr(tbl.HeaderRowRange.Cells(1, c).Value), vbLf, ""), vbCr, "")))
        If rawH = "new/reused" Or InStr(1, rawH, "new/reused", vbTextCompare) > 0 Then
            nrHeaderName = tbl.HeaderRowRange.Cells(1, c).Value
            Exit For
        End If
    Next c
    
    If nrHeaderName <> "" Then
        scName = "Slicer_NewReused_" & Replace(tbl.Name, " ", "_") & "_" & Format(Now, "hhnnss")
        
        On Error Resume Next
        Set sc = Nothing
        Set sc = ThisWorkbook.SlicerCaches.Add2(tbl, nrHeaderName, scName)
        On Error GoTo SlicerError
        
        If Not sc Is Nothing Then
            On Error Resume Next
            Set sl = sc.Slicers.Add(ws, , , "New/Reused", _
                                     slicerTop, slicerLeft, 220, 55)
            If Not sl Is Nothing Then
                sl.Style = "SlicerStyleLight1"
                sl.NumberOfColumns = 3
            End If
            On Error GoTo SlicerError
            slicerTop = slicerTop + 60
        End If
    End If
    
    ' === SLICER 3: Event Type (Rev8) ===
    If outputColMap.exists(LCase("Event Type")) Then
        Dim etHeaderName As String
        etHeaderName = ""
        For c = 1 To tbl.Range.Columns.Count
            rawH = LCase(Trim(Replace(Replace(CStr(tbl.HeaderRowRange.Cells(1, c).Value), vbLf, ""), vbCr, "")))
            If rawH = "event type" Then
                etHeaderName = tbl.HeaderRowRange.Cells(1, c).Value
                Exit For
            End If
        Next c
        
        If etHeaderName <> "" Then
            scName = "Slicer_EventType_" & Replace(tbl.Name, " ", "_") & "_" & Format(Now, "hhnnss")
            
            On Error Resume Next
            Set sc = Nothing
            Set sc = ThisWorkbook.SlicerCaches.Add2(tbl, etHeaderName, scName)
            On Error GoTo SlicerError
            
            If Not sc Is Nothing Then
                On Error Resume Next
                Set sl = sc.Slicers.Add(ws, , , "Event Type", _
                                         slicerTop, slicerLeft, 400, 55)
                If Not sl Is Nothing Then
                    sl.Style = "SlicerStyleLight1"
                    sl.NumberOfColumns = 5
                End If
                On Error GoTo SlicerError
            End If
        End If
    End If
    
    ' Debug: list all slicer caches after creation
    DebugLog "AddSlicers: existing caches AFTER:"
    For Each dbgSC In ThisWorkbook.SlicerCaches
        DebugLog "  cache='" & dbgSC.Name & "' slicers=" & dbgSC.Slicers.Count
    Next dbgSC
    
    Exit Sub
    
SlicerError:
    ' Slicers are non-critical - log and attempt to continue with remaining slicers
    DebugLog "AddSlicers warning: " & Err.Description & " (#" & Err.Number & ")"
    Err.Clear
    Resume Next
End Sub

'====================================================================
' APPLY NEW/REUSED/DEMO CONDITIONAL FORMATTING
'====================================================================

Private Sub ApplyNewReusedConditionalFormatting(ws As Worksheet, col As Long, _
                                                 dataStartRow As Long, dataRowCount As Long)
    Dim rng As Range, fc As FormatCondition
    Dim addr As String, firstRow As Long
    
    firstRow = dataStartRow + 1
    Set rng = ws.Range(ws.Cells(firstRow, col), ws.Cells(dataStartRow + dataRowCount - 1, col))
    addr = ws.Cells(firstRow, col).Address(False, False)
    rng.FormatConditions.Delete
    
    ' New = Green (SA9 light)
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & addr & "=""New""")
    fc.Interior.Color = RGB(209, 250, 229)
    fc.Font.Color = RGB(6, 95, 70)
    fc.Font.Bold = True
    fc.StopIfTrue = True

    ' Demo = Purple (SA9 light)
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & addr & "=""Demo""")
    fc.Interior.Color = RGB(237, 220, 255)
    fc.Font.Color = RGB(107, 33, 168)
    fc.Font.Bold = True
    fc.StopIfTrue = True

    ' Reused = Gray (SA9 light)
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & addr & "=""Reused""")
    fc.Interior.Color = RGB(235, 238, 242)
    fc.Font.Color = RGB(75, 85, 99)
    fc.StopIfTrue = True
End Sub

'====================================================================
' APPLY ESCALATED CONDITIONAL FORMATTING
'====================================================================

Private Sub ApplyEscalatedConditionalFormatting(ws As Worksheet, col As Long, _
                                                  dataStartRow As Long, dataRowCount As Long)
    Dim rng As Range, fc As FormatCondition
    Dim addr As String, firstRow As Long
    
    firstRow = dataStartRow + 1
    Set rng = ws.Range(ws.Cells(firstRow, col), ws.Cells(dataStartRow + dataRowCount - 1, col))
    addr = ws.Cells(firstRow, col).Address(False, False)
    rng.FormatConditions.Delete
    
    ' Escalated = Red (SA9 light)
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & addr & "=""Escalated""")
    fc.Interior.Color = RGB(254, 226, 226)
    fc.Font.Color = RGB(153, 27, 27)
    fc.Font.Bold = True
    fc.StopIfTrue = True

    ' Watched = Amber (SA9 light)
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & addr & "=""Watched""")
    fc.Interior.Color = RGB(254, 243, 199)
    fc.Font.Color = RGB(146, 64, 14)
    fc.Font.Bold = True
    fc.StopIfTrue = True
End Sub
'====================================================================
' CONDITIONAL FORMATTING SUBS
'====================================================================

Private Sub ApplyGapConditionalFormatting(ws As Worksheet, gapCol As Long, _
                                           dataStartRow As Long, dataRowCount As Long)
    Dim rng As Range, fc As FormatCondition, addr As String, firstRow As Long
    firstRow = dataStartRow + 1
    Set rng = ws.Range(ws.Cells(firstRow, gapCol), ws.Cells(dataStartRow + dataRowCount - 1, gapCol))
    addr = ws.Cells(firstRow, gapCol).Address(False, False)
    rng.FormatConditions.Delete
    
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & addr & ")," & addr & "<0)")
    fc.Interior.Color = RGB(254, 226, 226): fc.Font.Color = RGB(153, 27, 27): fc.StopIfTrue = True

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & addr & ")," & addr & ">=0," & addr & "<=4)")
    fc.Interior.Color = RGB(209, 250, 229): fc.Font.Color = RGB(6, 95, 70): fc.StopIfTrue = True

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & addr & ")," & addr & ">4)")
    fc.Interior.Color = RGB(254, 243, 199): fc.Font.Color = RGB(146, 64, 14): fc.StopIfTrue = True
End Sub

Private Sub ApplyDurationConditionalFormatting(ws As Worksheet, actualCol As Long, stdCol As Long, _
                                                dataStartRow As Long, dataRowCount As Long)
    Dim rng As Range, fc As FormatCondition
    Dim aAddr As String, sAddr As String, firstRow As Long
    firstRow = dataStartRow + 1
    Set rng = ws.Range(ws.Cells(firstRow, actualCol), ws.Cells(dataStartRow + dataRowCount - 1, actualCol))
    aAddr = ws.Cells(firstRow, actualCol).Address(False, False)
    sAddr = ws.Cells(firstRow, stdCol).Address(False, False)
    rng.FormatConditions.Delete
    
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & aAddr & "),ISNUMBER(" & sAddr & ")," & aAddr & "<" & sAddr & ")")
    fc.Interior.Color = RGB(254, 226, 226): fc.Font.Color = RGB(153, 27, 27): fc.StopIfTrue = True

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & aAddr & "),ISNUMBER(" & sAddr & ")," & aAddr & ">" & sAddr & "," & aAddr & "-" & sAddr & "<=4)")
    fc.Interior.Color = RGB(209, 250, 229): fc.Font.Color = RGB(6, 95, 70): fc.StopIfTrue = True

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & aAddr & "),ISNUMBER(" & sAddr & ")," & aAddr & "-" & sAddr & ">4)")
    fc.Interior.Color = RGB(254, 243, 199): fc.Font.Color = RGB(146, 64, 14): fc.StopIfTrue = True
End Sub

Private Sub ApplyCycleTimeConditionalFormatting(ws As Worksheet, cycleTimeCol As Long, _
                                                 dataStartRow As Long, dataRowCount As Long)
    Dim rng As Range, fc As FormatCondition, addr As String, firstRow As Long
    firstRow = dataStartRow + 1
    Set rng = ws.Range(ws.Cells(firstRow, cycleTimeCol), ws.Cells(dataStartRow + dataRowCount - 1, cycleTimeCol))
    addr = ws.Cells(firstRow, cycleTimeCol).Address(False, False)
    rng.FormatConditions.Delete
    
    ' Green: <= threshold (from S1) - SA9 light
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & addr & ")," & addr & "<=$S$1)")
    fc.Interior.Color = RGB(209, 250, 229): fc.Font.Color = RGB(6, 95, 70): fc.StopIfTrue = True

    ' Red: > threshold - SA9 light
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & addr & ")," & addr & ">$S$1)")
    fc.Interior.Color = RGB(254, 226, 226): fc.Font.Color = RGB(153, 27, 27): fc.StopIfTrue = True
End Sub

Private Sub ApplyPreInstallConditionalFormatting(ws As Worksheet, preInstallCol As Long, _
                                                  newReusedCol As Long, dataStartRow As Long, dataRowCount As Long)
    Dim rng As Range, fc As FormatCondition
    Dim dAddr As String, nrAddr As String, firstRow As Long
    firstRow = dataStartRow + 1
    Set rng = ws.Range(ws.Cells(firstRow, preInstallCol), ws.Cells(dataStartRow + dataRowCount - 1, preInstallCol))
    dAddr = ws.Cells(firstRow, preInstallCol).Address(False, False)
    nrAddr = ws.Cells(firstRow, newReusedCol).Address(False, False)
    rng.FormatConditions.Delete
    
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & dAddr & ")," & dAddr & ">=TODAY()," & dAddr & "<=TODAY()+7)")
    fc.Font.Color = RGB(0, 0, 255): fc.Font.Bold = True: fc.StopIfTrue = True

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & dAddr & ")," & dAddr & ">=TODAY()," & dAddr & "<=TODAY()+30)")
    fc.Font.Color = RGB(128, 0, 128): fc.StopIfTrue = True

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & dAddr & ")," & dAddr & "<TODAY())")
    fc.Interior.Color = RGB(240, 240, 242): fc.Font.Color = RGB(120, 120, 130): fc.StopIfTrue = True

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(" & nrAddr & "=""New"",NOT(ISNUMBER(" & dAddr & ")))")
    fc.Font.Color = CLR_CT_MISS: fc.Font.Bold = True: fc.StopIfTrue = True
End Sub

Private Sub ApplyEstCARConditionalFormatting(ws As Worksheet, estCARCol As Long, _
                                              dataStartRow As Long, dataRowCount As Long)
    Dim rng As Range, fc As FormatCondition, addr As String, firstRow As Long
    firstRow = dataStartRow + 1
    Set rng = ws.Range(ws.Cells(firstRow, estCARCol), ws.Cells(dataStartRow + dataRowCount - 1, estCARCol))
    addr = ws.Cells(firstRow, estCARCol).Address(False, False)
    rng.FormatConditions.Delete
    
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & addr & ")," & addr & ">=TODAY()," & addr & "<=TODAY()+7)")
    fc.Font.Color = RGB(0, 0, 255): fc.Font.Bold = True: fc.StopIfTrue = True

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & addr & ")," & addr & ">=TODAY()," & addr & "<=TODAY()+30)")
    fc.Font.Color = RGB(128, 0, 128): fc.StopIfTrue = True

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & addr & ")," & addr & "<TODAY())")
    fc.Interior.Color = RGB(240, 240, 242): fc.Font.Color = RGB(120, 120, 130): fc.StopIfTrue = True
End Sub

'====================================================================
' APPLY DATE FORMATTING
'====================================================================

Private Sub ApplyDateFormatting(ws As Worksheet, sirfisHeaders As Range, outputColMap As Object, _
                                 dataStartRow As Long, dataRowCount As Long)
    Dim header As Range, colIdx As Long, dateColRange As Range
    Dim dateCFRange As Range  ' Union of all date columns needing CF
    
    For Each header In sirfisHeaders
        If outputColMap.exists(LCase(header.Value)) Then
            colIdx = outputColMap(LCase(header.Value))
            If IsDateHeader(header.Value) Then
                ws.Columns(colIdx).NumberFormat = "mm/dd/yyyy"
                If LCase(header.Value) Like "*start*" Or LCase(header.Value) Like "*finish*" Then
                    Set dateColRange = ws.Range(ws.Cells(dataStartRow + 1, colIdx), _
                                                ws.Cells(dataStartRow + dataRowCount - 1, colIdx))
                    If dateCFRange Is Nothing Then
                        Set dateCFRange = dateColRange
                    Else
                        Set dateCFRange = Union(dateCFRange, dateColRange)
                    End If
                End If
            End If
        End If
    Next header
    
    ' Apply CF once to the entire union range (3 rules instead of 3 x N columns)
    If Not dateCFRange Is Nothing Then
        ApplyDateConditionalFormatting dateCFRange
    End If
End Sub

Private Sub ApplyDateConditionalFormatting(rng As Range)
    Dim fc As FormatCondition, addr As String
    addr = rng.Cells(1, 1).Address(False, False)
    rng.FormatConditions.Delete
    
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & addr & ")," & addr & ">=TODAY()," & addr & "<=TODAY()+7)")
    fc.Font.Color = RGB(0, 0, 255): fc.StopIfTrue = True

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & addr & ")," & addr & ">=TODAY()," & addr & "<=TODAY()+30)")
    fc.Font.Color = RGB(128, 0, 128): fc.StopIfTrue = True

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & addr & ")," & addr & "<TODAY())")
    fc.Interior.Color = RGB(240, 240, 242): fc.Font.Color = RGB(120, 120, 130): fc.StopIfTrue = True
End Sub

'====================================================================
' APPLY GATING CONDITIONAL FORMATTING (fixed relative addresses)
'====================================================================

Private Sub ApplyGatingConditionalFormatting(ws As Worksheet, gatingDict As Object, _
                                              sirfisHeaders As Range, outputColMap As Object, dataStartRow As Long)
    Dim gatingGroups As Object, group As Object
    Dim headerName As Variant, gatingValue As String
    Dim gatingLetter As String, gatingNumber As Long
    Dim groupKey As Variant, groupKeys As Variant, groupIdx As Long
    Dim gatingColIndex As Long, gatedColIndex As Long
    Dim gatingColLetter As String, gatedColLetter As String
    Dim fc As FormatCondition, dateColRange As Range
    Dim formula As String, fallbackFormula As String
    Dim gatingHeaderName As String, gatedHeaderName As String
    Dim gatedPriority As Long, checkPriority As Long
    Dim checkHeaderName As String, checkColLetter As String
    Dim lastDataRow As Long, firstDataRow As Long
    
    lastDataRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    firstDataRow = dataStartRow + 1
    
    Set gatingGroups = CreateObject("Scripting.Dictionary")
    
    For Each headerName In gatingDict.Keys
        gatingValue = gatingDict(headerName)
        gatingLetter = Left(gatingValue, 1)
        If Len(gatingValue) > 1 And IsNumeric(Mid(gatingValue, 2)) Then
            gatingNumber = CInt(Mid(gatingValue, 2))
            groupKey = gatingLetter
            If Not gatingGroups.exists(groupKey) Then
                Set gatingGroups(groupKey) = CreateObject("Scripting.Dictionary")
            End If
            gatingGroups(groupKey)(gatingNumber) = headerName
        End If
    Next headerName
    
    groupKeys = gatingGroups.Keys
    
    For groupIdx = LBound(groupKeys) To UBound(groupKeys)
        groupKey = groupKeys(groupIdx)
        Set group = gatingGroups(groupKey)
        
        If group.exists(1) Then
            gatingHeaderName = group(1)
            gatingColIndex = 0
            If outputColMap.exists(LCase(gatingHeaderName)) Then
                gatingColIndex = outputColMap(LCase(gatingHeaderName))
                gatingColLetter = ColLetter(gatingColIndex)
            End If
            If gatingColIndex = 0 Then GoTo NextGroup
            
            For gatedPriority = 2 To 10
                If group.exists(gatedPriority) Then
                    gatedHeaderName = group(gatedPriority)
                    gatedColIndex = 0
                    If outputColMap.exists(LCase(gatedHeaderName)) Then
                        gatedColIndex = outputColMap(LCase(gatedHeaderName))
                        gatedColLetter = ColLetter(gatedColIndex)
                    End If
                    If gatedColIndex = 0 Then GoTo NextGatedPriority
                    
                    ' Use relative addresses (no $ prefix on row)
                    If gatedPriority = 2 Then
                        formula = "=AND(ISNUMBER(" & gatingColLetter & firstDataRow & ")," & _
                                  "ISNUMBER(" & gatedColLetter & firstDataRow & ")," & _
                                  gatingColLetter & firstDataRow & ">" & gatedColLetter & firstDataRow & ")"
                    Else
                        fallbackFormula = "=AND(ISNUMBER(" & gatingColLetter & firstDataRow & ")"
                        For checkPriority = 2 To gatedPriority - 1
                            If group.exists(checkPriority) Then
                                checkHeaderName = group(checkPriority)
                                If outputColMap.exists(LCase(checkHeaderName)) Then
                                    checkColLetter = ColLetter(outputColMap(LCase(checkHeaderName)))
                                    fallbackFormula = fallbackFormula & ",NOT(ISNUMBER(" & checkColLetter & firstDataRow & "))"
                                End If
                            End If
                        Next checkPriority
                        fallbackFormula = fallbackFormula & ",ISNUMBER(" & gatedColLetter & firstDataRow & ")," & _
                                          gatingColLetter & firstDataRow & ">" & gatedColLetter & firstDataRow & ")"
                        formula = fallbackFormula
                    End If
                    
                    Set dateColRange = ws.Range(ws.Cells(firstDataRow, gatingColIndex), ws.Cells(lastDataRow, gatingColIndex))
                    Set fc = dateColRange.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
                    fc.Font.Color = RGB(220, 53, 69)
                    fc.priority = 1
                    Exit For
                End If
NextGatedPriority:
            Next gatedPriority
        End If
NextGroup:
    Next groupIdx
End Sub

'====================================================================
' CENTER CALCULATED COLUMNS
'====================================================================

Private Sub CenterCalculatedColumns(ws As Worksheet, dataStartRow As Long, dataRowCount As Long, _
                                      actualStartCol As Long, totalColCount As Long, baseColCount As Long)
    Dim lastDataRow As Long
    lastDataRow = dataStartRow + dataRowCount - 1
    
    ' Center all milestone + additional columns
    If actualStartCol > 0 And actualStartCol <= totalColCount Then
        ws.Range(ws.Cells(dataStartRow, actualStartCol), _
                 ws.Cells(lastDataRow, totalColCount)).HorizontalAlignment = xlCenter
    End If
    
    ' Also center the Group column and any short-value base columns
    ws.Range(ws.Cells(dataStartRow, 1), ws.Cells(lastDataRow, 1)).HorizontalAlignment = xlCenter  ' Row nums if any
End Sub

'====================================================================
' ULTRA-POLISHED FORMATTING
'====================================================================

Private Sub ApplyWorkingSheetFormatting(ws As Worksheet, dataStartRow As Long, dataRowCount As Long, _
                                         totalColCount As Long, actualStartCol As Long, stdStartCol As Long, _
                                         milestoneCount As Long, additionalColStart As Long, _
                                         stdEndCol As Long, gapStartCol As Long, gapEndCol As Long, _
                                         Optional outputColMap As Object = Nothing)
    Dim headerRange As Range, dataRange As Range
    Dim i As Long, j As Long
    Dim lastDataRow As Long
    
    lastDataRow = dataStartRow + dataRowCount - 1
    
    ' === TITLE BAR (Row 1) and SUBTITLE BAR (Row 2) ===
    Dim lastVisibleCol As Long
    lastVisibleCol = ws.Cells(dataStartRow, ws.Columns.Count).End(xlToLeft).Column
    If lastVisibleCol < totalColCount Then lastVisibleCol = totalColCount
    TISCommon.ApplyTitleBar ws, lastVisibleCol, "TIS Commitment Tracker"
    TISCommon.ApplySubtitleBar ws, lastVisibleCol, _
        "Single source of truth  |  TIS columns are auto-updated  |  User columns (green headers) are yours  |  " & TIS_VERSION

    ' === HEADER ROW - Base formatting ===
    Dim headerRow As Long: headerRow = dataStartRow
    Set headerRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, totalColCount))
    With headerRange
        .Font.Bold = True
        .Font.Size = 9
        .Font.name = "Segoe UI"
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(44, 62, 80)  ' Default fallback
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 48
    End With

    ' === ZONE-COLORED HEADERS ===
    ' Zone boundary variables (used also for category bar and grouping)
    Dim ourStartC As Long: ourStartC = 0
    Dim ourEndC As Long: ourEndC = 0
    Dim tisStartC As Long: tisStartC = 0
    Dim tisEndC As Long: tisEndC = 0
    Dim nrCol As Long: nrCol = 0
    Dim lastDataC As Long: lastDataC = 0

    If Not outputColMap Is Nothing Then
        ' Detect zone boundaries using outputColMap keys
        ' Physical column order: [Identity] [TIS Dates] [Our Dates] [Analysis] [User Fields]

        ' --- Find Our Dates zone (Set through WhatIf) ---
        If outputColMap.exists(LCase(TIS_COL_OUR_SET)) Then ourStartC = outputColMap(LCase(TIS_COL_OUR_SET))
        If ourStartC = 0 Then ourStartC = 7  ' fallback

        If outputColMap.exists(LCase(TIS_COL_WHATIF)) Then ourEndC = outputColMap(LCase(TIS_COL_WHATIF))
        If ourEndC = 0 And outputColMap.exists(LCase(TIS_COL_HEALTH)) Then ourEndC = outputColMap(LCase(TIS_COL_HEALTH))
        If ourEndC = 0 Then ourEndC = ourStartC + 11  ' fallback

        ' --- Find Identity zone end (last of Site/Entity Code/Entity Type/CEID/Group/Event Type) ---
        Dim identityEndC As Long: identityEndC = 0
        Dim idKeys As Variant, idIdx As Long
        idKeys = Array("site", LCase(HEADER_ENTITY_CODE), "entity type", "ceid", "group", LCase(HEADER_EVENT_TYPE))
        For idIdx = LBound(idKeys) To UBound(idKeys)
            If outputColMap.exists(CStr(idKeys(idIdx))) Then
                If outputColMap(CStr(idKeys(idIdx))) > identityEndC Then identityEndC = outputColMap(CStr(idKeys(idIdx)))
            End If
        Next idIdx
        If identityEndC = 0 Then identityEndC = ourStartC - 1  ' fallback

        ' --- TIS Dates zone: between identity end and Our Dates start ---
        tisStartC = identityEndC + 1
        tisEndC = ourStartC - 1

        ' --- Find New/Reused column as Analysis/User boundary ---
        If outputColMap.exists("new/reused") Then nrCol = outputColMap("new/reused")
        If nrCol = 0 Then
            ' Fallback: scan keys for partial match
            Dim mapKey As Variant
            For Each mapKey In outputColMap.keys
                If InStr(1, Replace(CStr(mapKey), vbLf, ""), "new/reused", vbTextCompare) > 0 Then
                    nrCol = outputColMap(mapKey)
                    Exit For
                End If
            Next mapKey
        End If

        ' --- Find last data column (BOD2 or BOD1) ---
        If outputColMap.exists("bod2") Then
            lastDataC = outputColMap("bod2")
        ElseIf outputColMap.exists("bod1") Then
            lastDataC = outputColMap("bod1")
        Else
            lastDataC = totalColCount  ' fallback
        End If

        ' === APPLY ZONE HEADER COLORS ===
        ' 1. Identity zone
        If identityEndC >= 1 Then
            TISCommon.ApplyZoneHeader ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, identityEndC)), ZONE_IDENTITY_BG, ZONE_IDENTITY_FG
        End If

        ' 2. TIS Dates zone (between identity and Our Dates)
        If tisEndC >= tisStartC Then
            TISCommon.ApplyZoneHeader ws.Range(ws.Cells(headerRow, tisStartC), ws.Cells(headerRow, tisEndC)), ZONE_TIS_BG, ZONE_TIS_FG
        End If

        ' 3. Our Dates zone (Set through WhatIf)
        TISCommon.ApplyZoneHeader ws.Range(ws.Cells(headerRow, ourStartC), ws.Cells(headerRow, ourEndC)), ZONE_USER_BG, ZONE_USER_FG

        ' 4. Analysis zone (after Our Dates through before New/Reused)
        Dim analysisStartC As Long: analysisStartC = ourEndC + 1
        Dim analysisEndC As Long
        If nrCol > 0 Then
            analysisEndC = nrCol - 1
        Else
            analysisEndC = analysisStartC + 10  ' fallback
        End If
        If analysisEndC >= analysisStartC Then
            TISCommon.ApplyZoneHeader ws.Range(ws.Cells(headerRow, analysisStartC), ws.Cells(headerRow, analysisEndC)), ZONE_CALC_BG, ZONE_CALC_FG
        End If

        ' 5. User Fields zone (New/Reused through BOD2)
        If nrCol > 0 And lastDataC >= nrCol Then
            TISCommon.ApplyZoneHeader ws.Range(ws.Cells(headerRow, nrCol), ws.Cells(headerRow, lastDataC)), ZONE_USER_BG, ZONE_USER_FG
        End If

        ' === ZONE CATEGORY BAR (Row above headers) ===
        Dim catRow As Long: catRow = headerRow - 1
        ws.Rows(catRow).RowHeight = 16

        If identityEndC >= 1 Then
            TISCommon.ApplyZoneCategoryLabel ws, catRow, 1, identityEndC, "IDENTITY", ZONE_IDENTITY_BG
        End If
        If tisEndC >= tisStartC Then
            TISCommon.ApplyZoneCategoryLabel ws, catRow, tisStartC, tisEndC, "TIS DATES (auto-updated)", ZONE_TIS_BG
        End If
        TISCommon.ApplyZoneCategoryLabel ws, catRow, ourStartC, ourEndC, "OUR COMMITMENT DATES (editable)", ZONE_USER_BG
        If analysisEndC >= analysisStartC Then
            TISCommon.ApplyZoneCategoryLabel ws, catRow, analysisStartC, analysisEndC, "MILESTONE ANALYSIS", ZONE_CALC_BG
        End If
        If nrCol > 0 And lastDataC >= nrCol Then
            TISCommon.ApplyZoneCategoryLabel ws, catRow, nrCol, lastDataC, "USER FIELDS", ZONE_USER_BG
        End If
    End If

    ' Header separators
    For j = 1 To totalColCount
        With ws.Cells(dataStartRow, j).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .Color = RGB(80, 100, 120)
        End With
    Next j

    ' === DATA ROWS - Light zebra stripes (SA9) ===
    For i = dataStartRow + 1 To lastDataRow
        If (i - dataStartRow) Mod 2 = 0 Then
            ws.Range(ws.Cells(i, 1), ws.Cells(i, totalColCount)).Interior.Color = RGB(248, 249, 252)
        Else
            ws.Range(ws.Cells(i, 1), ws.Cells(i, totalColCount)).Interior.Color = RGB(255, 255, 255)
        End If
    Next i

    ' === SECTION HEADERS - SA9 accent colors ===
    If milestoneCount > 0 And actualStartCol > 0 Then
        ' Actual Duration - Soft blue
        ws.Range(ws.Cells(dataStartRow, actualStartCol), _
                 ws.Cells(dataStartRow, actualStartCol + milestoneCount)).Interior.Color = RGB(59, 130, 246)

        ' STD Duration - Soft green
        If stdStartCol > 0 And stdEndCol >= stdStartCol Then
            ws.Range(ws.Cells(dataStartRow, stdStartCol), _
                     ws.Cells(dataStartRow, stdEndCol)).Interior.Color = RGB(34, 197, 94)
        End If

        ' Gap - Soft violet
        If gapStartCol > 0 And gapEndCol >= gapStartCol Then
            ws.Range(ws.Cells(dataStartRow, gapStartCol), _
                     ws.Cells(dataStartRow, gapEndCol)).Interior.Color = RGB(139, 92, 246)
        End If
    End If

    ' Additional columns - Soft teal
    If additionalColStart > 0 And additionalColStart <= totalColCount Then
        ws.Range(ws.Cells(dataStartRow, additionalColStart), _
                 ws.Cells(dataStartRow, totalColCount)).Interior.Color = RGB(20, 184, 166)
    End If

    ' === TABLE BORDERS - Light clean lines (SA9) ===
    Set dataRange = ws.Range(ws.Cells(dataStartRow, 1), ws.Cells(lastDataRow, totalColCount))

    ' Subtle inner grid
    With dataRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = RGB(230, 232, 240)
    End With
    With dataRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = RGB(230, 232, 240)
    End With

    ' Clean outer border
    With dataRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(44, 62, 80)
    End With
    With dataRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(44, 62, 80)
    End With
    With dataRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(44, 62, 80)
    End With
    With dataRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(44, 62, 80)
    End With
    
    ' === SECTION DIVIDER LINES (slightly heavier between sections) ===
    If actualStartCol > 1 Then
        ApplySectionDivider ws, dataStartRow, lastDataRow, actualStartCol
    End If
    If stdStartCol > 0 Then
        ApplySectionDivider ws, dataStartRow, lastDataRow, stdStartCol
    End If
    If gapStartCol > 0 Then
        ApplySectionDivider ws, dataStartRow, lastDataRow, gapStartCol
    End If
    If additionalColStart > 0 Then
        ApplySectionDivider ws, dataStartRow, lastDataRow, additionalColStart
    End If
    
    ' Auto-fit data columns only (not entire sheet — expensive at scale)
    Dim afLastC As Long
    afLastC = ws.Cells(dataStartRow, ws.Columns.Count).End(xlToLeft).Column
    If afLastC > 0 Then ws.Range(ws.Columns(1), ws.Columns(afLastC)).AutoFit
    
    If milestoneCount > 0 And actualStartCol > 0 Then
        For j = actualStartCol To actualStartCol + milestoneCount
            ws.Columns(j).ColumnWidth = 9
        Next j
        If stdStartCol > 0 And stdEndCol >= stdStartCol Then
            For j = stdStartCol To stdEndCol
                ws.Columns(j).ColumnWidth = 9
            Next j
        End If
        If gapStartCol > 0 And gapEndCol >= gapStartCol Then
            For j = gapStartCol To gapEndCol
                ws.Columns(j).ColumnWidth = 8
            Next j
        End If
    End If
    
    ' Set default font for data area
    ws.Range(ws.Cells(dataStartRow + 1, 1), ws.Cells(lastDataRow, totalColCount)).Font.name = "Segoe UI"
    ws.Range(ws.Cells(dataStartRow + 1, 1), ws.Cells(lastDataRow, totalColCount)).Font.Size = 10

    ' === ADDITIONAL COLUMN GROUPING (TIS dates and user fields) ===
    If Not outputColMap Is Nothing Then
        ' Group TIS date columns (collapsible)
        If tisStartC > 0 And tisEndC > tisStartC Then
            On Error Resume Next
            ws.Columns(ColLetter(tisStartC) & ":" & ColLetter(tisEndC)).Group
            ws.Outline.ShowLevels ColumnLevels:=1  ' Start collapsed
            On Error GoTo 0
        End If

        ' Group user fields (SOC through BOD2)
        Dim socCol As Long: socCol = 0
        If outputColMap.exists("soc available") Then socCol = outputColMap("soc available")
        If socCol > 0 And lastDataC > socCol Then
            On Error Resume Next
            ws.Columns(ColLetter(socCol) & ":" & ColLetter(lastDataC)).Group
            On Error GoTo 0
        End If
    End If
End Sub

'====================================================================
' APPLY SECTION DIVIDER LINE
'====================================================================

Private Sub ApplySectionDivider(ws As Worksheet, startRow As Long, endRow As Long, col As Long)
    With ws.Range(ws.Cells(startRow, col), ws.Cells(endRow, col)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(44, 62, 80)
    End With
End Sub

'====================================================================
' SORTING WITH HELPER COLUMNS
' Delegates to TISCommon.SortWithHelperColumns for the actual sort logic.
'====================================================================

Private Sub ApplySortingWithHelperColumns(ws As Worksheet, sortDict As Object, _
                                           outputColMap As Object, dataStartRow As Long)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    SortWithHelperColumns ws, sortDict, outputColMap, dataStartRow, lastRow
End Sub

'====================================================================
' SORT WORKING SHEET BY STATUS + PROJECT START DATE
' Sorts: Active first (by earliest Our Date), Cancelled last.
' Uses temporary helper columns that are deleted after sort.
'====================================================================

Private Sub SortWorkingSheetByStatus(ws As Worksheet, outputColMap As Object, _
                                      dataStartRow As Long)
    Dim wsLastRow As Long
    wsLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If wsLastRow <= dataStartRow Then Exit Sub  ' Need at least 2 data rows

    Dim wsMaxCol As Long
    wsMaxCol = ws.Cells(dataStartRow, ws.Columns.Count).End(xlToLeft).Column

    ' Find Status column from outputColMap
    Dim statusCol As Long
    statusCol = 0
    If outputColMap.exists(LCase(TIS_COL_STATUS)) Then statusCol = outputColMap(LCase(TIS_COL_STATUS))
    If statusCol = 0 Then Exit Sub

    ' Find Our Date columns for MIN calculation
    Dim ourDateHeaders As Variant
    ourDateHeaders = Array(TIS_COL_OUR_SET, TIS_COL_OUR_SL1, TIS_COL_OUR_SL2, _
                           TIS_COL_OUR_SQ, TIS_COL_OUR_CONVS, TIS_COL_OUR_CONVF, _
                           TIS_COL_OUR_MRCLS, TIS_COL_OUR_MRCLF)

    ' Build MIN formula refs
    Dim minParts As String
    minParts = ""
    Dim oi As Long
    For oi = LBound(ourDateHeaders) To UBound(ourDateHeaders)
        Dim ourCol As Long
        ourCol = 0
        If outputColMap.exists(LCase(CStr(ourDateHeaders(oi)))) Then
            ourCol = outputColMap(LCase(CStr(ourDateHeaders(oi))))
        End If
        If ourCol > 0 Then
            If minParts <> "" Then minParts = minParts & ","
            minParts = minParts & TISCommon.ColLetter(ourCol) & (dataStartRow + 1)
        End If
    Next oi

    If minParts = "" Then Exit Sub

    ' Add temporary helper column for project start date
    Dim helperCol As Long
    helperCol = wsMaxCol + 1
    ws.Cells(dataStartRow, helperCol).Value = "_SortHelper"

    ' Write MIN formula to first data row
    Dim firstDataRow As Long
    firstDataRow = dataStartRow + 1
    ws.Cells(firstDataRow, helperCol).Formula = "=MIN(" & minParts & ")"

    ' FillDown to all data rows
    If wsLastRow > firstDataRow Then
        ws.Cells(firstDataRow, helperCol).Copy
        ws.Range(ws.Cells(firstDataRow + 1, helperCol), ws.Cells(wsLastRow, helperCol)).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False
    End If

    ' Convert formulas to values for sorting
    Dim helperRange As Range
    Set helperRange = ws.Range(ws.Cells(firstDataRow, helperCol), ws.Cells(wsLastRow, helperCol))
    helperRange.Value = helperRange.Value

    ' Add temporary status sort helper (numeric: Active=1, Completed=2, On Hold=3, Non IQ=4, Cancelled=5)
    Dim statusHelperCol As Long
    statusHelperCol = helperCol + 1
    ws.Cells(dataStartRow, statusHelperCol).Value = "_StatusSort"

    Dim sr As Long
    Dim statusArr As Variant
    statusArr = ws.Range(ws.Cells(firstDataRow, statusCol), ws.Cells(wsLastRow, statusCol)).Value
    Dim statusSortArr() As Variant
    ReDim statusSortArr(1 To wsLastRow - dataStartRow, 1 To 1)
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
        .SetRange ws.Range(ws.Cells(dataStartRow, 1), ws.Cells(wsLastRow, statusHelperCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' Delete helper columns
    ws.Range(ws.Columns(helperCol), ws.Columns(statusHelperCol)).Delete
End Sub

'====================================================================
' ROW PASSES FILTERS (TISLoader parity - OR within same column)
'====================================================================

Private Function RowPassesFilters(tisData As Variant, rowIndex As Long, _
                                   headerDict As Object, filterDict As Object, _
                                   pastDateFilterDict As Object, _
                                   excludeSystemsList As Collection, _
                                   siteColIdx As Long, entityCodeColIdx As Long, eventTypeColIdx As Long) As Boolean
    Dim hKey As Variant, colIndex As Long
    Dim filterValues As Variant
    Dim cellValue As String
    Dim matchFound As Boolean, j As Long
    Dim excludeEntry As Variant
    Dim rowSite As String, rowEntityCode As String, rowEventType As String
    
    RowPassesFilters = True
    
    ' Check exclude systems list first
    If Not excludeSystemsList Is Nothing Then
        If excludeSystemsList.Count > 0 Then
            ' Get row values for comparison (case-insensitive)
            If siteColIdx > 0 Then rowSite = LCase(Trim(CStr(tisData(rowIndex, siteColIdx)))) Else rowSite = ""
            If entityCodeColIdx > 0 Then rowEntityCode = LCase(Trim(CStr(tisData(rowIndex, entityCodeColIdx)))) Else rowEntityCode = ""
            If eventTypeColIdx > 0 Then rowEventType = LCase(Trim(CStr(tisData(rowIndex, eventTypeColIdx)))) Else rowEventType = ""
            
            For Each excludeEntry In excludeSystemsList
                ' excludeEntry is an array: (0)=site, (1)=entityCode, (2)=eventType
                If rowSite = excludeEntry(0) And _
                   rowEntityCode = excludeEntry(1) And _
                   rowEventType = excludeEntry(2) Then
                    RowPassesFilters = False
                    Exit Function
                End If
            Next excludeEntry
        End If
    End If
    
    ' Check regular filters (AND between columns, OR within same column)
    For Each hKey In filterDict.Keys
        If headerDict.exists(hKey) Then
            colIndex = headerDict(hKey)
            cellValue = LCase(Trim(CStr(tisData(rowIndex, colIndex))))
            filterValues = filterDict(hKey)
            
            matchFound = False
            For j = LBound(filterValues) To UBound(filterValues)
                If cellValue = LCase(Trim(CStr(filterValues(j)))) Then
                    matchFound = True
                    Exit For
                End If
            Next j
            
            If Not matchFound Then
                RowPassesFilters = False
                Exit Function
            End If
        End If
    Next hKey
    
    ' Check past date filters
    For Each hKey In pastDateFilterDict.Keys
        If headerDict.exists(hKey) Then
            colIndex = headerDict(hKey)
            If IsDate(tisData(rowIndex, colIndex)) Then
                If CDate(tisData(rowIndex, colIndex)) < Date Then
                    RowPassesFilters = False
                    Exit Function
                End If
            End If
        End If
    Next hKey
End Function

'====================================================================
' APPLY COLUMN GROUPING (from Definitions Grouped column)
'====================================================================

Private Sub ApplyColumnGrouping(ws As Worksheet, groupedDict As Object, outputColMap As Object)
    Dim groupNumbers As Object
    Dim headerName As Variant, groupNum As Long
    Dim outputCol As Long
    Dim gKey As Variant, gCols As Object, minCol As Long, maxCol As Long
    Dim c As Variant
    
    ' Collect all unique group numbers and their columns
    Set groupNumbers = CreateObject("Scripting.Dictionary")
    
    Dim grpMapKey As String
    For Each headerName In groupedDict.Keys
        groupNum = groupedDict(headerName)
        grpMapKey = LCase(CStr(headerName))
        If outputColMap.exists(grpMapKey) Then
            outputCol = outputColMap(grpMapKey)
            If Not groupNumbers.exists(groupNum) Then
                Set groupNumbers(groupNum) = CreateObject("Scripting.Dictionary")
            End If
            groupNumbers(groupNum)(outputCol) = True
        End If
    Next headerName
    
    ' For each group number, find min and max column and group them
    For Each gKey In groupNumbers.Keys
        Set gCols = groupNumbers(gKey)
        minCol = 99999
        maxCol = 0
        For Each c In gCols.Keys
            If CLng(c) < minCol Then minCol = CLng(c)
            If CLng(c) > maxCol Then maxCol = CLng(c)
        Next c
        
        If minCol < maxCol Then
            GroupColumnsCollapsed ws, minCol, maxCol
        End If
    Next gKey
End Sub
'====================================================================
' GROUP COLUMNS AND COLLAPSE
'====================================================================

Private Sub GroupColumnsCollapsed(ws As Worksheet, startCol As Long, endCol As Long)
    On Error Resume Next
    ws.Columns(startCol).Resize(1, endCol - startCol + 1).group
    ws.Outline.ShowLevels ColumnLevels:=1   ' Collapse to level 1
    On Error GoTo 0
End Sub

'====================================================================
' INSERT GROUP COLUMN FORMULA
'====================================================================

Private Sub InsertGroupColumnFormula(ws As Worksheet, groupCol As Long, entityTypeCol As Long, _
                                      dataStartRow As Long, dataRowCount As Long)
    Dim firstDataRow As Long, lastDataRow As Long
    Dim entityTypeLetter As String
    Dim formula As String
    
    firstDataRow = dataStartRow + 1
    lastDataRow = dataStartRow + dataRowCount - 1
    entityTypeLetter = ColLetter(entityTypeCol)
    
    ' VLOOKUP formula to CEIDs sheet
    formula = "=IFERROR(VLOOKUP(" & entityTypeLetter & firstDataRow & ",CEIDs!A:B,2,FALSE),"""")"
    
    ' Insert formula in first data row
    ws.Cells(firstDataRow, groupCol).formula = formula
    
    ' Fill down to all data rows
    If lastDataRow > firstDataRow Then
        ws.Range(ws.Cells(firstDataRow, groupCol), ws.Cells(lastDataRow, groupCol)).FillDown
    End If
End Sub
'====================================================================
' FIND COMMITTED WORKING SHEET (exact name only)
' The committed copy is always "Working Sheet" (no suffix).
' Versioned sheets ("Working Sheet 2") are drafts pending review.
'====================================================================

Private Function FindLatestWorkingSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_WORKING)
    On Error GoTo 0
    Set FindLatestWorkingSheet = ws
End Function

'====================================================================
' IMPORT FROM OLD SHEET - MAIN IMPORT ORCHESTRATOR
' Matches projects by Site|Entity Code|Event Type composite key.
' Imports user-entered data, marks new/edited projects, archives removed.
'====================================================================


'====================================================================
' IMPORT USER DATA FROM OLD SHEET (simplified — no Gantt/NIF import)
' Only imports user-entered columns and marks changes.
' Gantt and NIF are regenerated fresh by calling their builders.
'====================================================================

Private Sub ImportUserDataFromOldSheet(oldSheet As Worksheet, newSheet As Worksheet, _
                                outputColMap As Object, sirfisHeaders As Range, _
                                headerDict As Object, _
                                dataStartRow As Long, dataRowCount As Long, totalColCount As Long)
    Dim oldHeaderRow As Long, oldLastRow As Long, oldLastCol As Long
    Dim newLastRow As Long
    Dim i As Long, c As Long
    Dim oldHeaderMap As Object, oldKeyMap As Object, newKeyMap As Object
    Dim projectKey As String
    Dim oldSiteCol As Long, oldEntityCodeCol As Long, oldEventTypeCol As Long
    Dim newSiteCol As Long, newEntityCodeCol As Long, newEventTypeCol As Long
    Dim cellVal As String
    
    oldHeaderRow = TISCommon.FindHeaderRow(oldSheet)
    If oldHeaderRow = 0 Then Exit Sub
    
    oldLastRow = oldSheet.Cells(oldSheet.Rows.Count, 1).End(xlUp).Row
    oldLastCol = oldSheet.Cells(oldHeaderRow, oldSheet.Columns.Count).End(xlToLeft).Column
    newLastRow = dataStartRow + dataRowCount - 1
    
    ' Build old sheet header map
    Set oldHeaderMap = CreateObject("Scripting.Dictionary")
    For c = 1 To oldLastCol
        cellVal = LCase(Trim(Replace(Replace(CStr(oldSheet.Cells(oldHeaderRow, c).Value), vbLf, ""), vbCr, "")))
        If cellVal <> "" And Not oldHeaderMap.exists(cellVal) Then
            oldHeaderMap(cellVal) = c
        End If
    Next c
    
    ' Find key columns
    oldSiteCol = 0: oldEntityCodeCol = 0: oldEventTypeCol = 0
    If oldHeaderMap.exists("site") Then oldSiteCol = oldHeaderMap("site")
    If oldHeaderMap.exists(LCase(HEADER_ENTITY_CODE)) Then oldEntityCodeCol = oldHeaderMap(LCase(HEADER_ENTITY_CODE))
    If oldHeaderMap.exists(LCase(HEADER_EVENT_TYPE)) Then oldEventTypeCol = oldHeaderMap(LCase(HEADER_EVENT_TYPE))
    If oldEntityCodeCol = 0 Then Exit Sub
    
    newSiteCol = 0: newEntityCodeCol = 0: newEventTypeCol = 0
    If outputColMap.exists("site") Then newSiteCol = outputColMap("site")
    If outputColMap.exists(LCase(HEADER_ENTITY_CODE)) Then newEntityCodeCol = outputColMap(LCase(HEADER_ENTITY_CODE))
    If outputColMap.exists(LCase(HEADER_EVENT_TYPE)) Then newEventTypeCol = outputColMap(LCase(HEADER_EVENT_TYPE))
    If newEntityCodeCol = 0 Then Exit Sub
    
    ' Build key -> row maps
    ' Performance: bulk-read key columns into arrays to avoid per-cell COM calls
    Dim oldSiteArr As Variant, oldECArr As Variant, oldETArr As Variant
    Dim oldDataRows As Long
    oldDataRows = oldLastRow - oldHeaderRow
    If oldDataRows > 0 Then
        If oldSiteCol > 0 Then oldSiteArr = oldSheet.Range(oldSheet.Cells(oldHeaderRow + 1, oldSiteCol), oldSheet.Cells(oldLastRow, oldSiteCol)).Value
        If oldEntityCodeCol > 0 Then oldECArr = oldSheet.Range(oldSheet.Cells(oldHeaderRow + 1, oldEntityCodeCol), oldSheet.Cells(oldLastRow, oldEntityCodeCol)).Value
        If oldEventTypeCol > 0 Then oldETArr = oldSheet.Range(oldSheet.Cells(oldHeaderRow + 1, oldEventTypeCol), oldSheet.Cells(oldLastRow, oldEventTypeCol)).Value
    End If

    Set oldKeyMap = CreateObject("Scripting.Dictionary")
    For i = 1 To oldDataRows
        Dim okS As String, okEC As String, okET As String
        okS = "": okEC = "": okET = ""
        If oldSiteCol > 0 Then okS = LCase(Trim(CStr(oldSiteArr(i, 1))))
        If oldEntityCodeCol > 0 Then okEC = LCase(Trim(CStr(oldECArr(i, 1))))
        If oldEventTypeCol > 0 Then okET = LCase(Trim(CStr(oldETArr(i, 1))))
        projectKey = okS & "|" & okEC & "|" & okET
        If projectKey <> "||" And Not oldKeyMap.exists(projectKey) Then
            oldKeyMap(projectKey) = oldHeaderRow + i  ' Store actual sheet row
        End If
    Next i

    Dim newSiteArr As Variant, newECArr As Variant, newETArr As Variant
    Dim newDataRows As Long
    newDataRows = newLastRow - dataStartRow
    If newDataRows > 0 Then
        If newSiteCol > 0 Then newSiteArr = newSheet.Range(newSheet.Cells(dataStartRow + 1, newSiteCol), newSheet.Cells(newLastRow, newSiteCol)).Value
        If newEntityCodeCol > 0 Then newECArr = newSheet.Range(newSheet.Cells(dataStartRow + 1, newEntityCodeCol), newSheet.Cells(newLastRow, newEntityCodeCol)).Value
        If newEventTypeCol > 0 Then newETArr = newSheet.Range(newSheet.Cells(dataStartRow + 1, newEventTypeCol), newSheet.Cells(newLastRow, newEventTypeCol)).Value
    End If

    Set newKeyMap = CreateObject("Scripting.Dictionary")
    For i = 1 To newDataRows
        Dim nkS As String, nkEC As String, nkET As String
        nkS = "": nkEC = "": nkET = ""
        If newSiteCol > 0 Then nkS = LCase(Trim(CStr(newSiteArr(i, 1))))
        If newEntityCodeCol > 0 Then nkEC = LCase(Trim(CStr(newECArr(i, 1))))
        If newEventTypeCol > 0 Then nkET = LCase(Trim(CStr(newETArr(i, 1))))
        projectKey = nkS & "|" & nkEC & "|" & nkET
        If projectKey <> "||" And Not newKeyMap.exists(projectKey) Then
            newKeyMap(projectKey) = dataStartRow + i  ' Store actual sheet row
        End If
    Next i

    ' Import user data and detect changes
    Dim userImportHeaders As Variant
    userImportHeaders = Array("escalated", "ship" & vbLf & "date", "soc" & vbLf & "available", _
                              "soc" & vbLf & "uploaded?", "staffed?", "comments", "watch", _
                              "bod1", "bod2", _
                              LCase(TIS_COL_OUR_SET), LCase(TIS_COL_OUR_SL1), _
                              LCase(TIS_COL_OUR_SL2), LCase(TIS_COL_OUR_SQ), LCase(TIS_COL_OUR_CONVS), _
                              LCase(TIS_COL_OUR_CONVF), LCase(TIS_COL_OUR_MRCLS), LCase(TIS_COL_OUR_MRCLF), _
                              LCase(TIS_COL_STATUS), LCase(TIS_COL_LOCK), _
                              LCase(TIS_COL_WHATIF))

    Dim excludeFromCompare As Object
    Set excludeFromCompare = CreateObject("Scripting.Dictionary")
    excludeFromCompare("site") = True
    excludeFromCompare(LCase(HEADER_ENTITY_CODE)) = True
    excludeFromCompare(LCase(HEADER_EVENT_TYPE)) = True
    excludeFromCompare("published") = True

    ' Performance: Bulk-read old sheet data into array for fast comparison
    Dim oldMaxCol As Long
    oldMaxCol = oldSheet.Cells(oldHeaderRow, oldSheet.Columns.Count).End(xlToLeft).Column
    Dim oldDataArr As Variant
    If oldLastRow > oldHeaderRow Then
        oldDataArr = oldSheet.Range(oldSheet.Cells(oldHeaderRow, 1), oldSheet.Cells(oldLastRow, oldMaxCol)).Value
    End If

    ' Bulk-read new sheet data into array
    Dim newMaxCol As Long
    newMaxCol = newSheet.Cells(dataStartRow, newSheet.Columns.Count).End(xlToLeft).Column
    If newMaxCol < totalColCount Then newMaxCol = totalColCount
    Dim newDataArr As Variant
    If newLastRow >= dataStartRow Then
        newDataArr = newSheet.Range(newSheet.Cells(dataStartRow, 1), newSheet.Cells(newLastRow, newMaxCol)).Value
    End If

    ' Pre-resolve import column indices ONCE (not per row)
    Dim importColOldArr() As Long, importColNewArr() As Long
    Dim numImportHeaders As Long
    numImportHeaders = UBound(userImportHeaders) - LBound(userImportHeaders) + 1
    ReDim importColOldArr(LBound(userImportHeaders) To UBound(userImportHeaders))
    ReDim importColNewArr(LBound(userImportHeaders) To UBound(userImportHeaders))
    Dim importHeader As Variant
    Dim ihi As Long
    For ihi = LBound(userImportHeaders) To UBound(userImportHeaders)
        importColOldArr(ihi) = FindHeaderColInSheet(oldSheet, oldHeaderRow, CStr(userImportHeaders(ihi)), oldLastCol)
        importColNewArr(ihi) = FindHeaderColInSheet(newSheet, dataStartRow, CStr(userImportHeaders(ihi)), totalColCount)
    Next ihi

    ' Pre-resolve Completed column for migration
    Dim oldCompCol As Long
    oldCompCol = FindHeaderColInSheet(oldSheet, oldHeaderRow, "completed", oldLastCol)
    Dim newStatusCol As Long
    newStatusCol = FindHeaderColInSheet(newSheet, dataStartRow, LCase(TIS_COL_STATUS), totalColCount)

    ' Collect changed cells for batch orange fill (Fix 2)
    Dim changedCells As New Collection  ' Each item = Array(newRow, newColIdx, oldValFormatted)

    Dim nKey As Variant, oldRow As Long, newRow As Long
    Dim header As Range
    Dim oldColIdx As Long, newColIdx As Long
    Dim oldVal As String, newVal As String
    Dim headerLower As String
    Dim importColOld As Long, importColNew As Long

    For Each nKey In newKeyMap.Keys
        newRow = newKeyMap(nKey)

        If oldKeyMap.exists(nKey) Then
            oldRow = oldKeyMap(nKey)

            ' Array index offsets: oldDataArr row 1 = oldHeaderRow, so data row offset = oldRow - oldHeaderRow + 1
            ' But row 1 in the array IS oldHeaderRow, so oldRow maps to array index (oldRow - oldHeaderRow + 1)
            Dim oldArrIdx As Long
            oldArrIdx = oldRow - oldHeaderRow + 1  ' +1 because array includes header row
            Dim newArrIdx As Long
            newArrIdx = newRow - dataStartRow + 1  ' +1 because array includes header row

            ' Change detection on TIS-sourced columns (reads from arrays, not cells)
            Dim compareMapKey As String
            For Each header In sirfisHeaders
                headerLower = LCase(header.Value)
                compareMapKey = headerLower
                If Not excludeFromCompare.exists(headerLower) Then
                    If outputColMap.exists(compareMapKey) And oldHeaderMap.exists(headerLower) Then
                        newColIdx = outputColMap(compareMapKey)
                        oldColIdx = oldHeaderMap(headerLower)
                        ' Read from arrays instead of cells
                        Dim oldCellRaw As Variant, newCellRaw As Variant
                        If oldColIdx <= oldMaxCol Then oldCellRaw = oldDataArr(oldArrIdx, oldColIdx) Else oldCellRaw = ""
                        If newColIdx <= newMaxCol Then newCellRaw = newDataArr(newArrIdx, newColIdx) Else newCellRaw = ""
                        newVal = LCase(Trim(CStr(newCellRaw)))
                        oldVal = LCase(Trim(CStr(oldCellRaw)))
                        Dim valChanged As Boolean
                        valChanged = False
                        If IsDate(newCellRaw) And IsDate(oldCellRaw) Then
                            valChanged = (CLng(CDate(newCellRaw)) <> CLng(CDate(oldCellRaw)))
                        ElseIf newVal <> oldVal Then
                            valChanged = True
                        End If
                        If valChanged Then
                            ' Collect for batch orange fill + comment (applied after loop)
                            changedCells.Add Array(newRow, newColIdx, CStr(oldCellRaw))
                        End If
                    End If
                End If
            Next header

            ' Import user-entered data (using pre-resolved column indices and arrays)
            For ihi = LBound(userImportHeaders) To UBound(userImportHeaders)
                importColOld = importColOldArr(ihi)
                importColNew = importColNewArr(ihi)
                If importColOld > 0 And importColNew > 0 Then
                    Dim oldImportRaw As Variant
                    If importColOld <= oldMaxCol Then oldImportRaw = oldDataArr(oldArrIdx, importColOld) Else oldImportRaw = ""
                    If CStr(oldImportRaw) <> "" Then
                        newSheet.Cells(newRow, importColNew).Value = oldImportRaw
                        If IsDate(oldImportRaw) Then
                            newSheet.Cells(newRow, importColNew).NumberFormat = "mm/dd/yyyy"
                        End If
                    End If
                End If
            Next ihi

            ' Migrate Completed=TRUE -> Status="Completed" (Rev14: Completed column merged into Status)
            If oldCompCol > 0 And newStatusCol > 0 Then
                Dim compVal As String
                If oldCompCol <= oldMaxCol Then compVal = LCase(Trim(CStr(oldDataArr(oldArrIdx, oldCompCol)))) Else compVal = ""
                If compVal = "true" Or compVal = "1" Then
                    ' Only set Completed if Status is still "Active" (don't override Cancelled etc.)
                    Dim curStatus As String
                    curStatus = LCase(Trim(CStr(newSheet.Cells(newRow, newStatusCol).Value)))
                    If curStatus = "active" Or curStatus = "" Then
                        newSheet.Cells(newRow, newStatusCol).Value = "Completed"
                    End If
                ElseIf compVal = "non iq" Or compVal = "noniq" Then
                    Dim curStatus2 As String
                    curStatus2 = LCase(Trim(CStr(newSheet.Cells(newRow, newStatusCol).Value)))
                    If curStatus2 = "active" Or curStatus2 = "" Then
                        newSheet.Cells(newRow, newStatusCol).Value = "On Hold"
                    End If
                End If
            End If

            oldKeyMap.Remove nKey
        Else
            ' New project marker
            If newEntityCodeCol > 0 Then
                ApplyNewProjectBorder newSheet.Cells(newRow, newEntityCodeCol)
            End If
        End If
    Next nKey

    ' === Batch apply orange fill + comments for changed cells (Fix 2) ===
    ' 1. Build Union range for orange fill — single Interior.Color call
    If changedCells.Count > 0 Then
        Dim orangeRange As Range
        Dim ci As Long
        For ci = 1 To changedCells.Count
            Dim cellInfo As Variant
            cellInfo = changedCells(ci)
            If orangeRange Is Nothing Then
                Set orangeRange = newSheet.Cells(cellInfo(0), cellInfo(1))
            Else
                Set orangeRange = Union(orangeRange, newSheet.Cells(cellInfo(0), cellInfo(1)))
            End If
        Next ci
        If Not orangeRange Is Nothing Then
            orangeRange.Interior.Color = CLR_CHANGE_FILL
        End If

        ' 2. Write comments (still per-cell but acceptable with ScreenUpdating=False)
        For ci = 1 To changedCells.Count
            cellInfo = changedCells(ci)
            Dim commentText As String
            commentText = "[" & Format(Date, "YYYY-MM-DD") & "] Changed from: " & CStr(cellInfo(2))
            Dim targetCell As Range
            Set targetCell = newSheet.Cells(cellInfo(0), cellInfo(1))
            On Error Resume Next
            If targetCell.Comment Is Nothing Then
                targetCell.AddComment commentText
            Else
                Dim tExisting As String
                tExisting = targetCell.Comment.Text
                If Len(tExisting) + Len(commentText) + 1 < 1024 Then
                    targetCell.Comment.Text tExisting & vbLf & commentText
                End If
            End If
            On Error GoTo 0
        Next ci
    End If

    ' Rev14: Carry forward removed projects with Status="Cancelled"
    ' Fix 4: Bulk copy each removed row using array I/O
    If oldKeyMap.Count > 0 Then
        Dim statusColNew As Long
        statusColNew = 0
        If outputColMap.exists(LCase(TIS_COL_STATUS)) Then statusColNew = outputColMap(LCase(TIS_COL_STATUS))

        ' Build old-to-new column map ONCE (not per row)
        Dim oldToNewCol As Object
        Set oldToNewCol = CreateObject("Scripting.Dictionary")
        Dim remC As Long, remOldHdrLower As String, remNewMatchCol As Long
        For remC = 1 To oldLastCol
            remOldHdrLower = LCase(Trim(Replace(Replace(CStr(oldSheet.Cells(oldHeaderRow, remC).Value), vbLf, ""), vbCr, "")))
            If remOldHdrLower <> "" Then
                remNewMatchCol = FindHeaderColInSheet(newSheet, dataStartRow, remOldHdrLower, totalColCount)
                If remNewMatchCol > 0 Then oldToNewCol(remC) = remNewMatchCol
            End If
        Next remC

        ' Determine max column span for bulk row copy
        Dim maxOldC As Long, maxNewC As Long
        maxOldC = 0: maxNewC = 0
        Dim remOldCKey As Variant
        For Each remOldCKey In oldToNewCol.Keys
            If CLng(remOldCKey) > maxOldC Then maxOldC = CLng(remOldCKey)
            If oldToNewCol(remOldCKey) > maxNewC Then maxNewC = oldToNewCol(remOldCKey)
        Next remOldCKey

        Dim appendRow As Long
        appendRow = dataStartRow + dataRowCount  ' first row after current data
        Dim remKey As Variant, remOldRow As Long
        For Each remKey In oldKeyMap.Keys
            remOldRow = oldKeyMap(remKey)
            ' Bulk-read entire old row into array, then map columns to new row array
            Dim srcRowArr As Variant
            Dim remOldArrIdx As Long
            remOldArrIdx = remOldRow - oldHeaderRow + 1
            ' Build destination row array
            If maxNewC > 0 And oldToNewCol.Count > 0 Then
                Dim destRowArr() As Variant
                ReDim destRowArr(1 To 1, 1 To maxNewC)
                ' Initialize to Empty
                Dim dri As Long
                For dri = 1 To maxNewC
                    destRowArr(1, dri) = Empty
                Next dri
                ' Map old columns to new columns using the cached old data array
                For Each remOldCKey In oldToNewCol.Keys
                    Dim srcC As Long, dstC As Long
                    srcC = CLng(remOldCKey)
                    dstC = oldToNewCol(remOldCKey)
                    If srcC <= oldMaxCol Then
                        destRowArr(1, dstC) = oldDataArr(remOldArrIdx, srcC)
                    End If
                Next remOldCKey
                ' Bulk-write the mapped row
                newSheet.Range(newSheet.Cells(appendRow, 1), newSheet.Cells(appendRow, maxNewC)).Value = destRowArr
            End If
            ' Set Status = "Cancelled"
            If statusColNew > 0 Then
                newSheet.Cells(appendRow, statusColNew).Value = "Cancelled"
            End If
            appendRow = appendRow + 1
        Next remKey
    End If
End Sub

'====================================================================
' FIND HEADER COLUMN IN SHEET
'====================================================================

Private Function FindHeaderColInSheet(ws As Worksheet, headerRow As Long, _
                                       searchText As String, maxCol As Long) As Long
    Dim c As Long, cellVal As String
    Dim searchLower As String
    searchLower = LCase(Trim(Replace(Replace(searchText, vbLf, ""), vbCr, "")))
    FindHeaderColInSheet = 0
    For c = 1 To maxCol
        cellVal = LCase(Trim(Replace(Replace(CStr(ws.Cells(headerRow, c).Value), vbLf, ""), vbCr, "")))
        If cellVal = searchLower Then FindHeaderColInSheet = c: Exit Function
    Next c
End Function

'====================================================================
' APPLY NEW PROJECT BORDER (blue)
'====================================================================

Private Sub ApplyNewProjectBorder(cell As Range)
    With cell.Borders(xlEdgeLeft): .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(0, 0, 255): End With
    With cell.Borders(xlEdgeRight): .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(0, 0, 255): End With
    With cell.Borders(xlEdgeTop): .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(0, 0, 255): End With
    With cell.Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(0, 0, 255): End With
End Sub

'====================================================================
' APPLY CHANGED CELL BORDER (yellow)
'====================================================================

Private Sub ApplyChangedCellBorder(cell As Range)
    With cell.Borders(xlEdgeLeft): .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(255, 152, 0): End With
    With cell.Borders(xlEdgeRight): .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(255, 152, 0): End With
    With cell.Borders(xlEdgeTop): .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(255, 152, 0): End With
    With cell.Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(255, 152, 0): End With
End Sub

'====================================================================
' ARCHIVE REMOVED PROJECTS to "Removed Systems" sheet
'====================================================================

Private Sub ArchiveRemovedProjects(oldSheet As Worksheet, oldHeaderRow As Long, _
                                    oldLastCol As Long, oldKeyMap As Object)
    Dim wsArchive As Worksheet
    Dim archiveRow As Long, j As Long
    Dim nowStamp As String
    Dim oKey As Variant, oldRow As Long
    
    nowStamp = Format(Now(), "mm/dd/yyyy hh:nn:ss")
    
    If Not TISCommon.SheetExists(ThisWorkbook, SHEET_REMOVED) Then
        Set wsArchive = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsArchive.Name = SHEET_REMOVED
        For j = 1 To oldLastCol
            wsArchive.Cells(1, j).Value = oldSheet.Cells(oldHeaderRow, j).Value
        Next j
        wsArchive.Cells(1, oldLastCol + 1).Value = "Date Removed"
        With wsArchive.Range(wsArchive.Cells(1, 1), wsArchive.Cells(1, oldLastCol + 1))
            .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(139, 0, 0)
            .HorizontalAlignment = xlCenter: .WrapText = True: .RowHeight = 30
        End With
        archiveRow = 2
    Else
        Set wsArchive = ThisWorkbook.Sheets(SHEET_REMOVED)
        archiveRow = wsArchive.Cells(wsArchive.Rows.Count, 1).End(xlUp).Row + 1
        If archiveRow < 2 Then archiveRow = 2
    End If
    
    For Each oKey In oldKeyMap.Keys
        oldRow = oldKeyMap(oKey)
        wsArchive.Range(wsArchive.Cells(archiveRow, 1), wsArchive.Cells(archiveRow, oldLastCol)).Value = _
            oldSheet.Range(oldSheet.Cells(oldRow, 1), oldSheet.Cells(oldRow, oldLastCol)).Value
        wsArchive.Cells(archiveRow, oldLastCol + 1).Value = nowStamp
        archiveRow = archiveRow + 1
    Next oKey
    
    ' AutoFit only used columns on the archive sheet
    Dim archLastC As Long
    archLastC = wsArchive.Cells(1, wsArchive.Columns.Count).End(xlToLeft).Column
    If archLastC > 0 Then wsArchive.Range(wsArchive.Columns(1), wsArchive.Columns(archLastC)).AutoFit
End Sub

'====================================================================
' INSTALL SHEET EVENTS - Worksheet_Calculate handler for slicer responsiveness
' When a slicer/filter changes on a ListObject table, the sheet's Calculate
' event fires. We use this to force dashboard rows (2-5) to recalculate,
' which updates the SUMPRODUCT(SUBTOTAL(...)) counter formulas.
'====================================================================

Private Sub InstallSheetEvents(ws As Worksheet)
    Dim vbComp As Object
    Dim codeModule As Object
    Dim eventCode As String

    On Error GoTo EventInstallError

    ' Access the sheet's code module
    Set vbComp = ThisWorkbook.VBProject.VBComponents(ws.CodeName)
    Set codeModule = vbComp.CodeModule

    ' Check which handlers already exist
    Dim i As Long
    Dim hasCalcHandler As Boolean: hasCalcHandler = False
    Dim hasChangeHandler As Boolean: hasChangeHandler = False
    For i = 1 To codeModule.CountOfLines
        If InStr(codeModule.Lines(i, 1), "Worksheet_Calculate") > 0 Then hasCalcHandler = True
        If InStr(codeModule.Lines(i, 1), "Worksheet_Change") > 0 Then hasChangeHandler = True
    Next i

    ' Insert Worksheet_Calculate handler if not present
    If Not hasCalcHandler Then
        eventCode = vbCrLf & _
                    "Private Sub Worksheet_Calculate()" & vbCrLf & _
                    "    ' Force dashboard rows to recalculate when slicers/filters change" & vbCrLf & _
                    "    Static inCalc As Boolean" & vbCrLf & _
                    "    If inCalc Then Exit Sub" & vbCrLf & _
                    "    ' Guard: only fire if counter row has content" & vbCrLf & _
                    "    If IsEmpty(Me.Cells(3, 1)) And IsEmpty(Me.Cells(4, 1)) Then Exit Sub" & vbCrLf & _
                    "    inCalc = True" & vbCrLf & _
                    "    On Error Resume Next" & vbCrLf & _
                    "    Application.EnableEvents = False" & vbCrLf & _
                    "    Me.Range(""2:6"").Calculate" & vbCrLf & _
                    "    Application.EnableEvents = True" & vbCrLf & _
                    "    On Error GoTo 0" & vbCrLf & _
                    "    inCalc = False" & vbCrLf & _
                    "End Sub" & vbCrLf
        codeModule.AddFromString eventCode
    End If

    ' Insert Worksheet_Change handler for Lock? real-time enforcement
    If Not hasChangeHandler Then
        eventCode = vbCrLf & "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf
        eventCode = eventCode & "    On Error GoTo ExitHandler" & vbCrLf
        eventCode = eventCode & "    Dim lockCol As Long" & vbCrLf
        eventCode = eventCode & "    lockCol = 0" & vbCrLf
        eventCode = eventCode & "    On Error Resume Next" & vbCrLf
        eventCode = eventCode & "    lockCol = Application.Range(""LOCK_COL"").Value" & vbCrLf
        eventCode = eventCode & "    On Error GoTo ExitHandler" & vbCrLf
        eventCode = eventCode & "    If lockCol = 0 Then Exit Sub" & vbCrLf
        eventCode = eventCode & "    If Target.Column <> lockCol Then Exit Sub" & vbCrLf
        eventCode = eventCode & "    If Target.Row < 16 Then Exit Sub" & vbCrLf
        eventCode = eventCode & "    Dim ourStart As Long, ourEnd As Long" & vbCrLf
        eventCode = eventCode & "    On Error Resume Next" & vbCrLf
        eventCode = eventCode & "    ourStart = Application.Range(""OUR_DATE_START"").Value" & vbCrLf
        eventCode = eventCode & "    ourEnd = Application.Range(""OUR_DATE_END"").Value" & vbCrLf
        eventCode = eventCode & "    On Error GoTo ExitHandler" & vbCrLf
        eventCode = eventCode & "    If ourStart = 0 Or ourEnd = 0 Then Exit Sub" & vbCrLf
        eventCode = eventCode & "    Application.EnableEvents = False" & vbCrLf
        eventCode = eventCode & "    Dim r As Long, c As Long" & vbCrLf
        eventCode = eventCode & "    For r = Target.Row To Target.Row + Target.Rows.Count - 1" & vbCrLf
        eventCode = eventCode & "        Dim lockVal As String" & vbCrLf
        eventCode = eventCode & "        lockVal = LCase(Trim(CStr(Me.Cells(r, lockCol).Value)))" & vbCrLf
        eventCode = eventCode & "        Dim isLocked As Boolean" & vbCrLf
        eventCode = eventCode & "        isLocked = (lockVal = ""true"")" & vbCrLf
        eventCode = eventCode & "        For c = ourStart To ourEnd" & vbCrLf
        eventCode = eventCode & "            Me.Cells(r, c).Locked = isLocked" & vbCrLf
        eventCode = eventCode & "        Next c" & vbCrLf
        eventCode = eventCode & "    Next r" & vbCrLf
        eventCode = eventCode & "ExitHandler:" & vbCrLf
        eventCode = eventCode & "    Application.EnableEvents = True" & vbCrLf
        eventCode = eventCode & "End Sub" & vbCrLf
        codeModule.AddFromString eventCode
    End If

    Exit Sub

EventInstallError:
    ' If VBProject access is denied (Trust Center), use alternative approach:
    ' Set up an Application-level event handler instead
    DebugLog "Could not install sheet events: " & Err.Description
    DebugLog "Dashboard may not auto-update with slicers. Ensure 'Trust access to VBA project' is enabled."
End Sub

'====================================================================
' REV14: ACTIVATE WHATIF MODE
' For rows with a WhatIf date:
'   1. Compute delta = WhatIf date - project start date
'      (project start = first non-null Our Date, excluding SDD)
'   2. Back up original Our Date values to a hidden "WhatIf Backup" sheet
'   3. Shift all Our Dates by delta (except SDD — SDD never shifts)
'   4. Rebuild Gantt (which reads Our Dates)
' Our Dates are temporarily overwritten. DeactivateWhatIfMode restores them.
'====================================================================

Public Sub ActivateWhatIfMode()
    On Error GoTo ErrorHandler
    Dim appSt As AppState
    appSt = SaveAppState()
    SetPerformanceMode

    ' Save viewport position — restored in Cleanup before ScreenUpdating re-enables,
    ' preventing the scroll-position jump caused by Sheets.Add + Gantt rebuild.
    Dim savedScrollRow As Long, savedScrollCol As Long
    savedScrollRow = ActiveWindow.ScrollRow
    savedScrollCol = ActiveWindow.ScrollColumn

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then GoTo Cleanup

    ' Guard: if already in WhatIf mode (backup exists), restore first
    Dim existingBak As Worksheet
    On Error Resume Next
    Set existingBak = ThisWorkbook.Worksheets("WhatIf_Backup")
    On Error GoTo ErrorHandler
    If Not existingBak Is Nothing Then
        ' Already in WhatIf mode — restore first before re-activating
        DeactivateWhatIfMode
        Set existingBak = Nothing
    End If

    Dim headerRow As Long
    headerRow = TISCommon.FindHeaderRow(ws)
    If headerRow = 0 Then GoTo Cleanup

    Dim lastRow As Long, maxCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    maxCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastRow <= headerRow Then GoTo Cleanup

    ' Build column map
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    Dim c As Long, hv As String
    For c = 1 To maxCol
        hv = LCase(Trim(Replace(Replace(CStr(ws.Cells(headerRow, c).Value), vbLf, ""), vbCr, "")))
        If hv <> "" And Not colMap.exists(hv) Then colMap(hv) = c
    Next c

    ' Find WhatIf column
    Dim whatIfCol As Long
    whatIfCol = 0
    If colMap.exists(LCase(TIS_COL_WHATIF)) Then whatIfCol = colMap(LCase(TIS_COL_WHATIF))
    If whatIfCol = 0 Then GoTo Cleanup

    ' Our Date column positions (all are shiftable — no SDD in Our Dates)
    Dim ourColKeys(0 To 7) As String
    ourColKeys(0) = LCase(TIS_COL_OUR_SET)
    ourColKeys(1) = LCase(TIS_COL_OUR_SL1)
    ourColKeys(2) = LCase(TIS_COL_OUR_SL2)
    ourColKeys(3) = LCase(TIS_COL_OUR_SQ)
    ourColKeys(4) = LCase(TIS_COL_OUR_CONVS)
    ourColKeys(5) = LCase(TIS_COL_OUR_CONVF)
    ourColKeys(6) = LCase(TIS_COL_OUR_MRCLS)
    ourColKeys(7) = LCase(TIS_COL_OUR_MRCLF)

    Dim ourCols(0 To 7) As Long
    Dim oi As Long
    For oi = 0 To 7
        ourCols(oi) = 0
        If colMap.exists(ourColKeys(oi)) Then ourCols(oi) = colMap(ourColKeys(oi))
    Next oi

    ' Find Status column (skip non-Active rows)
    Dim statusCol As Long
    statusCol = 0
    If colMap.exists(LCase(TIS_COL_STATUS)) Then statusCol = colMap(LCase(TIS_COL_STATUS))

    ' --- Step 1: Back up Our Dates to hidden sheet ---
    Dim wsBak As Worksheet
    Dim bakName As String
    bakName = "WhatIf_Backup"

    ' Delete existing backup if present
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(bakName).Delete
    Application.DisplayAlerts = True
    Err.Clear
    On Error GoTo ErrorHandler

    Set wsBak = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsBak.Name = bakName
    wsBak.Visible = xlSheetVeryHidden
    ' Restore focus and viewport — Sheets.Add changes ActiveSheet even with ScreenUpdating=False
    ws.Activate
    ActiveWindow.ScrollRow = savedScrollRow
    ActiveWindow.ScrollColumn = savedScrollCol

    ' Copy Our Date range to backup (header + data)
    Dim ourStartCol As Long, ourEndCol As Long
    ourStartCol = ourCols(0)  ' Set
    ourEndCol = ourCols(7)    ' MRCL.F
    If ourStartCol > 0 And ourEndCol > 0 Then
        ws.Range(ws.Cells(headerRow, ourStartCol), ws.Cells(lastRow, ourEndCol)).Copy _
            Destination:=wsBak.Cells(1, 1)
    End If

    ' --- Step 2: Compute delta and shift Our Dates (bulk array I/O) ---
    ' ourStartCol / ourEndCol already set above for backup
    If ourStartCol = 0 Or ourEndCol = 0 Then GoTo Cleanup

    ' Bulk-read the entire Our Date range into a 2D array
    Dim ourData As Variant
    ourData = ws.Range(ws.Cells(headerRow + 1, ourStartCol), _
                        ws.Cells(lastRow, ourEndCol)).Value

    Dim rowCount As Long
    rowCount = lastRow - headerRow
    Dim shiftedCount As Long
    shiftedCount = 0

    Dim r As Long, mi As Long, si As Long
    Dim colOffset As Long
    Dim wsRow As Long
    Dim wifDate As Variant
    Dim projectStartIdx As Long
    Dim delta As Long

    For r = 1 To rowCount
        wsRow = headerRow + r

        ' Skip non-Active rows
        If statusCol > 0 Then
            If LCase(Trim(CStr(ws.Cells(wsRow, statusCol).Value))) <> "active" Then GoTo NextWIFRow
        End If

        wifDate = ws.Cells(wsRow, whatIfCol).Value
        If Not IsDate(wifDate) Then GoTo NextWIFRow

        ' Find project start: first non-null Our Date in array
        projectStartIdx = 0
        For si = 0 To 7
            If ourCols(si) > 0 Then
                colOffset = ourCols(si) - ourStartCol + 1  ' 1-based array index
                If IsDate(ourData(r, colOffset)) Then
                    projectStartIdx = colOffset
                    Exit For
                End If
            End If
        Next si
        If projectStartIdx = 0 Then GoTo NextWIFRow

        ' Compute delta
        delta = CLng(CDate(wifDate)) - CLng(CDate(ourData(r, projectStartIdx)))
        shiftedCount = shiftedCount + 1

        ' Shift all Our Dates in the array
        For mi = 0 To 7
            If ourCols(mi) > 0 Then
                colOffset = ourCols(mi) - ourStartCol + 1
                If IsDate(ourData(r, colOffset)) Then
                    ourData(r, colOffset) = CDate(ourData(r, colOffset)) + delta
                End If
            End If
        Next mi
NextWIFRow:
    Next r

    If shiftedCount = 0 Then
        ' No WhatIf dates entered — clean up backup and exit silently
        On Error Resume Next
        Application.DisplayAlerts = False
        wsBak.Delete
        Application.DisplayAlerts = True
        If Err.Number <> 0 Then
            DebugLog "WARNING: Could not delete WhatIf backup: " & Err.Description
            Err.Clear
            wsBak.Visible = xlSheetVeryHidden
        End If
        On Error GoTo ErrorHandler
        GoTo Cleanup
    End If

    ' Write shifted dates back — ONLY for rows that were actually shifted
    ' Use bulk row write: one Range.Value per shifted row (fast, preserves format)
    Dim rowArr(1 To 1, 1 To 1) As Variant
    Dim colCount As Long
    colCount = ourEndCol - ourStartCol + 1
    For r = 1 To rowCount
        wsRow = headerRow + r
        ' Skip non-Active rows
        If statusCol > 0 Then
            If LCase(Trim(CStr(ws.Cells(wsRow, statusCol).Value))) <> "active" Then GoTo NextWriteRow
        End If
        wifDate = ws.Cells(wsRow, whatIfCol).Value
        If Not IsDate(wifDate) Then GoTo NextWriteRow
        ' This row was shifted — write the entire row of Our Dates at once
        Dim rowData() As Variant
        ReDim rowData(1 To 1, 1 To colCount)
        For mi = 1 To colCount
            rowData(1, mi) = ourData(r, mi)
        Next mi
        ws.Range(ws.Cells(wsRow, ourStartCol), ws.Cells(wsRow, ourEndCol)).Value = rowData
        ' Ensure date format
        ws.Range(ws.Cells(wsRow, ourStartCol), ws.Cells(wsRow, ourEndCol)).NumberFormat = "mm/dd/yyyy"
NextWriteRow:
    Next r

    ' --- Step 3: Rebuild Gantt with shifted dates ---
    ' Save ALL data column widths (Gantt rebuild resets them)
    Dim savedWidths() As Double
    Dim cwi As Long
    Dim dataLastCol As Long
    dataLastCol = maxCol  ' Save all columns up to header scan boundary
    If dataLastCol < 1 Then dataLastCol = 100
    ReDim savedWidths(1 To dataLastCol)
    For cwi = 1 To dataLastCol
        savedWidths(cwi) = ws.Columns(cwi).ColumnWidth
    Next cwi

    On Error Resume Next
    GanttBuilder.BuildGantt silent:=True, targetSheet:=ws
    If Err.Number <> 0 Then
        DebugLog "WARNING: GanttBuilder failed in WhatIf: " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    ' Restore data column widths
    For cwi = 1 To dataLastCol
        ws.Columns(cwi).ColumnWidth = savedWidths(cwi)
    Next cwi

    ' Recompute Health with shifted dates
    PopulateHealthColumn ws, colMap, headerRow, rowCount + 1

    ' --- Step 4: Visual indicator ---
    ' Ensure we're on the Working Sheet (Gantt rebuild may shift focus)
    ws.Activate
    ' Change sheet tab color to amber
    ws.Tab.Color = RGB(255, 183, 77)

    DebugLog "WhatIf Mode Active: " & shiftedCount & " system(s) shifted."

    GoTo Cleanup

ErrorHandler:
    MsgBox "Error in ActivateWhatIfMode: " & Err.Description, vbCritical
Cleanup:
    On Error Resume Next
    If Not ws Is Nothing Then
        ' Restore viewport before ScreenUpdating re-enables to prevent scroll jump
        ws.Activate
        ActiveWindow.ScrollRow = savedScrollRow
        ActiveWindow.ScrollColumn = savedScrollCol
    End If
    On Error GoTo 0
    RestoreAppState appSt
End Sub

'====================================================================
' REV14: DEACTIVATE WHATIF MODE
' Restores Our Dates from the hidden backup sheet, rebuilds Gantt.
'====================================================================

Public Sub DeactivateWhatIfMode()
    On Error GoTo ErrorHandler
    Dim appSt As AppState
    appSt = SaveAppState()
    SetPerformanceMode

    ' Save viewport position — restored in Cleanup before ScreenUpdating re-enables
    Dim savedScrollRow As Long, savedScrollCol As Long
    savedScrollRow = ActiveWindow.ScrollRow
    savedScrollCol = ActiveWindow.ScrollColumn

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then GoTo Cleanup

    ' Find backup sheet
    Dim wsBak As Worksheet
    On Error Resume Next
    Set wsBak = ThisWorkbook.Worksheets("WhatIf_Backup")
    On Error GoTo ErrorHandler

    If wsBak Is Nothing Then
        GoTo Cleanup
    End If

    Dim headerRow As Long
    headerRow = TISCommon.FindHeaderRow(ws)
    If headerRow = 0 Then GoTo Cleanup

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Find Our Date columns in Working Sheet
    Dim maxCol As Long
    maxCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    Dim c As Long, hv As String
    For c = 1 To maxCol
        hv = LCase(Trim(Replace(Replace(CStr(ws.Cells(headerRow, c).Value), vbLf, ""), vbCr, "")))
        If hv <> "" And Not colMap.exists(hv) Then colMap(hv) = c
    Next c

    Dim ourStartCol As Long, ourEndCol As Long
    ourStartCol = 0: ourEndCol = 0
    If colMap.exists(LCase(TIS_COL_OUR_SET)) Then ourStartCol = colMap(LCase(TIS_COL_OUR_SET))
    If colMap.exists(LCase(TIS_COL_OUR_MRCLF)) Then ourEndCol = colMap(LCase(TIS_COL_OUR_MRCLF))

    ' Restore from backup (values only, not formulas)
    If ourStartCol > 0 And ourEndCol > 0 Then
        Dim bakRows As Long
        ' Find last row across ALL backup columns (not just column 1 which might have gaps)
        Dim bakMaxCol As Long
        bakMaxCol = ourEndCol - ourStartCol + 1
        bakRows = 1
        Dim bkC As Long
        For bkC = 1 To bakMaxCol
            Dim bkLastR As Long
            bkLastR = wsBak.Cells(wsBak.Rows.Count, bkC).End(xlUp).Row
            If bkLastR > bakRows Then bakRows = bkLastR
        Next bkC
        If bakRows > 1 Then
            ' Copy values back (skip header row in backup = row 1, data starts row 2)
            Dim bakData As Variant
            bakData = wsBak.Range(wsBak.Cells(2, 1), wsBak.Cells(bakRows, ourEndCol - ourStartCol + 1)).Value
            Dim restoreEndRow As Long
            restoreEndRow = headerRow + bakRows - 1
            If restoreEndRow > lastRow Then restoreEndRow = lastRow
            ws.Range(ws.Cells(headerRow + 1, ourStartCol), _
                     ws.Cells(restoreEndRow, ourEndCol)).Value = bakData
            ws.Range(ws.Cells(headerRow + 1, ourStartCol), _
                     ws.Cells(restoreEndRow, ourEndCol)).NumberFormat = "mm/dd/yyyy"
        End If
    End If

    ' Delete backup sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    wsBak.Delete
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then
        DebugLog "WARNING: Could not delete WhatIf backup: " & Err.Description
        Err.Clear
        ' Force-hide and rename to mark as consumed
        wsBak.Name = "WhatIf_Consumed_" & Format(Now, "hhnnss")
        wsBak.Visible = xlSheetVeryHidden
    End If
    On Error GoTo ErrorHandler

    ' Save ALL data column widths before Gantt rebuild
    Dim savedWidths() As Double
    Dim cwi As Long
    Dim dataLastCol As Long
    dataLastCol = maxCol
    If dataLastCol < 1 Then dataLastCol = 100
    ReDim savedWidths(1 To dataLastCol)
    For cwi = 1 To dataLastCol
        savedWidths(cwi) = ws.Columns(cwi).ColumnWidth
    Next cwi

    ' Rebuild Gantt with original dates
    On Error Resume Next
    GanttBuilder.BuildGantt silent:=True, targetSheet:=ws
    If Err.Number <> 0 Then
        DebugLog "WARNING: GanttBuilder failed in restore: " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    ' Restore data column widths
    For cwi = 1 To dataLastCol
        ws.Columns(cwi).ColumnWidth = savedWidths(cwi)
    Next cwi

    ' Recompute Health (it reads Our Dates)
    Dim wsDataRowCount As Long
    wsDataRowCount = lastRow - headerRow
    PopulateHealthColumn ws, colMap, headerRow, wsDataRowCount + 1

    ' Ensure we're on the Working Sheet
    ws.Activate
    ' Reset tab color
    ws.Tab.ColorIndex = xlColorIndexNone

    DebugLog "Normal mode restored. Our Dates reverted to committed values."

    GoTo Cleanup

ErrorHandler:
    MsgBox "Error in DeactivateWhatIfMode: " & Err.Description, vbCritical
Cleanup:
    On Error Resume Next
    If Not ws Is Nothing Then
        ' Restore viewport before ScreenUpdating re-enables to prevent scroll jump
        ws.Activate
        ActiveWindow.ScrollRow = savedScrollRow
        ActiveWindow.ScrollColumn = savedScrollCol
    End If
    On Error GoTo 0
    RestoreAppState appSt
End Sub
