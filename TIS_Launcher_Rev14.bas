Attribute VB_Name = "TIS_Launcher"
'====================================================================
' TIS Launcher
'
' Provides clean public entry points for all TIS Tracker operations.
' Creates an Instructions sheet with embedded buttons and usage guide.
'
' Public Subs (appear in Macro menu):
'   Step1_LoadTIS          - Load/compare TIS files
'   Step2_BuildWorkingSheet - Create Working Sheet from TIS data
'   Step3_BuildGantt       - Add Gantt chart to Working Sheet
'   Step4_BuildNIF         - Add NIF assignments + HC analyzer
'   Step5_BuildDashboard   - Build management Dashboard with charts
'   Setup_Instructions     - Create/refresh the Instructions sheet
'   StripAllModules        - Remove all TIS VBA modules for clean re-import
'   LoadAllModules         - Import all .bas files from a user-selected folder
'
' Stable-naming change: VB_Name is now "TIS_Launcher" (no _RevNN suffix).
' Module filenames may carry a Rev suffix (e.g., TIS_Launcher_Rev11.bas)
' but the internal VB_Name is stable so cross-module calls require no
' update when the revision number increments.
'
' Rev11 changes from Rev10:
'   - VB_Name stabilised to "TIS_Launcher" (no _RevNN suffix)
'   - Step subs updated to use stable module names (no _RevNN suffix)
'   - LoadAllModules reads VB_Name from inside each .bas file (not filename)
'     so Rev-suffixed filenames map correctly to stable VB_Names
'   - StripAllModules replaced with pattern-based matching (no hardcoded list)
'   - Added ReadVBNameFromFile, IsTISModule, GetModuleBaseName helpers
'   - Setup_Instructions version label reads from TISCommon.TIS_VERSION
'====================================================================

Option Explicit

Private Const SHEET_INSTRUCTIONS As String = "TIS Tracker"
Private Const SELF_MODULE_NAME As String = "TIS_Launcher"

'====================================================================
' PUBLIC ENTRY POINTS
'====================================================================

Public Sub Step1_LoadTIS()
    TISLoader.LoadNewTIS
End Sub

Public Sub Step2_BuildWorkingSheet()
    WorkfileBuilder.CreateWorkFile
End Sub

Public Sub Step3_BuildGantt()
    GanttBuilder.BuildGantt
End Sub

Public Sub Step4_BuildNIF()
    NIF_Builder.BuildNIF
End Sub

Public Sub Step5_BuildDashboard()
    DashboardBuilder.BuildDashboard
End Sub

Public Sub ActivateWhatIf()
    WorkfileBuilder.ActivateWhatIfMode
End Sub

Public Sub DeactivateWhatIf()
    WorkfileBuilder.DeactivateWhatIfMode
End Sub

Public Sub ToggleWhatIf()
    ' Check if WhatIf backup exists — if yes, we're in WhatIf mode, so deactivate
    Dim wsBak As Worksheet
    On Error Resume Next
    Set wsBak = ThisWorkbook.Worksheets("WhatIf_Backup")
    On Error GoTo 0

    If Not wsBak Is Nothing Then
        WorkfileBuilder.DeactivateWhatIfMode
    Else
        WorkfileBuilder.ActivateWhatIfMode
    End If
End Sub

'====================================================================
' LOAD ALL MODULES
' Imports all .bas files from a user-selected folder.
' Existing modules with the same name are removed first.
' TIS_Launcher is skipped (cannot replace running module).
' Reads the true VB_Name from inside each file so that Rev-suffixed
' filenames (e.g., WorkfileBuilder_Rev12.bas) map to stable VB_Names
' (e.g., "WorkfileBuilder") for correct duplicate detection.
'====================================================================

Public Sub LoadAllModules()
    Dim vbProj As Object
    Dim fd As Object
    Dim folderPath As String
    Dim fileName As String
    Dim filePath As String
    Dim modName As String
    Dim vbComp As Object
    Dim imported As String
    Dim skipped As String
    Dim count As Long
    Dim skipCount As Long

    ' Get VBA Project access
    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0

    If vbProj Is Nothing Then
        MsgBox "Cannot access VBA Project." & vbCrLf & vbCrLf & _
               "Enable: File > Options > Trust Center > Trust Center Settings > " & _
               "Macro Settings > 'Trust access to the VBA project object model'", _
               vbCritical, "Access Denied"
        Exit Sub
    End If

    ' Folder picker
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Select folder containing .bas module files"
        .InitialFileName = ThisWorkbook.Path & "\"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    ' Ensure trailing separator
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' Scan for .bas files
    fileName = Dir(folderPath & "*.bas")
    If fileName = "" Then
        MsgBox "No .bas files found in:" & vbCrLf & folderPath, _
               vbExclamation, "No Modules Found"
        Exit Sub
    End If

    ' Confirmation
    Dim fileList As String
    Dim tempName As String
    tempName = fileName
    Do While tempName <> ""
        fileList = fileList & "  " & tempName & vbCrLf
        tempName = Dir()
    Loop

    If MsgBox("Import the following .bas files from:" & vbCrLf & _
              folderPath & vbCrLf & vbCrLf & _
              fileList & vbCrLf & _
              "Existing modules with matching names will be replaced." & vbCrLf & vbCrLf & _
              "Continue?", vbYesNo + vbQuestion, "Load All Modules") = vbNo Then
        Exit Sub
    End If

    count = 0
    skipCount = 0
    imported = ""
    skipped = ""

    ' Re-scan (Dir was consumed by the listing above)
    fileName = Dir(folderPath & "*.bas")

    Do While fileName <> ""
        filePath = folderPath & fileName

        ' Read true VB_Name from inside the file (not from filename).
        ' This supports stable VB_Names (e.g., "WorkfileBuilder") in files
        ' with Rev-suffixed filenames (e.g., "WorkfileBuilder_Rev12.bas").
        Dim fileNameNoExt As String
        fileNameNoExt = Left(fileName, Len(fileName) - 4)
        modName = ReadVBNameFromFile(filePath, fileNameNoExt)

        ' Skip self (cannot replace running module)
        If UCase(modName) = UCase(SELF_MODULE_NAME) Then
            skipCount = skipCount + 1
            skipped = skipped & "  " & ChrW(9888) & " " & fileName & " (running - skipped)" & vbCrLf
            GoTo NextFile
        End If

        ' Remove existing module with same name (if any)
        On Error Resume Next
        Set vbComp = vbProj.VBComponents(modName)
        On Error GoTo 0

        If Not vbComp Is Nothing Then
            If vbComp.Type = 1 Then  ' Standard module only
                vbProj.VBComponents.Remove vbComp
            End If
            Set vbComp = Nothing
        End If

        ' Import the .bas file
        On Error Resume Next
        vbProj.VBComponents.Import filePath
        If Err.Number = 0 Then
            count = count + 1
            imported = imported & "  " & ChrW(10003) & " " & modName & " (" & fileName & ")" & vbCrLf
        Else
            skipCount = skipCount + 1
            skipped = skipped & "  " & ChrW(10007) & " " & fileName & " (error: " & Err.Description & ")" & vbCrLf
            Err.Clear
        End If
        On Error GoTo 0

NextFile:
        fileName = Dir()
    Loop

    ' Summary
    Dim msg As String
    msg = "Imported " & count & " module(s):" & vbCrLf & vbCrLf & imported
    If skipCount > 0 Then
        msg = msg & vbCrLf & "Skipped " & skipCount & ":" & vbCrLf & skipped
    End If
    msg = msg & vbCrLf & "Done! You may need to re-run Setup_Instructions to refresh buttons."

    MsgBox msg, vbInformation, "Load Complete"
End Sub

'====================================================================
' STRIP ALL MODULES
' Removes all TIS Tracker VBA modules except TIS_Launcher (self).
' Uses pattern-based matching — automatically handles any Rev number
' (Rev9, Rev10, Rev11, Rev12, ...) without a hardcoded list.
' Safe to run before LoadAllModules for a clean re-import.
'====================================================================
Public Sub StripAllModules()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim removed As String
    Dim skippedMsg As String
    Dim count As Long
    Dim i As Long
    Dim modName As String
    Dim allNames() As String
    Dim nameCount As Long

    ' Confirmation
    If MsgBox("This will remove all TIS Tracker VBA modules (except TIS_Launcher) " & vbCrLf & _
              "from this workbook." & vbCrLf & vbCrLf & _
              "You can then re-import fresh .bas files using 'Load All Modules'." & vbCrLf & vbCrLf & _
              "Continue?", vbYesNo + vbExclamation, "Strip All Modules") = vbNo Then
        Exit Sub
    End If

    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0

    If vbProj Is Nothing Then
        MsgBox "Cannot access VBA Project." & vbCrLf & vbCrLf & _
               "Enable: File > Options > Trust Center > Trust Center Settings > " & _
               "Macro Settings > 'Trust access to the VBA project object model'", _
               vbCritical, "Access Denied"
        Exit Sub
    End If

    ' Collect all standard module names first.
    ' Cannot safely remove while iterating the VBComponents collection.
    nameCount = 0
    ReDim allNames(0 To vbProj.VBComponents.Count - 1)
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = 1 Then  ' vbext_ct_StdModule
            allNames(nameCount) = vbComp.Name
            nameCount = nameCount + 1
        End If
    Next vbComp

    count = 0
    removed = ""
    skippedMsg = ""

    For i = 0 To nameCount - 1
        modName = allNames(i)

        ' Self-protect: never remove the running Launcher module
        If UCase(modName) = UCase(SELF_MODULE_NAME) Then
            skippedMsg = skippedMsg & "  " & ChrW(9888) & " " & modName & " (running - kept)" & vbCrLf
            GoTo NextComp
        End If

        ' Pattern-based match: strips _RevNN suffix, checks base name
        If IsTISModule(modName) Then
            On Error Resume Next
            Set vbComp = vbProj.VBComponents(modName)
            On Error GoTo 0
            If Not vbComp Is Nothing Then
                On Error Resume Next
                vbProj.VBComponents.Remove vbComp
                If Err.Number = 0 Then
                    count = count + 1
                    removed = removed & "  " & ChrW(10003) & " " & modName & vbCrLf
                Else
                    skippedMsg = skippedMsg & "  " & ChrW(10007) & " " & modName & " (error: " & Err.Description & ")" & vbCrLf
                    Err.Clear
                End If
                On Error GoTo 0
                Set vbComp = Nothing
            End If
        End If
NextComp:
    Next i

    Dim msg As String
    If count > 0 Then
        msg = "Removed " & count & " module(s):" & vbCrLf & vbCrLf & removed
    Else
        msg = "No TIS modules found to remove." & vbCrLf & vbCrLf
    End If
    If skippedMsg <> "" Then
        msg = msg & vbCrLf & "Kept:" & vbCrLf & skippedMsg
    End If
    msg = msg & vbCrLf & "Use 'Load All Modules' to import fresh .bas files."

    MsgBox msg, vbInformation, "Strip Complete"
End Sub

'====================================================================
' IS TIS MODULE
' Returns True if the given VBA module name belongs to the TIS Tracker.
' Matches both stable names (e.g., "WorkfileBuilder") and any
' Rev-suffixed variant (e.g., "WorkfileBuilder_Rev11", "WorkfileBuilder_Rev9").
' TIS_Launcher is intentionally included — old Rev-named Launcher
' modules are stripped, while the running stable-named module is
' protected by the self-skip check in StripAllModules before this runs.
'====================================================================
Private Function IsTISModule(modName As String) As Boolean
    Dim baseName As String
    baseName = UCase(GetModuleBaseName(modName))
    Select Case baseName
        Case "TISCOMMON", _
             "TISLOADER", _
             "WORKFILEBUILDER", _
             "GANTTBUILDER", _
             "NIF_BUILDER", _
             "DASHBOARDBUILDER", _
             "RAMPALIGNMENT", _
             "HCHEATMAP", _
             "TIS_LAUNCHER"
            IsTISModule = True
        Case Else
            IsTISModule = False
    End Select
End Function

'====================================================================
' GET MODULE BASE NAME
' Strips a trailing _RevNN suffix (case-insensitive) from a VBA module name.
' Returns the base name unchanged if no such suffix is present.
'
' Examples:
'   "WorkfileBuilder_Rev11"  ->  "WorkfileBuilder"
'   "TIS_Launcher_Rev11"     ->  "TIS_Launcher"
'   "TISCommon"              ->  "TISCommon"
'   "WorkfileBuilder"        ->  "WorkfileBuilder"
'====================================================================
Private Function GetModuleBaseName(modName As String) As String
    Dim i As Integer
    Dim suffix As String

    ' Walk backwards looking for "_Rev" followed by one or more digits
    For i = Len(modName) - 4 To 1 Step -1  ' minimum "_Rev1" is 5 chars
        If UCase(Mid(modName, i, 4)) = "_REV" Then
            suffix = Mid(modName, i + 4)
            If Len(suffix) > 0 Then
                Dim allDigits As Boolean
                Dim k As Integer
                allDigits = True
                For k = 1 To Len(suffix)
                    If Mid(suffix, k, 1) < "0" Or Mid(suffix, k, 1) > "9" Then
                        allDigits = False
                        Exit For
                    End If
                Next k
                If allDigits Then
                    GetModuleBaseName = Left(modName, i - 1)
                    Exit Function
                End If
            End If
        End If
    Next i

    GetModuleBaseName = modName  ' no _RevNN suffix found
End Function

'====================================================================
' READ VB_NAME FROM FILE
' Reads the Attribute VB_Name value from the first 5 lines of a .bas file.
' Returns the filename-sans-extension as fallback if attribute not found.
' This is used instead of deriving the module name from the filename,
' so that stable VB_Names (e.g., "WorkfileBuilder") are correctly
' matched even when the filename has a Rev suffix (e.g., "WorkfileBuilder_Rev12.bas").
'====================================================================
Private Function ReadVBNameFromFile(filePath As String, fileNameNoExt As String) As String
    Dim fileNum As Integer
    Dim lineText As String
    Dim i As Integer
    Dim posOpen As Integer
    Dim posClose As Integer

    On Error GoTo Fallback

    fileNum = FreeFile
    Open filePath For Input As #fileNum

    For i = 1 To 5
        If EOF(fileNum) Then Exit For
        Line Input #fileNum, lineText
        lineText = Trim(lineText)
        If lineText Like "Attribute VB_Name = *" Then
            posOpen = InStr(lineText, Chr(34))
            If posOpen > 0 Then
                posClose = InStr(posOpen + 1, lineText, Chr(34))
                If posClose > posOpen Then
                    ReadVBNameFromFile = Mid(lineText, posOpen + 1, posClose - posOpen - 1)
                    Close #fileNum
                    Exit Function
                End If
            End If
        End If
    Next i

    Close #fileNum
    ReadVBNameFromFile = fileNameNoExt
    Exit Function

Fallback:
    On Error GoTo 0
    If fileNum > 0 Then
        On Error Resume Next
        Close #fileNum
        On Error GoTo 0
    End If
    ReadVBNameFromFile = fileNameNoExt
End Function

'====================================================================
' SETUP INSTRUCTIONS SHEET
' Creates a polished instructions/control panel sheet with buttons
'====================================================================

Public Sub Setup_Instructions()
    Dim ws As Worksheet
    Dim btn As Button
    Dim r As Long

    ' Delete existing if present
    If TISCommon.SheetExists(ThisWorkbook, SHEET_INSTRUCTIONS) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(SHEET_INSTRUCTIONS).Delete
        Application.DisplayAlerts = True
    End If

    ' Create as first sheet
    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = SHEET_INSTRUCTIONS
    ws.Tab.Color = RGB(44, 62, 80)

    ' Light theme
    ws.Cells.Font.Name = "Calibri"
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 55
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("F").ColumnWidth = 55

    ' -- TITLE --
    r = 2
    ws.Cells(r, 2).Value = "TIS Tracker"
    With ws.Cells(r, 2)
        .Font.Size = 24: .Font.Bold = True: .Font.Color = RGB(44, 62, 80)
    End With
    ws.Cells(r, 3).Value = TISCommon.TIS_VERSION
    With ws.Cells(r, 3)
        .Font.Size = 12: .Font.Color = RGB(160, 160, 170): .HorizontalAlignment = xlLeft
    End With

    r = 3
    ws.Cells(r, 2).Value = "Tool Installation System Tracker"
    With ws.Cells(r, 2)
        .Font.Size = 10: .Font.Color = RGB(120, 130, 140): .Font.Italic = True
    End With

    ' -- WORKFLOW BUTTONS --
    r = 5
    ws.Cells(r, 2).Value = "Workflow"
    With ws.Cells(r, 2)
        .Font.Size = 14: .Font.Bold = True: .Font.Color = RGB(44, 62, 80)
    End With

    ' Separator line
    r = 6
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(200, 210, 220)
    End With

    ' Button 1: Load TIS
    r = 8
    CreateWorkflowButton ws, r, 2, "Step1_LoadTIS", _
        ChrW(9654) & " Step 1: Load TIS", RGB(52, 152, 219), _
        "Load a new TIS file or compare with existing TIS data."

    ' Button 2: Build Working Sheet
    r = 11
    CreateWorkflowButton ws, r, 2, "Step2_BuildWorkingSheet", _
        ChrW(9654) & " Step 2: Build Working Sheet", RGB(46, 204, 113), _
        "Rebuilds the Working Sheet in-place (preserves external formula links). Old data backed up as 'Old [date]'. Auto-builds Gantt, NIF, and Dashboard."

    ' Button 3: Build Gantt
    r = 14
    CreateWorkflowButton ws, r, 2, "Step3_BuildGantt", _
        ChrW(9654) & " Step 3: Build Gantt", RGB(155, 89, 182), _
        "Adds/rebuilds the Gantt chart on the Working Sheet. Uses formula-based rendering with conditional formatting."

    ' Button 4: Build NIF
    r = 17
    CreateWorkflowButton ws, r, 2, "Step4_BuildNIF", _
        ChrW(9654) & " Step 4: Build NIF", RGB(230, 126, 34), _
        "Adds NIF assignment columns and HC Analyzer tables (separate New/Reuse Available and Gap tables). Restores previous assignments if available."

    ' Button 5: Build Dashboard
    r = 20
    CreateWorkflowButton ws, r, 2, "Step5_BuildDashboard", _
        ChrW(9654) & " Step 5: Build Dashboard", RGB(0, 188, 153), _
        "Builds management Dashboard with KPI cards (incl. Conversions, Completed), group-filterable counters, collapsible CEID drill-down, HC gap analysis, and date-filterable charts."

    ' -- UTILITIES --
    r = 24
    ws.Cells(r, 2).Value = "Utilities"
    With ws.Cells(r, 2)
        .Font.Size = 14: .Font.Bold = True: .Font.Color = RGB(44, 62, 80)
    End With

    r = 25
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(200, 210, 220)
    End With

    ' Button: Load All Modules
    r = 27
    CreateWorkflowButton ws, r, 2, "LoadAllModules", _
        ChrW(10145) & " Load All Modules", RGB(41, 128, 185), _
        "Import all .bas module files from a selected folder. Replaces existing modules with same name. Use after Strip to reload fresh code."

    ' Button: Strip All Modules
    r = 30
    CreateWorkflowButton ws, r, 2, "StripAllModules", _
        ChrW(9888) & " Strip All Modules", RGB(192, 57, 43), _
        "Removes all TIS VBA modules (except Launcher) for clean re-import. Requires VBA project trust access."


    ' -- CONFIGURATION GUIDE --
    r = 5
    ws.Cells(r, 5).Value = "Configuration"
    With ws.Cells(r, 5)
        .Font.Size = 14: .Font.Bold = True: .Font.Color = RGB(44, 62, 80)
    End With

    r = 6
    With ws.Range(ws.Cells(r, 5), ws.Cells(r, 6)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(200, 210, 220)
    End With

    ' Definitions sheet guide
    r = 8
    WriteGuideSection ws, r, 5, "Definitions Sheet", _
        "Column A: Header names (sirfisheaders) to include from TIS" & vbLf & _
        "Column B: Filter values (space-separated for OR, AND across rows)" & vbLf & _
        "Column C: Sort priority (1=highest) or 'X' for past-date exclusion" & vbLf & _
        "Column D: Gating codes (A1, A2 = sequential gates)" & vbLf & _
        "Column E: Exclude systems (EntityCode|EventType|Site)" & vbLf & _
        "Column F: Milestone tokens (A1|A2 = start/end pair for duration)" & vbLf & _
        "Column G: Milestone display name (e.g., 'Set - SL1')" & vbLf & _
        "Column H: 'X' to include column in Gantt view" & vbLf & _
        "Column J: Group number for column collapsing" & vbLf & _
        "Cell S1: Cycle Time threshold (default: 85 days)" & vbLf & _
        "Cell S2: Gantt start date override (leave blank for auto-calculate)"

    ' CEIDs sheet guide
    r = 21
    WriteGuideSection ws, r, 5, "CEIDs Sheet", _
        "Column A: Entity Type (lookup key)" & vbLf & _
        "Column B: Group name (e.g., 'Litho', 'Etch')" & vbLf & _
        "Used for Group VLOOKUP column and NIF group filtering."

    ' New-Reused sheet guide
    r = 25
    WriteGuideSection ws, r, 5, "New-Reused Sheet", _
        "Column A: Entity Code" & vbLf & _
        "Column B: 'New' or leave blank (blank = Reused)" & vbLf & _
        "Systems with Event Type 'Demo' auto-detected as Demo." & vbLf & _
        "Determines system classification for counters and Gantt text color."

    ' Working Sheet features
    r = 30
    WriteGuideSection ws, r, 5, "Working Sheet Features", _
        "In-place rebuild: external formula links preserved across rebuilds" & vbLf & _
        "Old data backed up as 'Old [YYYY-MM-DD]' sheet for review" & vbLf & _
        "Group Slicer: filters data table and HC analyzer tables" & vbLf & _
        "Time Frame: enter weeks to filter counters by upcoming projects" & vbLf & _
        "Counters: Total, New, Reused, Demo, Conversions, Completed" & vbLf & _
        "Gantt: auto-generated from milestone definitions (formula-based)" & vbLf & _
        "NIF: separate New/Reuse Available HC and Gap tables"

    ' Dashboard features
    r = 37
    WriteGuideSection ws, r, 5, "Dashboard Features", _
        "KPI Cards: Total, New, Reused, Demo, CT Miss, Escalated, Watched, Conversions, Completed" & vbLf & _
        "Group Filter: dropdown to filter KPI card values by group" & vbLf & _
        "System Counters: New, Reused, Total, Demo, Conversions, Completed with collapsible CEIDs" & vbLf & _
        "HC Gap Analysis: per-group New/Reuse Need, Available, and Gap breakdown" & vbLf & _
        "Chart Start Date: selectable start date for activity and systems charts" & vbLf & _
        "Monthly Activity: stacked bar chart (New/Reused/Demo systems added)" & vbLf & _
        "Active Systems: stacked area chart (systems over time)" & vbLf & _
        "Group Breakdown: stacked bar (systems per group by type)" & vbLf & _
        "Escalation Tracker: summary table with status tracking"

    ' Freeze and select
    ws.Rows(1).RowHeight = 10
    ws.Cells(1, 1).Select

    ws.Activate

    MsgBox "Instructions sheet created!" & vbCrLf & vbCrLf & _
           "Use the buttons on the 'TIS Tracker' sheet to run each step.", _
           vbInformation, "Setup Complete"
End Sub

'====================================================================
' CREATE WORKFLOW BUTTON (helper)
'====================================================================

Private Sub CreateWorkflowButton(ws As Worksheet, row As Long, col As Long, _
                                   macroName As String, caption As String, _
                                   accentColor As Long, description As String)
    Dim btn As Button
    Dim btnLeft As Double, btnTop As Double

    ' Button
    btnLeft = ws.Cells(row, col).Left
    btnTop = ws.Cells(row, col).Top

    Set btn = ws.Buttons.Add(btnLeft, btnTop, 170, 24)
    btn.OnAction = macroName
    btn.Characters.Text = caption
    btn.Font.Name = "Calibri"
    btn.Font.Size = 10
    btn.Font.Bold = True
    btn.Name = "btn_" & Replace(macroName, ".", "_")

    ' Description below button
    ws.Cells(row + 1, col).Value = description
    With ws.Cells(row + 1, col)
        .Font.Size = 8: .Font.Color = RGB(120, 130, 140)
        .WrapText = True
    End With
    ws.Rows(row + 1).RowHeight = 28
End Sub

'====================================================================
' WRITE GUIDE SECTION (helper)
'====================================================================

Private Sub WriteGuideSection(ws As Worksheet, row As Long, col As Long, _
                                title As String, body As String)
    ws.Cells(row, col).Value = title
    With ws.Cells(row, col)
        .Font.Size = 11: .Font.Bold = True: .Font.Color = RGB(44, 62, 80)
    End With

    ws.Cells(row + 1, col).Value = body
    With ws.Cells(row + 1, col)
        .Font.Size = 8: .Font.Color = RGB(80, 90, 100)
        .WrapText = True: .VerticalAlignment = xlTop
    End With

    ' Auto-height for wrapped text
    Dim lineCount As Long
    lineCount = Len(body) - Len(Replace(body, vbLf, "")) + 1
    ws.Rows(row + 1).RowHeight = Application.WorksheetFunction.Max(lineCount * 13, 26)
End Sub
