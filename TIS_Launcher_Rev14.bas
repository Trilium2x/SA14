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
'   - Setup_Instructions version label reads from verStr
'====================================================================

Option Explicit

Private Const SHEET_CONTROLS As String = "Controls"
Private Const SHEET_INSTRUCTIONS As String = "TIS Tracker"  ' legacy name
Private Const SELF_MODULE_NAME As String = "TIS_Launcher"

' Inline fallbacks so Launcher can compile and run Setup_Controls
' even before TISCommon is loaded (first-time bootstrap scenario).
' When TISCommon IS loaded, these are shadowed by the Public constants.
#If False Then
    ' These exist solely to prevent compile errors if TISCommon is absent.
    ' The actual values come from TISCommon Public constants at runtime.
#End If
Private Const FALLBACK_BG As Long = 3349260       ' RGB(12, 27, 51)
Private Const FALLBACK_ACCENT As Long = 12491862   ' RGB(86, 156, 190)
Private Const FALLBACK_VERSION As String = "Rev14"

'====================================================================
' PUBLIC ENTRY POINTS
'====================================================================

' --- TIS Management ---
Public Sub LoadCompareTIS()
    TISLoader.LoadAndCompareTIS
End Sub

Public Sub UpdateTISToWS()
    TISLoader.ApplyTISToWorkingSheet
End Sub

' Legacy entry point (calls both steps for backward compatibility)
Public Sub Step1_LoadTIS()
    TISLoader.LoadNewTIS
End Sub

' --- Build ---
Public Sub BuildWorkingSheet()
    WorkfileBuilder.CreateWorkFile
End Sub

Public Sub BuildGantt()
    GanttBuilder.BuildGantt
End Sub

Public Sub BuildNIF()
    NIF_Builder.BuildNIF
End Sub

Public Sub BuildDashboard()
    DashboardBuilder.BuildDashboard
End Sub

' Legacy wrappers (backward compatibility)
Public Sub Step2_BuildWorkingSheet(): BuildWorkingSheet: End Sub
Public Sub Step3_BuildGantt(): BuildGantt: End Sub
Public Sub Step4_BuildNIF(): BuildNIF: End Sub
Public Sub Step5_BuildDashboard(): BuildDashboard: End Sub

' --- Scenario ---
Public Sub ToggleWhatIf()
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

Public Sub ActivateWhatIf(): WorkfileBuilder.ActivateWhatIfMode: End Sub
Public Sub DeactivateWhatIf(): WorkfileBuilder.DeactivateWhatIfMode: End Sub

' --- Reports ---
Public Sub RunRampForGroup()
    ' Read group from Controls sheet dropdown cell
    Dim ctrlWs As Worksheet
    On Error Resume Next
    Set ctrlWs = ThisWorkbook.Worksheets(SHEET_CONTROLS)
    On Error GoTo 0
    If ctrlWs Is Nothing Then
        RampAlignment.BuildRampAlignment
        Exit Sub
    End If
    Dim grpCell As Range
    Set grpCell = Nothing
    On Error Resume Next
    Set grpCell = ctrlWs.Range("RAMP_GROUP_SELECT")
    On Error GoTo 0
    If grpCell Is Nothing Or grpCell.Value = "" Or grpCell.Value = "All" Then
        RampAlignment.BuildRampAlignment
    Else
        ' Set the RampAlignment dropdown named range to the selected group,
        ' then call Generate (it reads from the named range, takes no arguments)
        On Error Resume Next
        Dim raDD As Range
        Set raDD = ThisWorkbook.Names("RampAlign_DropdownCell").RefersToRange
        If Not raDD Is Nothing Then raDD.Value = CStr(grpCell.Value)
        On Error GoTo 0
        RampAlignment.RampAlignment_Generate
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
' Uses pattern-based matching -- automatically handles any Rev number
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
' TIS_Launcher is intentionally included -- old Rev-named Launcher
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
    Setup_Controls
End Sub


Public Sub Setup_Controls()
    Dim ws As Worksheet
    Dim r As Long

    ' Use inline constants (no TISCommon dependency -- Launcher must compile standalone)
    Dim clrBG As Long, clrAccent As Long, verStr As String
    clrBG = FALLBACK_BG
    clrAccent = FALLBACK_ACCENT
    verStr = FALLBACK_VERSION

    ' Delete existing Controls or legacy sheet
    Dim delNames As Variant, dn As Variant
    delNames = Array(SHEET_CONTROLS, "TIS Tracker")
    For Each dn In delNames
        Dim chkWs As Worksheet
        Set chkWs = Nothing
        On Error Resume Next
        Set chkWs = ThisWorkbook.Worksheets(CStr(dn))
        On Error GoTo 0
        If Not chkWs Is Nothing Then
            Application.DisplayAlerts = False
            chkWs.Delete
            Application.DisplayAlerts = True
        End If
    Next dn

    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = SHEET_CONTROLS
    ws.Tab.Color = clrBG

    ' --- Dark background ---
    ws.Cells.Interior.Color = clrBG
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Color = RGB(226, 232, 240)
    Dim ci As Long
    For ci = 1 To 12
        ws.Columns(ci).ColumnWidth = Choose(ci, 2, 12, 16, 2, 12, 16, 2, 14, 16, 2, 14, 16)
    Next ci

    Dim labelClr As Long: labelClr = RGB(148, 163, 184)

    ' === TITLE BAR ===
    r = 2
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Merge
    ws.Cells(r, 2).Value = "CONTROLS"
    With ws.Cells(r, 2): .Font.Size = 22: .Font.Bold = True: End With
    ws.Range(ws.Cells(r, 8), ws.Cells(r, 12)).Merge
    ws.Cells(r, 8).Value = "Raizing fabs " & verStr
    With ws.Cells(r, 8): .Font.Size = 12: .Font.Color = labelClr: .HorizontalAlignment = xlRight: End With
    r = 3
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Merge
    ws.Cells(r, 2).Value = "Tool Installation System Tracker"
    With ws.Cells(r, 2): .Font.Size = 9: .Font.Color = labelClr: .Font.Italic = True: End With

    ' === TIS MANAGEMENT card (blue theme) ===
    Dim clrBlue As Long: clrBlue = RGB(52, 152, 219)
    Dim clrGreen As Long: clrGreen = RGB(46, 184, 92)
    Dim clrAmber As Long: clrAmber = RGB(255, 183, 77)
    Dim clrSlate As Long: clrSlate = RGB(71, 85, 105)

    r = 5: WriteCardTitle ws, r, 2, "TIS MANAGEMENT", clrBlue
    r = 7: CreateControlButton ws, r, 2, "TIS_Launcher.LoadCompareTIS", "Load / Compare TIS", clrBlue
    r = 9
    ws.Cells(r, 2).Value = "Loads new TIS, compares with previous."
    With ws.Cells(r, 2): .Font.Size = 8: .Font.Color = labelClr: .WrapText = True: End With
    r = 11: CreateControlButton ws, r, 2, "TIS_Launcher.UpdateTISToWS", "Update TIS to WS", clrBlue
    r = 13
    ws.Cells(r, 2).Value = "Applies TIS changes to Working Sheet."
    With ws.Cells(r, 2): .Font.Size = 8: .Font.Color = labelClr: .WrapText = True: End With
    r = 15
    ws.Cells(r, 2).Value = "Status"
    With ws.Cells(r, 2): .Font.Size = 9: .Font.Bold = True: End With
    r = 16
    ws.Cells(r, 2).Value = "TIS loaded:": ws.Cells(r, 3).Value = "--"
    With ws.Cells(r, 2): .Font.Size = 8: .Font.Color = labelClr: End With
    With ws.Cells(r, 3): .Font.Size = 8: End With
    r = 17
    ws.Cells(r, 2).Value = "WS synced:": ws.Cells(r, 3).Value = "--"
    With ws.Cells(r, 2): .Font.Size = 8: .Font.Color = labelClr: End With
    With ws.Cells(r, 3): .Font.Size = 8: End With
    ' Named ranges for TISLoader to update status display
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="TIS_STATUS_DATE", RefersTo:="=" & ws.Name & "!$C$16"
    ThisWorkbook.Names.Add Name:="TIS_STATUS_SYNC", RefersTo:="=" & ws.Name & "!$C$17"
    On Error GoTo 0

    ' === BUILD card (green theme) ===
    r = 5: WriteCardTitle ws, r, 8, "BUILD", clrGreen
    r = 7: CreateControlButton ws, r, 8, "TIS_Launcher.BuildWorkingSheet", "Build Working Sheet", clrGreen
    r = 9: CreateControlButton ws, r, 8, "TIS_Launcher.BuildGantt", "Build Gantt", clrGreen
    r = 11: CreateControlButton ws, r, 8, "TIS_Launcher.BuildNIF", "Build NIF", clrGreen
    r = 13: CreateControlButton ws, r, 8, "TIS_Launcher.BuildDashboard", "Build Dashboard", clrGreen

    ' === SCENARIO card (amber theme) ===
    r = 20: WriteCardTitle ws, r, 2, "SCENARIO", clrAmber
    r = 22: CreateControlButton ws, r, 2, "TIS_Launcher.ToggleWhatIf", "Toggle WhatIf", clrAmber
    r = 24
    ws.Cells(r, 2).Value = "WhatIf Mode:"
    ws.Cells(r, 3).Value = "OFF"
    With ws.Cells(r, 2): .Font.Size = 8: .Font.Color = labelClr: End With
    With ws.Cells(r, 3): .Font.Size = 9: .Font.Bold = True: .Font.Color = RGB(46, 204, 113): End With

    ' === REPORTS card (blue theme) ===
    r = 20: WriteCardTitle ws, r, 8, "REPORTS", clrBlue
    r = 22
    ws.Cells(r, 8).Value = "Group:"
    With ws.Cells(r, 8): .Font.Size = 9: .Font.Color = labelClr: End With
    ws.Cells(r, 9).Value = "All"
    With ws.Cells(r, 9)
        .Font.Size = 9: .Font.Bold = True: .Font.Color = RGB(30, 41, 59)
        .Interior.Color = RGB(226, 232, 240): .HorizontalAlignment = xlCenter
    End With
    ' Populate group dropdown from Working Sheet
    Dim grpList As String: grpList = "All"
    ' Populate group list (inline -- no TISCommon dependency for bootstrap)
    Dim wsData As Worksheet: Set wsData = Nothing
    On Error Resume Next
    Set wsData = FindWorkingSheetInline()
    On Error GoTo 0
    If Not wsData Is Nothing Then
        Dim grpCol As Long, grpHdr As Long
        grpHdr = FindHeaderRowInline(wsData)
        If grpHdr > 0 Then
            grpCol = FindHeaderColInline(wsData, grpHdr, "Group")
            If grpCol > 0 Then
                Dim grpDict As Object: Set grpDict = CreateObject("Scripting.Dictionary")
                Dim gr As Long
                For gr = grpHdr + 1 To wsData.Cells(wsData.Rows.Count, grpCol).End(xlUp).Row
                    Dim gv As String: gv = Trim(CStr(wsData.Cells(gr, grpCol).Value & ""))
                    If gv <> "" And Not grpDict.exists(gv) Then grpDict(gv) = True: grpList = grpList & "," & gv
                Next gr
            End If
        End If
    End If
    On Error Resume Next
    ws.Cells(r, 9).Validation.Delete
    ws.Cells(r, 9).Validation.Add Type:=xlValidateList, Formula1:=grpList
    ws.Cells(r, 9).Validation.InCellDropdown = True
    ThisWorkbook.Names.Add Name:="RAMP_GROUP_SELECT", RefersTo:="=" & ws.Name & "!" & ws.Cells(r, 9).Address
    On Error GoTo 0
    r = 24: CreateControlButton ws, r, 8, "TIS_Launcher.RunRampForGroup", "Ramp Alignment", clrBlue

    ' === UTILITIES card (slate theme) ===
    r = 28: WriteCardTitle ws, r, 2, "UTILITIES", clrSlate
    r = 30: CreateControlButton ws, r, 2, "LoadAllModules", "Load Modules", clrSlate
    r = 32: CreateControlButton ws, r, 2, "StripAllModules", "Strip Modules", clrSlate

    ' === REFERENCE card ===
    r = 28: WriteCardTitle ws, r, 8, "REFERENCE", labelClr
    r = 30: WriteRefLine ws, r, 8, "Definitions", "Headers, filters, sort, milestones, gating"
    r = 32: WriteRefLine ws, r, 8, "CEIDs", "Entity Type -> Group mapping"
    r = 34: WriteRefLine ws, r, 8, "New-Reused", "Entity Code -> New/Reused/Demo"
    r = 36: WriteRefLine ws, r, 8, "Milestones", "STD durations per CEID per phase"
    r = 38: WriteRefLine ws, r, 8, "Working Sheet", "Our Dates, Status, Health, WhatIf, Gantt, NIF"
    r = 40: WriteRefLine ws, r, 8, "Dashboard", "KPIs, HC Gap, Charts, Escalation Tracker"

    ws.Rows(1).RowHeight = 8
    ws.Cells(1, 1).Select
    ws.Activate
    MsgBox "Controls panel created.", vbInformation, "Raizing fabs"
End Sub

'====================================================================
' CONTROL PANEL HELPERS
'====================================================================

Private Sub CreateControlButton(ws As Worksheet, row As Long, col As Long, _
                                  macroName As String, caption As String, accentColor As Long)
    Dim shp As Shape
    Dim btnLeft As Double, btnTop As Double
    btnLeft = ws.Cells(row, col).Left
    btnTop = ws.Cells(row, col).Top
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, 180, 26)
    shp.Name = "ctrl_" & Replace(Replace(macroName, ".", "_"), " ", "_")
    shp.OnAction = macroName
    shp.Placement = xlFreeFloating
    shp.Shadow.Visible = msoFalse
    shp.Line.Visible = msoFalse
    shp.Fill.ForeColor.RGB = accentColor
    On Error Resume Next
    shp.Adjustments(1) = 0.3
    On Error GoTo 0
    With shp.TextFrame
        .Characters.Text = caption
        .Characters.Font.Name = "Segoe UI"
        .Characters.Font.Size = 10
        .Characters.Font.Bold = True
        .Characters.Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .MarginLeft = 4: .MarginRight = 4
        .MarginTop = 2: .MarginBottom = 2
    End With
End Sub

Private Sub WriteCardTitle(ws As Worksheet, row As Long, col As Long, _
                             title As String, accentColor As Long)
    ws.Cells(row, col).Value = title
    With ws.Cells(row, col)
        .Font.Size = 11: .Font.Bold = True: .Font.Color = accentColor
    End With
    With ws.Range(ws.Cells(row + 1, col), ws.Cells(row + 1, col + 1)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(51, 65, 85)
    End With
End Sub

Private Sub WriteRefLine(ws As Worksheet, row As Long, col As Long, _
                           title As String, desc As String)
    ws.Cells(row, col).Value = title
    With ws.Cells(row, col): .Font.Size = 9: .Font.Bold = True: .Font.Color = RGB(226, 232, 240): End With
    ws.Cells(row + 1, col).Value = desc
    With ws.Cells(row + 1, col): .Font.Size = 8: .Font.Color = RGB(148, 163, 184): .WrapText = True: End With
End Sub

' Legacy helper (kept for backward compatibility)
Private Sub CreateWorkflowButton(ws As Worksheet, row As Long, col As Long, _
                                   macroName As String, caption As String, _
                                   accentColor As Long, description As String)
    CreateControlButton ws, row, col, macroName, caption, accentColor
    ws.Cells(row + 1, col).Value = description
    With ws.Cells(row + 1, col): .Font.Size = 8: .Font.Color = RGB(120, 130, 140): .WrapText = True: End With
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

'====================================================================
' INLINE HELPERS (no TISCommon dependency -- for bootstrap)
'====================================================================

Private Function FindWorkingSheetInline() As Worksheet
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If s.Name Like "Working Sheet*" Then
            Set FindWorkingSheetInline = s: Exit Function
        End If
    Next s
    Set FindWorkingSheetInline = Nothing
End Function

Private Function FindHeaderRowInline(ws As Worksheet) As Long
    Dim rr As Long, cc As Long, hv2 As String
    For rr = 1 To 20
        For cc = 1 To 50
            hv2 = LCase(Trim(CStr(ws.Cells(rr, cc).Value & "")))
            If hv2 = "ceid" Or hv2 = "entity code" Or hv2 = "site" Then
                FindHeaderRowInline = rr: Exit Function
            End If
        Next cc
    Next rr
    FindHeaderRowInline = 0
End Function

Private Function FindHeaderColInline(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim cc2 As Long, lc2 As Long, hv3 As String
    lc2 = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For cc2 = 1 To lc2
        hv3 = LCase(Trim(CStr(ws.Cells(headerRow, cc2).Value & "")))
        If hv3 = LCase(headerName) Then FindHeaderColInline = cc2: Exit Function
    Next cc2
    FindHeaderColInline = 0
End Function
