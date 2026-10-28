Attribute VB_Name = "OperationsRefresh"
Option Explicit
Option Compare Text

' Operations Workload Dashboard - Windows desktop Excel.
' Import this standard module into the supplied workbook, save as .xlsm,
' then run InstallOperationsDashboard once.
' No VBA-project access, add-ins, Windows APIs, or extra references required.

Private Const MODULE_VERSION As String = "1.5.4"
Private Const ERR_BASE As Long = vbObjectError + 2400
Private Const SECURITY_FORCE_DISABLE As Long = 3
Private Const DATA_COLUMNS As Long = 29
Private Const UNASSIGNED_TEAM As String = "[UNASSIGNED]"
Private Const COLOR_NAVY As Long = &H5D3617
Private Const COLOR_LINE As Long = &HE1D5CB
Private Const COLOR_INPUT As Long = &HDAEFE2
Private Const COLOR_AMBER As Long = &HCCF2FF
Private Const COLOR_GREY As Long = &HE6E6E7
Private Const TEAM_META_COLUMN As Long = 52
Private Const RADAR_DATA_COLUMN As Long = 58
Private Const RADAR_DETAIL_COLUMN As Long = 11
Private Const RADAR_LAST_COLUMN As Long = 18
Private Const RADAR_MAX_INDEX As Double = 800 ' 100 = 1.0x benchmark; cap at 8.0x.

Private mBusy As Boolean
Private mTeams As Object, mMembers As Object, mFields As Object, mStats As Object
Private mReadKeys As Object, mFingerprints As Object, mParticipants As Object
Private mRadarBenchmarks As Object
Private mTeamOrder As Collection, mSources As Collection, mAccounts As Collection, mIssues As Collection
Private mAliases As Collection, mManagerAliases As Collection, mFileAssignments As Collection
Private mRecords As Collection, mLogs As Collection, mAmountMetrics As Collection
Private mSourceBook As Workbook
Private mReportDate As Date
Private mFolder As String
Private mDataLastRow As Long
Private mWarnings As Long
Private mRunID As String
Private mAmountThreshold As Double

' TYPING-ORDER COPY OF VERSION 1.5.4
' Keep all procedures in ONE standard module in the practice workbook's VBA project.
' MANUAL ENTRY: do NOT type the Attribute VB_Name line above; it is .bas import metadata.
' Insert > Module, use F4 to name the module OperationsRefresh, and type the Option lines,
' all Private Const declarations and all Private module variables before any procedure.
' Save the practice workbook as .xlsm before running installer/appearance checkpoints.
' Version 1.5.4 caps the light radar at 8.0x; its adjacent raw-value table remains uncapped.
' The five TypingCheck... macros are optional learning checkpoints, not production refreshes.
' Compile with Debug > Compile VBAProject at the end of EACH stage. Do not add later-stage
' callers early, replace missing functions with empty stubs, or remove Private declarations.
' For line-by-line tracing, put the cursor inside a checkpoint and use F8; use Ctrl+G for output.
' Buttons created in Stage 3 cannot refresh until RefreshOperations is typed in Stage 8.
' TypingCheckEmptyOutputs deliberately replaces practice outputs with zero-entry previews.
' Use a practice copy and remove all TypingCheck... macros when learning is complete.
' Native desktop Excel compilation/execution of this file has not been performed here.


' ========================================================================
' STAGE 1: FOUNDATIONS
' ========================================================================
' Compile, then run TypingCheckBasics. Read its output in the Immediate Window (Ctrl+G).
' Procedure typing order:
' 01. Fail
' 02. TextOf
' 03. CleanName
' 04. NormalizeHeader
' 05. NewMap
' 06. ValueOf
' 07. RequiredText
' 08. LiteralValue

Private Sub Fail(ByVal message As String)
    Err.Raise ERR_BASE, "OperationsRefresh", message
End Sub

Private Function TextOf(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    TextOf = Trim$(Replace(CStr(value), ChrW(160), " "))
End Function

Private Function CleanName(ByVal value As String) As String
    value = Replace(value, ChrW(160), " ")
    value = Replace(value, vbTab, " "): value = Replace(value, vbCr, " "): value = Replace(value, vbLf, " ")
    Do While InStr(value, "  ") > 0: value = Replace(value, "  ", " "): Loop
    CleanName = Trim$(value)
End Function

Private Function NormalizeHeader(ByVal value As String) As String
    Dim i As Long, ch As String, result As String
    value = LCase$(value)
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then result = result & ch
    Next i
    NormalizeHeader = result
End Function

Private Function NewMap(ParamArray pairs() As Variant) As Object
    Dim i As Long, result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    If (UBound(pairs) + 1) Mod 2 <> 0 Then Fail "Internal mapping requires key/value pairs."
    For i = 0 To UBound(pairs) Step 2: result.Add CStr(pairs(i)), pairs(i + 1): Next i
    Set NewMap = result
End Function

Private Function ValueOf(ByVal row As Object, ByVal key As String) As Variant
    If row.Exists(key) Then ValueOf = row(key)
End Function

Private Function RequiredText(ByVal row As Object, ByVal key As String) As String
    RequiredText = TextOf(ValueOf(row, key))
    If Len(RequiredText) = 0 Then Fail "Missing configuration value: " & key & " (row " & TextOf(ValueOf(row, "@row")) & ")."
End Function

Private Function LiteralValue(ByVal value As Variant) As Variant
    Dim text As String
    LiteralValue = value
    If VarType(value) <> vbString Then Exit Function
    text = LTrim$(CStr(value))
    If Len(text) > 0 Then
        Select Case Left$(text, 1)
            Case "=", "+", "-", "@": LiteralValue = "'" & CStr(value)
        End Select
    End If
End Function

' Optional temporary checkpoint for this stage.

Public Sub TypingCheckBasics()
    Dim map As Object
    Set map = NewMap("Member", "002")
    Debug.Print "Member:", map("member")
    Debug.Print "Header:", NormalizeHeader("Age (days)")
    Debug.Print "Name:", CleanName("  Kate" & ChrW(160) & "  Smith  ")
    Debug.Print "Literal text:", LiteralValue("=1+1")
End Sub

' ========================================================================
' STAGE 2: README AND APPEARANCE
' ========================================================================
' On a saved practice .xlsm, create empty Dashboard, Configuration and Report Mapping tabs. Run 
' RepairOperationsDashboardAppearance; check ReadMe, tab colours and Arial.
' Procedure typing order:
' 01. SheetOrNothing
' 02. EnsureSheetExists
' 03. FitReadMeRows
' 04. PutRow
' 05. ColumnWidths
' 06. SeedReadMeSheet
' 07. ApplyWorkbookPresentation
' 08. RequireMacroWorkbook
' 09. RequiredSheet
' 10. CreateRecoveryCopy
' 11. RepairOperationsDashboardAppearance

Private Function SheetOrNothing(ByVal book As Workbook, ByVal name As String) As Worksheet
    On Error Resume Next
    Set SheetOrNothing = book.Worksheets(name)
    On Error GoTo 0
End Function

Private Sub EnsureSheetExists(ByVal name As String)
    Dim ws As Worksheet
    Set ws = SheetOrNothing(ThisWorkbook, name)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = name
    End If
End Sub

Private Sub FitReadMeRows(ByVal ws As Worksheet)
    Dim row As Variant
    For Each row In Array(5, 6, 7, 8, 9, 10, 15, 16, 17, 18, 23, 24, 25, 26, 27, 28)
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 2)).Rows.AutoFit
        ws.Rows(row).RowHeight = Application.Max(ws.Rows(row).RowHeight, IIf(row < 11, 30, 36))
    Next row
End Sub

Private Sub PutRow(ByVal ws As Worksheet, ByVal row As Long, ByVal column As Long, ByVal values As Variant)
    Dim c As Long, data() As Variant, count As Long
    count = UBound(values) - LBound(values) + 1
    ReDim data(1 To 1, 1 To count)
    For c = 1 To count: data(1, c) = LiteralValue(values(LBound(values) + c - 1)): Next c
    ws.Cells(row, column).Resize(1, count).Value2 = data
End Sub

Private Sub ColumnWidths(ByVal ws As Worksheet, ParamArray pairs() As Variant)
    Dim i As Long
    For i = 0 To UBound(pairs) Step 2: ws.Range(CStr(pairs(i))).ColumnWidth = CDbl(pairs(i + 1)): Next i
End Sub

Private Sub SeedReadMeSheet(ByVal ws As Worksheet)
    Dim row As Variant
    ' ReadMe A1:H32 is generated content; leave notes outside this area intact.
    With ws.Range("A1:H32")
        .UnMerge: .Clear
        .Font.Name = "Arial": .Font.Size = 10: .Font.Color = RGB(31, 31, 31)
        .Interior.Color = vbWhite: .VerticalAlignment = xlCenter: .RowHeight = 18
    End With
    ws.Range("A2").Value2 = "How to use this workbook"
    With ws.Range("A2:H2")
        .Font.Size = 16: .Font.Bold = True: .Font.Color = RGB(31, 78, 120): .RowHeight = 30
    End With
    With ws.Range("A3:H3").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(183, 201, 214)
    End With
    PutRow ws, 5, 1, Array("1", "Review Configuration and replace the fictional teams, managers, and members.")
    PutRow ws, 6, 1, Array("2", "Set each team's applicable reports to Yes or No; set the source folder, reporting date, and amount threshold.")
    PutRow ws, 7, 1, Array("3", "Update Report Mapping when source filenames, account tabs, headers, aliases, or tracked OMS issues change.")
    PutRow ws, 8, 1, Array("4", "Click Refresh all teams on Dashboard. Dashboard remains the compact executive view for report and team workload.")
    PutRow ws, 9, 1, Array("5", "On Team Summary, set Show and Radar to Yes or No, then click Apply team selection. Selected radars and their detail tables appear below the manager totals.")
    PutRow ws, 10, 1, Array("6", "Check Validation and Refresh Log after every refresh before using the results; save the workbook after review.")
    PutRow ws, 13, 1, Array("Refresh status")
    PutRow ws, 15, 1, Array("Current version", "VBA module " & MODULE_VERSION & " installed. Run Refresh all teams after checking Configuration.")
    PutRow ws, 16, 1, Array("Configuration changes", "Amount-threshold formulas update immediately; radar benchmarks and other configuration changes take effect after Refresh all teams.")
    PutRow ws, 17, 1, Array("Refresh button", "Refresh reads reports without saving them, validates before commit, and creates a recovery copy.")
    PutRow ws, 18, 1, Array("Production use", "Replace all fictional configuration and validate exact source headers against sanitized workbooks.")
    PutRow ws, 21, 1, Array("Dashboard and summary rules")
    PutRow ws, 23, 1, Array("Amount threshold", "For Chase, Chase CAD, Internal, and HSBC Internal, include records where the absolute amount meets the editable Amount Threshold on Configuration.")
    PutRow ws, 24, 1, Array("OMS issues", "Extract the two-digit issue from positions 6-7 of Issue. Add or activate OmsIssueDefinitionsTable rows to track more issues and their nature.")
    PutRow ws, 25, 1, Array("eTR", "Count only rows whose mapped Approval Date in column A is a valid date. Comment-only rows are ignored; matching eTR files are appended.")
    PutRow ws, 26, 1, Array("Radar colours", "White chart background with pale report sectors. Team actual is solid orange; the benchmark is dashed blue. Tables remain neutral; grey identifies non-BAU reports.")
    PutRow ws, 27, 1, Array("Membership exceptions", "When observed ownership differs from configured membership, include a review row under the observed manager and keep its entries in manager totals. Check Team Members settings.")
    PutRow ws, 28, 1, Array("Radar benchmark", "For each metric, the benchmark averages the lowest 20% of eligible employee values (headcount rounded up). The adjacent table shows raw actual and benchmark counts/member, ages and absolute amounts. The radar is capped at 8.0x. Refresh after input changes.")
    PutRow ws, 30, 1, Array("Data flow")
    PutRow ws, 32, 1, Array("Standard folder", ChrW(&H2192), "Report mappings", ChrW(&H2192), _
        "Normalized data", ChrW(&H2192), "Combined dashboard and validation", "")
    ColumnWidths ws, "A:A", 18, "B:B", 72, "C:C", 18, "D:D", 6, "E:E", 20, "F:F", 6, "G:G", 28, "H:H", 5
    With ws.Range("A5:A10")
        .Interior.Color = RGB(217, 234, 247): .Font.Size = 11
        .Font.Bold = True: .Font.Color = RGB(31, 78, 120): .HorizontalAlignment = xlCenter
    End With
    ws.Range("A15:A18,A23:A28").Font.Bold = True
    ws.Range("A15:B18,A23:B28,B5:B10,G32").WrapText = True
    For Each row In Array(13, 21, 30)
        With ws.Range(ws.Cells(row, 1), ws.Cells(row, 8))
            .Interior.Color = RGB(217, 234, 247): .Font.Size = 11
            .Font.Bold = True: .Font.Color = RGB(31, 78, 120): .RowHeight = 24
        End With
    Next row
    With ws.Range("A32:H32")
        .Interior.Color = RGB(234, 242, 248): .Font.Bold = True
        .Font.Color = RGB(31, 78, 120): .HorizontalAlignment = xlCenter: .RowHeight = 36
        .BorderAround Weight:=xlThin, Color:=RGB(183, 201, 214)
    End With
    FitReadMeRows ws
End Sub

Private Sub ApplyWorkbookPresentation()
    Dim name As Variant, ws As Worksheet
    For Each name In Array("Dashboard", "Team Summary", "Employee Summary", "Configuration", _
                          "Report Mapping", "Data >>", "Normalized Data", "Validation", "ReadMe", "Refresh Log")
        Set ws = SheetOrNothing(ThisWorkbook, CStr(name))
        If Not ws Is Nothing Then
            ' A protected mapping sheet can still be read during refresh.
            If Not ws.ProtectContents Then
                ws.UsedRange.Font.Name = "Arial"
                If name = "Configuration" Then ws.Range("A6:F9").Font.Size = 10
            End If
            Select Case name
                Case "Dashboard": ws.Tab.Color = RGB(31, 78, 120)
                Case "Team Summary", "Employee Summary": ws.Tab.Color = RGB(91, 155, 213)
                Case "Configuration": ws.Tab.Color = RGB(255, 217, 102)
                Case "Report Mapping": ws.Tab.Color = RGB(165, 165, 165)
                Case "Data >>": ws.Tab.Color = RGB(127, 140, 141)
                Case "ReadMe": ws.Tab.Color = RGB(112, 173, 71)
                Case Else: ws.Tab.ColorIndex = xlColorIndexNone
            End Select
        End If
    Next name
End Sub

Private Sub RequireMacroWorkbook()
    If ThisWorkbook.ReadOnly Then Fail "Open an editable dashboard workbook before running this macro."
    If Len(ThisWorkbook.Path) = 0 Then Fail "Save the dashboard locally as an Excel Macro-Enabled Workbook (.xlsm) first."
    If ThisWorkbook.FileFormat <> xlOpenXMLWorkbookMacroEnabled Then _
        Fail "Save As Excel Macro-Enabled Workbook (.xlsm) before installing or refreshing."
    If ThisWorkbook.Date1904 Then Fail "The dashboard template must use the standard 1900 date system. Source workbooks may use either system."
End Sub

Private Function RequiredSheet(ByVal name As String) As Worksheet
    Set RequiredSheet = SheetOrNothing(ThisWorkbook, name)
    If RequiredSheet Is Nothing Then Fail "Required dashboard tab is missing: " & name
End Function

Private Function CreateRecoveryCopy() As String
    Dim fso As Object, folder As String, stem As String, path As String, suffix As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.BuildPath(ThisWorkbook.Path, "Dashboard_Backups")
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    stem = fso.GetBaseName(ThisWorkbook.Name) & "_before_refresh_" & Format$(Now, "yyyymmdd_hhnnss")
    path = fso.BuildPath(folder, stem & ".xlsm")
    Do While fso.FileExists(path)
        suffix = suffix + 1
        path = fso.BuildPath(folder, stem & "_" & CStr(suffix) & ".xlsm")
    Loop
    ThisWorkbook.SaveCopyAs Filename:=path
    CreateRecoveryCopy = path
End Function

Public Sub RepairOperationsDashboardAppearance()
    Dim ws As Worksheet, name As Variant, backupPath As String, stage As String
    On Error GoTo Failed
    stage = "checking workbook"
    RequireMacroWorkbook
    If ThisWorkbook.ProtectStructure Then Fail "Unprotect workbook structure before appearance repair."
    For Each name In Array("Dashboard", "Team Summary", "Employee Summary", "Configuration", _
                          "Report Mapping", "Data >>", "Normalized Data", "Validation", "ReadMe", "Refresh Log")
        Set ws = SheetOrNothing(ThisWorkbook, CStr(name))
        If Not ws Is Nothing Then
            If ws.ProtectContents Then Fail "Unprotect the dashboard sheet before appearance repair: " & name
        End If
    Next name
    Set ws = RequiredSheet("Dashboard")
    Set ws = RequiredSheet("Configuration")
    Set ws = RequiredSheet("Report Mapping")
    stage = "creating recovery copy"
    backupPath = CreateRecoveryCopy()
    stage = "creating ReadMe"
    EnsureSheetExists "ReadMe"
    SeedReadMeSheet RequiredSheet("ReadMe")
    stage = "applying tab colours and fonts"
    ApplyWorkbookPresentation
    RequiredSheet("ReadMe").Activate
    MsgBox "Appearance repair " & MODULE_VERSION & " completed for " & ThisWorkbook.Name & "." & vbCrLf & _
           "ReadMe rebuilt; tab colours and Arial fonts applied. Configuration values are retained." & vbCrLf & _
           "Recovery copy: " & backupPath, vbInformation, "Dashboard appearance repaired"
    Exit Sub
Failed:
    MsgBox "Appearance repair " & MODULE_VERSION & " stopped while " & stage & "." & vbCrLf & _
           "Workbook: " & ThisWorkbook.Name & vbCrLf & "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Dashboard appearance repair stopped"
End Sub

' ========================================================================
' STAGE 3: WORKBOOK INSTALLATION AND SETTINGS
' ========================================================================
' Run InstallOperationsDashboard. Check all nine tabs, configuration tables and buttons. Choose the 
' dummy-report folder. Refresh is not available yet.
' Procedure typing order:
' 01. SectionBand
' 02. HeaderStyle
' 03. AddConfigTable
' 04. SetTitle
' 05. NoteBand
' 06. SeedConfigurationSheet
' 07. SeedReportMappingSheet
' 08. EnsureWorkbookScaffold
' 09. SeedField
' 10. DefaultFieldRows
' 11. SettingCell
' 12. TableOrNothing
' 13. AddActionButton
' 14. LastUsedRow
' 15. InstallOperationsDashboard
' 16. ChooseOperationsReportFolder

Private Sub SectionBand(ByVal ws As Worksheet, ByVal row As Long, ByVal lastCol As Long, _
                        ByVal title As String, Optional ByVal regionBand As Boolean = False)
    With ws.Range(ws.Cells(row, 1), ws.Cells(row, lastCol))
        .Merge: .Value2 = LiteralValue(title)
        .Font.Name = "Arial": .Font.Size = IIf(regionBand, 12, 10): .Font.Bold = True
        .Interior.Color = IIf(regionBand, COLOR_NAVY, RGB(221, 235, 247))
        .Font.Color = IIf(regionBand, vbWhite, COLOR_NAVY)
        .RowHeight = 24
    End With
End Sub

Private Sub HeaderStyle(ByVal target As Range)
    With target
        .Font.Name = "Arial": .Font.Size = 10: .Font.Bold = True: .Font.Color = vbWhite
        .Interior.Color = COLOR_NAVY: .WrapText = True: .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).Color = COLOR_LINE
    End With
End Sub

Private Sub AddConfigTable(ByVal ws As Worksheet, ByVal startRow As Long, ByVal title As String, _
                           ByVal name As String, ByVal headers As Variant, ByVal rows As Collection)
    Dim c As Long, r As Long, values As Variant, lo As ListObject
    If ws.ProtectContents Then Fail "Unprotect Report Mapping before installation."
    SectionBand ws, startRow, UBound(headers) - LBound(headers) + 1, title
    PutRow ws, startRow + 1, 1, headers
    r = startRow + 2
    For Each values In rows
        PutRow ws, r, 1, values
        r = r + 1
    Next values
    c = UBound(headers) - LBound(headers) + 1
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(r - 1, c)), , xlYes)
    lo.Name = name: lo.TableStyle = "TableStyleMedium2"
    HeaderStyle lo.HeaderRowRange
    lo.HeaderRowRange.RowHeight = 32
    lo.DataBodyRange.Font.Color = RGB(0, 70, 160)
    lo.DataBodyRange.WrapText = True
    lo.DataBodyRange.Rows.AutoFit
    lo.Range.Font.Name = "Arial": lo.Range.Font.Size = 10
End Sub

Private Sub SetTitle(ByVal ws As Worksheet, ByVal title As String, ByVal lastCol As Long)
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol))
        .Merge: .Value2 = title
        .Font.Name = "Arial": .Font.Size = 16: .Font.Bold = True
        .Font.Color = vbWhite: .Interior.Color = COLOR_NAVY
        .RowHeight = 30
    End With
End Sub

Private Sub NoteBand(ByVal ws As Worksheet, ByVal row As Long, ByVal lastCol As Long, ByVal text As String)
    With ws.Range(ws.Cells(row, 1), ws.Cells(row, lastCol))
        .Merge: .Value2 = LiteralValue(text): .WrapText = True
        .Font.Name = "Arial": .Font.Size = 10: .Font.Color = RGB(89, 89, 89)
        .Interior.Color = RGB(242, 246, 250): .RowHeight = 30
    End With
End Sub

Private Sub SeedConfigurationSheet(ByVal ws As Worksheet)
    Dim rows As Collection
    SetTitle ws, "Configuration", 11
    NoteBand ws, 4, 11, "Replace the fictional starter rows before production use. Configuration is applied during refresh."
    PutRow ws, 6, 1, Array("Source folder", ThisWorkbook.Path, "", "Amount threshold", 1000, "Absolute credit/debit amount")
    PutRow ws, 7, 1, Array("Reporting date", Date, "", "Used by Chase, Chase CAD, Internal, and HSBC Internal.")
    PutRow ws, 8, 1, Array("File selection rule", "Latest matching file; append all eTR files")
    PutRow ws, 9, 1, Array("Refresh scope", "All active teams")
    ws.Range("B6:B7,E6").Interior.Color = COLOR_INPUT
    ws.Range("B7").NumberFormat = "mm/dd/yyyy"

    Set rows = New Collection
    rows.Add Array("Yes", "NY", "NY-A", "NY Alpha", "Morgan Lee", "Yes", "Yes", "Yes", "Yes", "Yes", "Yes")
    rows.Add Array("Yes", "Toronto", "TOR-A", "Toronto Alpha", "Jordan Patel", "Yes", "No", "No", "Yes", "Yes", "Yes")
    AddConfigTable ws, 11, "Teams and applicable reports", "TeamsTable", _
        Array("Active", "Region", "Team ID", "Team Name", "Manager", "OMS", "Chase", "Chase CAD", "Internal", "Internal HSBC", "eTR"), rows

    Set rows = New Collection
    rows.Add Array("Yes", "M001", "NY-A", "Tom", DateSerial(Year(Date), 1, 1), "", "Starter member")
    rows.Add Array("Yes", "M002", "NY-A", "Amy", DateSerial(Year(Date), 1, 1), "", "Starter member")
    rows.Add Array("Yes", "M011", "TOR-A", "Kate", DateSerial(Year(Date), 1, 1), "", "Starter member")
    rows.Add Array("Yes", "M014", "TOR-A", "Bob", DateSerial(Year(Date), 1, 1), "", "Starter member")
    AddConfigTable ws, 18, "Team members", "MembersTable", _
        Array("Active", "Member ID", "Team ID", "Member Name", "Effective From", "Effective To", "Notes"), rows

    ws.Columns("A:K").AutoFit
End Sub

Private Sub SeedReportMappingSheet(ByVal ws As Worksheet)
    Dim rows As Collection
    SetTitle ws, "Report Mapping", 9
    NoteBand ws, 4, 9, "Technical settings. Update filename patterns, worksheet names, account tabs, and aliases when source layouts change."

    Set rows = New Collection
    rows.Add Array("Yes", "OMS_v1", "OMS", "Both", "OMS_Report_*.xlsx", "Sheet1", 5, "No", "Employee alias")
    rows.Add Array("Yes", "CHASE_v1", "Chase", "NY", "Chase_Report_20*.xlsx", "Current Day", 2, "No", "Employee alias")
    rows.Add Array("Yes", "CHASE_CAD_v1", "Chase CAD", "NY", "Chase_Report_CAD_*.xlsx", "Sheet1", 2, "No", "Employee alias")
    rows.Add Array("Yes", "INTERNAL_v1", "Internal", "Both", "Internal_Summary_*.xlsx", "Configured account tabs", 1, "No", "Employee alias")
    rows.Add Array("Yes", "INTERNAL_HSBC_v1", "Internal HSBC", "Both", "Internal_HSBC_Summary_*.xlsx", "Configured account tabs", 1, "No", "Employee alias")
    rows.Add Array("Yes", "ETR_v1", "eTR", "Both", "eTR_*.xlsx", "Sheet1", 7, "Yes", "File context + alias")
    AddConfigTable ws, 7, "Source definitions", "SourceDefinitionsTable", _
        Array("Active", "Layout ID", "Report", "Region Rule", "Filename Pattern", "Worksheet Rule", "Header Row", "Multiple Files", "Assignment Method"), rows

    Set rows = New Collection
    rows.Add Array("Yes", "Internal", "1026830", "CAD", "Standard", "Internal 1026830 - CAD Standard")
    rows.Add Array("Yes", "Internal", "4005716", "USD", "Standard", "Internal 4005716 - USD Standard")
    rows.Add Array("Yes", "Internal HSBC", "1164458", "CAD", "HSBC", "HSBC Internal 1164458 - CAD")
    rows.Add Array("Yes", "Internal HSBC", "4061826", "USD", "HSBC", "HSBC Internal 4061826 - USD")
    AddConfigTable ws, 17, "Internal account tabs and dashboard labels", "AccountTabsTable", _
        Array("Active", "Report", "Account Tab", "Currency", "Account Type", "Dashboard Label"), rows

    Set rows = New Collection
    rows.Add Array("No", "eTR", "eTR_Toronto_Team_Alpha_*.xlsx", "Toronto", "TOR-A")
    rows.Add Array("No", "eTR", "eTR_NY_Team_Beta_*.xlsx", "NY", "NY-A")
    AddConfigTable ws, 26, "Team-specific files", "FileAssignmentsTable", _
        Array("Active", "Report", "Filename Pattern", "Region", "Team ID"), rows

    Set rows = New Collection
    rows.Add Array("Yes", "OMS", "*", "Tom", "M001", "")
    rows.Add Array("Yes", "OMS", "*", "Bob", "M014", "")
    rows.Add Array("Yes", "Chase", "*", "Tom", "M001", "")
    rows.Add Array("Yes", "Chase CAD", "*", "Tom", "M001", "")
    rows.Add Array("Yes", "Internal", "*", "Bob", "M014", "")
    rows.Add Array("Yes", "Internal HSBC", "*", "Bob", "M014", "")
    rows.Add Array("Yes", "eTR", "*Toronto*", "Kate", "M011", "Toronto eTR file")
    rows.Add Array("Yes", "eTR", "*NY*", "Tom", "M001", "NY eTR file")
    AddConfigTable ws, 34, "Employee aliases", "AliasesTable", _
        Array("Active", "Report", "Source Context", "Source Value", "Member ID", "Notes"), rows

    Set rows = New Collection
    rows.Add Array("Yes", 10, "CCMS & ACBS facility discrepancy", "Editable issue definition")
    rows.Add Array("Yes", 12, "CCMS & ACBS facility discrepancy", "Editable issue definition")
    rows.Add Array("Yes", 19, "Past due", "Editable issue definition")
    rows.Add Array("Yes", 20, "Past due loan fee", "Matches the supplied issue 20 example")
    rows.Add Array("Yes", 21, "Past due", "Editable issue definition")
    rows.Add Array("Yes", 22, "Past due", "Editable issue definition")
    rows.Add Array("No", "", "", "Available for another OMS issue")
    AddConfigTable ws, 48, "OMS issue definitions and available dashboard slots", "OmsIssueDefinitionsTable", _
        Array("Active", "Issue Code", "Issue Nature", "Notes"), rows

    ws.Columns("A:I").AutoFit
End Sub

Private Sub EnsureWorkbookScaffold()
    Dim name As Variant, ws As Worksheet
    For Each name In Array("Dashboard", "Team Summary", "Employee Summary", "Configuration", _
                          "Report Mapping", "Data >>", "Normalized Data", "Validation", "ReadMe")
        EnsureSheetExists CStr(name)
    Next name
    For Each name In Array("Configuration", "Report Mapping", "Data >>")
        Set ws = RequiredSheet(CStr(name))
        If WorksheetFunction.CountA(ws.Cells) = 0 Then
            Select Case name
                Case "Configuration": SeedConfigurationSheet ws
                Case "Report Mapping": SeedReportMappingSheet ws
                Case "Data >>"
                    SetTitle ws, "Data >>", 5
                    NoteBand ws, 4, 5, "System output sheets follow this divider. Edit Configuration and Report Mapping only."
            End Select
        End If
    Next name
End Sub

Private Sub SeedField(ByVal rows As Collection, ByVal layoutID As String, ByVal field As String, _
                      ByVal headers As String, ByVal required As Boolean, Optional ByVal defaultValue As String = "")
    rows.Add Array("Yes", layoutID, field, headers, 0, IIf(required, "Yes", "No"), defaultValue)
End Sub

Private Function DefaultFieldRows() As Collection
    Dim rows As New Collection, layout As Variant
    SeedField rows, "OMS_v1", "Reference", "SRF", True
    SeedField rows, "OMS_v1", "Employee", "Servicing User", True
    SeedField rows, "OMS_v1", "ItemDate", "Start Date", False
    SeedField rows, "OMS_v1", "AgeDays", "Age (days)|Age Days|Age", False
    SeedField rows, "OMS_v1", "Status", "Status", False, "Issue"
    SeedField rows, "OMS_v1", "Client", "Client|Borrower", False
    SeedField rows, "OMS_v1", "Issue", "Issue", True
    For Each layout In Array("CHASE_v1", "CHASE_CAD_v1")
        SeedField rows, CStr(layout), "Reference", IIf(layout = "CHASE_v1", "UID", "ACCOUNT|UID"), True
        SeedField rows, CStr(layout), "Employee", "Admin Name|Analyst", True
        SeedField rows, CStr(layout), "ItemDate", "POST|VALUE", False
        SeedField rows, CStr(layout), "AgeDays", IIf(layout = "CHASE_v1", "Days OS|Days O|Days Outstanding", "Days O|Days OS|Days Outstanding"), False
        SeedField rows, CStr(layout), "Amount", IIf(layout = "CHASE_v1", "AMT|AMOUNT", "AMOUNT|AMT"), True
        SeedField rows, CStr(layout), "Currency", "", False, IIf(layout = "CHASE_v1", "USD", "CAD")
        SeedField rows, CStr(layout), "Client", IIf(layout = "CHASE_v1", "Borrower|Service Name", "Service Name|Borrower"), False
        SeedField rows, CStr(layout), "Comment", "DETAIL|Comment", False
        SeedField rows, CStr(layout), "Status", "Comment", False
    Next layout
    For Each layout In Array("INTERNAL_v1", "INTERNAL_HSBC_v1")
        SeedField rows, CStr(layout), "Reference", "UID", True
        SeedField rows, CStr(layout), "Employee", "Analyst|Admin Name", True
        SeedField rows, CStr(layout), "ItemDate", "Date", False
        SeedField rows, CStr(layout), "AgeDays", "Day|Days|Age Days", False
        SeedField rows, CStr(layout), "Amount", "Balance|Amount", True
        SeedField rows, CStr(layout), "Currency", "Currency|Current", False
        SeedField rows, CStr(layout), "Client", "Client Name|Borrower", False
        SeedField rows, CStr(layout), "Comment", "Description", False
        SeedField rows, CStr(layout), "Status", "Disposition on Transaction", False
    Next layout
    SeedField rows, "ETR_v1", "Reference", "SRF|CCMS FacId", True
    SeedField rows, "ETR_v1", "Employee", "Servicing User|Admin User", True
    SeedField rows, "ETR_v1", "ItemDate", "Approval Date", False
    SeedField rows, "ETR_v1", "Client", "Borrower|Single Name", False
    SeedField rows, "ETR_v1", "Status", "", False, "New"
    For Each layout In Array("OMS_v1", "CHASE_v1", "CHASE_CAD_v1", "INTERNAL_v1", "INTERNAL_HSBC_v1", "ETR_v1")
        ' Disabled-by-default optional ownership fields. Map only when authoritative.
        SeedField rows, CStr(layout), "Team", "", False
        SeedField rows, CStr(layout), "Manager", "", False
        SeedField rows, CStr(layout), "Region", "", False
    Next layout
    Set DefaultFieldRows = rows
End Function

Private Function SettingCell(ByVal label As String) As Range
    Dim ws As Worksheet, cell As Range
    Set ws = RequiredSheet("Configuration")
    For Each cell In ws.Range("A1:F15").Cells
        If TextOf(cell.Value2) = label Then
            Set SettingCell = cell.Offset(0, 1)
            Exit Function
        End If
    Next cell
    Fail "Configuration setting is missing: " & label
End Function

Private Function TableOrNothing(ByVal name As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = name Then Set TableOrNothing = lo: Exit Function
        Next lo
    Next ws
End Function

Private Sub AddActionButton(ByVal ws As Worksheet, ByVal name As String, ByVal anchor As String, _
                            ByVal caption As String, ByVal macroName As String)
    Dim button As Object, target As Range
    Set target = ws.Range(anchor)
    On Error Resume Next
    Set button = ws.Buttons(name)
    On Error GoTo 0
    If button Is Nothing Then
        Set button = ws.Buttons.Add(target.Left, target.Top, 155, 25)
        button.Name = name
    End If
    button.Caption = caption
    button.Font.Name = "Arial": button.Font.Size = 10
    button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
    button.Left = target.Left: button.Top = target.Top
    button.Width = 155: button.Height = 25
    If name = "TeamSelectionButton" Then
        button.Placement = xlMove
    Else
        button.Placement = xlFreeFloating
    End If
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If Not found Is Nothing Then LastUsedRow = found.Row
End Function

Public Sub InstallOperationsDashboard()
    Dim ws As Worksheet, lo As ListObject, startRow As Long
    Dim rows As Collection, backupPath As String, name As Variant
    On Error GoTo Failed
    RequireMacroWorkbook
    If ThisWorkbook.ProtectStructure Then Fail "Unprotect workbook structure before installation."
    For Each name In Array("Dashboard", "Team Summary", "Employee Summary", "Configuration", _
                          "Report Mapping", "Data >>", "Normalized Data", "Validation", "ReadMe", "Refresh Log")
        Set ws = SheetOrNothing(ThisWorkbook, CStr(name))
        If Not ws Is Nothing Then
            If ws.ProtectContents Then Fail "Unprotect the dashboard sheet before installation: " & name
        End If
    Next name
    backupPath = CreateRecoveryCopy()
    EnsureWorkbookScaffold
    SeedReadMeSheet RequiredSheet("ReadMe")
    ApplyWorkbookPresentation
    Set ws = RequiredSheet("Report Mapping")
    Set lo = TableOrNothing("FieldMappingsTable")
    If lo Is Nothing Then
        Set rows = DefaultFieldRows()
        startRow = LastUsedRow(ws) + 3
        AddConfigTable ws, startRow, "Source column mappings", "FieldMappingsTable", _
            Array("Active", "Layout ID", "Field", "Header Aliases", "Column Fallback", "Required", "Default Value"), rows
    End If
    Set lo = TableOrNothing("ManagerAliasesTable")
    If lo Is Nothing Then
        Set rows = New Collection
        rows.Add Array("", "", "", "", "", "Enable only for authoritative source-manager fields")
        startRow = LastUsedRow(ws) + 3
        AddConfigTable ws, startRow, "Source manager aliases", "ManagerAliasesTable", _
            Array("Active", "Report", "Source Context", "Source Manager", "Team ID", "Notes"), rows
    End If
    Set ws = RequiredSheet("Configuration")
    With SettingCell("Amount threshold").Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
             Operator:=xlGreaterEqual, Formula1:="0"
        .IgnoreBlank = False
        .ErrorTitle = "Invalid threshold"
        .ErrorMessage = "Enter a non-negative numeric amount."
        .ShowError = True
    End With
    AddActionButton RequiredSheet("Dashboard"), "OpsRefreshButton", "Q5", _
        "Refresh all teams", "RefreshOperations"
    AddActionButton ws, "OpsFolderButton", "H6", "Choose report folder", "ChooseOperationsReportFolder"
    ApplyWorkbookPresentation
    MsgBox "Installation complete. Review the new column mappings on Report Mapping, choose the source folder, " & _
           "and save this .xlsm workbook." & vbCrLf & "Recovery copy: " & backupPath, vbInformation, "Operations dashboard"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Installation stopped"
End Sub

Public Sub ChooseOperationsReportFolder()
    Dim picker As FileDialog
    On Error GoTo Failed
    Set picker = Application.FileDialog(msoFileDialogFolderPicker)
    picker.Title = "Select the folder containing current outstanding reports"
    picker.AllowMultiSelect = False
    If picker.Show = -1 Then SettingCell("Source folder").Value2 = picker.SelectedItems(1)
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Folder selection"
End Sub

' ========================================================================
' STAGE 4: CONFIGURATION AND BUSINESS RULES
' ========================================================================
' Run RunOperationsSelfTests, then TypingCheckConfiguration. It prints active team, member, 
' source-layout and account-tab counts.
' Procedure typing order:
' 01. InitializeRun
' 02. DigitsOnly
' 03. ParseDateValue
' 04. MemberEffective
' 05. BuildAmountMetrics
' 06. TableRows
' 07. YesFlag
' 08. ActiveRows
' 09. NewStats
' 10. ReportNames
' 11. CanonReport
' 12. CanonRegion
' 13. IsKnownField
' 14. CollectionHasLayout
' 15. ParseAmount
' 16. LoadConfiguration
' 17. WildMatch
' 18. PatternSpecificity
' 19. AliasCandidates
' 20. UniqueKey
' 21. Quoted
' 22. Criterion
' 23. SheetRuleMatches
' 24. ExtractIssueCode
' 25. AssertTest
' 26. WeightedMean
' 27. TestBusinessRules
' 28. RunOperationsSelfTests

Private Sub InitializeRun()
    Set mTeams = NewMap()
    Set mMembers = NewMap()
    Set mFields = NewMap()
    Set mStats = NewMap()
    Set mReadKeys = NewMap()
    Set mFingerprints = NewMap()
    Set mParticipants = NewMap()
    Set mRadarBenchmarks = NewMap()
    Set mTeamOrder = New Collection
    Set mSources = New Collection
    Set mAccounts = New Collection
    Set mIssues = New Collection
    Set mAliases = New Collection
    Set mManagerAliases = New Collection
    Set mFileAssignments = New Collection
    Set mRecords = New Collection
    Set mLogs = New Collection
    Set mAmountMetrics = New Collection
    mWarnings = 0
    mRunID = Format$(Now, "yyyymmdd_hhnnss")
End Sub

Private Function DigitsOnly(ByVal text As String) As Boolean
    Dim i As Long, ch As String
    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    DigitsOnly = True
End Function

Private Function ParseDateValue(ByVal value As Variant, ByVal date1904 As Boolean) As Variant
    Dim text As String, parts As Variant, i As Long, serial As Double
    Dim yearNumber As Long, monthNumber As Long, dayNumber As Long, result As Date
    ParseDateValue = Empty
    If IsEmpty(value) Or IsNull(value) Or IsError(value) Then Exit Function
    On Error GoTo InvalidDate
    If VarType(value) = vbDate Then ParseDateValue = DateValue(CDate(value)): Exit Function
    If VarType(value) <> vbString Then
        If Not IsNumeric(value) Or VarType(value) = vbBoolean Then Exit Function
        serial = CDbl(value)
        If date1904 Then serial = serial + 1462
        If serial < 1 Or serial > 2958465 Then Exit Function
        If Not date1904 Then
            If serial = 60 Then Exit Function ' Excel's fictitious 1900-02-29 is not a VBA date.
            If serial < 60 Then serial = serial + 1
        End If
        ParseDateValue = DateValue(CDate(Fix(serial)))
        Exit Function
    End If
    text = TextOf(value)
    If Len(text) = 0 Then Exit Function
    parts = Split(Replace(text, "-", "/"), "/")
    If UBound(parts) <> 2 Then Exit Function
    For i = 0 To 2
        If Len(parts(i)) = 0 Or Len(parts(i)) > 4 Then Exit Function
        If Not DigitsOnly(CStr(parts(i))) Then Exit Function
    Next i
    If Len(parts(0)) = 4 Then
        yearNumber = CLng(parts(0)): monthNumber = CLng(parts(1)): dayNumber = CLng(parts(2))
    ElseIf Len(parts(2)) = 4 Then
        monthNumber = CLng(parts(0)): dayNumber = CLng(parts(1)): yearNumber = CLng(parts(2))
    Else
        Exit Function
    End If
    If yearNumber < 1900 Or yearNumber > 9999 Or monthNumber < 1 Or monthNumber > 12 Or dayNumber < 1 Or dayNumber > 31 Then Exit Function
    result = DateSerial(yearNumber, monthNumber, dayNumber)
    If Year(result) <> yearNumber Or Month(result) <> monthNumber Or Day(result) <> dayNumber Then Exit Function
    ParseDateValue = result
    Exit Function
InvalidDate:
    ParseDateValue = Empty
End Function

Private Function MemberEffective(ByVal row As Object) As Boolean
    Dim fromDate As Variant, toDate As Variant
    fromDate = ParseDateValue(ValueOf(row, "Effective From"), ThisWorkbook.Date1904)
    toDate = ParseDateValue(ValueOf(row, "Effective To"), ThisWorkbook.Date1904)
    If Len(TextOf(ValueOf(row, "Effective From"))) > 0 And IsEmpty(fromDate) Then Fail "Invalid member Effective From date."
    If Len(TextOf(ValueOf(row, "Effective To"))) > 0 And IsEmpty(toDate) Then Fail "Invalid member Effective To date."
    If Not IsEmpty(fromDate) And Not IsEmpty(toDate) Then
        If CDate(toDate) < CDate(fromDate) Then Fail "Member Effective To is before Effective From."
    End If
    MemberEffective = True
    If Not IsEmpty(fromDate) Then If mReportDate < CDate(fromDate) Then MemberEffective = False
    If Not IsEmpty(toDate) Then If mReportDate > CDate(toDate) Then MemberEffective = False
End Function

Private Sub BuildAmountMetrics()
    Dim account As Object
    mAmountMetrics.Add NewMap("report", "Chase", "account", "", "label", "Chase USD Report")
    mAmountMetrics.Add NewMap("report", "Chase CAD", "account", "", "label", "Chase CAD Report")
    For Each account In mAccounts
        mAmountMetrics.Add NewMap("report", account("Report"), "account", account("Account Tab"), "label", account("Dashboard Label"))
    Next account
End Sub

Private Function TableRows(ByVal name As String) As Collection
    Dim lo As ListObject, result As New Collection, row As Object
    Dim data As Variant, headers As Variant, r As Long, c As Long, hasData As Boolean
    Set lo = TableOrNothing(name)
    If lo Is Nothing Then Fail "Configuration table is missing: " & name & ". Run InstallOperationsDashboard first."
    If lo.DataBodyRange Is Nothing Then Set TableRows = result: Exit Function
    data = lo.DataBodyRange.Value2: headers = lo.HeaderRowRange.Value2
    For r = 1 To UBound(data, 1)
        Set row = NewMap(): hasData = False
        row.Add "@row", lo.DataBodyRange.Row + r - 1
        For c = 1 To UBound(data, 2)
            If IsError(data(r, c)) Then Fail "Configuration-cell error in " & name & ", row " & row("@row")
            row.Add TextOf(headers(1, c)), data(r, c)
            If Len(TextOf(data(r, c))) > 0 Then hasData = True
        Next c
        If hasData Then result.Add row
    Next r
    Set TableRows = result
End Function

Private Function YesFlag(ByVal value As Variant) As Boolean
    YesFlag = (TextOf(value) = "Yes")
End Function

Private Function ActiveRows(ByVal name As String) As Collection
    Dim result As New Collection, rows As Collection, row As Object
    Set rows = TableRows(name)
    For Each row In rows
        If YesFlag(ValueOf(row, "Active")) Then result.Add row
    Next row
    Set ActiveRows = result
End Function

Private Function NewStats() As Object
    Set NewStats = NewMap("files", NewMap(), "raw", 0, "normalized", 0, "duplicates", 0)
End Function

Private Function ReportNames() As Variant
    ReportNames = Array("OMS", "eTR", "Chase", "Chase CAD", "Internal", "Internal HSBC")
End Function

Private Function CanonReport(ByVal value As String) As String
    Dim report As Variant
    For Each report In ReportNames()
        If CleanName(value) = report Then CanonReport = CStr(report): Exit Function
    Next report
    Fail "Unknown report name in configuration: " & value
End Function

Private Function CanonRegion(ByVal value As String) As String
    Select Case NormalizeHeader(value)
        Case "ny", "newyork": CanonRegion = "NY"
        Case "toronto", "to": CanonRegion = "Toronto"
        Case Else: Fail "Region must be NY or Toronto: " & value
    End Select
End Function

Private Function IsKnownField(ByVal field As String) As Boolean
    Dim item As Variant
    For Each item In Array("Reference", "Employee", "ItemDate", "AgeDays", "Amount", "Currency", "Client", "Comment", "Status", "Issue", "Team", "Manager", "Region")
        If field = item Then IsKnownField = True: Exit Function
    Next item
End Function

Private Function CollectionHasLayout(ByVal rows As Collection, ByVal id As String) As Boolean
    Dim row As Object
    For Each row In rows
        If row("Layout ID") = id Then CollectionHasLayout = True: Exit Function
    Next row
End Function

Private Function ParseAmount(ByVal value As Variant) As Variant
    Dim text As String, negative As Boolean, sign As Double, i As Long, ch As String
    Dim digits As Long, dots As Long, groups As Variant, integerPart As String, decimalAt As Long
    ParseAmount = Empty
    If IsEmpty(value) Or IsNull(value) Or IsError(value) Then Exit Function
    If VarType(value) <> vbString Then
        If IsNumeric(value) And VarType(value) <> vbBoolean And VarType(value) <> vbDate Then ParseAmount = CDbl(value)
        Exit Function
    End If
    text = TextOf(value)
    If Len(text) = 0 Then Exit Function
    text = Replace(text, ChrW(160), ""): text = Replace(text, " ", "")
    text = Replace(text, "$", "")
    If Len(text) >= 2 Then
        If Left$(text, 1) = "(" And Right$(text, 1) = ")" Then
            negative = True: text = Mid$(text, 2, Len(text) - 2)
        End If
    End If
    sign = 1
    If Left$(text, 1) = "-" Then
        If negative Then Exit Function
        sign = -1: text = Mid$(text, 2)
    ElseIf Left$(text, 1) = "+" Then
        text = Mid$(text, 2)
    End If
    If negative Then sign = -1
    If Len(text) = 0 Then Exit Function
    decimalAt = InStr(text, ".")
    If decimalAt > 0 Then integerPart = Left$(text, decimalAt - 1) Else integerPart = text
    If InStr(text, ",") > 0 Then
        If decimalAt > 0 Then If InStr(decimalAt, text, ",") > 0 Then Exit Function
        groups = Split(integerPart, ",")
        If Len(groups(0)) < 1 Or Len(groups(0)) > 3 Then Exit Function
        For i = 1 To UBound(groups)
            If Len(groups(i)) <> 3 Then Exit Function
        Next i
        text = Replace(text, ",", "")
    End If
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits + 1
        ElseIf ch = "." Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        Else
            Exit Function
        End If
    Next i
    If digits = 0 Then Exit Function
    On Error GoTo InvalidAmount
    ParseAmount = sign * Val(text)
    Exit Function
InvalidAmount:
    ParseAmount = Empty
End Function

Private Sub LoadConfiguration()
    Dim rows As Collection, row As Object, flags As Object, d As Object
    Dim fields As Collection, names As Object, usedCodes As Object
    Dim report As Variant, id As String, layoutID As String, v As Variant
    Dim amount As Variant, active As Boolean, rDate As Variant
    Set names = NewMap()
    Set usedCodes = NewMap()
    v = SettingCell("Amount threshold").Value2
    If VarType(v) = vbString Or VarType(v) = vbBoolean Then Fail "Configuration: Amount threshold must be a numeric cell, not text."
    amount = ParseAmount(v)
    If IsEmpty(amount) Then Fail "Configuration: Amount threshold must be numeric."
    If CDbl(amount) < 0 Then Fail "Configuration: Amount threshold cannot be negative."
    mAmountThreshold = CDbl(amount)
    rDate = ParseDateValue(SettingCell("Reporting date").Value2, ThisWorkbook.Date1904)
    If IsEmpty(rDate) Then Fail "Configuration: enter a valid Reporting date."
    mReportDate = CDate(rDate)
    mFolder = TextOf(SettingCell("Source folder").Value2)
    If Len(mFolder) = 0 Then Fail "Choose a source folder on Configuration."
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(mFolder) Then Fail "Source folder not found: " & mFolder
    For Each row In ActiveRows("TeamsTable")
        id = RequiredText(row, "Team ID")
        If id = UNASSIGNED_TEAM Then Fail "Reserved team ID: " & id
        If mTeams.Exists(id) Then Fail "Duplicate active Team ID: " & id
        Set d = NewMap("id", id, "name", RequiredText(row, "Team Name"), _
                       "manager", RequiredText(row, "Manager"), "region", CanonRegion(RequiredText(row, "Region")))
        If names.Exists(d("name")) Then Fail "Active Team Name must be unique: " & d("name")
        names.Add d("name"), True
        Set flags = NewMap()
        For Each report In ReportNames()
            v = ValueOf(row, CStr(report))
            If TextOf(v) <> "Yes" And TextOf(v) <> "No" Then Fail "Set every report flag to Yes or No for team " & id
            flags.Add CStr(report), YesFlag(v)
        Next report
        d.Add "flags", flags
        mTeams.Add id, d
        mTeamOrder.Add id
    Next row
    If mTeams.Count = 0 Then Fail "Configuration has no active teams."
    For Each row In ActiveRows("MembersTable")
        If MemberEffective(row) Then
            id = RequiredText(row, "Member ID")
            If mMembers.Exists(id) Then Fail "Duplicate effective Member ID: " & id
            If Not mTeams.Exists(RequiredText(row, "Team ID")) Then Fail "Member " & id & " points to a missing/inactive team."
            Set d = NewMap("id", id, "name", RequiredText(row, "Member Name"), "team", RequiredText(row, "Team ID"))
            mMembers.Add id, d
        End If
    Next row
    For Each row In ActiveRows("SourceDefinitionsTable")
        layoutID = RequiredText(row, "Layout ID")
        If CollectionHasLayout(mSources, layoutID) Then Fail "Duplicate active Layout ID: " & layoutID
        row("Report") = CanonReport(RequiredText(row, "Report"))
        v = ValueOf(row, "Region Rule")
        If Len(TextOf(v)) > 0 And TextOf(v) <> "Both" Then v = CanonRegion(TextOf(v))
        row("Region Rule") = TextOf(v)
        RequiredText row, "Filename Pattern"
        RequiredText row, "Worksheet Rule"
        If Not IsNumeric(ValueOf(row, "Header Row")) Then Fail "Header Row must be a positive integer for " & layoutID
        If CDbl(row("Header Row")) < 1 Or CDbl(row("Header Row")) <> Fix(CDbl(row("Header Row"))) Then Fail "Invalid Header Row for " & layoutID
        If TextOf(ValueOf(row, "Multiple Files")) <> "Yes" And TextOf(ValueOf(row, "Multiple Files")) <> "No" Then Fail "Multiple Files must be Yes or No for " & layoutID
        mSources.Add row
        If Not mStats.Exists(row("Report")) Then mStats.Add row("Report"), NewStats()
    Next row
    If mSources.Count = 0 Then Fail "No active source definitions."
    For Each row In ActiveRows("AccountTabsTable")
        row("Report") = CanonReport(RequiredText(row, "Report"))
        If row("Report") <> "Internal" And row("Report") <> "Internal HSBC" Then Fail "AccountTabsTable supports internal reports only."
        RequiredText row, "Account Tab"
        RequiredText row, "Dashboard Label"
        For Each d In mAccounts
            If d("Report") = row("Report") And d("Account Tab") = row("Account Tab") Then Fail "Duplicate account mapping: " & row("Account Tab")
        Next d
        mAccounts.Add row
    Next row
    Set rows = TableRows("OmsIssueDefinitionsTable")
    For Each row In rows
        active = YesFlag(ValueOf(row, "Active"))
        row("Active") = IIf(active, "Yes", "No")
        If active Then
            v = ValueOf(row, "Issue Code")
            If Not IsNumeric(v) Then Fail "Active OMS issue codes must be integers from 0 to 99."
            If CDbl(v) < 0 Or CDbl(v) > 99 Or Fix(CDbl(v)) <> CDbl(v) Then Fail "Invalid OMS issue code."
            If usedCodes.Exists(CStr(CLng(v))) Then Fail "Duplicate active OMS issue code: " & v
            usedCodes.Add CStr(CLng(v)), True
            RequiredText row, "Issue Nature"
        End If
        mIssues.Add row
    Next row
    If mIssues.Count = 0 Then Fail "Add at least one OMS issue definition row."
    Set mAliases = ActiveRows("AliasesTable")
    For Each row In mAliases
        row("Report") = CanonReport(RequiredText(row, "Report"))
        RequiredText row, "Source Value"
        RequiredText row, "Source Context"
        RequiredText row, "Member ID"
        ' Aliases to inactive/out-of-period members remain visible as unmapped records.
    Next row
    Set mFileAssignments = ActiveRows("FileAssignmentsTable")
    For Each row In mFileAssignments
        row("Report") = CanonReport(RequiredText(row, "Report"))
        RequiredText row, "Filename Pattern"
        If Not mTeams.Exists(RequiredText(row, "Team ID")) Then Fail "File assignment points to a missing/inactive Team ID."
        If Len(TextOf(ValueOf(row, "Region"))) > 0 Then
            If CanonRegion(row("Region")) <> mTeams(row("Team ID"))("region") Then Fail "File-assignment Region and Team ID disagree."
        End If
    Next row
    Set mManagerAliases = ActiveRows("ManagerAliasesTable")
    For Each row In mManagerAliases
        row("Report") = CanonReport(RequiredText(row, "Report"))
        RequiredText row, "Source Manager"
        RequiredText row, "Source Context"
        If Not mTeams.Exists(RequiredText(row, "Team ID")) Then Fail "Manager alias points to a missing/inactive team."
    Next row
    Set rows = ActiveRows("FieldMappingsTable")
    For Each row In rows
        layoutID = RequiredText(row, "Layout ID")
        RequiredText row, "Field"
        If Not IsKnownField(row("Field")) Then Fail "Unknown canonical field: " & row("Field")
        If Not mFields.Exists(layoutID) Then
            Set fields = New Collection
            mFields.Add layoutID, fields
        Else
            Set fields = mFields(layoutID)
        End If
        For Each d In fields
            If d("Field") = row("Field") Then Fail "Duplicate field mapping: " & layoutID & " / " & row("Field")
        Next d
        fields.Add row
    Next row
    For Each row In mSources
        If Not mFields.Exists(row("Layout ID")) Then Fail "No field mappings for " & row("Layout ID")
    Next row
    BuildAmountMetrics
End Sub

Private Function WildMatch(ByVal value As String, ByVal pattern As String) As Boolean
    On Error GoTo BadPattern
    WildMatch = LCase$(value) Like LCase$(pattern)
    Exit Function
BadPattern:
    Fail "Invalid wildcard pattern: " & pattern
End Function

Private Function PatternSpecificity(ByVal pattern As String) As Long
    Dim i As Long, ch As String
    For i = 1 To Len(pattern)
        ch = Mid$(pattern, i, 1)
        If ch <> "*" And ch <> "?" And ch <> "#" And ch <> "[" And ch <> "]" Then PatternSpecificity = PatternSpecificity + 1
    Next i
End Function

Private Function AliasCandidates(ByVal rows As Collection, ByVal report As String, ByVal fileName As String, _
                                 ByVal nameField As String, ByVal value As String, ByVal idField As String) As Object
    Dim row As Object, result As Object, best As Long, score As Long
    Set result = NewMap(): best = -1
    For Each row In rows
        If row("Report") = report And CleanName(row(nameField)) = value Then
            If WildMatch(fileName, row("Source Context")) Then
                score = PatternSpecificity(row("Source Context"))
                If score > best Then Set result = NewMap(): best = score
                If score = best Then result(row(idField)) = True
            End If
        End If
    Next row
    Set AliasCandidates = result
End Function

Private Function UniqueKey(ByVal map As Object) As String
    Dim keys As Variant
    If map.Count = 1 Then keys = map.Keys: UniqueKey = CStr(keys(0))
End Function

Private Function Quoted(ByVal text As String) As String
    Quoted = """" & Replace(text, """", """""") & """"
End Function

Private Function Criterion(ByVal text As String) As String
    text = Replace(text, "~", "~~")
    text = Replace(text, "*", "~*")
    text = Replace(text, "?", "~?")
    Criterion = Quoted(text)
End Function

Private Function SheetRuleMatches(ByVal name As String, ByVal rule As String) As Boolean
    Dim pattern As Variant
    For Each pattern In Split(rule, "|")
        If WildMatch(name, Trim$(CStr(pattern))) Then SheetRuleMatches = True: Exit Function
    Next pattern
End Function

Private Function ExtractIssueCode(ByVal text As String) As Variant
    Dim code As String
    ExtractIssueCode = Empty
    text = TextOf(text)
    If Len(text) < 7 Then Exit Function
    code = Mid$(text, 6, 2)
    If DigitsOnly(code) Then ExtractIssueCode = CLng(code)
End Function

Private Sub AssertTest(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Fail "Self-test: " & message
End Sub

Private Function WeightedMean(ByVal sums As Variant, ByVal counts As Variant) As Variant
    Dim i As Long, total As Double, count As Long
    WeightedMean = Empty
    For i = LBound(sums) To UBound(sums)
        total = total + CDbl(sums(i)): count = count + CLng(counts(i))
    Next i
    If count > 0 Then WeightedMean = total / count
End Function

Private Sub TestBusinessRules()
    Dim map As Object, nested As Object, aliases As New Collection
    Set nested = NewMap("value", 42)
    Set map = NewMap("Nested", nested, "blank", Empty, "Member", "002")
    AssertTest map.Count = 3 And map("member") = "002", "Case-insensitive maps preserve member IDs."
    AssertTest IsEmpty(map("blank")) And map("nested")("value") = 42, "Maps preserve Empty and object values."
    Set map = NewMap()
    AssertTest map.Count = 0 And UniqueKey(map) = "", "Empty maps are supported."
    AssertTest UniqueKey(NewMap("002", True)) = "002", "Unique keys preserve leading zeroes."
    aliases.Add NewMap("Report", "eTR", "Source Value", "Kate", "Source Context", "*", "Member ID", "001")
    aliases.Add NewMap("Report", "eTR", "Source Value", "Kate", "Source Context", "*Toronto*", "Member ID", "002")
    Set map = AliasCandidates(aliases, "eTR", "eTR_Toronto.xlsx", "Source Value", "Kate", "Member ID")
    AssertTest UniqueKey(map) = "002", "Specific aliases override generic aliases."
    AssertTest Abs(CDbl(ParseAmount(999))) < 1000, "Exclude 999 credit."
    AssertTest Abs(CDbl(ParseAmount(-999))) < 1000, "Exclude 999 debit."
    AssertTest Abs(CDbl(ParseAmount(1000))) >= 1000, "Include 1,000 credit."
    AssertTest Abs(CDbl(ParseAmount(-1000))) >= 1000, "Include 1,000 debit."
    AssertTest Abs(CDbl(ParseAmount(0))) >= 0, "A zero threshold includes zero numeric balances."
    AssertTest ParseAmount("($1,250.50)") = -1250.5, "Parentheses debit amount."
    AssertTest ParseAmount("-1,250.50") = -1250.5, "Signed debit amount."
    AssertTest IsEmpty(ParseAmount("N/A")), "Reject non-numeric amounts."
    AssertTest IsEmpty(ParseAmount("1,5")), "Reject ambiguous decimal-comma text."
    AssertTest ExtractIssueCode("P1 l020 Loan - Past Due Loan Fee - Performing") = 20, "OMS positions 6-7."
    AssertTest ExtractIssueCode("PI 1012 Fac discrepancy") = 12, "OMS issue 12."
    AssertTest IsEmpty(ExtractIssueCode("Missing")), "Malformed OMS code."
    AssertTest CleanName("  Kate" & ChrW(160) & "  Smith  ") = "Kate Smith", "Clean employee spaces."
    AssertTest NormalizeHeader(" Age (days) ") = "agedays", "Normalize source headers."
    AssertTest WildMatch("eTR_NY_Team_Beta_2026.xlsx", "etr_ny_*.xlsx"), "Case-insensitive file patterns."
    AssertTest SheetRuleMatches("Current Day", "Sheet1|Current Day"), "Alternative worksheet names."
    AssertTest WeightedMean(Array(10, 270), Array(1, 9)) = 28, "Entry-weighted manager average, not average of analyst averages."
    AssertTest IsEmpty(WeightedMean(Array(0, 0), Array(0, 0))), "No-data average must be blank."
    AssertTest ParseDateValue("2026-10-02", False) = DateSerial(2026, 10, 2), "ISO dates."
    AssertTest ParseDateValue("10/02/2026", False) = DateSerial(2026, 10, 2), "US report dates."
    AssertTest IsEmpty(ParseDateValue("02/30/2026", False)), "Invalid calendar dates."
    Set map = NewMap("ItemDate", "2026-10-02", "Reference", "SRF1")
    AssertTest IsImportableSourceRow(map, "eTR", False), "eTR rows with valid Approval Dates are imported."
    Set map = NewMap("ItemDate", "Issue comment", "Reference", "Comment for prior eTR item")
    AssertTest Not IsImportableSourceRow(map, "eTR", False), "eTR comment rows without valid Approval Dates are ignored."
    AssertTest Criterion("A*?") = Quoted("A~*~?"), "Literal wildcard criteria."
    AssertTest LiteralValue("=1+1") = "'=1+1", "Treat imported text as text, not formulas."
End Sub

Public Sub RunOperationsSelfTests()
    On Error GoTo Failed
    TestBusinessRules
    MsgBox "Business-rule self-tests passed: amount boundaries, debit/credit amounts, OMS positions, " & _
           "eTR Approval Dates, number parsing, header normalization, wildcards, and weighted averages.", vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbCritical, "Self-test failed"
End Sub

' Optional temporary checkpoint for this stage.

Public Sub TypingCheckConfiguration()
    On Error GoTo Failed
    InitializeRun
    LoadConfiguration
    Debug.Print "Teams:", mTeams.Count
    Debug.Print "Members:", mMembers.Count
    Debug.Print "Source layouts:", mSources.Count
    Debug.Print "Account tabs:", mAccounts.Count
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Configuration checkpoint"
End Sub

' ========================================================================
' STAGE 5: FILES, REPORT IMPORT AND OWNERSHIP
' ========================================================================
' Run TypingCheckFiles to see matched filenames. This checkpoint enumerates files without opening 
' reports. Full import runs in Stage 8.
' Procedure typing order:
' 01. FileNameOf
' 02. IsExcelExtension
' 03. LogEntry
' 04. LogInfo
' 05. SelectSourceFiles
' 06. LogWarning
' 07. ResolveFieldColumn
' 08. ResolveMemberID
' 09. ResolveFileTeam
' 10. ResolveManagerTeam
' 11. ResolveTeamToken
' 12. EnsureUnassignedTeam
' 13. ResolveAssignment
' 14. BauFlag
' 15. IsAmountReport
' 16. AddNormalizedRecord
' 17. LastUsedColumn
' 18. KeyPart
' 19. RowContainsData
' 20. IsFooterRow
' 21. IsImportableSourceRow
' 22. EtrFingerprint
' 23. ReadSourceSheet
' 24. ImportFile
' 25. ImportAllReports
' 26. Participant
' 27. ParticipantKey
' 28. BuildParticipants

Private Function FileNameOf(ByVal path As String) As String
    FileNameOf = CreateObject("Scripting.FileSystemObject").GetFileName(path)
End Function

Private Function IsExcelExtension(ByVal name As String) As Boolean
    Dim extension As String
    extension = LCase$(CreateObject("Scripting.FileSystemObject").GetExtensionName(name))
    IsExcelExtension = extension = "xlsx" Or extension = "xlsm" Or extension = "xls" Or extension = "xlsb"
End Function

Private Sub LogEntry(ByVal level As String, ByVal action As String, ByVal file As String, _
                     ByVal sheet As String, ByVal row As Long, ByVal detail As String)
    If mLogs Is Nothing Then Exit Sub
    mLogs.Add Array(mRunID, Now, level, action, file, sheet, row, detail)
End Sub

Private Sub LogInfo(ByVal action As String, ByVal file As String, ByVal sheet As String, ByVal row As Long, ByVal detail As String)
    LogEntry "INFO", action, file, sheet, row, detail
End Sub

Private Function SelectSourceFiles(ByVal source As Object) As Collection
    Dim fso As Object, folder As Object, file As Object, names As Object
    Dim keys As Variant, i As Long, j As Long, tmp As Variant
    Dim result As New Collection, newest As String, newestTime As Date
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(mFolder)
    Set names = NewMap()
    For Each file In folder.Files
        If Left$(file.Name, 2) <> "~$" Then
            If LCase$(file.Path) <> LCase$(ThisWorkbook.FullName) Then
                If IsExcelExtension(file.Name) And WildMatch(file.Name, source("Filename Pattern")) Then
                    names.Add file.Path, file.DateLastModified
                    If Len(newest) = 0 Then
                        newest = file.Path: newestTime = file.DateLastModified
                    ElseIf file.DateLastModified > newestTime Then
                        newest = file.Path: newestTime = file.DateLastModified
                    ElseIf file.DateLastModified = newestTime Then
                        If StrComp(file.Path, newest, vbTextCompare) > 0 Then newest = file.Path
                    End If
                End If
            End If
        End If
    Next file
    If names.Count = 0 Then Set SelectSourceFiles = result: Exit Function
    If Not YesFlag(source("Multiple Files")) Then
        result.Add newest
        If names.Count > 1 Then LogInfo "Latest file selected", FileNameOf(newest), "", 0, _
            CStr(names.Count) & " matches; newest modified file used. Archive old reports outside the input folder."
    Else
        keys = names.Keys
        For i = LBound(keys) To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If StrComp(CStr(keys(i)), CStr(keys(j)), vbTextCompare) > 0 Then
                    tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
                End If
            Next j
        Next i
        For i = LBound(keys) To UBound(keys): result.Add CStr(keys(i)): Next i
    End If
    Set SelectSourceFiles = result
End Function

Private Sub LogWarning(ByVal action As String, ByVal file As String, ByVal sheet As String, ByVal row As Long, ByVal detail As String)
    mWarnings = mWarnings + 1
    LogEntry "REVIEW", action, file, sheet, row, detail
End Sub

Private Function ResolveFieldColumn(ByVal field As Object, ByVal headers As Object, ByVal lastCol As Long, _
                                    ByVal path As String, ByVal sheetName As String) As Long
    Dim alias As Variant, key As String, indexes As Collection, fallback As Variant
    For Each alias In Split(TextOf(ValueOf(field, "Header Aliases")), "|")
        key = NormalizeHeader(CStr(alias))
        If Len(key) > 0 Then
            If headers.Exists(key) Then
                Set indexes = headers(key)
                If indexes.Count > 1 Then Fail "Ambiguous duplicate header '" & alias & "' in " & FileNameOf(path) & " / " & sheetName
                ResolveFieldColumn = CLng(indexes(1))
                Exit Function
            End If
        End If
    Next alias
    fallback = ValueOf(field, "Column Fallback")
    If Len(TextOf(fallback)) > 0 Then
        If Not IsNumeric(fallback) Then Fail "Column Fallback must be numeric or 0."
        If CDbl(fallback) < 0 Or CDbl(fallback) <> Fix(CDbl(fallback)) Then Fail "Invalid Column Fallback."
        If CDbl(fallback) > 0 Then
            If CDbl(fallback) > lastCol Then Fail "Fallback column is outside the source data."
            ResolveFieldColumn = CLng(fallback)
            LogWarning "Fallback column used", FileNameOf(path), sheetName, 0, _
                field("Field") & " used column " & fallback & ". Confirm the layout."
            Exit Function
        End If
    End If
    If YesFlag(ValueOf(field, "Required")) Then Fail "Missing required header for " & field("Field") & " in " & FileNameOf(path) & " / " & sheetName
End Function

Private Function ResolveMemberID(ByVal report As String, ByVal fileName As String, ByVal employee As String, _
                                 ByVal observedTeam As String, ByVal regionRule As String) As String
    Dim candidates As Object, id As Variant, member As Object
    Set candidates = AliasCandidates(mAliases, report, fileName, "Source Value", employee, "Member ID")
    If candidates.Count > 1 Then Fail "Ambiguous employee aliases for '" & employee & "' in " & fileName
    ResolveMemberID = UniqueKey(candidates)
    If Len(ResolveMemberID) > 0 Then Exit Function
    ' Prefer a literal Member ID, then an unambiguous name in source context.
    If mMembers.Exists(employee) Then ResolveMemberID = employee: Exit Function
    For Each id In mMembers.Keys
        Set member = mMembers(id)
        If CleanName(member("name")) = employee Then
            If Len(observedTeam) > 0 Then
                If member("team") = observedTeam Then candidates(id) = True
            ElseIf regionRule = "NY" Or regionRule = "Toronto" Then
                If mTeams(member("team"))("region") = regionRule Then candidates(id) = True
            Else
                candidates(id) = True
            End If
        End If
    Next id
    ResolveMemberID = UniqueKey(candidates)
    If Len(ResolveMemberID) > 0 Then Exit Function
    ' An employee under another manager may still have one unique global name.
    Set candidates = NewMap()
    For Each id In mMembers.Keys
        If CleanName(mMembers(id)("name")) = employee Then candidates(id) = True
    Next id
    ResolveMemberID = UniqueKey(candidates)
End Function

Private Function ResolveFileTeam(ByVal report As String, ByVal fileName As String) As String
    Dim row As Object, best As Long, score As Long, team As String
    best = -1
    For Each row In mFileAssignments
        If row("Report") = report And WildMatch(fileName, row("Filename Pattern")) Then
            score = PatternSpecificity(row("Filename Pattern"))
            If score > best Then
                best = score: team = row("Team ID")
            ElseIf score = best Then
                If team <> row("Team ID") Then Fail "Ambiguous file-team assignment for " & fileName
            End If
        End If
    Next row
    ResolveFileTeam = team
End Function

Private Function ResolveManagerTeam(ByVal report As String, ByVal fileName As String, ByVal managerName As String) As String
    Dim id As Variant, candidates As Object
    If Len(managerName) = 0 Then Exit Function
    Set candidates = AliasCandidates(mManagerAliases, report, fileName, "Source Manager", CleanName(managerName), "Team ID")
    If candidates.Count = 0 Then
        For Each id In mTeamOrder
            If CleanName(mTeams(id)("manager")) = CleanName(managerName) Then candidates(id) = True
        Next id
    End If
    If candidates.Count <> 1 Then Fail "Source manager '" & managerName & "' does not identify one team in " & fileName & ". Configure ManagerAliasesTable."
    ResolveManagerTeam = UniqueKey(candidates)
End Function

Private Function ResolveTeamToken(ByVal token As String) As String
    Dim id As Variant, result As String
    If Len(token) = 0 Then Exit Function
    If mTeams.Exists(token) Then ResolveTeamToken = token: Exit Function
    For Each id In mTeamOrder
        If CleanName(mTeams(id)("name")) = CleanName(token) Then
            If Len(result) > 0 Then Fail "Ambiguous source team name: " & token
            result = CStr(id)
        End If
    Next id
    If Len(result) = 0 Then Fail "Unknown source Team value: " & token
    ResolveTeamToken = result
End Function

Private Sub EnsureUnassignedTeam()
    Dim d As Object, flags As Object, report As Variant
    If mTeams.Exists(UNASSIGNED_TEAM) Then Exit Sub
    Set flags = NewMap()
    For Each report In ReportNames(): flags.Add CStr(report), False: Next report
    Set d = NewMap("id", UNASSIGNED_TEAM, "name", "Unassigned - check mappings", _
                   "region", "Unassigned", "manager", "Unassigned", "flags", flags)
    mTeams.Add UNASSIGNED_TEAM, d: mTeamOrder.Add UNASSIGNED_TEAM
End Sub

Private Function ResolveAssignment(ByVal report As String, ByVal fileName As String, ByVal employee As String, _
                                   ByVal row As Object, ByVal regionRule As String) As Object
    Dim d As Object, member As Object, observedTeam As String, rowTeam As String
    Dim managerTeam As String, fileTeam As String, memberID As String, configuredTeam As String
    Dim rowRegion As String
    rowTeam = ResolveTeamToken(TextOf(ValueOf(row, "Team")))
    managerTeam = ResolveManagerTeam(report, fileName, TextOf(ValueOf(row, "Manager")))
    fileTeam = ResolveFileTeam(report, fileName)
    If Len(rowTeam) > 0 Then
        observedTeam = rowTeam
    ElseIf Len(managerTeam) > 0 Then
        observedTeam = managerTeam
    Else
        observedTeam = fileTeam
    End If
    If Len(rowTeam) > 0 And Len(managerTeam) > 0 Then
        If rowTeam <> managerTeam Then Fail "Source Team and Manager resolve to different teams in " & fileName & ". Correct the mapping."
    End If
    memberID = ResolveMemberID(report, fileName, employee, observedTeam, regionRule)
    If Len(memberID) > 0 Then
        If mMembers.Exists(memberID) Then
            Set member = mMembers(memberID)
            configuredTeam = member("team")
        Else
            memberID = ""
        End If
    End If
    If Len(observedTeam) = 0 Then observedTeam = configuredTeam
    If Len(observedTeam) = 0 Then
        EnsureUnassignedTeam
        observedTeam = UNASSIGNED_TEAM
    End If
    Set d = NewMap("team", observedTeam, "member", memberID, "configuredTeam", configuredTeam, _
                   "configuredTeamName", "", "employee", employee, "status", "OK", "note", "")
    If Len(memberID) > 0 Then
        d("employee") = member("name")
        d("configuredTeamName") = mTeams(configuredTeam)("name")
        If observedTeam <> configuredTeam Then
            d("status") = "Team membership mismatch"
            d("note") = memberID & " (" & member("name") & ") is configured under " & mTeams(configuredTeam)("name") & _
                ", not " & mTeams(observedTeam)("name") & ". Please check Team Members setting."
        End If
    Else
        d("status") = "Unmapped employee"
        d("note") = "Source employee '" & employee & "' has no unique effective member mapping. Please check Employee aliases and Team Members setting."
    End If
    If observedTeam = UNASSIGNED_TEAM Then
        d("status") = "Unassigned team"
        d("note") = d("note") & " No team could be resolved. Configure an authoritative Team/Manager field, a file assignment, or a member alias."
    End If
    rowRegion = TextOf(ValueOf(row, "Region"))
    If Len(rowRegion) > 0 Then
        If observedTeam <> UNASSIGNED_TEAM Then
            If CanonRegion(rowRegion) <> mTeams(observedTeam)("region") Then Fail "Source Region and resolved Team disagree in " & fileName
        End If
    End If
    Set ResolveAssignment = d
End Function

Private Function BauFlag(ByVal teamID As String, ByVal report As String) As Boolean
    If mTeams.Exists(teamID) Then BauFlag = CBool(mTeams(teamID)("flags")(report))
End Function

Private Function IsAmountReport(ByVal report As String) As Boolean
    IsAmountReport = report = "Chase" Or report = "Chase CAD" Or report = "Internal" Or report = "Internal HSBC"
End Function

Private Sub AddNormalizedRecord(ByVal row As Object, ByVal source As Object, ByVal fileName As String, _
                                ByVal sheetName As String, ByVal sourceRow As Long, ByVal accountNumber As String, _
                                ByVal accountCurrency As String, ByVal date1904 As Boolean)
    Dim rec As Object, assignment As Object, amount As Variant, age As Variant, itemDate As Variant
    Dim issue As Variant, reportName As String, employeeName As String, recCurrency As String, stat As Object
    reportName = source("Report")
    employeeName = CleanName(TextOf(ValueOf(row, "Employee")))
    If Len(employeeName) = 0 Then employeeName = "[Unassigned employee]"
    amount = ParseAmount(ValueOf(row, "Amount"))
    If IsAmountReport(reportName) And IsEmpty(amount) Then Fail "Missing/invalid Amount: " & fileName & " / " & sheetName & " / row " & sourceRow
    itemDate = ParseDateValue(ValueOf(row, "ItemDate"), date1904)
    If Len(TextOf(ValueOf(row, "ItemDate"))) > 0 And IsEmpty(itemDate) Then _
        Fail "Invalid ItemDate: " & fileName & " / " & sheetName & " / row " & sourceRow
    age = ParseAmount(ValueOf(row, "AgeDays"))
    If Len(TextOf(ValueOf(row, "AgeDays"))) > 0 And IsEmpty(age) Then _
        Fail "Invalid AgeDays: " & fileName & " / " & sheetName & " / row " & sourceRow
    If IsEmpty(age) Then
        If Not IsEmpty(itemDate) Then
            age = DateDiff("d", CDate(itemDate), mReportDate)
        ElseIf IsAmountReport(reportName) Then
            Fail "An amount report needs AgeDays or ItemDate: " & fileName & " / row " & sourceRow
        End If
    End If
    If Not IsEmpty(age) Then
        If CDbl(age) < 0 Then Fail "Negative age: " & fileName & " / " & sheetName & " / row " & sourceRow
    End If
    Set assignment = ResolveAssignment(reportName, fileName, employeeName, row, source("Region Rule"))
    recCurrency = TextOf(ValueOf(row, "Currency"))
    If Len(recCurrency) = 0 Then recCurrency = accountCurrency
    issue = Empty
    If reportName = "OMS" Then
        issue = ExtractIssueCode(TextOf(ValueOf(row, "Issue")))
        If IsEmpty(issue) Then LogWarning "Unrecognized OMS issue code", fileName, sheetName, sourceRow, _
            "Positions 6-7 are not two digits. Row retained in all-OMS totals but not any tracked-issue column."
    End If
    Set rec = NewMap("id", "R" & Format$(mRecords.Count + 1, "000000"), "report", reportName, _
        "team", assignment("team"), "region", mTeams(assignment("team"))("region"), _
        "teamName", mTeams(assignment("team"))("name"), "manager", mTeams(assignment("team"))("manager"), _
        "member", assignment("member"), "employee", assignment("employee"), "sourceEmployee", employeeName, _
        "date", itemDate, "age", age, "status", ValueOf(row, "Status"), "amount", amount, "currency", recCurrency, _
        "client", ValueOf(row, "Client"), "reference", ValueOf(row, "Reference"), _
        "comment", ValueOf(row, IIf(reportName = "OMS", "Issue", "Comment")), _
        "file", fileName, "sheet", sheetName, "sourceRow", sourceRow, "importStatus", assignment("status"), _
        "issue", issue, "configuredTeam", assignment("configuredTeam"), "configuredTeamName", assignment("configuredTeamName"), _
        "note", assignment("note"), "account", accountNumber, "bau", IIf(BauFlag(assignment("team"), reportName), "BAU", "Non-BAU"))
    mRecords.Add rec
    Set stat = mStats(reportName)
    stat("normalized") = CLng(stat("normalized")) + 1
    If rec("importStatus") <> "OK" Then LogWarning rec("importStatus"), fileName, sheetName, sourceRow, rec("note")
End Sub

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If Not found Is Nothing Then LastUsedColumn = found.Column
End Function

Private Function KeyPart(ByVal value As String) As String
    KeyPart = CStr(Len(value)) & ":" & value
End Function

Private Function RowContainsData(ByVal row As Object) As Boolean
    Dim field As Variant
    For Each field In Array("Reference", "Employee", "Client", "Amount", "ItemDate", "Issue")
        If Len(TextOf(ValueOf(row, CStr(field)))) > 0 Then RowContainsData = True: Exit Function
    Next field
End Function

Private Function IsFooterRow(ByVal row As Object) As Boolean
    Dim field As Variant, value As String
    For Each field In Array("Reference", "Employee", "Client")
        value = NormalizeHeader(TextOf(ValueOf(row, CStr(field))))
        If value = "total" Or value = "grandtotal" Or value = "subtotal" Then IsFooterRow = True: Exit Function
    Next field
End Function

Private Function IsImportableSourceRow(ByVal row As Object, ByVal report As String, _
                                       ByVal date1904 As Boolean) As Boolean
    If Not RowContainsData(row) Or IsFooterRow(row) Then Exit Function
    If report = "eTR" Then
        If IsEmpty(ParseDateValue(ValueOf(row, "ItemDate"), date1904)) Then Exit Function
    End If
    IsImportableSourceRow = True
End Function

Private Function EtrFingerprint(ByVal row As Object) As String
    Dim field As Variant, result As String
    For Each field In Array("Reference", "Employee", "ItemDate", "Client", "Status", "Team", "Manager", "Region")
        result = result & KeyPart(CleanName(TextOf(ValueOf(row, CStr(field)))))
    Next field
    EtrFingerprint = result
End Function

Private Sub ReadSourceSheet(ByVal ws As Worksheet, ByVal source As Object, ByVal path As String, _
                            ByVal accountNumber As String, ByVal accountCurrency As String)
    Dim headerRow As Long, lastRow As Long, lastCol As Long, r As Long
    Dim data As Variant, fieldCols As Object, row As Object, seenKey As String
    Dim headers As Object, indexes As Collection, field As Object, col As Long, v As Variant
    Dim fields As Collection, stat As Object, fingerprint As String
    headerRow = CLng(source("Header Row"))
    lastRow = LastUsedRow(ws): lastCol = LastUsedColumn(ws)
    If lastRow < headerRow Or lastCol = 0 Then Fail "Header row missing: " & FileNameOf(path) & " / " & ws.Name
    data = ws.Range(ws.Cells(headerRow, 1), ws.Cells(Application.Max(lastRow, headerRow + 1), lastCol)).Value2
    Set headers = NewMap()
    For col = 1 To lastCol
        seenKey = NormalizeHeader(TextOf(data(1, col)))
        If Len(seenKey) > 0 Then
            If Not headers.Exists(seenKey) Then
                Set indexes = New Collection: headers.Add seenKey, indexes
            Else
                Set indexes = headers(seenKey)
            End If
            indexes.Add col
        End If
    Next col
    Set fieldCols = NewMap()
    Set fields = mFields(source("Layout ID"))
    For Each field In fields
        col = ResolveFieldColumn(field, headers, lastCol, path, ws.Name)
        fieldCols.Add field("Field"), col
    Next field
    If Not fieldCols.Exists("Employee") Then Fail "Employee field must be mapped for " & source("Layout ID")
    If fieldCols("Employee") = 0 Then Fail "Employee column not found in " & FileNameOf(path) & " / " & ws.Name
    If IsAmountReport(source("Report")) Then
        If Not fieldCols.Exists("Amount") Then Fail "Amount field must be mapped for " & source("Layout ID")
        If fieldCols("Amount") = 0 Then Fail "Amount column not found in " & FileNameOf(path) & " / " & ws.Name
    End If
    If source("Report") = "OMS" Then
        If Not fieldCols.Exists("Issue") Then Fail "Issue field must be mapped for OMS."
        If fieldCols("Issue") = 0 Then Fail "Issue column not found in " & FileNameOf(path)
    End If
    If source("Report") = "eTR" Then
        If Not fieldCols.Exists("ItemDate") Then Fail "eTR ItemDate must be mapped to Approval Date."
        If fieldCols("ItemDate") <> 1 Then _
            Fail "eTR ItemDate must map to Approval Date in column A: " & FileNameOf(path) & " / " & ws.Name
    End If
    Set stat = mStats(source("Report"))
    For r = 2 To UBound(data, 1)
        Set row = NewMap()
        For Each field In fields
            col = fieldCols(field("Field"))
            If col > 0 Then
                v = data(r, col)
                If IsError(v) Then Fail "Source-cell error: " & FileNameOf(path) & " / " & ws.Name & " / row " & (r + headerRow - 1)
            Else
                v = ValueOf(field, "Default Value")
            End If
            row.Add field("Field"), v
        Next field
        If IsImportableSourceRow(row, source("Report"), mSourceBook.Date1904) Then
            seenKey = KeyPart(path) & KeyPart(ws.Name) & CStr(r + headerRow - 1)
            If mReadKeys.Exists(seenKey) Then Fail "Overlapping source definitions import the same row twice: " & FileNameOf(path) & " / " & ws.Name
            mReadKeys.Add seenKey, True
            stat("raw") = CLng(stat("raw")) + 1
            If source("Report") = "eTR" Then
                fingerprint = EtrFingerprint(row)
                If mFingerprints.Exists(fingerprint) Then
                    If mFingerprints(fingerprint) <> FileNameOf(path) Then
                        stat("duplicates") = CLng(stat("duplicates")) + 1
                        LogWarning "Possible eTR duplicate", FileNameOf(path), ws.Name, r + headerRow - 1, _
                            "Identical mapped fields also found in " & mFingerprints(fingerprint) & ". Both rows retained; verify exports do not overlap."
                    End If
                Else
                    mFingerprints.Add fingerprint, FileNameOf(path)
                End If
            End If
            AddNormalizedRecord row, source, FileNameOf(path), ws.Name, r + headerRow - 1, _
                                accountNumber, accountCurrency, mSourceBook.Date1904
        End If
    Next r
    LogInfo "Sheet imported", FileNameOf(path), ws.Name, 0, source("Layout ID") & " / cached cell values read; links were not updated."
End Sub

Private Sub ImportFile(ByVal path As String, ByVal source As Object)
    Dim openBook As Workbook, ws As Worksheet, account As Object, matched As Long
    Dim stat As Object, fileID As String, hadError As Long, errorText As String
    On Error GoTo Failed
    For Each openBook In Application.Workbooks
        If StrComp(openBook.FullName, path, vbTextCompare) = 0 Then _
            Fail "Close the source workbook before refresh: " & FileNameOf(path)
    Next openBook
    Application.AutomationSecurity = SECURITY_FORCE_DISABLE
    Set mSourceBook = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, _
        Password:="", WriteResPassword:="", IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
    Set stat = mStats(source("Report"))
    fileID = LCase$(path)
    If Not stat("files").Exists(fileID) Then stat("files").Add fileID, True
    If TextOf(source("Worksheet Rule")) = "Configured account tabs" Then
        For Each account In mAccounts
            If account("Report") = source("Report") Then
                Set ws = SheetOrNothing(mSourceBook, account("Account Tab"))
                If ws Is Nothing Then Fail "Missing account tab " & account("Account Tab") & " in " & FileNameOf(path)
                ReadSourceSheet ws, source, path, account("Account Tab"), TextOf(ValueOf(account, "Currency"))
                matched = matched + 1
            End If
        Next account
    Else
        For Each ws In mSourceBook.Worksheets
            If SheetRuleMatches(ws.Name, source("Worksheet Rule")) Then
                ReadSourceSheet ws, source, path, "", ""
                matched = matched + 1
            End If
        Next ws
    End If
    If matched = 0 Then Fail "No configured data sheet found in " & FileNameOf(path) & ". Rule: " & source("Worksheet Rule")
    mSourceBook.Close SaveChanges:=False
    Set mSourceBook = Nothing
    Exit Sub
Failed:
    hadError = Err.Number: errorText = Err.Description
    On Error Resume Next
    If Not mSourceBook Is Nothing Then mSourceBook.Close SaveChanges:=False
    Set mSourceBook = Nothing
    On Error GoTo 0
    Err.Raise hadError, "ImportFile", errorText
End Sub

Private Sub ImportAllReports()
    Dim source As Object, files As Collection, path As Variant, fileCount As Long
    For Each source In mSources
        Set files = SelectSourceFiles(source)
        If files.Count = 0 Then Fail "No source file matched layout " & source("Layout ID") & ": " & source("Filename Pattern")
        For Each path In files
            Application.StatusBar = "Operations dashboard: " & FileNameOf(CStr(path))
            ImportFile CStr(path), source
            fileCount = fileCount + 1
        Next path
    Next source
    LogInfo "Files imported", "", "", 0, CStr(fileCount) & " configured source-file selections."
End Sub

Private Function Participant(ByVal memberID As String, ByVal employee As String, ByVal sourceEmployee As String, _
                             ByVal note As String, ByVal review As Boolean) As Object
    Set Participant = NewMap("member", memberID, "employee", employee, "sourceEmployee", sourceEmployee, "note", note, "review", review)
End Function

Private Function ParticipantKey(ByVal p As Object) As String
    If Len(p("member")) > 0 Then
        ParticipantKey = "M:" & p("member")
    Else
        ParticipantKey = "U:" & CleanName(p("sourceEmployee"))
    End If
End Function

Private Sub BuildParticipants()
    Dim id As Variant, memberID As Variant, list As Collection, p As Object, member As Object
    Dim rec As Object, seen As Object, key As String
    For Each id In mTeamOrder
        Set list = New Collection: Set seen = NewMap()
        For Each memberID In mMembers.Keys
            Set member = mMembers(memberID)
            If member("team") = CStr(id) Then
                Set p = Participant(memberID, member("name"), "", "", False)
                list.Add p: seen.Add ParticipantKey(p), p
            End If
        Next memberID
        For Each rec In mRecords
            If rec("team") = CStr(id) Then
                Set p = Participant(rec("member"), rec("employee"), rec("sourceEmployee"), rec("note"), rec("importStatus") <> "OK")
                key = ParticipantKey(p)
                If Not seen.Exists(key) Then
                    list.Add p: seen.Add key, p
                ElseIf rec("importStatus") <> "OK" Then
                    Set p = seen(key): p("review") = True: p("note") = rec("note")
                End If
            End If
        Next rec
        mParticipants.Add CStr(id), list
    Next id
End Sub

' Optional temporary checkpoint for this stage.

Public Sub TypingCheckFiles()
    Dim layout As Object, selected As Collection, item As Variant
    On Error GoTo Failed
    InitializeRun
    LoadConfiguration
    For Each layout In mSources
        Set selected = SelectSourceFiles(layout)
        Debug.Print layout("Report"), selected.Count
        For Each item In selected
            Debug.Print "  ", item
        Next item
    Next layout
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Filename checkpoint"
End Sub

' ========================================================================
' STAGE 6: NORMALIZED DATA AND FORMULA CONSTRUCTION
' ========================================================================
' Run TypingCheckFormula. It prints a bounded COUNTIFS expression for eTR; it does not write 
' spreadsheet data.
' Procedure typing order:
' 01. ResetOutput
' 02. FreezeAt
' 03. AddContainsFormat
' 04. WriteNormalizedData
' 05. DataRange
' 06. PairList
' 07. ParticipantCriteria
' 08. CountExpr
' 09. MetricCriteria
' 10. MetricCount
' 11. MetricAgeAverage
' 12. OmsIssueCount
' 13. TrackedOmsFormula
' 14. IssueHeaderFormula
' 15. TeamReportCount
' 16. ParticipantAgedAmountFormula
' 17. EscapedSelector
' 18. SelectedFormula
' 19. SelectedCount
' 20. ReportCriteria
' 21. DashboardReportCriteria
' 22. SelectedDashboardCount
' 23. SelectedReportCount
' 24. SelectedAllReports
' 25. ReportRegionCount

Private Sub ResetOutput(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.ListObjects.Count To 1 Step -1: ws.ListObjects(i).Unlist: Next i
    ws.UsedRange.UnMerge
    ws.UsedRange.Clear
    ws.UsedRange.Rows.RowHeight = 18
End Sub

Private Sub FreezeAt(ByVal ws As Worksheet, ByVal rows As Long, ByVal columns As Long)
    ThisWorkbook.Activate: ws.Activate
    If Application.ActiveWindow Is Nothing Then Exit Sub
    With Application.ActiveWindow
        .FreezePanes = False: .SplitColumn = columns: .SplitRow = rows: .FreezePanes = True
        .DisplayGridlines = False
    End With
End Sub

Private Sub AddContainsFormat(ByVal target As Range, ByVal text As String, ByVal fill As Long, ByVal font As Long)
    Dim condition As FormatCondition
    Set condition = target.FormatConditions.Add(Type:=xlTextString, String:=text, TextOperator:=xlContains)
    condition.Interior.Color = fill: condition.Font.Color = font
End Sub

Private Sub WriteNormalizedData()
    Dim ws As Worksheet, values() As Variant, rec As Object, r As Long, fields As Variant
    Dim c As Long, lo As ListObject, body As Range
    Set ws = RequiredSheet("Normalized Data")
    ResetOutput ws
    fields = Array("id", "report", "region", "team", "teamName", "manager", "member", "employee", _
        "sourceEmployee", "date", "age", "status", "amount", "currency", "client", "reference", _
        "comment", "file", "sheet", "sourceRow", "importStatus", "", "issue", "configuredTeam", _
        "configuredTeamName", "note", "account", "bau", "")
    PutRow ws, 1, 1, Array("Record ID", "Report", "Region", "Team ID", "Team Name", "Manager", _
        "Member ID", "Employee", "Source Employee", "Item Date", "Age Days", "Status / Category", _
        "Amount", "Currency", "Client / Borrower", "Reference", "Comment / Issue", "Source File", _
        "Source Sheet", "Source Row", "Import Status", "Meets Amount Threshold", "OMS Issue Code", _
        "Configured Team ID", "Configured Team Name", "Membership Note", "Account Number", "BAU Status", "OMS Tracked")
    mDataLastRow = Application.Max(2, mRecords.Count + 1)
    If mRecords.Count > 0 Then
        ReDim values(1 To mRecords.Count, 1 To DATA_COLUMNS)
        r = 0
        For Each rec In mRecords
            r = r + 1
            For c = 1 To DATA_COLUMNS
                If Len(fields(c - 1)) > 0 Then values(r, c) = LiteralValue(rec(fields(c - 1)))
            Next c
        Next rec
        ws.Range("A2").Resize(mRecords.Count, DATA_COLUMNS).Value2 = values
    End If
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:AC" & mDataLastRow), , xlYes)
    lo.Name = "NormalizedDataTable": lo.TableStyle = "TableStyleMedium2"
    Set body = ws.Range("V2:V" & mDataLastRow)
    body.FormulaR1C1 = "=IF(RC1="""","""",IF(OR(RC2=""Chase"",RC2=""Chase CAD"",RC2=""Internal"",RC2=""Internal HSBC"")," & _
        "IF(ISNUMBER(RC13),IF(ABS(RC13)>='Configuration'!" & SettingCell("Amount threshold").Address(True, True, xlR1C1) & _
        ",""Yes"",""No""),""No""),""""))"
    ' One reusable flag keeps tracked-OMS totals compact even with many codes.
    ws.Range("AC2:AC" & mDataLastRow).FormulaR1C1 = _
        "=IF(RC1="""","""",IF(RC2<>""OMS"","""",IF(ISNUMBER(RC23)," & _
        "IF(COUNTIFS(OmsIssueDefinitionsTable[Active],""Yes"",OmsIssueDefinitionsTable[Issue Code],RC23)>0,""Yes"",""No""),""No"")))"
    With ws.Range("A1:AC" & mDataLastRow)
        .Font.Name = "Arial": .Font.Size = 10: .VerticalAlignment = xlCenter
    End With
    HeaderStyle ws.Range("A1:AC1")
    ws.Rows(1).RowHeight = 32
    ws.Range("J2:J" & mDataLastRow).NumberFormat = "mm/dd/yyyy"
    ws.Range("K2:K" & mDataLastRow).NumberFormat = "0.0"
    ws.Range("M2:M" & mDataLastRow).NumberFormat = "#,##0.00;[Red]-#,##0.00"
    ws.Range("W2:W" & mDataLastRow).NumberFormat = "00"
    ColumnWidths ws, "G:G", 12, "H:I", 18, "A:F", 17, "J:K", 13, "L:L", 24, _
                 "M:N", 16, "O:Q", 28, "R:R", 40, "S:AC", 21, "Z:Z", 65
    AddContainsFormat ws.Range("U2:U" & mDataLastRow), "mismatch", COLOR_AMBER, RGB(156, 87, 0)
    AddContainsFormat ws.Range("U2:U" & mDataLastRow), "Unmapped", RGB(244, 204, 204), RGB(156, 0, 6)
    AddContainsFormat ws.Range("U2:U" & mDataLastRow), "Unassigned", RGB(244, 204, 204), RGB(156, 0, 6)
    AddContainsFormat ws.Range("AB2:AB" & mDataLastRow), "Non-BAU", COLOR_GREY, RGB(102, 102, 102)
    FreezeAt ws, 1, 2
End Sub

Private Function DataRange(ByVal column As String) As String
    DataRange = "'Normalized Data'!$" & column & "$2:$" & column & "$" & mDataLastRow
End Function

Private Function PairList(ParamArray pairs() As Variant) As String
    Dim i As Long, result As String
    For i = LBound(pairs) To UBound(pairs) Step 2
        If Len(result) > 0 Then result = result & ","
        result = result & DataRange(CStr(pairs(i))) & "," & Criterion(CStr(pairs(i + 1)))
    Next i
    PairList = result
End Function

Private Function ParticipantCriteria(ByVal p As Object) As String
    If p Is Nothing Then Exit Function
    If Len(p("member")) > 0 Then
        ParticipantCriteria = "," & PairList("G", p("member"))
    Else
        ParticipantCriteria = "," & PairList("I", p("sourceEmployee"))
    End If
End Function

Private Function CountExpr(ByVal pairs As String) As String
    CountExpr = "COUNTIFS(" & pairs & ")"
End Function

Private Function MetricCriteria(ByVal teamID As String, ByVal p As Object, ByVal metric As Object) As String
    Dim pairs As String
    pairs = PairList("B", metric("report"), "D", teamID, "V", "Yes") & ParticipantCriteria(p)
    If Len(metric("account")) > 0 Then pairs = pairs & "," & PairList("AA", metric("account"))
    MetricCriteria = pairs
End Function

Private Function MetricCount(ByVal teamID As String, ByVal p As Object, ByVal metric As Object) As String
    MetricCount = CountExpr(MetricCriteria(teamID, p, metric))
End Function

Private Function MetricAgeAverage(ByVal teamID As String, ByVal p As Object, ByVal metric As Object) As String
    Dim pairs As String, ageCount As String, ageSum As String
    pairs = MetricCriteria(teamID, p, metric)
    ageCount = CountExpr(pairs & "," & DataRange("K") & "," & Quoted(">=0") & "," & DataRange("K") & "," & Quoted("<>"))
    ageSum = "SUMIFS(" & DataRange("K") & "," & pairs & ")"
    MetricAgeAverage = "=IF(" & ageCount & "=0,""""," & ageSum & "/" & ageCount & ")"
End Function

Private Function OmsIssueCount(ByVal teamID As String, ByVal p As Object, ByVal issue As Object) As String
    Dim mappingRow As Long, pairs As String
    mappingRow = CLng(issue("@row"))
    pairs = PairList("B", "OMS", "D", teamID) & ParticipantCriteria(p) & _
        "," & DataRange("W") & ",'Report Mapping'!$B$" & mappingRow
    OmsIssueCount = "IF('Report Mapping'!$A$" & mappingRow & "<>""Yes"",0," & CountExpr(pairs) & ")"
End Function

Private Function TrackedOmsFormula(ByVal teamID As String, ByVal p As Object) As String
    TrackedOmsFormula = "=" & CountExpr(PairList("B", "OMS", "D", teamID, "AC", "Yes") & ParticipantCriteria(p))
End Function

Private Function IssueHeaderFormula(ByVal issue As Object) As String
    Dim row As Long: row = CLng(issue("@row"))
    IssueHeaderFormula = "=IF('Report Mapping'!$A$" & row & "<>""Yes"",""Add OMS issue"",""OMS ""&TEXT('Report Mapping'!$B$" & _
        row & ",""00"")&"" - ""&'Report Mapping'!$C$" & row & ")"
End Function

Private Function TeamReportCount(ByVal teamID As String, ByVal report As String, ByVal p As Object, Optional ByVal aged As Boolean = False) As String
    Dim pairs As String
    pairs = PairList("B", report, "D", teamID) & ParticipantCriteria(p)
    If IsAmountReport(report) Then pairs = pairs & "," & PairList("V", "Yes")
    If aged Then pairs = pairs & "," & DataRange("K") & ","">=30"""
    TeamReportCount = CountExpr(pairs)
End Function

Private Function ParticipantAgedAmountFormula(ByVal teamID As String, ByVal p As Object) As String
    Dim report As Variant, result As String
    For Each report In Array("Chase", "Chase CAD", "Internal", "Internal HSBC")
        If Len(result) > 0 Then result = result & "+"
        result = result & TeamReportCount(teamID, CStr(report), p, True)
    Next report
    ParticipantAgedAmountFormula = "=" & result
End Function

Private Function EscapedSelector(ByVal address As String) As String
    EscapedSelector = "SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" & address & ",""~"",""~~""),""*"",""~*""),""?"",""~?"")"
End Function

Private Function SelectedFormula(ByVal pairs As String) As String
    Dim teamPairs As String, regionPairs As String
    teamPairs = "," & DataRange("E") & "," & EscapedSelector("$E$5")
    regionPairs = "," & DataRange("C") & ",$B$5"
    SelectedFormula = "IF($B$5=""All"",IF($E$5=""All""," & CountExpr(pairs) & "," & CountExpr(pairs & teamPairs) & ")," & _
        "IF($E$5=""All""," & CountExpr(pairs & regionPairs) & "," & CountExpr(pairs & regionPairs & teamPairs) & "))"
End Function

Private Function SelectedCount(ByVal additionalPairs As Variant) As String
    Dim pairs As String, i As Long
    pairs = DataRange("A") & "," & Quoted("<>")
    For i = LBound(additionalPairs) To UBound(additionalPairs) Step 2
        pairs = pairs & "," & DataRange(CStr(additionalPairs(i))) & "," & Criterion(CStr(additionalPairs(i + 1)))
    Next i
    SelectedCount = SelectedFormula(pairs)
End Function

Private Function ReportCriteria(ByVal report As String, Optional ByVal aged As Boolean = False) As String
    ReportCriteria = PairList("B", report)
    If IsAmountReport(report) Then ReportCriteria = ReportCriteria & "," & PairList("V", "Yes")
    If aged Then ReportCriteria = ReportCriteria & "," & DataRange("K") & ","">=30"""
End Function

Private Function DashboardReportCriteria(ByVal report As String, Optional ByVal ageBand As Long = 0) As String
    Dim pairs As String
    pairs = ReportCriteria(report)
    If report = "OMS" Then pairs = pairs & "," & PairList("AC", "Yes")
    Select Case ageBand
        Case 1: pairs = pairs & "," & PairList("K", ">=0", "K", "<=30")
        Case 2: pairs = pairs & "," & PairList("K", ">30", "K", "<=60")
        Case 3: pairs = pairs & "," & PairList("K", ">60", "K", "<=90")
        Case 4: pairs = pairs & "," & PairList("K", ">90")
    End Select
    DashboardReportCriteria = pairs
End Function

Private Function SelectedDashboardCount(ByVal report As String, Optional ByVal ageBand As Long = 0) As String
    SelectedDashboardCount = SelectedFormula(DashboardReportCriteria(report, ageBand))
End Function

Private Function SelectedReportCount(ByVal report As String, ByVal aged As Boolean) As String
    SelectedReportCount = SelectedFormula(ReportCriteria(report, aged))
End Function

Private Function SelectedAllReports(Optional ByVal aged As Boolean = False) As String
    Dim report As Variant, result As String
    For Each report In ReportNames()
        If Len(result) > 0 Then result = result & "+"
        result = result & SelectedReportCount(CStr(report), aged)
    Next report
    SelectedAllReports = "=" & result
End Function

Private Function ReportRegionCount(ByVal report As String, ByVal region As String) As String
    ReportRegionCount = CountExpr(ReportCriteria(report) & "," & PairList("C", region))
End Function

' Optional temporary checkpoint for this stage.

Public Sub TypingCheckFormula()
    mDataLastRow = 11
    Debug.Print "=" & CountExpr(PairList("B", "eTR"))
    ' Expected: =COUNTIFS('Normalized Data'!$B$2:$B$11,"eTR")
    ' RefreshOperations will replace this test range with the real data bounds.
End Sub

' ========================================================================
' STAGE 7: DASHBOARD AND SUMMARIES
' ========================================================================
' Run TypingCheckEmptyOutputs on the practice copy. It backs up and renders zero-entry summaries from 
' your configuration. Review missing-import notices are expected.
' Procedure typing order:
' 01. RadarReportColor
' 02. GreyRange
' 03. ApplyDashboardBau
' 04. ReapplyTeamSummaryBau
' 05. ColumnTotal
' 06. TotalStyle
' 07. WriteDashboardTeam
' 08. PutCard
' 09. SetWorkbookName
' 10. HasRegion
' 11. UpdateSelectorLists
' 12. IsValidTeamFilter
' 13. ActiveTeamCount
' 14. RegionDisplay
' 15. ReportDisplayName
' 16. ExistingTeamSelections
' 17. ExistingTeamRadarSelections
' 18. TeamMemberCount
' 19. IsTrackedIssueValue
' 20. RadarRecordIncluded
' 21. RadarMetricRaw
' 22. TeamRadarMetric
' 23. MemberRadarMetric
' 24. Top20Benchmark
' 25. RadarIndex
' 26. RadarLabels
' 27. RadarReports
' 28. RadarMeasures
' 29. RadarDetailLabel
' 30. RadarDetailRow
' 31. WriteRadarDetails
' 32. AddRadarReportBackgrounds
' 33. StyleTeamRadar
' 34. BuildTeamRadar
' 35. ApplyTeamSummaryVisibility
' 36. ApplyTeamSummarySelection
' 37. WriteDashboard
' 38. FinishSheet
' 39. ReviewCount
' 40. WriteTeamSummary
' 41. WriteEmployeeSummary
' 42. ActiveLayoutCount
' 43. PutCheck
' 44. WriteValidation
' 45. UpdateReadMe
' 46. WriteRefreshLog

Private Function RadarReportColor(ByVal report As String) As Long
    Select Case report
        Case "Chase": RadarReportColor = RGB(84, 143, 255)
        Case "Chase CAD": RadarReportColor = RGB(64, 211, 205)
        Case "Internal": RadarReportColor = RGB(126, 207, 103)
        Case "Internal HSBC": RadarReportColor = RGB(184, 135, 238)
        Case "OMS": RadarReportColor = RGB(255, 171, 64)
        Case "eTR": RadarReportColor = RGB(242, 103, 103)
        Case Else: RadarReportColor = RGB(110, 120, 132)
    End Select
End Function

Private Sub GreyRange(ByVal target As Range)
    target.Interior.Color = COLOR_GREY
    target.Font.Color = RGB(102, 102, 102)
End Sub

Private Sub ApplyDashboardBau(ByVal ws As Worksheet, ByVal teamID As String, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim c As Long, metric As Object, issue As Object
    c = 2
    For Each metric In mAmountMetrics
        If Not BauFlag(teamID, metric("report")) Then GreyRange ws.Range(ws.Cells(firstRow, c), ws.Cells(lastRow, c))
        c = c + 1
    Next metric
    For Each issue In mIssues
        If Not BauFlag(teamID, "OMS") Then GreyRange ws.Range(ws.Cells(firstRow, c), ws.Cells(lastRow, c))
        c = c + 1
    Next issue
    If Not BauFlag(teamID, "eTR") Then GreyRange ws.Range(ws.Cells(firstRow, c), ws.Cells(lastRow, c))
End Sub

Private Sub ReapplyTeamSummaryBau(ByVal ws As Worksheet)
    Dim metaRow As Long, lastMeta As Long, sectionStart As Long, managerRow As Long
    Dim teamID As String
    lastMeta = ws.Cells(ws.Rows.Count, TEAM_META_COLUMN).End(xlUp).Row
    For metaRow = 2 To lastMeta
        teamID = TextOf(ws.Cells(metaRow, TEAM_META_COLUMN).Value2)
        If Len(teamID) > 0 Then
            If mParticipants.Exists(teamID) Then
                sectionStart = CLng(ws.Cells(metaRow, TEAM_META_COLUMN + 1).Value2)
                managerRow = sectionStart + 2 + mParticipants(teamID).Count
                ApplyDashboardBau ws, teamID, sectionStart + 1, managerRow
            End If
        End If
    Next metaRow
End Sub

Private Function ColumnTotal(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal totalRow As Long, ByVal column As Long) As String
    ColumnTotal = "=0"
    If totalRow > firstRow Then ColumnTotal = "=SUM(" & ws.Range(ws.Cells(firstRow, column), ws.Cells(totalRow - 1, column)).Address(False, False) & ")"
End Function

Private Sub TotalStyle(ByVal target As Range)
    With target
        .Font.Bold = True: .Interior.Color = RGB(221, 235, 247)
        .Borders(xlEdgeTop).LineStyle = xlDouble: .Borders(xlEdgeTop).Color = COLOR_NAVY
    End With
End Sub

Private Sub WriteDashboardTeam(ByVal ws As Worksheet, ByVal teamID As String, ByRef startRow As Long, ByVal lastCol As Long)
    Dim team As Object, list As Collection, p As Object, metric As Object, issue As Object
    Dim headerRow As Long, firstRow As Long, managerRow As Long, row As Long, c As Long, reviews As Long
    Set team = mTeams(teamID): Set list = mParticipants(teamID)
    SectionBand ws, startRow, lastCol, team("name") & "   Manager: " & team("manager")
    ws.Rows(startRow).RowHeight = 21
    headerRow = startRow + 1
    ws.Cells(headerRow, 1).Value2 = "Employee / Manager"
    c = 2
    For Each metric In mAmountMetrics
        ws.Cells(headerRow, c).Value2 = LiteralValue(metric("label")): c = c + 1
    Next metric
    For Each issue In mIssues
        ws.Cells(headerRow, c).Formula = IssueHeaderFormula(issue): c = c + 1
    Next issue
    ws.Cells(headerRow, c).Value2 = "eTR Outstanding"
    HeaderStyle ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol))
    ws.Rows(headerRow).RowHeight = 42
    firstRow = headerRow + 1: row = firstRow
    For Each p In list
        If p("review") Then
            ws.Cells(row, 1).Value2 = LiteralValue(p("employee") & " (Review team setting)")
            ws.Cells(row, 1).Interior.Color = COLOR_AMBER
            ws.Cells(row, 1).Font.Color = RGB(156, 87, 0)
            reviews = reviews + 1
        Else
            ws.Cells(row, 1).Value2 = LiteralValue(p("employee"))
        End If
        c = 2
        For Each metric In mAmountMetrics
            ws.Cells(row, c).Formula = "=" & MetricCount(teamID, p, metric): c = c + 1
        Next metric
        For Each issue In mIssues
            ws.Cells(row, c).Formula = "=" & OmsIssueCount(teamID, p, issue): c = c + 1
        Next issue
        ws.Cells(row, c).Formula = "=" & TeamReportCount(teamID, "eTR", p)
        row = row + 1
    Next p
    managerRow = row
    ws.Cells(managerRow, 1).Value2 = LiteralValue("Manager total: " & team("manager") & _
        IIf(reviews > 0, " (" & reviews & " team-setting reviews)", ""))
    For c = 2 To lastCol
        ws.Cells(managerRow, c).Formula = ColumnTotal(ws, firstRow, managerRow, c)
    Next c
    TotalStyle ws.Range(ws.Cells(managerRow, 1), ws.Cells(managerRow, lastCol))
    ws.Cells(managerRow, 1).WrapText = True
    ws.Rows(managerRow).RowHeight = IIf(reviews > 0, 26, 18)
    ws.Range(ws.Cells(firstRow, 2), ws.Cells(managerRow, lastCol)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(firstRow, 2), ws.Cells(managerRow, lastCol)).HorizontalAlignment = xlRight
    ApplyDashboardBau ws, teamID, headerRow, managerRow
    startRow = managerRow + 1
End Sub

Private Sub PutCard(ByVal ws As Worksheet, ByVal address As String, ByVal label As String, ByVal formula As String)
    Dim target As Range, title As Range, value As Range
    Set target = ws.Range(address)
    Set title = target.Rows(1): Set value = target.Offset(1, 0).Resize(target.Rows.Count - 1, target.Columns.Count)
    title.Merge: value.Merge
    title.Value2 = label: title.Font.Size = 10: title.Font.Bold = True
    value.Formula = formula: value.Font.Size = 22: value.Font.Bold = True
    value.NumberFormat = "#,##0"
    target.Interior.Color = RGB(221, 235, 247): target.Font.Color = COLOR_NAVY
    target.HorizontalAlignment = xlCenter: target.VerticalAlignment = xlCenter
    ws.Rows(target.Row).RowHeight = 20
    ws.Rows(target.Row + 1).RowHeight = 24: ws.Rows(target.Row + 2).RowHeight = 24
End Sub

Private Sub SetWorkbookName(ByVal name As String, ByVal refersTo As String)
    Dim existing As Excel.Name
    On Error Resume Next
    Set existing = ThisWorkbook.Names(name)
    On Error GoTo 0
    If existing Is Nothing Then
        ThisWorkbook.Names.Add Name:=name, RefersTo:=refersTo
    Else
        existing.RefersTo = refersTo
    End If
End Sub

Private Function HasRegion(ByVal region As String) As Boolean
    Dim id As Variant
    For Each id In mTeamOrder
        If mTeams(id)("region") = region Then HasRegion = True: Exit Function
    Next id
End Function

Private Sub UpdateSelectorLists()
    Dim ws As Worksheet, id As Variant, row As Long, lastRow As Long
    Set ws = RequiredSheet("Configuration")
    lastRow = Application.Max(LastUsedRow(ws), mTeams.Count + 2, 5)
    ws.Range("O2:P" & lastRow).ClearContents
    ws.Range("O1").Value2 = "Dashboard regions": ws.Range("P1").Value2 = "Dashboard teams"
    ws.Range("O2:O4").Value2 = Application.Transpose(Array("All", "NY", "Toronto"))
    If HasRegion("Unassigned") Then ws.Range("O5").Value2 = "Unassigned"
    ws.Range("P2").Value2 = "All": row = 3
    For Each id In mTeamOrder
        ws.Cells(row, 16).Value2 = LiteralValue(mTeams(id)("name")): row = row + 1
    Next id
    SetWorkbookName "OpsRegionList", "='Configuration'!$O$2:$O$" & IIf(HasRegion("Unassigned"), 5, 4)
    SetWorkbookName "OpsTeamList", "='Configuration'!$P$2:$P$" & row - 1
    ws.Range("O1:P1").Font.Bold = True
    ws.Columns("O:P").ColumnWidth = 25
End Sub

Private Function IsValidTeamFilter(ByVal name As String) As Boolean
    Dim id As Variant
    If name = "All" Then IsValidTeamFilter = True: Exit Function
    For Each id In mTeamOrder
        If mTeams(id)("name") = name Then IsValidTeamFilter = True: Exit Function
    Next id
End Function

Private Function ActiveTeamCount() As Long
    ActiveTeamCount = mTeams.Count
    If mTeams.Exists(UNASSIGNED_TEAM) Then ActiveTeamCount = ActiveTeamCount - 1
End Function

Private Function RegionDisplay(ByVal region As String) As String
    RegionDisplay = region
    If region = "NY" Then RegionDisplay = "New York"
End Function

Private Function ReportDisplayName(ByVal report As String) As String
    Select Case report
        Case "OMS": ReportDisplayName = "OMS tracked issues"
        Case "eTR": ReportDisplayName = "eTR"
        Case "Chase": ReportDisplayName = "Chase USD"
        Case "Chase CAD": ReportDisplayName = "Chase CAD"
        Case "Internal": ReportDisplayName = "Internal accounts"
        Case "Internal HSBC": ReportDisplayName = "HSBC internal accounts"
    End Select
End Function

Private Function ExistingTeamSelections(ByVal ws As Worksheet) As Object
    Dim result As Object, lo As ListObject, row As Long, teamID As String
    Set result = NewMap()
    Set lo = TableOrNothing("TeamSelectionTable")
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            For row = 1 To lo.DataBodyRange.Rows.Count
                teamID = TextOf(lo.ListColumns("Team ID").DataBodyRange.Cells(row, 1).Value2)
                If Len(teamID) > 0 Then result(teamID) = YesFlag(lo.ListColumns("Show").DataBodyRange.Cells(row, 1).Value2)
            Next row
        End If
    End If
    Set ExistingTeamSelections = result
End Function

Private Function ExistingTeamRadarSelections(ByVal ws As Worksheet) As Object
    Dim result As Object, lo As ListObject, radarColumn As ListColumn
    Dim row As Long, teamID As String
    Set result = NewMap()
    Set lo = TableOrNothing("TeamSelectionTable")
    If Not lo Is Nothing Then
        On Error Resume Next
        Set radarColumn = lo.ListColumns("Radar")
        On Error GoTo 0
        If Not lo.DataBodyRange Is Nothing Then
            For row = 1 To lo.DataBodyRange.Rows.Count
                teamID = TextOf(lo.ListColumns("Team ID").DataBodyRange.Cells(row, 1).Value2)
                If Len(teamID) > 0 Then
                    If radarColumn Is Nothing Then
                        result(teamID) = True
                    Else
                        result(teamID) = YesFlag(radarColumn.DataBodyRange.Cells(row, 1).Value2)
                    End If
                End If
            Next row
        End If
    End If
    Set ExistingTeamRadarSelections = result
End Function

Private Function TeamMemberCount(ByVal teamID As String) As Long
    Dim memberID As Variant
    For Each memberID In mMembers.Keys
        If mMembers(memberID)("team") = teamID Then TeamMemberCount = TeamMemberCount + 1
    Next memberID
End Function

Private Function IsTrackedIssueValue(ByVal value As Variant) As Boolean
    Dim issue As Object
    If IsEmpty(value) Or Not IsNumeric(value) Then Exit Function
    For Each issue In mIssues
        If YesFlag(ValueOf(issue, "Active")) Then
            If CLng(ValueOf(issue, "Issue Code")) = CLng(value) Then IsTrackedIssueValue = True: Exit Function
        End If
    Next issue
End Function

Private Function RadarRecordIncluded(ByVal rec As Object, ByVal report As String, _
                                     ByVal teamID As String, ByVal memberID As String) As Boolean
    If rec("report") <> report Then Exit Function
    If Len(teamID) > 0 Then If rec("team") <> teamID Then Exit Function
    If Len(memberID) > 0 Then
        If rec("member") <> memberID Then Exit Function
        If rec("team") <> mMembers(memberID)("team") Then Exit Function
    End If
    If IsAmountReport(report) Then
        If IsEmpty(rec("amount")) Then Exit Function
        If Abs(CDbl(rec("amount"))) < mAmountThreshold Then Exit Function
    ElseIf report = "OMS" Then
        If Not IsTrackedIssueValue(rec("issue")) Then Exit Function
    End If
    RadarRecordIncluded = True
End Function

Private Function RadarMetricRaw(ByVal report As String, ByVal measure As String, _
                                ByVal teamID As String, ByVal memberID As String) As Variant
    Dim rec As Object, itemCount As Long, observed As Long, total As Double
    For Each rec In mRecords
        If RadarRecordIncluded(rec, report, teamID, memberID) Then
            itemCount = itemCount + 1
            Select Case measure
                Case "Age"
                    If Not IsEmpty(rec("age")) Then total = total + CDbl(rec("age")): observed = observed + 1
                Case "AbsAmount"
                    If Not IsEmpty(rec("amount")) Then total = total + Abs(CDbl(rec("amount"))): observed = observed + 1
                Case "Count"
                Case Else
                    Fail "Unknown radar measure: " & measure
            End Select
        End If
    Next rec
    If measure = "Count" Then
        RadarMetricRaw = CDbl(itemCount)
    ElseIf observed > 0 Then
        RadarMetricRaw = total / observed
    End If
End Function

Private Function TeamRadarMetric(ByVal teamID As String, ByVal report As String, ByVal measure As String) As Variant
    Dim value As Variant, members As Long
    If Not BauFlag(teamID, report) Then Exit Function
    value = RadarMetricRaw(report, measure, teamID, "")
    If measure = "Count" Then
        members = TeamMemberCount(teamID)
        If members = 0 Then Exit Function
        TeamRadarMetric = CDbl(value) / members
    ElseIf Not IsEmpty(value) Then
        TeamRadarMetric = value
    End If
End Function

Private Function MemberRadarMetric(ByVal memberID As String, ByVal report As String, ByVal measure As String) As Variant
    Dim teamID As String
    teamID = mMembers(memberID)("team")
    If Not BauFlag(teamID, report) Then Exit Function
    MemberRadarMetric = RadarMetricRaw(report, measure, teamID, memberID)
End Function

Private Function Top20Benchmark(ByVal report As String, ByVal measure As String) As Variant
    Dim values() As Double, memberID As Variant, value As Variant
    Dim n As Long, i As Long, j As Long, take As Long, total As Double, swap As Double, key As String
    key = report & "|" & measure
    If mRadarBenchmarks.Exists(key) Then Top20Benchmark = mRadarBenchmarks(key): Exit Function
    If mMembers.Count = 0 Then mRadarBenchmarks(key) = Empty: Exit Function
    ReDim values(1 To mMembers.Count)
    For Each memberID In mMembers.Keys
        If BauFlag(mMembers(memberID)("team"), report) Then
            value = MemberRadarMetric(CStr(memberID), report, measure)
            If measure = "Count" Or Not IsEmpty(value) Then
                n = n + 1: values(n) = CDbl(value)
            End If
        End If
    Next memberID
    If n = 0 Then mRadarBenchmarks(key) = Empty: Exit Function
    For i = 2 To n
        swap = values(i): j = i - 1
        Do While j >= 1
            If values(j) <= swap Then Exit Do
            values(j + 1) = values(j): j = j - 1
        Loop
        values(j + 1) = swap
    Next i
    take = (n + 4) \ 5
    For i = 1 To take: total = total + values(i): Next i
    Top20Benchmark = total / take
    mRadarBenchmarks(key) = Top20Benchmark
End Function

Private Function RadarIndex(ByVal actual As Variant, ByVal benchmark As Variant) As Variant
    Dim score As Double
    If IsEmpty(actual) Or IsEmpty(benchmark) Then RadarIndex = CVErr(xlErrNA): Exit Function
    If CDbl(benchmark) = 0 Then
        If CDbl(actual) = 0 Then RadarIndex = 100 Else RadarIndex = RADAR_MAX_INDEX
        Exit Function
    End If
    score = 100 * CDbl(actual) / CDbl(benchmark)
    If score > RADAR_MAX_INDEX Then score = RADAR_MAX_INDEX
    RadarIndex = score
End Function

Private Function RadarLabels(ByVal region As String) As Variant
    If region = "NY" Then
        RadarLabels = Array("OMS count/member", "OMS avg age", "Chase count/member", "Chase avg value", _
            "Chase avg age", "CAD count/member", "CAD avg value", "CAD avg age", _
            "eTR count/member", "eTR avg age")
    ElseIf region = "Toronto" Then
        RadarLabels = Array("OMS count/member", "OMS avg age", "Internal count/member", "Internal avg value", _
            "Internal avg age", "HSBC count/member", "HSBC avg value", "HSBC avg age", _
            "eTR count/member", "eTR avg age")
    End If
End Function

Private Function RadarReports(ByVal region As String) As Variant
    If region = "NY" Then
        RadarReports = Array("OMS", "OMS", "Chase", "Chase", "Chase", "Chase CAD", "Chase CAD", "Chase CAD", "eTR", "eTR")
    ElseIf region = "Toronto" Then
        RadarReports = Array("OMS", "OMS", "Internal", "Internal", "Internal", _
            "Internal HSBC", "Internal HSBC", "Internal HSBC", "eTR", "eTR")
    End If
End Function

Private Function RadarMeasures(ByVal region As String) As Variant
    If region = "NY" Or region = "Toronto" Then
        RadarMeasures = Array("Count", "Age", "Count", "AbsAmount", "Age", "Count", "AbsAmount", "Age", "Count", "Age")
    End If
End Function

Private Function RadarDetailLabel(ByVal report As String, ByVal measure As String) As String
    Dim label As String
    label = report
    If report = "OMS" Then label = "OMS tracked"
    If report = "Chase" Then label = "Chase USD"
    If report = "Internal HSBC" Then label = "HSBC Internal"
    Select Case measure
        Case "Count": label = label & " - items / member"
        Case "Age": label = label & " - avg age (days)"
        Case "AbsAmount": label = label & " - avg absolute amount"
    End Select
    RadarDetailLabel = label
End Function

Private Sub RadarDetailRow(ByVal ws As Worksheet, ByVal row As Long)
    ' Three presentation columns, without widening the employee table above.
    ws.Range(ws.Cells(row, RADAR_DETAIL_COLUMN), ws.Cells(row, RADAR_DETAIL_COLUMN + 3)).Merge
    ws.Range(ws.Cells(row, RADAR_DETAIL_COLUMN + 4), ws.Cells(row, RADAR_DETAIL_COLUMN + 5)).Merge
    ws.Range(ws.Cells(row, RADAR_DETAIL_COLUMN + 6), ws.Cells(row, RADAR_LAST_COLUMN)).Merge
End Sub

Private Sub WriteRadarDetails(ByVal ws As Worksheet, ByVal teamID As String, ByVal topRow As Long, _
                              ByVal firstHelper As Long, ByVal reports As Variant, ByVal measures As Variant, _
                              ByRef bottomRow As Long)
    Dim i As Long, row As Long, sourceRow As Long, col As Long, sourceCol As Long
    Dim count As Long, sourceAddress As String, numberStyle As String, note As String
    count = UBound(reports) - LBound(reports) + 1
    bottomRow = topRow + count + 4
    With ws.Range(ws.Cells(topRow, RADAR_DETAIL_COLUMN), ws.Cells(bottomRow, RADAR_LAST_COLUMN))
        .Interior.Color = vbWhite: .Font.Name = "Arial": .Font.Size = 10
        .Font.Color = RGB(51, 65, 85): .VerticalAlignment = xlCenter
        .WrapText = True: .Rows.RowHeight = 26
    End With
    With ws.Range(ws.Cells(topRow, RADAR_DETAIL_COLUMN), ws.Cells(topRow, RADAR_LAST_COLUMN))
        .Merge: .Value2 = "Radar detail": .Font.Bold = True: .Font.Color = COLOR_NAVY
        .RowHeight = 22
    End With
    RadarDetailRow ws, topRow + 1
    ws.Cells(topRow + 1, RADAR_DETAIL_COLUMN).Value2 = "Metric"
    ws.Cells(topRow + 1, RADAR_DETAIL_COLUMN + 4).Value2 = "Team actual"
    ws.Cells(topRow + 1, RADAR_DETAIL_COLUMN + 6).Value2 = "Top 20% benchmark"
    With ws.Range(ws.Cells(topRow + 1, RADAR_DETAIL_COLUMN), ws.Cells(topRow + 1, RADAR_LAST_COLUMN))
        .Interior.Color = RGB(241, 245, 249): .Font.Bold = True: .RowHeight = 32
        .HorizontalAlignment = xlCenter
        .Borders(xlEdgeBottom).Color = COLOR_LINE: .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    ws.Cells(topRow + 1, RADAR_DETAIL_COLUMN).HorizontalAlignment = xlLeft
    ws.Cells(topRow + 1, RADAR_DETAIL_COLUMN + 4).Font.Color = RGB(196, 75, 17)
    ws.Cells(topRow + 1, RADAR_DETAIL_COLUMN + 6).Font.Color = RGB(30, 64, 175)
    For i = LBound(reports) To UBound(reports)
        row = topRow + 2 + i - LBound(reports): sourceRow = firstHelper + i - LBound(reports)
        RadarDetailRow ws, row
        ws.Cells(row, RADAR_DETAIL_COLUMN).Value2 = LiteralValue(RadarDetailLabel(CStr(reports(i)), CStr(measures(i))))
        numberStyle = "#,##0.00"
        If measures(i) = "Age" Then numberStyle = "0.0"
        If measures(i) = "AbsAmount" Then numberStyle = "#,##0"
        For col = 0 To 1
            sourceCol = RADAR_DATA_COLUMN + 1 + col
            sourceAddress = ws.Cells(sourceRow, sourceCol).Address(True, True)
            With ws.Cells(row, RADAR_DETAIL_COLUMN + 4 + 2 * col)
                .Formula = "=IF(ISNUMBER(" & sourceAddress & ")," & sourceAddress & ",""n/a"")"
                .NumberFormat = numberStyle: .HorizontalAlignment = xlRight
            End With
        Next col
        With ws.Range(ws.Cells(row, RADAR_DETAIL_COLUMN), ws.Cells(row, RADAR_LAST_COLUMN))
            .Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            If Not BauFlag(teamID, CStr(reports(i))) Then GreyRange .Cells
        End With
    Next i
    For i = 0 To 2
        row = topRow + count + 2 + i
        Select Case i
            Case 0: note = "Benchmark = 1.0x. Radar capped at 8.0x; raw values above are uncapped."
            Case 1: note = "Counts are per configured member. n/a means unavailable or non-BAU."
            Case 2
                If mTeams(teamID)("region") = "NY" Then
                    note = "Amounts: Chase USD in USD; Chase CAD in CAD. Refresh after input changes."
                Else
                    note = "Internal amounts combine source currencies; no FX conversion. Refresh after input changes."
                End If
        End Select
        With ws.Range(ws.Cells(row, RADAR_DETAIL_COLUMN), ws.Cells(row, RADAR_LAST_COLUMN))
            .Merge: .Value2 = note: .RowHeight = 24: .Font.Color = RGB(100, 116, 139)
        End With
    Next i
End Sub

Private Function AddRadarReportBackgrounds(ByVal radar As Chart, ByVal categoryRange As Range, _
                                           ByVal reports As Variant) As Long
    Dim band As Series, bandValues() As Variant
    Dim i As Long, j As Long, pointCount As Long, created As Long
    Dim reportName As String, previousReport As String
    pointCount = UBound(reports) - LBound(reports) + 1
    ' Background bands are cosmetic and must never interrupt a refresh.
    On Error Resume Next
    For i = LBound(reports) To UBound(reports)
        reportName = CStr(reports(i))
        If reportName <> previousReport Then
            ReDim bandValues(1 To pointCount)
            For j = 1 To pointCount
                If CStr(reports(LBound(reports) + j - 1)) = reportName Then
                    bandValues(j) = RADAR_MAX_INDEX
                Else
                    bandValues(j) = 0
                End If
            Next j
            Set band = Nothing
            Set band = radar.SeriesCollection.NewSeries
            If Not band Is Nothing Then
                With band
                    .Name = reportName & " background"
                    .XValues = categoryRange
                    .Values = bandValues
                    .ChartType = xlRadarFilled
                    .Format.Fill.Visible = msoTrue: .Format.Fill.Solid
                    .Format.Fill.ForeColor.RGB = RadarReportColor(reportName)
                    .Format.Fill.Transparency = 0.9
                    .Format.Line.Visible = msoFalse
                    .MarkerStyle = xlMarkerStyleNone
                End With
                created = created + 1
            End If
            previousReport = reportName
        End If
    Next i
    On Error GoTo 0
    AddRadarReportBackgrounds = created
End Function

Private Sub StyleTeamRadar(ByVal radar As Chart, ByVal teamSeries As Series, _
                           ByVal benchmarkSeries As Series, ByVal teamName As String, _
                           ByVal backgroundCount As Long)
    Dim i As Long
    ' Everything in this procedure is cosmetic and must never interrupt a refresh.
    On Error Resume Next
    radar.HasTitle = True
    radar.ChartTitle.Text = teamName & vbLf & "Workload vs top 20% benchmark"
    radar.HasLegend = True: radar.Legend.Position = xlLegendPositionBottom
    radar.ChartArea.Format.Fill.Solid
    radar.ChartArea.Format.Fill.ForeColor.RGB = vbWhite
    radar.ChartArea.Format.Line.ForeColor.RGB = RGB(226, 232, 240)
    radar.PlotArea.Format.Fill.Solid
    radar.PlotArea.Format.Fill.ForeColor.RGB = vbWhite
    radar.PlotArea.Format.Line.Visible = msoFalse
    With teamSeries
        .Name = "Team actual"
        .Format.Fill.Visible = msoTrue: .Format.Fill.Solid
        .Format.Fill.ForeColor.RGB = RGB(196, 75, 17): .Format.Fill.Transparency = 0.92
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(196, 75, 17)
        .Format.Line.Weight = 3.5: .Format.Line.DashStyle = msoLineSolid
        .Border.Color = RGB(196, 75, 17)
        .Border.LineStyle = xlContinuous: .Border.Weight = xlThick
        .MarkerStyle = xlMarkerStyleCircle: .MarkerSize = 9
        .MarkerForegroundColor = RGB(196, 75, 17): .MarkerBackgroundColor = vbWhite
    End With
    With benchmarkSeries
        .Name = "Top 20% benchmark (1.0x)"
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(30, 64, 175)
        .Format.Line.Weight = 2.75: .Format.Line.DashStyle = msoLineDash
        .Border.Color = RGB(30, 64, 175)
        .Border.LineStyle = xlDash: .Border.Weight = xlMedium
        .MarkerStyle = xlMarkerStyleDiamond: .MarkerSize = 6
        .MarkerForegroundColor = RGB(30, 64, 175)
        .MarkerBackgroundColor = vbWhite
    End With
    With radar.Axes(xlValue)
        .MinimumScale = 0: .MaximumScale = RADAR_MAX_INDEX: .MajorUnit = 200
        .TickLabelPosition = xlTickLabelPositionNone
        .Format.Line.ForeColor.RGB = RGB(203, 213, 225)
        .MajorGridlines.Format.Line.ForeColor.RGB = RGB(203, 213, 225)
        .MajorGridlines.Format.Line.Transparency = 0.15
    End With
    With radar.ChartArea.Font
        .Name = "Arial": .Size = 10: .Color = RGB(51, 65, 85)
    End With
    With radar.ChartTitle.Format.TextFrame2.TextRange.Font
        .Name = "Arial": .Size = 12: .Bold = msoTrue
        .Fill.ForeColor.RGB = COLOR_NAVY
    End With
    With radar.Legend.Font
        .Name = "Arial": .Size = 10: .Color = RGB(51, 65, 85)
    End With
    ' Keep only the actual and benchmark series in the legend.
    For i = backgroundCount To 1 Step -1
        radar.Legend.LegendEntries(i).Delete
    Next i
    On Error GoTo 0
End Sub

Private Function BuildTeamRadar(ByVal ws As Worksheet, ByVal teamID As String, ByVal topRow As Long, _
                                ByVal leftCol As Long, ByRef helperRow As Long, ByVal chartIndex As Long, _
                                ByRef bottomRow As Long) As String
    Dim labels As Variant, reports As Variant, measures As Variant, actual As Variant, benchmark As Variant
    Dim firstHelper As Long, i As Long, chartName As String, chartWidth As Double, chartHeight As Double
    Dim chartBox As ChartObject, radar As Chart, teamSeries As Series, benchmarkSeries As Series
    Dim backgroundCount As Long
    labels = RadarLabels(mTeams(teamID)("region"))
    If IsEmpty(labels) Then Exit Function
    reports = RadarReports(mTeams(teamID)("region")): measures = RadarMeasures(mTeams(teamID)("region"))
    firstHelper = helperRow
    For i = LBound(labels) To UBound(labels)
        actual = TeamRadarMetric(teamID, CStr(reports(i)), CStr(measures(i)))
        benchmark = Top20Benchmark(CStr(reports(i)), CStr(measures(i)))
        ws.Cells(helperRow, RADAR_DATA_COLUMN).Value2 = LiteralValue(CStr(labels(i)))
        If Not IsEmpty(actual) Then ws.Cells(helperRow, RADAR_DATA_COLUMN + 1).Value2 = actual
        If Not IsEmpty(benchmark) Then ws.Cells(helperRow, RADAR_DATA_COLUMN + 2).Value2 = benchmark
        ws.Cells(helperRow, RADAR_DATA_COLUMN + 3).Value = RadarIndex(actual, benchmark)
        If IsEmpty(benchmark) Then
            ws.Cells(helperRow, RADAR_DATA_COLUMN + 4).Value = CVErr(xlErrNA)
        Else
            ws.Cells(helperRow, RADAR_DATA_COLUMN + 4).Value2 = 100
        End If
        helperRow = helperRow + 1
    Next i
    WriteRadarDetails ws, teamID, topRow, firstHelper, reports, measures, bottomRow
    chartWidth = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(topRow, RADAR_DETAIL_COLUMN - 2)).Width
    chartHeight = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(bottomRow, leftCol)).Height
    chartName = "OpsTeamRadar_" & CStr(chartIndex)
    Set chartBox = ws.ChartObjects.Add(ws.Cells(topRow, leftCol).Left, ws.Cells(topRow, leftCol).Top, chartWidth, chartHeight)
    chartBox.Name = chartName: chartBox.Placement = xlMove
    Set radar = chartBox.Chart
    radar.ChartType = xlRadarFilled: radar.PlotVisibleOnly = False: radar.DisplayBlanksAs = xlNotPlotted
    backgroundCount = AddRadarReportBackgrounds(radar, _
        ws.Range(ws.Cells(firstHelper, RADAR_DATA_COLUMN), ws.Cells(helperRow - 1, RADAR_DATA_COLUMN)), reports)
    Set teamSeries = radar.SeriesCollection.NewSeries
    With teamSeries
        .Name = "Team actual"
        .XValues = ws.Range(ws.Cells(firstHelper, RADAR_DATA_COLUMN), ws.Cells(helperRow - 1, RADAR_DATA_COLUMN))
        .Values = ws.Range(ws.Cells(firstHelper, RADAR_DATA_COLUMN + 3), ws.Cells(helperRow - 1, RADAR_DATA_COLUMN + 3))
    End With
    Set benchmarkSeries = radar.SeriesCollection.NewSeries
    With benchmarkSeries
        .Name = "Top 20% benchmark (1.0x)"
        .XValues = teamSeries.XValues
        .Values = ws.Range(ws.Cells(firstHelper, RADAR_DATA_COLUMN + 4), ws.Cells(helperRow - 1, RADAR_DATA_COLUMN + 4))
    End With
    StyleTeamRadar radar, teamSeries, benchmarkSeries, CStr(mTeams(teamID)("name")), backgroundCount
    BuildTeamRadar = chartName
    helperRow = helperRow + 1
End Function

Private Sub ApplyTeamSummaryVisibility(ByVal ws As Worksheet, ByVal selections As Object)
    Dim row As Long, lastMeta As Long
    Dim teamID As String, startRow As Long, endRow As Long, chartName As String, showTeam As Boolean
    Dim chartBox As ChartObject
    lastMeta = ws.Cells(ws.Rows.Count, TEAM_META_COLUMN).End(xlUp).Row
    For row = 2 To lastMeta
        teamID = TextOf(ws.Cells(row, TEAM_META_COLUMN).Value2)
        If Len(teamID) > 0 Then
            startRow = CLng(ws.Cells(row, TEAM_META_COLUMN + 1).Value2)
            endRow = CLng(ws.Cells(row, TEAM_META_COLUMN + 2).Value2)
            chartName = TextOf(ws.Cells(row, TEAM_META_COLUMN + 3).Value2)
            showTeam = True
            If selections.Exists(teamID) Then showTeam = CBool(selections(teamID))
            ws.Rows(startRow & ":" & endRow).Hidden = Not showTeam
            If Len(chartName) > 0 Then
                Set chartBox = Nothing
                On Error Resume Next
                Set chartBox = ws.ChartObjects(chartName)
                On Error GoTo 0
                If Not chartBox Is Nothing Then chartBox.Visible = showTeam
            End If
        End If
    Next row
End Sub

Public Sub ApplyTeamSummarySelection()
    Dim oldScreen As Boolean, stateSaved As Boolean, selectionTable As ListObject
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating: stateSaved = True
    Application.ScreenUpdating = False
    Set selectionTable = TableOrNothing("TeamSelectionTable")
    If selectionTable Is Nothing Then Fail "Refresh all teams before applying a Team Summary selection."
    If mTeams Is Nothing Or mParticipants Is Nothing Or mAmountMetrics Is Nothing Or mIssues Is Nothing Or mRecords Is Nothing Or mRadarBenchmarks Is Nothing Then
        Fail "Run Refresh all teams once before changing Team Summary radar selections."
    End If
    WriteTeamSummary
    RequiredSheet("Team Summary").Calculate
    RequiredSheet("Team Summary").Activate
    Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    If stateSaved Then Application.ScreenUpdating = oldScreen
    MsgBox Err.Description, vbExclamation, "Team selection"
End Sub

Private Sub WriteDashboard()
    Dim ws As Worksheet, regionFilter As String, teamFilter As String, r As Long, report As Variant
    Dim id As Variant, team As Object, c As Long, lastCol As Long, overviewLast As Long
    Set ws = RequiredSheet("Dashboard")
    regionFilter = TextOf(ws.Range("B5").Value2): teamFilter = TextOf(ws.Range("E5").Value2)
    If regionFilter <> "All" And regionFilter <> "NY" And regionFilter <> "Toronto" And regionFilter <> "Unassigned" Then regionFilter = "All"
    If Not IsValidTeamFilter(teamFilter) Then teamFilter = "All"
    ResetOutput ws
    Do While ws.ChartObjects.Count > 0: ws.ChartObjects(1).Delete: Loop
    lastCol = 9
    SetTitle ws, "Operations Workload Dashboard", lastCol
    NoteBand ws, 4, lastCol, "Use the Region and Team filters. OMS shows tracked issues only. Age buckets exclude blank ages; Total includes all reportable items. Grey means non-BAU."
    PutRow ws, 5, 1, Array("Region", regionFilter, "", "Team", teamFilter, "", "Reporting date", mReportDate)
    ws.Range("H5").NumberFormat = "mm/dd/yyyy"
    ws.Range("B5").Interior.Color = COLOR_INPUT: ws.Range("E5").Interior.Color = COLOR_INPUT
    UpdateSelectorLists
    ws.Range("B5").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OpsRegionList"
    ws.Range("E5").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=OpsTeamList"
    SectionBand ws, 7, 6, "Workload by report"
    PutRow ws, 8, 1, Array("Report", "Age <= 30 days", "Age 31-60 days", "Age 61-90 days", "Age > 90 days", "Total")
    HeaderStyle ws.Range("A8:F8")
    r = 9
    For Each report In ReportNames()
        ws.Cells(r, 1).Value2 = IIf(report = "OMS", "OMS - tracked OMS items", ReportDisplayName(CStr(report)))
        For c = 2 To 5
            ws.Cells(r, c).Formula = "=" & SelectedDashboardCount(CStr(report), c - 1)
        Next c
        ws.Cells(r, 6).Formula = "=" & SelectedDashboardCount(CStr(report))
        ws.Cells(r, 1).Font.Bold = True
        r = r + 1
    Next report
    ws.Range("B9:F14").NumberFormat = "#,##0"
    SectionBand ws, 17, 9, "Team workload"
    PutRow ws, 18, 1, Array("Region", "Team", "Manager", "OMS tracked issues", _
        "Chase USD", "Chase CAD", "Internal", "Internal HSBC", "eTR")
    HeaderStyle ws.Range("A18:I18")
    ws.Range("A18:I18").RowHeight = 32
    r = 19
    For Each id In mTeamOrder
        Set team = mTeams(id)
        PutRow ws, r, 1, Array(team("region"), team("name"), team("manager"))
        ws.Cells(r, 4).Formula = TrackedOmsFormula(CStr(id), Nothing)
        ws.Cells(r, 5).Formula = "=" & TeamReportCount(CStr(id), "Chase", Nothing)
        ws.Cells(r, 6).Formula = "=" & TeamReportCount(CStr(id), "Chase CAD", Nothing)
        ws.Cells(r, 7).Formula = "=" & TeamReportCount(CStr(id), "Internal", Nothing)
        ws.Cells(r, 8).Formula = "=" & TeamReportCount(CStr(id), "Internal HSBC", Nothing)
        ws.Cells(r, 9).Formula = "=" & TeamReportCount(CStr(id), "eTR", Nothing)
        For c = 4 To 9
            report = Choose(c - 3, "OMS", "Chase", "Chase CAD", "Internal", "Internal HSBC", "eTR")
            If Not BauFlag(CStr(id), CStr(report)) Then GreyRange ws.Cells(r, c)
        Next c
        r = r + 1
    Next id
    overviewLast = r - 1
    ws.Columns(1).ColumnWidth = 38
    ws.Columns("B:C").ColumnWidth = 22: ws.Columns("D:I").ColumnWidth = 18
    ws.Range(ws.Cells(19, 1), ws.Cells(overviewLast, 9)).Borders(xlInsideHorizontal).Color = COLOR_LINE
    AddActionButton ws, "OpsRefreshButton", "Q5", "Refresh all teams", "RefreshOperations"
    FinishSheet ws, overviewLast, lastCol
    FreezeAt ws, 5, 1
End Sub

Private Sub FinishSheet(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Font.Name = "Arial"
    ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, lastCol)).Font.Size = 10
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).VerticalAlignment = xlCenter
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False: ws.PageSetup.FitToPagesWide = 1: ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Address
End Sub

Private Function ReviewCount(ByVal teamID As String) As Long
    Dim p As Object
    For Each p In mParticipants(teamID)
        If p("review") Then ReviewCount = ReviewCount + 1
    Next p
End Function

Private Sub WriteTeamSummary()
    Dim ws As Worksheet, selections As Object, radarSelections As Object
    Dim id As Variant, team As Object, lo As ListObject
    Dim row As Long, selectRow As Long, selectLast As Long, lastCol As Long, visualLastCol As Long, c As Long
    Dim sectionStart As Long, sectionEnd As Long, metaRow As Long, helperRow As Long, chartName As String
    Dim formatRow As Long, managerRow As Long, participantRows As Double
    Dim showValue As String, radarValue As String, showTeam As Boolean, showRadar As Boolean, anyRadar As Boolean
    Set ws = RequiredSheet("Team Summary")
    Set selections = ExistingTeamSelections(ws)
    Set radarSelections = ExistingTeamRadarSelections(ws)
    lastCol = 1 + mAmountMetrics.Count + mIssues.Count + 1
    If lastCol >= TEAM_META_COLUMN Then Fail "Too many Team Summary columns for the reserved chart-data area."
    For Each id In mTeamOrder: participantRows = participantRows + mParticipants(id).Count: Next id
    If participantRows + 40 * mTeams.Count + 100 >= ws.Rows.Count Then Fail "Too many rows for Team Summary."
    ws.UsedRange.Rows.Hidden = False
    ResetOutput ws
    Do While ws.ChartObjects.Count > 0: ws.ChartObjects(1).Delete: Loop
    For Each id In mTeamOrder
        showTeam = True: showRadar = True
        If selections.Exists(CStr(id)) Then showTeam = CBool(selections(CStr(id)))
        If radarSelections.Exists(CStr(id)) Then showRadar = CBool(radarSelections(CStr(id)))
        If showTeam And showRadar Then
            If mTeams(id)("region") = "NY" Or mTeams(id)("region") = "Toronto" Then anyRadar = True: Exit For
        End If
    Next id
    visualLastCol = lastCol
    If anyRadar Then visualLastCol = Application.Max(lastCol, RADAR_LAST_COLUMN)
    ' Fix widths before creating charts so their geometry matches the final sheet.
    ws.Columns(1).ColumnWidth = 22
    For c = 2 To visualLastCol: ws.Columns(c).ColumnWidth = 10.5: Next c
    SetTitle ws, "Team Summary", visualLastCol
    NoteBand ws, 4, visualLastCol, _
        "Set Show and Radar to Yes or No, then Apply team selection. Selected radars and numeric details appear below the team totals. Radar = No keeps a compact table-only view. Grey means non-BAU."
    SectionBand ws, 6, 6, "Team selection"
    PutRow ws, 7, 1, Array("Show", "Radar", "Region", "Team ID", "Team", "Manager")
    HeaderStyle ws.Range("A7:F7")
    selectRow = 8
    For Each id In mTeamOrder
        Set team = mTeams(id)
        showValue = "Yes": radarValue = "Yes"
        If selections.Exists(CStr(id)) Then If Not CBool(selections(CStr(id))) Then showValue = "No"
        If radarSelections.Exists(CStr(id)) Then If Not CBool(radarSelections(CStr(id))) Then radarValue = "No"
        PutRow ws, selectRow, 1, Array(showValue, radarValue, team("region"), id, team("name"), team("manager"))
        selectRow = selectRow + 1
    Next id
    selectLast = selectRow - 1
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A7:F" & selectLast), , xlYes)
    lo.Name = "TeamSelectionTable": lo.TableStyle = "TableStyleMedium2"
    HeaderStyle lo.HeaderRowRange
    With ws.Range("A8:B" & selectLast)
        .Interior.Color = COLOR_INPUT: .Font.Color = RGB(0, 70, 160): .HorizontalAlignment = xlCenter
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="Yes" & Application.International(xlListSeparator) & "No"
        End With
    End With
    ws.Range("C8:F" & selectLast).WrapText = True
    ws.Rows("8:" & selectLast).AutoFit
    PutRow ws, 1, TEAM_META_COLUMN, Array("Team ID", "Start row", "End row", "Chart name")
    PutRow ws, 1, RADAR_DATA_COLUMN, Array("Metric", "Team raw", "Top 20% raw", "Team index", "Benchmark index")
    metaRow = 2: helperRow = 2: row = selectLast + 5
    For Each id In mTeamOrder
        sectionStart = row
        WriteDashboardTeam ws, CStr(id), row, lastCol
        showTeam = True: showRadar = True: chartName = vbNullString
        If selections.Exists(CStr(id)) Then showTeam = CBool(selections(CStr(id)))
        If radarSelections.Exists(CStr(id)) Then showRadar = CBool(radarSelections(CStr(id)))
        sectionEnd = row - 1
        If showTeam And showRadar Then _
            chartName = BuildTeamRadar(ws, CStr(id), row + 1, 2, helperRow, metaRow - 1, sectionEnd)
        PutRow ws, metaRow, TEAM_META_COLUMN, Array(id, sectionStart, sectionEnd, chartName)
        metaRow = metaRow + 1
        If Len(chartName) > 0 Then
            row = sectionEnd + 2
        Else
            row = sectionEnd + 1
        End If
    Next id
    ws.Range(ws.Cells(1, 1), ws.Cells(row, visualLastCol)).Font.Name = "Arial"
    ws.Range(ws.Cells(1, 1), ws.Cells(row, visualLastCol)).VerticalAlignment = xlCenter
    ws.Columns("AZ:BJ").Hidden = True
    Set selections = ExistingTeamSelections(ws)
    ApplyTeamSummaryVisibility ws, selections
    FinishSheet ws, row, visualLastCol
    ws.Range("A7:F" & selectLast).Font.Size = 9
    ' Keep employee tables compact without shrinking the new radar-detail tables.
    For formatRow = 2 To metaRow - 1
        id = CStr(ws.Cells(formatRow, TEAM_META_COLUMN).Value2)
        sectionStart = CLng(ws.Cells(formatRow, TEAM_META_COLUMN + 1).Value2)
        managerRow = sectionStart + 2 + mParticipants(id).Count
        ws.Range(ws.Cells(sectionStart, 1), ws.Cells(managerRow, lastCol)).Font.Size = 9
    Next formatRow
    ReapplyTeamSummaryBau ws
    ' Column widths can increase wrapped selection-row heights. Position the
    ' button only after the table has reached its final rendered geometry.
    ws.Rows("8:" & selectLast).AutoFit
    AddActionButton ws, "TeamSelectionButton", "A" & CStr(selectLast + 2), _
        "Apply team selection", "ApplyTeamSummarySelection"
    FreezeAt ws, 4, 1
End Sub

Private Sub WriteEmployeeSummary()
    Dim ws As Worksheet, id As Variant, p As Object, list As Collection, row As Long, c As Long, report As Variant
    Set ws = RequiredSheet("Employee Summary"): ResetOutput ws
    SetTitle ws, "Employee Summary", 14
    PutRow ws, 4, 1, Array("Status", "Member ID", "Region", "Team", "Employee", "Membership Note", _
        "Summary total (tracked OMS)", "OMS tracked issues", "Chase USD at threshold", "Chase CAD at threshold", _
        "Internal at threshold", "HSBC Internal at threshold", "eTR Outstanding", "Amount-threshold items aged 30+")
    HeaderStyle ws.Range("A4:N4"): ws.Rows(4).RowHeight = 48
    row = 5
    For Each id In mTeamOrder
        Set list = mParticipants(id)
        For Each p In list
            PutRow ws, row, 1, Array(IIf(p("review"), "Review", "Configured"), p("member"), mTeams(id)("region"), _
                mTeams(id)("name"), p("employee"), p("note"))
            ws.Cells(row, 8).Formula = TrackedOmsFormula(CStr(id), p)
            c = 9
            For Each report In Array("Chase", "Chase CAD", "Internal", "Internal HSBC", "eTR")
                ws.Cells(row, c).Formula = "=" & TeamReportCount(CStr(id), CStr(report), p)
                If Not BauFlag(CStr(id), CStr(report)) Then GreyRange ws.Cells(row, c)
                c = c + 1
            Next report
            ws.Cells(row, 7).Formula = "=SUM(H" & row & ":M" & row & ")"
            ws.Cells(row, 14).Formula = ParticipantAgedAmountFormula(CStr(id), p)
            If Not BauFlag(CStr(id), "OMS") Then GreyRange ws.Cells(row, 8)
            If p("review") Then ws.Range(ws.Cells(row, 1), ws.Cells(row, 6)).Interior.Color = COLOR_AMBER
            row = row + 1
        Next p
    Next id
    ColumnWidths ws, "A:A", 15, "B:C", 13, "D:E", 22, "F:F", 55, "G:N", 22
    If row > 5 Then
        ws.Range("F5:F" & row - 1).WrapText = True
        ws.Range("A5:F" & row - 1).Rows.AutoFit
        ws.Range("G5:N" & row - 1).NumberFormat = "#,##0"
    End If
    FinishSheet ws, row, 14
    FreezeAt ws, 4, 5
End Sub

Private Function ActiveLayoutCount(ByVal report As String) As Long
    Dim source As Object
    For Each source In mSources
        If source("Report") = report Then ActiveLayoutCount = ActiveLayoutCount + 1
    Next source
End Function

Private Sub PutCheck(ByVal ws As Worksheet, ByVal row As Long, ByVal label As String, ByVal formula As String, _
                     ByVal action As String, Optional ByVal informationOnly As Boolean = False)
    ws.Cells(row, 1).Value2 = label: ws.Cells(row, 2).Formula = formula
    If informationOnly Then
        ws.Cells(row, 3).Value2 = "Information"
    Else
        ws.Cells(row, 3).Formula = "=IF(B" & row & "=0,""OK"",""Review"")"
    End If
    ws.Cells(row, 4).Value2 = action
End Sub

Private Sub WriteValidation()
    Dim ws As Worksheet, row As Long, report As Variant, stat As Object, reviewRows As Long, id As Variant
    Dim duplicateCandidates As Long, issueInvalid As Long, rec As Object, missingReports As Long
    Set ws = RequiredSheet("Validation"): ResetOutput ws
    SetTitle ws, "Validation", 8
    PutRow ws, 4, 1, Array("Report", "Active layouts", "Loaded files", "Raw rows", "Normalized rows", _
        "Included at threshold", "Excluded below threshold", "Status")
    HeaderStyle ws.Range("A4:H4")
    row = 5
    For Each report In ReportNames()
        ws.Cells(row, 1).Value2 = ReportDisplayName(CStr(report))
        ws.Cells(row, 2).Value2 = ActiveLayoutCount(CStr(report))
        If mStats.Exists(CStr(report)) Then
            Set stat = mStats(CStr(report))
            ws.Cells(row, 3).Value2 = stat("files").Count
            ws.Cells(row, 4).Value2 = stat("raw")
            duplicateCandidates = duplicateCandidates + CLng(stat("duplicates"))
        Else
            ws.Cells(row, 3).Value2 = 0: ws.Cells(row, 4).Value2 = 0
            missingReports = missingReports + 1
        End If
        ws.Cells(row, 5).Formula = "=" & CountExpr(PairList("B", CStr(report)))
        If IsAmountReport(CStr(report)) Then
            ws.Cells(row, 6).Formula = "=" & CountExpr(PairList("B", CStr(report), "V", "Yes"))
            ws.Cells(row, 7).Formula = "=" & CountExpr(PairList("B", CStr(report), "V", "No"))
        Else
            ws.Cells(row, 6).Value2 = "n.a.": ws.Cells(row, 7).Value2 = "n.a."
        End If
        ws.Cells(row, 8).Formula = "=IF(B" & row & "=0,""Not loaded"",IF(AND(C" & row & ">0,D" & row & "=E" & row & "),""OK"",""Review""))"
        row = row + 1
    Next report
    For Each id In mTeamOrder: reviewRows = reviewRows + ReviewCount(CStr(id)): Next id
    For Each rec In mRecords
        If rec("report") = "OMS" And IsEmpty(rec("issue")) Then issueInvalid = issueInvalid + 1
    Next rec
    SectionBand ws, 13, 4, "Configuration and data checks"
    PutRow ws, 14, 1, Array("Check", "Value", "Status", "Action")
    HeaderStyle ws.Range("A14:D14")
    PutCheck ws, 15, "Unmapped employees", "=" & CountExpr(PairList("U", "Unmapped employee")), "Review aliases and effective Team Members."
    PutCheck ws, 16, "Team membership mismatch records", "=" & CountExpr(PairList("U", "Team membership mismatch")), "Review the employee under the observed manager."
    PutCheck ws, 17, "Review employee/team rows", "=" & CStr(reviewRows), "Review highlighted rows in all summaries."
    PutCheck ws, 18, "Non-BAU entries", "=" & CountExpr(PairList("AB", "Non-BAU")), "Grey cells retain counts.", True
    PutCheck ws, 19, "Possible overlapping eTR rows", "=" & CStr(duplicateCandidates), "Rows retained. Remove overlapping exports and refresh."
    PutCheck ws, 20, "Source reconciliation issues", "=COUNTIFS(H5:H10,""Review"")", "Compare source files, headers, and row counts."
    PutCheck ws, 21, "Unassigned-team records", "=" & CountExpr(PairList("D", UNASSIGNED_TEAM)), "Review the unassigned group and source/file/employee mappings."
    PutCheck ws, 22, "Unrecognized OMS issue codes", "=" & CStr(issueInvalid), "Positions 6-7 must contain two digits."
    PutCheck ws, 23, "Reports without active layouts", "=" & CStr(missingReports), "Activate a source layout if that report is needed.", True
    AddContainsFormat ws.Range("C15:C23"), "Review", RGB(244, 204, 204), RGB(156, 0, 6)
    AddContainsFormat ws.Range("H5:H10"), "Review", RGB(244, 204, 204), RGB(156, 0, 6)
    ColumnWidths ws, "A:A", 38, "B:B", 18, "C:C", 17, "D:D", 58, "E:H", 21
    ws.Range("D15:D23").WrapText = True: ws.Rows("15:23").AutoFit
    FinishSheet ws, 23, 8
End Sub

Private Sub UpdateReadMe()
    Dim ws As Worksheet
    Set ws = RequiredSheet("ReadMe")
    SeedReadMeSheet ws
    ws.Range("B15").Value2 = "VBA refresh " & MODULE_VERSION & ". Last successful refresh: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Range("B16").Value2 = "Amount-threshold formulas update immediately; radar benchmarks and other configuration changes take effect after Refresh all teams."
    ws.Range("B17").Value2 = "Inputs opened read-only, macros disabled for programmatic opens, external links not updated. Recovery copy created before results were changed."
    ws.Range("B24").Value2 = "OMS numbers are extracted from positions 6-7. Add or activate OmsIssueDefinitionsTable rows for more tracked-issue columns."
    ws.Range("B26").Value2 = "White radars use pale report sectors, solid orange actual lines and dashed blue benchmark lines. Tables remain neutral. Non-BAU grey never suppresses employee-table counts."
    FitReadMeRows ws
    ApplyWorkbookPresentation
End Sub

Private Sub WriteRefreshLog(ByVal outcome As String, ByVal backupPath As String)
    Dim ws As Worksheet, row As Long, value As Variant, lastRow As Long
    Set ws = SheetOrNothing(ThisWorkbook, "Refresh Log")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=RequiredSheet("Validation"))
        ws.Name = "Refresh Log"
    End If
    If ws.ProtectContents Then Exit Sub
    ws.Tab.ColorIndex = xlColorIndexNone
    ws.Range("A1").Value2 = "Refresh audit log - " & outcome
    ws.Range("A1").Font.Name = "Arial"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    PutRow ws, 3, 1, Array("Run ID", "Time", "Level", "Action", "Source file", "Source sheet", "Source row", "Details")
    HeaderStyle ws.Range("A3:H3")
    row = Application.Max(4, LastUsedRow(ws) + 1)
    If row + mLogs.Count >= ws.Rows.Count Then Fail "Refresh Log is full. Archive the log using your approved process."
    For Each value In mLogs
        PutRow ws, row, 1, value: row = row + 1
    Next value
    PutRow ws, row, 1, Array(mRunID, Now, outcome, "Run outcome", "", "", 0, _
        IIf(Len(backupPath) > 0, "Recovery copy: " & backupPath, "Outputs not committed; see preceding errors."))
    lastRow = row
    ws.Range("B4:B" & lastRow).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ColumnWidths ws, "A:B", 23, "C:C", 13, "D:D", 28, "E:E", 45, "F:F", 22, "G:G", 12, "H:H", 85
    ws.Range("A4:H" & lastRow).Font.Name = "Arial": ws.Range("A4:H" & lastRow).Font.Size = 10
    ws.Range("H4:H" & lastRow).WrapText = True: ws.Rows("4:" & lastRow).AutoFit
End Sub

' Optional temporary checkpoint for this stage.

Public Sub TypingCheckEmptyOutputs()
    Dim backupPath As String, sheet As Worksheet
    On Error GoTo Failed
    RequireMacroWorkbook
    InitializeRun
    LoadConfiguration
    BuildParticipants
    backupPath = CreateRecoveryCopy()
    WriteNormalizedData
    WriteDashboard
    WriteTeamSummary
    WriteEmployeeSummary
    WriteValidation
    ApplyWorkbookPresentation
    With ThisWorkbook.Worksheets("ReadMe")
        .Range("B15").Value2 = "Typing checkpoint: zero-entry layout preview; no reports were imported."
        FitReadMeRows ThisWorkbook.Worksheets("ReadMe")
    End With
    For Each sheet In ThisWorkbook.Worksheets
        sheet.Calculate
    Next sheet
    MsgBox "Zero-entry preview complete. These are practice results, not imported workloads." & vbCrLf & _
           "Recovery copy: " & backupPath, vbInformation, "Layout checkpoint"
    Exit Sub
Failed:
    MsgBox Err.Description & vbCrLf & "Recovery copy: " & backupPath, vbExclamation, "Layout checkpoint"
End Sub

' ========================================================================
' STAGE 8: FULL REFRESH AND FINAL ACCEPTANCE
' ========================================================================
' Run RefreshOperations against dummy reports. Use VerifyDummyReportRefresh only with the supplied 
' full dummy configuration, reporting date 10/02/2026 and threshold 1000.
' Procedure typing order:
' 01. CheckOutputPermissions
' 02. RefreshOperations
' 03. VerifyDummyReportRefresh

Private Sub CheckOutputPermissions()
    Dim name As Variant, ws As Worksheet, totalRows As Double
    If ThisWorkbook.ProtectStructure Then Fail "Unprotect workbook structure before refresh."
    For Each name In Array("Dashboard", "Team Summary", "Employee Summary", "Normalized Data", "Validation", "ReadMe", "Configuration")
        Set ws = RequiredSheet(CStr(name))
        If ws.ProtectContents Then Fail "Unprotect the output/configuration sheet before refresh: " & name
    Next name
    Set ws = SheetOrNothing(ThisWorkbook, "Refresh Log")
    If Not ws Is Nothing Then If ws.ProtectContents Then Fail "Unprotect Refresh Log before refresh."
    If 2 + mAmountMetrics.Count + mIssues.Count >= TEAM_META_COLUMN Then Fail "Too many Team Summary columns for the reserved chart-data area."
    ' Table rows are small in normal use; check the expanded summaries before clearing them.
    totalRows = mMembers.Count + 40 * mTeams.Count + 100
    If totalRows >= RequiredSheet("Team Summary").Rows.Count Then Fail "Too many configured members for Team Summary."
End Sub

Public Sub RefreshOperations()
    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean
    Dim oldAlerts As Boolean, oldSecurity As Long, oldStatus As Variant
    Dim stateSaved As Boolean, commitStarted As Boolean, backupPath As String
    Dim failureText As String, savedErr As Long, calcSheet As Worksheet
    Dim participantRows As Double, teamID As Variant
    If mBusy Then
        MsgBox "A refresh is already running.", vbInformation
        Exit Sub
    End If
    mBusy = True
    Set mLogs = Nothing
    On Error GoTo Failed
    RequireMacroWorkbook
    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldSecurity = Application.AutomationSecurity
    oldStatus = Application.StatusBar
    stateSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = SECURITY_FORCE_DISABLE
    InitializeRun
    TestBusinessRules
    LoadConfiguration
    CheckOutputPermissions
    Application.StatusBar = "Operations dashboard: validating current reports..."
    ImportAllReports
    BuildParticipants
    If mRecords.Count >= RequiredSheet("Normalized Data").Rows.Count Then Fail "Too many imported rows for one Excel sheet."
    For Each teamID In mTeamOrder: participantRows = participantRows + mParticipants(teamID).Count: Next teamID
    If participantRows + 40 * mTeams.Count + 100 >= RequiredSheet("Team Summary").Rows.Count Then Fail "Too many employee/team rows for Team Summary."
    ' No existing output is changed until all files and configuration pass.
    backupPath = CreateRecoveryCopy()
    commitStarted = True
    Application.StatusBar = "Operations dashboard: rebuilding summaries..."
    WriteNormalizedData
    WriteDashboard
    WriteTeamSummary
    WriteEmployeeSummary
    WriteValidation
    UpdateReadMe
    LogInfo "Refresh complete", "", "", 0, CStr(mRecords.Count) & " source rows imported."
    WriteRefreshLog "SUCCESS", backupPath
    ' Calculate the threshold column before its dependent summaries.
    RequiredSheet("Normalized Data").Calculate
    For Each calcSheet In ThisWorkbook.Worksheets: calcSheet.Calculate: Next calcSheet
    RequiredSheet("Dashboard").Activate
    GoTo CleanUp
Failed:
    savedErr = Err.Number
    failureText = Err.Description
    On Error Resume Next
    If Not mSourceBook Is Nothing Then mSourceBook.Close SaveChanges:=False
    Set mSourceBook = Nothing
    If Not mLogs Is Nothing Then
        LogEntry "ERROR", "Refresh stopped", "", "", 0, failureText
        WriteRefreshLog "FAILED", backupPath
    End If
    On Error GoTo 0
CleanUp:
    On Error Resume Next
    If stateSaved Then
        Application.AutomationSecurity = oldSecurity
        Application.Calculation = oldCalc
        Application.DisplayAlerts = oldAlerts
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
        Application.StatusBar = oldStatus
    End If
    mBusy = False
    On Error GoTo 0
    If savedErr <> 0 Then
        If commitStarted Then
            MsgBox "Refresh stopped while writing results. Do not save over your last good copy." & vbCrLf & _
                   failureText & vbCrLf & "Restore the recovery copy: " & backupPath, vbCritical, "Refresh failed"
        Else
            MsgBox "Refresh stopped. Existing dashboard results were not replaced." & vbCrLf & _
                   failureText & vbCrLf & "See Refresh Log for details.", vbExclamation, "Refresh failed"
        End If
    Else
        MsgBox "All teams refreshed: " & Format$(mRecords.Count, "#,##0") & " imported rows." & vbCrLf & _
               CStr(mWarnings) & " review message(s). Check Validation and Refresh Log, then save the workbook." & _
               vbCrLf & "Recovery copy: " & backupPath, vbInformation, "Refresh complete"
    End If
End Sub

Public Sub VerifyDummyReportRefresh()
    Dim ws As Worksheet, values As Variant, r As Long, included As Long
    Dim mismatch As Long, reportCounts As Object, key As Variant
    On Error GoTo Failed
    Set ws = RequiredSheet("Normalized Data")
    ws.Calculate
    values = ws.Range("A1:AB" & LastUsedRow(ws)).Value2
    Set reportCounts = NewMap()
    AssertTest UBound(values, 1) - 1 = 63, "Dummy import must have 63 normalized rows."
    AssertTest CDbl(SettingCell("Amount threshold").Value2) = 1000, "Set the dummy threshold to 1,000 before this check."
    For r = 2 To UBound(values, 1)
        key = TextOf(values(r, 2))
        If Not reportCounts.Exists(key) Then reportCounts.Add key, 0
        reportCounts(key) = CLng(reportCounts(key)) + 1
        If Not IsAmountReport(CStr(key)) Or TextOf(values(r, 22)) = "Yes" Then included = included + 1
        If TextOf(values(r, 21)) = "Team membership mismatch" Then mismatch = mismatch + 1
    Next r
    AssertTest included = 62, "Dummy reportable total must be 62."
    AssertTest mismatch = 4, "Dummy report must show four membership mismatches."
    AssertTest reportCounts("OMS") = 6 And reportCounts("eTR") = 8, "Dummy OMS/eTR row counts."
    AssertTest reportCounts("Chase") = 8 And reportCounts("Chase CAD") = 5, "Dummy Chase row counts."
    AssertTest reportCounts("Internal") = 24 And reportCounts("Internal HSBC") = 12, "Dummy internal row counts."
    MsgBox "Dummy refresh checks passed: 63 rows, 62 reportable, four membership exceptions." & vbCrLf & _
           "The supplied OMS sample contains issue 07, so 10/12/19/20/21/22 correctly show zero. Activate 07 to test extension.", vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbCritical, "Dummy refresh check"
End Sub
