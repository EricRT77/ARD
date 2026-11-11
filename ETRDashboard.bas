Attribute VB_Name = "ETRDashboard"
Option Explicit

' eTR dashboard v2.1. Desktop Excel for Windows, 32- or 64-bit.
' Start with a blank workbook, save as .xlsm, import this module,
' then run ETR_InstallDashboard. No template, extra references or add-ins.
' Creates Dashboard, Configuration and Original Data automatically.
' Source files are read-only. Original Data retains the last successful
' selected-sheet snapshot (all rows/columns, values and number formats).

Private Const ERR_BASE As Long = vbObjectError + 741
Private Const CHUNK_SIZE As Long = 50000
Private Const MAX_AGE As Long = 2147483646
Private Const MAX_BUCKETS As Long = 12
Private Const DEFAULT_SNAPSHOT_CELLS As Long = 10000000
Private Const SNAPSHOT_CHUNK_CELLS As Long = 250000
Private mBusy As Boolean
Private mLastError As String

Private Type AgeBucket
    Label As String
    MinAge As Long
    MaxAge As Long
    OpenEnded As Boolean
    SortOrder As Long
End Type

Private Type Settings
    HeaderRow As Long
    FirstDataRow As Long
    DescriptionHeader As String
    TransitHeader As String
    TRidHeader As String
    ApprovalDateHeader As String
    DescriptionFilter As String
    FilterMode As String
    SourceSheet As String
    UnmappedPolicy As String
    TransitKeyMode As String
    MaxDataRows As Long
    MaxSnapshotCells As Long
    BucketCount As Long
    Buckets() As AgeBucket
End Type

Private Type RunMetrics
    ReadRows As Long
    BlankRows As Long
    ExcludedRows As Long
    GrossRows As Long
    UniquePairs As Long
    GlobalIds As Long
    UnmappedRows As Long
    UnmappedTransits As Long
    MaximumCell As Long
    SourcePath As String
    SourceSheet As String
    SnapshotLastRow As Long
    SnapshotLastColumn As Long
    AsOfDate As Date
    CompletedAt As Date
End Type

Public Sub RefreshETRDashboard()
    Dim chooser As Object, chosen As String
    If mBusy Then Exit Sub
    On Error GoTo Failed
    Set chooser = Application.FileDialog(3) ' msoFileDialogFilePicker
    With chooser
        .Title = "Select an eTR source workbook"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel workbooks", "*.xlsx;*.xlsm;*.xlsb"
        If .Show <> -1 Then Exit Sub ' Cancellation does not change the workbook.
        chosen = .SelectedItems(1)
    End With
    Call ETR_RefreshFile(chosen, True)
    Exit Sub
Failed:
    MsgBox "Unable to select a source file: " & Err.Description, vbExclamation, "eTR dashboard"
End Sub

Public Sub ETR_ShowConfiguration()
    ThisWorkbook.Worksheets("Configuration").Activate
End Sub

Public Sub ETR_ShowOriginalData()
    ThisWorkbook.Worksheets("Original Data").Activate
End Sub

Public Sub ETR_InstallDashboard()
    Call ETR_InitializeWorkbook(True)
End Sub

' Idempotent setup. Also callable without dialogs by acceptance tests.
Public Function ETR_InitializeWorkbook(Optional ByVal showMessages As Boolean = True) As Boolean
    Dim cfgSheet As Worksheet, dashSheet As Worksheet, dataSheet As Worksheet, reusable As Worksheet
    Dim cfg As Settings, metrics As RunMetrics, regions As Object, orders As Object
    Dim gross As Object, distinct As Object, oldScreen As Boolean, oldEvents As Boolean
    Dim oldAlerts As Boolean, stateSaved As Boolean, errorText As String, needsDashboard As Boolean, initializedData As Boolean
    If mBusy Then
        mLastError = "An eTR operation is already in progress.": Exit Function
    End If
    mBusy = True: mLastError = vbNullString
    On Error GoTo Failed
    If ThisWorkbook.ProtectStructure Then Fail "Unprotect the workbook structure before setup."
    If ThisWorkbook.ReadOnly Then Fail "Save an editable copy before setup."
    Set cfgSheet = WorksheetOrNothing("Configuration")
    Set dashSheet = WorksheetOrNothing("Dashboard")
    Set dataSheet = WorksheetOrNothing("Original Data")
    If Not dashSheet Is Nothing Then
        If dashSheet.ProtectContents Or dashSheet.ProtectDrawingObjects Then Fail "Unprotect the Dashboard sheet before setup."
        If Not SheetIsBlank(dashSheet) Then
            If CleanText(dashSheet.Range("B2").Value2, "Dashboard title") <> "eTR heatmap dashboard" Then _
                Fail "Dashboard already contains other content. Rename that sheet before setup."
        End If
    End If
    If Not cfgSheet Is Nothing Then
        If cfgSheet.ProtectContents Then Fail "Unprotect the Configuration sheet before setup."
    End If
    If Not dataSheet Is Nothing Then
        If Not SheetIsBlank(dataSheet) Then
            If SheetRole(dataSheet) <> "SourceSnapshot" Then _
                Fail "Original Data already contains other content. Rename that sheet before setup."
        End If
    End If
    If ThisWorkbook.Worksheets.Count = 1 And dashSheet Is Nothing Then
        If ThisWorkbook.Worksheets(1).Name <> "Configuration" And ThisWorkbook.Worksheets(1).Name <> "Original Data" Then
            If SheetIsBlank(ThisWorkbook.Worksheets(1)) Then Set reusable = ThisWorkbook.Worksheets(1)
        End If
    End If
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts: stateSaved = True
    Application.ScreenUpdating = False: Application.EnableEvents = False
    If cfgSheet Is Nothing Then
        Set cfgSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        cfgSheet.Name = "Configuration"
    End If
    If SheetIsBlank(cfgSheet) Then BuildConfiguration cfgSheet
    UpgradeApprovalSettings cfgSheet
    LoadConfiguration cfg, regions, orders
    If dashSheet Is Nothing Then
        If reusable Is Nothing Then
            Set dashSheet = ThisWorkbook.Worksheets.Add(Before:=cfgSheet)
        Else
            Set dashSheet = reusable
        End If
        dashSheet.Name = "Dashboard"
    End If
    needsDashboard = SheetIsBlank(dashSheet)
    If dataSheet Is Nothing Then
        Set dataSheet = ThisWorkbook.Worksheets.Add(After:=cfgSheet)
        dataSheet.Name = "Original Data"
    End If
    If SheetIsBlank(dataSheet) Then
        BuildEmptyOriginalData dataSheet
        initializedData = True
    End If
    If needsDashboard Then
        Set gross = NewDictionary(vbBinaryCompare): Set distinct = NewDictionary(vbBinaryCompare)
        RenderDashboard dashSheet, cfg, regions, orders, gross, distinct, metrics, False
        dashSheet.Calculate
    Else
        AddControls dashSheet
        If CStr(dashSheet.Range("B12").Value2) = "n.a." Then
            dashSheet.Range("B5").Value2 = "No source loaded. Configure your regions, then use Refresh eTR file."
        ElseIf initializedData Then
            dashSheet.Range("B5").Value2 = "Original Data has no source snapshot yet. Refresh to rebuild both output tabs."
        End If
    End If
    MarkSheetRole dashSheet, "Dashboard"
    dashSheet.Visible = xlSheetVisible: cfgSheet.Visible = xlSheetVisible: dataSheet.Visible = xlSheetVisible
    If dashSheet.Index <> 1 Then dashSheet.Move Before:=ThisWorkbook.Sheets(1)
    cfgSheet.Move After:=dashSheet
    dataSheet.Move After:=cfgSheet
    cfgSheet.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False: ActiveWindow.SplitColumn = 0: ActiveWindow.SplitRow = 5
    ActiveWindow.FreezePanes = True
    dataSheet.Activate: ActiveWindow.DisplayGridlines = False
    dashSheet.Activate: ActiveWindow.DisplayGridlines = False
    ETR_InitializeWorkbook = True
    GoTo Finally
Failed:
    errorText = Err.Description: mLastError = errorText
Finally:
    On Error Resume Next
    If stateSaved Then
        Application.EnableEvents = oldEvents: Application.DisplayAlerts = oldAlerts
        Application.ScreenUpdating = oldScreen
    End If
    mBusy = False
    If showMessages Then
        If ETR_InitializeWorkbook Then
            MsgBox "Setup complete: Dashboard, Configuration and Original Data are ready." & vbCrLf & _
                "Enter your region mappings, select an eTR file, and save this workbook as .xlsm.", vbInformation, "eTR dashboard"
        Else
            MsgBox "Setup did not finish: " & errorText, vbExclamation, "eTR dashboard"
        End If
    End If
    On Error GoTo 0
End Function

Public Function ETR_LastError() As String
    ETR_LastError = mLastError
End Function

' Callable without dialogs for integration/testing. Does not save the dashboard file.
' Both published sheets are replaced together after validation and reconciliation.
Public Function ETR_RefreshFile(ByVal sourcePath As String, Optional ByVal showMessages As Boolean = True) As Boolean
    Dim cfg As Settings, metrics As RunMetrics
    Dim regionByTransit As Object, regionOrder As Object
    Dim gross As Object, distinct As Object
    Dim sourceBook As Workbook, sourceSheet As Worksheet, staged As Worksheet, stagedData As Worksheet, publishedData As Worksheet
    Dim sourceOwned As Boolean, stateSaved As Boolean
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim oldCalc As XlCalculation, oldSecurity As Long, oldStatus As Variant
    Dim headerCols() As Long, message As String
    If mBusy Then
        mLastError = "A refresh is already in progress."
        Exit Function
    End If
    mBusy = True
    mLastError = vbNullString
    On Error GoTo Failed
    metrics.AsOfDate = Date ' Capture once, even if a long refresh crosses midnight.
    ValidateSourcePath sourcePath
    LoadConfiguration cfg, regionByTransit, regionOrder
    If ThisWorkbook.ProtectStructure Then Fail "Unprotect the dashboard workbook structure before refreshing."
    If ThisWorkbook.ReadOnly Then Fail "The dashboard is read-only. Save an editable copy before refreshing."
    ' Check the existing Dashboard before opening any source.
    Set staged = ThisWorkbook.Worksheets("Dashboard")
    Set staged = Nothing
    Set publishedData = WorksheetOrNothing("Original Data")
    If publishedData Is Nothing Then Fail "Run ETR_InstallDashboard once to create the Original Data tab."
    If SheetRole(publishedData) <> "SourceSnapshot" Then Fail "Original Data is not an eTR-managed snapshot. Run setup in a blank workbook."

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation
    oldSecurity = Application.AutomationSecurity
    oldStatus = Application.StatusBar
    stateSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = 3 ' msoAutomationSecurityForceDisable
    Application.StatusBar = "eTR: opening source read-only..."
    Set sourceBook = Application.Workbooks.Open(Filename:=sourcePath, UpdateLinks:=0, _
        ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False, _
        Password:=vbNullString, WriteResPassword:=vbNullString)
    sourceOwned = True
    Set sourceSheet = ResolveSourceSheet(sourceBook, cfg, headerCols)
    metrics.SourcePath = sourceBook.FullName
    metrics.SourceSheet = sourceSheet.Name
    SnapshotBounds sourceSheet, cfg, metrics.SnapshotLastRow, metrics.SnapshotLastColumn
    Set gross = NewDictionary(vbBinaryCompare)
    Set distinct = NewDictionary(vbBinaryCompare)
    AggregateSource sourceSheet, headerCols, cfg, regionByTransit, gross, distinct, metrics
    Set stagedData = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    stagedData.Name = UnusedSheetName("_eTR_data")
    stagedData.Visible = xlSheetVeryHidden
    CopyOriginalData sourceSheet, stagedData, cfg, metrics
    sourceBook.Close SaveChanges:=False
    sourceOwned = False
    Set sourceBook = Nothing
    metrics.CompletedAt = Now
    Application.StatusBar = "eTR: preparing both heatmaps..."
    Set staged = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    staged.Name = UnusedSheetName("_eTR_stage")
    staged.Visible = xlSheetVeryHidden
    RenderDashboard staged, cfg, regionByTransit, regionOrder, gross, distinct, metrics
    staged.Calculate
    ' Independently reconcile displayed totals before replacing the published view.
    If CLng(staged.Range("B12").Value2) <> metrics.GrossRows Then Fail "Gross total reconciliation failed."
    If CLng(staged.Range("D12").Value2) <> metrics.UniquePairs Then Fail "Unique total reconciliation failed."
    If CLng(staged.Range("F12").Value2) <> metrics.GlobalIds Then Fail "Global TRid reconciliation failed."
    If metrics.UniquePairs > metrics.GrossRows Then Fail "Unique counts exceed gross counts."
    CommitOutputs staged, stagedData
    Set staged = Nothing
    Set stagedData = Nothing
    ETR_RefreshFile = True
    GoTo Finally

Failed:
    message = Err.Description
    If Len(message) = 0 Then message = "Unknown refresh error."
    mLastError = message
Finally:
    On Error Resume Next
    If sourceOwned Then sourceBook.Close SaveChanges:=False
    If Not staged Is Nothing Then
        Application.DisplayAlerts = False
        staged.Delete
    End If
    If Not stagedData Is Nothing Then
        Application.DisplayAlerts = False
        stagedData.Delete
    End If
    If stateSaved Then
        Application.CutCopyMode = False
        Application.AutomationSecurity = oldSecurity
        Application.Calculation = oldCalc
        Application.EnableEvents = oldEvents
        Application.DisplayAlerts = oldAlerts
        Application.ScreenUpdating = oldScreen
        Application.StatusBar = oldStatus
    End If
    mBusy = False
    If ETR_RefreshFile Then
        ThisWorkbook.Worksheets("Original Data").Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.FreezePanes = False: ActiveWindow.SplitColumn = 0
        ActiveWindow.SplitRow = cfg.HeaderRow: ActiveWindow.FreezePanes = True
        ThisWorkbook.Worksheets("Dashboard").Activate
        ActiveWindow.DisplayGridlines = False
        If showMessages Then MsgBox "Refresh complete." & vbCrLf & _
            Format$(metrics.GrossRows, "#,##0") & " gross rows; " & _
            Format$(metrics.UniquePairs, "#,##0") & " unique transit + TRid pairs." & vbCrLf & _
            "Save the dashboard to retain this snapshot.", vbInformation, "eTR dashboard"
    ElseIf showMessages Then
        MsgBox "Refresh stopped. The previous Dashboard and Original Data are unchanged." & vbCrLf & vbCrLf & _
            mLastError, vbExclamation, "eTR dashboard"
    End If
    On Error GoTo 0
End Function

Private Sub Fail(ByVal message As String)
    Err.Raise ERR_BASE, "eTR dashboard", message
End Sub

Private Function NewDictionary(ByVal compareMode As VbCompareMethod) As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
    NewDictionary.CompareMode = compareMode
End Function

Private Function CleanText(ByVal value As Variant, ByVal context As String) As String
    If IsError(value) Then Fail context & " contains an Excel error."
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    CleanText = Trim$(Replace(Replace(Replace(CStr(value), ChrW(160), " "), vbCr, " "), vbLf, " "))
    CleanText = Trim$(Replace(CleanText, vbTab, " "))
End Function

Private Function HeaderKey(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    HeaderKey = LCase$(Replace(CleanText(value, "Header"), " ", vbNullString))
End Function

Private Function WholeNumber(ByVal value As Variant, ByVal low As Long, ByVal high As Long, ByVal context As String) As Long
    Dim n As Double, txt As String
    txt = CleanText(value, context)
    If Len(txt) = 0 Or VarType(value) = vbBoolean Or VarType(value) = vbDate Then Fail context & " must be a whole number."
    If Not IsNumeric(value) Then Fail context & " must be a whole number."
    n = CDbl(value)
    If n < CDbl(low) Or n > CDbl(high) Then Fail context & " is outside the permitted range " & CStr(low) & " to " & CStr(high) & "."
    If n <> Fix(n) Then Fail context & " must be a whole number."
    WholeNumber = CLng(n)
End Function

Private Function RequiredSetting(ByVal items As Object, ByVal key As String, Optional ByVal allowBlank As Boolean = False) As Variant
    If Not items.Exists(key) Then Fail "Missing setting: " & key & "."
    If Not allowBlank Then
        If Len(CleanText(items(key), "Setting " & key)) = 0 Then Fail "Setting " & key & " is blank."
    End If
    RequiredSetting = items(key)
End Function

Private Function GetTable(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    On Error Resume Next
    Set GetTable = ws.ListObjects(tableName)
    On Error GoTo 0
    If GetTable Is Nothing Then Fail "Configuration table is missing: " & tableName & ". Run ETR_InstallDashboard in a blank workbook."
End Function

Private Function TableColumn(ByVal table As ListObject, ByVal name As String) As Long
    Dim col As ListColumn
    For Each col In table.ListColumns
        If HeaderKey(col.Name) = HeaderKey(name) Then
            If TableColumn <> 0 Then Fail "Duplicate column " & name & " in " & table.Name & "."
            TableColumn = col.Index
        End If
    Next col
    If TableColumn = 0 Then Fail "Missing column " & name & " in " & table.Name & "."
End Function

Private Function WorksheetOrNothing(ByVal name As String) As Worksheet
    Dim item As Object
    For Each item In ThisWorkbook.Sheets
        If StrComp(item.Name, name, vbTextCompare) = 0 Then
            If TypeOf item Is Worksheet Then
                Set WorksheetOrNothing = item
            Else
                Fail "The reserved name '" & name & "' belongs to a chart sheet. Rename it before setup."
            End If
            Exit Function
        End If
    Next item
End Function

Private Function SheetIsBlank(ByVal ws As Worksheet) As Boolean
    If ws.ListObjects.Count > 0 Or ws.Shapes.Count > 0 Or ws.Hyperlinks.Count > 0 Then Exit Function
    If ws.Comments.Count > 0 Then Exit Function
    SheetIsBlank = (Application.CountA(ws.UsedRange) = 0)
End Function

Private Function SheetRole(ByVal ws As Worksheet) As String
    Dim item As Object
    For Each item In ws.CustomProperties
        If item.Name = "eTRRole" Then SheetRole = CStr(item.Value): Exit Function
    Next item
End Function

Private Sub MarkSheetRole(ByVal ws As Worksheet, ByVal role As String)
    Dim item As Object
    For Each item In ws.CustomProperties
        If item.Name = "eTRRole" Then item.Delete: Exit For
    Next item
    ws.CustomProperties.Add Name:="eTRRole", Value:=role
End Sub

Private Sub AddListValidation(ByVal cell As Range, ByVal choices As Variant)
    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:=Join(choices, Application.International(xlListSeparator))
        .IgnoreBlank = False: .InCellDropdown = True: .ShowError = True
        .ErrorTitle = "Choose a listed option": .ErrorMessage = "Use the dropdown choices."
    End With
End Sub

Private Sub StyleInputTable(ByVal table As ListObject)
    table.TableStyle = "TableStyleMedium2"
    With table.HeaderRowRange
        .Interior.Color = RGB(24, 52, 73): .Font.Color = vbWhite: .Font.Bold = True
        .WrapText = True: .VerticalAlignment = xlCenter: .RowHeight = 34
    End With
    If Not table.DataBodyRange Is Nothing Then
        table.DataBodyRange.Interior.Color = RGB(255, 243, 204)
        table.DataBodyRange.Font.Color = RGB(23, 74, 118)
    End If
End Sub

Private Sub BuildConfiguration(ByVal ws As Worksheet)
    Dim defaults As Variant, hints As Variant, mappingHeader(1 To 1, 1 To 3) As Variant
    Dim bucketData As Variant, grid() As Variant, i As Long, j As Long, table As ListObject
    With ws.Range("B2:I36")
        .Font.Name = "Arial": .Font.Size = 10: .Font.Color = RGB(32, 52, 70)
        .RowHeight = 22: .VerticalAlignment = xlCenter
    End With
    ws.Tab.Color = RGB(105, 133, 160)
    ws.Columns("A").ColumnWidth = 3: ws.Columns("B").ColumnWidth = 28
    ws.Columns("C").ColumnWidth = 27: ws.Columns("D").ColumnWidth = 13
    ws.Columns("E").ColumnWidth = 3: ws.Columns("F").ColumnWidth = 23
    ws.Columns("G:H").ColumnWidth = 13: ws.Columns("I").ColumnWidth = 12
    PutMerged ws, 2, 2, 9, "Configuration"
    ws.Range("B2").Font.Size = 16: ws.Range("B2").Font.Bold = True: ws.Rows(2).RowHeight = 30
    PutMerged ws, 3, 2, 9, "Amber cells are editable. Changes take effect on the next successful refresh."
    defaults = Array(Array("HeaderRow", 2), Array("FirstDataRow", 3), _
        Array("DescriptionHeader", "Description"), Array("TransitHeader", "booking_transit"), _
        Array("TRidHeader", "TRid"), Array("ApprovalDateHeader", "ApprovalDate"), _
        Array("DescriptionFilter", "New"), Array("FilterMode", "Exact"), Array("SourceSheet", ""), _
        Array("UnmappedPolicy", "Include"), Array("TransitKeyMode", "Numeric"), _
        Array("MaxDataRows", 500000), Array("MaxSnapshotCells", DEFAULT_SNAPSHOT_CELLS))
    hints = Array("Source headers are in row 2 by default.", "First source record; must be below HeaderRow.", _
        "Match by header name; column A is the current position only.", _
        "Match by header name; column C is the current position only.", _
        "Match by header name; column H is the current position only.", _
        "Approval date, matched by header name. Age is today minus this date.", _
        "Exact and Contains use this text, ignoring case and outer spaces.", _
        "Exact, Contains or All. All includes every description.", _
        "Blank selects the only worksheet with all four required headers.", _
        "Include groups missing mappings under Unmapped; Error stops refresh.", _
        "Numeric: 00101 = 101. Text: preserve leading zeros; ignore case.", _
        "Maximum source row span beginning at FirstDataRow, including gaps.", _
        "Maximum cells in the complete Original Data snapshot, including blanks inside its bounds.")
    ReDim grid(1 To UBound(defaults) + 2, 1 To 2)
    grid(1, 1) = "Setting": grid(1, 2) = "Value"
    For i = 0 To UBound(defaults)
        grid(i + 2, 1) = defaults(i)(0): grid(i + 2, 2) = defaults(i)(1)
        PutMerged ws, i + 6, 4, 9, hints(i)
        With ws.Range(ws.Cells(i + 6, 4), ws.Cells(i + 6, 9))
            .Font.Color = RGB(90, 110, 127): .WrapText = True: .RowHeight = 30
        End With
    Next i
    ws.Range("B5:C18").Value2 = grid
    Set table = ws.ListObjects.Add(xlSrcRange, ws.Range("B5:C18"), , xlYes)
    table.Name = "tblSettings": StyleInputTable table
    table.ListColumns(1).DataBodyRange.Interior.Color = RGB(247, 249, 252)
    ws.Range("C17:C18").NumberFormat = "#,##0"
    AddListValidation ws.Range("C13"), Array("Exact", "Contains", "All")
    AddListValidation ws.Range("C15"), Array("Include", "Error")
    AddListValidation ws.Range("C16"), Array("Numeric", "Text")
    PutMerged ws, 20, 2, 9, "Age = refresh date minus ApprovalDate (calendar days). Source AgedDays is ignored."
    ws.Range("B20:I20").WrapText = True: ws.Rows(20).RowHeight = 30
    PutMerged ws, 22, 2, 4, "Transit to region mapping"
    PutMerged ws, 22, 6, 9, "Age buckets (inclusive days)"
    ws.Range("B22:I22").Font.Bold = True: ws.Rows(22).RowHeight = 28
    mappingHeader(1, 1) = "booking_transit": mappingHeader(1, 2) = "Region": mappingHeader(1, 3) = "RegionOrder"
    ws.Range("B23:D23").Value2 = mappingHeader
    ws.Range("B24").NumberFormat = "@"
    Set table = ws.ListObjects.Add(xlSrcRange, ws.Range("B23:D24"), , xlYes)
    table.Name = "tblTransitMapping": StyleInputTable table
    bucketData = Array(Array("BucketLabel", "MinAge", "MaxAge", "BucketOrder"), _
        Array("0-2 days", 0, 2, 1), Array("3-5 days", 3, 5, 2), Array("6-10 days", 6, 10, 3), _
        Array("11-20 days", 11, 20, 4), Array("21+ days", 21, Empty, 5))
    ReDim grid(1 To 6, 1 To 4)
    For i = 0 To 5
        For j = 0 To 3: grid(i + 1, j + 1) = bucketData(i)(j): Next j
    Next i
    ws.Range("F23:I28").Value2 = grid
    Set table = ws.ListObjects.Add(xlSrcRange, ws.Range("F23:I28"), , xlYes)
    table.Name = "tblAgeBuckets": StyleInputTable table
    With ws.Range("F30:I32")
        .Merge: .Value2 = "Use 1-12 buckets covering all whole days from 0, without gaps or overlaps. Leave only the final MaxAge blank."
        .WrapText = True: .Font.Color = RGB(90, 110, 127)
    End With
    With ws.Range("F34:I36")
        .Merge: .Value2 = "Add rows inside the mapping table. Each transit needs one region. Use one RegionOrder per region. Unmapped is reserved."
        .WrapText = True: .Font.Color = RGB(90, 110, 127)
    End With
    MarkSheetRole ws, "Configuration"
End Sub

Private Sub UpgradeApprovalSettings(ByVal ws As Worksheet)
    Dim table As ListObject, i As Long, keyCol As Long, valCol As Long, oldRow As Long
    Set table = GetTable(ws, "tblSettings")
    keyCol = TableColumn(table, "Setting"): valCol = TableColumn(table, "Value")
    If table.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To table.ListRows.Count
        If StrComp(CStr(table.DataBodyRange.Cells(i, keyCol).Value2), "ApprovalDateHeader", vbTextCompare) = 0 Then Exit Sub
        If StrComp(CStr(table.DataBodyRange.Cells(i, keyCol).Value2), "AgeHeader", vbTextCompare) = 0 Then oldRow = i
    Next i
    If oldRow = 0 Then Exit Sub
    table.DataBodyRange.Cells(oldRow, keyCol).Value2 = "ApprovalDateHeader"
    table.DataBodyRange.Cells(oldRow, valCol).Value2 = "ApprovalDate"
    If SheetRole(ws) = "Configuration" Then
        ws.Cells(table.DataBodyRange.Row + oldRow - 1, 4).Value2 = "Approval date, matched by header name. Age is today minus this date."
        ws.Range("B20").Value2 = "Age = refresh date minus ApprovalDate (calendar days). Source AgedDays is ignored."
    End If
End Sub

Private Sub BuildEmptyOriginalData(ByVal ws As Worksheet)
    With ws.Range("A1:H5")
        .Font.Name = "Arial": .Font.Size = 10: .Font.Color = RGB(32, 52, 70)
        .ColumnWidth = 16: .RowHeight = 24
    End With
    PutMerged ws, 1, 1, 8, "Original Data"
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True
    PutMerged ws, 3, 1, 8, "No source snapshot has been loaded. Use Refresh eTR file on Dashboard."
    PutMerged ws, 4, 1, 8, "This tab will retain all source rows and columns, including rows excluded by the filter."
    PutMerged ws, 5, 1, 8, "Values and number formats are retained. Refresh replaces this snapshot after validation."
    ws.Range("A3:H5").WrapText = True: ws.Rows("3:5").RowHeight = 30
    ws.Tab.Color = RGB(163, 175, 185)
    MarkSheetRole ws, "SourceSnapshot"
End Sub

Private Sub SnapshotBounds(ByVal ws As Worksheet, ByRef cfg As Settings, ByRef lastRow As Long, ByRef lastCol As Long)
    Dim found As Range
    Set found = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If found Is Nothing Then Fail "The source worksheet is empty."
    lastRow = found.Row
    Set found = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    lastCol = found.Column
    If lastRow - cfg.FirstDataRow + 1 > cfg.MaxDataRows Then _
        Fail "The complete source snapshot exceeds MaxDataRows. Check trailing content or increase the limit."
    If CDbl(lastRow) * CDbl(lastCol) > CDbl(cfg.MaxSnapshotCells) Then _
        Fail "The complete source snapshot exceeds MaxSnapshotCells. Check trailing content or increase the limit."
    If ws.Parent.Date1904 <> ThisWorkbook.Date1904 Then _
        Fail "Source and dashboard use different Excel date systems. Use a source with the same date system to preserve dates."
End Sub

Private Sub CopyOriginalData(ByVal source As Worksheet, ByVal target As Worksheet, ByRef cfg As Settings, ByRef metrics As RunMetrics)
    Dim firstRow As Long, lastRow As Long, chunkRows As Long, src As Range, dst As Range, formulas As Range
    Dim savedNumber As Long, savedMessage As String
    On Error GoTo Failed
    chunkRows = SNAPSHOT_CHUNK_CELLS \ metrics.SnapshotLastColumn
    If chunkRows < 1 Then chunkRows = 1
    If chunkRows > CHUNK_SIZE Then chunkRows = CHUNK_SIZE
    For firstRow = 1 To metrics.SnapshotLastRow Step chunkRows
        lastRow = firstRow + chunkRows - 1
        If lastRow > metrics.SnapshotLastRow Then lastRow = metrics.SnapshotLastRow
        Set src = source.Range(source.Cells(firstRow, 1), source.Cells(lastRow, metrics.SnapshotLastColumn))
        Set dst = target.Range(target.Cells(firstRow, 1), target.Cells(lastRow, metrics.SnapshotLastColumn))
        src.Copy
        dst.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        Application.StatusBar = "eTR: retaining Original Data through row " & Format$(lastRow, "#,##0") & "..."
    Next firstRow
    Set dst = target.Range(target.Cells(1, 1), target.Cells(metrics.SnapshotLastRow, metrics.SnapshotLastColumn))
    ' Paste only values and number formats; never import source code, formulas or connections.
    On Error Resume Next
    Set formulas = dst.SpecialCells(xlCellTypeFormulas)
    On Error GoTo Failed
    If Not formulas Is Nothing Then Fail "Original Data unexpectedly contains formulas. Refresh was not committed."
    If target.Hyperlinks.Count > 0 Then target.Hyperlinks.Delete
    dst.Font.Name = "Arial": dst.Font.Size = 10
    target.Range(target.Columns(1), target.Columns(metrics.SnapshotLastColumn)).ColumnWidth = 18
    With target.Range(target.Cells(cfg.HeaderRow, 1), target.Cells(cfg.HeaderRow, metrics.SnapshotLastColumn))
        .Interior.Color = RGB(24, 52, 73): .Font.Color = vbWhite: .Font.Bold = True
        .WrapText = True: .RowHeight = 36
    End With
    target.Tab.Color = RGB(163, 175, 185)
    MarkSheetRole target, "SourceSnapshot"
    target.CustomProperties.Add Name:="eTRSourcePath", Value:=metrics.SourcePath
    target.CustomProperties.Add Name:="eTRSourceSheet", Value:=metrics.SourceSheet
    target.CustomProperties.Add Name:="eTRSnapshotLastRow", Value:=CStr(metrics.SnapshotLastRow)
    target.CustomProperties.Add Name:="eTRSnapshotLastColumn", Value:=CStr(metrics.SnapshotLastColumn)
    ' Filtering is a convenience; data validation and snapshot copying must still succeed.
    On Error Resume Next
    target.Range(target.Cells(cfg.HeaderRow, 1), target.Cells(metrics.SnapshotLastRow, metrics.SnapshotLastColumn)).AutoFilter
    On Error GoTo 0
    Exit Sub
Failed:
    savedNumber = Err.Number: savedMessage = Err.Description
    Application.CutCopyMode = False
    Err.Raise savedNumber, "eTR Original Data", "Could not retain Original Data: " & savedMessage
End Sub

Private Sub LoadConfiguration(ByRef cfg As Settings, ByRef regionByTransit As Object, ByRef regionOrder As Object)
    Dim ws As Worksheet, table As ListObject, items As Object, wanted As Object
    Dim values As Variant, key As String, value As Variant, i As Long
    Dim keyCol As Long, valCol As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Configuration")
    On Error GoTo 0
    If ws Is Nothing Then Fail "Configuration is missing. Run ETR_InstallDashboard first."
    Set table = GetTable(ws, "tblSettings")
    keyCol = TableColumn(table, "Setting")
    valCol = TableColumn(table, "Value")
    If table.DataBodyRange Is Nothing Then Fail "The settings table is empty."
    values = table.DataBodyRange.Value2
    Set items = NewDictionary(vbTextCompare)
    For i = 1 To UBound(values, 1)
        key = CleanText(values(i, keyCol), "Setting name")
        value = values(i, valCol)
        If Len(key) = 0 Then
            If Len(CleanText(value, "Setting value")) > 0 Then Fail "A setting value has no setting name."
        Else
            If items.Exists(key) Then Fail "Duplicate setting: " & key & "."
            If IsError(value) Then Fail "Setting " & key & " contains an Excel error."
            items.Add key, value
        End If
    Next i
    cfg.HeaderRow = WholeNumber(RequiredSetting(items, "HeaderRow"), 1, 1048575, "HeaderRow")
    cfg.FirstDataRow = WholeNumber(RequiredSetting(items, "FirstDataRow"), cfg.HeaderRow + 1, 1048576, "FirstDataRow")
    cfg.DescriptionHeader = CleanText(RequiredSetting(items, "DescriptionHeader"), "DescriptionHeader")
    cfg.TransitHeader = CleanText(RequiredSetting(items, "TransitHeader"), "TransitHeader")
    cfg.TRidHeader = CleanText(RequiredSetting(items, "TRidHeader"), "TRidHeader")
    ' Older installations used AgeHeader. Never interpret that column as a date.
    cfg.ApprovalDateHeader = "ApprovalDate"
    If items.Exists("ApprovalDateHeader") Then _
        cfg.ApprovalDateHeader = CleanText(RequiredSetting(items, "ApprovalDateHeader"), "ApprovalDateHeader")
    cfg.DescriptionFilter = CleanText(RequiredSetting(items, "DescriptionFilter", True), "DescriptionFilter")
    cfg.FilterMode = LCase$(CleanText(RequiredSetting(items, "FilterMode"), "FilterMode"))
    cfg.SourceSheet = CleanText(RequiredSetting(items, "SourceSheet", True), "SourceSheet")
    cfg.UnmappedPolicy = LCase$(CleanText(RequiredSetting(items, "UnmappedPolicy"), "UnmappedPolicy"))
    cfg.TransitKeyMode = LCase$(CleanText(RequiredSetting(items, "TransitKeyMode"), "TransitKeyMode"))
    cfg.MaxDataRows = WholeNumber(RequiredSetting(items, "MaxDataRows"), 1, 1048574, "MaxDataRows")
    cfg.MaxSnapshotCells = DEFAULT_SNAPSHOT_CELLS
    If items.Exists("MaxSnapshotCells") Then _
        cfg.MaxSnapshotCells = WholeNumber(items("MaxSnapshotCells"), 4, 100000000, "MaxSnapshotCells")
    If cfg.FilterMode <> "exact" And cfg.FilterMode <> "contains" And cfg.FilterMode <> "all" Then Fail "FilterMode must be Exact, Contains or All."
    If cfg.FilterMode <> "all" And Len(cfg.DescriptionFilter) = 0 Then Fail "DescriptionFilter is required unless FilterMode is All."
    If cfg.UnmappedPolicy <> "include" And cfg.UnmappedPolicy <> "error" Then Fail "UnmappedPolicy must be Include or Error."
    If cfg.TransitKeyMode <> "numeric" And cfg.TransitKeyMode <> "text" Then Fail "TransitKeyMode must be Numeric or Text."
    Set wanted = NewDictionary(vbBinaryCompare)
    For Each value In Array(cfg.DescriptionHeader, cfg.TransitHeader, cfg.TRidHeader, cfg.ApprovalDateHeader)
        key = HeaderKey(value)
        If wanted.Exists(key) Then Fail "The four configured source headers must be distinct."
        wanted.Add key, True
    Next value
    LoadBuckets ws, cfg
    LoadMappings ws, cfg, regionByTransit, regionOrder
End Sub

Private Sub LoadBuckets(ByVal ws As Worksheet, ByRef cfg As Settings)
    Dim table As ListObject, data As Variant, i As Long, j As Long, n As Long
    Dim labelCol As Long, minCol As Long, maxCol As Long, orderCol As Long
    Dim labels As Object, orders As Object, tmp As AgeBucket, label As String
    Dim minText As String, maxText As String, orderText As String, context As String
    Set table = GetTable(ws, "tblAgeBuckets")
    labelCol = TableColumn(table, "BucketLabel"): minCol = TableColumn(table, "MinAge")
    maxCol = TableColumn(table, "MaxAge"): orderCol = TableColumn(table, "BucketOrder")
    If table.DataBodyRange Is Nothing Then Fail "At least one age bucket is required."
    data = table.DataBodyRange.Value2
    ReDim cfg.Buckets(1 To MAX_BUCKETS)
    Set labels = NewDictionary(vbTextCompare): Set orders = NewDictionary(vbBinaryCompare)
    For i = 1 To UBound(data, 1)
        context = "Age bucket table row " & CStr(i)
        label = CleanText(data(i, labelCol), context)
        minText = CleanText(data(i, minCol), context)
        maxText = CleanText(data(i, maxCol), context)
        orderText = CleanText(data(i, orderCol), context)
        If Len(label & minText & maxText & orderText) > 0 Then
            n = n + 1
            If n > MAX_BUCKETS Then Fail "Use at most " & CStr(MAX_BUCKETS) & " age buckets."
            If Len(label) = 0 Or Len(label) > 40 Then Fail context & ": label must contain 1 to 40 characters."
            If labels.Exists(label) Then Fail "Duplicate age bucket label: " & label & "."
            labels.Add label, True
            With cfg.Buckets(n)
                .Label = label
                .MinAge = WholeNumber(data(i, minCol), 0, MAX_AGE, context & " MinAge")
                .OpenEnded = (Len(maxText) = 0)
                If .OpenEnded Then
                    .MaxAge = MAX_AGE
                Else
                    .MaxAge = WholeNumber(data(i, maxCol), .MinAge, MAX_AGE - 1, context & " MaxAge")
                End If
                .SortOrder = WholeNumber(data(i, orderCol), 1, MAX_AGE, context & " BucketOrder")
                If orders.Exists(CStr(.SortOrder)) Then Fail "BucketOrder must be unique."
                orders.Add CStr(.SortOrder), True
            End With
        End If
    Next i
    If n = 0 Then Fail "At least one age bucket is required."
    cfg.BucketCount = n
    For i = 1 To n - 1
        For j = i + 1 To n
            If cfg.Buckets(j).SortOrder < cfg.Buckets(i).SortOrder Then
                tmp = cfg.Buckets(i): cfg.Buckets(i) = cfg.Buckets(j): cfg.Buckets(j) = tmp
            End If
        Next j
    Next i
    If cfg.Buckets(1).MinAge <> 0 Then Fail "The first age bucket must start at 0."
    For i = 1 To n
        If i < n Then
            If cfg.Buckets(i).OpenEnded Then Fail "Only the final age bucket may have a blank MaxAge."
            If cfg.Buckets(i + 1).MinAge <> cfg.Buckets(i).MaxAge + 1 Then Fail "Age buckets must be consecutive, with no gaps or overlaps, in BucketOrder."
        Else
            If Not cfg.Buckets(i).OpenEnded Then Fail "Leave the final age bucket MaxAge blank to cover all remaining ages."
        End If
    Next i
End Sub

Private Sub LoadMappings(ByVal ws As Worksheet, ByRef cfg As Settings, ByRef regionByTransit As Object, ByRef regionOrder As Object)
    Dim table As ListObject, data As Variant, i As Long
    Dim transCol As Long, regCol As Long, orderCol As Long, transit As String, region As String
    Dim ord As Long, regionNames As Object, context As String, t As String, o As String
    Set regionByTransit = NewDictionary(vbBinaryCompare)
    Set regionOrder = NewDictionary(vbTextCompare)
    Set regionNames = NewDictionary(vbTextCompare)
    Set table = GetTable(ws, "tblTransitMapping")
    transCol = TableColumn(table, "booking_transit"): regCol = TableColumn(table, "Region")
    orderCol = TableColumn(table, "RegionOrder")
    If table.DataBodyRange Is Nothing Then Exit Sub
    data = table.DataBodyRange.Value2
    For i = 1 To UBound(data, 1)
        context = "Transit mapping table row " & CStr(i)
        t = CleanText(data(i, transCol), context)
        region = CleanText(data(i, regCol), context)
        o = CleanText(data(i, orderCol), context)
        If Len(t & region & o) > 0 Then
            transit = TransitKey(data(i, transCol), cfg.TransitKeyMode, context)
            If Len(region) = 0 Or Len(region) > 60 Then Fail context & ": Region must contain 1 to 60 characters."
            If StrComp(region, "Unmapped", vbTextCompare) = 0 Then Fail "Unmapped is reserved; leave that transit out of the mapping table."
            ord = WholeNumber(data(i, orderCol), 1, MAX_AGE, context & " RegionOrder")
            If regionByTransit.Exists(transit) Then Fail "Duplicate transit mapping after normalization: " & transit & "."
            If regionOrder.Exists(region) Then
                If CLng(regionOrder(region)) <> ord Then Fail "Use one RegionOrder for region " & region & "."
                region = regionNames(region)
            Else
                regionOrder.Add region, ord
                regionNames.Add region, region
            End If
            regionByTransit.Add transit, region
        End If
    Next i
    regionOrder.Add "Unmapped", 2147483647
End Sub

Private Function TransitKey(ByVal value As Variant, ByVal mode As String, ByVal context As String) As String
    Dim text As String, i As Long, ch As String, n As Double
    text = CleanText(value, context & " booking_transit")
    If Len(text) = 0 Then Fail context & ": booking_transit is blank."
    If VarType(value) = vbBoolean Then Fail context & ": booking_transit cannot be TRUE/FALSE."
    If VarType(value) <> vbString Then
        If Not IsNumeric(value) Then Fail context & ": booking_transit must be a numeric or text identifier."
        n = CDbl(value)
        If n < 0 Or n > 999999999999999# Or n <> Fix(n) Then Fail context & ": numeric booking_transit must be a nonnegative integer of at most 15 digits. Store longer identifiers as text."
        text = Format$(n, "0")
    End If
    If mode = "numeric" Then
        For i = 1 To Len(text)
            ch = Mid$(text, i, 1)
            If ch < "0" Or ch > "9" Then Fail context & ": Numeric transit mode accepts only digits (no signs, decimals or separators)."
        Next i
        Do While Len(text) > 1
            If Left$(text, 1) <> "0" Then Exit Do
            text = Mid$(text, 2)
        Loop
    Else
        text = UCase$(text)
    End If
    If Len(text) > 80 Then Fail context & ": booking_transit exceeds 80 characters."
    TransitKey = text
End Function

Private Function TransactionKey(ByVal value As Variant, ByVal context As String) As String
    Dim text As String, n As Double
    text = CleanText(value, context & " TRid")
    If Len(text) = 0 Then Fail context & ": TRid is blank."
    If VarType(value) = vbBoolean Then Fail context & ": TRid cannot be TRUE/FALSE."
    If VarType(value) <> vbString Then
        If Not IsNumeric(value) Then Fail context & ": TRid must be text or a whole-number identifier."
        n = CDbl(value)
        If n < 0 Or n > 999999999999999# Or n <> Fix(n) Then Fail context & ": numeric TRid must be a nonnegative integer of at most 15 digits. Store longer identifiers as text."
        text = Format$(n, "0")
    End If
    If Len(text) > 255 Then Fail context & ": TRid exceeds 255 characters."
    TransactionKey = text ' Deliberately case-sensitive. Leading zeros in text TRids are preserved.
End Function

Private Sub ValidateSourcePath(ByVal sourcePath As String)
    Dim ext As String, book As Workbook, sourceName As String, normalizedPath As String
    If Len(Trim$(sourcePath)) = 0 Then Fail "No source file was selected."
    ext = LCase$(Mid$(sourcePath, InStrRev(sourcePath, ".") + 1))
    If ext <> "xlsx" And ext <> "xlsm" And ext <> "xlsb" Then Fail "Select an .xlsx, .xlsm or .xlsb workbook. Convert legacy .xls or CSV files to .xlsx first."
    normalizedPath = Replace(sourcePath, "/", "\")
    sourceName = Mid$(normalizedPath, InStrRev(normalizedPath, "\") + 1)
    If StrComp(sourcePath, ThisWorkbook.FullName, vbTextCompare) = 0 Or _
        StrComp(sourceName, ThisWorkbook.Name, vbTextCompare) = 0 Then Fail "Select a source workbook, not this dashboard."
    For Each book In Application.Workbooks
        ' OneDrive may expose FullName as a cloud URL even when the picker gives a local path.
        ' Excel cannot open two workbooks with the same base filename in one instance.
        If StrComp(sourcePath, book.FullName, vbTextCompare) = 0 Or _
            StrComp(sourceName, book.Name, vbTextCompare) = 0 Then _
            Fail "Close the selected source workbook first, saving your changes, then refresh."
    Next book
End Sub

Private Function ResolveSourceSheet(ByVal book As Workbook, ByRef cfg As Settings, ByRef columns() As Long) As Worksheet
    Dim ws As Worksheet, candidate As Worksheet, count As Long, matches As Boolean
    Dim testColumns() As Long, duplicate As String, foundDuplicate As String
    For Each ws In book.Worksheets
        If Len(cfg.SourceSheet) = 0 Or StrComp(ws.Name, cfg.SourceSheet, vbTextCompare) = 0 Then
            matches = MatchHeaders(ws, cfg, testColumns, duplicate)
            If matches Then
                count = count + 1
                Set candidate = ws
                foundDuplicate = duplicate
            ElseIf Len(cfg.SourceSheet) > 0 Then
                Fail "Worksheet '" & ws.Name & "' is missing required headers in row " & CStr(cfg.HeaderRow) & ": " & _
                    cfg.DescriptionHeader & ", " & cfg.TransitHeader & ", " & cfg.TRidHeader & ", " & cfg.ApprovalDateHeader & "."
            End If
        End If
    Next ws
    If count = 0 Then Fail "No source worksheet has all four configured headers in row " & CStr(cfg.HeaderRow) & ". Check SourceSheet and the header settings."
    If count > 1 Then Fail "More than one source worksheet matches. Set SourceSheet to the intended worksheet name."
    If Len(foundDuplicate) > 0 Then Fail "Duplicate required header on '" & candidate.Name & "': " & foundDuplicate & "."
    Call MatchHeaders(candidate, cfg, columns, duplicate)
    Set ResolveSourceSheet = candidate
End Function

Private Function MatchHeaders(ByVal ws As Worksheet, ByRef cfg As Settings, ByRef columns() As Long, ByRef duplicate As String) As Boolean
    Dim wanted As Variant, data As Variant, lastCol As Long, col As Long, i As Long, key As String
    ReDim columns(1 To 4)
    duplicate = vbNullString
    wanted = Array(cfg.DescriptionHeader, cfg.TransitHeader, cfg.TRidHeader, cfg.ApprovalDateHeader)
    lastCol = ws.Cells(cfg.HeaderRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 4 Then Exit Function
    data = ws.Range(ws.Cells(cfg.HeaderRow, 1), ws.Cells(cfg.HeaderRow, lastCol)).Value2
    For col = 1 To lastCol
        key = HeaderKey(data(1, col))
        For i = 1 To 4
            If key = HeaderKey(wanted(i - 1)) Then
                If columns(i) <> 0 Then duplicate = CStr(wanted(i - 1))
                columns(i) = col
            End If
        Next i
    Next col
    For i = 1 To 4
        If columns(i) = 0 Then Exit Function
    Next i
    MatchHeaders = True
End Function

Private Function ReadColumn(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal col As Long) As Variant
    Dim one(1 To 1, 1 To 1) As Variant
    If firstRow = lastRow Then
        one(1, 1) = ws.Cells(firstRow, col).Value2
        ReadColumn = one
    Else
        ReadColumn = ws.Range(ws.Cells(firstRow, col), ws.Cells(lastRow, col)).Value2
    End If
End Function

Private Function IsBlankValue(ByVal value As Variant) As Boolean
    If IsError(value) Then Exit Function
    IsBlankValue = (Len(CleanText(value, "Source cell")) = 0)
End Function

Private Function DescriptionMatches(ByVal text As String, ByRef cfg As Settings) As Boolean
    Select Case cfg.FilterMode
        Case "all": DescriptionMatches = True
        Case "exact": DescriptionMatches = (StrComp(text, cfg.DescriptionFilter, vbTextCompare) = 0)
        Case "contains": DescriptionMatches = (InStr(1, text, cfg.DescriptionFilter, vbTextCompare) > 0)
    End Select
End Function

Private Function ApprovalDay(ByVal value As Variant, ByVal date1904 As Boolean, ByVal asOfDate As Date, ByVal context As String) As Date
    Dim serial As Double, result As Date, text As String, yyyy As Long, mm As Long, dd As Long
    If IsError(value) Then Fail context & " ApprovalDate contains an Excel error."
    If IsEmpty(value) Or IsNull(value) Then Fail context & " ApprovalDate is blank."
    If VarType(value) = vbBoolean Then Fail context & " ApprovalDate must be a valid date."
    If VarType(value) = vbString Then
        text = Trim$(CStr(value))
        If Len(text) = 0 Then Fail context & " ApprovalDate is blank."
        If Not text Like "####-##-##" Then _
            Fail context & " ApprovalDate must be an Excel date or ISO text yyyy-mm-dd."
        yyyy = CLng(Left$(text, 4)): mm = CLng(Mid$(text, 6, 2)): dd = CLng(Right$(text, 2))
        If yyyy < 1900 Or mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Then GoTo InvalidDate
        On Error GoTo InvalidDate
        result = DateSerial(yyyy, mm, dd)
        On Error GoTo 0
        If Year(result) <> yyyy Or Month(result) <> mm Or Day(result) <> dd Then GoTo InvalidDate
    ElseIf IsNumeric(value) Then
        serial = Int(CDbl(value)) ' Ignore timestamps; measure calendar days.
        If date1904 Then
            If serial < 0 Then GoTo InvalidDate
            serial = serial + 1462
        Else
            If serial < 1 Or serial = 60 Then GoTo InvalidDate ' Excel's fictitious 1900-02-29.
            If serial < 60 Then serial = serial + 1 ' Excel/OLE dates before March 1900.
        End If
        If serial > 2958465 Then GoTo InvalidDate
        result = CDate(serial)
    Else
        GoTo InvalidDate
    End If
    If result > asOfDate Then Fail context & " ApprovalDate is in the future (after " & Format$(asOfDate, "yyyy-mm-dd") & ")."
    ApprovalDay = result
    Exit Function
InvalidDate:
    On Error GoTo 0
    Fail context & " ApprovalDate is not a valid calendar date."
End Function

Private Function BucketForAge(ByVal age As Long, ByRef cfg As Settings) As Long
    Dim i As Long
    For i = 1 To cfg.BucketCount
        If age >= cfg.Buckets(i).MinAge And age <= cfg.Buckets(i).MaxAge Then
            BucketForAge = i
            Exit Function
        End If
    Next i
    Fail "Calculated age " & CStr(age) & " does not match an age bucket."
End Function

Private Sub IncrementCount(ByVal counts As Object, ByVal transit As String, ByVal bucket As Long, ByVal nBuckets As Long)
    Dim values As Variant, fresh() As Long
    If Not counts.Exists(transit) Then
        ReDim fresh(1 To nBuckets)
        counts.Add transit, fresh
    End If
    values = counts(transit)
    values(bucket) = values(bucket) + 1
    counts(transit) = values
End Sub

Private Sub AggregateSource(ByVal ws As Worksheet, ByRef columns() As Long, ByRef cfg As Settings, _
        ByVal mapping As Object, ByVal gross As Object, ByVal distinct As Object, ByRef metrics As RunMetrics)
    Dim approvalById As Object, pairSeen As Object, unmapped As Object
    Dim lastRow As Long, endRow As Long, startRow As Long, i As Long, j As Long, r As Long
    Dim descriptions As Variant, transits As Variant, ids As Variant, approvals As Variant
    Dim context As String, description As String, transit As String, id As String, pair As String
    Dim approved As Date, age As Long, bucket As Long, firstSeen As Variant, key As Variant, counts As Variant
    Set approvalById = NewDictionary(vbBinaryCompare)
    Set pairSeen = NewDictionary(vbBinaryCompare)
    Set unmapped = NewDictionary(vbBinaryCompare)
    For i = 1 To 4
        r = ws.Cells(ws.Rows.Count, columns(i)).End(xlUp).Row
        If r > lastRow Then lastRow = r
    Next i
    If lastRow < cfg.FirstDataRow Then Exit Sub ' Valid empty source.
    If lastRow - cfg.FirstDataRow + 1 > cfg.MaxDataRows Then Fail "Source row span exceeds MaxDataRows (" & CStr(cfg.MaxDataRows) & "). Check trailing content or increase the limit."
    For startRow = cfg.FirstDataRow To lastRow Step CHUNK_SIZE
        endRow = startRow + CHUNK_SIZE - 1
        If endRow > lastRow Then endRow = lastRow
        descriptions = ReadColumn(ws, startRow, endRow, columns(1))
        transits = ReadColumn(ws, startRow, endRow, columns(2))
        ids = ReadColumn(ws, startRow, endRow, columns(3))
        approvals = ReadColumn(ws, startRow, endRow, columns(4))
        For i = 1 To endRow - startRow + 1
            r = startRow + i - 1
            context = "Source row " & CStr(r)
            If IsBlankValue(descriptions(i, 1)) And IsBlankValue(transits(i, 1)) And _
                    IsBlankValue(ids(i, 1)) And IsBlankValue(approvals(i, 1)) Then
                metrics.BlankRows = metrics.BlankRows + 1
            Else
                metrics.ReadRows = metrics.ReadRows + 1
                description = CleanText(descriptions(i, 1), context & " Description")
                If DescriptionMatches(description, cfg) Then
                    transit = TransitKey(transits(i, 1), cfg.TransitKeyMode, context)
                    id = TransactionKey(ids(i, 1), context)
                    approved = ApprovalDay(approvals(i, 1), ws.Parent.Date1904, metrics.AsOfDate, context)
                    age = DateDiff("d", approved, metrics.AsOfDate)
                    bucket = BucketForAge(age, cfg)
                    If approvalById.Exists(id) Then
                        firstSeen = approvalById(id)
                        If CDbl(firstSeen(0)) <> CDbl(approved) Then Fail context & ": TRid '" & id & _
                            "' has ApprovalDate " & Format$(approved, "yyyy-mm-dd") & "; row " & CStr(firstSeen(1)) & _
                            " has " & Format$(CDate(firstSeen(0)), "yyyy-mm-dd") & ". One TRid must have one approval date, even across transits."
                    Else
                        approvalById.Add id, Array(CDbl(approved), r)
                    End If
                    If Not mapping.Exists(transit) Then
                        If cfg.UnmappedPolicy = "error" Then Fail context & ": transit " & transit & " has no region mapping."
                        metrics.UnmappedRows = metrics.UnmappedRows + 1
                        If Not unmapped.Exists(transit) Then unmapped.Add transit, True
                    End If
                    IncrementCount gross, transit, bucket, cfg.BucketCount
                    metrics.GrossRows = metrics.GrossRows + 1
                    ' Length prefixes prevent composite-key collisions even when IDs contain punctuation.
                    pair = CStr(Len(transit)) & ":" & transit & CStr(Len(id)) & ":" & id
                    If Not pairSeen.Exists(pair) Then
                        pairSeen.Add pair, True
                        IncrementCount distinct, transit, bucket, cfg.BucketCount
                    End If
                Else
                    metrics.ExcludedRows = metrics.ExcludedRows + 1
                End If
            End If
        Next i
        Application.StatusBar = "eTR: read through source row " & Format$(endRow, "#,##0") & "..."
    Next startRow
    metrics.UniquePairs = pairSeen.Count
    metrics.GlobalIds = approvalById.Count
    metrics.UnmappedTransits = unmapped.Count
    If metrics.ReadRows <> metrics.GrossRows + metrics.ExcludedRows Then Fail "Source row reconciliation failed."
    For Each key In gross.Keys
        counts = gross(key)
        For j = 1 To cfg.BucketCount
            If counts(j) > metrics.MaximumCell Then metrics.MaximumCell = counts(j)
        Next j
    Next key
End Sub

Private Function RegionForTransit(ByVal transit As String, ByVal mapping As Object) As String
    If mapping.Exists(transit) Then RegionForTransit = mapping(transit) Else RegionForTransit = "Unmapped"
End Function

Private Function RegionRank(ByVal region As String, ByVal orders As Object) As Long
    If orders.Exists(region) Then RegionRank = CLng(orders(region)) Else RegionRank = 2147483647
End Function

Private Function CompareTransits(ByVal a As String, ByVal b As String, ByVal mapping As Object, ByVal orders As Object, ByVal mode As String) As Long
    Dim ra As String, rb As String, oa As Long, ob As Long
    ra = RegionForTransit(a, mapping): rb = RegionForTransit(b, mapping)
    oa = RegionRank(ra, orders): ob = RegionRank(rb, orders)
    If oa < ob Then CompareTransits = -1: Exit Function
    If oa > ob Then CompareTransits = 1: Exit Function
    CompareTransits = StrComp(ra, rb, vbTextCompare)
    If CompareTransits <> 0 Then Exit Function
    If mode = "numeric" Then
        If Len(a) < Len(b) Then CompareTransits = -1: Exit Function
        If Len(a) > Len(b) Then CompareTransits = 1: Exit Function
    End If
    CompareTransits = StrComp(a, b, vbBinaryCompare)
End Function

Private Sub SortTransits(ByRef keys As Variant, ByVal first As Long, ByVal last As Long, ByVal mapping As Object, ByVal orders As Object, ByVal mode As String)
    Dim i As Long, j As Long, pivot As String, tmp As Variant
    i = first: j = last: pivot = CStr(keys((first + last) \ 2))
    Do While i <= j
        Do While CompareTransits(CStr(keys(i)), pivot, mapping, orders, mode) < 0: i = i + 1: Loop
        Do While CompareTransits(CStr(keys(j)), pivot, mapping, orders, mode) > 0: j = j - 1: Loop
        If i <= j Then
            tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then SortTransits keys, first, j, mapping, orders, mode
    If i < last Then SortTransits keys, i, last, mapping, orders, mode
End Sub

Private Sub PutMerged(ByVal ws As Worksheet, ByVal row As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal value As Variant)
    With ws.Range(ws.Cells(row, firstCol), ws.Cells(row, lastCol))
        .Merge
        .NumberFormat = "@"
        .Value2 = value
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub RenderDashboard(ByVal ws As Worksheet, ByRef cfg As Settings, ByVal mapping As Object, ByVal orders As Object, _
        ByVal gross As Object, ByVal distinct As Object, ByRef metrics As RunMetrics, Optional ByVal hasSource As Boolean = True)
    Dim keys As Variant, lastCol As Long, viewLastCol As Long, grossTotal As Long, uniqueTotal As Long
    Dim starts As Variant, labels As Variant, captions As Variant, i As Long, col As Long, endCol As Long
    Dim uniqueStart As Long, filterText As String, overallLastRow As Long
    Dim regionCount As Long, previousRegion As String, currentRegion As String
    lastCol = cfg.BucketCount + 4
    viewLastCol = lastCol: If viewLastCol < 9 Then viewLastCol = 9
    If gross.Count > 0 Then
        keys = gross.Keys
        If gross.Count > 1 Then SortTransits keys, LBound(keys), UBound(keys), mapping, orders, cfg.TransitKeyMode
    End If
    ws.Tab.Color = RGB(24, 52, 73)
    ws.Columns(1).ColumnWidth = 3
    ws.Columns(2).ColumnWidth = 26
    ws.Columns(3).ColumnWidth = 13
    ws.Range(ws.Columns(4), ws.Columns(viewLastCol)).ColumnWidth = 13
    ' Styling is bounded to the generated report, not the entire worksheet.
    If gross.Count > 0 Then
        For i = LBound(keys) To UBound(keys)
            currentRegion = RegionForTransit(CStr(keys(i)), mapping)
            If currentRegion <> previousRegion Then regionCount = regionCount + 1
            previousRegion = currentRegion
        Next i
    End If
    overallLastRow = 2 * (gross.Count + regionCount) + 50
    If overallLastRow > 1048576 Then Fail "Too many transit rows for two heatmaps on one Excel worksheet."
    With ws.Range(ws.Cells(2, 2), ws.Cells(overallLastRow, viewLastCol))
        .Font.Name = "Arial": .Font.Size = 10: .Font.Color = RGB(32, 52, 70)
        .RowHeight = 22: .VerticalAlignment = xlCenter
    End With
    PutMerged ws, 2, 2, 6, "eTR heatmap dashboard"
    With ws.Range("B2:F3")
        .UnMerge: .Merge: .Font.Size = 16: .Font.Bold = True: .Font.Color = RGB(24, 52, 73)
    End With
    ws.Range(ws.Cells(4, 2), ws.Cells(4, viewLastCol)).Borders(xlEdgeBottom).Color = RGB(24, 52, 73)
    PutMerged ws, 5, 2, viewLastCol, "Last refresh succeeded. Dashboard and Original Data contain the same source snapshot."
    ws.Range(ws.Cells(5, 2), ws.Cells(5, viewLastCol)).Interior.Color = RGB(231, 238, 245)
    ws.Rows(5).RowHeight = 30
    ws.Cells(6, 2).Value2 = "Source file"
    PutMerged ws, 6, 3, viewLastCol, metrics.SourcePath
    ws.Range(ws.Cells(6, 3), ws.Cells(6, viewLastCol)).WrapText = True
    ws.Rows(6).RowHeight = 32
    ws.Cells(7, 2).Value2 = "Source worksheet"
    PutMerged ws, 7, 3, viewLastCol, metrics.SourceSheet
    ws.Cells(8, 2).Value2 = "Last successful refresh"
    PutMerged ws, 8, 3, viewLastCol, CDbl(metrics.CompletedAt)
    ws.Range("C8").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(9, 2).Value2 = "Applied filter"
    If cfg.FilterMode = "all" Then filterText = "All descriptions" Else filterText = cfg.DescriptionFilter & " (" & cfg.FilterMode & ")"
    PutMerged ws, 9, 3, viewLastCol, filterText & "; transit mode: " & cfg.TransitKeyMode
    If hasSource Then
        PutMerged ws, 10, 2, viewLastCol, "Age as of " & Format$(metrics.AsOfDate, "yyyy-mm-dd") & " = refresh date minus " & cfg.ApprovalDateHeader & " (calendar days)."
    Else
        PutMerged ws, 10, 2, viewLastCol, "Age = refresh date minus ApprovalDate (calendar days). Refresh to update ages."
    End If
    ws.Range(ws.Cells(10, 2), ws.Cells(10, viewLastCol)).Font.Color = RGB(90, 110, 127)
    ws.Range(ws.Cells(10, 2), ws.Cells(10, viewLastCol)).WrapText = True: ws.Rows(10).RowHeight = 26
    starts = Array(2, 4, 6, 8)
    labels = Array("Gross eTR rows", "Unique by transit", "Global TRids", "Duplicates removed")
    captions = Array("Qualifying source rows", "Distinct transit + TRid", "Distinct IDs overall", "Gross minus unique")
    For i = 0 To 3
        col = CLng(starts(i)): endCol = col + 1
        PutMerged ws, 11, col, endCol, labels(i)
        With ws.Range(ws.Cells(11, col), ws.Cells(11, endCol))
            .Interior.Color = RGB(231, 238, 245): .Font.Bold = True: .RowHeight = 25
        End With
        PutMerged ws, 12, col, endCol, 0
        With ws.Range(ws.Cells(12, col), ws.Cells(12, endCol))
            .NumberFormat = "#,##0": .Font.Size = 16: .Font.Bold = True: .RowHeight = 30
        End With
        PutMerged ws, 13, col, endCol, captions(i)
        With ws.Range(ws.Cells(13, col), ws.Cells(13, endCol))
            .Font.Size = 9: .Font.Color = RGB(90, 110, 127): .WrapText = True: .RowHeight = 25
        End With
    Next i
    PutMerged ws, 14, 2, viewLastCol, "Read " & Format$(metrics.ReadRows, "#,##0") & _
        " rows; filtered out " & Format$(metrics.ExcludedRows, "#,##0") & _
        "; blank rows " & Format$(metrics.BlankRows, "#,##0") & _
        "; unmapped " & Format$(metrics.UnmappedRows, "#,##0") & " rows across " & CStr(metrics.UnmappedTransits) & " transits."
    ws.Range(ws.Cells(14, 2), ws.Cells(14, viewLastCol)).WrapText = True
    ws.Rows(14).RowHeight = 30
    grossTotal = RenderHeatmap(ws, 16, "01  Gross eTR rows", keys, gross, mapping, cfg, metrics.MaximumCell, hasSource)
    uniqueStart = grossTotal + 4
    uniqueTotal = RenderHeatmap(ws, uniqueStart, "02  Unique eTR by transit", keys, distinct, mapping, cfg, metrics.MaximumCell, hasSource)
    ws.Range("B12").Formula = "=" & ws.Cells(grossTotal, lastCol).Address(False, False)
    ws.Range("D12").Formula = "=" & ws.Cells(uniqueTotal, lastCol).Address(False, False)
    ws.Range("F12").Value2 = metrics.GlobalIds
    ws.Range("H12").Formula = "=B12-D12"
    ws.Range("H13").NumberFormat = "General"
    ws.Range("H13").Formula = "=IF(B12=0,""Rate: n.a."",""Rate: ""&TEXT(H12/B12,""0.0%""))"
    If Not hasSource Then
        ws.Range("B5").Value2 = "No source loaded. Configure your regions, then use Refresh eTR file."
        ws.Range(ws.Cells(5, 2), ws.Cells(5, viewLastCol)).Interior.Color = RGB(255, 243, 204)
        ws.Range("C6").Value2 = "n.a.": ws.Range("C7").Value2 = "n.a."
        ws.Range("C8").NumberFormat = "@": ws.Range("C8").Value2 = "n.a."
        ws.Range("C9").Value2 = filterText & " (pending first refresh); transit mode: " & cfg.TransitKeyMode
        ws.Range("F12").Value2 = "n.a.": ws.Range("H12").Value2 = "n.a."
        ws.Range("H13").Value2 = "Gross minus unique"
        ws.Range("B14").Value2 = "Results remain unavailable until the first successful refresh. Original Data will retain all source rows."
    End If
    PutMerged ws, uniqueTotal + 3, 2, viewLastCol, "Unique totals count transit + TRid pairs. The same TRid in different transits contributes once in each."
    ws.Range(ws.Cells(uniqueTotal + 3, 2), ws.Cells(uniqueTotal + 3, viewLastCol)).WrapText = True
    ws.Rows(uniqueTotal + 3).RowHeight = 32
    AddControls ws
    MarkSheetRole ws, "Dashboard"
    ' Page setup is optional; printer availability cannot invalidate correct results.
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape: .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .PrintArea = ws.Range(ws.Cells(2, 2), ws.Cells(uniqueTotal + 3, viewLastCol)).Address
    End With
    On Error GoTo 0
End Sub

Private Function RenderHeatmap(ByVal ws As Worksheet, ByVal topRow As Long, ByVal title As String, ByVal keys As Variant, _
        ByVal counts As Object, ByVal mapping As Object, ByRef cfg As Settings, ByVal maximum As Long, ByVal hasSource As Boolean) As Long
    Dim lastCol As Long, titleLastCol As Long, r As Long, headerRow As Long, firstData As Long
    Dim i As Long, j As Long, firstIndex As Long, lastIndex As Long, subtotal As Long
    Dim region As String, transit As String, rowCounts As Variant, block As Variant
    Dim detailStart As Long, detailEnd As Long, detailCount As Long, summaryRows As Collection, item As Variant
    Dim header() As Variant
    lastCol = cfg.BucketCount + 4
    titleLastCol = lastCol: If titleLastCol < 9 Then titleLastCol = 9
    PutMerged ws, topRow, 2, titleLastCol, title
    With ws.Range(ws.Cells(topRow, 2), ws.Cells(topRow, titleLastCol))
        .Font.Size = 12: .Font.Bold = True: .RowHeight = 27
    End With
    PutMerged ws, topRow + 1, 2, titleLastCol, "Age from ApprovalDate (days). Shared heat-color scale: 0 to " & CStr(maximum) & ". Region rows are subtotals."
    If Not hasSource Then ws.Cells(topRow + 1, 2).Value2 = "Both heatmaps use the same region, transit and age-bucket layout."
    ws.Range(ws.Cells(topRow + 1, 2), ws.Cells(topRow + 1, titleLastCol)).Font.Color = RGB(90, 110, 127)
    headerRow = topRow + 2
    ReDim header(1 To 1, 1 To cfg.BucketCount + 3)
    header(1, 1) = "Region": header(1, 2) = "Transit"
    For j = 1 To cfg.BucketCount: header(1, j + 2) = cfg.Buckets(j).Label: Next j
    header(1, cfg.BucketCount + 3) = "Total"
    With ws.Range(ws.Cells(headerRow, 2), ws.Cells(headerRow, lastCol))
        .NumberFormat = "@": .Value2 = header
        .Interior.Color = RGB(24, 52, 73): .Font.Color = vbWhite: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .WrapText = True: .RowHeight = 34
    End With
    r = headerRow + 1: firstData = r
    Set summaryRows = New Collection
    If counts.Count > 0 Then
        i = LBound(keys)
        Do While i <= UBound(keys)
            firstIndex = i
            region = RegionForTransit(CStr(keys(i)), mapping)
            lastIndex = i
            Do While lastIndex < UBound(keys)
                If StrComp(RegionForTransit(CStr(keys(lastIndex + 1)), mapping), region, vbTextCompare) <> 0 Then Exit Do
                lastIndex = lastIndex + 1
            Loop
            subtotal = r: summaryRows.Add subtotal
            ws.Cells(subtotal, 2).NumberFormat = "@": ws.Cells(subtotal, 2).Value2 = region
            With ws.Range(ws.Cells(subtotal, 2), ws.Cells(subtotal, lastCol))
                .Interior.Color = RGB(231, 238, 245): .Font.Bold = True: .RowHeight = 30
            End With
            ws.Cells(subtotal, 2).WrapText = True
            If Len(region) > 28 Then ws.Rows(subtotal).RowHeight = 45
            detailStart = subtotal + 1
            detailCount = lastIndex - firstIndex + 1
            detailEnd = detailStart + detailCount - 1
            ReDim block(1 To detailCount, 1 To cfg.BucketCount + 1)
            For i = firstIndex To lastIndex
                transit = CStr(keys(i)): rowCounts = counts(transit)
                block(i - firstIndex + 1, 1) = transit
                For j = 1 To cfg.BucketCount
                    block(i - firstIndex + 1, j + 1) = CLng(rowCounts(j))
                Next j
            Next i
            ws.Range(ws.Cells(detailStart, 3), ws.Cells(detailEnd, 3)).NumberFormat = "@"
            ws.Range(ws.Cells(detailStart, 3), ws.Cells(detailEnd, lastCol - 1)).Value2 = block
            ws.Range(ws.Cells(detailStart, 4), ws.Cells(detailEnd, lastCol)).NumberFormat = "#,##0"
            ws.Range(ws.Cells(detailStart, 4), ws.Cells(detailEnd, lastCol)).HorizontalAlignment = xlRight
            ws.Range(ws.Cells(detailStart, lastCol), ws.Cells(detailEnd, lastCol)).FormulaR1C1 = "=SUM(RC4:RC[-1])"
            ws.Range(ws.Cells(detailStart, lastCol), ws.Cells(detailEnd, lastCol)).Font.Bold = True
            For j = 4 To lastCol - 1
                ws.Cells(subtotal, j).Formula = "=SUM(" & ws.Range(ws.Cells(detailStart, j), ws.Cells(detailEnd, j)).Address(False, False) & ")"
            Next j
            ws.Cells(subtotal, lastCol).FormulaR1C1 = "=SUM(RC4:RC[-1])"
            ApplyHeatColors ws.Range(ws.Cells(detailStart, 4), ws.Cells(detailEnd, lastCol - 1)), maximum
            r = detailEnd + 1
            i = lastIndex + 1
        Loop
    Else
        If hasSource Then
            PutMerged ws, r, 2, lastCol, "No qualifying records for this filter."
        Else
            PutMerged ws, r, 2, lastCol, "No source workbook has been loaded."
        End If
        r = r + 1
    End If
    ws.Cells(r, 2).Value2 = "Grand total"
    For j = 4 To lastCol - 1
        If Not hasSource Then
            ws.Cells(r, j).Value2 = "n.a."
        ElseIf counts.Count = 0 Then
            ws.Cells(r, j).Value2 = 0
        Else
            ' Only transit rows have a value in column C. Never sum region totals twice.
            ws.Cells(r, j).Formula = "=SUMIF(" & ws.Range(ws.Cells(firstData, 3), ws.Cells(r - 1, 3)).Address & _
                ",""<>""," & ws.Range(ws.Cells(firstData, j), ws.Cells(r - 1, j)).Address & ")"
        End If
    Next j
    If hasSource Then
        ws.Cells(r, lastCol).FormulaR1C1 = "=SUM(RC4:RC[-1])"
    Else
        ws.Cells(r, lastCol).Value2 = "n.a."
    End If
    ws.Range(ws.Cells(firstData, 4), ws.Cells(r, lastCol)).NumberFormat = "#,##0"
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, lastCol))
        .Interior.Color = RGB(24, 52, 73): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 28
    End With
    RenderHeatmap = r
End Function

Private Sub ApplyHeatColors(ByVal cells As Range, ByVal maximum As Long)
    Dim heatScale As ColorScale, high As Double
    high = maximum: If high < 1 Then high = 1
    Set heatScale = cells.FormatConditions.AddColorScale(3)
    With heatScale.ColorScaleCriteria(1)
        .Type = xlConditionValueNumber: .Value = 0: .FormatColor.Color = RGB(245, 248, 252)
    End With
    With heatScale.ColorScaleCriteria(2)
        .Type = xlConditionValueNumber: .Value = high / 2: .FormatColor.Color = RGB(249, 211, 144)
    End With
    With heatScale.ColorScaleCriteria(3)
        .Type = xlConditionValueNumber: .Value = high: .FormatColor.Color = RGB(240, 144, 117)
    End With
End Sub

Private Sub AddControls(ByVal ws As Worksheet)
    Dim button As Shape, area As Range
    On Error Resume Next
    ws.Shapes("eTR_Refresh").Delete
    ws.Shapes("eTR_Configure").Delete
    ws.Shapes("eTR_OriginalData").Delete
    On Error GoTo 0
    ws.Range("H2:I3").UnMerge
    ws.Range("H2:I3").ClearContents
    ws.Range("H2:I3").Interior.Pattern = xlNone
    Set area = ws.Range("H2:I3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    With button
        .Name = "eTR_Refresh": .OnAction = "RefreshETRDashboard"
        .TextFrame2.TextRange.Text = "Refresh eTR file"
        .TextFrame2.TextRange.Font.Name = "Arial": .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
        .TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        .TextFrame2.VerticalAnchor = 3
        .Fill.ForeColor.RGB = RGB(24, 52, 73): .Line.Visible = 0
        .AlternativeText = "Select an eTR workbook and refresh both heatmaps."
        .Placement = xlFreeFloating
    End With
    Set area = ws.Range("G2:G3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width - 5, area.Height)
    With button
        .Name = "eTR_Configure": .OnAction = "ETR_ShowConfiguration"
        .TextFrame2.TextRange.Text = "Configure"
        .TextFrame2.TextRange.Font.Name = "Arial": .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(24, 52, 73)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = 2: .TextFrame2.VerticalAnchor = 3
        .Fill.ForeColor.RGB = RGB(231, 238, 245): .Line.Visible = 0
        .Placement = xlFreeFloating
    End With
    Set area = ws.Range("B15:C15")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top + 1, 150, area.Height - 2)
    With button
        .Name = "eTR_OriginalData": .OnAction = "ETR_ShowOriginalData"
        .TextFrame2.TextRange.Text = "View Original Data"
        .TextFrame2.TextRange.Font.Name = "Arial": .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(24, 52, 73)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = 2: .TextFrame2.VerticalAnchor = 3
        .Fill.ForeColor.RGB = RGB(231, 238, 245): .Line.Visible = 0
        .AlternativeText = "Open the unfiltered source snapshot from the last successful refresh."
        .Placement = xlFreeFloating
    End With
End Sub

Private Function UnusedSheetName(ByVal prefix As String) As String
    Dim i As Long, candidate As String, ws As Worksheet
    For i = 1 To 9999
        candidate = prefix & CStr(i)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(candidate)
        On Error GoTo 0
        If ws Is Nothing Then UnusedSheetName = candidate: Exit Function
    Next i
    Fail "Cannot allocate a temporary worksheet name."
End Function

Private Sub CommitOutputs(ByVal staged As Worksheet, ByVal stagedData As Worksheet)
    Dim current As Worksheet, currentData As Worksheet, backupName As String, backupDataName As String
    Dim stagedName As String, stagedDataName As String
    Dim renamedOld As Boolean, renamedNew As Boolean, renamedOldData As Boolean, renamedNewData As Boolean, oldAlerts As Boolean
    Dim number As Long, message As String
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Rollback
    Set current = ThisWorkbook.Worksheets("Dashboard")
    Set currentData = ThisWorkbook.Worksheets("Original Data")
    backupName = UnusedSheetName("_eTR_previous")
    backupDataName = UnusedSheetName("_eTR_prior_data")
    stagedName = staged.Name
    stagedDataName = stagedData.Name
    staged.Visible = xlSheetVisible
    stagedData.Visible = xlSheetVisible
    staged.Move Before:=current
    stagedData.Move After:=ThisWorkbook.Worksheets("Configuration")
    current.Name = backupName
    renamedOld = True
    currentData.Name = backupDataName
    renamedOldData = True
    staged.Name = "Dashboard"
    renamedNew = True
    stagedData.Name = "Original Data"
    renamedNewData = True
    Application.DisplayAlerts = False
    ' Delete the old pair in one Excel operation, after both replacements are ready.
    ThisWorkbook.Worksheets(Array(backupName, backupDataName)).Delete
    On Error Resume Next
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    Exit Sub
Rollback:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = oldAlerts
    If renamedNewData Then stagedData.Name = stagedDataName
    If renamedNew Then staged.Name = stagedName
    If renamedOldData Then currentData.Name = "Original Data"
    If renamedOld Then current.Name = "Dashboard"
    current.Visible = xlSheetVisible
    currentData.Visible = xlSheetVisible
    On Error GoTo 0
    Err.Raise number, "eTR dashboard commit", message
End Sub
