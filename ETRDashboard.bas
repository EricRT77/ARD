Attribute VB_Name = "ETRDashboard"
Option Explicit

' eTR dashboard v1.0. Desktop Excel for Windows, 32- or 64-bit.
' Import this module into eTR_Dashboard_Template.xlsx, save as .xlsm,
' then run ETR_InstallDashboard. No reference libraries or add-ins needed.
' All source workbooks are opened read-only. No source data are retained.

Private Const ERR_BASE As Long = vbObjectError + 741
Private Const CHUNK_SIZE As Long = 50000
Private Const MAX_AGE As Long = 2147483646
Private Const MAX_BUCKETS As Long = 12
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
    AgeHeader As String
    DescriptionFilter As String
    FilterMode As String
    SourceSheet As String
    UnmappedPolicy As String
    TransitKeyMode As String
    MaxDataRows As Long
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

Public Sub ETR_InstallDashboard()
    Dim ws As Worksheet, cfg As Settings, regions As Object, orders As Object
    On Error GoTo Failed
    LoadConfiguration cfg, regions, orders
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    AddControls ws
    If CStr(ws.Range("B12").Value2) = "n.a." Then
        ws.Range("B5").Value2 = "No source loaded. Configure your regions, then use Refresh eTR file."
    End If
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    MsgBox "Refresh controls are ready. Configure your transit mappings, then select an eTR file.", vbInformation, "eTR dashboard"
    Exit Sub
Failed:
    MsgBox "Setup did not finish: " & Err.Description, vbExclamation, "eTR dashboard"
End Sub

Public Function ETR_LastError() As String
    ETR_LastError = mLastError
End Function

' Callable without dialogs for integration/testing. Does not save the dashboard file.
' Every failure before the final sheet swap leaves the prior Dashboard intact.
Public Function ETR_RefreshFile(ByVal sourcePath As String, Optional ByVal showMessages As Boolean = True) As Boolean
    Dim cfg As Settings, metrics As RunMetrics
    Dim regionByTransit As Object, regionOrder As Object
    Dim gross As Object, distinct As Object
    Dim sourceBook As Workbook, sourceSheet As Worksheet, staged As Worksheet
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
    ValidateSourcePath sourcePath
    LoadConfiguration cfg, regionByTransit, regionOrder
    If ThisWorkbook.ProtectStructure Then Fail "Unprotect the dashboard workbook structure before refreshing."
    If ThisWorkbook.ReadOnly Then Fail "The dashboard is read-only. Save an editable copy before refreshing."
    ' Check the existing Dashboard before opening any source.
    Set staged = ThisWorkbook.Worksheets("Dashboard")
    Set staged = Nothing

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
    Set gross = NewDictionary(vbBinaryCompare)
    Set distinct = NewDictionary(vbBinaryCompare)
    AggregateSource sourceSheet, headerCols, cfg, regionByTransit, gross, distinct, metrics
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
    CommitDashboard staged
    Set staged = Nothing
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
    If stateSaved Then
        Application.AutomationSecurity = oldSecurity
        Application.Calculation = oldCalc
        Application.EnableEvents = oldEvents
        Application.DisplayAlerts = oldAlerts
        Application.ScreenUpdating = oldScreen
        Application.StatusBar = oldStatus
    End If
    mBusy = False
    If ETR_RefreshFile Then
        ThisWorkbook.Worksheets("Dashboard").Activate
        ActiveWindow.DisplayGridlines = False
        If showMessages Then MsgBox "Refresh complete." & vbCrLf & _
            Format$(metrics.GrossRows, "#,##0") & " gross rows; " & _
            Format$(metrics.UniquePairs, "#,##0") & " unique transit + TRid pairs." & vbCrLf & _
            "Save the dashboard to retain this snapshot.", vbInformation, "eTR dashboard"
    ElseIf showMessages Then
        MsgBox "Refresh stopped. The previous dashboard is unchanged." & vbCrLf & vbCrLf & _
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
    If GetTable Is Nothing Then Fail "Configuration table is missing: " & tableName & ". Use the supplied template."
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

Private Sub LoadConfiguration(ByRef cfg As Settings, ByRef regionByTransit As Object, ByRef regionOrder As Object)
    Dim ws As Worksheet, table As ListObject, items As Object, wanted As Object
    Dim values As Variant, key As String, value As Variant, i As Long
    Dim keyCol As Long, valCol As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Configuration")
    On Error GoTo 0
    If ws Is Nothing Then Fail "Configuration sheet is missing. Use the supplied template."
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
    cfg.AgeHeader = CleanText(RequiredSetting(items, "AgeHeader"), "AgeHeader")
    cfg.DescriptionFilter = CleanText(RequiredSetting(items, "DescriptionFilter", True), "DescriptionFilter")
    cfg.FilterMode = LCase$(CleanText(RequiredSetting(items, "FilterMode"), "FilterMode"))
    cfg.SourceSheet = CleanText(RequiredSetting(items, "SourceSheet", True), "SourceSheet")
    cfg.UnmappedPolicy = LCase$(CleanText(RequiredSetting(items, "UnmappedPolicy"), "UnmappedPolicy"))
    cfg.TransitKeyMode = LCase$(CleanText(RequiredSetting(items, "TransitKeyMode"), "TransitKeyMode"))
    cfg.MaxDataRows = WholeNumber(RequiredSetting(items, "MaxDataRows"), 1, 1048574, "MaxDataRows")
    If cfg.FilterMode <> "exact" And cfg.FilterMode <> "contains" And cfg.FilterMode <> "all" Then Fail "FilterMode must be Exact, Contains or All."
    If cfg.FilterMode <> "all" And Len(cfg.DescriptionFilter) = 0 Then Fail "DescriptionFilter is required unless FilterMode is All."
    If cfg.UnmappedPolicy <> "include" And cfg.UnmappedPolicy <> "error" Then Fail "UnmappedPolicy must be Include or Error."
    If cfg.TransitKeyMode <> "numeric" And cfg.TransitKeyMode <> "text" Then Fail "TransitKeyMode must be Numeric or Text."
    Set wanted = NewDictionary(vbBinaryCompare)
    For Each value In Array(cfg.DescriptionHeader, cfg.TransitHeader, cfg.TRidHeader, cfg.AgeHeader)
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
    Dim ext As String, book As Workbook
    If Len(Trim$(sourcePath)) = 0 Then Fail "No source file was selected."
    ext = LCase$(Mid$(sourcePath, InStrRev(sourcePath, ".") + 1))
    If ext <> "xlsx" And ext <> "xlsm" And ext <> "xlsb" Then Fail "Select an .xlsx, .xlsm or .xlsb workbook. Convert legacy .xls or CSV files to .xlsx first."
    If StrComp(sourcePath, ThisWorkbook.FullName, vbTextCompare) = 0 Then Fail "Select a source workbook, not this dashboard."
    For Each book In Application.Workbooks
        If StrComp(sourcePath, book.FullName, vbTextCompare) = 0 Then Fail "Close the selected source workbook first, saving your changes, then refresh."
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
                    cfg.DescriptionHeader & ", " & cfg.TransitHeader & ", " & cfg.TRidHeader & ", " & cfg.AgeHeader & "."
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
    wanted = Array(cfg.DescriptionHeader, cfg.TransitHeader, cfg.TRidHeader, cfg.AgeHeader)
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

Private Function BucketForAge(ByVal age As Long, ByRef cfg As Settings) As Long
    Dim i As Long
    For i = 1 To cfg.BucketCount
        If age >= cfg.Buckets(i).MinAge And age <= cfg.Buckets(i).MaxAge Then
            BucketForAge = i
            Exit Function
        End If
    Next i
    Fail "AgedDays " & CStr(age) & " does not match an age bucket."
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
    Dim ageById As Object, pairSeen As Object, unmapped As Object
    Dim lastRow As Long, endRow As Long, startRow As Long, i As Long, j As Long, r As Long
    Dim descriptions As Variant, transits As Variant, ids As Variant, ages As Variant
    Dim context As String, description As String, transit As String, id As String, pair As String
    Dim age As Long, bucket As Long, firstSeen As Variant, key As Variant, counts As Variant
    Set ageById = NewDictionary(vbBinaryCompare)
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
        ages = ReadColumn(ws, startRow, endRow, columns(4))
        For i = 1 To endRow - startRow + 1
            r = startRow + i - 1
            context = "Source row " & CStr(r)
            If IsBlankValue(descriptions(i, 1)) And IsBlankValue(transits(i, 1)) And _
                    IsBlankValue(ids(i, 1)) And IsBlankValue(ages(i, 1)) Then
                metrics.BlankRows = metrics.BlankRows + 1
            Else
                metrics.ReadRows = metrics.ReadRows + 1
                description = CleanText(descriptions(i, 1), context & " Description")
                If DescriptionMatches(description, cfg) Then
                    transit = TransitKey(transits(i, 1), cfg.TransitKeyMode, context)
                    id = TransactionKey(ids(i, 1), context)
                    age = WholeNumber(ages(i, 1), 0, MAX_AGE, context & " AgedDays")
                    bucket = BucketForAge(age, cfg)
                    If ageById.Exists(id) Then
                        firstSeen = ageById(id)
                        If CLng(firstSeen(0)) <> age Then Fail context & ": TRid '" & id & _
                            "' has AgedDays " & CStr(age) & "; row " & CStr(firstSeen(1)) & _
                            " has " & CStr(firstSeen(0)) & ". One TRid must have one age, even across transits."
                    Else
                        ageById.Add id, Array(age, r)
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
    metrics.GlobalIds = ageById.Count
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
        ByVal gross As Object, ByVal distinct As Object, ByRef metrics As RunMetrics)
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
    PutMerged ws, 5, 2, viewLastCol, "Last refresh succeeded. Configuration changes apply on the next refresh."
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
    grossTotal = RenderHeatmap(ws, 16, "01  Gross eTR rows", keys, gross, mapping, cfg, metrics.MaximumCell)
    uniqueStart = grossTotal + 4
    uniqueTotal = RenderHeatmap(ws, uniqueStart, "02  Unique eTR by transit", keys, distinct, mapping, cfg, metrics.MaximumCell)
    ws.Range("B12").Formula = "=" & ws.Cells(grossTotal, lastCol).Address(False, False)
    ws.Range("D12").Formula = "=" & ws.Cells(uniqueTotal, lastCol).Address(False, False)
    ws.Range("F12").Value2 = metrics.GlobalIds
    ws.Range("H12").Formula = "=B12-D12"
    ws.Range("H13").NumberFormat = "General"
    ws.Range("H13").Formula = "=IF(B12=0,""Rate: n.a."",""Rate: ""&TEXT(H12/B12,""0.0%""))"
    PutMerged ws, uniqueTotal + 3, 2, viewLastCol, "Unique totals count transit + TRid pairs. The same TRid in different transits contributes once in each."
    ws.Range(ws.Cells(uniqueTotal + 3, 2), ws.Cells(uniqueTotal + 3, viewLastCol)).WrapText = True
    ws.Rows(uniqueTotal + 3).RowHeight = 32
    AddControls ws
    ' Page setup is optional; printer availability cannot invalidate correct results.
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape: .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .PrintArea = ws.Range(ws.Cells(2, 2), ws.Cells(uniqueTotal + 3, viewLastCol)).Address
    End With
    On Error GoTo 0
End Sub

Private Function RenderHeatmap(ByVal ws As Worksheet, ByVal topRow As Long, ByVal title As String, ByVal keys As Variant, _
        ByVal counts As Object, ByVal mapping As Object, ByRef cfg As Settings, ByVal maximum As Long) As Long
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
    PutMerged ws, topRow + 1, 2, titleLastCol, "AgedDays (inclusive). Shared heat-color scale: 0 to " & CStr(maximum) & ". Region rows are subtotals."
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
        PutMerged ws, r, 2, lastCol, "No qualifying records for this filter."
        r = r + 1
    End If
    ws.Cells(r, 2).Value2 = "Grand total"
    For j = 4 To lastCol - 1
        If counts.Count = 0 Then
            ws.Cells(r, j).Value2 = 0
        Else
            ' Only transit rows have a value in column C. Never sum region totals twice.
            ws.Cells(r, j).Formula = "=SUMIF(" & ws.Range(ws.Cells(firstData, 3), ws.Cells(r - 1, 3)).Address & _
                ",""<>""," & ws.Range(ws.Cells(firstData, j), ws.Cells(r - 1, j)).Address & ")"
        End If
    Next j
    ws.Cells(r, lastCol).FormulaR1C1 = "=SUM(RC4:RC[-1])"
    ws.Range(ws.Cells(firstData, 4), ws.Cells(r, lastCol)).NumberFormat = "#,##0"
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, lastCol))
        .Interior.Color = RGB(24, 52, 73): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 28
    End With
    RenderHeatmap = r
End Function

Private Sub ApplyHeatColors(ByVal cells As Range, ByVal maximum As Long)
    Dim scale As ColorScale, high As Double
    high = maximum: If high < 1 Then high = 1
    Set scale = cells.FormatConditions.AddColorScale(3)
    With scale.ColorScaleCriteria(1)
        .Type = xlConditionValueNumber: .Value = 0: .FormatColor.Color = RGB(245, 248, 252)
    End With
    With scale.ColorScaleCriteria(2)
        .Type = xlConditionValueNumber: .Value = high / 2: .FormatColor.Color = RGB(249, 211, 144)
    End With
    With scale.ColorScaleCriteria(3)
        .Type = xlConditionValueNumber: .Value = high: .FormatColor.Color = RGB(240, 144, 117)
    End With
End Sub

Private Sub AddControls(ByVal ws As Worksheet)
    Dim button As Shape, area As Range
    On Error Resume Next
    ws.Shapes("eTR_Refresh").Delete
    ws.Shapes("eTR_Configure").Delete
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

Private Sub CommitDashboard(ByVal staged As Worksheet)
    Dim current As Worksheet, backupName As String, stagedName As String
    Dim renamedOld As Boolean, renamedNew As Boolean, oldAlerts As Boolean
    Dim number As Long, message As String
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Rollback
    Set current = ThisWorkbook.Worksheets("Dashboard")
    backupName = UnusedSheetName("_eTR_previous")
    stagedName = staged.Name
    staged.Visible = xlSheetVisible
    staged.Move Before:=current
    current.Name = backupName
    renamedOld = True
    staged.Name = "Dashboard"
    renamedNew = True
    Application.DisplayAlerts = False
    current.Delete ' Final commit point; no fallible work follows.
    On Error Resume Next
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    Exit Sub
Rollback:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = oldAlerts
    If renamedNew Then staged.Name = stagedName
    If renamedOld Then current.Name = "Dashboard"
    current.Visible = xlSheetVisible
    On Error GoTo 0
    Err.Raise number, "eTR dashboard commit", message
End Sub
