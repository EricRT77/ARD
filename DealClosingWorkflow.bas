Attribute VB_Name = "DealClosingWorkflow"
Option Explicit

' Recreates the supplied one-page current / AI-assisted workflow as editable
' Excel shapes, text boxes and routed vector arrows. No external files needed.
' Import this .bas using Alt+F11 > File > Import File, then run
' CreateDealClosingWorkflow from Alt+F8. Save as .xlsm to keep the macro.
' Original source: Deal_Closing_Process_Map(4) - Current vs AI-Assisted.pdf, page 1.
' Updated from: New Microsoft Excel Worksheet.xlsm (saved user edits).
' All operational text is reference content from the diagram, not executable
' business logic. This macro draws the workflow; it does not perform its tasks.
' Routed arrows preserve the source positions; they do not auto-reroute when
' a node is moved. Use Shape Format > Selection Pane / Edit Points to adjust.
'
' Increase DRAW_SCALE to enlarge every shape and its text together.
Private Const DRAW_SCALE As Double = 1#
Private Const CANVAS_WIDTH As Double = 2790#
Private Const CANVAS_HEIGHT As Double = 2500#

Public Sub CreateDealClosingWorkflow()
    Dim wb As Workbook, ws As Worksheet
    Dim previousScreen As Boolean, previousEvents As Boolean
    Dim previousStatus As Variant, message As String
    On Error GoTo Failed
    previousScreen = Application.ScreenUpdating
    previousEvents = Application.EnableEvents
    previousStatus = Application.StatusBar
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Set wb = Workbooks.Add(xlWBATWorksheet)
    If wb.ProtectStructure Then Err.Raise vbObjectError + 2201, , _
        "The workbook structure is protected. Open a blank unprotected workbook."
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Creating the deal closing workflow..."
    Set ws = BlankCanvas(wb)
    ws.Activate
    ws.Range("A1:ED250").ColumnWidth = 4.5
    ws.Range("A1:ED250").RowHeight = 15
    ws.Tab.Color = RGB(18, 63, 107)
    DrawPart01 ws
    DrawPart02 ws
    DrawPart03 ws
    DrawPart04 ws
    DrawPart05 ws
    DrawPart06 ws
    DrawPart07 ws
    DrawPart08 ws
    DrawPart09 ws
    DrawPart10 ws
    DrawPart11 ws
    DrawPart12 ws
    DrawPart13 ws
    DrawPart14 ws
    DrawPart15 ws
    DrawPart16 ws
    DrawPart17 ws
    DrawPart18 ws
    DrawPart19 ws
    DrawPart20 ws
    DrawPart21 ws
    DrawPart22 ws
    DrawPart23 ws
    DrawPart24 ws
    DrawPart25 ws
    DrawPart26 ws
    DrawPart27 ws
    DrawPart28 ws
    DrawPart29 ws
    DrawPart30 ws
    DrawPart31 ws
    ApplySavedWorkbookChanges ws
    ConfigureView ws
    Application.StatusBar = previousStatus
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    MsgBox "The complete workflow is ready on '" & ws.Name & "'." & vbCrLf & _
        "Use Excel's zoom control to inspect the detail." & vbCrLf & _
        "Save as .xlsm to retain this macro.", vbInformation, "Workflow created"
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    Application.StatusBar = previousStatus
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    On Error GoTo 0
    MsgBox "The workflow could not be completed: " & message, vbExclamation, "Workflow"
End Sub

Private Function BlankCanvas(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet, candidate As String, suffix As Long
    candidate = "Deal Closing Workflow"
    suffix = 1
    Do While SheetExists(wb, candidate)
        suffix = suffix + 1
        candidate = "Deal Closing Workflow " & CStr(suffix)
    Loop
    ' Reuse only an entirely empty, unprotected active worksheet.
    If TypeOf wb.ActiveSheet Is Worksheet Then
        Set ws = wb.ActiveSheet
        If ws.ProtectContents Or ws.Shapes.Count > 0 Then Set ws = Nothing
        If Not ws Is Nothing Then
            If Application.CountA(ws.UsedRange) > 0 Then Set ws = Nothing
        End If
    End If
    If ws Is Nothing Then Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = candidate
    Set BlankCanvas = ws
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sh As Object
    For Each sh In wb.Sheets
        If StrComp(sh.Name, sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next sh
End Function

Private Sub ConfigureView(ByVal ws As Worksheet)
    Dim lastRow As Long, lastColumn As Long
    lastRow = 1
    Do While ws.Cells(lastRow, 1).Top < CANVAS_HEIGHT * DRAW_SCALE
        lastRow = lastRow + 1
    Loop
    lastColumn = 1
    Do While ws.Cells(1, lastColumn).Left < CANVAS_WIDTH * DRAW_SCALE
        lastColumn = lastColumn + 1
    Loop
    ' Printer-dependent settings must not prevent creation of the diagram.
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn)).Address
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = 18
        .RightMargin = 18
        .TopMargin = 18
        .BottomMargin = 18
        .CenterHorizontally = True
        .PrintGridlines = False
    End With
    On Error GoTo 0
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.Zoom = 50
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End If
    ws.Range("A1").Select
End Sub

Private Sub AddBox(ByVal ws As Worksheet, ByVal name As String, ByVal kind As Long, _
                   ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                   ByVal fillColor As Long, ByVal lineColor As Long, ByVal lineWeight As Double)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(kind, x * DRAW_SCALE, y * DRAW_SCALE, w * DRAW_SCALE, h * DRAW_SCALE)
    s.Name = name
    s.Placement = xlFreeFloating
    s.Shadow.Visible = msoFalse
    If kind = msoShapeRoundedRectangle Then s.Adjustments.Item(1) = 0.12
    SetStyle s, fillColor, lineColor, lineWeight
End Sub

Private Sub SetStyle(ByVal s As Shape, ByVal fillColor As Long, ByVal lineColor As Long, ByVal weight As Double)
    If fillColor < 0 Then
        s.Fill.Visible = msoFalse
    Else
        s.Fill.Visible = msoTrue
        s.Fill.Solid
        s.Fill.ForeColor.RGB = fillColor
        s.Fill.Transparency = 0
    End If
    If lineColor < 0 Then
        s.Line.Visible = msoFalse
    Else
        s.Line.Visible = msoTrue
        s.Line.ForeColor.RGB = lineColor
        s.Line.Weight = weight * DRAW_SCALE
        s.Line.Transparency = 0
    End If
End Sub

Private Sub AddPath(ByVal ws As Worksheet, ByVal name As String, ByVal data As String, _
                    ByVal fillColor As Long, ByVal lineColor As Long, ByVal weight As Double, ByVal dotted As Boolean)
    Dim fb As FreeformBuilder, s As Shape, segment As Variant, values As Variant
    Dim firstX As Double, firstY As Double
    For Each segment In Split(data, ";")
        values = Split(CStr(segment), ",")
        Select Case values(0)
            Case "M"
                firstX = Val(values(1)) * DRAW_SCALE
                firstY = Val(values(2)) * DRAW_SCALE
                Set fb = ws.Shapes.BuildFreeform(msoEditingCorner, firstX, firstY)
            Case "L"
                fb.AddNodes msoSegmentLine, msoEditingAuto, Val(values(1)) * DRAW_SCALE, Val(values(2)) * DRAW_SCALE
            Case "C"
                fb.AddNodes msoSegmentCurve, msoEditingCorner, _
                    Val(values(1)) * DRAW_SCALE, Val(values(2)) * DRAW_SCALE, _
                    Val(values(3)) * DRAW_SCALE, Val(values(4)) * DRAW_SCALE, _
                    Val(values(5)) * DRAW_SCALE, Val(values(6)) * DRAW_SCALE
            Case "H"
                fb.AddNodes msoSegmentLine, msoEditingAuto, firstX, firstY
        End Select
    Next segment
    Set s = fb.ConvertToShape
    s.Name = name
    s.Placement = xlFreeFloating
    s.Shadow.Visible = msoFalse
    SetStyle s, fillColor, lineColor, weight
    If dotted And lineColor >= 0 Then s.Line.DashStyle = msoLineRoundDot
End Sub

Private Sub AddLabel(ByVal ws As Worksheet, ByVal name As String, _
                     ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                     ByVal caption As String, ByVal size As Double, ByVal ink As Long, _
                     ByVal bold As Boolean, ByVal vertical As Boolean, ByVal centered As Boolean, _
                     ByVal spacing As Double, ByVal backColor As Long)
    Dim s As Shape, direction As Long
    direction = msoTextOrientationHorizontal
    If vertical Then direction = msoTextOrientationUpward
    Set s = ws.Shapes.AddTextbox(direction, x * DRAW_SCALE, y * DRAW_SCALE, w * DRAW_SCALE, h * DRAW_SCALE)
    s.Name = name
    s.Placement = xlFreeFloating
    s.Shadow.Visible = msoFalse
    SetStyle s, backColor, -1, 0
    With s.TextFrame2
        .AutoSize = msoAutoSizeNone
        .WordWrap = msoFalse
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorTop
        If vertical Then .VerticalAnchor = msoAnchorMiddle
        With .TextRange
            .Text = caption
            .Font.Name = "Arial"
            .Font.Size = size * DRAW_SCALE
            .Font.Bold = IIf(bold, msoTrue, msoFalse)
            .Font.Fill.ForeColor.RGB = ink
            With .ParagraphFormat
                .Alignment = IIf(centered, msoAlignCenter, msoAlignLeft)
                .SpaceBefore = 0
                .SpaceAfter = 0
                .LineRuleWithin = msoFalse
                .SpaceWithin = spacing * DRAW_SCALE
            End With
        End With
    End With
    s.AlternativeText = caption
End Sub

Private Sub GroupNode(ByVal ws As Worksheet, ByVal name As String, ByVal names As String)
    Dim a As Variant, members() As Variant, i As Long, s As Shape
    a = Split(names, "|")
    ReDim members(LBound(a) To UBound(a))
    For i = LBound(a) To UBound(a)
        members(i) = CStr(a(i))
    Next i
    Set s = ws.Shapes.Range(members).Group
    s.Name = name
    s.Placement = xlFreeFloating
End Sub

Private Sub DrawPart01(ByVal ws As Worksheet)
    AddBox ws, "Box_0", msoShapeRectangle, 14.998, 15.147, 2759.85, 39.015, 7028498, -1, 0
    AddBox ws, "Box_2", msoShapeRectangle, 14.998, 54.135, 2759.85, 24.75, 16117994, -1, 0
    AddBox ws, "Box_4", msoShapeRectangle, 14.998, 78.885, 2759.85, 390.015, 16777215, 8615784, 0.977
    AddBox ws, "Box_5", msoShapeRectangle, 14.998, 78.885, 48.735, 390.015, 16048875, 8615784, 0.977
    AddBox ws, "Box_9", msoShapeRectangle, 14.998, 468.9, 2759.85, 517.5, 16777215, 8615784, 0.977
    AddBox ws, "Box_10", msoShapeRectangle, 14.998, 468.9, 48.735, 517.5, 15397085, 8615784, 0.977
    AddBox ws, "Box_14", msoShapeRectangle, 14.998, 986.4, 2759.85, 157.5, 16777215, 8615784, 0.977
    AddBox ws, "Box_15", msoShapeRectangle, 14.998, 986.4, 48.735, 157.5, 16184561, 8615784, 0.977
    AddBox ws, "Box_18", msoShapeRectangle, 14.998, 1143.9, 2759.85, 532.35, 16777215, 8615784, 0.977
    AddBox ws, "Box_19", msoShapeRectangle, 14.998, 1143.9, 48.735, 532.35, 16184561, 8615784, 0.977
    AddBox ws, "Box_23", msoShapeRectangle, 14.998, 1676.25, 2759.85, 360, 16777215, 8615784, 0.977
    AddBox ws, "Box_24", msoShapeRectangle, 14.998, 1676.25, 48.735, 360, 16184561, 8615784, 0.977
    AddBox ws, "Box_27", msoShapeRectangle, 14.998, 2036.25, 2759.85, 194.985, 16777215, 8615784, 0.977
    AddBox ws, "Box_28", msoShapeRectangle, 14.998, 2036.25, 48.735, 194.985, 16184561, 8615784, 0.977
    AddBox ws, "Box_33", msoShapeRectangle, 14.998, 2231.1, 2759.85, 239.985, 16777215, 8615784, 0.977
    AddBox ws, "Box_34", msoShapeRectangle, 14.998, 2231.1, 48.735, 239.985, 16184561, 8615784, 0.977
    AddBox ws, "Box_38", msoShapeFlowchartOffpageConnector, 69.345, 1175.4, 71.235, 63.405, 16777215, 4750519, 1.202
    AddBox ws, "Box_41", msoShapeFlowchartTerminator, 166.905, 1190.7, 86.13, 33.75, 16777215, 8615784, 1.202
    AddBox ws, "Box_43", msoShapeRoundedRectangle, 121.86, 1306.8, 176.265, 71.235, 16777215, 8615784, 1.202
    AddBox ws, "Box_47", msoShapeDiamond, 354.855, 1288.257, 145.275, 108.048, 16777215, 8615784, 1.202
    AddBox ws, "Box_50", msoShapeRoundedRectangle, 363.735, 1170, 127.485, 75.015, 16777215, 8615784, 1.202
    AddBox ws, "Box_55", msoShapeRoundedRectangle, 585, 1303.2, 135, 78.75, 16577527, 10373232, 1.949
    AddBox ws, "Box_60", msoShapeDiamond, 774.48, 1288.257, 145.275, 108.048, 16777215, 8615784, 1.202
    AddBox ws, "Box_62", msoShapeRoundedRectangle, 774, 1170, 146.25, 75.015, 16777215, 8615784, 1.202
    AddBox ws, "Box_67", msoShapeFlowchartDelay, 783.45, 1023.75, 127.485, 82.485, 8630769, 4750519, 1.202
    AddBox ws, "Box_70", msoShapeDiamond, 976.95, 1288.26, 130.32, 108, 16777215, 8615784, 1.202
    AddBox ws, "Box_72", msoShapeRoundedRectangle, 987.75, 1173.6, 108.765, 67.5, 16577527, 10373232, 1.949
    AddBox ws, "Box_75", msoShapeFlowchartOffpageConnector, 1141.65, 1175.4, 63.765, 63.135, 16577527, 10373232, 1.949
    AddBox ws, "Box_77", msoShapeDiamond, 1211.442, 1288.24, 156.513, 108.065, 16777215, 8615784, 1.202
    AddBox ws, "Box_80", msoShapeFlowchartDelay, 1226.25, 1438.2, 127.485, 86.265, 8630769, 4750519, 1.202
    AddBox ws, "Box_83", msoShapeDiamond, 1402.693, 1281.039, 179.012, 123.051, 16777215, 8615784, 1.202
    AddBox ws, "Box_87", msoShapeRoundedRectangle, 1419.3, 1170, 146.25, 75.015, 16777215, 8615784, 1.202
    AddBox ws, "Box_92", msoShapeRoundedRectangle, 1434.15, 1447.65, 116.235, 67.5, 61695, 10373232, 1.949
    AddBox ws, "Box_95", msoShapeRoundedRectangle, 1426.5, 1582.65, 131.265, 67.5, 16777215, 8615784, 1.202
    AddBox ws, "Box_99", msoShapeDiamond, 1425.179, 1738.259, 134.026, 100.531, 16777215, 8615784, 1.202
End Sub

Private Sub DrawPart02(ByVal ws As Worksheet)
    AddBox ws, "Box_101", msoShapeRoundedRectangle, 1602.9, 1575, 183.735, 82.485, 8630769, 4750519, 1.202
    AddBox ws, "Box_106", msoShapeDiamond, 1620.495, 1738.215, 148.995, 100.62, 16777215, 8615784, 1.202
    AddBox ws, "Box_108", msoShapeDiamond, 1616.445, 2298.915, 156.51, 96.885, 16777215, 8615784, 1.202
    AddBox ws, "Box_111", msoShapeFlowchartDelay, 1466.1, 2304.45, 127.485, 86.265, 8630769, 4750519, 1.202
    AddBox ws, "Box_114", msoShapeDiamond, 1831.995, 1738.215, 152.775, 100.62, 16777215, 8615784, 1.202
    AddBox ws, "Box_117", msoShapeDiamond, 2045.733, 1738.253, 137.802, 100.537, 16777215, 8615784, 1.202
    AddBox ws, "Box_119", msoShapeDiamond, 2045.733, 1899.803, 137.802, 100.537, 16777215, 8615784, 1.202
    AddBox ws, "Box_121", msoShapeRectangle, 2202.75, 2096.1, 146.25, 82.485, 16777215, 8615784, 1.202
    AddBox ws, "Box_124", msoShapeRoundedRectangle, 2049.3, 2313.9, 131.265, 67.5, 16777215, 8615784, 1.202
    AddBox ws, "Box_126", msoShapeRoundedRectangle, 2347.2, 1755, 120.015, 67.5, 16777215, 8615784, 1.202
    AddBox ws, "Box_128", msoShapeFlowchartDelay, 2343.6, 1573.2, 127.485, 86.265, 8630769, 4750519, 1.202
    AddBox ws, "Box_130", msoShapeRoundedRectangle, 2335.95, 1441.8, 142.515, 78.75, 16777215, 8615784, 1.202
    AddBox ws, "Box_135", msoShapeRoundedRectangle, 2520, 1580.4, 135, 71.235, 16777215, 8615784, 1.202
    AddBox ws, "Box_139", msoShapeFlowchartOffpageConnector, 2682.9, 1584.45, 63.765, 63.135, 16577527, 10373232, 1.949
    AddBox ws, "Box_141", msoShapeRectangle, 110.61, 1443.6, 198.765, 75.015, 8342551, 8342551, 1.202
    AddBox ws, "Box_145", msoShapeRectangle, 747.9, 1438.2, 198.765, 116.235, 8342551, 8342551, 1.202
    AddBox ws, "Box_152", msoShapeRoundedRectangle, 1633.05, 1421.1, 288.765, 120.015, 8342551, 8342551, 1.202
    AddBox ws, "Box_159", msoShapeRoundedRectangle, 2525.4, 1434.15, 221.265, 93.735, 8342551, 8342551, 1.202
    AddBox ws, "Box_164", msoShapeRoundedRectangle, 2250, 2309.85, 239.985, 75.015, 8342551, 8342551, 1.202
    AddBox ws, "Box_168", msoShapeRectangle, 104.985, 558.9, 481.5, 99, 16185839, 7503624, 1.202
    AddBox ws, "Box_172", msoShapeRectangle, 637.2, 558.9, 481.5, 99, 16185839, 7503624, 1.202
    AddBox ws, "Box_176", msoShapeRectangle, 1170, 558.9, 481.5, 99, 16185839, 7503624, 1.202
    AddBox ws, "Box_180", msoShapeRectangle, 1702.35, 558.9, 481.5, 99, 16185839, 7503624, 1.202
    AddBox ws, "Box_184", msoShapeRectangle, 2234.7, 558.9, 481.5, 99, 16185839, 7503624, 1.202
    AddBox ws, "Box_188", msoShapeRectangle, 90, 724.5, 511.2, 180, 16185839, 12766619, 1.202
    AddBox ws, "Box_189", msoShapeRectangle, 622.35, 724.5, 511.2, 180, 16185839, 12766619, 1.202
    AddBox ws, "Box_190", msoShapeRectangle, 1154.7, 724.5, 511.2, 180, 16185839, 12766619, 1.202
    AddBox ws, "Box_191", msoShapeRectangle, 1687.5, 724.5, 511.2, 180, 16185839, 12766619, 1.202
    AddBox ws, "Box_192", msoShapeRectangle, 2219.85, 724.5, 511.2, 180, 16185839, 12766619, 1.202
    AddBox ws, "Box_193", msoShapeRoundedRectangle, 676.35, 1274.4, 49.5, 22.5, 10373232, 10373232, 1.202
    AddBox ws, "Box_195", msoShapeRoundedRectangle, 1053, 1145.25, 49.5, 22.5, 10373232, 10373232, 1.202
    AddBox ws, "Box_197", msoShapeRoundedRectangle, 1161.9, 1147.05, 49.5, 22.5, 10373232, 10373232, 1.202
    AddBox ws, "Box_199", msoShapeRoundedRectangle, 1507.05, 1418.85, 49.5, 22.5, 10373232, 10373232, 1.202
    AddBox ws, "Box_201", msoShapeRoundedRectangle, 2703.15, 1555.65, 49.5, 22.5, 10373232, 10373232, 1.202
    AddBox ws, "Box_207", msoShapeRectangle, 90, 171.135, 511.2, 205.515, 16577527, 14200515, 1.048
End Sub

Private Sub DrawPart03(ByVal ws As Worksheet)
    AddBox ws, "Box_208", msoShapeRectangle, 90, 171.135, 511.2, 59.985, 10373232, 10373232, 1.048
    AddBox ws, "Box_216", msoShapeRectangle, 622.35, 171.135, 511.2, 205.515, 16577527, 14200515, 1.048
    AddBox ws, "Box_217", msoShapeRectangle, 622.35, 171.135, 511.2, 59.985, 10373232, 10373232, 1.048
    AddBox ws, "Box_224", msoShapeRectangle, 1154.7, 171.135, 511.2, 205.515, 16577527, 14200515, 1.048
    AddBox ws, "Box_225", msoShapeRectangle, 1154.7, 171.135, 511.2, 59.985, 10373232, 10373232, 1.048
    AddBox ws, "Box_232", msoShapeRectangle, 1687.5, 171.135, 511.2, 205.515, 16577527, 14200515, 1.048
    AddBox ws, "Box_233", msoShapeRectangle, 1687.5, 171.135, 511.2, 59.985, 10373232, 10373232, 1.048
    AddBox ws, "Box_240", msoShapeRectangle, 2219.85, 171.135, 511.2, 205.515, 16577527, 14200515, 1.048
    AddBox ws, "Box_241", msoShapeRectangle, 2219.85, 171.135, 511.2, 59.985, 10373232, 10373232, 1.048
    AddBox ws, "Box_258", msoShapeRectangle, 90, 724.5, 511.2, 56.25, 7503624, 7503624, 1.048
    AddBox ws, "Box_267", msoShapeRectangle, 622.35, 724.5, 511.2, 56.25, 7503624, 7503624, 1.048
    AddBox ws, "Box_275", msoShapeRectangle, 1154.7, 724.5, 511.2, 56.25, 7503624, 7503624, 1.048
    AddBox ws, "Box_283", msoShapeRectangle, 1687.5, 724.5, 511.2, 56.25, 7503624, 7503624, 1.048
    AddBox ws, "Box_291", msoShapeRectangle, 2219.85, 724.5, 511.2, 56.25, 7503624, 7503624, 1.048
    AddPath ws, "Route_303_001", "M,141.885,1208.25;L,141.21,1208.25;L,141.21,1206.9;L,141.885,1206.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_303_002", "M,141.885,1208.25;L,141.21,1208.25;L,141.21,1206.9;L,141.885,1206.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_304_003", "M,210.645,1225.8;L,209.385,1225.8;L,209.385,1224.9;L,210.645,1224.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_304_004", "M,210.645,1225.8;L,209.385,1225.8;L,209.385,1224.9;L,210.645,1224.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_305_005", "M,299.385,1343.25;L,298.71,1343.25;L,298.71,1341.9;L,299.385,1341.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_305_006", "M,299.385,1343.25;L,298.71,1343.25;L,298.71,1341.9;L,299.385,1341.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_306_007", "M,428.13,1287;L,426.87,1287;L,426.87,1270.35;L,428.13,1270.35;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_306_008", "M,428.13,1256.85;L,426.87,1256.85;L,426.87,1253.25;L,428.13,1253.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_306_009", "M,428.13,1287;L,426.87,1287;L,426.87,1270.35;L,428.13,1270.35;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_306_010", "M,428.13,1256.85;L,426.87,1256.85;L,426.87,1253.25;L,428.13,1253.25;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_308_011", "M,363.15,1208.25;L,362.475,1208.25;L,362.475,1206.9;L,363.15,1206.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_308_012", "M,363.15,1208.25;L,362.475,1208.25;L,362.475,1206.9;L,363.15,1206.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_309_013", "M,535.05,1030.5;L,532.8,1030.5;L,532.8,669.6;L,535.05,669.6;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_309_014", "M,535.05,1030.5;L,532.8,1030.5;L,532.8,669.6;L,535.05,669.6;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_311_015", "M,721.35,1343.25;L,720.9,1343.25;L,720.9,1341.9;L,721.35,1341.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_311_016", "M,721.35,1343.25;L,720.9,1343.25;L,720.9,1341.9;L,721.35,1341.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_312_017", "M,847.8,1287;L,846.45,1287;L,846.45,1270.35;L,847.8,1270.35;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_312_018", "M,847.8,1256.85;L,846.45,1256.85;L,846.45,1253.25;L,847.8,1253.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_312_019", "M,847.8,1287;L,846.45,1287;L,846.45,1270.35;L,847.8,1270.35;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_312_020", "M,847.8,1256.85;L,846.45,1256.85;L,846.45,1253.25;L,847.8,1253.25;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_314_021", "M,847.8,1169.55;L,846.45,1169.55;L,846.45,1168.65;L,847.8,1168.65;H", _
        8739893, -1, 0, False
End Sub

Private Sub DrawPart04(ByVal ws As Worksheet)
    AddPath ws, "Route_314_022", "M,847.8,1169.55;L,846.45,1169.55;L,846.45,1168.65;L,847.8,1168.65;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_315_023", "M,783,1064.25;L,783,1065.6;L,782.1,1065.6;L,782.1,1064.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_315_024", "M,783,1064.25;L,783,1065.6;L,782.1,1065.6;L,782.1,1064.25;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_316_025", "M,937.35,1343.25;L,921.15,1343.25;L,921.15,1341.9;L,937.35,1341.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_316_026", "M,968.4,1343.25;L,951.75,1343.25;L,951.75,1341.9;L,968.4,1341.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_316_027", "M,937.35,1343.25;L,921.15,1343.25;L,921.15,1341.9;L,937.35,1341.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_316_028", "M,968.4,1343.25;L,951.75,1343.25;L,951.75,1341.9;L,968.4,1341.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_318_029", "M,1043.1,1287;L,1041.75,1287;L,1041.75,1270.8;L,1043.1,1270.8;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_318_030", "M,1043.1,1257.3;L,1041.75,1257.3;L,1041.75,1250.1;L,1043.1,1250.1;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_318_031", "M,1043.1,1287;L,1041.75,1287;L,1041.75,1270.8;L,1043.1,1270.8;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_318_032", "M,1043.1,1257.3;L,1041.75,1257.3;L,1041.75,1250.1;L,1043.1,1250.1;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_320_033", "M,1098.45,1208.25;L,1097.55,1208.25;L,1097.55,1206.9;L,1098.45,1206.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_320_034", "M,1098.45,1208.25;L,1097.55,1208.25;L,1097.55,1206.9;L,1098.45,1206.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_321_035", "M,1149.3,1343.25;L,1108.8,1343.25;L,1108.8,1341.9;L,1149.3,1341.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_321_036", "M,1202.85,1343.25;L,1167.75,1343.25;L,1167.75,1341.9;L,1202.85,1341.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_321_037", "M,1149.3,1343.25;L,1108.8,1343.25;L,1108.8,1341.9;L,1149.3,1341.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_321_038", "M,1202.85,1343.25;L,1167.75,1343.25;L,1167.75,1341.9;L,1202.85,1341.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_323_039", "M,1290.6,1429.65;L,1289.25,1429.65;L,1289.25,1425.6;L,1290.6,1425.6;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_323_040", "M,1290.6,1397.7;L,1290.6,1412.1;L,1289.25,1412.1;L,1289.25,1397.7;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_323_041", "M,1290.6,1429.65;L,1289.25,1429.65;L,1289.25,1425.6;L,1290.6,1425.6;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_323_042", "M,1290.6,1397.7;L,1290.6,1412.1;L,1289.25,1412.1;L,1289.25,1397.7;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_325_043", "M,1225.35,1481.85;L,1224.9,1481.85;L,1224.9,1480.5;L,1225.35,1480.5;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_325_044", "M,1225.35,1481.85;L,1224.9,1481.85;L,1224.9,1480.5;L,1225.35,1480.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_326_045", "M,1378.35,1343.25;L,1369.35,1343.25;L,1369.35,1341.9;L,1378.35,1341.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_326_046", "M,1378.35,1343.25;L,1369.35,1343.25;L,1369.35,1341.9;L,1378.35,1341.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_328_047", "M,1493.1,1279.35;L,1491.75,1279.35;L,1491.75,1270.8;L,1493.1,1270.8;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_328_048", "M,1493.1,1257.3;L,1491.75,1257.3;L,1491.75,1253.25;L,1493.1,1253.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_328_049", "M,1493.1,1279.35;L,1491.75,1279.35;L,1491.75,1270.8;L,1493.1,1270.8;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_328_050", "M,1493.1,1257.3;L,1491.75,1257.3;L,1491.75,1253.25;L,1493.1,1253.25;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_330_051", "M,1493.1,1438.65;L,1491.75,1438.65;L,1491.75,1432.8;L,1493.1,1432.8;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_330_052", "M,1493.1,1405.35;L,1493.1,1419.3;L,1491.75,1419.3;L,1491.75,1405.35;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_330_053", "M,1493.1,1438.65;L,1491.75,1438.65;L,1491.75,1432.8;L,1493.1,1432.8;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_330_054", "M,1493.1,1405.35;L,1493.1,1419.3;L,1491.75,1419.3;L,1491.75,1405.35;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_332_055", "M,1493.1,1516.5;L,1491.75,1516.5;L,1491.75,1516.05;L,1493.1,1516.05;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_332_056", "M,1493.1,1516.5;L,1491.75,1516.5;L,1491.75,1516.05;L,1493.1,1516.05;H", _
        -1, 8739893, 0.023, False
End Sub

Private Sub DrawPart05(ByVal ws As Worksheet)
    AddPath ws, "Route_333_057", "M,1493.1,1651.05;L,1491.75,1651.05;L,1491.75,1650.6;L,1493.1,1650.6;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_333_058", "M,1493.1,1651.05;L,1491.75,1651.05;L,1491.75,1650.6;L,1493.1,1650.6;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_338_059", "M,1695.6,1658.7;L,1694.25,1658.7;L,1694.25,1658.25;L,1695.6,1658.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_338_060", "M,1695.6,1658.7;L,1694.25,1658.7;L,1694.25,1658.25;L,1695.6,1658.25;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_339_061", "M,1790.55,1789.2;L,1770.75,1789.2;L,1770.75,1787.85;L,1790.55,1787.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_339_062", "M,1823.85,1789.2;L,1808.55,1789.2;L,1808.55,1787.85;L,1823.85,1787.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_339_063", "M,1790.55,1789.2;L,1770.75,1789.2;L,1770.75,1787.85;L,1790.55,1787.85;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_339_064", "M,1823.85,1789.2;L,1808.55,1789.2;L,1808.55,1787.85;L,1823.85,1787.85;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_341_065", "M,1695.6,2290.5;L,1694.25,2290.5;L,1694.25,1892.25;L,1695.6,1892.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_341_066", "M,1695.6,1840.5;L,1695.6,1878.75;L,1694.25,1878.75;L,1694.25,1840.5;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_341_067", "M,1695.6,2290.5;L,1694.25,2290.5;L,1694.25,1892.25;L,1695.6,1892.25;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_341_068", "M,1695.6,1840.5;L,1695.6,1878.75;L,1694.25,1878.75;L,1694.25,1840.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_343_069", "M,1695.6,2409.3;L,1694.25,2409.3;L,1694.25,2397.15;L,1695.6,2397.15;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_343_070", "M,1695.6,2409.3;L,1694.25,2409.3;L,1694.25,2397.15;L,1695.6,2397.15;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_346_071", "M,1803.15,2346.75;L,1803.15,2348.1;L,1774.35,2348.1;L,1774.35,2346.75;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_346_072", "M,1803.15,2346.75;L,1803.15,2348.1;L,1774.35,2348.1;L,1774.35,2346.75;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_348_073", "M,1909.35,1737;L,1908,1737;L,1908,1712.7;L,1909.35,1712.7;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_348_074", "M,1909.35,1737;L,1908,1737;L,1908,1712.7;L,1909.35,1712.7;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_350_075", "M,2007,1789.2;L,1986.3,1789.2;L,1986.3,1787.85;L,2007,1787.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_350_076", "M,2037.15,1789.2;L,2025,1789.2;L,2025,1787.85;L,2037.15,1787.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_350_077", "M,2007,1789.2;L,1986.3,1789.2;L,1986.3,1787.85;L,2007,1787.85;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_350_078", "M,2037.15,1789.2;L,2025,1789.2;L,2025,1787.85;L,2037.15,1787.85;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_352_079", "M,2227.5,1789.2;L,2185.2,1789.2;L,2185.2,1787.85;L,2227.5,1787.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_352_080", "M,2339.1,1789.2;L,2241.9,1789.2;L,2241.9,1787.85;L,2339.1,1787.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_352_081", "M,2227.5,1789.2;L,2185.2,1789.2;L,2185.2,1787.85;L,2227.5,1787.85;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_352_082", "M,2339.1,1789.2;L,2241.9,1789.2;L,2241.9,1787.85;L,2339.1,1787.85;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_354_083", "M,2115.45,1891.35;L,2114.1,1891.35;L,2114.1,1874.25;L,2115.45,1874.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_354_084", "M,2115.45,1840.5;L,2115.45,1860.75;L,2114.1,1860.75;L,2114.1,1840.5;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_354_085", "M,2115.45,1891.35;L,2114.1,1891.35;L,2114.1,1874.25;L,2115.45,1874.25;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_354_086", "M,2115.45,1840.5;L,2115.45,1860.75;L,2114.1,1860.75;L,2114.1,1840.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_356_087", "M,2213.1,1949.4;L,2213.1,1950.75;L,2185.2,1950.75;L,2185.2,1949.4;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_356_088", "M,2213.1,1949.4;L,2213.1,1950.75;L,2185.2,1950.75;L,2185.2,1949.4;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_358_089", "M,2115.45,2305.35;L,2114.1,2305.35;L,2114.1,2026.35;L,2115.45,2026.35;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_358_090", "M,2115.45,2001.6;L,2115.45,2012.85;L,2114.1,2012.85;L,2114.1,2001.6;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_358_091", "M,2115.45,2305.35;L,2114.1,2305.35;L,2114.1,2026.35;L,2115.45,2026.35;H", _
        -1, 8739893, 0.023, False
End Sub

Private Sub DrawPart06(ByVal ws As Worksheet)
    AddPath ws, "Route_358_092", "M,2115.45,2001.6;L,2115.45,2012.85;L,2114.1,2012.85;L,2114.1,2001.6;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_360_093", "M,2350.35,2137.95;L,2349.9,2137.95;L,2349.9,2136.6;L,2350.35,2136.6;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_360_094", "M,2350.35,2137.95;L,2349.9,2137.95;L,2349.9,2136.6;L,2350.35,2136.6;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_361_095", "M,2407.95,1754.55;L,2406.6,1754.55;L,2406.6,1753.65;L,2407.95,1753.65;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_361_096", "M,2407.95,1754.55;L,2406.6,1754.55;L,2406.6,1753.65;L,2407.95,1753.65;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_362_097", "M,2407.95,1571.85;L,2406.6,1571.85;L,2406.6,1555.65;L,2407.95,1555.65;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_362_098", "M,2407.95,1542.15;L,2406.6,1542.15;L,2406.6,1529.1;L,2407.95,1529.1;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_362_099", "M,2407.95,1571.85;L,2406.6,1571.85;L,2406.6,1555.65;L,2407.95,1555.65;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_362_100", "M,2407.95,1542.15;L,2406.6,1542.15;L,2406.6,1529.1;L,2407.95,1529.1;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_362_101", "M,2407.95,1572.3;L,2406.6,1572.3;L,2406.6,1571.85;L,2407.95,1571.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_364_102", "M,2335.5,1481.85;L,2334.6,1481.85;L,2334.6,1480.5;L,2335.5,1480.5;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_364_103", "M,2335.5,1481.85;L,2334.6,1481.85;L,2334.6,1480.5;L,2335.5,1480.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_365_104", "M,2511.45,1615.5;L,2511.45,1616.85;L,2506.5,1616.85;L,2506.5,1615.5;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_365_105", "M,2488.05,1615.5;L,2488.05,1616.85;L,2472.3,1616.85;L,2472.3,1615.5;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_365_106", "M,2511.45,1615.5;L,2511.45,1616.85;L,2506.5,1616.85;L,2506.5,1615.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_365_107", "M,2488.05,1615.5;L,2488.05,1616.85;L,2472.3,1616.85;L,2472.3,1615.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_367_108", "M,2655.9,1616.85;L,2655.45,1616.85;L,2655.45,1615.5;L,2655.9,1615.5;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_367_109", "M,2655.9,1616.85;L,2655.45,1616.85;L,2655.45,1615.5;L,2655.9,1615.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_368_110", "M,210.42,1379.25;L,209.61,1379.25;L,209.61,1378.8;L,210.42,1378.8;H", _
        10326919, -1, 0, False
    AddPath ws, "Route_368_111", "M,210.42,1379.25;L,209.61,1379.25;L,209.61,1378.8;L,210.42,1378.8;H", _
        -1, 10326919, 0.023, False
    AddPath ws, "Route_368_112", "M,210.42,1442.7;L,210.42,1443.15;L,209.61,1443.15;L,209.61,1442.7;H", _
        10326919, -1, 0, False
    AddPath ws, "Route_368_113", "M,210.42,1442.7;L,210.42,1443.15;L,209.61,1443.15;L,209.61,1442.7;H", _
        -1, 10326919, 0.023, False
    AddPath ws, "Route_369_114", "M,847.8,1436.85;L,847.8,1437.3;L,846.9,1437.3;L,846.9,1436.85;H", _
        10326919, -1, 0, False
    AddPath ws, "Route_369_115", "M,847.8,1436.85;L,847.8,1437.3;L,846.9,1437.3;L,846.9,1436.85;H", _
        -1, 10326919, 0.023, False
    AddPath ws, "Route_370_116", "M,1551.6,1481.4;L,1551.6,1481.4;L,1551.6,1480.95;L,1551.6,1480.95;H", _
        10326919, -1, 0, False
    AddPath ws, "Route_370_117", "M,1551.6,1481.4;L,1551.6,1481.4;L,1551.6,1480.95;L,1551.6,1480.95;H", _
        -1, 10326919, 0.023, False
    AddPath ws, "Route_370_118", "M,1632.15,1481.4;L,1631.7,1481.4;L,1631.7,1480.95;L,1632.15,1480.95;H", _
        10326919, -1, 0, False
    AddPath ws, "Route_370_119", "M,1632.15,1481.4;L,1631.7,1481.4;L,1631.7,1480.95;L,1632.15,1480.95;H", _
        -1, 10326919, 0.023, False
    AddPath ws, "Route_371_120", "M,2479.5,1481.4;L,2479.05,1481.4;L,2479.05,1480.95;L,2479.5,1480.95;H", _
        10326919, -1, 0, False
    AddPath ws, "Route_371_121", "M,2479.5,1481.4;L,2479.05,1481.4;L,2479.05,1480.95;L,2479.5,1480.95;H", _
        -1, 10326919, 0.023, False
    AddPath ws, "Route_371_122", "M,2524.95,1481.4;L,2524.5,1481.4;L,2524.5,1480.95;L,2524.95,1480.95;H", _
        10326919, -1, 0, False
    AddPath ws, "Route_371_123", "M,2524.95,1481.4;L,2524.5,1481.4;L,2524.5,1480.95;L,2524.95,1480.95;H", _
        -1, 10326919, 0.023, False
    AddPath ws, "Route_372_124", "M,2181.6,2347.65;L,2181.15,2347.65;L,2181.15,2347.2;L,2181.6,2347.2;H", _
        10326919, -1, 0, False
    AddPath ws, "Route_372_125", "M,2181.6,2347.65;L,2181.15,2347.65;L,2181.15,2347.2;L,2181.6,2347.2;H", _
        -1, 10326919, 0.023, False
    AddPath ws, "Route_372_126", "M,2249.1,2347.65;L,2248.65,2347.65;L,2248.65,2347.2;L,2249.1,2347.2;H", _
        10326919, -1, 0, False
End Sub

Private Sub DrawPart07(ByVal ws As Worksheet)
    AddPath ws, "Route_372_127", "M,2249.1,2347.65;L,2248.65,2347.65;L,2248.65,2347.2;L,2249.1,2347.2;H", _
        -1, 10326919, 0.023, False
    AddPath ws, "Route_373_128", "M,587.7,608.85;L,586.8,608.85;L,586.8,607.5;L,587.7,607.5;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_373_129", "M,587.7,608.85;L,586.8,608.85;L,586.8,607.5;L,587.7,607.5;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_374_130", "M,1120.05,608.85;L,1119.6,608.85;L,1119.6,607.5;L,1120.05,607.5;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_374_131", "M,1120.05,608.85;L,1119.6,608.85;L,1119.6,607.5;L,1120.05,607.5;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_375_132", "M,1652.4,608.85;L,1651.95,608.85;L,1651.95,607.5;L,1652.4,607.5;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_375_133", "M,1652.4,608.85;L,1651.95,608.85;L,1651.95,607.5;L,1652.4,607.5;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_376_134", "M,2185.2,608.85;L,2184.3,608.85;L,2184.3,607.5;L,2185.2,607.5;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_376_135", "M,2185.2,608.85;L,2184.3,608.85;L,2184.3,607.5;L,2185.2,607.5;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_377_136", "M,2475.9,658.8;L,2475,658.8;L,2475,658.35;L,2475.9,658.35;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_377_137", "M,2475.9,658.8;L,2475,658.8;L,2475,658.35;L,2475.9,658.35;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_377_138", "M,346.14,723.6;L,346.14,724.05;L,345.33,724.05;L,345.33,723.6;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_377_139", "M,346.14,723.6;L,346.14,724.05;L,345.33,724.05;L,345.33,723.6;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_378_140", "M,878.4,724.05;L,877.5,724.05;L,877.5,723.6;L,878.4,723.6;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_378_141", "M,878.4,724.05;L,877.5,724.05;L,877.5,723.6;L,878.4,723.6;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_379_142", "M,1410.75,723.6;L,1410.75,724.05;L,1410.3,724.05;L,1410.3,723.6;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_379_143", "M,1410.75,723.6;L,1410.75,724.05;L,1410.3,724.05;L,1410.3,723.6;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_380_144", "M,1943.55,723.6;L,1943.55,724.05;L,1942.65,724.05;L,1942.65,723.6;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_380_145", "M,1943.55,723.6;L,1943.55,724.05;L,1942.65,724.05;L,1942.65,723.6;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_381_146", "M,2475.9,723.6;L,2475.9,724.05;L,2475,724.05;L,2475,723.6;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_381_147", "M,2475.9,723.6;L,2475.9,724.05;L,2475,724.05;L,2475,723.6;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_303_148", "M,165.015,1207.35;L,159.165,1210.95;L,159.165,1204.2;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_303_149", "M,165.015,1207.35;L,159.165,1210.95;L,159.165,1204.2;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_304_150", "M,210.015,1305;L,206.64,1299.15;L,213.39,1299.15;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_304_151", "M,210.015,1305;L,206.64,1299.15;L,213.39,1299.15;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_305_152", "M,352.8,1342.35;L,346.95,1345.95;L,346.95,1339.2;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_305_153", "M,352.8,1342.35;L,346.95,1345.95;L,346.95,1339.2;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_306_154", "M,428.13,1287.9;L,428.04,1287.9;L,427.86,1287.9;L,427.68,1287.9;L,427.5,1287.45;L,427.32,1287.9;L,427.14,1287.9;L,426.96,1287.9;L,426.87,1287.9;L,426.87,1287;L,428.13,1287;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_306_155", "M,430.875,1252.8;L,424.125,1252.8;L,427.5,1246.95;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_306_156", "M,431.955,1253.25;L,423.045,1253.25;L,427.5,1245.6;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_306_157", "M,425.205,1251.9;L,429.795,1251.9;L,427.5,1247.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_306_158", "M,431.955,1253.25;L,423.045,1253.25;L,427.5,1245.6;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_306_159", "M,425.205,1251.9;L,429.795,1251.9;L,427.5,1247.85;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_308_160", "M,362.52,1207.35;L,324,1207.35;C,315,1207.35,315,1216.35,315,1216.35;L,315,1333.35;C,315,1342.35,324,1342.35,324,1342.35;L,346.32,1342.35", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_309_161", "M,535.05,1341.9;L,534.6,1341.9;L,534.6,1342.35;L,534.6,1342.8;L,534.6,1342.8;L,534.15,1343.25;L,534.15,1343.25;L,533.7,1343.25;L,533.7,1343.7;L,533.25,1343.7;L,533.25,1343.7;L,501.75,1343.7;L,501.75,1341.45;L,532.8,1341.45;L,532.8,1052.1;L,535.05,1052.1;H", _
        7503624, -1, 0, False
End Sub

Private Sub DrawPart08(ByVal ws As Worksheet)
    AddPath ws, "Route_309_162", "M,535.05,1341.9;L,534.6,1341.9;L,534.6,1342.35;L,534.6,1342.8;L,534.6,1342.8;L,534.15,1343.25;L,534.15,1343.25;L,533.7,1343.25;L,533.7,1343.7;L,533.25,1343.7;L,533.25,1343.7;L,501.75,1343.7;L,501.75,1341.45;L,532.8,1341.45;L,532.8,1052.1;L,535.05,1052.1;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_309_163", "M,501.75,1343.7;L,500.4,1343.7;L,500.4,1343.25;L,500.4,1342.8;L,500.85,1342.8;L,500.85,1342.8;L,500.85,1342.35;L,500.85,1342.35;L,500.85,1341.9;L,500.4,1341.9;L,500.4,1341.9;L,500.4,1341.45;L,501.75,1341.45;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_309_164", "M,535.95,668.25;L,531.45,668.25;L,533.7,661.95;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_309_165", "M,537.3,669.6;L,530.1,669.6;L,533.7,658.35;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_309_166", "M,533.25,667.35;L,534.15,667.35;L,533.7,665.55;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_309_167", "M,537.3,669.6;L,530.1,669.6;L,533.7,658.35;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_309_168", "M,533.25,667.35;L,534.15,667.35;L,533.7,665.55;H", _
        -1, 7503624, 0.023, False
    AddPath ws, "Route_311_169", "M,772.65,1342.35;L,766.8,1345.95;L,766.8,1339.2;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_311_170", "M,772.65,1342.35;L,766.8,1345.95;L,766.8,1339.2;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_312_171", "M,847.8,1287.9;L,847.8,1287.9;L,847.8,1287.9;L,847.35,1287.9;L,847.35,1287.45;L,846.9,1287.9;L,846.9,1287.9;L,846.9,1287.9;L,846.45,1287.9;L,846.45,1287;L,847.8,1287;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_312_172", "M,850.5,1252.8;L,843.75,1252.8;L,847.35,1246.95;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_312_173", "M,851.85,1253.25;L,842.85,1253.25;L,847.35,1245.6;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_312_174", "M,845.1,1251.9;L,849.6,1251.9;L,847.35,1247.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_312_175", "M,851.85,1253.25;L,842.85,1253.25;L,847.35,1245.6;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_312_176", "M,845.1,1251.9;L,849.6,1251.9;L,847.35,1247.85;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_314_177", "M,847.35,1107.9;L,850.5,1113.75;L,843.75,1113.75;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_314_178", "M,847.35,1107.9;L,850.5,1113.75;L,843.75,1113.75;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_315_179", "M,782.1,1065.15;L,661.5,1065.15;C,652.05,1064.7,652.05,1073.7,652.05,1073.7;L,652.5,1294.2", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_315_180", "M,652.5,1300.95;L,648.9,1295.1;L,655.65,1295.1;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_315_181", "M,652.5,1300.95;L,648.9,1295.1;L,655.65,1295.1;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_316_182", "M,921.6,1343.25;L,920.7,1343.25;L,920.7,1342.8;L,920.7,1342.8;L,920.7,1342.8;L,920.7,1342.35;L,920.7,1342.35;L,920.7,1341.9;L,920.7,1341.9;L,920.7,1341.9;L,921.6,1341.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_316_183", "M,975.15,1342.35;L,969.3,1345.95;L,969.3,1339.2;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_316_184", "M,976.5,1342.35;L,968.4,1346.85;L,968.4,1337.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_316_185", "M,969.75,1344.6;L,973.8,1342.35;L,969.75,1340.1;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_316_186", "M,976.5,1342.35;L,968.4,1346.85;L,968.4,1337.85;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_316_187", "M,969.75,1344.6;L,973.8,1342.35;L,969.75,1340.1;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_318_188", "M,1043.1,1287.9;L,1042.65,1287.9;L,1042.65,1287.9;L,1042.65,1287.9;L,1042.2,1287.9;L,1042.2,1287.9;L,1041.75,1287.9;L,1041.75,1287.9;L,1041.75,1287.9;L,1041.75,1287;L,1043.1,1287;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_318_189", "M,1045.8,1249.2;L,1039.05,1249.2;L,1042.2,1243.35;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_318_190", "M,1046.7,1250.1;L,1037.7,1250.1;L,1042.2,1242;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_318_191", "M,1039.95,1248.75;L,1044.45,1248.75;L,1042.2,1244.7;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_318_192", "M,1046.7,1250.1;L,1037.7,1250.1;L,1042.2,1242;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_318_193", "M,1039.95,1248.75;L,1044.45,1248.75;L,1042.2,1244.7;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_320_194", "M,1139.4,1207.35;L,1133.55,1210.95;L,1133.55,1204.2;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_320_195", "M,1139.4,1207.35;L,1133.55,1210.95;L,1133.55,1204.2;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_321_196", "M,1108.8,1343.25;L,1107.9,1343.25;L,1107.9,1342.8;L,1108.35,1342.8;L,1108.35,1342.8;L,1108.35,1342.35;L,1108.35,1342.35;L,1108.35,1341.9;L,1107.9,1341.9;L,1107.9,1341.9;L,1108.8,1341.9;H", _
        8739893, -1, 0, False
End Sub

Private Sub DrawPart09(ByVal ws As Worksheet)
    AddPath ws, "Route_321_197", "M,1209.6,1342.35;L,1203.75,1345.95;L,1203.75,1339.2;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_321_198", "M,1210.95,1342.35;L,1202.85,1346.85;L,1202.85,1337.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_321_199", "M,1204.2,1344.6;L,1208.25,1342.35;L,1204.2,1340.1;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_321_200", "M,1210.95,1342.35;L,1202.85,1346.85;L,1202.85,1337.85;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_321_201", "M,1204.2,1344.6;L,1208.25,1342.35;L,1204.2,1340.1;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_323_202", "M,1289.25,1397.25;L,1289.25,1397.25;L,1289.7,1397.25;L,1289.7,1397.25;L,1290.15,1397.25;L,1290.15,1397.25;L,1290.15,1397.25;L,1290.6,1397.25;L,1290.6,1397.7;L,1289.25,1397.7;L,1289.25,1397.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_323_203", "M,1293.3,1430.55;L,1289.7,1436.4;L,1286.55,1430.55;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_323_204", "M,1294.2,1429.65;L,1289.7,1437.3;L,1285.2,1429.65;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_323_205", "M,1289.7,1435.05;L,1291.95,1431;L,1287.45,1431;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_323_206", "M,1294.2,1429.65;L,1289.7,1437.3;L,1285.2,1429.65;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_323_207", "M,1289.7,1435.05;L,1291.95,1431;L,1287.45,1431;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_325_208", "M,1224.9,1481.4;L,1197.45,1481.4;C,1188.45,1481.4,1188.45,1472.4,1188.45,1472.4;L,1188.45,1351.35;C,1188.45,1342.35,1197.45,1342.35,1197.45,1342.35;L,1202.85,1342.35", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_325_209", "M,1209.6,1342.35;L,1203.75,1345.95;L,1203.75,1339.2;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_326_210", "M,1369.35,1343.25;L,1368.45,1343.25;L,1368.45,1342.8;L,1368.9,1342.8;L,1368.9,1342.8;L,1368.9,1342.35;L,1368.9,1342.35;L,1368.9,1341.9;L,1368.45,1341.9;L,1368.45,1341.9;L,1369.35,1341.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_326_211", "M,1400.85,1342.35;L,1396.8,1345.05;L,1396.8,1340.1;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_326_212", "M,1402.2,1342.35;L,1396.8,1345.5;L,1396.8,1344.15;L,1399.5,1342.35;L,1396.8,1340.55;L,1396.8,1339.2;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_326_213", "M,1402.2,1342.35;L,1396.8,1345.5;L,1396.8,1344.15;L,1399.5,1342.35;L,1396.8,1340.55;L,1396.8,1339.2;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_328_214", "M,1493.1,1280.25;L,1492.65,1280.25;L,1492.65,1280.25;L,1492.65,1280.25;L,1492.2,1280.25;L,1492.2,1280.25;L,1491.75,1280.25;L,1491.75,1280.25;L,1491.75,1280.25;L,1491.75,1279.35;L,1493.1,1279.35;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_328_215", "M,1495.8,1252.8;L,1489.05,1252.8;L,1492.2,1246.95;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_328_216", "M,1496.7,1253.25;L,1487.7,1253.25;L,1492.2,1245.6;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_328_217", "M,1489.95,1251.9;L,1494.45,1251.9;L,1492.2,1247.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_328_218", "M,1496.7,1253.25;L,1487.7,1253.25;L,1492.2,1245.6;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_328_219", "M,1489.95,1251.9;L,1494.45,1251.9;L,1492.2,1247.85;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_330_220", "M,1491.75,1404.45;L,1491.75,1404.9;L,1492.2,1404.9;L,1492.2,1404.9;L,1492.65,1404.9;L,1492.65,1404.9;L,1492.65,1404.45;L,1493.1,1404.45;L,1493.1,1405.35;L,1491.75,1405.35;L,1491.75,1404.45;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_330_221", "M,1495.8,1439.55;L,1492.2,1445.4;L,1489.05,1439.55;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_330_222", "M,1496.7,1438.65;L,1492.2,1446.3;L,1487.7,1438.65;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_330_223", "M,1492.2,1444.05;L,1494.45,1440;L,1489.95,1440;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_330_224", "M,1496.7,1438.65;L,1492.2,1446.3;L,1487.7,1438.65;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_330_225", "M,1492.2,1444.05;L,1494.45,1440;L,1489.95,1440;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_332_226", "M,1492.2,1580.4;L,1489.05,1574.55;L,1495.8,1574.55;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_332_227", "M,1492.2,1580.4;L,1489.05,1574.55;L,1495.8,1574.55;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_333_228", "M,1492.2,1736.55;L,1489.05,1730.7;L,1495.8,1730.7;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_333_229", "M,1492.2,1736.55;L,1489.05,1730.7;L,1495.8,1730.7;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_334_230", "M,1384.2,1779.75;L,1384.2,1781.1;L,1384.65,1782.45;L,1385.1,1783.35;L,1385.55,1784.7;L,1386.45,1785.6;L,1387.8,1786.5;L,1388.7,1786.95;L,1390.05,1787.85;L,1391.4,1787.85;L,1392.75,1787.85;L,1423.8,1787.85;L,1423.8,1789.2;L,1392.3,1789.2;L,1390.95,1789.2;L,1389.6,1788.75;L,1388.25,1788.3;L,1386.9,1787.4;L,1385.55,1786.5;L,1384.65,1785.15;L,1383.75,1784.25;L,1383.3,1782.45;L,1382.85,1781.1;L,1382.85,1779.75;L,1382.85,1739.25;L,1384.2,1739.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_334_231", "M,1418.4,1616.85;L,1392.75,1616.85;L,1391.4,1616.85;L,1390.05,1617.3;L,1388.7,1617.75;L,1387.8,1618.2;L,1386.45,1619.1;L,1385.55,1620.45;L,1385.1,1621.35;L,1384.65,1622.7;L,1384.2,1624.05;L,1384.2,1625.4;L,1384.2,1725.75;L,1382.85,1725.75;L,1382.85,1624.95;L,1382.85,1623.6;L,1383.3,1622.25;L,1383.75,1620.9;L,1384.65,1619.55;L,1385.55,1618.2;L,1386.9,1617.3;L,1388.25,1616.4;L,1389.6,1615.95;L,1390.95,1615.5;L,1392.3,1615.5;L,1418.4,1615.5;H", _
        8739893, -1, 0, False
End Sub

Private Sub DrawPart10(ByVal ws As Worksheet)
    AddPath ws, "Route_334_232", "M,1384.2,1779.75;L,1384.2,1781.1;L,1384.65,1782.45;L,1385.1,1783.35;L,1385.55,1784.7;L,1386.45,1785.6;L,1387.8,1786.5;L,1388.7,1786.95;L,1390.05,1787.85;L,1391.4,1787.85;L,1392.75,1787.85;L,1423.8,1787.85;L,1423.8,1789.2;L,1392.3,1789.2;L,1390.95,1789.2;L,1389.6,1788.75;L,1388.25,1788.3;L,1386.9,1787.4;L,1385.55,1786.5;L,1384.65,1785.15;L,1383.75,1784.25;L,1383.3,1782.45;L,1382.85,1781.1;L,1382.85,1779.75;L,1382.85,1739.25;L,1384.2,1739.25;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_334_233", "M,1418.4,1616.85;L,1392.75,1616.85;L,1391.4,1616.85;L,1390.05,1617.3;L,1388.7,1617.75;L,1387.8,1618.2;L,1386.45,1619.1;L,1385.55,1620.45;L,1385.1,1621.35;L,1384.65,1622.7;L,1384.2,1624.05;L,1384.2,1625.4;L,1384.2,1725.75;L,1382.85,1725.75;L,1382.85,1624.95;L,1382.85,1623.6;L,1383.3,1622.25;L,1383.75,1620.9;L,1384.65,1619.55;L,1385.55,1618.2;L,1386.9,1617.3;L,1388.25,1616.4;L,1389.6,1615.95;L,1390.95,1615.5;L,1392.3,1615.5;L,1418.4,1615.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_334_234", "M,1424.7,1788.3;L,1424.7,1788.3;L,1424.7,1788.3;L,1424.7,1788.75;L,1424.7,1788.75;L,1424.7,1789.2;L,1424.7,1789.2;L,1424.7,1789.2;L,1423.8,1789.2;L,1423.8,1787.85;L,1424.7,1787.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_334_235", "M,1424.7,1616.4;L,1418.85,1619.55;L,1418.85,1612.8;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_334_236", "M,1426.05,1616.4;L,1418.4,1620.45;L,1418.4,1611.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_334_237", "M,1419.75,1618.65;L,1423.35,1616.4;L,1419.75,1613.7;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_334_238", "M,1426.05,1616.4;L,1418.4,1620.45;L,1418.4,1611.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_334_239", "M,1419.75,1618.65;L,1423.35,1616.4;L,1419.75,1613.7;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_336_240", "M,1590.3,1779.75;L,1590.3,1781.1;L,1589.85,1782.45;L,1589.4,1784.25;L,1588.5,1785.15;L,1587.6,1786.5;L,1586.25,1787.4;L,1585.35,1788.3;L,1583.55,1788.75;L,1582.2,1789.2;L,1580.85,1789.2;L,1560.6,1789.2;L,1560.6,1787.85;L,1580.85,1787.85;L,1582.2,1787.85;L,1583.55,1787.85;L,1584.45,1786.95;L,1585.8,1786.5;L,1586.7,1785.6;L,1587.6,1784.7;L,1588.05,1783.35;L,1588.95,1782.45;L,1588.95,1781.1;L,1588.95,1779.75;L,1588.95,1739.25;L,1590.3,1739.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_336_241", "M,1594.8,1616.85;L,1593.9,1616.85;L,1593.45,1616.85;L,1592.55,1617.3;L,1592.1,1617.75;L,1591.65,1618.2;L,1591.2,1618.65;L,1590.75,1619.1;L,1590.75,1619.55;L,1590.3,1620.45;L,1590.3,1620.9;L,1590.3,1725.75;L,1588.95,1725.75;L,1588.95,1620.9;L,1589.4,1620;L,1589.4,1619.1;L,1589.85,1618.65;L,1590.3,1617.75;L,1590.75,1617.3;L,1591.2,1616.4;L,1592.1,1615.95;L,1593,1615.95;L,1593.9,1615.5;L,1594.35,1615.5;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_336_242", "M,1590.3,1779.75;L,1590.3,1781.1;L,1589.85,1782.45;L,1589.4,1784.25;L,1588.5,1785.15;L,1587.6,1786.5;L,1586.25,1787.4;L,1585.35,1788.3;L,1583.55,1788.75;L,1582.2,1789.2;L,1580.85,1789.2;L,1560.6,1789.2;L,1560.6,1787.85;L,1580.85,1787.85;L,1582.2,1787.85;L,1583.55,1787.85;L,1584.45,1786.95;L,1585.8,1786.5;L,1586.7,1785.6;L,1587.6,1784.7;L,1588.05,1783.35;L,1588.95,1782.45;L,1588.95,1781.1;L,1588.95,1779.75;L,1588.95,1739.25;L,1590.3,1739.25;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_336_243", "M,1594.8,1616.85;L,1593.9,1616.85;L,1593.45,1616.85;L,1592.55,1617.3;L,1592.1,1617.75;L,1591.65,1618.2;L,1591.2,1618.65;L,1590.75,1619.1;L,1590.75,1619.55;L,1590.3,1620.45;L,1590.3,1620.9;L,1590.3,1725.75;L,1588.95,1725.75;L,1588.95,1620.9;L,1589.4,1620;L,1589.4,1619.1;L,1589.85,1618.65;L,1590.3,1617.75;L,1590.75,1617.3;L,1591.2,1616.4;L,1592.1,1615.95;L,1593,1615.95;L,1593.9,1615.5;L,1594.35,1615.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_336_244", "M,1560.6,1789.2;L,1559.7,1789.2;L,1560.15,1789.2;L,1560.15,1789.2;L,1560.15,1788.75;L,1560.15,1788.75;L,1560.15,1788.3;L,1560.15,1788.3;L,1560.15,1788.3;L,1559.7,1787.85;L,1560.6,1787.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_336_245", "M,1601.1,1616.4;L,1595.25,1619.55;L,1595.25,1612.8;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_336_246", "M,1602.45,1616.4;L,1594.8,1620.45;L,1594.8,1611.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_336_247", "M,1595.7,1618.65;L,1599.75,1616.4;L,1595.7,1613.7;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_336_248", "M,1602.45,1616.4;L,1594.8,1620.45;L,1594.8,1611.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_336_249", "M,1595.7,1618.65;L,1599.75,1616.4;L,1595.7,1613.7;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_338_250", "M,1694.7,1736.55;L,1691.55,1730.7;L,1698.3,1730.7;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_338_251", "M,1694.7,1736.55;L,1691.55,1730.7;L,1698.3,1730.7;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_339_252", "M,1770.75,1789.2;L,1769.85,1789.2;L,1769.85,1789.2;L,1769.85,1789.2;L,1769.85,1788.75;L,1770.3,1788.75;L,1769.85,1788.3;L,1769.85,1788.3;L,1769.85,1788.3;L,1769.85,1787.85;L,1770.75,1787.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_339_253", "M,1830.15,1788.75;L,1824.3,1791.9;L,1824.3,1785.15;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_339_254", "M,1831.5,1788.75;L,1823.85,1793.25;L,1823.85,1784.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_339_255", "M,1824.75,1791;L,1828.8,1788.75;L,1824.75,1786.5;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_339_256", "M,1831.5,1788.75;L,1823.85,1793.25;L,1823.85,1784.25;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_339_257", "M,1824.75,1791;L,1828.8,1788.75;L,1824.75,1786.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_341_258", "M,1695.6,1840.5;L,1694.25,1840.5;L,1694.25,1839.6;L,1694.25,1839.6;L,1694.25,1839.6;L,1694.7,1839.6;L,1694.7,1839.6;L,1695.15,1839.6;L,1695.15,1839.6;L,1695.15,1839.6;L,1695.6,1839.6;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_341_259", "M,1698.3,2290.95;L,1694.7,2296.8;L,1691.55,2290.95;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_341_260", "M,1699.2,2290.5;L,1694.7,2298.15;L,1690.2,2290.5;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_341_261", "M,1694.7,2295.9;L,1696.95,2291.85;L,1692.45,2291.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_341_262", "M,1699.2,2290.5;L,1694.7,2298.15;L,1690.2,2290.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_341_263", "M,1694.7,2295.9;L,1696.95,2291.85;L,1692.45,2291.85;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_343_264", "M,1530.45,2398.95;L,1530.45,2424.6;L,1530.45,2425.95;L,1530.9,2427.3;L,1531.35,2428.65;L,1531.8,2429.55;L,1532.7,2430.45;L,1534.05,2431.35;L,1534.95,2432.25;L,1536.3,2432.7;L,1537.65,2433.15;L,1539,2433.15;L,1685.7,2433.15;L,1687.05,2433.15;L,1688.4,2432.7;L,1689.75,2432.25;L,1690.65,2431.35;L,1691.55,2430.45;L,1692.45,2429.55;L,1693.35,2428.65;L,1693.8,2427.3;L,1694.25,2425.95;L,1694.25,2424.6;L,1694.25,2422.8;L,1695.6,2422.8;L,1695.6,2424.6;L,1695.15,2426.4;L,1695.15,2427.75;L,1694.25,2429.1;L,1693.8,2430.45;L,1692.45,2431.35;L,1691.55,2432.7" & _
        ";L,1690.2,2433.15;L,1688.85,2434.05;L,1687.5,2434.05;L,1685.7,2434.5;L,1538.55,2434.5;L,1537.2,2434.05;L,1535.85,2434.05;L,1534.5,2433.15;L,1533.15,2432.7;L,1531.8,2431.35;L,1530.9,2430.45;L,1530,2429.1;L,1529.55,2427.75;L,1529.1,2426.4;L,1529.1,2424.6;L,1529.1,2398.95;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_343_265", "M,1530.45,2398.95;L,1530.45,2424.6;L,1530.45,2425.95;L,1530.9,2427.3;L,1531.35,2428.65;L,1531.8,2429.55;L,1532.7,2430.45;L,1534.05,2431.35;L,1534.95,2432.25;L,1536.3,2432.7;L,1537.65,2433.15;L,1539,2433.15;L,1685.7,2433.15;L,1687.05,2433.15;L,1688.4,2432.7;L,1689.75,2432.25;L,1690.65,2431.35;L,1691.55,2430.45;L,1692.45,2429.55;L,1693.35,2428.65;L,1693.8,2427.3;L,1694.25,2425.95;L,1694.25,2424.6;L,1694.25,2422.8;L,1695.6,2422.8;L,1695.6,2424.6;L,1695.15,2426.4;L,1695.15,2427.75;L,1694.25,2429.1;L,1693.8,2430.45;L,1692.45,2431.35;L,1691.55,2432.7" & _
        ";L,1690.2,2433.15;L,1688.85,2434.05;L,1687.5,2434.05;L,1685.7,2434.5;L,1538.55,2434.5;L,1537.2,2434.05;L,1535.85,2434.05;L,1534.5,2433.15;L,1533.15,2432.7;L,1531.8,2431.35;L,1530.9,2430.45;L,1530,2429.1;L,1529.55,2427.75;L,1529.1,2426.4;L,1529.1,2424.6;L,1529.1,2398.95;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_343_266", "M,1694.25,2396.7;L,1694.25,2396.7;L,1694.7,2396.7;L,1694.7,2396.7;L,1695.15,2396.7;L,1695.15,2396.7;L,1695.15,2396.7;L,1695.6,2396.7;L,1695.6,2397.15;L,1694.25,2397.15;L,1694.25,2396.7;H", _
        8739893, -1, 0, False
End Sub

Private Sub DrawPart11(ByVal ws As Worksheet)
    AddPath ws, "Route_343_267", "M,1533.15,2398.5;L,1526.4,2398.5;L,1530,2392.65;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_343_268", "M,1534.05,2398.95;L,1525.5,2398.95;L,1530,2391.3;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_343_269", "M,1527.3,2397.6;L,1532.25,2397.6;L,1530,2393.55;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_343_270", "M,1534.05,2398.95;L,1525.5,2398.95;L,1530,2391.3;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_343_271", "M,1527.3,2397.6;L,1532.25,2397.6;L,1530,2393.55;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_345_272", "M,1594.8,2348.1;L,1594.35,2348.1;L,1594.35,2348.1;L,1594.35,2348.1;L,1594.35,2347.65;L,1594.35,2347.65;L,1594.35,2347.65;L,1594.35,2347.65;L,1594.35,2347.2;L,1594.35,2347.2;L,1594.35,2347.2;L,1594.35,2346.75;L,1594.35,2346.75;L,1594.35,2346.75;L,1594.8,2346.75;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_345_273", "M,1594.8,2348.1;L,1594.35,2348.1;L,1594.35,2348.1;L,1594.35,2348.1;L,1594.35,2347.65;L,1594.35,2347.65;L,1594.35,2347.65;L,1594.35,2347.65;L,1594.35,2347.2;L,1594.35,2347.2;L,1594.35,2347.2;L,1594.35,2346.75;L,1594.35,2346.75;L,1594.35,2346.75;L,1594.8,2346.75;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_345_274", "M,1614.6,2347.65;L,1608.75,2350.8;L,1608.75,2344.05;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_345_275", "M,1614.6,2347.65;L,1608.75,2350.8;L,1608.75,2344.05;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_346_276", "M,1909.35,1847.7;L,1909.35,2338.65;L,1908.9,2340;L,1908.9,2341.35;L,1908,2342.7;L,1907.55,2344.05;L,1906.2,2345.4;L,1905.3,2346.3;L,1903.95,2347.2;L,1902.6,2347.65;L,1901.25,2348.1;L,1899.45,2348.1;L,1821.6,2348.1;L,1821.6,2346.75;L,1899.45,2346.75;L,1900.8,2346.75;L,1902.15,2346.3;L,1903.5,2345.85;L,1904.4,2345.4;L,1905.3,2344.5;L,1906.2,2343.15;L,1907.1,2342.25;L,1907.55,2340.9;L,1908,2339.55;L,1908,2338.2;L,1908,1847.7;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_346_277", "M,1909.35,1847.7;L,1909.35,2338.65;L,1908.9,2340;L,1908.9,2341.35;L,1908,2342.7;L,1907.55,2344.05;L,1906.2,2345.4;L,1905.3,2346.3;L,1903.95,2347.2;L,1902.6,2347.65;L,1901.25,2348.1;L,1899.45,2348.1;L,1821.6,2348.1;L,1821.6,2346.75;L,1899.45,2346.75;L,1900.8,2346.75;L,1902.15,2346.3;L,1903.5,2345.85;L,1904.4,2345.4;L,1905.3,2344.5;L,1906.2,2343.15;L,1907.1,2342.25;L,1907.55,2340.9;L,1908,2339.55;L,1908,2338.2;L,1908,1847.7;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_346_278", "M,1774.35,2346.75;L,1774.35,2348.1;L,1773.45,2348.1;L,1773.45,2348.1;L,1773.9,2347.65;L,1773.9,2347.65;L,1773.9,2347.65;L,1773.9,2347.2;L,1773.9,2347.2;L,1773.45,2346.75;L,1773.45,2346.75;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_346_279", "M,1912.05,1846.8;L,1905.3,1846.8;L,1908.45,1840.95;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_346_280", "M,1912.95,1847.7;L,1903.95,1847.7;L,1908.45,1839.6;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_346_281", "M,1906.2,1846.35;L,1910.7,1846.35;L,1908.45,1842.3;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_346_282", "M,1912.95,1847.7;L,1903.95,1847.7;L,1908.45,1839.6;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_346_283", "M,1906.2,1846.35;L,1910.7,1846.35;L,1908.45,1842.3;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_348_284", "M,1901.25,1615.5;L,1902.6,1615.95;L,1903.95,1616.4;L,1905.3,1617.3;L,1906.2,1618.2;L,1907.55,1619.55;L,1908,1620.9;L,1908.9,1622.25;L,1908.9,1623.6;L,1909.35,1624.95;L,1909.35,1699.2;L,1908,1699.2;L,1908,1625.4;L,1908,1624.05;L,1907.55,1622.7;L,1907.1,1621.35;L,1906.2,1620.45;L,1905.3,1619.1;L,1904.4,1618.2;L,1903.5,1617.75;L,1902.15,1617.3;L,1900.8,1616.85;L,1899.45,1616.85;L,1795.05,1616.85;L,1795.05,1615.5;L,1899.45,1615.5;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_348_285", "M,1901.25,1615.5;L,1902.6,1615.95;L,1903.95,1616.4;L,1905.3,1617.3;L,1906.2,1618.2;L,1907.55,1619.55;L,1908,1620.9;L,1908.9,1622.25;L,1908.9,1623.6;L,1909.35,1624.95;L,1909.35,1699.2;L,1908,1699.2;L,1908,1625.4;L,1908,1624.05;L,1907.55,1622.7;L,1907.1,1621.35;L,1906.2,1620.45;L,1905.3,1619.1;L,1904.4,1618.2;L,1903.5,1617.75;L,1902.15,1617.3;L,1900.8,1616.85;L,1899.45,1616.85;L,1795.05,1616.85;L,1795.05,1615.5;L,1899.45,1615.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_348_286", "M,1909.35,1737;L,1909.35,1737.9;L,1908.9,1737.9;L,1908.9,1737.9;L,1908.9,1737.45;L,1908.45,1737.45;L,1908.45,1737.45;L,1908,1737.9;L,1908,1737.9;L,1908,1737.9;L,1908,1737;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_348_287", "M,1794.6,1619.55;L,1788.75,1616.4;L,1794.6,1612.8;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_348_288", "M,1795.05,1620.45;L,1787.4,1616.4;L,1795.05,1611.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_348_289", "M,1789.65,1616.4;L,1793.7,1618.65;L,1793.7,1613.7;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_348_290", "M,1795.05,1620.45;L,1787.4,1616.4;L,1795.05,1611.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_348_291", "M,1789.65,1616.4;L,1793.7,1618.65;L,1793.7,1613.7;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_350_292", "M,1986.3,1789.2;L,1985.4,1789.2;L,1985.4,1789.2;L,1985.85,1789.2;L,1985.85,1788.75;L,1985.85,1788.75;L,1985.85,1788.3;L,1985.85,1788.3;L,1985.4,1788.3;L,1985.4,1787.85;L,1986.3,1787.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_350_293", "M,2043.9,1788.75;L,2038.05,1791.9;L,2038.05,1785.15;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_350_294", "M,2045.25,1788.75;L,2037.6,1793.25;L,2037.6,1784.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_350_295", "M,2038.5,1791;L,2042.55,1788.75;L,2038.5,1786.5;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_350_296", "M,2045.25,1788.75;L,2037.6,1793.25;L,2037.6,1784.25;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_350_297", "M,2038.5,1791;L,2042.55,1788.75;L,2038.5,1786.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_352_298", "M,2185.2,1789.2;L,2184.3,1789.2;L,2184.3,1789.2;L,2184.3,1789.2;L,2184.3,1788.75;L,2184.3,1788.75;L,2184.3,1788.3;L,2184.3,1788.3;L,2184.3,1788.3;L,2184.3,1787.85;L,2185.2,1787.85;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_352_299", "M,2345.4,1788.75;L,2339.55,1791.9;L,2339.55,1785.15;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_352_300", "M,2346.75,1788.75;L,2339.1,1793.25;L,2339.1,1784.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_352_301", "M,2340,1791;L,2344.05,1788.75;L,2340,1786.5;H", _
        8739893, -1, 0, False
End Sub

Private Sub DrawPart12(ByVal ws As Worksheet)
    AddPath ws, "Route_352_302", "M,2346.75,1788.75;L,2339.1,1793.25;L,2339.1,1784.25;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_352_303", "M,2340,1791;L,2344.05,1788.75;L,2340,1786.5;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_354_304", "M,2115.45,1840.5;L,2114.1,1840.5;L,2114.1,1839.6;L,2114.1,1839.6;L,2114.55,1839.6;L,2114.55,1839.6;L,2115,1839.6;L,2115,1839.6;L,2115,1839.6;L,2115.45,1839.6;L,2115.45,1839.6;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_354_305", "M,2118.15,1891.8;L,2115,1897.65;L,2111.4,1891.8;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_354_306", "M,2119.05,1891.35;L,2115,1899;L,2110.5,1891.35;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_354_307", "M,2115,1896.3;L,2117.25,1892.25;L,2112.3,1892.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_354_308", "M,2119.05,1891.35;L,2115,1899;L,2110.5,1891.35;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_354_309", "M,2115,1896.3;L,2117.25,1892.25;L,2112.3,1892.25;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_356_310", "M,2267.1,1949.4;L,2268.45,1949.4;L,2269.8,1949.85;L,2271.6,1950.3;L,2272.5,1951.2;L,2273.85,1952.1;L,2274.75,1953.45;L,2275.65,1954.8;L,2276.1,1956.15;L,2276.55,1957.5;L,2276.55,1958.85;L,2276.55,2088;L,2275.2,2088;L,2275.2,1958.85;L,2275.2,1957.5;L,2275.2,1956.15;L,2274.3,1955.25;L,2273.85,1953.9;L,2272.95,1953;L,2272.05,1952.1;L,2270.7,1951.65;L,2269.8,1951.2;L,2268.45,1950.75;L,2267.1,1950.75;L,2227.05,1950.75;L,2227.05,1949.4;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_356_311", "M,2267.1,1949.4;L,2268.45,1949.4;L,2269.8,1949.85;L,2271.6,1950.3;L,2272.5,1951.2;L,2273.85,1952.1;L,2274.75,1953.45;L,2275.65,1954.8;L,2276.1,1956.15;L,2276.55,1957.5;L,2276.55,1958.85;L,2276.55,2088;L,2275.2,2088;L,2275.2,1958.85;L,2275.2,1957.5;L,2275.2,1956.15;L,2274.3,1955.25;L,2273.85,1953.9;L,2272.95,1953;L,2272.05,1952.1;L,2270.7,1951.65;L,2269.8,1951.2;L,2268.45,1950.75;L,2267.1,1950.75;L,2227.05,1950.75;L,2227.05,1949.4;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_356_312", "M,2185.2,1949.4;L,2185.2,1950.75;L,2184.3,1950.75;L,2184.3,1950.3;L,2184.3,1950.3;L,2184.3,1950.3;L,2184.3,1949.85;L,2184.3,1949.85;L,2184.3,1949.4;L,2184.3,1949.4;L,2184.3,1949.4;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_356_313", "M,2276.1,2094.3;L,2272.5,2088.45;L,2279.25,2088.45;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_356_314", "M,2276.1,2095.65;L,2271.6,2088;L,2280.6,2088;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_356_315", "M,2273.85,2088.9;L,2276.1,2092.95;L,2278.35,2088.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_356_316", "M,2276.1,2095.65;L,2271.6,2088;L,2280.6,2088;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_356_317", "M,2273.85,2088.9;L,2276.1,2092.95;L,2278.35,2088.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_358_318", "M,2115.45,2001.6;L,2114.1,2001.6;L,2114.1,2000.7;L,2114.1,2000.7;L,2114.55,2001.15;L,2114.55,2001.15;L,2115,2001.15;L,2115,2001.15;L,2115,2001.15;L,2115.45,2000.7;L,2115.45,2000.7;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_358_319", "M,2118.15,2305.8;L,2115,2311.65;L,2111.4,2305.8;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_358_320", "M,2119.05,2305.35;L,2115,2313;L,2110.5,2305.35;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_358_321", "M,2115,2310.75;L,2117.25,2306.7;L,2112.3,2306.7;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_358_322", "M,2119.05,2305.35;L,2115,2313;L,2110.5,2305.35;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_358_323", "M,2115,2310.75;L,2117.25,2306.7;L,2112.3,2306.7;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_360_324", "M,2350.35,2137.5;L,2398.5,2137.5;C,2407.05,2137.5,2407.05,2128.5,2407.05,2128.5;L,2407.5,1830.6", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_360_325", "M,2407.5,1824.3;L,2410.65,1830.15;L,2403.9,1830.15;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_360_326", "M,2407.5,1824.3;L,2410.65,1830.15;L,2403.9,1830.15;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_361_327", "M,2407.5,1661.4;L,2410.65,1667.25;L,2403.9,1667.25;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_361_328", "M,2407.5,1661.4;L,2410.65,1667.25;L,2403.9,1667.25;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_362_329", "M,2410.65,1528.2;L,2403.9,1528.2;L,2407.5,1522.35;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_362_330", "M,2411.55,1529.1;L,2403,1529.1;L,2407.5,1521;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_362_331", "M,2404.8,1527.75;L,2409.75,1527.75;L,2407.5,1523.7;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_362_332", "M,2411.55,1529.1;L,2403,1529.1;L,2407.5,1521;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_362_333", "M,2404.8,1527.75;L,2409.75,1527.75;L,2407.5,1523.7;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_364_334", "M,2334.6,1481.4;L,2311.2,1481.4;C,2302.2,1481.4,2302.2,1490.4,2302.2,1490.4;L,2302.2,1607.4;C,2302.2,1616.4,2311.2,1616.4,2311.2,1616.4;L,2335.05,1616.4", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_364_335", "M,2341.8,1616.4;L,2335.95,1619.55;L,2335.95,1612.8;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_364_336", "M,2341.8,1616.4;L,2335.95,1619.55;L,2335.95,1612.8;H", _
        -1, 8739893, 1.274, False
End Sub

Private Sub DrawPart13(ByVal ws As Worksheet)
    AddPath ws, "Route_365_337", "M,2472.3,1616.85;L,2471.85,1616.85;L,2471.85,1616.85;L,2471.85,1616.4;L,2471.85,1616.4;L,2471.85,1616.4;L,2471.85,1616.4;L,2471.85,1615.95;L,2471.85,1615.95;L,2471.85,1615.95;L,2471.85,1615.5;L,2471.85,1615.5;L,2472.3,1615.5;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_365_338", "M,2517.75,1616.4;L,2511.9,1619.55;L,2511.9,1612.8;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_365_339", "M,2519.1,1616.4;L,2511.45,1620.45;L,2511.45,1611.9;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_365_340", "M,2512.8,1618.65;L,2516.85,1616.4;L,2512.8,1613.7;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_365_341", "M,2519.1,1616.4;L,2511.45,1620.45;L,2511.45,1611.9;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_365_342", "M,2512.8,1618.65;L,2516.85,1616.4;L,2512.8,1613.7;H", _
        -1, 8739893, 0.023, False
    AddPath ws, "Route_367_343", "M,2680.65,1616.4;L,2674.8,1619.55;L,2674.8,1612.8;H", _
        8739893, -1, 0, False
    AddPath ws, "Route_367_344", "M,2680.65,1616.4;L,2674.8,1619.55;L,2674.8,1612.8;H", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_369_345", "M,847.8,1397.7;L,846.9,1397.7;L,846.9,1397.25;L,846.9,1397.25;L,846.9,1397.25;L,847.35,1397.25;L,847.35,1397.25;L,847.8,1397.25;L,847.8,1397.25;H", _
        10326919, -1, 0, False
    AddPath ws, "Route_369_346", "M,847.8,1397.7;L,846.9,1397.7;L,846.9,1397.25;L,846.9,1397.25;L,846.9,1397.25;L,847.35,1397.25;L,847.35,1397.25;L,847.8,1397.25;L,847.8,1397.25;H", _
        -1, 10326919, 0.023, False
    AddPath ws, "Route_373_347", "M,635.4,608.4;L,629.55,611.55;L,629.55,604.8;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_373_348", "M,635.4,608.4;L,629.55,611.55;L,629.55,604.8;H", _
        -1, 7503624, 1.274, False
    AddPath ws, "Route_374_349", "M,1167.75,608.4;L,1161.9,611.55;L,1161.9,604.8;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_374_350", "M,1167.75,608.4;L,1161.9,611.55;L,1161.9,604.8;H", _
        -1, 7503624, 1.274, False
    AddPath ws, "Route_375_351", "M,1700.55,608.4;L,1694.7,611.55;L,1694.7,604.8;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_375_352", "M,1700.55,608.4;L,1694.7,611.55;L,1694.7,604.8;H", _
        -1, 7503624, 1.274, False
    AddPath ws, "Route_376_353", "M,2232.9,608.4;L,2227.05,611.55;L,2227.05,604.8;H", _
        7503624, -1, 0, False
    AddPath ws, "Route_376_354", "M,2232.9,608.4;L,2227.05,611.55;L,2227.05,604.8;H", _
        -1, 7503624, 1.274, False
    AddPath ws, "Route_377_355", "M,2475.45,658.8;L,2475.45,705.6;L,535.95,705.6", _
        -1, 7503624, 0.824, True
    AddPath ws, "Route_377_356", "M,535.95,705.6;C,535.95,704.25,535.05,703.35,533.7,703.35", _
        -1, 7503624, 0.824, True
    AddPath ws, "Route_377_357", "M,533.7,703.35;C,532.35,703.35,531.45,704.25,531.45,705.6", _
        -1, 7503624, 0.824, True
    AddPath ws, "Route_377_358", "M,531.45,705.6;L,531.45,705.6;L,345.735,705.6;L,345.735,723.6", _
        -1, 7503624, 0.824, True
    AddPath ws, "Route_378_359", "M,2475.45,658.8;L,2475.45,705.6;L,877.95,705.6;L,877.95,723.6", _
        -1, 7503624, 0.824, True
    AddPath ws, "Route_379_360", "M,2475.45,696.6;C,2475.45,705.6,2466.45,705.6,2466.45,705.6", _
        -1, 7503624, 0.824, True
    AddPath ws, "Route_379_361", "M,1419.75,705.6;C,1410.75,705.6,1410.75,714.6,1410.75,714.6", _
        -1, 7503624, 0.824, True
    AddPath ws, "Route_380_362", "M,1952.1,705.6;C,1943.1,705.6,1943.1,714.6,1943.1,714.6", _
        -1, 7503624, 0.824, True
    AddPath ws, "Route_381_363", "M,2475.45,658.8;L,2475.45,705.6;L,2475.45,723.6", _
        -1, 7503624, 0.824, True
    AddPath ws, "Route_303_364", "M,141.885,1207.35;L,158.535,1207.35", _
        -1, 8739893, 1.274, True
    AddPath ws, "Route_304_365", "M,210.015,1225.35;L,210.015,1298.25", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_305_366", "M,299.385,1342.35;L,346.32,1342.35", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_311_367", "M,721.35,1342.35;L,765.9,1342.35", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_314_368", "M,847.35,1168.65;L,847.35,1114.65", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_320_369", "M,1098.45,1207.35;L,1133.1,1207.35", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_332_370", "M,1492.2,1516.5;L,1492.2,1574.1", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_333_371", "M,1492.2,1651.05;L,1492.2,1729.8", _
        -1, 8739893, 1.274, False
End Sub

Private Sub DrawPart14(ByVal ws As Worksheet)
    AddPath ws, "Route_338_372", "M,1694.7,1658.7;L,1694.7,1729.8", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_345_373", "M,1594.8,2347.65;L,1607.85,2347.65", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_361_374", "M,2407.5,1753.65;L,2407.5,1667.7", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_367_375", "M,2655.9,1616.4;L,2674.35,1616.4", _
        -1, 8739893, 1.274, False
    AddPath ws, "Route_368_376", "M,210.015,1379.25;L,210.015,1442.7", _
        -1, 10326919, 0.824, True
    AddPath ws, "Route_369_377", "M,847.35,1397.7;L,847.35,1436.85", _
        -1, 10326919, 0.824, True
    AddPath ws, "Route_370_378", "M,1551.6,1481.4;L,1631.7,1481.4", _
        -1, 10326919, 0.824, True
    AddPath ws, "Route_371_379", "M,2479.5,1481.4;L,2524.5,1481.4", _
        -1, 10326919, 0.824, True
    AddPath ws, "Route_372_380", "M,2181.6,2347.65;L,2248.65,2347.65", _
        -1, 10326919, 0.824, True
    AddPath ws, "Route_373_381", "M,587.7,608.4;L,629.1,608.4", _
        -1, 7503624, 1.274, False
    AddPath ws, "Route_374_382", "M,1120.05,608.4;L,1161.45,608.4", _
        -1, 7503624, 1.274, False
    AddPath ws, "Route_375_383", "M,1652.4,608.4;L,1693.8,608.4", _
        -1, 7503624, 1.274, False
    AddPath ws, "Route_376_384", "M,2185.2,608.4;L,2226.6,608.4", _
        -1, 7503624, 1.274, False
    AddPath ws, "Route_379_385", "M,2475.45,658.8;L,2475.45,696.6", _
        -1, 7503624, 0.824, True
    AddPath ws, "Route_379_386", "M,2466.45,705.6;L,1419.75,705.6", _
        -1, 7503624, 0.824, True
    AddPath ws, "Route_379_387", "M,1410.75,714.6;L,1410.75,723.6", _
        -1, 7503624, 0.824, True
    AddPath ws, "Route_380_388", "M,2466.45,705.6;L,1952.1,705.6", _
        -1, 7503624, 0.824, True
    AddPath ws, "Route_380_389", "M,1943.1,714.6;L,1943.1,723.6", _
        -1, 7503624, 0.824, True
    AddLabel ws, "Text_001", 28.796, 23.845, 620.615, 21.502, _
        "DEAL CLOSING PROCESS FLOW - CURRENT VS AI-ASSISTED CA READER", _
        16.5, 16777215, True, False, False, 22.275, -1
    AddLabel ws, "Text_002", 28.796, 58.1, 760.727, 17.749, _
        "ALTERNATIVE CA SUPPORT FLOWS: CURRENT (PURPLE) / PROPOSED (TEAL) | SHARED CLOSING WORKFLOW BELOW", _
        12.75, 3287584, True, False, False, 17.213, -1
    AddLabel ws, "Text_003", 22.354, 220.015, 18, 107.545, _
        "DEAL MANAGER", _
        12, 4339238, True, True, True, 16.2, -1
    AddLabel ws, "Text_004", 38.478, 207.623, 18, 132.357, _
        "MANUAL CA REVIEW", _
        12, 4339238, True, True, True, 16.2, -1
    AddLabel ws, "Text_005", 22.354, 647.34, 18, 160.01, _
        "AI-ASSISTED CA READER", _
        12, 4339238, True, True, True, 16.2, -1
    AddLabel ws, "Text_006", 38.478, 622.047, 18, 210.953, _
        "PROPOSED + HUMAN VALIDATION", _
        12, 4339238, True, True, True, 16.2, -1
    AddLabel ws, "Text_007", 30.418, 1036.577, 18, 56.523, _
        "CB/CTM", _
        12, 4339238, True, True, True, 16.2, -1
    AddLabel ws, "Text_008", 22.354, 1242.664, 18, 334.186, _
        "DEAL MANAGER - CLOSING AND FULFILLMENT TEAM /", _
        12, 4339238, True, True, True, 16.2, -1
    AddLabel ws, "Text_009", 38.478, 1318.722, 18, 181.628, _
        "ANALYST - SERVICING TEAM", _
        12, 4339238, True, True, True, 16.2, -1
    AddLabel ws, "Text_010", 30.418, 1739.472, 18, 233.828, _
        "1st ASSOCIATE - VERIFICATION TEAM", _
        12, 4339238, True, True, True, 16.2, -1
    AddLabel ws, "Text_011", 14.286, 2073.449, 18, 116.751, _
        "2nd ASSOCIATE - ", _
        12, 4339238, True, True, True, 16.2, -1
    AddLabel ws, "Text_012", 30.418, 2086.708, 18, 93.592, _
        "VERIFICATION", _
        12, 4339238, True, True, True, 16.2, -1
    AddLabel ws, "Text_013", 46.533, 2111.622, 18, 43.928, _
        "TEAM", _
        12, 4339238, True, True, True, 16.2, -1
    AddLabel ws, "Text_014", 22.354, 2255.172, 18, 191.528, _
        "MANAGER / DIRECTOR - LOAN", _
        12, 4339238, True, True, True, 16.2, -1
    AddLabel ws, "Text_015", 38.478, 2294.322, 18, 113.228, _
        "SERVICING TEAM", _
        12, 4339238, True, True, True, 16.2, -1
    AddLabel ws, "Text_016", 69.735, 1187.979, 70.58, 31.55, _
        "Pre-Closing" & vbLf & "Day", _
        11.25, 3287584, False, False, True, 15.3, -1
    AddLabel ws, "Text_017", 191.46, 1199.533, 37.302, 17.002, _
        "Start", _
        12, 3287584, False, False, True, 16.2, -1
End Sub

Private Sub DrawPart15(ByVal ws As Worksheet)
    AddLabel ws, "Text_018", 129.99, 1319.829, 160.39, 46.4, _
        "Email CB/CTM/CCG Finance/" & vbLf & "CCG Credit for confirmation" & vbLf & "of CP's and executed CA", _
        11.25, 3287584, False, False, True, 14.85, -1
    AddLabel ws, "Text_019", 378.075, 1327.479, 98.472, 31.1, _
        "Executed CA and" & vbLf & "CP's received?", _
        11.25, 3287584, False, False, True, 14.85, -1
    AddLabel ws, "Text_020", 385.95, 1177.179, 83.105, 61.7, _
        "Follow up with" & vbLf & "CB/CTM/CCG" & vbLf & "Finance/CCG" & vbLf & "Credit", _
        11.25, 3287584, False, False, True, 15.3, -1
    AddLabel ws, "Text_021", 594.75, 1312.179, 115.255, 61.7, _
        "Verify Final" & vbLf & "Allocation on" & vbLf & "executed CA against" & vbLf & "Risk-approved eTR", _
        11.25, 3287584, False, False, True, 14.85, -1
    AddLabel ws, "Text_022", 800.4, 1334.533, 93.854, 17.002, _
        "Discrepancies?", _
        12, 3287584, False, False, True, 16.2, -1
    AddLabel ws, "Text_023", 784.65, 1177.179, 125.33, 61.7, _
        "Email CB/CTM/CCG" & vbLf & "Finance/CCG Credit to" & vbLf & "make any required" & vbLf & "changes", _
        11.25, 3287584, False, False, True, 15.3, -1
    AddLabel ws, "Text_024", 799.05, 1049.829, 96.369, 31.1, _
        "CB/CTM initiates" & vbLf & "changes via eTR", _
        11.25, 3287584, False, False, True, 14.85, -1
    AddLabel ws, "Text_025", 1011.45, 1334.533, 61.928, 17.002, _
        "Funding?", _
        12, 3287584, False, False, True, 16.2, -1
    AddLabel ws, "Text_026", 994.35, 1191.433, 96.456, 33.202, _
        "Setup Facility in" & vbLf & "ACBS", _
        12, 3287584, False, False, True, 16.2, -1
    AddLabel ws, "Text_027", 1154.55, 1195.033, 38.629, 17.002, _
        "Fees", _
        12, 3287584, False, False, True, 16.2, -1
    AddLabel ws, "Text_028", 1232.4, 1327.479, 115.01, 31.1, _
        "Notice received from" & vbLf & "Agent/Borrower?", _
        11.25, 3287584, False, False, True, 14.85, -1
    AddLabel ws, "Text_029", 1230.6, 1466.079, 118.737, 31.55, _
        "Follow up with Agent/" & vbLf & "Borrower", _
        11.25, 3287584, False, False, True, 15.3, -1
    AddLabel ws, "Text_030", 1429.95, 1319.829, 124.581, 46.4, _
        "NY: Over USD 25MM?" & vbLf & "LO: Over 10MM?" & vbLf & "TO: Over 25MM", _
        11.25, 3287584, False, False, True, 14.85, -1
    AddLabel ws, "Text_031", 1434.9, 1177.179, 114.897, 61.7, _
        "Advise Treasury and" & vbLf & "Cash Management" & vbLf & "of confirmation of" & vbLf & "CP's", _
        11.25, 3287584, False, False, True, 15.3, -1
    AddLabel ws, "Text_032", 1446.15, 1465.033, 93.021, 33.202, _
        "Input Loan via" & vbLf & "Drawdown", _
        12, 3287584, True, False, True, 16.2, -1
    AddLabel ws, "Text_033", 1434.45, 1593.429, 115.674, 46.4, _
        "Attain CTM/" & vbLf & "Manager Approval in" & vbLf & "Drawdown", _
        11.25, 3287584, False, False, True, 15.3, -1
    AddLabel ws, "Text_034", 1457.4, 1780.483, 69.962, 17.001, _
        "Approved?", _
        12, 3287584, False, False, True, 16.2, -1
    AddLabel ws, "Text_035", 1612.8, 1583.833, 141.852, 65.601, _
        "Input Wire in" & vbLf & "-  BESS for NY" & vbLf & "-  RFT for TO" & vbLf & "-  RIMMS / ACBS for LO", _
        12, 3287584, False, False, False, 16.2, -1
    AddLabel ws, "Text_036", 1644.15, 1781.079, 101.631, 16.25, _
        "Less than 25MM?", _
        11.25, 3287584, False, False, True, 15.188, -1
    AddLabel ws, "Text_037", 1643.7, 2332.329, 102.121, 31.1, _
        "Manager provides" & vbLf & "email approval", _
        11.25, 3287584, False, False, True, 14.85, -1
    AddLabel ws, "Text_038", 1488.45, 2332.329, 82.79, 31.55, _
        "Follow up with" & vbLf & "Manager", _
        11.25, 3287584, False, False, True, 15.3, -1
    AddLabel ws, "Text_039", 1856.55, 1773.429, 104.243, 31.55, _
        "Wire details match" & vbLf & "funding notice?", _
        11.25, 3287584, False, False, True, 15.3, -1
    AddLabel ws, "Text_040", 2066.25, 1781.079, 97.357, 16.25, _
        "100MM or more?", _
        11.25, 3287584, False, False, True, 15.188, -1
    AddLabel ws, "Text_041", 2068.5, 1942.179, 92.911, 16.25, _
        "1 Billion or more", _
        11.25, 3287584, False, False, True, 15.188, -1
    AddLabel ws, "Text_042", 2214.3, 2121.133, 123.766, 33.201, _
        "ADMRQ and Second" & vbLf & "verifier approves", _
        12, 3287584, False, False, True, 16.2, -1
    AddLabel ws, "Text_043", 2061.75, 2339.383, 106.14, 17.001, _
        "Payment Sentinel", _
        12, 3287584, False, False, True, 16.2, -1
    AddLabel ws, "Text_044", 2366.85, 1780.483, 81.154, 17.001, _
        "Release wire", _
        12, 3287584, False, False, True, 16.2, -1
    AddLabel ws, "Text_045", 2354.7, 1608.729, 105.514, 16.25, _
        "Funding / Wire Ref", _
        11.25, 3287584, False, False, True, 15.188, -1
    AddLabel ws, "Text_046", 2349.75, 1450.779, 115.566, 61.7, _
        "Get assistance from" & vbLf & "Cash Management" & vbLf & "only if funds are held" & vbLf & "by chase", _
        11.25, 3287584, False, False, True, 15.3, -1
    AddLabel ws, "Text_047", 2532.9, 1593.429, 109.13, 46.4, _
        "Provide Wire Ref to" & vbLf & "CB and Agent/" & vbLf & "Borrower", _
        11.25, 3287584, False, False, True, 14.85, -1
    AddLabel ws, "Text_048", 2692.65, 1603.633, 44.659, 17.001, _
        "Loans", _
        12, 3287584, False, False, True, 16.2, -1
    AddLabel ws, "Text_049", 120.33, 1458.429, 163.135, 46.4, _
        "Include funding request, if" & vbLf & "not received at pre-closing." & vbLf & "Agent/Business send an email", _
        11.25, 16777215, False, False, False, 15.3, -1
    AddLabel ws, "Text_050", 757.8, 1450.779, 150.848, 91.85, _
        "Note:" & vbLf & "-  Required changes before" & vbLf & "   proceeding:" & vbLf & "-  GL Unit, Undersubscribed" & vbLf & "   Commitments, Product" & vbLf & "   Options, Booking Points", _
        11.25, 16777215, False, False, False, 15.3, -1
    AddLabel ws, "Text_051", 1642.5, 1435.929, 224.8, 91.85, _
        "Note:" & vbLf & "For time-sensitive payments, Administrator" & vbLf & "proceeds to input wire immediately after" & vbLf & "CID (no need to wait for approval)." & vbLf & "Executed CA, CP's confirmation and" & vbLf & "Risk-approved eTR are held.", _
        11.25, 16777215, False, False, False, 15.3, -1
    AddLabel ws, "Text_052", 2535.3, 1450.779, 183.488, 61.7, _
        "Note:" & vbLf & "Call 416-842-6610" & vbLf & "Cash Management will raise funds" & vbLf & "from Fed if required", _
        11.25, 16777215, False, False, False, 15.3, -1
End Sub

Private Sub DrawPart16(ByVal ws As Worksheet)
    AddLabel ws, "Text_053", 2259.45, 2324.679, 186.266, 46.4, _
        "Note:" & vbLf & "Payment Sentinel requires Director" & vbLf & "level approval.", _
        11.25, 16777215, False, False, False, 15.3, -1
    AddLabel ws, "Text_054", 234.36, 576.91, 223.29, 63.05, _
        "1 - PREPARE CA" & vbLf & "Load the executed agreement" & vbLf & "Build a searchable page index", _
        15.75, 3287584, False, False, True, 21.15, -1
    AddLabel ws, "Text_055", 749.7, 576.91, 257.126, 63.05, _
        "2 - EXTRACT WITH AI" & vbLf & "25 fields + verbatim supporting text" & vbLf & "Flag missing / ambiguous values", _
        15.75, 3287584, False, False, True, 21.15, -1
    AddLabel ws, "Text_056", 1271.25, 576.91, 279.117, 63.05, _
        "3 - TRACE EVIDENCE" & vbLf & "Match supporting text to source pages" & vbLf & "Flag unmatched evidence", _
        15.75, 3287584, False, False, True, 21.15, -1
    AddLabel ws, "Text_057", 1809.45, 576.91, 267.705, 63.05, _
        "4 - DEAL MANAGER REVIEW" & vbLf & "Validate values and source evidence" & vbLf & "Resolve exceptions before reuse", _
        15.75, 3287584, False, False, True, 21.15, -1
    AddLabel ws, "Text_058", 2332.8, 576.91, 285.912, 63.05, _
        "5 - SAVE VALIDATED RECORD" & vbLf & "Reviewed values + source references" & vbLf & "One shared CA summary for CA1-CA5", _
        15.75, 3287584, False, False, True, 21.15, -1
    AddLabel ws, "Text_059", 682.5, 1277.24, 37.43, 17.749, _
        "CA1", _
        12.75, 16777215, True, False, True, 17.213, -1
    AddLabel ws, "Text_060", 1059.15, 1148.09, 37.43, 17.749, _
        "CA2", _
        12.75, 16777215, True, False, True, 17.213, -1
    AddLabel ws, "Text_061", 1168.05, 1149.89, 37.43, 17.749, _
        "CA3", _
        12.75, 16777215, True, False, True, 17.213, -1
    AddLabel ws, "Text_062", 1513.2, 1421.69, 37.43, 17.749, _
        "CA4", _
        12.75, 16777215, True, False, True, 17.213, -1
    AddLabel ws, "Text_063", 2709.3, 1558.49, 37.43, 17.748, _
        "CA5", _
        12.75, 16777215, True, False, True, 17.213, -1
    AddLabel ws, "Text_064", 105.795, 96.596, 534.28, 23, _
        "CURRENT PROCESS: MANUAL CA LOOKUP & EXTRACTION", _
        18, 10373232, True, False, False, 24.3, -1
    AddLabel ws, "Text_065", 105.795, 131.59, 740.827, 20.75, _
        "CA Reader project scope: 25 fields in total. Each process uses its relevant subset; subsets may overlap.", _
        15.75, 3287584, False, False, False, 21.263, -1
    AddLabel ws, "Text_066", 104.805, 180.28, 344.408, 41.9, _
        "CA1 - Verify Final Allocation on executed CA" & vbLf & "against Risk-approved eTR", _
        15.75, 16777215, True, False, False, 21.15, -1
    AddLabel ws, "Text_067", 109.575, 258.434, 261.346, 88.012, _
        "Open executed CA" & vbLf & "Locate the allocation field subset" & vbLf & "Extract and check final allocations" & vbLf & "against the Risk-approved eTR", _
        16.5, 3287584, False, False, False, 22.185, -1
    AddLabel ws, "Text_068", 637.2, 190.855, 226.298, 20.75, _
        "CA2 - Setup Facility in ACBS", _
        15.75, 16777215, True, False, False, 21.263, -1
    AddLabel ws, "Text_069", 641.7, 258.434, 287.905, 88.012, _
        "Return to executed CA" & vbLf & "Locate the facility setup field subset" & vbLf & "Read and interpret the required terms" & vbLf & "Manually enter facility data in ACBS", _
        16.5, 3287584, False, False, False, 22.185, -1
    AddLabel ws, "Text_070", 1169.55, 190.855, 92.984, 20.75, _
        "CA3 - Fees", _
        15.75, 16777215, True, False, False, 21.263, -1
    AddLabel ws, "Text_071", 1174.5, 258.434, 331.901, 88.012, _
        "Return to executed CA" & vbLf & "Locate the fee field subset" & vbLf & "Read fee terms and calculation rules" & vbLf & "Apply the required data in the Fees process", _
        16.5, 3287584, False, False, False, 22.185, -1
    AddLabel ws, "Text_072", 1701.9, 190.855, 247.786, 20.75, _
        "CA4 - Input Loan via Drawdown", _
        15.75, 16777215, True, False, False, 21.263, -1
    AddLabel ws, "Text_073", 1706.85, 258.434, 319.789, 88.012, _
        "Return to executed CA" & vbLf & "Locate the drawdown field subset" & vbLf & "Read terms needed for the borrowing" & vbLf & "Manually enter and check Drawdown data", _
        16.5, 3287584, False, False, False, 22.185, -1
    AddLabel ws, "Text_074", 2234.7, 190.855, 103.396, 20.75, _
        "CA5 - Loans", _
        15.75, 16777215, True, False, False, 21.263, -1
    AddLabel ws, "Text_075", 2239.2, 258.434, 316.601, 88.012, _
        "Return to executed CA" & vbLf & "Locate the loan field subset" & vbLf & "Read terms needed for loan processing" & vbLf & "Manually apply data in the Loans process", _
        16.5, 3287584, False, False, False, 22.185, -1
    AddLabel ws, "Text_076", 105.795, 395.459, 813.916, 21.502, _
        "Repeated effort: locate clauses / schedules - read and interpret - extract - manually enter or verify", _
        16.5, 10373232, True, False, False, 22.275, -1
    AddLabel ws, "Text_077", 105.795, 430.229, 906.934, 19.251, _
        "Current support is shown for comparison. The teal lane shows the proposed replacement at CA1-CA5; routing determines which uses occur.", _
        14.25, 7956312, False, False, False, 19.238, -1
    AddLabel ws, "Text_078", 105.795, 483.326, 732.207, 23, _
        "PROPOSED PROCESS: AI-ASSISTED EXTRACTION, HUMAN VALIDATION & REUSE", _
        18, 7503624, True, False, False, 24.3, -1
    AddLabel ws, "Text_079", 105.795, 518.41, 841.271, 20.75, _
        "Extract the 25-field scope once per CA version, validate the results, then reuse the relevant subset at each touchpoint.", _
        15.75, 3287584, False, False, False, 21.263, -1
    AddLabel ws, "Text_080", 104.85, 676.81, 502.733, 20.75, _
        "ONE VALIDATED CA RECORD - FIVE PROCESS-SPECIFIC USES", _
        15.75, 7503624, True, False, False, 21.263, -1
    AddLabel ws, "Text_081", 104.805, 731.71, 344.408, 41.9, _
        "CA1 - Verify Final Allocation on executed CA" & vbLf & "against Risk-approved eTR", _
        15.75, 16777215, True, False, False, 21.15, -1
    AddLabel ws, "Text_082", 109.575, 796.51, 271.107, 63.05, _
        "Reuse validated allocation fields" & vbLf & "Compare with Risk-approved eTR" & vbLf & "DM resolves allocation discrepancies", _
        15.75, 3287584, False, False, False, 21.15, -1
    AddLabel ws, "Text_083", 104.805, 877.349, 269.584, 19.252, _
        "Change: reusable allocation reference", _
        14.25, 7503624, True, False, False, 19.238, -1
    AddLabel ws, "Text_084", 637.2, 742.06, 226.298, 20.75, _
        "CA2 - Setup Facility in ACBS", _
        15.75, 16777215, True, False, False, 21.263, -1
    AddLabel ws, "Text_085", 641.7, 796.51, 246.233, 63.05, _
        "Reuse validated facility fields" & vbLf & "DM enters / checks data in ACBS" & vbLf & "Same setup controls continue", _
        15.75, 3287584, False, False, False, 21.15, -1
    AddLabel ws, "Text_086", 637.2, 877.349, 276.599, 19.252, _
        "Change: less manual source extraction", _
        14.25, 7503624, True, False, False, 19.238, -1
    AddLabel ws, "Text_087", 1169.55, 742.06, 92.984, 20.75, _
        "CA3 - Fees", _
        15.75, 16777215, True, False, False, 21.263, -1
End Sub

Private Sub DrawPart17(ByVal ws As Worksheet)
    AddLabel ws, "Text_088", 1174.5, 796.51, 301.451, 63.05, _
        "Reuse validated fee fields and evidence" & vbLf & "DM applies terms in the Fees process" & vbLf & "Fee processing remains subject to review", _
        15.75, 3287584, False, False, False, 21.15, -1
    AddLabel ws, "Text_089", 1169.55, 877.349, 294.356, 19.252, _
        "Change: reviewed fee terms ready for use", _
        14.25, 7503624, True, False, False, 19.238, -1
    AddLabel ws, "Text_090", 1701.9, 742.06, 247.786, 20.75, _
        "CA4 - Input Loan via Drawdown", _
        15.75, 16777215, True, False, False, 21.263, -1
    AddLabel ws, "Text_091", 1706.85, 796.51, 287.775, 63.05, _
        "Reuse validated drawdown fields" & vbLf & "DM inputs / checks Drawdown details" & vbLf & "Funding notice and approvals still apply", _
        15.75, 3287584, False, False, False, 21.15, -1
    AddLabel ws, "Text_092", 1701.9, 877.349, 275.839, 19.252, _
        "Change: validated borrowing reference", _
        14.25, 7503624, True, False, False, 19.238, -1
    AddLabel ws, "Text_093", 2234.7, 742.06, 103.396, 20.75, _
        "CA5 - Loans", _
        15.75, 16777215, True, False, False, 21.263, -1
    AddLabel ws, "Text_094", 2239.2, 796.51, 301.042, 63.05, _
        "Reuse validated loan fields and evidence" & vbLf & "DM applies terms in the Loans process" & vbLf & "Revisit CA for exceptions or amendments", _
        15.75, 3287584, False, False, False, 21.15, -1
    AddLabel ws, "Text_095", 2234.7, 877.349, 305.669, 19.252, _
        "Change: shared source for loan processing", _
        14.25, 7503624, True, False, False, 19.238, -1
    AddLabel ws, "Text_096", 105.795, 917.414, 715.95, 21.501, _
        "Key change: repeated CA searches become review and reuse of a validated CA summary.", _
        16.5, 7503624, True, False, False, 22.275, -1
    AddLabel ws, "Text_097", 105.795, 951.599, 885.745, 19.252, _
        "Deal Manager owns validation and exceptions. ACBS / Drawdown entry and approvals remain manual. Revalidate when the CA changes.", _
        14.25, 7956312, False, False, False, 19.238, -1
    AddLabel ws, "Text_098", 420.3, 1255.929, 26.346, 16.25, _
        "No", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_099", 518.85, 1029.626, 43.749, 23, _
        "Yes", _
        18, 3287584, False, False, False, 24.3, 16777215
    AddLabel ws, "Text_100", 837.9, 1255.929, 30.338, 16.25, _
        "Yes", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_101", 937.35, 1334.679, 26.346, 16.25, _
        "No", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_102", 1035, 1256.379, 26.346, 16.25, _
        "No", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_103", 1149.3, 1334.679, 30.338, 16.25, _
        "Yes", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_104", 1282.5, 1411.179, 26.346, 16.25, _
        "No", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_105", 1378.35, 1334.679, 30.337, 16.25, _
        "Yes", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_106", 1483.2, 1256.379, 30.338, 16.25, _
        "Yes", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_107", 1485, 1418.379, 26.346, 16.25, _
        "No", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_108", 1376.55, 1724.829, 26.346, 16.25, _
        "No", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_109", 1580.4, 1724.829, 30.337, 16.25, _
        "Yes", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_110", 1790.1, 1781.079, 30.337, 16.25, _
        "Yes", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_111", 1687.5, 1877.829, 26.346, 16.25, _
        "No", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_112", 1687.5, 2408.379, 26.346, 16.25, _
        "No", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_113", 1803.15, 2339.979, 30.337, 16.25, _
        "Yes", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_114", 1901.25, 1698.279, 26.346, 16.25, _
        "No", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_115", 2007, 1781.079, 30.338, 16.25, _
        "Yes", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_116", 2227.5, 1781.079, 26.346, 16.25, _
        "No", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_117", 2105.55, 1859.829, 30.338, 16.25, _
        "Yes", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_118", 2212.65, 1942.179, 26.346, 16.25, _
        "No", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_119", 2105.55, 2011.929, 30.338, 16.25, _
        "Yes", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_120", 2400.3, 1541.229, 26.346, 16.25, _
        "No", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    AddLabel ws, "Text_121", 2488.05, 1608.729, 30.338, 16.25, _
        "Yes", _
        11.25, 3287584, False, False, False, 15.188, 16777215
    GroupNode ws, "Node_38", "Box_38|Text_016"
End Sub

Private Sub DrawPart18(ByVal ws As Worksheet)
    GroupNode ws, "Node_41", "Box_41|Text_017"
    GroupNode ws, "Node_43", "Box_43|Text_018"
    GroupNode ws, "Node_47", "Box_47|Text_019"
    GroupNode ws, "Node_50", "Box_50|Text_020"
    GroupNode ws, "Node_55", "Box_55|Text_021"
    GroupNode ws, "Node_60", "Box_60|Text_022"
    GroupNode ws, "Node_62", "Box_62|Text_023"
    GroupNode ws, "Node_67", "Box_67|Text_024"
    GroupNode ws, "Node_70", "Box_70|Text_025"
    GroupNode ws, "Node_72", "Box_72|Text_026"
    GroupNode ws, "Node_75", "Box_75|Text_027"
    GroupNode ws, "Node_77", "Box_77|Text_028"
    GroupNode ws, "Node_80", "Box_80|Text_029"
    GroupNode ws, "Node_83", "Box_83|Text_030"
    GroupNode ws, "Node_87", "Box_87|Text_031"
    GroupNode ws, "Node_92", "Box_92|Text_032"
    GroupNode ws, "Node_95", "Box_95|Text_033"
    GroupNode ws, "Node_99", "Box_99|Text_034"
    GroupNode ws, "Node_101", "Box_101|Text_035"
    GroupNode ws, "Node_106", "Box_106|Text_036"
    GroupNode ws, "Node_108", "Box_108|Text_037"
    GroupNode ws, "Node_111", "Box_111|Text_038"
    GroupNode ws, "Node_114", "Box_114|Text_039"
    GroupNode ws, "Node_117", "Box_117|Text_040"
    GroupNode ws, "Node_119", "Box_119|Text_041"
    GroupNode ws, "Node_121", "Box_121|Text_042"
    GroupNode ws, "Node_124", "Box_124|Text_043"
    GroupNode ws, "Node_126", "Box_126|Text_044"
    GroupNode ws, "Node_128", "Box_128|Text_045"
    GroupNode ws, "Node_130", "Box_130|Text_046"
    GroupNode ws, "Node_135", "Box_135|Text_047"
    GroupNode ws, "Node_139", "Box_139|Text_048"
    GroupNode ws, "Node_141", "Box_141|Text_049"
    GroupNode ws, "Node_145", "Box_145|Text_050"
    GroupNode ws, "Node_152", "Box_152|Text_051"
End Sub

Private Sub DrawPart19(ByVal ws As Worksheet)
    GroupNode ws, "Node_159", "Box_159|Text_052"
    GroupNode ws, "Node_164", "Box_164|Text_053"
    GroupNode ws, "Node_168", "Box_168|Text_054"
    GroupNode ws, "Node_172", "Box_172|Text_055"
    GroupNode ws, "Node_176", "Box_176|Text_056"
    GroupNode ws, "Node_180", "Box_180|Text_057"
    GroupNode ws, "Node_184", "Box_184|Text_058"
    GroupNode ws, "Node_188", "Box_188|Box_258|Text_081|Text_082|Text_083"
    GroupNode ws, "Node_189", "Box_189|Box_267|Text_084|Text_085|Text_086"
    GroupNode ws, "Node_190", "Box_190|Box_275|Text_087|Text_088|Text_089"
    GroupNode ws, "Node_191", "Box_191|Box_283|Text_090|Text_091|Text_092"
    GroupNode ws, "Node_192", "Box_192|Box_291|Text_093|Text_094|Text_095"
    GroupNode ws, "Node_193", "Box_193|Text_059"
    GroupNode ws, "Node_195", "Box_195|Text_060"
    GroupNode ws, "Node_197", "Box_197|Text_061"
    GroupNode ws, "Node_199", "Box_199|Text_062"
    GroupNode ws, "Node_201", "Box_201|Text_063"
    GroupNode ws, "Node_207", "Box_207|Box_208|Text_066|Text_067"
    GroupNode ws, "Node_216", "Box_216|Box_217|Text_068|Text_069"
    GroupNode ws, "Node_224", "Box_224|Box_225|Text_070|Text_071"
    GroupNode ws, "Node_232", "Box_232|Box_233|Text_072|Text_073"
    GroupNode ws, "Node_240", "Box_240|Box_241|Text_074|Text_075"
    ws.Shapes("Route_303_001").ZOrder msoBringToFront
    ws.Shapes("Route_303_002").ZOrder msoBringToFront
    ws.Shapes("Route_304_003").ZOrder msoBringToFront
    ws.Shapes("Route_304_004").ZOrder msoBringToFront
    ws.Shapes("Route_305_005").ZOrder msoBringToFront
    ws.Shapes("Route_305_006").ZOrder msoBringToFront
    ws.Shapes("Route_306_007").ZOrder msoBringToFront
    ws.Shapes("Route_306_008").ZOrder msoBringToFront
    ws.Shapes("Route_306_009").ZOrder msoBringToFront
    ws.Shapes("Route_306_010").ZOrder msoBringToFront
    ws.Shapes("Route_308_011").ZOrder msoBringToFront
    ws.Shapes("Route_308_012").ZOrder msoBringToFront
    ws.Shapes("Route_309_013").ZOrder msoBringToFront
End Sub

Private Sub DrawPart20(ByVal ws As Worksheet)
    ws.Shapes("Route_309_014").ZOrder msoBringToFront
    ws.Shapes("Route_311_015").ZOrder msoBringToFront
    ws.Shapes("Route_311_016").ZOrder msoBringToFront
    ws.Shapes("Route_312_017").ZOrder msoBringToFront
    ws.Shapes("Route_312_018").ZOrder msoBringToFront
    ws.Shapes("Route_312_019").ZOrder msoBringToFront
    ws.Shapes("Route_312_020").ZOrder msoBringToFront
    ws.Shapes("Route_314_021").ZOrder msoBringToFront
    ws.Shapes("Route_314_022").ZOrder msoBringToFront
    ws.Shapes("Route_315_023").ZOrder msoBringToFront
    ws.Shapes("Route_315_024").ZOrder msoBringToFront
    ws.Shapes("Route_316_025").ZOrder msoBringToFront
    ws.Shapes("Route_316_026").ZOrder msoBringToFront
    ws.Shapes("Route_316_027").ZOrder msoBringToFront
    ws.Shapes("Route_316_028").ZOrder msoBringToFront
    ws.Shapes("Route_318_029").ZOrder msoBringToFront
    ws.Shapes("Route_318_030").ZOrder msoBringToFront
    ws.Shapes("Route_318_031").ZOrder msoBringToFront
    ws.Shapes("Route_318_032").ZOrder msoBringToFront
    ws.Shapes("Route_320_033").ZOrder msoBringToFront
    ws.Shapes("Route_320_034").ZOrder msoBringToFront
    ws.Shapes("Route_321_035").ZOrder msoBringToFront
    ws.Shapes("Route_321_036").ZOrder msoBringToFront
    ws.Shapes("Route_321_037").ZOrder msoBringToFront
    ws.Shapes("Route_321_038").ZOrder msoBringToFront
    ws.Shapes("Route_323_039").ZOrder msoBringToFront
    ws.Shapes("Route_323_040").ZOrder msoBringToFront
    ws.Shapes("Route_323_041").ZOrder msoBringToFront
    ws.Shapes("Route_323_042").ZOrder msoBringToFront
    ws.Shapes("Route_325_043").ZOrder msoBringToFront
    ws.Shapes("Route_325_044").ZOrder msoBringToFront
    ws.Shapes("Route_326_045").ZOrder msoBringToFront
    ws.Shapes("Route_326_046").ZOrder msoBringToFront
    ws.Shapes("Route_328_047").ZOrder msoBringToFront
    ws.Shapes("Route_328_048").ZOrder msoBringToFront
End Sub

Private Sub DrawPart21(ByVal ws As Worksheet)
    ws.Shapes("Route_328_049").ZOrder msoBringToFront
    ws.Shapes("Route_328_050").ZOrder msoBringToFront
    ws.Shapes("Route_330_051").ZOrder msoBringToFront
    ws.Shapes("Route_330_052").ZOrder msoBringToFront
    ws.Shapes("Route_330_053").ZOrder msoBringToFront
    ws.Shapes("Route_330_054").ZOrder msoBringToFront
    ws.Shapes("Route_332_055").ZOrder msoBringToFront
    ws.Shapes("Route_332_056").ZOrder msoBringToFront
    ws.Shapes("Route_333_057").ZOrder msoBringToFront
    ws.Shapes("Route_333_058").ZOrder msoBringToFront
    ws.Shapes("Route_338_059").ZOrder msoBringToFront
    ws.Shapes("Route_338_060").ZOrder msoBringToFront
    ws.Shapes("Route_339_061").ZOrder msoBringToFront
    ws.Shapes("Route_339_062").ZOrder msoBringToFront
    ws.Shapes("Route_339_063").ZOrder msoBringToFront
    ws.Shapes("Route_339_064").ZOrder msoBringToFront
    ws.Shapes("Route_341_065").ZOrder msoBringToFront
    ws.Shapes("Route_341_066").ZOrder msoBringToFront
    ws.Shapes("Route_341_067").ZOrder msoBringToFront
    ws.Shapes("Route_341_068").ZOrder msoBringToFront
    ws.Shapes("Route_343_069").ZOrder msoBringToFront
    ws.Shapes("Route_343_070").ZOrder msoBringToFront
    ws.Shapes("Route_346_071").ZOrder msoBringToFront
    ws.Shapes("Route_346_072").ZOrder msoBringToFront
    ws.Shapes("Route_348_073").ZOrder msoBringToFront
    ws.Shapes("Route_348_074").ZOrder msoBringToFront
    ws.Shapes("Route_350_075").ZOrder msoBringToFront
    ws.Shapes("Route_350_076").ZOrder msoBringToFront
    ws.Shapes("Route_350_077").ZOrder msoBringToFront
    ws.Shapes("Route_350_078").ZOrder msoBringToFront
    ws.Shapes("Route_352_079").ZOrder msoBringToFront
    ws.Shapes("Route_352_080").ZOrder msoBringToFront
    ws.Shapes("Route_352_081").ZOrder msoBringToFront
    ws.Shapes("Route_352_082").ZOrder msoBringToFront
    ws.Shapes("Route_354_083").ZOrder msoBringToFront
End Sub

Private Sub DrawPart22(ByVal ws As Worksheet)
    ws.Shapes("Route_354_084").ZOrder msoBringToFront
    ws.Shapes("Route_354_085").ZOrder msoBringToFront
    ws.Shapes("Route_354_086").ZOrder msoBringToFront
    ws.Shapes("Route_356_087").ZOrder msoBringToFront
    ws.Shapes("Route_356_088").ZOrder msoBringToFront
    ws.Shapes("Route_358_089").ZOrder msoBringToFront
    ws.Shapes("Route_358_090").ZOrder msoBringToFront
    ws.Shapes("Route_358_091").ZOrder msoBringToFront
    ws.Shapes("Route_358_092").ZOrder msoBringToFront
    ws.Shapes("Route_360_093").ZOrder msoBringToFront
    ws.Shapes("Route_360_094").ZOrder msoBringToFront
    ws.Shapes("Route_361_095").ZOrder msoBringToFront
    ws.Shapes("Route_361_096").ZOrder msoBringToFront
    ws.Shapes("Route_362_097").ZOrder msoBringToFront
    ws.Shapes("Route_362_098").ZOrder msoBringToFront
    ws.Shapes("Route_362_099").ZOrder msoBringToFront
    ws.Shapes("Route_362_100").ZOrder msoBringToFront
    ws.Shapes("Route_362_101").ZOrder msoBringToFront
    ws.Shapes("Route_364_102").ZOrder msoBringToFront
    ws.Shapes("Route_364_103").ZOrder msoBringToFront
    ws.Shapes("Route_365_104").ZOrder msoBringToFront
    ws.Shapes("Route_365_105").ZOrder msoBringToFront
    ws.Shapes("Route_365_106").ZOrder msoBringToFront
    ws.Shapes("Route_365_107").ZOrder msoBringToFront
    ws.Shapes("Route_367_108").ZOrder msoBringToFront
    ws.Shapes("Route_367_109").ZOrder msoBringToFront
    ws.Shapes("Route_368_110").ZOrder msoBringToFront
    ws.Shapes("Route_368_111").ZOrder msoBringToFront
    ws.Shapes("Route_368_112").ZOrder msoBringToFront
    ws.Shapes("Route_368_113").ZOrder msoBringToFront
    ws.Shapes("Route_369_114").ZOrder msoBringToFront
    ws.Shapes("Route_369_115").ZOrder msoBringToFront
    ws.Shapes("Route_370_116").ZOrder msoBringToFront
    ws.Shapes("Route_370_117").ZOrder msoBringToFront
    ws.Shapes("Route_370_118").ZOrder msoBringToFront
End Sub

Private Sub DrawPart23(ByVal ws As Worksheet)
    ws.Shapes("Route_370_119").ZOrder msoBringToFront
    ws.Shapes("Route_371_120").ZOrder msoBringToFront
    ws.Shapes("Route_371_121").ZOrder msoBringToFront
    ws.Shapes("Route_371_122").ZOrder msoBringToFront
    ws.Shapes("Route_371_123").ZOrder msoBringToFront
    ws.Shapes("Route_372_124").ZOrder msoBringToFront
    ws.Shapes("Route_372_125").ZOrder msoBringToFront
    ws.Shapes("Route_372_126").ZOrder msoBringToFront
    ws.Shapes("Route_372_127").ZOrder msoBringToFront
    ws.Shapes("Route_373_128").ZOrder msoBringToFront
    ws.Shapes("Route_373_129").ZOrder msoBringToFront
    ws.Shapes("Route_374_130").ZOrder msoBringToFront
    ws.Shapes("Route_374_131").ZOrder msoBringToFront
    ws.Shapes("Route_375_132").ZOrder msoBringToFront
    ws.Shapes("Route_375_133").ZOrder msoBringToFront
    ws.Shapes("Route_376_134").ZOrder msoBringToFront
    ws.Shapes("Route_376_135").ZOrder msoBringToFront
    ws.Shapes("Route_377_136").ZOrder msoBringToFront
    ws.Shapes("Route_377_137").ZOrder msoBringToFront
    ws.Shapes("Route_377_138").ZOrder msoBringToFront
    ws.Shapes("Route_377_139").ZOrder msoBringToFront
    ws.Shapes("Route_378_140").ZOrder msoBringToFront
    ws.Shapes("Route_378_141").ZOrder msoBringToFront
    ws.Shapes("Route_379_142").ZOrder msoBringToFront
    ws.Shapes("Route_379_143").ZOrder msoBringToFront
    ws.Shapes("Route_380_144").ZOrder msoBringToFront
    ws.Shapes("Route_380_145").ZOrder msoBringToFront
    ws.Shapes("Route_381_146").ZOrder msoBringToFront
    ws.Shapes("Route_381_147").ZOrder msoBringToFront
    ws.Shapes("Route_303_148").ZOrder msoBringToFront
    ws.Shapes("Route_303_149").ZOrder msoBringToFront
    ws.Shapes("Route_304_150").ZOrder msoBringToFront
    ws.Shapes("Route_304_151").ZOrder msoBringToFront
    ws.Shapes("Route_305_152").ZOrder msoBringToFront
    ws.Shapes("Route_305_153").ZOrder msoBringToFront
End Sub

Private Sub DrawPart24(ByVal ws As Worksheet)
    ws.Shapes("Route_306_154").ZOrder msoBringToFront
    ws.Shapes("Route_306_155").ZOrder msoBringToFront
    ws.Shapes("Route_306_156").ZOrder msoBringToFront
    ws.Shapes("Route_306_157").ZOrder msoBringToFront
    ws.Shapes("Route_306_158").ZOrder msoBringToFront
    ws.Shapes("Route_306_159").ZOrder msoBringToFront
    ws.Shapes("Route_308_160").ZOrder msoBringToFront
    ws.Shapes("Route_309_161").ZOrder msoBringToFront
    ws.Shapes("Route_309_162").ZOrder msoBringToFront
    ws.Shapes("Route_309_163").ZOrder msoBringToFront
    ws.Shapes("Route_309_164").ZOrder msoBringToFront
    ws.Shapes("Route_309_165").ZOrder msoBringToFront
    ws.Shapes("Route_309_166").ZOrder msoBringToFront
    ws.Shapes("Route_309_167").ZOrder msoBringToFront
    ws.Shapes("Route_309_168").ZOrder msoBringToFront
    ws.Shapes("Route_311_169").ZOrder msoBringToFront
    ws.Shapes("Route_311_170").ZOrder msoBringToFront
    ws.Shapes("Route_312_171").ZOrder msoBringToFront
    ws.Shapes("Route_312_172").ZOrder msoBringToFront
    ws.Shapes("Route_312_173").ZOrder msoBringToFront
    ws.Shapes("Route_312_174").ZOrder msoBringToFront
    ws.Shapes("Route_312_175").ZOrder msoBringToFront
    ws.Shapes("Route_312_176").ZOrder msoBringToFront
    ws.Shapes("Route_314_177").ZOrder msoBringToFront
    ws.Shapes("Route_314_178").ZOrder msoBringToFront
    ws.Shapes("Route_315_179").ZOrder msoBringToFront
    ws.Shapes("Route_315_180").ZOrder msoBringToFront
    ws.Shapes("Route_315_181").ZOrder msoBringToFront
    ws.Shapes("Route_316_182").ZOrder msoBringToFront
    ws.Shapes("Route_316_183").ZOrder msoBringToFront
    ws.Shapes("Route_316_184").ZOrder msoBringToFront
    ws.Shapes("Route_316_185").ZOrder msoBringToFront
    ws.Shapes("Route_316_186").ZOrder msoBringToFront
    ws.Shapes("Route_316_187").ZOrder msoBringToFront
    ws.Shapes("Route_318_188").ZOrder msoBringToFront
End Sub

Private Sub DrawPart25(ByVal ws As Worksheet)
    ws.Shapes("Route_318_189").ZOrder msoBringToFront
    ws.Shapes("Route_318_190").ZOrder msoBringToFront
    ws.Shapes("Route_318_191").ZOrder msoBringToFront
    ws.Shapes("Route_318_192").ZOrder msoBringToFront
    ws.Shapes("Route_318_193").ZOrder msoBringToFront
    ws.Shapes("Route_320_194").ZOrder msoBringToFront
    ws.Shapes("Route_320_195").ZOrder msoBringToFront
    ws.Shapes("Route_321_196").ZOrder msoBringToFront
    ws.Shapes("Route_321_197").ZOrder msoBringToFront
    ws.Shapes("Route_321_198").ZOrder msoBringToFront
    ws.Shapes("Route_321_199").ZOrder msoBringToFront
    ws.Shapes("Route_321_200").ZOrder msoBringToFront
    ws.Shapes("Route_321_201").ZOrder msoBringToFront
    ws.Shapes("Route_323_202").ZOrder msoBringToFront
    ws.Shapes("Route_323_203").ZOrder msoBringToFront
    ws.Shapes("Route_323_204").ZOrder msoBringToFront
    ws.Shapes("Route_323_205").ZOrder msoBringToFront
    ws.Shapes("Route_323_206").ZOrder msoBringToFront
    ws.Shapes("Route_323_207").ZOrder msoBringToFront
    ws.Shapes("Route_325_208").ZOrder msoBringToFront
    ws.Shapes("Route_325_209").ZOrder msoBringToFront
    ws.Shapes("Route_326_210").ZOrder msoBringToFront
    ws.Shapes("Route_326_211").ZOrder msoBringToFront
    ws.Shapes("Route_326_212").ZOrder msoBringToFront
    ws.Shapes("Route_326_213").ZOrder msoBringToFront
    ws.Shapes("Route_328_214").ZOrder msoBringToFront
    ws.Shapes("Route_328_215").ZOrder msoBringToFront
    ws.Shapes("Route_328_216").ZOrder msoBringToFront
    ws.Shapes("Route_328_217").ZOrder msoBringToFront
    ws.Shapes("Route_328_218").ZOrder msoBringToFront
    ws.Shapes("Route_328_219").ZOrder msoBringToFront
    ws.Shapes("Route_330_220").ZOrder msoBringToFront
    ws.Shapes("Route_330_221").ZOrder msoBringToFront
    ws.Shapes("Route_330_222").ZOrder msoBringToFront
    ws.Shapes("Route_330_223").ZOrder msoBringToFront
End Sub

Private Sub DrawPart26(ByVal ws As Worksheet)
    ws.Shapes("Route_330_224").ZOrder msoBringToFront
    ws.Shapes("Route_330_225").ZOrder msoBringToFront
    ws.Shapes("Route_332_226").ZOrder msoBringToFront
    ws.Shapes("Route_332_227").ZOrder msoBringToFront
    ws.Shapes("Route_333_228").ZOrder msoBringToFront
    ws.Shapes("Route_333_229").ZOrder msoBringToFront
    ws.Shapes("Route_334_230").ZOrder msoBringToFront
    ws.Shapes("Route_334_231").ZOrder msoBringToFront
    ws.Shapes("Route_334_232").ZOrder msoBringToFront
    ws.Shapes("Route_334_233").ZOrder msoBringToFront
    ws.Shapes("Route_334_234").ZOrder msoBringToFront
    ws.Shapes("Route_334_235").ZOrder msoBringToFront
    ws.Shapes("Route_334_236").ZOrder msoBringToFront
    ws.Shapes("Route_334_237").ZOrder msoBringToFront
    ws.Shapes("Route_334_238").ZOrder msoBringToFront
    ws.Shapes("Route_334_239").ZOrder msoBringToFront
    ws.Shapes("Route_336_240").ZOrder msoBringToFront
    ws.Shapes("Route_336_241").ZOrder msoBringToFront
    ws.Shapes("Route_336_242").ZOrder msoBringToFront
    ws.Shapes("Route_336_243").ZOrder msoBringToFront
    ws.Shapes("Route_336_244").ZOrder msoBringToFront
    ws.Shapes("Route_336_245").ZOrder msoBringToFront
    ws.Shapes("Route_336_246").ZOrder msoBringToFront
    ws.Shapes("Route_336_247").ZOrder msoBringToFront
    ws.Shapes("Route_336_248").ZOrder msoBringToFront
    ws.Shapes("Route_336_249").ZOrder msoBringToFront
    ws.Shapes("Route_338_250").ZOrder msoBringToFront
    ws.Shapes("Route_338_251").ZOrder msoBringToFront
    ws.Shapes("Route_339_252").ZOrder msoBringToFront
    ws.Shapes("Route_339_253").ZOrder msoBringToFront
    ws.Shapes("Route_339_254").ZOrder msoBringToFront
    ws.Shapes("Route_339_255").ZOrder msoBringToFront
    ws.Shapes("Route_339_256").ZOrder msoBringToFront
    ws.Shapes("Route_339_257").ZOrder msoBringToFront
    ws.Shapes("Route_341_258").ZOrder msoBringToFront
End Sub

Private Sub DrawPart27(ByVal ws As Worksheet)
    ws.Shapes("Route_341_259").ZOrder msoBringToFront
    ws.Shapes("Route_341_260").ZOrder msoBringToFront
    ws.Shapes("Route_341_261").ZOrder msoBringToFront
    ws.Shapes("Route_341_262").ZOrder msoBringToFront
    ws.Shapes("Route_341_263").ZOrder msoBringToFront
    ws.Shapes("Route_343_264").ZOrder msoBringToFront
    ws.Shapes("Route_343_265").ZOrder msoBringToFront
    ws.Shapes("Route_343_266").ZOrder msoBringToFront
    ws.Shapes("Route_343_267").ZOrder msoBringToFront
    ws.Shapes("Route_343_268").ZOrder msoBringToFront
    ws.Shapes("Route_343_269").ZOrder msoBringToFront
    ws.Shapes("Route_343_270").ZOrder msoBringToFront
    ws.Shapes("Route_343_271").ZOrder msoBringToFront
    ws.Shapes("Route_345_272").ZOrder msoBringToFront
    ws.Shapes("Route_345_273").ZOrder msoBringToFront
    ws.Shapes("Route_345_274").ZOrder msoBringToFront
    ws.Shapes("Route_345_275").ZOrder msoBringToFront
    ws.Shapes("Route_346_276").ZOrder msoBringToFront
    ws.Shapes("Route_346_277").ZOrder msoBringToFront
    ws.Shapes("Route_346_278").ZOrder msoBringToFront
    ws.Shapes("Route_346_279").ZOrder msoBringToFront
    ws.Shapes("Route_346_280").ZOrder msoBringToFront
    ws.Shapes("Route_346_281").ZOrder msoBringToFront
    ws.Shapes("Route_346_282").ZOrder msoBringToFront
    ws.Shapes("Route_346_283").ZOrder msoBringToFront
    ws.Shapes("Route_348_284").ZOrder msoBringToFront
    ws.Shapes("Route_348_285").ZOrder msoBringToFront
    ws.Shapes("Route_348_286").ZOrder msoBringToFront
    ws.Shapes("Route_348_287").ZOrder msoBringToFront
    ws.Shapes("Route_348_288").ZOrder msoBringToFront
    ws.Shapes("Route_348_289").ZOrder msoBringToFront
    ws.Shapes("Route_348_290").ZOrder msoBringToFront
    ws.Shapes("Route_348_291").ZOrder msoBringToFront
    ws.Shapes("Route_350_292").ZOrder msoBringToFront
    ws.Shapes("Route_350_293").ZOrder msoBringToFront
End Sub

Private Sub DrawPart28(ByVal ws As Worksheet)
    ws.Shapes("Route_350_294").ZOrder msoBringToFront
    ws.Shapes("Route_350_295").ZOrder msoBringToFront
    ws.Shapes("Route_350_296").ZOrder msoBringToFront
    ws.Shapes("Route_350_297").ZOrder msoBringToFront
    ws.Shapes("Route_352_298").ZOrder msoBringToFront
    ws.Shapes("Route_352_299").ZOrder msoBringToFront
    ws.Shapes("Route_352_300").ZOrder msoBringToFront
    ws.Shapes("Route_352_301").ZOrder msoBringToFront
    ws.Shapes("Route_352_302").ZOrder msoBringToFront
    ws.Shapes("Route_352_303").ZOrder msoBringToFront
    ws.Shapes("Route_354_304").ZOrder msoBringToFront
    ws.Shapes("Route_354_305").ZOrder msoBringToFront
    ws.Shapes("Route_354_306").ZOrder msoBringToFront
    ws.Shapes("Route_354_307").ZOrder msoBringToFront
    ws.Shapes("Route_354_308").ZOrder msoBringToFront
    ws.Shapes("Route_354_309").ZOrder msoBringToFront
    ws.Shapes("Route_356_310").ZOrder msoBringToFront
    ws.Shapes("Route_356_311").ZOrder msoBringToFront
    ws.Shapes("Route_356_312").ZOrder msoBringToFront
    ws.Shapes("Route_356_313").ZOrder msoBringToFront
    ws.Shapes("Route_356_314").ZOrder msoBringToFront
    ws.Shapes("Route_356_315").ZOrder msoBringToFront
    ws.Shapes("Route_356_316").ZOrder msoBringToFront
    ws.Shapes("Route_356_317").ZOrder msoBringToFront
    ws.Shapes("Route_358_318").ZOrder msoBringToFront
    ws.Shapes("Route_358_319").ZOrder msoBringToFront
    ws.Shapes("Route_358_320").ZOrder msoBringToFront
    ws.Shapes("Route_358_321").ZOrder msoBringToFront
    ws.Shapes("Route_358_322").ZOrder msoBringToFront
    ws.Shapes("Route_358_323").ZOrder msoBringToFront
    ws.Shapes("Route_360_324").ZOrder msoBringToFront
    ws.Shapes("Route_360_325").ZOrder msoBringToFront
    ws.Shapes("Route_360_326").ZOrder msoBringToFront
    ws.Shapes("Route_361_327").ZOrder msoBringToFront
    ws.Shapes("Route_361_328").ZOrder msoBringToFront
End Sub

Private Sub DrawPart29(ByVal ws As Worksheet)
    ws.Shapes("Route_362_329").ZOrder msoBringToFront
    ws.Shapes("Route_362_330").ZOrder msoBringToFront
    ws.Shapes("Route_362_331").ZOrder msoBringToFront
    ws.Shapes("Route_362_332").ZOrder msoBringToFront
    ws.Shapes("Route_362_333").ZOrder msoBringToFront
    ws.Shapes("Route_364_334").ZOrder msoBringToFront
    ws.Shapes("Route_364_335").ZOrder msoBringToFront
    ws.Shapes("Route_364_336").ZOrder msoBringToFront
    ws.Shapes("Route_365_337").ZOrder msoBringToFront
    ws.Shapes("Route_365_338").ZOrder msoBringToFront
    ws.Shapes("Route_365_339").ZOrder msoBringToFront
    ws.Shapes("Route_365_340").ZOrder msoBringToFront
    ws.Shapes("Route_365_341").ZOrder msoBringToFront
    ws.Shapes("Route_365_342").ZOrder msoBringToFront
    ws.Shapes("Route_367_343").ZOrder msoBringToFront
    ws.Shapes("Route_367_344").ZOrder msoBringToFront
    ws.Shapes("Route_369_345").ZOrder msoBringToFront
    ws.Shapes("Route_369_346").ZOrder msoBringToFront
    ws.Shapes("Route_373_347").ZOrder msoBringToFront
    ws.Shapes("Route_373_348").ZOrder msoBringToFront
    ws.Shapes("Route_374_349").ZOrder msoBringToFront
    ws.Shapes("Route_374_350").ZOrder msoBringToFront
    ws.Shapes("Route_375_351").ZOrder msoBringToFront
    ws.Shapes("Route_375_352").ZOrder msoBringToFront
    ws.Shapes("Route_376_353").ZOrder msoBringToFront
    ws.Shapes("Route_376_354").ZOrder msoBringToFront
    ws.Shapes("Route_377_355").ZOrder msoBringToFront
    ws.Shapes("Route_377_356").ZOrder msoBringToFront
    ws.Shapes("Route_377_357").ZOrder msoBringToFront
    ws.Shapes("Route_377_358").ZOrder msoBringToFront
    ws.Shapes("Route_378_359").ZOrder msoBringToFront
    ws.Shapes("Route_379_360").ZOrder msoBringToFront
    ws.Shapes("Route_379_361").ZOrder msoBringToFront
    ws.Shapes("Route_380_362").ZOrder msoBringToFront
    ws.Shapes("Route_381_363").ZOrder msoBringToFront
End Sub

Private Sub DrawPart30(ByVal ws As Worksheet)
    ws.Shapes("Route_303_364").ZOrder msoBringToFront
    ws.Shapes("Route_304_365").ZOrder msoBringToFront
    ws.Shapes("Route_305_366").ZOrder msoBringToFront
    ws.Shapes("Route_311_367").ZOrder msoBringToFront
    ws.Shapes("Route_314_368").ZOrder msoBringToFront
    ws.Shapes("Route_320_369").ZOrder msoBringToFront
    ws.Shapes("Route_332_370").ZOrder msoBringToFront
    ws.Shapes("Route_333_371").ZOrder msoBringToFront
    ws.Shapes("Route_338_372").ZOrder msoBringToFront
    ws.Shapes("Route_345_373").ZOrder msoBringToFront
    ws.Shapes("Route_361_374").ZOrder msoBringToFront
    ws.Shapes("Route_367_375").ZOrder msoBringToFront
    ws.Shapes("Route_368_376").ZOrder msoBringToFront
    ws.Shapes("Route_369_377").ZOrder msoBringToFront
    ws.Shapes("Route_370_378").ZOrder msoBringToFront
    ws.Shapes("Route_371_379").ZOrder msoBringToFront
    ws.Shapes("Route_372_380").ZOrder msoBringToFront
    ws.Shapes("Route_373_381").ZOrder msoBringToFront
    ws.Shapes("Route_374_382").ZOrder msoBringToFront
    ws.Shapes("Route_375_383").ZOrder msoBringToFront
    ws.Shapes("Route_376_384").ZOrder msoBringToFront
    ws.Shapes("Route_379_385").ZOrder msoBringToFront
    ws.Shapes("Route_379_386").ZOrder msoBringToFront
    ws.Shapes("Route_379_387").ZOrder msoBringToFront
    ws.Shapes("Route_380_388").ZOrder msoBringToFront
    ws.Shapes("Route_380_389").ZOrder msoBringToFront
    ws.Shapes("Text_098").ZOrder msoBringToFront
    ws.Shapes("Text_099").ZOrder msoBringToFront
    ws.Shapes("Text_100").ZOrder msoBringToFront
    ws.Shapes("Text_101").ZOrder msoBringToFront
    ws.Shapes("Text_102").ZOrder msoBringToFront
    ws.Shapes("Text_103").ZOrder msoBringToFront
    ws.Shapes("Text_104").ZOrder msoBringToFront
    ws.Shapes("Text_105").ZOrder msoBringToFront
    ws.Shapes("Text_106").ZOrder msoBringToFront
End Sub

Private Sub DrawPart31(ByVal ws As Worksheet)
    ws.Shapes("Text_107").ZOrder msoBringToFront
    ws.Shapes("Text_108").ZOrder msoBringToFront
    ws.Shapes("Text_109").ZOrder msoBringToFront
    ws.Shapes("Text_110").ZOrder msoBringToFront
    ws.Shapes("Text_111").ZOrder msoBringToFront
    ws.Shapes("Text_112").ZOrder msoBringToFront
    ws.Shapes("Text_113").ZOrder msoBringToFront
    ws.Shapes("Text_114").ZOrder msoBringToFront
    ws.Shapes("Text_115").ZOrder msoBringToFront
    ws.Shapes("Text_116").ZOrder msoBringToFront
    ws.Shapes("Text_117").ZOrder msoBringToFront
    ws.Shapes("Text_118").ZOrder msoBringToFront
    ws.Shapes("Text_119").ZOrder msoBringToFront
    ws.Shapes("Text_120").ZOrder msoBringToFront
    ws.Shapes("Text_121").ZOrder msoBringToFront
End Sub

' Changes captured from New Microsoft Excel Worksheet.xlsm, 2026-11-09.
' This routine uses the saved positions and styles; it does not open the source.
Private Sub ApplySavedWorkbookChanges(ByVal ws As Worksheet)
    Dim s As Shape
    ws.Shapes("Box_9").Line.ForeColor.RGB = RGB(8, 127, 114)
    Set s = ws.Shapes("Route_305_366").Duplicate
    s.Name = "Revision_Allocation_Link"
    SetSavedBounds s, 499.88504, 1348.349976, 81.11496, 3.599921
    SetStyle s, -1, 8739893, 1.274016
    s.Shadow.Visible = msoFalse
    s.Placement = xlFreeFloating
    s.ZOrder msoBringToFront

    Set s = ws.Shapes("Route_305_152").Duplicate
    s.Name = "Revision_Allocation_Arrow"
    SetSavedBounds s, 580.634277, 1345.252563, 5.850079, 6.750079
    SetStyle s, 8739893, -1, 1
    s.Shadow.Visible = msoFalse
    s.Placement = xlFreeFloating
    s.ZOrder msoBringToFront

    Set s = ws.Shapes("Route_348_285").Duplicate
    s.Name = "Revision_CA_Record_Return_Elbow"
    SetSavedBounds s, 2716.214844, 609.700684, 15.431496, 64.933464
    SetStyle s, 7503624, 7503624, 1
    s.Shadow.Visible = msoFalse
    s.Placement = xlFreeFloating
    s.ZOrder msoBringToFront

    Set s = ws.Shapes.AddConnector(msoConnectorStraight, _
        2731.250122 * DRAW_SCALE, 674.642853 * DRAW_SCALE, 565.785767 * DRAW_SCALE, 674.285767 * DRAW_SCALE)
    s.Name = "Revision_CA_Record_Return_Line"
    SetStyle s, -1, 7503624, 1
    s.Shadow.Visible = msoFalse
    s.Placement = xlMoveAndSize
    s.Line.BeginArrowheadStyle = 1
    s.Line.EndArrowheadStyle = 1
    s.Line.BeginArrowheadLength = msoArrowheadLengthMedium
    s.Line.EndArrowheadLength = msoArrowheadLengthMedium
    s.Line.BeginArrowheadWidth = msoArrowheadWidthMedium
    s.Line.EndArrowheadWidth = msoArrowheadWidthMedium
    s.ZOrder msoBringToFront

    Set s = ws.Shapes.AddConnector(msoConnectorStraight, _
        566.954468 * DRAW_SCALE, 674.54541 * DRAW_SCALE, 566.954468 * DRAW_SCALE, 1337.727234 * DRAW_SCALE)
    s.Name = "Revision_CA_Record_To_Allocation"
    SetStyle s, -1, 7503624, 1.5
    s.Shadow.Visible = msoFalse
    s.Placement = xlMoveAndSize
    s.Line.BeginArrowheadStyle = 1
    s.Line.EndArrowheadStyle = 2
    s.Line.BeginArrowheadLength = msoArrowheadLengthMedium
    s.Line.EndArrowheadLength = msoArrowheadLengthMedium
    s.Line.BeginArrowheadWidth = msoArrowheadWidthMedium
    s.Line.EndArrowheadWidth = msoArrowheadWidthMedium
    s.ZOrder msoBringToFront

    Set s = ws.Shapes("Text_099").Duplicate
    s.Name = "Revision_Yes_Allocation"
    SetSavedBounds s, 530.940918, 1350.535156, 44.067165, 23.454567
    SetStyle s, 16777215, -1, 0.75
    s.Shadow.Visible = msoFalse
    s.Placement = xlFreeFloating
    s.TextFrame2.TextRange.Font.Size = 11 * DRAW_SCALE
    s.ZOrder msoBringToFront

End Sub

Private Sub SetSavedBounds(ByVal s As Shape, ByVal x As Double, ByVal y As Double, _
                           ByVal w As Double, ByVal h As Double)
    s.LockAspectRatio = msoFalse
    s.Width = w * DRAW_SCALE
    s.Height = h * DRAW_SCALE
    s.Left = x * DRAW_SCALE
    s.Top = y * DRAW_SCALE
End Sub
