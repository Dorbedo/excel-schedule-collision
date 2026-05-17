Attribute VB_Name = "Calibration"
Option Explicit

Private Const CAL_SHEET    As String = "Kalibrierung"
Private Const TABLE_X      As String = "Kalibrierung_X"
Private Const TABLE_Y      As String = "Kalibrierung_Y"
Private Const RECT_PREFIX  As String = "Kalibrierung_"
Private Const MKR_PREFIX_X As String = "KalMkr_X_"
Private Const MKR_PREFIX_Y As String = "KalMkr_Y_"

' Places a semi-transparent calibration rectangle on the layout sheet.
' The user drags the corners to 2 known grid reference points, then calls
' ApplyCalibrationRect to read the positions and save the calibration.
Public Sub PlaceCalibrationRect(layoutName As String)
    Dim ws As Worksheet
    Set ws = LayoutImage.GetLayoutSheet(layoutName)
    If ws Is Nothing Then
        MsgBox "Layout '" & layoutName & "' nicht gefunden.", vbExclamation
        Exit Sub
    End If

    ' Remove any existing calibration rect for this layout
    Dim shapeName As String
    shapeName = RECT_PREFIX & SanitizeSheetName(layoutName)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    ' Default position: centred at 25%-75% of picture area
    Dim pic As Shape
    Set pic = LayoutImage.GetLayoutPicture(layoutName)

    Dim l As Double, t As Double, w As Double, h As Double
    If pic Is Nothing Then
        l = 50: t = 50: w = 200: h = 150
    Else
        l = pic.Left + pic.Width * 0.25
        t = pic.Top + pic.Height * 0.25
        w = pic.Width * 0.5
        h = pic.Height * 0.5
    End If

    Dim rect As Shape
    Set rect = ws.Shapes.AddShape(msoShapeRectangle, l, t, w, h)
    rect.Name = shapeName

    With rect
        .Fill.Transparency = 0.65
        .Fill.ForeColor.RGB = RGB(0, 120, 215)
        .Line.ForeColor.RGB = RGB(0, 80, 160)
        .Line.Weight = 2
    End With
    With rect.TextFrame2
        .TextRange.Text = "Kalibrierung" & Chr(10) & layoutName
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .VerticalAnchor = msoAnchorMiddle
    End With

    ws.Activate
    rect.Select

    MsgBox "Kalibrierungsrechteck platziert." & Chr(10) & Chr(10) & _
           "Ecken auf 2 bekannte Rasterpunkte ziehen," & Chr(10) & _
           "dann 'Kalibrierung " & Chr(252) & "bernehmen' ausf" & Chr(252) & "hren.", _
           vbInformation
End Sub

' Reads the calibration rectangle's corner positions (in sheet points),
' asks the user for grid labels and real distances for each axis,
' replaces the calibration table rows for this layout and removes the rect.
Public Sub ApplyCalibrationRect(layoutName As String)
    Dim ws As Worksheet
    Set ws = LayoutImage.GetLayoutSheet(layoutName)
    If ws Is Nothing Then
        MsgBox "Layout '" & layoutName & "' nicht gefunden.", vbExclamation
        Exit Sub
    End If

    Dim shapeName As String
    shapeName = RECT_PREFIX & SanitizeSheetName(layoutName)

    Dim rect As Shape
    On Error Resume Next
    Set rect = ws.Shapes(shapeName)
    On Error GoTo 0
    If rect Is Nothing Then
        MsgBox "Kein Kalibrierungsrechteck gefunden." & Chr(10) & _
               "Bitte zuerst 'Kalibrierung platzieren' ausf" & Chr(252) & "hren.", vbExclamation
        Exit Sub
    End If

    ' Corner positions in sheet points (origin = top-left of sheet / image)
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    x1 = rect.Left
    y1 = rect.Top
    x2 = rect.Left + rect.Width
    y2 = rect.Top + rect.Height

    ' Collect axis labels and real distances for both corners
    Dim xLbl1 As String, xM1 As Double, xLbl2 As String, xM2 As Double
    Dim yLbl1 As String, yM1 As Double, yLbl2 As String, yM2 As Double

    If Not AskAxisInput("X", 1, xLbl1, xM1) Then Exit Sub
    If Not AskAxisInput("X", 2, xLbl2, xM2) Then Exit Sub
    If Not AskAxisInput("Y", 1, yLbl1, yM1) Then Exit Sub
    If Not AskAxisInput("Y", 2, yLbl2, yM2) Then Exit Sub

    ' Write to calibration tables (existing rows for this layout are replaced)
    ReplaceCalibrationRows TABLE_X, layoutName, _
        Array(xLbl1, xLbl2), Array(x1, x2), Array(xM1, xM2)
    ReplaceCalibrationRows TABLE_Y, layoutName, _
        Array(yLbl1, yLbl2), Array(y1, y2), Array(yM1, yM2)

    ' Clean up rect and navigate to calibration sheet
    rect.Delete

    Dim calWs As Worksheet
    On Error Resume Next
    Set calWs = ThisWorkbook.Sheets(CAL_SHEET)
    On Error GoTo 0
    If Not calWs Is Nothing Then calWs.Activate

    MsgBox "Kalibrierung f" & Chr(252) & "r '" & layoutName & "' gespeichert.", vbInformation
End Sub

' ── Calibration markers ──────────────────────────────────────────────────────

' Places small labelled boxes on the layout sheet at each calibrated position.
' X markers appear along the top edge of the image, Y markers along the left edge.
Public Sub ShowCalibrationMarkers(layoutName As String)
    Dim ws As Worksheet
    Set ws = LayoutImage.GetLayoutSheet(layoutName)
    If ws Is Nothing Then
        MsgBox "Layout '" & layoutName & "' nicht gefunden.", vbExclamation
        Exit Sub
    End If

    HideCalibrationMarkers layoutName

    Dim calWs As Worksheet
    On Error Resume Next
    Set calWs = ThisWorkbook.Sheets(CAL_SHEET)
    On Error GoTo 0
    If calWs Is Nothing Then Exit Sub

    Dim pic As Shape
    Set pic = LayoutImage.GetLayoutPicture(layoutName)
    Dim imgLeft As Double, imgTop As Double
    If pic Is Nothing Then
        imgLeft = 0: imgTop = 0
    Else
        imgLeft = pic.Left: imgTop = pic.Top
    End If

    Dim tblX As ListObject, tblY As ListObject
    On Error Resume Next
    Set tblX = calWs.ListObjects(TABLE_X)
    Set tblY = calWs.ListObjects(TABLE_Y)
    On Error GoTo 0

    PlaceAxisMarkers ws, tblX, layoutName, True, imgLeft, imgTop
    PlaceAxisMarkers ws, tblY, layoutName, False, imgLeft, imgTop

    ws.Activate
End Sub

' Removes all calibration markers for the given layout from its sheet.
Public Sub HideCalibrationMarkers(layoutName As String)
    Dim ws As Worksheet
    Set ws = LayoutImage.GetLayoutSheet(layoutName)
    If ws Is Nothing Then Exit Sub

    Dim sanName As String
    sanName = SanitizeSheetName(layoutName)

    ' Collect names first — modifying the Shapes collection while iterating causes errors
    Dim names() As String
    ReDim names(ws.Shapes.Count - 1)
    Dim count As Integer
    count = 0

    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, Len(MKR_PREFIX_X & sanName)) = MKR_PREFIX_X & sanName Or _
           Left(shp.Name, Len(MKR_PREFIX_Y & sanName)) = MKR_PREFIX_Y & sanName Then
            names(count) = shp.Name
            count = count + 1
        End If
    Next shp

    Dim k As Integer
    For k = 0 To count - 1
        On Error Resume Next
        ws.Shapes(names(k)).Delete
        On Error GoTo 0
    Next k
End Sub

Private Sub PlaceAxisMarkers(ws As Worksheet, tbl As ListObject, layoutName As String, _
                              isX As Boolean, imgLeft As Double, imgTop As Double)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Const BOX_W  As Double = 22
    Const BOX_H  As Double = 14
    Const MARGIN As Double = 3

    Dim sanName As String
    sanName = SanitizeSheetName(layoutName)

    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim r As ListRow
        Set r = tbl.ListRows(i)
        If r.Range(1).Value <> layoutName Then GoTo NextRow

        Dim lbl As String, pos As Double, m As Double
        lbl = CStr(r.Range(2).Value)
        pos = CDbl(r.Range(3).Value)
        m   = CDbl(r.Range(4).Value)

        Dim l As Double, t As Double
        If isX Then
            l = pos - BOX_W / 2
            t = imgTop + MARGIN
        Else
            l = imgLeft + MARGIN
            t = pos - BOX_H / 2
        End If

        Dim mkr As Shape
        Set mkr = ws.Shapes.AddShape(msoShapeRectangle, l, t, BOX_W, BOX_H)

        If isX Then
            mkr.Name = MKR_PREFIX_X & sanName & "_" & lbl
            mkr.Fill.ForeColor.RGB = RGB(0, 120, 215)
        Else
            mkr.Name = MKR_PREFIX_Y & sanName & "_" & lbl
            mkr.Fill.ForeColor.RGB = RGB(0, 160, 80)
        End If

        With mkr
            .Fill.Transparency = 0.25
            .Line.Visible = msoFalse
            .Placement = xlFreeFloating
        End With
        With mkr.TextFrame2
            .TextRange.Text = lbl & Chr(10) & Format(m, "0.00") & "m"
            .TextRange.Font.Size = 6
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With

NextRow:
    Next i
End Sub

' ── Helpers ───────────────────────────────────────────────────────────────────

Private Function AskAxisInput(axis As String, pointNum As Integer, _
                               ByRef lbl As String, ByRef m As Double) As Boolean
    Dim raw As String
    raw = InputBox( _
        "Ecke " & CStr(pointNum) & " - " & axis & "-Achse" & Chr(10) & _
        "Format:  Bezeichnung, Abstand (m)" & Chr(10) & _
        "Beispiel:  A, 0.00", _
        axis & "-Achse  Punkt " & CStr(pointNum) & " / 2")
    If Trim(raw) = "" Then Exit Function

    On Error GoTo ParseErr
    Dim parts() As String
    parts = Split(raw, ",")
    If UBound(parts) <> 1 Then GoTo ParseErr
    lbl = Trim(parts(0))
    m   = CDbl(Trim(parts(1)))
    AskAxisInput = True
    Exit Function
ParseErr:
    MsgBox "Ung" & Chr(252) & "ltige Eingabe.", vbExclamation
End Function

Private Sub ReplaceCalibrationRows(tableName As String, layoutName As String, _
                                    labels As Variant, positions As Variant, meters As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CAL_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub

    ' Delete existing rows for this layout in reverse to avoid index shifting
    If Not tbl.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = tbl.ListRows.Count To 1 Step -1
            If tbl.ListRows(i).Range(1).Value = layoutName Then
                tbl.ListRows(i).Delete
            End If
        Next i
    End If

    ' Insert fresh rows
    Dim j As Integer
    For j = 0 To UBound(labels)
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add
        newRow.Range(1).Value = layoutName
        newRow.Range(2).Value = labels(j)
        newRow.Range(3).Value = positions(j)
        newRow.Range(4).Value = meters(j)
    Next j
End Sub

Private Function SanitizeSheetName(name As String) As String
    Dim result As String
    result = Left(Trim(name), 31)
    Dim ch As Variant
    For Each ch In Array("\", "/", "?", "*", "[", "]", ":")
        result = Join(Split(result, CStr(ch)), "_")
    Next ch
    SanitizeSheetName = result
End Function
