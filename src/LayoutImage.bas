Attribute VB_Name = "LayoutImage"
Option Explicit

Private Const CAL_SHEET As String = "Kalibrierung"
Private Const TABLE_X   As String = "Kalibrierung_X"
Private Const TABLE_Y   As String = "Kalibrierung_Y"

' ── Main workflow ─────────────────────────────────────────────────────────────

Public Sub RunImportLayoutImage()
    ' Step 1+2: File dialog
    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        FileFilter:="Bilder (*.jpg; *.jpeg; *.png),*.jpg;*.jpeg;*.png", _
        Title:="Layoutbild ausw" & Chr(228) & "hlen")
    If filePath = False Then Exit Sub

    ' Step 3: Layout name — default = filename without extension
    Dim layoutName As String
    layoutName = InputBox( _
        "Layoutbezeichnung anpassen:", _
        "Layout importieren", _
        FileBaseName(CStr(filePath)))
    If Trim(layoutName) = "" Then Exit Sub
    layoutName = Trim(layoutName)

    ' Embed image on dedicated sheet
    ImportLayoutImage layoutName, CStr(filePath)

    ' Step 4: If layout unknown in calibration table → collect 2 reference points
    If Not LayoutExistsInCalibration(layoutName) Then
        Dim ok As Boolean
        ok = CollectAndAddCalibrationPoints(layoutName)
        If Not ok Then
            MsgBox "Kalibrierung abgebrochen. Bitte Kalibrierungspunkte sp" & _
                   Chr(228) & "ter manuell eintragen.", vbExclamation
        End If
    End If

    ' Step 6: Navigate to calibration sheet for review
    StartCalibration layoutName
End Sub

' ── Image embedding ───────────────────────────────────────────────────────────

Public Sub ImportLayoutImage(layoutName As String, imagePath As String)
    Dim sheetName As String
    sheetName = SanitizeSheetName(layoutName)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    Else
        Dim shp As Shape
        For Each shp In ws.Shapes
            If shp.Type = msoPicture Then shp.Delete
        Next shp
    End If

    Dim pic As Shape
    Set pic = ws.Shapes.AddPicture( _
        Filename:=imagePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=0, Top:=0, _
        Width:=-1, Height:=-1)
    pic.Name = "Bild_" & sheetName
    pic.LockAspectRatio = msoTrue
    pic.Placement = xlFreeFloating
End Sub

' ── Calibration point collection ──────────────────────────────────────────────

' Asks the user for 2 reference points (X + Y per point) and adds them to the
' calibration tables. Returns False if the user cancels or input is invalid.
Private Function CollectAndAddCalibrationPoints(layoutName As String) As Boolean
    Dim xLabel(1) As String, xPixel(1) As Long, xMeter(1) As Double
    Dim yLabel(1) As String, yPixel(1) As Long, yMeter(1) As Double

    Dim i As Integer
    For i = 0 To 1
        Dim num As String
        num = CStr(i + 1)

        ' X axis input
        Dim xRaw As String
        xRaw = InputBox( _
            "Koordinationspunkt " & num & " - X-Achse" & Chr(10) & _
            "Format: Bezeichnung, Pixel, Abstand(m)" & Chr(10) & _
            "Beispiel:  A, 120, 0.00", _
            "Kalibrierung Punkt " & num & "/2")
        If Trim(xRaw) = "" Then Exit Function

        If Not ParseCalibInput(xRaw, xLabel(i), xPixel(i), xMeter(i)) Then
            MsgBox "Ung" & Chr(252) & "ltige Eingabe X-Achse Punkt " & num, vbExclamation
            Exit Function
        End If

        ' Y axis input
        Dim yRaw As String
        yRaw = InputBox( _
            "Koordinationspunkt " & num & " - Y-Achse" & Chr(10) & _
            "Format: Bezeichnung, Pixel, Abstand(m)" & Chr(10) & _
            "Beispiel:  1, 80, 0.00", _
            "Kalibrierung Punkt " & num & "/2")
        If Trim(yRaw) = "" Then Exit Function

        If Not ParseCalibInput(yRaw, yLabel(i), yPixel(i), yMeter(i)) Then
            MsgBox "Ung" & Chr(252) & "ltige Eingabe Y-Achse Punkt " & num, vbExclamation
            Exit Function
        End If
    Next i

    ' Step 5: Write to calibration tables
    AddCalibrationRow TABLE_X, layoutName, xLabel(0), xPixel(0), xMeter(0)
    AddCalibrationRow TABLE_X, layoutName, xLabel(1), xPixel(1), xMeter(1)
    AddCalibrationRow TABLE_Y, layoutName, yLabel(0), yPixel(0), yMeter(0)
    AddCalibrationRow TABLE_Y, layoutName, yLabel(1), yPixel(1), yMeter(1)

    CollectAndAddCalibrationPoints = True
End Function

Private Function ParseCalibInput(raw As String, ByRef lbl As String, _
                                  ByRef px As Long, ByRef m As Double) As Boolean
    On Error GoTo ParseErr
    Dim parts() As String
    parts = Split(raw, ",")
    If UBound(parts) <> 2 Then Exit Function
    lbl = Trim(parts(0))
    px  = CLng(Trim(parts(1)))
    m   = CDbl(Trim(parts(2)))
    ParseCalibInput = True
    Exit Function
ParseErr:
End Function

Private Sub AddCalibrationRow(tableName As String, layoutName As String, _
                               lbl As String, px As Long, m As Double)
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

    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    newRow.Range(1).Value = layoutName
    newRow.Range(2).Value = lbl
    newRow.Range(3).Value = px
    newRow.Range(4).Value = m
End Sub

' ── Calibration start ─────────────────────────────────────────────────────────

' Navigates to the calibration sheet and highlights the rows for this layout.
' Interpolation logic is a TODO — this currently serves as review step.
Public Sub StartCalibration(layoutName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CAL_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ws.Activate

    ' Select first matching row in Kalibrierung_X for visual feedback
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(TABLE_X)
    On Error GoTo 0
    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then
            Dim cell As Range
            For Each cell In tbl.ListColumns("Layout").DataBodyRange
                If cell.Value = layoutName Then
                    cell.Select
                    Exit For
                End If
            Next cell
        End If
    End If

    MsgBox "Layout '" & layoutName & "' importiert." & Chr(10) & _
           "Kalibrierungspunkte bitte pr" & Chr(252) & "fen und erg" & Chr(228) & "nzen.", _
           vbInformation
End Sub

' ── Helpers ───────────────────────────────────────────────────────────────────

Public Function GetLayoutSheet(layoutName As String) As Worksheet
    On Error Resume Next
    Set GetLayoutSheet = ThisWorkbook.Sheets(SanitizeSheetName(layoutName))
    On Error GoTo 0
End Function

Public Function GetLayoutPicture(layoutName As String) As Shape
    Dim ws As Worksheet
    Set ws = GetLayoutSheet(layoutName)
    If ws Is Nothing Then Exit Function
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Then Set GetLayoutPicture = shp: Exit Function
    Next shp
End Function

Private Function LayoutExistsInCalibration(layoutName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CAL_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(TABLE_X)
    On Error GoTo 0
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then Exit Function

    Dim cell As Range
    For Each cell In tbl.ListColumns("Layout").DataBodyRange
        If cell.Value = layoutName Then LayoutExistsInCalibration = True: Exit Function
    Next cell
End Function

Private Function FileBaseName(filePath As String) As String
    Dim name As String
    name = Mid(filePath, InStrRev(filePath, "\") + 1)
    Dim dot As Integer
    dot = InStrRev(name, ".")
    If dot > 0 Then name = Left(name, dot - 1)
    FileBaseName = name
End Function

Private Function SanitizeSheetName(name As String) As String
    Dim result As String
    result = Left(Trim(name), 31)
    Dim ch As Variant
    For Each ch In Array("\", "/", "?", "*", "[", "]", ":")
        result = Join(Split(result, CStr(ch)), "_")
    Next ch
    SanitizeSheetName = result
End Function
