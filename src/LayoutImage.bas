Attribute VB_Name = "LayoutImage"
Option Explicit

Public Sub RunImportLayoutImage()
    Dim layoutName As String
    layoutName = InputBox( _
        "Layoutname (muss mit Kalibrierungstabelle " & Chr(252) & "bereinstimmen):", _
        "Layout importieren")
    If Trim(layoutName) = "" Then Exit Sub

    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        FileFilter:="Bilder (*.jpg; *.jpeg; *.png),*.jpg;*.jpeg;*.png", _
        Title:="Layoutbild ausw" & Chr(228) & "hlen")
    If filePath = False Then Exit Sub

    ImportLayoutImage Trim(layoutName), CStr(filePath)
End Sub

' Embeds an image on a dedicated sheet named after the layout.
' Sheet name = sanitized layout name (matched against calibration table).
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
        ' Remove any previously embedded layout image
        Dim shp As Shape
        For Each shp In ws.Shapes
            If shp.Type = msoPicture Then shp.Delete
        Next shp
    End If

    ' Embed image (Width/Height = -1 preserves original dimensions)
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

    ws.Activate
    MsgBox "Layout '" & layoutName & "' wurde importiert.", vbInformation
End Sub

' Returns the sheet for a given layout name, or Nothing if not found.
Public Function GetLayoutSheet(layoutName As String) As Worksheet
    Dim sheetName As String
    sheetName = SanitizeSheetName(layoutName)
    On Error Resume Next
    Set GetLayoutSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
End Function

' Returns the embedded picture shape for a layout sheet, or Nothing if absent.
Public Function GetLayoutPicture(layoutName As String) As Shape
    Dim ws As Worksheet
    Set ws = GetLayoutSheet(layoutName)
    If ws Is Nothing Then Exit Function

    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Then
            Set GetLayoutPicture = shp
            Exit Function
        End If
    Next shp
End Function

' Sanitizes a string for use as an Excel sheet name (max 31 chars, no special chars).
Private Function SanitizeSheetName(name As String) As String
    Dim result As String
    result = Left(Trim(name), 31)
    Dim invalidChars As Variant
    invalidChars = Array("\", "/", "?", "*", "[", "]", ":")
    Dim c As Variant
    For Each c In invalidChars
        result = Join(Split(result, CStr(c)), "_")
    Next c
    SanitizeSheetName = result
End Function
