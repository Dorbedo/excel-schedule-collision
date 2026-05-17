Attribute VB_Name = "ColorHelper"
Option Explicit

' Converts HEX string (#FF0000 or FF0000) to Excel color Long. Returns -1 on error.
Public Function HexToLong(hexStr As String) As Long
    hexStr = Trim(hexStr)
    If Left(hexStr, 1) = "#" Then hexStr = Mid(hexStr, 2)
    If Len(hexStr) <> 6 Then HexToLong = -1: Exit Function

    On Error GoTo InvalidHex
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Left(hexStr, 2))
    g = CLng("&H" & Mid(hexStr, 3, 2))
    b = CLng("&H" & Right(hexStr, 2))
    HexToLong = RGB(r, g, b)
    Exit Function
InvalidHex:
    HexToLong = -1
End Function

' Converts RGB string "255,0,0" to Excel color Long. Returns -1 on error.
Public Function RGBStringToLong(rgbStr As String) As Long
    On Error GoTo InvalidRGB
    Dim parts() As String
    parts = Split(Trim(rgbStr), ",")
    If UBound(parts) <> 2 Then RGBStringToLong = -1: Exit Function

    Dim r As Long, g As Long, b As Long
    r = CLng(Trim(parts(0)))
    g = CLng(Trim(parts(1)))
    b = CLng(Trim(parts(2)))

    If r < 0 Or r > 255 Or g < 0 Or g > 255 Or b < 0 Or b > 255 Then
        RGBStringToLong = -1: Exit Function
    End If
    RGBStringToLong = RGB(r, g, b)
    Exit Function
InvalidRGB:
    RGBStringToLong = -1
End Function

' Converts Excel color Long to HEX string (#RRGGBB)
Public Function LongToHex(colorLong As Long) As String
    Dim r As Long, g As Long, b As Long
    r = colorLong Mod 256
    g = (colorLong \ 256) Mod 256
    b = (colorLong \ 65536) Mod 256
    LongToHex = "#" & Right("00" & Hex(r), 2) & Right("00" & Hex(g), 2) & Right("00" & Hex(b), 2)
End Function

' Processes the Farbe column:
'   HEX/RGB value in cell  -> applies fill color to cell
'   Empty cell with fill   -> reads fill color and writes HEX into cell
Public Sub ProcessColorColumn()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Vorgangsliste")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet 'Vorgangsliste' nicht gefunden.", vbExclamation
        Exit Sub
    End If

    Dim tbl As ListObject
    Set tbl = ws.ListObjects("Vorgangsliste")
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then Exit Sub

    Dim colorCol As ListColumn
    On Error Resume Next
    Set colorCol = tbl.ListColumns("Farbe")
    On Error GoTo 0
    If colorCol Is Nothing Then Exit Sub

    Dim cell As Range
    Dim cellText As String
    Dim colorVal As Long

    For Each cell In colorCol.DataBodyRange
        cellText = Trim(cell.Value)

        If cellText = "" Then
            ' Empty cell: read existing fill color if set
            If cell.Interior.ColorIndex <> xlNone Then
                cell.Value = LongToHex(cell.Interior.Color)
            End If
        ElseIf InStr(cellText, ",") > 0 Then
            ' RGB format: "255,0,0"
            colorVal = RGBStringToLong(cellText)
            If colorVal >= 0 Then cell.Interior.Color = colorVal
        Else
            ' HEX format: "#FF0000" or "FF0000"
            colorVal = HexToLong(cellText)
            If colorVal >= 0 Then cell.Interior.Color = colorVal
        End If
    Next cell

    MsgBox "Farben wurden verarbeitet.", vbInformation
End Sub
