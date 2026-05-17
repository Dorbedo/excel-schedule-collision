Attribute VB_Name = "ColorHelper"
Option Explicit

' Parses any supported color string into r, g, b, a components (all 0-255).
' Supported formats:
'   HEX: "#RRGGBB", "RRGGBB", "#AARRGGBB", "AARRGGBB"
'   RGB: "R,G,B" or "R,G,B,A"
' Returns True on success. Alpha defaults to 255 (opaque) when not specified.
Public Function ParseColorString(colorStr As String, _
                                  ByRef r As Long, ByRef g As Long, _
                                  ByRef b As Long, ByRef a As Long) As Boolean
    colorStr = Trim(colorStr)
    r = 0: g = 0: b = 0: a = 255

    On Error GoTo ParseError

    If InStr(colorStr, ",") > 0 Then
        ' RGB / RGBA format: "R,G,B" or "R,G,B,A"
        Dim parts() As String
        parts = Split(colorStr, ",")
        If UBound(parts) < 2 Or UBound(parts) > 3 Then GoTo ParseError
        r = CLng(Trim(parts(0)))
        g = CLng(Trim(parts(1)))
        b = CLng(Trim(parts(2)))
        If UBound(parts) = 3 Then a = CLng(Trim(parts(3)))
    Else
        ' HEX format
        Dim hexStr As String
        hexStr = colorStr
        If Left(hexStr, 1) = "#" Then hexStr = Mid(hexStr, 2)

        Select Case Len(hexStr)
            Case 6  ' RRGGBB
                r = CLng("&H" & Mid(hexStr, 1, 2))
                g = CLng("&H" & Mid(hexStr, 3, 2))
                b = CLng("&H" & Mid(hexStr, 5, 2))
            Case 8  ' AARRGGBB
                a = CLng("&H" & Mid(hexStr, 1, 2))
                r = CLng("&H" & Mid(hexStr, 3, 2))
                g = CLng("&H" & Mid(hexStr, 5, 2))
                b = CLng("&H" & Mid(hexStr, 7, 2))
            Case Else
                GoTo ParseError
        End Select
    End If

    If r < 0 Or r > 255 Or g < 0 Or g > 255 Or _
       b < 0 Or b > 255 Or a < 0 Or a > 255 Then GoTo ParseError

    ParseColorString = True
    Exit Function
ParseError:
    ParseColorString = False
End Function

' Converts Excel color Long (BGR) + explicit alpha to #AARRGGBB string.
' Excel cell fills are always opaque, so alpha from Interior.Color is always FF.
Public Function LongToHex(colorLong As Long) As String
    Dim r As Long, g As Long, b As Long
    r = colorLong And 255
    g = (colorLong \ 256) And 255
    b = (colorLong \ 65536) And 255
    LongToHex = "#FF" & Right("00" & Hex(r), 2) & _
                Right("00" & Hex(g), 2) & _
                Right("00" & Hex(b), 2)
End Function

' Processes the Farbe column:
'   HEX/RGB value in cell  -> applies RGB fill to cell (alpha preserved in text)
'   Empty cell with fill   -> reads fill color, writes #AARRGGBB (alpha=FF)
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
    Dim r As Long, g As Long, b As Long, a As Long

    For Each cell In colorCol.DataBodyRange
        cellText = Trim(cell.Value)

        If cellText = "" Then
            ' Empty cell: read existing fill color and write as #AARRGGBB
            If cell.Interior.ColorIndex <> xlNone Then
                cell.Value = LongToHex(cell.Interior.Color)
            End If
        Else
            ' Parse and apply — alpha is kept in cell text, only RGB applied to fill
            If ParseColorString(cellText, r, g, b, a) Then
                cell.Interior.Color = RGB(r, g, b)
            End If
        End If
    Next cell

    MsgBox "Farben wurden verarbeitet.", vbInformation
End Sub
