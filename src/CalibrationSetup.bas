Attribute VB_Name = "CalibrationSetup"
Option Explicit

Private Const SHEET_NAME  As String = "Kalibrierung"
Private Const TABLE_X     As String = "Kalibrierung_X"
Private Const TABLE_Y     As String = "Kalibrierung_Y"
Private Const COL_X_START As Integer = 1   ' Column A
Private Const COL_Y_START As Integer = 5   ' Column E  (gap of one empty column)

Public Sub CreateCalibrationSheet()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_NAME
    Else
        ws.Cells.Clear
    End If

    Call BuildAxisTable(ws, TABLE_X, COL_X_START, "X-Achse")
    Call BuildAxisTable(ws, TABLE_Y, COL_Y_START, "Y-Achse")

    ' Freeze top row so headers stay visible while scrolling
    ws.Activate
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(2, COL_X_START).Select

    MsgBox "Kalibrierungssheet '" & SHEET_NAME & "' wurde erstellt.", vbInformation
End Sub

Private Sub BuildAxisTable(ws As Worksheet, tableName As String, _
                            startCol As Integer, axisLabel As String)
    Const HDR_ROW As Integer = 1

    ' Section label above the table
    With ws.Cells(HDR_ROW, startCol)
        .Value = axisLabel
        .Font.Bold = True
        .Font.Size = 11
    End With

    ' Column headers in row 2
    ws.Cells(HDR_ROW + 1, startCol).Value = "Bezeichnung"
    ws.Cells(HDR_ROW + 1, startCol + 1).Value = "Pixelposition"
    ws.Cells(HDR_ROW + 1, startCol + 2).Value = "Abstand (m)"

    ' Create Excel table starting at header row
    Dim tblRange As Range
    Set tblRange = ws.Range( _
        ws.Cells(HDR_ROW + 1, startCol), _
        ws.Cells(HDR_ROW + 1, startCol + 2))

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=tblRange, _
        XlListObjectHasHeaders:=xlYes)
    tbl.Name = tableName
    tbl.TableStyle = "TableStyleLight9"

    ' Column widths
    ws.Columns(startCol).ColumnWidth = 16       ' Bezeichnung
    ws.Columns(startCol + 1).ColumnWidth = 16   ' Pixelposition
    ws.Columns(startCol + 2).ColumnWidth = 14   ' Abstand (m)

    ' Number format for pixel and meter columns (applied to data rows via column)
    ws.Columns(startCol + 1).NumberFormat = "0"
    ws.Columns(startCol + 2).NumberFormat = "0.00"
End Sub
