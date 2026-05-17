Attribute VB_Name = "TableSetup"
Option Explicit

Private Const SHEET_NAME As String = "Vorgangsliste"

Private Const COL_ID As Integer = 1
Private Const COL_LAYOUT As Integer = 2
Private Const COL_NAME As Integer = 3
Private Const COL_DURATION As Integer = 4
Private Const COL_START As Integer = 5
Private Const COL_END As Integer = 6
Private Const COL_PREDECESSOR As Integer = 7
Private Const COL_AREA As Integer = 8
Private Const COL_OWNER As Integer = 9
Private Const COL_STATUS As Integer = 10

Public Sub CreateTaskTable()
    Dim ws As Worksheet
    Dim tbl As ListObject

    ' Get or create the sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_NAME
    Else
        ws.Cells.Clear
    End If

    ' Write headers (matching MS Project German field names where applicable)
    With ws
        .Cells(1, COL_ID).Value = "ID"
        .Cells(1, COL_LAYOUT).Value = "Layoutname"
        .Cells(1, COL_NAME).Value = "Vorgangsname"
        .Cells(1, COL_DURATION).Value = "Dauer"
        .Cells(1, COL_START).Value = "Anfang"
        .Cells(1, COL_END).Value = "Ende"
        .Cells(1, COL_PREDECESSOR).Value = "Vorgänger"
        .Cells(1, COL_AREA).Value = "Bereich/Fläche"
        .Cells(1, COL_OWNER).Value = "Verantwortlicher"
        .Cells(1, COL_STATUS).Value = "Status"
    End With

    ' Create Excel table
    Set tbl = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range("A1:J1"), _
        XlListObjectHasHeaders:=xlYes _
    )
    tbl.Name = "Vorgangsliste"
    tbl.TableStyle = "TableStyleMedium2"

    ' Column widths
    ws.Columns(COL_ID).ColumnWidth = 5
    ws.Columns(COL_LAYOUT).ColumnWidth = 20
    ws.Columns(COL_NAME).ColumnWidth = 35
    ws.Columns(COL_DURATION).ColumnWidth = 10
    ws.Columns(COL_START).ColumnWidth = 13
    ws.Columns(COL_END).ColumnWidth = 13
    ws.Columns(COL_PREDECESSOR).ColumnWidth = 12
    ws.Columns(COL_AREA).ColumnWidth = 20
    ws.Columns(COL_OWNER).ColumnWidth = 22
    ws.Columns(COL_STATUS).ColumnWidth = 18

    ' Date format for Start and End
    ws.Columns(COL_START).NumberFormat = "DD.MM.YYYY"
    ws.Columns(COL_END).NumberFormat = "DD.MM.YYYY"

    ' Data validation for Status column
    Call AddStatusValidation(ws)

    ws.Activate
    MsgBox "Tabelle '" & SHEET_NAME & "' wurde erfolgreich erstellt.", vbInformation
End Sub

Private Sub AddStatusValidation(ws As Worksheet)
    Dim rng As Range
    ' Apply to data rows (row 2 onward, capped at row 1000)
    Set rng = ws.Range(ws.Cells(2, COL_STATUS), ws.Cells(1000, COL_STATUS))

    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="Geplant,Bestätigt,In Arbeit,Abgeschlossen,Kollision"
        .ShowDropDown = False
        .ErrorMessage = "Bitte einen gültigen Status aus der Liste wählen."
    End With
End Sub
