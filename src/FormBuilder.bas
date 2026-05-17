Attribute VB_Name = "FormBuilder"
Option Explicit

' Creates frmLayoutSelector programmatically in this workbook's VBA project.
' Run once after initial setup. Safe to re-run — existing form is replaced.
'
' Requires: Excel Options -> Trust Center -> Macro Settings ->
'           "Trust access to the VBA project object model" must be enabled.
Public Sub BuildLayoutSelectorForm()
    Dim vbp As Object
    On Error GoTo NoTrust
    Set vbp = ThisWorkbook.VBProject
    On Error GoTo 0

    ' Remove existing form
    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents("frmLayoutSelector")
    On Error GoTo 0

    ' Add UserForm (3 = vbext_ct_MSForm)
    Dim vbc As Object
    Set vbc = vbp.VBComponents.Add(3)
    vbc.Name = "frmLayoutSelector"

    ' Configure form dimensions and caption
    Dim frm As Object
    Set frm = vbc.Designer
    frm.Caption = "Layout ausw" & Chr(228) & "hlen"
    frm.Width = 240
    frm.Height = 110
    frm.StartUpPosition = 1  ' CenterOwner

    ' ComboBox — drop-down list only (Style = 2)
    Dim cbo As Object
    Set cbo = frm.Controls.Add("Forms.ComboBox.1", "cboLayout", True)
    cbo.Left = 8:  cbo.Top = 8
    cbo.Width = 214: cbo.Height = 20
    cbo.Style = 2

    ' OK button
    Dim btnOK As Object
    Set btnOK = frm.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    btnOK.Caption = "OK"
    btnOK.Left = 106: btnOK.Top = 40
    btnOK.Width = 56:  btnOK.Height = 20
    btnOK.Default = True

    ' Cancel button
    Dim btnCancel As Object
    Set btnCancel = frm.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    btnCancel.Caption = "Abbrechen"
    btnCancel.Left = 166: btnCancel.Top = 40
    btnCancel.Width = 60:  btnCancel.Height = 20
    btnCancel.Cancel = True

    ' Inject form event code (built line by line to avoid line-continuation limit)
    Dim c As String
    Dim q As String
    q = Chr(34)

    c = "Private mResult As String" & vbLf
    c = c & vbLf
    c = c & "Public Property Get SelectedLayout() As String" & vbLf
    c = c & "    SelectedLayout = mResult" & vbLf
    c = c & "End Property" & vbLf
    c = c & vbLf
    c = c & "Private Sub UserForm_Activate()" & vbLf
    c = c & "    mResult = " & q & q & vbLf
    c = c & "End Sub" & vbLf
    c = c & vbLf
    c = c & "Private Sub btnOK_Click()" & vbLf
    c = c & "    If cboLayout.ListIndex < 0 Then" & vbLf
    c = c & "        MsgBox " & q & "Bitte ein Layout ausw" & q & " & Chr(228) & " & q & "hlen." & q & ", vbExclamation" & vbLf
    c = c & "        Exit Sub" & vbLf
    c = c & "    End If" & vbLf
    c = c & "    mResult = cboLayout.Value" & vbLf
    c = c & "    Me.Hide" & vbLf
    c = c & "End Sub" & vbLf
    c = c & vbLf
    c = c & "Private Sub btnCancel_Click()" & vbLf
    c = c & "    mResult = " & q & q & vbLf
    c = c & "    Me.Hide" & vbLf
    c = c & "End Sub" & vbLf
    c = c & vbLf
    c = c & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbLf
    c = c & "    If CloseMode = vbFormControlMenu Then mResult = " & q & q & vbLf
    c = c & "End Sub"

    vbc.CodeModule.AddFromString c

    MsgBox "frmLayoutSelector wurde erstellt.", vbInformation
    Exit Sub

NoTrust:
    MsgBox "Fehler: Zugriff auf das VBA-Projekt nicht erlaubt." & vbLf & vbLf & _
           "Bitte aktivieren:" & vbLf & _
           "Excel Optionen => Trust Center => Makroeinstellungen" & vbLf & _
           Chr(34) & "Zugriff auf das VBA-Projektobjektmodell vertrauen" & Chr(34), _
           vbCritical
End Sub
