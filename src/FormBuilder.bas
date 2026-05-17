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

    ' Add controls — only names assigned here; all properties set in Initialize
    ' (Controls.Add returns the base Control interface via late binding,
    '  type-specific properties like Style/Default/Cancel must be set at runtime)
    Dim frm As Object
    Set frm = vbc.Designer
    Dim tmp As Object
    Set tmp = frm.Controls.Add("Forms.ComboBox.1", "cboLayout", True)
    Set tmp = frm.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    Set tmp = frm.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    Set tmp = Nothing

    ' Inject all code; Initialize handles form + control sizing and type-specific props
    Dim c As String
    Dim q As String
    q = Chr(34)

    c = "Private mResult As String" & vbLf
    c = c & vbLf
    c = c & "Private Sub UserForm_Initialize()" & vbLf
    c = c & "    Me.Caption = " & q & "Layout ausw" & q & " & Chr(228) & " & q & "hlen" & q & vbLf
    c = c & "    Me.Width = 240: Me.Height = 110" & vbLf
    c = c & "    Me.StartUpPosition = 1" & vbLf
    c = c & "    With cboLayout" & vbLf
    c = c & "        .Left = 8: .Top = 8: .Width = 214: .Height = 20" & vbLf
    c = c & "        .Style = 2" & vbLf
    c = c & "    End With" & vbLf
    c = c & "    With btnOK" & vbLf
    c = c & "        .Caption = " & q & "OK" & q & vbLf
    c = c & "        .Left = 106: .Top = 40: .Width = 56: .Height = 20" & vbLf
    c = c & "        .Default = True" & vbLf
    c = c & "    End With" & vbLf
    c = c & "    With btnCancel" & vbLf
    c = c & "        .Caption = " & q & "Abbrechen" & q & vbLf
    c = c & "        .Left = 166: .Top = 40: .Width = 60: .Height = 20" & vbLf
    c = c & "        .Cancel = True" & vbLf
    c = c & "    End With" & vbLf
    c = c & "End Sub" & vbLf
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
