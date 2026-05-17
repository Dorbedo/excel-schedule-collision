Attribute VB_Name = "Main"
Option Explicit

Public Sub RunSetup()
    TableSetup.CreateTaskTable
    CalibrationSetup.CreateCalibrationSheet
End Sub

' Run once after setup to create the layout selector dropdown form.
Public Sub RunBuildForms()
    FormBuilder.BuildLayoutSelectorForm
End Sub

Public Sub RunImportLayout()
    LayoutImage.RunImportLayoutImage
End Sub

Public Sub RunPlaceCalibrationRect()
    Dim layoutName As String
    layoutName = Calibration.PickLayoutName("Kalibrierung platzieren")
    If layoutName = "" Then Exit Sub
    Calibration.PlaceCalibrationRect layoutName
End Sub

Public Sub RunApplyCalibrationRect()
    Dim layoutName As String
    layoutName = Calibration.PickLayoutName("Kalibrierung " & Chr(252) & "bernehmen")
    If layoutName = "" Then Exit Sub
    Calibration.ApplyCalibrationRect layoutName
End Sub

Public Sub RunShowCalibrationMarkers()
    Dim layoutName As String
    layoutName = Calibration.PickLayoutName("Kalibrierungsmarker anzeigen")
    If layoutName = "" Then Exit Sub
    Calibration.ShowCalibrationMarkers layoutName
End Sub

Public Sub RunHideCalibrationMarkers()
    Dim layoutName As String
    layoutName = Calibration.PickLayoutName("Kalibrierungsmarker ausblenden")
    If layoutName = "" Then Exit Sub
    Calibration.HideCalibrationMarkers layoutName
End Sub

Public Sub RunColorProcessing()
    ColorHelper.ProcessColorColumn
End Sub

Public Sub RunCollisionCheck()
    CollisionChecker.CheckAll
End Sub
