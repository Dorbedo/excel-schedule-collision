Attribute VB_Name = "Main"
Option Explicit

Public Sub RunSetup()
    TableSetup.CreateTaskTable
    CalibrationSetup.CreateCalibrationSheet
End Sub

Public Sub RunImportLayout()
    LayoutImage.RunImportLayoutImage
End Sub

Public Sub RunPlaceCalibrationRect()
    Dim layoutName As String
    layoutName = InputBox("Layoutname:", "Kalibrierung platzieren")
    If Trim(layoutName) = "" Then Exit Sub
    Calibration.PlaceCalibrationRect Trim(layoutName)
End Sub

Public Sub RunApplyCalibrationRect()
    Dim layoutName As String
    layoutName = InputBox("Layoutname:", "Kalibrierung " & Chr(252) & "bernehmen")
    If Trim(layoutName) = "" Then Exit Sub
    Calibration.ApplyCalibrationRect Trim(layoutName)
End Sub

Public Sub RunColorProcessing()
    ColorHelper.ProcessColorColumn
End Sub

Public Sub RunCollisionCheck()
    CollisionChecker.CheckAll
End Sub
