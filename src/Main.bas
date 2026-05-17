Attribute VB_Name = "Main"
Option Explicit

Public Sub RunSetup()
    TableSetup.CreateTaskTable
    CalibrationSetup.CreateCalibrationSheet
End Sub

Public Sub RunColorProcessing()
    ColorHelper.ProcessColorColumn
End Sub

Public Sub RunCollisionCheck()
    CollisionChecker.CheckAll
End Sub
