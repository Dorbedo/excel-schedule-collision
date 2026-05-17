Attribute VB_Name = "Main"
Option Explicit

Public Sub RunSetup()
    TableSetup.CreateTaskTable
End Sub

Public Sub RunCollisionCheck()
    CollisionChecker.CheckAll
End Sub
