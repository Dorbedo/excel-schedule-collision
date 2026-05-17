Attribute VB_Name = "CollisionChecker"
Option Explicit

' Checks whether two time/layout ranges overlap
Public Function HasCollision(startA As Date, endA As Date, startB As Date, endB As Date) As Boolean
    HasCollision = (startA < endB) And (endA > startB)
End Function
