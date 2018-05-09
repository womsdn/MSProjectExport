Attribute VB_Name = "ClassTests"
Option Explicit
Public Sub test()

Dim mytable As cTable

Set mytable = New cTable
mytable.Create 10, 10

mytable.CellValue(2, 1) = 100
Debug.Print mytable.CellValue(1, 1)

End Sub
