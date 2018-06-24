Attribute VB_Name = "ClassTests"
Option Explicit
Public Sub test()

Dim mycResourcetimeScales As cResourceTimeScales
Dim datStart As Date
Dim datFinish As Date
Dim Rscs As New Collection
Dim Tsks As New Collection
Dim Tsunit As PjTimescaleUnit
Dim rsc As Resource
Dim tsk As Task

'myCol.Add ActiveProject.Resources(1)
Tsks.Add ActiveProject.Tasks(8)

For Each rsc In ActiveProject.Resources
    Rscs.Add rsc
Next

datStart = #7/3/2017#
datFinish = #6/29/2018#



Tsunit = pjTimescaleMonths

Debug.Print


Set mycResourcetimeScales = New cResourceTimeScales

mycResourcetimeScales.Create Rscs, Tsks, datStart, datFinish, Tsunit
'mycResourcetimeScales.Create rscs, Tsks


mycResourcetimeScales.Dump2MarkDown


End Sub

Public Sub RenameResources()

Dim rsc As Resource

For Each rsc In ActiveProject.Resources
    rsc.Name = "Employee " & rsc.ID
Next

End Sub
