Attribute VB_Name = "ClassTests"
Option Explicit
Public Sub test()

Dim mycResourcetimeScales As cResourceTimeScales
Dim datStart As Date
Dim datFinish As Date
Dim Rcss As Resources
Dim Tsks As Tasks
Dim Tsunit As PjTimescaleUnit

datStart = #7/3/2017#
datFinish = #6/29/2018#
Set Rcss = ActiveProject.Resources
Set Tsks = ActiveSelection.Tasks
Tsunit = pjTimescaleHalfYears

Debug.Print


Set mycResourcetimeScales = New cResourceTimeScales

mycResourcetimeScales.Create Rcss, Tsks, datStart, datFinish, Tsunit
'mycResourcetimeScales.Create Rcss, Tsks


mycResourcetimeScales.Dump2MarkDown


End Sub

Public Sub RenameResources()

Dim rsc As Resource

For Each rsc In ActiveProject.Resources
    rsc.Name = "Employee " & rsc.ID
Next

End Sub
