Attribute VB_Name = "ClassTests"
Option Explicit
Public Sub test()

Dim mycResourcetimeScales As cResourceTimeScales
Dim datStart As Date
Dim datFinish As Date
Dim Rcss As Resources
Dim Tsks As Tasks
Dim Tsunit As PjTimescaleUnit

datStart = #1/1/2018#
datFinish = #3/2/2018#
Set Rcss = ActiveProject.Resources
Set Tsks = ActiveSelection.Tasks
Tsunit = pjTimescaleMonths

Debug.Print


Set mycResourcetimeScales = New cResourceTimeScales

mycResourcetimeScales.Create Rcss, Tsks, datStart, datFinish, Tsunit


mycResourcetimeScales.Dump


End Sub

