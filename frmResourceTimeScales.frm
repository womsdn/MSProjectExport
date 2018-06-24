VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmResourceTimeScales 
   Caption         =   "Export resource planning"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5310
   OleObjectBlob   =   "frmResourceTimeScales.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmResourceTimeScales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tskTasks As Collection
Private rscResources As Collection
Private Sub cbStartExport_Click()
    'Initialize object
    Dim mycResourcetimeScales As cResourceTimeScales
    Set mycResourcetimeScales = New cResourceTimeScales
    
    'Create object with timescales
    mycResourcetimeScales.Create rscResources, tskTasks, dpStart, dpFinish, cmUnit.Value
    
    'Export object
    mycResourcetimeScales.Dump2MarkDown

End Sub
Public Sub UserForm_Initialize()
    Dim datFirstDate As Date
    Dim datLastDate As Date
    Dim tsk As Task
    Dim rsc As Resource
    Dim ass As Assignment
    
    Set tskTasks = New Collection
    Set rscResources = New Collection
    
    'Determine tasks and resources to inspect based on viewtype (task or resrouce) and selected items
    If ActiveProject.Views(ActiveProject.CurrentView).Type = pjTaskItem Then
        For Each tsk In ActiveSelection.Tasks
            tskTasks.Add tsk
        Next
        
        For Each rsc In ActiveProject.Resources
            rscResources.Add rsc
        Next
        
        'Determine first and last date, based on selected tasks
        datFirstDate = tskTasks(1).Start
        datLastDate = tskTasks(1).Finish
        For Each tsk In tskTasks
            If tsk.Start < datFirstDate Then
                datFirstDate = tsk.Start
            End If
            If tsk.Finish > datLastDate Then
                datLastDate = tsk.Finish
            End If
        Next
    ElseIf ActiveProject.Views(ActiveProject.CurrentView).Type = pjResourceItem Then
        
        For Each tsk In ActiveProject.Tasks
            tskTasks.Add tsk
        Next
        
        For Each rsc In ActiveSelection.Resources
            rscResources.Add rsc
        Next
        
        'Determine first and last date, based on assignments of selected resources
        datFirstDate = rscResources(1).Assignments(1).Start
        datLastDate = rscResources(1).Assignments(1).Finish
        For Each rsc In rscResources
            For Each ass In rsc.Assignments
                If ass.Start < datFirstDate Then
                    datFirstDate = ass.Start
                End If
                If ass.Finish > datLastDate Then
                    datLastDate = ass.Finish
                End If
            Next
        Next
    End If
    
    Me.dpStart = datFirstDate
    Me.dpFinish = datLastDate

    Me.cmUnit.AddItem (pjTimescaleMonths)
    Me.cmUnit.Column(1, 0) = ("Per month")
    Me.cmUnit.AddItem (pjTimescaleWeeks)
    Me.cmUnit.Column(1, 1) = ("Per week")
    Me.cmUnit.AddItem (pjTimescaleYears)
    Me.cmUnit.Column(1, 2) = ("Per year")
    Me.cmUnit.Value = pjTimescaleMonths


End Sub
