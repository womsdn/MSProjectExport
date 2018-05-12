VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cResourceTimeScales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pDicResources As Dictionary
Private Enum DetailType
    dtWork = pjAssignmentTimescaledWork
    dtActualwork = pjAssignmentTimescaledActualWork
    'dtBaselineWork = pjAssignmentTimescaledBaselineWork
End Enum

'Create multidimensional pDicResources dictionary with following structure:
'pDicResources(ResourceID)(TaskID)(DetailType)(TimeSlice)
'ResourceID - ID of the resources from rscResources
'TaskID - ID of all the tasks from tskTasks to which resource is assigned to
'Detail
Public Sub Create(rscResources As Resources, tskTasks As Tasks, datStart As Date, datFinish As Date, intTimeScaleUnit As PjTimescaleUnit)
    Dim rscResource As Resource
    Dim tskTask As Task
    Dim assAssignment As Assignment
    
    Set pDicResources = New Dictionary
    For Each rscResource In rscResources
        'Create dictionary item with resource ID as key and add the dictionary which will
        'contain all the tasks to which the resource has assignments and that are in tskTasks
        pDicResources.Add rscResource.ID, CreateObject("Scripting.Dictionary")
        For Each assAssignment In rscResource.Assignments
            For Each tskTask In tskTasks
                If assAssignment.TaskID = tskTask.ID Then
                    pDicResources(rscResource.ID).Add tskTask.ID, CreateObject("Scripting.Dictionary")
                    pDicResources(rscResource.ID)(tskTask.ID).Add dtWork, assAssignment.TimeScaleData(datStart, datFinish, dtWork, intTimeScaleUnit)
                    pDicResources(rscResource.ID)(tskTask.ID).Add dtActualwork, assAssignment.TimeScaleData(datStart, datFinish, dtActualwork, intTimeScaleUnit)
                End If
            Next
        Next
        If pDicResources(rscResource.ID).Count = 0 Then pDicResources.Remove (rscResource.ID)
    Next
End Sub

Public Sub Dump()
    Dim rscID As Variant
    Dim tskID As Variant
    
    For Each rscID In pDicResources.Keys
        Debug.Print ActiveProject.Resources(rscID).Name
        For Each tskID In pDicResources(rscID).Keys
            Debug.Print vbTab & ActiveProject.Tasks(tskID).Name
            Debug.Print vbTab & vbTab & pDicResources(rscID)(tskID)(dtWork)(1).StartDate & vbTab & pDicResources(rscID)(tskID)(dtWork)(1).Value
        Next
    Next
End Sub
