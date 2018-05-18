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
Private pDicResourcesAss As Dictionary
Private pDicResources As Dictionary
Private Enum DetailType
    dtWork = pjAssignmentTimescaledWork
    dtActualWork = pjAssignmentTimescaledActualWork
    dtBaselineWork = pjAssignmentTimescaledBaselineWork
    dtWorkAssTotal
    dtActualWorkAssTotal
    dtBaselineWorkAssTotal
End Enum

'Create multidimensional pDicResourcesAss dictionary with following structure:
'pDicResourcesAss(ResourceID)(TaskID)(DetailType)(TimeSlice)
'ResourceID - ID of the resources from rscResources
'TaskID - ID of all the tasks from tskTasks to which resource is assigned to
'Detail
Public Sub Create(rscResources As Resources, tskTasks As Tasks, datStart As Date, datFinish As Date, intTimeScaleUnit As PjTimescaleUnit)
    Dim rscResource As Resource
    Dim tskTask As Task
    Dim assAssignment As Assignment
    Dim tsvTSV As TimeScaleValue
    Dim tskID As Variant
    Dim intCounter As Integer
    
    Set pDicResourcesAss = New Dictionary
    Set pDicResources = New Dictionary
    For Each rscResource In rscResources
       
        'Create dictionary item with resource ID as key and add the dictionary which will
        'contain all the tasks to which the resource has assignments and that are in tskTasks
        pDicResourcesAss.Add rscResource.ID, CreateObject("Scripting.Dictionary")
        
        For Each assAssignment In rscResource.Assignments
            For Each tskTask In tskTasks
                If assAssignment.TaskID = tskTask.ID Then
                    pDicResourcesAss(rscResource.ID).Add tskTask.ID, CreateObject("Scripting.Dictionary") 'Add dictionary to hold data
                    
                    'Add dtWork key which contains Assignment work timescaledata
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtWork, assAssignment.TimeScaleData(datStart, datFinish, dtWork, intTimeScaleUnit)
                    'Add dtWorkAssTotal key which contains assignment work total
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtWorkAssTotal, 0
                    For Each tsvTSV In pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtWork)
                        pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtWorkAssTotal) = pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtWorkAssTotal) + Val(tsvTSV.Value)
                    Next
                    
                    'Add dtActualWork key which contains actual Assignment work timescaledata
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtActualWork, assAssignment.TimeScaleData(datStart, datFinish, dtActualWork, intTimeScaleUnit)
                    'Add dtActualWorkAssTotal key which contains actual assignment work total
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtActualWorkAssTotal, 0
                    For Each tsvTSV In pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtActualWork)
                        pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtActualWorkAssTotal) = pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtActualWorkAssTotal) + Val(tsvTSV.Value)
                    Next
                    
                    'Add dtBaselineWork key which contains baseline Assignment work timescaledata
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtBaselineWork, assAssignment.TimeScaleData(datStart, datFinish, dtBaselineWork, intTimeScaleUnit)
                    'Add dtBaselineWorkAssTotal key which contains baseline assignment work total
                    'Check if assignment start/finish falls completely in asked period (datStart - datFinish)
                    'This determines the way the assignment total baseline work will be calculated
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtBaselineWorkAssTotal, 0
                    If assAssignment.Start >= datStart And assAssignment.Finish <= datFinish Then
                        'If so then the total can be calculated with assassignment.baselinework
                        'This is preferred, because baseline timescale data is less reliable, because
                        'baseline data is data the user in most cases enters manually and he is
                        'more likely to meticulously administer assignment baselines than timescale baselines
                        'And if the Set Baseline function is used, assignment and timescale baseline data will be synchronized anyway
                        pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtBaselineWorkAssTotal) = assAssignment.BaselineWork
                    Else
                        'If not then the total baseline has to be calculated with the timescale data
                        For Each tsvTSV In pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtBaselineWork)
                            pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtBaselineWorkAssTotal) = pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtBaselineWorkAssTotal) + Val(tsvTSV.Value)
                        Next
                    End If
                    
                End If
            Next
        Next
        'If there are no assignments for the resource, then remove it from the dictionary
        If pDicResourcesAss(rscResource.ID).Count = 0 Then
            pDicResourcesAss.Remove (rscResource.ID)
        'else calculate the total work assignements for the resource
        Else
            'Initialize dictionary for rscResource totals and set start values
            pDicResources.Add rscResource.ID, CreateObject("Scripting.Dictionary")
            pDicResources(rscResource.ID).Add dtWork, 0
            pDicResources(rscResource.ID).Add dtActualWork, 0
            pDicResources(rscResource.ID).Add dtBaselineWork, 0
                      
            'Iterate through totals of assignments (keys in pDicResourcesAss of the rscResource)
            For Each tskID In pDicResourcesAss(rscResource.ID).Keys
                pDicResources(rscResource.ID)(dtWork) = pDicResources(rscResource.ID)(dtWork) + pDicResourcesAss(rscResource.ID)(tskID)(dtWorkAssTotal)
                pDicResources(rscResource.ID)(dtActualWork) = pDicResources(rscResource.ID)(dtActualWork) + pDicResourcesAss(rscResource.ID)(tskID)(dtActualWorkAssTotal)
                pDicResources(rscResource.ID)(dtBaselineWork) = pDicResources(rscResource.ID)(dtBaselineWork) + pDicResourcesAss(rscResource.ID)(tskID)(dtBaselineWorkAssTotal)
            Next
        End If
    Next
End Sub

Public Sub Dump()
    Dim rscID As Variant
    Dim tskID As Variant
    Dim tsv As TimeScaleValue
    Dim strString As String
    Dim intCounter As Integer
    
    For Each rscID In pDicResourcesAss.Keys
        Debug.Print ActiveProject.Resources(rscID).Name
        For Each tskID In pDicResourcesAss(rscID).Keys
            strString = ActiveProject.Tasks(tskID).Name & vbTab & pDicResourcesAss(rscID)(tskID)(dtWorkAssTotal) / IIf(ActiveProject.Resources(rscID).Type = pjResourceTypeWork, 60, 1)
            For intCounter = 1 To pDicResourcesAss(rscID)(tskID)(dtWork).Count
                strString = strString & vbTab & Val(pDicResourcesAss(rscID)(tskID)(dtWork)(intCounter).Value) / IIf(ActiveProject.Resources(rscID).Type = pjResourceTypeWork, 60, 1)
            Next
            Debug.Print strString
        Next
    Next
End Sub
