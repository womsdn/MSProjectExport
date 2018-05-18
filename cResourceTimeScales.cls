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
    dtwork = pjAssignmentTimescaledWork
    dtActualWork = pjAssignmentTimescaledActualWork
    dtBaseLineWork = pjAssignmentTimescaledBaselineWork
    dtworkasstotal
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
    Dim booAssWithinAskedPeriod As Integer
    
    Set pDicResourcesAss = New Dictionary
    Set pDicResources = New Dictionary
    For Each rscResource In rscResources
        booAssWithinAskedPeriod = True 'Sets default value which will determine how the total baseline work will be calulated further down
       
        'Create dictionary item with resource ID as key and add the dictionary which will
        'contain all the tasks to which the resource has assignments and that are in tskTasks
        pDicResourcesAss.Add rscResource.ID, CreateObject("Scripting.Dictionary")
        pDicResources.Add rscResource.ID, CreateObject("Scripting.Dictionary")
        pDicResources(rscResource.ID).Add dtBaseLineWork, 0
        
        For Each assAssignment In rscResource.Assignments
            For Each tskTask In tskTasks
                If assAssignment.TaskID = tskTask.ID Then
                    pDicResourcesAss(rscResource.ID).Add tskTask.ID, CreateObject("Scripting.Dictionary") 'Add dictionary to hold data
                    
                    'Add dtWork key which contains Assignment work timescaledata
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtwork, assAssignment.TimeScaleData(datStart, datFinish, dtwork, intTimeScaleUnit)
                    'Add dtWorkAssTotal key which contains assignment work total
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtworkasstotal, 0
                    For Each tsvTSV In pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtwork)
                        pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtworkasstotal) = pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtworkasstotal) + Val(tsvTSV.Value)
                    Next
                    
                    'Add dtActualWork key which contains actual Assignment work timescaledata
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtActualWork, assAssignment.TimeScaleData(datStart, datFinish, dtActualWork, intTimeScaleUnit)
                    'Add dtActualWorkAssTotal key which contains actual assignment work total
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtActualWorkAssTotal, 0
                    For Each tsvTSV In pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtActualWork)
                        pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtActualWorkAssTotal) = pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtActualWorkAssTotal) + Val(tsvTSV.Value)
                    Next
                    
                    'Add dtBaselineWork key which contains baseline Assignment work timescaledata
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtBaseLineWork, assAssignment.TimeScaleData(datStart, datFinish, dtBaseLineWork, intTimeScaleUnit)
                    'Add dtBaselineWorkAssTotal key which contains baseline assignment work total
                    'Check if assignment start/finish falls completely in asked period (datStart - datFinish)
                    'This determines the way the assignment total baseline work will be calculated
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtBaselineWorkAssTotal, 0
                    If assAssignment.Start >= datStart And assAssignment.Finish <= datFinish Then
                        pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtBaselineWorkAssTotal) = assAssignment.BaselineWork
                    Else
                        'If so then the total can be calculated with assassignment.baselinework
                        'This is preferred, because baseline timescale data is less reliable, because
                        'baseline data is data the user in most cases enters manually and he is
                        'more likely to meticulously administer assignment baselines than timescale baselines
                        'And if the Set Baseline function is used, the data will be synchronized anyway
                        For Each tsvTSV In pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtBaseLineWork)
                            pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtBaselineWorkAssTotal) = pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtBaselineWorkAssTotal) + Val(tsvTSV.Value)
                        Next
                    End If
                    
                    'Check if assignment start/finish falls completey in asked periode (datStart - datFinish)
                    'This determines the way the total baseline work will be calculated
                    If assAssignment.Start >= datStart And assAssignment.Finish <= datFinish Then
                        'If so then the total can be calculated with assassignment.baselinework
                        'This is preferred, because baseline timescale data is less reliable, because
                        'baseline data is data the user in most cases enters manually and he is
                        'more likely to meticulously administer assignment baselines than timescale baselines
                        'And if the Set Baseline function is used, the data will be synchronized anyway
                        pDicResources(rscResource.ID)(dtBaseLineWork) = pDicResources(rscResource.ID)(dtBaseLineWork) + assAssignment.BaselineWork
                    Else
                        'If not then the baseline has to be calculated with the timscale data, that happens further down the line if this boolean is set to false
                        booAssWithinAskedPeriod = False
                    End If
                End If
            Next
        Next
        'If there are no assignments for the resource, then remove it from the dictionary
        If pDicResourcesAss(rscResource.ID).Count = 0 Then
            pDicResourcesAss.Remove (rscResource.ID)
            pDicResources.Remove (rscResource.ID)
        'else calculate the total work assignements for the resources
        Else
            'Set start value
            pDicResources(rscResource.ID).Add dtwork, 0
            pDicResources(rscResource.ID).Add dtActualWork, 0
            
            'Recalculate baseline work if assAssignments are not within asked period
            If Not booAssWithinAskedPeriod Then pDicResources(rscResource.ID)(dtBaseLineWork) = 0
            
            'Iterate through assignments
            For Each tskID In pDicResourcesAss(rscResource.ID).Keys
                For intCounter = 1 To pDicResourcesAss(rscResource.ID)(tskID)(dtwork).Count
                    pDicResources(rscResource.ID)(dtwork) = pDicResources(rscResource.ID)(dtwork) + Val(pDicResourcesAss(rscResource.ID)(tskID)(dtwork)(intCounter).Value)
                    pDicResources(rscResource.ID)(dtActualWork) = pDicResources(rscResource.ID)(dtActualWork) + Val(pDicResourcesAss(rscResource.ID)(tskID)(dtActualWork)(intCounter).Value)
                    If Not booAssWithinAskedPeriod Then
                        pDicResources(rscResource.ID)(dtBaseLineWork) = pDicResources(rscResource.ID)(dtBaseLineWork) + Val(pDicResourcesAss(rscResource.ID)(tskID)(dtBaseLineWork)(intCounter).Value)
                    End If
                Next
                
                'If all assignments fall within the asked range (datStart - datPeriod) then the baseline work is equal to the total of
                
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
            strString = ActiveProject.Tasks(tskID).Name & vbTab & pDicResourcesAss(rscID)(tskID)(dtworkasstotal) / IIf(ActiveProject.Resources(rscID).Type = pjResourceTypeWork, 60, 1)
            For intCounter = 1 To pDicResourcesAss(rscID)(tskID)(dtwork).Count
                strString = strString & vbTab & Val(pDicResourcesAss(rscID)(tskID)(dtwork)(intCounter).Value) / IIf(ActiveProject.Resources(rscID).Type = pjResourceTypeWork, 60, 1)
            Next
            Debug.Print strString
        Next
    Next
End Sub
