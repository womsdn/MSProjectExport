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
Private pDicTotal As Dictionary
Private Enum DetailType
    dtWorkAssTSVs
    dtWorkAssTotal
    dtActualWorkAssTSVs
    dtActualWorkAssTotal
    dtBaselineWorkAssTSVs
    dtBaselineWorkAssTotal
    
    dtWorkRscTSVs
    dtWorkRscTotal
    dtActualWorkRscTSVs
    dtActualWorkRscTotal
    dtBaselineWorkRscTSVs
    dtBaselineWorkRscTotal
    
    dtWorkTSVs
    dtworktotal
    dtActualWorkTSVs
    dtActualWorktotal
    dtBaselineWorkTSVs
    dtBaselineWorktotal
End Enum

'Create multidimensional pDicResourcesAss dictionary with following structure:
'1. pDicResourcesAss
'   2. ResourceID - key for ResourceID, contains dictionary with assigned tasks
'       3. TaskID - key for ID of the assigned task, contains dictionary with various assignment data
'           dtWorkAssTSVs - key for collection of Timescalevalues for work on assigned task in asked period and per asked timescale unit
'           dtWorkAssTotal - key for integer with total work on assigned task in asked period
'           dtActualWorkAssTSVs - key for collection of timescalevalues for actual work on assigned task in asked period and per asked timescale unit
'           dtActualWorkAssTotal - key for integer with total actual work on assigned task in asked period
'           dtBaselineWorkAssTSVs - key for collection of baseline timescalevalues for work on assigned task in asked period and per asked timescale unit
'           dtBaselineWorkAssTotal - key for integer with total baseline work on assigned task in asked period
'
'1. pDicResources
'   2. ResourceID - key for resource, contains dictionary with various resource data
'       dtWorkRscTSVs - key for dictionary that contains Startdate, Work-value pairs for the resources in asked period and per timescale unit
'       dtWorkRscTotal - key for integer that contains total work for the resource in asked period
'       dtActualWorkRscTSVs - key for dictionary that contains Startdate, Actual Work-value pairs for the resources in asked period and per timescale unit
'       dtActualWorkRscTotal - key for integer that contains total actual work for the resource in asked period
'       dtBaselineWorkRscTSVs - - key for dictionary that contains Startdate, Baseline work-value pairs for the resources in asked period and per timescale unit
'       dtBaselineWorkTscTotal  - key for integer that contains total baseline work for the resource in asked period
'
'1. pDicTotal
'   dtWorkTSVs - key for dictionary that contains Startdate - Workvalue pairs for all asked resources in asked period and per timescale unit
'   dtWorkTotal - key for integer that contains total work for all asked resources in asked period
'   dtActualWorkTSVs - key for dictionary that contains Startdate - ActualWorkvalue pairs for all asked resources in asked period and per timescale unit
'   dtActualWorkTotal - key for integer that contains total actual work for all asked resources in asked period
'   dtBaselineWorkTSVs - key for dictionary that contains Startdate - BaselineWorkvalue pairs for all asked resources in asked period and per timescale unit
'   dtBaselineWorkTotal - key for integer that contains total Baseline work for all asked resources in asked period

Public Sub Create(rscResources As Resources, tskTasks As Tasks, datStart As Date, datFinish As Date, intTimeScaleUnit As PjTimescaleUnit)
    Dim rscResource As Resource
    Dim tskTask As Task
    Dim assAssignment As Assignment
    Dim tsvTSV As TimeScaleValue
    Dim tskID As Variant
    Dim rscID As Variant
    Dim intCounter As Integer
    Dim datTSVStart As Variant
    
    Set pDicResourcesAss = New Dictionary
    Set pDicResources = New Dictionary
    Set pDicTotal = New Dictionary
    
    For Each rscResource In rscResources
       
        'Create dictionary item with resource ID as key and add the dictionary which will
        'contain all the tasks to which the resource has assignments and that are in tskTasks
        pDicResourcesAss.Add rscResource.ID, CreateObject("Scripting.Dictionary")
        
        'Initialize dictionary for rscResource totals
        pDicResources.Add rscResource.ID, CreateObject("Scripting.Dictionary")
                        
        'Initialize dictionary for rscResource timescale totals and set start values
        pDicResources(rscResource.ID).Add dtWorkRscTSVs, CreateObject("Scripting.Dictionary")
        pDicResources(rscResource.ID).Add dtActualWorkRscTSVs, CreateObject("Scripting.Dictionary")
        pDicResources(rscResource.ID).Add dtBaselineWorkRscTSVs, CreateObject("Scripting.Dictionary")
                
        For Each assAssignment In rscResource.Assignments
            For Each tskTask In tskTasks
                If assAssignment.TaskID = tskTask.ID Then
                    pDicResourcesAss(rscResource.ID).Add tskTask.ID, CreateObject("Scripting.Dictionary") 'Add dictionary to hold data
                    
                    'Add dtWork key which contains Assignment work timescaledata
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtWorkAssTSVs, assAssignment.TimeScaleData(datStart, datFinish, pjAssignmentTimescaledWork, intTimeScaleUnit)
                    'Add dtWorkAssTotal key which contains assignment work total
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtWorkAssTotal, 0
                    For Each tsvTSV In pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtWorkAssTSVs)
                        pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtWorkAssTotal) = pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtWorkAssTotal) + Val(tsvTSV.Value)
                        'Add timescale values to the dictionary with total timescale value for the resource per timeunit
                        If pDicResources(rscResource.ID)(dtWorkRscTSVs).Exists(tsvTSV.StartDate) Then
                            pDicResources(rscResource.ID)(dtWorkRscTSVs)(tsvTSV.StartDate) = pDicResources(rscResource.ID)(dtWorkRscTSVs)(tsvTSV.StartDate) + Val(tsvTSV.Value)
                        Else
                            pDicResources(rscResource.ID)(dtWorkRscTSVs).Add tsvTSV.StartDate, Val(tsvTSV.Value)
                        End If
                    Next
                    
                    'Add dtActualWork key which contains actual Assignment work timescaledata
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtActualWorkAssTSVs, assAssignment.TimeScaleData(datStart, datFinish, pjAssignmentTimescaledActualWork, intTimeScaleUnit)
                    'Add dtActualWorkAssTotal key which contains actual assignment work total
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtActualWorkAssTotal, 0
                    For Each tsvTSV In pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtActualWorkAssTSVs)
                        pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtActualWorkAssTotal) = pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtActualWorkAssTotal) + Val(tsvTSV.Value)
                        'Add timescale values to the dictionary with total timescale value for the resource per timeunit
                        If pDicResources(rscResource.ID)(dtActualWorkRscTSVs).Exists(tsvTSV.StartDate) Then
                            pDicResources(rscResource.ID)(dtActualWorkRscTSVs)(tsvTSV.StartDate) = pDicResources(rscResource.ID)(dtActualWorkRscTSVs)(tsvTSV.StartDate) + Val(tsvTSV.Value)
                        Else
                            pDicResources(rscResource.ID)(dtActualWorkRscTSVs).Add tsvTSV.StartDate, Val(tsvTSV.Value)
                        End If
                    Next
                    
                    'Add dtBaselineWork key which contains baseline Assignment work timescaledata
                    pDicResourcesAss(rscResource.ID)(tskTask.ID).Add dtBaselineWorkAssTSVs, assAssignment.TimeScaleData(datStart, datFinish, pjAssignmentTimescaledBaselineWork, intTimeScaleUnit)
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
                        For Each tsvTSV In pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtBaselineWorkAssTSVs)
                            pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtBaselineWorkAssTotal) = pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtBaselineWorkAssTotal) + Val(tsvTSV.Value)
                        Next
                    End If
                    
                    'Add timescale values to the dictionary with total timescale value for the resource per timeunit
                    For Each tsvTSV In pDicResourcesAss(rscResource.ID)(tskTask.ID)(dtBaselineWorkAssTSVs)
                        If pDicResources(rscResource.ID)(dtBaselineWorkRscTSVs).Exists(tsvTSV.StartDate) Then
                            pDicResources(rscResource.ID)(dtBaselineWorkRscTSVs)(tsvTSV.StartDate) = pDicResources(rscResource.ID)(dtBaselineWorkRscTSVs)(tsvTSV.StartDate) + Val(tsvTSV.Value)
                        Else
                            pDicResources(rscResource.ID)(dtBaselineWorkRscTSVs).Add tsvTSV.StartDate, Val(tsvTSV.Value)
                        End If
                    Next
                    
                End If
            Next
        Next
        'If there are no assignments for the resource, then remove it from the dictionary
        If pDicResourcesAss(rscResource.ID).Count = 0 Then
            pDicResourcesAss.Remove (rscResource.ID)
            pDicResources.Remove (rscResource.ID)
        'else calculate the total work assignements for the resource
        Else
            'Set start values for rscResource totals
            pDicResources(rscResource.ID).Add dtWorkRscTotal, 0
            pDicResources(rscResource.ID).Add dtActualWorkRscTotal, 0
            pDicResources(rscResource.ID).Add dtBaselineWorkRscTotal, 0
                      
            'Iterate through totals of assignments (keys in pDicResourcesAss of the rscResource)
            For Each tskID In pDicResourcesAss(rscResource.ID).Keys
                pDicResources(rscResource.ID)(dtWorkRscTotal) = pDicResources(rscResource.ID)(dtWorkRscTotal) + pDicResourcesAss(rscResource.ID)(tskID)(dtWorkAssTotal)
                pDicResources(rscResource.ID)(dtActualWorkRscTotal) = pDicResources(rscResource.ID)(dtActualWorkRscTotal) + pDicResourcesAss(rscResource.ID)(tskID)(dtActualWorkAssTotal)
                pDicResources(rscResource.ID)(dtBaselineWorkRscTotal) = pDicResources(rscResource.ID)(dtBaselineWorkRscTotal) + pDicResourcesAss(rscResource.ID)(tskID)(dtBaselineWorkAssTotal)
            Next
        End If
    Next
    
    'If there are resources with assignments then calculate totals of al resources
    'Only count resources of work type (MS Project standard)
    If pDicResources.Count > 0 Then
        'Initialize dictionary for total values and timescales of all asked resources
        pDicTotal.Add dtWorkTSVs, CreateObject("Scripting.Dictionary")
        pDicTotal.Add dtworktotal, 0
        pDicTotal.Add dtActualWorkTSVs, CreateObject("Scripting.Dictionary")
        pDicTotal.Add dtActualWorktotal, 0
        pDicTotal.Add dtBaselineWorkTSVs, CreateObject("Scripting.Dictionary")
        pDicTotal.Add dtBaselineWorktotal, 0
        
        For Each rscID In pDicResources.Keys
            If ActiveProject.Resources(rscID).Type = pjResourceTypeWork Then
                'Loop through all work TSV's
                For Each datTSVStart In pDicResources(rscID)(dtWorkRscTSVs).Keys
                    If pDicTotal(dtWorkTSVs).Exists(datTSVStart) Then
                        pDicTotal(dtWorkTSVs)(datTSVStart) = pDicTotal(dtWorkTSVs)(datTSVStart) + pDicResources(rscID)(dtWorkRscTSVs)(datTSVStart)
                    Else
                        pDicTotal(dtWorkTSVs).Add datTSVStart, pDicResources(rscID)(dtWorkRscTSVs)(datTSVStart)
                    End If
                Next
                'Calculate work total
                pDicTotal(dtworktotal) = pDicTotal(dtworktotal) + pDicResources(rscID)(dtWorkRscTotal)
            
                'Loop through all actual work TSV's
                For Each datTSVStart In pDicResources(rscID)(dtActualWorkRscTSVs).Keys
                    If pDicTotal(dtActualWorkTSVs).Exists(datTSVStart) Then
                        pDicTotal(dtActualWorkTSVs)(datTSVStart) = pDicTotal(dtActualWorkTSVs)(datTSVStart) + pDicResources(rscID)(dtActualWorkRscTSVs)(datTSVStart)
                    Else
                        pDicTotal(dtActualWorkTSVs).Add datTSVStart, pDicResources(rscID)(dtActualWorkRscTSVs)(datTSVStart)
                    End If
                Next
                'Calculate actual work total
                pDicTotal(dtActualWorktotal) = pDicTotal(dtActualWorktotal) + pDicResources(rscID)(dtActualWorkRscTotal)
            
                'Loop through all baseline work TSV's
                For Each datTSVStart In pDicResources(rscID)(dtBaselineWorkRscTSVs).Keys
                    If pDicTotal(dtBaselineWorkTSVs).Exists(datTSVStart) Then
                        pDicTotal(dtBaselineWorkTSVs)(datTSVStart) = pDicTotal(dtBaselineWorkTSVs)(datTSVStart) + pDicResources(rscID)(dtBaselineWorkRscTSVs)(datTSVStart)
                    Else
                        pDicTotal(dtBaselineWorkTSVs).Add datTSVStart, pDicResources(rscID)(dtBaselineWorkRscTSVs)(datTSVStart)
                    End If
                Next
                'Calculate baseline work total
                pDicTotal(dtBaselineWorktotal) = pDicTotal(dtBaselineWorktotal) + pDicResources(rscID)(dtBaselineWorkRscTotal)
            End If
        Next
    End If

End Sub

Public Sub Dump()
    Dim rscID As Variant
    Dim tskID As Variant
    Dim dtStart As Variant
    Dim tsv As TimeScaleValue
    Dim strString As String
    Dim intCounter As Integer
    
    For Each rscID In pDicResourcesAss.Keys
        Debug.Print ActiveProject.Resources(rscID).Name
        For Each tskID In pDicResourcesAss(rscID).Keys
            strString = ActiveProject.Tasks(tskID).Name & vbTab
            For intCounter = 1 To pDicResourcesAss(rscID)(tskID)(dtWorkAssTSVs).Count
                strString = strString & vbTab & Round(Val(pDicResourcesAss(rscID)(tskID)(dtWorkAssTSVs)(intCounter).Value) / IIf(ActiveProject.Resources(rscID).Type = pjResourceTypeWork, 60, 1), 1)
            Next
            strString = strString & vbTab & Round(pDicResourcesAss(rscID)(tskID)(dtWorkAssTotal) / IIf(ActiveProject.Resources(rscID).Type = pjResourceTypeWork, 60, 1), 1)
            Debug.Print strString
        Next
        strString = vbTab
        For Each dtStart In pDicResources(rscID)(dtWorkRscTSVs).Keys
            strString = strString & vbTab & Round(pDicResources(rscID)(dtWorkRscTSVs)(dtStart) / IIf(ActiveProject.Resources(rscID).Type = pjResourceTypeWork, 60, 1), 1)
        Next
        Debug.Print strString & vbTab & Round(pDicResources(rscID)(dtWorkRscTotal) / IIf(ActiveProject.Resources(rscID).Type = pjResourceTypeWork, 60, 1), 1)
    Next
    strString = vbTab
    For Each dtStart In pDicTotal(dtWorkTSVs).Keys
        strString = strString & vbTab & Round(pDicTotal(dtWorkTSVs)(dtStart) / 60, 1)
    Next
    Debug.Print strString & vbTab & Round(pDicTotal(dtworktotal) / 60, 1)
End Sub
