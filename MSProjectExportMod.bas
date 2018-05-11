Attribute VB_Name = "MSProjectExportMod"
Option Explicit

'MSProjectExport
'https://github.com/womsdn/MSProjectExport
Public Sub MSProjectExport(Optional booFilter As Boolean = True, _
    Optional datStartDate As Date = #1/1/1901#, Optional datFinishDate As Date = #1/1/1901#, _
    Optional intTimeScaleUnit As PjTimescaleUnit = pjTimescaleMonths)
'booFilter determines if all tasks and/or resources are exported or only the selected ones
'datStartDate and datFinishdate determine the date range the export will contain
'intTimeScaleUnit determines scale for instance per month, per week, per year

'Variables for storing MS Project objects
Dim rscResources As Resources               'Stores all resources for which the data is to be exported
Dim rscResource As Resource                 'Holds individual resource that is being processed
Dim tskTasks As Tasks                       'Stores all tasks for which the data is to be exported
Dim tskTask As Task                         'Holds individual task that is being processed

'Variables for working with TimeScaleValues
Dim tsvTimeScaleValues As TimeScaleValues
Dim tsvTimeScaleValue As TimeScaleValue

'Table for storing the result
Dim tbTable As cTable
Dim intNrOfRows As Integer
Dim intNrOfColumns As Integer

'Various
Dim intRowCounter As Integer
Dim intColumnCounter As Integer

'Constants
Const intWorkDiv As Integer = 60    'Division factor when resource is of work type

'Initiate table
Set tbTable = New cTable

'Determine which tasks and resources will be exported, based on booFilter value and selected tasks/resources
If booFilter Then
    If ActiveProject.Views(ActiveProject.CurrentView).Type = pjResourceItem Then
        Set rscResources = ActiveSelection.Resources
        Set tskTasks = ActiveProject.Tasks
    ElseIf ActiveProject.Views(ActiveProject.CurrentView).Type = pjTaskItem Then
        Set rscResources = ActiveProject.Resources
        Set tskTasks = ActiveSelection.Tasks
    Else
        Set rscResources = ActiveProject.Resources
    End If
Else
    Set rscResources = ActiveProject.Resources
    Set tskTasks = ActiveProject.Tasks
End If

'If no start- and/or finish date is supplied, use the one of the total project
If datStartDate = #1/1/1901# Then datStartDate = ActiveProject.ProjectSummaryTask.Start
If datFinishDate = #1/1/1901# Then datFinishDate = ActiveProject.ProjectSummaryTask.Finish

'Determine number of rows and columns needed
intNrOfRows = rscResources.Count + 1 'Number of rows needed is equal to resources plus header
Set tsvTimeScaleValues = ActiveProject.ProjectSummaryTask.TimeScaleData(datStartDate, datFinishDate, pjTaskTimescaledWork, intTimeScaleUnit)
intNrOfColumns = tsvTimeScaleValues.Count + 1 'Number of columns needed is equal to found number of dates in range + 1

'Create the table
tbTable.Create intNrOfRows, intNrOfColumns


'Fill the header
intRowCounter = 1
intColumnCounter = 1
tbTable.CellValue(intRowCounter, intColumnCounter) = "Name"
For Each tsvTimeScaleValue In tsvTimeScaleValues
    intColumnCounter = intColumnCounter + 1
    Select Case intTimeScaleUnit
        Case pjTimescaleHalfYears, pjTimescaleMonths, pjTimescaleQuarters, pjTimescaleYears
            tbTable.CellValue(intRowCounter, intColumnCounter) = Format(tsvTimeScaleValue.StartDate, "mmm-yy")
        Case pjTimescaleHours, pjTimescaleMinutes
            tbTable.CellValue(intRowCounter, intColumnCounter) = Format(tsvTimeScaleValue.StartDate, "d-m-yy hh:mm")
        Case pjTimescaleNone, pjTimescaleThirdsOfMonths, pjTimescaleWeeks, pjTimescaleDays
            tbTable.CellValue(intRowCounter, intColumnCounter) = Format(tsvTimeScaleValue.StartDate, "d-m-yy")
        Case Else
            tbTable.CellValue(intRowCounter, intColumnCounter) = tsvTimeScaleValue.StartDate
    End Select
Next

intColumnCounter = 1
For Each rscResource In rscResources
    intRowCounter = intRowCounter + 1
    tbTable.CellValue(intRowCounter, 1) = rscResource.Name
Next

tbTable.Dump

End Sub
