Attribute VB_Name = "MSProjectExportMod"
Option Explicit
'MSProjectExport
'https://github.com/womsdn/MSProjectExport
Public Sub MSProjectExport(Optional booFilter As Boolean = True)
'booFilter determines if all tasks and/or resources are exported or only the selected ones

'Variables for storing MS Project objects
Dim rscResources As Resources       'Stores all resources for which the data is to be exported
Dim rscResource As Resource         'Holds individual resource that is being processed
Dim tskTasks As Tasks               'Stores all tasks for which the data is to be exported
Dim tskTask As Task                 'Holds individual task that is being processed

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



End Sub

Public Sub DictionaryTest()

Dim mydict As Dictionary


Set mydict = CreateObject("Scripting.Dictionary")

mydict.Add 1, CreateObject("Scripting.Dictionary")
mydict.Add 2, CreateObject("Scripting.dictionary")

mydict(1)(1) = "R1C1"
mydict(1)(2) = "R1C2"
mydict(1)(3) = "R1C3"
mydict(1)(4) = "R1C4"

mydict(2)(1) = "R2C1"
mydict(2)(2) = "R2C2"
mydict(2)(3) = "R2C3"
mydict(2)(4) = "R2C4"

Debug.Print




End Sub
