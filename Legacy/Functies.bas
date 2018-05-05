Attribute VB_Name = "Functies"
Option Explicit

'***************** Code Start ***************
'This code was originally written by Dev Ashish.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Dev Ashish
'
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_MAX = 10
Private Const VK_ESCAPE = &H1B

Private Declare Function apiGetAsyncKeyState Lib "User32" _
        Alias "GetAsyncKeyState" _
        (ByVal vKey As Long) _
        As Integer

Private Declare Function apiFindWindow Lib "User32" Alias _
    "FindWindowA" (ByVal strClass As String, _
    ByVal lpWindow As String) As Long

Private Declare Function apiSendMessage Lib "User32" Alias _
    "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal _
    wParam As Long, lParam As Long) As Long
    
Private Declare Function apiSetForegroundWindow Lib "User32" Alias _
    "SetForegroundWindow" (ByVal hwnd As Long) As Long
    
Private Declare Function apiShowWindow Lib "User32" Alias _
    "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    
Private Declare Function apiIsIconic Lib "User32" Alias _
    "IsIconic" (ByVal hwnd As Long) As Long


Public Function fIsAppRunning(ByVal strAppName As String, _
        Optional fActivate As Boolean) As Boolean
    Dim lngH As Long, strClassName As String
    Dim lngX As Long, lngTmp As Long
    Const WM_USER = 1024
    On Local Error GoTo fIsAppRunning_Err
    fIsAppRunning = False
    Select Case LCase$(strAppName)
        Case "excel":       strClassName = "XLMain"
        Case "word":        strClassName = "OpusApp"
        Case "access":      strClassName = "OMain"
        Case "powerpoint95": strClassName = "PP7FrameClass"
        Case "powerpoint97": strClassName = "PP97FrameClass"
        Case "notepad":     strClassName = "NOTEPAD"
        Case "paintbrush":  strClassName = "pbParent"
        Case "wordpad":     strClassName = "WordPadClass"
        Case Else:          strClassName = vbNullString
    End Select
    
    If strClassName = "" Then
        lngH = apiFindWindow(vbNullString, strAppName)
    Else
        lngH = apiFindWindow(strClassName, vbNullString)
    End If
    If lngH <> 0 Then
        apiSendMessage lngH, WM_USER + 18, 0, 0
        lngX = apiIsIconic(lngH)
        If lngX <> 0 Then
            lngTmp = apiShowWindow(lngH, SW_SHOWNORMAL)
        End If
        If fActivate Then
            lngTmp = apiSetForegroundWindow(lngH)
        End If
        fIsAppRunning = True
    End If
fIsAppRunning_Exit:
    Exit Function
fIsAppRunning_Err:
    fIsAppRunning = False
    Resume fIsAppRunning_Exit
End Function
'******************** Code End ****************
Sub Filter_Select()
'This Macro filters your gantt view to show only the
'tasks that you have selected. It is helpful when you want
'to present a certain groups of tasks or
'you want to show a set of tasks that would
'be difficult to filter for.
'This macro overwrites data in the flag5 field
'Thanks to Michael Edwards for catching the bug in this one.
'version history:
'v1.00 Feb, 2002
'v1.01 Mar 14, 2002 (corrected field used for filter)

'Copyright Jack Dahlgren Feb. 2002

Dim jTasks As Tasks
Dim jTask As Task
Dim jResources As Resources
Dim jResource   As Resource

If ActiveProject.Views(ActiveProject.CurrentView).Type = pjTaskItem Then
    'clear the flag field
    For Each jTask In ActiveProject.Tasks
        If Not jTask Is Nothing Then
            jTask.Flag5 = "Nee"
        End If
    Next jTask
        
    'set the flag field for the selected tasks
    Set jTasks = ActiveSelection.Tasks
    For Each jTask In jTasks
        If Not jTask Is Nothing Then
            jTask.Flag5 = "Ja"
        End If
    Next jTask
    
    'filter to show just the selected tasks
    FilterEdit Name:="select", TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Flag5", Test:="Is gelijk aan", Value:="Ja", ShowInMenu:=False, ShowSummaryTasks:=False
    FilterApply Name:="select"
ElseIf ActiveProject.Views(ActiveProject.CurrentView).Type = pjResourceItem Then
    'Clear the flag field
    For Each jResource In ActiveProject.Resources
        If Not jResource Is Nothing Then
            jResource.Flag5 = "Nee"
        End If
    Next jResource
    
    'set the flag field for the selected resources
    Set jResources = ActiveSelection.Resources
    For Each jResource In jResources
        If Not jResource Is Nothing Then
            jResource.Flag5 = "Ja"
        End If
    Next jResource

    'Filter to show just the selected resources
    FilterEdit Name:="select", TaskFilter:=False, Create:=True, OverwriteExisting:=True, FieldName:="Flag5", Test:="Is gelijk aan", Value:="Ja", ShowInMenu:=False, ShowSummaryTasks:=False
    FilterApply Name:="select"
End If

End Sub
'Geselecteerde taken kopi�ren naar het klembord
Sub CopyTasks(Optional intImageWidth As Integer = 20)
Dim TC As Task
Dim seltasks As Tasks
Dim StartDate As Date
Dim EndDate As Date
EndDate = #1/1/1980#
StartDate = #12/31/2020#

'Controleer of er een taak is geselecteerd, zo nee dan is het een ander resourceoverzicht
On Error GoTo MustBeResourcesThen
Debug.Print ActiveSelection.Tasks.Count
On Error GoTo 0



Set seltasks = ActiveSelection.Tasks
On Error GoTo skipTask:

For Each TC In seltasks
    If Not TC Is Nothing Then
        If TC.Start < StartDate Then
            StartDate = TC.Start
        End If
        If TC.Finish > EndDate Then
            EndDate = TC.Finish + 20
        End If
    End If
skipTask:
Next TC

' Adjust back one week
StartDate = DateAdd("ww", -1, StartDate)
EditCopyPicture Object:=False, ForPrinter:=0, SelectedRows:=1, _
    FromDate:=StartDate, ToDate:=EndDate, ScaleOption:=pjCopyPictureTimescale, maximagewidth:=intImageWidth

Exit Sub

MustBeResourcesThen:
    MsgBox "Huidig overzicht bevat geen taken", vbCritical, "Geen taken"
End Sub

Public Function GetLastActualDate(rcs As Resource) As String
'Deze functie retourneert een datum (in stringvorm) tot welke er voor een resource
'actuele waarden zijn geboekt.
'Hij werk in eenheden van weken, dus als op bv op maandag 31-1-2011 uren zijn geboekt
'retourneert deze functie dat uren zijn geboekt tot maandag 7-2-2011.
'Als er nog geen uren zijn geboekt, retourneert hij de waarde "NVT"

    Dim TSVs As TimeScaleValues
    Dim TSV As TimeScaleValue
    Dim intCounter As Integer

    Set TSVs = rcs.TimeScaleData(ActiveProject.ProjectSummaryTask.Start, _
        ActiveProject.ProjectSummaryTask.Finish, Type:=pjResourceTimescaledActualWork, _
        TimescaleUnit:=pjTimescaleWeeks)

    intCounter = TSVs.Count
    
    Do Until intCounter = 1 Or TSVs(intCounter) <> ""
        intCounter = intCounter - 1
    Loop

    If intCounter = 1 And TSVs(intCounter) = "" Then
        GetLastActualDate = "NVT"
    Else
        GetLastActualDate = TSVs(intCounter).EndDate
    End If

End Function
Public Sub FilterRelatedTasks()
    Dim jTask As Task
    Dim strTempStore As String
    

If ActiveProject.Views(ActiveProject.CurrentView).Type = pjTaskItem Then
    If ActiveSelection.Tasks.Count > 1 Then
        MsgBox "Er mag maar een taak geselecteerd zijn", vbOKOnly, "Meerdere taken geselecteerd"
    Else
        strTempStore = ActiveSelection.Tasks(1).ID
        
        'Clear de Flag5 indicatie
        For Each jTask In ActiveProject.Tasks
            If Not jTask Is Nothing Then
                jTask.Flag5 = "Nee"
            End If
        Next
                
        ActiveSelection.Tasks(1).Flag5 = "Ja"
        MarkPredecessorsRecursively ActiveSelection.Tasks(1)
        MarkSuccessorsRecursively ActiveSelection.Tasks(1)
        
        'Filter to show just the selected resources
        FilterEdit Name:="select", TaskFilter:=False, Create:=True, OverwriteExisting:=True, FieldName:="Flag5", Test:="Is gelijk aan", Value:="Ja", ShowInMenu:=False, ShowSummaryTasks:=False
        FilterApply Name:="select"
        
        Find "ID", "is gelijk aan", strTempStore
    End If
Else
    MsgBox "Geen taken geselecteerd", vbOKOnly, "Geen taken"
End If

End Sub

Public Sub MarkPredecessorsRecursively(tskTask As Task)
    Dim tskPredecessor As Task
    
    tskTask.Flag5 = "Ja"
    For Each tskPredecessor In tskTask.PredecessorTasks
        If Not tskPredecessor Is Nothing Then
            MarkPredecessorsRecursively tskPredecessor
        End If
    Next

End Sub

Public Sub MarkSuccessorsRecursively(tskTask As Task)
    Dim tskSuccessor As Task
    
    tskTask.Flag5 = "Ja"
    For Each tskSuccessor In tskTask.SuccessorTasks
        If Not tskSuccessor Is Nothing Then
            MarkSuccessorsRecursively tskSuccessor
        End If
    Next

End Sub


Public Sub FilterCurrentResourcesOverlap()
'Deze functie geeft alle taken weer worden waar de resources van de geselecteerde taak
'gedurende de looptijd daarvan ook aan werken.
    
    Dim jTask As Task
    Dim jAllTask As Task
    Dim assAllAssignment As Assignment
    Dim strTempStore As String
    Dim datStart As Date
    Dim datFinish As Date
    Dim resResources As Resources
    Dim resResource As Resource
    Dim assAssignments As Assignments
    Dim assAssignment As Assignment
    

    'Controleer of er een taak is geselecteerd, zo nee dan is het een ander resourceoverzicht
    If ActiveProject.Views(ActiveProject.CurrentView).Type = pjTaskItem Then
        strTempStore = ActiveSelection.Tasks(1).ID
        
        'Clear de Flag5 indicatie
        For Each jTask In ActiveProject.Tasks
            If Not jTask Is Nothing Then
                jTask.Flag5 = "Nee"
            End If
        Next
                
        
        For Each jTask In ActiveSelection.Tasks
            If Not jTask Is Nothing Then
                Debug.Print jTask.Name
                For Each assAssignment In jTask.Assignments
                    For Each jAllTask In ActiveProject.Tasks
                        If Not jAllTask Is Nothing Then
                            For Each assAllAssignment In jAllTask.Assignments
                                If assAllAssignment.ResourceID = assAssignment.ResourceID Then
                                    If Not (assAllAssignment.Finish < assAssignment.Start Or _
                                        assAllAssignment.Start > assAssignment.Finish) Then
                                        jAllTask.Flag5 = "Ja"
                                    End If
                                End If
                            Next
                        End If
                    Next
                Next
            End If
        Next
        
        'Filter to show just the selected resources
        FilterEdit Name:="select", TaskFilter:=False, Create:=True, OverwriteExisting:=True, FieldName:="Flag5", Test:="Is gelijk aan", Value:="Ja", ShowInMenu:=False, ShowSummaryTasks:=False
        FilterApply Name:="select"
        
        Find "ID", "is gelijk aan", strTempStore
    Else
        MsgBox "Geen taken geselecteerd", vbOKOnly, "Geen taken"
    End If
End Sub

Public Sub FilterCurrentTasksOverlap()
'Deze functie geeft alle overlappende taken weer van de geselecteerde taak
    
    Dim jTask As Task
    Dim jAllTask As Task
    Dim assAllAssignment As Assignment
    Dim strTempStore As String
    Dim datStart As Date
    Dim datFinish As Date
    Dim resResources As Resources
    Dim resResource As Resource
    Dim assAssignments As Assignments
    Dim assAssignment As Assignment
    

    If ActiveProject.Views(ActiveProject.CurrentView).Type = pjTaskItem Then
        strTempStore = ActiveSelection.Tasks(1).ID
        
        'Clear de Flag5 indicatie
        For Each jTask In ActiveProject.Tasks
            If Not jTask Is Nothing Then
                jTask.Flag5 = "Nee"
            End If
        Next
                
        For Each jTask In ActiveSelection.Tasks
            If Not jTask Is Nothing Then
                For Each jAllTask In ActiveProject.Tasks
                    If Not jAllTask Is Nothing Then
                        If Not (jAllTask.Finish < jTask.Start Or _
                            jAllTask.Start > jTask.Finish) Then
                            jAllTask.Flag5 = "Ja"
                        End If
                    End If
                Next
            End If
        Next
        
        'Filter to show just the selected resources
        FilterEdit Name:="select", TaskFilter:=False, Create:=True, OverwriteExisting:=True, FieldName:="Flag5", Test:="Is gelijk aan", Value:="Ja", ShowInMenu:=False, ShowSummaryTasks:=False
        FilterApply Name:="select"
        
        Find "ID", "is gelijk aan", strTempStore
Else
    MsgBox "Geen taken geselecteerd", vbOKOnly, "Geen taken"
End If
End Sub


Public Sub ResetFilter()
    'Controleer of er een taak is geselecteerd, zo nee dan is het een ander resourceoverzicht
    If ActiveProject.Views(ActiveProject.CurrentView).Type = pjTaskItem Then
        FilterApply "Alle taken"
    ElseIf ActiveProject.Views(ActiveProject.CurrentView).Type = pjResourceItem Then
        FilterApply "Alle resources"
    End If
End Sub

Public Sub TotalsPerResourcePerPhase(Optional booResourcesFilter As Boolean = False, _
    Optional booShowOutlineNumbers As Boolean = True, Optional datStartDate As Date = #1/1/1901#, Optional datFinishDate As Date = #1/1/1901#, _
    Optional intTimeScaleData As Integer = pjAssignmentTimescaledWork, Optional intTimeScaleUnit As Integer = pjTimescaleMonths, _
    Optional booColorMarkings As Boolean = True)

'Variables for storing Project objects
Dim ass As Assignment
Dim tsk As Task
Dim tskTasks As Tasks
Dim rcs As Resource
Dim booWorked As Boolean
Dim booAnyWork As Boolean
Dim intWorkDiv As Integer
Dim intAssCounter As Integer
Dim rscResources As Resources

'Variables for storing values
Dim sngRcsWorkBaseline As Single
Dim sngRcsWorkActual As Single
Dim sngRcsWorkRemaining As Single
Dim sngRcsWorkTotal As Single
Dim sngRcsWorkDeviation As Single

Dim lngRcsCostBaseline As Single
Dim lngRcsCostActual As Single
Dim lngRcsCostRemaining As Single
Dim lngRcsCostTotal As Single
Dim lngRcsCostDeviation As Single
Dim datStart As Date
Dim datFinish As Date
Dim datIntervalStart As Date
Dim datIntervalFinish As Date

'Variables for storing Excel objects
Dim appExcel As Excel.Application
Dim wbWorkbook As Excel.Workbook
Dim shSheet As Excel.Worksheet
Dim intRow As Integer
Dim intColumn As Integer
Dim intFase As Integer

'Variables for storing Timescalevalues
Dim tsvTimeScaleValues As TimeScaleValues
Dim tsvTimeScaleValue As TimeScaleValue
Dim tsvTimeScaleValuesActualWork As TimeScaleValues
Dim tsvTimeScaleValueActualWork As TimeScaleValue
Dim intColumnCounter As Integer
Dim intTeller As Integer

'Als het filter aanstaat en de huidige view is van het resourcestype selecteer dan alleen de geselecteerde resources
If booResourcesFilter Then
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

If fIsAppRunning("Excel") Then
    Set appExcel = GetObject(, "Excel.Application")
Else
    Set appExcel = CreateObject("Excel.Application")
End If

Set wbWorkbook = appExcel.Workbooks.Add
Set shSheet = wbWorkbook.ActiveSheet

With appExcel.ActiveWindow.ActiveCell
    intRow = intRow + 1
    .Cells(intRow, 1) = "Product"
    .Cells(intRow, 2) = "Begindatum"
    .Cells(intRow, 3) = "Einddatum"
    .Cells(intRow, 4) = "Werk baseline"
    .Cells(intRow, 5) = "Werk besteed"
    .Cells(intRow, 6) = "Werk nog nodig"
    .Cells(intRow, 7) = "Werk totaal"
    .Cells(intRow, 8) = "Werk afwijking"
    
    'Kolomkoppen voor timescalevalues toevoegen indien van toepassing
    If datStartDate <> #1/1/1901# And datFinishDate <> #1/1/1901# Then
        Set tsvTimeScaleValues = ActiveProject.Tasks(1).TimeScaleData(datStartDate, datFinishDate, pjTaskTimescaledWork, intTimeScaleUnit)
        For Each tsvTimeScaleValue In tsvTimeScaleValues
            intColumnCounter = intColumnCounter + 1
            .Cells(intRow, 9 + intColumnCounter) = tsvTimeScaleValue.StartDate
            Select Case intTimeScaleUnit
                Case pjTimescaleMonths
                    .Cells(intRow, 9 + intColumnCounter).NumberFormat = "[$-413]mmm-yy;@"
                Case pjTimescaleWeeks
                    .Cells(intRow, 9 + intColumnCounter).NumberFormat = "[$-413]dd-mm-yy;@"
                Case pjTimescaleYears
                    .Cells(intRow, 9 + intColumnCounter).NumberFormat = "[$-413]yyyy;@"
            End Select
        Next
    End If
    
    intRow = intRow + 1
    booAnyWork = False
    For Each rcs In rscResources
        If Not (rcs Is Nothing) Then
            booWorked = False
            
            'Controleer of de resources aan de taken uit het filter heeft gewerkt
            For Each ass In rcs.Assignments
                For Each tsk In tskTasks
                    If tsk.ID = ass.TaskID Then
                        booWorked = True
                        booAnyWork = True
                    End If
                Next
            Next
            
            'Zo ja, neem hem op in de spreadsheet
            If booWorked Then
                intAssCounter = 0
                .Cells(intRow, 1) = rcs.Name & " (Geboekt tot: " & GetLastActualDate(rcs) & ")": .Cells(intRow, 1).Font.Bold = True: intRow = intRow + 1
                If rcs.Type = pjResourceTypeWork Then
                    intWorkDiv = 60
                Else
                    intWorkDiv = 1
                End If
                For Each ass In rcs.Assignments
                    For Each tsk In tskTasks
                        If tsk.ID = ass.TaskID Then
                            intAssCounter = intAssCounter + 1
                            If tsk.Text10 = "" Then
                                .Cells(intRow, 1) = "- " & IIf(booShowOutlineNumbers, ActiveProject.Tasks(ass.TaskID).OutlineNumber & " ", "") & ass.TaskName
                            Else
                                .Cells(intRow, 1) = "- " & ActiveProject.Tasks(ass.TaskID).Text10
                            End If
                            .Cells(intRow, 2) = ActiveProject.Tasks(ass.TaskID).Start
                            .Cells(intRow, 2).NumberFormat = "ddd d/m/yy"
                            .Cells(intRow, 3) = ActiveProject.Tasks(ass.TaskID).Finish
                            .Cells(intRow, 3).NumberFormat = "ddd d/m/yy"
                            .Cells(intRow, 4) = ass.BaselineWork / intWorkDiv
                            .Cells(intRow, 5) = ass.ActualWork / intWorkDiv
                            .Cells(intRow, 6) = ass.RemainingWork / intWorkDiv
                            .Cells(intRow, 7) = ass.Work / intWorkDiv
                            .Cells(intRow, 8) = ass.WorkVariance / intWorkDiv
                            
                            'Kleurmarkeringen toevoegen
                            If booColorMarkings Then
                                Select Case CDate(Format(ActiveProject.Tasks(ass.TaskID).Start, "dd-mm-yyyy")) - Date
                                    Case Is < 0
                                        If ass.RemainingWork > 0 And ass.ActualWork = 0 Then
                                            .Cells(intRow, 2).Font.Color = RGB(255, 0, 0)
                                        End If
                                    Case Is < 14
                                        If ass.RemainingWork > 0 And ass.ActualWork = 0 Then .Cells(intRow, 2).Font.Color = RGB(247, 150, 70)
                                End Select
                                Select Case CDate(Format(ActiveProject.Tasks(ass.TaskID).Finish, "dd-mm-yyyy")) - Date
                                    Case Is < 0
                                        If ass.RemainingWork > 0 And ass.ActualWork = 0 Then .Cells(intRow, 3).Font.Color = RGB(255, 0, 0)
                                    Case Is < 14
                                        If ass.RemainingWork > 0 And ass.ActualWork = 0 Then .Cells(intRow, 3).Font.Color = RGB(247, 150, 70)
                                End Select
                            End If
                            
                            'Kolommen voor timescalevalues toevoegen indien van toepassing
                            If datStartDate <> #1/1/1901# And datFinishDate <> #1/1/1901# Then
                                Set tsvTimeScaleValues = ass.TimeScaleData(datStartDate, datFinishDate, intTimeScaleData, intTimeScaleUnit)
                                Set tsvTimeScaleValuesActualWork = ass.TimeScaleData(datStartDate, datFinishDate, pjAssignmentTimescaledActualWork, intTimeScaleUnit)
                                intColumnCounter = 0
                                For Each tsvTimeScaleValue In tsvTimeScaleValues
                                    intColumnCounter = intColumnCounter + 1
                                    If tsvTimeScaleValue.Value <> "" Then
                                        If intTimeScaleData = pjAssignmentTimescaledCost Then
                                            .Cells(intRow, 9 + intColumnCounter) = tsvTimeScaleValue.Value
                                            .Cells(intRow, 9 + intColumnCounter).NumberFormat = "� #,##0"
                                        Else
                                            .Cells(intRow, 9 + intColumnCounter) = tsvTimeScaleValue.Value / intWorkDiv
                                            .Cells(intRow, 9 + intColumnCounter).NumberFormat = "0"
                                            If intTimeScaleData = pjAssignmentTimescaledWork Then
                                                If tsvTimeScaleValue = tsvTimeScaleValuesActualWork(intColumnCounter) Then
                                                    With .Cells(intRow, 9 + intColumnCounter).Interior
                                                        .Pattern = xlSolid
                                                        .PatternColorIndex = xlAutomatic
                                                        .ThemeColor = xlThemeColorDark1
                                                        .TintAndShade = -0.14996795556505
                                                        .PatternTintAndShade = 0
                                                    End With
                                                ElseIf tsvTimeScaleValuesActualWork(intColumnCounter) > 0 And tsvTimeScaleValuesActualWork(intColumnCounter) <> "" Then
                                                    With .Cells(intRow, 9 + intColumnCounter).Interior
                                                        .Pattern = xlSolid
                                                        .PatternColorIndex = xlAutomatic
                                                        .ThemeColor = xlThemeColorDark1
                                                        .TintAndShade = -4.99893185216834E-02
                                                        .PatternTintAndShade = 0
                                                    End With
                                                End If
                                            End If
                                        End If
                                    Else
                                        .Cells(intRow, 9 + intColumnCounter) = 0
                                        If intTimeScaleData = pjAssignmentTimescaledCost Then
                                            .Cells(intRow, 9 + intColumnCounter).NumberFormat = "� #,##0"
                                        Else
                                            .Cells(intRow, 9 + intColumnCounter).NumberFormat = "0"
                                        End If
                                        
                                        
                                        .Cells(intRow, 9 + intColumnCounter) = 0
                                    End If
                                Next
                            End If
                            
                            intRow = intRow + 1
                        End If
                    Next
                    'End If
                Next
                .Cells(intRow, 1) = "Totaal " & rcs.Name
                .Cells(intRow, 4).FormulaR1C1 = "=SUBTOTAL(9, R[-" & intAssCounter & "]C:R[-1]C"
                .Cells(intRow, 5).FormulaR1C1 = "=SUBTOTAL(9, R[-" & intAssCounter & "]C:R[-1]C"
                .Cells(intRow, 6).FormulaR1C1 = "=SUBTOTAL(9, R[-" & intAssCounter & "]C:R[-1]C"
                .Cells(intRow, 7).FormulaR1C1 = "=SUBTOTAL(9, R[-" & intAssCounter & "]C:R[-1]C"
                .Cells(intRow, 8).FormulaR1C1 = "=SUBTOTAL(9, R[-" & intAssCounter & "]C:R[-1]C"
                
                
                'Voeg de totalen voor de timescalevalues toe indien van toepassing
                If intColumnCounter > 0 Then
                    For intTeller = 1 To intColumnCounter
                        .Cells(intRow, 9 + intTeller).FormulaR1C1 = "=SUBTOTAL(9, R[-" & intAssCounter & "]C:R[-1]C"
                        If intTimeScaleData = pjAssignmentTimescaledCost Then
                            .Cells(intRow, 9 + intTeller).NumberFormat = "� #,##0"
                        Else
                            .Cells(intRow, 9 + intTeller).NumberFormat = "0"
                        End If
                    Next
                End If
                
                
                intRow = intRow + 2
            End If
        End If
    Next
    
    'Als er ��n of meerdere personen aan de taken uit de filter hebben gewerkt, moeten ook de totaalregels opgenomen worden.
    If booAnyWork Then
        .Cells(intRow, 1) = "Totaal"
        .Cells(intRow, 4).FormulaR1C1 = "=SUBTOTAL(9, R3C:R[-1]C"
        .Cells(intRow, 5).FormulaR1C1 = "=SUBTOTAL(9, R3C:R[-1]C"
        .Cells(intRow, 6).FormulaR1C1 = "=SUBTOTAL(9, R3C:R[-1]C"
        .Cells(intRow, 7).FormulaR1C1 = "=SUBTOTAL(9, R3C:R[-1]C"
        .Cells(intRow, 8).FormulaR1C1 = "=SUBTOTAL(9, R3C:R[-1]C"
    
        'Voeg de totalen voor de timescalevalues toe indien van toepassing
        If intColumnCounter > 0 Then
            For intTeller = 1 To intColumnCounter
                .Cells(intRow, 9 + intTeller).FormulaR1C1 = "=SUBTOTAL(9, R3C:R[-1]C"
                If intTimeScaleData = pjAssignmentTimescaledCost Then
                    .Cells(intRow, 9 + intTeller).NumberFormat = "� #,##0"
                Else
                    .Cells(intRow, 9 + intTeller).NumberFormat = "0"
                End If
            Next
        End If
    End If
    shSheet.Columns("A:H").EntireColumn.AutoFit

End With

appExcel.Visible = True

End Sub
'Subroutine ResourceExportExcel
'Exporteert een resourceoverzicht naar Excel
Public Sub ResourceExportExcel(Optional intTaskId As Integer = 1, _
    Optional datStart As Date = #1/1/1901#, Optional datFinish As Date = #1/1/1901#, _
    Optional intDataTypeAsked As Integer = pjAssignmentTimescaledWork, _
    Optional intTimeUnit As Integer = pjTimescaleMonths, _
    Optional booResourcesFilter As Boolean = False, _
    Optional strNumberFormat As String = "0", _
    Optional strCurrencyFormat As String = "� #,##0_-", _
    Optional booBrief = True, _
    Optional booMetFunctie = True, _
    Optional booGemarkeerd = False)
    
    'Variabelen voor het werken met Excel
    Dim appExcel As Excel.Application
    Dim wbWorkbook As Excel.Workbook
    Dim shSheet As Excel.Worksheet
    Dim intRow As Integer
    Dim intColumn As Integer
    Dim intFixedColumns As Integer 'Aantal vaste kolomkoppen voor de offset van de timescalevalues
    
    'Variabelen voor het opslaan van de timescalevalues
    Dim tsvTimeScaleValues As TimeScaleValues
    Dim tsvTimeScaleValue As TimeScaleValue
    
    'Variabelen voor het opslaan van de resources
    Dim rscResources As Resources
    Dim rscResource As Resource

    'Variabelen voor het opslaan van de assignments
    Dim assAssignment As Assignment

    'Overige variabelen
    Dim intColumnCount As Integer 'Aantal kolommen van de timescales
    Dim intColumnCounter As Integer 'Teller voor de kolom
    Dim sngTimeScaleValue() As Single 'Matrix sngTimeScaleValue(intTellerResource, intTellerType, intTellerPeriode)voor het tijdelijk opslaan van de timescalevalue voor een tijdsunit
    Dim strWBSFilter As String 'String waarin het outlinenumber wordt opgeslagen van de gekozen taak
    Dim booAssignments As Boolean 'Boolean die aangeeft of er voor de huidige resource assignments zijn gevonden die aan de filter voldoen
    Dim booAnyAssignments As Boolean 'Boolean die aangeeft of voor de geselecteerde resources assignments zijn gevonden
    Dim intWorkDiv As Integer 'Deler om value om te zetten naar uren danwel eenheden
    Dim intWorkDivTsc As Integer 'Deler om timescalevalues om te zetten naar uren danwel eenheden
    Dim sngActualWork As Single 'Totaal gewerkte uren
    Dim sngRemainingWork As Single 'Werk nog nodig
    Dim sngTotalWork As Single 'Totaal aan huidig verwacht werk
    Dim sngTotalWorkBaseline As Single 'Werk volgens de baseline, gebaseerd op de assignments
    Dim sngTotalWorkBaselineTSVs As Single 'Werk volgens de baseline, gebaseerd op de timescalevalues
    Dim sngWorkVariance As Single 'Verwachte afwijking op baseline, gebaseerd op de assignment
    Dim sngWorkVarianceTSVs As Single 'Verwachte afwijking op baseline, gebaseerd op de timescalevalues
    Dim sngActualCost As Single 'Totaal bestede kosten
    Dim sngRemainingCost As Single 'Kosten nog nodig
    Dim sngTotalCost As Single 'Totaal aan huidig verwacht kosten
    Dim sngTotalCostBaseline As Single 'Kosten volgens de baseline
    Dim sngCostVariance As Single 'Verwachte afwijking op baseline
    Dim varHeader As Variant 'Variant voor kolomkoppen
    Dim intStartRow As Integer 'Regel waarop het totaliseren gestart moet worden
    Dim intTeller As Integer
    Dim intExportDataType As Integer 'Wordt gebruikt voor het exporteren van het juiste datatype naar Excel
    Dim intStartColumnTotal As Integer 'Wordt gebruikt als startkolom voor het berekenen van totalen
    Dim intFinishColumnTotal As Integer 'Wordt gebruikt als finishkolom voor het berekenen van totalen
    Dim varDataType As Variant 'Wordt bij het opvragen van de Timescalevalues gebruikt om alle waarden op te vragen
                                '0 = Werkelijk werk (pjAssignmentTimescaledActualWork)
                                '1 = Totaal werk (pjAssignmentTimescaledWork)
                                '2 = Baseline werk (pjAssignmentTimescaledBaselineWork)
                                '3 = Werkelijk kosten (pjAssignmentTimescaledActualCost)
                                '4 = Totaal kosten (pjAssignmentTimescaledCost)
                                '5 = Baseline kosten (pjAssignmentTimescaledBaselineCost)
                                '6 = Resterend Werk (1-0)
                                '7 = Verschil Werk (1-2)
                                '8 = Resterende kosten (4-3)
                                '9 = Verschil kosten (4-5)
    
    'Tellers voor het vullen van de sngTimeScaleValue matrix
    Dim intTellerResource As Integer
    Dim intTellerType As Integer
    Dim intTellerPeriode As Integer
    
    'Als het filter aanstaat en de huidige view is van het resourcestype selecteer dan alleen de geselecteerde resources
    If ActiveProject.Views(ActiveProject.CurrentView).Type = pjResourceItem And booResourcesFilter Then
        Set rscResources = ActiveSelection.Resources
    Else
        Set rscResources = ActiveProject.Resources
    End If

    'Bepaal een geldige begin- en einddatum wanneer deze niet of niet correct is opgegeven
    'zodat er met de timescalevalues gewerkt kan worden
    If datStart = #1/1/1901# Then datStart = ActiveProject.Tasks(intTaskId).Start
    If datFinish = #1/1/1901# Or datFinish < datStart Then datFinish = ActiveProject.Tasks(intTaskId).Finish

    'Haal het outlinenumber op van de bij aanroep aangegeven taak, zodat daarop gefilterd kan worden
    strWBSFilter = ActiveProject.Tasks(intTaskId).OutlineNumber

    'Bepaal het aantal kolommen voor de timescalevalues
    intColumnCount = ActiveProject.Tasks(intTaskId).TimeScaleData(datStart, datFinish, 3, intTimeUnit).Count
       
    'Verkrijg een Excel applicatie met worksheet
    If fIsAppRunning("Excel") Then
        Set appExcel = GetObject(, "Excel.Application")
    Else
        Set appExcel = CreateObject("Excel.Application")
    End If
    Set wbWorkbook = appExcel.Workbooks.Add
    Set shSheet = wbWorkbook.ActiveSheet
    
    With appExcel.ActiveWindow.ActiveCell
        'Stel spreadsheet kolomkoppen in
        intRow = 2
        intColumn = 1
        intStartRow = 3
        
        'Vul de kolomkoppen in, afhankelijk van de booBrief optie
        If booBrief Then
            For Each varHeader In Array("Naam/Functie", "Afdeling", "PrSrt", _
                                        "Tarief", "Uren totaal", "Kosten totaal")
                .Cells(intRow, intColumn) = varHeader
                intColumn = intColumn + 1
            Next
        Else
            For Each varHeader In Array("Naam", "Projectfunctie", "Afdeling", "PrSrt", _
                                        "Tarief", "Uren geboekt", "Uren nog nodig", "Uren totaal", "Uren baseline (TSV)", "Uren baseline (Ass)", _
                                        "Uren afwijking (TSV)", "Uren afwijking (Ass)", "Kosten besteed", "Kosten nog nodig", "Kosten totaal", _
                                        "Kosten baseline", "Kosten afwijking")
                .Cells(intRow, intColumn) = varHeader
                intColumn = intColumn + 1
            Next
        End If
        
        intFixedColumns = intColumn - 1 'Stel aantal vaste kolommen in
        'Geef op regel 1 aan om welke details het gaat
        .Cells(1, intFixedColumns + 2) = "Details " & _
            Switch(intDataTypeAsked = pjAssignmentTimescaledWork, "uren", _
            intDataTypeAsked = pjAssignmentTimescaledActualWork, "werkelijke uren", _
            intDataTypeAsked = pjAssignmentTimescaledBaselineWork, "uren baseline", _
            intDataTypeAsked = pjAssignmentTimescaledCost, "kosten", _
            intDataTypeAsked = pjAssignmentTimescaledActualCost, "werkelijke kosten", _
            intDataTypeAsked = pjAssignmentTimescaledBaselineCost, "kosten baseline")
        
        'Vul de kolomkoppen met de datums van de timescalevalues
        Set tsvTimeScaleValues = ActiveProject.Tasks(intTaskId).TimeScaleData(datStart, datFinish, pjTaskTimescaledWork, intTimeUnit)
        For Each tsvTimeScaleValue In tsvTimeScaleValues
            .Cells(intRow, intFixedColumns + 2 + intColumnCounter) = tsvTimeScaleValue.StartDate
            Select Case intTimeUnit
                Case pjTimescaleMonths
                    .Cells(intRow, intFixedColumns + 2 + intColumnCounter).NumberFormat = "[$-413]mmm-yy;@"
                Case pjTimescaleWeeks
                    .Cells(intRow, intFixedColumns + 2 + intColumnCounter).NumberFormat = "[$-413]dd-mm-yy;@"
                Case pjTimescaleYears
                    .Cells(intRow, intFixedColumns + 2 + intColumnCounter).NumberFormat = "[$-413]yyyy;@"
            End Select
            intColumnCounter = intColumnCounter + 1
        Next
    
        'Check per resource of er assignments zijn die aan het filter voldoen
        intTellerResource = 0
        ReDim sngTimeScaleValue(0, 0, 0) 'Maak de matrix voor de timescalevalues leeg
        ReDim sngTimeScaleValue(rscResources.Count - 1, 10, intColumnCount - 1) 'Matrix voor de timescalevalues
        For Each rscResource In rscResources
            Debug.Print rscResource.Name
            
            'Initialiseer de waarden voor de resource
            booAssignments = False 'Boolean die aangeeft of er taken zijn gevonden die aan het filter voldoen
            sngActualWork = 0
            sngRemainingWork = 0
            sngTotalWork = 0
            sngTotalWorkBaselineTSVs = 0
            sngTotalWorkBaseline = 0
            sngWorkVarianceTSVs = 0
            sngWorkVariance = 0
            
            sngActualCost = 0
            sngRemainingCost = 0
            sngTotalCost = 0
            sngTotalCostBaseline = 0
            sngCostVariance = 0
            
            'Stel de deler correct in aan de hand van het type resource
            If rscResource.Type = pjResourceTypeWork Then
                'Als het resource van het type werk is, moet er door 60 gedeeld worden om het aantal uren te krijgen
                intWorkDiv = 60
            Else
                'Als het om resource van het type materiaal gaat, hoeft er nooit door 60 gedeeld te worden.
                intWorkDiv = 1
            End If
            
            'Haal de gegevens uit alle assignments op en zet ze in de variabelen
            For Each assAssignment In rscResource.Assignments
                Debug.Print rscResource.Name, assAssignment.Task.Name, assAssignment.Work / 60, assAssignment.Cost, assAssignment.CostRateTable, rscResource.CostRateTables(assAssignment.CostRateTable + 1).PayRates(1).StandardRate
                'Controleer of de assignment voldoet aan het taakfilter wat is meegegeven dmv parameter ResourceExportExcelintTaskId
                If Left(ActiveProject.Tasks(assAssignment.TaskID).OutlineNumber, Len(strWBSFilter)) = strWBSFilter Then
                    booAssignments = True 'Er is tenminste ��n assigment voor de huidige resource gevonden die aan het filter voldoet
                    booAnyAssignments = True 'Er is tenminste ��n assignment gevonden voor alle resources
                    
                    'Haal de TimeScaleValues op van alle typen
                    intTellerType = 0
                    For Each varDataType In Array(pjAssignmentTimescaledActualWork, pjAssignmentTimescaledWork, pjAssignmentTimescaledBaselineWork, _
                            pjAssignmentTimescaledActualCost, pjAssignmentTimescaledCost, pjAssignmentTimescaledBaselineCost)
                        Set tsvTimeScaleValues = assAssignment.TimeScaleData(datStart, datFinish, varDataType, intTimeUnit)
                        intTellerPeriode = 0
                        For Each tsvTimeScaleValue In tsvTimeScaleValues
                            If tsvTimeScaleValue <> "" Then
                                Select Case varDataType
                                    Case pjAssignmentTimescaledActualWork
                                        sngTimeScaleValue(intTellerResource, intTellerType, intTellerPeriode) = _
                                            sngTimeScaleValue(intTellerResource, intTellerType, intTellerPeriode) + tsvTimeScaleValue.Value / intWorkDiv
                                        sngActualWork = sngActualWork + tsvTimeScaleValue.Value / intWorkDiv
                                    Case pjAssignmentTimescaledWork
                                        sngTimeScaleValue(intTellerResource, intTellerType, intTellerPeriode) = _
                                            sngTimeScaleValue(intTellerResource, intTellerType, intTellerPeriode) + tsvTimeScaleValue.Value / intWorkDiv
                                        sngTotalWork = sngTotalWork + tsvTimeScaleValue.Value / intWorkDiv
                                    Case pjAssignmentTimescaledBaselineWork
                                        sngTimeScaleValue(intTellerResource, intTellerType, intTellerPeriode) = _
                                            sngTimeScaleValue(intTellerResource, intTellerType, intTellerPeriode) + tsvTimeScaleValue.Value / intWorkDiv
                                        sngTotalWorkBaselineTSVs = sngTotalWorkBaselineTSVs + tsvTimeScaleValue.Value / intWorkDiv
                                    Case pjAssignmentTimescaledActualCost
                                        sngTimeScaleValue(intTellerResource, intTellerType, intTellerPeriode) = _
                                            sngTimeScaleValue(intTellerResource, intTellerType, intTellerPeriode) + tsvTimeScaleValue.Value
                                        sngActualCost = sngActualCost + tsvTimeScaleValue.Value
                                    Case pjAssignmentTimescaledCost
                                        sngTimeScaleValue(intTellerResource, intTellerType, intTellerPeriode) = _
                                            sngTimeScaleValue(intTellerResource, intTellerType, intTellerPeriode) + tsvTimeScaleValue.Value
                                        sngTotalCost = sngTotalCost + tsvTimeScaleValue.Value
                                    Case pjAssignmentTimescaledBaselineCost
                                        sngTimeScaleValue(intTellerResource, intTellerType, intTellerPeriode) = _
                                            sngTimeScaleValue(intTellerResource, intTellerType, intTellerPeriode) + tsvTimeScaleValue.Value
                                        sngTotalCostBaseline = sngTotalCostBaseline + tsvTimeScaleValue.Value
                                End Select
                            End If
                            intTellerPeriode = intTellerPeriode + 1
                        Next
                        intTellerType = intTellerType + 1
                    Next
                    sngTotalWorkBaseline = sngTotalWorkBaseline + assAssignment.BaselineWork / intWorkDiv
                    sngWorkVariance = sngWorkVariance + (assAssignment.Work - assAssignment.BaselineWork) / intWorkDiv
                End If
            Next
            
            'Berekende velden (Resterend, Afwijking) vullen
            For intTeller = 0 To intTellerPeriode - 1
                sngTimeScaleValue(intTellerResource, 6, intTeller) = sngTimeScaleValue(intTellerResource, 1, intTeller) - sngTimeScaleValue(intTellerResource, 0, intTeller)
                sngRemainingWork = sngRemainingWork + sngTimeScaleValue(intTellerResource, 6, intTeller)
                sngTimeScaleValue(intTellerResource, 7, intTeller) = sngTimeScaleValue(intTellerResource, 1, intTeller) - sngTimeScaleValue(intTellerResource, 2, intTeller)
                sngWorkVarianceTSVs = sngWorkVarianceTSVs + sngTimeScaleValue(intTellerResource, 7, intTeller)
                sngTimeScaleValue(intTellerResource, 8, intTeller) = sngTimeScaleValue(intTellerResource, 4, intTeller) - sngTimeScaleValue(intTellerResource, 3, intTeller)
                sngRemainingCost = sngRemainingCost + sngTimeScaleValue(intTellerResource, 8, intTeller)
                sngTimeScaleValue(intTellerResource, 9, intTeller) = sngTimeScaleValue(intTellerResource, 4, intTeller) - sngTimeScaleValue(intTellerResource, 5, intTeller)
                sngCostVariance = sngCostVariance + sngTimeScaleValue(intTellerResource, 9, intTeller)
            Next
            'Als er assignments zijn gevonden, druk deze dan af
            If booAssignments Then
                If booBrief Then
                    intRow = intRow + 1
                    If booMetFunctie Then
                        .Cells(intRow, 1) = rscResource.Text1 & vbCrLf & rscResource.Text4 'Naam resource & projectfunctie
                    Else
                        .Cells(intRow, 1) = rscResource.Text1
                    End If
                    .Cells(intRow, 2) = rscResource.Group 'Afdeling
                    .Cells(intRow, 3) = rscResource.Text2 'Prestatiesoort
                    .Cells(intRow, 4) = "'" & RatePerDate(rscResource, datStart, datFinish)
                    .Cells(intRow, 4).HorizontalAlignment = xlRight
                    .Cells(intRow, 5) = sngTotalWork
                    .Cells(intRow, 5).NumberFormat = strNumberFormat
                    .Cells(intRow, 6) = sngTotalCost
                    .Cells(intRow, 6).NumberFormat = strCurrencyFormat
                Else
                    intRow = intRow + 1
                    .Cells(intRow, 1) = rscResource.Text1 'Naam resource
                    .Cells(intRow, 2) = rscResource.Text4 'Projectfunctie
                    .Cells(intRow, 3) = rscResource.Group 'Afdeling
                    .Cells(intRow, 4) = rscResource.Text2 'Prestatiesoort
                    .Cells(intRow, 5) = "'" & RatePerDate(rscResource, datStart, datFinish)
                    .Cells(intRow, 5).HorizontalAlignment = xlRight
                    .Cells(intRow, 6) = sngActualWork
                    .Cells(intRow, 6).NumberFormat = strNumberFormat
                    .Cells(intRow, 7) = sngRemainingWork
                    .Cells(intRow, 7).NumberFormat = strNumberFormat
                    .Cells(intRow, 8) = sngTotalWork
                    .Cells(intRow, 8).NumberFormat = strNumberFormat
                    .Cells(intRow, 9) = sngTotalWorkBaselineTSVs
                    .Cells(intRow, 9).NumberFormat = strNumberFormat
                    .Cells(intRow, 10) = sngTotalWorkBaseline
                    .Cells(intRow, 10).NumberFormat = strNumberFormat
                    .Cells(intRow, 11) = sngWorkVarianceTSVs
                    .Cells(intRow, 11).NumberFormat = strNumberFormat
                    .Cells(intRow, 12) = sngWorkVariance
                    .Cells(intRow, 12).NumberFormat = strNumberFormat
                    .Cells(intRow, 13) = sngActualCost
                    .Cells(intRow, 13).NumberFormat = strCurrencyFormat
                    .Cells(intRow, 14) = sngRemainingCost
                    .Cells(intRow, 14).NumberFormat = strCurrencyFormat
                    .Cells(intRow, 15) = sngTotalCost
                    .Cells(intRow, 15).NumberFormat = strCurrencyFormat
                    .Cells(intRow, 16) = sngTotalCostBaseline
                    .Cells(intRow, 16).NumberFormat = strCurrencyFormat
                    .Cells(intRow, 17) = sngCostVariance
                    .Cells(intRow, 17).NumberFormat = strCurrencyFormat
                End If
                
                'Periodewaarden invullen
                intExportDataType = Switch(intDataTypeAsked = pjAssignmentTimescaledActualWork, 0, _
                    intDataTypeAsked = pjAssignmentTimescaledWork, 1, _
                    intDataTypeAsked = pjAssignmentTimescaledBaselineWork, 2, _
                    intDataTypeAsked = pjAssignmentTimescaledActualCost, 3, _
                    intDataTypeAsked = pjAssignmentTimescaledCost, 4, _
                    intDataTypeAsked = pjAssignmentTimescaledBaselineCost, 5)
                For intColumnCounter = 0 To intColumnCount - 1
                    .Cells(intRow, intFixedColumns + 2 + intColumnCounter) = sngTimeScaleValue(intTellerResource, intExportDataType, intColumnCounter)
                    Select Case intDataTypeAsked
                        Case pjAssignmentTimescaledCost, pjAssignmentTimescaledActualCost, pjAssignmentTimescaledBaselineCost
                            .Cells(intRow, intFixedColumns + 2 + intColumnCounter).NumberFormat = strCurrencyFormat
                        Case Else
                            .Cells(intRow, intFixedColumns + 2 + intColumnCounter).NumberFormat = strNumberFormat
                    End Select
                    
                    If booGemarkeerd And sngTimeScaleValue(intTellerResource, 1, intColumnCounter) <> 0 Then
                        'Cel markeren als er al uren geboekt zijn
                        If sngTimeScaleValue(intTellerResource, 0, intColumnCounter) = _
                            sngTimeScaleValue(intTellerResource, 1, intColumnCounter) Then
                            With .Cells(intRow, intFixedColumns + 2 + intColumnCounter).Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .ThemeColor = xlThemeColorDark1
                                .TintAndShade = -0.14996795556505
                                .PatternTintAndShade = 0
                            End With
                        ElseIf sngTimeScaleValue(intTellerResource, 0, intColumnCounter) > 0 Then
                            With .Cells(intRow, intFixedColumns + 2 + intColumnCounter).Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .ThemeColor = xlThemeColorDark1
                                .TintAndShade = -4.99893185216834E-02
                                .PatternTintAndShade = 0
                            End With
                        End If
                    End If
                Next
            End If
        intTellerResource = intTellerResource + 1
        Next
        
        If booAnyAssignments Then
            intStartColumnTotal = Switch(booBrief, 5, Not booBrief, 6)
            intFinishColumnTotal = Switch(booBrief, 6, Not booBrief, 17)
            intRow = intRow + 1
            .Cells(intRow, 1) = "Totaal"
            For intTeller = intStartColumnTotal To intFinishColumnTotal
                .Cells(intRow, intTeller).FormulaR1C1 = "=SUBTOTAL(9, R" & intStartRow & "C:R[-1]C"
                If intTeller < Switch(booBrief, 6, Not booBrief, 13) Then
                    .Cells(intRow, intTeller).NumberFormat = strNumberFormat
                Else
                    .Cells(intRow, intTeller).NumberFormat = strCurrencyFormat
                End If
                For intColumnCounter = 0 To intColumnCount - 1
                    .Cells(intRow, intFixedColumns + 2 + intColumnCounter) = "=SUBTOTAL(9, R" & intStartRow & "C:R[-1]C"
                    Select Case intDataTypeAsked
                        Case pjAssignmentTimescaledCost, pjAssignmentTimescaledActualCost, pjAssignmentTimescaledBaselineCost
                            .Cells(intRow, intFixedColumns + 2 + intColumnCounter).NumberFormat = strCurrencyFormat
                        Case Else
                            .Cells(intRow, intFixedColumns + 2 + intColumnCounter).NumberFormat = strNumberFormat
                    End Select
                Next
            Next
        End If
        
    End With
    shSheet.Columns.ColumnWidth = 80
    shSheet.Columns.EntireColumn.AutoFit
    shSheet.Columns(intFixedColumns + 1).ColumnWidth = 5
    shSheet.Cells.EntireColumn.AutoFit
    shSheet.Cells.VerticalAlignment = xlTop
    appExcel.Visible = True

End Sub
'Functie RatePerDate
'Geeft in ��n string de tarieven voor een resource in een bepaalde periode, gescheiden
'door een "/". Als geen periode is opgegeven wordt het huidige tarief gegeven
Public Function RatePerDate(rscResource As Resource, Optional datStartDate As Date = #1/1/1901#, _
    Optional datEndDate As Date = #1/1/1901#, Optional cstCostRateTableIndex As Integer = 1) As String
    Dim prtPayRate As PayRate 'Variabele voor het doorlopen van de payrates
    Dim strRatePerdate As String 'Tijdelijke variabele voor het kunnen bewerken van een gevonden payrate
    Dim arrRatePerdate() As String 'Tijdelijke array voor het kunnen bewerken van een gevonden payrate
    
    'Zet startdatum op nu indien er geen is opgegeven
    If datStartDate = #1/1/1901# Then datStartDate = Now()
    
    'Bepaal het tarief per startdatum
    For Each prtPayRate In rscResource.CostRateTables(cstCostRateTableIndex).PayRates
        If datStartDate >= prtPayRate.EffectiveDate Then
            If rscResource.Type = pjResourceTypeWork Then
                arrRatePerdate = Split(prtPayRate.StandardRate, "/")
                strRatePerdate = arrRatePerdate(0)
            Else
                strRatePerdate = CStr(prtPayRate.StandardRate)
            End If
            RatePerDate = strRatePerdate
        End If
    Next
    
    'Indien er een geldige einddatum is opgegeven, moeten ook de overige tarieven opgehaald worden
    'en aan het resultaat worden toegevoegd
    If datEndDate <> #1/1/1901# And datEndDate > datStartDate Then
        For Each prtPayRate In rscResource.CostRateTables(cstCostRateTableIndex).PayRates
            If prtPayRate.EffectiveDate > datStartDate And prtPayRate.EffectiveDate < datEndDate Then
                If rscResource.Type = pjResourceTypeWork Then
                    arrRatePerdate = Split(prtPayRate.StandardRate, "/")
                    strRatePerdate = arrRatePerdate(0)
                Else
                    strRatePerdate = CStr(prtPayRate.StandardRate)
                End If
                RatePerDate = RatePerDate & "/" & strRatePerdate
            End If
        Next
    End If
        
End Function
Public Sub FilterCurrentTasks(Optional datDate As Date = #1/1/1901#)
    'Deze functie filtert de taken waar op dit moment volgens de planning aan gewerkt zou moeten worden.

    Dim jTask As Task
    If datDate = #1/1/1901# Then
        datDate = Now
    End If

If ActiveProject.Views(ActiveProject.CurrentView).Type = pjTaskItem Then
    'Clear de Flag5 indicatie
    For Each jTask In ActiveProject.Tasks
        If Not jTask Is Nothing Then
            If jTask.Start < datDate And jTask.Finish > datDate Then
                jTask.Flag5 = "Ja"
            Else
                jTask.Flag5 = "Nee"
            End If
        End If
    Next
                
    For Each jTask In ActiveSelection.Tasks
        If Not jTask Is Nothing Then
            If jTask.Start < datDate And jTask.Finish > datDate Then
                jTask.Flag5 = "Ja"
            End If
        End If
    Next
                
    'Filter to show just the selected resources
    FilterEdit Name:="select", TaskFilter:=False, Create:=True, OverwriteExisting:=True, FieldName:="Flag5", Test:="Is gelijk aan", Value:="Ja", ShowInMenu:=False, ShowSummaryTasks:=True
    FilterApply Name:="select"
Else
    MsgBox "Deze functie werkt alleen in een takenoverzicht", vbOKOnly, "Geen taken"
End If

End Sub

Public Sub OWSKostenNaarExcel(Optional datStart As Date = #1/1/1901#, Optional datFinish As Date = #1/1/1901#)
'Met deze functie kunnen per taak, fase of tijdseenheid de uren per ontwikkelstraat getotaliseerd worden.
'Daarvoor dient in het veld "Text6" van de resources de ontwikkelstraat benoemd te zijn.
'De functie exporteert vervolgens het aantal uren per ontwikkelstraat per taak, fase of tijdseenheid naar Excel

Dim varOWS As Variant 'Variabele voor het doorlopen van alle ontwikkelstraten
Dim rscResource As Resource 'Variabele voor het doorlopen van alle resources
Dim dicOWSs As Dictionary 'Dictionary voor het opslaan van alle ontwikkelstraten
Dim tskTask As Task 'Variabele voor het doorlopen van alle taken
Dim tskTasks As Tasks 'Variabele waarin de te doorlopen taken worden opgeslagen
Dim intOWSnr As Integer 'Variabele voor het positioneren van de juiste waarde op de juiste regel in Excel
Dim intActnr As Integer 'Variabele voor het positioneren van de juiste waarde in de juiste kolom in Excel
Dim intNiveauTest As Integer 'Variabele voor het testen of de geselecteerde taken allemaal hetzelfde niveau hebben

'Variabelen voor het werken met Excel
Dim appExcel As Excel.Application
Dim wbWorkbook As Excel.Workbook
Dim shSheet As Excel.Worksheet

Set dicOWSs = New Dictionary

'Alle ingevulde ontwikkelstraten inventariseren
For Each rscResource In ActiveProject.Resources
    If rscResource.Text6 <> "" Then
        dicOWSs(rscResource.Text6) = 1
    End If
Next

If dicOWSs.Count > 0 Then

    'Bepalen welke taken nagelopen zullen worden
    'Als er maar ��n taak is geselecteerd, krijg je de waarden van die taak, behalve
    'wanneer die kinderen heeft, dan worden die genomen. Wanneer er meerdere taken
    'geselecteerd zijn, krijg je van die taken de ontwikkeluren. Als de geselecteerde
    'taken niet allemaal hetzelfde niveau hebben, krijg je een foutmelding en wordt
    'de functie afgebroken.
    If ActiveProject.Views(ActiveProject.CurrentView).Type = pjTaskItem Then
        If ActiveSelection.Tasks.Count = 1 Then
            If ActiveSelection.Tasks(1).OutlineChildren.Count > 0 Then
                Set tskTasks = ActiveSelection.Tasks(1).OutlineChildren
            Else
                Set tskTasks = ActiveSelection.Tasks
            End If
        Else
            intNiveauTest = ActiveSelection.Tasks(1).OutlineLevel
            For Each tskTask In ActiveSelection.Tasks
                If intNiveauTest <> tskTask.OutlineLevel Then
                    MsgBox ("Selectie bevat verschillende levels")
                    Exit Sub
                Else
                    intNiveauTest = tskTask.OutlineLevel
                End If
            Next
            Set tskTasks = ActiveSelection.Tasks
        End If
    Else
        Set tskTasks = ActiveProject.Tasks.Item(1)
    End If
    
    'Verkrijg een Excel applicatie met worksheet
    If fIsAppRunning("Excel") Then
        Set appExcel = GetObject(, "Excel.Application")
    Else
        Set appExcel = CreateObject("Excel.Application")
    End If
    Set wbWorkbook = appExcel.Workbooks.Add
    Set shSheet = wbWorkbook.ActiveSheet

    If datStart = #1/1/1901# Then datStart = ActiveProject.Start
    If datFinish = #1/1/1901# Then datFinish = ActiveProject.Finish
    
    'Per ontwikkelstraat alle taken doorlopen die aan het niveau voldoen, inclusief onderliggende
    With appExcel.ActiveWindow.ActiveCell
        For Each varOWS In dicOWSs.Keys()
            .Cells(intOWSnr + 2, 1) = varOWS
            intActnr = 0
            For Each tskTask In tskTasks
                If Not tskTask Is Nothing Then
                    .Cells(1, intActnr + 2) = tskTask.Name
                    .Cells(intOWSnr + 2, intActnr + 2) = GetTaskWorkByOWS(tskTask, varOWS, datStart, datFinish)
                    .Cells(intOWSnr + 2, intActnr + 2).NumberFormat = "0"
                    intActnr = intActnr + 1
                End If
            Next
            intOWSnr = intOWSnr + 1
        Next
    End With
    shSheet.Columns.ColumnWidth = 80
    shSheet.Columns.EntireColumn.AutoFit
    shSheet.Cells.EntireColumn.AutoFit
    shSheet.Cells.VerticalAlignment = xlTop
    appExcel.Visible = True

End If


End Sub

Public Function GetTaskWorkByOWS(tskTask As Task, strOWS, _
    Optional datStart As Date = #1/1/1901#, Optional datFinish As Date = #1/1/1901#, _
    Optional booRecursed As Boolean = False) As Single

Dim tsvTimeScaleValues As TimeScaleValues
Dim tsvTimeScaleValue As TimeScaleValue
Dim sngTemp As Single
Dim tskChildTask As Task
Dim assAssignment As Assignment

If Not booRecursed Then
    If datStart = #1/1/1901# Then datStart = tskTask.Start
    If datFinish = #1/1/1901# Then datFinish = tskTask.Finish
End If

For Each assAssignment In tskTask.Assignments
    If assAssignment.Resource.Type = pjResourceTypeWork And assAssignment.Resource.Text6 = strOWS Then
        Set tsvTimeScaleValues = assAssignment.TimeScaleData(datStart, datFinish, pjAssignmentTimescaledWork, _
            pjTimescaleDays)
        For Each tsvTimeScaleValue In tsvTimeScaleValues
            If tsvTimeScaleValue <> "" Then
                sngTemp = sngTemp + tsvTimeScaleValue.Value / 60
            End If
        Next
    End If
Next

For Each tskChildTask In tskTask.OutlineChildren
    sngTemp = sngTemp + GetTaskWorkByOWS(tskChildTask, strOWS, datStart, datFinish, True)
Next

GetTaskWorkByOWS = sngTemp

End Function


'Subroutine TaskExportExcel
'Exporteert een taakoverzicht naar Excel
Public Sub TaskExportExcel(Optional datStart As Date = #1/1/1901#, Optional datFinish As Date = #1/1/1901#, _
    Optional intDataTypeAsked As Integer = pjTaskTimescaledWork, _
    Optional intTimeUnit As Integer = pjTimescaleMonths, _
    Optional booTaskFilter As Boolean = False, _
    Optional strNumberFormat As String = "0", _
    Optional strCurrencyFormat As String = "� #,##0_-", _
    Optional booBrief = True, _
    Optional booGemarkeerd = False)

    'Variabelen voor het werken met Excel
    Dim appExcel As Excel.Application
    Dim wbWorkbook As Excel.Workbook
    Dim shSheet As Excel.Worksheet
    Dim intRow As Integer
    Dim intColumn As Integer
    Dim intFixedColumns As Integer 'Aantal vaste kolomkoppen voor de offset van de timescalevalues

    'Variabelen voor het opslaan van de timescalevalues
    Dim tsvTimeScaleValues As TimeScaleValues
    Dim tsvTimeScaleValue As TimeScaleValue

    'Variabelen voor het opslaan van de tasks
    Dim tskTasks As Tasks
    Dim tskTask As Task
    
        'Variabelen voor het opslaan van de assignments
    Dim assAssignment As Assignment

    'Overige variabelen
    Dim intColumnCount As Integer 'Aantal kolommen van de timescales
    Dim intColumnCounter As Integer 'Teller voor de kolom
    Dim sngTimeScaleValue() As Single 'Matrix sngTimeScaleValue(intTellertasks, intTellerType, intTellerPeriode)voor het tijdelijk opslaan van de timescalevalue voor een tijdsunit
    Dim booAssignments As Boolean 'Boolean die aangeeft of er voor de huidige tasks assignments zijn gevonden die aan de filter voldoen
    Dim booAnyAssignments As Boolean 'Boolean die aangeeft of voor de geselecteerde tasks assignments zijn gevonden
    Dim intWorkDiv As Integer 'Deler om value om te zetten naar uren danwel eenheden
    Dim intWorkDivTsc As Integer 'Deler om timescalevalues om te zetten naar uren danwel eenheden
    Dim sngActualWork As Single 'Totaal gewerkte uren
    Dim sngRemainingWork As Single 'Werk nog nodig
    Dim sngTotalWork As Single 'Totaal aan huidig verwacht werk
    Dim sngTotalWorkBaseline As Single 'Werk volgens de baseline
    Dim sngWorkVariance As Single 'Verwachte afwijking op baseline
    Dim sngActualCost As Single 'Totaal bestede kosten
    Dim sngRemainingCost As Single 'Kosten nog nodig
    Dim sngTotalCost As Single 'Totaal aan huidig verwacht kosten
    Dim sngTotalCostBaseline As Single 'Kosten volgens de baseline
    Dim sngCostVariance As Single 'Verwachte afwijking op baseline
    Dim varHeader As Variant 'Variant voor kolomkoppen
    Dim intStartRow As Integer 'Regel waarop het totaliseren gestart moet worden
    Dim intTeller As Integer
    Dim intExportDataType As Integer 'Wordt gebruikt voor het exporteren van het juiste datatype naar Excel
    Dim intStartColumnTotal As Integer 'Wordt gebruikt als startkolom voor het berekenen van totalen
    Dim intFinishColumnTotal As Integer 'Wordt gebruikt als finishkolom voor het berekenen van totalen
    Dim varDataType As Variant 'Wordt bij het opvragen van de Timescalevalues gebruikt om alle waarden op te vragen
                                '0 = Werkelijk werk (pjTaskTimescaledActualWork)
                                '1 = Totaal werk (pjTaskTimescaledWork)
                                '2 = Baseline werk (pjTaskTimescaledBaselineWork)
                                '3 = Werkelijk kosten (pjTaskTimescaledActualCost)
                                '4 = Totaal kosten (pjTaskTimescaledCost)
                                '5 = Baseline kosten (pjTaskTimescaledBaselineCost)
                                '6 = Resterend Werk (1-0)
                                '7 = Verschil Werk (1-2)
                                '8 = Resterende kosten (4-3)
                                '9 = Verschil kosten (4-5)

    'Tellers voor het vullen van de sngTimeScaleValue matrix
    Dim intTellerTask As Integer
    Dim intTellerType As Integer
    Dim intTellerPeriode As Integer

    'Als het filter aanstaat en de huidige view is van het taskstype selecteer dan alleen de geselecteerde tasks
    If ActiveProject.Views(ActiveProject.CurrentView).Type = pjTaskItem And booTaskFilter Then
        Set tskTasks = ActiveSelection.Tasks
    Else
        Set tskTasks = ActiveProject.Tasks
    End If

    'Bepaal een geldige begin- en einddatum wanneer deze niet of niet correct is opgegeven
    'zodat er met de timescalevalues gewerkt kan worden
    If datStart = #1/1/1901# Then datStart = ActiveProject.ProjectStart
    If datFinish = #1/1/1901# Or datFinish < datStart Then datFinish = ActiveProject.ProjectFinish

    'Bepaal het aantal kolommen voor de timescalevalues
    intColumnCount = tskTasks(1).TimeScaleData(datStart, datFinish, pjTaskTimescaledWork, intTimeUnit).Count
    
        'Verkrijg een Excel applicatie met worksheet
    If fIsAppRunning("Excel") Then
        Set appExcel = GetObject(, "Excel.Application")
    Else
        Set appExcel = CreateObject("Excel.Application")
    End If
    Set wbWorkbook = appExcel.Workbooks.Add
    Set shSheet = wbWorkbook.ActiveSheet

    With appExcel.ActiveWindow.ActiveCell
        'Stel spreadsheet kolomkoppen in
        intRow = 2
        intColumn = 1
        intStartRow = 3
        
        'Vul de kolomkoppen in, afhankelijk van de booBrief optie
        If booBrief Then
            For Each varHeader In Array("Taak")
                .Cells(intRow, intColumn) = varHeader
                intColumn = intColumn + 1
            Next
        Else
            For Each varHeader In Array("Taak", "Begindatum", "Einddatum")
                .Cells(intRow, intColumn) = varHeader
                intColumn = intColumn + 1
            Next
        End If
        
        intFixedColumns = intColumn - 1 'Stel aantal vaste kolommen in

        'Geef op regel 1 aan om welke details het gaat
        .Cells(1, intFixedColumns + 2) = "Details " & _
            Switch(intDataTypeAsked = pjTaskTimescaledWork, "uren", _
            intDataTypeAsked = pjTaskTimescaledActualWork, "werkelijke uren", _
            intDataTypeAsked = pjTaskTimescaledBaselineWork, "uren baseline", _
            intDataTypeAsked = pjTaskTimescaledCost, "kosten", _
            intDataTypeAsked = pjTaskTimescaledActualCost, "werkelijke kosten", _
            intDataTypeAsked = pjTaskTimescaledBaselineCost, "kosten baseline")

        'Vul de kolomkoppen met de datums van de timescalevalues
        Set tsvTimeScaleValues = tskTasks(1).TimeScaleData(datStart, datFinish, pjTaskTimescaledWork, intTimeUnit)
        For Each tsvTimeScaleValue In tsvTimeScaleValues
            .Cells(intRow, intFixedColumns + 2 + intColumnCounter) = tsvTimeScaleValue.StartDate
            Select Case intTimeUnit
                Case pjTimescaleMonths
                    .Cells(intRow, intFixedColumns + 2 + intColumnCounter).NumberFormat = "[$-413]mmm-yy;@"
                Case pjTimescaleWeeks
                    .Cells(intRow, intFixedColumns + 2 + intColumnCounter).NumberFormat = "[$-413]dd-mm-yy;@"
                Case pjTimescaleYears
                    .Cells(intRow, intFixedColumns + 2 + intColumnCounter).NumberFormat = "[$-413]yyyy;@"
            End Select
            intColumnCounter = intColumnCounter + 1
        Next

        'Check per task of er assignments zijn die aan het filter voldoen
        intTellerTask = 0
        ReDim sngTimeScaleValue(0, 0, 0) 'Maak de matrix voor de timescalevalues leeg
        ReDim sngTimeScaleValue(tskTasks.Count - 1, 10, intColumnCount - 1) 'Matrix voor de timescalevalues
        
        For Each tskTask In tskTasks
            'Initialiseer de waarden voor de task
            sngActualWork = 0
            sngRemainingWork = 0
            sngTotalWork = 0
            sngTotalWorkBaseline = 0
            sngWorkVariance = 0
            
            sngActualCost = 0
            sngRemainingCost = 0
            sngTotalCost = 0
            sngTotalCostBaseline = 0
            sngCostVariance = 0

            'Haal de TimeScaleValues op van alle typen
            intTellerType = 0
            intWorkDiv = 60
            For Each varDataType In Array(pjTaskTimescaledActualWork, pjTaskTimescaledWork, pjTaskTimescaledBaselineWork, _
                    pjTaskTimescaledActualCost, pjTaskTimescaledCost, pjTaskTimescaledBaselineCost)
                Set tsvTimeScaleValues = tskTask.TimeScaleData(datStart, datFinish, varDataType, intTimeUnit)
                intTellerPeriode = 0
                For Each tsvTimeScaleValue In tsvTimeScaleValues
                    If tsvTimeScaleValue <> "" Then
                        Select Case varDataType
                            Case pjTaskTimescaledActualWork
                                sngTimeScaleValue(intTellerTask, intTellerType, intTellerPeriode) = tsvTimeScaleValue.Value / intWorkDiv
                                sngActualWork = sngActualWork + tsvTimeScaleValue.Value / intWorkDiv
                            Case pjTaskTimescaledWork
                                sngTimeScaleValue(intTellerTask, intTellerType, intTellerPeriode) = tsvTimeScaleValue.Value / intWorkDiv
                                sngTotalWork = sngTotalWork + tsvTimeScaleValue.Value / intWorkDiv
                            Case pjTaskTimescaledBaselineWork
                                sngTimeScaleValue(intTellerTask, intTellerType, intTellerPeriode) = tsvTimeScaleValue.Value / intWorkDiv
                                sngTotalWorkBaseline = sngTotalWorkBaseline + tsvTimeScaleValue.Value / intWorkDiv
                            Case pjTaskTimescaledActualCost
                                sngTimeScaleValue(intTellerTask, intTellerType, intTellerPeriode) = tsvTimeScaleValue.Value
                                sngActualCost = sngActualCost + tsvTimeScaleValue.Value
                            Case pjTaskTimescaledCost
                                sngTimeScaleValue(intTellerTask, intTellerType, intTellerPeriode) = tsvTimeScaleValue.Value
                                sngTotalCost = sngTotalCost + tsvTimeScaleValue.Value
                            Case pjTaskTimescaledBaselineCost
                                sngTimeScaleValue(intTellerTask, intTellerType, intTellerPeriode) = tsvTimeScaleValue.Value
                                sngTotalCostBaseline = sngTotalCostBaseline + tsvTimeScaleValue.Value
                        End Select
                    End If
                    intTellerPeriode = intTellerPeriode + 1
                Next
                intTellerType = intTellerType + 1
            Next

            'Berekende velden (Resterend, Afwijking) vullen
            For intTeller = 0 To intTellerPeriode - 1
                sngTimeScaleValue(intTellerTask, 6, intTeller) = sngTimeScaleValue(intTellerTask, 1, intTeller) - sngTimeScaleValue(intTellerTask, 0, intTeller)
                sngRemainingWork = sngRemainingWork + sngTimeScaleValue(intTellerTask, 6, intTeller)
                sngTimeScaleValue(intTellerTask, 7, intTeller) = sngTimeScaleValue(intTellerTask, 1, intTeller) - sngTimeScaleValue(intTellerTask, 2, intTeller)
                sngWorkVariance = sngWorkVariance + sngTimeScaleValue(intTellerTask, 7, intTeller)
                sngTimeScaleValue(intTellerTask, 8, intTeller) = sngTimeScaleValue(intTellerTask, 4, intTeller) - sngTimeScaleValue(intTellerTask, 3, intTeller)
                sngRemainingCost = sngRemainingCost + sngTimeScaleValue(intTellerTask, 8, intTeller)
                sngTimeScaleValue(intTellerTask, 9, intTeller) = sngTimeScaleValue(intTellerTask, 4, intTeller) - sngTimeScaleValue(intTellerTask, 5, intTeller)
                sngCostVariance = sngCostVariance + sngTimeScaleValue(intTellerTask, 9, intTeller)
            Next
                        
            'Exporteer naar Excel
            If booBrief Then
                intRow = intRow + 1
                .Cells(intRow, 1) = String((tskTask.OutlineLevel - 1) * 2, " ") & tskTask.Name
            Else
                intRow = intRow + 1
                .Cells(intRow, 1) = String((tskTask.OutlineLevel - 1) * 2, " ") & tskTask.Name
                .Cells(intRow, 2) = tskTask.Start
                .Cells(intRow, 3) = tskTask.Finish
            End If
            
            'Periodewaarden invullen
            intExportDataType = Switch(intDataTypeAsked = pjTaskTimescaledActualWork, 0, _
                intDataTypeAsked = pjTaskTimescaledWork, 1, _
                intDataTypeAsked = pjTaskTimescaledBaselineWork, 2, _
                intDataTypeAsked = pjTaskTimescaledActualCost, 3, _
                intDataTypeAsked = pjTaskTimescaledCost, 4, _
                intDataTypeAsked = pjTaskTimescaledBaselineCost, 5)
            For intColumnCounter = 0 To intColumnCount - 1
                .Cells(intRow, intFixedColumns + 2 + intColumnCounter) = sngTimeScaleValue(intTellerTask, intExportDataType, intColumnCounter)
                Select Case intDataTypeAsked
                    Case pjTaskTimescaledCost, pjTaskTimescaledActualCost, pjTaskTimescaledBaselineCost
                        .Cells(intRow, intFixedColumns + 2 + intColumnCounter).NumberFormat = strCurrencyFormat
                    Case Else
                        .Cells(intRow, intFixedColumns + 2 + intColumnCounter).NumberFormat = strNumberFormat
                End Select
                
                If booGemarkeerd And sngTimeScaleValue(intTellerTask, 1, intColumnCounter) <> 0 Then
                    'Cel markeren als er al uren geboekt zijn
                    If sngTimeScaleValue(intTellerTask, 0, intColumnCounter) = _
                        sngTimeScaleValue(intTellerTask, 1, intColumnCounter) Then
                        With .Cells(intRow, intFixedColumns + 2 + intColumnCounter).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorDark1
                            .TintAndShade = -0.14996795556505
                            .PatternTintAndShade = 0
                        End With
                    ElseIf sngTimeScaleValue(intTellerTask, 0, intColumnCounter) > 0 Then
                        With .Cells(intRow, intFixedColumns + 2 + intColumnCounter).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorDark1
                            .TintAndShade = -4.99893185216834E-02
                            .PatternTintAndShade = 0
                        End With
                    End If
                End If
            Next
            If tskTask.Summary Then
                .Cells(intRow).EntireRow.Font.Bold = True
            End If
            
            intTellerTask = intTellerTask + 1
        Next

    End With
    shSheet.Columns.ColumnWidth = 80
    shSheet.Columns.EntireColumn.AutoFit
    shSheet.Columns(intFixedColumns + 1).ColumnWidth = 5
    shSheet.Cells.EntireColumn.AutoFit
    shSheet.Cells.VerticalAlignment = xlTop
    appExcel.Visible = True

End Sub

'Deze SubRoutine maakt een taak Resources onder de geselecteerde taak aan en verplaatst en totaliseert de assignments van alle
'onderliggende taken.
Public Function RollUpResources(Optional tsvScale As Integer = pjTimescaleMonths, Optional booDeleteAssignments As Boolean = False) As String

    Dim tskTasks As Tasks 'Alle op te rollen taken
    Dim tskTask As Task 'Taak per iteratie
    Dim dctFndWrk As New Scripting.Dictionary 'Matrix voor het opslaan van het gevonden werk per resource
    Dim dctActWrk As New Scripting.Dictionary 'Matrix voor het opslaan van het gevonden werkelijke werk per resource
    Dim dctTxtWrk As New Scripting.Dictionary 'Matrix voor het opslaan van de brongegevens in de assignment notes
    Dim varResourceID As Variant
    Dim tskTargetTask As Task
    Dim intTsvTeller As Integer
    Dim tsvTargetTSVsWrk As TimeScaleValues
    Dim tsvTargetTSVsActWrk As TimeScaleValues
    Dim assTargetAssignment As Assignment
    Dim txtFeedback As String
    
    'Verwijderen van assignments laten bevestigen
    If booDeleteAssignments Then
        If MsgBox("Hiermee worden alle assignments van onderliggende taken verwijderd!" & vbCrLf & _
            "Dit kan niet hersteld worden!", vbOKCancel, "Assignments verwijderen geselecteerd") = vbCancel Then
            RollUpResources = "Operatie afgebroken door gebruiker"
            GoTo Finish
        End If
    End If
    
    'Functie wordt alleen uitgevoerd in taakweergave, wanneer er maar ��n taak geselecteerd is
    If ActiveProject.Views(ActiveProject.CurrentView).Type = pjTaskItem Then
        On Error GoTo ErrorHandler
        If ActiveSelection.Tasks.Count = 1 Then
            On Error GoTo 0
            Set tskTask = ActiveSelection.Tasks(1)
            Set tskTargetTask = tskTask.OutlineChildren.Add("Rolled up resources " & tskTask.Name, tskTask.ID + 1)
            
            'Normaal gesproken krijgt een child task dezelfde startdatum, behalve wanneer de startdatum
            'van de oorspronkelijke taak voor de projectstartdatum ligt. Dus hierop moet worden gecheckt.
            If tskTargetTask.Start <> tskTask.Start Then
                tskTargetTask.Start = tskTask.Start
            End If
            
            tskTargetTask.Duration = tskTask.Duration
            
            'Haal alle assignments uit de taak en onderliggende taken op en zet deze in dictionaries
            'Iedere betrokken resource heeft 2 dictionaries, 1 voor gepland (dctfndwrk) en 1 voor
            'werkelijk gewerkt werk (dctactwrk)
            GetChildrensAssignments tskTask, dctFndWrk, dctActWrk, dctTxtWrk, tskTargetTask, txtFeedback, tsvScale, booDeleteAssignments
            
            'Zet de waarden uit de aangemaakte dictionaries in assignments van de nieuwe resourcetaak
            For Each varResourceID In dctFndWrk.Keys
                Set assTargetAssignment = tskTargetTask.Assignments.Add(tskTargetTask.ID, varResourceID)
                'assTargetAssignment.Work = 0
                assTargetAssignment.Notes = dctTxtWrk(varResourceID)
                Set tsvTargetTSVsWrk = assTargetAssignment.TimeScaleData(tskTargetTask.Start, tskTargetTask.Finish, pjAssignmentTimescaledWork, tsvScale)
                Set tsvTargetTSVsActWrk = assTargetAssignment.TimeScaleData(tskTargetTask.Start, tskTargetTask.Finish, pjAssignmentTimescaledActualWork, tsvScale)

                'Gepland werk op 0 zetten
                For intTsvTeller = 0 To UBound(dctFndWrk(varResourceID))
                    tsvTargetTSVsWrk(intTsvTeller + 1) = 0
                Next


                For intTsvTeller = 0 To UBound(dctFndWrk(varResourceID))
                    'Als de geplande uren gevuld zijn
                        If Not IsEmpty(dctFndWrk(varResourceID)(intTsvTeller)) Then
                            'Vul de geplande uren
                            tsvTargetTSVsWrk(intTsvTeller + 1) = dctFndWrk(varResourceID)(intTsvTeller)
                            'Als de geplande uren gelijk zijn aan de gewerkte uren
                            If dctFndWrk(varResourceID)(intTsvTeller) = dctActWrk(varResourceID)(intTsvTeller) And Not IsEmpty(dctActWrk(varResourceID)(intTsvTeller)) Then
                                'Vul de gewerkte uren
                                tsvTargetTSVsActWrk(intTsvTeller + 1) = dctActWrk(varResourceID)(intTsvTeller)
                            End If
                    End If
                Next
            Next
            RollUpResources = "Operatie uitgevoerd: " & dctFndWrk.Count & " assignments in taak """ & tskTargetTask.Name & """ opgenomen." & vbCrLf & _
                vbCrLf & "Waarschuwingen:" & txtFeedback
        Else
            On Error GoTo 0
            RollUpResources = "Meer dan ��n taak geselecteerd"
        End If
    Else
        RollUpResources = "Werkt alleen in taakweergave"
    End If

GoTo Finish

ErrorHandler:
    RollUpResources = "Error"

Finish:
    On Error GoTo 0

End Function

Public Sub GetChildrensAssignments(ByVal tskTask As Task, dctFndWrk As Scripting.Dictionary, dctActWrk As Scripting.Dictionary, _
    dctTxtWrk As Scripting.Dictionary, tskTargetTask As Task, Optional txtFeedback As String = "", _
    Optional tsvScale = pjTimescaleWeeks, Optional booDelAssignment As Boolean = False)

    Dim assTaskAssignments As Assignments
    Dim assAssignment As Assignment
    Dim tskChildTask As Task
    Dim tsvSrcWrk As TimeScaleValues
    Dim tsvSrcActWrk As TimeScaleValues
    Dim intTeller As Integer
    Dim assTargetAssignment As Assignment
    Dim intAssTeller As Integer
    Dim intTsvTeller As Integer
    Dim varTmpArray() As Variant
    Dim varTargetFndTsvs() As Variant
    Dim varTargetActTsvs() As Variant
    Dim txtWrk As String
    Dim intWorkDiv As Integer
    
    Set assTaskAssignments = tskTask.Assignments
    If tskTask.ID <> tskTargetTask.ID Then
        For intAssTeller = 1 To tskTask.Assignments.Count 'Doorloop alle assignments van de taak
            Set assAssignment = tskTask.Assignments(intAssTeller)
            If assAssignment.CostRateTable <> 0 Then
                txtFeedback = txtFeedback & vbCrLf & "De assignment van """ & assAssignment.ResourceName & """ op taak """ & tskTask.Name & _
                    """ heeft een afwijkende costrate, de totaalkosten van """ & _
                    tskTargetTask.Name & """ kunnen daardoor afwijken van de oorspronkelijke kosten."
            End If
            
            intWorkDiv = IIf(assAssignment.ResourceType = pjResourceTypeWork, 60, 1)
            
            Set tsvSrcWrk = assAssignment.TimeScaleData(tskTargetTask.Start, tskTargetTask.Finish, pjAssignmentTimescaledWork, tsvScale) 'Haal alle werk timescalevalues op vd assignment
            Set tsvSrcActWrk = assAssignment.TimeScaleData(tskTargetTask.Start, tskTargetTask.Finish, pjAssignmentTimescaledActualWork, tsvScale) 'Haal alle werkelijke gewerkte timescalevalues op vd assignment
            
            If Not dctFndWrk.Exists(assAssignment.ResourceID) Then 'Controleer of de resource van de ID al bestaat in de target
                'Zoniet, maak nieuwe arrays aan voor het opslaan van de geplande en gewerkte uren en de tekstnotes waar in de
                'brongegevens worden opgeslagen en koppel deze aan de dictionary vd resource
                ReDim varTargetFndTsvs(tsvSrcWrk.Count - 1)
                dctFndWrk.Add assAssignment.ResourceID, varTargetFndTsvs
                ReDim varTargetActTsvs(tsvSrcActWrk.Count - 1)
                dctActWrk.Add assAssignment.ResourceID, varTargetActTsvs
                
                txtWrk = "Oorspronkelijke assignments (kan via clipboard in Excel geplakt worden)" & vbCrLf & _
                    "Resource" & vbTab & "Taaknaam" & vbTab & "CategorieWrk" & vbTab & "Begindatum" & vbTab & "Einddatum"
                For intTsvTeller = 1 To tsvSrcWrk.Count
                    txtWrk = txtWrk & vbTab & tsvSrcWrk(intTsvTeller).StartDate
                Next
                dctTxtWrk.Add assAssignment.ResourceID, txtWrk
            End If
            
            'De arrays kunnen niet direct in de dictionary bewerkt worden, dus worden ze tijdelijk opgeslagen in variants
            varTargetFndTsvs = dctFndWrk(assAssignment.ResourceID)
            varTargetActTsvs = dctActWrk(assAssignment.ResourceID)
            txtWrk = dctTxtWrk(assAssignment.ResourceID)
            
            'Tel de timescalevalues (alleen de gevulde) op bij de reeds bestaande timescalevalues van de target
            For intTsvTeller = 1 To tsvSrcWrk.Count
                If tsvSrcWrk(intTsvTeller) <> "" Then
                    varTargetFndTsvs(intTsvTeller - 1) = varTargetFndTsvs(intTsvTeller - 1) + CDbl(tsvSrcWrk(intTsvTeller))
                    If tsvSrcWrk(intTsvTeller) <> tsvSrcActWrk(intTsvTeller) Then
                        txtFeedback = txtFeedback & vbCrLf & "Gewerkte " & IIf(tsvSrcActWrk(intTsvTeller) = "", 0, tsvSrcActWrk(intTsvTeller)) / intWorkDiv & " uren van """ & _
                            assAssignment.ResourceName & """ op taak """ & tskTask.Name & """ in de periode " & _
                            tsvSrcWrk(intTsvTeller).StartDate & " tm " & tsvSrcWrk(intTsvTeller).EndDate & _
                            " zijn genegeerd, omdat deze niet gelijk zijn aan de geplande uren (" & _
                            tsvSrcWrk(intTsvTeller) / intWorkDiv & ")."
                    End If
                End If
                If tsvSrcActWrk(intTsvTeller) <> "" Then
                    varTargetActTsvs(intTsvTeller - 1) = varTargetActTsvs(intTsvTeller - 1) + CDbl(tsvSrcActWrk(intTsvTeller))
                End If
            Next
            
            'Voeg regel toe aan notes met geplande uren
            txtWrk = txtWrk & vbCrLf & assAssignment.ResourceName & vbTab & tskTask.Name & vbTab & "Gepland werk" & vbTab & _
                assAssignment.Start & vbTab & assAssignment.Finish
            For intTsvTeller = 0 To UBound(varTargetFndTsvs)
                If tsvSrcWrk(intTsvTeller + 1) = "" Then
                    txtWrk = txtWrk & vbTab & tsvSrcWrk(intTsvTeller + 1)
                Else
                    txtWrk = txtWrk & vbTab & tsvSrcWrk(intTsvTeller + 1) / intWorkDiv
                End If
            Next
            
            'Voeg regel toe aan notes met gewerkte uren
            txtWrk = txtWrk & vbCrLf & assAssignment.ResourceName & vbTab & tskTask.Name & vbTab & "Gewerkt werk" & vbTab & _
                assAssignment.Start & vbTab & assAssignment.Finish
            For intTsvTeller = 0 To UBound(varTargetActTsvs)
                If tsvSrcActWrk(intTsvTeller + 1) = "" Then
                    txtWrk = txtWrk & vbTab & tsvSrcActWrk(intTsvTeller + 1)
                Else
                    txtWrk = txtWrk & vbTab & tsvSrcActWrk(intTsvTeller + 1) / intWorkDiv
                End If
            Next
            
            'Gewijzigde arrays weer in de dictionary opslaan
            dctFndWrk(assAssignment.ResourceID) = varTargetFndTsvs
            dctActWrk(assAssignment.ResourceID) = varTargetActTsvs
            dctTxtWrk(assAssignment.ResourceID) = txtWrk
            
        Next
                
        If booDelAssignment Then
            For intAssTeller = 1 To tskTask.Assignments.Count
                tskTask.Assignments(1).Delete
            Next
        End If
                
        'Doe hetzelfde voor alle onderliggende taken
        For Each tskChildTask In tskTask.OutlineChildren
            GetChildrensAssignments tskChildTask, dctFndWrk, dctActWrk, dctTxtWrk, tskTargetTask, txtFeedback, tsvScale, booDelAssignment
        Next
        
    End If
End Sub



