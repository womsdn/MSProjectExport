VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPlanningPerRscProdExcel 
   Caption         =   "Planning per resource per product"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   OleObjectBlob   =   "frmPlanningPerRscProdExcel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPlanningPerRscProdExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnEsc_Click()
    Unload Me
End Sub

Private Sub btnStart_Click()
    If Me.chkTimeScaleData Then
        TotalsPerResourcePerPhase Me.chkFilterResources, Me.chkOutlineNumbers, CDate(Me.txtBegindatum), CDate(Me.txtEinddatum), Int(Me.cmbTimeScaleData), Int(Me.cmbEenheid), Me.chkColorMarkings
    Else
        TotalsPerResourcePerPhase Me.chkFilterResources, Me.chkOutlineNumbers, booColorMarkings:=Me.chkColorMarkings
    End If
End Sub

Private Sub chkFilterResources_Change()
    Call VulData
End Sub

Private Sub chkTimeScaleData_Click()
    If Me.chkTimeScaleData Then
        Me.cmbEenheid.Enabled = True
        Me.cmbTimeScaleData.Enabled = True
        Me.txtBegindatum.Enabled = True
        Me.txtEinddatum.Enabled = True
    Else
        Me.cmbEenheid.Enabled = False
        Me.cmbTimeScaleData.Enabled = False
        Me.txtBegindatum.Enabled = False
        Me.txtEinddatum.Enabled = False
    End If
End Sub


Private Sub UserForm_Initialize()

    Dim tskTask As Task
    Dim intTeller As Integer
    Dim rcsResources As Resources

    'Zet standaard de optie aan voor gedetailleerde planningsdata
    Me.chkTimeScaleData = True
    
    'Zet de kleurmarkeringen standaard uit
    Me.chkColorMarkings = False
            
    'Zet de optie die aangeeft dat de outlinenumbers opgenomen moeten worden standaard aan.
    chkOutlineNumbers = True
    
    'Stel de keuzelijst met eenheden in
    Me.cmbEenheid.AddItem (pjTimescaleMonths)
    Me.cmbEenheid.Column(1, 0) = ("Per maand")
    Me.cmbEenheid.AddItem (pjTimescaleWeeks)
    Me.cmbEenheid.Column(1, 1) = ("Per week")
    Me.cmbEenheid.AddItem (pjTimescaleYears)
    Me.cmbEenheid.Column(1, 2) = ("Per jaar")
    Me.cmbEenheid.Value = pjTimescaleMonths 'Default keuze
    
    'Stel de keuzelijst met op te vragen data in
    Me.cmbTimeScaleData.AddItem (pjAssignmentTimescaledWork)
    Me.cmbTimeScaleData.Column(1, 0) = "Werk"
    Me.cmbTimeScaleData.AddItem (pjAssignmentTimescaledActualWork)
    Me.cmbTimeScaleData.Column(1, 1) = "Werkelijke hoeveelheid werk"
    Me.cmbTimeScaleData.AddItem (pjAssignmentTimescaledBaselineWork)
    Me.cmbTimeScaleData.Column(1, 2) = "Baseline hoeveelheid werk"
    Me.cmbTimeScaleData.AddItem (pjAssignmentTimescaledCost)
    Me.cmbTimeScaleData.Column(1, 3) = "Kosten"
    
    Me.cmbTimeScaleData.Value = pjAssignmentTimescaledWork
        
    'Als er resources geselecteerd zijn, gaat het om die
    If ActiveProject.Views(ActiveProject.CurrentView).Type = pjResourceItem Then
        If ActiveSelection.Resources.Count > 1 Then
            chkFilterResources = True
        Else
            chkFilterResources = False
        End If
    ElseIf ActiveProject.Views(ActiveProject.CurrentView).Type = pjTaskItem Then
        If ActiveSelection.Tasks.Count > 1 Then
            chkFilterResources = True
        Else
            chkFilterResources = False
        End If
    End If
    chkFilterResources.Visible = True
    Call VulData

End Sub
Public Sub VulData()
    
    Dim datFirstDate As Date
    Dim datLastDate As Date
    Dim tskTasks As Tasks
    Dim tsk As Task
    Dim rcsResources As Resources
    Dim rcs As Resource
    Dim ass As Assignment
    
    If ActiveProject.Views(ActiveProject.CurrentView).Type = pjTaskItem Then
        If Me.chkFilterResources Then
            Set tskTasks = ActiveSelection.Tasks
        Else
            Set tskTasks = ActiveProject.Tasks
        End If
        
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
        If Me.chkFilterResources Then
            Set rcsResources = ActiveSelection.Resources
        Else
            Set rcsResources = ActiveProject.Resources
        End If
        datFirstDate = rcsResources(1).Assignments(1).Start
        datLastDate = rcsResources(1).Assignments(1).Finish
        For Each rcs In rcsResources
            For Each ass In rcs.Assignments
                If ass.Start < datFirstDate Then
                    datFirstDate = ass.Start
                End If
                If ass.Finish > datLastDate Then
                    datLastDate = ass.Finish
                End If
            Next
        Next
    End If
    
    Me.txtBegindatum = datFirstDate
    Me.txtEinddatum = datLastDate
 
End Sub

Private Sub txtBegindatum_AfterUpdate()
    On Error GoTo errhandler
    
    Dim datTestdate As Date
    
    datTestdate = CDate(Me.txtBegindatum)
    
Exit Sub
    
errhandler:
    MsgBox "Geen geldige datum"
    On Error Resume Next
    Call VulData
End Sub
Private Sub txtBegindatum_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    txtBegindatum = Format(Now() - Weekday(Now(), vbMonday) + 1, "d-m-yyyy") & " 08:00:00"
End Sub

Private Sub txtEinddatum_AfterUpdate()
    On Error GoTo errhandler
    
    Dim datTestdate As Date
    
    datTestdate = CDate(Me.txtEinddatum)
    
Exit Sub
    
errhandler:
    MsgBox "Geen geldige datum"
    On Error Resume Next
    Call VulData
End Sub
