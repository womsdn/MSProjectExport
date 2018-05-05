VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPlanningPerRscExcel 
   Caption         =   "Planning per resource"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "frmPlanningPerRscExcel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPlanningPerRscExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub booUitgebreid_Change()
    If booUitgebreid Then
        Me.booMetFunctie.Visible = False
    Else
        Me.booMetFunctie.Visible = True
    End If
    
End Sub
Private Sub btnEsc_Click()
    Unload Me
End Sub

Private Sub btnStart_Click()
    If Me.cmbFilter <> "" Then
        ResourceExportExcel Me.cmbFilter.Column(2), CDate(Me.txtBegindatum), CDate(Me.txtEinddatum), Me.cmbTimeScaleData, Me.cmbEenheid, Me.chkFilterResources, , , Not Me.booUitgebreid, Me.booMetFunctie, Me.booGemarkeerd
    Else
        MsgBox ("Geen taak/product geselecteerd")
    End If
'    Unload Me
End Sub


Private Sub cmbFilter_Change()
    Call VulData
End Sub

Private Sub txtBegindatum_AfterUpdate()
    On Error GoTo errhandler
    
    Dim datTestdate As Date
    
    datTestdate = CDate(Me.txtBegindatum)
    
Exit Sub
    
errhandler:
    MsgBox "Geen geldige datum"
    On Error Resume Next
    Me.txtBegindatum = ActiveProject.Tasks(Val(Me.cmbFilter.Column(2))).Start
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
    Me.txtEinddatum = ActiveProject.Tasks(Val(Me.cmbFilter.Column(2))).Finish
End Sub


Private Sub UserForm_Initialize()

    Dim tskTask As Task
    Dim intTeller As Integer
    Dim rcsResources As Resources

    'Vul de selectie keuzelijst met de taken in het project
    For Each tskTask In ActiveProject.Tasks
        If Not tskTask Is Nothing Then 'Controleer of er geen sprake is van een lege taak
            Me.cmbFilter.AddItem (tskTask.OutlineNumber)
            Me.cmbFilter.Column(1, intTeller) = tskTask.OutlineNumber & " - " & tskTask.Name
            Me.cmbFilter.Column(2, intTeller) = tskTask.ID
            intTeller = intTeller + 1
        End If
    Next
    
    'Als de keuzelijst gevuld is, stel dan de default keuze in
    If intTeller > 0 Then
        Me.cmbFilter.Value = Me.cmbFilter.Column(0, 0)
        Me.cmbFilter.SetFocus
        Me.cmbFilter.SelStart = 0
        Me.cmbFilter.SelLength = Len(Me.cmbFilter.Column(1, 0))
    End If
    
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
    Me.cmbTimeScaleData.AddItem (pjAssignmentTimescaledActualCost)
    Me.cmbTimeScaleData.Column(1, 4) = "Werkelijke kosten"
    Me.cmbTimeScaleData.AddItem (pjAssignmentTimescaledBaselineCost)
    Me.cmbTimeScaleData.Column(1, 5) = "Baseline hoeveelheid kosten"
    Me.cmbTimeScaleData.Value = pjAssignmentTimescaledWork
        
    'Bepaal de resources over welke de begin- en einddatum opgehaald moeten worden
    On Error GoTo MustBeTasksThen
    
    'Als er resources geselecteerd zijn, gaat het om die
    If ActiveSelection.Resources.Count > 1 Then
        chkFilterResources = True
    Else
        chkFilterResources = False
    End If
    chkFilterResources.Visible = True
    Set rcsResources = ActiveSelection.Resources
    Call VulData
    
    '
    
    Exit Sub
    
MustBeTasksThen:
    'Als er geen resources geselecteerd zijn, moeten alle resources genomen worden
    chkFilterResources.Visible = False
    Set rcsResources = ActiveProject.Resources
    Call VulData

End Sub
Public Sub VulData()

    On Error Resume Next
    Me.txtBegindatum = ActiveProject.Tasks(Val(Me.cmbFilter.Column(2))).Start
    Me.txtEinddatum = ActiveProject.Tasks(Val(Me.cmbFilter.Column(2))).Finish
 
End Sub


