Attribute VB_Name = "Module1"
Sub ShowCreateTIP()
    frmInformatieTIPExporteren.Show
End Sub
Sub ShowBoekingsregels()
    frmBoekingsregels.Show
End Sub
Sub CalcInternCostsMacro()
    frmCalcInternalCost.Show
End Sub
Sub CallCopyTasks()
    Call CopyTasks
End Sub
Sub CallFilter_Select()
    Call Filter_Select
End Sub
Sub CallFilterRelatedTasks()
    Call FilterRelatedTasks
End Sub
Sub CallPlanningPerRscProdExcel()
    frmPlanningPerRscProdExcel.Show
End Sub
Sub CallPlanningPerRscProdCSV()
    frmPlanningPerRscProdCSV.Show
End Sub
Sub CallPlanningPerRscExcel()
    frmPlanningPerRscExcel.Show
End Sub
Sub CallPlanningPerRscCSV()
    frmPlanningPerRscCSV.Show
End Sub
Sub CallFilterCurrentTasks()
    frmPickCurrentDate.Show
End Sub
Sub CallOWSKostenNaarExcel()
    Call OWSKostenNaarExcel
End Sub
Sub CallTaskExportExcel()
    Call frmPlanningperTskExcel.Show
End Sub
Sub CallRollUpResources()
    Call frmRollupResources.Show
End Sub

