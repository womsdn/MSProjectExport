VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Project_Open(ByVal pj As Project)
    CreateCustomRibbon
End Sub

Private Sub DeleteCustomRibbon()
ribbonXml = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">" & "<mso:ribbon></mso:ribbon></mso:customUI>"
ActiveProject.SetCustomUI (ribbonXml)
End Sub
Private Sub CreateCustomRibbon()
    Dim ribbonXml As String
    
    ribbonXml = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">"
    ribbonXml = ribbonXml + "  <mso:ribbon>"
    ribbonXml = ribbonXml + "    <mso:qat/>"
    ribbonXml = ribbonXml + "    <mso:tabs>"
    ribbonXml = ribbonXml + "      <mso:tab id=""highlightTab"" label=""Macros"" insertBeforeQ=""mso:TabFormat"">"
    ribbonXml = ribbonXml + "        <mso:group id=""Filters"" label=""Filters"" autoScale=""true"">"
    ribbonXml = ribbonXml + "          <mso:button id=""ResetFilter"" label=""Reset Filter"" imageMso=""FilterToggleFilter"" onAction=""ResetFilter""/>"
    ribbonXml = ribbonXml + "          <mso:button id=""Filter_Select"" label=""Geselecteerde taken"" imageMso=""FilterToggleFilter"" onAction=""Filter_Select""/>"
    ribbonXml = ribbonXml + "          <mso:button id=""FilterCurrentResourcesOverlap"" label=""Taken met overlappende resources"" imageMso=""FilterToggleFilter"" onAction=""FilterCurrentResourcesOverlap""/>"
    ribbonXml = ribbonXml + "          <mso:button id=""FilterCurrentTasks"" label=""Actuele taken"" imageMso=""FilterToggleFilter"" onAction=""CallFilterCurrentTasks""/>"
    ribbonXml = ribbonXml + "          <mso:button id=""FilterCurrentTasksOverlap"" label=""Overlappende taken"" imageMso=""FilterToggleFilter"" onAction=""FilterCurrentTasksOverlap""/>"
    ribbonXml = ribbonXml + "          <mso:button id=""FilterRelatedTasks"" label=""Gerelateerde taken"" imageMso=""FilterToggleFilter"" onAction=""FilterRelatedTasks""/>"
    ribbonXml = ribbonXml + "          <mso:button id=""CallRollUpResources"" label=""Cumuleer resource assignments..."" imageMso=""CollapseAll"" onAction=""CallRollUpResources""/>"
    ribbonXml = ribbonXml + "        </mso:group>"
    ribbonXml = ribbonXml + "        <mso:group id=""Export"" label=""Export"" autoScale=""true"">"
    ribbonXml = ribbonXml + "          <mso:button id=""CallPlanningPerRscProdExcel"" label=""Planning per resource/product"" imageMso=""ExportExcel"" onAction=""CallPlanningPerRscProdExcel""/>"
    ribbonXml = ribbonXml + "          <mso:button id=""CallPlanningPerRscExcel"" label=""Planning per resource"" imageMso=""ExportExcel"" onAction=""CallPlanningPerRscExcel""/>"
    ribbonXml = ribbonXml + "          <mso:button id=""CallTaskExportExcel"" label=""Planning per taak"" imageMso=""ExportExcel"" onAction=""CallTaskExportExcel""/>"
    ribbonXml = ribbonXml + "          <mso:button id=""CallCopyTasks"" label=""Kopieer naar klembord"" imageMso=""Copy"" onAction=""CallCopyTasks""/>"
    ribbonXml = ribbonXml + "          <mso:button id=""CallOWSKostenNaarExcel"" label=""Exporteer OWS uren naar Excel"" imageMso=""ExportExcel"" onAction=""CallOWSKostenNaarExcel""/>"
    ribbonXml = ribbonXml + "        </mso:group>"
    ribbonXml = ribbonXml + "      </mso:tab>"
    ribbonXml = ribbonXml + "    </mso:tabs>"
    ribbonXml = ribbonXml + "  </mso:ribbon>"
    ribbonXml = ribbonXml + "</mso:customUI>"
    
   
    ActiveProject.SetCustomUI (ribbonXml)

End Sub










