VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'https://www.google.nl/search?q=vba+custom+object+use+class+as+property&rlz=1C9BKJA_enNL790NL790&oq=vba+custom+object+use+class+as+property&aqs=chrome..69i57.48046j0j7&hl=nl&sourceid=chrome-mobile&ie=UTF-8
Option Explicit
Private pCells() As cCell

Private Sub Class_Initialize()
End Sub

Public Sub Create(intNrOfRows As Integer, intNrOfColumns As Integer)
Dim intCounterRow As Integer
Dim intCounterColumn As Integer

ReDim pCells(intNrOfRows - 1, intNrOfColumns - 1)

For intCounterRow = 0 To intNrOfRows - 1
    For intCounterColumn = 0 To intNrOfColumns - 1
        Set pCells(intCounterRow, intCounterColumn) = New cCell
    Next
Next
        

End Sub

Public Property Let CellValue(intRow As Integer, intcolumn As Integer, pCellvalue)
    If pCells(intRow, intcolumn) Is Nothing Then
        Debug.Print
    End If
    pCells(intRow, intcolumn).Value = pCellvalue
End Property

Public Property Get CellValue(intRow As Integer, intcolumn As Integer) As Variant
    CellValue = pCells(intRow, intcolumn).Value
End Property


