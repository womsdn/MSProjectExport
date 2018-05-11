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
Public Property Let CellValue(intRow As Integer, intcolumn As Integer, pCellvalue)
    If pCells(intRow - 1, intcolumn - 1) Is Nothing Then
        Debug.Print
    End If
    pCells(intRow - 1, intcolumn - 1).Value = pCellvalue
End Property
Public Property Get CellValue(intRow As Integer, intcolumn As Integer) As Variant
    CellValue = pCells(intRow - 1, intcolumn - 1).Value
End Property
Public Sub Create(intNrOfRows As Integer, intNrOfColumns As Integer)
    Dim intCounterRow As Integer
    Dim intCounterColumn As Integer
    
    ReDim pCells(intNrOfRows, intNrOfColumns)
    
    For intCounterRow = 0 To intNrOfRows - 1
        For intCounterColumn = 0 To intNrOfColumns - 1
            Set pCells(intCounterRow, intCounterColumn) = New cCell
        Next
    Next
End Sub
Public Sub Dump()
    Dim intCounterRow As Integer
    Dim intCounterColumn As Integer
    Dim strRow As String
    
    For intCounterRow = 0 To UBound(pCells, 1) - 1
        strRow = ""
        For intCounterColumn = 0 To UBound(pCells, 2) - 1
            strRow = strRow & IIf(strRow = "", "", vbTab) & pCells(intCounterRow, intCounterColumn).Value
        Next
        Debug.Print strRow
    Next
    
    
End Sub
