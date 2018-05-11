VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim pValue As Variant
Dim pBold As Boolean

Property Let Value(sValue As Variant)
    pValue = sValue
End Property

Property Get Value()
    Value = pValue
End Property

Private Sub Class_Initialize()
    pBold = False
End Sub
