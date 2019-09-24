VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
Public Enum R5Component
    HydroComp
    Custom
    TripVariable
    TripLogical
    Cntrlvar
    Header
    Comment1
    Comment2
    Comment3
    Force
End Enum

Public Enum CntrlvarTypes
    CntrlvarSum
    CntrlvarDiv
    CntrlvarMult
    CntrlvarDiffrend
End Enum

Public Sub WriteToFile(ByRef ts As Object, ByVal wC As Boolean)
End Sub

Public Function ComponentInfo() As String
End Function

Public Property Get ObjectType() As R5Component
End Property

Public Property Get RowBegin() As Integer
End Property

Public Property Get RowEnd() As Integer
End Property
