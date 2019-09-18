VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionConnections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Private pConnections As Collection
'
'
' PROPERTIES

'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set pConnections = New Collection
End Sub

Public Sub Add(ByVal NewConnection As PropertiesHydroCompConnection)
    pConnections.Add NewConnection
End Sub

Public Function Item(index As Variant) As PropertiesHydroCompConnection
Attribute Item.VB_UserMemId = 0
    Set Item = pConnections(index)
End Function

Public Function Count() As Long
    Count = pConnections.Count
End Function
'
'
' CLASS METHODS

'
'
' STATIC METHODS

