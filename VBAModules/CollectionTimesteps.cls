VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionTimesteps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Private pTimesteps As Collection

Private pHydroSystem As HydroSystem
'
'
' PROPERTIES
Public Property Get HydroSystem() As HydroSystem
    Set HydroSystem = pHydroSystem
End Property

Public Property Set HydroSystem(arg As HydroSystem)
    Set pHydroSystem = arg
End Property
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set pTimesteps = New Collection
End Sub

Public Sub Add(ByRef inputRows() As Variant)
    Dim NewTimestep As New PropertiesSettingsTimestep
    NewTimestep.Create inputRows, HydroSystem
    pTimesteps.Add NewTimestep
End Sub

Public Function Item(index As Variant) As PropertiesSettingsTimestep
Attribute Item.VB_UserMemId = 0
    Set Item = pTimesteps(index)
End Function

Public Function Count() As Integer
    Count = pTimesteps.Count
End Function
'
'
' CLASS METHODS

'
'
' STATIC METHODS

