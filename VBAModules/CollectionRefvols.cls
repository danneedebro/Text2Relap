VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionRefvols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Private pRefvols As Collection

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
    Set pRefvols = New Collection
End Sub

Public Sub Add(ByRef inputRows() As Variant)
    Dim NewRefvol As New PropertiesSettingsRefvol
    NewRefvol.Create inputRows, HydroSystem
    pRefvols.Add NewRefvol
End Sub

Public Function Item(index As Variant) As PropertiesSettingsRefvol
Attribute Item.VB_UserMemId = 0
    Set Item = pRefvols(index)
End Function

Public Function Count() As Integer
    Count = pRefvols.Count
End Function
'
'
' CLASS METHODS

'
'
' STATIC METHODS

