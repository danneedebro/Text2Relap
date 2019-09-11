VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertiesHydroCompSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PropertiesHydroCompSegment.cls
'
'
Option Explicit
'
'
' MEMBER VARIABLES
Public Volumes As Integer
Public VolumeFirst As Integer
Public VolumeLast As Integer
Public Junctions As Integer
Public JunctionFirst As Integer
Public JunctionLast As Integer
Public ID As String
Public Dx As Double
Public Area As Double
Public AngleVert As Double
Public AngleAzi As Double
Public KLossPos As Double
Public KLossNeg As Double
Public VolumeM2 As Double
Public Elevchange As Double
Public Name As String
Public Drawing As String
Public ForceNumber As Integer
Public Roughness As Double
Public VFlags As String
Public JFlags As String
Public Init As PropertiesHydroCompInit
'
'
' PROPERTIES
Public Property Get length() As String
    With Me
        length = .Dx * (.VolumeLast - .VolumeFirst + 1)
    End With
End Property
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set Init = New PropertiesHydroCompInit
End Sub
'
'
' CLASS METHODS
Public Function Repr() As String
    With Me
        Repr = "ID=" & .ID & ", VolumeFirst=" & .VolumeFirst & ", VolumeLast=" & .VolumeLast & _
               ", Length=" & .length & ", dx=" & .Dx & ", Area=" & .Area
    End With
End Function
'
'
' STATIC METHODS











