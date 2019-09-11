VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertiesSettingsTimestep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PropertiesSettingsTimestep.cls
'
'
Option Explicit
'
'
' MEMBER VARIABLES
Private pHydroSystem As HydroSystem

Public TimeEnd As Double
Public TimestepMin As Double
Public TimestepMax As Double
Public ControlOption As String
Public FrequencyMinor As Long
Public FrequencyMajor As Long
Public FrequencyRestart As Long
'
'
' PROPERTIES
Private Property Get Inputdeck() As Text2Relap
    Set Inputdeck = pHydroSystem.Inputdeck
End Property
'
'
' CLASS EVENT METHODS
Public Sub Create(inputRows() As Variant, ByRef ParentSystem As HydroSystem)
    On Error GoTo ErrorHandler

    ' Associate with parent hydro system
    Set pHydroSystem = ParentSystem
    
    ' Set defaults
    With Me
        .TimeEnd = 10
        .TimestepMin = 0.000000000001
        .TimestepMax = 0.01
        .ControlOption = "11001"
        .FrequencyMinor = 10
        .FrequencyMajor = 1000000
        .FrequencyRestart = Round(0.5 / .TimestepMax, 0)
    End With
    
    ' Read input
    TimeEnd = Inputdeck.CheckInput(inputRows(1, 2), Numeric:=True, valueDefault:=1#, causeError:=True, FieldName:="Timestep TimeEnd", word:=2)
    TimestepMax = Inputdeck.CheckInput(inputRows(1, 3), Numeric:=True, valueDefault:=0.01, causeError:=True, FieldName:="Timestep dtMax", word:=3)
    FrequencyMinor = Inputdeck.CheckInput(inputRows(1, 4), Numeric:=True, valueDefault:=10, causeError:=True, FieldName:="Timestep fMinor", word:=4)
    FrequencyRestart = Inputdeck.CheckInput(inputRows(1, 5), Numeric:=True, valueDefault:=10, causeError:=True, FieldName:="Timestep fRestart", word:=5)
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".Create, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
'
'
' CLASS METHODS

'
'
' STATIC METHODS


