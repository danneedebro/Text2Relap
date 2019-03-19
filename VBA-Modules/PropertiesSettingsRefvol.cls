VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertiesSettingsRefvol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PropertiesSettingsRefvol.cls
'
'
Option Explicit
'
' MEMBER VARIABLES
Private pHydroSystem As HydroSystem

Public CCC As Integer
Public Volume As Integer
Public Elevation As Double
Public SystemName As String
Public Fluid As String
Public SystemInformationFlag As Integer
'
'
' PROPERTIES
Private Property Get Inputdeck() As Text2Relap
    Set Inputdeck = pHydroSystem.Inputdeck
End Property

Public Property Get ReferenceVolume() As String
    ReferenceVolume = Format(Me.CCC, "#000") & Format(Me.Volume, "#00") & "0000"
End Property
'
'
' CLASS EVENT METHODS
Public Sub Create(inputRows() As Variant, ByRef ParentSystem As HydroSystem)
    On Error GoTo ErrorHandler

    ' Associate with parent hydro system
    Set pHydroSystem = ParentSystem
    
    CCC = Inputdeck.CheckInput(inputRows(1, 2), Numeric:=True, valueDefault:=1, causeError:=True, FieldName:="Refvol CCC", word:=2)
    Volume = Inputdeck.CheckInput(inputRows(1, 3), Numeric:=True, valueDefault:=1, causeError:=True, FieldName:="Refvol Volume", word:=3)
    Elevation = Inputdeck.CheckInput(inputRows(1, 4), Numeric:=True, valueDefault:=0#, causeError:=True, FieldName:="Refvol Elevation", word:=4)
    SystemName = inputRows(1, 5)
    Fluid = Inputdeck.CheckInput(inputRows(1, 6), Numeric:=False, valueDefault:="h2o", causeError:=True, FieldName:="Refvol Fluid type", word:=6)
    
    Exit Sub
ErrorHandler:
    err.Raise err.Number, TypeName(Me) & ".Create, " & err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub
'
'
' CLASS METHODS

'
'
' STATIC METHODS




