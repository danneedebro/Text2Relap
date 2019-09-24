VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertiesHydroCompInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PropertiesHydroCompInit.cls
'
'
Option Explicit
'
'
' MEMBER VARIABLES
Private pHydroSystem As HydroSystem

Public EbtFlag As String
Public Pressure As Double
Public Temperature As Double
Public Quality As Double
Public param1 As String
Public ID As String
Public FlowLiq As Double
Public FlowGas As Double
Public InterfVel As Double
Public InitVelocity As Boolean     ' if true 'flowLiq' and 'flowGas' is given as velocities instead of massflows
Public CntrlWord As Integer
'
'
' PROPERTIES
Private Property Get InputDeck() As Text2Relap
    Set InputDeck = pHydroSystem.InputDeck
End Property

'
'
' CLASS EVENT METHODS
Public Sub Create(inputRow() As Variant, ByRef ParentSystem As HydroSystem)
        
    ' Associate with parent hydro system
    Set pHydroSystem = ParentSystem
    
    ' Set defaults
    Me.CntrlWord = 1
    
    Dim PressureRead As Double, TemperatureRead As Double, QualityRead As Double
    
    
    PressureRead = InputDeck.CheckInput(inputRow(1, 2), Numeric:=True, valueDefault:=100000#, causeError:=True, FieldName:="Pressure", word:=2)
    TemperatureRead = InputDeck.CheckInput(inputRow(1, 3), Numeric:=True, valueDefault:=293.15, causeError:=True, FieldName:="Temperature", word:=3)
    
    Me.ID = inputRow(1, 4)
    
    If LCase(inputRow(1, 1)) = "init" Then
        EbtFlag = "103"
    ElseIf LCase(inputRow(1, 1)) = "initgas" Then
        EbtFlag = "104"
        Pressure = PressureRead
        Temperature = TemperatureRead
        If IsEmpty(inputRow(1, 5)) Then
            Quality = 1#
        Else
            Me.Quality = InputDeck.CheckInput(inputRow(1, 5), Numeric:=True, valueDefault:=1#, causeError:=True, FieldName:="Quality", word:=5)
        End If
        
    End If
    
    If TemperatureRead >= 0 And TemperatureRead <= 1 And Me.EbtFlag = "103" Then
        Me.EbtFlag = "102"
        Me.Quality = TemperatureRead
        Me.Pressure = PressureRead
    ElseIf PressureRead >= 0 And PressureRead <= 1 And Me.EbtFlag = "103" Then
        Me.EbtFlag = "101"
        Me.Quality = Pressure
        Me.Temperature = TemperatureRead
    Else
        Me.Pressure = PressureRead
        Me.Temperature = TemperatureRead
    End If
    
End Sub
'
'
' CLASS METHODS

'
'
' STATIC METHODS



