VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComponentCntrlvar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ComponentCntrlvar.cls
'
'
Option Explicit

Implements iComponent

'
'
' MEMBER VARIABLES
Private pHydroSystem As HydroSystem

Private pCntrlvarType As Integer
Private pCntrlvarNumber As Integer
Private pID As String
Public ScalingFactor As Double
Public InitialValue As Double
Public InitialValueFlag As Integer
Public LimiterControl As Integer
Public MinValue As Double
Public MaxValue As Double
Public Text As String

Private pRowBegin As Integer
Private pRowEnd As Integer
'
'
' SHARED PROPERTIES AND METHODS
Private Property Get iComponent_ObjectType() As R5Component
    iComponent_ObjectType = Cntrlvar
End Property

Private Function iComponent_ComponentInfo() As String
    iComponent_ComponentInfo = "cntrlvar no: " & Me.CntrlvarNumber
End Function

Private Property Get iComponent_RowBegin() As Integer
    iComponent_RowBegin = pRowBegin
End Property

Private Property Get iComponent_RowEnd() As Integer
    iComponent_RowEnd = pRowEnd
End Property

Private Sub iComponent_WriteToFile(ts As Object, ByVal wC As Boolean)
    WriteToFile ts, wC
End Sub
'
'
' PROPERTIES
Public Property Get HydroSystem() As HydroSystem
    Set HydroSystem = pHydroSystem
End Property

Public Property Get Settings() As PropertiesSettings
    Set Settings = HydroSystem.Settings
End Property

Public Property Get CntrlvarNumber() As Integer
    CntrlvarNumber = pCntrlvarNumber
End Property

Public Property Let CntrlvarNumber(arg As Integer)
    pCntrlvarNumber = arg
End Property

Public Property Get ID() As String
    ID = pID
End Property
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    
End Sub

Public Sub Create(inputRows() As Variant, ByRef ParentSystem As HydroSystem, ByVal RowBegin As Integer, ByVal RowEnd As Integer)
    On Error GoTo ErrorHandler
    Set pHydroSystem = ParentSystem
    pRowBegin = RowBegin
    pRowEnd = RowEnd
        
    pID = inputRows(1, 2)
    CntrlvarNumber = HydroSystem.Settings.CurrentCntrlvarNum
    
    Exit Sub
ErrorHandler:
    err.Raise err.Number, TypeName(Me) & ".Create, " & err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub
'
'
' CLASS METHODS
Public Sub WriteToFile(ByRef ts As Object, ByVal wC As Boolean)
' Action: Writes the force components
'
' Inputs:   ts = A TextStream object (included in the MS Scripting Runtime)
'           wC = writeComments, True if comments are to be written, false otherwise
'
    Dim s As New ResourceSprintf
    Dim i As Integer, j As Integer, k As Integer
    
    Dim CntrlvarStr As String
    Select Case Settings.CntrlvarFormat
        Case 999
            CntrlvarStr = "205%03d%02d"
        Case 9999
            CntrlvarStr = "205%04d%01d"
    End Select  ' CCCVV0000
    
    ts.WriteLine s.sprintf(CntrlvarStr & "  %-8s  %-8s  %7.3g  %7.3g  %1d", CntrlvarNumber, 0, ID, pCntrlvarType, ScalingFactor, InitialValue, InitialValueFlag)
    
    
End Sub


'
'
' STATIC METHODS






