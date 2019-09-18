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

'Private pCntrlvarType As Integer
Private pCCC As Integer
Private pID As String
Private pCntrlvarType As CntrlvarTypes
Public ScalingFactor As Double
Public InitialValue As Double
Public InitialValueFlag As Integer
Public LimiterControl As Integer
Public MinValue As Double
Public MaxValue As Double
Public CntrlvarFormat As Integer
Public CvInput As Collection
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
Public Property Get CCC() As Integer
    CCC = pCCC
End Property

Public Property Let CCC(arg As Integer)
    pCCC = arg
End Property

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(arg As String)
    pID = arg
End Property
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set CvInput = New Collection
End Sub

Public Sub Create(CntrlvarType As CntrlvarTypes, Optional ScalingFactor As Double = 1#, Optional InitialValue As Double = 0#)
    On Error GoTo ErrorHandler
    
    pRowBegin = 0
    pRowEnd = 0
    With Me
        .ScalingFactor = ScalingFactor
        .InitialValue = InitialValue
    End With
        
    pCntrlvarType = CntrlvarType
    pID = "Id"
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".Create, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
'
'
' CLASS METHODS
Public Sub AddCvInput(ParamArray Args() As Variant)
' Action: Add values to CvInput
    Dim i As Integer
    For i = LBound(Args) To UBound(Args)
        CvInput.Add Args(i)
    Next i
End Sub


Public Sub WriteToFile(ByRef ts As Object, ByVal wC As Boolean)
' Action: Writes the force components
'
' Inputs:   ts = A TextStream object (included in the MS Scripting Runtime)
'           wC = writeComments, True if comments are to be written, false otherwise
'
    On Error GoTo ErrorHandler
    
    Dim s As New ResourceSprintf
    Dim i As Integer, j As Integer, k As Integer
    Dim t As New ResourceTablePrint
    
    Dim cvCCC As ResourceTablePrint, cvNN As ResourceTablePrint
    
    Dim cvString As String
    cvString = IIf(CntrlvarFormat = 999, "205%03d%02d", "205%04d%d")
    
    
    Dim CntrlvarStr As String
    Select Case pCntrlvarType
        Case CntrlvarSum
            ts.WriteLine s.sprintf(cvString & " %-8s  %-8s  %7.3g  %7.3g  %1d", CCC, 0, ID, "sum", ScalingFactor, InitialValue, InitialValueFlag)
            ts.WriteLine s.sprintf(cvString & " %8.6f % 7.5f  %-8s  %s", CCC, 1, CvInput(1), CvInput(2), CvInput(3), CvInput(4))
            For i = 5 To CvInput.Count Step 3   ' NOTE: Strange bug forces the need to convert to double
                ts.WriteLine s.sprintf(cvString & "          % 7.5f  %-8s  %s", CCC, (i + 1) / 3, CvInput(i), CvInput(i + 1), CvInput(i + 2))
            Next i
            
            
        Case Else
            
    End Select  ' CCCVV0000
    
    'ts.WriteLine s.sprintf(CntrlvarStr & "  %-8s  %-8s  %7.3g  %7.3g  %1d", CntrlvarNumber, 0, ID, pCntrlvarType, ScalingFactor, InitialValue, InitialValueFlag)
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".WriteToFile, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
