VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComponentTrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ComponentTrip.cls
'
'
Option Explicit

Implements iComponent

'
'
' MEMBER VARIABLES
Private pHydroSystem As HydroSystem

Private pTripType As Integer   ' 1 = Variable trip,  2 = Logical trip
Private pTripNumber As Integer
Private pID As String
Private Operator As String
Private AdditiveConstant As Double
Private Latch As String
Private TripIDs(1 To 2) As String
Private CompIDs(1 To 2) As String    ' mflowj, p, etc  OR Hydrocomp Unique-ID (PIPE_12)
Private PlotvarnumCodes(1 To 2) As String

Private pRowBegin As Integer
Private pRowEnd As Integer
'
'
' SHARED PROPERTIES AND METHODS
Private Property Get iComponent_ObjectType() As R5Component
    If pTripType = 1 Then
        iComponent_ObjectType = TripVariable
    Else
        iComponent_ObjectType = TripLogical
    End If
End Property

Private Function iComponent_ComponentInfo() As String
    If pTripType = 1 Then
        iComponent_ComponentInfo = "Variable trip no: " & Me.TripNumber
    Else
        iComponent_ComponentInfo = "Logical trip no: " & Me.TripNumber
    End If
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

Public Property Get TripNumber() As Integer
    TripNumber = pTripNumber
End Property

Public Property Let TripNumber(arg As Integer)
    pTripNumber = arg
End Property

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(arg As String)
    pID = arg
End Property

Public Property Get TripType() As Integer
    TripType = pTripType
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
    
    Select Case LCase(inputRows(1, 1))
        Case "tripvar"
            Me.TripNumber = Settings.CurrentTripNumVariable
            Settings.IncrementTripVariable
            pTripType = 1
            CompIDs(1) = inputRows(1, 3)
            PlotvarnumCodes(1) = inputRows(1, 4)
            CompIDs(2) = inputRows(1, 6)
            PlotvarnumCodes(2) = inputRows(1, 7)
            Operator = inputRows(1, 5)
            AdditiveConstant = inputRows(1, 8)
            Latch = inputRows(1, 9)
        Case "triplog"
            Me.TripNumber = Settings.CurrentTripNumLogical
            Settings.IncrementTripLogical
            pTripType = 2
            TripIDs(1) = inputRows(1, 3)
            Operator = inputRows(1, 4)
            TripIDs(2) = inputRows(1, 5)
            Latch = inputRows(1, 6)
    End Select
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".Create, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
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
    On Error GoTo ErrorHandler
    Dim s As New ResourceSprintf
    Dim i As Integer, j As Integer, k As Integer
    Dim CurrHydroComp As ComponentHydro
    
    Dim TripStr As String
    Select Case Settings.TripFormat
        Case 0    ' 401-599 and 601-799
            TripStr = s.sprintf("%-8d", TripNumber)
        Case 1, 2 '20600010
            TripStr = s.sprintf("206%05d", TripNumber)
    End Select  ' CCCVV0000
    
    
    Select Case TripType
        Case 1 ' Variable trip
            Dim Plotvars(1 To 2) As String
            Dim Plotnums(1 To 2) As String
            
            Dim CurrCCC As Integer
            
            For i = 1 To 2
                CurrCCC = -1
                
                Dim FindID As Variant
                FindID = HydroSystem.Components.FindID(CompIDs(i))
                
                If FindID(1) = True Then
                    CurrCCC = FindID(4)
                ElseIf CompIDs(i) = "" Then
                    ' Do nothing
                Else
                    MsgBox "Error component with ID=" & CompIDs(i) & " not found", vbExclamation
                End If
                
                                
                ' Parse PlotvarnumCodes(i)
                Plotvars(i) = Split(PlotvarnumCodes(i), "-")(0)
                Plotnums(i) = Split(PlotvarnumCodes(i), "-")(1)
                
                Plotnums(i) = Replace(Plotnums(i), "CCC", CStr(CurrCCC))
                
            Next i
            ts.WriteLine s.sprintf("%s  %8s  %9s  %2s  %8s  %9s  %#7.3g  %s", TripStr, Plotvars(1), Plotnums(1), Operator, Plotvars(2), Plotnums(2), AdditiveConstant, Latch)
            ts.WriteLine ""
        
        Case 2 ' Logical trip
            Dim Trips(1 To 2) As Integer
                    
            ' Loop through all trips
            For i = 1 To HydroSystem.Components.Count
                Dim CurrTrip As ComponentTrip
                If HydroSystem.Components(i).ObjectType = TripLogical Or HydroSystem.Components(i).ObjectType = TripVariable Then
                    Set CurrTrip = HydroSystem.Components(i)
                
                    With CurrTrip
                        If .ID = TripIDs(1) Then
                            Trips(1) = .TripNumber
                        ElseIf .ID = TripIDs(2) Then
                            Trips(2) = .TripNumber
                        End If
                    End With
                End If
            Next i
            ts.WriteLine s.sprintf("%s  %5d  %2s  %5d  %s", TripStr, Trips(1), Operator, Trips(2), Latch)
            ts.WriteLine ""
    End Select
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".WriteToFile, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
