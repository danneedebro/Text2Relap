VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionComponents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Private pComponents As Collection
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
    Set pComponents = New Collection
End Sub

Public Sub Create(ByRef Parent As HydroSystem)
    Set pHydroSystem = Parent
End Sub

Public Sub AddObject(ByVal R5Comp As Object, Optional ByVal Key As String)
    If Key = "" Then
        pComponents.Add R5Comp
    Else
        pComponents.Add R5Comp, Key
    End If
End Sub

Public Sub Add(ByVal Item As R5Component, ByRef inputRows() As Variant, ByVal SourceRow1 As Integer, ByVal SourceRow2 As Integer)
    On Error GoTo ErrorHandler
    
    Dim Key As String
    
    If Item = HydroComp Then
        Dim NewHydroComp As New ComponentHydro
        
        Key = "CCC" & Format(HydroSystem.Settings.CurrentCCC, "000")
        pComponents.Add NewHydroComp, Key
        NewHydroComp.Create inputRows, HydroSystem, SourceRow1, SourceRow2
        HydroSystem.Settings.IncrementCCC
        
    ElseIf Item = Custom Then
        Dim NewCustomComp As New ComponentCustom
        pComponents.Add NewCustomComp
        NewCustomComp.Create inputRows, HydroSystem, SourceRow1
        
    ElseIf Item = TripLogical Or Item = TripVariable Then
        Dim NewTripComp As New ComponentTrip
        NewTripComp.Create inputRows, HydroSystem, SourceRow1, SourceRow2
        Key = "TRIP=" & NewTripComp.ID
        pComponents.Add NewTripComp, Key
        
    ElseIf Item = Comment1 Or Item = Comment2 Or Item = Comment3 Then
        Dim NewComment As New ComponentComment
        NewComment.Create inputRows, SourceRow1, SourceRow2
        pComponents.Add NewComment
        
    ElseIf Item = Header Then
        Dim NewHeader As New ComponentHeader
        pComponents.Add NewHeader, "Header"
        NewHeader.Create HydroSystem, SourceRow1, SourceRow2
    
    
    Else
        MsgBox "error"
    
    End If
    
    Exit Sub
ErrorHandler:
    If err.Number = 457 Then
        HydroSystem.Inputdeck.HighlightCurrentLines
        Dim answ
        
        answ = MsgBox("Critical error: " & Key & " already taken", vbCritical, "CCC already occupied")
        HydroSystem.Inputdeck.Warnings = True
        HydroSystem.Inputdeck.Abort = True
        Exit Sub
    Else
        err.Raise err.Number, TypeName(Me) & ".Add, " & err.Source, err.Description, err.HelpFile, err.HelpContext
    End If
End Sub

Public Function Item(index As Variant) As iComponent
Attribute Item.VB_UserMemId = 0
    Set Item = pComponents(index)
End Function

Public Function Count() As Long
    Count = pComponents.Count
End Function

Public Function FindID(ByVal ID As String, Optional ByVal FindLast As Boolean = False) As Variant
' Action: Finds the componentindex
'
' Inputs:
'   ID = Unique ID of a hydro comp or pipe segment
'
' Outputs:
'   FindID(1) = Boolean, true if found
'   FindID(2) = Component index
'   FindID(3) = Segment index
'   FindID(4) = Component number (CCC)

    On Error GoTo ErrorHandler
    Dim i As Integer, j As Integer
    Dim CurrHydroComp As ComponentHydro
    Dim Result(1 To 4) As Variant
    Result(1) = False
    
    For i = 1 To Me.Count
        If Me(i).ObjectType <> HydroComp Then GoTo Next_i
        Set CurrHydroComp = Me(i)
        For j = 1 To CurrHydroComp.Segments.Count
            If CurrHydroComp.Segments(j).ID = ID Then
                Result(1) = True
                Result(2) = i
                Result(3) = j
                Result(4) = CurrHydroComp.CCC
                If FindLast = False Then Exit For
            End If
        Next j
Next_i:
    Next i
    
    FindID = Result
    
    Exit Function
ErrorHandler:
    err.Raise err.Number, TypeName(Me) & ".FindID, " & err.Source, err.Description, err.HelpFile, err.HelpContext
End Function
'
'
' CLASS METHODS

'
'
' STATIC METHODS

