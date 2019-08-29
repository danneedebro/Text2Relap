VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HydroSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLASS HydroSystem
'
'
Option Explicit

' MEMBER VARIABLES
Private pInputDeck As Text2Relap             ' Reference to parent object that reads the input
Private pComponents As CollectionComponents  ' Contains all printable components
Private pSettings As PropertiesSettings      ' Contains all the model settings
'
'
' PROPERTIES
Public Property Get Inputdeck() As Text2Relap
    If IsMissing(pInputDeck) Then MsgBox "Inputdeck missing"
    Set Inputdeck = pInputDeck
End Property

Public Property Set Inputdeck(arg As Text2Relap)
    Set pInputDeck = arg
End Property

Public Property Get Settings() As PropertiesSettings
    Set Settings = pSettings
End Property

Public Property Get Components() As CollectionComponents
    Set Components = pComponents
End Property

Public Property Get Trip(ByVal TripId As String) As ComponentTrip
    Set Trip = pComponents("TRIP=" & TripId)
End Property

Public Property Get Force(ByVal ForceId As String) As ComponentForce
    ' If keyword given, try accessing it, if it doesn't exist, create it
    On Error GoTo ErrorHandler
    Dim firstError As Boolean
    firstError = True
    Set Force = pComponents(ForceId)
    Exit Property
ErrorHandler:
    If err.Number = 5 And firstError = True Then
        firstError = False
        Dim NewForce As New ComponentForce
        NewForce.Create Me
        pComponents.AddObject NewForce, ForceId
        'Force(Key).Create Me
        Resume
    End If
End Property
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set pSettings = New PropertiesSettings
    pSettings.Create Me
    
    Set pComponents = New CollectionComponents
    pComponents.Create Me
End Sub

Public Sub Create(ByRef CallingInputDeck As Text2Relap)
' Action: Only action is to make a connection to main class Text2Relap
'
    On Error GoTo ErrorHandler
    Set pInputDeck = CallingInputDeck
    
    Exit Sub
ErrorHandler:
    err.Raise err.Number, TypeName(Me) & ".Create, " & err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub

Public Sub DoPreRead()
' Action: Does actions po
'
    On Error GoTo ErrorHandler

    Components.Add Header, NewInputRow, 7, 15
    
    Dim i As Integer
    Dim NewComment As ComponentComment
    Dim NewTrip As ComponentTrip
    Dim TripInput() As Variant
    
    TripInput = Me.NewInputRow(Word1:="tripvar", Word4:="time-0", Word5:="ge", _
                               Word7:="null-0", Word8:=9999999, Word9:="n")
    
    For i = 1 To 4
        If i = 1 Then
            Components.Add Comment3, NewInputRow, 7, 15
            Set NewComment = Components(Components.Count)
            NewComment.CommentType = 3
            NewComment.Text = vbNewLine & "* Default trips for valves"
        ElseIf i = 3 Then
            Components.Add Comment3, NewInputRow, 7, 15
            Set NewComment = Components(Components.Count)
            NewComment.CommentType = 3
            NewComment.Text = vbNewLine & "* Default trips for pumps"
        End If
        
        NewInputRow TripInput, Word2:="DFLT" & CStr(i)
        If i = 4 Then NewInputRow TripInput, Word5:="lt"
        
        Components.Add TripVariable, TripInput, 7, 15
        
    Next i
    
    Exit Sub
ErrorHandler:
    err.Raise err.Number, TypeName(Me) & ".DoPreRead, " & err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub

Public Sub DoPostRead()
' Action:
    On Error GoTo ErrorHandler

    Dim i As Integer
    
    ' Set Trip number on default trips (created before tripformat was set)
    For i = 1 To 4
        Select Case Settings.TripFormat
            Case 0: Trip("DFLT" & CStr(i)).TripNumber = 400 + i
            Case 1: Trip("DFLT" & CStr(i)).TripNumber = i * 10
            Case 2: Trip("DFLT" & CStr(i)).TripNumber = i
        End Select
    Next i
    
    ' Set Custom initial conditions
    Settings.Inits.AssignValues
    
    ' Establish junction connections
    MakeConnections
        
    Exit Sub
ErrorHandler:
    err.Raise err.Number, TypeName(Me) & ".DoPostRead, " & err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub
'
'
' CLASS METHODS
Public Function ComponentExists(ByVal CCC As Integer, Optional Vol As Integer = 0) As Boolean
    Dim FoundItem As Boolean
    Dim i As Integer, j As Integer
    Dim CurrHydroComp As ComponentHydro
    
    For i = 1 To Me.Components.Count
        If Components(i).ObjectType <> HydroComp Then GoTo Next_i
        Set CurrHydroComp = Components(i)
        With CurrHydroComp
        
            If .CCC = CCC Then
                If Vol <> 0 Then  ' If also search for volume
                    FoundItem = False
                    For j = 1 To .Segments.Count
                        If Vol >= .Segments(j).VolumeFirst And Vol <= .Segments(j).VolumeLast Then
                            FoundItem = True
                        End If
                    Next j
                Else
                    FoundItem = True
                End If
                
            End If
        End With
Next_i:
    Next i
    
    ComponentExists = FoundItem
    
End Function

Public Function ReplaceTrips(ByVal InputString As String) As String
' Action: If InputString contains the Substring "[TRIP:" it parses out the ID/Name of the trip ("[TRIP:tripID]") and
'         replaces this with the trip number
'
    Dim pos1 As Integer, pos2 As Integer
    Dim TripId As String
    Dim i As Integer
    Const LookFor = "[TRIP:"
    Dim CurrTrip As ComponentTrip
    
    pos2 = 1 ' Sets the starting position
    
    Do
        pos1 = InStr(1, InputString, LookFor)
        If pos1 = 0 Then Exit Do
        
        pos2 = InStr(pos1, InputString, "]")
        If pos2 = 0 Then Exit Do
        
        ' Parse out trip ID
        TripId = Mid(InputString, pos1 + Len(LookFor), pos2 - pos1 - Len(LookFor))
        
        ' Fetch tripnumber for trip with id "TripId"
        On Error GoTo TripIdNotFound
        Set CurrTrip = Trip(TripId)
        InputString = Left(InputString, pos1 - 1) + Format(CurrTrip.TripNumber, "@@@@@") & Right(InputString, Len(InputString) - pos2)
        
    Loop
    
    ReplaceTrips = InputString
    Exit Function
TripIdNotFound:
    ReplaceTrips = InputString
End Function


Public Function FindSegment(ByVal CurrentConnection As PropertiesHydroCompConnection, ByVal JunctionIndex As Integer) As PropertiesHydroCompConnection
' Action: Returns a junction connection where "CCC" and "Volume" are identified from the ID. Loops through all hydro components
'         and looks for the ID
'
' Inputs:
'   CurrentConnection = The "connection"-object that are to be connected (Fields CCC and Volume blank and only a ID is present)
'   JunctionIndex = 1 or 2 depending if its a "from-connection" (1) or a "to-connection" (2)

    On Error GoTo ErrorHandler

    Dim FindFirst As Boolean
    
    ' Depending on if its a from- or to-junction look for the first or last occurance of an ID.
    ' Usually there are only supposed to be one unique ID but if a pipe segment doesn't fit in a pipe component
    ' there can several identical unique IDs
    With CurrentConnection
        If JunctionIndex = 1 Then
            If .Side = 1 Then FindFirst = True
            If .Side = 2 Or .Side = 3 Or .Side = 4 Or .Side = 5 Or .Side = 6 Then FindFirst = False
        ElseIf JunctionIndex = 2 Then
            If .Side = 1 Or .Side = 3 Or .Side = 4 Or .Side = 5 Or .Side = 6 Then FindFirst = True
            If .Side = 2 Then FindFirst = False
        End If
    End With
    
    Dim i As Integer, j As Integer
    Dim CCC As Integer
    Dim SegmentIndex As Integer
    Dim CurrHydroComp As ComponentHydro
    
    For i = 1 To Components.Count
        If Components(i).ObjectType <> HydroComp Then GoTo Next_i
        Set CurrHydroComp = Me.Components(i)
        For j = 1 To CurrHydroComp.Segments.Count
            If CurrHydroComp.Segments(j).ID = CurrentConnection.ID Then
                CurrentConnection.CCC = CurrHydroComp.CCC
                If FindFirst = True Then
                    CurrentConnection.Volume = CurrHydroComp.Segments(j).VolumeFirst
                    GoTo ExitLoops
                Else
                    CurrentConnection.Volume = CurrHydroComp.Segments(j).VolumeLast
                End If
            End If
        Next j
Next_i:
    Next i
ExitLoops:
    Set FindSegment = CurrentConnection
    
    Exit Function
ErrorHandler:
    err.Raise err.Number, TypeName(Me) & ".FindSegment, " & err.Source, err.Description, err.HelpFile, err.HelpContext
End Function

Public Sub MakeConnections()
' Action: Connects junctions so that instead of refering to a ID they refer to the component number
'         of the component
'
'
    On Error GoTo ErrorHandler

    Dim i As Integer, j As Integer, k As Integer
    Dim lookForID1 As String
    Dim foundConnection As Boolean
    Dim CurrHydroComp As ComponentHydro
    
    ' Loop through all components
    For i = 1 To Me.Components.Count
        If Components(i).ObjectType <> HydroComp Then GoTo nextIteration_i
        Set CurrHydroComp = Me.Components(i)
        If CurrHydroComp.Info.Family <> JunctionComponent Then GoTo nextIteration_i
        ' Loop through both connections for the junction component
        For j = 1 To 2
            Dim FindFirst As Boolean
            
            With CurrHydroComp.Connections(j)
                If .CCC <> 0 Then GoTo NextIteration_j  ' If junction already connected
                foundConnection = False
                
                Dim NewConnection As PropertiesHydroCompConnection
                Set NewConnection = Me.FindSegment(CurrHydroComp.Connections(j), j)
                
                If NewConnection.CCC = 0 Then
                    MsgBox "Error: Did not find connection for junction " & CurrHydroComp.Info.MainTypeStr & "-" & CurrHydroComp.CCC
                Else
                    .CCC = NewConnection.CCC
                    .Volume = NewConnection.Volume
                End If
             End With
NextIteration_j:
        Next j
nextIteration_i:
    Next i
    
    Exit Sub
ErrorHandler:
    err.Raise err.Number, TypeName(Me) & ".MakeConnections(Comp=" & CStr(i) & "), " & err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub


'
'
' STATIC METHODS
Public Function NewInputRow(Optional ByRef InputRowToChange As Variant, Optional ByVal RowToChange As Integer = 1, Optional ByVal Word1 As Variant, _
                            Optional ByVal Word2 As Variant, Optional ByVal Word3 As Variant, _
                            Optional ByVal Word4 As Variant, Optional ByVal Word5 As Variant, Optional ByVal Word6 As Variant, _
                            Optional ByVal Word7 As Variant, Optional ByVal Word8 As Variant, Optional ByVal Word9 As Variant, _
                            Optional ByVal Word10 As Variant, Optional ByVal Word11 As Variant, Optional ByVal Word12 As Variant, _
                            Optional ByVal Word13 As Variant, Optional ByVal Word14 As Variant, Optional ByVal Word15 As Variant, _
                            Optional ByVal Word16 As Variant, Optional ByVal Word17 As Variant, Optional ByVal Word18 As Variant, _
                            Optional ByVal Word19 As Variant, Optional ByVal Word20 As Variant, Optional ByVal Word21 As Variant, _
                            Optional ByVal Word22 As Variant) As Variant()
    
    Dim Result(1 To 1, 1 To 22) As Variant
    Dim i As Integer
    
    If Not IsMissing(Word1) Then Result(1, 1) = Word1
    If Not IsMissing(Word2) Then Result(1, 2) = Word2
    If Not IsMissing(Word3) Then Result(1, 3) = Word3
    If Not IsMissing(Word4) Then Result(1, 4) = Word4
    If Not IsMissing(Word5) Then Result(1, 5) = Word5
    If Not IsMissing(Word6) Then Result(1, 6) = Word6
    If Not IsMissing(Word7) Then Result(1, 7) = Word7
    If Not IsMissing(Word8) Then Result(1, 8) = Word8
    If Not IsMissing(Word9) Then Result(1, 9) = Word9
    If Not IsMissing(Word10) Then Result(1, 10) = Word10
    
    If Not IsMissing(Word11) Then Result(1, 11) = Word11
    If Not IsMissing(Word12) Then Result(1, 12) = Word12
    If Not IsMissing(Word13) Then Result(1, 13) = Word13
    If Not IsMissing(Word14) Then Result(1, 14) = Word14
    If Not IsMissing(Word15) Then Result(1, 15) = Word15
    If Not IsMissing(Word16) Then Result(1, 16) = Word16
    If Not IsMissing(Word17) Then Result(1, 17) = Word17
    If Not IsMissing(Word18) Then Result(1, 18) = Word18
    If Not IsMissing(Word19) Then Result(1, 19) = Word19
    If Not IsMissing(Word20) Then Result(1, 20) = Word20
    
    If Not IsMissing(Word21) Then Result(1, 21) = Word21
    If Not IsMissing(Word22) Then Result(1, 22) = Word22
    
    If IsMissing(InputRowToChange) = True Then  '
        NewInputRow = Result
    Else
        If VarType(InputRowToChange) = vbArray + vbVariant Then
            If UBound(InputRowToChange, 2) <> 22 Or LBound(InputRowToChange, 2) <> 1 Then
                MsgBox "Should have size (1 to 22)"
                Exit Function
            Else
                For i = LBound(InputRowToChange, 2) To UBound(InputRowToChange, 2)
                    If IsEmpty(Result(1, i)) = False Then InputRowToChange(RowToChange, i) = Result(1, i)
                Next i
            End If
        Else
            MsgBox "Should be a variant array"
            Exit Function
        End If
    End If
End Function



