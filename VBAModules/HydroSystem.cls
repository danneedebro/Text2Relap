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
Public Property Get InputDeck() As Text2Relap
    If IsMissing(pInputDeck) Then MsgBox "Inputdeck missing"
    Set InputDeck = pInputDeck
End Property

Public Property Set InputDeck(arg As Text2Relap)
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
    If Err.Number = 5 And firstError = True Then
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
    Err.Raise Err.Number, TypeName(Me) & ".Create, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub DoPreRead()
' Action: Does actions po
'
    On Error GoTo ErrorHandler

    Components.Add Header, NewInputRow, 7, 15
    
    Dim i As Integer
    Dim newComment As ComponentComment
    Dim NewTrip As ComponentTrip
    Dim TripInput() As Variant
    
    TripInput = Me.NewInputRow(Word1:="tripvar", Word4:="time-0", Word5:="ge", _
                               Word7:="null-0", Word8:=9999999, Word9:="n")
    
    For i = 1 To 4
        If i = 1 Then
            Components.Add Comment3, NewInputRow, 7, 15
            Set newComment = Components(Components.Count)
            newComment.CommentType = 3
            newComment.Text = vbNewLine & "* Default trips for valves"
        ElseIf i = 3 Then
            Components.Add Comment3, NewInputRow, 7, 15
            Set newComment = Components(Components.Count)
            newComment.CommentType = 3
            newComment.Text = vbNewLine & "* Default trips for pumps"
        End If
        
        NewInputRow TripInput, Word2:="DFLT" & CStr(i)
        If i = 4 Then NewInputRow TripInput, Word5:="lt"
        
        Components.Add TripVariable, TripInput, 7, 15
        
    Next i
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".DoPreRead, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
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
    
    ' Construct force
    MakeForces
        
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".DoPostRead, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
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
    
    
    ' Loop through each hydro component and their segments looking for a segment with specified id
    Dim component As ComponentHydro
    Dim segment As PropertiesHydroCompSegment
    
    For Each component In Components.Subset(HydroComp)
    
        For Each segment In component.Segments
            If segment.ID = CurrentConnection.ID Then
                CurrentConnection.CCC = component.CCC
                If FindFirst = True Then
                    CurrentConnection.Volume = segment.VolumeFirst
                    GoTo ExitLoops
                Else
                    CurrentConnection.Volume = segment.VolumeLast
                End If
            End If
        Next segment

    Next component
    
ExitLoops:
    Set FindSegment = CurrentConnection
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".FindSegment, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
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
    Dim component As ComponentHydro
    Dim connection As PropertiesHydroCompConnection
    For Each component In Components.Subset(HydroComp)
    
        If component.Info.Family <> JunctionComponent Then GoTo NextComponent
        
        ' Loop through both connections for the junction component
        j = 0
        For Each connection In component.Connections
            j = j + 1
            Dim FindFirst As Boolean
            
            If connection.CCC <> 0 Then GoTo NextConnection  ' If junction already connected
            foundConnection = False
            
            Dim NewConnection As PropertiesHydroCompConnection
            Set NewConnection = Me.FindSegment(connection, j)
            
            If NewConnection.CCC = 0 Then
                MsgBox "Error: Did not find connection for junction " & component.Info.MainTypeStr & "-" & component.CCC
            Else
                connection.CCC = NewConnection.CCC
                connection.Volume = NewConnection.Volume
            End If
            
NextConnection:
        Next connection
NextComponent:
    Next component
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".MakeConnections(Comp=" & "), " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub MakeForces()
' Action: Construct forces by looping through hydrocomponents and create the cntrlvars
'         to calculate the unbalanced pipe forces
'
    On Error GoTo ErrorHandler

    Dim i As Integer, j As Integer, k As Integer
    Dim s As New ResourceSprintf
    
    Dim forceMatrix(1 To 9999) As New Collection
    Dim forceInfo As New Collection
    
    ' Loop through components looking for force numbers in pipes
    ' TODO: If last segment in pipe, look for components that connects to this last volume
    Dim component As ComponentHydro
    Dim segment As PropertiesHydroCompSegment
    For Each component In Components.Subset(HydroComp)
    
        For Each segment In component.Segments
            If segment.ForceNumber > 0 And segment.ForceNumber <= 9999 Then
                forceMatrix(segment.ForceNumber).Add component.CCC
                forceMatrix(segment.ForceNumber).Add segment.JunctionFirst
                forceMatrix(segment.ForceNumber).Add segment.JunctionLast
                forceMatrix(segment.ForceNumber).Add segment.Dx
            End If
        Next segment
        
    Next component
    
    '20522200 ID-0005     sum      1.0000    0.0000     0
    '20522201 0.000000 1.00000  cntrlvar  224
    '20522300 ID-0005     sum      1.0000    0.0000     0
    '20522301 0.000000 0.20000  mflowj    104260000
    '20522302          0.20000  mflowj    104270000
    '20522303          0.20000  mflowj    104280000
    '20522304          0.20000  mflowj    104290000
    '20522305          0.20000  mflowj    104300000
    '20522306          0.20000  mflowj    104310000
    '20522307          0.20000  mflowj    104320000
    '20522308          0.20000  mflowj    104330000
    '20522309          0.20000  mflowj    104340000
    '20522310          0.20000  mflowj    104350000
    '20522311          0.20000  mflowj    104360000
    '20522312          0.20000  mflowj    104370000
    '20522313          0.20000  mflowj    104380000
    '20522314          0.20000  mflowj    104390000
    '20522315          0.20000  mflowj    104400000
    '20522316          0.20000  mflowj    104410000
    '20522317          0.20000  mflowj    104420000
    '20522318          0.20000  mflowj    104430000
    '20522319          0.20000  mflowj    104440000
    '20522320          0.20000  mflowj    104450000
    '20522400 ID-0005     sum      1.0000    0.0000     0
    '20522401 0.000000 1.00000  cntrlvar  223
    '20522402          0.20000  mflowj    104460000
    '20522403          0.20000  mflowj    104470000
    '20522404          0.20000  mflowj    104480000
    '20522405          0.20000  mflowj    104490000
    '20522406          0.20000  mflowj    104500000
    '20522500 ID-0005 sum  1.000 0.0 0
    '20522501 0.000000 1.00000  cntrlvar  224
    '20522502         -1.00000  cntrlvar  222
    '20500500 ID-0005 div -1.000 0.0 0
    '20500501 dt  0  cntrlvar  225
    
    
    ' Loop through collected information
    Dim Cv1 As ComponentCntrlvar, CvN As ComponentCntrlvar
    Dim CommentForce As ComponentComment
    For i = 1 To UBound(forceMatrix)
        With forceMatrix(i)
            If .Count = 0 Then GoTo NextPotentialForce
            
            Set CommentForce = New ComponentComment
            Components.AddObject CommentForce
            CommentForce.CommentType = 3
            CommentForce.Text = vbNewLine & "* Force " & i
            
            
            ' First cntrlvar to store old value of mass flow summation
            Set Cv1 = New ComponentCntrlvar
            Components.AddObject Cv1
            Cv1.Create CntrlvarSum, CCC:=Settings.CurrentCntrlvarNum, CntrlvarFormat:=Settings.CntrlvarFormat
            Settings.IncrementCntrlvar
            
            
            Dim cardNum As Integer, cardNumMax As Integer, cvCCCPrev As Integer
            Set CvN = New ComponentCntrlvar
            Components.AddObject CvN
            CvN.Create CntrlvarSum, CCC:=Settings.CurrentCntrlvarNum, CntrlvarFormat:=Settings.CntrlvarFormat
            Settings.IncrementCntrlvar
            cvCCCPrev = CvN.CCC
            
            cardNum = 0
            cardNumMax = IIf(Settings.CntrlvarFormat = 999, 21, 9) ' Max number of
            
            ' Loop from 'junctionFirst' to 'junctionLast'
            For j = 1 To forceMatrix(i).Count Step 4
                For k = .Item(j + 1) To .Item(j + 2)
                    cardNum = cardNum + 1
                    If cardNum = 1 Then
                        CvN.AddCvInput 0#, .Item(j + 3), "mflowj", s.sprintf("%03d%02d0000", .Item(j), k)
                    ElseIf cardNum = cardNumMax Then
                        Set CvN = New ComponentCntrlvar
                        Components.AddObject CvN
                        CvN.Create CntrlvarSum, CCC:=Settings.CurrentCntrlvarNum, CntrlvarFormat:=Settings.CntrlvarFormat
                        Settings.IncrementCntrlvar
                        CvN.AddCvInput 0#, 1#, "cntrlvar", CStr(cvCCCPrev)
                        CvN.AddCvInput .Item(j + 3), "mflowj", s.sprintf("%03d%02d0000", .Item(j), k)
                        cvCCCPrev = CvN.CCC
                        cardNum = 1
                    Else
                        CvN.AddCvInput .Item(j + 3), "mflowj", s.sprintf("%03d%02d0000", .Item(j), k)
                    End If
                    
                    'Debug.Print s.sprintf("%7.5f  mflowj %03d%02d0000", .Item(j + 3), .Item(j), k)
                Next k
            Next j
            
            ' Add input to the first cntrlvar that stores sum(mflowj)_old
            Cv1.AddCvInput 0#, 1, "cntrlvar", CStr(cvCCCPrev)
            
            ' Add the cntrlvar that sums: sum(mflowj) - sum(mflowj)_old
            Set CvN = New ComponentCntrlvar
            Components.AddObject CvN
            CvN.Create CntrlvarSum, CCC:=Settings.CurrentCntrlvarNum, CntrlvarFormat:=Settings.CntrlvarFormat
            Settings.IncrementCntrlvar
            CvN.AddCvInput 0#, 1#, "cntrlvar", CStr(cvCCCPrev)
            CvN.AddCvInput -1#, "cntrlvar", CStr(Cv1.CCC)
            cvCCCPrev = CvN.CCC
            
            ' Add the div component
            Set CvN = New ComponentCntrlvar
            Components.AddObject CvN
            CvN.Create CntrlvarDiv, CCC:=i, CntrlvarFormat:=Settings.CntrlvarFormat, ScalingFactor:=-1#
            Settings.IncrementCntrlvar
            
            CvN.AddCvInput "dt", "0", "cntrlvar", CStr(cvCCCPrev)
            
            'Debug.Print "   CCC=" & .Item(j) & ", j1="
            
        End With
NextPotentialForce:
    Next i
    
    Exit Sub
ErrorHandler:
    Stop
    Err.Raise Err.Number, TypeName(Me) & ".MakeForces(), " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
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
    
    Dim result(1 To 1, 1 To 22) As Variant
    Dim i As Integer
    
    If Not IsMissing(Word1) Then result(1, 1) = Word1
    If Not IsMissing(Word2) Then result(1, 2) = Word2
    If Not IsMissing(Word3) Then result(1, 3) = Word3
    If Not IsMissing(Word4) Then result(1, 4) = Word4
    If Not IsMissing(Word5) Then result(1, 5) = Word5
    If Not IsMissing(Word6) Then result(1, 6) = Word6
    If Not IsMissing(Word7) Then result(1, 7) = Word7
    If Not IsMissing(Word8) Then result(1, 8) = Word8
    If Not IsMissing(Word9) Then result(1, 9) = Word9
    If Not IsMissing(Word10) Then result(1, 10) = Word10
    
    If Not IsMissing(Word11) Then result(1, 11) = Word11
    If Not IsMissing(Word12) Then result(1, 12) = Word12
    If Not IsMissing(Word13) Then result(1, 13) = Word13
    If Not IsMissing(Word14) Then result(1, 14) = Word14
    If Not IsMissing(Word15) Then result(1, 15) = Word15
    If Not IsMissing(Word16) Then result(1, 16) = Word16
    If Not IsMissing(Word17) Then result(1, 17) = Word17
    If Not IsMissing(Word18) Then result(1, 18) = Word18
    If Not IsMissing(Word19) Then result(1, 19) = Word19
    If Not IsMissing(Word20) Then result(1, 20) = Word20
    
    If Not IsMissing(Word21) Then result(1, 21) = Word21
    If Not IsMissing(Word22) Then result(1, 22) = Word22
    
    If IsMissing(InputRowToChange) = True Then  '
        NewInputRow = result
    Else
        If VarType(InputRowToChange) = vbArray + vbVariant Then
            If UBound(InputRowToChange, 2) <> 22 Or LBound(InputRowToChange, 2) <> 1 Then
                MsgBox "Should have size (1 to 22)"
                Exit Function
            Else
                For i = LBound(InputRowToChange, 2) To UBound(InputRowToChange, 2)
                    If IsEmpty(result(1, i)) = False Then InputRowToChange(RowToChange, i) = result(1, i)
                Next i
            End If
        Else
            MsgBox "Should be a variant array"
            Exit Function
        End If
    End If
End Function



