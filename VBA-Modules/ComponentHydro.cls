VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComponentHydro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ComponentHydro.cls
'
'
Option Explicit

Implements iComponent
'
'
' MEMBER VARIABLES
Private pSegments As CollectionSegments
Private pConnections As CollectionConnections
Public Init As PropertiesHydroCompInit
Public Info As PropertiesHydroCompInfo
Private ComponentNumber As Integer
Public Filename As String
Private pName As String

Private pHydroSystem As HydroSystem
Private pJunction As ComponentHydroJunction
Private pPipe As ComponentHydroPipe
Private pSnglvol As ComponentHydroSingleVol

Private pRowBegin As Integer
Private pRowEnd As Integer

'
'
' SHARED PROPERTIES AND METHODS
Private Property Get iComponent_ObjectType() As R5Component
    iComponent_ObjectType = HydroComp
End Property

Private Function iComponent_ComponentInfo() As String
    Dim OutputStr As String
    Dim i As Integer
    OutputStr = "HydroComp, CCC=" & Me.CCC & ", type=" & Me.Info.MainTypeStr
    If Me.Info.MainType = pipe Then
        Dim length As Double
        For i = 1 To Me.Segments.Count: length = length + Me.Segments(i).Volumes * Me.Segments(i).Dx: Next i
        OutputStr = OutputStr & ", vols=" & Me.Segments(Me.Segments.Count).VolumeLast & ", length=" & Format(length, "0.00") & " m"
    ElseIf Me.Info.MainType = valve Then
        OutputStr = OutputStr & ", subtype=" & Me.Info.SubTypeStr
    ElseIf Me.Info.MainType = pump Then
        OutputStr = OutputStr & ", subtype=" & Me.Info.SubTypeStr
    End If
    
    If Me.Info.Family = JunctionComponent Then
        OutputStr = OutputStr & ", from=" & Me.Connections(1).ConnectionCode & ", to=" & Me.Connections(2).ConnectionCode
    End If
    
    iComponent_ComponentInfo = OutputStr
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
Public Property Get Inputdeck() As Text2Relap
    Set Inputdeck = pHydroSystem.Inputdeck
End Property

Public Property Get HydroSystem() As HydroSystem
    Set HydroSystem = pHydroSystem
End Property

Public Property Set HydroSystem(ByRef arg As HydroSystem)
    Set pHydroSystem = arg
End Property

Public Property Get Settings() As PropertiesSettings
    Set Settings = pHydroSystem.Settings
End Property

Public Property Get Child() As Object
    If Me.Info.Family = JunctionComponent Then
        Set Child = pJunction
    ElseIf Me.Info.Family = PipeComponent Then
        Set Child = pPipe
    ElseIf Me.Info.Family = SingleVolumeComponent Then
        Set Child = pSnglvol
    End If
End Property

Public Property Get Connections() As CollectionConnections
    Set Connections = pConnections
End Property

Public Property Get Segments() As CollectionSegments
    Set Segments = pSegments
End Property

Public Property Get CCC() As Integer
    CCC = ComponentNumber
End Property

Public Property Let CCC(ByVal arg As Integer)
    ComponentNumber = arg
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(arg As String)
    Dim Pos As Integer
    Pos = InStr(1, arg, ".inc")
    If Pos > 0 Then
        Filename = arg
        pName = Left(arg, Pos - 1)
    Else
        Filename = ""
        pName = Left(arg, 8)  ' Max 8 chars
    End If
    
End Property

Public Property Get RowBegin() As Integer
    RowBegin = pRowBegin
End Property

Public Property Get RowEnd() As Integer
    RowEnd = pRowEnd
End Property

'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set Info = New PropertiesHydroCompInfo
    Set pSegments = New CollectionSegments
    Set pConnections = New CollectionConnections
End Sub

Public Sub Create(inputRows() As Variant, ByRef ParentSystem As HydroSystem, ByVal RowBegin As Integer, ByVal RowEnd As Integer)
' Action: Creates a new hydrodynamic component from input
'
    On Error GoTo ErrorHandler
    Dim Word1 As Variant
    Word1 = inputRows(1, 1)
    
    pRowBegin = RowBegin
    pRowEnd = RowEnd
    
    Me.Name = inputRows(1, 10)
    Info.Name = inputRows(1, 10)
    Info.SetType Word1
    
    Set Me.HydroSystem = ParentSystem
    Set Init = Settings.Inits.CurrentInit
    CCC = HydroSystem.Settings.CurrentCCC
    
    If Info.Family = JunctionComponent Then
        Set pJunction = New ComponentHydroJunction
        pJunction.Create inputRows, Me
    ElseIf Info.Family = PipeComponent Then
        Set pPipe = New ComponentHydroPipe
        pPipe.Create inputRows, Me
    ElseIf Info.Family = SingleVolumeComponent Then
        Set pSnglvol = New ComponentHydroSingleVol
        pSnglvol.Create inputRows, Me
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".Create, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
'
'
' CLASS METHODS
Public Sub WriteToFile(ByRef ts As Object, ByVal wC As Boolean)
    On Error GoTo ErrorHandler
    If Info.Family = JunctionComponent Then
        pJunction.WriteToFile ts, wC
    ElseIf Info.Family = PipeComponent Then
        pPipe.WriteToFile ts, wC
    ElseIf Info.Family = SingleVolumeComponent Then
        pSnglvol.WriteToFile ts, wC
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".WriteToFile, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function ReplaceWords(ByVal InputString As String) As String
' Action: Replaces certain words with corresponding values
'
'
    Dim s As New ResourceSprintf
    Dim t As New ResourceTablePrint
    t.SetDefaultValues
    Dim Replacements(1 To 21) As Variant
    Dim i As Integer
    
    ' Replace trips
    If InStr(1, InputString, "[TRIP:") <> 0 Then InputString = HydroSystem.ReplaceTrips(InputString)
        
    For i = -10 To 10
        InputString = Replace(InputString, s.sprintf("[CCC%+d]", i), s.sprintf("%03d", Me.CCC + i))
    Next i
    
    With Me.Segments(1)
        Replacements(1) = Array("[CCC]", "CCC", s.sprintf("%03d", Me.CCC))
        Replacements(2) = Array("[AREA]", "AREA", t.GetFormattedValue(.Area, Area))
        Replacements(3) = Array("[KPOS]", "KPOS", t.GetFormattedValue(.KLossPos, KLossPos))
        Replacements(4) = Array("[KNEG]", "KNEG", t.GetFormattedValue(.KLossNeg, KLossNeg))
        Replacements(5) = Array("[DRAWING]", "DRAWING", s.sprintf("%s", .Drawing))
        Replacements(6) = Array("[TLPVBFE]", "TLPVBFE", t.GetFormattedValue(.VFlags, VFlags))
        Replacements(7) = Array("[JEFVCAHS]", "JEFVCAHS", t.GetFormattedValue(.JFlags, JFlags))
        Replacements(8) = Array("[INITPRESS]", "INITPRESS", t.GetFormattedValue(.Init.Pressure, InitPressure))
        Replacements(9) = Array("[INITTEMP]", "INITTEMP", t.GetFormattedValue(.Init.Temperature, InitTemperature))
        Replacements(10) = Array("[INITQUALITY]", "INITQUALITY", t.GetFormattedValue(.Init.Quality, InitQuality))
        Replacements(11) = Array("[EBTFLAG]", "EBTFLAG", t.GetFormattedValue(.Init.EbtFlag, EbtFlag))
        Replacements(12) = Array("[INITFLIQ]", "INITFLIQ", t.GetFormattedValue(.Init.FlowLiq, InitFlowLiq))
        Replacements(13) = Array("[INITFGAS]", "INITFGAS", t.GetFormattedValue(.Init.FlowGas, InitFlowGas))
        Replacements(14) = Array("[INITIVEL]", "INITIVEL", t.GetFormattedValue(.Init.InterfVel, InitInterFVel))
        Replacements(15) = Array("[DX]", "DX", t.GetFormattedValue(.Dx, Dx))
        Replacements(16) = Array("[VOLUME]", "VOLUME", t.GetFormattedValue(.VolumeM2, Volume))
        Replacements(17) = Array("[ANGLE1]", "ANGLE1", t.GetFormattedValue(.AngleAzi, AngleAzi))
        Replacements(18) = Array("[ANGLE2]", "ANGLE2", t.GetFormattedValue(.AngleVert, AngleVert))
        Replacements(19) = Array("[ELEVCHANGE]", "ELEVCHANGE", t.GetFormattedValue(.Elevchange, Elevchange))
        Replacements(20) = Array("[NAME]", "NAME", t.GetFormattedValue(Me.Name, Name_HydroComp))
        Replacements(21) = Array("[CWORD]", "CWORD", t.GetFormattedValue(.Init.CntrlWord, CWord))
    End With
    
    For i = LBound(Replacements) To UBound(Replacements)
        If IsEmpty(Replacements(i)) = False Then
            InputString = Replace(InputString, Replacements(i)(0), Replacements(i)(2))
            InputString = Replace(InputString, Replacements(i)(1), Replacements(i)(2))
        End If
    Next i
    
    ' Replace global replacements word stored under Settings
    For i = 1 To Settings.Replacements.Count
        InputString = Replace(InputString, Settings.Replacements.FindWord(i), Settings.Replacements.ReplaceWord(i))
    Next i
    
    ReplaceWords = InputString
    Exit Function
'errorHandler:
    'ReplaceWords = "<Error in HydroComponent.ReplaceWords>"
End Function
'
'
' STATIC METHODS






