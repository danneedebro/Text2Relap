VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComponentHydroJunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ComponentHydroJunction.cls
'
'
Option Explicit
'
' MEMBER VARIABLES
' Member variables - Shared
Private pBase As ComponentHydro
'
' Member variables - Specific for a junction type component
Private Type ValveInput
    Time() As Variant
    ValveStem() As Variant
    CntrlvarNumber As Integer
    TableNumber As Integer
End Type

Private Type PumpInput
    RatedFlow As Double
    RatedHead As Double
    Time() As Variant
    RatedSpeed() As Variant
    RatedTorque As Double
End Type

Private Type TmdpjunInput
    Time() As Variant
    Flowrate() As Variant
End Type

Private pValve As ValveInput
Private pPump As PumpInput
Private pTmdpjun As TmdpjunInput
'
'
' PROPERTIES
' Properties - Shared for all hydro components
Public Property Get Inputdeck() As Text2Relap
    Set Inputdeck = pBase.Inputdeck
End Property

Public Property Get HydroSystem() As HydroSystem
    Set HydroSystem = pBase.HydroSystem
End Property

Public Property Get Settings() As PropertiesSettings
    Set Settings = pBase.Settings
End Property

Public Property Get Connections() As CollectionConnections
    Set Connections = pBase.Connections
End Property

Public Property Get Segments() As CollectionSegments
    Set Segments = pBase.Segments
End Property

Public Property Get Info() As PropertiesHydroCompInfo
    Set Info = pBase.Info
End Property

Public Property Set Info(arg As PropertiesHydroCompInfo)
    Set pBase.Info = arg
End Property

Public Property Get CCC() As Integer
    CCC = pBase.CCC
End Property

Public Property Let CCC(arg As Integer)
    pBase.CCC = arg
End Property

Public Property Get Filename() As String
    Filename = pBase.Filename
End Property

Public Property Get Name() As String
    Name = pBase.Name
End Property
'
' Properties - Specific for a junction type component
' (none)

'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set pBase = New ComponentHydro
    ReDim pValve.Time(0)
    ReDim pValve.ValveStem(0)
    ReDim pPump.Time(0)
    ReDim pPump.RatedSpeed(0)
    ReDim pTmdpjun.Time(0)
    ReDim pTmdpjun.Flowrate(0)
End Sub
'
'
' CLASS METHODS
Public Sub Create(InputData As InputObject, ByRef BaseComponent As ComponentHydro)
' Action: Contructor for a junction-type-component
'
' Inputs:   inputRows = A Variant array with inputs from the worksheet this type
'                       of component
'           BaseComponent = A reference to the base hydrodynamic component
'
    On Error GoTo ErrorHandler
    
    ' Set the Base component object
    Set pBase = BaseComponent
    
    ' Create two junctions
    Dim i As Integer
    For i = 1 To 2
        pBase.Connections.Add New PropertiesHydroCompConnection
        Connections(i).ID = Inputdeck.CheckInput(InputData(1, 12 + i), Numeric:=False, valueDefault:=0, causeError:=True, FieldName:="Junction connection-ID", word:=12 + i, IsFatal:=True)
        Connections(i).Side = Inputdeck.CheckInput(InputData(1, 14 + i), True, valueDefault:=3 - i, causeError:=True, FieldName:="Junction connection-Side", word:=14 + i, IsFatal:=True)
    Next i

    ' Create one segment with values read into it
    pBase.Segments.Add New PropertiesHydroCompSegment
    
    ' Set default values for different components
    Dim Defaults As New PropertiesHydroCompSegment
    Defaults.VFlags = "0000000"    ' tlpvbfe
    Defaults.JFlags = "00000000"   ' jefvcahs
    If Me.Info.MainType = valve Then
        Defaults.JFlags = "00000100"
    End If
    
    With Segments(1)
        If Me.Info.MainType = pump Then
            .Dx = Inputdeck.CheckInput(InputData(1, 3), Numeric:=True, valueDefault:=0.2, causeError:=True, FieldName:="Dx", word:=3)
            .AngleVert = Inputdeck.CheckInput(InputData(1, 6), Numeric:=True, valueDefault:=0#, causeError:=True, FieldName:="Vertical Angle", word:=6)
            .AngleAzi = Inputdeck.CheckInput(InputData(1, 7), Numeric:=True, valueDefault:=0#, causeError:=True, FieldName:="Azimuthal Angle", word:=7)
            .Volumes = 1
        End If
        .ID = InputData(1, 2)
        .Area = Inputdeck.CheckInput(InputData(1, 5), Numeric:=True, valueDefault:=0, causeError:=True, FieldName:="Area", word:=5)
        .KLossPos = Inputdeck.CheckInput(InputData(1, 8), Numeric:=True, valueDefault:=0, causeError:=True, FieldName:="K+", word:=8)
        .KLossNeg = Inputdeck.CheckInput(InputData(1, 9), Numeric:=True, valueDefault:=0, causeError:=True, FieldName:="K-", word:=9)
        .Name = Inputdeck.CheckInput(InputData(1, 10), Numeric:=False, valueDefault:="Name", causeError:=False, FieldName:="Name", word:=10)
        .Drawing = Inputdeck.CheckInput(InputData(1, 11), Numeric:=False, valueDefault:="", causeError:=False, FieldName:="Drawing", word:=11)
        .VFlags = Inputdeck.CheckInput(InputData(1, 21), Numeric:=False, valueDefault:=Defaults.VFlags, causeError:=False, FieldName:="Volume flags", word:=21, IsBinaryString:=True)
        .JFlags = Inputdeck.CheckInput(InputData(1, 22), Numeric:=False, valueDefault:=Defaults.JFlags, causeError:=False, FieldName:="Junction flags", word:=22, IsBinaryString:=True)
        Set .Init = pBase.Init
    End With
    
    ' Read specific input for junction type components
    Dim ErrorMessages As String
    Dim s As New ResourceStringManipulation
    Dim Word20 As Variant
    If Me.Info.MainType <> sngljun Then
        Word20 = s.ParseBrackedExpression(InputData(1, 20), True)
    
        ' If second bit is set input is not a multiple of two
        If s.CheckBitSet(Word20(0), 2) = True Then ErrorMessages = ErrorMessages & "Input Error: Input is Not a multiple of two." & vbNewLine & vbNewLine
        If s.CheckBitSet(Word20(0), 3) = True Then ErrorMessages = ErrorMessages & "Input Error: Wrong format." & vbNewLine & vbNewLine
    End If
    
    If Me.Info.MainType = valve Then
        pValve.Time = Word20(1)
        pValve.ValveStem = Word20(2)
        If Me.Info.SubType <> srvvlv Then
            If UBound(pValve.Time) > 0 Then ErrorMessages = ErrorMessages & "Input Error: Entering the valve stem as a table is only permissable for a srvvlv component. Setting initial vlvstem=" & CStr(pValve.ValveStem(0)) & vbNewLine & vbNewLine
        End If
        
        If Me.Info.SubType = srvvlv Then
            pValve.CntrlvarNumber = Settings.CurrentCntrlvarNum
            pValve.TableNumber = Settings.CurrentTableNum
            Settings.IncrementCntrlvar
            Settings.IncrementTableNum
            
        ElseIf Me.Info.SubType = trpvlv Then
            
        End If
        
        If s.CheckBitSet(Word20(0), 4) = True Then ErrorMessages = ErrorMessages & "Input Error: Vlvstem (Y-value) > 1" & vbNewLine & vbNewLine
        
    ElseIf Me.Info.MainType = tmdpjun Then
        pTmdpjun.Time = Word20(1)
        pTmdpjun.Flowrate = Word20(2)
    ElseIf Me.Info.MainType = pump Then
        With pPump
            .RatedFlow = Inputdeck.CheckInput(InputData(1, 18), True, 0, False)
            .RatedHead = Inputdeck.CheckInput(InputData(1, 19), True, 0, False)
            .Time = Word20(1)
            .RatedSpeed = Word20(2)
            .RatedTorque = .RatedHead * 998.5 * 9.81 * .RatedFlow / IIf(.RatedSpeed(0) <> 0, .RatedSpeed(0), 1)
        
            If UBound(pPump.RatedSpeed) > 0 Then
            If .Time(0) <> -1 Then ErrorMessages = ErrorMessages & "Input Error: If entering the pump speed as a time vs speed table, make sure to let the first value pair be " & _
                                        "t=-1 and N=RatedSpeed" & vbNewLine & "Ex ""[-1;157;0;0;5;157;15;157]""" & vbNewLine & vbNewLine
            End If
        End With
    End If
    
    ' If input errors occured
    If ErrorMessages <> "" Then
        Inputdeck.HighlightCurrentLines 20
        Dim answ
        answ = MsgBox(ErrorMessages & vbNewLine & "Continue?", vbExclamation + vbOKCancel)
        Inputdeck.Warnings = True
        If answ = vbCancel Then
            Inputdeck.Abort = True
            Exit Sub
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".Create, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function ReplaceWords(ByVal InputString As String) As String
' Action: Replaces certain words with corresponding values
'
'
    On Error GoTo ErrorHandler

    Dim s As New ResourceSprintf
    Dim t As New ResourceTablePrint
    t.SetDefaultValues
    Dim Replacements(1 To 10) As Variant
    Dim i As Integer
    
    ' Replace keywords with properties specific to junction type components
    With Me.Segments(1)
        Replacements(1) = Array("[NNNVV000F]", "NNNVV000F", Me.Connections(1).ConnectionCode)
        Replacements(2) = Array("[MMMVV000F]", "MMMVV000F", Me.Connections(2).ConnectionCode)
        Replacements(3) = Array("[NNNVV]", "NNNVV", s.sprintf("%03d%02d", Me.Connections(1).CCC, Me.Connections(1).Volume))
        Replacements(4) = Array("[MMMVV]", "MMMVV", s.sprintf("%03d%02d", Me.Connections(2).CCC, Me.Connections(2).Volume))
    End With
        
    If Me.Info.MainType = pump Then
        Replacements(5) = Array("[PUMPFLOW]", "PUMPFLOW", t.GetFormattedValue(pPump.RatedFlow, PumpFlow))
        Replacements(6) = Array("[PUMPHEAD]", "PUMPHEAD", t.GetFormattedValue(pPump.RatedHead, PumpHead))
        Replacements(7) = Array("[PUMPRPM]", "PUMPRPM", t.GetFormattedValue(pPump.RatedSpeed(0), PumpSpeed))
        Replacements(8) = Array("[PUMPTORQUE]", "PUMPTORQUE", t.GetFormattedValue(pPump.RatedTorque, PumpTorque))
        
        For i = 0 To UBound(pPump.RatedSpeed)
            InputString = Replace(InputString, s.sprintf("[TIME-%d]", i + 1), t.GetFormattedValue(pPump.Time(i), Time))
            InputString = Replace(InputString, s.sprintf("[PUMPVEL-%d]", i + 1), t.GetFormattedValue(pPump.RatedSpeed(i), PumpSpeed))
        Next i
        
    ElseIf Me.Info.MainType = tmdpjun Then
        For i = 0 To UBound(pTmdpjun.Flowrate)
            InputString = Replace(InputString, s.sprintf("[TIME-%d]", i + 1), t.GetFormattedValue(pTmdpjun.Time(i), Time))
            InputString = Replace(InputString, s.sprintf("[MASSFLOW-%d]", i + 1), t.GetFormattedValue(pTmdpjun.Flowrate(i), Massflow))
        Next i
        
    ElseIf Me.Info.SubType = srvvlv Then
        Replacements(5) = Array("[VALVESTEM]", "VALVESTEM", t.GetFormattedValue(pValve.ValveStem(0), ValveOpen))
        Replacements(6) = Array("[NNN]", "XYQTWSA", s.sprintf("%03d", pValve.CntrlvarNumber))
        Replacements(7) = Array("[NNNN]", "XYQTWSA", s.sprintf("%04d", pValve.CntrlvarNumber))
        Replacements(8) = Array("[N]", "XYQTWSA", s.sprintf("%4d", pValve.CntrlvarNumber))
        Replacements(9) = Array("[TTT]", "XYQTWSA", s.sprintf("%03d", pValve.TableNumber))
        Replacements(10) = Array("[T]", "XYQTWSA", s.sprintf("%3d", pValve.TableNumber))
        
        For i = 0 To UBound(pValve.ValveStem)
            InputString = Replace(InputString, s.sprintf("[TIME-%d]", i + 1), t.GetFormattedValue(pValve.Time(i), Time))
            InputString = Replace(InputString, s.sprintf("[VALVESTEM-%d]", i + 1), t.GetFormattedValue(pValve.ValveStem(i), ValveOpen))
        Next i
    ElseIf Me.Info.SubType = chkvlv Then
        Replacements(5) = Array("[CHKVLVINIT]", "CHKVLVINIT", s.sprintf("%2d", 1 - pValve.ValveStem(0)))
    ElseIf Me.Info.MainType = valve Then
        Replacements(5) = Array("[VALVESTEM]", "VALVESTEM", t.GetFormattedValue(pValve.ValveStem(0), ValveOpen))
        
    End If

    
    For i = LBound(Replacements) To UBound(Replacements)
        If IsEmpty(Replacements(i)) = False Then
            InputString = Replace(InputString, Replacements(i)(0), Replacements(i)(2))
            InputString = Replace(InputString, Replacements(i)(1), Replacements(i)(2))
        End If
    Next i

    ' Replace properties that are shared for all hydro components
    InputString = pBase.ReplaceWords(InputString)
    ReplaceWords = InputString

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".ReplaceWords, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function



Public Sub WriteToFile(ByRef ts As Object, ByVal wC As Boolean, Optional ByVal GenerateIncludeFile As Boolean = False)
' Action: Writes the component to a file stored in the TextStream-object
'
' Inputs:   ts = A TextStream object (included in the MS Scripting Runtime)
'           wC = writeComments, True if comments are to be written, false otherwise
'           GenerateIncludeFile = Write default include file (send arg to fixNum to for
'                                 instance write out CCC instead of componentnumber
'
    On Error GoTo ErrorHandler
    
    If Me.Filename <> "" And GenerateIncludeFile = False Then
        WriteCustom ts, wC
        Exit Sub
    End If
    
    Dim t As New ResourceTablePrint
    t.SetDefaultValues
    t.OutputKeywords = IIf(GenerateIncludeFile = True, 1, 0)
    t.OutputKeywordBrackets = 1
    
    ' Set commonly used arguments
    Dim pCCC As ResourceTablePrint
    Set pCCC = t.Args(CCC, CCC_Code, OutputKeywordBrackets:=0, OutputKeywordIndexes:=0)
    Dim card As ResourceTablePrint
    Set card = t.Args(0, CardNumberHydro)
    
    ' HEADER
    t.AddLine "*-------------------------------------------------------------------"
    t.AddLine "*        Name      Component"
    t.AddLine "{ }{  }  {NAME  }  {TYPE}", pCCC, card, t.Args(Me.Name, Name_HydroComp), t.Args(Me.Info.MainTypeStr)
    t.AddLine "*"
    t.AddLine "* Drawing: {}", t.Args(Me.Segments(1).Drawing, Drawing)
    t.AddLine "*-------------------------------------------------------------------"
    ts.WriteLine t.WriteAndClear
    
    ' Sngljun or valve
    If Me.Info.MainType = sngljun Or Me.Info.MainType = valve Then
        With Segments(1)
            
            ' CARD 101-109
            t.AddLine "*        From.comp    To.comp     Area     K+     K-  jefvcahs"
            t.AddLine "{ }{  }  {To}  {From}  {Area}  {KPOS}  {KNEG}  {JFLAGS}", pCCC, card.Update(101), t.Args(Connections(1).ConnectionCode, _
                                                                             ConnectionCode1), t.Args(Connections(2).ConnectionCode, ConnectionCode2), _
                                                                             t.Args(.Area, Area), t.Args(.KLossPos, KLossNeg), t.Args(.KLossNeg, KLossNeg), _
                                                                             t.Args(.JFlags, JFlags)
            t.AddLine "*"
            ts.WriteLine t.WriteAndClear ' Write to file
        
            ' CARD 201 Junction initial conditions
            t.AddLine "*        cword  Liq.flow  Vap.flow  intf.vel"
            t.AddLine "{ }{  }      {}   {     }   {     }   {     }", pCCC, card.Update(201), t.Args(.Init.CntrlWord, CWord), t.Args(.Init.FlowLiq, InitFlowLiq), t.Args(.Init.FlowGas, InitFlowGas), t.Args(.Init.InterfVel, InitInterFVel)
            t.AddLine "*"
            ts.WriteLine t.WriteAndClear
            
            Select Case Info.MainType
                Case ComponentType.valve
                    WriteValve ts, wC, GenerateIncludeFile
                Case ComponentType.tmdpjun
                    WriteTmdpjun ts, wC, GenerateIncludeFile
                Case ComponentType.pump
                    WritePump ts, wC, GenerateIncludeFile
            End Select
        End With
    
    ' Tmdpjun
    ElseIf Me.Info.MainType = tmdpjun Then
        WriteTmdpjun ts, wC, GenerateIncludeFile
        
    ' Pump
    ElseIf Me.Info.MainType = pump Then
        WritePump ts, wC, GenerateIncludeFile
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".WriteToFile, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub WriteValve(ByRef ts As Object, ByVal wC As Boolean, Optional ByVal GenerateIncludeFile As Boolean = False)
' Action: Writes a valve component
'
    On Error GoTo ErrorHandler

    Dim s As New ResourceStringManipulation
    Dim i As Integer
        
    Dim t As New ResourceTablePrint
    t.SetDefaultValues
    t.OutputKeywords = IIf(GenerateIncludeFile = True, 1, 0)
    t.OutputKeywordBrackets = 1
    
    ' Set commonly used arguments
    Dim pCCC As ResourceTablePrint
    Set pCCC = t.Args(CCC, CCC_Code, OutputKeywordBrackets:=0, OutputKeywordIndexes:=0)
    Dim card As ResourceTablePrint
    Set card = t.Args(0, CardNumberHydro)
    
    ' CARD 300
    t.AddLine "*        valve.type"
    t.AddLine "{ }{  }  {}", pCCC, card.Update(300), t.Args(Me.Info.SubTypeStr)
    t.AddLine "*"
    ts.WriteLine t.WriteAndClear
    
    ' MOTOR VALVE (MTRVLV)
    If Me.Info.SubType = mtrvlv Then
        Dim DefaultOpenTrip As Integer, DefaultCloseTrip As Integer
        If Settings.TripFormat = 1 Then
            DefaultOpenTrip = 1
            DefaultCloseTrip = 2
        ElseIf Settings.TripFormat = 2 Then
            DefaultOpenTrip = 1
            DefaultCloseTrip = 2
        Else
            DefaultOpenTrip = 401
            DefaultCloseTrip = 402
        End If
        
        ' CARD 301
        t.AddLine "*        trp.open  trp.close  rate  init.pos"
        t.AddLine "{ }{  }  {   }     {   }      0.50  {}", pCCC, card.Update(301), t.Args(DefaultOpenTrip, FStr:="%5d", Key:="TRIP:DFLT1"), _
                                                                  t.Args(DefaultCloseTrip, FStr:="%5d", Key:="TRIP:DFLT2"), t.Args(pValve.ValveStem(0), ValveOpen)
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear
        
        ' CARD 401-499 - CV-TABLE
        t.AddLine "* Cv-table"
        t.AddLine "*        norm.area  Forward.Cv  Reverse.Cv"
        t.AddLine "*{ }{  }      {   }     {     }     {     }", pCCC, card.Update(401), t.Args(0#, FStr:="%5.3f"), t.Args(0#, FStr:="%7.2f"), t.Args(0#, FStr:="%7.2f")
        t.AddLine "*{ }{  }      {   }     {     }     {     }", pCCC, card.Update(402), t.Args(1#, FStr:="%5.3f"), t.Args(2000#, FStr:="%7.2f"), t.Args(2000#, FStr:="%7.2f")
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear
            
    ' SERVO VALVE (SRVVLV)
    ElseIf Me.Info.SubType = srvvlv Then
        ' CARD 301
        t.AddLine "*        cntrlvar"
        t.AddLine "{ }{  }  {  }", pCCC, card.Update(301), t.Args(pValve.CntrlvarNumber, Key:="N", FStr:="%4d")
        t.AddLine "*"
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear
        
        If Settings.CntrlvarFormat = 9999 Then
            t.AddLine "205{  }0  Func  function  1.000  {}  1", t.Args(pValve.CntrlvarNumber, Key:="NNNN", FStr:="%04d"), t.Args(pValve.ValveStem(0), ValveOpen)
            t.AddLine "205{  }1  time  0  {}", t.Args(pValve.CntrlvarNumber, Key:="NNNN", FStr:="%04d"), t.Args(pValve.TableNumber, Key:="T", FStr:="%3d")
        Else
            t.AddLine "205{ }00  Func  function  1.000  {}  1", t.Args(pValve.CntrlvarNumber, Key:="NNN", FStr:="%03d"), t.Args(pValve.ValveStem(0), ValveOpen)
            t.AddLine "205{ }01  time  0  {}", t.Args(pValve.CntrlvarNumber, Key:="NNN", FStr:="%03d"), t.Args(pValve.TableNumber, Key:="T", FStr:="%3d")
        End If
        ts.WriteLine t.WriteAndClear
                
        t.AddLine "* vlvarea = f(t)*A"
        t.AddLine "202{ }00   normarea", t.Args(pValve.TableNumber, Key:="TTT", FStr:="%03d")
        t.AddLine "*            Time(s)   Opening"
        
        t.OutputKeywordIndexes = 1  ' Resets index-counter
        For i = 0 To UBound(pValve.ValveStem)
            t.AddLine "202{ }{}      {}    {}", t.Args(pValve.TableNumber, Key:="TTT", FStr:="%03d", OutputKeywordIndexes:=0), _
                                                t.Args(i + 1, FStr:="%02d"), t.Args(pValve.Time(i), Time), _
                                                t.Args(pValve.ValveStem(i), ValveOpen, Key:="VALVESTEM")
        Next i
        t.OutputKeywordIndexes = 0
        
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear
    
 
    ' CHECK VALVE (CHKVLV)
    ElseIf Me.Info.SubType = chkvlv Then
        ' CARD 301
        t.AddLine "*        Type  Init   dP   Leak"
        t.AddLine "{ }{  }     0    {}  0.0   0.00", pCCC, card.Update(301), t.Args(1 - pValve.ValveStem(0), Key:="CHKVLVINIT", FStr:="%2d")
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear
        
    ' INERTIAL SWING CHECK VALVE (INRVLV)
    ElseIf Me.Info.SubType = inrvlv Then
        Dim FlapperRadius As Double, FlapperThickness As Double, FlapperVolume As Double, FlapperMass As Double
        Dim MomentLength As Double, Inertia As Double, CrackingPressure As Double, LeakFraction As Double
        Dim AngleMax As Double, AngleMin As Double, AngleInit As Double
        
        FlapperRadius = Sqr(Me.Segments(1).Area / 3.14159265)
        FlapperThickness = 0.2 * FlapperRadius
        FlapperVolume = 3.14159265 * FlapperRadius ^ 2 * FlapperThickness
        FlapperMass = FlapperVolume * 7800
        MomentLength = FlapperRadius
        Inertia = FlapperMass * FlapperRadius ^ 2 / 4 + FlapperMass * MomentLength ^ 2
        AngleMax = 75
        AngleMin = 0
        AngleInit = pValve.ValveStem(0) * AngleMax
    
        ' CARD 301
        t.AddLine "*        Latch  Init  P.crack  Leak.fr  Angle.init  Angle.min"
        t.AddLine "{ }{  }      0     0     {  }    {   }        {  }       {  }", pCCC, card.Update(301), t.Args(CrackingPressure, FStr:="%4.1f"), _
                                                                                   t.Args(LeakFraction, FStr:="%5.3f"), t.Args(AngleInit, FStr:="%4.1f"), t.Args(AngleMin, FStr:="%4.1f")
        ' CARD 302
        t.AddLine "*"
        t.AddLine "*        Angle.max   Inertia  Ang.vel  Mom.len  Flap.rad  Subm.weight"
        t.AddLine "{ }{  }       {  }  {   }    {   }   {    }   {    }       {   }", pCCC, card.Update(302), t.Args(AngleMax, FStr:="%4.1f"), _
                                                                                                t.Args(Inertia, FStr:="%5.3g"), t.Args(0#, FStr:="%5.3f"), _
                                                                                                t.Args(MomentLength, FStr:="%6.4f"), t.Args(FlapperRadius, _
                                                                                                FStr:="%6.4g"), t.Args(FlapperMass, FStr:="%5.3g")
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear

    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".WriteValve, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Sub

Private Sub WriteTmdpjun(ByRef ts As Object, ByVal wC As Boolean, Optional ByVal GenerateIncludeFile As Boolean = False)
' Action: Writes a time-dependant junction component
'
    On Error GoTo ErrorHandler

    Dim t As New ResourceTablePrint
    t.SetDefaultValues
    t.OutputKeywords = IIf(GenerateIncludeFile = True, 1, 0)
    t.OutputKeywordBrackets = 1
    
    ' Set commonly used arguments
    Dim pCCC As ResourceTablePrint
    Set pCCC = t.Args(CCC, CCC_Code, OutputKeywordBrackets:=0, OutputKeywordIndexes:=0)
    Dim card As ResourceTablePrint
    Set card = t.Args(0, CardNumberHydro)
    
    With Segments(1)
        
        ' CARD 101-109
        t.AddLine "*        From.comp    To.comp      Area  jefvcahs"
        t.AddLine "{ }{  }  {       }  {       }   {     }  {      }", pCCC, card.Update(101), t.Args(Connections(1).ConnectionCode, ConnectionCode1), _
                                                         t.Args(Connections(2).ConnectionCode, ConnectionCode2), t.Args(.Area, Area), t.Args(.JFlags, JFlags)
        t.AddLine "*"
        t.AddLine "*        cword  tbl.trip  plotvar  plotnum"
        t.AddLine "{ }{  }      1         0     time        0", pCCC, card.Update(200)
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear
        
        t.AddLine "*            Time  Liq.flow  Vap.flow  intf.vel"
        
        t.OutputKeywordIndexes = 1 ' Resets index-counter
        Dim i As Integer
        For i = 0 To UBound(pTmdpjun.Time)
            t.AddLine "{ }{  }  {      }   {     }   {     }   {     }", pCCC, card.Update(201 + i), t.Args(pTmdpjun.Time(i), Time), _
                                                                         t.Args(pTmdpjun.Flowrate(i), Massflow), t.Args(0#, Massflow, OutputKeywords:=0), _
                                                                         t.Args(0#, Massflow, OutputKeywords:=0)
        Next i
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear ' Write to file
        
    End With
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".WriteTmdpjun, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Private Sub WritePump(ByRef ts As Object, ByVal wC As Boolean, Optional ByVal GenerateIncludeFile As Boolean = False)
' Action: Writes a pump component
'
    On Error GoTo ErrorHandler

    Dim t As New ResourceTablePrint
    t.SetDefaultValues
    t.OutputKeywords = IIf(GenerateIncludeFile = True, 1, 0)
    t.OutputKeywordBrackets = 1
    
    ' Set commonly used arguments
    Dim pCCC As ResourceTablePrint
    Set pCCC = t.Args(CCC, CCC_Code, OutputKeywordBrackets:=0, OutputKeywordIndexes:=0)
    Dim card As ResourceTablePrint
    Set card = t.Args(0, CardNumberHydro)
    
    Dim TripA As Integer, TripB As Integer
    If Settings.TripFormat = 1 Then
        TripA = 3
        TripB = 4
    ElseIf Settings.TripFormat = 2 Then
        TripA = 3
        TripB = 4
    Else
        TripA = 403
        TripB = 404
    End If
    
    
    With Segments(1)
        
        ' CARD 101
        t.AddLine "*** GEOMETRI ***"
        t.AddLine "*           Area         Dx   Volume  Azi.ang    Angle  Elev.ch  tlpvbfe"
        t.AddLine "{ }{  }  {     }  {       }  {     }  {     }  {   }  {     }  {     }", pCCC, card.Update(101), t.Args(.Area, Area), t.Args(.Dx, Dx), _
                                                                                              t.Args(.VolumeM2, Volume), t.Args(.AngleAzi, AngleAzi), _
                                                                                              t.Args(.AngleVert, AngleVert), t.Args(.Elevchange, Elevchange), t.Args(.VFlags, VFlags)
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear
        
        ' CARD 108 and 109
        t.AddLine "*** JUNCTIONS ***"
        t.AddLine "*        From.comp     Area     K+     K-  jefvcahs"
        t.AddLine "{ }{  }  {       }  {     }  {   }  {   }  {      }", pCCC, card.Update(108), t.Args(Connections(1).ConnectionCode, ConnectionCode1), _
                                                                         t.Args(.Area, Area), t.Args(0#, KLossPos, OutputKeywords:=0), t.Args(0#, KLossNeg, OutputKeywords:=0), _
                                                                         t.Args(.JFlags, JFlags)
        t.AddLine "*"
        t.AddLine "*          To.comp     Area     K+     K-  jefvcahs"
        t.AddLine "{ }{  }  {       }  {     }  {   }  {   }  {      }", pCCC, card.Update(109), t.Args(Connections(2).ConnectionCode, ConnectionCode2), _
                                                                         t.Args(.Area, Area), t.Args(.KLossPos, KLossPos), t.Args(.KLossNeg, KLossNeg), _
                                                                         t.Args(.JFlags, JFlags)
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear
        
        ' CARD 200, 201, 202
        t.AddLine "*** INITIAL CONDITIONS ***"
        Select Case Right(.Init.EbtFlag, 1)
            Case "1"
                t.AddLine "*        ebt   Temp    x_s"
                t.AddLine "{ }{  }  { }  {       }  {    }", pCCC, card.Update(200), t.Args(.Init.EbtFlag, EbtFlag), t.Args(.Init.Temperature, InitTemperature), _
                                                                   t.Args(.Init.Quality, InitQuality)
            Case "2"
                t.AddLine "*        ebt   Pressure    x_s"
                t.AddLine "{ }{  }  { }  {       }  {    }", pCCC, card.Update(200), t.Args(.Init.EbtFlag, EbtFlag), t.Args(.Init.Pressure, InitPressure), _
                                                                   t.Args(.Init.Quality, InitQuality)
            Case "3"
                t.AddLine "*        ebt   Pressure    Temp"
                t.AddLine "{ }{  }  { }  {       }  {    }", pCCC, card.Update(200), t.Args(.Init.EbtFlag, EbtFlag), t.Args(.Init.Pressure, InitPressure), _
                                                                   t.Args(.Init.Temperature, InitTemperature)
            Case "4"
                t.AddLine "*        ebt   Pressure    Temp   x_s"
                t.AddLine "{ }{  }  { }  {       }  {    }   { }", pCCC, card.Update(200), t.Args(.Init.EbtFlag, EbtFlag), t.Args(.Init.Pressure, InitPressure), _
                                                                   t.Args(.Init.Temperature, InitTemperature), t.Args(.Init.Quality, InitQuality)
        End Select
        t.AddLine "*"
        t.AddLine "*        cword  Liq.flow  Gas.flow  Interf.vel"
        t.AddLine "{ }{  }     {}    {}    {}", pCCC, card.Update(201), t.Args(.Init.CntrlWord, CWord), t.Args(.Init.FlowLiq, InitFlowLiq), _
                                                        t.Args(.Init.FlowGas, InitFlowGas), t.Args(.Init.InterfVel, InitInterFVel)
        t.AddLine "*"
        t.AddLine "*        cword  Liq.flow  Gas.flow  Interf.vel"
        t.AddLine "{ }{  }     {}   {}    {}", pCCC, card.Update(202), t.Args(.Init.CntrlWord, CWord), t.Args(.Init.FlowLiq, InitFlowLiq), _
                                                        t.Args(.Init.FlowGas, InitFlowGas), t.Args(.Init.InterfVel, InitInterFVel)
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear
        
        ' CARD 301 - 303
        t.AddLine "*** PUMP DESCRIPTION ***"
        t.AddLine "*        pmp.tbl  2phInd  2phDiff  Trq.tbl  t.dep  trip.n  rev.ind"
        t.AddLine "{ }{  }        0      -1       -3       -1      0   {   }        1", pCCC, card.Update(301), t.Args(TripA, FStr:="%5d", Key:="TRIP:DFLT3")
        t.AddLine "*"
        t.AddLine "*        r.speed  v.ratio  r.flow  r.head   r.trq  inertia      rho"
        t.AddLine "{ }{  }   {    }    0.000  {    }  {    }  {    }    10.00   998.50", pCCC, card.Update(302), t.Args(pPump.RatedSpeed(0), PumpSpeed), _
                                                                                         t.Args(pPump.RatedFlow, PumpFlow), t.Args(pPump.RatedHead, PumpHead), _
                                                                                         t.Args(pPump.RatedTorque, PumpTorque)
        t.AddLine "* ..contd"
        t.AddLine "*        r.m.trq   TF2   TF0   TF1   TF3"
        t.AddLine "{ }{  }    0.000  0.00  0.00  0.00  0.00", pCCC, card.Update(303)
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear
        
        
        ' ---------------- PUMP CURVES ------------------------------
        ' -----------------------------------------------------------
        t.AddLine "*** HOMOLOGOUS PUMP CURVES ***"
                
        Dim CurveType As Integer, CurveRegime As Integer, i As Integer
        Dim CardNum As Integer
    
        Dim Pumpcurves As Variant
        Pumpcurves = GetPumpcurves()
    
        Dim tmpArr As Variant, tmpArr2 As Variant, tmpArr3 As Variant
        tmpArr = Array("AN", "VN", "AD", "VD", "AT", "VT", "AR", "VR")
        tmpArr2 = Array("Q>0", "Q>0", "Q<0", "Q<0", "Q<0", "Q<0", "Q>0", "Q>0")
        tmpArr3 = Array("N>0", "N>0", "N>0", "N>0", "N<0", "N<0", "N<0", "N<0")
        
        CardNum = 1000
        For CurveType = 1 To 2
            For CurveRegime = 1 To 8
                CardNum = CardNum + 100
                t.AddLine "*"
                t.AddLine "* {}{} ({}, {})", IIf(CurveType = 1, "H", "B"), CStr(tmpArr(CurveRegime - 1)), CStr(tmpArr2(CurveRegime - 1)), CStr(tmpArr3(CurveRegime - 1))
                t.AddLine "{ }{  }  {}  {}", pCCC, card.Update(CardNum), CurveType, CurveRegime
                If CurveRegime Mod 2 = 0 Then  ' 2=HVN, 4=HVD, 6=HVT, 8=HVR  (a/v and h/v^2)
                    t.AddLine "*             a/v     {}/v^2", IIf(CurveType = 1, "h", "b")
                Else  ' 1=HAN, 3=HAD, 5=HAT, 7=HAR  (v/a and h/a^2)
                    t.AddLine "*             v/a     {}/a^2", IIf(CurveType = 1, "h", "b")
                End If
                For i = 0 To UBound(Pumpcurves(CurveType, CurveRegime, 1))
                    t.AddLine "{ }{  }   {}   {}", pCCC, card.Update(CardNum + i + 1), t.Args(Pumpcurves(CurveType, CurveRegime, 1)(i), FStr:="% 6.4f"), t.Args(Pumpcurves(CurveType, CurveRegime, 2)(i), FStr:="% 6.4f")
                Next i
                
                ts.WriteLine t.WriteAndClear
            Next CurveRegime
        Next CurveType
        
        ' CARD 6100 - Pump speed control
        t.AddLine "*"
        t.AddLine "*** PUMP VELOCITY CONTROL ***"
        t.AddLine "*"
        t.AddLine "*        trip  plotvar  plotnum"
        t.AddLine "{ }{  } {   }  time     0", pCCC, card.Update(6100), t.Args(TripB, FStr:="%5d", Key:="TRIP:DFLT4")
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear ' Write to file
        
        ' CARD 6101-6199 - Pump speed control table
        t.AddLine "*            Time  Pump.vel"
        t.OutputKeywordIndexes = 1
        For i = 0 To UBound(pPump.Time)
            t.AddLine "{ }{  }  {      }     {   }", pCCC, card.Update(6101 + i), t.Args(pPump.Time(i), Time), t.Args(pPump.RatedSpeed(i), PumpSpeed, Key:="PUMPVEL")
        Next i
        t.AddLine "*"
        ts.WriteLine t.WriteAndClear ' Write to file
        
    End With
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".WritePump, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Private Sub WriteCustom(ByRef ts As Object, ByVal wC As Boolean)
' Action: Writes a custom component
'
    ' Activate Microsoft Scripting runtime Tools...References...Microsoft scripting runtime
    On Error GoTo ErrorHandler
    Dim fso As Scripting.FileSystemObject
    Dim tsRead As Scripting.TextStream
    Dim tsWrite As Scripting.TextStream
    
    Dim Includefile As New ResourceFileObject
    Includefile.CreateByParts pBase.HydroSystem.Settings.IncludeFolder, Me.Filename
    
    Set fso = New Scripting.FileSystemObject
    
    ' If include file doesn't exist, create it
    If Includefile.FileExists = False Then
        Dim answ, CompType As String
        CompType = Me.Info.MainTypeStr & "-" & Me.Info.SubTypeStr
        Inputdeck.HighlightCurrentLines 10, pBase.RowBegin
        answ = MsgBox("Warning: Custom hydrodynamic component '" & Includefile.FullPath & "' doesn't exist. A generic " & CompType & "-type includefile will be created", vbExclamation + vbOKCancel, "Include file not found")
        Inputdeck.Warnings = True
        If answ = vbCancel Then
            Inputdeck.Abort = True
            Exit Sub
        End If
        If Includefile.FolderExists = False Then Includefile.CreateFolders ' Create subfolders if they don't exist
        Set tsWrite = fso.CreateTextFile(Includefile.FullPath, True)
        Me.WriteToFile tsWrite, wC, GenerateIncludeFile:=True
        tsWrite.Close
    End If
    
    ' Open include file for reading
    Set tsRead = fso.OpenTextFile(Includefile.FullPath, ForReading)
   
    ' Read all lines of the file and write them to file
    Do While tsRead.AtEndOfStream <> True
        ts.WriteLine ReplaceWords(tsRead.ReadLine)
    Loop
    
    tsRead.Close
    Set fso = Nothing
    
    Exit Sub
ErrorHandler:
    MsgBox "Error writing/creating Custom hydro component " & Me.Filename
    If Not tsWrite Is Nothing Then tsWrite.Close
    If Not tsRead Is Nothing Then tsRead.Close
    
    Err.Raise Err.Number, TypeName(Me) & ".WriteCustom, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
'
'
' STATIC METHODS
Private Function GetPumpcurves() As Variant
' Action: Returns a homologous pump curve
'
' Output:
'  Pumpcurve(CurveType, CurveRegime, indepVar or depVar)(index)
'       CurveType:  1=HEAD CURVE,  2=TORQUE CURVE
'       CurveRegime: 1=HAN or BAN, 2=HVN or BVN, 3=HAD or BAD, 4=HVD or BVD
'                    5=HAT or BAT, 6=HVT or BVT, 7=HAR or BAR, 8=HVR or BVR
'       Independant var: 1 = v/a for CurveRegime=1, 3, 5, 7
'                            a/v for CurveRegime=2, 4, 6, 8
'       Dependant: 2 = h/a^2 or b/a^2 for CurveRegime=1, 3, 5, 7
'                      h/v^2 or b/v^2 for CurveRegime=2, 4, 6, 8
'       Index: from 0 to N with up to 99 points
'
    Dim Pumpcurves(1 To 2, 1 To 8, 1 To 2) As Variant
    
    ' HEAD CURVES
    ' HAN
    Pumpcurves(1, 1, 1) = Split("0.0000,0.1000,0.2000,0.3000,0.4000,0.5000,0.6000,0.7000,0.8000,0.9000,1.0000", ",")
    Pumpcurves(1, 1, 2) = Split("1.5700,1.5396,1.5251,1.4890,1.4620,1.4288,1.3751,1.3265,1.2315,1.1231,1.0000", ",")

    Pumpcurves(1, 2, 1) = Split("0.0000,0.1000,0.2000,0.3000,0.4000,0.5000,0.6000,0.7000,0.8000,0.9000,1.0000", ",")
    Pumpcurves(1, 2, 2) = Split("-1.2200,-1.0563,-0.8924,-0.7325,-0.5512,-0.3678,-0.1490,0.0889,0.3745,0.6760,1.0000", ",")
    
    Pumpcurves(1, 3, 1) = Split("0.0000,-0.1000,-0.2000,-0.3000,-0.4000,-0.5000,-0.6000,-0.7000,-0.8000,-0.9000,-1.0000", ",")
    Pumpcurves(1, 3, 2) = Split("1.5700,1.6203,1.7061,1.8314,1.9845,2.1703,2.3721,2.5629,2.7728,3.0120,3.2800", ",")
    
    Pumpcurves(1, 4, 1) = Split("0.0000,-0.1000,-0.2000,-0.3000,-0.4000,-0.5000,-0.6000,-0.7000,-0.8000,-0.9000,-1.0000", ",")
    Pumpcurves(1, 4, 2) = Split("1.3000,1.3692,1.4538,1.5626,1.7061,1.8742,2.0777,2.3243,2.6064,2.9248,3.2800", ",")
    
    Pumpcurves(1, 5, 1) = Split("0.0000,0.1000,0.2000,0.3000,0.4000,0.5000,0.6000,0.7000,0.8000,0.9000,1.0000", ",")
    Pumpcurves(1, 5, 2) = Split("0.3300,0.3319,0.3301,0.3305,0.3280,0.3258,0.3264,0.3576,0.4656,0.5934,0.7400", ",")
    
    Pumpcurves(1, 6, 1) = Split("0.0000,0.1000,0.2000,0.3000,0.4000,0.5000,0.6000,0.7000,0.8000,0.9000,1.0000", ",")
    Pumpcurves(1, 6, 2) = Split("1.3000,1.2525,1.2054,1.1545,1.0616,0.9413,0.8578,0.8197,0.7700,0.7460,0.7400", ",")
    
    Pumpcurves(1, 7, 1) = Split("0.0000,-0.1000,-0.2000,-0.3000,-0.4000,-0.5000,-0.6000,-0.7000,-0.8000,-0.9000,-1.0000", ",")
    Pumpcurves(1, 7, 2) = Split("0.3300,0.3102,0.2825,0.2139,0.1322,0.0038,-0.1806,-0.4019,-0.5989,-0.7959,-1.0000", ",")
    
    Pumpcurves(1, 8, 1) = Split("0.0000,-0.1000,-0.2000,-0.3000,-0.4000,-0.5000,-0.6000,-0.7000,-0.8000,-0.9000,-1.0000", ",")
    Pumpcurves(1, 8, 2) = Split("-1.2200,-1.2986,-1.2208,-1.1620,-1.1182,-1.0859,-1.0618,-1.0432,-1.0280,-1.0141,-1.0000", ",")

    ' TORQUE CURVES
    Pumpcurves(2, 1, 1) = Split("0.0000,0.1000,0.2000,0.3000,0.4000,0.5000,0.6000,0.7000,0.8000,0.9000,1.0000", ",")
    Pumpcurves(2, 1, 2) = Split("0.9900,0.9134,0.8966,0.9045,0.9312,0.9679,0.9933,1.0282,1.0236,1.0141,1.0000", ",")
    
    Pumpcurves(2, 2, 1) = Split("0.0000,0.1000,0.2000,0.3000,0.4000,0.5000,0.6000,0.7000,0.8000,0.9000,1.0000", ",")
    Pumpcurves(2, 2, 2) = Split("-1.3500,-1.1184,-0.8865,-0.6447,-0.3595,-0.1092,0.1099,0.3125,0.5485,0.7741,1.0000", ",")
    
    Pumpcurves(2, 3, 1) = Split("0.0000,-0.1000,-0.2000,-0.3000,-0.4000,-0.5000,-0.6000,-0.7000,-0.8000,-0.9000,-1.0000", ",")
    Pumpcurves(2, 3, 2) = Split("0.9900,1.0806,1.1997,1.3594,1.5456,1.7610,1.9851,2.2349,2.5320,2.8596,3.2200", ",")
    
    Pumpcurves(2, 4, 1) = Split("0.0000,-0.1000,-0.2000,-0.3000,-0.4000,-0.5000,-0.6000,-0.7000,-0.8000,-0.9000,-1.0000", ",")
    Pumpcurves(2, 4, 2) = Split("1.5700,1.6506,1.7291,1.8166,1.9256,2.0750,2.2550,2.4585,2.6940,2.9468,3.2200", ",")
    
    Pumpcurves(2, 5, 1) = Split("0.0000,0.1000,0.2000,0.3000,0.4000,0.5000,0.6000,0.7000,0.8000,0.9000,1.0000", ",")
    Pumpcurves(2, 5, 2) = Split("-0.7700,-0.6840,-0.5940,-0.4896,-0.3832,-0.2741,-0.1453,-0.0152,0.1277,0.3142,0.5400", ",")
    
    Pumpcurves(2, 6, 1) = Split("0.0000,0.1000,0.2000,0.3000,0.4000,0.5000,0.6000,0.7000,0.8000,0.9000,1.0000", ",")
    Pumpcurves(2, 6, 2) = Split("1.5700,1.4790,1.3867,1.2851,1.1818,1.0749,0.9807,0.8944,0.7679,0.6523,0.5400", ",")
    
    Pumpcurves(2, 7, 1) = Split("0.0000,-0.1000,-0.2000,-0.3000,-0.4000,-0.5000,-0.6000,-0.7000,-0.8000,-0.9000,-1.0000", ",")
    Pumpcurves(2, 7, 2) = Split("-0.7700,-0.8830,-1.0365,-1.2731,-1.6626,-2.1892,-2.6536,-3.0393,-3.4776,-3.9634,-4.5000", ",")
    
    Pumpcurves(2, 8, 1) = Split("0.0000,-0.1000,-0.2000,-0.3000,-0.4000,-0.5000,-0.6000,-0.7000,-0.8000,-0.9000,-1.0000", ",")
    Pumpcurves(2, 8, 2) = Split("-1.3500,-1.5365,-1.7569,-2.0176,-2.3038,-2.6212,-3.0261,-3.4714,-3.8452,-4.1816,-4.5000", ",")

    GetPumpcurves = Pumpcurves

End Function

