VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComponentHydroPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ComponentHydroPipe.cls
'
'
Option Explicit
'
' MEMBER VARIABLES
' Member variables - Shared
Private pBase As ComponentHydro
'
' Member variables - Specific for a pipe type component
' (none)

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

' Properties - Specific for a pipe type component
Private Property Get NumberOfVolumes() As Integer
    NumberOfVolumes = Me.Segments(Me.Segments.Count).VolumeLast
End Property

'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set pBase = New ComponentHydro
End Sub

Public Sub Create(inputRows() As Variant, ByRef BaseComponent As ComponentHydro)
' Action: Contructor for a junction-type-component
'
' Inputs:   inputRows = A Variant array with inputs from the worksheet this type
'                       of component
'           BaseComponent = A reference to the base hydrodynamic component
'
    On Error GoTo ErrorHandler
    ' Set the Base component object
    Set pBase = BaseComponent
    Dim CurrPipeCompInd As Integer
    CurrPipeCompInd = HydroSystem.Components.Count
    
    
    Dim newPipe() As Variant

    ' Create one segment with values read into it
    Dim i As Integer, j As Integer, k As Integer
    Dim length As Double, Dx As Double, DxActual As Double, Volumes As Integer, VolumesTotal As Integer
    For i = 1 To UBound(inputRows, 1)
        pBase.Segments.Add New PropertiesHydroCompSegment
        
        ' Calculate number of volumes for pipe segment
        length = Inputdeck.CheckInput(inputRows(i, 3), Numeric:=True, valueDefault:=1, causeError:=True, FieldName:="Length", word:=3)
        Dx = Inputdeck.CheckInput(inputRows(i, 4), Numeric:=True, valueDefault:=0.2, causeError:=True, FieldName:="Dx", word:=4)
        DxActual = CalculateDx(length, Dx)
        Volumes = length / DxActual
        VolumesTotal = VolumesTotal + Volumes
        
        Dim PipeOverflow As Boolean
        If VolumesTotal > 99 Then
            PipeOverflow = True
            ReDim newPipe(1 To UBound(inputRows, 1) - i + 1, 1 To UBound(inputRows, 2))
            
            ' Copy values in inputRows to inputRowsNew
            For j = 1 To UBound(newPipe, 1)
                For k = 1 To UBound(newPipe, 2)
                    newPipe(j, k) = inputRows(i + j - 1, k)
                Next k
            Next j
            
            ' Change length and dx for first segment in inputRowsNew by subtracting
            ' length of the part of the segment that fitted in the pipe
            Dim VolumesOverFlow As Integer
            VolumesOverFlow = VolumesTotal - 99
            newPipe(1, 3) = VolumesOverFlow * DxActual
            newPipe(1, 4) = DxActual
            newPipe(1, 8) = inputRows(i, 8) ' K+
            newPipe(1, 9) = inputRows(i, 9) ' K-
            VolumesTotal = 99
            Volumes = Volumes - VolumesOverFlow
        End If
        
        With Me.Segments(i)
            .VolumeFirst = VolumesTotal - Volumes + 1
            .VolumeLast = VolumesTotal
            .Volumes = Volumes
            .JunctionFirst = .VolumeFirst
            ' Set last junction equal to last volume if not the last segment. Then subtract one
            If i = UBound(inputRows, 1) Or .VolumeLast = 99 Then
                .JunctionLast = .VolumeLast - 1
            Else
                .JunctionLast = .VolumeLast
            End If
            .Junctions = .JunctionLast - .JunctionFirst + 1
            .ID = inputRows(i, 2)
            .Dx = DxActual
            .Area = Inputdeck.CheckInput(inputRows(i, 5), Numeric:=True, valueDefault:=0, causeError:=True, FieldName:="Area", word:=5)
            .AngleVert = Inputdeck.CheckInput(inputRows(i, 6), Numeric:=True, valueDefault:=0, causeError:=True, FieldName:="Vertical angle", word:=6)
            .AngleAzi = Inputdeck.CheckInput(inputRows(i, 7), Numeric:=True, valueDefault:=0, causeError:=True, FieldName:="Vertical angle")
            If PipeOverflow = False Then ' If pipe overflow, transfer K+ / K- to newPipe
                .KLossPos = Inputdeck.CheckInput(inputRows(i, 8), Numeric:=True, valueDefault:=0, causeError:=True, FieldName:="K+", word:=8)
                .KLossNeg = Inputdeck.CheckInput(inputRows(i, 9), Numeric:=True, valueDefault:=0, causeError:=True, FieldName:="K+", word:=9)
            End If
            .Name = inputRows(i, 10)
            .Drawing = inputRows(i, 11)
            .Roughness = Inputdeck.CheckInput(inputRows(i, 17), Numeric:=True, valueDefault:=0, causeError:=True, FieldName:="Roughness", word:=17)
            .VFlags = Inputdeck.CheckInput(inputRows(i, 21), Numeric:=False, valueDefault:="0000000", causeError:=False, FieldName:="Volume flags", word:=21, IsBinaryString:=True)
            .JFlags = Inputdeck.CheckInput(inputRows(i, 22), Numeric:=False, valueDefault:="00000000", causeError:=False, FieldName:="Junction flags", word:=22, IsBinaryString:=True)
            Set .Init = pBase.Init
            
            .ForceNumber = Inputdeck.CheckInput(inputRows(i, 12), Numeric:=True, valueDefault:=0, causeError:=False, FieldName:="Force number", word:=12)
            If .ForceNumber <> 0 Then
                Dim ForceStr As String
                ForceStr = "Force" & CStr(.ForceNumber)
                HydroSystem.Force(ForceStr).ForceNumber = .ForceNumber
                HydroSystem.Force(ForceStr).ComponentIndexes.Add CurrPipeCompInd
                HydroSystem.Force(ForceStr).SegmentIndexes.Add i
            End If
        End With
        
        If PipeOverflow = True Then Exit For
        
    Next i
    
    ' If you have pipe overflow (more than 99 volumes):
    '   1. Create a connecting junction
    '   2. Create a new pipe component
    If PipeOverflow = True Then
        Dim newJunction() As Variant
        ReDim newJunction(1, 1 To UBound(inputRows, 2))
        Dim FromCode As String, ToCode As String
        
        pBase.Settings.IncrementCCC
        
        FromCode = Chr(34) & Format(Settings.CurrentCCC - Settings.CCCIncrement, "000") & Format(99, "00\0\0\0\2") & Chr(34)
        ToCode = Chr(34) & Format(Settings.CurrentCCC + Settings.CCCIncrement, "000") & Format(1, "00\0\0\0\1") & Chr(34)
        
        newJunction(1, 1) = "Junction"
        newJunction(1, 5) = 0#
        newJunction(1, 8) = 0#
        newJunction(1, 9) = 0#
        newJunction(1, 10) = inputRows(i, 10)
        newJunction(1, 13) = FromCode
        newJunction(1, 14) = ToCode
        newJunction(1, 15) = 2
        newJunction(1, 16) = 1
        newJunction(1, 21) = "tlpvbfe"
        newJunction(1, 22) = inputRows(i, 22)
        HydroSystem.Components.Add HydroComp, newJunction, pBase.RowBegin, pBase.RowEnd
        
        ' Add a pipe component
        HydroSystem.Components.Add HydroComp, newPipe, pBase.RowBegin, pBase.RowEnd
        Settings.IncrementCCC Decrement:=True
    End If
    
    Exit Sub
ErrorHandler:
    err.Raise err.Number, TypeName(Me) & ".Create, " & err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub
'
'
' CLASS METHODS
Public Sub WriteToFile(ByRef ts As Object, ByVal wC As Boolean, Optional ByVal GenerateIncludeFile As Boolean = False)
' Action: Writes the component to a file stored in the TextStream-object
'
' Inputs:   ts = A TextStream object (included in the MS Scripting Runtime)
'           wC = writeComments, True if comments are to be written, false otherwise
'
    On Error GoTo ErrorHandler:

    Dim SegCnt As Integer
    Dim i As Integer
    SegCnt = Me.Segments.Count

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
    t.AddLine "*"
    ts.WriteLine t.WriteAndClear
   
    ' CARD 1 -Number of volumes
    t.AddLine "*        Volumes"
    t.AddLine "{ }{  }       {}", pCCC, card.Update(1), t.Args(NumberOfVolumes, FStr:="%2d", Key:="VOLUMES")
    
    ' CARD 101-199 - Area
    t.AddLine "*** GEOMETRY ***"
    t.AddLine "*"
    t.AddLine "*           Area  Vol        Diameter"
    t.OutputKeywordIndexes = 1
    Dim ActualDiameter As Double
    For i = 1 To SegCnt
        With Me.Segments(i)
            ActualDiameter = Sqr(CDbl(t.GetFormattedValue(.Area, Area)) * 4 / 3.14159265) * 1000
            t.AddLine "{ }{  }  {}   {}  *   {} mm", pCCC, card.Update(100 + i), t.Args(.Area, Area), t.Args(.VolumeLast, Vols), t.Args(ActualDiameter, FStr:="%8.2f")
        End With
    Next i
    t.AddLine "*"
    
    ' CARD 301-399 - Dx
    t.AddLine "*               Dx  Vol     ID            Length  DWG"
    t.OutputKeywordIndexes = 1
    Dim ActualLength As Double
    For i = 1 To SegCnt
        With Me.Segments(i)
            ActualLength = CDbl(t.GetFormattedValue(.Dx, Dx)) * .Volumes
            t.AddLine "{ }{  }  {}   {}  *  {}  {} m  {}", pCCC, card.Update(300 + i), t.Args(.Dx, Dx), t.Args(.VolumeLast, Vols), t.Args(.ID, ID), _
                                                           t.Args(ActualLength, FStr:="%6.4f"), t.Args(.Drawing, Drawing)
        End With
    Next i
    t.AddLine "*"
    
    ' CARD 501-599 - Azimuthal angle
    t.AddLine "*        Azi.ang  Vol"
    t.OutputKeywordIndexes = 1
    For i = 1 To SegCnt
        With Me.Segments(i)
            t.AddLine "{ }{  }  {}   {}", pCCC, card.Update(500 + i), t.Args(.AngleAzi, AngleAzi), t.Args(.VolumeLast, Vols)
        End With
    Next i
    t.AddLine "*"
    
    ' CARD 601-699 - Vertical/inclination angle
    t.AddLine "*        Vrt.ang  Vol"
    t.OutputKeywordIndexes = 1
    For i = 1 To SegCnt
        With Me.Segments(i)
            t.AddLine "{ }{  }  {}   {}", pCCC, card.Update(600 + i), t.Args(.AngleVert, AngleVert), t.Args(.VolumeLast, Vols)
        End With
    Next i
    t.AddLine "*"
    
    ' CARD 801-899 - Pipe Volume X-Coordinate Friction Data
    t.AddLine "*        Roughness  Hyd.dia  Vol"
    t.OutputKeywordIndexes = 1
    For i = 1 To SegCnt
        With Me.Segments(i)
            t.AddLine "{ }{  }  {}   {}   {}", pCCC, card.Update(800 + i), t.Args(.Roughness, Roughness), t.Args(0, HydDiameter), t.Args(.VolumeLast, Vols)
        End With
    Next i
    t.AddLine "*"
    
    ' CARD 901-999 - Pipe Junction Loss Coefficients
    If NumberOfVolumes > 1 Then
        t.AddLine "*** LOSS COEFFICIENTS ***"
        t.AddLine "*           K+     K-  Junc"
        t.OutputKeywordIndexes = 1
        Dim tmpCnt As Integer
        tmpCnt = 1
        For i = 1 To Me.Segments.Count
            With Me.Segments(i)
                If .Junctions > 1 Then
                    t.AddLine "{ }{  }  {}  {}    {}", pCCC, card.Update(900 + tmpCnt), t.Args(0#, KLossPos), t.Args(0#, KLossNeg), t.Args(.JunctionLast - 1, Juncs)
                    tmpCnt = tmpCnt + 1
                    t.AddLine "{ }{  }  {}  {}    {}", pCCC, card.Update(900 + tmpCnt), t.Args(.KLossPos, KLossPos), t.Args(.KLossNeg, KLossNeg), t.Args(.JunctionLast, Juncs)
                    tmpCnt = tmpCnt + 1
                ElseIf .Junctions = 1 Then
                    t.AddLine "{ }{  }  {}  {}    {}", pCCC, card.Update(900 + tmpCnt), t.Args(.KLossPos, KLossPos), t.Args(.KLossNeg, KLossNeg), t.Args(.JunctionLast, Juncs)
                    tmpCnt = tmpCnt + 1
                ElseIf .Junctions = 0 Then
                    ' Write nothing
                End If
            End With
        Next i
        t.AddLine "*"
    End If
    
    ' CARD 1001-1099 - Pipe Volume X-Coordinate Control Flags
    t.AddLine "*** FLAGS ***"
    t.AddLine "*        tlpvbfe  Vol"
    t.OutputKeywordIndexes = 1
    For i = 1 To SegCnt
        With Me.Segments(i)
            t.AddLine "{ }{  }  {}   {}", pCCC, card.Update(1000 + i), t.Args(.VFlags, VFlags), t.Args(.VolumeLast, Vols)
        End With
    Next i
    t.AddLine "*"
    
    ' CARD 1101-1199 - Pipe Junction Control Flags
    If NumberOfVolumes > 1 Then
        t.AddLine "*        jefvcahs  Junc"
        t.OutputKeywordIndexes = 1
        For i = 1 To SegCnt
            With Me.Segments(i)
                If .Junctions > 0 Then
                    t.AddLine "{ }{  }  {}    {}", pCCC, card.Update(1100 + i), t.Args(.JFlags, JFlags), t.Args(.JunctionLast, Juncs)
                End If
            End With
        Next i
        t.AddLine "*"
    End If
    
    ' CARD 1201-1299 - Pipe Volume Initial Conditions
    t.AddLine "*** INITIAL CONDITIONS ***"
    For i = 1 To SegCnt
        With Me.Segments(i)
            Select Case Right(.Init.EbtFlag, 1)
                Case "1"
                    If i = 1 Then t.AddLine "*        ebt    Temp         x_s  W5  W6  W7  Vol": t.OutputKeywordIndexes = 1
                    t.AddLine "{ }{  }  {}  {}  {}  0.  0.  0.   {}", pCCC, card.Update(1200 + i), t.Args(.Init.EbtFlag, EbtFlag), t.Args(.Init.Temperature, InitTemperature), _
                                                                     t.Args(.Init.Quality, InitQuality), t.Args(.VolumeLast, Vols)
                Case "2"
                    If i = 1 Then t.AddLine "*        ebt   Pressure         x_s  W5  W6  W7  Vol": t.OutputKeywordIndexes = 1
                    t.AddLine "{ }{  }  {}  {}  {}  0.  0.  0.   {}", pCCC, card.Update(1200 + i), t.Args(.Init.EbtFlag, EbtFlag), t.Args(.Init.Pressure, InitPressure), _
                                                                     t.Args(.Init.Quality, InitQuality), t.Args(.VolumeLast, Vols)
                Case "3"
                    If i = 1 Then t.AddLine "*        ebt   Pressure    Temp  W5  W6  W7  Vol": t.OutputKeywordIndexes = 1
                    t.AddLine "{ }{  }  {}  {}  {}  0.  0.  0.   {}", pCCC, card.Update(1200 + i), t.Args(.Init.EbtFlag, EbtFlag), t.Args(.Init.Pressure, InitPressure), _
                                                                     t.Args(.Init.Temperature, InitTemperature), t.Args(.VolumeLast, Vols)
                Case "4"
                    If i = 1 Then t.AddLine "*        ebt   Pressure    Temp         x_s  W6  W7  Vol": t.OutputKeywordIndexes = 1
                    t.AddLine "{ }{  }  {}  {}  {}  {}  0.  0.  {}", pCCC, card.Update(1200 + i), t.Args(.Init.EbtFlag, EbtFlag), t.Args(.Init.Pressure, InitPressure), _
                                                                     t.Args(.Init.Temperature, InitTemperature), t.Args(.Init.Quality, InitQuality), t.Args(.VolumeLast, Vols)
                Case Else
            End Select
         End With
    Next i
    t.AddLine "*"
    
    ' CARD 1301-1399 - Pipe Junction Initial Conditions
    If NumberOfVolumes > 1 Then
        t.AddLine "*        cword  (1=mflow, 0=vel)"
        t.OutputKeywordIndexes = 0
        t.AddLine "{ }{  }      {}", pCCC, card.Update(1300), t.Args(Me.Segments(1).Init.CntrlWord, CWord)
        t.AddLine "*"
        t.AddLine "*        Liq.flow  Vap.flow  intf.vel  Junc"
        t.OutputKeywordIndexes = 1
        For i = 1 To Me.Segments.Count
            With Me.Segments(i)
                If .Junctions > 0 Then
                    t.AddLine "{ }{  }   {}   {}   {}    {}", pCCC, card.Update(1300 + i), t.Args(.Init.FlowLiq, InitFlowLiq), t.Args(.Init.FlowGas, InitFlowGas), t.Args(.Init.InterfVel, InitInterFVel), t.Args(.JunctionLast, Juncs)
                End If
            End With
        Next i
        t.AddLine "*"
    End If

    ts.WriteLine t.WriteAndClear
    
        

    
    
    
    Exit Sub
ErrorHandler:
    err.Raise err.Number, TypeName(Me) & ".WriteToFile, " & err.Source, err.Description, err.HelpFile, err.HelpContext
    
End Sub

'
'
' STATIC METHODS
Private Function CalculateDx(ByVal L As Double, ByVal Dx As Double) As Double
'***************************************************************************
'Purpose:   Calculates a actual element length dx preserving the length while
'           at the same time matches the desired dx
'
'Inputs:    Length = Length of the pipe segment
'           Dx = Desired element length
'
'Outputs:   A dx that preserves the total length
'***************************************************************************
    Dim actualDx, diffLargeDx, diffSmallDx, largeDx, smallDx

    If L <= Dx Then
        actualDx = L
    Else
        largeDx = L / Int(((L + 0.00001) / Dx))
        smallDx = L / Int(((L + 0.00001) / Dx) + 1)

        diffLargeDx = Abs(Dx - largeDx)
        diffSmallDx = Abs(Dx - smallDx)

        If diffSmallDx <= 0.5 * diffLargeDx Then
            actualDx = smallDx
        Else
            actualDx = largeDx
        End If
    End If

    CalculateDx = actualDx
End Function


