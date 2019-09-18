VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComponentHydroSingleVol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ComponentHydroSingleVol.cls
'
'
Option Explicit
'
' MEMBER VARIABLES
' Member variables - Shared
Private pBase As ComponentHydro
'
' Member variables - Specific for a single volume type component

Private Type TmdpvolInput
    Time1() As Variant
    Pressure() As Variant
    Time2() As Variant
    Temperature() As Variant
    TimePressTemp() As Variant
    EbtFlag As String
End Type

Private pTmdpvol As TmdpvolInput
'
'
' PROPERTIES
' Properties - Shared for all hydro components
Public Property Get InputDeck() As Text2Relap
    Set InputDeck = pBase.InputDeck
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
' Properties - Specific for a single volume type component
' (none)
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set pBase = New ComponentHydro
End Sub

Public Sub Create(inputRows() As Variant, ByRef BaseComponent As ComponentHydro)
' Action: Contructor for a singlevolume-type-component
'
' Inputs:   inputRows = A Variant array with inputs from the worksheet this type
'                       of component
'           BaseComponent = A reference to the base hydrodynamic component
'
    On Error GoTo ErrorHandler
    ' Set the Base component object
    Set pBase = BaseComponent
    
    ' Create one segment with values read into it
    pBase.Segments.Add New PropertiesHydroCompSegment
    
    With Me.Segments(1)
        .VolumeFirst = 1
        .VolumeLast = 1
        .Volumes = 1
        .ID = inputRows(1, 2)
        .Dx = InputDeck.CheckInput(inputRows(1, 3), Numeric:=True, valueDefault:=1#, causeError:=True, FieldName:="Length", word:=3)
        .Area = InputDeck.CheckInput(inputRows(1, 5), Numeric:=True, valueDefault:=0#, causeError:=True, FieldName:="Area", word:=5)
        .AngleVert = InputDeck.CheckInput(inputRows(1, 6), Numeric:=True, valueDefault:=0#, causeError:=True, FieldName:="Vertical angle", word:=6)
        .AngleAzi = InputDeck.CheckInput(inputRows(1, 7), Numeric:=True, valueDefault:=0#, causeError:=True, FieldName:="Azimuthal angle", word:=7)
        .Name = inputRows(1, 10)
        .Drawing = inputRows(1, 11)
        .Elevchange = .Dx * Sin(.AngleVert * 3.14159265 / 180)
        .VFlags = InputDeck.CheckInput(inputRows(1, 21), Numeric:=False, valueDefault:="0000000", causeError:=False, FieldName:="Volume flags", word:=21, IsBinaryString:=True)
        .JFlags = InputDeck.CheckInput(inputRows(1, 22), Numeric:=False, valueDefault:="00000000", causeError:=False, FieldName:="Junction flags", word:=22, IsBinaryString:=True)
        Set .Init = pBase.Init
    End With
    
    ' Read specific input for single volume type components
    Dim ErrorMessages1 As String
    Dim ErrorMessages2 As String
    Dim s As New ResourceStringManipulation
    Dim Word17 As Variant
    Dim Word18 As Variant
    Dim answ
    If Me.Info.MainType = tmdpvol Then
        Word17 = s.ParseBrackedExpression(inputRows(1, 17), True)
        Word18 = s.ParseBrackedExpression(inputRows(1, 18), True)
        ' If second bit is set input is not a multiple of two
        If s.CheckBitSet(Word17(0), 2) = True Then ErrorMessages1 = ErrorMessages1 & "Input Error: Input is Not a multiple of two." & vbNewLine & vbNewLine
        If s.CheckBitSet(Word17(0), 3) = True Then ErrorMessages1 = ErrorMessages1 & "Input Error: Wrong format." & vbNewLine & vbNewLine
        If s.CheckBitSet(Word17(0), 5) = True Then ErrorMessages1 = ErrorMessages1 & "Input Error: Time vector is not increasing" & vbNewLine & vbNewLine
        
        If s.CheckBitSet(Word18(0), 2) = True Then ErrorMessages2 = ErrorMessages2 & "Input Error: Input is Not a multiple of two." & vbNewLine & vbNewLine
        If s.CheckBitSet(Word18(0), 3) = True Then ErrorMessages2 = ErrorMessages2 & "Input Error: Wrong format." & vbNewLine & vbNewLine
        If s.CheckBitSet(Word18(0), 5) = True Then ErrorMessages2 = ErrorMessages2 & "Input Error: Time vector is not increasing" & vbNewLine & vbNewLine
    End If
    
    
    ReDim pTmdpvol.Time1(0)
    ReDim pTmdpvol.Pressure(0)
    ReDim pTmdpvol.Time2(0)
    ReDim pTmdpvol.Temperature(0)
    If Me.Info.MainType = tmdpvol Then
        Dim CombinedPTArray As Variant
        pTmdpvol.TimePressTemp = s.MergeTwoSets(Word17(1), Word17(2), Word18(1), Word18(2))
    End If
    
    ' If input errors occured
    If ErrorMessages1 <> "" Then
        InputDeck.HighlightCurrentLines 17
        answ = MsgBox(ErrorMessages1 & vbNewLine & "Continue?", vbExclamation + vbOKCancel)
        If answ = vbCancel Then
            InputDeck.Abort = True
            Exit Sub
        End If
    End If
    
    ' If input errors occured
    If ErrorMessages2 <> "" Then
        InputDeck.HighlightCurrentLines 18
        answ = MsgBox(ErrorMessages2 & vbNewLine & "Continue?", vbExclamation + vbOKCancel)
        InputDeck.Warnings = True
        If answ = vbCancel Then
            InputDeck.Abort = True
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
    Dim s As New ResourceSprintf
    Dim t As New ResourceTablePrint
    t.SetDefaultValues
    Dim Replacements(1 To 5) As Variant
    Dim i As Integer
    
    ' Replace keywords with properties specific to junction type components
    With Me.Segments(1)
        Replacements(1) = Array("[NNNVV000F]", "NNNVV000F", Me.Connections(1).ConnectionCode)
        Replacements(2) = Array("[MMMVV000F]", "MMMVV000F", Me.Connections(2).ConnectionCode)
    End With
        
    If Me.Info.MainType = tmdpvol Then
        
        
        ' Replace time-data for tmdpvol
        For i = 0 To UBound(pTmdpvol.TimePressTemp)
            InputString = Replace(InputString, s.sprintf("[TIME-%d]", i + 1), t.GetFormattedValue(pTmdpvol.TimePressTemp(i, 0), Time))
            InputString = Replace(InputString, s.sprintf("[PRESSURE-%d]", i + 1), t.GetFormattedValue(pTmdpvol.TimePressTemp(i, 1), Pressure))
            InputString = Replace(InputString, s.sprintf("[TEMPERATURE-%d]", i + 1), t.GetFormattedValue(pTmdpvol.TimePressTemp(i, 2), Temperature))
        Next i

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

End Function

'
'
' CLASS METHODS
Public Sub WriteToFile(ByRef ts As Object, ByVal wC As Boolean, Optional ByVal GenInc As Boolean = False)
' Action: Writes the component to a file stored in the TextStream-object
'
' Inputs:   ts = A TextStream object (included in the MS Scripting Runtime)
'           wC = writeComments, True if comments are to be written, false otherwise
'           GenInc = Write default include file (send arg to fixNum to for instance write out CCC instead of componentnumber
'
    On Error GoTo ErrorHandler

    If Me.Filename <> "" And GenInc = False Then
        WriteCustom ts, wC
        Exit Sub
    End If
    
    Dim t As New ResourceTablePrint
    t.SetDefaultValues
    t.OutputKeywords = IIf(GenInc = True, 1, 0)
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
    
    With Segments(1)
        ' Card 101-102 - Single-Volume X-Coordinate Volume Data
        t.AddLine "*           Area     Length   Volume  Azi.angle  Vert.angle  Elev.ch"
        t.AddLine "{ }{  }  {     }  {       }  {     }    {       }     {        }  {     }", pCCC, card.Update(101), t.Args(.Area, Area), t.Args(.Dx, Dx), _
                                                                                          t.Args(.VolumeM2, Volume), t.Args(.AngleAzi, AngleAzi), _
                                                                                          t.Args(.AngleVert, AngleVert), t.Args(.Elevchange, Elevchange)
        t.AddLine "*"
        t.AddLine "*          Roughn.  Hyd.dia  tlpvbfe"
        t.AddLine "{ }{  }  {       }   {     }  {     }", pCCC, card.Update(102), t.Args(.Roughness, Roughness), t.Args(0#, HydDiameter), t.Args(.VFlags, VFlags)
        t.AddLine "*"
        
        ' SINGLE VOLUME SPECIFIC INPUT
        If Me.Info.MainType = snglvol Then
            
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
            ts.WriteLine t.WriteAndClear ' Write to file
        
        ' TIME-DEPENDANT VOLUME SPECIFIC INPUT
        ElseIf Me.Info.MainType = tmdpvol Then
            
            ' Merge time-dependant data into a common vector
            'Dim s As New ResourceStringManipulation
            'Dim CombinedPTArray As Variant
            'CombinedPTArray = s.MergeTwoSets(pTmdpvol.Time1, pTmdpvol.Pressure, pTmdpvol.Time2, pTmdpvol.Temperature)
            
            ' If 0<=P(t)<= 1  ebt=101,  If 0<=T(t)<=1  ebt=102
            Dim i As Integer
            Dim PressureMax As Double, PressureMin As Double, TemperatureMax As Double, TemperatureMin As Double
            For i = LBound(pTmdpvol.TimePressTemp, 1) To UBound(pTmdpvol.TimePressTemp, 1)
                If pTmdpvol.TimePressTemp(i, 1) < PressureMin Or i = LBound(pTmdpvol.TimePressTemp, 1) Then PressureMin = pTmdpvol.TimePressTemp(i, 1)
                If pTmdpvol.TimePressTemp(i, 1) > PressureMax Or i = LBound(pTmdpvol.TimePressTemp, 1) Then PressureMax = pTmdpvol.TimePressTemp(i, 1)
                If pTmdpvol.TimePressTemp(i, 2) < TemperatureMin Or i = LBound(pTmdpvol.TimePressTemp, 1) Then TemperatureMin = pTmdpvol.TimePressTemp(i, 2)
                If pTmdpvol.TimePressTemp(i, 2) > TemperatureMax Or i = LBound(pTmdpvol.TimePressTemp, 1) Then TemperatureMax = pTmdpvol.TimePressTemp(i, 2)
            Next i
            If Right(.Init.EbtFlag, 1) = "4" Then
                ' Do nothing
            ElseIf PressureMin >= 0 And PressureMax <= 1 Then
                .Init.EbtFlag = Left(.Init.EbtFlag, 2) & "1"
            ElseIf TemperatureMin >= 0 And TemperatureMax <= 1 Then
                .Init.EbtFlag = Left(.Init.EbtFlag, 2) & "2"
            Else
                .Init.EbtFlag = Left(.Init.EbtFlag, 2) & "3"
            End If
        
            t.AddLine "*        ebt  tbl.trip  cntrlvar  num"
            t.AddLine "{ }{  }  { }", pCCC, card.Update(200), t.Args(.Init.EbtFlag, EbtFlag)
            t.AddLine "*"
            
            ' Time-dependant data
            Select Case Right(.Init.EbtFlag, 1)
                Case "1"
                    t.AddLine "*            Time   Temp        x_s"
                    t.OutputKeywordIndexes = 1
                    For i = LBound(pTmdpvol.TimePressTemp) To UBound(pTmdpvol.TimePressTemp)
                        t.AddLine "{ }{  }  {}  {}  {}", pCCC, card.Update(201 + i), t.Args(pTmdpvol.TimePressTemp(i, 0), Time), t.Args(pTmdpvol.TimePressTemp(i, 1), Pressure), _
                                                         t.Args(pTmdpvol.TimePressTemp(i, 2), Temperature)
                    Next i
                Case "2"
                    t.AddLine "*            Time   Pressure    x_s"
                    t.OutputKeywordIndexes = 1
                    For i = LBound(pTmdpvol.TimePressTemp) To UBound(pTmdpvol.TimePressTemp)
                        t.AddLine "{ }{  }  {}  {}  {}", pCCC, card.Update(201 + i), t.Args(pTmdpvol.TimePressTemp(i, 0), Time), t.Args(pTmdpvol.TimePressTemp(i, 1), Pressure), _
                                                         t.Args(pTmdpvol.TimePressTemp(i, 2), InitQuality)
                    Next i
                Case "3"
                    t.AddLine "*            Time   Pressure    Temp"
                    t.OutputKeywordIndexes = 1
                    For i = LBound(pTmdpvol.TimePressTemp) To UBound(pTmdpvol.TimePressTemp)
                        t.AddLine "{ }{  }  {}  {}  {}", pCCC, card.Update(201 + i), t.Args(pTmdpvol.TimePressTemp(i, 0), Time), t.Args(pTmdpvol.TimePressTemp(i, 1), Pressure), _
                                                         t.Args(pTmdpvol.TimePressTemp(i, 2), Temperature)
                    Next i
                Case "4"
                    t.AddLine "*            Time   Pressure    Temp   x_s"
                    t.OutputKeywordIndexes = 1
                    For i = LBound(pTmdpvol.TimePressTemp) To UBound(pTmdpvol.TimePressTemp)
                        t.AddLine "{ }{  }  {}  {}  {}  {}", pCCC, card.Update(201 + i), t.Args(pTmdpvol.TimePressTemp(i, 0), Time), t.Args(pTmdpvol.TimePressTemp(i, 1), Pressure), _
                                                         t.Args(pTmdpvol.TimePressTemp(i, 2), Temperature), t.Args(.Init.Quality, InitQuality, OutputKeywordIndexes:=0)
                    Next i
                Case Else
            End Select
            
            ts.WriteLine t.WriteAndClear ' Write to file
        End If
        
    End With
    
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".WriteValve, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
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
        answ = MsgBox("Warning: Custom hydrodynamic component '" & Includefile.FullPath & "' doesn't exist. A generic " & CompType & "-type includefile will be created", vbExclamation + vbOKCancel, "Include file not found")
        InputDeck.Warnings = True
        If answ = vbCancel Then
            InputDeck.Abort = True
            Exit Sub
        End If
        If Includefile.FolderExists = False Then Includefile.CreateFolders ' Create subfolders if they don't exist
        Set tsWrite = fso.CreateTextFile(Includefile.FullPath, True)
        Me.WriteToFile tsWrite, wC, GenInc:=True
        tsWrite.Close
    End If
    
    ' Open include file for reading
    Set tsRead = fso.OpenTextFile(Includefile.FullPath, ForReading)
   
    ' Read all lines of the file and write them to file
    Do While tsRead.AtEndOfStream <> True
        ts.WriteLine tsRead.ReadLine
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

