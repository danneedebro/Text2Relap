VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResourceTablePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Public Enum EnumDefaults
    None ' Must be at top
    Area
    Time
    CCC_Code
    CardNumberHydro
    C_Code
    AngleVert
    AngleAzi
    Dx
    KLossPos
    KLossNeg
    Roughness
    JFlags
    VFlags
    Pressure
    Temperature
    Quality
    Massflow
    InitVelocity
    InitMassflow
    InitFlowLiq
    InitFlowGas
    InitInterFVel
    InitPressure
    InitTemperature
    InitQuality
    Elevchange
    HydDiameter
    PumpFlow
    PumpHead
    PumpSpeed
    PumpTorque
    ValveOpen
    Density
    Diameter_mm
    EbtFlag
    Vols
    Juncs
    Volume
    ConnectionCode1
    ConnectionCode2
    Drawing
    Name_HydroComp
    CWord
    ID
End Enum

Private pLines As Collection

Private pFormatStr As String
Private pValue As Variant
Private pKeyword As String

Private Defaults() As Variant
Private FirstLine As Integer

Private pOutputKeywords As Integer
Private pOutputKeywordIndexes As Integer
Private pOutputKeywordBrackets As Integer
Private pOutputComments As Integer
'
'
' PROPERTIES
Public Property Get OutputComments() As Integer
    OutputComments = pOutputComments
End Property

Public Property Let OutputComments(arg As Integer)
    pOutputComments = arg
End Property

Public Property Get OutputKeywordBrackets() As Integer
    OutputKeywordBrackets = pOutputKeywordBrackets
End Property

Public Property Let OutputKeywordBrackets(arg As Integer)
    pOutputKeywordBrackets = arg
End Property

Public Property Get OutputKeywordIndexes() As Integer
    OutputKeywordIndexes = pOutputKeywordIndexes
End Property

Public Property Let OutputKeywordIndexes(arg As Integer)
    If arg = 1 Then FirstLine = pLines.Count
    pOutputKeywordIndexes = arg
End Property

Public Property Get OutputKeywords() As Integer
    OutputKeywords = pOutputKeywords
End Property

Public Property Let OutputKeywords(arg As Integer)
    pOutputKeywords = arg
End Property

Public Property Get Keyword() As String
    Keyword = pKeyword
End Property

Public Property Let Keyword(arg As String)
    pKeyword = arg
End Property

Public Property Get Value() As Variant
    Value = pValue
End Property

Public Property Let Value(arg As Variant)
    pValue = arg
End Property

Public Property Get FormatStr() As String
    FormatStr = pFormatStr
End Property

Public Property Let FormatStr(arg As String)
    pFormatStr = arg
End Property


'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    OutputKeywords = -1
    OutputKeywordIndexes = -1
    OutputComments = -1
    OutputKeywordBrackets = -1
    Set pLines = New Collection
    ReDim Defaults(0)
End Sub

Public Sub SetDefaultValues()
' Action: Set default keyword and formatstring for different properties
    ReDim Defaults(50)
    
    Defaults(0) = Array(EnumDefaults.Area, "AREA", "%7.3g")
    Defaults(1) = Array(EnumDefaults.Time, "TIME", "%8.4f")
    Defaults(2) = Array(EnumDefaults.CCC_Code, "CCC", "%03d")
    Defaults(3) = Array(EnumDefaults.C_Code, "C", "%3d")
    Defaults(4) = Array(EnumDefaults.AngleVert, "ANGLE2", "%7.3f")
    Defaults(5) = Array(EnumDefaults.AngleAzi, "ANGLE1", "%7.3f")
    Defaults(6) = Array(EnumDefaults.Dx, "DX", "%9.7f")
    Defaults(7) = Array(EnumDefaults.KLossPos, "KPOS", "%5.3f")
    Defaults(8) = Array(EnumDefaults.KLossNeg, "KNEG", "%5.3f")
    Defaults(9) = Array(EnumDefaults.Roughness, "ROUGHNESS", "%9.3e")
    Defaults(10) = Array(EnumDefaults.JFlags, "JEFVCAHS", "%8s")
    Defaults(11) = Array(EnumDefaults.VFlags, "TLPVBFE", "%7s")
    Defaults(12) = Array(EnumDefaults.Pressure, "PRESSURE", "%9.3e")
    Defaults(13) = Array(EnumDefaults.Temperature, "TEMPERATURE", "%6.2f")
    Defaults(14) = Array(EnumDefaults.Quality, "QUALITY", "%7.3g")
    Defaults(15) = Array(EnumDefaults.Massflow, "MASSFLOW", "%7.3f")
    Defaults(16) = Array(EnumDefaults.CWord, "CWORD", "%1d")
    Defaults(17) = Array(EnumDefaults.ID, "ID", "%-10s")
    Defaults(18) = Array(EnumDefaults.InitPressure, "INITPRESS", "%9.3e")
    Defaults(19) = Array(EnumDefaults.InitTemperature, "INITTEMP", "%6.2f")
    Defaults(20) = Array(EnumDefaults.InitQuality, "INITQUALITY", "%10.4e")
    Defaults(21) = Array(EnumDefaults.Elevchange, "ELEVCHANGE", "%7.5f")
    Defaults(22) = Array(EnumDefaults.HydDiameter, "HYD.DIA", "%6.4f")
    Defaults(23) = Array(EnumDefaults.PumpFlow, "PUMPFLOW", "%6.4f")
    Defaults(24) = Array(EnumDefaults.PumpHead, "PUMPHEAD", "%6.3f")
    Defaults(25) = Array(EnumDefaults.PumpSpeed, "PUMPRPM", "%6.2f")
    Defaults(26) = Array(EnumDefaults.PumpTorque, "PUMPTORQUE", "%6.3f")
    Defaults(27) = Array(EnumDefaults.ValveOpen, "VALVESTEM", "%5.4f")
    Defaults(28) = Array(EnumDefaults.Density, "DENSITY", "%7.2f")
    Defaults(29) = Array(EnumDefaults.Diameter_mm, "DIAMETER", "%6.2f")
    Defaults(30) = Array(EnumDefaults.EbtFlag, "EBTFLAG", "%3d")
    Defaults(31) = Array(EnumDefaults.Vols, "VOLS", "%2d")
    Defaults(32) = Array(EnumDefaults.Juncs, "JUNCS", "%2d")
    Defaults(33) = Array(EnumDefaults.Volume, "VOLUME", "%7.5f")
    Defaults(34) = Array(EnumDefaults.ConnectionCode1, "NNNVV000F", "%9s")
    Defaults(35) = Array(EnumDefaults.ConnectionCode2, "MMMVV000F", "%9s")
    Defaults(36) = Array(EnumDefaults.Drawing, "DRAWING", "%-12s")
    Defaults(37) = Array(EnumDefaults.CardNumberHydro, "", "%04d")
    Defaults(38) = Array(EnumDefaults.Name_HydroComp, "NAME", "%-8s")
    Defaults(39) = Array(EnumDefaults.InitFlowLiq, "INITFLIQ", "%7.3f")
    Defaults(40) = Array(EnumDefaults.InitFlowGas, "INITFGAS", "%7.3f")
    Defaults(41) = Array(EnumDefaults.InitInterFVel, "INITIVEL", "%7.3f")
    Defaults(42) = Array(EnumDefaults.None, "", "")
End Sub
'
'
' CLASS METHODS
Public Function WriteAndClear() As String
' Action: Returns content of pLines
'
    Dim OutputString As String
    
    Dim i As Integer
    For i = 1 To pLines.Count
        OutputString = OutputString & pLines(i) & IIf(i < pLines.Count, vbNewLine, "")
    Next i
    
    Set pLines = New Collection
    WriteAndClear = OutputString
    
End Function

Public Function AddLine(ByVal InputString As String, ParamArray Args() As Variant)
' Action: Replaces placeholders with values or keyword (using WriteLine function)
'         and stores it in pLines collection
'
    pLines.Add WriteLine(InputString, Args)
End Function

Public Function WriteLine(ByVal InputString As String, ParamArray Args() As Variant) As String
' Action: Replaces placeholders with values or keyword and returns resulting string
'
    Dim i As Integer, j As Integer
    Dim CurrChar As String
    Dim BracketCount As Integer
    Dim Start As Integer
    Dim CurrArg As ResourceTablePrint
    Dim s As New ResourceSprintf
    Dim IndirectArray As Boolean
    Dim LBnd As Integer, UBnd As Integer
    
    If IsMissing(Args) = True Then
        GoTo LastLine
    ElseIf VarType(Args) = vbArray + vbVariant Then
        If IsMissing(Args(0)) = True Then GoTo LastLine
        LBnd = LBound(Args(0))
        UBnd = UBound(Args(0))
        IndirectArray = True
    Else
        LBnd = LBound(Args)
        UBnd = UBound(Args)
        IndirectArray = False
    End If
    
    For i = LBnd To UBnd
        If IndirectArray = True Then
            If TypeName(Me) = TypeName(Args(0)(i)) Then
                Set CurrArg = Args(0)(i)
            Else
                Set CurrArg = New ResourceTablePrint
                CurrArg.Value = CStr(Args(0)(i))
            End If
        Else
            If TypeName(Me) = TypeName(Args(0)(i)) Then
                Set CurrArg = Args(i)
            Else
                Set CurrArg = New ResourceTablePrint
                CurrArg.Value = CStr(Args(i))
            End If
        End If
        BracketCount = 0
        Start = 0
        
        For j = 1 To Len(InputString)
            CurrChar = Mid(InputString, j, 1)
            If CurrChar = "{" Then
                BracketCount = BracketCount + 1
                Start = j
            End If
            If CurrChar = "}" And BracketCount = 1 Then
                Dim CurrKeyword As String
                Dim CurrValue As String
                Dim UseBrackets As Boolean, UseKeywordIndexes As Boolean, UseKeyword As Boolean
                
                If CurrArg.OutputKeywordBrackets <> -1 Then
                    UseBrackets = IIf(CurrArg.OutputKeywordBrackets = 1, True, False)
                ElseIf Me.OutputKeywordBrackets = -1 Then
                    UseBrackets = False  ' Default
                ElseIf Me.OutputKeywordBrackets <> -1 Then
                    UseBrackets = IIf(Me.OutputKeywordBrackets = 1, True, False)
                End If
                
                If CurrArg.OutputKeywordIndexes <> -1 Then
                    UseKeywordIndexes = IIf(CurrArg.OutputKeywordIndexes = 1, True, False)
                ElseIf Me.OutputKeywordIndexes = -1 Then
                    UseKeywordIndexes = False
                ElseIf Me.OutputKeywordIndexes <> -1 Then
                    UseKeywordIndexes = IIf(Me.OutputKeywordIndexes = 1, True, False)
                End If
                
                If CurrArg.OutputKeywords <> -1 Then
                    UseKeyword = IIf(CurrArg.OutputKeywords = 1, True, False)
                ElseIf Me.OutputKeywords = -1 Then
                    UseKeyword = False
                ElseIf Me.OutputKeywords <> -1 Then
                    UseKeyword = IIf(Me.OutputKeywords = 1, True, False)
                End If
                
                ' TODO: Override with CurrArg.Settings
                CurrKeyword = IIf(UseBrackets = True, "[", "") & CurrArg.Keyword & IIf(UseKeywordIndexes = True, "-" & CStr(pLines.Count - FirstLine + 1), "") & _
                              IIf(UseBrackets = True, "]", "")
                CurrValue = IIf(CurrArg.FormatStr = "", CStr(CurrArg.Value), s.sprintf(CurrArg.FormatStr, CurrArg.Value))
                InputString = Left(InputString, Start - 1) & _
                              IIf(UseKeyword = False Or CurrArg.Keyword = "", CurrValue, CurrKeyword) & _
                              Mid(InputString, j + 1)
                Exit For
            End If
            
        Next j
    Next i
    
LastLine:
    WriteLine = InputString
End Function

Public Function Args(ByVal Value As Variant, Optional OutputType As EnumDefaults = -1, Optional ByVal Key As String = "", _
                     Optional ByVal FStr As String = "", Optional ByVal OutputKeywordBrackets As Integer = -1, Optional ByVal OutputKeywords As Integer = -1, _
                     Optional ByVal OutputKeywordIndexes As Integer = -1, Optional ByVal OutputComments As Integer = -1) As ResourceTablePrint
    Dim NewArg As New ResourceTablePrint
    Dim i As Integer
    
    NewArg.Value = Value
    NewArg.OutputKeywordBrackets = OutputKeywordBrackets
    NewArg.OutputKeywords = OutputKeywords
    NewArg.OutputKeywordIndexes = OutputKeywordIndexes
    NewArg.OutputComments = OutputComments
    
    If UBound(Defaults) = 0 Then
        MsgBox "Default values not initialized. Use 'SetDefaults' to initiate them", vbCritical
        Stop
        Set Args = NewArg
        Exit Function
    End If
    
    ' If default values should be used
    If OutputType <> -1 Then
        For i = LBound(Defaults) To UBound(Defaults)
            If IsEmpty(Defaults(i)) = False Then
                If Defaults(i)(0) = OutputType Then
                    NewArg.Keyword = Defaults(i)(1)
                    NewArg.FormatStr = Defaults(i)(2)
                    Exit For
                End If
            End If
        Next i
    End If
    
    
    If Key <> "" Then NewArg.Keyword = Key
    If FStr <> "" Then NewArg.FormatStr = FStr
    
    Set Args = NewArg
End Function

Public Function Update(Optional ByVal Value As Variant, Optional OutputType As EnumDefaults, Optional ByVal Keyword As String = "", Optional ByVal FormatStr As String = "") As ResourceTablePrint
    If IsMissing(Value) = False Then Me.Value = Value
    Set Update = Me
End Function

Public Sub Settings(Optional ByVal OutputKeywordBrackets As Integer = -1, Optional ByVal OutputKeywords As Integer = -1, _
                    Optional ByVal OutputKeywordIndexes As Integer = -1, Optional ByVal OutputComments As Integer = -1)
    If OutputKeywordBrackets <> -1 Then Me.OutputKeywordBrackets = OutputKeywordBrackets
    If OutputKeywords <> -1 Then Me.OutputKeywords = OutputKeywords
    If OutputKeywordIndexes <> -1 Then Me.OutputKeywordIndexes = OutputKeywordIndexes
    If OutputComments <> -1 Then Me.OutputComments = OutputComments
    
    If OutputKeywordIndexes = 1 Then FirstLine = pLines.Count
End Sub

Public Function GetDefaultFormat(OutputType As EnumDefaults) As String
    
    Dim i As Integer
    Dim Result As String
    
    If UBound(Defaults) = 0 Then
        MsgBox "Default values not initialized. Use 'SetDefaults' to initiate them", vbCritical
        Stop
        GetDefaultFormat = ""
        Exit Function
    End If
    
    For i = LBound(Defaults) To UBound(Defaults)
        If IsEmpty(Defaults(i)) = False Then
            If Defaults(i)(0) = OutputType Then
                Result = Defaults(i)(2)
                Exit For
            End If
        End If
    Next i
    
    GetDefaultFormat = Result
    
End Function

Public Function GetFormattedValue(ByVal Value As Variant, ByVal OutputType As EnumDefaults) As String
    
    Dim i As Integer
    Dim Result As String
    Dim s As New ResourceSprintf
    
    If UBound(Defaults) = 0 Then
        MsgBox "Default values not initialized. Use 'SetDefaults' to initiate them", vbCritical
        Stop
        GetFormattedValue = CStr(Value)
        Exit Function
    End If
    
    For i = LBound(Defaults) To UBound(Defaults)
        If IsEmpty(Defaults(i)) = False Then
            If Defaults(i)(0) = OutputType Then
                Result = s.sprintf(Defaults(i)(2), Value)
                Exit For
            End If
        End If
    Next i
    
    GetFormattedValue = Result
    
End Function
'
'
' STATIC METHODS


