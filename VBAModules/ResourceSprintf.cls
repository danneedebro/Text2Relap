VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResourceSprintf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------ResourceSprintf.cls------------------------
' -----------------------------------------------------------------
'                Name:  sprintf (ResourceSprintf)
'             Project:  -
'                Type:  Class module
'              Author:  Daniel Edebro
' Last (sign.) update:  2019-02-09
'         Description:  Mimics the C++-sprintf function that formats numbers in strings
'
' Key Methods/Functions (*=public):
' + sprintf*                Main method mimicing the C++-sprintf function
'
' + GetFormatSpecProperty*  Fetches a certain format spec property (fieldwidth, precision etc)
'                           from a format spec string
'
' + FormatNum               Private function that converts a value using a Format spec UDT
'
' + GetFormatSpec           Returns a UDT with format specifications from
'                           a format spec string (%6.2f, %s, etc)
'
' -----------------------------------------------------------------
' USAGE:
' Dim s as New ResourceSprintf
' Debug.Print s.sprintf("Hello %s %03d","World", 2)
'
' -----------------------------------------------------------------
' REFERENCE
' Format  %{Identifier}${Flags}{Fieldwidth}.{Precision}
'
' Identifier: Order for processing the function input arguments
' Fieldwidth: Minimum number of characters to print
' Precision: for %f, %e: Number of digits to print
'            for %g: Number of significant digits
'            for %s: Numbers of letters to include
'
' Conversion characters
'   %d, %i      Integer
'   %u          Unsigned integer
'   %f          Float fixed point notation
'   %e %E       Exponential notation
'   %g %G       The more compact of %f and %e
'   %s          Strings
'   %o, %x, %X  Octal (base8), and Hex (base16) representation of number
'
' Flags
'   -           Left align
'   +           Always print a plus sign
'   <space>     Print a space for positive numbers
'   0           Pad with zeros
'   #           Modifier (for %f, %e print decimal point even if precision
'               is zero. For %o and %x, write prefix 0 or 0x
'
'
'
' -----------------------------------------------------------------
' NOTES:
' Custom pad chars not implemented %'.8.3f
'
Option Explicit
'
' MEMBER VARIABLES
Public Enum EnumFormatSpecProperties
    Invalid
    ConversionChar
    Fieldwidth
    Precision
    FlagLeftAlign
    FlagSign
    FlagLeadingZeros
    FlagHash
    FlagSpace
    Identifier
    PadChar
End Enum

Private Type TFormatSpec
    Invalid As Boolean
    ConversionChar As String
    Fieldwidth As Integer
    Precision As Integer
    FlagLeftAlign As Boolean
    FlagSign As Boolean
    FlagLeadingZeros As Boolean
    FlagHash As Boolean
    FlagSpace As Boolean
    Identifier As Integer
    PadChar As String
End Type

Private Const CONVERSION_CHARS = "diufeEgGsoxX"
Private Const CONVERSION_FLAGS = "-+ 0#"  ' - = Left-justify, + = always-print-plus-sign, " " = insert-space-instead-of-plus-sign
                                          ' 0 = pad-with-zeros, # = special
'
'
' PROPERTIES

'
'
' CLASS EVENT METHODS

'
'
' CLASS METHODS

'
'
' STATIC METHODS
Public Function sprintf(ByVal InputString As String, ParamArray Args() As Variant) As String
' Action: Public function that given a value (or string) formats it according to a
'         sprintf-type format specification string (%4.2f, %03d, etc)
    Dim i As Integer, j As Integer
    Dim StartIndex As Integer, StopIndex As Integer
    Dim cnt As Integer
    Dim FormatSpecStr As String, NewValue As String
    Dim FormatSpec As TFormatSpec
    Dim OutputString As String
    Dim CharTrigger As String, CharCur As String
    cnt = 0
    i = 0
    
    ' Copy Values in Args to Values because Args take on different forms dep on input:
    ' sprintf("Hello %s %d", "World", 2)  -> Args(0)="World", Args(1)=2
    ' sprintf("Hello %s %d")              -> Args = Missing
    ' Args=Array("Hello", 2)
    ' sprintf("Hello %s %d", Args)        -> Args(0)(0)="World", Args(0)(1)=2
    ' Args=Array()
    ' sprintf("Hello %s %d", Args)        -> Args(0) = Missing
    Dim Values() As Variant
    Dim NoArgs As Boolean
    
    If IsMissing(Args) Then
        NoArgs = True
    ElseIf VarType(Args(0)) = vbArray + vbVariant Then
        If IsMissing(Args(0)) = True Then
            NoArgs = True
        Else
            ReDim Values(UBound(Args(0)) - LBound(Args(0)))
            For i = 0 To UBound(Values)
                Values(i) = Args(0)(LBound(Args(0)) + i)
            Next i
        End If
    Else
        ReDim Values(UBound(Args))
        For i = 0 To UBound(Values)
            Values(i) = Args(i)
        Next i
    End If

    
    ' Loop through InputString and search for trigger char "%" (or \). When found, continue looping
    ' through InputString until a conversion character is found (d, f, s, e, etc). Send that string
    ' to GetFormatSpec method. If Formatspec is valid this with from the formatted string that FormatNum method returns
    i = 0
    Do
        i = i + 1
        CharCur = Mid(InputString, i, 1)
        If InStr(1, "%\", CharCur) > 0 Then
            CharTrigger = Mid(InputString, i, 1)
            For j = i + 1 To Len(InputString)
                CharCur = Mid(InputString, j, 1)
                
                ' If conversion character read ("d", "f", "e", "E", "s", etc)
                If CharTrigger = "%" And InStr(1, CONVERSION_CHARS, CharCur) > 0 And NoArgs = False Then
                    FormatSpecStr = Mid(InputString, i, j - i + 1)
                    
                    ' Parse the format spec using GetFormatSpec
                    FormatSpec = GetFormatSpec(FormatSpecStr)
                    
                    ' If format spec valid, format number according to it using FormatNum and
                    ' insert the formatted number inside InputString. Place loop-marker att the end
                    If FormatSpec.Invalid = False Then
                        If FormatSpec.Identifier > 0 And FormatSpec.Identifier - 1 <= UBound(Values) Then
                            NewValue = FormatNum(Values(FormatSpec.Identifier - 1), FormatSpec)
                        ElseIf FormatSpec.Identifier = 0 And cnt <= UBound(Values) Then
                            NewValue = FormatNum(Values(cnt), FormatSpec)
                            cnt = cnt + 1
                        Else
                            NewValue = FormatSpecStr
                        End If
                        InputString = Left(InputString, i - 1) & NewValue & Right(InputString, Len(InputString) - j)
                        i = i - 1 + Len(NewValue)
                        Exit For
                    Else: Exit For  ' If not a valid format spec, do nothing
                    End If
                    
                ' If "\n, "\b" etc
                ElseIf j = i + 1 And CharTrigger = "\" Then
                    Select Case CharCur
                        Case "n": NewValue = vbNewLine
                        Case "t": NewValue = vbTab
                        Case "b": NewValue = vbBack ' TODO: Maybe change this to actually remove last char?
                        Case "r": NewValue = vbCr  ' Not sure that this does what it should (Carrige return=return to col 1 of same row)
                        Case "\": NewValue = "\"
                        Case Else
                            NewValue = ""
                    End Select
                    InputString = Left(InputString, i - 1) & NewValue & Right(InputString, Len(InputString) - j)
                    i = i - 1 + Len(NewValue)
                    Exit For
                    
                ' If Escape sequence %% or \\
                ElseIf j = i + 1 And CharCur = CharTrigger Then ' Escape sequence %% or \\
                    InputString = Left(InputString, i - 1) & CharTrigger & Right(InputString, Len(InputString) - j)
                    i = i - 1 + Len(CharTrigger)
                    Exit For
                End If
            Next j
        End If
    Loop While i < Len(InputString)

    sprintf = InputString

End Function

Private Function GetFormatSpec(FormatSpec As String) As TFormatSpec
' Action: Returns a format spec as a custom type from a sprintf-type
'         format spec ("%4.2f", "%03d", etc)
'
' Input:
'         FormatSpec = String containing a single formatspec string (%03d etc)
'
' Outputs:
'         A UDT containing the properties of the format specification but easy to read
'
    Dim SubStr As String
    Dim tmpVal As Double
    Dim Parts() As String
    Dim Result As TFormatSpec
    
    
    With Result
        ' Set defaults
        .Precision = -1  ' Means depends
    
        ' Check if basic requirement for a format string exists (starts with % and ends with proper conv.char)
        If Left(FormatSpec, 1) <> "%" Or InStr(1, CONVERSION_CHARS, Right(FormatSpec, 1)) = 0 Then
            .Invalid = True
            GoTo TheEnd
        End If
        
        .ConversionChar = Right(FormatSpec, 1)  ' Know this is correct
        
        SubStr = Mid(FormatSpec, 2, Len(FormatSpec) - 2)
        
        If SubStr = "" Then GoTo TheEnd  ' if %f  or %d  etc
            
        Parts = Split(SubStr, "$")

        If UBound(Parts) > 0 Then  ' The identifier sign '$' is present
            If IsNumeric(Parts(0)) = True Then .Identifier = CInt(Parts(0))
            SubStr = Parts(1)
        Else
            SubStr = Parts(0)
        End If
        
        If SubStr = "" Then GoTo TheEnd  ' if %2$f  or %$1d  etc
        
        Parts = Split(SubStr, ".")  ' Split remaining part with '.'
        
        If UBound(Parts) > 0 Then ' Parts(0) contains flags if any, Parts(1) precision
            If IsNumeric(Parts(1)) Then
                .Precision = CInt(Parts(1))
            Else
                .Invalid = True
            End If
            SubStr = Parts(0)  ' Continue parsing the first part
        Else    ' Parts(0) contains flags, no precision given (ie -1)
            '.Precision = -1
        End If
        
        ' Parse part before precision operator '.' looking for flags
        Dim i As Integer
        For i = 1 To Len(SubStr)
            If InStr(1, CONVERSION_FLAGS, Mid(SubStr, i, 1)) > 0 Then
                Select Case Mid(SubStr, i, 1)
                    Case "-": .FlagLeftAlign = True
                    Case "+": .FlagSign = True
                    Case " ": .FlagSpace = True
                    Case "#": .FlagHash = True
                    Case "0": .FlagLeadingZeros = True
                End Select
                If i = Len(SubStr) Then SubStr = ""
            Else
                SubStr = Right(SubStr, Len(SubStr) - i + 1) ' Set save remaining to continue parsing fieldwidth
                Exit For
            End If
        Next i
        
        ' Check Fieldwidth part part between flags (+-0) and dot in %+-012.2f
        If SubStr = "" Then
            .Fieldwidth = 0
        ElseIf IsNumeric(SubStr) = True Then
            .Fieldwidth = CInt(SubStr)
        Else
            .Invalid = True
        End If
    End With

TheEnd:
    GetFormatSpec = Result
    
End Function

Private Function FormatNum(ByVal Value As Variant, ByRef FormatSpec As TFormatSpec) As String
' Action: Returns a string that is formatted according to the format specifications given by the
'         custom type 'TFormatSpec' (containing conversion character, fieldwidth, precision,
'         alignment- and other flags)
' Inputs:
'           Value = A number or a string
'           FormatSpec = A custom type containing information about the conversion parsed from
'                        a sprintf-type format specification string (%5.2f, %03d, etc)
' Outputs:
'           A formatted value (or string) according to specifications
'           <ERROR> if error
'
    Dim NumStr As String
    Dim FormatStr As String
    Dim SignPos As String, SignNeg As String
  
    With FormatSpec
        If .Invalid = True Then
            FormatNum = "<ERROR>"
            Exit Function
        End If
        
        ' If flag + and flag <space> is given it is assumed that a plus sign is wanted
        If .FlagSign = True Then
            SignPos = "+"
        ElseIf .FlagSpace = True Then
            SignPos = " "
        Else
            SignPos = ""
        End If
        SignNeg = "-"
        
        ' Set default precision when none is given
        Select Case .ConversionChar
            Case "f"
                If .Precision = -1 Then   ' if precision not specified, %f  or %5f
                    FormatStr = "0." & String(8, "#")
                Else
                    FormatStr = IIf(.Precision = 0, IIf(.FlagHash = True, "0.", "0"), "0.") & _
                                String(.Precision, "0")
                End If
                NumStr = Format(Value, SignPos & FormatStr & ";" & SignNeg & FormatStr)
            Case "e", "E"
                If .Precision = -1 Then
                    FormatStr = "0." & String(8, "#") & .ConversionChar & "+00"
                Else
                    FormatStr = IIf(.Precision = 0, IIf(.FlagHash = True, "0.", "0"), "0.") & _
                                String(.Precision, "0") & .ConversionChar & "+00"
                End If
                NumStr = Format(Value, SignPos & FormatStr & ";" & SignNeg & FormatStr)
            Case "g", "G"
                If .Precision = -1 Then  ' If precision not given, try which of %e or %f is shorter
                    Dim NumStr2 As String
                    ' %f
                    FormatStr = "0." & String(8, "#")
                    NumStr = Format(Value, SignPos & FormatStr & ";" & SignNeg & FormatStr)
                    ' %e
                    FormatStr = "0." & String(8, "#") & "e+00"
                    NumStr2 = Format(Value, SignPos & FormatStr & ";" & SignNeg & FormatStr)
                    
                    If Len(NumStr2) < Len(NumStr) Then NumStr = NumStr2
                Else  ' if precision is given then %f is shorter
                    If .Precision = 0 Then
                        FormatStr = "0e+00"
                    Else
                        FormatStr = "0." & String(.Precision - 1, "0") & "e+00"
                    End If
                    NumStr = Format(Value, SignPos & FormatStr & ";" & SignNeg & FormatStr)
                    
                    NumStr2 = NumberOfSignificantDigits(Value, .Precision, AlwaysPlotDecimalPoint:=.FlagHash, AlwaysPlotPlusSign:=.FlagSign)
                                        
                    If Len(NumStr2) < Len(NumStr) Then NumStr = NumStr2
                    
                End If
                
            Case "d", "i", "u"
                If .ConversionChar = "u" And IsNumeric(Value) Then Value = Abs(CInt(Value))
                If .Precision = -1 Then .Precision = 0
                FormatStr = IIf(.Precision = 0, "0", "0.") & String(.Precision, "0")
                NumStr = Format(Value, SignPos & FormatStr & ";" & SignNeg & FormatStr)
            Case "s"
                If .Precision <= 0 Then NumStr = CStr(Value) Else NumStr = Left(CStr(Value), .Precision)
            Case "o", "x", "X"
                If .ConversionChar = "o" And IsNumeric(Value) Then NumStr = Oct(CInt(Value))
                If .ConversionChar = "x" And IsNumeric(Value) Then NumStr = Format(Hex(CInt(Value)), "<")
                If .ConversionChar = "X" And IsNumeric(Value) Then NumStr = Format(Hex(CInt(Value)), ">")
        End Select
        
        ' Pad with characters if requested fieldwidth is > len(NumStr)
        If Len(NumStr) < .Fieldwidth Then
            If .PadChar = "" Then .PadChar = " " ' Catch error
            If .FlagLeadingZeros = True Then .PadChar = "0"
            
            ' if sign=true then place the +/- sign before the zeros
            If .FlagLeadingZeros = True And (Left(NumStr, 1) = "+" Or Left(NumStr, 1) = "-") Then
                NumStr = Left(NumStr, 1) & String(.Fieldwidth - Len(NumStr), "0") & Right(NumStr, Len(NumStr) - 1)
            ElseIf .FlagLeadingZeros = True And .FlagLeftAlign = False Then
                NumStr = String(.Fieldwidth - Len(NumStr), "0") & NumStr
            ElseIf .FlagLeadingZeros = True And .FlagLeftAlign = True Then
                NumStr = NumStr & String(.Fieldwidth - Len(NumStr), "0")
            ElseIf .FlagLeftAlign = True Then
                NumStr = NumStr & String(.Fieldwidth - Len(NumStr), .PadChar)
            Else
                NumStr = String(.Fieldwidth - Len(NumStr), .PadChar) & NumStr
            End If
        End If
    End With
    FormatNum = NumStr
End Function

Public Function GetFormatSpecProperty(ByVal FormatSpecString As String, FormatSpecProperty As EnumFormatSpecProperties) As Variant
' Action: Utility function that returns properties (custom types can't be returned in public methods)
'
    Dim FormatSpec As TFormatSpec
    FormatSpec = GetFormatSpec(FormatSpecString)
    
    With FormatSpec
        Select Case FormatSpecProperty
            Case Invalid:           GetFormatSpecProperty = .Invalid
            Case ConversionChar:    GetFormatSpecProperty = .ConversionChar
            Case Fieldwidth:        GetFormatSpecProperty = .Fieldwidth
            Case Precision:         GetFormatSpecProperty = .Precision
            Case FlagLeftAlign:     GetFormatSpecProperty = .FlagLeftAlign
            Case FlagSign:          GetFormatSpecProperty = .FlagSign
            Case FlagLeadingZeros:  GetFormatSpecProperty = .FlagLeadingZeros
            Case FlagHash:          GetFormatSpecProperty = .FlagHash
            Case FlagSpace:         GetFormatSpecProperty = .FlagSpace
            Case Identifier:        GetFormatSpecProperty = .Identifier
            Case PadChar:           GetFormatSpecProperty = .PadChar
            Case Else:              GetFormatSpecProperty = "Unknown"
        End Select
    End With
End Function

Function NumberOfSignificantDigits(ByVal Value As Variant, ByVal SignificantDigits As Integer, _
                                   Optional ByVal AlwaysPlotDecimalPoint As Boolean = False, _
                                   Optional ByVal AlwaysPlotPlusSign As Boolean = False) As String
' Action: Returns a string with Value formatted with N significant digits
'
    Dim Exponent As Integer, Coeff As String
    Dim NumStr As String

    ' If value not numeric, return Value as string
    If IsNumeric(Value) = False Then
        NumberOfSignificantDigits = CStr(Value)
        Exit Function
    ElseIf CDbl(Value) = 0 Then
        NumberOfSignificantDigits = "0." & String(SignificantDigits - 1, "0")
        Exit Function
    End If
    
    NumStr = Format(CDbl(Value), "+" & String(SignificantDigits, "0") & "e+00;-" & String(SignificantDigits, "0") & "e+00")
    
    ' Parse out exponent and coeff and pad coeff with zeros  coeff*10^exponent
    Exponent = CInt(Right(NumStr, 3))
    Coeff = Left(NumStr, SignificantDigits + 1)
    
    ' If exponent negative, shift decimal point as many steps as the exponent
    If Exponent < 0 Then
        Coeff = Format(Coeff, String(Abs(Exponent) + 1, "0"))
        NumStr = Left(Coeff, Len(Coeff) + Exponent) & "." & Right(Coeff, Abs(Exponent))
    ' If exponent positive, add as many zeros as the exponent
    ElseIf Exponent > 0 Then
        NumStr = IIf(AlwaysPlotPlusSign = True, Coeff, Format(Coeff)) & String(Exponent, "0") & IIf(AlwaysPlotDecimalPoint = True, ".", "")
    ' If exponent = 0 then do nothing (or add decimal point)
    Else
        NumStr = IIf(AlwaysPlotPlusSign = True, Coeff, Format(Coeff)) & IIf(AlwaysPlotDecimalPoint = True, ".", "")
    End If
    
    NumberOfSignificantDigits = NumStr
End Function
'
'
' UNIT TESTING


