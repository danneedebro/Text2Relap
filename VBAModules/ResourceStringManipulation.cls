VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResourceStringManipulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ResourceStringManipulation.cls
'
'
Option Explicit




Public Function ParseBrackedExpression(ByVal InputString As String, Numeric As Boolean) As Variant
' Action: Reads a string on format [t1;y1;t2;y2,...,tN;yN] and returns it as
'         an array
'
' Inputs:
'           InputString = The string to be parsed
'           Numeric = True if input is supposed to be numeric
'
' Output:
'           Binar
    On Error GoTo ErrorHandler
    Dim WrongFormat As Boolean, NotAMultipleOfTwo As Boolean, NoBrackets As Boolean
    Dim TimeVectorNotIncresing As Boolean, YValueGreaterThanOne As Boolean
    
    Dim output(0 To 2) As Variant
    Dim XValues() As Variant
    Dim YValues() As Variant
    
    Dim i As Integer
    Dim nRows As Integer
    
    
    ' Check if string is contained within brackets
    If Left(InputString, 1) = "[" And Right(InputString, 1) = "]" Then
        Dim strArr() As String
        strArr = Split(Mid(InputString, 2, Len(InputString) - 2), ";")
        
        ' If
        If (UBound(strArr) + 1) Mod 2 <> 0 Then NotAMultipleOfTwo = True
        nRows = Int((UBound(strArr) + 1) / 2)
        ReDim XValues(nRows - 1)
        ReDim YValues(nRows - 1)

        For i = 0 To UBound(XValues)
            XValues(i) = strArr(i * 2)
            YValues(i) = strArr(i * 2 + 1)
            If Numeric = True And IsNumeric(XValues(i)) = False Then WrongFormat = True
            If Numeric = True And IsNumeric(YValues(i)) = False Then WrongFormat = True
            If Numeric = True And IsNumeric(YValues(i)) = True Then
                If CDbl(YValues(i)) > 1 Then YValueGreaterThanOne = True
            End If
            If i > 0 Then
                If XValues(i) <= XValues(i - 1) Then TimeVectorNotIncresing = True
            End If
        Next i
    Else
        NoBrackets = True
        ReDim XValues(0)
        ReDim YValues(0)
        If Numeric = True And IsNumeric(InputString) = False Then WrongFormat = True
        XValues(0) = 0
        YValues(0) = InputString
    End If
    
    ' If output supposed to be numeric convert to double
    If Numeric = True Then
        For i = LBound(XValues) To UBound(XValues)
            If IsNumeric(XValues(i)) = True Then
                XValues(i) = CDbl(XValues(i))
            Else
                XValues(i) = -99
            End If
            If IsNumeric(YValues(i)) = True Then
                YValues(i) = CDbl(YValues(i))
            Else
                YValues(i) = -99
            End If
        Next i
    End If
    
    
    ' Set error / info codes
    output(0) = 0
    If NoBrackets = True Then output(0) = output(0) + BinToInt("1")  ' Not a bracked expression
    If NotAMultipleOfTwo = True Then output(0) = output(0) + BinToInt("10")  ' Not a multiple of two
    If WrongFormat = True Then output(0) = output(0) + BinToInt("100")  ' Wrong format (string if supposed to be numeric)
    If YValueGreaterThanOne = True Then output(0) = output(0) + BinToInt("1000")  ' Wrong format (string if supposed to be numeric)
    If TimeVectorNotIncresing = True Then output(0) = output(0) + BinToInt("10000")  ' Time vector not increasing
    
    output(1) = XValues
    output(2) = YValues
    ParseBrackedExpression = output

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".ParseBrackedExpression(""" & InputString & """), " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function MergeTwoSets(Time1 As Variant, Y1 As Variant, Time2 As Variant, Y2 As Variant) As Variant
' Action: Merges two sets of arrays
'
' Inputs:
'           Time1 = Time vector 1
'           Y1 = y-Value vector 1 (same size as Time1)
'           Time2 = Time vector 2
'           Y2 = y-Value vector 2 (same size as Time2)
'
' Output:
'           An array with N rows and 3 columns
'
    Dim NewArray() As Variant
    
    ReDim NewArray(UBound(Time1) + UBound(Time2) + 1, 2)
    
    Dim i As Integer, j As Integer, k As Integer, ind As Integer
    Dim N As Integer, SetID As Integer
    Dim CurrTime As Double, CurrY As Double
    
    Dim CurrSetTime As Variant, CurrSetY As Variant
    Dim OtherSetTime As Variant, OtherSetY As Variant
    Dim ColumnKnown As Integer, ColumnUnknown As Integer
    
    ' Outer loop Set
    For SetID = 1 To 2
    
        If SetID = 1 Then
            ColumnKnown = 1
            ColumnUnknown = 2
            CurrSetTime = Time1
            CurrSetY = Y1
            OtherSetTime = Time2
            OtherSetY = Y2
        Else
            ColumnKnown = 2
            ColumnUnknown = 1
            CurrSetTime = Time2
            CurrSetY = Y2
            OtherSetTime = Time1
            OtherSetY = Y1
        End If
    
        ' Loop through Set and fill in both known values and interpolate unknown values
        ' from the other set
        For i = 0 To UBound(CurrSetTime)
            ind = (SetID - 1) * (UBound(OtherSetTime) + 1) + i
            CurrTime = CurrSetTime(i)
            CurrY = CurrSetY(i)
            NewArray(ind, 0) = CurrTime
            NewArray(ind, ColumnKnown) = CurrY
            
            If CurrTime <= OtherSetTime(0) Then
                NewArray(ind, ColumnUnknown) = OtherSetY(0)
            ElseIf CurrTime >= OtherSetTime(UBound(OtherSetTime)) Then
                NewArray(ind, ColumnUnknown) = OtherSetY(UBound(OtherSetY))
            Else
                ' Interpolate in values from other set
                For j = 0 To UBound(OtherSetTime) - 1
                    If CurrTime < OtherSetTime(j + 1) And CurrTime >= OtherSetTime(j) Then
                        NewArray(ind, ColumnUnknown) = OtherSetY(j) + (OtherSetY(j + 1) - OtherSetY(j)) * (CurrTime - OtherSetTime(j)) / (OtherSetTime(j + 1) - OtherSetTime(j))
                        Exit For
                    End If
                Next j
            End If
        Next i
    
    Next SetID
   
    Dim tmpVal As Double
    Dim IsSorted As Boolean
    
    Dim cnt As Integer
    
    ' Sort list
    Do
        cnt = cnt + 1
        IsSorted = True
        For i = 0 To UBound(NewArray) - 1
            If NewArray(i, 0) > NewArray(i + 1, 0) Then
                For j = 0 To UBound(NewArray, 2)
                    tmpVal = NewArray(i, j)
                    NewArray(i, j) = NewArray(i + 1, j)
                    NewArray(i + 1, j) = tmpVal
                Next j
                IsSorted = False
            End If
        Next i
    Loop While IsSorted = False
    
    MergeTwoSets = NewArray

End Function

Public Function IntToBin(ByVal Number As Integer, Optional ByVal NumberOfBits As Integer = 1, Optional Prefix As Boolean = False) As String
' Action: Converts number to binary string representation
'
    Dim BinaryString As String
    BinaryString = String(NumberOfBits, "0")
    
    Dim CurrentBit As Integer
    CurrentBit = 1
    
    Dim TmpNum As Integer
    Do
        TmpNum = Number Mod 2
        If CurrentBit > NumberOfBits Then
            BinaryString = CStr(TmpNum) & BinaryString
        Else
            BinaryString = Left(BinaryString, NumberOfBits - CurrentBit) & CStr(TmpNum) & Right(BinaryString, CurrentBit - 1)
        End If
        CurrentBit = CurrentBit + 1
        Number = Int(Number / 2)
    Loop Until Number = 0
    IntToBin = IIf(Prefix = True, "0b", "") & BinaryString
End Function

Public Function BinToInt(ByVal BinaryNumber As String) As Integer
' Action: Converts binary string to integer
'
    Dim i As Integer
    Dim Exponent As Integer, coefficient As Integer
    For i = Len(BinaryNumber) To 1 Step -1
        Exponent = Len(BinaryNumber) - i
        coefficient = CInt(Mid(BinaryNumber, i, 1))
        BinToInt = BinToInt + coefficient * 2 ^ Exponent
    Next i
End Function

Public Function CheckBitSet(ByVal Number As Integer, ByVal BitToCheck As Integer) As Boolean
' Action: Checks if bit is set
'
    Dim BinaryNumber As String
    BinaryNumber = IntToBin(Number, BitToCheck)
    If Left(Right(BinaryNumber, BitToCheck), 1) = 1 Then
        CheckBitSet = True
    Else
        CheckBitSet = False
    End If
    
End Function
