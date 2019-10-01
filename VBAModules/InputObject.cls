VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Private pSheetname As String
Private pRow1 As Long
Private pRow2 As Long
Private pData As Variant
'
'
' PROPERTIES
Public Property Get RowFirst() As Long
    RowFirst = pRow1
End Property

Public Property Get RowLast() As Long
    RowLast = pRow2
End Property

Public Property Get SheetName() As String
    SheetName = pSheetname
End Property

Public Property Let SheetName(arg As String)
    pSheetname = arg
End Property

Public Property Let Replace(Row As Long, Col As Long, arg As Variant)
    pData(Row, Col) = arg
End Property

Public Property Get Sheet() As Worksheet
    Set Sheet = ThisWorkbook.Worksheets(pSheetname)
End Property

Public Sub SetData(arg As Variant)
    pData = arg
End Sub

Public Property Get Data(Optional FirstRow As Long = -1, Optional LastRow As Long = -1) As Variant
    Dim newData() As Variant
    Dim NumberOfRows As Long
    FirstRow = IIf(FirstRow = -1, LBound(pData, 1), FirstRow)
    LastRow = IIf(LastRow = -1, UBound(pData, 1), LastRow)
    NumberOfRows = LastRow - FirstRow + 1
    
    ReDim newData(1 To NumberOfRows, LBound(pData, 2) To UBound(pData, 2))
    Dim i As Long, j As Long
    For i = FirstRow To LastRow
        For j = LBound(pData, 2) To UBound(pData, 2)
            newData(i - FirstRow + 1, j) = pData(i, j)
        Next j
    Next i
    Data = newData
End Property



'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    ReDim pData(1 To 1, 1 To 22)
End Sub

Public Sub Create(inputRange As Range)
    pSheetname = inputRange.Worksheet.Name
    pRow1 = inputRange.Row
    pRow2 = pRow1 + inputRange.Rows.Count - 1
    pData = inputRange
End Sub

Public Sub CreateFromParts(SheetName As String, row1 As Long, row2 As Long, Optional Data As Variant)
    If IsMissing(Data) Then
        Dim emptyVariant(1 To 1, 1 To 22) As Variant
        pData = emptyVariant
    Else
        pData = Data
    End If
    pSheetname = SheetName
    pRow1 = row1
    pRow2 = row2
End Sub

Public Function item(Row As Long, Col As Long) As Variant
Attribute item.VB_UserMemId = 0
    item = pData(Row, Col)
End Function


Public Function NumberOfRows() As Integer
    NumberOfRows = UBound(pData, 1) - LBound(pData, 1) + 1
End Function

Public Function NumberOfColumns() As Integer
    NumberOfColumns = UBound(pData, 2) - LBound(pData, 2) + 1
End Function

Public Sub Highlight()
' Action: Highlights the selected range
    With Sheet
        .Range(.Cells(pRow1, 1), .Cells(pRow2, UBound(pData, 2))).Select
    End With
End Sub

Public Sub WriteData()
    Dim i As Long, j As Long
    Dim rowString As String
    For i = LBound(pData, 1) To UBound(pData, 1)
        rowString = "Row " & i & ": "
        For j = LBound(pData, 2) To UBound(pData, 2)
            rowString = rowString & pData(i, j) & IIf(j < UBound(pData, 2), ", ", "")
        Next j
        Debug.Print rowString
    Next i
End Sub

Public Function MakeCopy(Optional OnlyRowN As Long = -1, Optional FirstRow As Long = -1, Optional LastRow As Long = -1) As InputObject
' Action: Returns a copy of itself
'
    Dim copyOfSelf As New InputObject
    
    Dim i As Long, j As Long
    Dim sourceRow1 As Long, sourceRow2 As Long
    
    
    Dim newData As Variant
    
    If OnlyRowN <> -1 Then
        ReDim newData(1 To 1, 1 To NumberOfColumns)
        For i = LBound(pData, 2) To UBound(pData, 2)
            newData(1, i) = pData(OnlyRowN, i)
        Next i
    
        sourceRow1 = Me.RowFirst + OnlyRowN - 1
        sourceRow2 = Me.RowFirst + OnlyRowN - 1
    ElseIf FirstRow = -1 And LastRow = -1 Then
        ReDim newData(LBound(pData, 1) To UBound(pData, 1), LBound(pData, 2) To UBound(pData, 2))
        For i = LBound(pData, 1) To UBound(pData, 1)
            For j = LBound(pData, 2) To UBound(pData, 2)
                newData(i, j) = pData(i, j)
            Next j
        Next i
        sourceRow1 = Me.RowFirst
        sourceRow2 = Me.RowLast
    Else
        FirstRow = IIf(FirstRow = -1, 1, FirstRow)
        LastRow = IIf(LastRow = -1, UBound(pData, 1), LastRow)
        ReDim newData(1 To LastRow - FirstRow + 1, LBound(pData, 2) To UBound(pData, 2))
        For i = 1 To LastRow - FirstRow + 1
            For j = LBound(pData, 2) To UBound(pData, 2)
                newData(i, j) = pData(FirstRow + i - 1, j)
            Next j
        Next i
        sourceRow1 = Me.RowFirst + FirstRow - 1
        sourceRow2 = Me.RowFirst + LastRow - 1
    End If
    
    With copyOfSelf
        .CreateFromParts pSheetname, sourceRow1, sourceRow2, newData
    End With
    
    Set MakeCopy = copyOfSelf
    
End Function

Public Sub SetDataFromWords(Optional ByVal Word1 As Variant, _
                            Optional ByVal Word2 As Variant, Optional ByVal Word3 As Variant, _
                            Optional ByVal Word4 As Variant, Optional ByVal Word5 As Variant, Optional ByVal Word6 As Variant, _
                            Optional ByVal Word7 As Variant, Optional ByVal Word8 As Variant, Optional ByVal Word9 As Variant, _
                            Optional ByVal Word10 As Variant, Optional ByVal Word11 As Variant, Optional ByVal Word12 As Variant, _
                            Optional ByVal Word13 As Variant, Optional ByVal Word14 As Variant, Optional ByVal Word15 As Variant, _
                            Optional ByVal Word16 As Variant, Optional ByVal Word17 As Variant, Optional ByVal Word18 As Variant, _
                            Optional ByVal Word19 As Variant, Optional ByVal Word20 As Variant, Optional ByVal Word21 As Variant, _
                            Optional ByVal Word22 As Variant)
    
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
    
    SetData result
End Sub


