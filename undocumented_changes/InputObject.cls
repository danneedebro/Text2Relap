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
Public Property Get Sheet() As Worksheet
    Set Sheet = ThisWorkbook.Worksheets(pSheetname)
End Property

Public Sub SetData(arg As Variant)
    pData = arg
End Sub

Public Property Get Data(Optional FirstRow As Long = -1, Optional LastRow As Long = -1) As Variant
    Dim newData() As Variant
    Dim numberOfRows As Long
    FirstRow = IIf(FirstRow = -1, LBound(pData, 1), FirstRow)
    LastRow = IIf(LastRow = -1, UBound(pData, 1), LastRow)
    numberOfRows = LastRow - FirstRow + 1
    
    ReDim newData(1 To numberOfRows, LBound(pData, 2) To UBound(pData, 2))
    Dim i As Long, j As Long
    For i = FirstRow To LastRow
        For j = LBound(pData, 2) To UBound(pData, 2)
            newData(i - FirstRow + 1, j) = pData(i, j)
        Next j
    Next i
    
    Debug.Print vbNewLine & "newData is:"
    Dim rowString As String
    For i = LBound(newData, 1) To UBound(newData, 1)
        rowString = "Row " & i & ": "
        For j = LBound(newData, 2) To UBound(newData, 2)
            rowString = rowString & newData(i, j) & IIf(j < UBound(newData, 2), ", ", "")
        Next j
        Debug.Print rowString
    Next i
    
    
    Data = newData

End Property



'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    
End Sub

Public Sub Create(InputRange As Range)
    pSheetname = InputRange.Worksheet.Name
    pRow1 = InputRange.Row
    pRow2 = pRow1 + InputRange.Rows.Count - 1
    pData = InputRange
End Sub

Public Sub CreateFromParts(Data As Variant, SheetName As String, row1 As Long, row2 As Long)
    pSheetname = SheetName
    pRow1 = row1
    pRow2 = row2
    pData = Data
End Sub

Public Function Item(Row As Long, Col As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Item = pData(Row, Col)
End Function

Public Function Count() As Integer
    Count = UBound(pData, 1) - LBound(pData, 1) + 1
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
    
    Dim row1 As Long, row2 As Long
    row1 = IIf(FirstRow = -1, pRow1, pRow1 + FirstRow - 1)
    row2 = IIf(LastRow = -1, pRow2, pRow1 + LastRow - 1)
    
    If OnlyRowN <> -1 Then
        row1 = pRow1 + OnlyRowN - 1
        row2 = pRow1 + OnlyRowN - 1
    End If
    
    With copyOfSelf
        .CreateFromParts Data(FirstRow:=FirstRow, LastRow:=LastRow), pSheetname, row1, row2
    End With
    
    Set MakeCopy = copyOfSelf
    
End Function
'
'
' CLASS METHODS

'
'
' STATIC METHODS


