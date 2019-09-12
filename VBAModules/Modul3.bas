Attribute VB_Name = "Modul3"
Sub Makro5()
Attribute Makro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro5 Makro
'

'
    Range("F18").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="""HEJ"";""DÅ"";""APA"""
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub


Sub debuggg()
    Dim inputObj As New InputObject
    inputObj.Create Range("A1:V1000")
    'inputObj.CreateFromParts "324", 1, 15
    
    Dim inputObjChild As InputObject
    
    Set inputObjChild = inputObj.MakeCopy(OnlyRowN:=83)
    Set inputObjChild = inputObj.MakeCopy()
    Set inputObjChild = inputObj.MakeCopy(FirstRow:=83, LastRow:=84)
    Debug.Print inputObjChild.NumberOfColumns
    Debug.Print inputObjChild.NumberOfRows
    Debug.Print inputObjChild.RowFirst
    Debug.Print inputObjChild.RowLast
    
    
    
    
End Sub



Public Sub My_Split()

Dim z As Variant

z = Split(Replace(Join(Filter(Split(Replace(Replace(Selection.Value, ")", "^#"), "(", "#^"), "#"), "^"), "|"), "^", ""), "|")

Selection.Offset(0, 1).Resize(, UBound(z) + 1) = z

End Sub


