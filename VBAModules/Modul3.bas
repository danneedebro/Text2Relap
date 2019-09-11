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
    inputObj.Create Range("A50:B53")
    
    inputObj.MakeCopy(OnlyRowN:=3).Highlight
    Debug.Print inputObj.NumberOfColumns
    Debug.Print inputObj.NumberOfRows
    Debug.Print inputObj.RowFirst
    Debug.Print inputObj.RowLast
    
    Dim apa As HydroSystem
    
End Sub



Sub testinit()

    Dim apa As Variant
    Debug.Print "Hej"
    ReDim apa(1 To 1, 1 To 22)
    Debug.Print "Då"

End Sub
