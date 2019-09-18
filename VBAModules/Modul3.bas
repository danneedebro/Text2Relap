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
Dim t As New ResourceTablePrint
    t.SetDefaultValues
    t.AddLine "*"
End Sub
