Attribute VB_Name = "Modul2"
Function Lista() As Variant

    Dim Result(1 To 3) As Variant
    
    Result(1) = "HEJ"
    Result(2) = "D�"
    Result(3) = "GE"



End Function


Sub Makro2()
Attribute Makro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro2 Makro
'

'
    ActiveSheet.Buttons.Add(658.5, 108, 123.75, 48).Select
    Selection.OnAction = "AddPipe"
End Sub
Sub Makro3()
Attribute Makro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro3 Makro
'

'
    ActiveSheet.Shapes.AddShape(msoShapeOval, 495, 392.6470866142, 25.5882677165, _
        23.8234645669).Select
    ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 295.5882677165, _
        397.0588188976, 495, 404.5588188976).Select
        Selection.ShapeRange.line.EndArrowheadStyle = msoArrowheadOpen
    Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes("Oval 1"), _
        3
End Sub
