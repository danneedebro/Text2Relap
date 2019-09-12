Attribute VB_Name = "Text2RelapWriteStripRequest"
Option Explicit

Private Function GetSheetIndex(Optional Question As String = "Select sheet") As Integer
' Action: Prompts the user to select the current sheet and returns the sheet index
'
' Example:
'    Dim shtInd As Integer
'    shtInd = GetSheetIndex()
'    If shtInd = -1 Then
'       Exit Sub
'    else
'       msgbox Sheets(shtInd).Name
'    End if
'
    Dim i As Integer
    Dim ws As Worksheet
    Dim shtInd As Integer
    Dim answ
    Question = Question & vbNewLine & vbNewLine

    For i = 1 To Worksheets.Count
        Question = Question & i & " = '" & Worksheets(i).Name & "'" & vbNewLine
    Next i
    
SelectSheet:
    Dim tmp As String
    tmp = InputBox(Question, "Select sheet")
    If tmp = "" Then   ' Cancel is pressed
        GetSheetIndex = -1
        Exit Function
    ElseIf IsNumeric(tmp) = True Then
        If CInt(tmp) > Worksheets.Count Then
            answ = MsgBox("Select a value between 1 and " & Worksheets.Count, vbCritical + vbOKCancel)
            If answ = vbCancel Then GetSheetIndex = -1 Else GoTo SelectSheet
        Else
            GetSheetIndex = CInt(tmp)
        End If
    Else
        answ = MsgBox("Input a NUMERIC value between 1 and " & Worksheets.Count, vbCritical + vbOKCancel)
            If answ = vbCancel Then GetSheetIndex = -1 Else GoTo SelectSheet
    End If
    
End Function


Sub WriteStripRequestFile()
' Action: Generates a strip request file
'
    Dim i As Integer, j As Integer
    
    Dim shtInd As Integer
    shtInd = GetSheetIndex()
    If shtInd = -1 Then Exit Sub
    
    Dim InputDeck As Text2Relap
    Set InputDeck = NewInputdeck(Worksheets(shtInd).Name)
    Debug.Print Worksheets(shtInd).Name
    Debug.Print InputDeck.HydroSystem.Components.Count
    
    Dim junctionCodes As New Collection
    Dim volumeCodes As New Collection
    Dim hydroCompCurr As ComponentHydro
    Dim s As New ResourceSprintf
    
    With InputDeck.HydroSystem
        For i = 1 To .Components.Count
            'Debug.Print Inputdeck.HydroSystem.Components(i).ComponentInfo
            'Debug.Print "   " & Inputdeck.HydroSystem.Components(i).ObjectType
            If .Components(i).ObjectType = HydroComp Then
                Set hydroCompCurr = .Components(i)
            
                With hydroCompCurr
                    If .Info.Family = JunctionComponent Then
                        junctionCodes.Add s.sprintf("%03d%02d0000", .CCC, 0)
                    ElseIf .Info.Family = PipeComponent Then
                        volumeCodes.Add s.sprintf("%03d%02d0000", .CCC, 1)
                    ElseIf .Info.Family = SingleVolumeComponent Then
                        volumeCodes.Add s.sprintf("%03d%02d0000", .CCC, 1)
                    End If
                End With
            End If
        Next i
    End With
    
    
    Dim stripRequestGroups As Range
    Set stripRequestGroups = Range("A5:D28")
    Dim cardInd As Integer
    cardInd = 1000
    Dim plotvarCurr As String
 
    For i = 1 To stripRequestGroups.Rows.Count
        plotvarCurr = stripRequestGroups(i, 2)
        Select Case plotvarCurr
            Case "mflowj", "vlvstem"
                For j = 1 To junctionCodes.Count
                    Debug.Print s.sprintf("%04d  %8s  %s", cardInd, plotvarCurr, junctionCodes(j))
                    cardInd = cardInd + 1
                Next j
            Case "p", "voidg", "voidf"
                For j = 1 To volumeCodes.Count
                    Debug.Print s.sprintf("%04d  %8s  %s", cardInd, plotvarCurr, volumeCodes(j))
                    cardInd = cardInd + 1
                Next j
        End Select
    Next i


End Sub
