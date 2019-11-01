Attribute VB_Name = "Text2RelapWriteStripRequest"
Option Explicit

Public Sub WriteStripRequest()
' Action: Writes a strip request file with optional decorators used by batch plot script THistPlot
'
    ' Select input sheet
    Dim shtInd As Integer
    shtInd = GetSheetIndex("Select sheet that contains Text2Relap input")
    If shtInd = -1 Then Exit Sub
    
    ' Read input sheet
    Dim InputDeck As Text2Relap
    Set InputDeck = NewInputdeck(Sheets(shtInd).Name, -1)
    If InputDeck.ReadOk = False Then Exit Sub
    
    
    ' Construct collections with the plotnums of the volumes, junctions and valves of the system.
    ' Later used
    Dim VolumesFirstAndLast As New Collection, VolumesFirst As New Collection, Junctions As New Collection, Valves As New Collection, Pumps As New Collection
    Dim Forces As New Collection
    Dim s As New ResourceSprintf
    
    Dim component As ComponentHydro
    For Each component In InputDeck.HydroSystem.Components.Subset(HydroComp)
        If component.Info.Family = JunctionComponent Then
            If component.Info.MainType = pump Then
                Pumps.Add CStr(component.CCC)
            ElseIf component.Info.MainType = valve Then
                Valves.Add CStr(component.CCC)
                Junctions.Add s.sprintf("%03d000000", component.CCC)
            Else
                Junctions.Add s.sprintf("%03d000000", component.CCC)
            End If
            
        ElseIf component.Info.Family = PipeComponent Then
            VolumesFirstAndLast.Add s.sprintf("%03d%02d0000", component.CCC, 1)
            VolumesFirstAndLast.Add s.sprintf("%03d%02d0000", component.CCC, component.Segments(component.Segments.Count).VolumeLast)
            VolumesFirst.Add s.sprintf("%03d%02d0000", component.CCC, 1)
            
            'Add forces
            Dim segment As PropertiesHydroCompSegment
            For Each segment In component.Segments
                If segment.ForceNumber > 0 And segment.ForceNumber <= 9999 Then
                    On Error Resume Next
                    Forces.Add segment.ForceNumber, Key:=CStr(segment.ForceNumber)
                    On Error GoTo 0
                    
                End If
            Next segment
            
        
        ElseIf component.Info.Family = SingleVolumeComponent Then
            VolumesFirstAndLast.Add s.sprintf("%03d010000", component.CCC)
            VolumesFirst.Add s.sprintf("%03d010000", component.CCC)
        End If
    Next component
   
    
    ' Read what plotvars to strip out from current sheet
    Dim sht As Worksheet
    Set sht = ThisWorkbook.ActiveSheet
    
    Dim InputRange As Range
    Set InputRange = sht.Range(sht.Cells(1, 1), sht.Cells(sht.Cells(sht.Rows.Count, "A").End(xlUp).Row, 5))
    
    Dim i As Integer
    Dim stripRequestCard As Long
    Dim plotnum As Variant
    stripRequestCard = 1000
    Dim ts As New ResourceTextStreamDummy
    
    ' Write header
    ts.WriteLine "=Stripfil"
    ts.WriteLine "100     strip fmtout"
    ts.WriteLine "0000103 0"
    
    For i = 1 To InputRange.Rows.Count
        
        ' Loop through all lines
        Select Case LCase(InputRange(i, 1))
            
            ' Inserts a strip request card
            Case "channels"
                Dim plotalf As String
                plotalf = LCase(InputRange(i, 2))
                Dim collectionToLoop As Collection
                Select Case plotalf
                    Case "mflowj", "velfj"
                        Set collectionToLoop = Junctions
                    Case "vlvstem"
                        Set collectionToLoop = Valves
                    Case "p"
                        Set collectionToLoop = VolumesFirstAndLast
                    Case "pmpvel"
                        Set collectionToLoop = Pumps
                    Case "forces"
                        plotalf = "cntrlvar"
                        Set collectionToLoop = Forces
                    Case Else
                        Set collectionToLoop = New Collection
                End Select
            
                For Each plotnum In collectionToLoop
                    stripRequestCard = stripRequestCard + 1
                    ts.WriteLine stripRequestCard & " " & plotalf & " " & plotnum
                Next plotnum
                
            ' Input for plot request file decorators used by THistPlot below
            
            ' A plot group
            Case "group"
                ts.WriteLine vbNewLine & "*<GROUP>"
            
            Case "plot"
                ts.WriteLine "*<PLOT>"
            
            ' XInterval XMin XMax
            Case "xint"
                ts.WriteLine s.sprintf("*XInt: %f %f", InputRange(i, 2), InputRange(i, 3))
                
            Case "yint"
                ts.WriteLine s.sprintf("*YInt: %f %f", InputRange(i, 2), InputRange(i, 3))
                
            Case "title"
                ts.WriteLine s.sprintf("*Title: %s", InputRange(i, 2))
                
            Case "ylabel"
                ts.WriteLine s.sprintf("*YLabel: %s", InputRange(i, 2))
                
            Case "xlabel"
                ts.WriteLine s.sprintf("*XLabel: %s", InputRange(i, 2))
                
            Case "yscale"
                ts.WriteLine s.sprintf("*YScale: %f", InputRange(i, 2))
                
            Case "yoffset"
                ts.WriteLine s.sprintf("*YOffset: %f", InputRange(i, 2))
                
            Case "xscale"
                ts.WriteLine s.sprintf("*XScale: %f", InputRange(i, 2))
                
            Case "xoffset"
                ts.WriteLine s.sprintf("*XOffset: %f", InputRange(i, 2))
                
            Case "yspanmin"
                ts.WriteLine s.sprintf("*YSpanMin: %f", InputRange(i, 2))
            
            Case "Curve"
                ts.WriteLine s.sprintf("*Curve: %s %s", InputRange(i, 2), InputRange(i, 3))
            
            Case "labeldefault"
                ts.WriteLine s.sprintf("*XYLabelDefaults: %s %s", InputRange(i, 2), InputRange(i, 3))
            
        End Select
    Next i
    
    ts.WriteLine ".end"
    
    
    ' Display the result
    UserForm1.TextBox1.Text = ts.TextStream
    UserForm1.Show
    
    
End Sub





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

