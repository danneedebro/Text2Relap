Attribute VB_Name = "Text2RelapMain"
Option Explicit

Private Const TESTMATRIX = "TestMatrix"
Private Const TESTMATRIX_SHEET = "runMatrix"
Private Const TESTMATRIX_CURRENT_SET = "CurrentIndex"
Private Const TESTMATRIX_CURRENT_FILEPATH = "CurrentFilename"
Private Const TESTMATRIX_FILENAMES = "TestMatrixFilenames"
Private Const TESTMATRIX_VARIABLE_NAMES = "TestMatrixVariableNames"
Private Const TESTMATRIX_VARIABLE_VALUES = "TestMatrixVariableValues"

Private Type TFileWriteStatus
    Message As String
    FileWritten As Boolean
    Warnings As Boolean
    Abort As Boolean
    FullPath As String
    RelativePath As String
    Filename As String
End Type

Function NewInputdeck(Optional ByVal InputWorksheet As String = "", Optional ByVal LastRow As Integer = -1) As Text2Relap
' Action: Factory function for a Text2Relap input deck
'
    Dim Inputdeck As New Text2Relap
    Inputdeck.Create InputWorksheet, LastRow
    Set NewInputdeck = Inputdeck
End Function

Sub CreateInputFile()
' Action: Create file
'
    Dim CurrFilename As String
    Dim FileWriteStatus As TFileWriteStatus
    CurrFilename = Range(TESTMATRIX_CURRENT_FILEPATH).Value
    FileWriteStatus = ReadInputAndWriteToFile(CurrFilename)
    With FileWriteStatus
        If .FileWritten = True Then
            If .Warnings = True Then
                MsgBox "Warning: File """ & .Filename & """ created with warnings. Review input" & vbNewLine & vbNewLine & "Full path: " & .FullPath, vbExclamation
            Else
                MsgBox "Info: File """ & .Filename & """ created successfully" & vbNewLine & vbNewLine & "Full path: " & .FullPath, vbInformation
            End If
        Else
            MsgBox "Info: Errors during reading of """ & .Filename & """. No file created", vbExclamation
        End If
    End With
End Sub

Sub CreateInputFiles()
' Action: Create multiple files
'
    Dim LoadCase1 As Variant, LoadCase2 As Variant
    Dim i As Integer
    Dim CurrFilename As String
    Dim FileWriteStatus() As TFileWriteStatus
    
    LoadCase1 = InputBox(Prompt:="Från case with index:", Title:="Start", Default:="1")
    If IsNumeric(LoadCase1) = False Or LoadCase1 = vbNullString Then
        Exit Sub
    Else
        LoadCase1 = CInt(LoadCase1)
    End If

    LoadCase2 = InputBox(Prompt:="...case:", Title:="Start", Default:="1")
    If IsNumeric(LoadCase2) = False Or LoadCase2 = vbNullString Then
        Exit Sub
    Else
        LoadCase2 = CInt(LoadCase2)
    End If


    ReDim FileWriteStatus(LoadCase1 To LoadCase2)
    For i = LoadCase1 To LoadCase2
        Range("CurrentIndex").Value = i
        CurrFilename = Range("CurrentFilename").Value
        FileWriteStatus(i) = ReadInputAndWriteToFile(CurrFilename)
        If FileWriteStatus(i).FileWritten = False Then
            Dim answ
            answ = MsgBox("Abort?", vbQuestion + vbYesNo)
            If answ = vbYes Then Exit For
        End If
    Next i
    
    Dim ResultString As String
    For i = LoadCase1 To LoadCase2
        ResultString = ResultString & "Case index: " & CStr(i) & vbNewLine & FileWriteStatus(i).Message & vbNewLine & vbNewLine
    Next i
    MsgBox ResultString
End Sub


Function ReadInputAndWriteToFile(ByVal Filename As String) As TFileWriteStatus
' Action: Creates a input file
'
    Dim Result As TFileWriteStatus
    
    Dim Inputdeck As Text2Relap
    Set Inputdeck = NewInputdeck(ActiveSheet.Name)
    Dim InputFile As New ResourceFileObject
    
    Result.Filename = Filename
    
    If Inputdeck.ReadOk = True Then
        If InStr(1, Filename, ":") > 0 Then   ' If Full path supplied
            InputFile.CreateByParts Filename
        Else ' If relative path
            InputFile.CreateByParts ThisWorkbook.Path, Filename
        End If
        
        If InputFile.FolderExists = False Then
            InputFile.CreateFolders
        End If
    
        Result.RelativePath = InputFile.getRelativePath(ThisWorkbook.Path)
        Result.FullPath = InputFile.FullPath
        Result.Filename = InputFile.Filename
    
        If Inputdeck.WriteToFile(InputFile.FullPath) = True Then
            Result.FileWritten = True
            If Inputdeck.Warnings = False Then
                Result.Message = "Info: " & InputFile.FullPath & " created successfully"
                Result.Warnings = False
            Else
                Result.Message = "Warning: " & InputFile.FullPath & " created with warnings"
                Result.Warnings = True
            End If
        Else
            Result.FileWritten = False
            Result.Warnings = True
            Result.Message = "Error: " & InputFile.FullPath & " NOT created successfully"
        End If
    Else
        Result.FileWritten = False
        Result.Warnings = True
        Result.Abort = False
    End If
    
    ReadInputAndWriteToFile = Result
    
End Function


' TODO:
    ' Fundera på Heat structures
    ' TMDPVOL WriteToFile - TidsDubletter i när man sammanfogat Pressure och Temperature
    ' Belys external connections-junctions, dvs sådana som inte kopplar till angränsande volymer
    ' Merge
    

Sub ResetCondFormatting()
' Action: Resets the format conditions for the sheet
'
    Dim Inputdeck As Text2Relap
    Set Inputdeck = NewInputdeck(ActiveSheet.Name)
    Inputdeck.ResetFormatConditions
End Sub

'Sub MakeStripFile()
' Action: Creates a strip request file from input
'
'    Const MAKESTRIPFILE_WORKSHEET = "Tools"
'    Const MAKESTRIPFILE_INPUT = "A5:D30"
'
'    Dim InputRange As Range
'    Dim i As Integer, j As Integer, k As Integer
'    Dim cnt As Integer
'    Set InputRange = Worksheets(MAKESTRIPFILE_WORKSHEET).Range(MAKESTRIPFILE_INPUT)'
'
'    cnt = 1000
'    Dim Inputdeck As New Text2Relap
'    If Inputdeck.ReadOk = True Then
'
'        For i = 1 To InputRange.Rows.Count
'            With Inputdeck.HydroSystem
'                Select Case InputRange(i, 2)
'                    ' Junction quantities
'                    Case "mflowj", "voidfj", "voidgj", "velfj"
'                        For j = 1 To .Components.Count
'                            With .Component(j)
'                                If .Info.Family = JunctionComponent Then
'                                    cnt = cnt + 1
'                                    Debug.Print CStr(cnt) & "  " & "mflowj" & "  " & Format(.CCC, "000") & "000000"
'                                End If
'                            End With
'                        Next j
'                    ' Volume quantities
'                    Case "p", "voidg", "voidf"
'                        For j = 1 To .Components.Count
'                            With .Component(j)
'                                If .Info.Family = PipeComponent Then
'                                    cnt = cnt + 1
'                                    Debug.Print CStr(cnt) & "  " & "p" & "  " & Format(.CCC, "000") & "010000"
'                                    Debug.Print CStr(cnt) & "  " & "p" & "  " & Format(.CCC, "000") & Format(.Segment(.Segments.Count).VolumeLast, "00\0\0\0\0")
'                                ElseIf .Info.Family = SingleVolumeComponent Then
'                                    cnt = cnt + 1
'                                    Debug.Print CStr(cnt) & "  " & "p" & "  " & Format(.CCC, "000") & "010000"
'                                End If
'                            End With
'                        Next j
                    ' Forces
'                    Case "force"
'                        For j = 1 To .Forces.Count
'                            cnt = cnt + 1
'                            Debug.Print CStr(cnt) & "  " & "cntrlvar" & "  " & CStr(.Force(Index:=j).ForceNumber)
'                        Next j
'
'                    Case Else
                        
'                End Select
'            End With
'        Next i
'    End If
'End Sub

Sub ProbeOutput()
Attribute ProbeOutput.VB_ProcData.VB_Invoke_Func = "I\n14"
' Action: Probe Output by selecting cells with rows that contain a plotted output
'
    Dim SelectedRange As Range
    Dim Row1 As Integer, Row2 As Integer
    
    If TypeName(Selection) = "Range" Then
        Set SelectedRange = Selection
        With SelectedRange
            Row1 = .Rows(1).row
            Row2 = .Rows(.Rows.Count).row
        End With

        Dim Inputdeck As Text2Relap
        Set Inputdeck = NewInputdeck(ActiveSheet.Name)
        If Inputdeck.ReadOk = True Then
            Inputdeck.ProbeInput Row1, Row2, 0
            Inputdeck.ProbeInput Row1, Row2, 3
        Else
            MsgBox "Failed to read"
        End If
    End If

End Sub

Sub ModelSummary()
' Action: Writes a model summary
'
    Dim Inputdeck As Text2Relap
    Set Inputdeck = NewInputdeck(ActiveSheet.Name, -1)
    If Inputdeck.ReadOk = True Then
        Inputdeck.ModelSummary
    End If
End Sub

Private Sub TurnOffScreenUpdate(Optional TurnOff As Boolean = True)
    If TurnOff = False Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    Else
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        Application.EnableEvents = False
    End If
End Sub

Sub AddPipe()
' Action: Adds one or more pipe segments at the rows of the selected cells
'
'
    Dim CurrRow As Integer, currRowCnt As Integer, Word1 As String
    Dim question
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select one or more cells where you want to add new pipe segments", vbExclamation, "Insert pipe segment"
        Exit Sub
    End If
    
    CurrRow = Selection.row
    currRowCnt = Selection.Rows.Count
    
    Word1 = Cells(CurrRow, 1)
    
    TurnOffScreenUpdate True
    
    If Word1 = "Pipe" Then
        question = MsgBox("Insert " & CStr(currRowCnt) & " pipe segments BELOW row " + CStr(CurrRow) + " with the same properties as '" & _
                          Cells(CurrRow, 2) & "'?", vbYesNoCancel, "Insert pipe segment")
        If question <> vbYes Then
            Exit Sub
        End If
        Rows(CStr(CurrRow) & ":" & CStr(CurrRow + currRowCnt - 1)).Select
        Selection.Insert Shift:=xlUp, copyorigin:=xlFormatFromLeftOrAbove
        Rows(CurrRow + currRowCnt).Select
        Selection.Copy
        Rows(CStr(CurrRow) & ":" & CStr(CurrRow + currRowCnt - 1)).Select
        ActiveSheet.Paste
        Rows(CStr(CurrRow + 1) & ":" & CStr(CurrRow + currRowCnt - 1 + 1)).Select
    Else
        question = MsgBox("Insert " & CStr(currRowCnt) & " pipe segments ON row " + CStr(CurrRow) + " ?. ", vbYesNoCancel, "Insert pipe segments")
        If question <> vbYes Then
            Exit Sub
        End If
        
        Range(Cells(CurrRow, 1), Cells(CurrRow + currRowCnt - 1, 1)) = "Pipe"
        Range(Cells(CurrRow, 2), Cells(CurrRow + currRowCnt - 1, 2)).Formula = "=CONCATENATE(""PIPE_"",ROW())"
        Range(Cells(CurrRow, 4), Cells(CurrRow + currRowCnt - 1, 4)).Formula = "=dx"
        Range(Cells(CurrRow, 7), Cells(CurrRow + currRowCnt - 1, 9)) = 0#
        Range(Cells(CurrRow, 10), Cells(CurrRow + currRowCnt - 1, 10)) = "Pipe"
        Range(Cells(CurrRow, 11), Cells(CurrRow + currRowCnt - 1, 16)) = "-"
        Range(Cells(CurrRow, 17), Cells(CurrRow + currRowCnt - 1, 17)).Formula = "=roughness"
        Range(Cells(CurrRow, 18), Cells(CurrRow + currRowCnt - 1, 22)) = "-"
    End If

    TurnOffScreenUpdate False

End Sub




Sub AddJunction()
' Action: Adds a single junction at selected row
'
'
    Dim CurrRow As Integer, currRowCnt As Integer, Word1 As String
    Dim question
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell or a row where you want to add a new SNGLJUN", vbExclamation, "Insert single junction"
        Exit Sub
    End If
    
    CurrRow = Selection.row
    
    
    Word1 = Cells(CurrRow, 1)
    
    If Word1 = "Pipe" Then
        question = MsgBox("Insert a new sngljun BELOW row " + CStr(CurrRow) + "?", vbYesNoCancel, "Insert junction")
        If question <> vbYes Then
            Exit Sub
        End If
        Rows(CurrRow + 1 & ":" & CurrRow + 3).Select
        Selection.Insert Shift:=xlUp, copyorigin:=xlFormatFromLeftOrAbove
        CurrRow = CurrRow + 2
        Range(Cells(CurrRow, 1), Cells(CurrRow, 1)) = "Junction"
        Range(Cells(CurrRow, 2), Cells(CurrRow, 2)).Formula = "=CONCATENATE(""JUNC_"",ROW())"
        Range(Cells(CurrRow, 3), Cells(CurrRow, 4)) = "-"
        Range(Cells(CurrRow, 5), Cells(CurrRow, 5)) = 0#      ' Area = 0 för inre junction
        Range(Cells(CurrRow, 6), Cells(CurrRow, 7)) = "-"     '
        Range(Cells(CurrRow, 8), Cells(CurrRow, 9)) = 0#      ' K+  K-
        Range(Cells(CurrRow, 10), Cells(CurrRow, 10)) = "junction"      ' Namn
        Range(Cells(CurrRow, 11), Cells(CurrRow, 11)) = Cells(CurrRow - 2, 11)   ' Ritning   (samma som pipe för inre junction)
        Range(Cells(CurrRow, 12), Cells(CurrRow, 12)) = "-"       ' Kraftnr
        Range(Cells(CurrRow, 13), Cells(CurrRow, 13)).Formula = "=OFFSET($A$1,ROW()-3,1)"
        Range(Cells(CurrRow, 14), Cells(CurrRow, 14)).Formula = "=OFFSET($A$1,ROW()+1,1)"
        Range(Cells(CurrRow, 15), Cells(CurrRow, 15)) = 2
        Range(Cells(CurrRow, 16), Cells(CurrRow, 16)) = 1
        Range(Cells(CurrRow, 17), Cells(CurrRow, 22)) = "-"
    Else
        question = MsgBox("Insert a new sngljun ON row " + CStr(CurrRow) + "?", vbYesNoCancel, "Insert junction")
        If question <> vbYes Then
            Exit Sub
        End If
        Range(Cells(CurrRow, 1), Cells(CurrRow, 1)) = "Junction"
        Range(Cells(CurrRow, 2), Cells(CurrRow, 2)).Formula = "=CONCATENATE(""JUNC_"",ROW())"
        Range(Cells(CurrRow, 3), Cells(CurrRow, 4)) = "-"
        Range(Cells(CurrRow, 5), Cells(CurrRow, 5)) = 0#      ' Area = 0 för inre junction
        Range(Cells(CurrRow, 6), Cells(CurrRow, 7)) = "-"     '
        Range(Cells(CurrRow, 8), Cells(CurrRow, 9)) = 0#      ' K+  K-
        Range(Cells(CurrRow, 10), Cells(CurrRow, 10)) = "junction"      ' Namn
        Range(Cells(CurrRow, 11), Cells(CurrRow, 11)) = "-"   ' Ritning
        Range(Cells(CurrRow, 12), Cells(CurrRow, 12)) = "-"       ' Kraftnr
        Range(Cells(CurrRow, 13), Cells(CurrRow, 13)).Formula = "=OFFSET($A$1,ROW()-3,1)"
        Range(Cells(CurrRow, 14), Cells(CurrRow, 14)).Formula = "=OFFSET($A$1,ROW()+1,1)"
        Range(Cells(CurrRow, 15), Cells(CurrRow, 15)) = 2
        Range(Cells(CurrRow, 16), Cells(CurrRow, 16)) = 1
        Range(Cells(CurrRow, 17), Cells(CurrRow, 22)) = "-"
    End If

End Sub



Sub AddTmdpvol()
' Action: Adds a time-dependant volume at selected rows
'
   
    Dim CurrRow As Integer, currRowCnt As Integer, Word1 As String
    Dim question
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell or a row where you want to add a new TMDPVOL", vbExclamation, "Insert tmdpvol"
        Exit Sub
    End If
    
    CurrRow = Selection.row
    
    Word1 = Cells(CurrRow, 1)
    
    If Word1 <> "" Then
        question = MsgBox("Insert a time-dependent volume on row " + CStr(CurrRow) + "?", vbYesNoCancel, "Insert time-dependent volume")
        If question <> vbYes Then
            Exit Sub
        End If
    Else
        question = MsgBox("Insert a time-dependent volume on row " + CStr(CurrRow) + "?", vbYesNoCancel, "Insert time-dependent volume")
        If question <> vbYes Then
            Exit Sub
        End If
    
    End If
    
    Range(Cells(CurrRow, 1), Cells(CurrRow, 1)) = "Tmdpvol"
    Range(Cells(CurrRow, 2), Cells(CurrRow, 2)).Formula = "=CONCATENATE(""TMDV_"",ROW())"
    Range(Cells(CurrRow, 3), Cells(CurrRow, 3)) = 1#        ' Längd = 1.000 m
    Range(Cells(CurrRow, 4), Cells(CurrRow, 4)) = "-"       ' dx = "-"
    Range(Cells(CurrRow, 5), Cells(CurrRow, 5)) = 1#        ' Area = 1.000 m2
    Range(Cells(CurrRow, 6), Cells(CurrRow, 7)) = 0#        ' Vinklar
    Range(Cells(CurrRow, 8), Cells(CurrRow, 9)) = "-"       ' K+  K-
    Range(Cells(CurrRow, 10), Cells(CurrRow, 10)) = "TDVol" ' Namn
    Range(Cells(CurrRow, 11), Cells(CurrRow, 11)) = "-"     ' Ritning
    Range(Cells(CurrRow, 12), Cells(CurrRow, 16)) = "-"
    Range(Cells(CurrRow, 17), Cells(CurrRow, 17)) = 100000# ' Tryck i Pa
    Range(Cells(CurrRow, 18), Cells(CurrRow, 18)) = 293.15  ' Temp i K
    Range(Cells(CurrRow, 19), Cells(CurrRow, 22)) = "-"

End Sub



Sub AddFlowPath()
' Action: Adds a new flowpath (a comment, followed by a "Relapnr" and "Init" block
'
'
    Dim CurrRow As Integer, currRowCnt As Integer, Word1 As String, descrString As String, relapNr As String
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell where you want to add a new flowpath", vbExclamation, "Insert flowpath"
        Exit Sub
    End If
    
    CurrRow = Selection.row
    
    Word1 = Cells(CurrRow, 1)
    
    Dim question
    If Word1 <> "" Then
        question = MsgBox("Insert new flowpath on row " + CStr(CurrRow) + "?", vbYesNoCancel, "Insert new flowpath")
        If question <> vbYes Then
            Exit Sub
        End If
    Else
        question = MsgBox("Insert new flowpath on row " + CStr(CurrRow) + "?", vbYesNoCancel, "Insert new flowpath")
        If question <> vbYes Then
            Exit Sub
        End If
    
    End If
    
    descrString = InputBox(Prompt:="Description", Title:="New flowpath-Description", Default:="Flowpath N: From XXX to YYY")
    relapNr = InputBox(Prompt:="Start component numbering", Title:="New flowpath-CCC start", Default:="100")
    Rows(CurrRow & ":" & CurrRow + 1).Select
    Selection.Insert Shift:=xlDown, copyorigin:=xlFormatFromLeftOrAbove
    CurrRow = CurrRow
    
    Range(Cells(CurrRow, 1), Cells(CurrRow, 1)) = "* " & descrString
    Range(Cells(CurrRow + 1, 1), Cells(CurrRow + 1, 1)) = "Relapnr"
    Range(Cells(CurrRow + 1, 2), Cells(CurrRow + 1, 2)) = CInt(relapNr)
    Range(Cells(CurrRow + 2, 1), Cells(CurrRow + 2, 1)) = "Init"
    Range(Cells(CurrRow + 2, 2), Cells(CurrRow + 2, 2)) = 100000#
    Range(Cells(CurrRow + 2, 3), Cells(CurrRow + 2, 3)) = 293.15

End Sub

Sub AddVariable()
' Action: Inserts a Test matrix variable lookup in current cell
'
    Dim CurrRow As Integer, currRowCnt As Integer, Word1 As String, variable As String
    Dim variableList As Range
    Dim question
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell where you want to add a new variable lookup", vbExclamation, "Insert variable lookup"
        Exit Sub
    End If
    
    CurrRow = Selection.row
    
    Word1 = Cells(CurrRow, 1)
    
    question = MsgBox("Insert a variable lookup at cell """ + CStr(Selection.Address) + """?", vbYesNoCancel, "Insert variable lookup")
    If question <> vbYes Then Exit Sub
        
    Set variableList = Range(TESTMATRIX_VARIABLE_NAMES)
    
    Dim tmpStr As String
    Dim i As Integer
    tmpStr = ""
    For i = 1 To variableList.Columns.Count
        If variableList(1, i) <> "" Then tmpStr = tmpStr & variableList(1, i) & ", "
    Next i
    Debug.Print tmpStr
    
    variable = InputBox(Prompt:=tmpStr, Title:="Choose variable", Default:=tmpStr)
    
    Selection.Formula = "=INDEX(" & TESTMATRIX_VARIABLE_VALUES & ", " & TESTMATRIX_CURRENT_SET & _
                        ",MATCH(" & Chr(34) & variable & Chr(34) & "," & TESTMATRIX_VARIABLE_NAMES & ",0))"

End Sub

Sub AddTripVariable()
' Action: Adds a variable trip
'
    Dim CurrRow As Integer, currRowCnt As Integer, Word1 As String
    Dim question
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell or a row where you want to add a new TRIP", vbExclamation, "Insert trip"
        Exit Sub
    End If
    
    CurrRow = Selection.row
    
    Word1 = Cells(CurrRow, 1)
    
    If Word1 <> "" Then
        question = MsgBox("Insert a variable trip on row " + CStr(CurrRow) + "?", vbYesNoCancel, "Insert variable trip")
        If question <> vbYes Then
            Exit Sub
        End If
    End If
    
    TurnOffScreenUpdate True
    Range(Cells(CurrRow, 1), Cells(CurrRow, 1)) = "TripVar"
    Range(Cells(CurrRow, 2), Cells(CurrRow, 2)).Formula = "=CONCATENATE(""TRIP_"",ROW())"
    Range(Cells(CurrRow, 3), Cells(CurrRow, 3)) = "<ID>"
    Range(Cells(CurrRow, 4), Cells(CurrRow, 4)) = "mflowj-CCC010000"
    Range(Cells(CurrRow, 5), Cells(CurrRow, 5)) = "ge"
    Range(Cells(CurrRow, 6), Cells(CurrRow, 6)) = "<ID>"
    Range(Cells(CurrRow, 7), Cells(CurrRow, 7)) = "null-0"
    Range(Cells(CurrRow, 8), Cells(CurrRow, 8)) = 0#
    Range(Cells(CurrRow, 9), Cells(CurrRow, 9)) = "n"
    TurnOffScreenUpdate False

End Sub

Sub AddTripLogical()
' Action: Adds a logical trip
'
    Dim CurrRow As Integer, currRowCnt As Integer, Word1 As String
    Dim question
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell or a row where you want to add a new TRIP", vbExclamation, "Insert trip"
        Exit Sub
    End If
    
    CurrRow = Selection.row
    
    Word1 = Cells(CurrRow, 1)
    
    If Word1 <> "" Then
        question = MsgBox("Insert a logical trip on row " + CStr(CurrRow) + "?", vbYesNoCancel, "Insert variable trip")
        If question <> vbYes Then
            Exit Sub
        End If
    End If
    
    TurnOffScreenUpdate True
    
    Range(Cells(CurrRow, 1), Cells(CurrRow, 1)) = "TripLog"
    Range(Cells(CurrRow, 2), Cells(CurrRow, 2)).Formula = "=CONCATENATE(""TRIP_"",ROW())"
    Range(Cells(CurrRow, 3), Cells(CurrRow, 3)) = "<TRIP-ID1>"
    Range(Cells(CurrRow, 4), Cells(CurrRow, 4)) = "and"
    Range(Cells(CurrRow, 5), Cells(CurrRow, 5)) = "<TRIP-ID2>"
    Range(Cells(CurrRow, 6), Cells(CurrRow, 6)) = "n"
    
    TurnOffScreenUpdate False

End Sub


Sub dublicateCurrLoadCase()
' Funktion som dublicerar aktuellt lastfall i runMatrix
'
'
    Dim CurrRow As Integer, loadCase As String, NewLoadCase As String
    Dim readCol As Integer
    
    With Range(TESTMATRIX)
        If TypeName(Selection) <> "Range" Then
            MsgBox "Select a single cell or row to dublicate a load definition", vbExclamation, "Dublicate load case in Test matrix"
            Exit Sub
        ElseIf Selection.Worksheet.Name <> TESTMATRIX_SHEET Then
            MsgBox "Select a load case in worksheet """ & TESTMATRIX_SHEET & """ to dublicate", vbExclamation, "Dublicate load case in Test matrix"
            Exit Sub
        ElseIf ActiveCell.row < .Rows(1).row Or ActiveCell.row > .Rows(.Rows.Count).row Or ActiveCell.Column < .Columns(1).Column Or ActiveCell.Column > .Columns(.Columns.Count).Column Then
            MsgBox "Outside range"
            Exit Sub
        End If
    End With
    
    readCol = 2   ' Kolumn där lastbeteckningen står
    
    CurrRow = ActiveCell.row
    
    loadCase = Cells(CurrRow, readCol)
    
    Dim question
    question = MsgBox("Dublicate row """ + loadCase + """?. ", vbYesNoCancel, "Dublicera case")
    If question <> vbYes Then
        Exit Sub
    End If
    
    NewLoadCase = InputBox(Prompt:="New label", Title:="New label", Default:=loadCase)
    
    ' Select row below, insert new row
    Rows(CurrRow + 1).Select
    Selection.Insert Shift:=xlDown, copyorigin:=xlFormatFromRightOrBelow ' CopyOrigin:=xlFormatFromLeftOrAbove
    Rows(CurrRow).Select
    Selection.Copy
    Rows(CurrRow + 1).Select
    ActiveSheet.Paste
    
    Cells(CurrRow + 1, readCol) = NewLoadCase
End Sub



Sub AddLoopCheck()
' Action: Adds a loop check
'
    Dim i As Integer, j As Integer
    Dim s As New ResourceSprintf
    
    Dim Inputdeck As New Text2Relap
    If Inputdeck.ReadOk = False Then Exit Sub
    
    Dim FirstComp As Boolean, LastComp As Boolean
    Dim LastCompRow As Integer, LastCompIndex As Integer
    LastCompRow = -1
    With Inputdeck.HydroSystem
        For i = 1 To .Components.Count
            With .Components(i)
                If .ObjectType = HydroComp Then
                    If FirstComp = True Then
                        FirstComp = False
                        Range(Cells(.RowBegin, 26), Cells(.RowBegin, 26)) = "X"
                    End If
                    LastCompRow = .RowEnd
                    LastCompIndex = i
                    For j = .RowBegin To .RowEnd + 1
                        Range(Cells(j, 24), Cells(j, 24)).Formula = s.sprintf("=IF(OR(A%1$d=""Pipe"", A%1$d=""Tmdpvol""), X%2$d+C%1$d*SIN(F%1$d*PI()/180), X%2$d)", j, j - 1)
                    Next j
                ElseIf .ObjectType = Comment1 Then
                    FirstComp = True
                    If LastCompRow <> -1 Then
                        Range(Cells(LastCompRow, 26), Cells(LastCompRow, 26)) = "Y"
                    End If
                End If
            End With
        Next i
    End With
End Sub
