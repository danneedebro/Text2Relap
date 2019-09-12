Attribute VB_Name = "Text2RelapMain"
Option Explicit

' TODO:
' Fundera på Heat structures
' TMDPVOL WriteToFile - TidsDubletter i när man sammanfogat Pressure och Temperature
' Belys external connections-junctions, dvs sådana som inte kopplar till angränsande volymer
' Merge


Public Const TESTMATRIX = "TestMatrix"
Public Const TESTMATRIX_SHEET = "runMatrix"
Public Const TESTMATRIX_CURRENT_SET = "CurrentIndex"
Public Const TESTMATRIX_CURRENT_FILEPATH = "CurrentFilename"
Public Const TESTMATRIX_FILENAMES = "TestMatrixFilenames"
Public Const TESTMATRIX_VARIABLE_NAMES = "TestMatrixVariableNames"
Public Const TESTMATRIX_VARIABLE_VALUES = "TestMatrixVariableValues"


Private Type TFileWriteStatus
    message As String
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
    Application.CalculateFull
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
        Application.CalculateFull
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
        ResultString = ResultString & "Case index: " & CStr(i) & vbNewLine & FileWriteStatus(i).message & vbNewLine & vbNewLine
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
                Result.message = "Info: " & InputFile.FullPath & " created successfully"
                Result.Warnings = False
            Else
                Result.message = "Warning: " & InputFile.FullPath & " created with warnings"
                Result.Warnings = True
            End If
        Else
            Result.FileWritten = False
            Result.Warnings = True
            Result.message = "Error: " & InputFile.FullPath & " NOT created successfully"
        End If
    Else
        Result.FileWritten = False
        Result.Warnings = True
        Result.Abort = False
    End If
    
    ReadInputAndWriteToFile = Result
    
End Function


Sub ProbeOutput()
Attribute ProbeOutput.VB_ProcData.VB_Invoke_Func = "I\n14"
' Action: Probe Output by selecting cells with rows that contain a plotted output
'
    Dim SelectedRange As Range
    Dim row1 As Integer, row2 As Integer
    
    If TypeName(Selection) = "Range" Then
        Set SelectedRange = Selection
        With SelectedRange
            row1 = .Rows(1).Row
            row2 = .Rows(.Rows.Count).Row
        End With

        Dim Inputdeck As Text2Relap
        Set Inputdeck = NewInputdeck(ActiveSheet.Name)
        If Inputdeck.ReadOk = True Then
            Inputdeck.ProbeInput row1, row2, 0
            Inputdeck.ProbeInput row1, row2, 3
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

