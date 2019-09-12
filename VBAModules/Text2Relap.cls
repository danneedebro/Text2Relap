VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Text2Relap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --------------------------- TEXT2RELAP --------------------------
' -----------------------------------------------------------------
'                Name:  Text2Relap
'             Project:  Text2Relap
'                Type:  Class module
'              Author:  Daniel Edebro
'         Description:  Creates RELAP5 input files from a simplified input
'
' -----------------------------------------------------------------
' USAGE:
' Dim Inputdeck as New Text2Relap  ' Reads input, connects junctions and create forces
' If Inputdeck.ReadOk = true
'   Inputdeck.WriteToFile("Case1.i")
' End if
'
' ' Loop through all read components
' With Inputdeck.HydroSystem
'   For i = 1 to .Components.Count
'       Debug.print "CCC=" & .Component(i).CCC & ", type=" & .Component(i).Info.MainTypeStr
'   Next i
' End With
'
'
'
'
Option Explicit
'
' MEMBER VARIABLES
Private Const DEBUG_ME As Boolean = True
Private Const DEFAULT_INPUT_WORKSHEET As String = "text2relap"
Private Const DEFAULT_INPUT_RANGE_LASTROW As Integer = 1000

Private Const TEXT2RELAP_VERSION = "2.0.0"
Private Const TEXT2RELAP_BRANCH = "beta.2"

Private InputRange As Range
Private pHydroSystem As HydroSystem
Private IsInitialized As Boolean
Private currLine1 As Long
Private currLine2 As Long

Private pAbort As Boolean
Private pWarnings As Boolean

Private pIsInitialized As Boolean
Private pReadOk As Boolean

Private pInputSheet As String
Private pInputRange As String

Private pSource As InputObject

' Variables relating to Conditional formatting
Private Type FConds
    Word1 As String
    WordsColor(1 To 22) As Boolean
    WordsGrey(1 To 22) As Boolean
    ColorColor As Long
    colorGrey As Long
    NumberOfWords As Integer
End Type

Private pFCondInput() As FConds


'
'
' PROPERTIES
Public Property Get HydroSystem() As HydroSystem
    Set HydroSystem = pHydroSystem
End Property

Private Property Let InputSheet(arg As String)
    Dim i As Integer
    Dim WorksheetExists As Boolean
    For i = 1 To ThisWorkbook.Worksheets.Count
        If ThisWorkbook.Worksheets(i).Name = arg Then WorksheetExists = True
    Next i
    If WorksheetExists = False Then
        
    Else
        pInputSheet = arg
    End If
End Property

Public Property Let Abort(arg As Boolean)
    pAbort = arg
End Property

Public Property Get Warnings() As Boolean
    Warnings = pWarnings
End Property

Public Property Let Warnings(arg As Boolean)
    pWarnings = arg
End Property

Public Property Get ReadOk() As Boolean
    ReadOk = pReadOk
End Property

Public Property Get Text2RelapVersion() As String
    Text2RelapVersion = TEXT2RELAP_VERSION
End Property

Public Property Get Text2RelapBranch() As String
    Text2RelapBranch = TEXT2RELAP_BRANCH
End Property

Public Property Get Source() As InputObject
    Set Source = pSource
End Property

'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    pIsInitialized = False
End Sub

Public Sub Create(Optional ByVal InputWorksheet As String = "", Optional ByVal LastRow As Integer = -1)
' Action: Initialize the input data by first reading it and making junction
'         connections
'
    On Error GoTo ErrorHandler
    Dim StatusLog As String

    ' Set Input Worksheet
    If InputWorksheet = "" Then InputSheet = DEFAULT_INPUT_WORKSHEET Else InputSheet = InputWorksheet
    
    ' Set Input range
    If LastRow = -1 Then LastRow = DEFAULT_INPUT_RANGE_LASTROW
    With ThisWorkbook.Worksheets(pInputSheet)
        Set InputRange = .Range(.Cells(1, 1), .Cells(LastRow, 22))
    End With
    
    Set pSource = New InputObject
    pSource.CreateFromParts pInputSheet, 1, CDbl(LastRow)
    
    ' Create HydroSystem
    If DEBUG_ME = True Then Debug.Print "0. CREATING"
    Set pHydroSystem = New HydroSystem
    HydroSystem.Create Me
    If DEBUG_ME = True Then Debug.Print "1. CREATING"
    HydroSystem.DoPreRead
    
    '
    If DEBUG_ME = True Then Debug.Print "1. READING INPUT" & vbNewLine
    If ReadInput() = True Then
        pReadOk = True
        
        If DEBUG_ME = True Then Debug.Print "3. POST READ OPERATIONS"
        HydroSystem.DoPostRead
        IsInitialized = True
    Else
        pReadOk = False
        IsInitialized = False
    End If
    
    Exit Sub
ErrorHandler:
    Dim answ
    Me.HighlightCurrentLines
    Err.Source = TypeName(Me) & ".Init, " & Err.Source
    
    answ = MsgBox("Unexpected error in """ & Err.Source & """" & vbNewLine & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description & vbNewLine & vbNewLine & "Continue reading input?", vbCritical + vbYesNo)
    If answ = vbNo Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Else
        Resume Next
    End If

    Debug.Print "   Fail"
    pReadOk = False
End Sub
'
'
' CLASS METHODS
Private Function ReadInput() As Boolean
' Action: Reads the input data line by line and depending on the first word
'         performs different actions
'
    On Error GoTo ErrorHandler
    Dim inputRow() As Variant, inputRows() As Variant
    Dim Word1
    Dim i As Long, j As Long
    Dim IgnoreActive As Boolean
    
    Dim InputData As New InputObject
    InputData.Create InputRange
    
    ' Sweep 1 - Read general settings
    For i = 1 To InputData.NumberOfRows
        ' If user cancels further reading in subroutine
        If pAbort = True Then
            ReadInput = False
            Exit Function
        End If
    
        'Debug.Print "Word1 = " & InputData(CLng(i), CLng(1))
    
        'inputRow = RangeToArray(InputRange, i, i)
        Word1 = LCase(InputData(i, 1))
        currLine1 = i
        currLine2 = i
        
        Select Case Word1
            Case "title"
                HydroSystem.Settings.Title = InputData(i, 2)
            Case "include"
                HydroSystem.Settings.IncludeFolder = InputData(i, 2)
            Case "tripformat"
                HydroSystem.Settings.TripFormat = CheckInput(InputData(i, 2), Numeric:=True, valueDefault:=0, causeError:=True, FieldName:="Trip format", word:=2)
            Case "cvformat"
                HydroSystem.Settings.CntrlvarFormat = CheckInput(InputData(i, 2), Numeric:=True, valueDefault:=999, causeError:=True, FieldName:="Cntrlvar format", word:=2)
            Case "forcecalc"
                HydroSystem.Settings.ForceCalcMethod = CheckInput(InputData(i, 2), Numeric:=True, valueDefault:=2, causeError:=True, FieldName:="Force calculation method", word:=2)
            Case "comments"
                If InputData(i, 2) = 0 Then
                    HydroSystem.Settings.WriteComments = False
                Else
                    HydroSystem.Settings.WriteComments = True
                End If
        End Select
    
    Next i
    
    
    ' Sweep 2 - Read components and other things
    For i = 1 To InputData.NumberOfRows
        ' If user cancels further reading in subroutine
        If pAbort = True Then
            ReadInput = False
            Exit Function
        End If
        
        If IsEmpty(InputData(i, 1)) = True Then GoTo Next_i
        
        'inputRow = RangeToArray(InputRange, i, i)
        Word1 = LCase(InputData(i, 1))
        currLine1 = i
        currLine2 = i
        
        ' Debug
        If DEBUG_ME = True Then
            Debug.Print "Reading row " + CStr(i) + ": " + Word1
            If i = 25 Then
                Debug.Print "   Place Breakpoint here"
            End If
        End If
        
        Select Case Word1
            Case "ignore": IgnoreActive = True
            Case "/ignore": IgnoreActive = False
        End Select
        
        If IgnoreActive = True Then GoTo Next_i
        
        Select Case Word1
            Case "relapnr"
                HydroSystem.Settings.CurrentCCC = InputData(i, 2)
                
            Case "init", "initgas"
                HydroSystem.Settings.Inits.Add InputData.MakeCopy(i)
                
            Case "pipe"
                For j = i + 1 To InputRange.Rows.Count
                    If LCase(InputRange(j, 1)) <> "pipe" Then
                        inputRows = RangeToArray(InputRange, i, j - 1)
                        Exit For
                    End If
                Next j
                i = j - 1
                currLine2 = i
                HydroSystem.Components.Add HydroComp, InputData.MakeCopy(FirstRow:=currLine1, LastRow:=currLine2)
                
            Case "tmdpvol", "snglvol", "junction", "mtrvlv", "srvvlv", "inrvlv", _
                 "chkvlv", "trpvlv", "pump", "tmdpjun"
                HydroSystem.Components.Add HydroComp, InputData.MakeCopy(i)
                
            Case "custom"
                HydroSystem.Components.Add Custom, InputData.MakeCopy(i)
                
            Case "triggerwords", "replacements"
                HydroSystem.Settings.Replacements.Add InputData.MakeCopy(i)
                
            Case "timestep"
                HydroSystem.Settings.Timesteps.Add InputData.MakeCopy(i)
                
            Case "refvol"
                HydroSystem.Settings.Refvols.Add InputData.MakeCopy(i)
                
            Case "cntrlvarnr"
                HydroSystem.Settings.CurrentCntrlvarNum = CheckInput(InputData(i, 2), Numeric:=True, valueDefault:=HydroSystem.Settings.CurrentCCC, causeError:=True, FieldName:="CntrlvarNr", word:=2)
            
            Case "header"
                HydroSystem.Settings.Filename = InputData(i, 2)
                
            Case "develop"
                HydroSystem.Settings.DevelopmentControl = InputData(i, 2)
            
            Case "tripvar"
                HydroSystem.Components.Add TripVariable, InputData.MakeCopy(i)
            
            Case "triplog"
                HydroSystem.Components.Add TripLogical, InputData.MakeCopy(i)
            
            Case Else
                If Left(Word1, 1) = "*" Then
                    HydroSystem.Components.Add Comment1, InputData.MakeCopy(i)
                End If
                
        End Select
Next_i:
    Next i
    
    ReadInput = True

    Exit Function
ErrorHandler:
    Dim answ
    Me.HighlightCurrentLines
    Err.Source = TypeName(Me) & ".ReadInput(line=" & currLine1 & "), " & Err.Source
    
    answ = MsgBox("Unexpected error in """ & Err.Source & """" & vbNewLine & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description & vbNewLine & vbNewLine & "Continue reading input?", vbCritical + vbYesNo)
    If answ = vbNo Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Else
        Resume Next
    End If

End Function

Public Function CheckInput(ByVal Value As Variant, ByVal Numeric As Boolean, ByVal valueDefault As Variant, ByVal causeError As Boolean, _
                           Optional ByVal FieldName As String = "", Optional ByVal word As Integer = -1, Optional ByVal IsFatal As Boolean = False, _
                           Optional IsBinaryString As Boolean = False)
' Action: Checks input values
'
'Inputs:
'           value = The input that is to be checked
'           Numeric = True if value is supposed to be numeric, false if string input
'           valueDefault = Default value if value not correct
'           causeError = Produce error message
'
'Outputs:
'           value = if value correct
'           valueDefault = if value incorrect
'
    On Error GoTo ErrorHandler
    Dim NewValue As Variant
    Dim InputError As Boolean
    Dim answ
    
    If IsEmpty(Value) Then
        
        CheckInput = valueDefault
        InputError = True
    Else
        If Numeric = True Then
            If IsNumeric(Value) = True Then
                CheckInput = Value  ' Correct value
            Else
                CheckInput = valueDefault  ' Incorrect
                InputError = True
            End If
        Else
            ' Check if binary string
            If IsBinaryString = True Then
                Dim i As Integer
                For i = 1 To Len(Value)
                    If Mid(Value, i, 1) = "0" Or Mid(Value, i, 1) = "1" Then
                        CheckInput = Value
                    Else ' If string contains characters other than "0" or "1" - break loop and return default value
                        CheckInput = valueDefault
                        Exit For
                    End If
                Next i
                
            Else
                CheckInput = Value
            End If
        End If
        
    End If

    If causeError = True And InputError = True Then
        Me.HighlightCurrentLines word
        
        If IsFatal = True Then
            answ = MsgBox("Fatal Input error: """ & FieldName & """-value is empty or wrong format.", vbCritical, "Fatal Input error")
            Me.Abort = True
        Else
            Me.Warnings = True
            answ = MsgBox("Input error: """ & FieldName & """-value is empty or wrong format. Setting default value " & CStr(valueDefault) & vbNewLine & vbNewLine & "Continue?", vbExclamation + vbOKCancel, "Input error")
            If answ = vbCancel Then
                Err.Raise 0, "", "Input error: """ & FieldName & """-value is empty or wrong format."
                Me.Abort = True
            End If
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".CheckInput(Value:='" & Value & "'), " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

Public Sub HighlightCurrentLines(Optional ByVal word As Integer = -1, Optional ByVal SelectRow As Integer = -1)
' Action: Higlights a certain row
'
    Dim FirstRow As Integer, LastRow As Integer
    
    If SelectRow = -1 Then
        FirstRow = currLine1
        LastRow = currLine2
    Else
        FirstRow = SelectRow
        LastRow = SelectRow
    End If
    

    Debug.Print "Hello"
    If word = -1 Or word > InputRange.Columns.Count Then
        Range(InputRange.Rows(FirstRow), InputRange.Rows(LastRow)).Select
    Else
        If currLine2 > currLine1 Then
            Range(InputRange.Rows(FirstRow), InputRange.Rows(LastRow)).Columns(word).Select
        Else
            InputRange.Rows(FirstRow).Columns(word).Select
        End If
    End If
End Sub


Public Function WriteToFile(ByVal Filename As String) As Boolean
' Action: Writes the input file to file
'
'
    On Error GoTo ErrorHandler_1
    
    Dim WriteComments As Boolean
    WriteComments = HydroSystem.Settings.WriteComments
    
    ' Activate Microsoft Scripting runtime Tools...References...Microsoft scripting runtime
    Dim fso As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    
    Set fso = New Scripting.FileSystemObject
    Set ts = fso.CreateTextFile(Filename)
    
    On Error GoTo ErrorHandler_2
    Dim i As Integer
    For i = 1 To HydroSystem.Components.Count
        
        If HydroSystem.Components(i).ObjectType = Force Then GoTo Next_i
        
        If pAbort = True Then
            WriteToFile = False
            Exit Function
        End If
        
        ' Debug
        If DEBUG_ME = True Then Debug.Print "Writing component " & CStr(i) & " out of " & CStr(HydroSystem.Components.Count) & ": " & HydroSystem.Components(i).ComponentInfo
        
        If i = 1 And DEBUG_ME = True Then
            Debug.Print "Break point here"
        End If
        
        HydroSystem.Components(i).WriteToFile ts, WriteComments
        If WriteComments = False Then ts.WriteLine ""
Next_i:
    Next i
    
    ' Create and write forces
    ts.WriteLine ""
    ts.WriteLine "* ------------ FORCES ------------"
    ts.WriteLine ""
    
    Dim CurrForce As ComponentForce
    
    On Error GoTo ErrorHandler_3
    For i = 1 To HydroSystem.Components.Count
        If HydroSystem.Components(i).ObjectType = Force Then
            Set CurrForce = HydroSystem.Components(i)
            If DEBUG_ME = True Then Debug.Print "Writing force " & CStr(CurrForce.ForceNumber)
            CurrForce.WriteToFile ts, WriteComments
            'HydroSystem.Force(index:=i).WriteToFile ts, WriteComments
        End If
    Next i
    
    ' Mark end of input file
    ts.WriteBlankLines 2
    ts.WriteLine ".end"
    
    ts.Close
    Set fso = Nothing
    
    WriteToFile = True
    
    Exit Function
ErrorHandler_1: ' Error during creation of textstream object
    Err.Source = TypeName(Me) & ".WriteToFile [Error creating FileObject] , " & Err.Source
    GoTo ErrorHandler
ErrorHandler_2: ' Error during writing files
    Err.Source = TypeName(Me) & ".WriteToFile [Error writing component(" & i & ")='" & HydroSystem.Components(i).ObjectType & "'], " & Err.Source
    GoTo ErrorHandler
ErrorHandler_3:
    Err.Source = TypeName(Me) & ".WriteToFile [Error writing component(" & i & ")='F" & CurrForce.ForceNumber & "'], " & Err.Source
    GoTo ErrorHandler
    
ErrorHandler:
    Dim answ
    answ = MsgBox("Unexpected error in """ & Err.Source & """" & vbNewLine & vbNewLine & "Number: " & Err.Number & vbNewLine & "Description: " & Err.Description & vbNewLine & vbNewLine & "Continue writing to file?", vbCritical + vbYesNo, "Error writing to file")
    If answ = vbNo Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Else
        Resume Next
    End If
    MsgBox ""
    Err.Raise Err.Number, TypeName(Me) & ".WriteToFile(Filename:='" & Filename & "'), " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Public Sub ProbeInput(ByVal row1 As Integer, ByVal row2 As Integer, OutputType As Integer)
' Action: Probe input
'
' Inputs:
'   Row1 {Int}: First row of selected range
'   Row2 {Int}: Last row of selected range
'   OutputType {Int}: 0 = Output short component list to msgbox
'                     1 = Output text that is written to msgbox
'
    Dim i As Integer
    Dim WriteComments As Boolean
    Dim ts As New ResourceTextStreamDummy
    Dim Pos As Integer, posPrev As Integer
    Dim ResidualString As String
    Dim answ
    Dim AllOutput As String
    Dim CurrHydroComp As ComponentHydro
    
    WriteComments = HydroSystem.Settings.WriteComments

    For i = 1 To HydroSystem.Components.Count
        With HydroSystem.Components(i)
            If (.RowBegin >= row1 And .RowBegin <= row2) Or (.RowEnd >= row1 And .RowEnd <= row2) Or (.RowBegin <= row1 And .RowEnd >= row2) Then
                If OutputType = 0 Then
                    
                    ts.WriteLine .ComponentInfo
                ElseIf OutputType = 1 Then
                    .WriteToFile ts, WriteComments
                    ResidualString = ts.TextStream
                    
                    posPrev = 0
                    Pos = 1
                    Do
                        Pos = InStr(posPrev + 1, ResidualString, vbNewLine)
                        If Pos > 1024 Or Pos = 0 Then
                            'Debug.Print "NEW BATCH"
                            answ = MsgBox(Left(ResidualString, posPrev), vbOKCancel, .ComponentInfo)
                            If answ = vbCancel Then Exit Sub
                            'Debug.Print Left(ResidualString, posPrev)
                            ResidualString = Mid(ResidualString, posPrev + 1)
                            posPrev = 0
                        Else
                            posPrev = Pos
                        End If
                    Loop While Pos <> 0
                    ts.Clear
                ElseIf OutputType = 2 Then
                    .WriteToFile ts, WriteComments
                    AllOutput = AllOutput & ts.TextStream & vbNewLine
                ElseIf OutputType = 3 Then
                    If .ObjectType = HydroComp Then
                        Set CurrHydroComp = HydroSystem.Components(i)
                        CurrHydroComp.Child.WriteToFile ts, WriteComments, False
                        If DEBUG_ME = True Then CurrHydroComp.Child.WriteToFile ts, WriteComments, True ' Debug purpose
                    Else
                        .WriteToFile ts, WriteComments
                    End If
                    'AllOutput = AllOutput & ts.TextStream & vbNewLine
                End If
            End If
        End With
    Next i
    
    If OutputType = 0 Then
        If ts.TextStream = "" Then
            MsgBox "Current selection contain no probeable input. Select rows that contains a component that prints to output file", vbInformation
        Else
            MsgBox ts.TextStream
        End If
    ElseIf OutputType = 2 Or OutputType = 3 Then
        UserForm1.TextBox1.Text = ts.TextStream
        UserForm1.Show
    End If
End Sub

Public Sub ModelSummary()
' Action: Writes out a model summary
'
    Dim i As Integer, j As Integer
    Dim ResultString As String
    
    ' Stores forces, CCCs and .inc-filenames in current flow path
    Dim ListForces As New Collection
    Dim ListCCCs As New Collection
    Dim ListIncs As New Collection
    
    Dim FlowPathLength As Double  ' total length of flowpath
    
    Dim CurrentFlowPath As String
    CurrentFlowPath = "Header"
    
    Dim CurrentFlowPathComment As ComponentComment
    Dim CurrentHydroComponent As ComponentHydro
    Dim CurrentCustomComponent As ComponentCustom
    

    For i = 1 To HydroSystem.Components.Count
        With HydroSystem.Components(i)
            If .ObjectType = Comment1 Or i = HydroSystem.Components.Count Then
                Dim ComponentList As String, ForceList As String, IncludeList As String
                
                ' Shortens to 100-110 instead of 100, 101, 102, .., 110
                For j = 1 To ListCCCs.Count
                    If j = 1 Then
                        ComponentList = ListCCCs(j)
                    ElseIf j > 1 And ListCCCs(j) - ListCCCs(j - 1) <> 1 Then
                        ComponentList = ComponentList & "-" & ListCCCs(j - 1) & ", " & ListCCCs(j)
                    ElseIf j = ListCCCs.Count Then
                        ComponentList = ComponentList & "-" & ListCCCs(j - 1)
                    End If
                Next j
                Set ListCCCs = New Collection
            
                ' Shortens to 1-10 instead of 1, 2, 3, 3, 3, 4, 4
                For j = 1 To ListForces.Count
                    If j = 1 Then
                        ForceList = ListForces(j)
                    ElseIf j > 1 And ListForces(j) - ListForces(j - 1) > 1 Then
                        ForceList = ForceList & "-" & ListForces(j - 1) & ", " & ListForces(j)
                    ElseIf j = ListForces.Count Then
                        ForceList = ForceList & "-" & ListForces(j - 1)
                    End If
                Next j
                Set ListForces = New Collection
                
                For j = 1 To ListIncs.Count
                    IncludeList = IIf(j = 1, "", IncludeList & ", ") & ListIncs(j)
                Next j
                Set ListIncs = New Collection
            
                ResultString = ResultString & CurrentFlowPath & vbNewLine & _
                              "   CCC: " & ComponentList & vbNewLine & _
                              "  .inc: " & IncludeList & vbNewLine & _
                              "Forces: " & ForceList & vbNewLine & _
                              "Length: " & FlowPathLength & " m" & vbNewLine & vbNewLine
                
                ComponentList = ""  ' Clear component list
                ForceList = ""
                FlowPathLength = 0#
                
                ' Fetch new comment
                If .ObjectType = Comment1 Then
                    Set CurrentFlowPathComment = HydroSystem.Components(i)
                    CurrentFlowPath = CurrentFlowPathComment.Text
                End If
                
            ' If hydro component, add its CCC- and forcenumbers to list
            ElseIf .ObjectType = HydroComp Then
                
                Set CurrentHydroComponent = HydroSystem.Components(i)
                
                With CurrentHydroComponent
                    ListCCCs.Add .CCC
                    If .Filename <> "" Then ListIncs.Add .Filename
                    
                    For j = 1 To CurrentHydroComponent.Segments.Count
                        With .Segments(j)
                            FlowPathLength = FlowPathLength + .Dx * .Volumes
                            If .ForceNumber <> 0 Then ListForces.Add .ForceNumber
                        End With
                    Next j
                End With
                
            ' If custom component, add its filename to list
            ElseIf .ObjectType = Custom Then
                Set CurrentCustomComponent = HydroSystem.Components(i)
                ListIncs.Add CurrentCustomComponent.Filename
                ComponentList = ComponentList & IIf(ComponentList = "", "", ", ") & "Custom=" & CurrentCustomComponent.Filename
                
            End If
        End With
    Next i
    
    UserForm1.TextBox1.Text = ResultString
    UserForm1.Show
    
End Sub
'
'
' STATIC METHODS
Private Function RangeToArray(ByRef MainRange As Range, StartRow As Long, EndRow As Long) As Variant
' Action: Returns a reduced variant 2-dim array from input range
'
' Inputs:   MainRange = A MxN Range to be converted to array
'           row1 = Starting row
'           row2 = Ending row
'
' Output:   A variant with as many columns as 'MainRange'
'
    Dim subRange As Range
    Dim arr() As Variant
    With MainRange
        Set subRange = .Range(.Cells(StartRow, 1), .Cells(EndRow, .Columns.Count))
    End With
    arr = subRange
    RangeToArray = arr
End Function


























