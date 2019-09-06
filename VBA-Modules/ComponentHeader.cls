VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComponentHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ComponentHeader.cls
'
'
Option Explicit

Implements iComponent
'
'
' MEMBER VARIABLES
Private pHydroSystem As HydroSystem

Private pRowBegin As Integer
Private pRowEnd As Integer
'
'
' SHARED PROPERTIES AND METHODS
Private Property Get iComponent_ObjectType() As R5Component
    iComponent_ObjectType = Header
End Property

Private Function iComponent_ComponentInfo() As String
    iComponent_ComponentInfo = "Header"
End Function

Private Property Get iComponent_RowBegin() As Integer
    iComponent_RowBegin = pRowBegin
End Property

Private Property Get iComponent_RowEnd() As Integer
    iComponent_RowEnd = pRowEnd
End Property

Private Sub iComponent_WriteToFile(ts As Object, ByVal wC As Boolean)
    WriteToFile ts, wC
End Sub
'
'
' PROPERTIES
Public Property Get HydroSystem() As HydroSystem
    Set HydroSystem = pHydroSystem
End Property

Public Property Get Settings() As PropertiesSettings
    Set Settings = HydroSystem.Settings
End Property
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    
End Sub

Public Sub Create(ByRef ParentSystem As HydroSystem, ByVal RowBegin As Integer, ByVal RowEnd As Integer)
    On Error GoTo ErrorHandler
    pRowBegin = RowBegin
    pRowEnd = RowEnd
    
    Set pHydroSystem = ParentSystem
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".Create, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
'
'
' CLASS METHODS
Public Function ReplaceWords(ByVal InputString As String) As String
' Action: Replaces certain words with corresponding values
'
'
    Dim s As New ResourceSprintf
    Dim Replacements(1 To 5) As Variant
    
    ' Replace keywords with properties specific to junction type components
    Replacements(1) = Array("[TEXT2RELAPVERSION]", HydroSystem.Inputdeck.Text2RelapVersion)
    Replacements(2) = Array("[TEXT2RELAPBRANCH]", HydroSystem.Inputdeck.Text2RelapBranch)
    Replacements(3) = Array("[CURRENTDATE]", Format(Date, "YYYY-MM-dd"))
    Replacements(4) = Array("[CURRENTTIME]", Format(Time, "hh:mm:ss"))
    Replacements(5) = Array("[TITLE]", Settings.Title)
    
    
    Dim i As Integer
    For i = LBound(Replacements) To UBound(Replacements)
        InputString = Replace(InputString, Replacements(i)(0), Replacements(i)(1))
    Next i
    
    ' Replace refvol strings
    For i = 1 To Settings.Refvols.Count
        With Settings.Refvols(i)
            InputString = Replace(InputString, s.sprintf("[REFVOL-%d]", i), s.sprintf("%9s", .ReferenceVolume))
            InputString = Replace(InputString, s.sprintf("[ELEV-%d]", i), s.sprintf("%7.3f", .Elevation))
            InputString = Replace(InputString, s.sprintf("[FLUIDTYPE-%d]", i), s.sprintf("%6s", .Fluid))
            InputString = Replace(InputString, s.sprintf("[SYSTEMNAME-%d]", i), s.sprintf("%8s", .SystemName))
        End With
    Next i
    
    For i = 1 To Settings.Timesteps.Count
        With Settings.Timesteps(i)
            InputString = Replace(InputString, s.sprintf("[TIMEEND-%d]", i), s.sprintf("%8.3f", .TimeEnd))
            InputString = Replace(InputString, s.sprintf("[DTMIN-%d]", i), s.sprintf("%9.3e", .TimestepMin))
            InputString = Replace(InputString, s.sprintf("[DTMAX-%d]", i), s.sprintf("%9.3e", .TimestepMax))
            InputString = Replace(InputString, s.sprintf("[SSDTT-%d]", i), s.sprintf("%5s", .ControlOption))
            InputString = Replace(InputString, s.sprintf("[FMINOR-%d]", i), s.sprintf("%5d", .FrequencyMinor))
            InputString = Replace(InputString, s.sprintf("[FMAJOR-%d]", i), s.sprintf("%7d", .FrequencyMajor))
            InputString = Replace(InputString, s.sprintf("[FRESTART-%d]", i), s.sprintf("%7d", .FrequencyRestart))
        End With
    Next i

    ' Replace global replacements word stored under Settings
    For i = 1 To Settings.Replacements.Count
        InputString = Replace(InputString, Settings.Replacements.FindWord(i), Settings.Replacements.ReplaceWord(i))
    Next i
    
    ReplaceWords = InputString

End Function

Public Sub WriteToFile(ByRef ts As Object, ByVal wC As Boolean, Optional ByVal GenerateIncludeFile As Boolean = False)
' Action: Writes the component to a file stored in the TextStream-object
'
' Inputs:   ts = A TextStream object (included in the MS Scripting Runtime)
'           wC = writeComments, True if comments are to be written, false otherwise
'
    Dim i As Integer
    
    If Settings.Filename <> "" And GenerateIncludeFile = False Then
        WriteCustom ts, wC
        Exit Sub
    End If
    
    Dim t As New ResourceTablePrint
    t.SetDefaultValues
    t.OutputKeywords = IIf(GenerateIncludeFile = True, 1, 0)
    t.OutputKeywordBrackets = 1
    
    ' Set commonly used arguments
    Dim card As ResourceTablePrint
    Set card = t.Args(0, FStr:="%03d")
    
    
    t.AddLine "* Generated with text2relap {} {}", t.Args(HydroSystem.Inputdeck.Text2RelapVersion, Key:="TEXT2RELAPVERSION"), _
                                                   t.Args(HydroSystem.Inputdeck.Text2RelapBranch, Key:="TEXT2RELAPBRANCH")
    t.AddLine "* {} {}", t.Args(Format(Date, "YYYY-MM-dd"), Key:="CURRENTDATE"), t.Args(Format(Time, "hh:mm:ss"), Key:="CURRENTTIME")
    t.AddLine "*"
    t.AddLine "={}", t.Args(Settings.Title, Key:="TITLE")
    t.AddLine ""
    t.AddLine "* Developmental Model Control"
    If Settings.DevelopmentControl = "" Then
        t.AddLine "*001     0"
    Else
        t.AddLine "001 {}", t.Args(Settings.DevelopmentControl, Key:="DEVCONTROL")
    End If
    t.AddLine "*"
    t.AddLine "* Problem type and option"
    t.AddLine "100     new   transnt"
    t.AddLine ""
    t.AddLine "* Input check or run option"
    t.AddLine "101     run"
    t.AddLine ""
    t.AddLine "* Units selection: SI"
    t.AddLine "102    si si"
    t.AddLine ""
    t.AddLine "*      Refvol.    Elev.   Fluid      Name"
    t.OutputKeywordIndexes = 1
    For i = 1 To Settings.Refvols.Count
        With Settings.Refvols(i)
            If HydroSystem.ComponentExists(.CCC, .Volume) = False Then
                MsgBox "Error: Reference volume '" & .ReferenceVolume & "' doesn't exist.", vbCritical
            End If
            t.AddLine "{ }  {}  {}  {}  {}", card.Update(119 + i), t.Args(.ReferenceVolume, Key:="REFVOL"), t.Args(.Elevation, FStr:="%7.3f", Key:="ELEV"), t.Args(.Fluid, Key:="FLUIDTYPE"), t.Args(.SystemName, Key:="SYSTEMNAME")
        End With
    Next i
    t.OutputKeywordIndexes = 0
    t.AddLine ""
    t.AddLine "* CPU remaining card / Allocated CPU time"
    t.AddLine "105     10.0  20.0  360000."
    t.AddLine ""
    t.AddLine "110  nitrogen"
    t.AddLine "115  1.0"
    t.AddLine "*"
    t.AddLine "* Time step and plot frequency"
    t.AddLine "*       t.end     dt.min     dt.max  ssdtt  f.minor  f.major  f.restart"
    t.OutputKeywordIndexes = 1
    For i = 1 To Settings.Timesteps.Count
        With Settings.Timesteps(i)
            t.AddLine "{ }  {}  {}  {}  {}  {}  {}  {}", card.Update(200 + i), t.Args(.TimeEnd, FStr:="%8.3f", Key:="TIMEEND"), t.Args(.TimestepMin, FStr:="%9.3e", Key:="DTMIN"), _
                                                         t.Args(.TimestepMax, FStr:="%9.3e", Key:="DTMAX"), t.Args(.ControlOption, FStr:="%5s", Key:="SSDTT"), _
                                                         t.Args(.FrequencyMinor, FStr:="%5d", Key:="FMINOR"), t.Args(.FrequencyMajor, FStr:="%7d", Key:="FMAJOR"), _
                                                         t.Args(.FrequencyRestart, FStr:="%7d", Key:="FRESTART")

        End With
    Next i
    t.OutputKeywordIndexes = 0
    t.AddLine "*"
    t.AddLine "*"
    t.AddLine ""
    t.AddLine "20500000  " & IIf(Settings.CntrlvarFormat = 9999, "9999", "999")
    t.AddLine ""
    Select Case Settings.TripFormat
        Case 1
            t.AddLine "20600000 expanded"
        Case 2
            t.AddLine "20600000 expbig"
    End Select
    t.AddLine ""
    t.AddLine "* --------------------------------------------------"
    t.AddLine "*                       Trips"
    t.AddLine "* --------------------------------------------------"
    t.AddLine "*"
    ts.WriteLine t.WriteAndClear

End Sub

Private Sub WriteCustom(ByRef ts As Object, ByVal wC As Boolean)
' Action: Writes a custom header file
'
    ' Activate Microsoft Scripting runtime Tools...References...Microsoft scripting runtime
    On Error GoTo ErrorHandler
    Dim fso As Scripting.FileSystemObject
    Dim tsRead As Scripting.TextStream
    Dim tsWrite As Scripting.TextStream
    
    Dim Includefile As New ResourceFileObject
    Includefile.CreateByParts Settings.IncludeFolder, Settings.Filename
    
    Set fso = New Scripting.FileSystemObject
    
    ' If include file doesn't exist, create it
    If Includefile.FileExists = False Then
        Dim answ
        answ = MsgBox("Warning: Custom header file '" & Includefile.FullPath & "' doesn't exist. A generic header file will be created", vbExclamation + vbOKCancel, "Include file not found")
        HydroSystem.Inputdeck.Warnings = True
        If answ = vbCancel Then
            HydroSystem.Inputdeck.Abort = True
            Exit Sub
        End If
        If Includefile.FolderExists = False Then Includefile.CreateFolders ' Create subfolders if they don't exist
        Set tsWrite = fso.CreateTextFile(Includefile.FullPath, True)
        Me.WriteToFile tsWrite, wC, GenerateIncludeFile:=True
        tsWrite.Close
    End If
    
    ' Open include file for reading
    Set tsRead = fso.OpenTextFile(Includefile.FullPath, ForReading)
   
    ' Read all lines of the file and write them to file
    Do While tsRead.AtEndOfStream <> True
        ts.WriteLine ReplaceWords(tsRead.ReadLine)
    Loop
    
    tsRead.Close
    Set fso = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "Error writing/creating Custom header file " & Settings.Filename
    If Not tsWrite Is Nothing Then tsWrite.Close
    If Not tsRead Is Nothing Then tsRead.Close
End Sub

'
'
' STATIC METHODS


