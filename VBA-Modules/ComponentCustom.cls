VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComponentCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ComponentCustom.cls
'
'
Option Explicit

Implements iComponent
'
'
' MEMBER VARIABLES
Public Filename As String
Public LookForID As String

Public HydroSystem As HydroSystem
Public Init As PropertiesHydroCompInit

Private pRowBegin As Integer
Private pRowEnd As Integer
'
'
' SHARED PROPERTIES AND METHODS
Private Property Get iComponent_ObjectType() As R5Component
    iComponent_ObjectType = Custom
End Property

Private Function iComponent_ComponentInfo() As String
    iComponent_ComponentInfo = "Custom component"
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
Public Property Get Inputdeck() As Text2Relap
    Set Inputdeck = HydroSystem.Inputdeck
End Property

Public Property Get Settings() As PropertiesSettings
    Set Settings = HydroSystem.Settings
End Property

'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    'Set HydroSystem = New HydroSystem
End Sub

Public Sub Create(inputRows() As Variant, ByRef System As HydroSystem, ByVal RowBegin As Integer)
    On Error GoTo ErrorHandler
    Filename = inputRows(1, 2)
    LookForID = inputRows(1, 3)
    
    pRowBegin = RowBegin
    pRowEnd = RowBegin
    
    Set Me.HydroSystem = System
    Set Init = Settings.Inits.CurrentInit
    
    Exit Sub
ErrorHandler:
    err.Raise err.Number, TypeName(Me) & ".Create, " & err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub
'
'
' CLASS METHODS
Public Sub WriteToFile(ByRef ts As Object, ByVal wC As Boolean)
' Action: Writes the component to a file stored in the TextStream-object
'
' Inputs:   ts = A TextStream object (included in the MS Scripting Runtime)
'           wC = writeComments, True if comments are to be written, false otherwise
'
' Action: Writes a custom component
'
    On Error GoTo ErrorHandler

    ' Activate Microsoft Scripting runtime Tools...References...Microsoft scripting runtime
    Dim fso As Scripting.FileSystemObject
    Dim tsRead As Scripting.TextStream
    Dim tsWrite As Scripting.TextStream
    
    Dim Includefile As New ResourceFileObject
    Includefile.CreateByParts HydroSystem.Settings.IncludeFolder, Me.Filename
    
    Set fso = New Scripting.FileSystemObject
    
    ' If include file doesn't exist, create it
    If Includefile.FileExists = False Then
        Dim answ
        answ = MsgBox("Warning: Custom component " & vbNewLine & Chr(34) & Includefile.FullPath & Chr(34) & vbNewLine & " doesn't exist. An empty custom component includefile will be created", vbExclamation + vbOKCancel, "Include file not found")
        Inputdeck.Warnings = True
        If answ = vbCancel Then
            Inputdeck.Abort = True
            Exit Sub
        End If
        If Includefile.FolderExists = False Then Includefile.CreateFolders ' Create subfolders if they don't exist
        Set tsWrite = fso.CreateTextFile(Includefile.FullPath, True)

        tsWrite.WriteLine "* " & Me.Filename
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
    MsgBox "Error writing/creating Custom component " & Me.Filename
    If Not tsWrite Is Nothing Then tsWrite.Close
    If Not tsRead Is Nothing Then tsRead.Close
End Sub
'
'
' STATIC METHODS
Private Function ReplaceWords(ByVal InputString As String) As String
' Action: Replaces certain words with corresponding values
'
'
    ' Replace global replacements word stored under Settings
    Dim i As Integer
    For i = 1 To Settings.Replacements.Count
        InputString = Replace(InputString, Settings.Replacements.FindWord(i), Settings.Replacements.ReplaceWord(i))
    Next i
    
    ReplaceWords = InputString

End Function
