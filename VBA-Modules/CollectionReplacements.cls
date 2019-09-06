VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionReplacements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Private pReplacementsFind As Collection
Private pReplacementsReplace As Collection
Private pHydroSystem As HydroSystem
'
'
' PROPERTIES
Public Property Get HydroSystem() As HydroSystem
    Set HydroSystem = pHydroSystem
End Property

Public Property Set HydroSystem(arg As HydroSystem)
    Set pHydroSystem = arg
End Property
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set pReplacementsFind = New Collection
    Set pReplacementsReplace = New Collection
End Sub

Public Sub Add(ByRef inputRows() As Variant)
    On Error GoTo ErrorHandler:

    ' Parse text
    Dim s As New ResourceStringManipulation
    Dim Word2 As Variant
    Dim ErrorMessages As String
    
    Word2 = s.ParseBrackedExpression(inputRows(1, 2), Numeric:=False)
    
    ' If second bit is set input is not a multiple of two
    If s.CheckBitSet(Word2(0), 2) = True Then ErrorMessages = ErrorMessages & "Input Error: Input is Not a multiple of two." & vbNewLine & vbNewLine
    'If s.CheckBitSet(word20(0), 3) = True Then ErrorMessages = ErrorMessages & "Input Error: Wrong format." & vbNewLine & vbNewLine
    
    ' If input errors occured
    If ErrorMessages <> "" Then
        HydroSystem.Inputdeck.HighlightCurrentLines 2
        Dim answ
        answ = MsgBox(ErrorMessages & vbNewLine & "Continue?", vbExclamation + vbOKCancel)
        If answ = vbCancel Then
            HydroSystem.Inputdeck.Abort = True
            Exit Sub
        End If
    End If
    
    ' Append read list to current list
    Dim i As Integer
    For i = LBound(Word2(1)) To UBound(Word2(1))
        pReplacementsFind.Add Word2(1)(i)
        pReplacementsReplace.Add Word2(2)(i)
    Next i
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".Add, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function FindWord(index As Variant) As String
    FindWord = pReplacementsFind(index)
End Function

Public Function ReplaceWord(index As Variant) As String
    ReplaceWord = pReplacementsReplace(index)
End Function

Public Function Count() As Integer
    Count = pReplacementsFind.Count
End Function
'
'
' CLASS METHODS

'
'
' STATIC METHODS

