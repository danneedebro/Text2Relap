VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComponentComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ComponentComment.cls
'
'
Option Explicit

Implements iComponent
'
'
' MEMBER VARIABLES
Public Text As String
Public CommentType As Integer   ' 1 = New flow path comment,  2 = More discrete comment,  3 = Exact string that was put in
Private pRowBegin As Integer
Private pRowEnd As Integer

'
'
' SHARED PROPERTIES AND METHODS
Private Property Get iComponent_ObjectType() As R5Component
    If Me.CommentType = 1 Then
        iComponent_ObjectType = Comment1
    ElseIf Me.CommentType = 2 Then
        iComponent_ObjectType = Comment2
    Else
        iComponent_ObjectType = Comment3
    End If
End Property

Private Function iComponent_ComponentInfo() As String
    iComponent_ComponentInfo = "Comment"
End Function

Private Property Get iComponent_RowBegin() As Integer
    iComponent_RowBegin = pRowBegin
End Property

Private Property Get iComponent_RowEnd() As Integer
    iComponent_RowEnd = pRowEnd
End Property

Private Property Get iComponent_Source() As InputObject
    Set iComponent_Source = New InputObject
End Property

Private Sub iComponent_WriteToFile(ts As Object, ByVal wC As Boolean)
    WriteToFile ts, wC
End Sub
'
'
' PROPERTIES

'
'
' CLASS EVENT METHODS
Public Sub Create(InputData As InputObject)
    On Error GoTo ErrorHandler
    pRowBegin = InputData.RowFirst
    pRowEnd = InputData.RowLast
    If Left(InputData(1, 1), 2) = "**" Then
        CommentType = 2
        Text = Right(InputData(1, 1), Len(InputData(1, 1)) - 2)
    ElseIf Left(InputData(1, 1), 1) = "*" Then
        CommentType = 1
        Text = Right(InputData(1, 1), Len(InputData(1, 1)) - 1)
    Else
        CommentType = 3
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".Create, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
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
    On Error GoTo ErrorHandler

    If CommentType = 1 Then
        ts.WriteBlankLines 6
        ts.WriteLine "*"
        ts.WriteLine "*  " & Me.Text
        ts.WriteLine "*"
        ts.WriteLine "*"
    ElseIf CommentType = 2 Then
        ts.WriteLine "*-------------------------------------------------------------------"
        ts.WriteLine "*  " & Me.Text
        ts.WriteLine "*-------------------------------------------------------------------"
    ElseIf CommentType = 3 Then
        ts.WriteLine Me.Text
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".WriteToFile, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

