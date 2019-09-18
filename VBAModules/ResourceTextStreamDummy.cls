VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResourceTextStreamDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ResourceTextStreamDummy.cls
'
' Action: Mimics the methods in a text stream object in order to get the same output to something else than
'         a text file (for instance a msgbox or a textbox)
'
Option Explicit

Private pTextStream As String

Public Property Get TextStream() As String
    TextStream = pTextStream
End Property

Public Sub WriteLine(ByVal InputString As String)
    pTextStream = pTextStream & InputString & vbNewLine
End Sub

Public Sub WriteBlankLines(ByVal NumberOfBlankLines As Integer)
    pTextStream = pTextStream & String(NumberOfBlankLines, vbNewLine)
End Sub

Public Sub Clear()
    pTextStream = ""
End Sub
