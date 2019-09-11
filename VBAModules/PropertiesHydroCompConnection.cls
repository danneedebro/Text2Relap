VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertiesHydroCompConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PropertiesHydroCompConnection.cls
'
'
Option Explicit
'
'
' MEMBER VARIABLES
Public CCC As Integer
Public Volume As Integer
Public Side As Integer
Private pID As String
'
'
' PROPERTIES
Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(arg As String)
    If Left(arg, 1) = Chr(34) And Right(arg, 1) = Chr(34) And Len(arg) = Len(Chr(34) & "CCCVV000F" & Chr(34)) Then
        On Error GoTo ErrorHandler
        Me.CCC = CInt(Mid(arg, 2, 3))
        Me.Volume = CInt(Mid(arg, 5, 2))
        Me.Side = CInt(Mid(11, 1))
    Else
        pID = arg
    End If
    Exit Property
ErrorHandler:
    MsgBox "Error"
End Property

'
'
' CLASS EVENT METHODS

'
'
' CLASS METHODS
Public Function ConnectionCode() As String
    ConnectionCode = Format(CCC, "#000") & Format(Volume, "#00") & Format(Side, "#0000")
End Function
'
'
' STATIC METHODS









