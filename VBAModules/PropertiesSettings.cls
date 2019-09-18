VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertiesSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PropertiesSettings.cls
'
'
Option Explicit
'
' MEMBER VARIABLES
Private pHydroSystem As HydroSystem

Public Title As String
Public DevelopmentControl As String
Public Filename As String
Private pIncludeFolder As String
Public WriteComments As Boolean

Public ForceCalcMethod As Integer

' Variables relating to CCC (Hydro comp number)
Private pCurrentCCC As Integer
Private pCCCIncrement As Integer

' Variables relating to cntrlvars
Private pCurrentCntrlvarNum As Integer
Private pCntrlvarFormat As Integer

' Variables relating to trips
Private pTripFormat As Integer
Private pCurrentTripNumVariable As Long
Private pCurrentTripNumLogical As Long
Private pTripIncrement As Integer

' Variables relating to tables
Private pCurrentTableNum As Integer

' Variables relating to Replacement words
Private pReplacements As CollectionReplacements

' Variables relating to initialization
Private pInits As CollectionInits

Private pTimesteps As CollectionTimesteps
Private pRefvols As CollectionRefvols
'
'
' PROPERTIES
Public Property Get HydroSystem() As HydroSystem
    Set HydroSystem = pHydroSystem
End Property

Public Property Get Refvols() As CollectionRefvols
    Set Refvols = pRefvols
End Property

Public Property Get Timesteps() As CollectionTimesteps
    Set Timesteps = pTimesteps
End Property

Public Property Get Replacements() As CollectionReplacements
    Set Replacements = pReplacements
End Property

Public Property Get Inits() As CollectionInits
    Set Inits = pInits
End Property

Public Property Get IncludeFolder() As String
    IncludeFolder = pIncludeFolder
End Property

Public Property Let IncludeFolder(arg As String)
    Dim f As New ResourceFileObject
    
    If InStr(1, arg, ":") > 0 Then
        pIncludeFolder = arg
    Else
        f.CreateByParts ThisWorkbook.Path, arg
        pIncludeFolder = f.FullPath
    End If
End Property

Public Property Get TripFormat() As Integer
    TripFormat = pTripFormat
End Property

Public Property Let TripFormat(arg As Integer)
    If arg = 0 Then  ' Normal trip numbers (401-599 Variable trips, 601-799 Logical trips)
        pCurrentTripNumVariable = 405
        pCurrentTripNumLogical = 601
        pTripIncrement = 1
        pTripFormat = 0
    ElseIf arg = 1 Then
        pCurrentTripNumVariable = 50
        pCurrentTripNumLogical = 10010
        pTripIncrement = 10
        pTripFormat = 1
    ElseIf arg = 2 Then
        pCurrentTripNumVariable = 5
        pCurrentTripNumLogical = 10001
        pTripIncrement = 1
        pTripFormat = 2
    Else
        MsgBox "Error: Incorrect tripformat (0=Normal, 1=EXPANDED, 2=EXPBIG). Using 0=Normal", vbInformation
        pCurrentTripNumVariable = 405
        pCurrentTripNumLogical = 601
        pTripIncrement = 1
    End If
End Property

Public Property Get CurrentTripNumVariable() As Long
    CurrentTripNumVariable = pCurrentTripNumVariable
End Property

Public Property Let CurrentTripNumVariable(arg As Long)
    pCurrentTripNumVariable = arg
End Property

Public Property Get CurrentTripNumLogical() As Long
    CurrentTripNumLogical = pCurrentTripNumLogical
End Property

Public Property Let CurrentTripNumLogical(arg As Long)
    pCurrentTripNumLogical = arg
End Property

Public Property Get CntrlvarFormat() As Integer
    CntrlvarFormat = pCntrlvarFormat
End Property

Public Property Let CntrlvarFormat(arg As Integer)
    If arg = 999 Then
        pCntrlvarFormat = 999
        
    ElseIf arg = 9999 Then
        pCntrlvarFormat = 9999
    Else
        MsgBox "Error: Incorrect cntrlvar format (999=Normal (1-999), 9999=EXPANDED (1-9999). Using 999=Normal", vbInformation
        pCntrlvarFormat = 9999
    End If
End Property

Public Property Get CCCIncrement() As Integer
    CCCIncrement = pCCCIncrement
End Property

Public Property Get CurrentCCC() As Integer
    CurrentCCC = pCurrentCCC
End Property

Public Property Let CurrentCCC(arg As Integer)
    If arg = 0 Then ' If no number is given, round up until the nearest five
        pCurrentCCC = Round((pCurrentCCC + 5) / 10 + 0.5, 0) * 10
    Else
        pCurrentCCC = arg
    End If
End Property

Public Property Get CurrentCntrlvarNum() As Integer
    CurrentCntrlvarNum = pCurrentCntrlvarNum
End Property

Public Property Let CurrentCntrlvarNum(arg As Integer)
    pCurrentCntrlvarNum = arg
End Property

Public Property Let CntrlvarStart(arg As Integer)
    pCurrentCntrlvarNum = arg
End Property

Public Property Get CurrentTableNum() As Integer
    CurrentTableNum = pCurrentTableNum
End Property


'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set pTimesteps = New CollectionTimesteps
    
    Set pRefvols = New CollectionRefvols
    
    Set pReplacements = New CollectionReplacements
    
    Set pInits = New CollectionInits
    
    ' Set default values
    With Me
        .IncludeFolder = "include\"
        .CntrlvarFormat = 999
        .CntrlvarStart = 200
        .TripFormat = 0
        .WriteComments = True
        .CurrentTripNumLogical = 601
        .CurrentTripNumVariable = 405
    End With
    pCurrentTableNum = 1
    pCCCIncrement = 1
    
End Sub

Public Sub Create(ByRef ParentSystem As HydroSystem)
' Action: Instantiate Settings object and newly created custom collection objects with a reference to
'         the current hydrosystem
'
    Set pHydroSystem = ParentSystem
    Set pTimesteps.HydroSystem = ParentSystem
    Set pRefvols.HydroSystem = ParentSystem
    Set pReplacements.HydroSystem = ParentSystem
    Set pInits.HydroSystem = ParentSystem
End Sub
'
'
' CLASS METHODS
Public Sub IncrementCCC(Optional ByVal Decrement As Boolean = False)
' Action: Increments the current component number (CCC) for hydrodynamic components
'
    If Decrement = False Then
        pCurrentCCC = pCurrentCCC + pCCCIncrement
    Else
        pCurrentCCC = pCurrentCCC - pCCCIncrement
    End If
End Sub

Public Sub IncrementTripVariable()
' Action: Increments the current tripnumber for variable trips
'
    pCurrentTripNumVariable = pCurrentTripNumVariable + pTripIncrement
End Sub

Public Sub IncrementTripLogical()
' Action: Increments the current tripnumber for logical trips
'
    pCurrentTripNumLogical = pCurrentTripNumLogical + pTripIncrement
End Sub

Public Sub IncrementCntrlvar()
' Action: Increments the current cntrlvarnumber
    pCurrentCntrlvarNum = pCurrentCntrlvarNum + 1
End Sub

Public Sub IncrementTableNum()
' Action: Increments the current cntrlvarnumber
    pCurrentTableNum = pCurrentTableNum + 1
End Sub
'
'
' STATIC METHODS




