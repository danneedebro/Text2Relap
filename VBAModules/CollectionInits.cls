VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionInits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' MEMBER VARIABLES
Private pInits As Collection
Private pCurrentInit As PropertiesHydroCompInit

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
    Set pInits = New Collection
    
    ' Initial conditions
    Set pCurrentInit = New PropertiesHydroCompInit
    pCurrentInit.EbtFlag = 103
    pCurrentInit.Pressure = 100000#
    pCurrentInit.Temperature = 293.15
End Sub

Public Sub Add(ByRef inputRows() As Variant)
    On Error GoTo ErrorHandler:
    Dim NewInit As New PropertiesHydroCompInit
    NewInit.Create inputRows, HydroSystem
     
    ' If no ID is given, change CurrentInit, otherwise add to Inits-collection
    If NewInit.ID = "" Then
        Set pCurrentInit = NewInit
    Else
        NewInit.ID = inputRows(1, 4)
        pInits.Add NewInit
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".Add, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Function Item(index As Variant) As PropertiesHydroCompInit
Attribute Item.VB_UserMemId = 0
    Set Item = pInits(index)
End Function

Public Function CurrentInit() As PropertiesHydroCompInit
    Set CurrentInit = pCurrentInit
End Function

Public Function Count() As Integer
    Count = pInits.Count
End Function

Public Sub AssignValues()
' Action: Assigns custom initialisation to selected components
'
    On Error GoTo ErrorHandler:
    Dim i As Integer, j As Integer, k As Integer
    Dim LookForID As String
    Dim CurrHydroComp As ComponentHydro
    
    For i = 1 To Me.Count
        LookForID = Me(i).ID
        For j = 1 To HydroSystem.Components.Count
            If HydroSystem.Components(j).ObjectType <> HydroComp Then GoTo Next_j
            Set CurrHydroComp = HydroSystem.Components(j)
            With CurrHydroComp
                For k = 1 To .Segments.Count
                    If .Segments(k).ID = LookForID Then
                        Set .Segments(k).Init = Me(i)
                    End If
                Next k
            End With
Next_j:
        Next j
    Next i
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".AssignValues, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
'
'
' CLASS METHODS

