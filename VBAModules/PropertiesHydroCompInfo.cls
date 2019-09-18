VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertiesHydroCompInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PropertiesHydroCompInfo.cls
'
'
Option Explicit
'
'
' MEMBER VARIABLES
Public Enum ComponentFamily
    JunctionComponent
    PipeComponent
    SingleVolumeComponent
End Enum



Public Enum ComponentType
    snglvol
    tmdpvol
    sngljun
    tmdpjun
    pipe
    annulus     ' Not implemented
    prizer      ' Not implemented
    canchan     ' Not implemented
    Branch      ' Not implemented
    separatr    ' Not implemented
    jetmixer    ' Not implemented
    turbine     ' Not implemented
    eccmix      ' Not implemented
    valve
    pump
    mtpljun     ' Not implemented
    accum       ' Not implemented
    deleteComp  ' Not implemented
End Enum

Public Enum ComponentSubType
    None
    mtrvlv
    trpvlv
    chkvlv
    inrvlv
    srvvlv
    endvlv
    rlfvlv
End Enum

Private pFamily As ComponentFamily
Private pMainType As ComponentType
Private pSubType As ComponentSubType
Private pName As String
Public UseIncludeFile As Boolean
Public IncludeFilename As String
'
'
' PROPERTIES
Public Property Get Family() As ComponentFamily
    Family = pFamily
End Property

Public Property Let Family(arg As ComponentFamily)
    pFamily = arg
End Property

Public Property Get MainType() As ComponentType
    MainType = pMainType
End Property

Public Property Let MainType(arg As ComponentType)
    pMainType = arg
End Property

Public Property Get MainTypeStr() As String
    On Error GoTo ErrorHandler
    MainTypeStr = Switch(pMainType = pipe, "pipe", pMainType = snglvol, "snglvol", pMainType = tmdpvol, "tmdpvol", _
                         pMainType = sngljun, "sngljun", pMainType = valve, "valve", pMainType = pump, "pump", _
                         pMainType = tmdpjun, "tmdpjun")
    Exit Property
ErrorHandler:
    MainTypeStr = "unknown"
End Property

Public Property Get SubType() As ComponentSubType
    SubType = pSubType
End Property

Public Property Let SubType(arg As ComponentSubType)
    pSubType = arg
End Property

Public Property Get SubTypeStr() As String
    On Error GoTo ErrorHandler
    SubTypeStr = Switch(pSubType = chkvlv, "chkvlv", pSubType = endvlv, "endvlv", pSubType = inrvlv, "inrvlv", _
                        pSubType = mtrvlv, "mtrvlv", pSubType = None, "none", pSubType = rlfvlv, "rlfvlv", _
                        pSubType = srvvlv, "srvvlv", pSubType = trpvlv, "trpvlv")
    Exit Property
ErrorHandler:
    SubTypeStr = "unknown"
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(arg As String)
    Dim Pos As Integer
    Pos = InStr(1, arg, ".inc")
    If Pos > 0 Then
        UseIncludeFile = True
        IncludeFilename = arg
        pName = Left(arg, Pos - 1)
    Else
        UseIncludeFile = False
        IncludeFilename = ""
        pName = arg
    End If
End Property
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    
End Sub

Public Sub SetType(ByVal Word1 As Variant)
    Select Case LCase(Word1)
        Case "pipe"
            Family = PipeComponent
            MainType = pipe
            SubType = None
        Case "tmdpvol"
            Family = SingleVolumeComponent
            MainType = tmdpvol
            SubType = None
        Case "snglvol"
            Family = SingleVolumeComponent
            MainType = snglvol
            SubType = None
        Case "junction", "sngljun"
            Family = JunctionComponent
            MainType = sngljun
            SubType = None
        Case "mtrvlv"
            Family = JunctionComponent
            MainType = valve
            SubType = mtrvlv
        Case "trpvlv"
            Family = JunctionComponent
            MainType = valve
            SubType = trpvlv
        Case "inrvlv"
            Family = JunctionComponent
            MainType = valve
            SubType = inrvlv
        Case "chkvlv"
            Family = JunctionComponent
            MainType = valve
            SubType = chkvlv
        Case "srvvlv"
            Family = JunctionComponent
            MainType = valve
            SubType = srvvlv
        Case "endvalve"
            Family = JunctionComponent
            MainType = valve
            SubType = endvlv
        Case "rlfvlv"
            Family = JunctionComponent
            MainType = valve
            SubType = rlfvlv
        Case "pump"
            Family = JunctionComponent
            MainType = pump
            SubType = None
        Case "tmdpjun"
            Family = JunctionComponent
            MainType = tmdpjun
            SubType = None
        Case Else
        
    End Select
End Sub
'
'
' CLASS METHODS

'
'
' STATIC METHODS

