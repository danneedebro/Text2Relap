VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComponentForce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ComponentForce.cls
'
'
Option Explicit

Implements iComponent

'
'
' MEMBER VARIABLES
Private pComponentIndexes As Collection
Private pSegmentIndexes As Collection
Private pForceNumber As Integer

Private pRowBegin As Integer
Private pRowEnd As Integer

Private pHydroSystem As HydroSystem
'
'
' SHARED PROPERTIES AND METHODS
Private Property Get iComponent_ObjectType() As R5Component
    iComponent_ObjectType = Force
End Property

Private Function iComponent_ComponentInfo() As String
    iComponent_ComponentInfo = "Force"
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

'Private Function iComponent_ComponentInfo() As String
'    iComponent_ComponentInfo = "Force"
'End Function

'Private Sub iComponent_WriteToFile(ts As Object, ByVal wC As Boolean)
'    WriteToFile ts, wC
'End Sub
'
'
' PROPERTIES
Public Property Get HydroSystem() As HydroSystem
    Set HydroSystem = pHydroSystem
End Property

Public Property Get Settings() As PropertiesSettings
    Set Settings = HydroSystem.Settings
End Property

Public Property Get ForceNumber() As Integer
    ForceNumber = pForceNumber
End Property

Public Property Let ForceNumber(arg As Integer)
    pForceNumber = arg
End Property

Public Property Get ComponentIndexes() As Collection
    Set ComponentIndexes = pComponentIndexes
End Property

Public Property Get SegmentIndexes() As Collection
    Set SegmentIndexes = pSegmentIndexes
End Property
'
'
' CLASS EVENT METHODS
Private Sub Class_Initialize()
    Set pComponentIndexes = New Collection
    Set pSegmentIndexes = New Collection
End Sub

Public Sub Create(ByRef ParentSystem As HydroSystem)
    On Error GoTo ErrorHandler
    Set pHydroSystem = ParentSystem
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".Create, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
'
'
' CLASS METHODS
Public Sub WriteToFile(ByRef ts As Object, ByVal wC As Boolean)
' Action: Writes the force components
'
' Inputs:   ts = A TextStream object (included in the MS Scripting Runtime)
'           wC = writeComments, True if comments are to be written, false otherwise
'
    On Error GoTo ErrorHandler

    Dim i As Integer, j As Integer, k As Integer, m As Integer
    Dim s As New ResourceStringManipulation
    Dim CurrForce As Integer
    Dim Plotvars As Collection, Plotnums As Collection, Coeffs As Collection
    Dim CurrHydroComp As ComponentHydro
    Dim ConnectingJunctionComponent As ComponentHydro
    Dim CurrConnection As PropertiesHydroCompConnection
    
    Set Plotvars = New Collection
    Set Plotnums = New Collection
    Set Coeffs = New Collection
    
    Dim ComponentIndex As Integer
    Dim SegmentIndex As Integer
    
    For i = 1 To Me.ComponentIndexes.Count
        ComponentIndex = ComponentIndexes(i)
        SegmentIndex = SegmentIndexes(i)
        
        Set CurrHydroComp = HydroSystem.Components(ComponentIndex)
        
    
        For j = CurrHydroComp.Segments(SegmentIndex).JunctionFirst To CurrHydroComp.Segments(SegmentIndex).JunctionLast
            Plotvars.Add "mflowj"
            Plotnums.Add Format(CurrHydroComp.CCC, "000") & Format(j, "00") & "0000"
            Coeffs.Add CurrHydroComp.Segments(SegmentIndex).Dx
        Next j
        
        ' If number of junctions less than number of volumes, look for the junction that connects to
        ' the last volume
        If CurrHydroComp.Segments(SegmentIndex).Volumes <> CurrHydroComp.Segments(SegmentIndex).Junctions Then
            Dim LookForVolume As Integer, LookForCCC As Integer
            LookForCCC = CurrHydroComp.CCC
            LookForVolume = CurrHydroComp.Segments(SegmentIndex).VolumeLast
            
            For k = 1 To HydroSystem.Components.Count
                If HydroSystem.Components(k).ObjectType <> HydroComp Then GoTo Next_k
                Set ConnectingJunctionComponent = HydroSystem.Components(k)
                For m = 1 To ConnectingJunctionComponent.Connections.Count
                    If ConnectingJunctionComponent.Connections(m).CCC = LookForCCC And ConnectingJunctionComponent.Connections(m).Volume = LookForVolume Then
                        Plotvars.Add "mflowj"
                        If ConnectingJunctionComponent.Info.MainType = pump Then
                            Plotnums.Add Format(ConnectingJunctionComponent.CCC, "000") & Format(m, "00") & "0000"
                        Else
                            Plotnums.Add Format(ConnectingJunctionComponent.CCC, "000") & "000000"
                        End If
                        Coeffs.Add -(-1) ^ m * CurrHydroComp.Segments(SegmentIndex).Dx  ' Add set coeff negative if a connects-to-junction
                    End If
                Next m
Next_k:
            Next k
        End If
 
    Next i
    
    Dim CntrlvarString As String
    CntrlvarString = ""
    
    CntrlvarString = CntrlvarString & CntrlvarSum(Format(ForceNumber, "\I\D-0000"), 1#, Plotvars, Plotnums, Coeffs)
    Set Plotvars = New Collection
    Set Plotnums = New Collection
    Set Coeffs = New Collection
    Plotvars.Add "cntrlvar"
    Plotnums.Add Settings.CurrentCntrlvarNum - 1
    Coeffs.Add 1#
    Plotvars.Add "cntrlvar"
    Plotnums.Add Settings.CurrentCntrlvarNum
    Coeffs.Add -1#
    CntrlvarString = CntrlvarString & CntrlvarSum(Format(ForceNumber, "\I\D-0000"), 1#, Plotvars, Plotnums, Coeffs)
    CntrlvarString = CntrlvarString & CntrlvarDiv(Format(ForceNumber, "\I\D-0000"), 1#, "cntrlvar", _
                                                  CStr(Settings.CurrentCntrlvarNum - 1), "time", "0", Cntrlvar:=ForceNumber)
    Settings.IncrementCntrlvar
    Settings.IncrementCntrlvar
    ts.WriteLine CntrlvarString

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, TypeName(Me) & ".WriteToFile, " & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub WriteForcesDiff()
    
End Sub
'
'
' STATIC METHODS
Private Function CntrlvarSum(ByVal Name As String, ByVal Coeff As Double, ByVal Plotvars As Collection, ByVal Plotnums As Collection, ByVal Coeffs As Collection) As String
' Action: Returns a string with a Sum operator
'
    Dim s As New ResourceSprintf
    Dim OutputStr As String
    Dim Cntrlvar As Integer
    Dim RowsMax As Integer
    Dim FormatStr As String
    
    If Plotvars.Count <> Plotnums.Count Or Plotvars.Count <> Coeffs.Count Or Plotnums.Count <> Coeffs.Count Then
        MsgBox "Error in ComponentForce.CntrlvarSum: Number of elements of Plotvars, Plotnums and Coeffs not equal", vbCritical
        CntrlvarSum = "<ERROR>"
        Exit Function
    End If
    
    Cntrlvar = Settings.CurrentCntrlvarNum
    If Settings.CntrlvarFormat = 999 Then
        RowsMax = 20
        FormatStr = "205%03d%02d"
    Else
        RowsMax = 8
        FormatStr = "205%04d%01d"
    End If
    
    
    OutputStr = OutputStr & s.sprintf(FormatStr & " %-8s sum  1.000 0.0 0\n", Cntrlvar, 0, Name)
    
    Dim cnt As Integer
    cnt = 0
    Do While True
        cnt = cnt + 1
        If Plotvars.Count = 0 Or Plotnums.Count = 0 Or Coeffs.Count = 0 Then
            Settings.IncrementCntrlvar
            CntrlvarSum = OutputStr
            Exit Do
        End If
        
        If cnt = 1 Then
            OutputStr = OutputStr & s.sprintf(FormatStr & " 0.000000 % 7.5f  %-8s  %s\n", Cntrlvar, cnt, Coeffs(1), Plotvars(1), Plotnums(1))
        Else
            OutputStr = OutputStr & s.sprintf(FormatStr & "          % 7.5f  %-8s  %s\n", Cntrlvar, cnt, Coeffs(1), Plotvars(1), Plotnums(1))
        End If
        Plotvars.Remove 1
        Plotnums.Remove 1
        Coeffs.Remove 1
        
        If cnt >= RowsMax And Plotvars.Count > 0 Then
            Plotvars.Add "cntrlvar", before:=1
            Plotnums.Add Format(Cntrlvar), before:=1
            Coeffs.Add 1#, before:=1
                    
            Settings.IncrementCntrlvar
            OutputStr = OutputStr & CntrlvarSum(Name, 1#, Plotvars, Plotnums, Coeffs)
            Exit Do
        End If
    Loop
    
    CntrlvarSum = OutputStr
    
End Function

Private Function CntrlvarDiv(ByVal Name As String, Coeff As Double, ByVal NumerPlotvar As String, ByVal NumerPlotnum As String, _
                             ByVal DenomPlotvar As String, ByVal DenomPlotnum As String, Optional ByVal Cntrlvar As Integer = -1) As String
    Dim s As New ResourceSprintf
    Dim OutputStr As String
    Dim FormatStr As String
    
    ' If no cntrlvar specified, use next number
    If Cntrlvar = -1 Then
        Cntrlvar = Settings.CurrentCntrlvarNum
        Settings.IncrementCntrlvar
    End If
    If Settings.CntrlvarFormat = 999 Then FormatStr = "205%03d%02d" Else FormatStr = "205%04d%01d"
    
    OutputStr = OutputStr & s.sprintf(FormatStr & " %-8s div %7.5f 0.0 0\n", Cntrlvar, 0, Name, Coeff)
    OutputStr = OutputStr & s.sprintf(FormatStr & " %s   %s  %s  %s\n", Cntrlvar, 1, DenomPlotvar, DenomPlotnum, NumerPlotvar, NumerPlotnum)
    
    CntrlvarDiv = OutputStr
    
End Function





