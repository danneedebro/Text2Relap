Attribute VB_Name = "Modul1"


' TODO: Implementera Components
'           Components.Add Type, inputrows
' TODO: Utöka iComponents med Create-funktion
' Properties i HydroSystem:  HydroComponent("CCC") som hittar
'                            Trip("ID",number)
'                            Force(number)
'                            Force

Sub Test()

    Dim t As New HydroSystem
    't.Create
    'Set t.Inputdeck = New Text2Relap
    'Dim wrongArr(1, 22) As String
    
    'Dim inputRow() As Variant
    'inputRow = t.NewInputRow(Word1:="*")
    't.NewInputRow inputRow, Word2:="HEJ"
    't.NewInputRow wrongArr, Word2:="HEJ"
    
    't.AddComponent t.NewInputRow(Word1:="Pipe"), 3, 3
        
    Dim apa As New Collection
    apa.Add 5, "Hej"
    
    Debug.Print apa("Hej")
    
    
    'Dim Components As New CollectionComponents
    'Components.Add New HydroSystem, "Apa"
    'Debug.Print Components("Apa").ObjectType
    'tv(1).

End Sub


Sub Iterate()
    Range("H4").FormulaLocal = "=sin(1)"
End Sub



Sub TestScope()

    Dim a As New PropertiesSettingsTimestep
    a.TimeEnd = 1.111
    
    Debug.Print a.TimeEnd
    
    Debug.Print Hej(a)
    
    Debug.Print a.TimeEnd
    
    
    Dim coll As New Collection
    coll.Add 4, "Item1"
    coll.Add 5, "Item1"

End Sub


Function Hej(ByVal Tstep As PropertiesSettingsTimestep)
    
    Dim apa As New PropertiesSettingsTimestep
    Set apa = Tstep
    
    apa.TimeEnd = 12#
    Hej = 5

End Function
