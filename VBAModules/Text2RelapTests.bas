Attribute VB_Name = "Text2RelapTests"
Option Explicit

Private Sub TestInputObject()
    
    Dim InputData As New InputObject
    
    
    Debug.Assert InputData.NumberOfRows = 1
    Debug.Assert InputData.NumberOfColumns = 22
    
    ' Create an object from Range
    InputData.Create Range("A5:C20")
    Debug.Assert InputData.RowFirst = 5
    Debug.Assert InputData.RowLast = 20
    Debug.Assert InputData.NumberOfRows = 16
    Debug.Assert InputData.NumberOfColumns = 3
    
    ' Create an child object from this one
    Dim InputDataChild As InputObject
    Set InputDataChild = InputData.MakeCopy(OnlyRowN:=2) ' Row A6:C6
    Debug.Assert InputDataChild.RowFirst = 6
    Debug.Assert InputDataChild.RowLast = 6
    Debug.Assert InputDataChild(1, 1) = InputData(2, 1)
    
    Set InputDataChild = InputData.MakeCopy(FirstRow:=2) ' Row A6:C20
    Debug.Assert InputDataChild.RowFirst = 6
    Debug.Assert InputDataChild.RowLast = 20
    
    Set InputDataChild = InputData.MakeCopy(LastRow:=15) ' Row A5:C19
    Debug.Assert InputDataChild.RowFirst = 5
    Debug.Assert InputDataChild.RowLast = 19
    
    ' Create an empty
    Set InputDataChild = New InputObject
    InputDataChild.CreateFromParts InputData.SheetName, 1, 50
    Debug.Assert InputDataChild.RowFirst = 1
    Debug.Assert InputDataChild.RowLast = 50
    Debug.Assert InputDataChild.NumberOfRows = 1
    
    ' Create copy of child
    Dim InputDataChild2 As InputObject
    Set InputDataChild2 = InputDataChild.MakeCopy
    Debug.Assert InputDataChild2.RowFirst = 1
    Debug.Assert InputDataChild2.RowLast = 50
    InputDataChild2.SetDataFromWords Word2:="Hello"
    Debug.Assert InputDataChild2(1, 2) = "Hello"
    
End Sub


Private Sub TestHydroComponents()
    Dim InputDeck As New Text2Relap
    InputDeck.Create ThisWorkbook.ActiveSheet.Name
    
    Dim HydroSystem As New HydroSystem
    HydroSystem.Create InputDeck
    
    Dim NewJunction As New ComponentHydro
    
    Dim InputData As New InputObject
    InputData.CreateFromParts "", 15, 15
    InputData.SetDataFromWords Word1:="Junction"
    
    NewJunction.Create InputData, HydroSystem


End Sub
