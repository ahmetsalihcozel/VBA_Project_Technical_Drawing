Public Sub AddNewLayer()

Dim NewSymbols As AcadLayer
Dim OldSymbols As AcadLayer
Dim SiralamaObje As AcadEntity

On Error Resume Next

ThisDrawing.Layers.Add ("Old_Symbols")
ThisDrawing.Layers.Add ("New_Symbols")


ThisDrawing.ActiveLayer = ThisDrawing.Layers("New_Symbols")
ThisDrawing.ActiveLayer.color = acWhite
ThisDrawing.ActiveLayer.Linetype = Continuous
ThisDrawing.ActiveLayer.Description = "Contains New Replaced Symbols"

ThisDrawing.ActiveLayer = ThisDrawing.Layers("Old_Symbols")
ThisDrawing.ActiveLayer.color = acCyan
ThisDrawing.ActiveLayer.Linetype = Continuous
ThisDrawing.ActiveLayer.Description = "Contains Old Replaced Symbols"

End Sub
