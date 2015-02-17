Sub SignificantFigures()
Dim rng As Range
Dim strFiguresCell As String
Dim strFormula As String

strFiguresCell = "K2"
strFormula = ""

For Each rng In Selection
strFormula = rng.formula
If IsNumeric(rng) And (rng <> 0) Then
strFormula = Mid(rng.formula, 2, 9999)
rng.Value = "=Round(" & strFormula & "," & strFiguresCell & "-(INT(LOG(" & strFormula & "))+1))"
End If
Next rng

End Sub