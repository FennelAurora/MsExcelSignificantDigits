Sub ReverseSignificantFigures()
Dim arrSplitBefore() As String
Dim arrSplitAfter() As String
Dim strBefore As String
Dim strAfter As String
Dim strOriginalFormula As String

strBefore = "-(INT(LOG("
strAfter = "))+1))"

For Each rng In Selection
If IsNumeric(rng) And (InStr(rng.formula, strBefore) > 0) Then
arrSplitBefore = Split(rng.formula, strBefore)
arrSplitAfter = Split(arrSplitBefore(1), strAfter)
strOriginalFormula = arrSplitAfter(0)

rng.Value = "=" & strOriginalFormula
End If
Next rng