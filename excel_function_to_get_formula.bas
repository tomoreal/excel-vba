'excel function to get formula
'指定したセルの数式をセルに表示する関数
Public Function getformula(ByVal target As Excel.Range) As Variant
    getformula = target.Cells(1).Formula
End Function
