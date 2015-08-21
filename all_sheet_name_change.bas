'1番目のシートに書いた名前で、全シート名を変更する
'シート名は、A列に縦に2行目から書く。
Sub all_sheet_name_change()
Dim i As Integer
For i = 2 To Worksheets.Count
 Sheets(i).Name = Cells(i, 1).Value
Next
End Sub
