Sub hyperlink_to_files_in_dir()
    'ダイアログで指定したディレクトリのファイルへのハイパーリンクをsheet1に作成する
    Dim PATH As String
    Dim FileName As String
    Dim i, j As Integer
    
    'フォルダの選択
    '複数のWord文書に連続して処理を施すマクロ http://stabucky.com/wp/archives/3004
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "フォルダを選択"
        .AllowMultiSelect = False
        If .Show = -1 Then
            PATH = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With
    
    With Worksheets("Sheet1") 'シート名指定
        i = 1 '一覧表の開始セル行番号
        j = 1 '一覧表の開始セル列番号
        FileName = Dir(PATH & "\*.*")
        Do Until FileName = ""
            .Hyperlinks.Add _
                Anchor:=.Cells(i, j), _
                address:=PATH & "\" & FileName, _
            TextToDisplay:=FileName
            i = i + 1
            FileName = Dir
        Loop
    End With
End Sub
