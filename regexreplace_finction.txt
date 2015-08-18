' WSHの正規表現を使った置換
' http://www.hi-ho.ne.jp/tetsuzo/windows/wsh/regexp.htm　から引用したがページ消滅
' http://d.hatena.ne.jp/s-n-k/20081007/1223395593 からVBScript.RegExp参照法を引用
' 関数名をgoogle spreadsheetと一致させ相互に動く様にした。
' 使用法： REGEXREPLACE(置換対象、befor正規表現、after正規表現)
' 2002/01/24

Function REGEXREPLACE(text_Original As String, text_Search As String, text_Replace As String) As String

    ' 参照設定で Microsoft VBScript Regular Expressions 5.5の設定をしない場合
    Dim reg   As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    ' 参照設定で Microsoft VBScript Regular Expressions 5.5 を追加した場合
    'Dim reg
    'Set reg = New RegExp
    
    Dim strText As String
    Dim text_result As String
    
    reg.Global = True
    reg.ignoreCase = False
    reg.pattern = text_Search
    
    If reg.Test(text_Original) = True Then
        text_result = reg.Replace(text_Original, text_Replace)
    Else
        text_result = text_Original
    End If
    
    REGEXREPLACE = text_result

End Function
