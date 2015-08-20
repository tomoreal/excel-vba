Sub set_header_date_title_pageno()
' ブック内の全シートに、ヘッダ、フッタを設定する。
    For Each myObj In Sheets
        With myObj.PageSetup
            .LeftHeader = "&D  &T"
            .CenterHeader = "&A"
            .RightHeader = "&F"
            .LeftFooter = ""
            .CenterFooter = "&P/&N"
            .RightFooter = ""
        End With
    Next
End Sub
