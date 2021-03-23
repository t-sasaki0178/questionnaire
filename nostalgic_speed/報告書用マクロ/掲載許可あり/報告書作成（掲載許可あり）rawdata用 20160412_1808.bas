Attribute VB_Name = "Module1"
Sub フォント変換_メイリオ9pt()
Attribute フォント変換_メイリオ9pt.VB_ProcData.VB_Invoke_Func = " \n14"

    Selection.CurrentRegion.Select
    
    With Selection.Font
        .Name = "メイリオ"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Selection.Columns.AutoFit
    Selection.Borders.LineStyle = xlContinuous
    
    Dim a() As Variant
    
    MaxCol = Cells(1, Columns.Count).End(xlToLeft).Column
    ReDim a(MaxCol)
    
    For i = 1 To MaxCol
        a(i) = Cells(1, i).Value
        If a(i) = "SAMPLEID" Then
            Cells(1, i) = "回答者ID"
        
        Else
            If a(i) = "ANSWERDATE" Then
                Cells(1, i) = "回答日時"
            End If
        End If
    Next
    
End Sub
Sub 都道府県名変換()

    Dim 都道府県名 As Variant
    Dim MinRow As Single, MaxRow As Single, MinCol As Single, MaxCol As Single
    Dim a As Single

    都道府県名 = Array("", "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県", "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", _
    "東京都", "神奈川県", "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県", "静岡県", "愛知県", "三重県", "滋賀県", "京都府", _
    "大阪府", "兵庫県", "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県", "徳島県", "香川県", "愛媛県", "高知県", "福岡県", _
    "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県")
    
    MinRow = ActiveCell.Row
    MinCol = ActiveCell.Column

    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column

    
    For i = 0 To MaxRow - MinRow
        a = Cells(MinRow + i, MinCol).Value
        Cells(MinRow + i, MinCol) = 都道府県名(a)
    Next
    
    Cells(MinRow - 1, MinCol) = "都道府県"

    Range(Cells(MinRow - 1, MinCol), Cells(MaxRow, MinCol)).Select
    Selection.Columns.AutoFit

End Sub
Sub 男女変換()

    Dim 男女 As Variant
    Dim MinRow As Single, MaxRow As Single, MinCol As Single, MaxCol As Single
    Dim a As Single

    男女 = Array("", "男性", "女性")
    
    MinRow = ActiveCell.Row
    MinCol = ActiveCell.Column

    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column

    
    For i = 0 To MaxRow - MinRow
        a = Cells(MinRow + i, MinCol).Value
        Cells(MinRow + i, MinCol) = 男女(a)
    Next
    
    Cells(MinRow - 1, MinCol) = "性別"

    Range(Cells(MinRow - 1, MinCol), Cells(MaxRow, MinCol)).Select
    Selection.Columns.AutoFit

End Sub
Sub 職業名変換()

    Dim 職業名 As Variant
    Dim MinRow As Single, MaxRow As Single, MinCol As Single, MaxCol As Single
    Dim a As Single

    職業名 = Array("", "自営業", "会社経営者・役員", "会社員", "公務員・団体職員", "パート・アルバイト", "専業主婦 (主夫)", "学生", "無職", "その他")
    
    MinRow = ActiveCell.Row
    MinCol = ActiveCell.Column

    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column

    
    For i = 0 To MaxRow - MinRow
        a = Cells(MinRow + i, MinCol).Value
        Cells(MinRow + i, MinCol) = 職業名(a)
    Next
    
    Cells(MinRow - 1, MinCol) = "職業"
    Cells(MinRow - 1, MinCol + 1) = "職業（その他）"

    Range(Cells(MinRow - 1, MinCol), Cells(MaxRow, MinCol + 1)).Select
    Selection.Columns.AutoFit

End Sub
Sub 年代別変換()

    Dim 年代別 As Variant
    Dim MinRow As Single, MaxRow As Single, MinCol As Single, MaxCol As Single
    Dim a As Single

    年代別 = Array("", "～19歳", "20～24歳", "25～29歳", "30～34歳", "35～39歳", "40～44歳", "45～49歳", "50～54歳", "55～59歳" _
    , "60～64歳", "65～69歳", "70歳～")
    
    MinRow = ActiveCell.Row
    MinCol = ActiveCell.Column

    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column

    
    For i = 0 To MaxRow - MinRow
        a = Cells(MinRow + i, MinCol).Value
        Cells(MinRow + i, MinCol) = 年代別(a)
    Next
    
    Cells(MinRow - 1, MinCol) = "年代"

    Range(Cells(MinRow - 1, MinCol), Cells(MaxRow, MinCol)).Select
    Selection.Columns.AutoFit

End Sub
Sub メール配信変換()

    Dim メール As Variant
    Dim MinRow As Single, MaxRow As Single, MinCol As Single, MaxCol As Single
    Dim a As Single

    メール = Array("", "希望する", "希望しない")
    
    MinRow = ActiveCell.Row
    MinCol = ActiveCell.Column

    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column

    
    For i = 0 To MaxRow - MinRow
        a = Cells(MinRow + i, MinCol).Value
        Cells(MinRow + i, MinCol) = メール(a)
    Next
    
    Cells(MinRow - 1, MinCol) = "メール配信"

    Range(Cells(MinRow - 1, MinCol), Cells(MaxRow, MinCol)).Select
    Selection.Columns.AutoFit

End Sub
Sub 掲載許可変換()

    Dim 掲載 As Variant
    Dim MinRow As Single, MaxRow As Single, MinCol As Single, MaxCol As Single
    Dim a As Single

    掲載 = Array("", "×", "")
    
    MinRow = ActiveCell.Row
    MinCol = ActiveCell.Column

    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column

    
    For i = 0 To MaxRow - MinRow
        a = Cells(MinRow + i, MinCol).Value
        Cells(MinRow + i, MinCol) = 掲載(a)
    Next
    
    Cells(MinRow - 1, MinCol) = "掲載許可"

    Range(Cells(MinRow - 1, MinCol), Cells(MaxRow, MinCol)).Select
    Selection.Columns.AutoFit

End Sub
Sub 氏名欄()

    Dim 氏名 As Variant
    Dim MinRow As Single, MinCol As Single

    氏名 = Array("", "氏名(漢字）", "氏名（よみ）")
    
    MinRow = ActiveCell.Row
    MinCol = ActiveCell.Column

    Cells(1, MinCol) = 氏名(1)
    Cells(1, MinCol + 1) = 氏名(2)


End Sub
Sub 住所欄()

    Dim MinRow As Single, MinCol As Single
    
    MinRow = ActiveCell.Row
    MinCol = ActiveCell.Column

    Cells(1, MinCol) = "住所"

End Sub
Sub メールアドレス欄()

    Dim MinRow As Single, MinCol As Single
    
    MinRow = ActiveCell.Row
    MinCol = ActiveCell.Column

    Cells(1, MinCol) = "メールアドレス"

End Sub
Sub 郵便番号欄()

    Dim MinRow As Single, MaxRow As Single, MinCol As Single, MaxCol As Single
    
    MinRow = ActiveCell.Row
    MinCol = ActiveCell.Column
    
    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column
    
    Range(Cells(MinRow, MinCol), Cells(MaxRow, MinCol)).Select
    Selection.NumberFormatLocal = "000-0000"
    
    Cells(1, MinCol) = "郵便番号"
    
    Range(Cells(1, MinCol), Cells(MaxRow, MinCol)).Select
    Selection.Columns.AutoFit

End Sub
Sub 電話番号欄()

    Dim MinRow As Single, MaxRow As Single, MinCol As Single, MaxCol As Single
    
    MinRow = ActiveCell.Row
    MinCol = ActiveCell.Column
    
    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column
    
    For i = 2 To MaxRow
    
    Cells(i, MinCol).Select
        
        If Len(Cells(i, MinCol)) = 10 Then

            Selection.NumberFormatLocal = "000-0000-0000"
        Else
            Cells(i, MinCol).Select
            Selection.NumberFormatLocal = "0000-00-0000"
        End If
            
    Next
    
    Cells(1, MinCol) = "電話番号"

    Range(Cells(1, MinCol), Cells(MaxRow, MinCol)).Select
    Selection.Columns.AutoFit

End Sub













