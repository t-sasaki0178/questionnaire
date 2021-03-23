Attribute VB_Name = "Module111"
Sub 都道府県エリア分け()

    Dim OriSheetname As String, MakeSheetname As String     '元のシート名及び作業シート名
    Dim MinRow As Long, MinCol As Long, MaxRow As Long, MaxCol As Long
    Dim Title_MinRow As Long, Table_MinRow As Long
    Dim TABLE() As Variant
    Dim ColumnWidth As Variant
     
    OriSheetname = "N％表"
    MakeSheetname = "作業シート"
    
    Sheets(OriSheetname).Select
   
    '作業シートのタイトル部分の開始行番号
    Title_MinRow = 3
    '作業シートの表組の開始行番号
    Table_MinRow = Title_MinRow + 2
    '元シートの開始列番号
    MinRow = 2
    '元シートの開始列番号
    MinCol = 2
    
    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column
    
    Worksheets(OriSheetname).Select
    
    Columns(2).Select
        With Selection
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = True
            .ReadingOrder = xlContext
        End With
    
    '----- 作業用シートを作成 -----
    With Worksheets.Add()
        .Name = MakeSheetname
    End With
    
    '----- TABLEの文字列を含むセルを特定する処理 -----
    ReDim TABLE(MaxRow - MinRow)
    
    '「TABLE」を含むセルの1行下の表範囲を選択
    For i = 0 To MaxRow - MinRow
    
        '元シート側を選択
        Sheets(OriSheetname).Select

        TABLE(i) = Cells(MinRow + i, MinCol + 1)
    
        If TABLE(i) Like "*都道府県*" Then
            Cells(MinRow + i, MinCol).Select
            Selection.CurrentRegion.Select
        
            '元シートのタイトルを削除した範囲選択及びコピー
            With Selection
                .Offset(1, 0).Select
                .Resize(Selection.Rows.Count - 1).Select
                .Rows.AutoFit
                .Copy
            End With
        
        '作業シート側を選択
        Worksheets(MakeSheetname).Select
        
        '作業シートA1に選択した範囲をコピー
        Cells(1, 1).Select
        ActiveSheet.Paste
        
        '列幅調整
        ColumnWidth = Array("", 10, 67.22, 3.33, 8.56, 8.56)
        
        For j = 1 To UBound(ColumnWidth)
            Columns(j).Select
                Selection.ColumnWidth = ColumnWidth(j)
        Next j
        
        '文末の改行コードを削除する
        Call 文末の改行コードを削除(Title_MinRow, Table_MinRow, MinCol, MakeSheetname)
        'ユニコード番号を文字に変更する
        Call ユニコード表記を文字に変換(Title_MinRow, Table_MinRow, MinCol, MakeSheetname)
        'エリア情報をcvsデータから読み出し、記入する
        Call 都道府県配列処理
        'ピボットテーブル作成
        Call Pivottable
        'エリア別の表を作成
        Call 表を作る(MakeSheetname)
        'ピボットグラフを創る
        Call ピボットグラフを創る(MakeSheetname)
        'ピボットグラフを装飾する
        Call グラフ装飾
        
        Cells(2, 12).Select
        Selection.CurrentRegion.Select
        'Selection.Offset(1, 0).Select
        'Selection.Resize(Selection.Rows.Count - 1).Select
        Selection.Copy

        'PowerPointへの貼りこみする
        Call PPt_Paste
        
        ActiveSheet.ChartObjects("グラフ 1").Activate

        'PowerPointへの貼りこみする
        Call PPt_Paste
        
        Sheets(MakeSheetname).Select
        Cells(1, 1).Select
        Selection.CurrentRegion.Clear
        
        End If
    Next i
    
    '作業シートの削除
    Sheets(MakeSheetname).Select
    ActiveWindow.SelectedSheets.Delete
    
End Sub
Private Sub 文末の改行コードを削除(Title_MinRow As Long, Table_MinRow As Long, MinCol As Long, MakeSheetname As String)

    Dim buf As String, MaxRow As Long, MaxCol As Long
    
    Sheets(MakeSheetname).Select
    
    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(Table_MinRow, Columns.Count).End(xlToLeft).Column
    
    Dim 処理対象セル As Long
    処理対象セル = 2
    
    For i = 1 To MaxRow
        buf = Cells(i, 処理対象セル).Value
    
        Do While Right(buf, 1) = vbLf
            If Right(buf, 1) = vbLf Then
                buf = Left(buf, Len(buf) - 1)
            End If
        Loop
    
        Cells(i, 処理対象セル) = buf
    Next
End Sub
Private Sub ユニコード表記を文字に変換(Title_MinRow As Long, Table_MinRow As Long, MinCol As Long, MakeSheetname As String)

    Dim MaxRow As Single, MaxCol As Single
    Dim i As Single
    Dim Sentence As String
    Dim Fpoint As Single, Bpoint As Single
    Dim Length As Single
    Dim Unicode As String, UnicodeNum As Long
    Dim 処理対象セル As Long
    処理対象セル = 2
    
    Sheets(MakeSheetname).Select

    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(Table_MinRow, Columns.Count).End(xlToLeft).Column
    
    Cells(Table_MinRow, MaxCol).Select

    For i = Title_MinRow To MaxRow
        Do
        Sentence = Cells(i, 処理対象セル).Value
            If Sentence Like "*&*#*" = True Then
                    Fpoint = InStr(Sentence, "&#")
                    Bpoint = InStr(Sentence, ";")
                    If Fpoint <> 0 Or Bpoint <> 0 Then
                        Length = Bpoint - Fpoint
                        'キャラクターユニコード番号を抽出
                        Unicode = Mid(Sentence, Fpoint, Length + 1)
                        UnicodeNum = Right(Mid(Sentence, Fpoint, Length), Length - 2)
                        '文字置き換え
                        Sentence = Replace(Sentence, Unicode, WorksheetFunction.Unichar(UnicodeNum))
                        'セルに設置
                        Cells(i, 処理対象セル) = Sentence
                    End If
            Else
            Fpoint = 0
            End If
        Loop While Fpoint <> 0
    Next i
    
End Sub
Private Sub PPt_Paste()

    Dim ppApp As Object, ppPst As Object, ppSld As Object
    Dim ppW As Single, ppH As Single, i As Long
    
    'PowerPoint レイアウト番号、拡張メタファイル形式
    Const ppLayoutBlank = 12
    Const ppPasteEnhancedMetafile = 2
    
    Application.Wait [Now() + "0:00:00.2"]
    
    '選択範囲コピー
    Selection.Copy
 
    'PowerPointを開く
    Set ppApp = CreateObject("PowerPoint.Application")
    Set ppPst = ppApp.ActivePresentation
  
    'PowerPointスライド数取得
    i = ppPst.Slides.Count
  
    'PowerPointスライド追加
    Set ppSld = ppPst.Slides.Add(Index:=i + 1, Layout:=12)
    
    '形式を選択して貼り付け
    With ppSld.Shapes.PasteSpecial(ppPasteEnhancedMetafile)(1)
        .Left = 19 '横位置(調整してください)
        .Top = 62 '縦位置(調整してください)
        .Width = ppPst.SlideMaster.Width - 38 '幅(調整してください)
    End With
   
End Sub
Private Sub 都道府県配列処理()
    Dim OpenFileName As String, FileName As String, Path As String
    Dim buf As String
    Dim tmp As Variant, n As Long, tmp2(46, 1) As Variant
    Dim todoufuken As Long
    
    OpenFileName = Application.GetOpenFilename("CSVファイル,*.csv?")
    
    If OpenFileName <> "False" Then
        FileName = Dir(OpenFileName)
        Path = Replace(OpenFileName, FileName, "")
    Else
        MsgBox "キャンセルされました"
    End If

    Open Path + FileName For Input As #1
    
    Do Until EOF(1)
        Line Input #1, buf
        tmp = Split(buf, ",")
    Loop
            
    Close #1
    
    todoufuken = 47
    
    For i = 0 To todoufuken - 1
    
        tmp2(i, 0) = tmp(i * 2)
        tmp2(i, 1) = tmp(i * 2 + 1)

    Next

    For i = 0 To todoufuken - 1
        For j = 0 To todoufuken - 1
            If Cells(4 + i, 2) = tmp2(j, 0) Then
                Cells(4 + i, 3) = tmp2(j, 1)
            End If
        Next
    Next

End Sub
Private Sub Pivottable()

    'ピボットテーブル作成

    'タイトル「エリア」を記入
    Cells(3, 3).Value = "エリア"
    
    'ピボットキャッシュ宣言
    Range(Cells(3, 2), Cells(50, 5)).Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="作業シート!R3C2:R50C5", Version:=6) _
        .CreatePivotTable TableDestination:="作業シート!R4C8", TableName:="ピボットテーブル1", DefaultVersion:=6
        
    'ピボットテーブル作成
    Sheets("作業シート").Select
    
    With ActiveSheet.PivotTables("ピボットテーブル1")
    
        With .PivotFields("エリア")
        
            .Orientation = xlRowField
            .Position = 1
            
        End With
        
        With .PivotFields("単一回答")
        
            .Orientation = xlRowField
            .Position = 2
            
        End With
        
        With ActiveSheet.PivotTables("ピボットテーブル1")
    
            .AddDataField ActiveSheet.PivotTables("ピボットテーブル1").PivotFields("Ｎ"), "合計 / Ｎ", xlSum
            .AddDataField ActiveSheet.PivotTables("ピボットテーブル1").PivotFields("％"), "合計 / ％", xlSum
            
            With .PivotFields("合計 / ％")
            
                .Calculation = xlPercentOfTotal
                .NumberFormat = "0.0%"
                
            End With
            
            With .PivotFields("エリア")
            
                .ShowDetail = False
                
            End With
            
        End With
        
    End With
    
End Sub
Private Sub 表を作る(MakeSheetname As String)

    Dim strPivotName As String  'ピボットテーブル名
    Dim DataArea As Range       'ピボットテーブルのセル範囲
    Dim aryData As Variant      'データ範囲から値を取り込む2次元配列
    Dim MinRow As Long, MinCol As Long, MaxRow As Long, MaxCol As Long, FieldRow As Long
       
    'コピー先の列の幅調整
    Columns(12).Select
        Selection.ColumnWidth = 10
    
    Columns(13).Select
        Selection.ColumnWidth = 70.55
    
    Columns("N:O").Select
        Selection.ColumnWidth = 8.56
    
    '％の書式設定
    Columns(15).Select
        Selection.NumberFormatLocal = "#,##0.0_ ;[赤]-#,##0.0 "
    
    '------------ピボットテーブルからセルへコピー-----------
    strPivotName = "ピボットテーブル1"
    
    With Sheets(MakeSheetname)
    'ピボットテーブルのデータ範囲取得
        Set DataArea = .PivotTables(strPivotName).TableRange1
        
    'セル範囲から2次元配列作成
        aryData = DataArea
  
    'セル範囲に転記
        DataArea.Offset(-1, 5) = aryData
        
    End With
     
    '--------------表の罫線を引く--------------
    MinRow = 3
    MinCol = 13
    
    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column
    
    FieldRow = MaxRow - MinRow
    
    '表の周囲の罫線を実線で引く
    With Range(Cells(MinRow - 1, MinCol - 1), Cells(MaxRow, MaxCol))
        .Borders.LineStyle = xlContinuous
        
    '表内部の横線を消す
        .Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    End With
    
    'データ範囲の横線をヘアラインで引く
    Range(Cells(MinRow - 1, MinCol + 1), Cells(MaxRow - 1, MaxCol)).Borders(xlInsideHorizontal).Weight = xlHairline
    
    'タイトルの周囲の罫線を実線で引く
    Range(Cells(MinRow - 1, MinCol - 1), Cells(MinRow, MaxCol)).Borders(xlEdgeBottom).Weight = xlThin
    
    '
    Range(Cells(MinRow - 1, MinCol), Cells(MaxRow, MinCol)).Borders(xlInsideVertical).Weight = xlHairline
    Range(Cells(MinRow - 1, MaxCol - 1), Cells(MaxRow, MaxCol)).Borders(xlInsideVertical).Weight = xlHairline
    
    '全体の横線を二重線で引く
    With Range(Cells(MaxRow, MinCol - 1), Cells(MaxRow, MaxCol))
        .Borders(xlEdgeTop).LineStyle = xlDouble
    '背景色を塗る
        .Interior.Color = RGB(204, 255, 255)
    End With
    
    '--------------文字の調整--------------
    'エリアのナンバーを削除する
    For i = 0 To FieldRow
    
        Cells(MinRow + i, MinCol) = Mid(Cells(MinRow + i, MinCol), 4)
        
            If IsNumeric(Cells(MinRow + i, MaxCol)) = True Then
                Cells(MinRow + i, MaxCol) = Cells(MinRow + i, MaxCol) * 100
            End If
        
    Next i
    'ナンバーを振る処理
    For i = 1 To FieldRow - 1
    
        Cells(MinRow + i, MinCol - 1) = i
        
    Next i
     
    'タイトル部分をクリア
    Range(Cells(MinRow - 1, MinCol - 1), Cells(MinRow, MaxCol)).ClearContents
    
    '全体に
    Cells(MaxRow, MinCol).Value = "全体"
    
    'セルの結合
    Range(Cells(MinRow - 1, MinCol - 1), Cells(MinRow, MinCol - 1)).Merge
    
    '---------------文字のコピー-----------
    'Q
    With Cells(MinRow - 1, MinCol - 1)
    
        .Value = Cells(2, 1).Value
        
        With .Font
            .Name = "Arial Black"
            .Size = 9
        End With
        
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
    End With
    
    'タイトル「あなたがお住まいの都道府県をお知らせください。」
    With Cells(MinRow - 1, MinCol)
    
        .Value = Cells(2, 2).Value
        With .Font
                .Name = "ＭＳ Ｐゴシック"
                .Size = 9
                .Bold = True
        End With
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        
    End With
    
    '単一回答
    With Cells(MinRow, MinCol)
    
        .Value = Cells(3, 2).Value
        With .Font
                .Name = "ＭＳ Ｐゴシック"
                .Size = 8
        End With
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        
    End With
    
    'N
    With Cells(MinRow, MinCol + 1)
        .Value = Cells(3, 4).Value
        With .Font
                .Name = "ＭＳ Ｐゴシック"
                .Size = 8
        End With
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    '%
    With Cells(MinRow, MaxCol)
        .Value = Cells(3, 5).Value
        With .Font
                .Name = "ＭＳ Ｐゴシック"
                .Size = 8
        End With
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

End Sub
Private Sub ピボットグラフを創る(MakeSheetname As String)

    Dim strPivotName As String  'ピボットテーブル名
    Dim DataArea As Range       'ピボットテーブルのセル範囲
    Dim aryData As Variant      'データ範囲から値を取り込む2次元配列
    Dim MinRow As Long, MinCol As Long, MaxRow As Long, MaxCol As Long, FieldRow As Long

    'ピボッドテーブル操作
    With ActiveSheet.PivotTables("ピボットテーブル1")
        '合計を隠す
        .PivotFields("合計 / Ｎ").Orientation = xlHidden
        'エリアを降順に
        .PivotFields("エリア").AutoSort xlDescending, "エリア"
    End With

    '------------ピボットグラフ用に表をコピー-----------
    strPivotName = "ピボットテーブル1"
    
    With Sheets(MakeSheetname)
    'ピボットテーブルのデータ範囲取得
        Set DataArea = .PivotTables(strPivotName).TableRange1
        
    'セル範囲から2次元配列作成
        aryData = DataArea
  
    'セル範囲に転記
        DataArea.Offset(22, 0) = aryData
        DataArea.Offset(22, 0).Borders.LineStyle = xlContinuous

    End With
    
    '--------------文字の調整--------------
    MinRow = 26
    MinCol = 8
    
    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column
    
    FieldRow = MaxRow - MinRow

    'エリアのナンバーを削除する
    For i = 0 To FieldRow
    
        Cells(MinRow + i, MinCol) = Mid(Cells(MinRow + i, MinCol), 4)
        
            If IsNumeric(Cells(MinRow + i, MaxCol)) = True Then
                Cells(MinRow + i, MaxCol) = Cells(MinRow + i, MaxCol) * 100
            End If
        
    Next i
    
    'ナンバーを振る処理
    'For i = 1 To FieldRow - 1
    '
    '    Cells(MinRow + i, MinCol - 1) = i
        
    'Next i
    
    '書式設定
    Range(Cells(MinRaw + 1, MaxCol), Cells(MaxRow, MaxCol)).NumberFormatLocal = "#,##0_ ;[赤]-#,##0 "
    
    With Range(Cells(MinRow, MaxCol - 1), Cells(MaxRow, MaxCol))
        With .Font
                .Name = "ＭＳ Ｐゴシック"
                .Size = 8
        End With
    End With

End Sub
Private Sub グラフ装飾()

    Dim MinRow As Long, MinCol As Long, MaxRow As Long, MaxCol As Long, FieldRow As Long
    
    MinRow = 26
    MinCol = 8
    
    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column
    
    FieldRow = MaxRow - MinRow

    Range(Cells(MinRow, MinCol), Cells(MaxRow, MaxCol)).Select
    
    'グラフ作成　設置場所、大きさ指定
    With ActiveSheet.Shapes
        .AddChart2(216, xlBarClustered).Select
    End With
    
    With ActiveChart
        .HasAxis(xlValue) = True
        
        With .Axes(xlValue)
            .MaximumScale = 100
            .MajorTickMark = xlOutside
            .TickLabels.NumberFormatLocal = "G/標準!%"
        End With
        
    End With
    
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Weight = 0.25
    End With
    
    'グラフタイトル選択
    ActiveChart.ChartTitle.Select
    
    'グラフのタイトル　文字の大きさ
    With ActiveSheet.ChartObjects(1).Chart
        .HasTitle = True
        .ChartTitle.Text = "[" & Cells(2, 1).Value & "]" & Cells(2, 2).Value & vbCr & "(n=" & Format(Cells(51, 4).Value) & ")"
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 10
            .Bold = False
        End With
    End With
    
    'グラフタイトル位置設定
     With ActiveChart.ChartTitle
      '---タイトル上端位置をグラフエリアの上端に設定
        .Top = 0
      '---タイトル左端位置をグラフエリアの左端に設定
        .Left = 0
      'グラフのタイトル　左寄せ
        .HorizontalAlignment = xlHAlignLeft
    End With

    With Selection.Format
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
    
    With ActiveChart
        .SetElement (msoElementDataLabelOutSideEnd)
        .ApplyDataLabels
        
        With .FullSeriesCollection(1)
            .DataLabels.Select
            .HasLeaderLines = False
        End With
        
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = True
            .MajorGridlines.Select
        End With
        
    End With
    
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Weight = 0.25
    End With
    
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.NumberFormat = "0.0_ "
    
    ActiveChart.FullSeriesCollection(1).Select
    
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 0.25
    End With
    
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(51, 102, 255)
        .BackColor.RGB = RGB(51, 102, 255)
        .TwoColorGradient msoGradientHorizontal, 3
    End With
    
    Selection.Format.Fill.Visible = msoTrue

    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 1
    End With
    
    With ActiveChart
        .ChartGroups(1).GapWidth = 40
        .ChartArea.Select
        .PlotArea.Select
    End With
    
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Weight = 0.25
    End With
    
    ActiveChart.Axes(xlCategory).Select
    
    'With Selection.Format.Line
    '    .Visible = msoTrue
    '    .Weight = 0.25
    'End With
    
    Selection.MajorTickMark = xlInside
    
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Weight = 0.25
    End With
        
    ActiveChart.PlotArea.Select
    
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With

    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    
    'With Selection.Format.Line
    '    .Visible = msoTrue
    '  .Weight = 0.25
    'End With

    ActiveSheet.Shapes("グラフ 1").Width = 543.5433070866
    
    ActiveChart.PlotArea.Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    
    ActiveChart.Axes(xlValue).Select
    Selection.TickLabelPosition = xlHigh
 
End Sub

