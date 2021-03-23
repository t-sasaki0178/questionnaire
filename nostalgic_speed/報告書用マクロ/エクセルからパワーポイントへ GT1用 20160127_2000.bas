Attribute VB_Name = "Module111"
Sub 共通GT1処理用()

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

        TABLE(i) = Cells(MinRow + i, MinCol)
    
        If TABLE(i) Like "*TABLE*" Then
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
        
        Cells(1, 1).Select
        Selection.CurrentRegion.Select
        Selection.Offset(1, 0).Select
        Selection.Resize(Selection.Rows.Count - 1).Select
        Selection.Copy

        'PowerPointへの貼りこみする
        Call PPt_Paste
        
        Sheets("グラフ").Select
        ActiveSheet.ChartObjects(Replace(TABLE(i), "TABLE", "GRAPH")).Activate

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
