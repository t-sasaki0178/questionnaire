Attribute VB_Name = "Module111"
Sub 二つ以上の回答用_掲載許可あり()

    Dim Total As Single         '記述件数
    Dim TrimAllText As Variant
    Dim QTitle As String, QTitles() As Variant      '設問名
    Dim OriSheetname As String, Sheetname1 As String      '元のシート名
    Dim MakeSheetname As String
    Dim MinRow As Integer, MinCol As Integer
    Dim Title_MinRow As Integer, Table_MinRow As Integer
    Dim Title_MinRow_OriTbl As Integer, Table_MinRow_OriTbl As Integer
    Dim MinRow_OriTbl As Integer, MinCol_OriTbl As Integer
     
    Dim OpenFileName As String, FileName As String, Path As String
    Dim CellNum As Integer
    
    OriSheetname = ActiveSheet.Name
    MakeSheetname = "コメント"
   
    '元データのタイトル開始行
    Title_MinRow_OriTbl = 4
    '元データの表の開始行
    Table_MinRow_OriTbl = Title_MinRow_OriTbl + 1
    '元データの開始列
    MinCol_OriTbl = 2
    '新設シートのタイトル開始行
    Title_MinRow = 3
    '新設シートの表の開始行
    Table_MinRow = Title_MinRow + 2
    '新設シートの開始列
    MinCol = 2
    j = 1
    
    Dim mySheet As Worksheet

    Set mySheet = ActiveWorkbook.Worksheets(OriSheetname)
    mySheet.Copy after:=Worksheets(OriSheetname)
    ActiveSheet.Name = "Original"
    
    Call ファイルパス取得(FileName, Path)
    'Call 元シートのセル数取得(Title_MinRow_OriTbl, MinCol_OriTbl, CellNum)
    CellNum = 1
    CellNum = CellNum + 1
    
    Do
    
    MakeSheetname = "コメント" + Format(j)
    
    Call 設問ナンバーの取得(OriSheetname, TrimAllText)
    Call タイトル取得(OriSheetname, Total, QTitle, QTitles(), Title_MinRow_OriTbl, MinCol_OriTbl)
    Call コメント用シート作成(MakeSheetname)
    Call コピペ(OriSheetname, MakeSheetname, Table_MinRow, MinCol, Table_MinRow_OriTbl, MinCol_OriTbl)
    Call 掲載許可(Table_MinRow, MinCol, MakeSheetname, FileName, Path)
    Call 文末の改行コードを削除(Title_MinRow, Table_MinRow, MinCol, MakeSheetname)
    Call 空白行削除(Title_MinRow, MinCol)
    Call ユニコード表記を文字に変換(Title_MinRow, Table_MinRow, MinCol)
    Call コメント用枠つくり(MakeSheetname, Total, QTitle, QTitles(), TrimAllText, Title_MinRow, Table_MinRow, MinCol)
    Call 列幅調整(MakeSheetname, Table_MinRow, MinCol)
    Call タイトル挿入(Title_MinRow, Table_MinRow, MinCol)
    
    Sheets(OriSheetname).Select   '元のシートを選択
    MaxRow = Cells(Rows.Count, MinCol_OriTbl).End(xlUp).Row
    MaxCol = Cells(Table_MinRow_OriTbl, Columns.Count).End(xlToLeft).Column
    
    Range(Cells(Table_MinRow_OriTbl - 1, MinCol_OriTbl + 1 + j), Cells(MaxRow, MinCol_OriTbl + 1 + j)).Select
    Application.CutCopyMode = False
    Selection.Cut
    
    Range(Cells(Table_MinRow_OriTbl - 1, MinCol_OriTbl + 1), Cells(MaxRow, MinCol_OriTbl + 1)).Select
    Selection.Insert Shift:=xlToRight
    
    j = j + 1

    Loop While j <> CellNum
    
    Sheets(OriSheetname).Select
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    
    Sheets("Original").Select
    ActiveSheet.Name = OriSheetname

End Sub
Private Sub ファイルパス取得(FileName As String, Path As String)
    Dim OpenFileName As String
    
    OpenFileName = Application.GetOpenFilename("CSVファイル,*.csv?")
    
    If OpenFileName <> "False" Then
        FileName = Dir(OpenFileName)
        Path = Replace(OpenFileName, FileName, "")
    Else
        MsgBox "キャンセルされました"
    End If
    
End Sub
Private Sub 設問ナンバーの取得(OriSheetname As String, TrimAllText As Variant)

  Dim strRet As String
  Dim intLoop As Integer
  Dim strChar As String

  strRet = ""
  TrimAllText = OriSheetname
  
  If TrimAllText Like "*_*" Then
    TrimAllText = Left(TrimAllText, InStr(TrimAllText, "_") - 1)
  End If
  
  If TrimAllText Like "*S*" Then
    TrimAllText = Left(TrimAllText, InStr(TrimAllText, "S") - 1)
  End If

  For intLoop = 1 To Len(TrimAllText)
    strChar = Mid(TrimAllText, intLoop, 1)
    If IsNumeric(strChar) Then
      strRet = strRet & strChar
    End If
  Next intLoop
 
  TrimAllText = strRet
  
End Sub
Private Sub タイトル取得(OriSheetname As String, Total As Single, QTitle As String, _
QTitles() As Variant, Title_MinRow_OriTbl As Integer, MinCol_OriTbl As Integer)

    Sheets(OriSheetname).Select   '元のシートを選択
    
    'Total = Range(Cells(Title_MinRow_OriTbl + 1, MinCol_OriTbl).Address) '記述件数の取得
    
    Sheets(OriSheetname).Select   '元のシートを選択
    MaxRow = Cells(Rows.Count, MinCol_OriTbl).End(xlUp).Row
    MaxCol = Cells(Title_MinRow_OriTbl, Columns.Count).End(xlToLeft).Column
    
    ReDim QTitles(MaxCol - MinCol_OriTbl)
    
    QTitle = Replace(Range(Cells(Title_MinRow_OriTbl, MinCol_OriTbl + 1).Address), vbLf, "")
    Point = InStr(QTitle, "【")  '【が設問名の何文字目にあるか
    QTitles(0) = Left(QTitle, Point - 1)
    
    For i = 1 To MaxCol - MinCol_OriTbl
    
        QTitle = Replace(Range(Cells(Title_MinRow_OriTbl, MinCol_OriTbl + i).Address), vbLf, "") '設問名取得
    
        Point = InStr(QTitle, "【")  '【が設問名の何文字目にあるか
    
        QTitle = Right(QTitle, Len(QTitle) - Point)
    
        Point = InStr(QTitle, "【") '【が設問名の何文字目にあるか
    
        QTitles(i) = Left(QTitle, Point - 1)
        
    Next
    
End Sub
Private Sub コメント用シート作成(MakeSheetname As String)
Attribute コメント用シート作成.VB_ProcData.VB_Invoke_Func = " \n14"

    'コメント用シート作成
    Dim NewWorkSheet As Worksheet
    
    On Error GoTo MyError
    
    Set NewWorkSheet = Worksheets.Add()
    NewWorkSheet.Name = MakeSheetname
    
    'Length = InputBox("半分時は35.09。点数は8.09")
    
    '列の幅の調整
    Sheets(MakeSheetname).Select

    'Columns(1).Select
    '    Selection.ColumnWidth = 8.09    'ID欄
    'Columns(2).Select
    '    Selection.ColumnWidth = 8.09    'ID欄
    'Columns(3).Select
    '    Selection.ColumnWidth = Length       '掲載欄
    'Columns(4).Select
    '    Selection.ColumnWidth = 70.18 - Length 'コメント欄
    Exit Sub

MyError:
    MsgBox "同じシート名があるよ"
        
End Sub
Private Sub コピペ(OriSheetname As String, MakeSheetname As String, Table_MinRow As Integer, MinCol As Integer, _
 Table_MinRow_OriTbl As Integer, MinCol_OriTbl As Integer)

    Dim MaxRow As Variant, MaxCol As Variant
    
    Sheets(OriSheetname).Select
    
    MaxRow = Cells(Rows.Count, MinCol_OriTbl).End(xlUp).Row
    MaxCol = Cells(Table_MinRow_OriTbl, Columns.Count).End(xlToLeft).Column
 
    'コピー範囲指定
    Range(Cells(Table_MinRow_OriTbl, MinCol_OriTbl), Cells(MaxRow, MaxCol)).Select
    
    Application.CutCopyMode = False
    Selection.Copy
    
    'ペースト範囲指定
    Sheets(MakeSheetname).Select
    
    Range(Cells(Table_MinRow, MinCol), Cells(MaxRow, MaxCol)).Select
    ActiveSheet.Paste
    
    Columns(3).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    'ソート
    Worksheets(MakeSheetname).Activate

       Worksheets(MakeSheetname).Rows(Table_MinRow & ":" & MaxRow) _
       .Sort Key1:=Range("B3"), order1:=xlAscending
      
End Sub
Private Sub 掲載許可(Table_MinRow As Integer, MinCol As Integer, MakeSheetname As String, _
FileName As String, Path As String)

    Dim buf As String, tmp As Variant, n As Long, tmp2() As Variant
    
    Open Path + FileName For Input As #1
    n = 0
    
    Do Until EOF(1)
        Line Input #1, buf
        n = n + 1
    Loop
        
    Close #1
    
    ReDim tmp2(n + 1, 1)
    
    Open Path + FileName For Input As #1
    a = Cells(Table_MinRow + i, MinCol).Value
    n = 0
    
    Do Until EOF(1)
        Line Input #1, buf
            
        tmp = Split(buf, ",")
        tmp2(n, 0) = n + (a - 1)
        tmp2(n, 1) = tmp(1)
            
        n = n + 1
        Loop
    Close #1

    Sheets(MakeSheetname).Select
    
    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    
    For i = 0 To MaxRow
        a = Cells(Table_MinRow + i, MinCol).Value
            For j = 1 To n - 1
                b = tmp2(j, 0)

                    If a = b Then
                        Cells(Table_MinRow + i, MinCol + 1) = tmp2(j, 1)

                    End If
            Next j
    Next i

    '1を×に変換する処理
    Columns(3).Select
    Selection.Replace What:="1", Replacement:="×", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
End Sub
Private Sub 文末の改行コードを削除(Title_MinRow As Integer, Table_MinRow As Integer, MinCol As Integer, MakeSheetname As String)

    Dim buf As String, MaxRow As Integer, MaxCol As Integer
    
    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(Table_MinRow, Columns.Count).End(xlToLeft).Column
    
    For i = 1 To MaxRow - Table_MinRow
    
        For j = MinCol + 1 To MaxCol
    
            buf = Cells(i, j).Value
    
            Do While Right(buf, 1) = vbLf
    
                If Right(buf, 1) = vbLf Then
                buf = Left(buf, Len(buf) - 1)
            End If
    
            Loop
    
        Cells(i, j) = buf
        
        Next j
    
    Next i
    
End Sub
Private Sub 空白行削除(Title_MinRow As Integer, MinCol As Integer)

    Dim MaxRow As Integer, MaxCol As Integer
    Dim i As Integer
    
    Do
    
    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(Title_MinRow, Columns.Count).End(xlToLeft).Column
    Max = MaxRow - MinRow
    counter = 0

    For i = 0 To MaxRow - Title_MinRow
        If Cells(Title_MinRow + i, MinCol + 1) = "" And Cells(Title_MinRow + i, MinCol + 2) = "" And _
         Cells(Title_MinRow + i, MinCol + 3) = "" Then
            counter = counter + 1
        End If
    Next i
    
    For i = 0 To MaxRow - Title_MinRow
        If Cells(Title_MinRow + i, MinCol + 1) = "" And Cells(Title_MinRow + i, MinCol + 2) = "" And _
         Cells(Title_MinRow + i, MinCol + 3) = "" Then
            Rows(Title_MinRow + i).Select
            Selection.Delete
        End If
    Next i
    
    Loop While counter <> 0
    
End Sub
Private Sub ユニコード表記を文字に変換(Title_MinRow As Integer, Table_MinRow As Integer, MinCol As Integer)

    Dim MaxRow As Single, MaxCol As Single
    Dim i As Single
    Dim Sentence As String
    Dim Fpoint As Single, Bpoint As Single
    Dim Length As Single
    Dim Unicode As String, UnicodeNum As Long

    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(Table_MinRow, Columns.Count).End(xlToLeft).Column
    
    Cells(Table_MinRow, MaxCol).Select

    For i = 1 To MaxRow - Table_MinRow
    
        For j = MinCol + 1 To MaxCol
    
            buf = Cells(i, j).Value
    
        Cells(i, j) = buf
        
        Next j
    
    Next i


    
    For i = Title_MinRow To MaxRow
        For j = MinCol + 1 To MaxCol
        
        Do
        Sentence = Cells(i, j).Value
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
                        Cells(i, j) = Sentence
                    End If
        
            Else
            Fpoint = 0
            End If
        Loop While Fpoint <> 0
        
        Next j
    Next i
End Sub
Private Sub コメント用枠つくり(MakeSheetname As String, Total As Single, QTitle As String, QTitles() As Variant, _
 TrimAllText As Variant, Title_MinRow As Integer, Table_MinRow As Integer, MinCol As Integer)

    Sheets(MakeSheetname).Select
    Rows("3:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Columns(3).Select
    Selection.ColumnWidth = 3    '掲載欄

    
    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(Table_MinRow, Columns.Count).End(xlToLeft).Column
    
    Range(Cells(Title_MinRow, MinCol), Cells(MaxRow, MaxCol)).Borders.LineStyle = xlDash
    Range(Cells(Title_MinRow, MinCol), Cells(Title_MinRow + 1, MaxCol)).BorderAround Weight:=xlThin
    Range(Cells(Table_MinRow, MinCol), Cells(MaxRow, MaxCol)).BorderAround Weight:=xlThin
    Range(Cells(Title_MinRow, MinCol + 1), Cells(MaxRow, MinCol + 1)).BorderAround Weight:=xlThin
    Range(Cells(Title_MinRow, MinCol + 1), Cells(MaxRow, MinCol + 1)).BorderAround Weight:=xlThin
    
    Range(Cells(Title_MinRow, MinCol), Cells(Title_MinRow + 1, MinCol)).MergeCells = True
    Range(Cells(Title_MinRow, MinCol + 1), Cells(Title_MinRow + 1, MinCol + 1)).MergeCells = True
    Range(Cells(Title_MinRow, MinCol + 2), Cells(Title_MinRow, MaxCol)).MergeCells = True
    
    Cells(Title_MinRow, MinCol) = "Q" + TrimAllText
    
    With Cells(Title_MinRow, MinCol).Font
        .Name = "Arial Black"
        .Size = 9
        .Bold = True
    End With
    
    With Cells(Title_MinRow, MinCol)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Cells(Title_MinRow, MinCol + 1) = "掲載"
    
    With Cells(Title_MinRow, MinCol + 1).Font
        .Name = "Arial Black"
        .Size = 8
        .Bold = True
    End With
    
    With Cells(Title_MinRow, MinCol + 1)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Cells(Title_MinRow, MinCol + 2) = Format(QTitles(0))

    With Cells(Title_MinRow, MinCol + 2).Font
        .Name = "ＭＳ Ｐゴシック"
        .Size = 9
        .Bold = True
    End With
    
    With Cells(Title_MinRow, MinCol + 2)
       .HorizontalAlignment = xlLeft
       .VerticalAlignment = xlCenter
    End With
    
    For i = 1 To UBound(QTitles())
        Cells(Title_MinRow + 1, MinCol + 1 + i) = QTitles(i)
    
         With Cells(Title_MinRow + 1, MinCol + 1 + i).Font
            .Name = "Arial Black"
            .Size = 8
        End With
    
        With Cells(Title_MinRow + 1, MinCol + 1 + i)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    Next
   
    'ID欄の右寄せ
    With Range(Cells(Table_MinRow, MinCol), Cells(MaxRow, MinCol))
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    
    '掲載欄のセンタリング
    With Range(Cells(Table_MinRow, MinCol + 1), Cells(MaxRow, MinCol + 1))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
End Sub
Private Sub タイトル挿入(Title_MinRow As Integer, Table_MinRow As Integer, MinCol As Integer)

    Dim TotalH As String, MaxRow As Single, MinRow As Single
    Dim Title As Range
    
    'MinRow = 3
    'MinCol = 2
    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(Title_MinRow, Columns.Count).End(xlToLeft).Column
    Startcell = Title_MinRow
    TotalH = 0

    
    For i = 1 To MaxRow
    
        If TotalH > 585 Then
            
            'タイトル部分コピー
            Range(Cells(Title_MinRow, MinCol), Cells(Title_MinRow + 1, MaxCol)).Select
            Application.CutCopyMode = False
            Selection.Copy
            
            '挿入部分範囲選択
            Range(Cells(i, MinCol), Cells(i, MaxCol)).Select
            Selection.Insert Shift:=xlDown
            
            'タイトル＋表組範囲選択
            Range(Cells(Startcell, 2), Cells(i - 1, MaxCol)).Select
            
            '0.5秒ウエイト
            Application.Wait [Now() + "0:00:00.5"]
            
            'パワーポイントにコピー＆ペースト
            Call PPt_Paste
            
            Startcell = i
            
            j = j + 2
            
            TotalH = Range(Cells(Title_MinRow, MinCol), Cells(Title_MinRow + 1, MaxCol)).Height
        
        End If
        
            TotalH = TotalH + Cells(i, MaxCol).Height
        
    Next i
    
    Range(Cells(Startcell, 2), Cells(i - 1 + j, MaxCol)).Select
    Call PPt_Paste
    
End Sub
Private Sub PPt_Paste()

    Dim ppApp As Object, ppPst As Object, ppSld As Object
    Dim ppW As Single, ppH As Single, i As Integer
    
    'PowerPoint レイアウト番号、拡張メタファイル形式
    Const ppLayoutBlank = 12
    Const ppPasteEnhancedMetafile = 2
    
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
Sub 元シートのセル数取得(Title_MinRow_OriTbl As Integer, MinCol_OriTbl As Integer, CellNum As Integer)

    Dim Title() As Variant, Titlef() As Variant
    
    MaxRow = Cells(Rows.Count, MinCol_OriTbl).End(xlUp).Row
    MaxCol = Cells(Title_MinRow_OriTbl, Columns.Count).End(xlToLeft).Column
    
    ReDim Title(MaxCol - MinCol_OriTbl), Titlef(MaxCol - MinCol_OriTbl)

    For i = 0 To MaxCol - MinCol_OriTbl
        Title(i) = Cells(Title_MinRow_OriTbl, MinCol_OriTbl + i).Value
        If Title(i) Like "*_*" Then
            Titlef(i) = Left(Title(i), InStr(Title(i), "_") - 1)
        End If
    Next
    
    i = 0
    
    If Titlef(1) = "" Then
        CellNum = 2
        Exit Sub
    End If
    
    Do
    i = i + 1
    Loop While Titlef(i) = Titlef(1)

    CellNum = i
    
End Sub
Private Sub 列幅調整(MakeSheetname As String, Table_MinRow As Integer, MinCol As Integer)

Dim a() As Variant, 最大() As Variant, 幅広さ() As Variant
Dim 開始対象列 As Long
Dim 差の総和 As Single, ゼロ以外の総和 As Single

MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
MaxCol = Cells(Table_MinRow, Columns.Count).End(xlToLeft).Column

Length = 70.18
一文字の長さ = 1.65
開始対象列 = 4
差の総和 = 0
ゼロ以外の総和 = 0

ReDim a(MaxRow - Table_MinRow, MaxCol - 開始対象列) As Variant
ReDim 最大(MaxCol - 開始対象列) As Variant, 差(MaxCol - 開始対象列) As Variant

For j = 0 To MaxRow - Table_MinRow
    For i = 0 To MaxCol - 開始対象列
        a(j, i) = Len(Cells(Table_MinRow + j, 開始対象列 + i).Value)
    Next
Next

For j = 0 To MaxRow - Table_MinRow
    For i = 0 To MaxCol - 開始対象列
        If a(j, i) > 最大(i) Then
            最大(i) = a(j, i)
        End If
    Next
Next


回数 = UBound(最大) + 1
セル内最大文字数 = Length / 回数 / 一文字の長さ

For i = 0 To UBound(最大)

    If 最大(i) > セル内最大文字数 Then
        最大(i) = セル内最大文字数
        差(i) = 0
        Columns(開始対象列 + i).ColumnWidth = Length / (回数)
            
    Else
        差(i) = セル内最大文字数 - 最大(i)
        Columns(開始対象列 + i).ColumnWidth = 最大(i) * 一文字の長さ
    End If
    
Next

'0より上の差（i)の合計
For i = 0 To UBound(最大)

    '差（i）の総和
    差の総和 = 差の総和 + 差(i)
    
    If 差(i) = 0 Then
        ゼロ以外の総和 = ゼロ以外の総和 + 1
    End If
    
Next

Dim 列長さの補正値 As Single

列幅の合計値 = 0

If ゼロ以外の総和 > 0 Then
'差（i)の合計÷差（i）の総和
    列の長さの補正値 = 差の総和 / ゼロ以外の総和
End If

For i = 0 To UBound(最大)

    'もし差が0でなければ
    If 差(i) = 0 Then
    'コラムの長さ＝コラムの長さ＋差（i)の合計÷差（i）の総和
       Columns(開始対象列 + i).ColumnWidth = Columns(開始対象列 + i).ColumnWidth + 列の長さの補正値
       
       対象セルの幅 = Columns(開始対象列 + i).ColumnWidth

    'でなければそのまま
    End If
    
Next

For i = 0 To UBound(最大)

       対象セルの幅 = Columns(開始対象列 + i).ColumnWidth
       列幅の合計値 = 列幅の合計値 + 対象セルの幅

Next

補正値 = (Length - 列幅の合計値) / 回数

For i = 0 To UBound(最大)
    Columns(開始対象列 + i).ColumnWidth = Columns(開始対象列 + i).ColumnWidth + 補正値
Next

    Cells(Table_MinRow, MinCol).Select
    Selection.CurrentRegion.Select
    Selection.Rows.AutoFit
    
End Sub

