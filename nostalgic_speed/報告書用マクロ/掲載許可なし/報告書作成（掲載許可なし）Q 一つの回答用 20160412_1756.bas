Attribute VB_Name = "Module1"
Sub 一つの回答用_掲載許可なし()

    Dim Total As Single         '記述件数
    Dim TrimAllText As Variant
    Dim QTitle As String         '設問名
    Dim Sheetname1 As String       '元のシート名

    Call シート名取得(Sheetname1)
    Call コメント用シート作成
    Call 設問ナンバーの取得(Sheetname1, TrimAllText)
    Call タイトル取得(Sheetname1, Total, QTitle)
    Call コピペ(Sheetname1)
    Call コメント用枠つくり(Sheetname1, Total, QTitle, TrimAllText)
    Call 文末の改行コードを削除
    Call ユニコード表記を文字に変換
    Call タイトル挿入

    
End Sub
Private Sub シート名取得(Sheetname1 As String)

    Sheetname1 = ActiveSheet.Name

End Sub
Private Sub コメント用シート作成()
Attribute コメント用シート作成.VB_ProcData.VB_Invoke_Func = " \n14"

    'コメント用シート作成
    Dim NewWorkSheet As Worksheet
    
    Set NewWorkSheet = Worksheets.Add()
    NewWorkSheet.Name = "コメント"
    
    '列の幅の調整
    Sheets("コメント").Select

    Columns("B:B").Select
        Selection.ColumnWidth = 8.09    'ID欄
    Columns("C:C").Select
        Selection.ColumnWidth = 73.18   'コメント欄
    
End Sub
Private Sub タイトル取得(Sheetname1 As String, Total As Single, QTitle As String)

    Sheets(Sheetname1).Select   '元のシートを選択
    
    Total = Range(Cells(2, 3).Address)  '記述件数の取得
    
    QTitle = Replace(Range(Cells(4, 3).Address), vbLf, "")   '設問名取得
    
    Point = InStr(QTitle, "【")  '【が設問名の何文字目にあるか
    
    QTitle = Left(QTitle, Point - 1)  '【より前の設問名を取得
    
End Sub
Private Sub コピペ(Sheetname1 As String)

    Dim MinRow As Variant
    Dim MinCol As Variant
    Dim MaxRow As Variant
    Dim MaxCol As Variant
    
    Sheets(Sheetname1).Select
    
    MinRow = 3
    MinCol = 2
    
    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column
 
    'コピー範囲指定、コピー
    Range(Cells(MinRow + 2, MinCol), Cells(MaxRow, MinCol + 1)).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    'ペースト範囲指定、ペースト
    Sheets("コメント").Select
    Range(Cells(MinRow + 2, MinCol), Cells(MaxRow, MinCol + 1)).Select
    ActiveSheet.Paste
    
    'ソート
    Worksheets("コメント").Activate
    Worksheets("コメント").Range(Cells(MinRow + 2, MinCol), Cells(MaxRow, MinCol + 1)).Sort Key1:=Range("B3"), order1:=xlAscending
    
End Sub
Private Sub コメント用枠つくり(Sheetname1 As String, Total As Single, QTitle As String, TrimAllText As Variant)

    Dim MinRow As Single, MinCol As Single
    
    Sheets("コメント").Select

    MinRow = 3
    MinCol = 2
    
    Range(Cells(MinRow, MinCol), Cells(MinRow + 1, MinCol + 1)).BorderAround Weight:=xlThin
    Range(Cells(MinRow, MinCol + 1), Cells(MinRow + 1, MinCol + 1)).BorderAround Weight:=xlThin
    
    Range(Cells(MinRow, MinCol), Cells(MinRow + 1, MinCol)).MergeCells = True

    Cells(MinRow, MinCol) = "Q" + TrimAllText
    
    With Cells(MinRow, MinCol).Font
        .Name = "Arial Black"
        .Size = 9
        .Bold = True
    End With
    
    With Cells(MinRow, MinCol)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Cells(MinRow, MinCol + 1) = Format(QTitle)
    
    With Cells(MinRow, MinCol + 1).Font
        .Name = "ＭＳ Ｐゴシック"
        .Size = 9
        .Bold = True
    End With
    
    Cells(MinRow, MinCol + 1).Select
    
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Cells(MinRow + 1, MinCol + 1) = "記述式"
    
    With Cells(MinRow + 1, MinCol + 1).Font
        .Name = "ＭＳ Ｐゴシック"
        .Size = 8
    End With


    Range(Cells(MinRow + 2, MinCol), Cells(MinRow + 2 + Total, MinCol + 1)).Borders.LineStyle = xlDash
    Range(Cells(MinRow + 2, MinCol), Cells(MinRow + 2 + Total, MinCol + 1)).Borders.Weight = xlHairline
    
    Range(Cells(MinRow + 2, MinCol), Cells(MinRow + 1 + Total, MinCol + 1)).BorderAround Weight:=xlThin
    
    Range(Cells(MinRow + 2, MinCol), Cells(MinRow + 2 + Total, MinCol)).HorizontalAlignment = xlRight
    Range(Cells(MinRow + 2, MinCol), Cells(MinRow + 2 + Total, MinCol)).VerticalAlignment = xlCenter
    
    Range(Cells(MinRow, MinCol + 1), Cells(MinRow + 1 + Total, MinCol + 1)).BorderAround Weight:=xlThin
    
End Sub
Private Sub 設問ナンバーの取得(Sheetname1 As String, TrimAllText As Variant)

  Dim strRet As String
  Dim intLoop As Integer
  Dim strChar As String

  strRet = ""
  TrimAllText = Sheetname1
  
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
Sub 文末の改行コードを削除()

    Dim buf As String, MaxRow As Single, MinRow As Single
    
    MaxRow = Cells(Rows.Count, ActiveCell.Column).End(xlUp).Row
    MaxCol = Cells(ActiveCell.Row, Columns.Count).End(xlToLeft).Column
    
    For i = 1 To MaxRow
    
        Cells(i, MaxCol).Select
    
        buf = Cells(ActiveCell.Row, ActiveCell.Column)
    
        Do While Right(buf, 1) = vbLf
    
            If Right(buf, 1) = vbLf Then
    
                buf = Left(buf, Len(buf) - 1)
        
            End If
    
        Loop
    
        Cells(ActiveCell.Row, ActiveCell.Column) = buf
    
    Next
    
End Sub

Sub ユニコード表記を文字に変換()
    Dim MaxRow As Single, MaxCol As Single
    Dim i As Single
    Dim Sentence As String
    Dim Fpoint As Single, Bpoint As Single
    Dim Length As Single
    Dim Unicode As String, UnicodeNum As Long
    
    MinRow = 5
    MinCol = 3

    MaxRow = Cells(Rows.Count, MinCol).End(xlUp).Row
    MaxCol = Cells(MinRow, Columns.Count).End(xlToLeft).Column
    
    Cells(MinRow, MaxCol).Select

    For i = 3 To MaxRow

        Do
    
        Sentence = Cells(i, MaxCol).Value
            
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
                        Cells(i, MaxCol) = Sentence
    
                    Else
    
                    End If
        
            Else
        
            Fpoint = 0
            
            End If
    
        Loop While Fpoint <> 0

    Next i
    
End Sub
Private Sub タイトル挿入()

    Dim TotalH As String, MaxRow As Single, MinRow As Single
    Dim Title As Range
    
    MinRow = 3
    MinCol = 2
    MaxRow = Cells(Rows.Count, ActiveCell.Column).End(xlUp).Row
    MaxCol = Cells(ActiveCell.Row, Columns.Count).End(xlToLeft).Column
    Startcell = MinRow
    TotalH = 0

    
    For i = 1 To MaxRow
    
        If TotalH > 585 Then
        
            Range(Cells(MinRow, MinCol), Cells(MinRow + 1, MaxCol)).Select
            Application.CutCopyMode = False
            Selection.Copy
            
            Range(Cells(i, MinCol), Cells(i, MaxCol)).Select
            Selection.Insert Shift:=xlDown
            
            Range(Cells(Startcell, 2), Cells(i - 1, MaxCol)).Select
            
            '0.5秒ウエイト
            Application.Wait [Now() + "0:00:00.5"]
            
            Call PPt_Paste
            
            Startcell = i
            
            J = J + 2
            
            TotalH = Range(Cells(3, 1), Cells(4, MaxCol)).Height
        
        End If
        
            TotalH = TotalH + Cells(i, MaxCol).Height
        
    Next i
    
    Range(Cells(Startcell, 2), Cells(i - 1 + J, MaxCol)).Select
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





