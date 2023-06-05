Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : ③変更住所録で①原簿と②archivesを更新する
' --------------------------------------+-----------------------------------------
' | @moduleName : m2_住所変更処理
' | @Version    : v1.0.0
' | @update     : 2023/06/02
' | @written    : 2023/06/02
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |  （１）単独レコードを新規シート（①new）へ移動する
' |  （２）重複レコードは、次のチェックを行い、変更を反映し、新規シート（①new）へ移動する
' |     ⅰ）(42)姓名keyが同じレコードは、重複レコードであること
' |     ⅱ）(54)識別区分が「1」(①原簿)もしくは「2」(②archives)のレコードに対し
' |       「3」(③変更住所録)の変更項目（空白でない項目）をコピーし、新規シート（①new）へ移動する
' |     ⅲ）コピーする変更項目と同じ内容のものであることを確認します。
' |     ⅳ）「2」(②archives)のレコードの③変更住所録の削除日の西暦年が「9999」のレコードは、①原簿として移動します
' |
' |
' --------------------------------------+----------------------------------------
' |  命名規則の統一
' |     Public変数  先頭を大文字    ≡ pascalCase
' |     private変数 先頭を小文字    ≡ camelCase
' |     定数        全て大文字、区切り文字は、アンダースコア(_) ≡ snake_case
' |     引数        接頭語(p_)をつけ、camelCaseに準ずる
' --------------------------------------+-----------------------------------------
'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' 共通有効シートサイズ（データ部のみの領域）
'
Private Wb                              As Workbook         ' このブック
' 作業シート work の定義
Private wsWrk                           As Worksheet
Private wrkX, wrkXmin, wrkXmax          As Long             ' i≡x 列　column
Private wrkY, wrkYmin, wrkYmax          As Long             ' j≡y 行　row
' Private wrkCnt                          As Long             ' 作業レコードの件数=修正前+変更

' ③new シートの定義
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i≡x 列　column
Private newY, newYmin, newYmax          As Long             ' j≡y 行　row
' Private new1Cnt                         As Long             ' new①原簿レコードの件数
' Private new2Cnt                         As Long             ' new③archivesレコードの件数

' ' 構造体の宣言
' Type cntTbl
'     old                                 As long     ' ①原簿
'     arv                                 As long     ' ②archive
'     trn                                 As long     ' ③変更住所録
'     wrk                                 As long     ' 作業レコードの件数=修正前+変更
'     new1                                As long     ' new①原簿レコードの件数
'     new2                                As long     ' new③archivesレコードの件数
'     new3                                As Long     ' newの変更住所録で新規レコード
' End Type
' Dim cnt                                 As cntTbl


Public Sub m2_住所変更処理_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     workシートの前後のレコードを比較
' --------------------------------------+-----------------------------------------
    Dim cnt                             As cntTbl
    Dim x, y                            As Long
    Dim CloseingMsg                     As String
    Dim w_rate, w_mod                   As Integer      ' 進捗率 / 表示間隔
    Dim i, iMin, iMax                   As Long         ' 同一レコードの範囲(列 col x)
    Dim j, jMin, jMax                   As Long         ' 同一レコードの範囲(行 row y)
    Dim r                               As Long         ' 変更項目の列番号
'
' ---Procedure Division ----------------+-----------------------------------------
'
' オブジェクト変数の定義（共通）
    Set Wb = ThisWorkbook
    ' 表の大きさを得る
    ' 作業シート（このシート）の初期値
    Set wsWrk = Wb.Worksheets("work")                       ' 作業用シートのため固定
    wrkYmin = YMIN
    wrkXmin = XMIN
    wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）3列目（"C")名前列で計測
    wrkXmax = wsWrk.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）   ' ヘッダー行 3行目で計測
    
' ③new シートの初期値
    Set wsNew = Wb.Worksheets(Range("C_newSheet").Value)                    ' 新住所録シート
    newYmin = YMIN
    newXmin = XMIN
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）
' --------------------------------------+-----------------------------------------
' step1. 単独レコードのみ先に新住所録シートへ移動する    
    newY = newYmin
    For y = wrkYmin To wrkYmax
        If wsWrk.Cells(y, PKEY_X).Value <> wsWrk.Cells(y + 1, PKEY_X).Value Then
            wsWrk.Cells(y, CHECKED_X) = "NA"
            wsWrk.Rows(y).Copy Destination:=wsNew.Rows(newY)
            wsWrk.Rows(y).ClearContents
            
            Select Case wsNew.Cells(newY, MASTER_X).Value
                Case 1
                    cnt.new1 = cnt.new1 + 1
                Case 2
                    cnt.new2 = cnt.new2 + 1
                Case 3
                    cnt.new3 = cnt.new3 + 1
                Case Else
                    MsgBox "識別区分エラー=" & wsNew.Cells(newY, MASTER_X).Value
                    End
            End Select
            newY = newY + 1
            GoTo Next_R
        End If

' 同一keyの更新処理
' step1:(42)key姓名昇順、(54)識別区分(1…①原簿, 2…②archives, 3…③変更住所録)降順にソートし、単独レコードの行を除く
' step2:(42)もしくは(54)のレコードを(54)識別区分を 0 としてコピーする（afterレコード）
' step3:再度、step1と同様にソートを行う
' step4:同一キーの(42)key姓名が、3→1or2→0　の順に並ぶので、変更項目を 0 レコードを更新し、新住所録シートへコピーする

' ' オブジェクト変数の定義（共通）
'     Sheets("work").Activate
' ' 表の大きさを得る
' ' 作業シート（このシート）の初期値
'     Set wsWrk = Wb.Worksheets("work")
'     wrkYmin = YMIN
'     wrkXmin = XMIN
'     wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）6列目（"F")名前列で計測
'     wrkXmax = wsWrk.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）   ' ヘッダー行 3行目で計測
'     wrkCnt = wrkYmax - wrkYmin + 1

' 昇順ソート　key: (39)姓名key(昇順)、(54)識別区分:BA列(降順）
    With ActiveSheet                '対象シートをアクティブにする
        .Sort.SortFields.Clear      '並び替え条件をクリア
        '項目1
        .Sort.SortFields.Add2 _
             Key:=.Range(PKEY_RNG) _
            , SortOn:=xlSortOnValues _
            , Order:=xlAscending _
            , DataOption:=xlSortNormal
        '項目2
        .Sort.SortFields.Add2 _
             Key:=.Range(MASTER_RNG) _
            , SortOn:=xlSortOnValues _
            , Order:=xlDescending _
            , DataOption:=xlSortNormal
'並び替えを実行する
        With .Sort
            .SetRange Range(Cells(wrkYmin - 1, wrkXmin), Cells(wrkYmax, wrkXmax))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
' 表の大きさを得る
    newYmin = YMIN
    newXmin = XMIN
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）
' (54)識別区分を 0 の afterレコードをコピーする
    Dim addY as long = wrkYmax          ' 追加する行
    For y = wrkYmin To wrkYmax step 2
        If wsWrk.Cells(y, PKEY_X).Value = wsWrk.Cells(y + 1, PKEY_X).Value Then
            wsWrk.Cells(y+1, CHECKED_X) = "added"
            addY = addY + 1
            wsWrk.Rows(y + 1).Copy Destination:=wsWrk.Rows(addY)
            wsWrk.cells(addY, MASTER_RNG) = 0
        Else
            msgbox "重複キーは、２レコードのルール違反。要確認！！"
            Stop
            END
        End If
    next y
stop
' 上書き項目：(6)名前～(15)方書
       For r = 6 To 15
            If wsWrk.Cells(y, r).Value <> "" Then
                wsWrk.Cells(y + 1, r).Value = wsWrk.Cells(y, r).Value
            End If
        Next r

' 上書き項目：(23)その他1～(26)備考
       For r = 23 To 26
            If wsWrk.Cells(y, r).Value <> "" Then
                wsWrk.Cells(y + 1, r).Value = wsWrk.Cells(y, r).Value
            End If
        Next r

' 上書き項目：(36)更新内容～(41)削除日
        For r = 36 To 41
            If wsWrk.Cells(y, r).Value <> "" Then
                wsWrk.Cells(y + 1, r).Value = wsWrk.Cells(y, r).Value
            End If
        Next r
' 同一キーなので一つ飛ばし
    y = y + 1
    
Next_R:
    Next y

' グループ項目：(16)携帯電話～(19)会社電話
        Dim r1 As Long
        Dim sameCnt As Long
        sameCnt = 0
' 変更項目数をカウント
        For r = 16 To 19
            If wsWrk.Cells(y, r).Value <> "" Then
                sameCnt = sameCnt + 1
            End If
        Next r
        
' 変更内容が現行と同一の内容かチェック
        For r = 16 To 19
            If wsWrk.Cells(y, r).Value <> "" Then
                For r1 = 16 To 19
                    If wsWrk.Cells(y, r).Value = wsWrk.Cells(y + 1, r1).Value Then
                        wsWrk.Cells(y + 1, r1).Value = ""
                        sameCnt = sameCnt - 1
                        Exit For
                    End If
                Next r1
            End If
        Next r
' 違う内容のものを空いてるセルにコピー
        If sameCnt <> 0 Then
            For r = 16 To 19
                If wsWrk.Cells(y, r).Value <> "" Then
                    For r1 = 16 To 19
                        If wsWrk.Cells(y + 1, r1).Value = "" Then
                            wsWrk.Cells(y + 1, r1).Value = wsWrk.Cells(y, r).Value
                            wsWrk.Cells(y, r).Value = ""
                            
                            Exit For
                        End If
                    Next r1
                End If
            Next r
        End If
        
        wsWrk.Rows(y).ClearContents
        wsWrk.Cells(y + 1, CHECKED_X) = "Mod"
        wsWrk.Rows(y + 1).Copy Destination:=wsNew.Rows(newY)
        wsWrk.Rows(y + 1).ClearContents
            
        Select Case wsNew.Cells(newY, MASTER_X).Value
            Case 1
                cnt.new1 = cnt.new1 + 1
            Case 2
                cnt.new2 = cnt.new2 + 1
            Case 3
                cnt.new3 = cnt.new3 + 1
            Case Else
                MsgBox "識別区分エラー=" & wsNew.Cells(newY, MASTER_X).Value
                End
        End Select
        newY = newY + 1
        y = y + 1   ' 同一keyが二つあるので、一つindexをくりあげる
        
        
Stop
        
End Sub

' グループ項目：(20)携帯メール～(22)会社メール

''
''
''
''
''
''            jMin = trnY                 ' 同一keyの最初の行(Row, y)
''
''            Do While wsWrk.Cells(y, PKEY_X).Value = wsWrk.Cells(y + 1, PKEY_X).Value
''                wsWrk.Rows(y).Copy Destination:=wsTrn.Rows(trnY)
''                trnCnt = trnCnt + 1
''                wsTrn.Cells(trnY, 42).Value = trnCnt
''                trnY = trnY + 1
''
''
''                wsWrk.Activate
''                wsWrk.Cells(y, CHECKED_X) = "③trn"
''                y = y + 1
''            Loop
''            wsWrk.Activate
''            wsWrk.Cells(y, CHECKED_X) = "③trn"
''            wsTrn.Activate
''            wsWrk.Rows(y).Copy Destination:=wsTrn.Rows(trnY)
''            trnCnt = trnCnt + 1
''            wsTrn.Cells(trnY, 42).Value = trnCnt
''' 最新のレコードを統合コピーの候補「⑨-999」とする
''            Rows(trnY & ":" & trnY).Select
''            Selection.Copy
''            trnY = trnY + 1
''            Rows(trnY & ":" & trnY).Select
''            ActiveSheet.Paste
''            Application.CutCopyMode = False     ' コピー状態の解除
'''            Rows(trnY).Insert
''            Cells(trnY, 1) = "⑨-999"
''            Cells(trnY, 42) = ""
''            jMax = trnY
''
''' チェックマーク行追加
''            trnY = trnY + 1
''            Cells(trnY, 1) = "???"
''            Rows(trnY & ":" & trnY).Interior.ColorIndex = xlNone    ' 色の初期化
''            trnY = trnY + 1
''' 統合手順の自動化 / 統合候補［⑨-999］が空白の項目は、過去のデータから持ってくる
''            For i = INPUTX_FROM To INPUTX_TO + 9
''                If Cells(jMax, i).Value = "" Or _
''                   Cells(jMax, i).Value = "　" Or _
''                   Cells(jMax, i).Value = " " Then
''                    For j = jMin To jMax - 1
''                        If Cells(j, i).Value <> "" Then
''                            Cells(jMax, i).Value = Cells(j, i).Value        ' ⑨-999 行
''                            Cells(jMax + 1, i).Value = Cells(j, 1).Value    ' ???　行
''
''                        End If
''                    Next j
''                End If
''            Next i
''        End If


'
' ' 4.件数整理（EOFレコードは除く）
'     CloseingMsg = "統合前件数" & Chr(9) & "＝ " & wrkCnt & Chr(13) & _
'                   "統合後件数" & Chr(9) & "＝ " & newCnt & Chr(13) & _
'                   "  削除件数" & Chr(9) & "＝ " & oldCnt & Chr(13) & _
'                   "  目視件数" & Chr(9) & "＝ " & trnCnt & Chr(13)
'
''     ' Debug.Print cntAllMsg
'
'
'' 終了処理
'    MsgBox CloseingMsg
'
'    Call 後処理_R("住所録マージプログラムは正常終了しました。" & Chr(13) & CloseingMsg)



'If y = 16 Then
'MsgBox y
'' Debug.Print "|wrk:" & wrkY & "=" & Left(wsWrk.Cells(wrkY, 3), 10) & Chr(9) & _
''             "|new:" & newY & "=" & Left(wsNew.Cells(newY, 3), 10) & Chr(9) & _
''             "|old:" & oldY & "=" & Left(wsold.Cells(oldY, 3), 10) & Chr(9) & _
''             "|trn:" & trnY & "=" & Left(wstrn.Cells(trnY, 3), 10)
'End If


