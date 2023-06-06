Attribute VB_Name = "m3_変更レコード処理"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : 変更レコードで①原簿または②archivesレコードを更新する
' --------------------------------------+-----------------------------------------
' | @moduleName : m3_変更レコード処理
' | @Version    : v1.0.0
' | @update     : 2023/06/06
' | @written    : 2023/06/06
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |  （１）重複レコードは、次のチェックを行い、変更を反映し、新規シート（①new）へ移動する
' |     ⅰ）(42)姓名keyが同じレコードは、重複レコードであること
' |     ⅱ）(54)識別区分が「1」(①原簿)もしくは「2」(②archives)のレコードに対し
' |       「3」(③変更住所録)の変更項目（空白でない項目）をコピーし、新規シート（①new）へ移動する
' |     ⅲ）コピーする変更項目と現行が同じ内容のものでないことを確認します。
' |         同じ内容であれば、変更なしとしてあつかうこと。
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

' ③new シートの定義
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i≡x 列　column
Private newY, newYmin, newYmax          As Long             ' j≡y 行　row

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


Public Sub m3_変更レコード処理_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     workシートに残ったレコードを更新
' --------------------------------------+-----------------------------------------
    Dim cnt                             As cntTbl
    Dim x, y, z                         As Long
    Dim CloseingMsg                     As String
    Dim w_rate, w_mod                   As Integer      ' 進捗率 / 表示間隔
    Dim i, iMin, iMax                   As Long         ' 同一レコードの範囲(列 col x)
    Dim j, jMin, jMax                   As Long         ' 同一レコードの範囲(行 row y)
    Dim r                               As Long         ' 変更項目の列番号

    Dim sw_change                       As Boolean      ' true ≡ 変更箇所有り　/　False ≡ 〃　無し

'
' ---Procedure Division ----------------+-----------------------------------------
'
' オブジェクト変数の定義（共通）
    Set Wb = ThisWorkbook
    ' 表の大きさを得る
    ' 作業シート（このシート）の初期値
    Set wsWrk = Wb.Worksheets("work")                                       ' 作業用シートのため固定
    wrkYmin = YMIN
    wrkXmin = XMIN
    wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）3列目（"C")名前列で計測
    wrkXmax = wsWrk.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）ヘッダー行 3行目で計測
    
' ③new シートの初期値
    Set wsNew = Wb.Worksheets(Range("C_newSheet").Value)                    ' 新住所録シート
    newYmin = YMIN
    newXmin = XMIN
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）

' --------------------------------------+-----------------------------------------
' 同一キーの(42)key姓名が、3→1or2→0　の順に並ぶので、変更項目を 0 レコードを更新し、新住所録シートへコピーする
    For j = wrkYmin To wrkYmax Step 3
        sw_change = False
        For i = 6 To 41
            Select Case i
' 上書き項目：(6)名前～(15)方書、(23)その他1～(26)備考
                Case 6 To 15, 23 To 26                                      
                    If wsWrk.Cells(j, i).Value <> "" Then
                        If wsWrk.Cells(j, i).Value <> wsWrk.Cells(j + 1, i).Value Then
                            wsWrk.Cells(j + 2, i).Value = wsWrk.Cells(j, i).Value
                            wsWrk.Cells(j, i).Font.Color = rgbSnow          ' 文字色：スノー    #fafaff#
                            wsWrk.Cells(j, i).Interior.Color = rgbDarkRed   ' 背景色：濃い赤    #00008b#
                            sw_change = True
                        End If
                    End If
' 管理項目：(36)更新内容～(41)削除日
                Case 36 To 41                                               
                    If wsWrk.Cells(j, i).Value <> "" Then
                        If wsWrk.Cells(j, i).Value <> wsWrk.Cells(j + 1, i).Value Then
                            
                        End If
                    End If
' グループ項目：(16)携帯電話～(19)会社電話
                Case 16 To 19
                    call modifyItem_R(j, 16, 19, sw_change)


' グループ項目：(20)携帯メール～(22)会社メール                          
                Case 20 To 22
                    call modifyItem_R(j, 20, 22, sw_change)

                Case Else
            End Select
        Next i
' 変更した項目があるときは、管理項目も更新する
        If sw_change Then
            For i = 36 To 41
                wsWrk.Cells(j + 2, i).Value = wsWrk.Cells(j, i).Value
            Next i
        End If
    Next j

End Sub


Private Sub modifyItem_R(ByVal p_j As long _
                       , ByVal p_from As long _
                       , ByVal p_to As long _
                       , ByRef p_modifySw As Boolean)
' --------------------------------------+-----------------------------------------
' | @function   : 変更レコードで更新
' | @moduleName : modifyItem_R
' | @remarks
' | 引数の意味
' | 引　数：p_j           行位置
' | 引　数：p_from        列位置の開始
' | 引　数：p_to          列位置の終了
' | 戻り値：p_modifySw    変更有り ≡ True  、変更なし ≡ False
' |
' --------------------------------------+-----------------------------------------
    Dim i, ii                           As long
    Dim sameCnt                         As long = 0
'
' ---Procedure Division ----------------+-----------------------------------------
'
' --------------------------------------+-----------------------------------------
'   グループ項目：(16)携帯電話～(19)会社電話
'   グループ項目：(20)携帯メール～(22)会社メール
' --------------------------------------+-----------------------------------------

'                       wsWrk.Cells(j, i).Font.Color = rgbSnow          ' 文字色：スノー    #fafaff#
'                       wsWrk.Cells(j, i).Interior.Color = rgbDarkRed   ' 背景色：濃い赤    #00008b#

'                        Else
'                            wsWrk.Cells(j, CHECKED_X).Value = "same"
' 変更項目数をカウント
    For i = p_from To p_to
        If wsWrk.Cells(p_j, i).Value <> "" Then
            sameCnt = sameCnt + 1
        End If
    Next i

' 変更内容が現行と同一の内容かチェック
    For i = p_from To p_to
        If wsWrk.Cells(p_j, i).Value <> "" Then
            For ii = p_from To p_to
                If wsWrk.Cells(p_j, i).Value = wsWrk.Cells(p_j + 1, ii).Value Then
                    wsWrk.Cells(p_j, i).Value = ""   '同じ値が既にあるので、変更項目は、消去
                    wsWrk.Cells(j, i).Interior.Color = rgbSnow      ' 背景色：スノー    #fafaff#
                    sameCnt = sameCnt - 1
                    Exit For
                End If
            Next ii
        End If
    Next i

' 違う内容のものを空いてるセルにコピー
    If sameCnt <> 0 Then
        For i = p_from To p_to
            If wsWrk.Cells(p_j, i).Value <> "" Then
                For ii = p_from To p_to
                    If wsWrk.Cells(p_j + 1, ii).Value = "" Then
                        wsWrk.Cells(p_j + 1, ii).Value = wsWrk.Cells(p_j, i).Value
                        wsWrk.Cells(p_j, i).Value = ""
                        sameCnt = sameCnt - 1
                        Exit For
                    End If
                Next ii
            End If
        Next i
    End If

    if sameCnt <> 0 Then
        msgbox "変更しきれなかった項目が残っています＝"　& sameCnt
        stop
    end if
    
'    wsWrk.Rows(j).ClearContents
    wsWrk.Cells(p_j + 1, CHECKED_X) = "Mod"
    newYmax = newYmax + 1
    wsWrk.Rows(p_j + 1).Copy Destination:=wsNew.Rows(newYmax)
'    wsWrk.Rows(p_j + 1).ClearContents

    Select Case wsNew.Cells(newYmax, MASTER_X).Value
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

       

Stop
        


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




