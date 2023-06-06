Attribute VB_Name = "m3_変更レコード処理"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : 変更レコードで�@原簿または�Aarchivesレコードを更新する
' --------------------------------------+-----------------------------------------
' | @moduleName : m3_変更レコード処理
' | @Version    : v1.0.0
' | @update     : 2023/06/06
' | @written    : 2023/06/06
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |  （１）重複レコードは、次のチェックを行い、変更を反映し、新規シート（�@new）へ移動する
' |     �@）(42)姓名keyが同じレコードは、重複レコードであること
' |     �A）(54)識別区分が「1」(�@原簿)もしくは「2」(�Aarchives)のレコードに対し
' |       「3」(�B変更住所録)の変更項目（空白でない項目）をコピーし、新規シート（�@new）へ移動する
' |     �B）コピーする変更項目と現行が同じ内容のものでないことを確認します。
' |         同じ内容であれば、変更なしとしてあつかうこと。
' |     �C）「2」(�Aarchives)のレコードの�B変更住所録の削除日の西暦年が「9999」のレコードは、�@原簿として移動します
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

' �Bnew シートの定義
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i≡x 列　column
Private newY, newYmin, newYmax          As Long             ' j≡y 行　row

' ' 構造体の宣言
' Type cntTbl
'     old                                 As long     ' �@原簿
'     arv                                 As long     ' �Aarchive
'     trn                                 As long     ' �B変更住所録
'     wrk                                 As long     ' 作業レコードの件数=修正前+変更
'     new1                                As long     ' new�@原簿レコードの件数
'     new2                                As long     ' new�Barchivesレコードの件数
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

    Dim sw_change                       as Boolean      ' true ≡ 変更箇所有り　/　False ≡ 〃　無し

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
    
' �Bnew シートの初期値
    Set wsNew = Wb.Worksheets(Range("C_newSheet").Value)                    ' 新住所録シート
    newYmin = YMIN
    newXmin = XMIN
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）

' --------------------------------------+-----------------------------------------
' 同一キーの(42)key姓名が、3→1or2→0　の順に並ぶので、変更項目を 0 レコードを更新し、新住所録シートへコピーする
    for j = wrkYmin to wrkYmax step 3
        sw_change = False
        for i = 6 to 41
            select case i
                case 6 to 15, 23 to 26, 36 to 41        ' 上書き項目
                    If wsWrk.Cells(j, i).Value <> "" Then
                        If wsWrk.Cells(j, i).Value <> wsWrk.Cells(j + 1, i).Value
                            wsWrk.Cells(j + 2, i).Value = wsWrk.Cells(j, i).Value
                            sw_change = true
                        Else
                            wsWrk.Cells(j, CHECKED_X).Value = "same"
                            wsWrk.Cells(j, i).Font.Color = rgbSnow          ' 文字色：スノー    #fafaff#
                            wsWrk.Cells(j, i).Interior.Color = rgbDarkRed   ' 背景色：濃い赤    #00008b#
                        End If
                    End If
                case Else
            end select
        next i
    next j


 



' ' 上書き項目：(6)名前〜(15)方書
'       For r = 6 To 15
'            If wsWrk.Cells(y, r).Value <> "" Then
'                wsWrk.Cells(y + 2, r).Value = wsWrk.Cells(y, r).Value
'            End If
'        Next r
'
'' 上書き項目：(23)その他1〜(26)備考
'       For r = 23 To 26
'            If wsWrk.Cells(y, r).Value <> "" Then
'                wsWrk.Cells(y + 1, r).Value = wsWrk.Cells(y, r).Value
'            End If
'        Next r
'
'' 上書き項目：(36)更新内容〜(41)削除日
'        For r = 36 To 41
'            If wsWrk.Cells(y, r).Value <> "" Then
'                wsWrk.Cells(y + 1, r).Value = wsWrk.Cells(y, r).Value
'            End If
'        Next r
'' 同一キーなので一つ飛ばし
'    y = y + 1
'
'Next_R:
'    Next y
'
'' グループ項目：(16)携帯電話〜(19)会社電話
'        Dim r1 As Long
'        Dim sameCnt As Long
'        sameCnt = 0
'' 変更項目数をカウント
'        For r = 16 To 19
'            If wsWrk.Cells(y, r).Value <> "" Then
'                sameCnt = sameCnt + 1
'            End If
'        Next r
'
'' 変更内容が現行と同一の内容かチェック
'        For r = 16 To 19
'            If wsWrk.Cells(y, r).Value <> "" Then
'                For r1 = 16 To 19
'                    If wsWrk.Cells(y, r).Value = wsWrk.Cells(y + 1, r1).Value Then
'                        wsWrk.Cells(y + 1, r1).Value = ""
'                        sameCnt = sameCnt - 1
'                        Exit For
'                    End If
'                Next r1
'            End If
'        Next r
'' 違う内容のものを空いてるセルにコピー
'        If sameCnt <> 0 Then
'            For r = 16 To 19
'                If wsWrk.Cells(y, r).Value <> "" Then
'                    For r1 = 16 To 19
'                        If wsWrk.Cells(y + 1, r1).Value = "" Then
'                            wsWrk.Cells(y + 1, r1).Value = wsWrk.Cells(y, r).Value
'                            wsWrk.Cells(y, r).Value = ""
'
'                            Exit For
'                        End If
'                    Next r1
'                End If
'            Next r
'        End If
'
'        wsWrk.Rows(y).ClearContents
'        wsWrk.Cells(y + 1, CHECKED_X) = "Mod"
'        wsWrk.Rows(y + 1).Copy Destination:=wsNew.Rows(newY)
'        wsWrk.Rows(y + 1).ClearContents
'
'        Select Case wsNew.Cells(newY, MASTER_X).Value
'            Case 1
'                cnt.new1 = cnt.new1 + 1
'            Case 2
'                cnt.new2 = cnt.new2 + 1
'            Case 3
'                cnt.new3 = cnt.new3 + 1
'            Case Else
'                MsgBox "識別区分エラー=" & wsNew.Cells(newY, MASTER_X).Value
'                End
'        End Select
'        newY = newY + 1
'        y = y + 1   ' 同一keyが二つあるので、一つindexをくりあげる
'
'
'Stop
        
End Sub

' グループ項目：(20)携帯メール〜(22)会社メール

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
''                wsWrk.Cells(y, CHECKED_X) = "�Btrn"
''                y = y + 1
''            Loop
''            wsWrk.Activate
''            wsWrk.Cells(y, CHECKED_X) = "�Btrn"
''            wsTrn.Activate
''            wsWrk.Rows(y).Copy Destination:=wsTrn.Rows(trnY)
''            trnCnt = trnCnt + 1
''            wsTrn.Cells(trnY, 42).Value = trnCnt
''' 最新のレコードを統合コピーの候補「�H-999」とする
''            Rows(trnY & ":" & trnY).Select
''            Selection.Copy
''            trnY = trnY + 1
''            Rows(trnY & ":" & trnY).Select
''            ActiveSheet.Paste
''            Application.CutCopyMode = False     ' コピー状態の解除
'''            Rows(trnY).Insert
''            Cells(trnY, 1) = "�H-999"
''            Cells(trnY, 42) = ""
''            jMax = trnY
''
''' チェックマーク行追加
''            trnY = trnY + 1
''            Cells(trnY, 1) = "???"
''            Rows(trnY & ":" & trnY).Interior.ColorIndex = xlNone    ' 色の初期化
''            trnY = trnY + 1
''' 統合手順の自動化 / 統合候補［�H-999］が空白の項目は、過去のデータから持ってくる
''            For i = INPUTX_FROM To INPUTX_TO + 9
''                If Cells(jMax, i).Value = "" Or _
''                   Cells(jMax, i).Value = "　" Or _
''                   Cells(jMax, i).Value = " " Then
''                    For j = jMin To jMax - 1
''                        If Cells(j, i).Value <> "" Then
''                            Cells(jMax, i).Value = Cells(j, i).Value        ' �H-999 行
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




