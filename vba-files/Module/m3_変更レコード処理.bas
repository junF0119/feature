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
'' 構造体の宣言
'Public Type cntTbl
'    old                                 As Long     ' ①原簿
'    arv                                 As Long     ' ②archive
'    trn                                 As Long     ' ③変更住所録
'    wrk                                 As Long     ' work
'    new1                                As Long     ' newの原簿レコード
'    new2                                As Long     ' newのarchivwレコード
'    new3                                As Long     ' newの変更住所録で新規レコード
'    mod                                 As Long     ' 変更レコード
'    add                                 As Long     ' 新規レコード
'End Type
'Public Cnt                              As cntTbl

Public Sub m3_変更レコード処理_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     workシートに残ったレコードを更新
' --------------------------------------+-----------------------------------------
'    Dim Cnt                             As cntTbl
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
' 同一キーの(42)key姓名が、3→1or2→0　の順に並ぶので、変更項目で 0 レコードを更新し、新住所録シートへコピーする
    For j = wrkYmin To wrkYmax Step 3
        wsWrk.Range(Cells(j + 1, 1), Cells(j + 1, wrkXmax)).Font.Color = rgbSnow             ' 文字色：スノー    #fafaff#
        wsWrk.Range(Cells(j + 1, 1), Cells(j + 1, wrkXmax)).Interior.Color = rgbDodgerBlue   ' 背景色：ドジャーブルー    #ff901e#

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
                            wsWrk.Cells(j + 1, i).Font.Color = rgbSnow        ' 文字色：スノー    #fafaff#
                            wsWrk.Cells(j + 1, i).Interior.Color = rgbDarkRed ' 背景色：濃い赤    #00008b#
                            wsWrk.Cells(j + 2, i).Font.Color = rgbSnow        ' 文字色：スノー    #fafaff#
                            wsWrk.Cells(j + 2, i).Interior.Color = rgbDarkRed ' 背景色：濃い赤    #00008b#

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
                Case 16
                    Call modifyItem_R(j, i, 16, 19, sw_change)

' グループ項目：(20)携帯メール～(22)会社メール
                Case 20
                    Call modifyItem_R(j, i, 20, 22, sw_change)

                Case Else
            End Select
        Next i
        
' 変更した項目があるときは、管理項目も更新する
        If sw_change Then
            For i = 36 To 41
                wsWrk.Cells(j + 2, i).Value = wsWrk.Cells(j, i).Value
            Next i
        End If
        
 ' 更新した行を新住所録シートへコピー
        wsWrk.Cells(j, CHECKED_X) = "trn"
        wsWrk.Cells(j + 1, CHECKED_X) = "before"
        newYmax = newYmax + 1
        wsWrk.Rows(j + 2).Copy Destination:=wsNew.Rows(newYmax)

'        wsWrk.Range(Cells(j + 2, 1), Cells(j + 2, wrkXmax)).Font.Color = rgbSnow             ' 文字色：スノー    #fafaff#
'        wsWrk.Range(Cells(j + 2, 1), Cells(j + 2, wrkXmax)).Interior.Color = rgbDodgerBlue   ' 背景色：ドジャーブルー    #ff901e#
    
        Select Case wsWrk.Cells(j + 1, MASTER_X).Value
            Case 1
                Cnt.new1 = Cnt.new1 + 1
            Case 2
                Cnt.new2 = Cnt.new2 + 1
            Case 3
                Cnt.new3 = Cnt.new3 + 1
            Case Else
                MsgBox "識別区分エラー=" & wsNew.Cells(newY, MASTER_X).Value
                End
        End Select
        
        If sw_change Then
            wsNew.Cells(newYmax, CHECKED_X) = "Modify"
            Cnt.mod = Cnt.mod + 1
        Else
            wsNew.Cells(newYmax, CHECKED_X) = "Add"
            Cnt.add = Cnt.add + 1
        End If
    
    Next j

End Sub


Private Sub modifyItem_R(ByVal p_j As Long _
                       , ByVal p_i As Long _
                       , ByVal p_from As Long _
                       , ByVal p_to As Long _
                       , ByRef p_modifySw As Boolean)
' --------------------------------------+-----------------------------------------
' | @function   : 変更レコードで更新
' | @moduleName : modifyItem_R
' | @remarks
' | 引数の意味
' | 引　数：p_j           行位置
' | 引　数：p_i           列位置
' | 引　数：p_from        列位置の開始
' | 引　数：p_to          列位置の終了
' | 戻り値：p_modifySw    変更有り ≡ True  、変更なし ≡ False
' |
' --------------------------------------+-----------------------------------------
    Dim Cnt                             As cntTbl
    Dim x, xx                          As Long
    Dim sameCnt                         As Long
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
    sameCnt = 0
    For x = p_from To p_to
        If wsWrk.Cells(p_j, x).Value <> "" Then
            sameCnt = sameCnt + 1
        End If
    Next x

' 変更内容が現行と同一の内容かチェック
    For x = p_from To p_to
        If wsWrk.Cells(p_j, x).Value <> "" Then
            For xx = p_from To p_to
                If wsWrk.Cells(p_j, x).Value = wsWrk.Cells(p_j + 2, xx).Value Then
'                    wsWrk.Cells(p_j, x).Value = ""                         ' 同じ値が既にあるので、変更項目は、消去
                    wsWrk.Cells(p_j, x).Font.Color = rgbNavy                ' 文字色：ネイビー  #800000#
                    wsWrk.Cells(p_j, x).Interior.Color = rgbSnow            ' 背景色：スノー    #fafaff#
                    wsWrk.Cells(p_j + 1, x).Font.Color = rgbNavy              ' 文字色：ネイビー  #800000#
                    wsWrk.Cells(p_j + 1, x).Interior.Color = rgbSnow          ' 背景色：スノー    #fafaff#
                    sameCnt = sameCnt - 1
                    Exit For
                End If
            Next xx
        End If
    Next x

' 違う内容のものを空いてるセルにコピー
    If sameCnt <> 0 Then
        For x = p_from To p_to
            If wsWrk.Cells(p_j, x).Value <> "" Then
                For xx = p_from To p_to
                    If wsWrk.Cells(p_j + 2, xx).Value = "" Then
                        wsWrk.Cells(p_j + 2, xx).Value = wsWrk.Cells(p_j, x).Value
'                        wsWrk.Cells(p_j, x).Value = ""                 ' 更新できたので、消去
                        wsWrk.Cells(p_j, x).Font.Color = rgbSnow          ' 文字色：スノー    #fafaff#
                        wsWrk.Cells(p_j, x).Interior.Color = rgbDarkRed   ' 背景色：濃い赤    #00008b#
                        wsWrk.Cells(p_j + 1, xx).Font.Color = rgbSnow        ' 文字色：スノー    #fafaff#
                        wsWrk.Cells(p_j + 1, xx).Interior.Color = rgbDarkRed ' 背景色：濃い赤    #00008b#
                        wsWrk.Cells(p_j + 2, xx).Font.Color = rgbSnow        ' 文字色：スノー    #fafaff#
                        wsWrk.Cells(p_j + 2, xx).Interior.Color = rgbDarkRed ' 背景色：濃い赤    #00008b#
                        p_modifySw = True
                        sameCnt = sameCnt - 1
                        Exit For
                    End If
                Next xx
            End If
        Next x
    End If

    If sameCnt <> 0 Then
        MsgBox "変更しきれなかった項目が残っています＝" & sameCnt
        Stop
    End If

End Sub
        
