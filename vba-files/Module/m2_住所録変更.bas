Attribute VB_Name = "m2_住所録変更"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : ③変更住所録で①原簿と②archivesを更新する
' --------------------------------------+-----------------------------------------
' | @moduleName : m2_住所録変更
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
' ③new シートの定義
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i≡x 列　column
Private newY, newYmin, newYmax          As Long             ' j≡y 行　row
Private new1Cnt                         As Long             ' new①原簿レコードの件数
Private new2Cnt                         As Long             ' new③archivesレコードの件数

' ' 構造体の宣言
' Type cntTbl
'     old                                 As long     ' ①原簿
'     arv                                 As long     ' ②archive
'     trn                                 As long     ' ③変更住所録
'     wrk                                 As long     ' work
'     new1                                As long     ' newの原簿レコード
'     new2                                As long     ' newのarchivwレコード
' End Type
' dim cnt                             as cntTbl


Public Sub m2_住所録変更_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     workシートの前後のレコードを比較
' --------------------------------------+-----------------------------------------
    dim cnt                             as cntTbl
    Dim x, y                            As Long
    Dim CloseingMsg                     As String
    Dim w_rate, w_mod                   As Integer      ' 進捗率 / 表示間隔
    Dim i, iMin, iMax                   As Long         ' 同一レコードの範囲(列 col x)
    Dim j, jMin, jMax                   As Long         ' 同一レコードの範囲(行 row y)
    dim r                               as long         ' 変更項目の列番号
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
    Set wsNew = Wb.Worksheets(Range("C_newSheet").Value)    ' 新住所録シート
    newYmin = YMIN
    newXmin = XMIN
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）
    
    For y = wrkYmin To wrkYmax
        w_rate = Int((y - YMIN) / (wrkYmax - YMIN) * 100)
        If w_rate >= 10 Then
            w_mod = w_rate Mod 10
            If w_mod = 0 Then
                Call バー表示("進捗率= " & CStr(w_rate) & "%")
            End If
        End If
        
        If wsWrk.Cells(y, PKEY_X).Value <> wsWrk.Cells(y + 1, PKEY_X).Value Then
            wsWrk.Rows(y).Copy Destination:=wsNew.Rows(newY)
            wsWrk.Activate
            wsWrk.Cells(y, CHECKED_X) = "①new"
            newY = newY + 1
            if wsWrk.Cells(y, MASTER_X).Value = 1 Then
                cnt.new1 = cnt.new1 + 1
            Else
                cnt.new2 = cnt.new2 + 1
            end If
            goto Next_R
        end if
' 同一keyの更新処理
' 上書き項目：(6)名前～(15)方書
' 上書き項目：(23)その他1～(26)備考
' 上書き項目：(36)更新内容～(41)削除日
        for r = 6 to 15, 23 to 26, 36 to 41
            If wsWrk.Cells(y, r).Value <> "" then
                wsWrk.Cells(y + 1, r).Value = wsWrk.Cells(y, r).Value
            end if
        next r 
' グループ項目：(16)携帯電話～(19)会社電話
        for r = 16 to 19
            If wsWrk.Cells(y, r).Value <> "" then
                dim r1 as long
                dim sw_copy as boolean
                sw_copy = false
                for r1 = 16 to 19
                    if wsWrk.Cells(y + 1, r1).Value = "" then
                        wsWrk.Cells(y + 1, r1).Value = wsWrk.Cells(y, r).Value
                        sw_copy = true
                        exit for
                    end if
                next r1
            end if
        next r

stop

' グループ項目：(20)携帯メール～(22)会社メール






            jMin = trnY                 ' 同一keyの最初の行(Row, y)
            
            Do While wsWrk.Cells(y, PKEY_X).Value = wsWrk.Cells(y + 1, PKEY_X).Value
                wsWrk.Rows(y).Copy Destination:=wsTrn.Rows(trnY)
                trnCnt = trnCnt + 1
                wsTrn.Cells(trnY, 42).Value = trnCnt
                trnY = trnY + 1
                
                
                wsWrk.Activate
                wsWrk.Cells(y, CHECKED_X) = "③trn"
                y = y + 1
            Loop
            wsWrk.Activate
            wsWrk.Cells(y, CHECKED_X) = "③trn"
            wsTrn.Activate
            wsWrk.Rows(y).Copy Destination:=wsTrn.Rows(trnY)
            trnCnt = trnCnt + 1
            wsTrn.Cells(trnY, 42).Value = trnCnt
' 最新のレコードを統合コピーの候補「⑨-999」とする
            Rows(trnY & ":" & trnY).Select
            Selection.Copy
            trnY = trnY + 1
            Rows(trnY & ":" & trnY).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False     ' コピー状態の解除
'            Rows(trnY).Insert
            Cells(trnY, 1) = "⑨-999"
            Cells(trnY, 42) = ""
            jMax = trnY
           
' チェックマーク行追加
            trnY = trnY + 1
            Cells(trnY, 1) = "???"
            Rows(trnY & ":" & trnY).Interior.ColorIndex = xlNone    ' 色の初期化
            trnY = trnY + 1
' 統合手順の自動化 / 統合候補［⑨-999］が空白の項目は、過去のデータから持ってくる
            For i = INPUTX_FROM To INPUTX_TO + 9
                If Cells(jMax, i).Value = "" Or _
                   Cells(jMax, i).Value = "　" Or _
                   Cells(jMax, i).Value = " " Then
                    For j = jMin To jMax - 1
                        If Cells(j, i).Value <> "" Then
                            Cells(jMax, i).Value = Cells(j, i).Value        ' ⑨-999 行
                            Cells(jMax + 1, i).Value = Cells(j, 1).Value    ' ???　行
                            
                        End If
                    Next j
                End If
            Next i
        End If
Next_R:
    Next y



 ' 4.件数整理（EOFレコードは除く）
     CloseingMsg = "統合前件数" & Chr(9) & "＝ " & wrkCnt & Chr(13) & _
                   "統合後件数" & Chr(9) & "＝ " & newCnt & Chr(13) & _
                   "  削除件数" & Chr(9) & "＝ " & oldCnt & Chr(13) & _
                   "  目視件数" & Chr(9) & "＝ " & trnCnt & Chr(13)
                    
'     ' Debug.Print cntAllMsg
    

' 終了処理
    MsgBox CloseingMsg
    
    Call 後処理_R("住所録マージプログラムは正常終了しました。" & Chr(13) & CloseingMsg)

End Sub

'If y = 16 Then
'MsgBox y
'' Debug.Print "|wrk:" & wrkY & "=" & Left(wsWrk.Cells(wrkY, 3), 10) & Chr(9) & _
''             "|new:" & newY & "=" & Left(wsNew.Cells(newY, 3), 10) & Chr(9) & _
''             "|old:" & oldY & "=" & Left(wsold.Cells(oldY, 3), 10) & Chr(9) & _
''             "|trn:" & trnY & "=" & Left(wstrn.Cells(trnY, 3), 10)
'End If


