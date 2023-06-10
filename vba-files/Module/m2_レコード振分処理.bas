Attribute VB_Name = "m2_レコード振分処理"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : 変更対象でないレコードを�@原簿と�Aarchivesから抜き出す
' --------------------------------------+-----------------------------------------
' | @moduleName : m2_レコード振分処理
' | @Version    : v1.0.0
' | @update     : 2023/06/02
' | @written    : 2023/06/02
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |  （１）単独レコードを新規シート（�@new）へ移動する
' |  （２）重複レコードは、次のチェックを行い、変更を反映し、新規シート（�@new）へ移動する
' |     �@）(42)姓名keyが同じレコードは、重複レコードであること
' |     �A）(54)識別区分が「1」(�@原簿)もしくは「2」(�Aarchives)のレコードに対し
' |       「3」(�B変更住所録)の変更項目（空白でない項目）をコピーし、新規シート（�@new）へ移動する
' |     �B）コピーする変更項目と同じ内容のものであることを確認します。
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
' Private wrkCnt                          As Long             ' 作業レコードの件数=修正前+変更

' �Bnew シートの定義
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i≡x 列　column
Private newY, newYmin, newYmax          As Long             ' j≡y 行　row
' Private new1Cnt                         As Long             ' new�@原簿レコードの件数
' Private new2Cnt                         As Long             ' new�Barchivesレコードの件数

'' 構造体の宣言
'Public Type cntTbl
'    old                                 As Long     ' �@原簿
'    arv                                 As Long     ' �Aarchive
'    trn                                 As Long     ' �B変更住所録
'    wrk                                 As Long     ' work
'    new1                                As Long     ' newの原簿レコード
'    new2                                As Long     ' newのarchivwレコード
'    new3                                As Long     ' newの変更住所録で新規レコード
'    mod                                 As Long     ' 変更レコード
'    add                                 As Long     ' 新規レコード
'End Type
'Public Cnt                              As cntTbl


Public Sub m2_レコード振分処理_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     workシートの前後のレコードを比較
' --------------------------------------+-----------------------------------------
'    Dim Cnt                             As cntTbl
    Dim x, y, z                         As Long
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
    
' �Bnew シートの初期値
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
                    Cnt.new1 = Cnt.new1 + 1
                Case 2
                    Cnt.new2 = Cnt.new2 + 1
                Case 3
                    Cnt.new3 = Cnt.new3 + 1
                    wsNew.Cells(newY, CHECKED_X) = "add"
                Case Else
                    MsgBox "識別区分エラー=" & wsNew.Cells(newY, MASTER_X).Value
                    End
            End Select
            newY = newY + 1
        Else
            y = y + 1   ' 同一キーなので1行スキップ
        End If
    Next y


' 同一keyの更新処理
' step1:(42)key姓名昇順、(54)識別区分(1…�@原簿, 2…�Aarchives, 3…�B変更住所録)降順にソートし、単独レコードの行を除く
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
    
    wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）
' (54)識別区分を 0 の afterレコードをコピーする
    Dim addY As Long          ' 追加する行
    wrkY = wrkYmax

    For y = wrkYmin To wrkYmax - 1 Step 2
        If wsWrk.Cells(y, PKEY_X).Value = wsWrk.Cells(y + 1, PKEY_X).Value Then
            wsWrk.Cells(y + 1, CHECKED_X) = "before"
            wrkY = wrkY + 1
            wsWrk.Rows(y + 1).Copy Destination:=wsWrk.Rows(wrkY)
            wsWrk.Cells(wrkY, CHECKED_X) = "after"
            wsWrk.Cells(wrkY, Range(MASTER_RNG).Column) = 0
        Else
            MsgBox "重複キーは、２レコードのルール違反。要確認！！"
            Stop
            End
        End If
    Next y

   wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）
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

End Sub

