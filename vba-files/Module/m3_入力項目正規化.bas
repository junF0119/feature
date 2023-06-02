Attribute VB_Name = "m3_入力項目正規化"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : 入力項目を正規化する（モジュール分割版）
' --------------------------------------+-----------------------------------------
' | @moduleName : m3_入力項目正規化
' | @Version    : v1.0.0
' | @update     : 2023/05/25
' | @written    : 2023/05/25
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |　このJobは、次の処理を行い、登録データの整合性を検証し、自動で修復する。
' |1.1　　Jobの初期処理ととして、更新前のシートがコピーされた時点で、このプログラムのバックアップを保存します。
' |このプログラムは、入力データの保全から入力ファイルは読むだけで、更新は行っていません。
' |万一、元のデータが壊れたときなどには、コピーしたシートから復元することができます。
' |
' |1.2　チェックレベルは、問題なし、自動修復、マニュアル修正でチック欄にマークを付す
' |
' |1.3　　［修正完了］ボタンを押下することで、更新後のシートへ修正後のレコードがコピーされる
' |1.4　　コピー後は、それぞれのシートをバージョンと日付を変更し、それぞれのフォルダーへExportする
' |
' |     プログラム構造
' |         1. 初期処理
' |             初期処理_R()
' |             1.1 既存シートのクリア
' |                 importClear_R()
' |             1.2 外部のマスターのシートを取り込む…… M-①新住所録原簿 / M-②Archives
' |                 importSheet_R()(
' |         2. データの整合性検証…… (53)PrimaryKey / (42)key姓名
' |             キー項目_R()
' |             2.1 シートの配列化 …… ①原簿 / ②archives
' |                 arrSet_R()
' |             2.2 キー項目の重複チェック
' |                 duplicateChk_F()
' |                     quickSort_R()
' |             2.3 キー項目のNull値チェック
' |                 nullKeyChk_F()
' |         3.　データ項目の正規化……キー項目を除く入力データの正規化
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

Private sw_errorChk                     As Boolean          ' true…エラー無し、false…エラー有り

' 構造体の宣言
Type pkeyStruct
    sortKey                             As Variant          ' quick sort用キー
    primaryKey                          As Integer          ' (53)PrimaryKey
    nameKey                             As String           ' (42)key姓名
    sheetName                           As String           ' シート名
    rowAddress                          As Integer          ' レコードの行(row)位置
End Type
Private ary()                           As pkeyStruct      ' 構造体の一元次元動的配列（sortKey,…）
Private j, jMax                         As Long            ' 配列 ary() のインデックス
Private primaryKeyMax                   As Long            ' primaryKeyの最大値

Public Sub 入力項目正規化_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' | @function   : 入力項目を正規化し、入力データのばらつきを訂正する
' --------------------------------------+-----------------------------------------
' | @moduleName : 入力項目正規化_R
' | @Version    : v1.0.0
' | @update     : 2023/05/25
' | @written    : 2023/05/25
' | @remarks
' |     エラーなし…true
' |     エラーあり…false
' |　   エラーがあるときの措置
' |　       ①?欄に　？　マークを付す
' |　       ②ErrCnt をカウント
' |　       ③チェック項目ごとにエラーがあるときは、プログラムを中断し、
' |          マニュアルで元データを修正し、このプログラムを再実行する
' |
' |
' --------------------------------------+-----------------------------------------
    Dim y, yMax                         As Long
    Dim sw_result                       As Boolean
    Dim errMsg                          As String
    Dim err42Cnt                        As Long
    Dim err53Cnt                        As Long

    Dim pkeyStruct                      As pkeyStruct   ' オブジェクトの定義（構造体）
' 構造体の宣言
' Type pkeyStruct
'     sortKey       As Variant          ' quick sort用キー
'     primaryKey    As Integer          ' (53)PrimaryKey
'     nameKey       As string           ' (42)key姓名
'     sheetName     As String           ' シート名
'     rowAddress    As Integer          ' レコードの行(row)位置
' End Type
'
' ---Procedure Division ----------------+-----------------------------------------
'
    sw_result = True            ' エラーなし
    ErrCnt = 0
    primaryKeyMax = 0
' --------------------------------------+-----------------------------------------
' 1.①原簿、②archiveシートからデータをary()テーブルに設定する
' --------------------------------------+-----------------------------------------
' 動的配列の初期化(配列0は使わない、ヘッダー部は除く）
'   dim dataCnt As Long            ' 有効データ件数
'    ReDim ary(dataCnt)
'    dataCnt = SrcYmax + ArvYmax - (YMIN - 1) * 2
'    MsgBox dataCnt & " " & LBound(ary) & "-" & UBound(ary)

    ReDim ary(SrcYmax + arvYmax - (YMIN - 1) * 2 - 1)   ' index= 0 - のため　-1　する
' ①原簿 + ②archives を配列 ary() へコピー
    jMax = -1
    Call arrSet_R(SrcCnt, SrcYmin, SrcYmax, Range("C_SrcSheet").Value)     ' ①原簿シートを配列 ary() へコピー
    Call arrSet_R(arvCnt, arvYmin, arvYmax, Range("C_arvSheet").Value)      ' 続けて、②archives をコピー

' --------------------------------------+-----------------------------------------
' 2.キー項目が null のレコードがないか探す　(42)key姓名 / (53)PrimaryKey
' --------------------------------------+-----------------------------------------
    err42Cnt = 0
    err53Cnt = 0
    If Not nullKeyChk_F(err42Cnt, err53Cnt) Then
        errMsg = "|null レコードがありました。" & Chr(13) _
               & "|(42)key姓名" & Chr(9) & "＝ " & err42Cnt & Chr(13) _
               & "|(53)PrimaryKey" & Chr(9) & "＝ " & err53Cnt & Chr(13) _
               & "| エラー箇所を修復したので確認してください。" & Chr(13) _
               & "| 正しければ、原本を直接手修正するか、コピーしてください。" & Chr(13) _
               & "| このプログラムは強制終了します。"
        MsgBox errMsg
        End
    Else
        StatusBarMsg = "キーが Null のレコードありません。" & Chr(13) & _
                "プログラムを継続します。"
        Call putStatusBar(StatusBarMsg)
    End If

' --------------------------------------+-----------------------------------------
' 3.(53)PrimaryKeyで昇順ソート後、(53)PrimaryKeyの重複チェックを実施
' --------------------------------------+-----------------------------------------
    If Not duplicateChk_F("(53)PrimaryKey") Then
        MsgBox "(53)PrimaryKeyに重複が、" & ErrCnt & " 件あります。修正してください。" & Chr(13) & _
                "このプログラムは、強制終了します。"
        End
    Else
        StatusBarMsg = "(53)PrimaryKeyに重複は、ありません。" & Chr(13) & _
                "プログラムを継続します。"
        Call putStatusBar(StatusBarMsg)
    End If


' --------------------------------------+-----------------------------------------
' 4.(42)key姓名で昇順ソート後、(42)key姓名の重複チェックを実施
' --------------------------------------+-----------------------------------------
    If Not duplicateChk_F("(42)key姓名") Then
        MsgBox "(42)key姓名に重複が、" & ErrCnt & " 件あります。修正してください。" & Chr(13) & _
                "このプログラムは、強制終了します。"
        End
    Else
        StatusBarMsg = "(42)key姓名に重複は、ありません。" & Chr(13) & _
                "プログラムを継続します。"
        Call putStatusBar(StatusBarMsg)
    End If
    
  
End Sub

Private Sub arrSet_R(ByRef p_cnt As Long, ByVal p_yMin As Long, ByVal p_yMax As Long, ByVal p_sheetName As String)
' --------------------------------------+-----------------------------------------
' |  PrimaryKeyをary配列の格納
' --------------------------------------+-----------------------------------------
' 構造体の宣言
' Type pkeyStruct
'     sortKey                             as variant          ' quick sort用キー
'     primaryKey                          As Integer          ' (53)PrimaryKey
'     nameKey                             as string           ' (42)key姓名
'     sheetName                           As String           ' シート名
'     rowAddress                          As Integer          ' レコードの行(row)位置
' End Type

    Dim pkey                            As pkeyStruct   ' オブジェクトの定義（構造体）
'
' ---Procedure Division ----------------+-----------------------------------------
'
    p_cnt = 0
    For j = p_yMin To p_yMax    '(行)
        jMax = jMax + 1
        pkey.primaryKey = Wb.Worksheets(p_sheetName).Cells(j, PRIMARYKEY_X)     ' BA列（53）
        pkey.nameKey = Wb.Worksheets(p_sheetName).Cells(j, PKEY_X)              ' AP列（42）
        pkey.sheetName = p_sheetName
        pkey.rowAddress = j
        ary(jMax) = pkey
' primaryKeyの最大値を得る
        If pkey.primaryKey > primaryKeyMax Then
            primaryKeyMax = pkey.primaryKey
        End If
        p_cnt = p_cnt + 1
    Next j
    
End Sub

Private Function nullKeyChk_F(ByRef p_err42Cnt As Long, ByRef p_err53Cnt As Long)
' --------------------------------------+-----------------------------------------
' | @function  指定フィールドにnullデータがないことの検証
' --------------------------------------+-----------------------------------------
' | @moduleName: nullKeyChk_F
' | @remarks
' |   エラーなし…true
' |   エラーあり…false
' |　エラーがあるときの措置
' |　①?欄に　？　マークを付す
' |　②ErrCnt にカウントする
' |　③候補データを指定フィールにセットする
' |
' | 引数の意味
' | 戻り値：p_err42Cnt: (42)key姓名がnullの件数
' | 戻り値：p_err53Cnt: (53)PrimaryKeyがnullの件数
' |
' | 構造体の宣言
' |  Type pkeyStruct
' |      sortKey                             as variant          ' quick sort用キー
' |      primaryKey                          As Integer          ' (53)PrimaryKey
' |      nameKey                             as string           ' (42)key姓名
' |      sheetName                           As String           ' シート名
' |      rowAddress                          As Integer          ' レコードの行(row)位置
' |  End Type
' |
' --------------------------------------+-----------------------------------------
    Dim y, yMax                         As Long
    Dim sw_result                       As Boolean
    Dim pkeyStruct                      As pkeyStruct   ' オブジェクトの定義（構造体）
    Dim w_contents                      As String
    Dim w_fullName                      As Variant      ' 姓名：藤縄　潤
    Dim w_firstName                     As String       ' 名前：潤
    Dim w_familyName                    As String       ' 姓：藤縄
    
    Dim debugText                       As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
    sw_result = True            ' エラーなし
    p_err42Cnt = 0
    p_err53Cnt = 0
    
' 指定キーのNull値チェック
    For y = LBound(ary) To UBound(ary)
' (42)key姓名のチェック＆修復
        If ary(y).nameKey = "" Then
            p_err42Cnt = p_err42Cnt + 1
            sw_result = False
            w_contents = Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, 6)
            w_fullName = Split(w_contents, " ")   ' (6)名前　区切り文字：半角の空白
            w_familyName = w_fullName(0)
            w_firstName = w_fullName(1)
            Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, PKEY_X) = w_familyName & w_firstName
            Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, CHECKED_X) = Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, CHECKED_X) & "◆"

        End If
        If ary(y).primaryKey = 0 Then
            p_err53Cnt = p_err53Cnt + 1
            sw_result = False
            primaryKeyMax = primaryKeyMax + 1
            Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, PRIMARYKEY_X) = primaryKeyMax
            Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, CHECKED_X) = Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, CHECKED_X) & "●"
            
        End If

    Next y
    
' 戻り値
    nullKeyChk_F = sw_result

End Function

Private Function duplicateChk_F(ByVal p_sortKey As Variant) As Boolean
' --------------------------------------+-----------------------------------------
' | @function  指定フィールドにデータの重複がないことの検証
' --------------------------------------+-----------------------------------------
' | @moduleName: duplicateChk_R
' | @remarks
' |   エラーなし…true
' |   エラーあり…false
' |　エラーがあるときの措置
' |　①?欄に　？　マークを付す
' |　②cntErr にカウントする
' |
' --------------------------------------+-----------------------------------------
    Dim y, yMax                         As Long
    Dim sw_result                       As Boolean
    
    Dim pkeyStruct                      As pkeyStruct   ' オブジェクトの定義（構造体）
    Dim zlogMsg                         As String
    Dim z                               As Long
'
' ---Procedure Division ----------------+-----------------------------------------
'
    sw_result = True            ' エラーなし
    ErrCnt = 0
    Call quickSort_R(ary(), p_sortKey, LBound(ary), UBound(ary), xlAscending)
' sortKeyの重複チェック
    For y = LBound(ary) To UBound(ary) - 1
        If ary(y).sortKey = 0 Or ary(y).sortKey = "" Then
            GoTo SkipRow
        End If
        If ary(y).sortKey = ary(y + 1).sortKey Then
            sw_result = False
            ErrCnt = ErrCnt + 1
            Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, CHECKED_X) = Sheets(ary(y + 1).sheetName).Cells(ary(y + 1).rowAddress, PKEY_X).Value  ' 相手のker姓名を表示
            Sheets(ary(y + 1).sheetName).Cells(ary(y + 1).rowAddress, CHECKED_X) = Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, PKEY_X).Value  '       〃
        End If
SkipRow:
    Next y
    
' 結果
    duplicateChk_F = sw_result

End Function

Private Sub quickSort_R(ByRef argAry() As pkeyStruct, _
                        ByVal p_keyName As String, _
                        ByVal p_lngMin As Long, _
                        ByVal p_lngMax As Long, _
                        Optional sOrder As XlSortOrder = xlAscending)
' --------------------------------------+-----------------------------------------
' | @function  : 構造体配列のクイックソート
' --------------------------------------+-----------------------------------------
' | @moduleName: quickSortk_R
' | @remarks
' | 引数の意味
' | 戻り値：ary() ソート後の配列
' | 引　数：p_keyName: ソートキーの項目名
' | 引　数：p_lngMin：配列添字の最小値LBound
' | 引　数：p_lngMax: 配列添字の最大値UBound
' | 構造体の宣言
' |  Type pkeyStruct
' |      sortKey                             as variant          ' quick sort用キー
' |      primaryKey                          As Integer          ' (53)PrimaryKey
' |      nameKey                             as string           ' (42)key姓名
' |      sheetName                           As String           ' シート名
' |      rowAddress                          As Integer          ' レコードの行(row)位置
' |  End Type
' |
' --------------------------------------+-----------------------------------------
    Dim i                               As Long
    Dim j                               As Long
    Dim vBase                           As pkeyStruct
    Dim vTemp                           As pkeyStruct
    Dim vSwap                           As pkeyStruct
'
' ---Procedure Division ----------------+-----------------------------------------
'
' --------------------------------------+-----------------------------------------
' |  (1) sort key の名前を指定し、それをary()のsortKeyへセットする
' --------------------------------------+-----------------------------------------
    For j = p_lngMin To p_lngMax
        Select Case p_keyName
            Case "(53)PrimaryKey"
                ary(j).sortKey = ary(j).primaryKey     ' BA列（53）
            Case "(42)key姓名"
                ary(j).sortKey = ary(j).nameKey        ' AP列（42）
            Case Else
                MsgBox "プログラムのバグです。" & Chr(13) & _
                "p_keyName=" & p_keyName & "は、定義されていません。" & Chr(13) & _
                "終了します。"
                End
        End Select

    Next j
' バブルソート
    For i = p_lngMax To p_lngMin Step -1
        For j = p_lngMin To i - 1
            If argAry(j).sortKey > argAry(j + 1).sortKey Then
                vSwap = argAry(j)
                argAry(j) = argAry(j + 1)
                argAry(j + 1) = vSwap
            End If
        Next j
    Next i
'

End Sub

'Dim zz As Long
'Dim debugText As String
'Call debug2text("", "open")
'For zz = 0 To SrcYmax + (ArvYmax - (YMIN - 1) * 2) - 1
'debugText = "|bb=" & zz & _
'            "|sortKey=" & ary(zz).sortKey & Chr(9) & _
'            "|primaryKey=" & ary(zz).primaryKey & Chr(9) & _
'            "|nameKey=" & ary(zz).nameKey & Chr(9) & _
'            "|sheetName=" & ary(zz).sheetName & Chr(9) & Chr(9) & _
'            "|rowAddress=" & ary(zz).rowAddress
'Call debug2text(debugText)
'Next zz
'Call debug2text("", "close")
'Stop




