Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : 住所録更新処理（モジュール分割版）
' --------------------------------------+-----------------------------------------
' | @moduleName : m0_住所録更新処理
' | @Version    : v1.0.1
' | @update     : 2023/06/01
' | @written    : 2023/05/30
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
' | プログラム構造
' |     1. 初期処理
' |         1.1 既存シートのクリア
' |             importClear_R()
' |         1.2 外部のマスターのシートを取り込む…… M-①新住所録原簿 / M-②Archives
' |             importSheet_R()
' |
' |     2. 重複キーチェック
' |         2.1 重複チェック…… (53)PrimaryKey / (42)key姓名
' |             keyCheck_F()
' |                 arrSet_R()
' |                 duplicateChk_F()
' |                     quickSort_R()
' |
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
Public Const PKEY_RNG                   As String = "AP3"   ' Keyのセル番号
Public Const PKEY_X                     As Long = 42        ' Keyの列番号"AP"
Public Const PSEIMEI_X                  As Long = 6         ' 作業域の最大行数計測の列番号"C"(名前)
Public Const PDEL_X                     As Long = 41        ' 削除日の列番号"AO"
Public Const XMIN                       As Long = 1         ' 開始列
Public Const XMAX                       As Long = 53        ' 最終列
Public Const YMIN                       As Long = 4         ' 開始行　∵ヘッダー部を除く
Public Const yMax                       As Long = 1999      ' 最大行　∵このプログラムであつかう最大行
Public Const INPUTX_FROM                As Long = 6         ' 入力項目開始列"F"
Public Const INPUTX_TO                  As Long = 26        ' 入力項目終了列"Z"
Public Const CHECKED_X                  As Long = 1         ' チェック欄（自由）
Public Const PRIMARYKEY_X               As Long = 53        ' PrimaryKeyの列"BA"
Public Const MASTER_RNG                 As String = "BB3"   ' workシート専用「識別区分」のセル番号"BB3"
Public Const MASTER_X                   As Long = 54        ' workシート専用「識別区分」の列番号"BB"

' ①原簿シートの定義
Public Wb                               As Workbook         ' このブック
Public wsSrc                            As Worksheet
Public SrcX, SrcXmin, SrcXmax           As Long             ' i≡x 列　column
Public SrcY, SrcYmin, SrcYmax           As Long             ' j≡y 行　row
Public SrcCnt                           As Long             ' レコード全件の件数
' ②archives シートの定義 ∵ 削除レコード
Public wsArv                            As Worksheet
Public arvX, arvXmin, arvXmax           As Long             ' i≡x 列　column
Public arvY, arvYmin, arvYmax           As Long             ' j≡y 行　row
Public arvCnt                           As Long             ' 削除レコードの件数
' ③目視 シートの定義
Public WsEye                            As Worksheet
Public EyeX, EyeXmin, EyeXmax           As Long             ' i≡x 列　column
Public EyeY, EyeYmin, EyeYmax           As Long             ' j≡y 行　row
Public EyeCnt                           As Long             ' 目視レコードの件数
' debug2Fileのfil番号
Public FileNum                          As Long
' --------------------------------------+-----------------------------------------
' 構造体の宣言
Type cntTbl
    old                                 As Long     ' ①原簿
    arv                                 As Long     ' ②archive
    trn                                 As Long     ' ③変更住所録
    wrk                                 As Long     ' work
    new1                                As Long     ' newの原簿レコード
    new2                                As Long     ' newのarchivwレコード
    new3                                As Long     ' newの変更住所録で新規レコード
End Type
' --------------------------------------+-----------------------------------------
'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +

Public Sub m0_住所録更新処理_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |
' | プログラム構造
' |     1. 初期処理
' |         1.1 既存シートのクリア
' |             importClear_R()
' |         1.2 外部のマスターのシートを取り込む…… M-①新住所録原簿 / M-②Archives
' |             importSheet_R()
' |
' |     2. キー項目のチェック…… (53)PrimaryKey / (42)key姓名
' |         2.1 重複チェック
' |             keyCheck_F()
' |                 arrSet_R()
' |                 duplicateChk_F()
' |                     quickSort_R()
' |         2.2 Null値チェック
' |
' |
' --------------------------------------+-----------------------------------------


'
' ---Procedure Division ----------------+-----------------------------------------
'
    Call m1_初期化処理_R("")
    
    Call m2_レコード振分処理_R("")

    Call m3_変更レコード処理_R("")
    
    Call m9_終了処理_R("")
    

End Sub


'If y = 16 Then
'MsgBox y
'Debug.Print "|wrk:" & wrkY & "=" & Left(wsWrk.Cells(wrkY, 3), 10) & Chr(9) & _
'            "|new:" & newY & "=" & Left(wsNew.Cells(newY, 3), 10) & Chr(9) & _
'            "|arv:" & arvY & "=" & Left(wsArv.Cells(arvY, 3), 10) & Chr(9) & _
'            "|eye:" & eyeY & "=" & Left(wsEye.Cells(eyeY, 3), 10)
'End If



