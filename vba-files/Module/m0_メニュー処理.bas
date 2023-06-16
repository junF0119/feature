Attribute VB_Name = "m0_メニュー処理"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : Jobを実行させるときの標準的メニュー処理（標準版）
' --------------------------------------+-----------------------------------------
' | @moduleName : m0_メニュー処理
' | @Version    : v1.2.0
' | @updaten    : 2023/05/31
' | @written    : 2023/04/21
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |
' | プログラム構造
' |     1. 前処理（システム共通）
' |         1.1 システムに関するPublic変数の取得
' |         1.2 オープニングメッセージの表示
' |         1.3 処理前の当該ブックのバックアップ出力
' |     2. 前処理2（アプリケーション共通）
' |         2.1 定数の設定
' |
' |
' --------------------------------------+-----------------------------------------
' |  命名規則の統一
' |     Public変数  先頭を大文字    ≡ pascalCase
' |     private変数 先頭を小文字    ≡ camelCase
' |     定数        全て大文字、区切り文字は、アンダースコア(_) ≡ snake_case
' |     引数        接頭語(p_)をつけ、camelCaseに準ずる
' --------------------------------------+-----------------------------------------
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' アプリケーション定数の定義
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

' --------------------------------------+-----------------------------------------
' 構造体の宣言
Public Type cntTbl
    old                                 As Long     ' ①原簿
    arv                                 As Long     ' ②archive
    trn                                 As Long     ' ③変更住所録
    wrk                                 As Long     ' work
    new1                                As Long     ' newの原簿レコード
    new2                                As Long     ' newのarchivwレコード
    new3                                As Long     ' newの変更住所録で新規レコード
    mod                                 As Long     ' 変更レコード
    Add                                 As Long     ' 新規レコード
End Type
Public Cnt                              As cntTbl
'
' --------------------------------------------------------------------------------
'   ※private変数(当該モジュール内のプロシージャ間で共有）
'     頭文字を小文字にする
' 個別定義

Public Sub メニュー処理(p_menu As Integer)
' --------------------------------------+-----------------------------------------
' |     メイン処理
' |  [メニュー]sheetのボタンのクリックで、メインプログラムは呼び出される
' |　引数で渡されたメニュー番号 Menu で処理を識別し実行する
' |
' | プログラム構造
' |     1. 前処理（システム共通）
' |         1.1 システムに関するPublic変数の取得
' |         1.2 オープニングメッセージの表示
' |         1.3 処理前の当該ブックのバックアップ出力
' |     2. 前処理2（アプリケーション共通）
' |         2.1 定数の設定
' |
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' |
' |　Ref. ボタンのマクロの書式例
' |　　　'#tm89.01-封筒宛先印刷-v9.3.4-20201028.xlsm'!'封筒印刷 1'
' |　　　'#tm89.01-封筒宛先印刷-v9.3.4-20201028.xlsm'!'封筒印刷 2'
' |
' |　Ref. VBAのコーディング例
' |　　　Public Sub 封筒印刷(Menu As Integer)
' |
' --------------------------------------+-----------------------------------------

'
' ---Procedure Division ----------------+-----------------------------------------
'
    MenuNum = p_menu
' 1. 前処理（システム共通）
    NumCnt = 0
    OpeningMsg = ""
    CloseingMsg = ""
    StatusBarMsg = ""
    Call 前処理_R("")
    Cnt.old = 0                         ' ①原簿
    Cnt.arv = 0                         ' ②archive
    Cnt.trn = 0                         ' ③変更住所録
    Cnt.wrk = 0                         ' work
    Cnt.new1 = 0                        ' newの原簿レコード
    Cnt.new2 = 0                        ' newのarchivwレコード
    Cnt.new3 = 0                        ' newの変更住所録で新規レコード
    Cnt.mod = 0                         ' 変更レコード
    Cnt.Add = 0                         ' 新規レコード
  
    Select Case MenuNum

        Case 1          ' Step1 新住所録の更新処理
        
            Call m1_初期化処理_R("")
            Call m2_レコード振分処理_R("")
            Call m3_変更レコード処理_R("")
            Call m9_終了処理_R("")
            
        Case 2          ' Step2 更新済み新住所録Export
        
            MsgBox "Step2が呼ばれました。"
            
        Case Else
            IsMsgPush ("プログラムのバグです。 中止します。")
            End
    End Select

End Sub



