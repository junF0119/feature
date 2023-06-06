Attribute VB_Name = "CM01_メニュー処理"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : Jobを実行させるときの標準的メニュー処理（標準版）
' --------------------------------------+-----------------------------------------
' | @moduleName : CM01_メニュー処理
' | @Version    : v1.2.0
' | @updaten    : 2023/05/31
' | @written    : 2023/04/21
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |  「DocInfo(削除不可)」シートを参照
' --------------------------------------+-----------------------------------------
' |  命名規則の統一
' |     Public変数  先頭を大文字    ≡ pascalCase
' |     private変数 先頭を小文字    ≡ camelCase
' |     定数        全て大文字、区切り文字は、アンダースコア(_) ≡ snake_case
' |     引数        接頭語(p_)をつけ、camelCaseに準ずる
' --------------------------------------+-----------------------------------------
'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
'
'
' --------------------------------------------------------------------------------
'   ※private変数(当該モジュール内のプロシージャ間で共有）
'     頭文字を小文字にする
' 個別定義

Public Sub m1_メニュー処理_R(p_menu As Integer)
' --------------------------------------+-----------------------------------------
' |     メイン処理
' |  [メニュー]sheetのボタンのクリックで、メインプログラムは呼び出される
' |　引数で渡されたメニュー番号 Menu で処理を識別し実行する
' |
' |　Ref. ボタンのマクロの書式例
' |　　　'#tm89.01-封筒宛先印刷-v9.3.4-20201028.xlsm'!'封筒印刷 1'
' |　　　'#tm89.01-封筒宛先印刷-v9.3.4-20201028.xlsm'!'封筒印刷 2'
' |
' |　Ref. VBAのコーディング例
' |　　　Public Sub 封筒印刷(Menu As Integer)
' |
' |
' --------------------------------------+-----------------------------------------

 '
' ---Procedure Division ----------------+-----------------------------------------

    MenuNum = p_menu
    NumCnt = 0
    OpeningMsg = ""
    CloseingMsg = ""
    StatusBarMsg = ""
  
    Select Case MenuNum
        Case 1
            Call m0_新住所録原簿更新処理_R("")
        Case Else
            IsMsgPush ("プログラムのバグです。 中止します。")
            End
    End Select

End Sub


