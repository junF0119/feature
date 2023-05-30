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
Public P_backupFile                     As String       ' 実行前ファイルの保存用フォルダのフルパス
Public P_fullPath                       As String       ' 実行Excelのフルパス+ファイル名 ≡ Thisworkbook
Public P_pathName                       As String       ' 実行Excelのフルパス
Public P_fileName                       As String       ' 実行Excelのファイル名
' ディレクトリ構造のパスと名前
Public P_rootPath                       As String       ' システムフォルダの親ディレクトリのルートパス
Public P_sysPath                        As String       ' システムフォルダまでのフルパス
Public P_sysName                        As String       ' システムフォルダの名前
Public P_subSysPath                     As String       ' サブシステムフォルダまでのフルパス
Public P_subSysName                     As String       ' サブシステムフォルダの名前
' 実行プログラムの情報
Public P_sysSymbol                      As String       ' システムシンボル
Public P_prgName                        As String       ' 実行Excelのプログラム名
Public P_version                        As String       ' vx.x.x
Public P_update                         As String       ' yyyymmdd
' プログラム実行時の日時情報
Public P_nowY                           As Integer      ' 今日の年（数字）
Public P_nowM                           As Integer      ' 今日の月（数字）
Public P_nowD                           As Integer      ' 今日の日（数字）
Public P_timeStart                      As Variant      ' プログラム開始の日付と時刻
Public P_timeStop                       As Variant      ' プログラム終了の日付と時刻
Public P_timeLap                        As Variant      ' プログラム実行の所要時間
Public P_nendoYYYY                      As Integer      ' 今年度（西暦）
' プログラム制御
Public P_mode                           As String       ' 操作モード insert / inquiry / modify / erase / clear / end
                                                        ' マクロ名のボタン番号指定方法　○○○.xlsm'!'処理名 n'　<== ボタン n
Public P_menuNum                        As Integer      ' シートボタンの処理番号
Public P_cnt                            As Long         ' 処理件数
Public P_cntErr                         As Long
' プログラム開始・終了メッセージ
Public P_openingMsg                     As String       ' プログラム開始メッセージ
Public P_closeingMsg                    As String       ' プログラム正常終了メッセージ
'
' --------------------------------------------------------------------------------
'   ※private変数(当該モジュール内のプロシージャ間で共有）
'     頭文字を大文字にする
' 個別定義

Public Sub メニュー処理(p_menu As Integer)
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
            Call 新住所録更新処理_R("")
        Case Else
            IsMsgPush ("プログラムのバグです。 中止します。")
            End
    End Select

End Sub

