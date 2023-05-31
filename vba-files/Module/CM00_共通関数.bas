Attribute VB_Name = "CM00_共通関数"
Option Explicit
' --------------------------------------------------------------------------------
' | @function   : 標準的に利用されるモジュールを標準関数として統合
' --------------------------------------+-----------------------------------------
' | @moduleName : CM00_共通関数"
' | @Version    : v3.0.2
' | @update     : 2023/05/31
' | @written    : 2023/04/21
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |
' | @Program naming rule
' |  ProgramID: xx.99.99-xxxxxxxxx-vx.y.z-yyyymmdd.suffix   x.y.z≡version number
' |              |--xx(SystemSymbl)
' |              |----99(Subsystem#)
' |              |------.(priod≡separator)
' |              |-------99(Module#)
' |
' | @統合VBA内訳
' |      1.proc_H2_前処理
' |      2.proc_H9_後処理
' |      3.util02_標準モジュール一括Export
' |      4.util03_バックアップ書込
' |      5.util08_IsExitsSheet
' |      6.util09_IsMsgbox
' |      7.util10_IsMsgPush
' |      8.util15_IsExitsFolderDir
' |      9.util16_IsExitsFileDir
' |     10.util17_selectFile
' |     11.util18_selectFolder
' |     12.util19_getLatestFile
' |     13.util20_getFolderPath_F
' |     14.util23_ファイル名を入力
' |     15.util24_R1C1形式2A1形式変換_F
' |     16.util29_ブックプロパティ取得
' |     17.util30_getImportSheet
' |     18.util31_get共通変数_F
' |     19.util00_ステータスバー表示・消去
' |
' | @環境設定
' |　　① ファイルタブ→オプション→セキュリティセンター→セキュリティセンターの
' |　     設定ボタン→マクロの設定→
' |       VBAプロジェクトオブジェクトモデルへのアクセスを信頼する(V)」をチェック
' |    ② VBA画面のツールメニュー→参照設定→ 次のライブラリファイルをチェック
' |
' |     □Visual Basic For Applications
' |     □Microsoft Excel 15.0 Object Library
' |     □OLE Automatiom
' |     □Microsoft Scripting Runtime
' |     □Microsoft Scripting Libary
' |     □Microsoft Visual Basic for Application Extensibilly 5.3
' |     □Windows Script Host Object Model
' |
' | @ディレクトリ構造図 ≡ フォルダ構造図
' |     1       1       1       1       1
' |  root/      RootPath (システムフォルダの親フォルダ　ルートからシステムフォルダの前までのフルパス）
' |     │
' |     ├ システム名　 SysPath（フルパス）　SysName（フォルダ名）
' |     │      │
' |     │      ├ サブシステム名   SubSysPath（フルパス）　SubSysName（フォルダ名） （親ディレクトリ≡親フォルダ　../　ParentFolder)
' |     │      │      │
' |            │      ├ !Repository(開発環境)
' |                   │
' |   コピペ用記号集   ├ 1.運用の手引き
' |-----------------  │
' |     ┣ ┠ ┝ ├       ├ 2.マスター群
' |     │ ┃           │
' |     ┫ ┨ ┥ ┤       ├ 3.実行プログラム  folderPath（フルパス） folderName（フォルダ名） (カレントディレクトリ≡基点フォルダ　　./　)
' |     ┌ ┏ ┓ ┐       │      │
' |     └ ┗ ┛ ┘       │      ├ tmXX.YY.ZZ-○○○-vX.Y.Z-yyyymmdd.xlms
' |                   │      ├ tmXX.YY.ZZ-△△△-vX.Y.Z-yyyymmdd.xlms
' |                   │
' |                   ├ 7.管理ツール
' |                   │
' |                   ├ 8.コンテンツ
' |                   │
' |                   ├ 9.ドキュメント
' |                   │
' --------------------------------------+----------------------------------------
' |  命名規則の統一
' |     Public変数  先頭を大文字    ≡ pascalCase
' |     private変数 先頭を小文字    ≡ camelCase
' |     定数        全て大文字、区切り文字は、アンダースコア(_) ≡ snake_case
' |     引数        接頭語(p_)をつけ、camelCaseに準ずる
' --------------------------------------+-----------------------------------------
'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +

'
'   ※public変数(当該プロジェクト内のモジュール間で共有)は、最初に呼ばれるプロシジャーに定義
'
Public BackupFile                       As String       ' 実行前ファイルの保存用フォルダのフルパス
Public fullPath                         As String       ' 実行Excelのフルパス+ファイル名 ≡ Thisworkbook
Public PathName                         As String       ' 実行Excelのフルパス
Public FileName                         As String       ' 実行Excelのファイル名
' ディレクトリ構造のパスと名前
Public RootPath                         As String       ' システムフォルダの親ディレクトリのルートパス
Public SysPath                          As String       ' システムフォルダまでのフルパス
Public SysName                          As String       ' システムフォルダの名前
Public SubSysPath                       As String       ' サブシステムフォルダまでのフルパス
Public SubSysName                       As String       ' サブシステムフォルダの名前
' 実行プログラムの情報
Public SysSymbol                        As String       ' システムシンボル
Public PrgName                          As String       ' 実行Excelのプログラム名
Public Version                          As String       ' vx.x.x
Public Update                           As String       ' yyyymmdd
' プログラム実行時の日時情報
Public nowY                             As Integer      ' 今日の年（数字）
Public nowM                             As Integer      ' 今日の月（数字）
Public nowD                             As Integer      ' 今日の日（数字）
Public TimeStart                        As Variant      ' プログラム開始の日付と時刻
Public TimeStop                         As Variant      ' プログラム終了の日付と時刻
Public TimeLap                          As Variant      ' プログラム実行の所要時間
Public NendoYYYY                        As Integer      ' 今年度（西暦）
' プログラム制御
Public Mode                             As String       ' 操作モード insert / inquiry / modify / erase / clear / end
                                                        ' マクロ名のボタン番号指定方法　○○○.xlsm'!'処理名 n'　<== ボタン n
Public MenuNum                          As Integer      ' シートボタンの処理番号
Public NumCnt                           As Long         ' 処理件数
Public ErrCnt                           As Long         ' エラー件数
' プログラム開始・終了メッセージ
Public OpeningMsg                       As String       ' プログラム開始メッセージ
Public CloseingMsg                      As String       ' プログラム正常終了メッセージ
Public StatusBarMsg                     As String       ' Excelステータスバー（最下辺）に表示するメッセージ
'
'

Sub get共通変数_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' | @function   : ジョブの初期処理（標準版）
' --------------------------------------+-----------------------------------------
' | @moduleName : util31_get共通変数
' | @Version    : v1.0.0
' | @update     : 2023/05/04
' | @written    : 2023/05/04
' | @remarks
' |     ユーザ定義のPublic変数の初期値設定
' |
' --------------------------------------+-----------------------------------------
    Dim nowYMD                          As Date
    Dim rc                              As Long
    Dim temp, temp1, temp2, temp3, temp4 As Variant
    Dim x                               As Long
'
' ---Procedure Division ----------------+-----------------------------------------
'
   On Error Resume Next  ' エラーでも次の行から処理を続行する
       
    Calculate                                           ' [DockInfo]　最新状態に更新　パスを正しくするため

    Application.DisplayAlerts = False                   ' waringを止める
    Application.ScreenUpdating = True                   ' 処理中の画面を消す / 入力内容がシートで確認するには、リアルで更新するため　false
    Application.Calculation = xlCalculationManual       ' 手動計算に変更
    
    Mode = ""
    
' ここから実行時間のカウントを開始
    TimeStart = Time                                  ' 処理時間計測
    TimeStop = TimeStart
    TimeLap = TimeStop - TimeStart
    nowYMD = Now()               ' 今日の日付から年、月、日、月末日を分割
    nowY = Year(nowYMD)
    nowM = month(nowYMD)
    nowD = Day(nowYMD)

' ファイル名書式： .\.\SysID.xx.xx_programName-vX.Y.Z_yyyymmdd.sufix
    fullPath = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
    PathName = ActiveWorkbook.Path
    FileName = ActiveWorkbook.Name
    temp = Split(PathName, "\")
    RootPath = temp(0)                ' システムフォルダの親ディレクトリ
    For x = 1 To UBound(temp) - 3
        RootPath = RootPath & "\" & temp(x)
    Next x

' サブシステムフォルダ（実行モジュールの親ディレクトリ
    temp = Split(PathName, "\")
    SubSysPath = temp(0)                ' システムフォルダの親ディレクトリ
    For x = 1 To UBound(temp) - 1
        SubSysPath = SubSysPath & "\" & temp(x)
    Next x

' SystemSymbolの抽出              tmXX.YY.ZZ-ProgramName-vn.m.l-yyyymmdd
    temp = Split(FileName, "-")
    temp1 = Split(temp(0), ".")
    SysSymbol = temp1(0)
'    SysSymbol = Left(SysSymbol, 4)
' プログラム名、version、更新日の抽出
    PrgName = temp(1)
    Version = temp(2)
    Update = temp(3)

End Sub

Sub 前処理_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' | @function   : ジョブの初期処理（標準版）
' --------------------------------------+-----------------------------------------
' | @moduleName : proc_H2_前処理
' | @Version    : v3.2.0
' | @update     : 2023/05/17
' | @written    : 2020/12/29
' | @remarks
' |     1.プログラムで共通の初期値、環境変数を設定等の事前準備
' |         (1) 現在のパスを取得
' |         (2) 当該プログラムのバックアップを書き出す
' |         (3) ステータスバーの表示
' |
' --------------------------------------+-----------------------------------------
'
' ---Procedure Division ----------------+-----------------------------------------
'
   On Error Resume Next  ' エラーでも次の行から処理を続行する
   
    Call get共通変数_R("")
    
    IsMsgbox (OpeningMsg)
    
    Call putStatusBar(StatusBarMsg)
    
    Call バックアップ書込("")

End Sub


Sub 後処理_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' | @function   :終了処理の要約（標準版）
' --------------------------------------+-----------------------------------------
' | @moduleName : proc_H9_後処理
' | @Version    : v3.0.0
' | @update     : 2201/01/02
' | @written    : 2020/12/29
' | @remarks
' |     終了処理
' --------------------------------------+-----------------------------------------
    Dim msgText                             As String
'
' ---Procedure Division ----------------+-----------------------------------------
'

    Application.ScreenUpdating = True '消した顔面を表示する
    TimeStop = Time
    TimeLap = TimeStop - TimeStart

    If CloseingMsg <> "" Then
        msgText = CloseingMsg & Chr(13) & Chr(13) _
            & "-------------------------------------------------------------" & Chr(13) _
            & "       ［Job summary］" & Chr(13) _
            & "-------------------------------------------------------------" & Chr(13) _
            & "|Backup(Before)" & Chr(9) & "⇒ " & BackupFile & Chr(13) _
            & "|処理時間" & Chr(9) & "⇒ " & Minute(TimeLap) & "分" & Second(TimeLap) & "秒" & Chr(13) _
            & "|処理件数" & Chr(9) & "⇒ " & NumCnt & " 件"
      

        IsMsgPush (msgText)
    End If
    
    Calculate                   ' 画面を最新状態に更新
    Application.StatusBar = False
    Application.DisplayAlerts = True
    ActiveWorkbook.Save

End Sub

Public Sub 標準モジュール一括Export()
' --------------------------------------+-----------------------------------------
' | @function   : 標準モジュール等を一括してExportするマクロ
' --------------------------------------+-----------------------------------------
' | @moduleName : util02_標準モジュール一括Export
' | @Version    : v3.0.0  α、β版に対応 　　区切り文字：- に統一
' | @update     : 2020/04/25
' | @written    : 2019/08/01
' | @remarks
' |      https://vbabeginner.net/標準モジュール等の一括エクスポート/
' --------------------------------------+-----------------------------------------
   
    Dim fso                             As Object
    Dim prgFullPath                     As String           '// 絶対パスのファイル名 ≡ 絶対パス+ファイル名
    Dim prgPathName                     As String           '// 絶対パス
    Dim prgFileName                     As String           '// ファイル名 fileName Form: n.m.l xxxxxxxxx-v?_yyyymmdd.xlsx   ?≡version number
    Dim prgDocNumber                    As String
    Dim prgDocName                      As String
    Dim prgVersion                      As String
    Dim prgUpdate                       As String

    Dim myDir                           As String
    Dim myBook                          As String
    Dim mySheet                         As String
    Dim i                               As Long
    Dim iMax                            As Long
    Dim x1, x2, x3, x4, x5              As Long
    Dim xα, xβ, xv                    As Long
   
   
    Dim nowYMD                          As Date
    Dim nowY                            As Integer
    Dim nowM                            As Integer
    Dim nowD                            As Integer
'                                       +
    Dim module                          As VBComponent      '// モジュール
    Dim moduleList                      As VBComponents     '// VBAプロジェクトの全モジュール
    Dim extension                       As String           '// モジュールの拡張子
    Dim sPath                           As String           '// 処理対象ブックのパス
    Dim sFilePath                       As String           '// エクスポートファイルパス
    Dim TargetBook                      As Object           '// 処理対象ブックオブジェクト
    Dim saveDir                         As String           '// VBAの保存フォルダ

    Dim fullPath                        As String
    Dim sysSybl                         As String           '// システムシンボル　\!Program(????)
    Dim l, lMax                         As Long
'
' ---Procedure Division ----------------+-----------------------------------------
'
  
    On Error Resume Next  ' エラーが発生しても、次の行から処理を続行する
' --------------------------------------+-----------------------------------------
' |     自分が置かれているフォルダ名（カレントディレクトリ名）、
' |     自分のファイル名・シート名を取得する
' --------------------------------------+-----------------------------------------

    Call get共通変数_R("")
'    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fso = New FileSystemObject          ' インスタンス化

    myDir = ActiveWorkbook.Path
    myBook = ActiveWorkbook.Name
    mySheet = ActiveSheet.Name

    prgPathName = myDir
    prgFullPath = myDir & "\" & myBook
    prgFileName = myBook
' ProgramID: xx99.99-xxxxxxxxx-vx.y.z-yyyymmdd.suffix   x.y.z≡version number を分割
'             |--xx(SystemSymbl)
'             |----99(Subsystem#)
'             |------.(priod≡separator)
'             |-------99(Module#)
'
    iMax = Len(prgFileName)
'
    x1 = InStr(1, prgFileName, ".")      ' SystemSymbol・・半角 . の位置
    x2 = InStr(x1, prgFileName, "-")     ' Module# ・・・・半角 スペース の位置

                                        ' docName ・・・・半角 -v の位置
                                        ' α版､β版､v版 (RC版)
    xα = InStr(x2, prgFileName, "-α")
    xβ = InStr(x2, prgFileName, "-β")
    xv = InStr(x2, prgFileName, "-v")

    If xα <> 0 Then
        x3 = xα
    ElseIf xβ <> 0 Then
        x3 = xβ
    ElseIf xv <> 0 Then
        x3 = xv
    Else
        x3 = 0
    End If

    x4 = InStr(x3 + 1, prgFileName, "-")   ' version ・・・・半角 - の位置
    x5 = InStrRev(prgFileName, ".")      ' update・・・・・半角 - の位置

    prgDocNumber = Left(prgFileName, x2 - 1)  ' ≡　Module#
    prgDocName = Mid(prgFileName, x2 + 1, x3 - x2 - 1)
    prgVersion = Mid(prgFileName, x3 + 2, x4 - x3 - 2)
    prgUpdate = Mid(prgFileName, x4 + 1, x5 - x4 - 1)

' --------------------------------------------------------------------------------
' |     今日の日付から年、月、日、月末日を分割
' --------------------------------------------------------------------------------

    nowYMD = Now()               ' 今日の日付から年、月、日、月末日を分割

    nowY = Year(nowYMD)
    nowM = month(nowYMD)
    nowD = Day(nowYMD)
   
'   Call R_DocInfoGet           ' ファイルのパス情報（DocInfo)からProgramID等を得る

   '// ブックが開かれていない場合は個人用マクロブック（personal.xlsb）を対象とする
    If (Workbooks.Count = 1) Then
        Set TargetBook = ThisWorkbook
   '// ブックが開かれている場合は表示しているブックを対象とする
    Else
        Set TargetBook = ActiveWorkbook
    End If

    sPath = TargetBook.Path & "\!VBAmodules(" & prgDocNumber & ")-v" & prgVersion

    If Dir(sPath, vbDirectory) = "" Then  ' フォルダがないときは、作成する
        MkDir sPath
    End If


   '// 処理対象ブックのモジュール一覧を取得
    Set moduleList = TargetBook.VBProject.VBComponents

   '// VBAプロジェクトに含まれる全てのモジュールをループ
    For Each module In moduleList
       '// クラス
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
       '// フォーム
        ElseIf (module.Type = vbext_ct_MSForm) Then
           '// .frxも一緒にエクスポートされる
            extension = "frm"
       '// 標準モジュール
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
       '// その他
        Else
           '// エクスポート対象外のため次ループへ
            GoTo CONTINUE
        End If

       '// エクスポート実施
        sFilePath = sPath & "\" & module.Name & "-v" & prgVersion & "-" & prgUpdate & "." & extension
        Call module.Export(sFilePath)

       '// 出力先確認用ログ出力
       'Debug.Print sFilePath
CONTINUE:
    Next

    MsgBox "VBA標準モジュール" & "-v" & prgVersion & "-" & prgUpdate & "　群を一括Export完了"
End Sub

Private Sub R_DocInfoGet()
' --------------------------------------+-----------------------------------------
' |     自分が置かれているフォルダ名（カレントディレクトリ名）、
' |     自分のファイル名・シート名を取得する
' --------------------------------------+-----------------------------------------
'                                       +
    Dim myDir                           As String
    Dim myBook                          As String
    Dim mySheet                         As String
    Dim i                               As Long
    Dim iMax                            As Long
    Dim x1, x2, x3, x4, x5              As Long
    Dim xα, xβ, xv                    As Long

' ---Procedure Division ---------------+------------------------------------------
'
   On Error Resume Next  ' エラーが発生しても、次の行から処理を続行する
' --------------------------------------+-----------------------------------------
' |     自分が置かれているフォルダ名（カレントディレクトリ名）、
' |     自分のファイル名・シート名を取得する
' --------------------------------------+-----------------------------------------

'    Set fso = CreateObject("Scripting.FileSystemObject")
   Set fso = New FileSystemObject          ' インスタンス化

   myDir = ActiveWorkbook.Path
   myBook = ActiveWorkbook.Name
   mySheet = ActiveSheet.Name

   prgPathName = myDir
   prgFullPath = myDir & "\" & myBook
   prgFileName = myBook
' ProgramID: xx99.99-xxxxxxxxx-vx.y.z-yyyymmdd.suffix   x.y.z≡version number を分割
'             |--xx(SystemSymbl)
'             |----99(Subsystem#)
'             |------.(priod≡separator)
'             |-------99(Module#)
'
   iMax = Len(prgFileName)
'
   x1 = InStr(1, prgFileName, ".")      ' SystemSymbol・・半角 . の位置
   x2 = InStr(x1, prgFileName, "-")     ' Module# ・・・・半角 スペース の位置

                                        ' docName ・・・・半角 -v の位置
                                        ' α版､β版､v版 (RC版)
   xα = InStr(x2, prgFileName, "-α")
   xβ = InStr(x2, prgFileName, "-β")
   xv = InStr(x2, prgFileName, "-v")

   If xα <> 0 Then
       x3 = xα
   ElseIf xβ <> 0 Then
       x3 = xβ
   ElseIf xv <> 0 Then
       x3 = xv
   Else
       x3 = 0
   End If

   x4 = InStr(x3 + 1, prgFileName, "-")   ' version ・・・・半角 - の位置
   x5 = InStrRev(prgFileName, ".")      ' update・・・・・半角 - の位置

   prgDocNumber = Left(prgFileName, x2 - 1)  ' ≡　Module#
   prgDocName = Mid(prgFileName, x2 + 1, x3 - x2 - 1)
   prgVersion = Mid(prgFileName, x3 + 2, x4 - x3 - 2)
   prgUpdate = Mid(prgFileName, x4 + 1, x5 - x4 - 1)

' --------------------------------------------------------------------------------
' |     今日の日付から年、月、日、月末日を分割
' --------------------------------------------------------------------------------

   nowYMD = Now()               ' 今日の日付から年、月、日、月末日を分割

   nowY = Year(nowYMD)
   nowM = month(nowYMD)
   nowD = Day(nowYMD)

End Sub

Public Sub バックアップ書込(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' | @function   : 当該Excelの修正前バックアップを作る
' --------------------------------------+-----------------------------------------
' | @moduleName : util03_バックアップ書込
' | @Version    : v2.1.0
' | @update     : 2020/04/11
' | @written    : 2019/08/01
' | @remarks
' |     当該Excelの修正前バックアップを作る
' --------------------------------------+-----------------------------------------
    Dim saveDir                         As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set fso = New FileSystemObject          ' インスタンス化

    saveDir = PathName & "\" & SysSymbol & "-backup"

    If Dir(saveDir, vbDirectory) = "" Then  ' フォルダがないときは、作成する
        MkDir saveDir
    End If

    BackupFile = "backup-" & Format(Now(), "yyyy-mm-dd_hhmmss") & "_" & FileName
'バックアップ後も同じファイルを使うためには、SaveCopyAs を使う
    ActiveWorkbook.SaveCopyAs saveDir & "\" & BackupFile            ' バックアップ後も同じファイルを使うためには、　SaveCopyAs を使う

End Sub

Public Function IsExistSheet(p_sheetName) As Boolean
' --------------------------------------+-----------------------------------------
' | @function   : シートの存在をチェック（標準版）
' --------------------------------------+-----------------------------------------
' | @moduleName : util08_IsExitsSheet
' | @Version    : v1.0.0
' | @update     : 2020/04/25
' | @written    : 2020/04/25
' | @remarks
' |     メイン処理
' |　引数で渡されたシートの存在をチェックする
' |　存在する　≡ thru
' |　存在しない≡ false
' |
' --------------------------------------+-----------------------------------------
    Dim objWorksheet                    As Worksheet
'
' ---Procedure Division ----------------+-----------------------------------------
'

'
     On Error GoTo NotExists
    
    Set objWorksheet = ThisWorkbook.Sheets(p_sheetName)
    
    IsExistSheet = True
    
    Exit Function
    
NotExists:
    IsExistSheet = False
End Function

Public Function IsMsgbox(p_helloMsg) As Boolean
' --------------------------------------+-----------------------------------------
' | @function   : メッセージ表示の応答により対応付き　msgbox （標準版）
' --------------------------------------+-----------------------------------------
' | @moduleName : util09_IsMsgbox
' | @Version    : v2.0.0
' | @update     : 2020/12/25
' | @written    : 2020/04/06
' | @remarks
' |             IsMsgbox("メッセージ")
' |
' |　引数のメッセージを表示し、応答により処理を分ける
' |　yes　≡ thru
' |　no   ≡ false
' |
' --------------------------------------+-----------------------------------------
    Dim rc                              As VbMsgBoxResult       ' 列挙体
'
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'      標準ボタンの指定（ )
'  ---------------------------------------------------
'     定数　　　　　  値  　Enter キーで実行するキー
'  ------------------ ---- ---------------------------
'   vbDefaultButton1    0   第1ボタン
'   vbDefaultButton2   256  第2ボタン
'   vbDefaultButton3   512  第3ボタン
'   vbDefaultButton4   768  第4ボタン
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'       標準アイコンの指定
'  --------------- ------------  ------------- ---- -------------------------------
'      種類            画像         定数        値    説明
'  --------------- ------------  ------------- ---- -------------------------------
'   エラーアイコン vba_msgbox23  vbCritical     16  エラーメッセージで使用します。
'   疑問符アイコン vba_msgbox26  vbQuestion     32  ヘルプウィンドウを開くメッセージボックスで使用します。
'   警告アイコン   vba_msgbox24  vbExclamation  48  警告のメッセージで使用します。
'   情報アイコン   vba_msgbox25  vbInformation  64  大事な情報を伝えるメッセージで使用します。

'
' ---Procedure Division ----------------+-----------------------------------------
'
'
        If p_helloMsg = "" Then
            p_helloMsg = "処理を開始します。(by IsMsgbox)" & Chr(13) & "キャンセルで強制終了します。"
        End If
        rc = MsgBox(p_helloMsg & vbNewLine, vbOKCancel + vbQuestion + vbDefaultButton1)
        If rc = vbCancel Then
            MsgBox "処理を強制終了します。"
            IsMsgbox = False
            End
        Else
            IsMsgbox = True
        End If

    
End Function

Public Sub IsMsgPush(ByVal Msg As String)
' --------------------------------------+-----------------------------------------
' | @function   : 自動で閉じるメッセージするマクロ
' --------------------------------------+-----------------------------------------
' | @moduleName : util10_IsMsgPush
' | @Version    : v1.1.0
' | @update     : 2020/04/26
' | @written    : 2019/09/01
' | @remarks
' |     自動で閉じるメッセージ　開発 → VisualBasic → ツール → 参照設定 → Windows Script Host Object Modelをon(レ）
' --------------------------------------+-----------------------------------------
    Dim wsh                             As Object  'IWshRuntimeLibrary.WshShell
'
' ---Procedure Division ----------------+-----------------------------------------
'
    
    Set wsh = CreateObject("Wscript.Shell")

    wsh.Popup _
        Text:=Msg & vbNewLine & "～このメッセージは、１０秒後に自動的に消えます～", _
        SecondsToWait:=10, _
        Title:="", _
        Type:=vbOKOnly + vbInformation

    Set wsh = Nothing

End Sub

Public Function IsExitsFolderDir(p_sFolderPath) As Boolean
' --------------------------------------+-----------------------------------------
' | @function   : フォルダの存在をチェック（標準版）
' --------------------------------------+-----------------------------------------
' | @moduleName : util15_IsExitsFolderDir
' | @Version    : v1.2.0
' | @update     : 2023/05/31
' | @written    : 2020/12/09
' | @remarks
' |　引数で渡されたフォルダの存在をチェックする
' |　存在する　≡ thru
' |　存在しない≡ false
' |
' --------------------------------------+-----------------------------------------
    Dim result
'
' ---Procedure Division ----------------+-----------------------------------------
'
    result = Dir(p_sFolderPath, vbDirectory)
    If (result = "") Then
        IsExitsFolderDir = False    ' フォルダが存在しない
    Else
        IsExitsFolderDir = True     ' フォルダが存在する
    End If
    
End Function

Public Function IsExistFileDir(p_sFilePath) As Boolean
' --------------------------------------+-----------------------------------------
' | @function   :ファイルの存在をチェック（標準版）
' --------------------------------------+-----------------------------------------
' | @moduleName : util16_IsExitsFileDir
' | @Version    : v1.0.1
' | @update     : 2023/05/18
' | @written    : 2020/03/21
' | @remarks
' |　引数で渡されたファイルの存在をチェックする
' |　存在する　≡ thru
' |　存在しない≡ false
' |
' --------------------------------------+-----------------------------------------
    Dim result
'
' ---Procedure Division ----------------+-----------------------------------------
'
'
    result = Dir(p_sFilePath)
    If (result = "") Then
        IsExistFileDir = False
    Else
        IsExistFileDir = True
    End If
    
End Function

Public Function selectFile(ByVal P_title As String)
' --------------------------------------+-----------------------------------------
' | @function   : [ファイルを開く]ダイアログボックスでファイルを選択します
' --------------------------------------+-----------------------------------------
' | @moduleName : util17_selectFile
' | @Version    : v1.0.0
' | @update     : 2020/01/02
' | @written    : 2021/01/02
' | @remarks
' |     フォルダをフルパスで取得し、フルパスと親フォルダ名を返す
' --------------------------------------+-----------------------------------------
    Dim tbl                             As Variant
    Dim openFolder                      As String

    Dim z, zMin, zMax                   As Long
    Dim fullPath                        As String
    Dim folderName                      As String
    Dim delimiterChar                   As String
'
' ---Procedure Division ----------------+-----------------------------------------
'

'　[ファイルを開く]ダイアログボックスで対象Excelを選択します

        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = P_title
            If .Show = True Then
                openFolder = .SelectedItems(1)
            End If
        End With
        
        tbl = ""
        tbl = Split(openFolder, "\")
        zMin = LBound(tbl, 1)           ' 開始行
        zMax = UBound(tbl, 1)           ' 最終行
        delimiterChar = ""
        For z = zMin To zMax
            If tbl(z) <> "" Then
                If fullPath <> "" Then  ' パスの先頭は、ドライブ文字なので、区切り文字￥はつけない
                    delimiterChar = "\"
                End If
                fullPath = fullPath + delimiterChar + tbl(z)
            End If
        Next z
        
        folderName = tbl(zMax)
    
' 成功
    getFolderPath_F = Array(fullPath, folderName)

End Function

Public Function selectFolder(ByVal P_title As String) As Variant
' --------------------------------------+-----------------------------------------
' | @function   : [ファイルを開く]ダイアログボックスでフォルダを選択します
' --------------------------------------+-----------------------------------------
' | @moduleName : util18_selectFolder
' | @Version    : v1.0.0
' | @update     : 2020/01/02
' | @written    : 2021/01/02
' | @remarks
' |     [ファイルを開く]ダイアログボックスでフォルダを選択し,
' |     親フォルダのフルパスと、そのフォルダ名を返す
' --------------------------------------+-----------------------------------------
    Dim tbl                             As Variant
    Dim openFolder                      As String

    Dim z, zMin, zMax                   As Long
    Dim fullPath                        As String
    Dim folderName                      As String
    Dim delimiterChar                   As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
'　[ファイルを開く]ダイアログボックスで対象Excelを選択します

        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = P_title
            If .Show = True Then
                openFolder = .SelectedItems(1)
            End If
        End With
        
        tbl = ""
        tbl = Split(openFolder, "\")
        zMin = LBound(tbl, 1)           ' 開始行
        zMax = UBound(tbl, 1)           ' 最終行
        delimiterChar = ""
        For z = zMin To zMax
            If tbl(z) <> "" Then
                If fullPath <> "" Then  ' パスの先頭は、ドライブ文字なので、区切り文字￥はつけない
                    delimiterChar = "\"
                End If
                fullPath = fullPath + delimiterChar + tbl(z)
            End If
        Next z
        
        folderName = tbl(zMax)
    
' 成功
    selectFolder = Array(fullPath, folderName)

End Function

Public Function getLatestFile(ByVal FolderPath As String) As String
' --------------------------------------+-----------------------------------------
' | @function   : フォルダにある最新のファイルを調べる
' --------------------------------------+-----------------------------------------
' | @moduleName : util19_getLatestFile
' | @Version    : v1.0.0
' | @update     : 2020/01/02
' | @written    : 2021/01/02
' | @remarks
' |     フォルダにある最新のファイルを調べる
' --------------------------------------+-----------------------------------------
    Dim Buf                             As String
    Dim fList(99)                       As String
    Dim i, iMin, iMax                   As Long
    Dim latestFile                      As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
     
    iMin = LBound(fList, 1)           ' 開始行
    iMax = UBound(fList, 1)           ' 最終行
    i = 0
    Buf = Dir(FolderPath & "\*.xlsm")
    Do While Buf <> ""
        i = i + 1
        fList(i) = Buf
        Buf = Dir()
    Loop
    bubbleSort (fList)
    latestFile = fList(iMin)
    For i = iMin + 1 To iMax
        If fList(i) <> "" Then
            If latestFile < fList(i) Then
                latestFile = fList(i)
            End If
        End If
    Next i

' 成功
    getLatestFile = latestFile

End Function

Public Function getFolderPath_F(ByVal P_title As String)
' --------------------------------------+-----------------------------------------
' | @function   : [ファイルを開く]ダイアログボックスでファイルを選択します
' --------------------------------------+-----------------------------------------
' | @moduleName : util20_getFolderPath_F
' | @Version    : v1.0.0
' | @update     : 2020/01/02
' | @written    : 2021/01/02
' | @remarks
' |     フォルダをフルパスで取得し、フルパスと親フォルダ名を返す
' --------------------------------------+-----------------------------------------
    Dim tbl                             As Variant
    Dim openFolder                      As String

    Dim z, zMin, zMax                   As Long
    Dim fullPath                        As String
    Dim folderName                      As String
    Dim delimiterChar                   As String
'
' ---Procedure Division ----------------+-----------------------------------------
'

'　[ファイルを開く]ダイアログボックスで対象Excelを選択します

        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = P_title
            If .Show = True Then
                openFolder = .SelectedItems(1)
            End If
        End With
        
        tbl = ""
        tbl = Split(openFolder, "\")
        zMin = LBound(tbl, 1)           ' 開始行
        zMax = UBound(tbl, 1)           ' 最終行
        delimiterChar = ""
        For z = zMin To zMax
            If tbl(z) <> "" Then
                If fullPath <> "" Then  ' パスの先頭は、ドライブ文字なので、区切り文字￥はつけない
                    delimiterChar = "\"
                End If
                fullPath = fullPath + delimiterChar + tbl(z)
            End If
        Next z
        
        folderName = tbl(zMax)
    
' 成功
    getFolderPath_F = Array(fullPath, folderName)

End Function

Public Function ファイル名を入力_F(ByVal P_MSG As String) As String
' --------------------------------------+-----------------------------------------
' | @function   : ファイル名を得る
' --------------------------------------+-----------------------------------------
' | @moduleName : util23_ファイル名を入力
' | @Version    : v1.0.0
' | @update     : 2021/06/22
' | @written    : 2021/06/22
' | @remarks
' |     戻り値：fullPath
' |     目的のファイル名（フルパス）を得る
' |
' --------------------------------------+-----------------------------------------
    Dim OpenPathFileName                As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
'　[ファイルを開く]ダイアログボックスで対象Excelを選択します
    OpenPathFileName = Application.GetOpenFilename("Excelファイル,*.xl*", , P_MSG)
        If OpenPathFileName = "False" Then
            MsgBox "選択したEXCELがエラーです。強制終了します。"
        End                             ' 処理の強制終了
    End If
    
 
' 成功
    ファイル名を得る_F = OpenPathFileName

End Function

Public Function toA1_F(ByVal StrR1C1 As String) As String
' --------------------------------------+-----------------------------------------
' | @function   : セルのR1C1形式をA1形式に変換
' --------------------------------------+-----------------------------------------
' | @moduleName : util24_R1C1形式2A1形式変換_F
' | @Version    : v1.0.0
' | @update     : 2021/06/24
' | @written    : 2021/06/24
' | @remarks
' |     戻り値：String型
' --------------------------------------+-----------------------------------------
'
' ---Procedure Division ----------------+-----------------------------------------
'
    toA1_F = Application.ConvertFormula( _
        Formula:=StrR1C1, _
        fromReferenceStyle:=xlR1C1, _
        toreferencestyle:=xlA1, _
        toabsolute:=xlRelative)

End Function

Public Sub util29_ブックプロパティ取得(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' | @function   : Bookのプロファイルを取得
' --------------------------------------+-----------------------------------------
' | @moduleName : util29_ブックプロパティ取得
' | @Version    : v1.0.0
' | @update     : 2021/04/07
' | @written    : 2022/04/07
' | @remarks
' |  「DocInfo(削除不可)」シートを参照
' --------------------------------------+-----------------------------------------
'   ※public変数(当該プロジェクト内のモジュール間で共有)は、最初に呼ばれるプロシジャーに定義
'     接頭語に P_ をつける
'
'
' ---Procedure Division ----------------+-----------------------------------------
'
    MsgBox "Title: " & ThisWorkbook.BuiltinDocumentProperties("Title") & vbNewLine _
        & "Author: " & ThisWorkbook.BuiltinDocumentProperties("Author") & vbNewLine _
        & "Subject: " & ThisWorkbook.BuiltinDocumentProperties("Subject") & vbNewLine _
        & "Keywords: " & ThisWorkbook.BuiltinDocumentProperties("Keywords") & vbNewLine _
        & "Category: " & ThisWorkbook.BuiltinDocumentProperties("Category") & vbNewLine _
        & "Comments: " & ThisWorkbook.BuiltinDocumentProperties("Comments")
End Sub

Public Function getImportSheet_F(ByVal p_importFile As String, ByVal p_importSheet As String, ByVal p_saveSheet As String, ByVal p_childPath As String, ByVal p_importMsg As String)
' --------------------------------------+-----------------------------------------
' | @function   : 別Excelのシートをimportするマクロ
' |             　importExcelがないか、エラーのときはエックスプローラで選択
' --------------------------------------+-----------------------------------------
' | @moduleName : util30_getImportSheet
' | @Version    : v1.0.0
' | @update     : 2023/04/29
' | @written    : 2023/04/29
' | @remarks
' |
' --------------------------------------+-----------------------------------------
    Dim wb                              As Workbook
    Dim ws                              As Worksheet
    Dim wbImp                           As Workbook
    Dim sw_FalseTrue                    As Boolean
    Dim w_path                          As String
'
' ---Procedure Division ----------------+-----------------------------------------
'

' 1.importするシートを格納するシート［work］は、事前に削除
' --------------------------------------+-----------------------------------------
'
    Set wb = ActiveWorkbook
    If IsExistSheet(p_saveSheet) Then
        wb.Worksheets(p_saveSheet).Delete            ' 以前のシートを削除
    End If
' ファイル指定の有無チェック　/　指定したファイルがなかったら、エクスプローラで指定させる
    sw_FalseTrue = False
    If p_importFile <> "" Then
        If IsExistFileDir(p_importFile) Then
            sw_FalseTrue = True
        Else
            sw_FalseTrue = False
        End If
    End If
'　[ファイルを開く]ダイアログボックスで対象Excelを選択します
    If sw_FalseTrue = False Then
        If p_childPath = "" Then
            w_path = PathName
        Else
            w_path = SubSysPath & "\" & p_childPath
        End If
        
        ChDir w_path                                    ' プログラムのあるフォルダを指定
        p_importFile = Application.GetOpenFilename("Excelファイル,*.xl*", , p_importMsg)
        sw_FalseTrue = True
    End If
    
' 外部Excelファイルを開き、importシートを作業シート work へコピー
    Workbooks.Open p_importFile
    Set wbImp = ActiveWorkbook
    wbImp.Worksheets(p_importSheet).Copy after:=wb.Worksheets(1)
    
    Set ws = ActiveSheet
    ws.Name = p_saveSheet
    ws.Tab.Color = RGB(0, 112, 192)     ' 青
    wbImp.Close saveChanges:=False      ' 保存しないでclose

' オブジェクト変数の解放
    Set wbImp = Nothing
    Set wb = Nothing
    Set ws = Nothing
' 戻り値
    getImportSheet_F = p_importFile

End Function

Public Sub putStatusBar(ByVal p_statusBarMsg As String)
' --------------------------------------------------------------------------------
' | @function   : EXCELのステータスバーに進行状況を表示
' --------------------------------------+-----------------------------------------
' | @moduleName : util00_ステータスバー表示・消去
' | @Version    : v1.1.0
' | @update     : 2023/05/17
' | @written    : 2021/12/11
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |
' |
' --------------------------------------+-----------------------------------------
'
' ---Procedure Division ----------------+-----------------------------------------
'/
'      EXCELのステータスバーに進行状況を表示 / 消去

    If p_statusBarMsg = "" Then
        Application.StatusBar = False   '      EXCELのステータスバーを消去
    Else
        Application.StatusBar = Format(Now(), "m/d hh:mm") & "：" & p_statusBarMsg & " を処理中です。"
    End If
    Calculate                         ' 　最新状態に更新
End Sub

Public Sub debug2text(ByVal p_text As String, Optional P_mode As String = "print")
' --------------------------------------------------------------------------------
' | @function   : Debug.Printをテキストファイルに出力
' --------------------------------------+-----------------------------------------
' | @moduleName : util33_debug2File
' | @Version    : v1.0.0
' | @update     : 2023/05/22
' | @written    : 2023/05/22
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |     引数: p_test
' |     引数: p_mode
' |         open  ≡ テキストファイルの open 処理
' |         print ≡ テキストの出力(Default)
' |         close ≡ テキストファイルの　close 処理
' |
' |
' |
' --------------------------------------+-----------------------------------------
'
  Dim dt                                As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
    Select Case P_mode
        Case "open"
' 処理日時を取得
            dt = Format(Now, "yyyymmdd_hhmmss") ' 現在の日時をyyyymmdd_hhmmss形式で取得
' 使用可能なファイル番号を取得
            FileNum = FreeFile()
' 常に新規ファイルを作成、ファイル名の重複を防ぐため処理日時をファイル名に追記
            Open ThisWorkbook.Path & "\" & "loggPrint" & "_" & dt & ".txt" For Output As #FileNum
        
        Case "close"
' ログファイルCLOSE
            Close #FileNum
        
        Case "print"
' ログ出力
            Print #FileNum, Now & " " & p_text
            Debug.Print Now & " " & p_text
            
        Case Else
' ログ出力
            Print #FileNum, Now & " " & p_text
            Debug.Print Now & " " & p_text
    End Select
        
    
    

'' 処理日時を取得
'    dt = Format(Now, "yyyymmdd_hhmmss") ' 現在の日時をyyyymmdd_hhmmss形式で取得
'' ログファイルのパスを設定
'    filePath = SubSysPath & "\" & "DebugPrint.txt"
'' 使用可能なファイル番号を取得
'    fileNo = FreeFile()
'' 常に新規ファイルを作成、ファイル名の重複を防ぐため処理日時をファイル名に追記
'    Open filePath & "_" & dt & ".txt" For Output As #fileNo
'' ログ出力
'    Print #fileNo, Now & " " & p_text
'' ログファイルCLOSE
'    Close #fileNo
'
    
End Sub

' @ReadMe 標準フォーム（例）
' --------------------------------------+-----------------------------------------
' | @function   : ジョブの初期処理（標準版）
' --------------------------------------+-----------------------------------------
' | @moduleName : util31_get共通変数
' | @Version    : v1.0.0
' | @update     : 2023/05/04
' | @written    : 2023/05/04
' | @remarks
' |     ユーザ定義のPublic変数の初期値設定
' |
' |   ※public変数(当該プロジェクト内のモジュール間で共有)は、最初に呼ばれるプロシジャーに定義
' |     接頭語に P_ をつける
' --------------------------------------+-----------------------------------------
'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
'
' ---Procedure Division ----------------+-----------------------------------------
'

' closeingMsg = "|①原簿シート" & Chr(9) & "＝ " & SrcCnt & Chr(13) _
'             & "|②archives" & Chr(9) & "＝ " & ArvCnt & Chr(13) _
'             & "|③目視" & Chr(9) & "＝ " & EyeCnt & Chr(13) _
'             & "| エラー" & Chr(9) & "＝ " & ErrCnt & Chr(13)
  
