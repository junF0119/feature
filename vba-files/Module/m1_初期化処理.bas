Attribute VB_Name = "m1_初期化処理"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : 初期化処理（モジュール分割版）
' --------------------------------------+-----------------------------------------
' | @moduleName : m1_初期化処理
' | @Version    : v1.0.0
' | @update     : 2023/06/02
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
' |        プログラム構造
' |            1. 初期処理
' |                1.1 既存シートのクリア
' |                    importClear_R()
' |                1.2 外部のマスターのシートを取り込む…… M-①新住所録原簿 / M-②Archives
' |                    importSheet_R()
' |            2. データの整合性検証
' |                2.1  重複チェック…… (53)PrimaryKey / (42)key姓名
' |                    keyCheck_F()
' |                        arrSet_R()
' |                        duplicateChk_F()
' |                            quickSort_R()
' --------------------------------------+-----------------------------------------
' |[使用するブックとシート]
' |  ファイル名         シート名        意味
' |　M-①新住所録原簿    ①原簿           Active住所録+InActive住所録
' |　M-②Archives       ②archives       InActiveになってから３年以上の住所録/削除対象の住所録
' |　T-③変更住所録      ③変更住所録    　追加・変更・削除になった住所録
' |       〃           原稿正規化       本システムのフォーマットに編集したレコード
' |       〃           原稿             本システムと異なるフォーマットの住所録
' |　M-⑨ラベル一覧      ⑨label          住所録をグループ化するためのリスト
' |
' |
' --------------------------------------+-----------------------------------------
' |  命名規則の統一
' |     Public変数  先頭を大文字    ≡ pascalCase
' |     private変数 先頭を小文字    ≡ camelCase
' |     定数        全て大文字、区切り文字は、アンダースコア(_) ≡ snake_case
' |     引数        接頭語(p_)をつけ、camelCaseに準ずる
' --------------------------------------+-----------------------------------------
'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +

' オブジェクト変数の定義
Private Wb                              As Workbook         ' このブック
' ①原簿シートの定義
Private wsOld                           As Worksheet
Private oldX, oldXmin, oldXmax          As Long             ' i≡x 列　column
Private oldY, oldYmin, oldYmax          As Long             ' j≡y 行　row
Private oldCnt                          As Long             ' 修正前レコードの件数
' ②archivesシートの定義
Private wsArv                           As Worksheet
Private arvX, arvXmin, arvXmax          As Long             ' i≡x 列　column
Private arvY, arvYmin, arvYmax          As Long             ' j≡y 行　row
Private arvCnt                          As Long             ' 修正前レコードの件数
' ③変更住所録シートの定義
Private wsTrn                           As Worksheet
Private trnX, trnXmin, trnXmax          As Long             ' i≡x 列　column
Private trnY, trnYmin, trnYmax          As Long             ' j≡y 行　row
Private trnCnt                          As Long             ' 変更レコードの件数
' 新住所録シートの定義
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i≡x 列　column
Private newY, newYmin, newYmax          As Long             ' j≡y 行　row
Private newCnt                          As Long             ' 修正後レコードの件数
' 作業シート work の定義
Private wsWrk                           As Worksheet
Private wrkX, wrkXmin, wrkXmax          As Long             ' i≡x 列　column
Private wrkY, wrkYmin, wrkYmax          As Long             ' j≡y 行　row
Private wrkCnt                          As Long             ' 作業レコードの件数=修正前+変更
 
Public Sub m1_初期化処理_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     レコードの状態ごとにそれぞれのシートに振り分ける
' --------------------------------------+-----------------------------------------
    Dim x, y                            As Long
    Dim w_rate, w_mod                   As Integer      ' 進捗率 / 表示間隔
    Dim i, iMin, iMax                   As Long         ' 同一レコードの範囲(列 col x)
    Dim j, jMin, jMax                   As Long         ' 同一レコードの範囲(行 row y)
    Dim inExcelpath                     As String

' ' 構造体の宣言
' Type cntTbl
'     old                                 As long     ' ①原簿
'     arv                                 As long     ' ②archive
'     trn                                 As long     ' ③変更住所録
'     wrk                                 As long     ' work
'     new1                                As long     ' newの原簿レコード
'     new2                                As long     ' newのarchivwレコード
' End Type
    Dim cnt                             As cntTbl
'
' ---Procedure Division ----------------+-----------------------------------------
'
' 1.1 前処理（共通）
            
    OpeningMsg = "「新住所録原簿の更新処理」プログラムを開始します。"
    StatusBarMsg = OpeningMsg
    Call 前処理_R("")

' 1.2 初期設定処理
    
' オブジェクト変数の定義（共通）
    Set Wb = ThisWorkbook
    Set wsOld = Wb.Worksheets(Range("C_oldSheet").Value)        ' ①原簿シート
    Set wsArv = Wb.Worksheets(Range("C_arvSheet").Value)        ' ②archivesシート
    Set wsTrn = Wb.Worksheets(Range("C_trnSheet").Value)        ' ③変更住所録シート
    Set wsNew = Wb.Worksheets(Range("C_newSheet").Value)        ' 新住所録シート
    Set wsWrk = Wb.Worksheets("work")                           ' 作業用シートのため固定
    
 ' 既存シートのクリア
    Call importClear_R(Range("C_oldSheet"))                     ' ①原簿シートのクリア
    Call importClear_R(Range("C_arvSheet"))                     ' ②archivesシートのクリア
    Call importClear_R(Range("C_trnSheet"))                     ' ③変更住所録シートのクリア
    Call importClear_R(Range("C_newSheet"))                     ' 新住所録シートのクリア
    Call importClear_R("work")                                  ' 作業用シートのクリア

' カウントをゼロ
    oldCnt = 0
    arvCnt = 0
    trnCnt = 0
    newCnt = 0
    wrkCnt = 0

' 1.3 外部Excelから取り込む

' M-①新住所録原簿を取り込み、戻り値を得る
    Call importSheet_R(Range("C_oldMst").Value, Range("C_oldSheet").Value, "M-①新住所録原簿を選択してください。", _
                       inExcelpath, oldYmax, oldXmax)
    Range("C_oldMst").Value = inExcelpath
    oldYmin = YMIN
    oldXmin = XMIN

' M-②archivesを取り込み、戻り値を得る
    Call importSheet_R(Range("C_arvMst").Value, Range("C_arvSheet").Value, "M-②Archivesを選択してください。", _
                       inExcelpath, arvYmax, arvXmax)
    Range("C_arvMst").Value = inExcelpath
    arvYmin = YMIN
    arvXmin = XMIN

' T-③変更住所録を取り込み、戻り値を得る
    Call importSheet_R(Range("C_trnMst").Value, Range("C_trnSheet").Value, "T-③変更住所録を選択してください。", _
                       inExcelpath, trnYmax, trnXmax)
    Range("C_trnMst").Value = inExcelpath
    trnYmin = YMIN
    trnXmin = XMIN

' 1.4 取り込んだシートに(54)識別区分:BA列を付加し workシート　に統合し、(42)key姓名/(54)識別区分で昇順ソートする
'   (54)識別区分:BA列　①原簿シート＝1、②archivesシート＝2、③変更住所録シート＝3
    j = 0
    jMin = oldYmin
    jMax = oldYmax
    wrkY = YMIN
    wrkYmin = YMIN
    wrkXmin = XMIN
' ①原簿シート
    For j = jMin To jMax
        wsOld.Range(Cells(j, oldXmin).Address, Cells(j, oldXmax).Address).Copy
        wsWrk.Range(Cells(wrkY, wrkXmin).Address).PasteSpecial _
                                                  Paste:=xlPasteValues _
                                                , Operation:=xlNone _
                                                , SkipBlanks:=False _
                                                , Transpose:=False
                                                
        Application.CutCopyMode = False                     ' コピー状態の解除
        wsWrk.Cells(wrkY, 54) = 1                           '  (54)識別区分:BA列
        wrkY = wrkY + 1
        oldCnt = oldCnt + 1
    Next j
' ②archivesシート
    jMin = arvYmin
    jMax = arvYmax
    For j = jMin To jMax
        wsArv.Range(Cells(j, arvXmin).Address, Cells(j, arvXmax).Address).Copy
        wsWrk.Range(Cells(wrkY, wrkXmin).Address).PasteSpecial _
                                                  Paste:=xlPasteValues _
                                                , Operation:=xlNone _
                                                , SkipBlanks:=False _
                                                , Transpose:=False
                                                            
        Application.CutCopyMode = False                     ' コピー状態の解除
        wsWrk.Cells(wrkY, 54) = 2                           '  (54)識別区分:BA列
        wrkY = wrkY + 1
        arvCnt = arvCnt + 1
    Next j

' ③変更住所録シート
    jMin = trnYmin
    jMax = trnYmax
    For j = jMin To jMax
        wsTrn.Range(Cells(j, trnXmin).Address, Cells(j, trnXmax).Address).Copy
        wsWrk.Range(Cells(wrkY, wrkXmin).Address).PasteSpecial _
                                                  Paste:=xlPasteValues _
                                                , Operation:=xlNone _
                                                , SkipBlanks:=False _
                                                , Transpose:=False
                                                            
        Application.CutCopyMode = False                     ' コピー状態の解除
        wsWrk.Cells(wrkY, 54) = 3                           '  (54)識別区分:BA列
        wrkY = wrkY + 1
        trnCnt = trnCnt + 1
    Next j

' オブジェクト変数の定義（共通）
    Sheets("work").Activate
' 表の大きさを得る
' 作業シート（このシート）の初期値
    Set wsWrk = Wb.Worksheets("work")
    wrkYmin = YMIN
    wrkXmin = XMIN
    wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）6列目（"F")名前列で計測
    wrkXmax = wsWrk.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）   ' ヘッダー行 3行目で計測
    wrkCnt = wrkYmax - wrkYmin + 1

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
' シート別のレコード件数をPublic変数にセット
    cnt.old = oldCnt    ' ①原簿
    cnt.arv = arvCnt    ' ②archive
    cnt.trn = trnCnt    ' ③変更住所録
    cnt.wrk = wrkCnt    ' work
    cnt.new1 = newCnt   ' newの原簿レコード
    cnt.new2 = newCnt   ' newのarchivesレコード

End Sub

Private Sub importClear_R(ByVal p_sheetName As String)
' --------------------------------------+-----------------------------------------
' | ※既存のシートを削除し、importシートをコピーすると、importするシートの名前定義も
' | 一緒にコピーされ、名前定義の重複でロジックに不具合が生じるので、既存シートの削除
' | でなくシートのクリアとfieldのコピーで対応することに変更する。
' |
' --------------------------------------+-----------------------------------------
'
' ---Procedure Division ----------------+-----------------------------------------
'
    Sheets(p_sheetName).Activate
'  シートに関係なく、データ域を一律クリア（ヘッダー域は除く）
    Range(Cells(YMIN, XMIN), Cells(yMax, XMAX + 1)).Select
    Selection.ClearContents

End Sub

Private Sub importSheet_R(ByVal p_excelFile As String, ByVal p_objSheet As String, ByVal p_openFileMsg As String, _
                          ByRef p_srcFile As String, ByRef p_yMax As Long, ByRef p_xMax As Long)
' --------------------------------------+-----------------------------------------
' | @function   : コピー元のシートをこのブックの同じ名前のシート へコピー
' | @moduleName : m1_初期化処理
' | @remarks
' | 引数の意味
' | 引　数：p_excelFile   コピー元のExcel File
' | 引　数：p_objSheet    コピーするシート名＝コピー先のシート名
' | 引　数：p_openFileMsg ファイル選択のエクスプローラに表示するメッセージ
' | 戻り値：p_srccFile    コピー元のExcelFileの絶対パスとファイル名
' | 戻り値：p_yMax        コピーしたシートの最終行の位置
' | 戻り値：p_xMax        コピーしたシートの最終列の位置
' |
' --------------------------------------+-----------------------------------------
    Dim wbTmp                           As Workbook     ' コピーもとのExcelファイルシート
    Dim childPath                       As String       ' コピー元のExcelの絶対パス
    Dim srcFile                         As String       ' コピー元のExcelファイル名（絶対パス付き）
    Dim sw_naFile                       As Boolean      ' ファイル有 ≡ True ファイル無 ≡ False
    Dim sw_naFolder                     As Boolean      ' フォルダ有 ≡ True フォルダ無 ≡ False
    
    Dim i, y                            As Long
    Dim absolutePath                    As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
' --------------------------------------+-----------------------------------------
' 1.import Excel ファイルの読み込み
' --------------------------------------+-----------------------------------------

' フォルダ指定の有無チェック　/　指定したフォルダがなかったら、エクスプローラで指定させる
    childPath = Range("C_childPath")
    sw_naFolder = False
    sw_naFile = False
' パス指定があるときは、フォルダの存在をチェック
    If childPath <> "" Then
        If IsExitsFolderDir(SubSysPath & "\" & childPath) Then
            sw_naFolder = True
        Else
            childPath = ""
            Range("C_childPath") = ""
        End If
    End If
' ファイル指定の有無チェック　/　指定したファイルがなかったら、エクスプローラで指定させる
    If p_excelFile <> "" Then
        If IsExistFileDir(p_excelFile) Then
            sw_naFile = True
            srcFile = p_excelFile
        Else
            srcFile = ""
        End If
    End If
    
'　[ファイルを開く]ダイアログボックスで対象Excelを選択します
    If sw_naFile = False Then
        If childPath = "" Then
            absolutePath = PathName
        Else
            absolutePath = SubSysPath & "\" & childPath
        End If
        ChDir absolutePath                                    ' プログラムのあるフォルダを指定
        srcFile = Application.GetOpenFilename("Excelファイル,*.xl*", , p_openFileMsg)
        sw_naFile = True
    End If
    
' 外部Excelファイルを開き、importシートを作業シート work へコピー
    Workbooks.Open srcFile
    Set wbTmp = ActiveWorkbook
'    ActiveSheet.ShowAllData         ' フィルタ解除
    wbTmp.Sheets(p_objSheet).Range(Cells(YMIN, XMIN).Address, Cells(yMax, XMAX).Address).Copy
    Wb.Sheets(p_objSheet).Range(Cells(YMIN, XMIN).Address).PasteSpecial _
                                                           Paste:=xlPasteValues _
                                                         , Operation:=xlNone _
                                                         , SkipBlanks:=False _
                                                         , Transpose:=False
    
   
    
    Application.CutCopyMode = False                         ' コピー状態の解除
    wbTmp.Close saveChanges:=False                          ' 保存しないでclose
' 表の大きさを得る
    p_srcFile = srcFile
    p_yMax = Wb.Worksheets(p_objSheet).Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row            ' 最終行（縦方向）(6)名前（"F")で計測
    p_xMax = Wb.Worksheets(p_objSheet).Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column   ' 最終列（横方向）   ' ヘッダー行 3行目で計測
' オブジェクト変数の解放
    Set wbTmp = Nothing

End Sub

