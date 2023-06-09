Attribute VB_Name = "m1_初期化処理"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : 初期化処理（モジュール分割版）
' --------------------------------------+-----------------------------------------
' | @moduleName : m1_初期化処理
' | @Version    : v1.0.0
' | @update     : 2023/05/30
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
' |                1.2 外部のマスターのシートを取り込む…… M-�@新住所録原簿 / M-�AArchives
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
' |　M-�@新住所録原簿    �@原簿           Active住所録+InActive住所録
' |　M-�AArchives       �Aarchives       InActiveになってから３年以上の住所録/削除対象の住所録
' |　T-�B変更住所録      �B変更住所録    　追加・変更・削除になった住所録
' |       〃           原稿正規化       本システムのフォーマットに編集したレコード
' |       〃           原稿             本システムと異なるフォーマットの住所録
' |　M-�Hラベル一覧      �Hlabel          住所録をグループ化するためのリスト
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
' �@原簿シートの定義
Private wsOld                           As Worksheet
Private oldX, oldXmin, oldXmax          As Long             ' i≡x 列　column
Private oldY, oldYmin, oldYmax          As Long             ' j≡y 行　row
Private oldCnt                          As Long             ' 修正前レコードの件数
' �B変更住所録シートの定義
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

Public Sub 初期化処理_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     レコードの状態ごとにそれぞれのシートに振り分ける
' --------------------------------------+-----------------------------------------
    Dim x, y                            As Long
    Dim w_rate, w_mod                   As Integer      ' 進捗率 / 表示間隔
    Dim i, iMin, iMax                   As Long         ' 同一レコードの範囲(列 col x)
    Dim j, jMin, jMax                   As Long         ' 同一レコードの範囲(行 row y)
    Dim inExcelpath                     As String

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
    Set wsOld = Wb.Worksheets(Range("C_oldSheet").Value)        ' �@原簿シート
    Set wsTrn = Wb.Worksheets(Range("C_trnSheet").Value)        ' �B変更住所録シート
    Set WsNew = Wb.Worksheets(Range("C_newSheet").Value)        ' 新住所録シート
    Set WsWrk = Wb.Worksheets("work")                           ' 作業用シートのため固定
    
 ' 既存シートのクリア
    Call importClear_R(Range("C_oldSheet"))                     ' �@原簿シートのクリア
    Call importClear_R(Range("C_trnSheet"))                     ' �B変更住所録シートのクリア
    Call importClear_R(Range("C_newSheet"))                     ' 新住所録シートのクリア
    Call importClear_R("work")                                  ' 作業用シートのクリア

' カウントをゼロ
    oldCnt = 0
    trnCnt = 0
    newCnt = 0
    wrkCnt = 0

' 1.3 外部Excelから取り込む

  ' M-�@新住所録原簿を取り込み、戻り値を得る
    Call importSheet_R(Range("C_oldMst").Value, Range("C_oldSheet").Value, "M-�@新住所録原簿を選択してください。", _
                       inExcelpath, SrcYmax, SrcXmax)
    Range("C_oldMst").Value = inExcelpath
    oldYmin = YMIN
    oldXmin = XMIN

' T-�A変更住所録を取り込み、戻り値を得る
    Call importSheet_R(Range("C_trnMst").Value, Range("C_trnSheet").Value, "T-�B変更住所録を選択してください。", _
                       inExcelpath, TrnYmax, TrnXmax)
    Range("C_trnMst").Value = inExcelpath
    trnYmin = YMIN
    trnXmin = XMIN

' 1.4 取り込んだシートに(54)識別区分:BA列を付加し workシート　に統合し、(42)key姓名/(54)識別区分で昇順ソートする
'   (54)識別区分:BA列　�@原簿シート＝1、�B変更住所録＝3　
    j = 0
    jMin = oldYmin
    jMax = oldYmax
    wrkY = wrkYmin
    for j = jMin to jMax
        wsOld.range(cells(oldXmin,j),cells(oldMax,j)) copy 
        wsWrk.range(cells(wrkXmin,wrkY),cells(wrkMax,wrkY)).PasteSpecial _
                            Paste:=xlPasteValues _          ' 値の貼り付け
                          , Operation:=xlNone _             ' 演算して貼り付けは、しない
                          , SkipBlanks:=False _             ' 空白セルは、無視しない
                          , Transpose:=False                ' 行列を入れ替えない
        Application.CutCopyMode = False                     ' コピー状態の解除
        wrkY = wrkY + 1
        wsWrk.cells(54,wrkY) = 1                            '  (54)識別区分:BA列
    next j

    jMin = trnYmin
    jMax = trnYmax
    for j = jMin to jMax
        wsTrn.range(cells(trnXmin,j),cells(trnMax,j)) copy 
        wsWrk.range(cells(wrkXmin,wrkY),cells(wrkMax,wrkY)).PasteSpecial _
                            Paste:=xlPasteValues _          ' 値の貼り付け
                          , Operation:=xlNone _             ' 演算して貼り付けは、しない
                          , SkipBlanks:=False _             ' 空白セルは、無視しない
                          , Transpose:=False                ' 行列を入れ替えない
        Application.CutCopyMode = False                     ' コピー状態の解除
        wrkY = wrkY + 1
        wsWrk.cells(54,wrkY) = 3                            '  (54)識別区分:BA列
    next j
' オブジェクト変数の定義（共通）
    Set wb = ThisWorkbook
    ' 表の大きさを得る
    ' 作業シート（このシート）の初期値
    Set wsWrk = wb.Worksheets("work")
    wrkYmin = YMIN
    wrkXmin = XMIN
    wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）3列目（"C")名前列で計測
    wrkXmax = wsWrk.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）   ' ヘッダー行 3行目で計測
    wrkCnt = wrkYmax - wrkYmin + 1
' 昇順ソート　key: (39)姓名key(昇順)、(54)識別区分:BA列(降順）
    With ActiveSheet                '対象シートをアクティブにする
        .Sort.SortFields.Clear      '並び替え条件をクリア
        '項目1
        .Sort.SortFields.Add _
             Key:=.Range(PKEY_RNG) _    ' (39)姓名key(昇順)
            ,SortOn:=xlSortOnValues _
            ,Order:=xlAscending _
            ,DataOption:=xlSortNormal
        '項目2
        .Sort.SortFields.Add _
             Key:=.Cells(54, 3) _       ' (54)識別区分:BA列(降順）
            ,SortOn:=xlSortOnValues, _
            ,Order:=xlDescending, _
            ,DataOption:=xlSortNormal
'並び替えを実行する
        With .Sort
            .SetRange Range(Cells(YMIN - 1, XMIN), Cells(wrkYmax, XMAX))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With

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
    Range(Cells(YMIN, XMIN), Cells(yMax, XMAX)).Select
    Selection.ClearContents

End Sub

Private Sub importSheet_R(ByVal p_excelFile As String, ByVal p_objSheet As String, p_openFileMsg As String, _
                          ByRef p_srcFile As String, ByRef p_yMax As Long, ByRef p_Xmax As Long)
' --------------------------------------+-----------------------------------------
' | @function   : コピー元のシートをこのブックの同じ名前のシート へコピー
' | @moduleName : m1_初期化処理
' | @remarks
' | 引数の意味
' | 引　数：p_excelFile   コピー元のExcel File
' | 引　数：p_objSheet    コピーするシート名＝コピー先のシート名
' | 引　数：p_openFileMsg ファイル選択のエクスプローラに表示するメッセージ
' | 戻り値：p_srccFile    コピー元のExcelFileの絶対パスとファイル名
' | 戻り値：p_Ymax        コピーしたシートの最終行の位置
' | 戻り値：p_Xmax        コピーしたシートの最終列の位置
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
        If IsExitsFolderDir(childPath) Then
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
    wbTmp.Sheets(p_objSheet).Range(Cells(YMIN, XMIN), Cells(yMax, XMAX)).Copy
    Wb.Worksheets(p_objSheet).Range(Cells(YMIN, XMIN), Cells(yMax, XMAX)).PasteSpecial _
                            Paste:=xlPasteValues _          ' 値の貼り付け
                          , Operation:=xlNone _             ' 演算して貼り付けは、しない
                          , SkipBlanks:=False _             ' 空白セルは、無視しない
                          , Transpose:=False                ' 行列を入れ替えない
    Application.CutCopyMode = False                         ' コピー状態の解除
    wbTmp.Close saveChanges:=False                          ' 保存しないでclose
' 表の大きさを得る
    p_srcFile = srcFile
    p_yMax = Wb.Worksheets(p_objSheet).Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row            ' 最終行（縦方向）1列目（"A")で計測
    p_Xmax = Wb.Worksheets(p_objSheet).Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column   ' 最終列（横方向）   ' ヘッダー行 3行目で計測
' オブジェクト変数の解放
    Set wbTmp = Nothing

End Sub

