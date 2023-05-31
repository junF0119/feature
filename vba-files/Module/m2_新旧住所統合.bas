Attribute VB_Name = "m2_新旧住所統合"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : 新旧の住所録をマージする
' --------------------------------------+-----------------------------------------
' | @moduleName : m2_新旧住所統合
' | @Version    : v1.0.0
' | @update     : 2023/05/10
' | @written    : 2023/05/10
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |  （１）単独レコードを新規シート（①new）へ移動する
' |  （２）重複レコードは、変更住所録を新規シート（①new）へ移動する
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
Const PKEY_RNG                          As String = "AM3"   ' Keyのセル番号
Const PKEY_X                            As Long = 39        ' Keyの列番号"AP"
Const PSEIMEI_X                         As Long = 3         ' 作業域の最大行数計測の列番号"C"(名前)
Const PDEL_X                            As Long = 38        ' 削除日の列番号"AL"
Const XMIN                              As Long = 1         ' 開始列
Const XMAX                              As Long = 42        ' 最終列
Const YMIN                              As Long = 4         ' 開始行　∵ヘッダー部を除く
Const yMax                              As Long = 1999      ' 最大行　∵このプログラムであつかう最大行
Const INPUTX_FROM                       As Long = 3         ' 入力項目開始列
Const INPUTX_TO                         As Long = 23        ' 入力項目終了列
Const CHECKED_X                         As Long = 40        ' チェック欄（自由）

Private wb                              As Workbook         ' このブック
' ①oldSheet シートの定義
Private wsOld                           As Worksheet
Private oldX, oldXmin, oldXmax          As Long             ' i≡x 列　column
Private oldY, oldYmin, oldYmax          As Long             ' j≡y 行　row
Private oldCnt                          As Long             ' 削除レコードの件数
' ②trnSheet シートの定義
Private wsTrn                           As Worksheet
Private trnX, trnXmin, trnXmax          As Long             ' i≡x 列　column
Private trnY, trnYmin, trnYmax          As Long             ' j≡y 行　row
Private trnCnt                          As Long             ' 目視レコードの件数
' ③new シートの定義
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i≡x 列　column
Private newY, newYmin, newYmax          As Long             ' j≡y 行　row
Private newCnt                          As Long             ' 単独レコードの件数

Public Sub 単独抽出処理_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     レコードの状態ごとにそれぞれのシートに振り分ける
' --------------------------------------+-----------------------------------------
    Dim x, y                            As Long
    Dim CloseingMsg                     As String
    Dim w_rate, w_mod                   As Integer      ' 進捗率 / 表示間隔
    Dim i, iMin, iMax                   As Long         ' 同一レコードの範囲(列 col x)
    Dim j, jMin, jMax                   As Long         ' 同一レコードの範囲(行 row y)
    
'
' ---Procedure Division ----------------+-----------------------------------------
'
' 初期処理
    Call 前処理_R("住所録マージ" & Chr(13) & " 　プログラムを開始します。")
    Call バー表示("住所録マージのプログラムを開始します。")
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
' 昇順ソート　key: (1)ラベルID / (39)姓名key
    With ActiveSheet                '対象シートをアクティブにする
        .Sort.SortFields.Clear      '並び替え条件をクリア
        '項目1
        .Sort.SortFields.Add _
            Key:=.Cells(1, 3), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        '項目2
        .Sort.SortFields.Add _
            Key:=.Range(PKEY_RNG), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
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
    
' ①new シートの初期値
    Set wsNew = wb.Worksheets("①new")
    newYmin = YMIN
    newXmin = XMIN
    wsNew.Activate
    wsNew.Rows(YMIN & ":" & yMax).Select                                    ' クリア
    Selection.ClearContents
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）
    newCnt = 0
    newY = newYmin
' ②archives シートの初期値 ∵ 削除レコード
    Set wsOld = wb.Worksheets("②archives")
    oldYmin = YMIN
    oldXmin = XMIN
    wsOld.Activate
    wsOld.Rows(YMIN & ":" & yMax).Select                                    ' クリア
    Selection.ClearContents
    oldYmax = wsOld.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）
    oldXmax = wsOld.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）
    oldCnt = 0
    oldY = oldYmin
' ③trnChk シートの初期値
    Set wsTrn = wb.Worksheets("③trnChk")
    trnYmin = YMIN
    trnXmin = XMIN
    trnY = trnYmin
    wsTrn.Activate
    wsTrn.Rows(YMIN & ":" & yMax).Select                                    ' クリア
    Selection.ClearContents
    trnYmax = wsTrn.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）
    trnXmax = wsTrn.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）
    trnCnt = 0
    trnY = trnYmin
' シートを元の状態（標準）に戻す
    Sheets("template").Select
    Range(Cells(YMIN, XMIN), Cells(yMax + 1, XMAX)).Copy
    Sheets("③trnChk").Select
    Range(Cells(YMIN, XMIN), Cells(yMax + 1, XMAX)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False     ' コピー状態の解除
    
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
'            wsWrk.Rows(y).Select
'            Selection.ClearContents
            newY = newY + 1
            newCnt = newCnt + 1
        Else
' 同一keyはskip
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


