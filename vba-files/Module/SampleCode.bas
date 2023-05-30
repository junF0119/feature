Attribute VB_Name = "SampleCode"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function  マージ処理（標準版）
' --------------------------------------+-----------------------------------------
' | @moduleName : CM10_マージ処理
' | @Version    : v1.0.20
' | @update     : 2023/05/05
' | @written    : 2023/04/19
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |
' |
' |   ※public変数(当該プロジェクト内のモジュール間で共有)は、最初に呼ばれるプロシジャーに定義
' |     接頭語に P_ をつける
' --------------------------------------+-----------------------------------------
'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' 共通有効シートサイズ（データ部のみの領域）
Const HV                                As String = "９－ＥＯＦ" ' HightValueの代替値
Const PKEY_RNG                          As String = "AP3"   ' machingKeyのセル番号
Const PKEY_X                            As Long = 42        ' machingKeyの列番号"AP"
Const PKEY_SHIFTJIS                     As Long = 29        ' shift-jisのKey
Const PSEIMEI_X                         As Long = 3         ' 作業域の最大行数計測の列番号"C"(名前)
Const PDEL_X                            As Long = 38        ' 削除日の列番号"AL"
Const XMIN                              As Long = 1         ' 開始列
Const XMAX                              As Long = PKEY_X + 1 ' 最終列
Const YMIN                              As Long = 4         ' 開始行　∵ヘッダー部を除く
Const INPUTX_FROM                       As Long = 3         ' 入力開始項目
Const INPUTX_TO                         As Long = 23        ' 入力終了項目
' 外部Excelファイルの定義　①Tranzaction / ②oldMaster
Private wbSrc                           As Workbook
Private wsSrc                           As Worksheet
Private srcFile                         As String           ' 外部Excelファイルのパス
Private srcSheet                        As String           ' 　〃　のimportするシート（①trn / ②old）
Private srcMsg                          As String           ' ファイルを選択するときのメッセージ
Private srcImport                       As String           ' import先のシート

' 作業シート work の定義
Private wbWork                          As Workbook         ' このブック
Private wsWork                          As Worksheet
Private workX, workXmin, workXmax       As Long             ' i≡x 列　column
Private workY, workYmin, workYmax       As Long             ' j≡y 行　row

' ①trn シートの定義
Private wsTrn                           As Worksheet
Private trnX, trnXmin, trnXmax          As Long             ' i≡x 列　column
Private trnY, trnYmin, trnYmax          As Long             ' j≡y 行　row
Private cntTrn                          As Long             ' 追加変更のレコード件数
' ②old シートの定義
Private wsOld                           As Worksheet
Private oldX, oldXmin, oldXmax          As Long             ' i≡x 列　column
Private oldY, oldYmin, oldYmax          As Long             ' j≡y 行　row
Private cntOld                          As Long             ' 旧マスターのレコード件数
' ③new シートの定義
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i≡x 列　column
Private newY, newYmin, newYmax          As Long             ' j≡y 行　row
Private cntNew                          As Long             ' 更新済みマスターの件数
Private cntMatch                        As Long             ' マッチしたレコード件数
' update シートの定義 ∵ trn = old & trn内容 ≠ oldの内容
Private wsUp                            As Worksheet
Private upX, upXmin, upXmax             As Long             ' i≡x 列　column
Private upY, upYmin, upYmax             As Long             ' j≡y 行　row
Private cntUp                           As Long             ' マスタの変更があったレコード件数
' archives シートの定義 ∵ trn > old 削除レコード
Private wsArv                           As Worksheet
Private ArvX, ArvXmin, ArvXmax          As Long             ' i≡x 列　column
Private ArvY, ArvYmin, ArvYmax          As Long             ' j≡y 行　row
Private cntArv                          As Long             ' 削除レコードの件数

Public Sub マージ処理_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     ①Tranzaction と ②oldMaster を一対一のマッチングを行い、③newMaster を出力
' --------------------------------------+-----------------------------------------
    Dim i                               As Long
'
' ---Procedure Division ----------------+-----------------------------------------
'
' オブジェクト変数の定義（共通）
    Set wbWork = ThisWorkbook
    Set wsTrn = wbWork.Worksheets(Range("C_trnImport").Value)
    Set wsOld = wbWork.Worksheets(Range("C_oldImport").Value)
    Set wsNew = wbWork.Worksheets(Range("C_newImport").Value)
    Set wsUp = wbWork.Worksheets(Range("C_update").Value)
    Set wsArv = wbWork.Worksheets(Range("C_archive").Value)
'    Set wsWork = wbWork.Worksheets(Range("C_work").Value)  ' シートを削除するためimportSheet_Rで定義
    
' ①trn シートのクリア
    Sheets("メニュー").Activate
    Call importClear_R(Range("C_trnImport"))
' ②old シートのクリア
    Sheets("メニュー").Activate
    Call importClear_R(Range("C_oldImport"))
' ③new シートのクリア
    Sheets("メニュー").Activate
    Call importClear_R(Range("C_newImport"))
' ④update シートのクリア
    Sheets("メニュー").Activate
    Call importClear_R(Range("C_update"))
' ⑤Archives シートのクリア
    Sheets("メニュー").Activate
    Call importClear_R(Range("C_archive"))

' 1.①Tranzaction Excel を「①trn」シートにImport
    Sheets("メニュー").Activate
    srcFile = Range("C_trnFile")                            ' Tranzaction Excel パスを指定
    srcSheet = Range("C_trn")                               ' 〃 ブックの住所録シート名を（①trn / ②old）指定
    srcImport = Range("C_trnImport")                        ' import先のシート名
    srcMsg = "追加・変更のある［①Tranzaction］ Excel を選択してください。"
    Call importSheet_R("")
    
    Range("C_trnFile") = srcFile                            ' 選択したファイルのパスを入力欄へセット
    trnYmin = YMIN
    trnXmin = XMIN
    trnYmax = workYmax                                      ' 最終行（縦方向）
    trnXmax = workXmax                                      ' 最終列（横方向）

' 2.②oldMaster Excel を「②old」シートにImport
    Sheets("メニュー").Activate
    srcFile = Range("C_oldMst")                             ' old Master Excel パスを指定
    srcSheet = Range("C_old")                               ' 〃 ブックの住所録シート名を（①trn / ②old）指定
    srcImport = Range("C_oldImport")                        ' import先のシート名
    srcMsg = "現行住所録［②oldMaster］ Excel を選択してください。"
    Call importSheet_R("")
    
    Range("C_oldMst") = srcFile                             ' 選択したファイルのパスを入力欄へセット
    oldYmin = YMIN
    oldXmin = XMIN
    oldYmax = workYmax                                      ' 最終行（縦方向）
    oldXmax = workXmax                                      ' 最終列（横方向）

' 3.①Transaction と ②oldMaster を「key姓名」で一対一のマッチングを行う
    Call maching_R("")

' 4.件数整理（EOFレコードは除く）
    CloseingMsg = "trn件数" & Chr(9) & "＝ " & cntTrn - 1 & Chr(13) & _
                    "old件数" & Chr(9) & "＝ " & cntOld - 1 & Chr(13) & _
                    "new件数" & Chr(9) & "＝ " & cntNew & Chr(13) & _
                    "変更無し" & Chr(9) & "＝ " & cntMatch & Chr(13) & _
                    "変更有り" & Chr(9) & "＝ " & cntUp & Chr(13) & _
                    "追加" & Chr(9) & "＝ " & cntTrn & Chr(13) & _
                    "削除" & Chr(9) & "＝ " & cntArv & Chr(13)
                    
    ' Debug.Print cntAllMsg

End Sub

Private Sub importClear_R(ByVal p_sheetName As String)
' --------------------------------------+-----------------------------------------
' | ※既存のシートを削除し、importシートをコピーすると、importするシートの名前定義も
' | 一緒にコピーされ、名前定義の重複でロジックに不具合が生じるので、既存シートの削除
' | でなくシートのクリアとfieldのコピーで対応することに変更する。
' |
' --------------------------------------+-----------------------------------------
    Dim wsTemp                          As Worksheet
    Dim tempX, tempXmin, tempXmax       As Long             ' i≡x 列　column
    Dim tempY, tempYmin, tempYmax       As Long             ' j≡y 行　row
'
' ---Procedure Division ----------------+-----------------------------------------
'
    Sheets(p_sheetName).Activate
'  シートに関係なく、一律クリア
    Range(Cells(YMIN, XMIN), Cells(1000, 100)).Select
    Selection.ClearContents

End Sub

Private Sub importSheet_R(ByVal dummy As Variant) '
' --------------------------------------+-----------------------------------------
' |  作業用シート work のI/O
' --------------------------------------+-----------------------------------------
' | 処理1:［srcFile］のExcelファイルの［srcSheet］シートをこのブックの［work］シートにimportする。
' |     　※前提：Excelシートのフォーマットは同じとする。
' | 処理2: 4行目以降を「key姓名」で昇順ソートする。
' | 処理3: ソート後に［work］シートを［srcSheet］へ上書きコピーする。
' |
' --------------------------------------+-----------------------------------------
' workシートの定数セット
    Set wsWork = wbWork.Worksheets(Range("C_work").Value)
    workYmin = YMIN                                         ' j≡y 行　row
    workXmin = XMIN                                         ' i≡x 列　column
    workYmax = workYmax                                     ' 最終行（縦方向）
    workXmax = workXmax                                     ' 最終列（横方向）

    Dim sw_FalseTrue                    As Boolean
    Dim i, y                            As Long
    Dim contentsPath                    As String
    
'
' ---Procedure Division ----------------+-----------------------------------------

' 1.import Excel ファイルの設定　/　①Tranzaction  ②oldMaster
' --------------------------------------+-----------------------------------------
' import シートの削除
    
    If IsExistSheet("work") Then
        wbWork.Worksheets("work").Delete                    ' 以前のシートを削除
    End If
' ファイル指定の有無チェック　/　指定したファイルがなかったら、エクスプローラで指定させる
    sw_FalseTrue = False
    If srcFile <> "" Then
        If IsExistFileDir(srcFile) Then
            sw_FalseTrue = True
        Else
            sw_FalseTrue = False
        End If
    End If
'　[ファイルを開く]ダイアログボックスで対象Excelを選択します
    If sw_FalseTrue = False Then
        If Range("C_childPath").Value = "" Then
            contentsPath = PathName
        Else
            contentsPath = SubSysPath & "\" & Range("C_childPath").Value
        End If
        ChDir contentsPath                                    ' プログラムのあるフォルダを指定
        srcFile = Application.GetOpenFilename("Excelファイル,*.xl*", , srcMsg)
        sw_FalseTrue = True
    End If
' 外部Excelファイルを開き、importシートを作業シート work へコピー
    Workbooks.Open srcFile
    Set wbSrc = ActiveWorkbook
    wbSrc.Worksheets(srcSheet).Copy after:=wbWork.Worksheets(1)
    ActiveSheet.Name = "work"
 
' 表の大きさを得る
    Set wsWork = wbWork.Worksheets("work")
    workYmax = wsWork.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row                   ' 最終行（縦方向）1列目（"A")で計測
    workXmax = wsWork.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column   ' 最終列（横方向）   ' ヘッダー行 3行目で計測
' 表の最終行の後にHV値を挿入
    workYmax = workYmax + 1
    wsWork.Cells(workYmax, 1) = HV
    wsWork.Cells(workYmax, PKEY_X) = HV     ' HV ≡ ９－ＥＯＦ　（cf.全角）
' shift-jisのkeyをunicodeに変換
    For y = YMIN To workYmax
        wsWork.Cells(y, PKEY_X) = StrConv(wsWork.Cells(y, PKEY_SHIFTJIS), vbUnicode)    ' Shift_JIS → UTF-16
    Next y

' 昇順ソート　key: 姓名key
    With ActiveSheet
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=.Range(PKEY_RNG), Order:=xlAscending
        .Sort.SetRange .Range(Cells(YMIN, XMIN), Cells(workYmax, XMAX))
        .Sort.Apply
    End With

' 最終行の再計算
    workYmax = wsWork.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row    ' 最終行（縦方向）1列目（"A")で計測

' シートwork を　import先のシートへ上書きコピー
    wbWork.Worksheets("work").Cells.Copy wbWork.Worksheets(srcImport).Range("A1")
    
' 保存しないでclose
    wbSrc.Close saveChanges:=False

' オブジェクト変数の解放
    Set wbSrc = Nothing
    Set wsSrc = Nothing
    Set wsWork = Nothing
     
End Sub

Private Sub maching_R(ByVal dummy As Variant)

'[ファイルの特性および前提]
'（１）マスターファイル（Mstと略す）
'    ①システムに必要な全項目のデータを有すること｡
'    ②keyの重複はないこと｡
'（２）トランザクションファイル(Trnと略す)
'    ①Trnのレコードフォーマットは､Mstと同じであること｡
'    ②Mstで変更になったレコードを有すること｡
'    ③keyと変更のあった項目を最低限有すること｡
'    ④変更のない項目・レコードも含むことがあること。
'（３）マッチングの判定基準
'┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
'┃ 一対一のマッチング処理
'┣━━━━━━━━━━━━━━━┯━━━━━━━┯━━━━━━━┯━━━━━━━┯━━━━━━━┯━━━━━━━┯━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
'┃#    compare   │  trn  │  old  │  new  │ update│  arv  │   付加条件
'┠───────────────┼───────┼───────┼───────┼───────┼───────┼─────────────────────────────────────────────────────
'┃1.1  trn = old │  (7)  │  7    │  (7)  │  N/A  │  N/A  │  変更無　全項目同じ⇒ trnをコピー
'┃1.2  trn = old │  9+   │  9    │  9+   │trn+old│  N/A  │  変更有　違う項目がある⇒要目視確認
'┃1.3  trn = old │  10x  │  10   │  N/A  │  N/A  │ 10x+10│  trnだ削除なので、trn,oldともarvへコピー
'┃1.4  trn = old │  EOF  │  EOF  │  N/A  │  N/A  │  N/A  │  EOF　プログラム終了
'┠───────────────┼───────┼───────┼───────┼───────┼───────┼─────────────────────────────────────────────────────
'┃2.1  trn < old │  13x  │  15   │  N/A  │  N/A  │  13x  │  trnが削除でoldに同じkeyがなかったので、arvへコピー
'┃2.2  trn < old │  1    │  N/A  │  1    │  N/A  │  N/A  │  oldに同じkeyがないので、追加としてnewへコピー
'┃2.3  trn < old │  16   │  EOF  │  16   │  N/A  │  N/A  │  old=EOFなので、追加としてnewへコピー
'┠───────────────┼───────┼───────┼───────┼───────┼───────┼─────────────────────────────────────────────────────
'┃3.1  trn > old │  N/A  │  5    │  5    │  N/A  │  N/A  │   変更無①
'┃3.2  trn > old │  14   │  15   │  14   │  N/A  │  N/A  │   追加
'┃3.3  trn > old │  EOF  │  15   │  15   │  N/A  │  N/A  │   trn=EOFなので変更なしで、newへコピー
'┗━━━━━━━━━━━━━━━┷━━━━━━━┷━━━━━━━┷━━━━━━━┷━━━━━━━┷━━━━━━━┷━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
'（4）マッチングkeyの取り扱い
'　　VBAでは、HighValue値( ≡ Allbit 1)（以下「HV」）を扱えないので、代替手段を次の通りとり、マッチングロジックを合理的にする。
'    ① machingKey = 1-もともとのキー　（例）キーが姓名で「山田　太郎」とすると、スペースを除いた　姓名　に
'    　接頭語として 1- を付けて、1-山田太郎　とする。
'    ② HV値として、9-EOF　を一番最後に追加する。
'
'
' --------------------------------------+-----------------------------------------
    Dim cntAllMsg                       As String
    Dim trnEof, oldEOF                  As Boolean
    Dim i                               As Long
'
' ---Procedure Division ----------------+-----------------------------------------
'
    Sheets("メニュー").Activate

' 表の大きさを得る
' 更新済みの新マスターと追加レコード
    Set wsNew = Worksheets(Range("C_newImport").Value)
    newYmin = YMIN
    newXmin = XMIN
    newYmax = newYmin                                       ' 最終行（縦方向）
    newXmax = XMAX                                          ' 最終列（横方向）
    cntNew = 0
    cntMatch = 0
    
' 旧マスタのデータ変更があったレコード～目視チェック用（新旧比較）
    Set wsUp = wbWork.Worksheets(Range("C_update").Value)
    upYmin = YMIN
    upXmin = XMIN
    upYmax = upYmin                                         ' 最終行（縦方向）
    upXmax = XMAX                                           ' 最終列（横方向）
    cntUp = 0                                               ' cntUp = cntNew - cntMatch
    
' 削除レコード～削除日が記載されたレコード
    Set wsArv = wbWork.Worksheets(Range("C_archive").Value)
    ArvYmin = YMIN
    ArvXmin = XMIN
    ArvYmax = ArvYmin                                       ' 最終行（縦方向）
    ArvXmax = XMAX                                          ' 最終列（横方向）
    cntArv = 0

    cntTrn = 0
    cntOld = 0
    trnEof = False
    oldEOF = False
    trnY = trnYmin
    oldY = oldYmin
    newY = newYmin
    upY = upYmin
    ArvY = ArvYmin
    
    Do Until trnEof = True And oldEOF = True
    
'If wsTrn.Cells(trnY, PKEY_X) = "１－カテナ船水" Then
'    wsTrn.Activate
''    MsgBox "trn=" & trnY
'End If
'
'If wsOld.Cells(oldY, PKEY_X) = "１－カテナ船水" Then
'    wsOld.Activate
''    MsgBox "trn=" & trnY
'End If
        
'Debug.Print ">trn:" & trnY & "…" & wsTrn.Cells(trnY, PKEY_X) & _
'            ">old:" & oldY & "…" & wsOld.Cells(oldY, PKEY_X) & _
'            ">new:" & newY & "…" & wsNew.Cells(newY, PKEY_X)
        
Debug.Print "trn:" & trnY & Chr(9) & _
            "|old:" & oldY & Chr(9) & _
            "|new:" & newY
            
        Select Case wsTrn.Cells(trnY, PKEY_X)               ' key姓名 /  削除日　x=PDEL_Xcol
            Case Is = wsOld.Cells(oldY, PKEY_X)             ' trn = old → match  trnをnewへコピー
                Call matchChk_R("")                     ' 変更内容がないかチェック

            Case Is > wsOld.Cells(oldY, PKEY_X)             ' trn > old → oldMasterのみ　newへそのままコピー
                
                wsOld.Rows(oldY).Copy Destination:=wsNew.Rows(newY)
                newY = newY + 1
                oldY = oldY + 1
                cntOld = cntOld + 1
                cntNew = cntNew + 1

            Case Is < wsOld.Cells(oldY, PKEY_X)             ' trn < Old → Transactionのみ 追加レコード


                wsTrn.Rows(trnY).Copy Destination:=wsNew.Rows(newY)
                trnY = trnY + 1
                newY = newY + 1
                cntTrn = cntTrn + 1
                cntNew = cntNew + 1
                    
        End Select
' EOF 判定
        If trnY > trnYmax Then
            trnEof = True
        End If
        If oldY > oldYmax Then
            oldEOF = True
        End If
    Loop
    

' 新マスター③new から削除レコードを⑤Archiveシートへ移動 & トレーラレコード 9-EOF を削除
    wsNew.Activate
' 表の大きさを得る
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row            ' 最終行（縦方向）3列目（"C")で計測
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column   ' 最終列（横方向）   ' ヘッダー行 3行目で計測
    
    ArvY = YMIN - 1
    For i = newYmin To newYmax
        If wsNew.Cells(i, PDEL_X) <> "" Then
            ArvY = ArvY + 1
            wsNew.Rows(i).Copy Destination:=wsArv.Rows(ArvY)
            wsNew.Rows(i).Select
            Selection.ClearContents
        End If
        If wsNew.Cells(i, PKEY_X) = "9-EOF" Then
            wsNew.Rows(i).Select
            Selection.ClearContents
        End If
    Next i

' マッチング終了/オブジェクト変数の解放
    Set wsTrn = Nothing
    Set wsOld = Nothing
    Set wsNew = Nothing
    Set wsArv = Nothing

End Sub


'' 昇順ソート　key: 姓名key
''    With ActiveSheet
''        .Sort.SortFields.Clear
''        .Sort.SortFields.Add Key:=.Range(PKEY_RNG), Order:=xlAscending
''        .Sort.SetRange .Range(Cells(YMIN, XMIN), Cells(workYmax, XMAX))
''        .Sort.Apply
''    End With
''
''    ActiveWindow.SmallScroll Down:=-15
''    Range("A3:AM657").Select
''    ActiveWorkbook.Worksheets("③new").Sort.SortFields.Clear
''    ActiveWorkbook.Worksheets("③new").Sort.SortFields.Add2 Key:=Range("AM4:AM657" _
''        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
''    With ActiveWorkbook.Worksheets("③new").Sort
''        .SetRange Range("A3:AM657")
''        .Header = xlYes
''        .MatchCase = False
''        .Orientation = xlTopToBottom
''        .SortMethod = xlPinYin
''        .Apply
''    End With
'
'
'    ActiveSheet.Sort.SortFields.Clear
'    ActiveSheet.Sort.SortFields.Add2 Key:=Range(PKEY_RNG) _
'        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With ActiveSheet.Sort
'        .SetRange Range(Cells(YMIN, XMIN), Cells(newYmax, newXmax))
'        .Header = xlGuess
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'    Sheets("③new").Activate
'If wsTrn.Cells(trnY, PKEY_X) = "1-ＳＰＡ服部" Then
'    wsTrn.Activate
''    MsgBox "trn=" & trnY
'End If
'If oldY > 25 Then
'    wsNew.Activate
''    MsgBox "old=" & oldY
'End If
'
'If wsOld.Cells(oldY, PKEY_X) = "1-ＳＳＫ永野" Then
'    wsOld.Activate
''    MsgBox "old=" & oldY
'End If
            
'Debug.Print ">trn:" & trnY & "…" & wsTrn.Cells(trnY, PKEY_X) & _
'            ">old:" & oldY & "…" & wsOld.Cells(oldY, PKEY_X) & _
'            ">new:" & newY & "…" & wsNew.Cells(newY, PKEY_X)
'
'Debug.Print "<trn:" & trnY & "…" & wsTrn.Cells(trnY, PKEY_X) & _
'            "<old:" & oldY & "…" & wsOld.Cells(oldY, PKEY_X) & _
'            "<new:" & newY & "…" & wsNew.Cells(newY, PKEY_X)

'Debug.Print "=trn:" & trnY & "…" & wsTrn.Cells(trnY, PKEY_X) & _
'            "=old:" & oldY & "…" & wsOld.Cells(oldY, PKEY_X) & _
'            "=new:" & newY & "…" & wsNew.Cells(newY, PKEY_X)
'

Private Sub matchChk_R(ByVal p_delRec As String)
' --------------------------------------+-----------------------------------------
' |     変更箇所のチェック
' |     変更があるときは、目視用に④updateシートへ出力
' --------------------------------------+-----------------------------------------
    Dim x                               As Long
    Dim trnRec                          As Variant
    Dim oldRec                          As Variant
'
' ---Procedure Division ----------------+-----------------------------------------
'
' 入力項目のみ結合比較
    trnRec = ""
    For x = INPUTX_FROM To INPUTX_TO
        trnRec = trnRec & wsTrn.Cells(trnY, x)              ' 文字列と数値の結合
    Next x
    oldRec = ""
    For x = INPUTX_FROM To INPUTX_TO
        oldRec = oldRec & wsOld.Cells(oldY, x)
    Next x
    
 If wsTrn.Cells(trnY, PKEY_X) = "1-ＳＰＡ服部" Then
    wsTrn.Activate
'    MsgBox "trn=" & trnY
End If
    
'結合した項目同士を比較し、等しければ変更なしで、そのまま新マスタへ登録
    If trnRec = oldRec Then
        wsTrn.Rows(trnY).Copy Destination:=wsNew.Rows(newY)

        oldY = oldY + 1
        trnY = trnY + 1
        newY = newY + 1
        cntTrn = cntTrn + 1
        cntOld = cntOld + 1
        cntNew = cntNew + 1
        cntMatch = cntMatch + 1
        Exit Sub
    End If
    
' --------------------------------------+-----------------------------------------
' |     updateシートへレコードをコピー   目視用
' |     変更があるときは、目視用に④updateシートへ出力
' |  ⅰ)レコードは、１行目にold、２行目にtrnの順にコピー
' |　ⅱ)３行目は、newでtrnをコピー
' |　ⅲ)oldにあり、trn(≡new)が空白の項目は、oldの項目をnewへコピー
' |　ⅳ)変更のあった項目には、４行目のその場所に「???」をセット
' |　ⅴ)new行の背景を赤、文字は白に変更
' |　ⅵ)目視で内容をチェック、修正はマニュアルで実施
' |　ⅶ)有効なnew行を「new」シートへコピー
' |　ⅷ)「new」シートを新規ブックとして出力
' |　ⅸ)ブック名は、統合した住所録の番号を付す
' |　　　①+②……+⑧
' |
' --------------------------------------+-----------------------------------------
    
    cntUp = cntUp + 1
    wsUp.Activate
    wsOld.Rows(oldY).Copy Destination:=wsUp.Rows(upY)
    wsUp.Cells(upY, XMAX + 1) = "old"
    wsTrn.Rows(trnY).Copy Destination:=wsUp.Rows(upY + 1)
    wsUp.Cells(upY + 1, XMAX + 1) = "trn"
    wsTrn.Rows(trnY).Copy Destination:=wsUp.Rows(upY + 2)
    wsUp.Cells(upY + 2, XMAX + 1) = "new"
    wsUp.Cells(upY + 3, XMAX + 1) = "???"
    
    Rows(upY + 2).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    For x = INPUTX_FROM To XMAX
        If wsUp.Cells(upY, x) <> "" And wsUp.Cells(upY + 1, x) = "" Then
            wsUp.Cells(upY + 2, x) = wsUp.Cells(upY, x)   ' old → new
            wsUp.Cells(upY + 3, x) = "???"
        End If
    Next x
    
    upY = upY + 4
    oldY = oldY + 1
    trnY = trnY + 1
'    newY = newY + 1
    cntTrn = cntTrn + 1
    cntOld = cntOld + 1
'    cntNew = cntNew + 1
    cntUp = cntUp + 1
    
End Sub



