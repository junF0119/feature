Attribute VB_Name = "m4_更新済住所録Export"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : ①原簿と②archivesレコードが混在した新住所録シートからそれぞれのExcelブックを出力
' --------------------------------------+-----------------------------------------
' | @moduleName : m4_住所録Export
' | @Version    : v1.0.0
' | @update     : 2023/06/11
' | @written    : 2023/06/11
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |  （１）重複レコードは、次のチェックを行い、変更を反映し、新規シート（①new）へ移動する
' |     ⅰ）(42)姓名keyが同じレコードは、重複レコードであること
' |     ⅱ）(54)識別区分が「1」(①原簿)もしくは「2」(②archives)のレコードに対し
' |       「3」(③変更住所録)の変更項目（空白でない項目）をコピーし、新規シート（①new）へ移動する
' |     ⅲ）コピーする変更項目と現行が同じ内容のものでないことを確認します。
' |         同じ内容であれば、変更なしとしてあつかうこと。
' |     ⅳ）「2」(②archives)のレコードの③変更住所録の削除日の西暦年が「9999」のレコードは、①原簿として移動します
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
Private Wb                              As Workbook         ' このブック
' ③new シートの定義
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i≡x 列　column
Private newY, newYmin, newYmax          As Long             ' j≡y 行　row
'' 構造体の宣言
'Public Type cntTbl
'    old                                 As Long     ' ①原簿
'    arv                                 As Long     ' ②archive
'    trn                                 As Long     ' ③変更住所録
'    wrk                                 As Long     ' work
'    new1                                As Long     ' newの原簿レコード
'    new2                                As Long     ' newのarchivwレコード
'    new3                                As Long     ' newの変更住所録で新規レコード
'    mod                                 As Long     ' 変更レコード
'    add                                 As Long     ' 新規レコード
'End Type
'Public Cnt                              As cntTbl

Public Sub m4_更新済住所録Export_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     新住所録シートから①住所録原簿と②ArchiveのExcelを出力
' --------------------------------------+-----------------------------------------
'
    Dim newCnt                          As Long: newCnt = 0
    Dim arvCnt                          As Long: arvCnt = 0
    Dim saveDir                         As String

'
' ---Procedure Division ----------------+-----------------------------------------
'
' オブジェクト変数の定義（共通）
    Set Wb = ThisWorkbook

' ③new シートの初期値 & 表の大きさを得る
    Set wsNew = Wb.Worksheets(Range("C_newSheet").Value)                    ' 新住所録シート
    newYmin = YMIN
    newXmin = XMIN
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' 最終行（縦方向）
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' 最終列（横方向）
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

    Sheets("新住所録").Select
    Sheets("新住所録").Copy Before:=Sheets(1)
    Rows("3:3").Select
    Selection.AutoFilter
    Rows("3:3").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("新住所録 (2)").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("新住所録 (2)").Sort.SortFields.Add2 key:=Range( _
        "BB4:BB801"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("新住所録 (2)").Sort
        .SetRange Range("A3:WWN801")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll ToRight:=34
    ActiveWindow.SmallScroll Down:=138
    Range("BB152:BB546").Select
    Selection.EntireRow.Delete
    Sheets("新住所録 (2)").Select
    Sheets("新住所録 (2)").Copy
    Application.Left = -1174.25
    Application.Top = 272.5
    Windows("zz2.1.1-新住所録更新-v1.3.0.take1-20230611.xlsm").Activate
    Sheets("⑨label").Select
    Sheets("⑨label").Copy Before:=Workbooks("Book1").Sheets(1)
    Application.Left = 163
    Application.Top = 97
    ActiveWorkbook.Names("C_ラベル一覧").Delete
    ActiveWorkbook.Names("pathName").Delete
    ActiveWorkbook.Names("pathName").Delete
    ActiveWorkbook.Names("update").Delete
    ActiveWorkbook.Names("update").Delete
    ActiveWorkbook.Names("updateTxt").Delete
    ActiveWorkbook.Names("updateTxt").Delete
    ActiveWorkbook.Names("updateTxt2").Delete
    ActiveWorkbook.Names("verShort").Delete
    ActiveWorkbook.Names("verShort").Delete
    ActiveWorkbook.Names("version").Delete
    ActiveWorkbook.Names("version").Delete
    ActiveWorkbook.Names("verUpdateTxt2").Delete
    ActiveWorkbook.Names("verUpdateTxt2").Delete
    Sheets("新住所録 (2)").Select
    Sheets("新住所録 (2)").Move Before:=Sheets(1)
    Sheets("新住所録 (2)").Select
    Sheets("新住所録 (2)").Name = "①原簿"
    ChDir "D:\Desktop\2.Job1-新住所録更新(zz2.1)-v1.3\1.1.inputData"
    ActiveWorkbook.SaveAs FileName:= _
        "D:\Desktop\2.Job1-新住所録更新(zz2.1)-v1.3\1.1.inputData\M-①新住所録原簿-v1.1.0-20230611.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End Sub



Sub Sample6()
    Dim i As Long, Target As String
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Target = Cells(i, 1)
        Sheets(2).Name = Target & "_予算"
        Sheets(3).Name = Target & "_実績"
        Sheets(Array(Target & "_予算", Target & "_実績")).Copy
        ActiveWorkbook.SaveAs "D:\Work\" & Target & ".xlsx"
        ActiveWorkbook.Close
    Next i
End Sub
















' --------------------------------------+-----------------------------------------
' 同一キーの(42)key姓名が、3→1or2→0　の順に並ぶので、変更項目で 0 レコードを更新し、新住所録シートへコピーする
    For j = wrkYmin To wrkYmax Step 3
        wsWrk.Range(Cells(j, 1), Cells(j, wrkXmax)).Font.Color = rgbBlack                   ' 文字色：黒　      #000000#
        wsWrk.Range(Cells(j, 1), Cells(j, wrkXmax)).Interior.Color = rgbKhaki               ' 背景色：カーキ    #8ce6f0#
        wsWrk.Range(Cells(j + 1, 1), Cells(j + 1, wrkXmax)).Font.Color = rgbSnow            ' 文字色：スノー    #fafaff#
        wsWrk.Range(Cells(j + 1, 1), Cells(j + 1, wrkXmax)).Interior.Color = rgbDodgerBlue  ' 背景色：ドジャーブルー    #ff901e#

        sw_change = False
        For i = 6 To 41
            Select Case i
' 上書き項目：(6)名前～(15)方書、(23)その他1～(26)備考
                Case 6 To 15, 23 To 26
                    If wsWrk.Cells(j, i).Value <> "" Then
                        If wsWrk.Cells(j, i).Value <> wsWrk.Cells(j + 1, i).Value Then
                            wsWrk.Cells(j + 2, i).Value = wsWrk.Cells(j, i).Value
                            wsWrk.Cells(j, i).Font.Color = rgbTeal            ' 文字色：青緑      #808000#
                            wsWrk.Cells(j, i).Interior.Color = rgbLightCoral  ' 背景色：薄いさんご #8080f0#
                            wsWrk.Cells(j + 1, i).Font.Color = rgbSnow        ' 文字色：スノー    #fafaff#
                            wsWrk.Cells(j + 1, i).Interior.Color = rgbDarkRed ' 背景色：濃い赤    #00008b#
                            wsWrk.Cells(j + 2, i).Font.Color = rgbSnow        ' 文字色：スノー    #fafaff#
                            wsWrk.Cells(j + 2, i).Interior.Color = rgbDarkRed ' 背景色：濃い赤    #00008b#

                            sw_change = True
                        End If
                    End If
                    
' 管理項目：(36)更新内容～(41)削除日
                Case 36 To 41
                    If wsWrk.Cells(j, i).Value <> "" Then
                        If wsWrk.Cells(j, i).Value <> wsWrk.Cells(j + 1, i).Value Then
                            
                        End If
                    End If
                    
' グループ項目：(16)携帯電話～(19)会社電話
                Case 16
                    Call modifyItem_R(j, i, 16, 19, sw_change)

' グループ項目：(20)携帯メール～(22)会社メール
                Case 20
                    Call modifyItem_R(j, i, 20, 22, sw_change)

                Case Else
            End Select
        Next i
        
' 変更した項目があるときは、管理項目も更新する
        If sw_change Then
            For i = 36 To 41
                wsWrk.Cells(j + 2, i).Value = wsWrk.Cells(j, i).Value
            Next i
        End If
        
 ' 更新した行を新住所録シートへコピー
        wsWrk.Cells(j, CHECKED_X) = "trn"
        wsWrk.Cells(j + 1, CHECKED_X) = "before"
        newYmax = newYmax + 1
        wsWrk.Rows(j + 2).Copy Destination:=wsNew.Rows(newYmax)
        wsNew.Cells(newYmax, newXmax) = wsWrk.Cells(j + 1, wrkXmax)

'        wsWrk.Range(Cells(j + 2, 1), Cells(j + 2, wrkXmax)).Font.Color = rgbSnow             ' 文字色：スノー    #fafaff#
'        wsWrk.Range(Cells(j + 2, 1), Cells(j + 2, wrkXmax)).Interior.Color = rgbDodgerBlue   ' 背景色：ドジャーブルー    #ff901e#
    
        Select Case wsWrk.Cells(j + 1, MASTER_X).Value
            Case 1
                Cnt.new1 = Cnt.new1 + 1
            Case 2
                Cnt.new2 = Cnt.new2 + 1
            Case 3
                Cnt.new3 = Cnt.new3 + 1
            Case Else
                MsgBox "識別区分エラー=" & wsNew.Cells(newY, MASTER_X).Value
                End
        End Select
        
' If newYmax = 884 Then
' Stop
' End If
'
        If sw_change Then
            wsNew.Cells(newYmax, CHECKED_X) = "Modify"
            Cnt.mod = Cnt.mod + 1
        Else
            wsNew.Cells(newYmax, CHECKED_X) = "Add"
            Cnt.Add = Cnt.Add + 1
        End If
    
    Next j

End Sub


Private Sub modifyItem_R(ByVal p_j As Long _
                       , ByVal p_i As Long _
                       , ByVal p_from As Long _
                       , ByVal p_to As Long _
                       , ByRef p_modifySw As Boolean)
' --------------------------------------+-----------------------------------------
' | @function   : 変更レコードで更新
' | @moduleName : modifyItem_R
' | @remarks
' | 引数の意味
' | 引　数：p_j           行位置
' | 引　数：p_i           列位置
' | 引　数：p_from        列位置の開始
' | 引　数：p_to          列位置の終了
' | 戻り値：p_modifySw    変更有り ≡ True  、変更なし ≡ False
' |
' --------------------------------------+-----------------------------------------
    Dim Cnt                             As cntTbl
    Dim x, xx                          As Long
    Dim sameCnt                         As Long
'
' ---Procedure Division ----------------+-----------------------------------------
'
' --------------------------------------+-----------------------------------------
'   グループ項目：(16)携帯電話～(19)会社電話
'   グループ項目：(20)携帯メール～(22)会社メール
' --------------------------------------+-----------------------------------------

'                       wsWrk.Cells(j, i).Font.Color = rgbSnow          ' 文字色：スノー    #fafaff#
'                       wsWrk.Cells(j, i).Interior.Color = rgbDarkRed   ' 背景色：濃い赤    #00008b#

'                        Else
'                            wsWrk.Cells(j, CHECKED_X).Value = "same"
'If p_j = 7 Then
'Stop
'End If
'

' 変更項目数をカウント
    sameCnt = 0
    For x = p_from To p_to
        If wsWrk.Cells(p_j, x).Value <> "" Then
            sameCnt = sameCnt + 1
        End If
    Next x

' 変更内容が現行と同一の内容かチェック
    For x = p_from To p_to
        If wsWrk.Cells(p_j, x).Value <> "" Then
            For xx = p_from To p_to
                If wsWrk.Cells(p_j, x).Value = wsWrk.Cells(p_j + 2, xx).Value Then
'                    wsWrk.Cells(p_j, x).Value = ""                         ' 同じ値が既にあるので、変更項目は、消去
                    wsWrk.Cells(p_j, x).Font.Strikethrough = True           ' 取り消し線を付ける
                    wsWrk.Cells(p_j, x).Font.Bold = True                    ' 太字に設定
                    wsWrk.Cells(p_j, x).Font.Color = rgbNavy                ' 文字色：ネイビー  #800000#
                    wsWrk.Cells(p_j, x).Interior.Color = rgbSnow            ' 背景色：スノー    #fafaff#
                    wsWrk.Cells(p_j + 1, xx).Font.Color = rgbNavy              ' 文字色：ネイビー  #800000#
                    wsWrk.Cells(p_j + 1, xx).Interior.Color = rgbSnow          ' 背景色：スノー    #fafaff#
                    sameCnt = sameCnt - 1
                    Exit For
                End If
            Next xx
        End If
    Next x

' 違う内容のものを空いてるセルにコピー
    If sameCnt <> 0 Then
        For x = p_from To p_to
            If wsWrk.Cells(p_j, x).Font.Strikethrough = False Then  ' 取り消し線の項目は、既に　登録済み
                If wsWrk.Cells(p_j, x).Value <> "" Then
                    For xx = p_from To p_to
                        If wsWrk.Cells(p_j + 2, xx).Value = "" Then
                            wsWrk.Cells(p_j + 2, xx).Value = wsWrk.Cells(p_j, x).Value
    '                        wsWrk.Cells(p_j, x).Value = ""                 ' 更新できたので、消去
                            wsWrk.Cells(p_j, x).Font.Color = rgbSnow          ' 文字色：スノー    #fafaff#
                            wsWrk.Cells(p_j, x).Interior.Color = rgbDarkRed   ' 背景色：濃い赤    #00008b#
                            wsWrk.Cells(p_j + 1, xx).Font.Color = rgbSnow        ' 文字色：スノー    #fafaff#
                            wsWrk.Cells(p_j + 1, xx).Interior.Color = rgbDarkRed ' 背景色：濃い赤    #00008b#
                            wsWrk.Cells(p_j + 2, xx).Font.Color = rgbSnow        ' 文字色：スノー    #fafaff#
                            wsWrk.Cells(p_j + 2, xx).Interior.Color = rgbDarkRed ' 背景色：濃い赤    #00008b#
                            p_modifySw = True
                            sameCnt = sameCnt - 1
                            Exit For
                        End If
                    Next xx
                End If
            End If
        Next x
    End If

    If sameCnt <> 0 Then
        MsgBox "変更しきれなかった項目が残っています＝" & sameCnt
        Stop
    End If

End Sub
        


