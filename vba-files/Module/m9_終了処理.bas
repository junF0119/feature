Attribute VB_Name = "m9_終了処理"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : 終了処理
' --------------------------------------+-----------------------------------------
' | @moduleName : m9_終了処理
' | @Version    : v1.1.0
' | @update     : 2023/05/22
' | @written    : 2023/05/16
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
' |                    importSheet_R()(
' |            2. データの整合性検証
' |                2.1  重複チェック…… (53)PrimaryKey / (42)key姓名
' |                    keyCheck_F()
' |                        arrSet_R()
' |                        duplicateChk_F()
' |                            quickSort_R()
' |                 2.2　キー項目のNull値チェックと復旧
' |
' |
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

Public Sub m9_終了処理_R(ByVal dummy As Variant)
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

'
' ---Procedure Division ----------------+-----------------------------------------
'
' 9.0 終了処理


    CloseingMsg = "|①原簿シート" & Chr(9) & "＝ " & Cnt.old & Chr(13) _
                & "|②archives" & Chr(9) & "＝ " & Cnt.arv & Chr(13) _
                & "|③変更住所録" & Chr(9) & "＝ " & Cnt.trn & Chr(13) _
                & "|作業レコード" & Chr(9) & "＝ " & Cnt.wrk & Chr(13) _
                & "|住所録(更新後)" & Chr(9) & "＝ " & Cnt.new1 + Cnt.new2 + Cnt.new3 & Chr(13) _
                & "| (内訳)①原稿" & Chr(9) & "＝ " & Cnt.new1 & Chr(13) _
                & "| (内訳)②archive" & Chr(9) & "＝ " & Cnt.new2 & Chr(13) _
                & "| (内訳)③新規" & Chr(9) & "＝ " & Cnt.new3 & Chr(13) _
                & "|変更レコード" & Chr(9) & "＝ " & Cnt.mod & Chr(13) _
                & "|追加レコード" & Chr(9) & "＝ " & Cnt.Add & Chr(13)
                
                
    Call 後処理_R(CloseingMsg & Chr(13) & "プログラムは正常終了しました。")
    

End Sub






