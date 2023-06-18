VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' --------------------------------------------------------------------------------
' | @クラス     : EXCELのステータスバーに進行状況を表示
' --------------------------------------+-----------------------------------------
' | @ClassName  : zStatusBar
' | @Version    : v1.0.0
' | @update     : 2023/06/18
' | @written    : 2023/06/18
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' | 　ステータスバー（Excelの左下にあるメッセージの表示エリア）表示にあたり、このクラスモジュールは
' | 次の処理を行う。
' |     1) 初期設定として、ステータスバーの残存メッセージを消去
' |     2) 進捗を進める数の指定
' |     3) プログレスバー (■■□□□□□□□□)
' |     4) キリ番だけ表示を更新する機能
' |
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' | コードテンプレート
' |    Dim stBar As New zStatusBar
' |    stBar.全処理数 = 1000
' |
' |    Dim i As Long
' |    For i = 1 To 1000
' |
' |        Call stBar.進捗をn進める(1)
' |        If stBar.Isキリ番(100) Then
' |            Call stBar.進捗を表示する
' |        End If
' |    Next
' |
' --------------------------------------+-----------------------------------------
'
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' プロパティの定義

Public TotCnt                           As Long
Public NowCnt                           As Long

'
' ---Procedure Division ----------------+-----------------------------------------
'
' メソッドの定義

Public Sub ViewProgress()
    Application.StatusBar = Format(Now(), "m/d hh:mm") & " (" _
                          & NowCnt & "/" & TotCnt & "  " & (NowCnt / TotCnt) * 100 & "%) …→　 " & progressBar_R
End Sub

Public Sub UpProgress(ByVal p_n As Long)
    NowCnt = NowCnt + p_n

'Application.Wait [Now()] + wait_sec / 86400
''wait_sec = 1 → 1秒
''wait_sec = 0.5 → 0.5秒

End Sub

Function Isキリ番(ByVal p_区切り値 As Long) As Boolean
    Dim w_ansInt                        As Integer
    Dim w_pInt                          As Integer
    
' ---Procedure Division ----------------+-----------------------------------------
    w_ansInt = Int(NowCnt)
    w_pInt = (p_区切り値)

    If w_ansInt < w_pInt Then
        Application.StatusBar = ""  'False   ' EXCELのステータスバーを消去
        Isキリ番 = False
    Else
        Isキリ番 = NowCnt Mod p_区切り値 = 0
    End If
End Function

Private Function progressBar_R() As String
    Dim onGoing                         As Long     ' ■の数
    
' ---Procedure Division ----------------+-----------------------------------------
    
    onGoing = NowCnt / TotCnt * 10
    If onGoing > 10 Then
        onGoing = 10
    ElseIf onGoing < 1 Then
        onGoing = 0
    End If
    
    progressBar_R = WorksheetFunction.Rept("■", onGoing) _
                  & WorksheetFunction.Rept("□", 10 - onGoing)

Application.Wait [Now()] + 0.2 / 86400
''wait_sec = 1 → 1秒
''wait_sec = 0.5 → 0.5秒


    Calculate                         ' 　最新状態に更新
    
End Function

