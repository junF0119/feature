VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' --------------------------------------+-----------------------------------------
' | @クラス名   : Lbl  ≡ C-ラベル一覧の行定義 / インスタンス
' --------------------------------------+-----------------------------------------
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' プロパティの定義

'Public LblID                            As Integer
Private lblId_                          As Integer
Public ラベル名                         As String
Public 接頭語                           As String
Public 結合子                           As String

Public Isラベル名                       As Boolean

'
' ---Procedure Division ----------------+-----------------------------------------
'
' メソッドの定義

Public Sub Greet()
    MsgBox Me.ラベル名 & "（" & Me.LblID & "）"
End Sub

Public Sub Initialize(ByVal values As Variant)

    Select Case TypeName(values)
        Case "Range"

            lblId_ = values(1).Value
            ラベル名 = values(2).Value
            接頭語 = values(3).Value
            結合子 = values(4).Value

        Case "Variant()"
            LblID = values(0)
            ラベル名 = values(1)
            接頭語 = values(2)
            結合子 = values(3)

    End Select


End Sub

'Public Sub Initialize(ByVal rng As Range)
'
'    lblId_ = rng(1).Value
'    ラベル名 = rng(2).Value
'    接頭語 = rng(3).Value
'    結合子 = rng(4).Value
'
'End Sub

'End Sub


' --------------------------------------+-----------------------------------------
' プライベート変数lblId_からラベルIDの値を取得
' --------------------------------------+-----------------------------------------
Public Property Get LblID() As String
    LblID = lblId_
End Property

' --------------------------------------+-----------------------------------------
' プライベート変数lblId_にラベルIDの値を設定
' --------------------------------------+-----------------------------------------
Public Property Let LblID(ByVal newLblID As String)
    If lblId_ <> 0 Then
        Debug.Print "既に登録してあるラベルIdは上書きできません"
    Else
        lblId_ = newLblID
    End If

End Property

'
'






