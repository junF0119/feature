VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : C-ラベル一覧の行定義
' --------------------------------------+-----------------------------------------
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' アプリケーション定数の定義
'
'Public LblID                              As String
Private lblId_                         As long
Public ラベル名                         As String
Public 接頭語                           As String
Public 結合子                           As String

'
' ---Procedure Division ----------------+-----------------------------------------
'

Public Sub Greet()
    MsgBox Me.ラベル名 & "（" & Me.lblId_ & "）"
End Sub

Public Sub Initialize(ByVal values As Variant)

    Select Case TypeName(values)
        Case "Range"
            lblId_ = values(1).Value
            ラベル名 = values(2).Value
            接頭語 = values(3).Value
            結合子 = values(4).Value
            
        Case "Variant()"
            lblId_ = values(0)
            ラベル名 = values(1)
            接頭語 = values(2)
            結合子 = values(3)
            
    End Select
    

End Sub

Public Property Get IsMale() As Boolean
    IsMale = (Me.接頭語 = "男")
End Property

Public Property Get LblID() As String
    LblID = lblId_
End Property

Public Property Let LblID(ByVal newLblID As String)
    If lblId_ <> "" Then
        Debug.Print "既に登録してあるラベルIdは上書きできません"
    Else
        lblId_ = newLblID
    End If
    
End Property








