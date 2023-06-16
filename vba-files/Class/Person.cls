VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : Class Lesson 14
' --------------------------------------+-----------------------------------------
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' アプリケーション定数の定義
'
'Public LblID                               As String
Private lblId_                             As String
Public 名前                             As String
Public 性別                             As String
Public 誕生日                           As Date

'
' ---Procedure Division ----------------+-----------------------------------------
'

Public Sub Greet()
    MsgBox Me.名前 & "です。　こんにちは！"
End Sub

Public Sub Initialize(ByVal values As Variant)

    Select Case TypeName(values)
        Case "Range"
            lblId_ = values(1).Value
            名前 = values(2).Value
            性別 = values(3).Value
'            誕生日 = values(4).Value
            
        Case "Variant()"
            lblId_ = values(0)
            名前 = values(1)
            性別 = values(2)
            誕生日 = values(3)
            
    End Select
    

End Sub

Public Property Get IsMale() As Boolean
    IsMale = (Me.性別 = "男")
End Property

Public Property Get LblID() As String
    LblID = lblId_
End Property

Public Property Let LblID(ByVal newLblID As String)
    If lblId_ <> "" Then
        Debug.Print "Idは上書きできません"
    Else
        lblId_ = newLblID
    End If
    
End Property








