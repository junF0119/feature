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
' �A�v���P�[�V�����萔�̒�`
'
'Public LblID                               As String
Private lblId_                             As String
Public ���O                             As String
Public ����                             As String
Public �a����                           As Date

'
' ---Procedure Division ----------------+-----------------------------------------
'

Public Sub Greet()
    MsgBox Me.���O & "�ł��B�@����ɂ��́I"
End Sub

Public Sub Initialize(ByVal values As Variant)

    Select Case TypeName(values)
        Case "Range"
            lblId_ = values(1).Value
            ���O = values(2).Value
            ���� = values(3).Value
'            �a���� = values(4).Value
            
        Case "Variant()"
            lblId_ = values(0)
            ���O = values(1)
            ���� = values(2)
            �a���� = values(3)
            
    End Select
    

End Sub

Public Property Get IsMale() As Boolean
    IsMale = (Me.���� = "�j")
End Property

Public Property Get LblID() As String
    LblID = lblId_
End Property

Public Property Let LblID(ByVal newLblID As String)
    If lblId_ <> "" Then
        Debug.Print "Id�͏㏑���ł��܂���"
    Else
        lblId_ = newLblID
    End If
    
End Property







