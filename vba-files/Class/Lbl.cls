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
' | @�N���X��   : Lbl  �� C-���x���ꗗ�̍s��` / �C���X�^���X
' --------------------------------------+-----------------------------------------
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' �v���p�e�B�̒�`

'Public LblID                            As Integer
Private lblId_                          As Integer
Public ���x����                         As String
Public �ړ���                           As String
Public �����q                           As String

Public Is���x����                       As Boolean

'
' ---Procedure Division ----------------+-----------------------------------------
'
' ���\�b�h�̒�`

Public Sub Greet()
    MsgBox Me.���x���� & "�i" & Me.LblID & "�j"
End Sub

Public Sub Initialize(ByVal values As Variant)

    Select Case TypeName(values)
        Case "Range"

            lblId_ = values(1).Value
            ���x���� = values(2).Value
            �ړ��� = values(3).Value
            �����q = values(4).Value

        Case "Variant()"
            LblID = values(0)
            ���x���� = values(1)
            �ړ��� = values(2)
            �����q = values(3)

    End Select


End Sub

'Public Sub Initialize(ByVal rng As Range)
'
'    lblId_ = rng(1).Value
'    ���x���� = rng(2).Value
'    �ړ��� = rng(3).Value
'    �����q = rng(4).Value
'
'End Sub

'End Sub


' --------------------------------------+-----------------------------------------
' �v���C�x�[�g�ϐ�lblId_���烉�x��ID�̒l���擾
' --------------------------------------+-----------------------------------------
Public Property Get LblID() As String
    LblID = lblId_
End Property

' --------------------------------------+-----------------------------------------
' �v���C�x�[�g�ϐ�lblId_�Ƀ��x��ID�̒l��ݒ�
' --------------------------------------+-----------------------------------------
Public Property Let LblID(ByVal newLblID As String)
    If lblId_ <> 0 Then
        Debug.Print "���ɓo�^���Ă��郉�x��Id�͏㏑���ł��܂���"
    Else
        lblId_ = newLblID
    End If

End Property

'
'





