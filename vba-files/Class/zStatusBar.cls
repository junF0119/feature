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
' | @�N���X     : EXCEL�̃X�e�[�^�X�o�[�ɐi�s�󋵂�\��
' --------------------------------------+-----------------------------------------
' | @ClassName  : zStatusBar
' | @Version    : v1.0.0
' | @update     : 2023/06/18
' | @written    : 2023/06/18
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' | �@�X�e�[�^�X�o�[�iExcel�̍����ɂ��郁�b�Z�[�W�̕\���G���A�j�\���ɂ�����A���̃N���X���W���[����
' | ���̏������s���B
' |     1) �����ݒ�Ƃ��āA�X�e�[�^�X�o�[�̎c�����b�Z�[�W������
' |     2) �i����i�߂鐔�̎w��
' |     3) �v���O���X�o�[ (��������������������)
' |     4) �L���Ԃ����\�����X�V����@�\
' |
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' | �R�[�h�e���v���[�g
' |    Dim stBar As New zStatusBar
' |    stBar.�S������ = 1000
' |
' |    Dim i As Long
' |    For i = 1 To 1000
' |
' |        Call stBar.�i����n�i�߂�(1)
' |        If stBar.Is�L����(100) Then
' |            Call stBar.�i����\������
' |        End If
' |    Next
' |
' --------------------------------------+-----------------------------------------
'
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' �v���p�e�B�̒�`

Public TotCnt                           As Long
Public NowCnt                           As Long

'
' ---Procedure Division ----------------+-----------------------------------------
'
' ���\�b�h�̒�`

Public Sub ViewProgress()
    Application.StatusBar = Format(Now(), "m/d hh:mm") & " (" _
                          & NowCnt & "/" & TotCnt & "  " & (NowCnt / TotCnt) * 100 & "%) �c���@ " & progressBar_R
End Sub

Public Sub UpProgress(ByVal p_n As Long)
    NowCnt = NowCnt + p_n

'Application.Wait [Now()] + wait_sec / 86400
''wait_sec = 1 �� 1�b
''wait_sec = 0.5 �� 0.5�b

End Sub

Function Is�L����(ByVal p_��؂�l As Long) As Boolean
    Dim w_ansInt                        As Integer
    Dim w_pInt                          As Integer
    
' ---Procedure Division ----------------+-----------------------------------------
    w_ansInt = Int(NowCnt)
    w_pInt = (p_��؂�l)

    If w_ansInt < w_pInt Then
        Application.StatusBar = ""  'False   ' EXCEL�̃X�e�[�^�X�o�[������
        Is�L���� = False
    Else
        Is�L���� = NowCnt Mod p_��؂�l = 0
    End If
End Function

Private Function progressBar_R() As String
    Dim onGoing                         As Long     ' ���̐�
    
' ---Procedure Division ----------------+-----------------------------------------
    
    onGoing = NowCnt / TotCnt * 10
    If onGoing > 10 Then
        onGoing = 10
    ElseIf onGoing < 1 Then
        onGoing = 0
    End If
    
    progressBar_R = WorksheetFunction.Rept("��", onGoing) _
                  & WorksheetFunction.Rept("��", 10 - onGoing)

Application.Wait [Now()] + 0.2 / 86400
''wait_sec = 1 �� 1�b
''wait_sec = 0.5 �� 0.5�b


    Calculate                         ' �@�ŐV��ԂɍX�V
    
End Function
