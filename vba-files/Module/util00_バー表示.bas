Attribute VB_Name = "util00_�o�[�\��"
Option Explicit
' --------------------------------------------------------------------------------
' | @function   EXCEL�̃X�e�[�^�X�o�[�ɐi�s�󋵂�\��
' |-------------------------------------------------------------------------------
' | @moduleName: �o�[�\��
' | @written 2021/12/11
' | @author Jun Fujinawa
' | @license N/A
' | @���O�ݒ�
' |
' |
' --------------------------------------------------------------------------------
'                                       +
Public Sub �o�[�\��(ByVal P_statusMsg As String)
'      EXCEL�̃X�e�[�^�X�o�[�ɐi�s�󋵂�\��

        Application.StatusBar = Format(Now(), "m/d hh:mm") & "�F" & P_statusMsg & " ���������ł��B"
End Sub
