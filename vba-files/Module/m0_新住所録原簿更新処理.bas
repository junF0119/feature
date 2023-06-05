Attribute VB_Name = "m0_�V�Z���^����X�V����"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : �V�Z���^����X�V�����i���W���[�������Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : m0_�V�Z���^����X�V����
' | @Version    : v1.0.1
' | @update     : 2023/06/01
' | @written    : 2023/05/30
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |�@����Job�́A���̏������s���A�o�^�f�[�^�̐����������؂��A�����ŏC������B
' |1.1�@�@Job�̏��������ƂƂ��āA�X�V�O�̃V�[�g���R�s�[���ꂽ���_�ŁA���̃v���O�����̃o�b�N�A�b�v��ۑ����܂��B
' |���̃v���O�����́A���̓f�[�^�̕ۑS������̓t�@�C���͓ǂނ����ŁA�X�V�͍s���Ă��܂���B
' |����A���̃f�[�^����ꂽ�Ƃ��Ȃǂɂ́A�R�s�[�����V�[�g���畜�����邱�Ƃ��ł��܂��B
' |
' |1.2�@�`�F�b�N���x���́A���Ȃ��A�����C���A�}�j���A���C���Ń`�b�N���Ƀ}�[�N��t��
' |
' |1.3�@�@�m�C�������n�{�^�����������邱�ƂŁA�X�V��̃V�[�g�֏C����̃��R�[�h���R�s�[�����
' |1.4�@�@�R�s�[��́A���ꂼ��̃V�[�g���o�[�W�����Ɠ��t��ύX���A���ꂼ��̃t�H���_�[��Export����
' |
' | �v���O�����\��
' |     1. ��������
' |         1.1 �����V�[�g�̃N���A
' |             importClear_R()
' |         1.2 �O���̃}�X�^�[�̃V�[�g����荞�ށc�c M-�@�V�Z���^���� / M-�AArchives
' |             importSheet_R()
' |
' |     2. �d���L�[�`�F�b�N
' |         2.1 �d���`�F�b�N�c�c (53)PrimaryKey / (42)key����
' |             keyCheck_F()
' |                 arrSet_R()
' |                 duplicateChk_F()
' |                     quickSort_R()
' |
' |
' |
' --------------------------------------+----------------------------------------
' |  �����K���̓���
' |     Public�ϐ�  �擪��啶��    �� pascalCase
' |     private�ϐ� �擪��������    �� camelCase
' |     �萔        �S�đ啶���A��؂蕶���́A�A���_�[�X�R�A(_) �� snake_case
' |     ����        �ړ���(p_)�����AcamelCase�ɏ�����
' --------------------------------------+-----------------------------------------
'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' ���ʗL���V�[�g�T�C�Y�i�f�[�^���݂̗̂̈�j
'
Public Const PKEY_RNG                   As String = "AP3"   ' Key�̃Z���ԍ�
Public Const PKEY_X                     As Long = 42        ' Key�̗�ԍ�"AP"
Public Const PSEIMEI_X                  As Long = 6         ' ��ƈ�̍ő�s���v���̗�ԍ�"C"(���O)
Public Const PDEL_X                     As Long = 41        ' �폜���̗�ԍ�"AO"
Public Const XMIN                       As Long = 1         ' �J�n��
Public Const XMAX                       As Long = 53        ' �ŏI��
Public Const YMIN                       As Long = 4         ' �J�n�s�@��w�b�_�[��������
Public Const yMax                       As Long = 1999      ' �ő�s�@�悱�̃v���O�����ł������ő�s
Public Const INPUTX_FROM                As Long = 6         ' ���͍��ڊJ�n��"F"
Public Const INPUTX_TO                  As Long = 26        ' ���͍��ڏI����"Z"
Public Const CHECKED_X                  As Long = 1         ' �`�F�b�N���i���R�j
Public Const PRIMARYKEY_X               As Long = 53        ' PrimaryKey�̗�"BA"
Public Const MASTER_RNG                 As String = "BB3"   ' work�V�[�g��p�u���ʋ敪�v�̃Z���ԍ�"BB3"

'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +

' �\���̂̐錾
Type cntTbl                                                 ' �e���R�[�h�̌������J�E���g
    old                                 As long = 0         '   �@����
    arv                                 As long = 0         '   �Aarchive
    trn                                 As long = 0         '   �B�ύX�Z���^
    wrk                                 As long = 0         '   work
    new1                                As long = 0         '   new�̌��냌�R�[�h
    new2                                As long = 0         '   new��archivw���R�[�h
    new3                                As long = 0         '   new�̕ύX�Z���^�ŐV�K���R�[�h
End Type
' Dim cnt                             As cntTbl


Public Sub m0_�V�Z���^����X�V����_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |
' | �v���O�����\��
' |     1. ��������
' |         1.1 �����V�[�g�̃N���A
' |             importClear_R()
' |         1.2 �O���̃}�X�^�[�̃V�[�g����荞�ށc�c M-�@�V�Z���^���� / M-�AArchives
' |             importSheet_R()
' |
' |     2. �L�[���ڂ̃`�F�b�N�c�c (53)PrimaryKey / (42)key����
' |         2.1 �d���`�F�b�N
' |             keyCheck_F()
' |                 arrSet_R()
' |                 duplicateChk_F()
' |                     quickSort_R()
' |         2.2 Null�l�`�F�b�N
' |
' |
' --------------------------------------+-----------------------------------------


'
' ---Procedure Division ----------------+-----------------------------------------
'
    Call m1_����������_R("")
    
    Call m2_�Z���ύX����_R("")
    
    Call m9_�I������_R("")
    

End Sub


'If y = 16 Then
'MsgBox y
'Debug.Print "|wrk:" & wrkY & "=" & Left(wsWrk.Cells(wrkY, 3), 10) & Chr(9) & _
'            "|new:" & newY & "=" & Left(wsNew.Cells(newY, 3), 10) & Chr(9) & _
'            "|arv:" & arvY & "=" & Left(wsArv.Cells(arvY, 3), 10) & Chr(9) & _
'            "|eye:" & eyeY & "=" & Left(wsEye.Cells(eyeY, 3), 10)
'End If



