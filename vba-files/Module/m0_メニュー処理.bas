Attribute VB_Name = "m0_���j���[����"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : Job�����s������Ƃ��̕W���I���j���[�����i�W���Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : m0_���j���[����
' | @Version    : v1.2.0
' | @updaten    : 2023/05/31
' | @written    : 2023/04/21
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |
' | �v���O�����\��
' |     1. �O�����i�V�X�e�����ʁj
' |         1.1 �V�X�e���Ɋւ���Public�ϐ��̎擾
' |         1.2 �I�[�v�j���O���b�Z�[�W�̕\��
' |         1.3 �����O�̓��Y�u�b�N�̃o�b�N�A�b�v�o��
' |     2. �O����2�i�A�v���P�[�V�������ʁj
' |         2.1 �萔�̐ݒ�
' |
' |
' --------------------------------------+-----------------------------------------
' |  �����K���̓���
' |     Public�ϐ�  �擪��啶��    �� pascalCase
' |     private�ϐ� �擪��������    �� camelCase
' |     �萔        �S�đ啶���A��؂蕶���́A�A���_�[�X�R�A(_) �� snake_case
' |     ����        �ړ���(p_)�����AcamelCase�ɏ�����
' --------------------------------------+-----------------------------------------
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' �A�v���P�[�V�����萔�̒�`
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
Public Const MASTER_X                   As Long = 54        ' work�V�[�g��p�u���ʋ敪�v�̗�ԍ�"BB"

' --------------------------------------+-----------------------------------------
' �\���̂̐錾
Public Type cntTbl
    old                                 As Long     ' �@����
    arv                                 As Long     ' �Aarchive
    trn                                 As Long     ' �B�ύX�Z���^
    wrk                                 As Long     ' work
    new1                                As Long     ' new�̌��냌�R�[�h
    new2                                As Long     ' new��archivw���R�[�h
    new3                                As Long     ' new�̕ύX�Z���^�ŐV�K���R�[�h
    mod                                 As Long     ' �ύX���R�[�h
    Add                                 As Long     ' �V�K���R�[�h
End Type
Public Cnt                              As cntTbl
'
' --------------------------------------------------------------------------------
'   ��private�ϐ�(���Y���W���[�����̃v���V�[�W���Ԃŋ��L�j
'     ���������������ɂ���
' �ʒ�`

Public Sub ���j���[����(p_menu As Integer)
' --------------------------------------+-----------------------------------------
' |     ���C������
' |  [���j���[]sheet�̃{�^���̃N���b�N�ŁA���C���v���O�����͌Ăяo�����
' |�@�����œn���ꂽ���j���[�ԍ� Menu �ŏ��������ʂ����s����
' |
' | �v���O�����\��
' |     1. �O�����i�V�X�e�����ʁj
' |         1.1 �V�X�e���Ɋւ���Public�ϐ��̎擾
' |         1.2 �I�[�v�j���O���b�Z�[�W�̕\��
' |         1.3 �����O�̓��Y�u�b�N�̃o�b�N�A�b�v�o��
' |     2. �O����2�i�A�v���P�[�V�������ʁj
' |         2.1 �萔�̐ݒ�
' |
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' |
' |�@Ref. �{�^���̃}�N���̏�����
' |�@�@�@'#tm89.01-����������-v9.3.4-20201028.xlsm'!'������� 1'
' |�@�@�@'#tm89.01-����������-v9.3.4-20201028.xlsm'!'������� 2'
' |
' |�@Ref. VBA�̃R�[�f�B���O��
' |�@�@�@Public Sub �������(Menu As Integer)
' |
' --------------------------------------+-----------------------------------------

'
' ---Procedure Division ----------------+-----------------------------------------
'
    MenuNum = p_menu
' 1. �O�����i�V�X�e�����ʁj
    NumCnt = 0
    OpeningMsg = ""
    CloseingMsg = ""
    StatusBarMsg = ""
    Call �O����_R("")
    Cnt.old = 0                         ' �@����
    Cnt.arv = 0                         ' �Aarchive
    Cnt.trn = 0                         ' �B�ύX�Z���^
    Cnt.wrk = 0                         ' work
    Cnt.new1 = 0                        ' new�̌��냌�R�[�h
    Cnt.new2 = 0                        ' new��archivw���R�[�h
    Cnt.new3 = 0                        ' new�̕ύX�Z���^�ŐV�K���R�[�h
    Cnt.mod = 0                         ' �ύX���R�[�h
    Cnt.Add = 0                         ' �V�K���R�[�h
  
    Select Case MenuNum

        Case 1          ' Step1 �V�Z���^�̍X�V����
        
            Call m1_����������_R("")
            Call m2_���R�[�h�U������_R("")
            Call m3_�ύX���R�[�h����_R("")
            Call m9_�I������_R("")
            
        Case 2          ' Step2 �X�V�ςݐV�Z���^Export
        
            MsgBox "Step2���Ă΂�܂����B"
            
        Case Else
            IsMsgPush ("�v���O�����̃o�O�ł��B ���~���܂��B")
            End
    End Select

End Sub



