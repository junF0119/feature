Attribute VB_Name = "CM01_���j���[����"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : Job�����s������Ƃ��̕W���I���j���[�����i�W���Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : CM01_���j���[����
' | @Version    : v1.2.0
' | @updaten    : 2023/05/31
' | @written    : 2023/04/21
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |  �uDocInfo(�폜�s��)�v�V�[�g���Q��
' --------------------------------------+-----------------------------------------
' |  �����K���̓���
' |     Public�ϐ�  �擪��啶��    �� pascalCase
' |     private�ϐ� �擪��������    �� camelCase
' |     �萔        �S�đ啶���A��؂蕶���́A�A���_�[�X�R�A(_) �� snake_case
' |     ����        �ړ���(p_)�����AcamelCase�ɏ�����
' --------------------------------------+-----------------------------------------
'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
'
Public P_backupFile                     As String       ' ���s�O�t�@�C���̕ۑ��p�t�H���_�̃t���p�X
Public P_fullPath                       As String       ' ���sExcel�̃t���p�X+�t�@�C���� �� Thisworkbook
Public P_pathName                       As String       ' ���sExcel�̃t���p�X
Public P_fileName                       As String       ' ���sExcel�̃t�@�C����
' �f�B���N�g���\���̃p�X�Ɩ��O
Public P_rootPath                       As String       ' �V�X�e���t�H���_�̐e�f�B���N�g���̃��[�g�p�X
Public P_sysPath                        As String       ' �V�X�e���t�H���_�܂ł̃t���p�X
Public P_sysName                        As String       ' �V�X�e���t�H���_�̖��O
Public P_subSysPath                     As String       ' �T�u�V�X�e���t�H���_�܂ł̃t���p�X
Public P_subSysName                     As String       ' �T�u�V�X�e���t�H���_�̖��O
' ���s�v���O�����̏��
Public P_sysSymbol                      As String       ' �V�X�e���V���{��
Public P_prgName                        As String       ' ���sExcel�̃v���O������
Public P_version                        As String       ' vx.x.x
Public P_update                         As String       ' yyyymmdd
' �v���O�������s���̓������
Public P_nowY                           As Integer      ' �����̔N�i�����j
Public P_nowM                           As Integer      ' �����̌��i�����j
Public P_nowD                           As Integer      ' �����̓��i�����j
Public P_timeStart                      As Variant      ' �v���O�����J�n�̓��t�Ǝ���
Public P_timeStop                       As Variant      ' �v���O�����I���̓��t�Ǝ���
Public P_timeLap                        As Variant      ' �v���O�������s�̏��v����
Public P_nendoYYYY                      As Integer      ' ���N�x�i����j
' �v���O��������
Public P_mode                           As String       ' ���샂�[�h insert / inquiry / modify / erase / clear / end
                                                        ' �}�N�����̃{�^���ԍ��w����@�@������.xlsm'!'������ n'�@<== �{�^�� n
Public P_menuNum                        As Integer      ' �V�[�g�{�^���̏����ԍ�
Public P_cnt                            As Long         ' ��������
Public P_cntErr                         As Long
' �v���O�����J�n�E�I�����b�Z�[�W
Public P_openingMsg                     As String       ' �v���O�����J�n���b�Z�[�W
Public P_closeingMsg                    As String       ' �v���O��������I�����b�Z�[�W
'
' --------------------------------------------------------------------------------
'   ��private�ϐ�(���Y���W���[�����̃v���V�[�W���Ԃŋ��L�j
'     ��������啶���ɂ���
' �ʒ�`

Public Sub ���j���[����(p_menu As Integer)
' --------------------------------------+-----------------------------------------
' |     ���C������
' |  [���j���[]sheet�̃{�^���̃N���b�N�ŁA���C���v���O�����͌Ăяo�����
' |�@�����œn���ꂽ���j���[�ԍ� Menu �ŏ��������ʂ����s����
' |
' |�@Ref. �{�^���̃}�N���̏�����
' |�@�@�@'#tm89.01-����������-v9.3.4-20201028.xlsm'!'������� 1'
' |�@�@�@'#tm89.01-����������-v9.3.4-20201028.xlsm'!'������� 2'
' |
' |�@Ref. VBA�̃R�[�f�B���O��
' |�@�@�@Public Sub �������(Menu As Integer)
' |
' |
' --------------------------------------+-----------------------------------------

 '
' ---Procedure Division ----------------+-----------------------------------------
       
    MenuNum = p_menu
    NumCnt = 0
    OpeningMsg = ""
    CloseingMsg = ""
    StatusBarMsg = ""
  
    Select Case MenuNum
        Case 1
            Call �V�Z���^�X�V����_R("")
        Case Else
            IsMsgPush ("�v���O�����̃o�O�ł��B ���~���܂��B")
            End
    End Select

End Sub

