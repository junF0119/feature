Attribute VB_Name = "m9_�I������"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : �I������
' --------------------------------------+-----------------------------------------
' | @moduleName : m9_�I������
' | @Version    : v1.1.0
' | @update     : 2023/05/22
' | @written    : 2023/05/16
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
' |        �v���O�����\��
' |            1. ��������
' |                1.1 �����V�[�g�̃N���A
' |                    importClear_R()
' |                1.2 �O���̃}�X�^�[�̃V�[�g����荞�ށc�c M-�@�V�Z���^���� / M-�AArchives
' |                    importSheet_R()(
' |            2. �f�[�^�̐���������
' |                2.1  �d���`�F�b�N�c�c (53)PrimaryKey / (42)key����
' |                    keyCheck_F()
' |                        arrSet_R()
' |                        duplicateChk_F()
' |                            quickSort_R()
' |                 2.2�@�L�[���ڂ�Null�l�`�F�b�N�ƕ���
' |
' |
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

Public Sub m9_�I������_R(ByVal dummy As Variant)
'' �\���̂̐錾
'Public Type cntTbl
'    old                                 As Long     ' �@����
'    arv                                 As Long     ' �Aarchive
'    trn                                 As Long     ' �B�ύX�Z���^
'    wrk                                 As Long     ' work
'    new1                                As Long     ' new�̌��냌�R�[�h
'    new2                                As Long     ' new��archivw���R�[�h
'    new3                                As Long     ' new�̕ύX�Z���^�ŐV�K���R�[�h
'    mod                                 As Long     ' �ύX���R�[�h
'    add                                 As Long     ' �V�K���R�[�h
'End Type
'Public Cnt                              As cntTbl

'
' ---Procedure Division ----------------+-----------------------------------------
'
' 9.0 �I������


    CloseingMsg = "|�@����V�[�g" & Chr(9) & "�� " & Cnt.old & Chr(13) _
                & "|�Aarchives" & Chr(9) & "�� " & Cnt.arv & Chr(13) _
                & "|�B�ύX�Z���^" & Chr(9) & "�� " & Cnt.trn & Chr(13) _
                & "|��ƃ��R�[�h" & Chr(9) & "�� " & Cnt.wrk & Chr(13) _
                & "|�Z���^(�X�V��)" & Chr(9) & "�� " & Cnt.new1 + Cnt.new2 + Cnt.new3 & Chr(13) _
                & "| (����)�@���e" & Chr(9) & "�� " & Cnt.new1 & Chr(13) _
                & "| (����)�Aarchive" & Chr(9) & "�� " & Cnt.new2 & Chr(13) _
                & "| (����)�B�V�K" & Chr(9) & "�� " & Cnt.new3 & Chr(13) _
                & "|�ύX���R�[�h" & Chr(9) & "�� " & Cnt.mod & Chr(13) _
                & "|�ǉ����R�[�h" & Chr(9) & "�� " & Cnt.Add & Chr(13)
                
                
    Call �㏈��_R(CloseingMsg & Chr(13) & "�v���O�����͐���I�����܂����B")
    

End Sub






