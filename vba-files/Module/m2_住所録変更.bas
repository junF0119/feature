Attribute VB_Name = "m2_�Z���^�ύX"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : �B�ύX�Z���^�Ň@����ƇAarchives���X�V����
' --------------------------------------+-----------------------------------------
' | @moduleName : m2_�Z���^�ύX
' | @Version    : v1.0.0
' | @update     : 2023/06/02
' | @written    : 2023/06/02
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |  �i�P�j�P�ƃ��R�[�h��V�K�V�[�g�i�@new�j�ֈړ�����
' |  �i�Q�j�d�����R�[�h�́A���̃`�F�b�N���s���A�ύX�𔽉f���A�V�K�V�[�g�i�@new�j�ֈړ�����
' |     �@�j(42)����key���������R�[�h�́A�d�����R�[�h�ł��邱��
' |     �A�j(54)���ʋ敪���u1�v(�@����)�������́u2�v(�Aarchives)�̃��R�[�h�ɑ΂�
' |       �u3�v(�B�ύX�Z���^)�̕ύX���ځi�󔒂łȂ����ځj���R�s�[���A�V�K�V�[�g�i�@new�j�ֈړ�����
' |     �B�j�R�s�[����ύX���ڂƓ������e�̂��̂ł��邱�Ƃ��m�F���܂��B
' |     �C�j�u2�v(�Aarchives)�̃��R�[�h�̇B�ύX�Z���^�̍폜���̐���N���u9999�v�̃��R�[�h�́A�@����Ƃ��Ĉړ����܂�
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
Private Wb                              As Workbook         ' ���̃u�b�N
' �Bnew �V�[�g�̒�`
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i��x ��@column
Private newY, newYmin, newYmax          As Long             ' j��y �s�@row
Private new1Cnt                         As Long             ' new�@���냌�R�[�h�̌���
Private new2Cnt                         As Long             ' new�Barchives���R�[�h�̌���

' ' �\���̂̐錾
' Type cntTbl
'     old                                 As long     ' �@����
'     arv                                 As long     ' �Aarchive
'     trn                                 As long     ' �B�ύX�Z���^
'     wrk                                 As long     ' work
'     new1                                As long     ' new�̌��냌�R�[�h
'     new2                                As long     ' new��archivw���R�[�h
' End Type
' dim cnt                             as cntTbl


Public Sub m2_�Z���^�ύX_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     work�V�[�g�̑O��̃��R�[�h���r
' --------------------------------------+-----------------------------------------
    dim cnt                             as cntTbl
    Dim x, y                            As Long
    Dim CloseingMsg                     As String
    Dim w_rate, w_mod                   As Integer      ' �i���� / �\���Ԋu
    Dim i, iMin, iMax                   As Long         ' ���ꃌ�R�[�h�͈̔�(�� col x)
    Dim j, jMin, jMax                   As Long         ' ���ꃌ�R�[�h�͈̔�(�s row y)
    dim r                               as long         ' �ύX���ڂ̗�ԍ�
'
' ---Procedure Division ----------------+-----------------------------------------
'
' �I�u�W�F�N�g�ϐ��̒�`�i���ʁj
    Set Wb = ThisWorkbook
    ' �\�̑傫���𓾂�
    ' ��ƃV�[�g�i���̃V�[�g�j�̏����l
    Set wsWrk = Wb.Worksheets("work")                       ' ��Ɨp�V�[�g�̂��ߌŒ�
    wrkYmin = YMIN
    wrkXmin = XMIN
    wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j3��ځi"C")���O��Ōv��
    wrkXmax = wsWrk.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' �ŏI��i�������j   ' �w�b�_�[�s 3�s�ڂŌv��
    
' �Bnew �V�[�g�̏����l
    Set wsNew = Wb.Worksheets(Range("C_newSheet").Value)    ' �V�Z���^�V�[�g
    newYmin = YMIN
    newXmin = XMIN
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' �ŏI��i�������j
    
    For y = wrkYmin To wrkYmax
        w_rate = Int((y - YMIN) / (wrkYmax - YMIN) * 100)
        If w_rate >= 10 Then
            w_mod = w_rate Mod 10
            If w_mod = 0 Then
                Call �o�[�\��("�i����= " & CStr(w_rate) & "%")
            End If
        End If
        
        If wsWrk.Cells(y, PKEY_X).Value <> wsWrk.Cells(y + 1, PKEY_X).Value Then
            wsWrk.Rows(y).Copy Destination:=wsNew.Rows(newY)
            wsWrk.Activate
            wsWrk.Cells(y, CHECKED_X) = "�@new"
            newY = newY + 1
            if wsWrk.Cells(y, MASTER_X).Value = 1 Then
                cnt.new1 = cnt.new1 + 1
            Else
                cnt.new2 = cnt.new2 + 1
            end If
            goto Next_R
        end if
' ����key�̍X�V����
' �㏑�����ځF(6)���O�`(15)����
' �㏑�����ځF(23)���̑�1�`(26)���l
' �㏑�����ځF(36)�X�V���e�`(41)�폜��
        for r = 6 to 15, 23 to 26, 36 to 41
            If wsWrk.Cells(y, r).Value <> "" then
                wsWrk.Cells(y + 1, r).Value = wsWrk.Cells(y, r).Value
            end if
        next r 
' �O���[�v���ځF(16)�g�ѓd�b�`(19)��Гd�b
        for r = 16 to 19
            If wsWrk.Cells(y, r).Value <> "" then
                dim r1 as long
                dim sw_copy as boolean
                sw_copy = false
                for r1 = 16 to 19
                    if wsWrk.Cells(y + 1, r1).Value = "" then
                        wsWrk.Cells(y + 1, r1).Value = wsWrk.Cells(y, r).Value
                        sw_copy = true
                        exit for
                    end if
                next r1
            end if
        next r

stop

' �O���[�v���ځF(20)�g�у��[���`(22)��Ѓ��[��






            jMin = trnY                 ' ����key�̍ŏ��̍s(Row, y)
            
            Do While wsWrk.Cells(y, PKEY_X).Value = wsWrk.Cells(y + 1, PKEY_X).Value
                wsWrk.Rows(y).Copy Destination:=wsTrn.Rows(trnY)
                trnCnt = trnCnt + 1
                wsTrn.Cells(trnY, 42).Value = trnCnt
                trnY = trnY + 1
                
                
                wsWrk.Activate
                wsWrk.Cells(y, CHECKED_X) = "�Btrn"
                y = y + 1
            Loop
            wsWrk.Activate
            wsWrk.Cells(y, CHECKED_X) = "�Btrn"
            wsTrn.Activate
            wsWrk.Rows(y).Copy Destination:=wsTrn.Rows(trnY)
            trnCnt = trnCnt + 1
            wsTrn.Cells(trnY, 42).Value = trnCnt
' �ŐV�̃��R�[�h�𓝍��R�s�[�̌��u�H-999�v�Ƃ���
            Rows(trnY & ":" & trnY).Select
            Selection.Copy
            trnY = trnY + 1
            Rows(trnY & ":" & trnY).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False     ' �R�s�[��Ԃ̉���
'            Rows(trnY).Insert
            Cells(trnY, 1) = "�H-999"
            Cells(trnY, 42) = ""
            jMax = trnY
           
' �`�F�b�N�}�[�N�s�ǉ�
            trnY = trnY + 1
            Cells(trnY, 1) = "???"
            Rows(trnY & ":" & trnY).Interior.ColorIndex = xlNone    ' �F�̏�����
            trnY = trnY + 1
' �����菇�̎����� / �������m�H-999�n���󔒂̍��ڂ́A�ߋ��̃f�[�^���玝���Ă���
            For i = INPUTX_FROM To INPUTX_TO + 9
                If Cells(jMax, i).Value = "" Or _
                   Cells(jMax, i).Value = "�@" Or _
                   Cells(jMax, i).Value = " " Then
                    For j = jMin To jMax - 1
                        If Cells(j, i).Value <> "" Then
                            Cells(jMax, i).Value = Cells(j, i).Value        ' �H-999 �s
                            Cells(jMax + 1, i).Value = Cells(j, 1).Value    ' ???�@�s
                            
                        End If
                    Next j
                End If
            Next i
        End If
Next_R:
    Next y



 ' 4.���������iEOF���R�[�h�͏����j
     CloseingMsg = "�����O����" & Chr(9) & "�� " & wrkCnt & Chr(13) & _
                   "�����㌏��" & Chr(9) & "�� " & newCnt & Chr(13) & _
                   "  �폜����" & Chr(9) & "�� " & oldCnt & Chr(13) & _
                   "  �ڎ�����" & Chr(9) & "�� " & trnCnt & Chr(13)
                    
'     ' Debug.Print cntAllMsg
    

' �I������
    MsgBox CloseingMsg
    
    Call �㏈��_R("�Z���^�}�[�W�v���O�����͐���I�����܂����B" & Chr(13) & CloseingMsg)

End Sub

'If y = 16 Then
'MsgBox y
'' Debug.Print "|wrk:" & wrkY & "=" & Left(wsWrk.Cells(wrkY, 3), 10) & Chr(9) & _
''             "|new:" & newY & "=" & Left(wsNew.Cells(newY, 3), 10) & Chr(9) & _
''             "|old:" & oldY & "=" & Left(wsold.Cells(oldY, 3), 10) & Chr(9) & _
''             "|trn:" & trnY & "=" & Left(wstrn.Cells(trnY, 3), 10)
'End If


