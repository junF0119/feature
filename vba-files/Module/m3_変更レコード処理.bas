Attribute VB_Name = "m3_�ύX���R�[�h����"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : �ύX���R�[�h�Ň@����܂��͇Aarchives���R�[�h���X�V����
' --------------------------------------+-----------------------------------------
' | @moduleName : m3_�ύX���R�[�h����
' | @Version    : v1.0.0
' | @update     : 2023/06/06
' | @written    : 2023/06/06
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |  �i�P�j�d�����R�[�h�́A���̃`�F�b�N���s���A�ύX�𔽉f���A�V�K�V�[�g�i�@new�j�ֈړ�����
' |     �@�j(42)����key���������R�[�h�́A�d�����R�[�h�ł��邱��
' |     �A�j(54)���ʋ敪���u1�v(�@����)�������́u2�v(�Aarchives)�̃��R�[�h�ɑ΂�
' |       �u3�v(�B�ύX�Z���^)�̕ύX���ځi�󔒂łȂ����ځj���R�s�[���A�V�K�V�[�g�i�@new�j�ֈړ�����
' |     �B�j�R�s�[����ύX���ڂƌ��s���������e�̂��̂łȂ����Ƃ��m�F���܂��B
' |         �������e�ł���΁A�ύX�Ȃ��Ƃ��Ă��������ƁB
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
' ��ƃV�[�g work �̒�`
Private wsWrk                           As Worksheet
Private wrkX, wrkXmin, wrkXmax          As Long             ' i��x ��@column
Private wrkY, wrkYmin, wrkYmax          As Long             ' j��y �s�@row

' �Bnew �V�[�g�̒�`
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i��x ��@column
Private newY, newYmin, newYmax          As Long             ' j��y �s�@row
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

Public Sub m3_�ύX���R�[�h����_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     work�V�[�g�Ɏc�������R�[�h���X�V
' --------------------------------------+-----------------------------------------
'    Dim Cnt                             As cntTbl
    Dim x, y, z                         As Long
    Dim CloseingMsg                     As String
    Dim w_rate, w_mod                   As Integer      ' �i���� / �\���Ԋu
    Dim i, iMin, iMax                   As Long         ' ���ꃌ�R�[�h�͈̔�(�� col x)
    Dim j, jMin, jMax                   As Long         ' ���ꃌ�R�[�h�͈̔�(�s row y)
    Dim r                               As Long         ' �ύX���ڂ̗�ԍ�

    Dim sw_change                       As Boolean      ' true �� �ύX�ӏ��L��@/�@False �� �V�@����

'
' ---Procedure Division ----------------+-----------------------------------------
'
' �I�u�W�F�N�g�ϐ��̒�`�i���ʁj
    Set Wb = ThisWorkbook
    ' �\�̑傫���𓾂�
    ' ��ƃV�[�g�i���̃V�[�g�j�̏����l
    Set wsWrk = Wb.Worksheets("work")                                       ' ��Ɨp�V�[�g�̂��ߌŒ�
    wrkYmin = YMIN
    wrkXmin = XMIN
    wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j3��ځi"C")���O��Ōv��
    wrkXmax = wsWrk.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' �ŏI��i�������j�w�b�_�[�s 3�s�ڂŌv��
    
' �Bnew �V�[�g�̏����l
    Set wsNew = Wb.Worksheets(Range("C_newSheet").Value)                    ' �V�Z���^�V�[�g
    newYmin = YMIN
    newXmin = XMIN
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' �ŏI��i�������j

' --------------------------------------+-----------------------------------------
' ����L�[��(42)key�������A3��1or2��0�@�̏��ɕ��Ԃ̂ŁA�ύX���ڂ� 0 ���R�[�h���X�V���A�V�Z���^�V�[�g�փR�s�[����
    For j = wrkYmin To wrkYmax Step 3
        wsWrk.Range(Cells(j + 1, 1), Cells(j + 1, wrkXmax)).Font.Color = rgbSnow             ' �����F�F�X�m�[    #fafaff#
        wsWrk.Range(Cells(j + 1, 1), Cells(j + 1, wrkXmax)).Interior.Color = rgbDodgerBlue   ' �w�i�F�F�h�W���[�u���[    #ff901e#

        sw_change = False
        For i = 6 To 41
            Select Case i
' �㏑�����ځF(6)���O�`(15)�����A(23)���̑�1�`(26)���l
                Case 6 To 15, 23 To 26
                    If wsWrk.Cells(j, i).Value <> "" Then
                        If wsWrk.Cells(j, i).Value <> wsWrk.Cells(j + 1, i).Value Then
                            wsWrk.Cells(j + 2, i).Value = wsWrk.Cells(j, i).Value
                            wsWrk.Cells(j, i).Font.Color = rgbSnow          ' �����F�F�X�m�[    #fafaff#
                            wsWrk.Cells(j, i).Interior.Color = rgbDarkRed   ' �w�i�F�F�Z����    #00008b#
                            wsWrk.Cells(j + 1, i).Font.Color = rgbSnow        ' �����F�F�X�m�[    #fafaff#
                            wsWrk.Cells(j + 1, i).Interior.Color = rgbDarkRed ' �w�i�F�F�Z����    #00008b#
                            wsWrk.Cells(j + 2, i).Font.Color = rgbSnow        ' �����F�F�X�m�[    #fafaff#
                            wsWrk.Cells(j + 2, i).Interior.Color = rgbDarkRed ' �w�i�F�F�Z����    #00008b#

                            sw_change = True
                        End If
                    End If
                    
' �Ǘ����ځF(36)�X�V���e�`(41)�폜��
                Case 36 To 41
                    If wsWrk.Cells(j, i).Value <> "" Then
                        If wsWrk.Cells(j, i).Value <> wsWrk.Cells(j + 1, i).Value Then
                            
                        End If
                    End If
                    
' �O���[�v���ځF(16)�g�ѓd�b�`(19)��Гd�b
                Case 16
                    Call modifyItem_R(j, i, 16, 19, sw_change)

' �O���[�v���ځF(20)�g�у��[���`(22)��Ѓ��[��
                Case 20
                    Call modifyItem_R(j, i, 20, 22, sw_change)

                Case Else
            End Select
        Next i
        
' �ύX�������ڂ�����Ƃ��́A�Ǘ����ڂ��X�V����
        If sw_change Then
            For i = 36 To 41
                wsWrk.Cells(j + 2, i).Value = wsWrk.Cells(j, i).Value
            Next i
        End If
        
 ' �X�V�����s��V�Z���^�V�[�g�փR�s�[
        wsWrk.Cells(j, CHECKED_X) = "trn"
        wsWrk.Cells(j + 1, CHECKED_X) = "before"
        newYmax = newYmax + 1
        wsWrk.Rows(j + 2).Copy Destination:=wsNew.Rows(newYmax)

'        wsWrk.Range(Cells(j + 2, 1), Cells(j + 2, wrkXmax)).Font.Color = rgbSnow             ' �����F�F�X�m�[    #fafaff#
'        wsWrk.Range(Cells(j + 2, 1), Cells(j + 2, wrkXmax)).Interior.Color = rgbDodgerBlue   ' �w�i�F�F�h�W���[�u���[    #ff901e#
    
        Select Case wsWrk.Cells(j + 1, MASTER_X).Value
            Case 1
                Cnt.new1 = Cnt.new1 + 1
            Case 2
                Cnt.new2 = Cnt.new2 + 1
            Case 3
                Cnt.new3 = Cnt.new3 + 1
            Case Else
                MsgBox "���ʋ敪�G���[=" & wsNew.Cells(newY, MASTER_X).Value
                End
        End Select
        
        If sw_change Then
            wsNew.Cells(newYmax, CHECKED_X) = "Modify"
            Cnt.mod = Cnt.mod + 1
        Else
            wsNew.Cells(newYmax, CHECKED_X) = "Add"
            Cnt.add = Cnt.add + 1
        End If
    
    Next j

End Sub


Private Sub modifyItem_R(ByVal p_j As Long _
                       , ByVal p_i As Long _
                       , ByVal p_from As Long _
                       , ByVal p_to As Long _
                       , ByRef p_modifySw As Boolean)
' --------------------------------------+-----------------------------------------
' | @function   : �ύX���R�[�h�ōX�V
' | @moduleName : modifyItem_R
' | @remarks
' | �����̈Ӗ�
' | ���@���Fp_j           �s�ʒu
' | ���@���Fp_i           ��ʒu
' | ���@���Fp_from        ��ʒu�̊J�n
' | ���@���Fp_to          ��ʒu�̏I��
' | �߂�l�Fp_modifySw    �ύX�L�� �� True  �A�ύX�Ȃ� �� False
' |
' --------------------------------------+-----------------------------------------
    Dim Cnt                             As cntTbl
    Dim x, xx                          As Long
    Dim sameCnt                         As Long
'
' ---Procedure Division ----------------+-----------------------------------------
'
' --------------------------------------+-----------------------------------------
'   �O���[�v���ځF(16)�g�ѓd�b�`(19)��Гd�b
'   �O���[�v���ځF(20)�g�у��[���`(22)��Ѓ��[��
' --------------------------------------+-----------------------------------------

'                       wsWrk.Cells(j, i).Font.Color = rgbSnow          ' �����F�F�X�m�[    #fafaff#
'                       wsWrk.Cells(j, i).Interior.Color = rgbDarkRed   ' �w�i�F�F�Z����    #00008b#

'                        Else
'                            wsWrk.Cells(j, CHECKED_X).Value = "same"

' �ύX���ڐ����J�E���g
    sameCnt = 0
    For x = p_from To p_to
        If wsWrk.Cells(p_j, x).Value <> "" Then
            sameCnt = sameCnt + 1
        End If
    Next x

' �ύX���e�����s�Ɠ���̓��e���`�F�b�N
    For x = p_from To p_to
        If wsWrk.Cells(p_j, x).Value <> "" Then
            For xx = p_from To p_to
                If wsWrk.Cells(p_j, x).Value = wsWrk.Cells(p_j + 2, xx).Value Then
'                    wsWrk.Cells(p_j, x).Value = ""                         ' �����l�����ɂ���̂ŁA�ύX���ڂ́A����
                    wsWrk.Cells(p_j, x).Font.Color = rgbNavy                ' �����F�F�l�C�r�[  #800000#
                    wsWrk.Cells(p_j, x).Interior.Color = rgbSnow            ' �w�i�F�F�X�m�[    #fafaff#
                    wsWrk.Cells(p_j + 1, x).Font.Color = rgbNavy              ' �����F�F�l�C�r�[  #800000#
                    wsWrk.Cells(p_j + 1, x).Interior.Color = rgbSnow          ' �w�i�F�F�X�m�[    #fafaff#
                    sameCnt = sameCnt - 1
                    Exit For
                End If
            Next xx
        End If
    Next x

' �Ⴄ���e�̂��̂��󂢂Ă�Z���ɃR�s�[
    If sameCnt <> 0 Then
        For x = p_from To p_to
            If wsWrk.Cells(p_j, x).Value <> "" Then
                For xx = p_from To p_to
                    If wsWrk.Cells(p_j + 2, xx).Value = "" Then
                        wsWrk.Cells(p_j + 2, xx).Value = wsWrk.Cells(p_j, x).Value
'                        wsWrk.Cells(p_j, x).Value = ""                 ' �X�V�ł����̂ŁA����
                        wsWrk.Cells(p_j, x).Font.Color = rgbSnow          ' �����F�F�X�m�[    #fafaff#
                        wsWrk.Cells(p_j, x).Interior.Color = rgbDarkRed   ' �w�i�F�F�Z����    #00008b#
                        wsWrk.Cells(p_j + 1, xx).Font.Color = rgbSnow        ' �����F�F�X�m�[    #fafaff#
                        wsWrk.Cells(p_j + 1, xx).Interior.Color = rgbDarkRed ' �w�i�F�F�Z����    #00008b#
                        wsWrk.Cells(p_j + 2, xx).Font.Color = rgbSnow        ' �����F�F�X�m�[    #fafaff#
                        wsWrk.Cells(p_j + 2, xx).Interior.Color = rgbDarkRed ' �w�i�F�F�Z����    #00008b#
                        p_modifySw = True
                        sameCnt = sameCnt - 1
                        Exit For
                    End If
                Next xx
            End If
        Next x
    End If

    If sameCnt <> 0 Then
        MsgBox "�ύX������Ȃ��������ڂ��c���Ă��܂���" & sameCnt
        Stop
    End If

End Sub
        
