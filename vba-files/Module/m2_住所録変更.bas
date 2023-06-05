Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : �B�ύX�Z���^�Ň@����ƇAarchives���X�V����
' --------------------------------------+-----------------------------------------
' | @moduleName : m2_�Z���ύX����
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
' ��ƃV�[�g work �̒�`
Private wsWrk                           As Worksheet
Private wrkX, wrkXmin, wrkXmax          As Long             ' i��x ��@column
Private wrkY, wrkYmin, wrkYmax          As Long             ' j��y �s�@row
' Private wrkCnt                          As Long             ' ��ƃ��R�[�h�̌���=�C���O+�ύX

' �Bnew �V�[�g�̒�`
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i��x ��@column
Private newY, newYmin, newYmax          As Long             ' j��y �s�@row
' Private new1Cnt                         As Long             ' new�@���냌�R�[�h�̌���
' Private new2Cnt                         As Long             ' new�Barchives���R�[�h�̌���

' ' �\���̂̐錾
' Type cntTbl
'     old                                 As long     ' �@����
'     arv                                 As long     ' �Aarchive
'     trn                                 As long     ' �B�ύX�Z���^
'     wrk                                 As long     ' ��ƃ��R�[�h�̌���=�C���O+�ύX
'     new1                                As long     ' new�@���냌�R�[�h�̌���
'     new2                                As long     ' new�Barchives���R�[�h�̌���
'     new3                                As Long     ' new�̕ύX�Z���^�ŐV�K���R�[�h
' End Type
' Dim cnt                                 As cntTbl


Public Sub m2_�Z���ύX����_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     work�V�[�g�̑O��̃��R�[�h���r
' --------------------------------------+-----------------------------------------
    Dim cnt                             As cntTbl
    Dim x, y                            As Long
    Dim CloseingMsg                     As String
    Dim w_rate, w_mod                   As Integer      ' �i���� / �\���Ԋu
    Dim i, iMin, iMax                   As Long         ' ���ꃌ�R�[�h�͈̔�(�� col x)
    Dim j, jMin, jMax                   As Long         ' ���ꃌ�R�[�h�͈̔�(�s row y)
    Dim r                               As Long         ' �ύX���ڂ̗�ԍ�
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
    Set wsNew = Wb.Worksheets(Range("C_newSheet").Value)                    ' �V�Z���^�V�[�g
    newYmin = YMIN
    newXmin = XMIN
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' �ŏI��i�������j
' --------------------------------------+-----------------------------------------
' step1. �P�ƃ��R�[�h�̂ݐ�ɐV�Z���^�V�[�g�ֈړ�����    
    newY = newYmin
    For y = wrkYmin To wrkYmax
        If wsWrk.Cells(y, PKEY_X).Value <> wsWrk.Cells(y + 1, PKEY_X).Value Then
            wsWrk.Cells(y, CHECKED_X) = "NA"
            wsWrk.Rows(y).Copy Destination:=wsNew.Rows(newY)
            wsWrk.Rows(y).ClearContents
            
            Select Case wsNew.Cells(newY, MASTER_X).Value
                Case 1
                    cnt.new1 = cnt.new1 + 1
                Case 2
                    cnt.new2 = cnt.new2 + 1
                Case 3
                    cnt.new3 = cnt.new3 + 1
                Case Else
                    MsgBox "���ʋ敪�G���[=" & wsNew.Cells(newY, MASTER_X).Value
                    End
            End Select
            newY = newY + 1
            GoTo Next_R
        End If

' ����key�̍X�V����
' step1:(42)key���������A(54)���ʋ敪(1�c�@����, 2�c�Aarchives, 3�c�B�ύX�Z���^)�~���Ƀ\�[�g���A�P�ƃ��R�[�h�̍s������
' step2:(42)��������(54)�̃��R�[�h��(54)���ʋ敪�� 0 �Ƃ��ăR�s�[����iafter���R�[�h�j
' step3:�ēx�Astep1�Ɠ��l�Ƀ\�[�g���s��
' step4:����L�[��(42)key�������A3��1or2��0�@�̏��ɕ��Ԃ̂ŁA�ύX���ڂ� 0 ���R�[�h���X�V���A�V�Z���^�V�[�g�փR�s�[����

' ' �I�u�W�F�N�g�ϐ��̒�`�i���ʁj
'     Sheets("work").Activate
' ' �\�̑傫���𓾂�
' ' ��ƃV�[�g�i���̃V�[�g�j�̏����l
'     Set wsWrk = Wb.Worksheets("work")
'     wrkYmin = YMIN
'     wrkXmin = XMIN
'     wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j6��ځi"F")���O��Ōv��
'     wrkXmax = wsWrk.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' �ŏI��i�������j   ' �w�b�_�[�s 3�s�ڂŌv��
'     wrkCnt = wrkYmax - wrkYmin + 1

' �����\�[�g�@key: (39)����key(����)�A(54)���ʋ敪:BA��(�~���j
    With ActiveSheet                '�ΏۃV�[�g���A�N�e�B�u�ɂ���
        .Sort.SortFields.Clear      '���ёւ��������N���A
        '����1
        .Sort.SortFields.Add2 _
             Key:=.Range(PKEY_RNG) _
            , SortOn:=xlSortOnValues _
            , Order:=xlAscending _
            , DataOption:=xlSortNormal
        '����2
        .Sort.SortFields.Add2 _
             Key:=.Range(MASTER_RNG) _
            , SortOn:=xlSortOnValues _
            , Order:=xlDescending _
            , DataOption:=xlSortNormal
'���ёւ������s����
        With .Sort
            .SetRange Range(Cells(wrkYmin - 1, wrkXmin), Cells(wrkYmax, wrkXmax))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
' �\�̑傫���𓾂�
    newYmin = YMIN
    newXmin = XMIN
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' �ŏI��i�������j
' (54)���ʋ敪�� 0 �� after���R�[�h���R�s�[����
    Dim addY as long = wrkYmax          ' �ǉ�����s
    For y = wrkYmin To wrkYmax step 2
        If wsWrk.Cells(y, PKEY_X).Value = wsWrk.Cells(y + 1, PKEY_X).Value Then
            wsWrk.Cells(y+1, CHECKED_X) = "added"
            addY = addY + 1
            wsWrk.Rows(y + 1).Copy Destination:=wsWrk.Rows(addY)
            wsWrk.cells(addY, MASTER_RNG) = 0
        Else
            msgbox "�d���L�[�́A�Q���R�[�h�̃��[���ᔽ�B�v�m�F�I�I"
            Stop
            END
        End If
    next y
stop
' �㏑�����ځF(6)���O�`(15)����
       For r = 6 To 15
            If wsWrk.Cells(y, r).Value <> "" Then
                wsWrk.Cells(y + 1, r).Value = wsWrk.Cells(y, r).Value
            End If
        Next r

' �㏑�����ځF(23)���̑�1�`(26)���l
       For r = 23 To 26
            If wsWrk.Cells(y, r).Value <> "" Then
                wsWrk.Cells(y + 1, r).Value = wsWrk.Cells(y, r).Value
            End If
        Next r

' �㏑�����ځF(36)�X�V���e�`(41)�폜��
        For r = 36 To 41
            If wsWrk.Cells(y, r).Value <> "" Then
                wsWrk.Cells(y + 1, r).Value = wsWrk.Cells(y, r).Value
            End If
        Next r
' ����L�[�Ȃ̂ň��΂�
    y = y + 1
    
Next_R:
    Next y

' �O���[�v���ځF(16)�g�ѓd�b�`(19)��Гd�b
        Dim r1 As Long
        Dim sameCnt As Long
        sameCnt = 0
' �ύX���ڐ����J�E���g
        For r = 16 To 19
            If wsWrk.Cells(y, r).Value <> "" Then
                sameCnt = sameCnt + 1
            End If
        Next r
        
' �ύX���e�����s�Ɠ���̓��e���`�F�b�N
        For r = 16 To 19
            If wsWrk.Cells(y, r).Value <> "" Then
                For r1 = 16 To 19
                    If wsWrk.Cells(y, r).Value = wsWrk.Cells(y + 1, r1).Value Then
                        wsWrk.Cells(y + 1, r1).Value = ""
                        sameCnt = sameCnt - 1
                        Exit For
                    End If
                Next r1
            End If
        Next r
' �Ⴄ���e�̂��̂��󂢂Ă�Z���ɃR�s�[
        If sameCnt <> 0 Then
            For r = 16 To 19
                If wsWrk.Cells(y, r).Value <> "" Then
                    For r1 = 16 To 19
                        If wsWrk.Cells(y + 1, r1).Value = "" Then
                            wsWrk.Cells(y + 1, r1).Value = wsWrk.Cells(y, r).Value
                            wsWrk.Cells(y, r).Value = ""
                            
                            Exit For
                        End If
                    Next r1
                End If
            Next r
        End If
        
        wsWrk.Rows(y).ClearContents
        wsWrk.Cells(y + 1, CHECKED_X) = "Mod"
        wsWrk.Rows(y + 1).Copy Destination:=wsNew.Rows(newY)
        wsWrk.Rows(y + 1).ClearContents
            
        Select Case wsNew.Cells(newY, MASTER_X).Value
            Case 1
                cnt.new1 = cnt.new1 + 1
            Case 2
                cnt.new2 = cnt.new2 + 1
            Case 3
                cnt.new3 = cnt.new3 + 1
            Case Else
                MsgBox "���ʋ敪�G���[=" & wsNew.Cells(newY, MASTER_X).Value
                End
        End Select
        newY = newY + 1
        y = y + 1   ' ����key�������̂ŁA���index�����肠����
        
        
Stop
        
End Sub

' �O���[�v���ځF(20)�g�у��[���`(22)��Ѓ��[��

''
''
''
''
''
''            jMin = trnY                 ' ����key�̍ŏ��̍s(Row, y)
''
''            Do While wsWrk.Cells(y, PKEY_X).Value = wsWrk.Cells(y + 1, PKEY_X).Value
''                wsWrk.Rows(y).Copy Destination:=wsTrn.Rows(trnY)
''                trnCnt = trnCnt + 1
''                wsTrn.Cells(trnY, 42).Value = trnCnt
''                trnY = trnY + 1
''
''
''                wsWrk.Activate
''                wsWrk.Cells(y, CHECKED_X) = "�Btrn"
''                y = y + 1
''            Loop
''            wsWrk.Activate
''            wsWrk.Cells(y, CHECKED_X) = "�Btrn"
''            wsTrn.Activate
''            wsWrk.Rows(y).Copy Destination:=wsTrn.Rows(trnY)
''            trnCnt = trnCnt + 1
''            wsTrn.Cells(trnY, 42).Value = trnCnt
''' �ŐV�̃��R�[�h�𓝍��R�s�[�̌��u�H-999�v�Ƃ���
''            Rows(trnY & ":" & trnY).Select
''            Selection.Copy
''            trnY = trnY + 1
''            Rows(trnY & ":" & trnY).Select
''            ActiveSheet.Paste
''            Application.CutCopyMode = False     ' �R�s�[��Ԃ̉���
'''            Rows(trnY).Insert
''            Cells(trnY, 1) = "�H-999"
''            Cells(trnY, 42) = ""
''            jMax = trnY
''
''' �`�F�b�N�}�[�N�s�ǉ�
''            trnY = trnY + 1
''            Cells(trnY, 1) = "???"
''            Rows(trnY & ":" & trnY).Interior.ColorIndex = xlNone    ' �F�̏�����
''            trnY = trnY + 1
''' �����菇�̎����� / �������m�H-999�n���󔒂̍��ڂ́A�ߋ��̃f�[�^���玝���Ă���
''            For i = INPUTX_FROM To INPUTX_TO + 9
''                If Cells(jMax, i).Value = "" Or _
''                   Cells(jMax, i).Value = "�@" Or _
''                   Cells(jMax, i).Value = " " Then
''                    For j = jMin To jMax - 1
''                        If Cells(j, i).Value <> "" Then
''                            Cells(jMax, i).Value = Cells(j, i).Value        ' �H-999 �s
''                            Cells(jMax + 1, i).Value = Cells(j, 1).Value    ' ???�@�s
''
''                        End If
''                    Next j
''                End If
''            Next i
''        End If


'
' ' 4.���������iEOF���R�[�h�͏����j
'     CloseingMsg = "�����O����" & Chr(9) & "�� " & wrkCnt & Chr(13) & _
'                   "�����㌏��" & Chr(9) & "�� " & newCnt & Chr(13) & _
'                   "  �폜����" & Chr(9) & "�� " & oldCnt & Chr(13) & _
'                   "  �ڎ�����" & Chr(9) & "�� " & trnCnt & Chr(13)
'
''     ' Debug.Print cntAllMsg
'
'
'' �I������
'    MsgBox CloseingMsg
'
'    Call �㏈��_R("�Z���^�}�[�W�v���O�����͐���I�����܂����B" & Chr(13) & CloseingMsg)



'If y = 16 Then
'MsgBox y
'' Debug.Print "|wrk:" & wrkY & "=" & Left(wsWrk.Cells(wrkY, 3), 10) & Chr(9) & _
''             "|new:" & newY & "=" & Left(wsNew.Cells(newY, 3), 10) & Chr(9) & _
''             "|old:" & oldY & "=" & Left(wsold.Cells(oldY, 3), 10) & Chr(9) & _
''             "|trn:" & trnY & "=" & Left(wstrn.Cells(trnY, 3), 10)
'End If


