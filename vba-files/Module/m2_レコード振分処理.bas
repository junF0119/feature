Attribute VB_Name = "m2_���R�[�h�U������"
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


Public Sub m2_���R�[�h�U������_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     work�V�[�g�̑O��̃��R�[�h���r
' --------------------------------------+-----------------------------------------
    Dim cnt                             As cntTbl
    Dim x, y, z                         As Long
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
        Else
            y = y + 1   ' ����L�[�Ȃ̂�1�s�X�L�b�v
        End If
    Next y


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
    
    wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j
' (54)���ʋ敪�� 0 �� after���R�[�h���R�s�[����
    Dim addY As Long          ' �ǉ�����s
    wrkY = wrkYmax

    For y = wrkYmin To wrkYmax Step 2
        If wsWrk.Cells(y, PKEY_X).Value = wsWrk.Cells(y + 1, PKEY_X).Value Then
            wsWrk.Cells(y + 1, CHECKED_X) = "before"
            wrkY = wrkY + 1
            wsWrk.Rows(y + 1).Copy Destination:=wsWrk.Rows(wrkY)
            wsWrk.Cells(wrkY, CHECKED_X) = "after"
            wsWrk.Cells(wrkY, Range(MASTER_RNG).Column) = 0
        Else
            MsgBox "�d���L�[�́A�Q���R�[�h�̃��[���ᔽ�B�v�m�F�I�I"
            Stop
            End
        End If
    Next y

   wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j
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

End Sub
