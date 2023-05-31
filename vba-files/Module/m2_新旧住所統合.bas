Attribute VB_Name = "m2_�V���Z������"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : �V���̏Z���^���}�[�W����
' --------------------------------------+-----------------------------------------
' | @moduleName : m2_�V���Z������
' | @Version    : v1.0.0
' | @update     : 2023/05/10
' | @written    : 2023/05/10
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |  �i�P�j�P�ƃ��R�[�h��V�K�V�[�g�i�@new�j�ֈړ�����
' |  �i�Q�j�d�����R�[�h�́A�ύX�Z���^��V�K�V�[�g�i�@new�j�ֈړ�����
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
Const PKEY_RNG                          As String = "AM3"   ' Key�̃Z���ԍ�
Const PKEY_X                            As Long = 39        ' Key�̗�ԍ�"AP"
Const PSEIMEI_X                         As Long = 3         ' ��ƈ�̍ő�s���v���̗�ԍ�"C"(���O)
Const PDEL_X                            As Long = 38        ' �폜���̗�ԍ�"AL"
Const XMIN                              As Long = 1         ' �J�n��
Const XMAX                              As Long = 42        ' �ŏI��
Const YMIN                              As Long = 4         ' �J�n�s�@��w�b�_�[��������
Const yMax                              As Long = 1999      ' �ő�s�@�悱�̃v���O�����ł������ő�s
Const INPUTX_FROM                       As Long = 3         ' ���͍��ڊJ�n��
Const INPUTX_TO                         As Long = 23        ' ���͍��ڏI����
Const CHECKED_X                         As Long = 40        ' �`�F�b�N���i���R�j

Private wb                              As Workbook         ' ���̃u�b�N
' �@oldSheet �V�[�g�̒�`
Private wsOld                           As Worksheet
Private oldX, oldXmin, oldXmax          As Long             ' i��x ��@column
Private oldY, oldYmin, oldYmax          As Long             ' j��y �s�@row
Private oldCnt                          As Long             ' �폜���R�[�h�̌���
' �AtrnSheet �V�[�g�̒�`
Private wsTrn                           As Worksheet
Private trnX, trnXmin, trnXmax          As Long             ' i��x ��@column
Private trnY, trnYmin, trnYmax          As Long             ' j��y �s�@row
Private trnCnt                          As Long             ' �ڎ����R�[�h�̌���
' �Bnew �V�[�g�̒�`
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i��x ��@column
Private newY, newYmin, newYmax          As Long             ' j��y �s�@row
Private newCnt                          As Long             ' �P�ƃ��R�[�h�̌���

Public Sub �P�ƒ��o����_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     ���R�[�h�̏�Ԃ��Ƃɂ��ꂼ��̃V�[�g�ɐU�蕪����
' --------------------------------------+-----------------------------------------
    Dim x, y                            As Long
    Dim CloseingMsg                     As String
    Dim w_rate, w_mod                   As Integer      ' �i���� / �\���Ԋu
    Dim i, iMin, iMax                   As Long         ' ���ꃌ�R�[�h�͈̔�(�� col x)
    Dim j, jMin, jMax                   As Long         ' ���ꃌ�R�[�h�͈̔�(�s row y)
    
'
' ---Procedure Division ----------------+-----------------------------------------
'
' ��������
    Call �O����_R("�Z���^�}�[�W" & Chr(13) & " �@�v���O�������J�n���܂��B")
    Call �o�[�\��("�Z���^�}�[�W�̃v���O�������J�n���܂��B")
' �I�u�W�F�N�g�ϐ��̒�`�i���ʁj
    Set wb = ThisWorkbook
    ' �\�̑傫���𓾂�
    ' ��ƃV�[�g�i���̃V�[�g�j�̏����l
    Set wsWrk = wb.Worksheets("work")
    wrkYmin = YMIN
    wrkXmin = XMIN
    wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j3��ځi"C")���O��Ōv��
    wrkXmax = wsWrk.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' �ŏI��i�������j   ' �w�b�_�[�s 3�s�ڂŌv��
    wrkCnt = wrkYmax - wrkYmin + 1
' �����\�[�g�@key: (1)���x��ID / (39)����key
    With ActiveSheet                '�ΏۃV�[�g���A�N�e�B�u�ɂ���
        .Sort.SortFields.Clear      '���ёւ��������N���A
        '����1
        .Sort.SortFields.Add _
            Key:=.Cells(1, 3), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        '����2
        .Sort.SortFields.Add _
            Key:=.Range(PKEY_RNG), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
'���ёւ������s����
        With .Sort
            .SetRange Range(Cells(YMIN - 1, XMIN), Cells(wrkYmax, XMAX))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
    
' �@new �V�[�g�̏����l
    Set wsNew = wb.Worksheets("�@new")
    newYmin = YMIN
    newXmin = XMIN
    wsNew.Activate
    wsNew.Rows(YMIN & ":" & yMax).Select                                    ' �N���A
    Selection.ClearContents
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' �ŏI��i�������j
    newCnt = 0
    newY = newYmin
' �Aarchives �V�[�g�̏����l �� �폜���R�[�h
    Set wsOld = wb.Worksheets("�Aarchives")
    oldYmin = YMIN
    oldXmin = XMIN
    wsOld.Activate
    wsOld.Rows(YMIN & ":" & yMax).Select                                    ' �N���A
    Selection.ClearContents
    oldYmax = wsOld.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j
    oldXmax = wsOld.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' �ŏI��i�������j
    oldCnt = 0
    oldY = oldYmin
' �BtrnChk �V�[�g�̏����l
    Set wsTrn = wb.Worksheets("�BtrnChk")
    trnYmin = YMIN
    trnXmin = XMIN
    trnY = trnYmin
    wsTrn.Activate
    wsTrn.Rows(YMIN & ":" & yMax).Select                                    ' �N���A
    Selection.ClearContents
    trnYmax = wsTrn.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j
    trnXmax = wsTrn.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' �ŏI��i�������j
    trnCnt = 0
    trnY = trnYmin
' �V�[�g�����̏�ԁi�W���j�ɖ߂�
    Sheets("template").Select
    Range(Cells(YMIN, XMIN), Cells(yMax + 1, XMAX)).Copy
    Sheets("�BtrnChk").Select
    Range(Cells(YMIN, XMIN), Cells(yMax + 1, XMAX)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False     ' �R�s�[��Ԃ̉���
    
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
'            wsWrk.Rows(y).Select
'            Selection.ClearContents
            newY = newY + 1
            newCnt = newCnt + 1
        Else
' ����key��skip
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


