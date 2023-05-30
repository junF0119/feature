Attribute VB_Name = "SampleCode"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function  �}�[�W�����i�W���Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : CM10_�}�[�W����
' | @Version    : v1.0.20
' | @update     : 2023/05/05
' | @written    : 2023/04/19
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |
' |
' |   ��public�ϐ�(���Y�v���W�F�N�g���̃��W���[���Ԃŋ��L)�́A�ŏ��ɌĂ΂��v���V�W���[�ɒ�`
' |     �ړ���� P_ ������
' --------------------------------------+-----------------------------------------
'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' ���ʗL���V�[�g�T�C�Y�i�f�[�^���݂̗̂̈�j
Const HV                                As String = "�X�|�d�n�e" ' HightValue�̑�֒l
Const PKEY_RNG                          As String = "AP3"   ' machingKey�̃Z���ԍ�
Const PKEY_X                            As Long = 42        ' machingKey�̗�ԍ�"AP"
Const PKEY_SHIFTJIS                     As Long = 29        ' shift-jis��Key
Const PSEIMEI_X                         As Long = 3         ' ��ƈ�̍ő�s���v���̗�ԍ�"C"(���O)
Const PDEL_X                            As Long = 38        ' �폜���̗�ԍ�"AL"
Const XMIN                              As Long = 1         ' �J�n��
Const XMAX                              As Long = PKEY_X + 1 ' �ŏI��
Const YMIN                              As Long = 4         ' �J�n�s�@��w�b�_�[��������
Const INPUTX_FROM                       As Long = 3         ' ���͊J�n����
Const INPUTX_TO                         As Long = 23        ' ���͏I������
' �O��Excel�t�@�C���̒�`�@�@Tranzaction / �AoldMaster
Private wbSrc                           As Workbook
Private wsSrc                           As Worksheet
Private srcFile                         As String           ' �O��Excel�t�@�C���̃p�X
Private srcSheet                        As String           ' �@�V�@��import����V�[�g�i�@trn / �Aold�j
Private srcMsg                          As String           ' �t�@�C����I������Ƃ��̃��b�Z�[�W
Private srcImport                       As String           ' import��̃V�[�g

' ��ƃV�[�g work �̒�`
Private wbWork                          As Workbook         ' ���̃u�b�N
Private wsWork                          As Worksheet
Private workX, workXmin, workXmax       As Long             ' i��x ��@column
Private workY, workYmin, workYmax       As Long             ' j��y �s�@row

' �@trn �V�[�g�̒�`
Private wsTrn                           As Worksheet
Private trnX, trnXmin, trnXmax          As Long             ' i��x ��@column
Private trnY, trnYmin, trnYmax          As Long             ' j��y �s�@row
Private cntTrn                          As Long             ' �ǉ��ύX�̃��R�[�h����
' �Aold �V�[�g�̒�`
Private wsOld                           As Worksheet
Private oldX, oldXmin, oldXmax          As Long             ' i��x ��@column
Private oldY, oldYmin, oldYmax          As Long             ' j��y �s�@row
Private cntOld                          As Long             ' ���}�X�^�[�̃��R�[�h����
' �Bnew �V�[�g�̒�`
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i��x ��@column
Private newY, newYmin, newYmax          As Long             ' j��y �s�@row
Private cntNew                          As Long             ' �X�V�ς݃}�X�^�[�̌���
Private cntMatch                        As Long             ' �}�b�`�������R�[�h����
' update �V�[�g�̒�` �� trn = old & trn���e �� old�̓��e
Private wsUp                            As Worksheet
Private upX, upXmin, upXmax             As Long             ' i��x ��@column
Private upY, upYmin, upYmax             As Long             ' j��y �s�@row
Private cntUp                           As Long             ' �}�X�^�̕ύX�����������R�[�h����
' archives �V�[�g�̒�` �� trn > old �폜���R�[�h
Private wsArv                           As Worksheet
Private ArvX, ArvXmin, ArvXmax          As Long             ' i��x ��@column
Private ArvY, ArvYmin, ArvYmax          As Long             ' j��y �s�@row
Private cntArv                          As Long             ' �폜���R�[�h�̌���

Public Sub �}�[�W����_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     �@Tranzaction �� �AoldMaster ����Έ�̃}�b�`���O���s���A�BnewMaster ���o��
' --------------------------------------+-----------------------------------------
    Dim i                               As Long
'
' ---Procedure Division ----------------+-----------------------------------------
'
' �I�u�W�F�N�g�ϐ��̒�`�i���ʁj
    Set wbWork = ThisWorkbook
    Set wsTrn = wbWork.Worksheets(Range("C_trnImport").Value)
    Set wsOld = wbWork.Worksheets(Range("C_oldImport").Value)
    Set wsNew = wbWork.Worksheets(Range("C_newImport").Value)
    Set wsUp = wbWork.Worksheets(Range("C_update").Value)
    Set wsArv = wbWork.Worksheets(Range("C_archive").Value)
'    Set wsWork = wbWork.Worksheets(Range("C_work").Value)  ' �V�[�g���폜���邽��importSheet_R�Œ�`
    
' �@trn �V�[�g�̃N���A
    Sheets("���j���[").Activate
    Call importClear_R(Range("C_trnImport"))
' �Aold �V�[�g�̃N���A
    Sheets("���j���[").Activate
    Call importClear_R(Range("C_oldImport"))
' �Bnew �V�[�g�̃N���A
    Sheets("���j���[").Activate
    Call importClear_R(Range("C_newImport"))
' �Cupdate �V�[�g�̃N���A
    Sheets("���j���[").Activate
    Call importClear_R(Range("C_update"))
' �DArchives �V�[�g�̃N���A
    Sheets("���j���[").Activate
    Call importClear_R(Range("C_archive"))

' 1.�@Tranzaction Excel ���u�@trn�v�V�[�g��Import
    Sheets("���j���[").Activate
    srcFile = Range("C_trnFile")                            ' Tranzaction Excel �p�X���w��
    srcSheet = Range("C_trn")                               ' �V �u�b�N�̏Z���^�V�[�g�����i�@trn / �Aold�j�w��
    srcImport = Range("C_trnImport")                        ' import��̃V�[�g��
    srcMsg = "�ǉ��E�ύX�̂���m�@Tranzaction�n Excel ��I�����Ă��������B"
    Call importSheet_R("")
    
    Range("C_trnFile") = srcFile                            ' �I�������t�@�C���̃p�X����͗��փZ�b�g
    trnYmin = YMIN
    trnXmin = XMIN
    trnYmax = workYmax                                      ' �ŏI�s�i�c�����j
    trnXmax = workXmax                                      ' �ŏI��i�������j

' 2.�AoldMaster Excel ���u�Aold�v�V�[�g��Import
    Sheets("���j���[").Activate
    srcFile = Range("C_oldMst")                             ' old Master Excel �p�X���w��
    srcSheet = Range("C_old")                               ' �V �u�b�N�̏Z���^�V�[�g�����i�@trn / �Aold�j�w��
    srcImport = Range("C_oldImport")                        ' import��̃V�[�g��
    srcMsg = "���s�Z���^�m�AoldMaster�n Excel ��I�����Ă��������B"
    Call importSheet_R("")
    
    Range("C_oldMst") = srcFile                             ' �I�������t�@�C���̃p�X����͗��փZ�b�g
    oldYmin = YMIN
    oldXmin = XMIN
    oldYmax = workYmax                                      ' �ŏI�s�i�c�����j
    oldXmax = workXmax                                      ' �ŏI��i�������j

' 3.�@Transaction �� �AoldMaster ���ukey�����v�ň�Έ�̃}�b�`���O���s��
    Call maching_R("")

' 4.���������iEOF���R�[�h�͏����j
    CloseingMsg = "trn����" & Chr(9) & "�� " & cntTrn - 1 & Chr(13) & _
                    "old����" & Chr(9) & "�� " & cntOld - 1 & Chr(13) & _
                    "new����" & Chr(9) & "�� " & cntNew & Chr(13) & _
                    "�ύX����" & Chr(9) & "�� " & cntMatch & Chr(13) & _
                    "�ύX�L��" & Chr(9) & "�� " & cntUp & Chr(13) & _
                    "�ǉ�" & Chr(9) & "�� " & cntTrn & Chr(13) & _
                    "�폜" & Chr(9) & "�� " & cntArv & Chr(13)
                    
    ' Debug.Print cntAllMsg

End Sub

Private Sub importClear_R(ByVal p_sheetName As String)
' --------------------------------------+-----------------------------------------
' | �������̃V�[�g���폜���Aimport�V�[�g���R�s�[����ƁAimport����V�[�g�̖��O��`��
' | �ꏏ�ɃR�s�[����A���O��`�̏d���Ń��W�b�N�ɕs���������̂ŁA�����V�[�g�̍폜
' | �łȂ��V�[�g�̃N���A��field�̃R�s�[�őΉ����邱�ƂɕύX����B
' |
' --------------------------------------+-----------------------------------------
    Dim wsTemp                          As Worksheet
    Dim tempX, tempXmin, tempXmax       As Long             ' i��x ��@column
    Dim tempY, tempYmin, tempYmax       As Long             ' j��y �s�@row
'
' ---Procedure Division ----------------+-----------------------------------------
'
    Sheets(p_sheetName).Activate
'  �V�[�g�Ɋ֌W�Ȃ��A�ꗥ�N���A
    Range(Cells(YMIN, XMIN), Cells(1000, 100)).Select
    Selection.ClearContents

End Sub

Private Sub importSheet_R(ByVal dummy As Variant) '
' --------------------------------------+-----------------------------------------
' |  ��Ɨp�V�[�g work ��I/O
' --------------------------------------+-----------------------------------------
' | ����1:�msrcFile�n��Excel�t�@�C���́msrcSheet�n�V�[�g�����̃u�b�N�́mwork�n�V�[�g��import����B
' |     �@���O��FExcel�V�[�g�̃t�H�[�}�b�g�͓����Ƃ���B
' | ����2: 4�s�ڈȍ~���ukey�����v�ŏ����\�[�g����B
' | ����3: �\�[�g��Ɂmwork�n�V�[�g���msrcSheet�n�֏㏑���R�s�[����B
' |
' --------------------------------------+-----------------------------------------
' work�V�[�g�̒萔�Z�b�g
    Set wsWork = wbWork.Worksheets(Range("C_work").Value)
    workYmin = YMIN                                         ' j��y �s�@row
    workXmin = XMIN                                         ' i��x ��@column
    workYmax = workYmax                                     ' �ŏI�s�i�c�����j
    workXmax = workXmax                                     ' �ŏI��i�������j

    Dim sw_FalseTrue                    As Boolean
    Dim i, y                            As Long
    Dim contentsPath                    As String
    
'
' ---Procedure Division ----------------+-----------------------------------------

' 1.import Excel �t�@�C���̐ݒ�@/�@�@Tranzaction  �AoldMaster
' --------------------------------------+-----------------------------------------
' import �V�[�g�̍폜
    
    If IsExistSheet("work") Then
        wbWork.Worksheets("work").Delete                    ' �ȑO�̃V�[�g���폜
    End If
' �t�@�C���w��̗L���`�F�b�N�@/�@�w�肵���t�@�C�����Ȃ�������A�G�N�X�v���[���Ŏw�肳����
    sw_FalseTrue = False
    If srcFile <> "" Then
        If IsExistFileDir(srcFile) Then
            sw_FalseTrue = True
        Else
            sw_FalseTrue = False
        End If
    End If
'�@[�t�@�C�����J��]�_�C�A���O�{�b�N�X�őΏ�Excel��I�����܂�
    If sw_FalseTrue = False Then
        If Range("C_childPath").Value = "" Then
            contentsPath = PathName
        Else
            contentsPath = SubSysPath & "\" & Range("C_childPath").Value
        End If
        ChDir contentsPath                                    ' �v���O�����̂���t�H���_���w��
        srcFile = Application.GetOpenFilename("Excel�t�@�C��,*.xl*", , srcMsg)
        sw_FalseTrue = True
    End If
' �O��Excel�t�@�C�����J���Aimport�V�[�g����ƃV�[�g work �փR�s�[
    Workbooks.Open srcFile
    Set wbSrc = ActiveWorkbook
    wbSrc.Worksheets(srcSheet).Copy after:=wbWork.Worksheets(1)
    ActiveSheet.Name = "work"
 
' �\�̑傫���𓾂�
    Set wsWork = wbWork.Worksheets("work")
    workYmax = wsWork.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row                   ' �ŏI�s�i�c�����j1��ځi"A")�Ōv��
    workXmax = wsWork.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column   ' �ŏI��i�������j   ' �w�b�_�[�s 3�s�ڂŌv��
' �\�̍ŏI�s�̌��HV�l��}��
    workYmax = workYmax + 1
    wsWork.Cells(workYmax, 1) = HV
    wsWork.Cells(workYmax, PKEY_X) = HV     ' HV �� �X�|�d�n�e�@�icf.�S�p�j
' shift-jis��key��unicode�ɕϊ�
    For y = YMIN To workYmax
        wsWork.Cells(y, PKEY_X) = StrConv(wsWork.Cells(y, PKEY_SHIFTJIS), vbUnicode)    ' Shift_JIS �� UTF-16
    Next y

' �����\�[�g�@key: ����key
    With ActiveSheet
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key:=.Range(PKEY_RNG), Order:=xlAscending
        .Sort.SetRange .Range(Cells(YMIN, XMIN), Cells(workYmax, XMAX))
        .Sort.Apply
    End With

' �ŏI�s�̍Čv�Z
    workYmax = wsWork.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row    ' �ŏI�s�i�c�����j1��ځi"A")�Ōv��

' �V�[�gwork ���@import��̃V�[�g�֏㏑���R�s�[
    wbWork.Worksheets("work").Cells.Copy wbWork.Worksheets(srcImport).Range("A1")
    
' �ۑ����Ȃ���close
    wbSrc.Close saveChanges:=False

' �I�u�W�F�N�g�ϐ��̉��
    Set wbSrc = Nothing
    Set wsSrc = Nothing
    Set wsWork = Nothing
     
End Sub

Private Sub maching_R(ByVal dummy As Variant)

'[�t�@�C���̓�������ёO��]
'�i�P�j�}�X�^�[�t�@�C���iMst�Ɨ����j
'    �@�V�X�e���ɕK�v�ȑS���ڂ̃f�[�^��L���邱�ơ
'    �Akey�̏d���͂Ȃ����ơ
'�i�Q�j�g�����U�N�V�����t�@�C��(Trn�Ɨ���)
'    �@Trn�̃��R�[�h�t�H�[�}�b�g�ͤMst�Ɠ����ł��邱�ơ
'    �AMst�ŕύX�ɂȂ������R�[�h��L���邱�ơ
'    �Bkey�ƕύX�̂��������ڂ��Œ���L���邱�ơ
'    �C�ύX�̂Ȃ����ځE���R�[�h���܂ނ��Ƃ����邱�ƁB
'�i�R�j�}�b�`���O�̔���
'����������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������
'�� ��Έ�̃}�b�`���O����
'����������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������
'��#    compare   ��  trn  ��  old  ��  new  �� update��  arv  ��   �t������
'����������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������
'��1.1  trn = old ��  (7)  ��  7    ��  (7)  ��  N/A  ��  N/A  ��  �ύX���@�S���ړ����� trn���R�s�[
'��1.2  trn = old ��  9+   ��  9    ��  9+   ��trn+old��  N/A  ��  �ύX�L�@�Ⴄ���ڂ�����˗v�ڎ��m�F
'��1.3  trn = old ��  10x  ��  10   ��  N/A  ��  N/A  �� 10x+10��  trn���폜�Ȃ̂ŁAtrn,old�Ƃ�arv�փR�s�[
'��1.4  trn = old ��  EOF  ��  EOF  ��  N/A  ��  N/A  ��  N/A  ��  EOF�@�v���O�����I��
'����������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������
'��2.1  trn < old ��  13x  ��  15   ��  N/A  ��  N/A  ��  13x  ��  trn���폜��old�ɓ���key���Ȃ������̂ŁAarv�փR�s�[
'��2.2  trn < old ��  1    ��  N/A  ��  1    ��  N/A  ��  N/A  ��  old�ɓ���key���Ȃ��̂ŁA�ǉ��Ƃ���new�փR�s�[
'��2.3  trn < old ��  16   ��  EOF  ��  16   ��  N/A  ��  N/A  ��  old=EOF�Ȃ̂ŁA�ǉ��Ƃ���new�փR�s�[
'����������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������
'��3.1  trn > old ��  N/A  ��  5    ��  5    ��  N/A  ��  N/A  ��   �ύX���@
'��3.2  trn > old ��  14   ��  15   ��  14   ��  N/A  ��  N/A  ��   �ǉ�
'��3.3  trn > old ��  EOF  ��  15   ��  15   ��  N/A  ��  N/A  ��   trn=EOF�Ȃ̂ŕύX�Ȃ��ŁAnew�փR�s�[
'����������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������
'�i4�j�}�b�`���Okey�̎�舵��
'�@�@VBA�ł́AHighValue�l( �� Allbit 1)�i�ȉ��uHV�v�j�������Ȃ��̂ŁA��֎�i�����̒ʂ�Ƃ�A�}�b�`���O���W�b�N�������I�ɂ���B
'    �@ machingKey = 1-���Ƃ��Ƃ̃L�[�@�i��j�L�[�������Łu�R�c�@���Y�v�Ƃ���ƁA�X�y�[�X���������@�����@��
'    �@�ړ���Ƃ��� 1- ��t���āA1-�R�c���Y�@�Ƃ���B
'    �A HV�l�Ƃ��āA9-EOF�@����ԍŌ�ɒǉ�����B
'
'
' --------------------------------------+-----------------------------------------
    Dim cntAllMsg                       As String
    Dim trnEof, oldEOF                  As Boolean
    Dim i                               As Long
'
' ---Procedure Division ----------------+-----------------------------------------
'
    Sheets("���j���[").Activate

' �\�̑傫���𓾂�
' �X�V�ς݂̐V�}�X�^�[�ƒǉ����R�[�h
    Set wsNew = Worksheets(Range("C_newImport").Value)
    newYmin = YMIN
    newXmin = XMIN
    newYmax = newYmin                                       ' �ŏI�s�i�c�����j
    newXmax = XMAX                                          ' �ŏI��i�������j
    cntNew = 0
    cntMatch = 0
    
' ���}�X�^�̃f�[�^�ύX�����������R�[�h�`�ڎ��`�F�b�N�p�i�V����r�j
    Set wsUp = wbWork.Worksheets(Range("C_update").Value)
    upYmin = YMIN
    upXmin = XMIN
    upYmax = upYmin                                         ' �ŏI�s�i�c�����j
    upXmax = XMAX                                           ' �ŏI��i�������j
    cntUp = 0                                               ' cntUp = cntNew - cntMatch
    
' �폜���R�[�h�`�폜�����L�ڂ��ꂽ���R�[�h
    Set wsArv = wbWork.Worksheets(Range("C_archive").Value)
    ArvYmin = YMIN
    ArvXmin = XMIN
    ArvYmax = ArvYmin                                       ' �ŏI�s�i�c�����j
    ArvXmax = XMAX                                          ' �ŏI��i�������j
    cntArv = 0

    cntTrn = 0
    cntOld = 0
    trnEof = False
    oldEOF = False
    trnY = trnYmin
    oldY = oldYmin
    newY = newYmin
    upY = upYmin
    ArvY = ArvYmin
    
    Do Until trnEof = True And oldEOF = True
    
'If wsTrn.Cells(trnY, PKEY_X) = "�P�|�J�e�i�D��" Then
'    wsTrn.Activate
''    MsgBox "trn=" & trnY
'End If
'
'If wsOld.Cells(oldY, PKEY_X) = "�P�|�J�e�i�D��" Then
'    wsOld.Activate
''    MsgBox "trn=" & trnY
'End If
        
'Debug.Print ">trn:" & trnY & "�c" & wsTrn.Cells(trnY, PKEY_X) & _
'            ">old:" & oldY & "�c" & wsOld.Cells(oldY, PKEY_X) & _
'            ">new:" & newY & "�c" & wsNew.Cells(newY, PKEY_X)
        
Debug.Print "trn:" & trnY & Chr(9) & _
            "|old:" & oldY & Chr(9) & _
            "|new:" & newY
            
        Select Case wsTrn.Cells(trnY, PKEY_X)               ' key���� /  �폜���@x=PDEL_Xcol
            Case Is = wsOld.Cells(oldY, PKEY_X)             ' trn = old �� match  trn��new�փR�s�[
                Call matchChk_R("")                     ' �ύX���e���Ȃ����`�F�b�N

            Case Is > wsOld.Cells(oldY, PKEY_X)             ' trn > old �� oldMaster�̂݁@new�ւ��̂܂܃R�s�[
                
                wsOld.Rows(oldY).Copy Destination:=wsNew.Rows(newY)
                newY = newY + 1
                oldY = oldY + 1
                cntOld = cntOld + 1
                cntNew = cntNew + 1

            Case Is < wsOld.Cells(oldY, PKEY_X)             ' trn < Old �� Transaction�̂� �ǉ����R�[�h


                wsTrn.Rows(trnY).Copy Destination:=wsNew.Rows(newY)
                trnY = trnY + 1
                newY = newY + 1
                cntTrn = cntTrn + 1
                cntNew = cntNew + 1
                    
        End Select
' EOF ����
        If trnY > trnYmax Then
            trnEof = True
        End If
        If oldY > oldYmax Then
            oldEOF = True
        End If
    Loop
    

' �V�}�X�^�[�Bnew ����폜���R�[�h���DArchive�V�[�g�ֈړ� & �g���[�����R�[�h 9-EOF ���폜
    wsNew.Activate
' �\�̑傫���𓾂�
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row            ' �ŏI�s�i�c�����j3��ځi"C")�Ōv��
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column   ' �ŏI��i�������j   ' �w�b�_�[�s 3�s�ڂŌv��
    
    ArvY = YMIN - 1
    For i = newYmin To newYmax
        If wsNew.Cells(i, PDEL_X) <> "" Then
            ArvY = ArvY + 1
            wsNew.Rows(i).Copy Destination:=wsArv.Rows(ArvY)
            wsNew.Rows(i).Select
            Selection.ClearContents
        End If
        If wsNew.Cells(i, PKEY_X) = "9-EOF" Then
            wsNew.Rows(i).Select
            Selection.ClearContents
        End If
    Next i

' �}�b�`���O�I��/�I�u�W�F�N�g�ϐ��̉��
    Set wsTrn = Nothing
    Set wsOld = Nothing
    Set wsNew = Nothing
    Set wsArv = Nothing

End Sub


'' �����\�[�g�@key: ����key
''    With ActiveSheet
''        .Sort.SortFields.Clear
''        .Sort.SortFields.Add Key:=.Range(PKEY_RNG), Order:=xlAscending
''        .Sort.SetRange .Range(Cells(YMIN, XMIN), Cells(workYmax, XMAX))
''        .Sort.Apply
''    End With
''
''    ActiveWindow.SmallScroll Down:=-15
''    Range("A3:AM657").Select
''    ActiveWorkbook.Worksheets("�Bnew").Sort.SortFields.Clear
''    ActiveWorkbook.Worksheets("�Bnew").Sort.SortFields.Add2 Key:=Range("AM4:AM657" _
''        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
''    With ActiveWorkbook.Worksheets("�Bnew").Sort
''        .SetRange Range("A3:AM657")
''        .Header = xlYes
''        .MatchCase = False
''        .Orientation = xlTopToBottom
''        .SortMethod = xlPinYin
''        .Apply
''    End With
'
'
'    ActiveSheet.Sort.SortFields.Clear
'    ActiveSheet.Sort.SortFields.Add2 Key:=Range(PKEY_RNG) _
'        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With ActiveSheet.Sort
'        .SetRange Range(Cells(YMIN, XMIN), Cells(newYmax, newXmax))
'        .Header = xlGuess
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'    Sheets("�Bnew").Activate
'If wsTrn.Cells(trnY, PKEY_X) = "1-�r�o�`����" Then
'    wsTrn.Activate
''    MsgBox "trn=" & trnY
'End If
'If oldY > 25 Then
'    wsNew.Activate
''    MsgBox "old=" & oldY
'End If
'
'If wsOld.Cells(oldY, PKEY_X) = "1-�r�r�j�i��" Then
'    wsOld.Activate
''    MsgBox "old=" & oldY
'End If
            
'Debug.Print ">trn:" & trnY & "�c" & wsTrn.Cells(trnY, PKEY_X) & _
'            ">old:" & oldY & "�c" & wsOld.Cells(oldY, PKEY_X) & _
'            ">new:" & newY & "�c" & wsNew.Cells(newY, PKEY_X)
'
'Debug.Print "<trn:" & trnY & "�c" & wsTrn.Cells(trnY, PKEY_X) & _
'            "<old:" & oldY & "�c" & wsOld.Cells(oldY, PKEY_X) & _
'            "<new:" & newY & "�c" & wsNew.Cells(newY, PKEY_X)

'Debug.Print "=trn:" & trnY & "�c" & wsTrn.Cells(trnY, PKEY_X) & _
'            "=old:" & oldY & "�c" & wsOld.Cells(oldY, PKEY_X) & _
'            "=new:" & newY & "�c" & wsNew.Cells(newY, PKEY_X)
'

Private Sub matchChk_R(ByVal p_delRec As String)
' --------------------------------------+-----------------------------------------
' |     �ύX�ӏ��̃`�F�b�N
' |     �ύX������Ƃ��́A�ڎ��p�ɇCupdate�V�[�g�֏o��
' --------------------------------------+-----------------------------------------
    Dim x                               As Long
    Dim trnRec                          As Variant
    Dim oldRec                          As Variant
'
' ---Procedure Division ----------------+-----------------------------------------
'
' ���͍��ڂ̂݌�����r
    trnRec = ""
    For x = INPUTX_FROM To INPUTX_TO
        trnRec = trnRec & wsTrn.Cells(trnY, x)              ' ������Ɛ��l�̌���
    Next x
    oldRec = ""
    For x = INPUTX_FROM To INPUTX_TO
        oldRec = oldRec & wsOld.Cells(oldY, x)
    Next x
    
 If wsTrn.Cells(trnY, PKEY_X) = "1-�r�o�`����" Then
    wsTrn.Activate
'    MsgBox "trn=" & trnY
End If
    
'�����������ړ��m���r���A��������ΕύX�Ȃ��ŁA���̂܂ܐV�}�X�^�֓o�^
    If trnRec = oldRec Then
        wsTrn.Rows(trnY).Copy Destination:=wsNew.Rows(newY)

        oldY = oldY + 1
        trnY = trnY + 1
        newY = newY + 1
        cntTrn = cntTrn + 1
        cntOld = cntOld + 1
        cntNew = cntNew + 1
        cntMatch = cntMatch + 1
        Exit Sub
    End If
    
' --------------------------------------+-----------------------------------------
' |     update�V�[�g�փ��R�[�h���R�s�[   �ڎ��p
' |     �ύX������Ƃ��́A�ڎ��p�ɇCupdate�V�[�g�֏o��
' |  �@)���R�[�h�́A�P�s�ڂ�old�A�Q�s�ڂ�trn�̏��ɃR�s�[
' |�@�A)�R�s�ڂ́Anew��trn���R�s�[
' |�@�B)old�ɂ���Atrn(��new)���󔒂̍��ڂ́Aold�̍��ڂ�new�փR�s�[
' |�@�C)�ύX�̂��������ڂɂ́A�S�s�ڂ̂��̏ꏊ�Ɂu???�v���Z�b�g
' |�@�D)new�s�̔w�i��ԁA�����͔��ɕύX
' |�@�E)�ڎ��œ��e���`�F�b�N�A�C���̓}�j���A���Ŏ��{
' |�@�F)�L����new�s���unew�v�V�[�g�փR�s�[
' |�@�G)�unew�v�V�[�g��V�K�u�b�N�Ƃ��ďo��
' |�@�H)�u�b�N���́A���������Z���^�̔ԍ���t��
' |�@�@�@�@+�A�c�c+�G
' |
' --------------------------------------+-----------------------------------------
    
    cntUp = cntUp + 1
    wsUp.Activate
    wsOld.Rows(oldY).Copy Destination:=wsUp.Rows(upY)
    wsUp.Cells(upY, XMAX + 1) = "old"
    wsTrn.Rows(trnY).Copy Destination:=wsUp.Rows(upY + 1)
    wsUp.Cells(upY + 1, XMAX + 1) = "trn"
    wsTrn.Rows(trnY).Copy Destination:=wsUp.Rows(upY + 2)
    wsUp.Cells(upY + 2, XMAX + 1) = "new"
    wsUp.Cells(upY + 3, XMAX + 1) = "???"
    
    Rows(upY + 2).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    For x = INPUTX_FROM To XMAX
        If wsUp.Cells(upY, x) <> "" And wsUp.Cells(upY + 1, x) = "" Then
            wsUp.Cells(upY + 2, x) = wsUp.Cells(upY, x)   ' old �� new
            wsUp.Cells(upY + 3, x) = "???"
        End If
    Next x
    
    upY = upY + 4
    oldY = oldY + 1
    trnY = trnY + 1
'    newY = newY + 1
    cntTrn = cntTrn + 1
    cntOld = cntOld + 1
'    cntNew = cntNew + 1
    cntUp = cntUp + 1
    
End Sub



