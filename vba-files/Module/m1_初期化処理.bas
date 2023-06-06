Attribute VB_Name = "m1_����������"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : �����������i���W���[�������Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : m1_����������
' | @Version    : v1.0.0
' | @update     : 2023/06/02
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
' |        �v���O�����\��
' |            1. ��������
' |                1.1 �����V�[�g�̃N���A
' |                    importClear_R()
' |                1.2 �O���̃}�X�^�[�̃V�[�g����荞�ށc�c M-�@�V�Z���^���� / M-�AArchives
' |                    importSheet_R()
' |            2. �f�[�^�̐���������
' |                2.1  �d���`�F�b�N�c�c (53)PrimaryKey / (42)key����
' |                    keyCheck_F()
' |                        arrSet_R()
' |                        duplicateChk_F()
' |                            quickSort_R()
' --------------------------------------+-----------------------------------------
' |[�g�p����u�b�N�ƃV�[�g]
' |  �t�@�C����         �V�[�g��        �Ӗ�
' |�@M-�@�V�Z���^����    �@����           Active�Z���^+InActive�Z���^
' |�@M-�AArchives       �Aarchives       InActive�ɂȂ��Ă���R�N�ȏ�̏Z���^/�폜�Ώۂ̏Z���^
' |�@T-�B�ύX�Z���^      �B�ύX�Z���^    �@�ǉ��E�ύX�E�폜�ɂȂ����Z���^
' |       �V           ���e���K��       �{�V�X�e���̃t�H�[�}�b�g�ɕҏW�������R�[�h
' |       �V           ���e             �{�V�X�e���ƈقȂ�t�H�[�}�b�g�̏Z���^
' |�@M-�H���x���ꗗ      �Hlabel          �Z���^���O���[�v�����邽�߂̃��X�g
' |
' |
' --------------------------------------+-----------------------------------------
' |  �����K���̓���
' |     Public�ϐ�  �擪��啶��    �� pascalCase
' |     private�ϐ� �擪��������    �� camelCase
' |     �萔        �S�đ啶���A��؂蕶���́A�A���_�[�X�R�A(_) �� snake_case
' |     ����        �ړ���(p_)�����AcamelCase�ɏ�����
' --------------------------------------+-----------------------------------------
'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +

' �I�u�W�F�N�g�ϐ��̒�`
Private Wb                              As Workbook         ' ���̃u�b�N
' �@����V�[�g�̒�`
Private wsOld                           As Worksheet
Private oldX, oldXmin, oldXmax          As Long             ' i��x ��@column
Private oldY, oldYmin, oldYmax          As Long             ' j��y �s�@row
Private oldCnt                          As Long             ' �C���O���R�[�h�̌���
' �Aarchives�V�[�g�̒�`
Private wsArv                           As Worksheet
Private arvX, arvXmin, arvXmax          As Long             ' i��x ��@column
Private arvY, arvYmin, arvYmax          As Long             ' j��y �s�@row
Private arvCnt                          As Long             ' �C���O���R�[�h�̌���
' �B�ύX�Z���^�V�[�g�̒�`
Private wsTrn                           As Worksheet
Private trnX, trnXmin, trnXmax          As Long             ' i��x ��@column
Private trnY, trnYmin, trnYmax          As Long             ' j��y �s�@row
Private trnCnt                          As Long             ' �ύX���R�[�h�̌���
' �V�Z���^�V�[�g�̒�`
Private wsNew                           As Worksheet
Private newX, newXmin, newXmax          As Long             ' i��x ��@column
Private newY, newYmin, newYmax          As Long             ' j��y �s�@row
Private newCnt                          As Long             ' �C���ヌ�R�[�h�̌���
' ��ƃV�[�g work �̒�`
Private wsWrk                           As Worksheet
Private wrkX, wrkXmin, wrkXmax          As Long             ' i��x ��@column
Private wrkY, wrkYmin, wrkYmax          As Long             ' j��y �s�@row
Private wrkCnt                          As Long             ' ��ƃ��R�[�h�̌���=�C���O+�ύX
 
Public Sub m1_����������_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     ���R�[�h�̏�Ԃ��Ƃɂ��ꂼ��̃V�[�g�ɐU�蕪����
' --------------------------------------+-----------------------------------------
    Dim x, y                            As Long
    Dim w_rate, w_mod                   As Integer      ' �i���� / �\���Ԋu
    Dim i, iMin, iMax                   As Long         ' ���ꃌ�R�[�h�͈̔�(�� col x)
    Dim j, jMin, jMax                   As Long         ' ���ꃌ�R�[�h�͈̔�(�s row y)
    Dim inExcelpath                     As String

' ' �\���̂̐錾
' Type cntTbl
'     old                                 As long     ' �@����
'     arv                                 As long     ' �Aarchive
'     trn                                 As long     ' �B�ύX�Z���^
'     wrk                                 As long     ' work
'     new1                                As long     ' new�̌��냌�R�[�h
'     new2                                As long     ' new��archivw���R�[�h
' End Type
    Dim cnt                             As cntTbl
'
' ---Procedure Division ----------------+-----------------------------------------
'
' 1.1 �O�����i���ʁj
            
    OpeningMsg = "�u�V�Z���^����̍X�V�����v�v���O�������J�n���܂��B"
    StatusBarMsg = OpeningMsg
    Call �O����_R("")

' 1.2 �����ݒ菈��
    
' �I�u�W�F�N�g�ϐ��̒�`�i���ʁj
    Set Wb = ThisWorkbook
    Set wsOld = Wb.Worksheets(Range("C_oldSheet").Value)        ' �@����V�[�g
    Set wsArv = Wb.Worksheets(Range("C_arvSheet").Value)        ' �Aarchives�V�[�g
    Set wsTrn = Wb.Worksheets(Range("C_trnSheet").Value)        ' �B�ύX�Z���^�V�[�g
    Set wsNew = Wb.Worksheets(Range("C_newSheet").Value)        ' �V�Z���^�V�[�g
    Set wsWrk = Wb.Worksheets("work")                           ' ��Ɨp�V�[�g�̂��ߌŒ�
    
 ' �����V�[�g�̃N���A
    Call importClear_R(Range("C_oldSheet"))                     ' �@����V�[�g�̃N���A
    Call importClear_R(Range("C_arvSheet"))                     ' �Aarchives�V�[�g�̃N���A
    Call importClear_R(Range("C_trnSheet"))                     ' �B�ύX�Z���^�V�[�g�̃N���A
    Call importClear_R(Range("C_newSheet"))                     ' �V�Z���^�V�[�g�̃N���A
    Call importClear_R("work")                                  ' ��Ɨp�V�[�g�̃N���A

' �J�E���g���[��
    oldCnt = 0
    arvCnt = 0
    trnCnt = 0
    newCnt = 0
    wrkCnt = 0

' 1.3 �O��Excel�����荞��

' M-�@�V�Z���^�������荞�݁A�߂�l�𓾂�
    Call importSheet_R(Range("C_oldMst").Value, Range("C_oldSheet").Value, "M-�@�V�Z���^�����I�����Ă��������B", _
                       inExcelpath, oldYmax, oldXmax)
    Range("C_oldMst").Value = inExcelpath
    oldYmin = YMIN
    oldXmin = XMIN

' M-�Aarchives����荞�݁A�߂�l�𓾂�
    Call importSheet_R(Range("C_arvMst").Value, Range("C_arvSheet").Value, "M-�AArchives��I�����Ă��������B", _
                       inExcelpath, arvYmax, arvXmax)
    Range("C_arvMst").Value = inExcelpath
    arvYmin = YMIN
    arvXmin = XMIN

' T-�B�ύX�Z���^����荞�݁A�߂�l�𓾂�
    Call importSheet_R(Range("C_trnMst").Value, Range("C_trnSheet").Value, "T-�B�ύX�Z���^��I�����Ă��������B", _
                       inExcelpath, trnYmax, trnXmax)
    Range("C_trnMst").Value = inExcelpath
    trnYmin = YMIN
    trnXmin = XMIN

' 1.4 ��荞�񂾃V�[�g��(54)���ʋ敪:BA���t���� work�V�[�g�@�ɓ������A(42)key����/(54)���ʋ敪�ŏ����\�[�g����
'   (54)���ʋ敪:BA��@�@����V�[�g��1�A�Aarchives�V�[�g��2�A�B�ύX�Z���^�V�[�g��3
    j = 0
    jMin = oldYmin
    jMax = oldYmax
    wrkY = YMIN
    wrkYmin = YMIN
    wrkXmin = XMIN
' �@����V�[�g
    For j = jMin To jMax
        wsOld.Range(Cells(j, oldXmin).Address, Cells(j, oldXmax).Address).Copy
        wsWrk.Range(Cells(wrkY, wrkXmin).Address).PasteSpecial _
                                                  Paste:=xlPasteValues _
                                                , Operation:=xlNone _
                                                , SkipBlanks:=False _
                                                , Transpose:=False
                                                
        Application.CutCopyMode = False                     ' �R�s�[��Ԃ̉���
        wsWrk.Cells(wrkY, 54) = 1                           '  (54)���ʋ敪:BA��
        wrkY = wrkY + 1
        oldCnt = oldCnt + 1
    Next j
' �Aarchives�V�[�g
    jMin = arvYmin
    jMax = arvYmax
    For j = jMin To jMax
        wsArv.Range(Cells(j, arvXmin).Address, Cells(j, arvXmax).Address).Copy
        wsWrk.Range(Cells(wrkY, wrkXmin).Address).PasteSpecial _
                                                  Paste:=xlPasteValues _
                                                , Operation:=xlNone _
                                                , SkipBlanks:=False _
                                                , Transpose:=False
                                                            
        Application.CutCopyMode = False                     ' �R�s�[��Ԃ̉���
        wsWrk.Cells(wrkY, 54) = 2                           '  (54)���ʋ敪:BA��
        wrkY = wrkY + 1
        arvCnt = arvCnt + 1
    Next j

' �B�ύX�Z���^�V�[�g
    jMin = trnYmin
    jMax = trnYmax
    For j = jMin To jMax
        wsTrn.Range(Cells(j, trnXmin).Address, Cells(j, trnXmax).Address).Copy
        wsWrk.Range(Cells(wrkY, wrkXmin).Address).PasteSpecial _
                                                  Paste:=xlPasteValues _
                                                , Operation:=xlNone _
                                                , SkipBlanks:=False _
                                                , Transpose:=False
                                                            
        Application.CutCopyMode = False                     ' �R�s�[��Ԃ̉���
        wsWrk.Cells(wrkY, 54) = 3                           '  (54)���ʋ敪:BA��
        wrkY = wrkY + 1
        trnCnt = trnCnt + 1
    Next j

' �I�u�W�F�N�g�ϐ��̒�`�i���ʁj
    Sheets("work").Activate
' �\�̑傫���𓾂�
' ��ƃV�[�g�i���̃V�[�g�j�̏����l
    Set wsWrk = Wb.Worksheets("work")
    wrkYmin = YMIN
    wrkXmin = XMIN
    wrkYmax = wsWrk.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j6��ځi"F")���O��Ōv��
    wrkXmax = wsWrk.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' �ŏI��i�������j   ' �w�b�_�[�s 3�s�ڂŌv��
    wrkCnt = wrkYmax - wrkYmin + 1

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
' �V�[�g�ʂ̃��R�[�h������Public�ϐ��ɃZ�b�g
    cnt.old = oldCnt    ' �@����
    cnt.arv = arvCnt    ' �Aarchive
    cnt.trn = trnCnt    ' �B�ύX�Z���^
    cnt.wrk = wrkCnt    ' work
    cnt.new1 = newCnt   ' new�̌��냌�R�[�h
    cnt.new2 = newCnt   ' new��archives���R�[�h

End Sub

Private Sub importClear_R(ByVal p_sheetName As String)
' --------------------------------------+-----------------------------------------
' | �������̃V�[�g���폜���Aimport�V�[�g���R�s�[����ƁAimport����V�[�g�̖��O��`��
' | �ꏏ�ɃR�s�[����A���O��`�̏d���Ń��W�b�N�ɕs���������̂ŁA�����V�[�g�̍폜
' | �łȂ��V�[�g�̃N���A��field�̃R�s�[�őΉ����邱�ƂɕύX����B
' |
' --------------------------------------+-----------------------------------------
'
' ---Procedure Division ----------------+-----------------------------------------
'
    Sheets(p_sheetName).Activate
'  �V�[�g�Ɋ֌W�Ȃ��A�f�[�^����ꗥ�N���A�i�w�b�_�[��͏����j
    Range(Cells(YMIN, XMIN), Cells(yMax, XMAX + 1)).Select
    Selection.ClearContents

End Sub

Private Sub importSheet_R(ByVal p_excelFile As String, ByVal p_objSheet As String, ByVal p_openFileMsg As String, _
                          ByRef p_srcFile As String, ByRef p_yMax As Long, ByRef p_xMax As Long)
' --------------------------------------+-----------------------------------------
' | @function   : �R�s�[���̃V�[�g�����̃u�b�N�̓������O�̃V�[�g �փR�s�[
' | @moduleName : m1_����������
' | @remarks
' | �����̈Ӗ�
' | ���@���Fp_excelFile   �R�s�[����Excel File
' | ���@���Fp_objSheet    �R�s�[����V�[�g�����R�s�[��̃V�[�g��
' | ���@���Fp_openFileMsg �t�@�C���I���̃G�N�X�v���[���ɕ\�����郁�b�Z�[�W
' | �߂�l�Fp_srccFile    �R�s�[����ExcelFile�̐�΃p�X�ƃt�@�C����
' | �߂�l�Fp_yMax        �R�s�[�����V�[�g�̍ŏI�s�̈ʒu
' | �߂�l�Fp_xMax        �R�s�[�����V�[�g�̍ŏI��̈ʒu
' |
' --------------------------------------+-----------------------------------------
    Dim wbTmp                           As Workbook     ' �R�s�[���Ƃ�Excel�t�@�C���V�[�g
    Dim childPath                       As String       ' �R�s�[����Excel�̐�΃p�X
    Dim srcFile                         As String       ' �R�s�[����Excel�t�@�C�����i��΃p�X�t���j
    Dim sw_naFile                       As Boolean      ' �t�@�C���L �� True �t�@�C���� �� False
    Dim sw_naFolder                     As Boolean      ' �t�H���_�L �� True �t�H���_�� �� False
    
    Dim i, y                            As Long
    Dim absolutePath                    As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
' --------------------------------------+-----------------------------------------
' 1.import Excel �t�@�C���̓ǂݍ���
' --------------------------------------+-----------------------------------------

' �t�H���_�w��̗L���`�F�b�N�@/�@�w�肵���t�H���_���Ȃ�������A�G�N�X�v���[���Ŏw�肳����
    childPath = Range("C_childPath")
    sw_naFolder = False
    sw_naFile = False
' �p�X�w�肪����Ƃ��́A�t�H���_�̑��݂��`�F�b�N
    If childPath <> "" Then
        If IsExitsFolderDir(SubSysPath & "\" & childPath) Then
            sw_naFolder = True
        Else
            childPath = ""
            Range("C_childPath") = ""
        End If
    End If
' �t�@�C���w��̗L���`�F�b�N�@/�@�w�肵���t�@�C�����Ȃ�������A�G�N�X�v���[���Ŏw�肳����
    If p_excelFile <> "" Then
        If IsExistFileDir(p_excelFile) Then
            sw_naFile = True
            srcFile = p_excelFile
        Else
            srcFile = ""
        End If
    End If
    
'�@[�t�@�C�����J��]�_�C�A���O�{�b�N�X�őΏ�Excel��I�����܂�
    If sw_naFile = False Then
        If childPath = "" Then
            absolutePath = PathName
        Else
            absolutePath = SubSysPath & "\" & childPath
        End If
        ChDir absolutePath                                    ' �v���O�����̂���t�H���_���w��
        srcFile = Application.GetOpenFilename("Excel�t�@�C��,*.xl*", , p_openFileMsg)
        sw_naFile = True
    End If
    
' �O��Excel�t�@�C�����J���Aimport�V�[�g����ƃV�[�g work �փR�s�[
    Workbooks.Open srcFile
    Set wbTmp = ActiveWorkbook
'    ActiveSheet.ShowAllData         ' �t�B���^����
    wbTmp.Sheets(p_objSheet).Range(Cells(YMIN, XMIN).Address, Cells(yMax, XMAX).Address).Copy
    Wb.Sheets(p_objSheet).Range(Cells(YMIN, XMIN).Address).PasteSpecial _
                                                           Paste:=xlPasteValues _
                                                         , Operation:=xlNone _
                                                         , SkipBlanks:=False _
                                                         , Transpose:=False
    
   
    
    Application.CutCopyMode = False                         ' �R�s�[��Ԃ̉���
    wbTmp.Close saveChanges:=False                          ' �ۑ����Ȃ���close
' �\�̑傫���𓾂�
    p_srcFile = srcFile
    p_yMax = Wb.Worksheets(p_objSheet).Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row            ' �ŏI�s�i�c�����j(6)���O�i"F")�Ōv��
    p_xMax = Wb.Worksheets(p_objSheet).Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column   ' �ŏI��i�������j   ' �w�b�_�[�s 3�s�ڂŌv��
' �I�u�W�F�N�g�ϐ��̉��
    Set wbTmp = Nothing

End Sub

