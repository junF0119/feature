Attribute VB_Name = "CM00_���ʊ֐�"
Option Explicit
' --------------------------------------------------------------------------------
' | @function   : �W���I�ɗ��p����郂�W���[����W���֐��Ƃ��ē���
' --------------------------------------+-----------------------------------------
' | @moduleName : CM00_���ʊ֐�"
' | @Version    : v3.0.2
' | @update     : 2023/05/31
' | @written    : 2023/04/21
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |
' | @Program naming rule
' |  ProgramID: xx.99.99-xxxxxxxxx-vx.y.z-yyyymmdd.suffix   x.y.z��version number
' |              |--xx(SystemSymbl)
' |              |----99(Subsystem#)
' |              |------.(priod��separator)
' |              |-------99(Module#)
' |
' | @����VBA����
' |      1.proc_H2_�O����
' |      2.proc_H9_�㏈��
' |      3.util02_�W�����W���[���ꊇExport
' |      4.util03_�o�b�N�A�b�v����
' |      5.util08_IsExitsSheet
' |      6.util09_IsMsgbox
' |      7.util10_IsMsgPush
' |      8.util15_IsExitsFolderDir
' |      9.util16_IsExitsFileDir
' |     10.util17_selectFile
' |     11.util18_selectFolder
' |     12.util19_getLatestFile
' |     13.util20_getFolderPath_F
' |     14.util23_�t�@�C���������
' |     15.util24_R1C1�`��2A1�`���ϊ�_F
' |     16.util29_�u�b�N�v���p�e�B�擾
' |     17.util30_getImportSheet
' |     18.util31_get���ʕϐ�_F
' |     19.util00_�X�e�[�^�X�o�[�\���E����
' |
' | @���ݒ�
' |�@�@�@ �t�@�C���^�u���I�v�V�������Z�L�����e�B�Z���^�[���Z�L�����e�B�Z���^�[��
' |�@     �ݒ�{�^�����}�N���̐ݒ聨
' |       VBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������(V)�v���`�F�b�N
' |    �A VBA��ʂ̃c�[�����j���[���Q�Ɛݒ聨 ���̃��C�u�����t�@�C�����`�F�b�N
' |
' |     ��Visual Basic For Applications
' |     ��Microsoft Excel 15.0 Object Library
' |     ��OLE Automatiom
' |     ��Microsoft Scripting Runtime
' |     ��Microsoft Scripting Libary
' |     ��Microsoft Visual Basic for Application Extensibilly 5.3
' |     ��Windows Script Host Object Model
' |
' | @�f�B���N�g���\���} �� �t�H���_�\���}
' |     1       1       1       1       1
' |  root/      RootPath (�V�X�e���t�H���_�̐e�t�H���_�@���[�g����V�X�e���t�H���_�̑O�܂ł̃t���p�X�j
' |     ��
' |     �� �V�X�e�����@ SysPath�i�t���p�X�j�@SysName�i�t�H���_���j
' |     ��      ��
' |     ��      �� �T�u�V�X�e����   SubSysPath�i�t���p�X�j�@SubSysName�i�t�H���_���j �i�e�f�B���N�g���ߐe�t�H���_�@../�@ParentFolder)
' |     ��      ��      ��
' |            ��      �� !Repository(�J����)
' |                   ��
' |   �R�s�y�p�L���W   �� 1.�^�p�̎����
' |-----------------  ��
' |     �� �� �� ��       �� 2.�}�X�^�[�Q
' |     �� ��           ��
' |     �� �� �� ��       �� 3.���s�v���O����  folderPath�i�t���p�X�j folderName�i�t�H���_���j (�J�����g�f�B���N�g���ߊ�_�t�H���_�@�@./�@)
' |     �� �� �� ��       ��      ��
' |     �� �� �� ��       ��      �� tmXX.YY.ZZ-������-vX.Y.Z-yyyymmdd.xlms
' |                   ��      �� tmXX.YY.ZZ-������-vX.Y.Z-yyyymmdd.xlms
' |                   ��
' |                   �� 7.�Ǘ��c�[��
' |                   ��
' |                   �� 8.�R���e���c
' |                   ��
' |                   �� 9.�h�L�������g
' |                   ��
' --------------------------------------+----------------------------------------
' |  �����K���̓���
' |     Public�ϐ�  �擪��啶��    �� pascalCase
' |     private�ϐ� �擪��������    �� camelCase
' |     �萔        �S�đ啶���A��؂蕶���́A�A���_�[�X�R�A(_) �� snake_case
' |     ����        �ړ���(p_)�����AcamelCase�ɏ�����
' --------------------------------------+-----------------------------------------
'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +

'
'   ��public�ϐ�(���Y�v���W�F�N�g���̃��W���[���Ԃŋ��L)�́A�ŏ��ɌĂ΂��v���V�W���[�ɒ�`
'
Public BackupFile                       As String       ' ���s�O�t�@�C���̕ۑ��p�t�H���_�̃t���p�X
Public fullPath                         As String       ' ���sExcel�̃t���p�X+�t�@�C���� �� Thisworkbook
Public PathName                         As String       ' ���sExcel�̃t���p�X
Public FileName                         As String       ' ���sExcel�̃t�@�C����
' �f�B���N�g���\���̃p�X�Ɩ��O
Public RootPath                         As String       ' �V�X�e���t�H���_�̐e�f�B���N�g���̃��[�g�p�X
Public SysPath                          As String       ' �V�X�e���t�H���_�܂ł̃t���p�X
Public SysName                          As String       ' �V�X�e���t�H���_�̖��O
Public SubSysPath                       As String       ' �T�u�V�X�e���t�H���_�܂ł̃t���p�X
Public SubSysName                       As String       ' �T�u�V�X�e���t�H���_�̖��O
' ���s�v���O�����̏��
Public SysSymbol                        As String       ' �V�X�e���V���{��
Public PrgName                          As String       ' ���sExcel�̃v���O������
Public Version                          As String       ' vx.x.x
Public Update                           As String       ' yyyymmdd
' �v���O�������s���̓������
Public nowY                             As Integer      ' �����̔N�i�����j
Public nowM                             As Integer      ' �����̌��i�����j
Public nowD                             As Integer      ' �����̓��i�����j
Public TimeStart                        As Variant      ' �v���O�����J�n�̓��t�Ǝ���
Public TimeStop                         As Variant      ' �v���O�����I���̓��t�Ǝ���
Public TimeLap                          As Variant      ' �v���O�������s�̏��v����
Public NendoYYYY                        As Integer      ' ���N�x�i����j
' �v���O��������
Public Mode                             As String       ' ���샂�[�h insert / inquiry / modify / erase / clear / end
                                                        ' �}�N�����̃{�^���ԍ��w����@�@������.xlsm'!'������ n'�@<== �{�^�� n
Public MenuNum                          As Integer      ' �V�[�g�{�^���̏����ԍ�
Public NumCnt                           As Long         ' ��������
Public ErrCnt                           As Long         ' �G���[����
' �v���O�����J�n�E�I�����b�Z�[�W
Public OpeningMsg                       As String       ' �v���O�����J�n���b�Z�[�W
Public CloseingMsg                      As String       ' �v���O��������I�����b�Z�[�W
Public StatusBarMsg                     As String       ' Excel�X�e�[�^�X�o�[�i�ŉ��Ӂj�ɕ\�����郁�b�Z�[�W
'
'

Sub get���ʕϐ�_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' | @function   : �W���u�̏��������i�W���Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : util31_get���ʕϐ�
' | @Version    : v1.0.0
' | @update     : 2023/05/04
' | @written    : 2023/05/04
' | @remarks
' |     ���[�U��`��Public�ϐ��̏����l�ݒ�
' |
' --------------------------------------+-----------------------------------------
    Dim nowYMD                          As Date
    Dim rc                              As Long
    Dim temp, temp1, temp2, temp3, temp4 As Variant
    Dim x                               As Long
'
' ---Procedure Division ----------------+-----------------------------------------
'
   On Error Resume Next  ' �G���[�ł����̍s���珈���𑱍s����
       
    Calculate                                           ' [DockInfo]�@�ŐV��ԂɍX�V�@�p�X�𐳂������邽��

    Application.DisplayAlerts = False                   ' waring���~�߂�
    Application.ScreenUpdating = True                   ' �������̉�ʂ����� / ���͓��e���V�[�g�Ŋm�F����ɂ́A���A���ōX�V���邽�߁@false
    Application.Calculation = xlCalculationManual       ' �蓮�v�Z�ɕύX
    
    Mode = ""
    
' ����������s���Ԃ̃J�E���g���J�n
    TimeStart = Time                                  ' �������Ԍv��
    TimeStop = TimeStart
    TimeLap = TimeStop - TimeStart
    nowYMD = Now()               ' �����̓��t����N�A���A���A�������𕪊�
    nowY = Year(nowYMD)
    nowM = month(nowYMD)
    nowD = Day(nowYMD)

' �t�@�C���������F .\.\SysID.xx.xx_programName-vX.Y.Z_yyyymmdd.sufix
    fullPath = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
    PathName = ActiveWorkbook.Path
    FileName = ActiveWorkbook.Name
    temp = Split(PathName, "\")
    RootPath = temp(0)                ' �V�X�e���t�H���_�̐e�f�B���N�g��
    For x = 1 To UBound(temp) - 3
        RootPath = RootPath & "\" & temp(x)
    Next x

' �T�u�V�X�e���t�H���_�i���s���W���[���̐e�f�B���N�g��
    temp = Split(PathName, "\")
    SubSysPath = temp(0)                ' �V�X�e���t�H���_�̐e�f�B���N�g��
    For x = 1 To UBound(temp) - 1
        SubSysPath = SubSysPath & "\" & temp(x)
    Next x

' SystemSymbol�̒��o              tmXX.YY.ZZ-ProgramName-vn.m.l-yyyymmdd
    temp = Split(FileName, "-")
    temp1 = Split(temp(0), ".")
    SysSymbol = temp1(0)
'    SysSymbol = Left(SysSymbol, 4)
' �v���O�������Aversion�A�X�V���̒��o
    PrgName = temp(1)
    Version = temp(2)
    Update = temp(3)

End Sub

Sub �O����_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' | @function   : �W���u�̏��������i�W���Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : proc_H2_�O����
' | @Version    : v3.2.0
' | @update     : 2023/05/17
' | @written    : 2020/12/29
' | @remarks
' |     1.�v���O�����ŋ��ʂ̏����l�A���ϐ���ݒ蓙�̎��O����
' |         (1) ���݂̃p�X���擾
' |         (2) ���Y�v���O�����̃o�b�N�A�b�v�������o��
' |         (3) �X�e�[�^�X�o�[�̕\��
' |
' --------------------------------------+-----------------------------------------
'
' ---Procedure Division ----------------+-----------------------------------------
'
   On Error Resume Next  ' �G���[�ł����̍s���珈���𑱍s����
   
    Call get���ʕϐ�_R("")
    
    IsMsgbox (OpeningMsg)
    
    Call putStatusBar(StatusBarMsg)
    
    Call �o�b�N�A�b�v����("")

End Sub


Sub �㏈��_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' | @function   :�I�������̗v��i�W���Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : proc_H9_�㏈��
' | @Version    : v3.0.0
' | @update     : 2201/01/02
' | @written    : 2020/12/29
' | @remarks
' |     �I������
' --------------------------------------+-----------------------------------------
    Dim msgText                             As String
'
' ---Procedure Division ----------------+-----------------------------------------
'

    Application.ScreenUpdating = True '��������ʂ�\������
    TimeStop = Time
    TimeLap = TimeStop - TimeStart

    If CloseingMsg <> "" Then
        msgText = CloseingMsg & Chr(13) & Chr(13) _
            & "-------------------------------------------------------------" & Chr(13) _
            & "       �mJob summary�n" & Chr(13) _
            & "-------------------------------------------------------------" & Chr(13) _
            & "|Backup(Before)" & Chr(9) & "�� " & BackupFile & Chr(13) _
            & "|��������" & Chr(9) & "�� " & Minute(TimeLap) & "��" & Second(TimeLap) & "�b" & Chr(13) _
            & "|��������" & Chr(9) & "�� " & NumCnt & " ��"
      

        IsMsgPush (msgText)
    End If
    
    Calculate                   ' ��ʂ��ŐV��ԂɍX�V
    Application.StatusBar = False
    Application.DisplayAlerts = True
    ActiveWorkbook.Save

End Sub

Public Sub �W�����W���[���ꊇExport()
' --------------------------------------+-----------------------------------------
' | @function   : �W�����W���[�������ꊇ����Export����}�N��
' --------------------------------------+-----------------------------------------
' | @moduleName : util02_�W�����W���[���ꊇExport
' | @Version    : v3.0.0  ���A���łɑΉ� �@�@��؂蕶���F- �ɓ���
' | @update     : 2020/04/25
' | @written    : 2019/08/01
' | @remarks
' |      https://vbabeginner.net/�W�����W���[�����̈ꊇ�G�N�X�|�[�g/
' --------------------------------------+-----------------------------------------
   
    Dim fso                             As Object
    Dim prgFullPath                     As String           '// ��΃p�X�̃t�@�C���� �� ��΃p�X+�t�@�C����
    Dim prgPathName                     As String           '// ��΃p�X
    Dim prgFileName                     As String           '// �t�@�C���� fileName Form: n.m.l xxxxxxxxx-v?_yyyymmdd.xlsx   ?��version number
    Dim prgDocNumber                    As String
    Dim prgDocName                      As String
    Dim prgVersion                      As String
    Dim prgUpdate                       As String

    Dim myDir                           As String
    Dim myBook                          As String
    Dim mySheet                         As String
    Dim i                               As Long
    Dim iMax                            As Long
    Dim x1, x2, x3, x4, x5              As Long
    Dim x��, x��, xv                    As Long
   
   
    Dim nowYMD                          As Date
    Dim nowY                            As Integer
    Dim nowM                            As Integer
    Dim nowD                            As Integer
'                                       +
    Dim module                          As VBComponent      '// ���W���[��
    Dim moduleList                      As VBComponents     '// VBA�v���W�F�N�g�̑S���W���[��
    Dim extension                       As String           '// ���W���[���̊g���q
    Dim sPath                           As String           '// �����Ώۃu�b�N�̃p�X
    Dim sFilePath                       As String           '// �G�N�X�|�[�g�t�@�C���p�X
    Dim TargetBook                      As Object           '// �����Ώۃu�b�N�I�u�W�F�N�g
    Dim saveDir                         As String           '// VBA�̕ۑ��t�H���_

    Dim fullPath                        As String
    Dim sysSybl                         As String           '// �V�X�e���V���{���@\!Program(????)
    Dim l, lMax                         As Long
'
' ---Procedure Division ----------------+-----------------------------------------
'
  
    On Error Resume Next  ' �G���[���������Ă��A���̍s���珈���𑱍s����
' --------------------------------------+-----------------------------------------
' |     �������u����Ă���t�H���_���i�J�����g�f�B���N�g�����j�A
' |     �����̃t�@�C�����E�V�[�g�����擾����
' --------------------------------------+-----------------------------------------

    Call get���ʕϐ�_R("")
'    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fso = New FileSystemObject          ' �C���X�^���X��

    myDir = ActiveWorkbook.Path
    myBook = ActiveWorkbook.Name
    mySheet = ActiveSheet.Name

    prgPathName = myDir
    prgFullPath = myDir & "\" & myBook
    prgFileName = myBook
' ProgramID: xx99.99-xxxxxxxxx-vx.y.z-yyyymmdd.suffix   x.y.z��version number �𕪊�
'             |--xx(SystemSymbl)
'             |----99(Subsystem#)
'             |------.(priod��separator)
'             |-------99(Module#)
'
    iMax = Len(prgFileName)
'
    x1 = InStr(1, prgFileName, ".")      ' SystemSymbol�E�E���p . �̈ʒu
    x2 = InStr(x1, prgFileName, "-")     ' Module# �E�E�E�E���p �X�y�[�X �̈ʒu

                                        ' docName �E�E�E�E���p -v �̈ʒu
                                        ' ���Ť���Ťv�� (RC��)
    x�� = InStr(x2, prgFileName, "-��")
    x�� = InStr(x2, prgFileName, "-��")
    xv = InStr(x2, prgFileName, "-v")

    If x�� <> 0 Then
        x3 = x��
    ElseIf x�� <> 0 Then
        x3 = x��
    ElseIf xv <> 0 Then
        x3 = xv
    Else
        x3 = 0
    End If

    x4 = InStr(x3 + 1, prgFileName, "-")   ' version �E�E�E�E���p - �̈ʒu
    x5 = InStrRev(prgFileName, ".")      ' update�E�E�E�E�E���p - �̈ʒu

    prgDocNumber = Left(prgFileName, x2 - 1)  ' �߁@Module#
    prgDocName = Mid(prgFileName, x2 + 1, x3 - x2 - 1)
    prgVersion = Mid(prgFileName, x3 + 2, x4 - x3 - 2)
    prgUpdate = Mid(prgFileName, x4 + 1, x5 - x4 - 1)

' --------------------------------------------------------------------------------
' |     �����̓��t����N�A���A���A�������𕪊�
' --------------------------------------------------------------------------------

    nowYMD = Now()               ' �����̓��t����N�A���A���A�������𕪊�

    nowY = Year(nowYMD)
    nowM = month(nowYMD)
    nowD = Day(nowYMD)
   
'   Call R_DocInfoGet           ' �t�@�C���̃p�X���iDocInfo)����ProgramID���𓾂�

   '// �u�b�N���J����Ă��Ȃ��ꍇ�͌l�p�}�N���u�b�N�ipersonal.xlsb�j��ΏۂƂ���
    If (Workbooks.Count = 1) Then
        Set TargetBook = ThisWorkbook
   '// �u�b�N���J����Ă���ꍇ�͕\�����Ă���u�b�N��ΏۂƂ���
    Else
        Set TargetBook = ActiveWorkbook
    End If

    sPath = TargetBook.Path & "\!VBAmodules(" & prgDocNumber & ")-v" & prgVersion

    If Dir(sPath, vbDirectory) = "" Then  ' �t�H���_���Ȃ��Ƃ��́A�쐬����
        MkDir sPath
    End If


   '// �����Ώۃu�b�N�̃��W���[���ꗗ���擾
    Set moduleList = TargetBook.VBProject.VBComponents

   '// VBA�v���W�F�N�g�Ɋ܂܂��S�Ẵ��W���[�������[�v
    For Each module In moduleList
       '// �N���X
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
       '// �t�H�[��
        ElseIf (module.Type = vbext_ct_MSForm) Then
           '// .frx���ꏏ�ɃG�N�X�|�[�g�����
            extension = "frm"
       '// �W�����W���[��
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
       '// ���̑�
        Else
           '// �G�N�X�|�[�g�ΏۊO�̂��ߎ����[�v��
            GoTo CONTINUE
        End If

       '// �G�N�X�|�[�g���{
        sFilePath = sPath & "\" & module.Name & "-v" & prgVersion & "-" & prgUpdate & "." & extension
        Call module.Export(sFilePath)

       '// �o�͐�m�F�p���O�o��
       'Debug.Print sFilePath
CONTINUE:
    Next

    MsgBox "VBA�W�����W���[��" & "-v" & prgVersion & "-" & prgUpdate & "�@�Q���ꊇExport����"
End Sub

Private Sub R_DocInfoGet()
' --------------------------------------+-----------------------------------------
' |     �������u����Ă���t�H���_���i�J�����g�f�B���N�g�����j�A
' |     �����̃t�@�C�����E�V�[�g�����擾����
' --------------------------------------+-----------------------------------------
'                                       +
    Dim myDir                           As String
    Dim myBook                          As String
    Dim mySheet                         As String
    Dim i                               As Long
    Dim iMax                            As Long
    Dim x1, x2, x3, x4, x5              As Long
    Dim x��, x��, xv                    As Long

' ---Procedure Division ---------------+------------------------------------------
'
   On Error Resume Next  ' �G���[���������Ă��A���̍s���珈���𑱍s����
' --------------------------------------+-----------------------------------------
' |     �������u����Ă���t�H���_���i�J�����g�f�B���N�g�����j�A
' |     �����̃t�@�C�����E�V�[�g�����擾����
' --------------------------------------+-----------------------------------------

'    Set fso = CreateObject("Scripting.FileSystemObject")
   Set fso = New FileSystemObject          ' �C���X�^���X��

   myDir = ActiveWorkbook.Path
   myBook = ActiveWorkbook.Name
   mySheet = ActiveSheet.Name

   prgPathName = myDir
   prgFullPath = myDir & "\" & myBook
   prgFileName = myBook
' ProgramID: xx99.99-xxxxxxxxx-vx.y.z-yyyymmdd.suffix   x.y.z��version number �𕪊�
'             |--xx(SystemSymbl)
'             |----99(Subsystem#)
'             |------.(priod��separator)
'             |-------99(Module#)
'
   iMax = Len(prgFileName)
'
   x1 = InStr(1, prgFileName, ".")      ' SystemSymbol�E�E���p . �̈ʒu
   x2 = InStr(x1, prgFileName, "-")     ' Module# �E�E�E�E���p �X�y�[�X �̈ʒu

                                        ' docName �E�E�E�E���p -v �̈ʒu
                                        ' ���Ť���Ťv�� (RC��)
   x�� = InStr(x2, prgFileName, "-��")
   x�� = InStr(x2, prgFileName, "-��")
   xv = InStr(x2, prgFileName, "-v")

   If x�� <> 0 Then
       x3 = x��
   ElseIf x�� <> 0 Then
       x3 = x��
   ElseIf xv <> 0 Then
       x3 = xv
   Else
       x3 = 0
   End If

   x4 = InStr(x3 + 1, prgFileName, "-")   ' version �E�E�E�E���p - �̈ʒu
   x5 = InStrRev(prgFileName, ".")      ' update�E�E�E�E�E���p - �̈ʒu

   prgDocNumber = Left(prgFileName, x2 - 1)  ' �߁@Module#
   prgDocName = Mid(prgFileName, x2 + 1, x3 - x2 - 1)
   prgVersion = Mid(prgFileName, x3 + 2, x4 - x3 - 2)
   prgUpdate = Mid(prgFileName, x4 + 1, x5 - x4 - 1)

' --------------------------------------------------------------------------------
' |     �����̓��t����N�A���A���A�������𕪊�
' --------------------------------------------------------------------------------

   nowYMD = Now()               ' �����̓��t����N�A���A���A�������𕪊�

   nowY = Year(nowYMD)
   nowM = month(nowYMD)
   nowD = Day(nowYMD)

End Sub

Public Sub �o�b�N�A�b�v����(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' | @function   : ���YExcel�̏C���O�o�b�N�A�b�v�����
' --------------------------------------+-----------------------------------------
' | @moduleName : util03_�o�b�N�A�b�v����
' | @Version    : v2.1.0
' | @update     : 2020/04/11
' | @written    : 2019/08/01
' | @remarks
' |     ���YExcel�̏C���O�o�b�N�A�b�v�����
' --------------------------------------+-----------------------------------------
    Dim saveDir                         As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set fso = New FileSystemObject          ' �C���X�^���X��

    saveDir = PathName & "\" & SysSymbol & "-backup"

    If Dir(saveDir, vbDirectory) = "" Then  ' �t�H���_���Ȃ��Ƃ��́A�쐬����
        MkDir saveDir
    End If

    BackupFile = "backup-" & Format(Now(), "yyyy-mm-dd_hhmmss") & "_" & FileName
'�o�b�N�A�b�v��������t�@�C�����g�����߂ɂ́ASaveCopyAs ���g��
    ActiveWorkbook.SaveCopyAs saveDir & "\" & BackupFile            ' �o�b�N�A�b�v��������t�@�C�����g�����߂ɂ́A�@SaveCopyAs ���g��

End Sub

Public Function IsExistSheet(p_sheetName) As Boolean
' --------------------------------------+-----------------------------------------
' | @function   : �V�[�g�̑��݂��`�F�b�N�i�W���Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : util08_IsExitsSheet
' | @Version    : v1.0.0
' | @update     : 2020/04/25
' | @written    : 2020/04/25
' | @remarks
' |     ���C������
' |�@�����œn���ꂽ�V�[�g�̑��݂��`�F�b�N����
' |�@���݂���@�� thru
' |�@���݂��Ȃ��� false
' |
' --------------------------------------+-----------------------------------------
    Dim objWorksheet                    As Worksheet
'
' ---Procedure Division ----------------+-----------------------------------------
'

'
     On Error GoTo NotExists
    
    Set objWorksheet = ThisWorkbook.Sheets(p_sheetName)
    
    IsExistSheet = True
    
    Exit Function
    
NotExists:
    IsExistSheet = False
End Function

Public Function IsMsgbox(p_helloMsg) As Boolean
' --------------------------------------+-----------------------------------------
' | @function   : ���b�Z�[�W�\���̉����ɂ��Ή��t���@msgbox �i�W���Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : util09_IsMsgbox
' | @Version    : v2.0.0
' | @update     : 2020/12/25
' | @written    : 2020/04/06
' | @remarks
' |             IsMsgbox("���b�Z�[�W")
' |
' |�@�����̃��b�Z�[�W��\�����A�����ɂ�菈���𕪂���
' |�@yes�@�� thru
' |�@no   �� false
' |
' --------------------------------------+-----------------------------------------
    Dim rc                              As VbMsgBoxResult       ' �񋓑�
'
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'      �W���{�^���̎w��i )
'  ---------------------------------------------------
'     �萔�@�@�@�@�@  �l  �@Enter �L�[�Ŏ��s����L�[
'  ------------------ ---- ---------------------------
'   vbDefaultButton1    0   ��1�{�^��
'   vbDefaultButton2   256  ��2�{�^��
'   vbDefaultButton3   512  ��3�{�^��
'   vbDefaultButton4   768  ��4�{�^��
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'       �W���A�C�R���̎w��
'  --------------- ------------  ------------- ---- -------------------------------
'      ���            �摜         �萔        �l    ����
'  --------------- ------------  ------------- ---- -------------------------------
'   �G���[�A�C�R�� vba_msgbox23  vbCritical     16  �G���[���b�Z�[�W�Ŏg�p���܂��B
'   �^�╄�A�C�R�� vba_msgbox26  vbQuestion     32  �w���v�E�B���h�E���J�����b�Z�[�W�{�b�N�X�Ŏg�p���܂��B
'   �x���A�C�R��   vba_msgbox24  vbExclamation  48  �x���̃��b�Z�[�W�Ŏg�p���܂��B
'   ���A�C�R��   vba_msgbox25  vbInformation  64  �厖�ȏ���`���郁�b�Z�[�W�Ŏg�p���܂��B

'
' ---Procedure Division ----------------+-----------------------------------------
'
'
        If p_helloMsg = "" Then
            p_helloMsg = "�������J�n���܂��B(by IsMsgbox)" & Chr(13) & "�L�����Z���ŋ����I�����܂��B"
        End If
        rc = MsgBox(p_helloMsg & vbNewLine, vbOKCancel + vbQuestion + vbDefaultButton1)
        If rc = vbCancel Then
            MsgBox "�����������I�����܂��B"
            IsMsgbox = False
            End
        Else
            IsMsgbox = True
        End If

    
End Function

Public Sub IsMsgPush(ByVal Msg As String)
' --------------------------------------+-----------------------------------------
' | @function   : �����ŕ��郁�b�Z�[�W����}�N��
' --------------------------------------+-----------------------------------------
' | @moduleName : util10_IsMsgPush
' | @Version    : v1.1.0
' | @update     : 2020/04/26
' | @written    : 2019/09/01
' | @remarks
' |     �����ŕ��郁�b�Z�[�W�@�J�� �� VisualBasic �� �c�[�� �� �Q�Ɛݒ� �� Windows Script Host Object Model��on(���j
' --------------------------------------+-----------------------------------------
    Dim wsh                             As Object  'IWshRuntimeLibrary.WshShell
'
' ---Procedure Division ----------------+-----------------------------------------
'
    
    Set wsh = CreateObject("Wscript.Shell")

    wsh.Popup _
        Text:=Msg & vbNewLine & "�`���̃��b�Z�[�W�́A�P�O�b��Ɏ����I�ɏ����܂��`", _
        SecondsToWait:=10, _
        Title:="", _
        Type:=vbOKOnly + vbInformation

    Set wsh = Nothing

End Sub

Public Function IsExitsFolderDir(p_sFolderPath) As Boolean
' --------------------------------------+-----------------------------------------
' | @function   : �t�H���_�̑��݂��`�F�b�N�i�W���Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : util15_IsExitsFolderDir
' | @Version    : v1.2.0
' | @update     : 2023/05/31
' | @written    : 2020/12/09
' | @remarks
' |�@�����œn���ꂽ�t�H���_�̑��݂��`�F�b�N����
' |�@���݂���@�� thru
' |�@���݂��Ȃ��� false
' |
' --------------------------------------+-----------------------------------------
    Dim result
'
' ---Procedure Division ----------------+-----------------------------------------
'
    result = Dir(p_sFolderPath, vbDirectory)
    If (result = "") Then
        IsExitsFolderDir = False    ' �t�H���_�����݂��Ȃ�
    Else
        IsExitsFolderDir = True     ' �t�H���_�����݂���
    End If
    
End Function

Public Function IsExistFileDir(p_sFilePath) As Boolean
' --------------------------------------+-----------------------------------------
' | @function   :�t�@�C���̑��݂��`�F�b�N�i�W���Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : util16_IsExitsFileDir
' | @Version    : v1.0.1
' | @update     : 2023/05/18
' | @written    : 2020/03/21
' | @remarks
' |�@�����œn���ꂽ�t�@�C���̑��݂��`�F�b�N����
' |�@���݂���@�� thru
' |�@���݂��Ȃ��� false
' |
' --------------------------------------+-----------------------------------------
    Dim result
'
' ---Procedure Division ----------------+-----------------------------------------
'
'
    result = Dir(p_sFilePath)
    If (result = "") Then
        IsExistFileDir = False
    Else
        IsExistFileDir = True
    End If
    
End Function

Public Function selectFile(ByVal P_title As String)
' --------------------------------------+-----------------------------------------
' | @function   : [�t�@�C�����J��]�_�C�A���O�{�b�N�X�Ńt�@�C����I�����܂�
' --------------------------------------+-----------------------------------------
' | @moduleName : util17_selectFile
' | @Version    : v1.0.0
' | @update     : 2020/01/02
' | @written    : 2021/01/02
' | @remarks
' |     �t�H���_���t���p�X�Ŏ擾���A�t���p�X�Ɛe�t�H���_����Ԃ�
' --------------------------------------+-----------------------------------------
    Dim tbl                             As Variant
    Dim openFolder                      As String

    Dim z, zMin, zMax                   As Long
    Dim fullPath                        As String
    Dim folderName                      As String
    Dim delimiterChar                   As String
'
' ---Procedure Division ----------------+-----------------------------------------
'

'�@[�t�@�C�����J��]�_�C�A���O�{�b�N�X�őΏ�Excel��I�����܂�

        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = P_title
            If .Show = True Then
                openFolder = .SelectedItems(1)
            End If
        End With
        
        tbl = ""
        tbl = Split(openFolder, "\")
        zMin = LBound(tbl, 1)           ' �J�n�s
        zMax = UBound(tbl, 1)           ' �ŏI�s
        delimiterChar = ""
        For z = zMin To zMax
            If tbl(z) <> "" Then
                If fullPath <> "" Then  ' �p�X�̐擪�́A�h���C�u�����Ȃ̂ŁA��؂蕶�����͂��Ȃ�
                    delimiterChar = "\"
                End If
                fullPath = fullPath + delimiterChar + tbl(z)
            End If
        Next z
        
        folderName = tbl(zMax)
    
' ����
    getFolderPath_F = Array(fullPath, folderName)

End Function

Public Function selectFolder(ByVal P_title As String) As Variant
' --------------------------------------+-----------------------------------------
' | @function   : [�t�@�C�����J��]�_�C�A���O�{�b�N�X�Ńt�H���_��I�����܂�
' --------------------------------------+-----------------------------------------
' | @moduleName : util18_selectFolder
' | @Version    : v1.0.0
' | @update     : 2020/01/02
' | @written    : 2021/01/02
' | @remarks
' |     [�t�@�C�����J��]�_�C�A���O�{�b�N�X�Ńt�H���_��I����,
' |     �e�t�H���_�̃t���p�X�ƁA���̃t�H���_����Ԃ�
' --------------------------------------+-----------------------------------------
    Dim tbl                             As Variant
    Dim openFolder                      As String

    Dim z, zMin, zMax                   As Long
    Dim fullPath                        As String
    Dim folderName                      As String
    Dim delimiterChar                   As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
'�@[�t�@�C�����J��]�_�C�A���O�{�b�N�X�őΏ�Excel��I�����܂�

        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = P_title
            If .Show = True Then
                openFolder = .SelectedItems(1)
            End If
        End With
        
        tbl = ""
        tbl = Split(openFolder, "\")
        zMin = LBound(tbl, 1)           ' �J�n�s
        zMax = UBound(tbl, 1)           ' �ŏI�s
        delimiterChar = ""
        For z = zMin To zMax
            If tbl(z) <> "" Then
                If fullPath <> "" Then  ' �p�X�̐擪�́A�h���C�u�����Ȃ̂ŁA��؂蕶�����͂��Ȃ�
                    delimiterChar = "\"
                End If
                fullPath = fullPath + delimiterChar + tbl(z)
            End If
        Next z
        
        folderName = tbl(zMax)
    
' ����
    selectFolder = Array(fullPath, folderName)

End Function

Public Function getLatestFile(ByVal FolderPath As String) As String
' --------------------------------------+-----------------------------------------
' | @function   : �t�H���_�ɂ���ŐV�̃t�@�C���𒲂ׂ�
' --------------------------------------+-----------------------------------------
' | @moduleName : util19_getLatestFile
' | @Version    : v1.0.0
' | @update     : 2020/01/02
' | @written    : 2021/01/02
' | @remarks
' |     �t�H���_�ɂ���ŐV�̃t�@�C���𒲂ׂ�
' --------------------------------------+-----------------------------------------
    Dim Buf                             As String
    Dim fList(99)                       As String
    Dim i, iMin, iMax                   As Long
    Dim latestFile                      As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
     
    iMin = LBound(fList, 1)           ' �J�n�s
    iMax = UBound(fList, 1)           ' �ŏI�s
    i = 0
    Buf = Dir(FolderPath & "\*.xlsm")
    Do While Buf <> ""
        i = i + 1
        fList(i) = Buf
        Buf = Dir()
    Loop
    bubbleSort (fList)
    latestFile = fList(iMin)
    For i = iMin + 1 To iMax
        If fList(i) <> "" Then
            If latestFile < fList(i) Then
                latestFile = fList(i)
            End If
        End If
    Next i

' ����
    getLatestFile = latestFile

End Function

Public Function getFolderPath_F(ByVal P_title As String)
' --------------------------------------+-----------------------------------------
' | @function   : [�t�@�C�����J��]�_�C�A���O�{�b�N�X�Ńt�@�C����I�����܂�
' --------------------------------------+-----------------------------------------
' | @moduleName : util20_getFolderPath_F
' | @Version    : v1.0.0
' | @update     : 2020/01/02
' | @written    : 2021/01/02
' | @remarks
' |     �t�H���_���t���p�X�Ŏ擾���A�t���p�X�Ɛe�t�H���_����Ԃ�
' --------------------------------------+-----------------------------------------
    Dim tbl                             As Variant
    Dim openFolder                      As String

    Dim z, zMin, zMax                   As Long
    Dim fullPath                        As String
    Dim folderName                      As String
    Dim delimiterChar                   As String
'
' ---Procedure Division ----------------+-----------------------------------------
'

'�@[�t�@�C�����J��]�_�C�A���O�{�b�N�X�őΏ�Excel��I�����܂�

        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = P_title
            If .Show = True Then
                openFolder = .SelectedItems(1)
            End If
        End With
        
        tbl = ""
        tbl = Split(openFolder, "\")
        zMin = LBound(tbl, 1)           ' �J�n�s
        zMax = UBound(tbl, 1)           ' �ŏI�s
        delimiterChar = ""
        For z = zMin To zMax
            If tbl(z) <> "" Then
                If fullPath <> "" Then  ' �p�X�̐擪�́A�h���C�u�����Ȃ̂ŁA��؂蕶�����͂��Ȃ�
                    delimiterChar = "\"
                End If
                fullPath = fullPath + delimiterChar + tbl(z)
            End If
        Next z
        
        folderName = tbl(zMax)
    
' ����
    getFolderPath_F = Array(fullPath, folderName)

End Function

Public Function �t�@�C���������_F(ByVal P_MSG As String) As String
' --------------------------------------+-----------------------------------------
' | @function   : �t�@�C�����𓾂�
' --------------------------------------+-----------------------------------------
' | @moduleName : util23_�t�@�C���������
' | @Version    : v1.0.0
' | @update     : 2021/06/22
' | @written    : 2021/06/22
' | @remarks
' |     �߂�l�FfullPath
' |     �ړI�̃t�@�C�����i�t���p�X�j�𓾂�
' |
' --------------------------------------+-----------------------------------------
    Dim OpenPathFileName                As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
'�@[�t�@�C�����J��]�_�C�A���O�{�b�N�X�őΏ�Excel��I�����܂�
    OpenPathFileName = Application.GetOpenFilename("Excel�t�@�C��,*.xl*", , P_MSG)
        If OpenPathFileName = "False" Then
            MsgBox "�I������EXCEL���G���[�ł��B�����I�����܂��B"
        End                             ' �����̋����I��
    End If
    
 
' ����
    �t�@�C�����𓾂�_F = OpenPathFileName

End Function

Public Function toA1_F(ByVal StrR1C1 As String) As String
' --------------------------------------+-----------------------------------------
' | @function   : �Z����R1C1�`����A1�`���ɕϊ�
' --------------------------------------+-----------------------------------------
' | @moduleName : util24_R1C1�`��2A1�`���ϊ�_F
' | @Version    : v1.0.0
' | @update     : 2021/06/24
' | @written    : 2021/06/24
' | @remarks
' |     �߂�l�FString�^
' --------------------------------------+-----------------------------------------
'
' ---Procedure Division ----------------+-----------------------------------------
'
    toA1_F = Application.ConvertFormula( _
        Formula:=StrR1C1, _
        fromReferenceStyle:=xlR1C1, _
        toreferencestyle:=xlA1, _
        toabsolute:=xlRelative)

End Function

Public Sub util29_�u�b�N�v���p�e�B�擾(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' | @function   : Book�̃v���t�@�C�����擾
' --------------------------------------+-----------------------------------------
' | @moduleName : util29_�u�b�N�v���p�e�B�擾
' | @Version    : v1.0.0
' | @update     : 2021/04/07
' | @written    : 2022/04/07
' | @remarks
' |  �uDocInfo(�폜�s��)�v�V�[�g���Q��
' --------------------------------------+-----------------------------------------
'   ��public�ϐ�(���Y�v���W�F�N�g���̃��W���[���Ԃŋ��L)�́A�ŏ��ɌĂ΂��v���V�W���[�ɒ�`
'     �ړ���� P_ ������
'
'
' ---Procedure Division ----------------+-----------------------------------------
'
    MsgBox "Title: " & ThisWorkbook.BuiltinDocumentProperties("Title") & vbNewLine _
        & "Author: " & ThisWorkbook.BuiltinDocumentProperties("Author") & vbNewLine _
        & "Subject: " & ThisWorkbook.BuiltinDocumentProperties("Subject") & vbNewLine _
        & "Keywords: " & ThisWorkbook.BuiltinDocumentProperties("Keywords") & vbNewLine _
        & "Category: " & ThisWorkbook.BuiltinDocumentProperties("Category") & vbNewLine _
        & "Comments: " & ThisWorkbook.BuiltinDocumentProperties("Comments")
End Sub

Public Function getImportSheet_F(ByVal p_importFile As String, ByVal p_importSheet As String, ByVal p_saveSheet As String, ByVal p_childPath As String, ByVal p_importMsg As String)
' --------------------------------------+-----------------------------------------
' | @function   : ��Excel�̃V�[�g��import����}�N��
' |             �@importExcel���Ȃ����A�G���[�̂Ƃ��̓G�b�N�X�v���[���őI��
' --------------------------------------+-----------------------------------------
' | @moduleName : util30_getImportSheet
' | @Version    : v1.0.0
' | @update     : 2023/04/29
' | @written    : 2023/04/29
' | @remarks
' |
' --------------------------------------+-----------------------------------------
    Dim wb                              As Workbook
    Dim ws                              As Worksheet
    Dim wbImp                           As Workbook
    Dim sw_FalseTrue                    As Boolean
    Dim w_path                          As String
'
' ---Procedure Division ----------------+-----------------------------------------
'

' 1.import����V�[�g���i�[����V�[�g�mwork�n�́A���O�ɍ폜
' --------------------------------------+-----------------------------------------
'
    Set wb = ActiveWorkbook
    If IsExistSheet(p_saveSheet) Then
        wb.Worksheets(p_saveSheet).Delete            ' �ȑO�̃V�[�g���폜
    End If
' �t�@�C���w��̗L���`�F�b�N�@/�@�w�肵���t�@�C�����Ȃ�������A�G�N�X�v���[���Ŏw�肳����
    sw_FalseTrue = False
    If p_importFile <> "" Then
        If IsExistFileDir(p_importFile) Then
            sw_FalseTrue = True
        Else
            sw_FalseTrue = False
        End If
    End If
'�@[�t�@�C�����J��]�_�C�A���O�{�b�N�X�őΏ�Excel��I�����܂�
    If sw_FalseTrue = False Then
        If p_childPath = "" Then
            w_path = PathName
        Else
            w_path = SubSysPath & "\" & p_childPath
        End If
        
        ChDir w_path                                    ' �v���O�����̂���t�H���_���w��
        p_importFile = Application.GetOpenFilename("Excel�t�@�C��,*.xl*", , p_importMsg)
        sw_FalseTrue = True
    End If
    
' �O��Excel�t�@�C�����J���Aimport�V�[�g����ƃV�[�g work �փR�s�[
    Workbooks.Open p_importFile
    Set wbImp = ActiveWorkbook
    wbImp.Worksheets(p_importSheet).Copy after:=wb.Worksheets(1)
    
    Set ws = ActiveSheet
    ws.Name = p_saveSheet
    ws.Tab.Color = RGB(0, 112, 192)     ' ��
    wbImp.Close saveChanges:=False      ' �ۑ����Ȃ���close

' �I�u�W�F�N�g�ϐ��̉��
    Set wbImp = Nothing
    Set wb = Nothing
    Set ws = Nothing
' �߂�l
    getImportSheet_F = p_importFile

End Function

Public Sub putStatusBar(ByVal p_statusBarMsg As String)
' --------------------------------------------------------------------------------
' | @function   : EXCEL�̃X�e�[�^�X�o�[�ɐi�s�󋵂�\��
' --------------------------------------+-----------------------------------------
' | @moduleName : util00_�X�e�[�^�X�o�[�\���E����
' | @Version    : v1.1.0
' | @update     : 2023/05/17
' | @written    : 2021/12/11
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |
' |
' --------------------------------------+-----------------------------------------
'
' ---Procedure Division ----------------+-----------------------------------------
'/
'      EXCEL�̃X�e�[�^�X�o�[�ɐi�s�󋵂�\�� / ����

    If p_statusBarMsg = "" Then
        Application.StatusBar = False   '      EXCEL�̃X�e�[�^�X�o�[������
    Else
        Application.StatusBar = Format(Now(), "m/d hh:mm") & "�F" & p_statusBarMsg & " ���������ł��B"
    End If
    Calculate                         ' �@�ŐV��ԂɍX�V
End Sub

Public Sub debug2text(ByVal p_text As String, Optional P_mode As String = "print")
' --------------------------------------------------------------------------------
' | @function   : Debug.Print���e�L�X�g�t�@�C���ɏo��
' --------------------------------------+-----------------------------------------
' | @moduleName : util33_debug2File
' | @Version    : v1.0.0
' | @update     : 2023/05/22
' | @written    : 2023/05/22
' | @author     : Jun Fujinawa
' | @license    : zStudio
' | @remarks
' |     ����: p_test
' |     ����: p_mode
' |         open  �� �e�L�X�g�t�@�C���� open ����
' |         print �� �e�L�X�g�̏o��(Default)
' |         close �� �e�L�X�g�t�@�C���́@close ����
' |
' |
' |
' --------------------------------------+-----------------------------------------
'
  Dim dt                                As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
    Select Case P_mode
        Case "open"
' �����������擾
            dt = Format(Now, "yyyymmdd_hhmmss") ' ���݂̓�����yyyymmdd_hhmmss�`���Ŏ擾
' �g�p�\�ȃt�@�C���ԍ����擾
            FileNum = FreeFile()
' ��ɐV�K�t�@�C�����쐬�A�t�@�C�����̏d����h�����ߏ����������t�@�C�����ɒǋL
            Open ThisWorkbook.Path & "\" & "loggPrint" & "_" & dt & ".txt" For Output As #FileNum
        
        Case "close"
' ���O�t�@�C��CLOSE
            Close #FileNum
        
        Case "print"
' ���O�o��
            Print #FileNum, Now & " " & p_text
            Debug.Print Now & " " & p_text
            
        Case Else
' ���O�o��
            Print #FileNum, Now & " " & p_text
            Debug.Print Now & " " & p_text
    End Select
        
    
    

'' �����������擾
'    dt = Format(Now, "yyyymmdd_hhmmss") ' ���݂̓�����yyyymmdd_hhmmss�`���Ŏ擾
'' ���O�t�@�C���̃p�X��ݒ�
'    filePath = SubSysPath & "\" & "DebugPrint.txt"
'' �g�p�\�ȃt�@�C���ԍ����擾
'    fileNo = FreeFile()
'' ��ɐV�K�t�@�C�����쐬�A�t�@�C�����̏d����h�����ߏ����������t�@�C�����ɒǋL
'    Open filePath & "_" & dt & ".txt" For Output As #fileNo
'' ���O�o��
'    Print #fileNo, Now & " " & p_text
'' ���O�t�@�C��CLOSE
'    Close #fileNo
'
    
End Sub

' @ReadMe �W���t�H�[���i��j
' --------------------------------------+-----------------------------------------
' | @function   : �W���u�̏��������i�W���Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : util31_get���ʕϐ�
' | @Version    : v1.0.0
' | @update     : 2023/05/04
' | @written    : 2023/05/04
' | @remarks
' |     ���[�U��`��Public�ϐ��̏����l�ݒ�
' |
' |   ��public�ϐ�(���Y�v���W�F�N�g���̃��W���[���Ԃŋ��L)�́A�ŏ��ɌĂ΂��v���V�W���[�ɒ�`
' |     �ړ���� P_ ������
' --------------------------------------+-----------------------------------------
'   +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
'
' ---Procedure Division ----------------+-----------------------------------------
'

' closeingMsg = "|�@����V�[�g" & Chr(9) & "�� " & SrcCnt & Chr(13) _
'             & "|�Aarchives" & Chr(9) & "�� " & ArvCnt & Chr(13) _
'             & "|�B�ڎ�" & Chr(9) & "�� " & EyeCnt & Chr(13) _
'             & "| �G���[" & Chr(9) & "�� " & ErrCnt & Chr(13)
  
