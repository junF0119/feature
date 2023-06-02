Attribute VB_Name = "m3_���͍��ڐ��K��"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : ���͍��ڂ𐳋K������i���W���[�������Łj
' --------------------------------------+-----------------------------------------
' | @moduleName : m3_���͍��ڐ��K��
' | @Version    : v1.0.0
' | @update     : 2023/05/25
' | @written    : 2023/05/25
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
' |     �v���O�����\��
' |         1. ��������
' |             ��������_R()
' |             1.1 �����V�[�g�̃N���A
' |                 importClear_R()
' |             1.2 �O���̃}�X�^�[�̃V�[�g����荞�ށc�c M-�@�V�Z���^���� / M-�AArchives
' |                 importSheet_R()(
' |         2. �f�[�^�̐��������؁c�c (53)PrimaryKey / (42)key����
' |             �L�[����_R()
' |             2.1 �V�[�g�̔z�� �c�c �@���� / �Aarchives
' |                 arrSet_R()
' |             2.2 �L�[���ڂ̏d���`�F�b�N
' |                 duplicateChk_F()
' |                     quickSort_R()
' |             2.3 �L�[���ڂ�Null�l�`�F�b�N
' |                 nullKeyChk_F()
' |         3.�@�f�[�^���ڂ̐��K���c�c�L�[���ڂ��������̓f�[�^�̐��K��
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

Private sw_errorChk                     As Boolean          ' true�c�G���[�����Afalse�c�G���[�L��

' �\���̂̐錾
Type pkeyStruct
    sortKey                             As Variant          ' quick sort�p�L�[
    primaryKey                          As Integer          ' (53)PrimaryKey
    nameKey                             As String           ' (42)key����
    sheetName                           As String           ' �V�[�g��
    rowAddress                          As Integer          ' ���R�[�h�̍s(row)�ʒu
End Type
Private ary()                           As pkeyStruct      ' �\���̂̈ꌳ�������I�z��isortKey,�c�j
Private j, jMax                         As Long            ' �z�� ary() �̃C���f�b�N�X
Private primaryKeyMax                   As Long            ' primaryKey�̍ő�l

Public Sub ���͍��ڐ��K��_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' | @function   : ���͍��ڂ𐳋K�����A���̓f�[�^�̂΂�����������
' --------------------------------------+-----------------------------------------
' | @moduleName : ���͍��ڐ��K��_R
' | @Version    : v1.0.0
' | @update     : 2023/05/25
' | @written    : 2023/05/25
' | @remarks
' |     �G���[�Ȃ��ctrue
' |     �G���[����cfalse
' |�@   �G���[������Ƃ��̑[�u
' |�@       �@?���Ɂ@�H�@�}�[�N��t��
' |�@       �AErrCnt ���J�E���g
' |�@       �B�`�F�b�N���ڂ��ƂɃG���[������Ƃ��́A�v���O�����𒆒f���A
' |          �}�j���A���Ō��f�[�^���C�����A���̃v���O�������Ď��s����
' |
' |
' --------------------------------------+-----------------------------------------
    Dim y, yMax                         As Long
    Dim sw_result                       As Boolean
    Dim errMsg                          As String
    Dim err42Cnt                        As Long
    Dim err53Cnt                        As Long

    Dim pkeyStruct                      As pkeyStruct   ' �I�u�W�F�N�g�̒�`�i�\���́j
' �\���̂̐錾
' Type pkeyStruct
'     sortKey       As Variant          ' quick sort�p�L�[
'     primaryKey    As Integer          ' (53)PrimaryKey
'     nameKey       As string           ' (42)key����
'     sheetName     As String           ' �V�[�g��
'     rowAddress    As Integer          ' ���R�[�h�̍s(row)�ʒu
' End Type
'
' ---Procedure Division ----------------+-----------------------------------------
'
    sw_result = True            ' �G���[�Ȃ�
    ErrCnt = 0
    primaryKeyMax = 0
' --------------------------------------+-----------------------------------------
' 1.�@����A�Aarchive�V�[�g����f�[�^��ary()�e�[�u���ɐݒ肷��
' --------------------------------------+-----------------------------------------
' ���I�z��̏�����(�z��0�͎g��Ȃ��A�w�b�_�[���͏����j
'   dim dataCnt As Long            ' �L���f�[�^����
'    ReDim ary(dataCnt)
'    dataCnt = SrcYmax + ArvYmax - (YMIN - 1) * 2
'    MsgBox dataCnt & " " & LBound(ary) & "-" & UBound(ary)

    ReDim ary(SrcYmax + arvYmax - (YMIN - 1) * 2 - 1)   ' index= 0 - �̂��߁@-1�@����
' �@���� + �Aarchives ��z�� ary() �փR�s�[
    jMax = -1
    Call arrSet_R(SrcCnt, SrcYmin, SrcYmax, Range("C_SrcSheet").Value)     ' �@����V�[�g��z�� ary() �փR�s�[
    Call arrSet_R(arvCnt, arvYmin, arvYmax, Range("C_arvSheet").Value)      ' �����āA�Aarchives ���R�s�[

' --------------------------------------+-----------------------------------------
' 2.�L�[���ڂ� null �̃��R�[�h���Ȃ����T���@(42)key���� / (53)PrimaryKey
' --------------------------------------+-----------------------------------------
    err42Cnt = 0
    err53Cnt = 0
    If Not nullKeyChk_F(err42Cnt, err53Cnt) Then
        errMsg = "|null ���R�[�h������܂����B" & Chr(13) _
               & "|(42)key����" & Chr(9) & "�� " & err42Cnt & Chr(13) _
               & "|(53)PrimaryKey" & Chr(9) & "�� " & err53Cnt & Chr(13) _
               & "| �G���[�ӏ����C�������̂Ŋm�F���Ă��������B" & Chr(13) _
               & "| ��������΁A���{�𒼐ڎ�C�����邩�A�R�s�[���Ă��������B" & Chr(13) _
               & "| ���̃v���O�����͋����I�����܂��B"
        MsgBox errMsg
        End
    Else
        StatusBarMsg = "�L�[�� Null �̃��R�[�h����܂���B" & Chr(13) & _
                "�v���O�������p�����܂��B"
        Call putStatusBar(StatusBarMsg)
    End If

' --------------------------------------+-----------------------------------------
' 3.(53)PrimaryKey�ŏ����\�[�g��A(53)PrimaryKey�̏d���`�F�b�N�����{
' --------------------------------------+-----------------------------------------
    If Not duplicateChk_F("(53)PrimaryKey") Then
        MsgBox "(53)PrimaryKey�ɏd�����A" & ErrCnt & " ������܂��B�C�����Ă��������B" & Chr(13) & _
                "���̃v���O�����́A�����I�����܂��B"
        End
    Else
        StatusBarMsg = "(53)PrimaryKey�ɏd���́A����܂���B" & Chr(13) & _
                "�v���O�������p�����܂��B"
        Call putStatusBar(StatusBarMsg)
    End If


' --------------------------------------+-----------------------------------------
' 4.(42)key�����ŏ����\�[�g��A(42)key�����̏d���`�F�b�N�����{
' --------------------------------------+-----------------------------------------
    If Not duplicateChk_F("(42)key����") Then
        MsgBox "(42)key�����ɏd�����A" & ErrCnt & " ������܂��B�C�����Ă��������B" & Chr(13) & _
                "���̃v���O�����́A�����I�����܂��B"
        End
    Else
        StatusBarMsg = "(42)key�����ɏd���́A����܂���B" & Chr(13) & _
                "�v���O�������p�����܂��B"
        Call putStatusBar(StatusBarMsg)
    End If
    
  
End Sub

Private Sub arrSet_R(ByRef p_cnt As Long, ByVal p_yMin As Long, ByVal p_yMax As Long, ByVal p_sheetName As String)
' --------------------------------------+-----------------------------------------
' |  PrimaryKey��ary�z��̊i�[
' --------------------------------------+-----------------------------------------
' �\���̂̐錾
' Type pkeyStruct
'     sortKey                             as variant          ' quick sort�p�L�[
'     primaryKey                          As Integer          ' (53)PrimaryKey
'     nameKey                             as string           ' (42)key����
'     sheetName                           As String           ' �V�[�g��
'     rowAddress                          As Integer          ' ���R�[�h�̍s(row)�ʒu
' End Type

    Dim pkey                            As pkeyStruct   ' �I�u�W�F�N�g�̒�`�i�\���́j
'
' ---Procedure Division ----------------+-----------------------------------------
'
    p_cnt = 0
    For j = p_yMin To p_yMax    '(�s)
        jMax = jMax + 1
        pkey.primaryKey = Wb.Worksheets(p_sheetName).Cells(j, PRIMARYKEY_X)     ' BA��i53�j
        pkey.nameKey = Wb.Worksheets(p_sheetName).Cells(j, PKEY_X)              ' AP��i42�j
        pkey.sheetName = p_sheetName
        pkey.rowAddress = j
        ary(jMax) = pkey
' primaryKey�̍ő�l�𓾂�
        If pkey.primaryKey > primaryKeyMax Then
            primaryKeyMax = pkey.primaryKey
        End If
        p_cnt = p_cnt + 1
    Next j
    
End Sub

Private Function nullKeyChk_F(ByRef p_err42Cnt As Long, ByRef p_err53Cnt As Long)
' --------------------------------------+-----------------------------------------
' | @function  �w��t�B�[���h��null�f�[�^���Ȃ����Ƃ̌���
' --------------------------------------+-----------------------------------------
' | @moduleName: nullKeyChk_F
' | @remarks
' |   �G���[�Ȃ��ctrue
' |   �G���[����cfalse
' |�@�G���[������Ƃ��̑[�u
' |�@�@?���Ɂ@�H�@�}�[�N��t��
' |�@�AErrCnt �ɃJ�E���g����
' |�@�B���f�[�^���w��t�B�[���ɃZ�b�g����
' |
' | �����̈Ӗ�
' | �߂�l�Fp_err42Cnt: (42)key������null�̌���
' | �߂�l�Fp_err53Cnt: (53)PrimaryKey��null�̌���
' |
' | �\���̂̐錾
' |  Type pkeyStruct
' |      sortKey                             as variant          ' quick sort�p�L�[
' |      primaryKey                          As Integer          ' (53)PrimaryKey
' |      nameKey                             as string           ' (42)key����
' |      sheetName                           As String           ' �V�[�g��
' |      rowAddress                          As Integer          ' ���R�[�h�̍s(row)�ʒu
' |  End Type
' |
' --------------------------------------+-----------------------------------------
    Dim y, yMax                         As Long
    Dim sw_result                       As Boolean
    Dim pkeyStruct                      As pkeyStruct   ' �I�u�W�F�N�g�̒�`�i�\���́j
    Dim w_contents                      As String
    Dim w_fullName                      As Variant      ' �����F����@��
    Dim w_firstName                     As String       ' ���O�F��
    Dim w_familyName                    As String       ' ���F����
    
    Dim debugText                       As String
'
' ---Procedure Division ----------------+-----------------------------------------
'
    sw_result = True            ' �G���[�Ȃ�
    p_err42Cnt = 0
    p_err53Cnt = 0
    
' �w��L�[��Null�l�`�F�b�N
    For y = LBound(ary) To UBound(ary)
' (42)key�����̃`�F�b�N���C��
        If ary(y).nameKey = "" Then
            p_err42Cnt = p_err42Cnt + 1
            sw_result = False
            w_contents = Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, 6)
            w_fullName = Split(w_contents, " ")   ' (6)���O�@��؂蕶���F���p�̋�
            w_familyName = w_fullName(0)
            w_firstName = w_fullName(1)
            Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, PKEY_X) = w_familyName & w_firstName
            Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, CHECKED_X) = Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, CHECKED_X) & "��"

        End If
        If ary(y).primaryKey = 0 Then
            p_err53Cnt = p_err53Cnt + 1
            sw_result = False
            primaryKeyMax = primaryKeyMax + 1
            Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, PRIMARYKEY_X) = primaryKeyMax
            Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, CHECKED_X) = Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, CHECKED_X) & "��"
            
        End If

    Next y
    
' �߂�l
    nullKeyChk_F = sw_result

End Function

Private Function duplicateChk_F(ByVal p_sortKey As Variant) As Boolean
' --------------------------------------+-----------------------------------------
' | @function  �w��t�B�[���h�Ƀf�[�^�̏d�����Ȃ����Ƃ̌���
' --------------------------------------+-----------------------------------------
' | @moduleName: duplicateChk_R
' | @remarks
' |   �G���[�Ȃ��ctrue
' |   �G���[����cfalse
' |�@�G���[������Ƃ��̑[�u
' |�@�@?���Ɂ@�H�@�}�[�N��t��
' |�@�AcntErr �ɃJ�E���g����
' |
' --------------------------------------+-----------------------------------------
    Dim y, yMax                         As Long
    Dim sw_result                       As Boolean
    
    Dim pkeyStruct                      As pkeyStruct   ' �I�u�W�F�N�g�̒�`�i�\���́j
    Dim zlogMsg                         As String
    Dim z                               As Long
'
' ---Procedure Division ----------------+-----------------------------------------
'
    sw_result = True            ' �G���[�Ȃ�
    ErrCnt = 0
    Call quickSort_R(ary(), p_sortKey, LBound(ary), UBound(ary), xlAscending)
' sortKey�̏d���`�F�b�N
    For y = LBound(ary) To UBound(ary) - 1
        If ary(y).sortKey = 0 Or ary(y).sortKey = "" Then
            GoTo SkipRow
        End If
        If ary(y).sortKey = ary(y + 1).sortKey Then
            sw_result = False
            ErrCnt = ErrCnt + 1
            Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, CHECKED_X) = Sheets(ary(y + 1).sheetName).Cells(ary(y + 1).rowAddress, PKEY_X).Value  ' �����ker������\��
            Sheets(ary(y + 1).sheetName).Cells(ary(y + 1).rowAddress, CHECKED_X) = Sheets(ary(y).sheetName).Cells(ary(y).rowAddress, PKEY_X).Value  '       �V
        End If
SkipRow:
    Next y
    
' ����
    duplicateChk_F = sw_result

End Function

Private Sub quickSort_R(ByRef argAry() As pkeyStruct, _
                        ByVal p_keyName As String, _
                        ByVal p_lngMin As Long, _
                        ByVal p_lngMax As Long, _
                        Optional sOrder As XlSortOrder = xlAscending)
' --------------------------------------+-----------------------------------------
' | @function  : �\���̔z��̃N�C�b�N�\�[�g
' --------------------------------------+-----------------------------------------
' | @moduleName: quickSortk_R
' | @remarks
' | �����̈Ӗ�
' | �߂�l�Fary() �\�[�g��̔z��
' | ���@���Fp_keyName: �\�[�g�L�[�̍��ږ�
' | ���@���Fp_lngMin�F�z��Y���̍ŏ��lLBound
' | ���@���Fp_lngMax: �z��Y���̍ő�lUBound
' | �\���̂̐錾
' |  Type pkeyStruct
' |      sortKey                             as variant          ' quick sort�p�L�[
' |      primaryKey                          As Integer          ' (53)PrimaryKey
' |      nameKey                             as string           ' (42)key����
' |      sheetName                           As String           ' �V�[�g��
' |      rowAddress                          As Integer          ' ���R�[�h�̍s(row)�ʒu
' |  End Type
' |
' --------------------------------------+-----------------------------------------
    Dim i                               As Long
    Dim j                               As Long
    Dim vBase                           As pkeyStruct
    Dim vTemp                           As pkeyStruct
    Dim vSwap                           As pkeyStruct
'
' ---Procedure Division ----------------+-----------------------------------------
'
' --------------------------------------+-----------------------------------------
' |  (1) sort key �̖��O���w�肵�A�����ary()��sortKey�փZ�b�g����
' --------------------------------------+-----------------------------------------
    For j = p_lngMin To p_lngMax
        Select Case p_keyName
            Case "(53)PrimaryKey"
                ary(j).sortKey = ary(j).primaryKey     ' BA��i53�j
            Case "(42)key����"
                ary(j).sortKey = ary(j).nameKey        ' AP��i42�j
            Case Else
                MsgBox "�v���O�����̃o�O�ł��B" & Chr(13) & _
                "p_keyName=" & p_keyName & "�́A��`����Ă��܂���B" & Chr(13) & _
                "�I�����܂��B"
                End
        End Select

    Next j
' �o�u���\�[�g
    For i = p_lngMax To p_lngMin Step -1
        For j = p_lngMin To i - 1
            If argAry(j).sortKey > argAry(j + 1).sortKey Then
                vSwap = argAry(j)
                argAry(j) = argAry(j + 1)
                argAry(j + 1) = vSwap
            End If
        Next j
    Next i
'

End Sub

'Dim zz As Long
'Dim debugText As String
'Call debug2text("", "open")
'For zz = 0 To SrcYmax + (ArvYmax - (YMIN - 1) * 2) - 1
'debugText = "|bb=" & zz & _
'            "|sortKey=" & ary(zz).sortKey & Chr(9) & _
'            "|primaryKey=" & ary(zz).primaryKey & Chr(9) & _
'            "|nameKey=" & ary(zz).nameKey & Chr(9) & _
'            "|sheetName=" & ary(zz).sheetName & Chr(9) & Chr(9) & _
'            "|rowAddress=" & ary(zz).rowAddress
'Call debug2text(debugText)
'Next zz
'Call debug2text("", "close")
'Stop




