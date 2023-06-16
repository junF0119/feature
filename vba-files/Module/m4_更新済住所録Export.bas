Attribute VB_Name = "m4_�X�V�ϏZ���^Export"
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : �@����ƇAarchives���R�[�h�����݂����V�Z���^�V�[�g���炻�ꂼ���Excel�u�b�N���o��
' --------------------------------------+-----------------------------------------
' | @moduleName : m4_�Z���^Export
' | @Version    : v1.0.0
' | @update     : 2023/06/11
' | @written    : 2023/06/11
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

Public Sub m4_�X�V�ϏZ���^Export_R(ByVal dummy As Variant)
' --------------------------------------+-----------------------------------------
' |     �V�Z���^�V�[�g����@�Z���^����ƇAArchive��Excel���o��
' --------------------------------------+-----------------------------------------
'
    Dim newCnt                          As Long: newCnt = 0
    Dim arvCnt                          As Long: arvCnt = 0
    Dim saveDir                         As String

'
' ---Procedure Division ----------------+-----------------------------------------
'
' �I�u�W�F�N�g�ϐ��̒�`�i���ʁj
    Set Wb = ThisWorkbook

' �Bnew �V�[�g�̏����l & �\�̑傫���𓾂�
    Set wsNew = Wb.Worksheets(Range("C_newSheet").Value)                    ' �V�Z���^�V�[�g
    newYmin = YMIN
    newXmin = XMIN
    newYmax = wsNew.Cells(Rows.Count, PSEIMEI_X).End(xlUp).Row              ' �ŏI�s�i�c�����j
    newXmax = wsNew.Cells(YMIN - 1, Columns.Count).End(xlToLeft).Column     ' �ŏI��i�������j
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

    Sheets("�V�Z���^").Select
    Sheets("�V�Z���^").Copy Before:=Sheets(1)
    Rows("3:3").Select
    Selection.AutoFilter
    Rows("3:3").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("�V�Z���^ (2)").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�V�Z���^ (2)").Sort.SortFields.Add2 key:=Range( _
        "BB4:BB801"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("�V�Z���^ (2)").Sort
        .SetRange Range("A3:WWN801")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll ToRight:=34
    ActiveWindow.SmallScroll Down:=138
    Range("BB152:BB546").Select
    Selection.EntireRow.Delete
    Sheets("�V�Z���^ (2)").Select
    Sheets("�V�Z���^ (2)").Copy
    Application.Left = -1174.25
    Application.Top = 272.5
    Windows("zz2.1.1-�V�Z���^�X�V-v1.3.0.take1-20230611.xlsm").Activate
    Sheets("�Hlabel").Select
    Sheets("�Hlabel").Copy Before:=Workbooks("Book1").Sheets(1)
    Application.Left = 163
    Application.Top = 97
    ActiveWorkbook.Names("C_���x���ꗗ").Delete
    ActiveWorkbook.Names("pathName").Delete
    ActiveWorkbook.Names("pathName").Delete
    ActiveWorkbook.Names("update").Delete
    ActiveWorkbook.Names("update").Delete
    ActiveWorkbook.Names("updateTxt").Delete
    ActiveWorkbook.Names("updateTxt").Delete
    ActiveWorkbook.Names("updateTxt2").Delete
    ActiveWorkbook.Names("verShort").Delete
    ActiveWorkbook.Names("verShort").Delete
    ActiveWorkbook.Names("version").Delete
    ActiveWorkbook.Names("version").Delete
    ActiveWorkbook.Names("verUpdateTxt2").Delete
    ActiveWorkbook.Names("verUpdateTxt2").Delete
    Sheets("�V�Z���^ (2)").Select
    Sheets("�V�Z���^ (2)").Move Before:=Sheets(1)
    Sheets("�V�Z���^ (2)").Select
    Sheets("�V�Z���^ (2)").Name = "�@����"
    ChDir "D:\Desktop\2.Job1-�V�Z���^�X�V(zz2.1)-v1.3\1.1.inputData"
    ActiveWorkbook.SaveAs FileName:= _
        "D:\Desktop\2.Job1-�V�Z���^�X�V(zz2.1)-v1.3\1.1.inputData\M-�@�V�Z���^����-v1.1.0-20230611.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
End Sub



Sub Sample6()
    Dim i As Long, Target As String
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Target = Cells(i, 1)
        Sheets(2).Name = Target & "_�\�Z"
        Sheets(3).Name = Target & "_����"
        Sheets(Array(Target & "_�\�Z", Target & "_����")).Copy
        ActiveWorkbook.SaveAs "D:\Work\" & Target & ".xlsx"
        ActiveWorkbook.Close
    Next i
End Sub
















' --------------------------------------+-----------------------------------------
' ����L�[��(42)key�������A3��1or2��0�@�̏��ɕ��Ԃ̂ŁA�ύX���ڂ� 0 ���R�[�h���X�V���A�V�Z���^�V�[�g�փR�s�[����
    For j = wrkYmin To wrkYmax Step 3
        wsWrk.Range(Cells(j, 1), Cells(j, wrkXmax)).Font.Color = rgbBlack                   ' �����F�F���@      #000000#
        wsWrk.Range(Cells(j, 1), Cells(j, wrkXmax)).Interior.Color = rgbKhaki               ' �w�i�F�F�J�[�L    #8ce6f0#
        wsWrk.Range(Cells(j + 1, 1), Cells(j + 1, wrkXmax)).Font.Color = rgbSnow            ' �����F�F�X�m�[    #fafaff#
        wsWrk.Range(Cells(j + 1, 1), Cells(j + 1, wrkXmax)).Interior.Color = rgbDodgerBlue  ' �w�i�F�F�h�W���[�u���[    #ff901e#

        sw_change = False
        For i = 6 To 41
            Select Case i
' �㏑�����ځF(6)���O�`(15)�����A(23)���̑�1�`(26)���l
                Case 6 To 15, 23 To 26
                    If wsWrk.Cells(j, i).Value <> "" Then
                        If wsWrk.Cells(j, i).Value <> wsWrk.Cells(j + 1, i).Value Then
                            wsWrk.Cells(j + 2, i).Value = wsWrk.Cells(j, i).Value
                            wsWrk.Cells(j, i).Font.Color = rgbTeal            ' �����F�F��      #808000#
                            wsWrk.Cells(j, i).Interior.Color = rgbLightCoral  ' �w�i�F�F�������� #8080f0#
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
        wsNew.Cells(newYmax, newXmax) = wsWrk.Cells(j + 1, wrkXmax)

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
        
' If newYmax = 884 Then
' Stop
' End If
'
        If sw_change Then
            wsNew.Cells(newYmax, CHECKED_X) = "Modify"
            Cnt.mod = Cnt.mod + 1
        Else
            wsNew.Cells(newYmax, CHECKED_X) = "Add"
            Cnt.Add = Cnt.Add + 1
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
'If p_j = 7 Then
'Stop
'End If
'

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
                    wsWrk.Cells(p_j, x).Font.Strikethrough = True           ' ����������t����
                    wsWrk.Cells(p_j, x).Font.Bold = True                    ' �����ɐݒ�
                    wsWrk.Cells(p_j, x).Font.Color = rgbNavy                ' �����F�F�l�C�r�[  #800000#
                    wsWrk.Cells(p_j, x).Interior.Color = rgbSnow            ' �w�i�F�F�X�m�[    #fafaff#
                    wsWrk.Cells(p_j + 1, xx).Font.Color = rgbNavy              ' �����F�F�l�C�r�[  #800000#
                    wsWrk.Cells(p_j + 1, xx).Interior.Color = rgbSnow          ' �w�i�F�F�X�m�[    #fafaff#
                    sameCnt = sameCnt - 1
                    Exit For
                End If
            Next xx
        End If
    Next x

' �Ⴄ���e�̂��̂��󂢂Ă�Z���ɃR�s�[
    If sameCnt <> 0 Then
        For x = p_from To p_to
            If wsWrk.Cells(p_j, x).Font.Strikethrough = False Then  ' ���������̍��ڂ́A���Ɂ@�o�^�ς�
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
            End If
        Next x
    End If

    If sameCnt <> 0 Then
        MsgBox "�ύX������Ȃ��������ڂ��c���Ă��܂���" & sameCnt
        Stop
    End If

End Sub
        


