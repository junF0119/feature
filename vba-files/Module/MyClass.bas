Attribute VB_Name = "MyClass"
Option Explicit
'
 '�V�[�g�̃I�u�W�F�N�g��: Sheet1��obj_Lbl �ɕύX

' ---Procedure Division ----------------+-----------------------------------------
'
Sub MyClass()
    
''��03
'    Dim p As Lbl
'    Set p = New Lbl
'
'    p.���x���� = "����"
'
    
''��04
'    Dim p As Lbl
'    Set p = New Lbl
'
'    p.LblID = 99
'    p.���x���� = "����"
'    p.Greet
'
'
''��05   ��s���̃f�[�^���N���X������
'    Dim p As Lbl: Set p = New Lbl
'
'    With Sheets("C-���x���ꗗ")
'        p.LblID = .Cells(3, 1).Value
'        p.���x���� = .Cells(3, 2).Value
'        p.�ړ��� = .Cells(3, 3).Value
'        p.�����q = .Cells(3, 4).Value
'    End With
'
''��06
'    Dim p As Lbl: Set p = New Lbl
'    Dim i As Long: i = 3
'    With Sheets("C-���x���ꗗ")
'        p.LblID = .Cells(i, 1).Value
'        p.���x���� = .Cells(i, 2).Value
'        p.�ړ��� = .Cells(i, 3).Value
'        p.�����q = .Cells(i, 4).Value
'        p.Is���x���� = (p.���x���� = "")
'    End With
'
'Stop
'    p.���x���� = "ABC"
'
'Stop
    
''��08
'    Dim p As Lbl: Set p = New Lbl
'
'    Dim i As Long: i = 3
'    With Sheets("C-���x���ꗗ")
'        p.LblID = .Cells(i, 1).Value
'        p.���x���� = .Cells(i, 2).Value
'        p.�ړ��� = .Cells(i, 3).Value
'        p.�����q = .Cells(i, 4).Value
'        p.Is���x���� = (p.���x���� = "")
'    End With
'
''    p.LblID = 99
'Stop

''��09
'    Dim p As Lbl: Set p = New Lbl
'
'    Dim i As Long: i = 5
'    With obj_Lbl             ' obj_Lbl �� �V�[�g��object��
'        p.LblID = .Cells(i, 1).Value
'        p.���x���� = .Cells(i, 2).Value
'        p.�ړ��� = .Cells(i, 3).Value
'        p.�����q = .Cells(i, 4).Value
'    End With
'
'    Debug.Print p.LblID
'    Stop

''��10 �����s�̃f�[�^���N���X�ň���
'    Dim i As Long: i = 5
'    Dim p As Lbl: Set p = New Lbl
'    With obj_Lbl             ' obj_Lbl �� �V�[�g��object��
'        p.Initialize .Range(.Cells(i, 1), .Cells(i, 4))
'    End With
'
'    Debug.Print p.LblID
'    Stop

''��11 �C���X�^���X�̏W�����R���N�V����������
'    Dim myLbls As Collection: Set myLbls = New Collection
'
'    Dim i As Long: i = 3
'
'    With obj_Lbl             ' obj_Lbl �� �V�[�g��object��
'        Do While .Cells(i, 1).Value <> ""
'            Dim p As Lbl: Set p = New Lbl
'            p.Initialize .Range(.Cells(i, 1), .Cells(i, 4))
'            myLbls.Add p, p.LblID
'            i = i + 1
'        Loop
'    End With
'
''    Debug.Print p.LblID
''    Stop
'
''��12
'    Debug.Print myLbls(5).���x����      ' �C���f�b�N�X�ł̎Q��
'    Debug.Print myLbls("10").���x����   ' �L�[�ł̎Q��
''�R���N�V�����̗v�f�ɂ��ă��[�v
'    For Each p In myLbls
'        Debug.Print p.LblID, p.���x����, p.�ړ���, p.�����q
'    Next p
'    Stop
    
''��13
'    Dim myLbls As Lbls: Set myLbls = New Lbls
'    Set myLbls.Items = New Collection
'
'    With obj_Lbl
'        Dim i As Long: i = 5
'        Do While .Cells(i, 1).Value <> ""
'            Dim p As Lbl: Set p = New Lbl
'            p.Initialize .Range(.Cells(i, 1), .Cells(i, 4))
'            myLbls.Items.Add p, p.LblID
'            i = i + 1
'        Loop
'    End With
'
'    Stop
    
''��14
'    Dim myLbls As Lbls: Set myLbls = New Lbls
''    Set myLbls.Items = New Collection
'
'    With obj_Lbl
'        Dim i As Long: i = 5
'        Do While .Cells(i, 1).Value <> ""
'            Dim p As Lbl: Set p = New Lbl
'            p.Initialize .Range(.Cells(i, 1), .Cells(i, 4))
'            myLbls.Items.Add p, p.LblID
'            i = i + 1
'        Loop
'    End With
'
'    Stop
    
''��15
'
'    Dim myLbls As Lbls: Set myLbls = New Lbls
'
'    Stop
    
''��16
'    Dim myLbls As Lbls: Set myLbls = New Lbls
'
'    With myLbls
'        Debug.Print .Item(1).���x����
'        .Item("3").Greet
'    End With
'    Stop
'
''��17
'    Dim myLbls As Lbls: Set myLbls = New Lbls
'
'    With myLbls
'        Debug.Print .Item(1).���x����
'        .Item("3").Greet
'    End With
'    Stop

''��18
'    Dim myLbls As Lbls: Set myLbls = New Lbls
'    Dim p As Lbl
'    Set p = myLbls.Add(Array("100", "abc", "XYZ", "-"))
'
'    With myLbls
'        .Remove 5
'        .Remove 10
'    End With
'    Stop

'��19
    Dim myLbls As Lbls: Set myLbls = New Lbls
    
    With myLbls
        .Add (Array("100", "abc", "XYZ", "-"))
        .Remove 5
        With .Item("10")
            .���x���� = "10.�����ђ�"
            .�ړ��� = "��"
            .�����q = "-"
        End With
        
        .ApplyToSheet
        
    End With
    
    Stop






Stop
    


    Stop

'' �R���N�V�����̗v�f�ɂ��ă��[�v
'    For Each p In myLabels
'        Debug.Print p.LblID, p.���O, p.����, p.�a����
'    Next p
    
    
'        p.Id = .Cells(I, 1).Value
'        p.���O = .Cells(I, 2).Value
'        p.���� = .Cells(I, 3).Value
'        p.�a���� = .Cells(I, 4).Value
    
    
'    Debug.Print p.Id
    
    
'    p.Id = "hoge"
'    p.���� = "F"
'    Stop
'    p.Greet
    

'
'    With myLabels
'        .Add (Array(58, "58.����", 58, "����"))
'        .Remove 1
'        .Add (Array(90, "90.Help", 90, "Help"))
'
'        .ApplyToSheet
'    End With
    
'    With myLabels
'        Debug.Print .Item(1).���O
'        .Item(2).Greet
'    End With
    
'
' ---Procedure Division ----------------+-----------------------------------------
'
'    Set myLabels.Items = New Collection
'    Sheets("C-���x���ꗗ").Activate
'
'    With Sheets("C-���x���ꗗ")
'        Dim i As Long: i = 2
'        Do While Cells(i, 1).Value <> ""
'            Dim p As Label: Set p = New Label
'            p.Initialize .Range(.Cells(i, 1), .Cells(i, 4))
'            myLabels.Items.add p, p.LblID
'            i = i + 1
'        Loop
'    End With



    
End Sub




