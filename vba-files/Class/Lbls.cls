VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lbls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' --------------------------------------+-----------------------------------------
' | @�N���X��   : Lbls   �� C-���x���ꗗ�̕\��`�@/�@�R���N�V����
' --------------------------------------+-----------------------------------------
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' �R���N�b�V�����̒�`
'
Private items_ As Collection
'
'Public Items As Collection

Private Sub Class_Initialize()
    Set items_ = New Collection

    With obj_Lbl
        Dim i As Long: i = 3
        Do While .Cells(i, 1).Value <> ""
            Me.Add .Range(.Cells(i, 1), .Cells(i, 4))
            i = i + 1
        Loop
     End With

End Sub

Public Property Get Item(ByVal key As Variant) As Lbl
     Set Item = items_.Item(key)
     
End Property

Public Function Add(ByVal values As Variant) As Lbl

    Dim p As Lbl: Set p = New Lbl
    p.Initialize values
    items_.Add p, p.LblID
    Set Add = p
    
End Function

Public Sub Remove(ByVal key As Variant)
     items_.Remove key
End Sub

Public Sub ApplyToSheet()

    With obj_Lbl
'        .Cells.ClearContents    ' �����ƒl�̂ݏ���
'        .Range(.Cells(1, 1), .Cells(1, 4)) = Array("���x��ID", "���x����", "�ړ���", "�����q")
'
        .Range(.Cells(3, 1), .Cells(100, 6)).ClearContents  ' �����ƒl�̂ݏ���
        
        Dim i As Long: i = 3
        Dim p As Lbl: Set p = New Lbl
        For Each p In items_
            .Range(.Cells(i, 1), .Cells(i, 4)) = Array(p.LblID, p.���x����, p.�ړ���, p.�����q) ' �l�̂݃R�s�[
            i = i + 1
        Next p
    End With

End Sub