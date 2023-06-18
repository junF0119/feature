VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' --------------------------------------+-----------------------------------------
' | @function   : Class Lesson 18
' --------------------------------------+-----------------------------------------
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' �A�v���P�[�V�����萔�̒�`
'
Private items_                          As Collection
'

Private Sub Class_Initialize()
    Set items_ = New Collection
    
    Sheets("C-?????").Activate
    
    With Sheets("C-?????")
        Dim i As Long: i = 2
        Do While Cells(i, 1).Value <> ""
            Me.Add .Range(.Cells(i, 1), .Cells(i, 4))
            i = i + 1
        Loop
    End With

End Sub

Public Property Get Item(ByVal key As Variant) As Label

    Set Item = items_.Item(key)

End Property

Public Function Add(ByVal values As Variant) As Label
    Dim p As Label: Set p = New Label
    
'    p.LblID = values(0)
'    p.���O = values(1)
'    p.���� = values(2)
'    p.�a���� = values(3)
    
    p.Initialize values
    
    items_.Add p, p.LblID
    Set Add = p

End Function

Public Sub Remove(ByVal key As Variant)
    items_.Remove key
End Sub


Public Sub ApplyToSheet()

    Sheets("C-?????").Activate
    
    With Sheets("C-?????")
        .Cells.Clear
        .Range(.Cells(1, 1), .Cells(1, 4)) = Array("Id", "���O", "����", "�a����")
        
        Dim i As Long: i = 2
        Dim p As Label
        For Each p In items_
            .Range(.Cells(i, 1), .Cells(i, 4)) = Array(p.LblID, p.���O, p.����, p.�a����)
            i = i + 1
        Next p
    End With

End Sub
