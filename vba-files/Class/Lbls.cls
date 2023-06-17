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
' | @クラス名   : Lbls   ≡ C-ラベル一覧の表定義　/　コレクション
' --------------------------------------+-----------------------------------------
' +   +   +   +   +   +   +   +   +   +   +   +   +   +   x   +   +   +   +   +   +
' コレクッションの定義
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
        .Cells.Clear
        .Range(.Cells(1, 1), .Cells(1, 4)) = Array("ラベルID", "ラベル名", "接頭語", "結合子")
        
        Dim i As Long: i = 3
        Dim p As Lbl: Set p = New Lbl
        For Each p In items_
            .Range(.Cells(i, 1), .Cells(i, 4)) = Array(p.LblID, p.ラベル名, p.接頭語, p.結合子)
            i = i + 1
        Next p
    End With

End Sub
