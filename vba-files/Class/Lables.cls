VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Labels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private items_ As Collection
'

Private Sub Class_Initialize()
    Set items_ = New Collection
    
    With Sheet1
        Dim i As Long: i = 2
        Do While .Cells(i, 1).Value <> ""
            Me.Add .Range(.Cells(i, 1), .Cells(i, 4))
            i = i + 1
        Loop
     End With
     
End Sub

Public Property Get Item(ByVal lblID As Variant) As Label
     Set Item = items_.Item(lblID)
     
End Property

Public Function Add(ByVal values As Variant) As Label

    Dim p As Lables: Set p = New Lablen
    p.Initialize values
    items_.Add p, p.LblID

    Set Add = p
    
End Function

Public Sub Remove(ByVal lblID As Variant)
     items_.Remove lblID
End Sub

Public Sub ApplyToSheet()

    Sheets("C-ラベル一覧").Activate
    
    With Sheets("C-ラベル一覧")
        .Cells.Clear
        .Range(.Cells(1, 1), .Cells(1, 4)) = Array("ラベルID", "ラベル名", "接頭語", "結合子")
        
        Dim i As Long: i = 2
        Dim p As Label
        For Each p In items_
            .Range(.Cells(i, 1), .Cells(i, 4)) = Array(p.lblID, p.ラベル名, p.接頭語, p.結合子)
            i = i + 1
        Next p
    End With

End Sub
