Attribute VB_Name = "MySub"
Option Explicit
'
' ---Procedure Division ----------------+-----------------------------------------
'
Sub MySub()
    
    Dim i As Long: i = 2
    Dim myLabels As Collection: Set myLabels = New Collection
    
'
' ---Procedure Division ----------------+-----------------------------------------
'
   
    Sheets("C-?????").Activate
    
    With Sheets("C-?????")
        Do While Cells(i, 1).Value <> ""
            Dim p As Label: Set p = New Label
            p.Initialize .Range(.Cells(i, 1), .Cells(i, 4))
            myLabels.Add p, p.LblID
            i = i + 1
        Loop
    End With
    
    Debug.Print myLabels(2).名前
    Debug.Print myLabels(1).誕生日
    
' コレクションの要素についてループ
    For Each p In myLabels
        Debug.Print p.LblID, p.名前, p.性別, p.誕生日
    Next p
    
    
'        p.Id = .Cells(I, 1).Value
'        p.名前 = .Cells(I, 2).Value
'        p.性別 = .Cells(I, 3).Value
'        p.誕生日 = .Cells(I, 4).Value
    
    
'    Debug.Print p.Id
    
    
'    p.Id = "hoge"
'    p.性別 = "F"
    Stop
'    p.Greet
    
End Sub


