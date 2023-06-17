Attribute VB_Name = "MyClass"
Option Explicit
'
 'シートのオブジェクト名: Sheet1⇒obj_Lbl に変更

' ---Procedure Division ----------------+-----------------------------------------
'
Sub MyClass()
    
''§03
'    Dim p As Lbl
'    Set p = New Lbl
'
'    p.ラベル名 = "分類"
'
    
''§04
'    Dim p As Lbl
'    Set p = New Lbl
'
'    p.LblID = 99
'    p.ラベル名 = "分類"
'    p.Greet
'
'
''§05   一行文のデータをクラス化する
'    Dim p As Lbl: Set p = New Lbl
'
'    With Sheets("C-ラベル一覧")
'        p.LblID = .Cells(3, 1).Value
'        p.ラベル名 = .Cells(3, 2).Value
'        p.接頭語 = .Cells(3, 3).Value
'        p.結合子 = .Cells(3, 4).Value
'    End With
'
''§06
'    Dim p As Lbl: Set p = New Lbl
'    Dim i As Long: i = 3
'    With Sheets("C-ラベル一覧")
'        p.LblID = .Cells(i, 1).Value
'        p.ラベル名 = .Cells(i, 2).Value
'        p.接頭語 = .Cells(i, 3).Value
'        p.結合子 = .Cells(i, 4).Value
'        p.Isラベル名 = (p.ラベル名 = "")
'    End With
'
'Stop
'    p.ラベル名 = "ABC"
'
'Stop
    
''§08
'    Dim p As Lbl: Set p = New Lbl
'
'    Dim i As Long: i = 3
'    With Sheets("C-ラベル一覧")
'        p.LblID = .Cells(i, 1).Value
'        p.ラベル名 = .Cells(i, 2).Value
'        p.接頭語 = .Cells(i, 3).Value
'        p.結合子 = .Cells(i, 4).Value
'        p.Isラベル名 = (p.ラベル名 = "")
'    End With
'
''    p.LblID = 99
'Stop

''§09
'    Dim p As Lbl: Set p = New Lbl
'
'    Dim i As Long: i = 5
'    With obj_Lbl             ' obj_Lbl ≡ シートのobject名
'        p.LblID = .Cells(i, 1).Value
'        p.ラベル名 = .Cells(i, 2).Value
'        p.接頭語 = .Cells(i, 3).Value
'        p.結合子 = .Cells(i, 4).Value
'    End With
'
'    Debug.Print p.LblID
'    Stop

''§10 複数行のデータをクラスで扱う
'    Dim i As Long: i = 5
'    Dim p As Lbl: Set p = New Lbl
'    With obj_Lbl             ' obj_Lbl ≡ シートのobject名
'        p.Initialize .Range(.Cells(i, 1), .Cells(i, 4))
'    End With
'
'    Debug.Print p.LblID
'    Stop

''§11 インスタンスの集合をコレクション化する
'    Dim myLbls As Collection: Set myLbls = New Collection
'
'    Dim i As Long: i = 3
'
'    With obj_Lbl             ' obj_Lbl ≡ シートのobject名
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
''§12
'    Debug.Print myLbls(5).ラベル名      ' インデックスでの参照
'    Debug.Print myLbls("10").ラベル名   ' キーでの参照
''コレクションの要素についてループ
'    For Each p In myLbls
'        Debug.Print p.LblID, p.ラベル名, p.接頭語, p.結合子
'    Next p
'    Stop
    
''§13
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
    
''§14
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
    
''§15
'
'    Dim myLbls As Lbls: Set myLbls = New Lbls
'
'    Stop
    
''§16
'    Dim myLbls As Lbls: Set myLbls = New Lbls
'
'    With myLbls
'        Debug.Print .Item(1).ラベル名
'        .Item("3").Greet
'    End With
'    Stop
'
''§17
'    Dim myLbls As Lbls: Set myLbls = New Lbls
'
'    With myLbls
'        Debug.Print .Item(1).ラベル名
'        .Item("3").Greet
'    End With
'    Stop

''§18
'    Dim myLbls As Lbls: Set myLbls = New Lbls
'    Dim p As Lbl
'    Set p = myLbls.Add(Array("100", "abc", "XYZ", "-"))
'
'    With myLbls
'        .Remove 5
'        .Remove 10
'    End With
'    Stop

'§19
    Dim myLbls As Lbls: Set myLbls = New Lbls
    
    With myLbls
        .Add (Array("100", "abc", "XYZ", "-"))
        .Remove 5
        With .Item("10")
            .ラベル名 = "10.桜美林中"
            .接頭語 = "中"
            .結合子 = "-"
        End With
        
        .ApplyToSheet
        
    End With
    
    Stop






Stop
    


    Stop

'' コレクションの要素についてループ
'    For Each p In myLabels
'        Debug.Print p.LblID, p.名前, p.性別, p.誕生日
'    Next p
    
    
'        p.Id = .Cells(I, 1).Value
'        p.名前 = .Cells(I, 2).Value
'        p.性別 = .Cells(I, 3).Value
'        p.誕生日 = .Cells(I, 4).Value
    
    
'    Debug.Print p.Id
    
    
'    p.Id = "hoge"
'    p.性別 = "F"
'    Stop
'    p.Greet
    

'
'    With myLabels
'        .Add (Array(58, "58.旅館", 58, "旅館"))
'        .Remove 1
'        .Add (Array(90, "90.Help", 90, "Help"))
'
'        .ApplyToSheet
'    End With
    
'    With myLabels
'        Debug.Print .Item(1).名前
'        .Item(2).Greet
'    End With
    
'
' ---Procedure Division ----------------+-----------------------------------------
'
'    Set myLabels.Items = New Collection
'    Sheets("C-ラベル一覧").Activate
'
'    With Sheets("C-ラベル一覧")
'        Dim i As Long: i = 2
'        Do While Cells(i, 1).Value <> ""
'            Dim p As Label: Set p = New Label
'            p.Initialize .Range(.Cells(i, 1), .Cells(i, 4))
'            myLabels.Items.add p, p.LblID
'            i = i + 1
'        Loop
'    End With



    
End Sub




