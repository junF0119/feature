Attribute VB_Name = "MySub2"
Option Explicit
'
' ---Procedure Division ----------------+-----------------------------------------
'
Sub MySub2()

    Dim myLabels As Labels: Set myLabels = New Labels
    
    Dim p As Label
        
    With myLabels
        .Add (Array(3, "����@���l", "�j", #8/21/1990#))
        .Remove 1
  
        .ApplyToSheet
    End With
    
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
    Stop
    
End Sub



