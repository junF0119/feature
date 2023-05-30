Attribute VB_Name = "getProperty"
Sub getProperty()
    Dim myStr As String
    Dim i As Long
    
    With ThisWorkbook
        '�g���݂̃v���p�e�B(BuiltinDocumentProperties)
        myStr = "[ BuilinDocumentProperties ]" & vbCrLf
        
        On Error Resume Next
        For i = 1 To .BuiltinDocumentProperties.Count
            With .BuiltinDocumentProperties.Item(i)
                myStr = myStr & .Name & ":" & .Value & vbCrLf
            End With
        Next i
        On Error GoTo 0
        
        '���[�U�ݒ�̃v���p�e�B(CustomDocumentProperties)
        myStr = myStr & vbCrLf & _
                "[ CustomDocumentProperties ]" & vbCrLf
        
        For i = 1 To .CustomDocumentProperties.Count
            With .CustomDocumentProperties.Item(i)
                myStr = myStr & .Name & ":" & .Value & vbCrLf
            End With
        Next i
    End With
    
    MsgBox myStr
End Sub
