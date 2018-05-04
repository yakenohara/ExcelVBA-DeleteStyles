Attribute VB_Name = "�s�v�X�^�C���S�폜"
Sub �s�v�X�^�C���S�폜()
    
    '�ϐ��錾
    Dim styleNames() As String
    Dim saveStyles() As Variant
    Dim haveToDelete As Boolean
    Dim saveStyle As Variant
    
    '�����l�ݒ�
    saveStyles = Array("Hyperlink", _
                       "Normal", _
                       "Followed Hyperlink")
    
    
    Application.ScreenUpdating = False
    
    numberOfStyles = ActiveWorkbook.Styles.Count
    
    ReDim styleNames(numberOfStyles)
    
    For i = 1 To numberOfStyles
        styleNames(i) = ActiveWorkbook.Styles(i).Name
    Next
    
    For i = 1 To numberOfStyles
        
        '�i���\��
        Application.StatusBar = "Progress:" & i & "/" & numberOfStyles
        
        haveToDelete = True
        For Each saveStyle In saveStyles
            
            If saveStyle = styleNames(i) Then '�c�������X�^�C���w��ł�������
                haveToDelete = False
                Exit For
            End If
            
        Next saveStyle
        
        If haveToDelete Then
            Call deleteStyle(ActiveWorkbook, styleNames(i))
        End If
        
    Next
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Done!"

End Sub

Private Sub deleteStyle(ByVal bk As Workbook, ByRef styleName As String)
    
    On Error GoTo whenDelMethodFailed
    
    bk.Styles(styleName).Delete
    Exit Sub

'�uStyle�N���X��Delete���\�b�h�����s���܂����v�̎�
whenDelMethodFailed:
    Exit Sub
End Sub
