Attribute VB_Name = "不要スタイル全削除"
Sub 不要スタイル全削除()
    
    '変数宣言
    Dim styleNames() As String
    Dim saveStyles() As Variant
    Dim haveToDelete As Boolean
    Dim saveStyle As Variant
    
    '初期値設定
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
        
        '進捗表示
        Application.StatusBar = "Progress:" & i & "/" & numberOfStyles
        
        haveToDelete = True
        For Each saveStyle In saveStyles
            
            If saveStyle = styleNames(i) Then '残したいスタイル指定であったら
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

'「StyleクラスのDeleteメソッドが失敗しました」の時
whenDelMethodFailed:
    Exit Sub
End Sub
