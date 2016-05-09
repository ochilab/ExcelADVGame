Public mode As String
Dim soundFile As String, rc As Long


Private Sub CommandButton1_Click()
   
   stopSound soundFile
    Me.Hide
    
    UserForm1.Show
    

End Sub

Private Sub UserForm_Activate()
    
    Dim dirName As String
    Dim sheet2 As Worksheet
    
   'データのディレクトリを（1,2）のセルで指定するようにした。
   Set sheet2 = Worksheets("Sheet2")
   dirName = sheet2.Cells(1, 2).Value
   
    soundFile = dirName & "\opening.mp3"
    playSound soundFile
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()



End Sub

Sub showImage(name As String)

    Dim sheet2 As Worksheet
    Set sheet2 = Worksheets("Sheet2")
    Debug.Print name
    For i = 1 To 100
        If sheet2.Cells(i, 1).Value = name Then
            DoEvents
        
            Image1.Picture = Nothing
            
            Image1.Picture = LoadPicture(dirName + sheet2.Cells(i, 2).Value)
            DoEvents
            Exit For
        End If
    Next


End Sub
'ウィンドウを閉じる
Private Sub UserForm_Terminate()
    stopSound soundFile 'サウンドを停止する
End Sub
