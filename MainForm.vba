Dim mode As String
Dim currentColumn As Integer
Dim pauseFlag As Boolean
Dim soundFile As String
Dim dirName As String



Private Sub CommandButton1_Click()

    nextAction 1


End Sub

Private Sub CommandButton2_Click()
    
  nextAction 2

End Sub


Private Sub CommandButton3_Click()

    nextAction 3

End Sub

Private Sub CommandButton4_Click()

    nextAction 4

End Sub

Private Sub CommandButton5_Click()

    nextAction 5


End Sub


Dim hp As Integer


Private Sub Label1_Click()
    pauseFlag = False
    CommandButton1.Enabled = True
    CommandButton2.Enabled = True
    CommandButton3.Enabled = True
    CommandButton4.Enabled = True


End Sub


Private Sub UserForm_Activate()
    
    
    hp = 10
    Label2.Caption = hp
    
    
     
    Dim sheet2 As Worksheet
    'データのディレクトリを（1,2）のセルで指定するようにした。
    Set sheet2 = Worksheets("Sheet2")
    dirName = sheet2.Cells(1, 2).Value
    
    
    Debug.Print "UserForm_Activate():" & dirName
    
    'データディレクトリの設定（これは必ずActivateで呼び出すこと
    initDir dirName
    
    
    '現在のモードを決定
    setCureentMode
    '選択肢を表示
    showSelection
    nextAction 0
    
   
    
End Sub

Private Sub UserForm_Click()



End Sub



Private Sub sound(file As String)
    
    If file = "off" Then
        stopSound soundFile
           
    Else
        Dim sheet2 As Worksheet
        Set sheet2 = Worksheets("Sheet2")
        Debug.Print name
        For i = 1 To 100
            If sheet2.Cells(i, 1).Value = file Then
                soundFile = dirName + sheet2.Cells(i, 2).Value
                Exit For
            End If
    Next
        playSound soundFile
    End If

    
    



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


Private Sub pause(pause As String)

    If pause = "on" Then
        pauseFlag = True
        CommandButton1.Enabled = False
        CommandButton2.Enabled = False
        CommandButton3.Enabled = False
        CommandButton4.Enabled = False
        
        
    Else
        pauseFlag = False
    
    End If
    
    Do While pauseFlag
        
        DoEvents 'おまじない
    
    Loop
    
    DoEvents

End Sub

' 吹き出しの表示
Private Sub showBalloon(msg As String)
    Label1.Caption = msg
End Sub

Private Sub showNext(nextMode As String)
    
    If nextMode = "OFF" Then
        
        CommandButton5.Enabled = False
        
    
    ElseIf nextMode = "on" Then
    
        CommandButton5.Enabled = True
    
    Else
    
        mode = nextMode
        UserForm_Activate
    End If

End Sub

'
' idには押された選択肢ボタンのIDが送られてくる
' 縦方向に記述するバージョンに変更(20160502)
'
Sub nextAction(id As Integer)

    Dim actions As Variant

    Dim action As String
    Dim sheet1 As Worksheet

    Dim num As Integer


    Set sheet1 = Worksheets("Sheet1")
    
    action = sheet1.Cells(currentColumn + id, 3)
    If action <> "" Then
        actions = Split(action, ",")
        num = UBound(actions)
    For i = 0 To num Step 2
        If actions(i) = "msg" Then
            MsgBox actions(i + 1)
        ElseIf actions(i) = "next" Then
            showNext CStr(actions(i + 1))
        
        ElseIf actions(i) = "lbl" Then
          showBalloon CStr(actions(i + 1))
                
        ElseIf actions(i) = "img" Then
            showImage CStr(actions(i + 1))
        
        ElseIf actions(i) = "snd" Then
            sound CStr(actions(i + 1))
        
        ElseIf actions(i) = "pause" Then
            pause CStr(actions(i + 1))
        ElseIf actions(i) = "msg2" Then
          showmsg2
        ElseIf actions(i) = "judge" Then
          judge CStr(actions(i + 1))
        
        Else
        
        End If
    Next
    End If

End Sub


Sub judge(num As Integer)


    hp = hp - num
    Label2.Caption = hp
    
    If hp = 0 Then
        MsgBox "脂肪"
    End If
    


End Sub

Sub showmsg2()


    MsgBox "近大！"

End Sub




'
'
'  縦方向に探索するように変更(20160502)
Private Sub setCureentMode()
    
    Dim sheet1 As Worksheet
    
    'スタート時の処理
    If mode = "" Then
        mode = "00"
    End If
    Set sheet1 = Worksheets("Sheet1")
'    With Worksheets("Sheet1")
        For i = 1 To 100
           If sheet1.Cells(i, 2).Value = mode Then
           ' If Worksheets(1).Cells(1, i).Value = mode Then
                currentColumn = i
                Exit For
            End If
        Next
    'End With
End Sub

'選択肢を表示するプログラム
' 縦方向に記述するバージョンに変更(20160502)
Private Sub showSelection()
    With Worksheets("Sheet1")
        CommandButton1.Caption = .Cells(currentColumn + 1, 2).Value
        CommandButton2.Caption = .Cells(currentColumn + 2, 2).Value
        CommandButton3.Caption = .Cells(currentColumn + 3, 2).Value
        CommandButton4.Caption = .Cells(currentColumn + 4, 2).Value
    End With
End Sub

'ウィンドウを閉じる
Private Sub UserForm_Terminate()
    stopSound soundFile 'サウンドを停止する
End Sub
