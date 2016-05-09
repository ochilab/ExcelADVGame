Dim dirName As String
Dim hp As Integer

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
(ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'データファイルのあるディレクトリを設定
Sub initDir(dir As String)
    dirName = dir
    Debug.Print "init()"
    
End Sub

'指定したファイルのサウンド再生をストップ
'soundFileはファイル名のみ（ディレクトリ名は与えない）
Sub stopSound(soundFile As String)

  Debug.Print "stopSound() " & dirName & soundFile
    mciSendString "Stop " & soundFile, "", 0, 0

End Sub

'指定したファイルのサウンドを再生
'soundFileはファイル名のみ（ディレクトリ名は与えない）
Sub playSound(soundFile As String)
    Debug.Print "playSound() " & dirName & soundFile
    If dir(soundFile) = "" Then
        MsgBox soundFile & vbCrLf & "がありません。", vbExclamation
        Exit Sub
    End If
    rc = mciSendString("Play " & dirName & soundFile, "", 0, 0)


End Sub
