On Error Resume Next

Dim installPath
Dim addInName
Dim addInFileName
Dim addInRegisterName
Dim objPowerPoint
Dim objAddin

'アドイン情報を設定
addInName = "PowerPoint配布資料作成"
addInFileName = "Menu.ppam"

IF MsgBox(addInName & " アドインをアンインストールしますか？", vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End IF

' アドインの登録名取得
addInRegisterName = Mid(addInFileName, 1, Len(addInFileName) - 5)

' PowerPointインスタンス化
Set objPowerPoint = CreateObject("PowerPoint.Application")

' アドイン登録解除
For i = 1 To objPowerPoint.Addins.Count
  Set objAddin = objPowerPoint.Addins.item(i)
  
  If objAddin.Name = addInRegisterName Then
    objAddin.AutoLoad = False
    objPowerPoint.Addins.Remove addInRegisterName
  End If
Next

' PowerPoint 終了
objPowerPoint.Quit

Set objAddin = Nothing
Set objAddins = Nothing
Set objPowerPoint = Nothing

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'インストール先パスの作成
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'ファイル削除
If objFileSys.FileExists(installPath) = True Then
  objFileSys.DeleteFile installPath , True
Else
  MsgBox "アドインファイルが存在しません。", vbExclamation
End If

Set objWshShell = Nothing
Set objFileSys = Nothing

IF Err.Number = 0 THEN
   MsgBox "アドインのアンインストールが終了しました。", vbInformation
ELSE
   MsgBox "エラーが発生しました。" & vbCrLF & "実行環境を確認してください。", vbExclamation
End IF