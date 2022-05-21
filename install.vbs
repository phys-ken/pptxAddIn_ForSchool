On Error Resume Next

Dim installPath
Dim addInName
Dim addInFileName
Dim objPowerPoint
Dim objAddin

'アドイン情報を設定
addInName = "PowerPoint配布資料作成"
addInFileName = "Menu.ppam" 

IF MsgBox(addInName & " アドインをインストールしますか？", vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End IF

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'インストール先パスの作成
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'ファイルコピー(上書き)
objFileSys.CopyFile  addInFileName ,installPath , True

Set objWshShell = Nothing
Set objFileSys = Nothing


' PowerPoint インスタンス化
Set objPowerPoint = CreateObject("PowerPoint.Application")

' アドイン登録
Set objAddin = objPowerPoint.AddIns.Add(installPath)
objAddin.AutoLoad = True

' PowerPoint 終了
objPowerPoint.Quit

Set objAddin = Nothing
Set objPowerPoint = Nothing

IF Err.Number = 0 THEN
   MsgBox "アドインのインストールが終了しました。", vbInformation
ELSE
   MsgBox "エラーが発生しました。" & vbCrLF & "実行環境を確認してください。", vbExclamation
End IF