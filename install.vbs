On Error Resume Next

Dim installPath
Dim addInName
Dim addInFileName
Dim objPowerPoint
Dim objAddin

'�A�h�C������ݒ�
addInName = "PowerPoint�z�z�����쐬"
addInFileName = "Menu.ppam" 

IF MsgBox(addInName & " �A�h�C�����C���X�g�[�����܂����H", vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End IF

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'�C���X�g�[����p�X�̍쐬
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'�t�@�C���R�s�[(�㏑��)
objFileSys.CopyFile  addInFileName ,installPath , True

Set objWshShell = Nothing
Set objFileSys = Nothing


' PowerPoint �C���X�^���X��
Set objPowerPoint = CreateObject("PowerPoint.Application")

' �A�h�C���o�^
Set objAddin = objPowerPoint.AddIns.Add(installPath)
objAddin.AutoLoad = True

' PowerPoint �I��
objPowerPoint.Quit

Set objAddin = Nothing
Set objPowerPoint = Nothing

IF Err.Number = 0 THEN
   MsgBox "�A�h�C���̃C���X�g�[�����I�����܂����B", vbInformation
ELSE
   MsgBox "�G���[���������܂����B" & vbCrLF & "���s�����m�F���Ă��������B", vbExclamation
End IF