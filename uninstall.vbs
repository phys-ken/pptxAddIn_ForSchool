On Error Resume Next

Dim installPath
Dim addInName
Dim addInFileName
Dim addInRegisterName
Dim objPowerPoint
Dim objAddin

'�A�h�C������ݒ�
addInName = "PowerPoint�z�z�����쐬"
addInFileName = "Menu.ppam"

IF MsgBox(addInName & " �A�h�C�����A���C���X�g�[�����܂����H", vbYesNo + vbQuestion) = vbNo Then
  WScript.Quit
End IF

' �A�h�C���̓o�^���擾
addInRegisterName = Mid(addInFileName, 1, Len(addInFileName) - 5)

' PowerPoint�C���X�^���X��
Set objPowerPoint = CreateObject("PowerPoint.Application")

' �A�h�C���o�^����
For i = 1 To objPowerPoint.Addins.Count
  Set objAddin = objPowerPoint.Addins.item(i)
  
  If objAddin.Name = addInRegisterName Then
    objAddin.AutoLoad = False
    objPowerPoint.Addins.Remove addInRegisterName
  End If
Next

' PowerPoint �I��
objPowerPoint.Quit

Set objAddin = Nothing
Set objAddins = Nothing
Set objPowerPoint = Nothing

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'�C���X�g�[����p�X�̍쐬
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'�t�@�C���폜
If objFileSys.FileExists(installPath) = True Then
  objFileSys.DeleteFile installPath , True
Else
  MsgBox "�A�h�C���t�@�C�������݂��܂���B", vbExclamation
End If

Set objWshShell = Nothing
Set objFileSys = Nothing

IF Err.Number = 0 THEN
   MsgBox "�A�h�C���̃A���C���X�g�[�����I�����܂����B", vbInformation
ELSE
   MsgBox "�G���[���������܂����B" & vbCrLF & "���s�����m�F���Ă��������B", vbExclamation
End IF