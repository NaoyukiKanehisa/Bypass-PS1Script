'==============================================================================================
'PowerShell�̎��s�|���V�[�Ɋ֌W�Ȃ��APowerShell�̃X�N���v�g�t�@�C��(.ps1)�����s����B
'[�g����]
'  �{VBScript�t�@�C��(.vbs)��PowerShell�̃X�N���v�g�t�@�C��(.ps1)���h���b�O&�h���b�v����B
'  �܂��́A�{VBScript�t�@�C��(.vbs)��PowerShell�̃X�N���v�g�t�@�C��(.ps1)�𓯈�f�B���N�g����
'  �z�u���A����t�@�C����(���g���q������)�ɂ���VBScript�t�@�C��(.vbs)�����s�B
'  �{�X�N���v�g���s���Ɉ������w�肷��ƁAPS1�X�N���v�g���֎��s�����̎󂯓n�����\�B
'   (���h���b�O&�h���b�v���͈����̎󂯓n���͕s��)

'PowerShell�E�C���h�E�\��(True/False)
displaywindow = True

'�h���b�O&�h���b�v���p(True/False)
dragdrop = False
'==============================================================================================

Set WshShell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

If (dragdrop = True) And (Wscript.Arguments.count > 0) Then
	ps1filename = Wscript.Arguments(0)
	Command = "powershell.exe -sta -WindowStyle Normal -ExecutionPolicy Bypass -File " & """" & ps1filename & """"
Else
	ps1filename = (Fso.BuildPath(Fso.GetParentFolderName(WScript.ScriptFullName),(Fso.GetBaseName(WScript.ScriptName) & ".ps1")))
	Command = "powershell.exe -sta -WindowStyle Normal -ExecutionPolicy Bypass -File " & """" & ps1filename & """"
	If Wscript.Arguments.count > 0 Then
		For i = 0 To (Wscript.Arguments.count -1)
			Command = Command & " """ & Wscript.Arguments(i) & """"
		Next
	End If
End If

If displaywindow = True Then
	WshShell.Run(Command),1,True
Else
	WshShell.Run(Command),0,True
End If

Set WshShell = Nothing
Set Fso = Nothing