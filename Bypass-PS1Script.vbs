'==============================================================================================
'PowerShellの実行ポリシーに関係なく、PowerShellのスクリプトファイル(.ps1)を実行する。
'[使い方]
'  本VBScriptファイル(.vbs)にPowerShellのスクリプトファイル(.ps1)をドラッグ&ドロップする。
'  または、本VBScriptファイル(.vbs)とPowerShellのスクリプトファイル(.ps1)を同一ディレクトリに
'  配置し、同一ファイル名(※拡張子を除く)にしてVBScriptファイル(.vbs)を実行。
'  本スクリプト実行時に引数を指定すると、PS1スクリプト内へ実行引数の受け渡しが可能。
'   (※ドラッグ&ドロップ時は引数の受け渡しは不可)

'PowerShellウインドウ表示(True/False)
displaywindow = True

'ドラッグ&ドロップ利用(True/False)
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