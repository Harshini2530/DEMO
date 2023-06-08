Function TimedBox(boxMessage,time,title)
Dim tempfolder, tempfile,objFSO
set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSHL = CreateObject("Wscript.Shell")
set tempfolder = objFSO.GetSpecialFolder(2)
tempfile = objFSO.Gettempname() &".hta"
With tempfolder.CreateTextFile(tempfile) 
	.Writeline "<html><HTA:APPLICATION sysMenu=""no""Scroll=""no"" Border=""dialog"">"&"<head><title>"& title &"</title><script language = ""VBScript"">"
	.Writeline "Sub Window_OnLoad"
	.Writeline "Window.moveTo 800,300"
	.Writeline "window.resizeTo 500,150"
	.Writeline "idTimer = window.setTimeout(""pausedSection""," & time & ",""VBScript"")"
	.Writeline "End Sub"
	.Writeline "Sub pausedSection"
	.Writeline "window.close"
	.Writeline "End Sub"
	.Writeline "</script></head><body><p align=""center""> "& boxmessage & "</p></body></html>"
	.Close
end With
objSHL.Run tempFolder & "\" & tempFile, 1, True
objFSO.DeleteFile tempFolder & "\" & tempFile
END Function

'TimedBox " Auto Select 043.25 should be executed. " ,3000,"AutoSelect"
CreateObject("WScript.Shell").Popup "Auto Select 043.25 should be executed!", 3, "AutoSelect"