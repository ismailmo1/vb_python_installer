'strappname must == foldername with py files!!
strAppName = "autocert_update"
strAppDesc=  "check new certficates"
'checks if python installed, installs dependencies with pip and creates shortcut to main
sub install()
	Const vbQuote = """"
	set objWso = CreateObject("Wscript.Shell")
	set oFso=CreateObject("Scripting.FileSystemObject")
	set exe = objWso.Exec("cmd /c python --version")

	'check python installation 
	If exe.StdErr.AtEndOfStream Then

		'pip install libraries from requirements.txt (same dir as this script)
		scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)	
		requirementsFile = vbQuote & scriptdir & "\requirements.txt" & vbQuote 		
		
		'cmd prompt to install pip and print when done
		strPip = "cmd /k " & "pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org -r "& requirementsFile & " &&echo.&echo.&echo.all done, you can close this window &echo."

		objWso.Run(strPip)		
		msgbox("don't press OK until the cmd prompt (black window) is all done")

		'create shortcut
		set oShellLink = objWso.CreateShortcut(scriptdir  & "\RUN_" & strAppName & ".lnk")

         	oShellLink.TargetPath = (scriptdir & "\" & strAppName & "\main.pyw")

         	oShellLink.Description = strAppDesc
         	oShellLink.WorkingDirectory = scriptdir & "\" & strAppName
         	oShellLink.Save

		
	Else
		Msgbox("python installation not found" & vbcrlf _
			& "download from python.org/downloads and try this again")
	End If
	
end sub

Call install