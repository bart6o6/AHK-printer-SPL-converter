#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#singleinstance force

SetTitleMatchMode,2 ;start of obtaining WinClass of Comarch. It changes often
WinGetClass,WinGetClassVariable, Comarch ERP XL,
VariableComarchClassName= % SubStr(WinGetClassVariable, 7, 8) 
;MsgBox, %VariableComarchClassName%^ ;end of obtaining WinClass of Comarch. Uncomment msgBox for details

MainMenuGuiClose(){
	ExitApp 
	}
/*-------------------------------------------------------------------------------------------------------
*/
Esc::
ExitApp


^r::
FileCreateDir, %a_scriptdir%\converted
FileCreateDir, %a_scriptdir%\converted\source
FileCreateDir, %a_scriptdir%\converted\pdf
FileCreateDir, %a_scriptdir%\converted\txt

sleep, 300
FileList := ""
Loop, Files, %a_scriptdir%/to convert/*.SPL
{
	FileList .= A_LoopFileLongPath "`n"
	
	;MsgBox, %name_no_ext%	
}	


Loop, Parse, FileList, `n,`r 
{
	Progress, %A_Index%, %A_LoopFileName%, Konwertowanie..., Konwertowanie
	if(A_LoopField = "")
	{
		msgBox, sukces
		return
	}
	
	
	;if(A_LoopField!="WYNIKI BADAN SKLADU CHEMICZNEGO")
	;{
	;	continue
	;}
	
	else
	{
		FileGetTime, DateFileCreated, %A_LoopField%, M
		;FileAppend, %A_LoopField% `n, C:\Users\bartek.gola\Desktop\Nowy folder 22222\list.txt
		;FileAppend, %name_no_ext% `n, C:\Users\bartek.gola\Desktop\Nowy folder 22222\list.txt
		SplitPath, A_LoopField, name, dir, ext, name_no_ext
		
		Run, cmd.exe,,
		WinWait, cmd.exe,,10
		sleep, 300
		;sendInput, cd %a_scriptdir%\SPLView\ {Enter} 
		;C:\Users\bartek.gola\Desktop\Nowy folder 22222\SPLView
		;sendInput, cd C:\Program Files (x86)\SPLView\ {Enter}
		send, "%a_scriptdir%\lib\SPLView\splview.exe" /pt "%A_LoopField%" "Microsoft Print to PDF" && exit 
		sleep, 200
		send, {Enter}
		WinWait, Zapisywanie wydruku jako,,10
		sleep, 500
		send, %a_scriptdir%\converted\pdf\%DateFileCreated%_%name_no_ext%.PDF
		sleep, 200
		sendInput, {Enter}
		sleep, 200
		
		;if(WinExist(Confirm Save As))
		;{
		;	WinActivate, Confirm Save As
		;	WinClose, Confirm Save As
		;	WinActivate, Confirm Save As
		;	Send, {Enter}
		;	WinClose, Zapisywanie wydruku jako
		;	WinActivate, SPL-Viewer
		;	Send, {Enter}
		;	WinClose, SPL-Viewer
		;	sleep, 200
		;}
		;ControlSend, cmd.exe, "%a_scriptdir%\lib\xpdf-tools-win-4.04\bin32\pdftotext.exe" -raw "%a_scriptdir%\converted\pdf\%DateFileCreated%_%name_no_ext%.PDF" "%a_scriptdir%\converted\txt\%DateFileCreated%_%name_no_ext%.txt"
		
		
		Run, cmd.exe,,
		WinWait, cmd.exe,,10
		sleep, 1500
		send, "%a_scriptdir%\lib\xpdf-tools-win-4.04\bin32\pdftotext.exe" -raw "%a_scriptdir%\converted\pdf\%DateFileCreated%_%name_no_ext%.PDF" "%a_scriptdir%\converted\txt\%DateFileCreated%_%name_no_ext%.txt" && exit
		sleep, 200
		send, {Enter}
		WinWaitClose, cmd.exe
		
		;sleep, 1000
		;WinClose, cmd.exe
		
		;sleep, 500
		;FileRead, content, %a_scriptdir%\converted\txt\%DateFileCreated%_%name_no_ext%.txt
		;sleep, 1000
		;loop, parse, Content, `n,`r
		;	
		;{
		;	;if( (InStr("%a_loopfield%", "WYNIKI BADA") != 1))
		;	;{
		;	;	MsgBox, dupa
		;	;	Continue
		;	;}
		;	if( (InStr("%a_loopfield%", "Sr.") = 1))
		;	{
		;		StringReplace, NewStr, "%a_loopfield%", "Sr.", "", All
		;		x=%NewStr%
		;		e .= x  . ";"
		;		Continue
		;	}
		;	x=%a_loopfield%
		;	e .= x . ";"
		;}
		;sleep, 1000
		;fileappend,%DateFileCreated%;%e%`n, zestawienie_wydrukow.csv
		;e=
	}
	
	FileMove, %A_LoopField%, %a_scriptdir%\converted\source
	FileMove, %a_scriptdir%\converted\source\%name_no_ext%.SPL, %a_scriptdir%\converted\source\%DateFileCreated%_%name_no_ext%.SPL ;rename file
	FileDelete, %a_scriptdir%/to convert/%name_no_ext%.SHD
	sleep, 100
}
until A_LoopField = ""