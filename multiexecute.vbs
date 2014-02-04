Sub RunApplication(ByVal sFile)

    Dim WShell : Set WShell = CreateObject("WScript.Shell")
	WShell.Run  sFile , 8, false

End Sub

RunApplication "iisLogBackup.vbs weblog_201312.zip u_ex1312*.log"
RunApplication "iisLogBackup.vbs weblog_201401.zip u_ex1401*.log"