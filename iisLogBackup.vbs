	'*****************************************************************************************************
	'	Class ���� �� �ɼ� ����	
	'*****************************************************************************************************
	Option Explicit
	Dim oCls
	SET oCls = New iisLogBackup


		'** iisLogBackup Ŭ���� ���ο��� �̷����� ���μ����� �α� ����Ʈ�� ����ϴ�
		'** 1 : �α� ����Ʈ ���� / 0 : �α� ����Ʈ ������ ����
		oCls.LogReportWrite = 1 


		'** �α� ������ ��� �ִ� ���� (������ \ �� �ٿ��ֽñ� �ٶ��ϴ�)
		oCls.SetLogFolder = "C:\inetpub\logs\LogFiles\W3SVC4\"
		

		'** ���������� �̵��� ����    ||   Default : ���� vbs ������ ����Ǵ� ����   "."
		'** ex) D:\�α�����\Ȩ������\2006��α�  ||  ������ \ �� ���ֽñ� �ٶ��ϴ� 
		oCls.SetMoveFolder = "D:\GitHub\iislogBackup" 

		
		'** iisLog �� ����Ǵ� ���� Ÿ���Դϴ� 
		'** iis ������ �α� ���� ������ ���� ���� / �ְ� ���� / ���� ���� �α� ������� �������ּ��� 
		'** 1 : ���� / 2 : �ְ� / 3 : ����
		oCls.LogType = 1
		
		'**  iisLogBackup�� �ǽ��� ������ �Ǵ� �ɼ��� �����մϴ�
		'** ������ ���� �ɼǴ�� ������ �����˴ϴ�.
		'**  m : ���� ��� ���� /   h : 15��(����) ��� ���� /  d: 1��(����) ���� / n : ���� �α� ��� (Default : ���糯¥)
		oCls.LogBackupOption = "m" 
		
		'** �α������� ����� ���ϸ��Դϴ�
		'** ���������� zip�����Դϴ�. (������ zip���� �ϰŵ翩)
		'** yyyy // mm // dd // ww �� ���ڰ� �����մϴ�.
		'** ���� ���� ����ϼŵ� �˴ϴ� - ���� �ɼ�(n)�� �ʿ���
		'** ex) WEB-yymmdd.zip   result : WEB-060125.zip
		'** ex) yyyymmdd.zip   result : 20060125.zip
		'** ex) yyyymmdd_Logbackup.zip   result : 20060125_LogBackup.zip
		'** ex) Log_yymmdd.zip   result : Log_060125.zip

		oCls.LogBackupFileName = "WEBLog_yyyymmdd.zip"

		'** ���� �α� ��� 
		'** �� ������Ƽ�� ���� ����� �ƴ� ������ �α׿� ���� ����� �����ϴµ� �����ϰ� ����Ͻ� �� �ֽ��ϴ�
		'** ���ڷ� ���� ��� �����ϵ��� ����
		'** ex) iislogbackup abc_yyyymm.zip u_ex1401*.log 

		IF Wscript.Arguments.Length > 0 THEN
			oCls.LogBackupOption = "n" 
			oCls.LogBackupFileName = Wscript.Arguments.Item(0)
			oCls.ForceLogFile = Wscript.Arguments.Item(1) '"u_ex1401*.log" 
		End If

		'** iisLogBackup Ŭ���� ���ο��� �̷����� ���μ����� �α� ����Ʈ�� ����ϴ�
		'** 1 : �α� ���� ���� / 0 : �α� ���� ���� ����
		
		oCls.LogFileDelete = 0 

		'** iisLogBackup Class Execute
		
		oCls.Exec()
		

		'** �� ������Ƽ�� ��� �Ŀ� ������ ���� �� �� �ֽ��ϴ�
		'** ������ �������ֽø� �ǰڽ��ϴ�
		'** �� CDont.NewMail �� �̿������� Windows 2000 (Adv) Server���� SMTP ���񽺰� ��ġ �Ǿ��־�� �մϴ�.
		'** Windows 2003 �� ���� DLL�� ����Ͻô���, CDO���Ϸ� Source�� �ణ �����ϼż� ����Ͻñ� �ٶ��ϴ�
		'** ���ÿ��� �ּ��� ���� �ϰ� ����Ͽ� �ּ���

		'** �� CDO sendmail �� �Լ��� ���� - oCls.OutSendToMail()

'		Dim strLogFileName : strLogFileName = oCls.LogFileName
'		oCls.Sendmail "�����»��<send@sendmail.com>", "�޴»��<receive@receive.com>", ToYMDDate(date()) & chr(9) & strLogFileName & "����� �Ǿ����ϴ�", "�ù�",1, 1, 0

	SET oCls = Nothing

 

'*****************************************************************************************************
'
'	�� �Ʒ� ���ʹ� Ŭ���� �����Դϴ�. 
'	�����Ͻ� �� �ִ� �ɷ��� ������ �е��� �����ϼż� �� ���� ����ϼŵ� �����մϴ� :)
'	�� �� ���� ���̵�� ������ �е� Contact ���ּ���!!!
'	
'	Class iisLogBackup Ver 1.2
'
'	�������� ������ ����� �� iisLogBackup Script
'	�α� ���� ���� �׸��� ������ ���콺 Ŭ���ϴ��� ������� �հ������� �� ������...
'	windows server administrator ���� ��ȭ�� �Ƚ��� �ٶ��......
'
'	2006. 01. 25.			modify : 2006. 12. 29.
'	Contact : Ssemi��	(http://www.ssemi.net)
'	
'	Dev Blog : http://ssemi.tistory.com
'*****************************************************************************************************


'*****************************************************************************************************
' Class iisLogBackup
'*****************************************************************************************************

CLASS iisLogBackup
	
	'----------------------------------------------------------------------------------------------------
	
	'** ���� ��ü ����
	Dim Shell
	Dim FSO
	Dim Folder
	Dim File
	Dim WshShell
	
	'** ���� Member ����
	Private CMD
	Private isLogFile
	Private m_FolderName
	Private m_moveFolderName
	Private m_LogFile
	Private PathHere 

	'** ������ ������Ƽ�� ���� ������ Member ����
    Private m_LogType
    Private m_LogBackupOption
	Private m_ForceLogFile
    Private m_LogBackupFileName
	Private m_logBackupAfterDelete
	Private m_LogReportWrite

	'----------------------------------------------------------------------------------------------------

	'** ������Ƽ�� ���� ������ �� ȣ��ȴ�.
    Public Property Let LogType(strArg)
		m_LogType = strArg
	End Property

    Public Property Let LogBackupOption(strArg)
		m_LogBackupOption = strArg
    End Property

    Public Property Let ForceLogFile(strArg)
		m_ForceLogFile = strArg
    End Property

	Public Property Let LogReportWrite(strArg)
		m_LogReportWrite = strArg
    End Property

    Public Property Let LogBackupFileName(strArg)
		m_LogBackupFileName = changeFormat(strArg)
    End Property

    Public Property Let LogFileDelete(strArg)
		m_logBackupAfterDelete = strArg
    End Property

	Public Property Let SetLogFolder(strArg) 
		IF FSO.FolderExists(strArg) Then
			SET Folder = FSO.GetFolder(strArg) 
		ELSE
			IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & ""& chr(9) & "[" & strArg &"] ������ �������� �ʽ��ϴ�")
		End IF
		m_FolderName = strArg
	End Property

	Public Property Let SetMoveFolder(strArg) 
		IF NOT FSO.FolderExists(strArg) Then
			FSO.CreateFolder(strArg) 
			IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & ""& chr(9) &"������ ������ �̵��� [" & strArg &"] ���� ����")
		End IF
		m_moveFolderName = strArg
	End Property	

	'----------------------------------------------------------------------------------------------------
	
	'** ������Ƽ���� ���� �о �� ȣ��ȴ�.
    Public Property Get LogFileName()
      LogFileName = m_LogBackupFileName
    End Property

	'----------------------------------------------------------------------------------------------------
	
	'** �α� ���� üũ
	Private Function ExistsLogFile()

		IF isObject(Folder) Then
			For Each File In Folder.Files 
				IF Right(LCase(File.Name), 3) = "log" Then 
					isLogFile = True
					Exit For
				ELSE
					isLogFile = False
				End IF
			Next
			SET Folder = Nothing
		End IF
	
		ExistsLogFile = isLogFile

	End Function 


	'** �α� ���� ������
	Private Function changeFormat(str)
	
		Dim tempstr, currentWeek
		
		Select Case Lcase(m_LogBackupOption)

			Case "m" :  '�Ѵ�
				tempstr = DateAdd("m", -1 , ToYMDDate(date()))
				currentWeek = "0" & Cstr(DatePart("ww", tempstr) - DatePart("ww", Year(tempstr) & "-" & Month(tempstr) & "-01") + 1)

				str = Replace(str, "yyyy", Split(tempstr, "-")(0))
				str = Replace(str, "yy", Right(Split(tempstr, "-")(0), 2))
				str = Replace(str, "mm", Split(tempstr, "-")(1))
				str = Replace(str, "ww", "")
				str = Replace(str, "dd", "")

			Case "h" : ' ����
				tempstr = DateAdd("d", -15 , ToYMDDate(date()))
				currentWeek = "0" & Cstr(DatePart("ww", tempstr) - DatePart("ww", Year(tempstr) & "-" & Month(tempstr) & "-01") + 1)

				str = Replace(str, "yyyy", Split(tempstr, "-")(0))
				str = Replace(str, "yy", Right(Split(tempstr, "-")(0), 2))
				str = Replace(str, "mm", Split(tempstr, "-")(1))
				IF Split(tempstr, "-")(2) > 15 Then
					str = Replace(str, "dd", "half_2")
					str = Replace(str, "ww", "half_2")
				ELSE
					str = Replace(str, "dd", "half_1")
					str = Replace(str, "ww", "half_1")
				End IF
				
			Case "d" :  ' 1��
				tempstr = DateAdd("d", -1 , ToYMDDate(date()))
				str = Replace(str, "yyyy", Split(tempstr, "-")(0))
				str = Replace(str, "yy", Right(Split(tempstr, "-")(0), 2))
				str = Replace(str, "mm", Split(tempstr, "-")(1))
				str = Replace(str, "dd", Split(tempstr, "-")(2))

			Case "n" :  ' ����
				currentWeek = "0" & Cstr(DatePart("ww", ToYMDDate(date())) - DatePart("ww", Year(ToYMDDate(date())) & "-" & Month(ToYMDDate(date())) & "-01") + 1)  & "w"
				str = Replace(str, "yyyy", Split(ToYMDDate(date()), "-")(0))
				str = Replace(str, "yy", Right(Split(ToYMDDate(date()), "-")(0), 2))
				str = Replace(str, "mm", Split(ToYMDDate(date()), "-")(1))
				IF m_LogType = 1 Then
					str = Replace(str, "dd", Split(ToYMDDate(date()), "-")(2))
				ELSEIF m_LogType = 2 Then
					str = Replace(str, "ww", currentWeek)
				End IF
				
				'force mode paste guid name = multi execute 
				'Dim TypeLib : Set TypeLib = CreateObject("Scriptlet.TypeLib")
				'str = replace(str, ".zip", "_"& Mid(TypeLib.Guid, 2, 4) & ".zip")
		End Select

		changeFormat = str

	End Function


	'** �α� ���� �ڵ������
	Private Sub LogFileCoordinator()

		Dim tmp, standard
		Dim y, m, d, w

		' �α׹���ɼ��� ���� ������ ����
		Select Case Lcase(m_LogBackupOption)
			Case "m" : '�Ѵ� month
				standard = DateAdd("m", -1 , ToYMDDate(date()))
			Case "h" :  '15�� ����
				standard = DateAdd("d", -15, ToYMDDate(date()))
			Case "d" :  '1�� day
				standard = DateAdd("d", -1 , ToYMDDate(date()))
			Case "n" :  '���� non
				standard = ToYMDDate(date())
		End Select

		y = Right(Year(standard), 2) : m = Month(standard) : d = Day(standard) : w = Cstr(DatePart("ww", standard) - DatePart("ww", Year(standard) & "-" & Month(standard) & "-01") + 1)
		IF Cstr(m) < 10 Then m = "0" & Cstr(m) ELSE m = CStr(m) End IF
		IF Cstr(d) < 10 Then d = "0" & Cstr(d) ELSE d = Cstr(d) End IF
		IF Cstr(w) < 10 Then w = "0" & Cstr(w) ELSE w = Cstr(w) End IF

		' �α׹�� Ÿ�����ϸ� ����
		Dim min , max, i 
		Select Case Lcase(m_LogBackupOption)
			Case "m" : 
				m_LogFile = "u_ex" & y & m & "*.log"
				
				Call LogFileBackup(m_LogBackupFileName, m_FolderName & m_LogFile, "a -tzip")

			Case "h" : 
				
				IF m_LogType = 1 Then
					IF Day(Standard) < 16 Then
						min = "01" : max = "15"
					ELSE
						min = "16" : max = Day(Dateserial(Year(standard), Month(standard)+1, 1) - 1) 
					End IF

					For i = min TO max
						IF i < 10 Then i = "0" & Cstr(i) 
						m_LogFile = "u_ex" & y & m & i & ".log"

						IF FSO.FileExists(m_FolderName & m_LogFile) And FSO.FileExists(m_moveFolderName & "\" & m_LogBackupFileName) Then
							Call LogFileBackup(m_LogBackupFileName, m_FolderName & m_LogFile, "u")
						ELSE
							Call LogFileBackup(m_LogBackupFileName, m_FolderName & m_LogFile, "a -tzip")
						End IF
					Next

				ELSEIF m_LogType = 2 Then
					Dim LastWeek, NowWeek
					LastWeek = Datepart("ww", DateAdd("d", -1, Dateserial(year(standard), month(standard)+1, 1))) - Datepart("ww", Dateserial(year(standard), month(standard), 1))  +1 
					NowWeek = Datepart("ww", standard)

					IF Cint(NowWeek) <= Cint(LastWeek / 2) Then
						min = "01" : max = "0" & Cstr(Cint(LastWeek / 2))
					ELSE
						min = "0" & Cstr(Cint(LastWeek / 2) +1) : max = "0"& Cstr(LastWeek)
					End IF


					For i = min TO max
						IF i < 10 Then i = "0" & Cstr(i) 
						m_LogFile = "u_ex" & y & m & i & ".log"

						IF FSO.FileExists(m_FolderName & m_LogFile) And FSO.FileExists(m_moveFolderName & "\" & m_LogBackupFileName) Then
							Call LogFileBackup(m_LogBackupFileName, m_FolderName & m_LogFile, "u")
						ELSE
							Call LogFileBackup(m_LogBackupFileName, m_FolderName & m_LogFile, "a -tzip")
						End IF
					Next

					'----------------------------------------------------------------------------------------
					' ex-source...min vs max 
'					m_LogFile = "ex" & y & m & min & ".log"
'					Call LogFileBackup(m_LogBackupFileName, m_FolderName & m_LogFile, "a -tzip")
'
'					For i = min + 1 TO max
'						IF i < 10 Then i = "0" & Cstr(i) 
'						m_LogFile = "ex" & y & m & i & ".log"
'						Call LogFileBackup(m_LogBackupFileName, m_FolderName & m_LogFile, "u")
'					Next
					'----------------------------------------------------------------------------------------
				ELSE
					IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() &  chr(9) & "iisLog�� ����� ���� Ÿ�� ������ [���� ����] or [�� ����]�� �ƴմϴ�")

				End IF


			Case "d" : 
				
				IF m_LogType = 1 Then
					m_LogFile = "u_ex" & y & m & d & ".log"
					Call LogFileBackup(m_LogBackupFileName, m_FolderName & m_LogFile, "a -tzip")
				ELSE
					IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() &  chr(9) & "iisLog�� ����� ���� Ÿ�� ������ [���� ����]�� �ƴմϴ�")
				End IF

			Case "n" : 
			
				IF Len(m_ForceLogFile) > 0 Then
					m_LogFile = m_ForceLogFile
				ELSE
					IF m_LogType = 1 Then
						m_LogFile = "u_ex" & y & m & d & ".log"
					ELSEIF m_LogType = 2 Then
						m_LogFile = "u_ex" & y & m & w & ".log"
					ELSE
						m_LogFile = "u_ex" & y & m & "*.log"
					End IF

				End IF

				Call LogFileBackup(m_LogBackupFileName, m_FolderName & m_LogFile, "a -tzip")

		End Select
		
	End Sub 



	'** �α� ���� ��� ���ν���
	Private Sub LogFileBackup(zip, target, typeOption)
		'IF FSO.FileExists(target) Then
			
			cmd  = "7z "& typeOption &" " & newFileName(zip) & " " & target
			IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & chr(9) & cmd)
			
			Shell.Run cmd , , True

			IF typeOption = "u" Then
				IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() &  chr(9) & "[" & target & "] ������ ["& zip &"] ���Ϸ� ���� <������Ʈ>")
			ELSE
				IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() &  chr(9) & "[" & target & "] ������ ["& zip &"] ���Ϸ� ���� ����")
			End IF

			' log file del
			IF m_logBackupAfterDelete = 1 Then
				FSO.DELETEFILE target, True
				IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() &  chr(9) & "[" & target & "] ������ �����Ͽ����ϴ�")
			End IF
		'ELSE
			'IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & chr(9) & "[" & target & "] ������ ã�� ���� �����ϴ�")
		'End IF
	End Sub


	'** �α� ��� ����
	Public Sub Exec()
		'���� ����� �� ���� �̸� ������ ���� 
		IF m_LogBackupOption = "n" AND Len(m_ForceLogFile) = 0  Then
			Wscript.Echo "No File parameter was passed ( force mode )"
			Call TerminateClass
			Wscript.Quit
		End IF
		
		PathHere = FSO.GetAbsolutePathName(".")

		isLogFile = ExistsLogFile()
		
		IF isLogFile Then
			Call LogFileCoordinator()
			
			'** Log File Move
			IF FSO.FileExists(m_LogBackupFileName) Then
				IF m_moveFolderName <> "." OR  m_moveFolderName <> PathHere Then '** Not Default value
					IF NOT FSO.FileExists(m_moveFolderName & "\" & m_LogBackupFileName) Then
						Call ErrorReport(NOW() & chr(9) & "Move " & PathHere & "\" & m_LogBackupFileName & ", " & m_moveFolderName & "\" )
						FSO.MoveFile PathHere & "\" & m_LogBackupFileName , m_moveFolderName & "\" 
						IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & chr(9) & "[" & PathHere & "\" & m_LogBackupFileName &"] ������ ["& m_moveFolderName &"]�� �̵��Ǿ����ϴ�")
					End IF
				End IF
			End IF

		ELSE
			IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & chr(9) & "["& m_FolderName &"] ������ �α� ������ �������� �ʽ��ϴ�")
		End IF

		IF m_LogReportWrite = 1 Then Call ErrorReport(String(100, "-"))

	End Sub

	'----------------------------------------------------------------------------------------------------

	'** Initialize EVENT
    Private Sub Class_Initialize
		Call InitClass
    End Sub

    '** Terminate EVENT
    Private Sub Class_Terminate
		Call TerminateClass
    End Sub

	Private Sub InitClass
		SET Shell = CreateObject("WScript.Shell") 
		SET FSO = CreateObject("Scripting.FileSystemObject") 
	End Sub

	Private Sub TerminateClass
		SET FSO = Nothing
		SET Shell = Nothing
	End Sub
	'----------------------------------------------------------------------------------------------------

	'** Error Report 
	Private Sub ErrorReport(str)
		Dim ReportFile : ReportFile = Left(Replace(ToYMDDate(date()), "-", ""), 6) & "Report.txt"
		IF NOT FSO.FileExists(ReportFile) Then
			FSO.CreateTextFile ReportFile, True
			SET F = FSO.OpenTextFile(ReportFile, 8, True)
				F.WriteLine String(100, "-")
				F.WriteLine chr(9) & "Windows IIS Log Backup Script Ver1.2 " & chr(9) & chr(9) & "2006. 12. 29."
				F.WriteLine vbCr
				F.WriteLine chr(9) & "Contact : Ssemi�� (http://www.ssemi.net) "
				F.WriteLine String(100, "-")
			SET F = Nothing
		End IF

		Dim F
		SET F = FSO.OpenTextFile(ReportFile, 8, True)
			F.WriteLine str
		SET F = Nothing
	End Sub

	'----------------------------------------------------------------------------------------------------
	'** SendMail 
	Public Sub Sendmail(strFrom, strTo, strSubject, strBody, bodyFormat, mailFormat, Importance)
		Dim objSendMail
		SET objSendMail = CreateObject("CDONTS.NewMail")
			objSendMail.From = strFrom
			objSendMail.To = strTo
			objSendMail.Subject = strSubject & " (" & ToYMDDate(date()) & ")"
			objSendMail.Body = strBody
			objSendMail.BodyFormat = bodyFormat '0 HTML / 1 TEXT
			objSendMail.MailFormat = mailFormat ' 0 MIME / 1 TEXT
			objSendMail.Importance = Importance ' 0 low / 1 normal / 2 importance
			objSendMail.Send
			SET objSendMail = Nothing
		IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & chr(9) & "["& strTo &"]�� ���� : ["& strSubject &"] ������ ���½��ϴ�")

	End Sub

	'** CDO use SendMail
	Public Function OutSendToMail(mailServer, FromUN, FromUA, strTo, strSubject, strBody)
		IF Len(ToUser) = 0 OR isNumeric(ToUser) Then Exit Function

		DIM iMsg
		DIM Flds
		DIM iConf

		SET iMsg = CreateObject("CDO.Message")
		SET iConf = CreateObject("CDO.Configuration")
		SET Flds = iConf.Fields

		Flds(cdoSendUsingMethod) = cdoSendUsingPort 
		Flds(cdoSMTPServer) = mailServer
		Flds(cdoSMTPServerPort) = 25
		Flds(cdoSMTPAuthenticate) = cdoAnonymous
		'Flds(cdoSendUserName) = "user"
		'Flds(cdoSendPassword) = "password"
		Flds(cdoURLGetLatestVersion) =  True
		Flds.Update

		With iMsg
			SET .Configuration = iConf
			.To = strTo
			.From = FromUN
			.Sender = FromUA
			.Subject = strSubject
			.TextBody = "" & strBody & ""
			.Send
		End With

		IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & chr(9) & "["& strTo &"]�� ���� : ["& strSubject &"] ������ ���½��ϴ�")

	End Function

	' YYYY-MM-DD �������� ����
	Function ToYMDDate(dt)
		dim s
		s = datepart("yyyy",dt)
		s = s & "-" & RIGHT("0" & datepart("m",dt),2)
		s = s & "-" & RIGHT("0" & datepart("d",dt),2)
		ToYMDDate = s
	End Function

	' duplicate file ó��
	Function newFileName(name)
		Dim Num : Num = 1
		Dim tempNum, removeNo, tmpFile
		tmpFile = name
		
		DO 
			IF FSO.FileExists(m_moveFolderName & "\" & name) THEN
				removeNo = Num - 1
				tempNum = "[" + CStr(removeNo) + "]"
				name = Replace(Replace(name, ".zip", ""), tempNum, "") & "[" & num & "].zip"
				Num = Num + 1
			ELSE
				Exit Do
			End IF
		Loop
		newFileName = name
	End Function
		
	' statement ?  t : f
	Function IIF(statement, t, f)
		if (statement) Then
			IIF = t
		Else
			IIF = f
		End If
	End Function
End Class

