	'*****************************************************************************************************
	'	Class 선언 및 옵션 설정	
	'*****************************************************************************************************
	Option Explicit
	Dim oCls
	SET oCls = New iisLogBackup


		'** iisLogBackup 클래스 내부에서 이뤄지는 프로세스를 로그 레포트로 남깁니다
		'** 1 : 로그 레포트 남김 / 0 : 로그 레포트 남기지 않음
		oCls.LogReportWrite = 1 


		'** 로그 파일이 들어 있는 폴더 (마지막 \ 꼭 붙여주시기 바랍니다)
		oCls.SetLogFolder = "C:\inetpub\logs\LogFiles\W3SVC4\"
		

		'** 압축파일이 이동될 폴더    ||   Default : 현재 vbs 파일이 실행되는 폴더   "."
		'** ex) D:\로그파일\홈페이지\2006년로그  ||  마지막 \ 꼭 빼주시기 바랍니다 
		oCls.SetMoveFolder = "D:\GitHub\iislogBackup" 

		
		'** iisLog 가 백업되는 파일 타입입니다 
		'** iis 서버의 로그 파일 설정이 일일 단위 / 주간 단위 / 월간 단위 로그 백업인지 선택해주세요 
		'** 1 : 일일 / 2 : 주간 / 3 : 월간
		oCls.LogType = 1
		
		'**  iisLogBackup을 실시할 기준이 되는 옵션을 설정합니다
		'** 각각의 기준 옵션대로 파일이 생성됩니다.
		'**  m : 월간 백업 기준 /   h : 15일(보름) 백업 기준 /  d: 1일(일일) 기준 / n : 강제 로그 백업 (Default : 현재날짜)
		oCls.LogBackupOption = "m" 
		
		'** 로그파일을 백업할 파일명입니다
		'** 파일형식은 zip형식입니다. (압축을 zip으로 하거든여)
		'** yyyy // mm // dd // ww 의 인자가 존재합니다.
		'** 인자 없이 사용하셔도 됩니다 - 강제 옵션(n)시 필요함
		'** ex) WEB-yymmdd.zip   result : WEB-060125.zip
		'** ex) yyyymmdd.zip   result : 20060125.zip
		'** ex) yyyymmdd_Logbackup.zip   result : 20060125_LogBackup.zip
		'** ex) Log_yymmdd.zip   result : Log_060125.zip

		oCls.LogBackupFileName = "WEBLog_yyyymmdd.zip"

		'** 강제 로그 백업 
		'** 이 프로퍼티는 단위 백업이 아닌 무제한 로그에 대한 백업을 진행하는데 유용하게 사용하실 수 있습니다
		'** 인자로 강제 모드 지정하도록 변경
		'** ex) iislogbackup abc_yyyymm.zip u_ex1401*.log 

		IF Wscript.Arguments.Length > 0 THEN
			oCls.LogBackupOption = "n" 
			oCls.LogBackupFileName = Wscript.Arguments.Item(0)
			oCls.ForceLogFile = Wscript.Arguments.Item(1) '"u_ex1401*.log" 
		End If

		'** iisLogBackup 클래스 내부에서 이뤄지는 프로세스를 로그 레포트로 남깁니다
		'** 1 : 로그 파일 삭제 / 0 : 로그 파일 삭제 안함
		
		oCls.LogFileDelete = 0 

		'** iisLogBackup Class Execute
		
		oCls.Exec()
		

		'** 이 프로퍼티는 백업 후에 메일을 보내 실 수 있습니다
		'** 각각을 설정해주시면 되겠습니다
		'** 단 CDont.NewMail 을 이용함으로 Windows 2000 (Adv) Server에서 SMTP 서비스가 설치 되어있어야 합니다.
		'** Windows 2003 일 경우는 DLL을 등록하시던지, CDO메일로 Source를 약간 수정하셔서 사용하시기 바랍니다
		'** 사용시에는 주석을 제거 하고 사용하여 주세요

		'** ※ CDO sendmail 용 함수가 있음 - oCls.OutSendToMail()

'		Dim strLogFileName : strLogFileName = oCls.LogFileName
'		oCls.Sendmail "보내는사람<send@sendmail.com>", "받는사람<receive@receive.com>", ToYMDDate(date()) & chr(9) & strLogFileName & "백업이 되었습니다", "냉무",1, 1, 0

	SET oCls = Nothing

 

'*****************************************************************************************************
'
'	※ 아래 부터는 클래스 내용입니다. 
'	수정하실 수 있는 능력이 있으신 분들은 수정하셔서 더 좋게 사용하셔도 무방합니다 :)
'	좀 더 좋은 아이디어 있으신 분들 Contact 해주세요!!!
'	
'	Class iisLogBackup Ver 1.2
'
'	귀차니즘 때문에 만들게 된 iisLogBackup Script
'	로그 파일 압축 그리고 삭제를 마우스 클릭하느라 힘들었던 손가락에게 이 영광을...
'	windows server administrator 에게 평화와 안식을 바라며......
'
'	2006. 01. 25.			modify : 2006. 12. 29.
'	Contact : Ssemi™	(http://www.ssemi.net)
'	
'	Dev Blog : http://ssemi.tistory.com
'*****************************************************************************************************


'*****************************************************************************************************
' Class iisLogBackup
'*****************************************************************************************************

CLASS iisLogBackup
	
	'----------------------------------------------------------------------------------------------------
	
	'** 각종 객체 변수
	Dim Shell
	Dim FSO
	Dim Folder
	Dim File
	Dim WshShell
	
	'** 공통 Member 변수
	Private CMD
	Private isLogFile
	Private m_FolderName
	Private m_moveFolderName
	Private m_LogFile
	Private PathHere 

	'** 실제로 프로퍼티의 값을 보관할 Member 변수
    Private m_LogType
    Private m_LogBackupOption
	Private m_ForceLogFile
    Private m_LogBackupFileName
	Private m_logBackupAfterDelete
	Private m_LogReportWrite

	'----------------------------------------------------------------------------------------------------

	'** 프로퍼티에 값을 설정할 때 호출된다.
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
			IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & ""& chr(9) & "[" & strArg &"] 폴더가 존재하지 않습니다")
		End IF
		m_FolderName = strArg
	End Property

	Public Property Let SetMoveFolder(strArg) 
		IF NOT FSO.FolderExists(strArg) Then
			FSO.CreateFolder(strArg) 
			IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & ""& chr(9) &"압축한 파일을 이동할 [" & strArg &"] 폴더 생성")
		End IF
		m_moveFolderName = strArg
	End Property	

	'----------------------------------------------------------------------------------------------------
	
	'** 프로퍼티에서 값을 읽어갈 때 호출된다.
    Public Property Get LogFileName()
      LogFileName = m_LogBackupFileName
    End Property

	'----------------------------------------------------------------------------------------------------
	
	'** 로그 파일 체크
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


	'** 로그 파일 포맷팅
	Private Function changeFormat(str)
	
		Dim tempstr, currentWeek
		
		Select Case Lcase(m_LogBackupOption)

			Case "m" :  '한달
				tempstr = DateAdd("m", -1 , ToYMDDate(date()))
				currentWeek = "0" & Cstr(DatePart("ww", tempstr) - DatePart("ww", Year(tempstr) & "-" & Month(tempstr) & "-01") + 1)

				str = Replace(str, "yyyy", Split(tempstr, "-")(0))
				str = Replace(str, "yy", Right(Split(tempstr, "-")(0), 2))
				str = Replace(str, "mm", Split(tempstr, "-")(1))
				str = Replace(str, "ww", "")
				str = Replace(str, "dd", "")

			Case "h" : ' 보름
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
				
			Case "d" :  ' 1일
				tempstr = DateAdd("d", -1 , ToYMDDate(date()))
				str = Replace(str, "yyyy", Split(tempstr, "-")(0))
				str = Replace(str, "yy", Right(Split(tempstr, "-")(0), 2))
				str = Replace(str, "mm", Split(tempstr, "-")(1))
				str = Replace(str, "dd", Split(tempstr, "-")(2))

			Case "n" :  ' 강제
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


	'** 로그 파일 코디네이터
	Private Sub LogFileCoordinator()

		Dim tmp, standard
		Dim y, m, d, w

		' 로그백업옵션을 통한 기준일 생성
		Select Case Lcase(m_LogBackupOption)
			Case "m" : '한달 month
				standard = DateAdd("m", -1 , ToYMDDate(date()))
			Case "h" :  '15일 보름
				standard = DateAdd("d", -15, ToYMDDate(date()))
			Case "d" :  '1일 day
				standard = DateAdd("d", -1 , ToYMDDate(date()))
			Case "n" :  '강제 non
				standard = ToYMDDate(date())
		End Select

		y = Right(Year(standard), 2) : m = Month(standard) : d = Day(standard) : w = Cstr(DatePart("ww", standard) - DatePart("ww", Year(standard) & "-" & Month(standard) & "-01") + 1)
		IF Cstr(m) < 10 Then m = "0" & Cstr(m) ELSE m = CStr(m) End IF
		IF Cstr(d) < 10 Then d = "0" & Cstr(d) ELSE d = Cstr(d) End IF
		IF Cstr(w) < 10 Then w = "0" & Cstr(w) ELSE w = Cstr(w) End IF

		' 로그백업 타켓파일명 생성
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
					IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() &  chr(9) & "iisLog의 백업된 파일 타입 설정이 [일일 기준] or [주 기준]이 아닙니다")

				End IF


			Case "d" : 
				
				IF m_LogType = 1 Then
					m_LogFile = "u_ex" & y & m & d & ".log"
					Call LogFileBackup(m_LogBackupFileName, m_FolderName & m_LogFile, "a -tzip")
				ELSE
					IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() &  chr(9) & "iisLog의 백업된 파일 타입 설정이 [일일 기준]이 아닙니다")
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



	'** 로그 파일 백업 프로시져
	Private Sub LogFileBackup(zip, target, typeOption)
		'IF FSO.FileExists(target) Then
			
			cmd  = "7z "& typeOption &" " & newFileName(zip) & " " & target
			IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & chr(9) & cmd)
			
			Shell.Run cmd , , True

			IF typeOption = "u" Then
				IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() &  chr(9) & "[" & target & "] 파일을 ["& zip &"] 파일로 압축 <업데이트>")
			ELSE
				IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() &  chr(9) & "[" & target & "] 파일을 ["& zip &"] 파일로 압축 성공")
			End IF

			' log file del
			IF m_logBackupAfterDelete = 1 Then
				FSO.DELETEFILE target, True
				IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() &  chr(9) & "[" & target & "] 파일을 삭제하였습니다")
			End IF
		'ELSE
			'IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & chr(9) & "[" & target & "] 파일을 찾을 수가 없습니다")
		'End IF
	End Sub


	'** 로그 백업 실행
	Public Sub Exec()
		'강제 모드일 때 파일 이름 없으면 에러 
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
						IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & chr(9) & "[" & PathHere & "\" & m_LogBackupFileName &"] 파일이 ["& m_moveFolderName &"]로 이동되었습니다")
					End IF
				End IF
			End IF

		ELSE
			IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & chr(9) & "["& m_FolderName &"] 폴더에 로그 파일이 존재하지 않습니다")
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
				F.WriteLine chr(9) & "Contact : Ssemi™ (http://www.ssemi.net) "
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
		IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & chr(9) & "["& strTo &"]로 제목 : ["& strSubject &"] 메일을 보냈습니다")

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

		IF m_LogReportWrite = 1 Then Call ErrorReport(NOW() & chr(9) & "["& strTo &"]로 제목 : ["& strSubject &"] 메일을 보냈습니다")

	End Function

	' YYYY-MM-DD 형식으로 변경
	Function ToYMDDate(dt)
		dim s
		s = datepart("yyyy",dt)
		s = s & "-" & RIGHT("0" & datepart("m",dt),2)
		s = s & "-" & RIGHT("0" & datepart("d",dt),2)
		ToYMDDate = s
	End Function

	' duplicate file 처리
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

