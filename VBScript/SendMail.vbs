Set args = Wscript.Arguments

strMailTo = args.Item(0) 
'msgbox strMailTo
strMailCC  = args.Item(1)
strSummaryFileName = args.Item(2)
g_tSummaryEnd_Time = args.Item(3)
g_tSummaryStart_Time = args.Item(4)
strRelease = args.Item(5)
strModuleName = args.Item(6)
g_SummaryTotal_TC = args.Item(7)
g_SummaryTotal_Pass = args.Item(8)
g_SummaryTotal_Fail = args.Item(9)
strEnvironment = args.Item(10)

'strMailTo = "Krishna.manubolu@thehartford.com"
'strMailCC = "Krishna.manubolu@thehartford.com"
'g_tSummaryStart_Time = Date
'g_tSummaryEnd_Time = Date
'strSummaryFileName = "C:\Selenium\WorkSpace\PC\Grid.txt"
'strRelease = "R2"
'strModuleName = "PC"
'strUserNameg_SummaryTotal_TC = "5"
'g_SummaryTotal_Pass = "5"
'g_SummaryTotal_Fail = "0"
'strEnvironment = "LTA"




Call fnSendSummarySnapshotEmail(strMailTo,strMailCC,strSummaryFileName,g_tSummaryEnd_Time,g_tSummaryStart_Time,strRelease,strModuleName,g_SummaryTotal_TC,g_SummaryTotal_Pass,g_SummaryTotal_Fail,strEnvironment)
'==================================================================================================
'
'
'==================================================================================================
Public Function fnSendSummarySnapshotEmail(strMailTo,strMailCC,strSummaryFileName,g_tSummaryEnd_Time,g_tSummaryStart_Time,strRelease,strModuleName,g_SummaryTotal_TC,g_SummaryTotal_Pass,g_SummaryTotal_Fail,strEnvironment)
			'On Error Resume Next
			strUserName = fnGetUserName()
			'strModuleName = "Policy Center"'fnGetModuleName()
			intSeconds = DateDiff("s",g_tSummaryStart_Time,g_tSummaryEnd_Time)
			strTime = fnSecondsToTime(intSeconds)
			Set objOutlook = CreateObject("Outlook.Application")
			Set objMail = objOutlook.CreateItem(0)
			With objMail
				objMail.to = strMailTo
				If strMailCC <> "" Then
						objMail.cc = strMailCC
				End If
				objMail.Subject = "Automation Execution Summary Snapshot - " & Date &" - " & Time
				'html and body tage,virtical bar
				SHTML = "<HTML><BODY><BR><BR><TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=70%><TR><TD BGCOLOR=RED WIDTH=100%></TD></TR></TABLE><br>"
				'wellsfargo logo
				SHTML = SHTML & "<TABLE WIDTH=70%><TR>"
				SHTML = SHTML & "<TD BGCOLOR=WHITE WIDTH=6%><IMG SRC='https://www.thehartford.com/sites/the_hartford/img/logo.png'></TD>"
				SHTML = SHTML & "<TD WIDTH=88% BGCOLOR=WHITE align=center><table width=100%><TR><td align=""center""><FONT FACE=Calibri COLOR=Black SINZE=5><B>COMPAS Automation Execution Summary Snapshot</B></FONT></td></tr><tr><table width=85% align=center><tr><td width=50% style='margin;0in;margin-bottom:.0001pt;text-align:left;font-size:10.5pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>" & strModuleName &"</B></FONT></td><td width=50% style='margin:0in;margin-bottom:.0001pt;text-align:right;font-size:10.5pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>"& strRelease & "</B></FONT></td></tr></tr></table></table></TD>"
				SHTML = SHTML & "<TD BGCOLOR=WHITE WIDTH=6%><IMG SRC='https://www.thehartford.com/sites/the_hartford/img/logo.png'></TD>"
				SHTML = SHTML & "</TR></TABLE>"
				'Details Bar
				SHTML =  SHTML & "<P style='margin:0in;margin-bottom:.0001pt;font-size:11.0pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>Here are the details of your batch execution in " & strEnvironment & " Environment:</p><BR>"
				'execution info
				SHTML = SHTML & "<TABLE style=""border:2px solid #C0C0C0"" WIDTH=55%><tr width=100%><table width=95% align =center>"
				SHTML = SHTML & "<TR WIDTH=100><table width=100%><TR><td width=50% height=15 style='font-size:9.0pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>&nbsp; Total Number of Test Cases</td><td width=50% height=15 style='font-size:9.0pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>&nbsp;" & g_SummaryTotal_TC &"</td></tr></table><table width=100%><tr><table align=left height=10 width=90%><tr><td bgcolor=#C0C0C0 width=100%></td></tr></table></tr></table></TR>"
				SHTML = SHTML & "<TR WIDTH=100><table width=100%><TR><td width=50% height=15 style='font-size:9.0pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>&nbsp; Total Test Cases Passed:</font></td><td width=50% height=15 style='font-size:9.0pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>&nbsp;" & g_SummaryTotal_Pass &"</font></td></tr></table><table width=100%><tr><table align=left height=10 width=90%><tr><td bgcolor=#C0C0C0 width=100%></td></tr></table></tr></table></TR>"
				SHTML = SHTML & "<TR WIDTH=100><table width=100%><TR><td width=50% height=15 style='font-size:9.0pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>&nbsp; Test Cases Failed:</b></font></td><td width=50% height=15 style='font-size:9.0pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>&nbsp;" & g_SummaryTotal_Fail &"</b></font></td></tr></table><table width=100%><tr><table align=left height=10 width=90%><tr><td bgcolor=#C0C0C0 width=100%></td></tr></table></tr></table></TR>"
				SHTML = SHTML & "<TR WIDTH=100><table width=100%><TR><td width=50% height=15 style='font-size:9.0pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>&nbsp; Test Execution Time:</b></font></td><td width=50% height=15 style='font-size:9.0pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>&nbsp;" & strTime &"</b></font></td></tr></table><table width=100%><tr><table align=left height=10 width=90%><tr><td bgcolor=#C0C0C0 width=100%></td></tr></table></tr></table></TR>"
				SHTML = SHTML & "<TR WIDTH=100><table width=100%><TR><td width=50% height=15 style='font-size:9.0pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>&nbsp; Test Execution Time:</b></font></td><td width=50% height=15 style='font-size:9.0pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>&nbsp;" & Date &" - "& Time &"</b></font></td></tr></table><table width=100%><tr><table align=left height=10 width=90%><tr><td bgcolor=#C0C0C0 width=100%></td></tr></table></tr></table></TR>"
				SHTML = SHTML & "<TR WIDTH=100><table width=100%><TR><td width=50% height=15 style='font-size:9.0pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>&nbsp; Executed By:</b></font></td><td width=50% height=15 style='font-size:9.0pt;font-family:""Calibri"",""sans-serif"";color:black;mso-font-kerning:12.0pt'>&nbsp;" & UCase(strUserName) &"</b></font></td></tr></table></TR>"
				SHTML = SHTML & "</table></tr></TABLE><br>"
				'vertical bar, html,body close
				SHTML = SHTML & "<TABLE BORDER=0 CELLPADDING=3 CELLSPACING=1 WIDTH=70%><TR><TD BGCOLOR=RED WIDTH=100%></TD></TR></TABLE></BODY></HTML>"
				objMail.htmlBody = SHTML
				objMail.Attachments.Add(strSummaryFileName)
				objMail.Send
			End With
			Set objMail = Nothing
			Set objOutlook = Nothing
End Function

'==================================================================================================
'
'
'==================================================================================================
Public Function fnGetUserName()
				Set objSysInfo= CreateObject("ADSystemInfo")
				strUser = objSysInfo.UserName
				Set objUser = GetObject("LDAP://" & strUser)
				strFullName = objUser.Get("CN")
				fnGetUserName = strFullName
End Function

'==================================================================================================
'
'
'==================================================================================================
Function fnSecondsToTime(ByVal intSeconds)
			Dim hours, minutes, seconds
			'calcultes whole hours (like a div operator)
			hours = intSeconds \ 3600
			'calculates theremaning number of seconds
			intSeconds = intSeconds Mod 3600
			'calculates the whole number of minutes in the remaining number of seconds
			minutes = intSeconds \ 60
			'Calculates the remaining number of seconds after taking the number of minutes
			seconds = intSeconds Mod 60
			'Returns as a string
			fnSecondsToTime = hours & ":" & minutes & ":" & seconds
End Function
