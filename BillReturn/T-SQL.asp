<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
	On Error Resume Next
	Session("chkTsql")=1
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close
	ippath=""
	if sys_City="台中縣" then
		ippath="10.114.9.58"
	elseif sys_City="雲林縣" then
		ippath="10.122.1.136"
	elseif sys_City="屏東縣" then
		ippath="10.134.1.55"
	elseif sys_City="台南市" then
		ippath="10.130.83.146"
	end if
	
	'ippath="220.128.140.193"
	DownFilePath="\Down\"
	UpFilePath="\Up\"

	if sys_City="台南市" or sys_City="雲林縣" then
		SQL_DownFilePath="\\"&ippath&"\d$\f\data\"
		SQL_UpFilePath="\\"&ippath&"\d$\f\data\result\"
		SQL_LogPath="\\"&ippath&"\d$\f\data\log\"
	else
		SQL_DownFilePath="\\"&ippath&"\f$\data\"
		SQL_UpFilePath="\\"&ippath&"\f$\data\result\"
		SQL_LogPath="\\"&ippath&"\f$\data\log\"
	end if

	nowTime=request("nowTime")
	Set Fso = CreateObject("Scripting.FileSystemObject")
	Set f = Fso.GetFolder(SQL_DownFilePath)
	Set fc = f.Files
	dirlist=""
	For Each f1 in fc
		If not ifnull(dirlist) Then dirlist=dirlist&","
		dirlist=dirlist&f1.name
	Next

	If Fso.FileExists(SQL_LogPath&gInitDt(date)&"@" & hour(nowTime)&".txt") then
		dirlist=""
	End if

	If Not ifnull(dirlist) Then
		set ftxt=Fso.CreateTextFile(SQL_LogPath & gInitDt(date)&"@" & hour(nowTime)&".txt", True)
		ftxt.close

		dirlist=dirlist&","
		dirlist=split(dirlist,",")

		For i=0 to Ubound(dirlist)-1
			sqlList=""
			If Fso.FileExists(SQL_LogPath&trim(dirlist(i))) then
				Fso.DeleteFile SQL_LogPath&trim(dirlist(i))
			end if
			Fso.MoveFile SQL_DownFilePath & trim(dirlist(i)), SQL_LogPath

			Set f=Fso.OpenTextFile(SQL_LogPath & trim(dirlist(i)),1,True,0)
			While Not f.AtEndOfStream
				sqlList=sqlList&f.ReadLine
			Wend

			If instr(Ucase(sqlList),"BILLNO:")>0 Then
				tempStr=split(sqlList,"@")
				tmpBatchNumber=replace(Ucase(trim(tempStr(0))),"BATCHNUMBER:","")
				tmpBillNo=replace(Ucase(trim(tempStr(1))),"BILLNO:","")

				strSQL="select FileName from DCILog where BillNo='"&tmpBillNo&"' and batchnumber='"&tmpBatchNumber&"'"
				set rsfile=conn.execute(strSQL)
				If not rsfile.eof Then
					Fso.CopyFile server.mappath(DownFilePath & trim(rsfile("FileName"))&".big"),SQL_UpFilePath
					Fso.CopyFile server.mappath(UpFilePath & trim(rsfile("FileName"))),SQL_UpFilePath
				End if
				rsfile.close
			elseIf Ucase(trim(left(sqlList,7)))<>"SELECT" Then
				strSQL=""
				If instr(sqlList,"GO;")>0 Then
					strSQL=split(sqlList,"GO;")
				End if

				If instr(sqlList,"GO;")<=0 Then
					conn.execute(sqlList)
				else
					For h=0 to Ubound(strSQL)-1
						conn.execute(strSQL(h))
					Next
				End if
				errmsg=""
				errmsg=errmsg&trim(dirlist(i))&chr(13)
				If conn.errors.count>0 Then
					for h =0 to conn.errors.count-1
						errmsg=errmsg&conn.errors.item(h)&"<hr>"&chr(13)
						errmsg=errmsg&Err.Description
					next
				else
					errmsg=errmsg&"已更新完成!!"
				end if

				Set objTextFile = Fso.CreateTextFile(SQL_UpFilePath & sys_City & " - " & gInitDt(date) & replace(nowTime,":","") & " - " & trim(dirlist(i)), True)
				objTextFile.WriteLine(errmsg)
			else
				set rs=conn.execute(sqlList)
				If conn.errors.count>0 Then
					errmsg=""
					errmsg=errmsg&trim(dirlist(i))&chr(13)
					for h =0 to conn.errors.count-1
						errmsg=errmsg&conn.errors.item(h)&"<hr>"&chr(13)
						errmsg=errmsg&Err.Description
					next
					Set objTextFile = Fso.CreateTextFile(SQL_UpFilePath & sys_City & " - " & gInitDt(date) & replace(nowTime,":","") & " - " & trim(dirlist(i)), True)
					objTextFile.WriteLine(errmsg)
				else
					If Not rs.eof Then
						tmpWriteStr=""
						For h=0 to rs.Fields.count-1
							If h>0 Then tmpWriteStr=tmpWriteStr& "	"
							 tmpWriteStr=tmpWriteStr&trim(rs.Fields.item(h).Name)
						next
						 tmpWriteStr=tmpWriteStr&Chr(13)

						Set objTextFile = Fso.CreateTextFile(SQL_UpFilePath & sys_City & " - " & gInitDt(date) & replace(nowTime,":","") & " - " & trim(dirlist(i)), True)
						objTextFile.WriteLine(tmpWriteStr)
						
						While Not rs.eof
							tmpWriteStr=""
							For h=0 to rs.Fields.count-1
								If h>0 Then tmpWriteStr=tmpWriteStr& "	"
								tmpWriteStr=tmpWriteStr& trim(rs.Fields.item(h).value)
							next
							tmpWriteStr=tmpWriteStr& Chr(13)
							objTextFile.WriteLine(tmpWriteStr)
							rs.movenext
						Wend
					else
						Set objTextFile = Fso.CreateTextFile(SQL_UpFilePath & sys_City & " - " & gInitDt(date) & replace(nowTime,":","") & " - " & trim(dirlist(i)), True)
						objTextFile.WriteLine("查無資料!!")
					End if
					rs.close
				end if
			End if
		next
		If Fso.FileExists(SQL_LogPath&gInitDt(date)&"@" & hour(nowTime)&".txt") then
			Fso.DeleteFile SQL_LogPath & gInitDt(date)&"@" & hour(nowTime)&".txt"
		end if
	End if
	Err.Clear%>
