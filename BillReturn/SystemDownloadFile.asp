<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
	strfiletitle="select value from Apconfigure where id=100"
	set rsfiletitle=conn.execute(strfiletitle)
	filetitle=trim(rsfiletitle("value"))
	rsfiletitle.close
	DownFilePath="\DownProcess\"
	UpFilePath="\UpProcess\"
	'MoveFilePath="\Down\"
	nowTime=request("nowTime")
	FileType=".txt"
	reUpFileName=""
	Set Fso=CreateObject("Scripting.FileSystemObject")
	strSQL="select distinct FileName,RecordDate from OldCaseBillMailHistory where FileName is not null and Status is null"
	set rstxt=conn.execute(strSQL)
	While Not rstxt.eof
		If Not Fso.FileExists(server.mappath(DownFilePath & trim(rstxt("FileName"))&".big")) then
			If DateDiff("h",rstxt("RecordDate"),now)>6 Then
				If trim(reUpFileName)<>"" Then reUpFileName=reUpFileName&"','"
				reUpFileName=reUpFileName&trim(rstxt("FileName"))
			End if
		else
			Set f=Fso.OpenTextFile(server.mappath(DownFilePath & trim(rstxt("FileName"))&".big"),1,True,0)
			While Not f.AtEndOfStream
				tempstr=f.ReadLine
				Sys_Status=trim(Mid(tempstr,258,2))
				Sys_SninDciFile=trim(Mid(tempstr,1,6))
				Sys_CarNo=trim(Mid(tempstr,7,9))
				Sys_BillNo=trim(Mid(tempstr,16,10))
				Sys_FileName=trim(rstxt("FileName"))

				strSQL="Update OldCaseBillMailHistory set Status='"&Sys_Status&"' where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and FileName='"&Sys_FileName&"' and SninDciFile='"&Sys_SninDciFile&"'"
				conn.execute(strSQL)
			Wend
			f.close
'			If Fso.FileExists(server.mappath(MoveFilePath & trim(rstxt("FileName"))&".big")) then
'				Fso.DeleteFile server.mappath(MoveFilePath & trim(rstxt("FileName"))&".big")
'			end if
'			Fso.MoveFile server.mappath(DownFilePath & trim(rstxt("FileName"))&".big"), server.mappath(MoveFilePath)&"\"
		end if
		rstxt.movenext
	Wend
	rstxt.close
	If not ifnull(reUpFileName) Then
		tmpFileName=split(reUpFileName,"','")
		For i=0 to Ubound(tmpFileName)
			If Fso.FileExists(server.mappath(UpFilePath & trim(tmpFileName(i)))) then
				Fso.DeleteFile server.mappath(UpFilePath & trim(tmpFileName(i)))
			end if
		Next		
		strSQL="select NVL(Max(FileNameSeq),6) as FileNameSeq from OldCaseBillMailHistory where RecordDate between TO_DATE('"&dateValue(nowTime)&" :00:00:00','YYYY/MM/DD HH24:MI:SS') and TO_DATE('"&dateValue(nowTime)&" :23:59:59','YYYY/MM/DD HH24:MI:SS')"
		set rs=conn.execute(strSQL)
		Sys_FileNameSeq=cdbl(rs("FileNameSeq"))+1
		If Sys_FileNameSeq<10 Then
			Sys_FileSeq=Sys_FileNameSeq
		else
			Sys_FileSeq=Chr((Sys_FileNameSeq+55))
		End if
		Sys_FileName=filetitle&gInitDT(dateValue(nowTime)) & Sys_FileSeq &".X.F"
		rs.close
		strSQL="select * from OldCaseBillMailHistory where FileName in('"&reUpFileName&"') and Status is null order by RecordDate"
		set rsup=conn.execute(strSQL)
		filecmt=1
		While Not rsup.eof
			strSQL="Update OldCaseBillMailHistory set SninDCIFile='"&right("00000"&(filecmt),5)&"',FileName='"&Sys_FileName&"',FileNameSeq='"&Sys_FileNameSeq&"',RecordDate="&funGetDate(nowTime,1)&" where BillNo='"&trim(rsup("BillNo"))&"' and CarNo='"&trim(rsup("CarNo"))&"' and FileName='"&trim(rsup("FileName"))&"' and SninDciFile='"&trim(rsup("SninDciFile"))&"'"

			conn.execute(strSQL)

			filecmt=filecmt+1
			rsup.movenext
		Wend
		rsup.close

		Set objTextFile = Fso.CreateTextFile(Server.mappath(UpFilePath & Sys_FileName), True)
		strSQL="select * from OldCaseBillMailHistory where FileName like '"&Sys_FileName&"%' order by SninDCIFile"
		set rstxt=conn.execute(strSQL)
		While not rstxt.eof
			strInput=""
			strInput=strInput&left(left(trim(rstxt("SninDCIFile")),5)&"               ",6)
			strInput=strInput&left(left(trim(rstxt("CarNo")),8)&"               ",9)
			strInput=strInput&left(left(trim(rstxt("BillNo")),9)&"               ",10)
			strInput=strInput&left(left(trim(rstxt("ReaSonID")),3)&"               ",4)
			strInput=strInput&left("               ",11)
			strInput=strInput&left("0"&"               ",2)
			strInput=strInput&left("X"&"               ",2)
			strInput=strInput&left(left(trim(rstxt("LoginID")),6)&"               ",7)
			strInput=strInput&left(left(trim(rstxt("DOCNumber")),9)&"               ",10)
			strInput=strInput&left(right("0"&gInitDT(trim(rstxt("ProcessDate"))),7)&"               ",8)
			objTextFile.WriteLine(strInput)
			rstxt.movenext
		Wend
		objTextFile.Close
	end if
%>