<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_RAREWORD.txt"

Response.AddHeader "Content-Disposition","attachment;filename=" &fname
Response.ContentType = "application/vnd.ms-txt"

strWhere=""
	If (not ifnull(Request("Sys_TD_RecordDate1"))) and (not ifnull(Request("Sys_TD_RecordDate2"))) Then
		TD_RecordDate1=gOutDT(request("Sys_TD_RecordDate1"))&" 0:0:0"
		TD_RecordDate2=gOutDT(request("Sys_TD_RecordDate2"))&" 23:59:59"

		If ifnull(strWhere) Then
			strWhere=" where TD_RecordDate between TO_DATE('"&TD_RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&TD_RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		else
			strWhere=strWhere&" and TD_RecordDate between TO_DATE('"&TD_RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&TD_RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		End if
	End if

	If not ifnull(Request("Sys_TD_RecordCity")) Then
		If ifnull(strWhere) Then
			strWhere=" where TD_RecordCity='"&trim(Request("Sys_TD_RecordCity"))&"'"
		else
			strWhere=strWhere&" and TD_RecordCity='"&trim(Request("Sys_TD_RecordCity"))&"'"
		End if
	End if

	If not ifnull(Request("Sys_TD_CARNO")) Then
		If ifnull(strWhere) Then
			strWhere=" where TD_CARNO='"&trim(Request("Sys_TD_CARNO"))&"'"
		else
			strWhere=strWhere&" and TD_CARNO='"&trim(Request("Sys_TD_CARNO"))&"'"
		End if
	End if

	If not ifnull(Request("Sys_TD_PROCESS")) Then
		If ifnull(strWhere) Then
			strWhere=" where TD_PROCESS='"&trim(Request("Sys_TD_PROCESS"))&"'"
		else
			strWhere=strWhere&" and TD_PROCESS='"&trim(Request("Sys_TD_PROCESS"))&"'"
		End if
	End if

	strSQL="update TDDT_RAREWORD set TD_PROCESS='1'"&strWhere
	conn.execute(strSQL)
	
	strSQL="select TD_SN,TD_CARNO,TD_OwnerName,TD_ADDRESS from TDDT_RAREWORD"&strWhere&" order by TD_RecordCity,TD_RecordDate DESC"
	set rsfound=conn.execute(strSQL)
	filecnt=0
	While Not rsfound.Eof
		filecnt=filecnt+1
		response.write filecnt &","&rsfound("TD_CARNO")&","&rsfound("TD_OwnerName")&","&rsfound("TD_ADDRESS")&vbnewline
		rsfound.MoveNext
	Wend
	rsfound.close

set rsfound=nothing
conn.close
set conn=nothing
%>