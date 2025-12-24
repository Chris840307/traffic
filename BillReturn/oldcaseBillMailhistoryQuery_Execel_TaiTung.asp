<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
	fMnoth=month(now)
	if fMnoth<10 then fMnoth="0"&fMnoth
	fDay=day(now)
	if fDay<10 then	fDay="0"&fDay
	fname=year(now)&fMnoth&fDay&"_舊資料送達紀錄(不上傳).xls"
	Response.AddHeader "Content-Disposition", "filename="&fname
	response.contenttype="application/x-msexcel; charset=MS950"

	strSQL=request("Sys_SQL")
	set rsfound=conn.execute(strSQL)
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE> 舊資料上傳查詢(不上傳)</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
</HEAD>
<BODY>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td>
			<table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
				<tr align="center">
					<th>批號</th>
					<th>序號</th>
					<th>註記日期</th>
					<th>單號</th>
					<th>車號</th>
					<th>送達原因</th>
					<th>檔名</th>
					<th>狀態</th>
				</tr><%
					While Not rsfound.eof
						response.write "<tr>"
						response.write "<td>"&trim(rsfound("BatchNo"))&"</td>"
						response.write "<td>"&trim(rsfound("SninDCIFile"))&"</td>"
						response.write "<td>"&gInitDT(trim(rsfound("RecordDate")))&"</td>"
						response.write "<td>"&trim(rsfound("BillNo"))&"</td>"
						response.write "<td>"&trim(rsfound("CarNo"))&"</td>"
						response.write "<td>"&trim(rsfound("ReaSonName"))&"</td>"
						response.write "<td>"&trim(rsfound("FileName"))&"</td>"
						response.write "<td>"&trim(rsfound("StatusName"))&"</td>"
						response.write "</tr>"
						rsfound.movenext
					Wend%>
				</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>