<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_"&Trim(request("date1"))&"建檔日_行人攤販_舉發單資料.xls"

Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
'Response.AddHeader "Content-Disposition", "attachment;filename=" &fname
'response.contenttype="application/vnd.ms-txt" 

Server.ScriptTimeout = 650000

%>
<%
	'檢查是否可進入本系統
	'AuthorityCheck(234)

	strSQL="select a.sn,a.billno,a.Carno,a.BillTypeID,a.IllegalDate,a.BillMemID1,a.BillMem1,a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.BillMemID4,a.BillMem4,a.Rule1,a.Rule2,a.IllegalAddress,a.MemberStation,a.BillFillDate,a.BillUnitiD,a.DeallineDate,a.CarSimpleID,a.SignType,a.RecordStateID,a.Recorddate,a.RecordMemberID,a.note,a.Driver,a.DriverID from Passerbase a where a.Recorddate between to_date('"&gOutDT(request("date1"))&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&gOutDT(request("date2"))&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
	" and billno is not null and RecordStateid=0"
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單查詢</title>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
</head>
<body>
<table width="100%" border="1">
<tr><td>告發單號</td><td>車號</td><td>違規時間</td><td>違規路段</td><td>違反法條一</td><td>違反法條二</td><td>違規人證號</td><td>違規人姓名</tr>
<%  
					If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
					Response.flush
%>
	<tr>
<%
					DciReturnStation=""
					CaseInDate=""
					IllegalMemID=""
					IllegalMem=""
					IllegalAddress=""
					OwnerName=""
					OwnerAddress=""
					DciCarTypeID=""
					SecondAddress=""
					
					%><td><%
					'告發單號
					response.write rsfound("BillNo")&"&nbsp;"
					%></td><td><%'車號
					response.write trim(rsfound("Carno"))&"&nbsp;"
					
					%></td><td><%'違歸時間
					if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
						response.write year(rsfound("IllegalDate"))-1911&"/"&Month(rsfound("IllegalDate"))&"/"&day(rsfound("IllegalDate"))&" "&hour(rsfound("IllegalDate"))&":"&minute(rsfound("IllegalDate"))&":00&nbsp;"
					end if	
					%></td><td><%'違規路段
					response.write trim(rsfound("IllegalAddress"))
					%></td><td><%'違反法條
					response.write trim(rsfound("Rule1"))&"&nbsp;"
					%></td><td><%'法條二
					response.write rsfound("rule2")&"&nbsp;"
					%></td><td><%'違規人證號
					
					response.write rsfound("DriverID")&"&nbsp;"
					%></td><td><%'違規人姓名
					response.write rsfound("Driver")
					%></td>
	</tr>
					<%
				
				'response.write vbCrLf
				rsfound.MoveNext
				Wend
				rsfound.close
				set rsfound=nothing
				%>
				
</body>
</html>
<%
conn.close
set conn=nothing
%>