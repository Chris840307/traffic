<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_Log紀錄列表.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

	strwhere=""
	strwhereM=""
	if request("Sys_CreditID")<>"" then
		strwhereM=" where CreditID = '"&request("Sys_CreditID")&"'"
	end if
	if request("Sys_ActionChName")<>"" then
		if strwhereM<>"" then
			strwhereM=strwhereM&" and ChName like '%"&request("Sys_ActionChName")&"%'"
		else
			strwhereM=" where ChName like '%"&request("Sys_ActionChName")&"%'"
		end if
	end If
	If strwhereM<>"" Then
		strwhere=strwhere&" and ActionMemberID in (select MemberID from MemberData "&strwhereM&")"
	End If 
	if request("Sys_IP")<>"" then
			strwhere=strwhere&" and ActionIP = '"&request("Sys_IP")&"'"
	end if
	if request("ActionDate")<>"" then
		ArgueDate1=gOutDT(request("ActionDate"))&" 0:0:0"
		ArgueDate2=gOutDT(request("ActionDate2"))&" 23:59:59"
			strwhere=strwhere&" and ActionDate between "&funGetDate(ArgueDate1,1)&" and "&funGetDate(ArgueDate2,1)
	end if
	if request("Sys_TypeID")<>"" then
			strwhere=strwhere&" and TypeID="&request("Sys_TypeID")
	end If
	If Trim(request("KeyWord"))<>"" Then
			strwhere=strwhere&" and ActionContent like '%"&Trim(request("KeyWord"))&"%' "
	End If 
	If Trim(request("ActionUnit"))<>"" Then
		strwhere=strwhere&" and ActionMemberID in (select MemberID from MemberData where UnitID in (select UnitID from UnitInfo where UnitTypeID='"&Trim(request("ActionUnit"))&"')) "
	End If 

	If Trim(request("chkSmith"))<>"" Then
		strwhere=strwhere&" and not (ActionContent like '%Billno=%' or ActionContent like '%Billno =%' or ActionContent like '%CarNo =%' or ActionContent like '%CarNo=%' or ActionContent like '%DriverID =%' or ActionContent like '%DriverID=%') "
	End If 

	strSQL="select * from Log where sn is not null "&strwhere&" order by ActionDate"
		
	set rsfound=conn.execute(strSQL)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>申訴案件</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="100%" height="100%" border="1">
	<tr>
		<td align="center"><strong>申訴案件紀錄列表</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" height="100%" border="1" cellpadding="4" cellspacing="1">
				<tr>
					<td>日期</td>
					<td>姓名</td>
					<td>身份證號</td>
					<td>IP</td>
					<td>類別</td>
					<td>內容</td>
				</tr><%
					while Not rsfound.eof
						response.write "<tr>"
						response.write "<td nowrap>"&gArrDT(rsfound("ActionDate"))&" "&Timevalue(rsfound("ActionDate"))&"</td>"
						response.write "<td nowrap>"&rsfound("ActionChName")&"</td>"
						response.write "<td nowrap>"
						CreditIDTmp=""
						If Trim(rsfound("ActionMemberID"))<>"" And Not IsNull(rsfound("ActionMemberID")) Then
							str="select * from MemberData where Memberid="&Trim(rsfound("ActionMemberID"))
							Set rs=conn.execute(str)
							If Not rs.eof Then
								response.write Trim(rs("ChName"))
								CreditIDTmp=left(Trim(rs("CreditID")),4)&"******"
							End If
							rs.close
							Set rs=Nothing 
						End If 
						response.write "</td>"
						response.write "<td nowrap>"&rsfound("ActionIP")&"</td>"
						response.write "<td nowrap>"
						If Trim(rsfound("TypeID"))="355" Then
							response.write "查詢"
						elseIf Trim(rsfound("TypeID"))="356" Then
							response.write "快速查詢"
						elseIf Trim(rsfound("TypeID"))="357" Then
							response.write "登入異常"
						elseIf Trim(rsfound("TypeID"))="358" Then
							response.write "帳號封鎖"
						elseIf Trim(rsfound("TypeID"))="359" Then
							response.write "人員異動"
						elseIf Trim(rsfound("TypeID"))="360" Then
							response.write "列印"
						elseIf Trim(rsfound("TypeID"))<>"" And Not IsNull(rsfound("TypeID")) then
							str="select * from Code where id="&Trim(rsfound("TypeID"))&" and TypeID=12"
							Set rs=conn.execute(str)
							If Not rs.eof Then
								response.write Trim(rs("Content"))
							End If
							rs.close
							Set rs=Nothing 
						End If 
						response.write "</td>"
						response.write "<td>"&Replace(rsfound("ActionContent"),"""","'")&"</td>"
						response.write "</tr>"
						rsfound.movenext
					wend%>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%conn.close%>