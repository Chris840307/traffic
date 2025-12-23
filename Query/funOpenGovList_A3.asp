<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="../smsx.cab#Version=6,1,432,1">
</object>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
body {font-family:新細明體;font-size:9pt}
.style1 {font-family:新細明體; line-height:21px; font-size: 18px}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>公告清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>

<%
Server.ScriptTimeout = 800
Response.flush
	Function getNameHidden(Names)
		strName=""
		if Mid(Names,2,2)="@@" then
			arrFont=split(Names,"@@")
			for FontRoop=0 to ubound(arrFont)
				If FontRoop=1 then
					strName=strName&"＊"
				ElseIf FontRoop=0 Or FontRoop=ubound(arrFont) Or InStr(arrFont(FontRoop),".png")=0 Then
					strName=strName&arrFont(FontRoop)
				Else
					strName=strName&"@@"&arrFont(FontRoop)&"@@"
				End If 
			Next
		elseif Mid(Names,1,2)="@@" then
			arrFont=split(Names,"@@")
			for FontRoop=0 to ubound(arrFont)
				If UBound(arrFont)>=3 Then	
					If InStr(arrFont(3),"png")>0 And Trim(arrFont(2))="" Then 
						cnt1=3
					Else
						cnt1=2
					End If 
				Else
					cnt1=2
				End If 
				If FontRoop=cnt1 Then
					If cnt1=2 And Len(Trim(arrFont(FontRoop)))>1 Then 
						strName=strName&"＊"&Right(Trim(arrFont(FontRoop)),Len(Trim(arrFont(FontRoop)))-1)
					Else 
						strName=strName&"＊"
					End If 					
				ElseIf FontRoop=0 Or FontRoop=ubound(arrFont) Or InStr(arrFont(FontRoop),".png")=0 Then
					strName=strName&arrFont(FontRoop)
				Else
					strName=strName&"@@"&arrFont(FontRoop)&"@@"
				End If 
			Next
		Else
			If Len(Names)=1 Then
				strName=Left(Names,1)
			ElseIf Len(Names)>1 then
				strName=Left(Names,1)&"＊"&Right(Names,Len(Names)-2)
			End If 
		End If 
		getNameHidden=strName
	End Function

'權限
'AuthorityCheck(234)
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

	if sys_City="台南市" then
		PageCount=20
	else
		PageCount=28
	end if
%>
<%
	strwhere=request("SQLstr")
%>
</head>
<body>
<form name=myForm method="post">
<%
PrintSN=0
PageNum=1
if sys_City="花蓮縣" then
	CloseDciReturnStatusID="a.DciReturnStatusID in ('S','N','k') "
else
	CloseDciReturnStatusID="a.DciReturnStatusID in ('S','N') "
end if
if sys_City="台東縣" then
	strSQL="select a.BillSN,a.BillNO,a.CarNO,e.Owner,e.DciReturnStation,f.BillTypeID,f.Rule1,f.Rule2,f.Rule3,f.Rule4,f.MemberStation,g.UserMarkResonID" &_
		",case when f.BillTypeID='1' then f.MemberStation when f.BillTypeID='2' then e.DcireturnStation end as St" &_
		" from (select a.BillSN,a.BillNo,a.CarNo,a.BillTypeID from DciLog a where a.BillSN is not null "&strwhere&" and a.ExchangeTypeID='N')" &_
		" a,BillBaseDCIReturn e,BillBase f,BillMailHistory g" &_
		" where a.BillSN=f.SN" &_
		" and f.RecordStateID=0 and f.SN=g.BillSn" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo" &_
		" and e.ExchangeTypeID='W' and ((e.Status in ('Y','S','n','L') and f.BillTypeID='2') or f.BillTypeID<>'2')" &_
		" and g.UserMarkResonID in ('1','2','3','4','8','M','K','L','O','P','Q','5','6','7','T')" &_
		" order by St,g.UserMarkDate"
else
	strSQL="select a.BillSN,a.BillNO,a.CarNO,e.Driver,e.Owner,e.DciReturnStation,f.BillTypeID,f.Rule1,f.Rule2,f.Rule3,f.Rule4,f.MemberStation,g.UserMarkResonID" &_
		" from (select a.BillSN,a.BillNo,a.CarNo,a.BillTypeID from DciLog a where a.BillSN is not null "&strwhere&" and a.ExchangeTypeID='N' and "&CloseDciReturnStatusID&")" &_
		" a,BillBaseDCIReturn e,BillBase f,BillMailHistory g" &_
		" where a.BillSN=f.SN" &_
		" and f.RecordStateID=0 and f.SN=g.BillSn" &_
		" and a.BillNo=e.BillNO and a.CarNo=e.CarNo" &_
		" and e.ExchangeTypeID='W' and ((e.Status in ('Y','S','n','L') and f.BillTypeID='2') or f.BillTypeID<>'2')" &_
		" and g.UserMarkResonID in ('1','2','3','4','8','M','K','L','O','P','Q','5','6','7','T')" &_
		" order by g.UserMarkDate"
end if
	set rs1=conn.execute(strSQL)
	If Not rs1.Bof Then rs1.MoveFirst 
	While Not rs1.Eof
	if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>
	<table width="100%" border="0" cellpadding="1" cellspacing="0" align="center">
	<tr>
	<td align="right">
	<center><span class="style1"><%
	strCity="select * from Apconfigure where ID=35"
	set rsCity=conn.execute(strCity)
	if not rsCity.eof then
		response.write rsCity("Value")
	end if
	rsCity.close
	set rsCity=nothing

	if sys_City="台南市" then
		strUnitName="select UnitName from UnitInfo where UnitID in (select UnitTypeID from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"')"
		set rsUnitName=conn.execute(strUnitName)
		if not rsUnitName.eof then
			response.write rsUnitName("UnitName")
		end if
		rsUnitName.close
		set rsUnitName=nothing
	else
		strUnitName="select UnitName from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
		set rsUnitName=conn.execute(strUnitName)
		if not rsUnitName.eof then
			response.write rsUnitName("UnitName")
		end if
		rsUnitName.close
		set rsUnitName=nothing
	end If
	if sys_City="花蓮縣" Then
		response.write "舉發違反道路交通管理事件通知單郵寄無法送達清冊"
	Else
		response.write "違規通知單郵寄無法送達清冊"
	End If 
	%></span></center>
	列印日期：<%=now%>
	<br>
	列印人員：<%
	strChName="select ChName from Memberdata where MemberID="&Session("User_ID")
	set rsChName=conn.execute(strChName)
	if not rsChName.eof then
		response.write rsChName("ChName")
	end if
	rsChName.close
	set rsChName=nothing
	%>
	<br>
	備註：因郵寄無法投遞，經查詢電腦資料，住址未辦理變更登記。
	</td>
	</tr>
	</table>
	<table width="100%" border="1" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center" width="5%" height="20">編號</td>
			<td align="center" width="9%">告發單號</td>
			<td align="center" width="9%">車號</td>
			<td align="center" width="25%">車主<%
			if sys_City="澎湖縣" Or sys_City="嘉義縣" then
				response.write "(駕駛人)"
			end if
			%>姓名</td>
			<td align="center" width="22%">違反法條代碼</td>
			<td align="center" width="13%">退件原因</td>
			<td align="center" width="17%">應到案處所</td>
		</tr>
<%		
	for i=1 to PageCount
		if rs1.eof then exit for
		PrintSN=PrintSN+1
%>		<tr>
			<td height="27"><%
			'編號
			response.write PrintSN
			%></td>
			<td><%
			'告發單號
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNO"))
			else
				response.write "&nbsp;"
			end if
			
			%></td>
			<td><%
			'車號
			if trim(rs1("CarNO"))<>"" and not isnull(rs1("CarNO")) then
				response.write trim(rs1("CarNO"))
			else
				response.write "&nbsp;"
			end if	
						
			%></td>
			<td><%
			'車主姓名
		if sys_City="澎湖縣" Or sys_City="嘉義市" Or sys_City="嘉義縣" then
			if trim(rs1("BillTypeID"))="2" then
				if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
					response.write funcCheckFont(rs1("Owner"),17,1)
				else
					response.write "&nbsp;"
				end if	
			else
				if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
					response.write "("&funcCheckFont(rs1("Driver"),17,1)&")"
				else
					response.write "&nbsp;"
				end if	
			end if
		else
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) Then
				If sys_City="高雄市" then 
					response.write funcCheckFont(getNameHidden(Trim(rs1("Owner"))),17,1)
				Else
					response.write funcCheckFont(rs1("Owner"),17,1)
				End If 
			else
				response.write "&nbsp;"
			end if	
		end if
			%></td>
			<td><%
			'法條一
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if	
			'法條二
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write "/"&trim(rs1("Rule2"))
			end if	
			'法條三
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				response.write "/"&trim(rs1("Rule3"))
			end if	
			'法條四
			if trim(rs1("BillTypeID"))="1" then
				if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
					response.write "/"&trim(rs1("Rule4"))
				end if	
			end if
			%></td>
			<td><%
			'退件原因
			strCode="select Content from DCIcode where TypeID=7 and ID='"&trim(rs1("UserMarkResonID"))&"'"
			set rsCode=conn.execute(strCode)
			if not rsCode.eof then
				ReturnReason=trim(rsCode("Content"))
			end if
			rsCode.close
			set rsCode=nothing

			if ReturnReason="" then
				response.write "&nbsp;"
			else
				response.write ReturnReason
			end if
			%></td>
			<td>
			<%
			if trim(rs1("BillTypeID"))="1" then
				strSqlStationName="select DCIstationName from Station where DCIstationID='"&trim(rs1("MemberStation"))&"'"
				set rsSN=conn.execute(strSqlStationName)
				if not rsSN.eof then
'					if StationNameArray="" then
'						StationNameArray=trim(rsSN("DCIstationName"))
'					else
'						StationNameArray=StationNameArray&","&trim(rsSN("DCIstationName"))
'					end if
					response.write trim(rsSN("DCIstationName"))
				end if
				rsSN.close
				set rsSN=nothing
			else	
				strSqlStationName="select DCIstationName from Station where DCIstationID='"&trim(rs1("DcireturnStation"))&"'"
				set rsSN=conn.execute(strSqlStationName)
				if not rsSN.eof then
'					if StationNameArray="" then
'						StationNameArray=trim(rsSN("DCIstationName"))
'					else
'						StationNameArray=StationNameArray&","&trim(rsSN("DCIstationName"))
'					end if
					response.write trim(rsSN("DCIstationName"))
				end if
				rsSN.close
				set rsSN=nothing
			end if
			%>
			</td>
		</tr>
<%			
		rs1.MoveNext
		next
%>
	</table>
	<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%
	Wend
	rs1.close
	set rs1=nothing
%>
</form>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}

printWindow(true,7,5.08,5.08,5.08);
</script>
<%conn.close%>