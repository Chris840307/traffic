<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>戶籍地址補正車籍資料列表</title>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
'權限
'AuthorityCheck(234)

fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_戶籍地址補正車籍資料清冊.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<%
Server.ScriptTimeout = 800
Response.flush
%>
<%
RecordDate=split(gInitDT(date),"-")
	strwhere=Session("PrintCarDataSQLxls")
	dcitype=trim(request("dcitype"))
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=Nothing

'	strdata=" and (substr(e.ownerid,2,1)<>'A' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'S' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'D' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'F' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'G' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'H' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'J' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'K' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'L' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'Z' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'X' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'C' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'V' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'B' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'N' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'M' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'Q' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'W' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'E' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'R' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'T' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'Y' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'U' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'I' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'O' "
'	strdata=strdata&" and substr(e.ownerid,2,1)<>'P' "
'	strdata=strdata&")"
'
'	strdata2=" and (substr(e.ownerid,1,1)='A' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='S' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='D' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='F' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='G' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='H' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='J' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='K' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='L' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='Z' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='X' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='C' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='V' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='B' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='N' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='M' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='Q' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='W' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='E' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='R' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='T' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='Y' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='U' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='I' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='O' "
'	strdata2=strdata2&" or substr(e.ownerid,1,1)='P' "
'	strdata2=strdata2&")"
	strdata2=strdata2&"  and (substr(e.ownerid,2,1) in ('1','2','3','4','5','6','7','8','9','0'))"
	strdata2=strdata2&" and (substr(e.ownerid,1,1) in ('A','S','D','F','G','H','J','K','L','Z','X','C','V','B','N','M','Q','W','E','R','T','Y','U','I','O','P',' '))"

	strSQL="select distinct e.billno,e.ownerid,a.SN,a.CarSimpleID,a.IllegalDate,a.Rule1,a.Rule2,a.IllegalAddress,a.RuleSpeed,a.IllegalSpeed,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillNo,a.RuleVer,e.CarNo,e.DCIReturnCarType,e.A_Name,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerAddress,e.OwnerZip,e.Nwner,e.NwnerID,e.NwnerAddress,e.NwnerZip,e.DCIReturnCarStatus from DCILog c,MemberData b,BillBase a,DCIReturnStatus d,BillBaseDCIReturn e where c.BillSN=a.SN and e.ExchangeTypeID='A' and e.Status='S' and a.CarNo=e.CarNo (+) and c.ExchangeTypeID=d.DCIActionID(+) and c.DCIReturnStatusID=d.DCIReturn(+) and c.RecordMemberID=b.MemberID(+) and a.RecordStateID=0 "&strdata&strdata2&" and (e.ownernotifyaddress is null or e.ownernotifyaddress='') "&strwhere&" order by a.RecordDate"
	set rsfound=conn.execute(strSQL)

	strCnt="select count(*) as cnt from (select distinct a.SN,a.CarSimpleID,a.IllegalDate,a.Rule1,a.Rule2,a.IllegalAddress,a.RuleSpeed,a.IllegalSpeed,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillNo,a.RuleVer,e.CarNo,e.DCIReturnCarType,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerAddress,e.OwnerZip,e.DCIReturnCarStatus from DCILog c,MemberData b,BillBase a,DCIReturnStatus d,BillBaseDCIReturn e where c.BillSN=a.SN and e.ExchangeTypeID='A' and e.Status='S' and a.CarNo=e.CarNo (+) and c.ExchangeTypeID=d.DCIActionID(+) and c.DCIReturnStatusID=d.DCIReturn(+) and c.RecordMemberID=b.MemberID(+) and a.RecordStateID=0 "&strdata&strdata2&" and (e.ownernotifyaddress is null or e.ownernotifyaddress='') "&strwhere&")"
	set Dbrs=conn.execute(strCnt)
	DBsum=Dbrs("cnt")
	Dbrs.close
	tmpSQL=strwhere
'response.write strSQL
%>

</head>
<body>
<form name=myForm method="post">
	<table width="100%" border="1" cellpadding="4" cellspacing="1">
		<tr>
			<td colspan="5" align="center">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr>
					<td colspan="5" align="center">
						<font size="3"><strong>戶籍地址補正車籍資料清冊</strong></font>
						(共 <%=DBsum%> 筆)
					</td>
				</tr>
				<tr>
					<td colspan="5" align="right">
						印表單位：<%
						UnitID=Session("Unit_ID")
						strUnit="select UnitName from UnitInfo where UnitID='"&UnitID&"'"
						set rsUnit=conn.execute(strUnit)
						if not rsUnit.eof then
							response.write trim(rsUnit("UnitName"))
						end if
						rsUnit.close
						set rsUnit=nothing
						%>
					</td>
				</tr>
				<tr>
					<td colspan="5" align="right">
						印表時間：<%=year(now)-1911%> - <%=month(now)%> - <%=day(now)%> - <%=hour(now)%> : <%=minute(now)%>
					</td>
				</tr>
			</table>
			</td>
		</tr>

		<tr>
			<td width="60" height="38">單號</td>
			<td width="38">車號</td>
			<td width="38">證號</td>
			<td width="65">車主姓名</td>
			<td width="250">車主地址</td>
		</tr>
		<%	ListSN=0
			if Not rsfound.eof then rsfound.move DBcnt
			While Not rsfound.Eof
				ListSN=ListSN+1
%>				<tr bgcolor="#ffffff">
					<td height="38" align="left"><%="&nbsp;"&rsfound("Billno")%></td>
					<td><%="&nbsp;"&rsfound("CarNo")%></td>
					<td><%="&nbsp;"&rsfound("ownerID")%></td>
					<td><%=funcCheckFont(rsfound("Owner"),20,0)%></td>
					<td><%
					'車主地址
					if (trim(rsfound("OwnerAddress"))<>"" and not isnull(rsfound("OwnerAddress"))) then
						response.write trim(rsfound("OwnerZip"))&funcCheckFont(rsfound("OwnerAddress"),20,0)
					end if
					%></td>

				</tr>
<%
			rsfound.MoveNext
		Wend
		rsfound.close
		set rsfound=nothing
		%>
		</tr>
	</table>
</form>
</body>
</html>
<%conn.close%>