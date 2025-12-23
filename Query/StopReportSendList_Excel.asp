<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>催繳資料清冊</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {
	font-size: 14pt;
	line-height: 20px;
}
.style2 {
	font-size: 11pt;
}

-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<%

'權限
'AuthorityCheck(234)

Server.ScriptTimeout=6000

RecordDate=split(gInitDT(date),"-")

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

	PageCount=21
	strwhere=trim(request("SQLstr"))

	strCnt="select count(*) as cnt from (select distinct a.*,e.DciReturnCarColor,e.Owner,e.OwnerAddress,e.OwnerZip,e.A_Name from BillBase a,DciLog b,BillBaseDciReturn e where a.CarNo=e.CarNo and e.ExchangeTypeID='A' and e.Status='S' and a.Sn=b.BillSn and a.RecordStateID=0 "&strwhere&")"
	set Dbrs=conn.execute(strCnt)
	DBsum=Dbrs("cnt")
	if trim(Dbrs("cnt"))="0" then
		pagecnt=1
	else
		pagecnt=fix(Cint(Dbrs("cnt"))/PageCount+0.9999999)
	end if
	Dbrs.close

	tmpSQL=strwhere

	strSQL="select distinct a.*,e.DciReturnCarColor,e.Owner,e.OwnerAddress,e.OwnerZip,e.DriverHomeZip,e.DriverHomeAddress,e.A_Name,e.OWNERNOTIFYADDRESS from BillBase a,DciLog b,BillBaseDciReturn e where a.CarNo=e.CarNo and e.ExchangeTypeID='A' and e.Status='S' and a.Sn=b.BillSn and a.RecordStateID=0 "&strwhere&" order by a.CarNo,a.RecordDate"
	set rsfound=conn.execute(strSQL)
'response.write strwhere
%>

</head>
<body>

<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;


承辦人：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				單位主管：
<%
mailSN=0
If Not rsfound.Bof Then rsfound.MoveFirst 
While Not rsfound.Eof
if mailSN>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
%>

	<table width="100%" border="0" cellspacing="0">
		<tr>
			<td height="10"></td>
		</tr>
		<tr>
		<td width="40%">
		<span class="style2">
		列印日期：<%=now%><br>
		列印單位：<%
		strPrintUnit="select UnitName from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
		set rsPrintUnit=conn.execute(strPrintUnit)
		if not rsPrintUnit.eof then
			response.write rsPrintUnit("UnitName")
		end if
		rsPrintUnit.close
		set rsPrintUnit=nothing
		%><br>
		列印人員：<%
		strUserID="select Chname from MemberData where MemberID='"&trim(Session("User_ID"))&"'"
		set rsUserID=conn.execute(strUserID)
		if not rsUserID.eof then
			response.write rsUserID("ChName")
		end if
		rsUserID.close
		set rsUserID=nothing
		%>
		</span>
		</td>
		<td width="30%">
		<span class="style1">催繳資料清冊</span>		
		</td>
		<td width="30%" align="right">
		<span class="style2">
		頁次 <%=fix(mailSN/PageCount)+1%> of <%=pagecnt%> &nbsp; &nbsp; 
		</span>
		</td>
		</tr>
	</table>
	<br>
	<table width="100%" border="0" cellspacing="0">
		<tr>
		<td width="14%"><span class="style2">停車日期時間</span></td>
		<td width="15%"><span class="style2">車號</span></td>
		<td width="56%"><span class="style2">停車地點</span></td>
		<td width="15%"><span class="style2">應繳停車費</span></td>
		</tr>
		<tr>
		<td><span class="style2">車種</span></td>
		<td><span class="style2">車主姓名</span></td>
		<td><span class="style2">地址</span></td>
		<td></td>
		</tr>
	</table>
	<hr>
<%		for i=1 to PageCount
			if rsfound.eof then exit for
			mailSN=mailSN+1
%>
	<table width="100%" border="0" cellspacing="0">
	<tr>
		<td width="14%"><span class="style2"><%
		response.write year(rsfound("IllegalDate"))-1911&"/"& month(rsfound("IllegalDate"))& "/" &day(rsfound("IllegalDate"))
		response.write "&nbsp; "&right("00"&hour(rsfound("IllegalDate")),2)&":"&right("00"&minute(rsfound("IllegalDate")),2)
		%></span></td>
		<td width="15%"><span class="style2"><%
		response.write rsfound("CarNo")
		%></span></td>
		<td width="56%"><span class="style2"><%
		response.write funcCheckFont(trim(rsfound("IllegalAddress")),17,1)		
		%></span></td>
		<td width="15%"><span class="style2"><%
		response.write trim(rsfound("Forfeit1"))		
		%></span></td>
	</tr>
	<tr>
		<td><span class="style2">
		<%
		if trim(rsfound("CarSimpleID"))="1" then
			response.write "汽車"
		elseif trim(rsfound("CarSimpleID"))="2" then
			response.write "拖車"
		elseif trim(rsfound("CarSimpleID"))="3" then
			response.write "重機"
		elseif trim(rsfound("CarSimpleID"))="4" then
			response.write "輕機"
		elseif trim(rsfound("CarSimpleID"))="6" then
			response.write "臨時車牌"
		end if
		%>
		</span></td>
		<td><span class="style2">
		<%
		If Not isNull(rsfound("Owner")) then
			If InStr(trim(rsfound("Owner")),"@@")>0 Then
				response.write funcCheckFont(trim(rsfound("Owner")),17,1)
			else
				response.write funcCheckFont(left(trim(rsfound("Owner")),7),17,1)
			End If
		End if
		%>
		</span></td>
		<td><span class="style2">
		<%
		if trim(rsfound("OWNERNOTIFYADDRESS"))<>"" and not isnull(rsfound("OWNERNOTIFYADDRESS")) then
			response.write "(通)"&funcCheckFont(trim(rsfound("OWNERNOTIFYADDRESS")),17,1)
		elseif trim(rsfound("DriverHomeAddress"))<>"" and not isnull(rsfound("DriverHomeAddress")) then
			response.write "(戶)"&trim(rsfound("DriverHomeZip"))&funcCheckFont(trim(rsfound("DriverHomeAddress")),17,1)
		else
			response.write "(車)"&trim(rsfound("OwnerZip"))&funcCheckFont(trim(rsfound("OwnerAddress")),17,1)
		end if
		%>
		</span></td>
		
	</tr>
		<tr>
			<td height="7"></td>
		</tr>
	</table>
<%			
		rsfound.MoveNext
		next

		response.flush
Wend
rsfound.close
set rsfound=nothing
%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(true,5.08,5.08,5.08,5.08);
</script>
<%conn.close%>