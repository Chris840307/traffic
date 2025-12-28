<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
'檢查是否可進入本系統
If Not ifnull(request("Sys_SendBillSN")) Then

	sys_billsn=request("Sys_SendBillSN")
elseif Not ifnull(request("hd_BillSN")) Then

	sys_billsn=request("hd_BillSN")
else

	sys_billsn=request("BillSN")
End If 

tmp_billsn=split(sys_billsn,",")

sys_billsn=""

For i = 0 to Ubound(tmp_billsn)

	If i >0 then

		If i mod 100 = 0 Then

			sys_billsn=sys_billsn&"@"
		elseif sys_billsn<>"" then

			sys_billsn=sys_billsn&","
		end If 
	end if

	sys_billsn=sys_billsn&tmp_billsn(i)

Next

tmpSQL=""

If Ubound(tmp_billsn) >= 100 Then

	sys_billsn=split(sys_billsn,"@")
	
	For i = 0 to Ubound(sys_billsn)
		
		If tmpSQL <>"" Then tmpSQL=tmpSQL&" union all "
		
		tmpSQL=tmpSQL&"select sn from passerbase where sn in("&sys_billsn(i)&")"
	Next

else

	tmpSQL="select sn from passerbase where sn in("&sys_billsn&")"

End if 

BasSQL="("&tmpSQL&") tmpPasser"

orderstr=request("orderstr")

strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
rsUInfo.close

strSQL="select WordNum from UnitInfo Where UnitID='"&Session("Unit_ID")&"'"
set rs=conn.execute(strSQL)
If not rs.eof Then
	If Not ifnull(rs("WordNum")) Then BillPageUnit=trim(rs("WordNum"))
end if
rs.close

strSQL="select a.chName,b.Content from MemberData a,(select ID,Content from Code where TypeID=4 ) b where MemberID='"&session("User_ID")&"' and a.JobID=b.ID(+)"
set mem=conn.execute(strSQL)
If not mem.eof Then
	chName=mem("chName")
	JobName=mem("Content")
end if
If ifnull(JobName) Then jobName="警員"
mem.close

strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
Sys_UnitID=trim(rsUnit("UnitID"))
Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
rsUnit.close

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set unit=conn.Execute(strSQL)
DB_UnitID=trim(unit("UnitID"))
DB_UnitName=trim(unit("UnitName"))
DB_UnitTel=trim(unit("Tel"))
DB_BankName=trim(unit("BankName"))
DB_BankAccount=trim(unit("BankAccount"))
DB_ManageMemberName=trim(unit("ManageMemberName"))
unit.close

strSql="select a.SN,a.Driver,a.DriverID,a.IllegalDate,a.Note," &_
		"(Select OpenGovNumber from PasserJude where billsn=a.sn) JudeOGN," &_
		"(Select SendNumber from PasserSend where billsn=a.sn) SendNumber," &_
		"(Select SendDate from PasserSend where billsn=a.sn) SendDate," &_
		"(Select ForFeit from PasserSend where billsn=a.sn) ForFeit," &_
		"(Select MakeSureDate from PasserSend where billsn=a.sn) MakeSureDate," &_
		"(Select LimitDate from PasserSend where billsn=a.sn) LimitDate" &_
		" from PasserBase a where Exists(select 'Y' from "&BasSQL&" where SN=a.SN) "&orderstr
PrintDate=split(gArrDT(date),"-")
set rsfound=conn.execute(strSql)
fileCnt=0
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>慢車行人道路障礙案件明細表</title>
<style type="text/css">
<!--
.style1 {font-family: "標楷體"; font-size: 14px; }
.style2 {font-family: "標楷體"; font-size: 20px; }
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>
<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%While Not rsfound.eof
	if cint(fileCnt)<>0 then response.write "<div class=""PageNext""></div>"
	if Not rsfound.eof then SendDate=split(gArrDT(rsfound("SendDate")),"-")%>
<center>
<span width="95%" class="style2"><%=thenPasserCity&replace(DB_UnitName,trim(thenPasserCity),"")%>移送行政執行『案件明細表』</span><span width="5%" class="style1"><%=SendDate(0)%>年<%=SendDate(1)%>月<%=SendDate(2)%>日</span>
</center>
<table width="100%" border="1" cellpadding="0" cellspacing="0">
	<tr>
		<th class="style1" nowrap>編 號</th>
		<th class="style1" nowrap>移送案號(總局文號）</th>
		<th class="style1" nowrap>執行處分
<br>之原機關</th>
		<th class="style1" nowrap>發文字號(分局處分書-裁決書字號）</th>
		<th class="style1" nowrap>義務人</th>
		<th class="style1" nowrap>身分證字號(義務人)</th>
		<th class="style1" nowrap>義務發生之日期</th>
		<th class="style1" nowrap>繳納期間<br>
屆滿日</th>
		<th class="style1" nowrap>應 納 金 額</th>
		<th class="style1" nowrap>備 考</th>
	</tr><%
		'if Not rsfound.eof then rsfound.move fileCnt+1
		For i=1 to 50
			if rsfound.eof then exit for
			fileCnt=fileCnt+1
			'MakeSureDate=gInitDT(DateAdd("d",20,rsfound("ArrivedDate")))
			'LimitDate=gInitDT(DateAdd("d",35,rsfound("ArrivedDate")))
			response.write "<tr>"
			response.write "<td class=""style1"">"&fileCnt& "&nbsp;</td>"
			response.write "<td class=""style1"">"&rsfound("SendNumber")&"&nbsp;</td>"
			response.write "<td class=""style1"">"&trim(replace(DB_UnitName,trim(thenPasserCity),""))&"&nbsp;</td>"
			response.write "<td class=""style1"">"&BillPageUnit&"裁字第"&rsfound("JudeOGN")&"號"&"&nbsp;</td>"
			response.write "<td class=""style1"" nowrap>"&trim(rsfound("Driver"))&"&nbsp;</td>"
			response.write "<td class=""style1"">"&trim(rsfound("DriverID"))&"&nbsp;</td>"
			response.write "<td class=""style1"">"&gInitDT(trim(rsfound("IllegalDate")))&"&nbsp;</td>"
			response.write "<td class=""style1"">"&gInitDT(trim(rsfound("LimitDate")))&"&nbsp;</td>"
			response.write "<td class=""style1"">"&trim(rsfound("ForFeit"))&"&nbsp;</td>"
			response.write "<td class=""style1"">"&trim(rsfound("Note"))&"&nbsp;</td>"
			response.write "</tr>"		
			rsfound.movenext
		next%>
</table>
<br>
<%Wend%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(false,0,0,0,0);
</script>
<%
rsfound.close
conn.close
set conn=nothing
%>
