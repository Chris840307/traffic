<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strSQL="select * from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
DB_UnitName=replace(trim(rsUnit("UnitName")),"交通組","")
DB_UnitTel=trim(rsUnit("Tel"))
DB_ManageMemberName=trim(rsUnit("ManageMemberName"))
rsUnit.close

strSql="select a.SN as BillSN,a.BillNo,a.Driver,a.DriverBirth,a.DriverID,a.DriverZip,a.DriverAddress,a.IllegalDate,a.IllegalAddress,a.DealLineDate,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillUnitID,b.OpenGovNumber as JudeOGN,b.AgentName as JudeAgentName,b.AgentSex as JudeAgentSex,b.AgentBirth as JudeAgentBirth,b.AgentID as JudeAgentID,b.AgentAddress as JudeAgentAddress,c.OpenGovNumber as UrgeOGN,c.UrgeTypeID,d.OpenGovNumber,d.BigUnitBossName,d.SubUnitSecBossName,d.SendNumber,d.SendDate,d.Agent,d.AgentBirthDate,d.AgentID,d.AgentAddress,d.ForFeit,d.MakeSureDate,d.LimitDate,d.AttatchJude,d.AttatchUrge,d.AttatchFortune,d.AttatchGround,d.AttatchRegister,d.AttatchFileList,d.AttatchTable,d.ATTATPOSTAGE,d.SafeToExit,d.SAFEACTION,d.SAFEASSURE,d.SAFEDETAIN,d.SAFESHUTSHOP from PasserBase a,PasserJude b,PasserUrge c,PasserSend d where a.SN="&trim(request("PBillSN"))&" and a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and a.SN=c.BillSN(+) and a.BillNo=c.BillNo(+) and a.SN=d.BillSN(+) and a.BillNo=d.BillNo(+)"
PrintDate=split(gArrDT(date),"-")
set rsfound=conn.execute(strSql)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>無標題文件</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-family: "標楷體"; font-size: 14px; }
.style2 {font-family: "標楷體"; font-size: 25px; }
.style3 {font-family: "標楷體"; font-size: 16px; }
-->
</style>
</head>

<body>
<div id="Layer1" class="style1" style="position:absolute; left:550px; top:50px; width:202px; height:36px; z-index:5"><%=rsfound("SendNumber")%></div><%'移送案號%>
<div id="Layer2" class="style2" style="position:absolute; left:250px; top:80px; width:202px; height:36px; z-index:5"><%=DB_UnitName%></div>
<div id="Layer3" class="style3" style="position:absolute; left:475px; top:130px; width:202px; height:36px; z-index:5"><%=PrintDate(0)&"年"&PrintDate(1)&"月"&PrintDate(2)&"日"%></div><%'發文日期%>
<div id="Layer4" class="style3" style="position:absolute; left:475px; top:150px; width:202px; height:36px; z-index:5"><%=rsfound("OpenGovNumber")%></div><%'發文字號%>
<div id="Layer5" class="style3" style="position:absolute; left:160px; top:215px; width:202px; height:36px; z-index:5"><%=rsfound("Driver")%></div><%'姓名或名稱%>
<div id="Layer6" class="style3" style="position:absolute; left:460px; top:215px; width:202px; height:36px; z-index:5"><%
	if trim(rsfound("Agent"))<>"" then
		response.write rsfound("Agent")
	else
		response.write rsfound("JudeAgentName")
	end if
%></div><%'法定代理人%>
<div id="Layer7" class="style3" style="position:absolute; left:160px; top:250px; width:202px; height:36px; z-index:5"><%
	if trim(rsfound("DriverBirth"))<>"" then
		DriverBirth=split(gArrDT(rsfound("DriverBirth")),"-")
		response.write DriverBirth(0)&"年"&DriverBirth(1)&"月"&DriverBirth(2)&"日"
	end if
%></div><%'出生年月日%>
<div id="Layer8" class="style3" style="position:absolute; left:460px; top:250px; width:202px; height:36px; z-index:5"><%
	if trim(rsfound("AgentBirthDate"))<>"" then
		AgentBirthDate=split(gArrDT(rsfound("AgentBirthDate")),"-")
	else
		AgentBirthDate=split(gArrDT(rsfound("JudeAgentBirth")),"-")
	end if
	if trim(AgentBirthDate(0))<>"" then
		response.write "　"&AgentBirthDate(0)&"年"&AgentBirthDate(1)&"月"&AgentBirthDate(2)&"日"
	end if
%></div><%'出生年月日%>
<div id="Layer9" class="style3" style="position:absolute; left:160px; top:350px; width:202px; height:36px; z-index:5"><%=rsfound("DriverID")%></div><%'身分證統一編號或%>
<div id="Layer10" class="style3" style="position:absolute; left:460px; top:350px; width:202px; height:36px; z-index:5"><%
	if trim(rsfound("AgentID"))<>"" then
		response.write rsfound("AgentID")
	else
		response.write rsfound("JudeAgentID")
	end if
%></div><%'身分證統一編號或%>
<div id="Layer11" class="style3" style="position:absolute; left:170px; top:400px; width:202px; height:36px; z-index:5"><%=trim(rsfound("DriverZip"))&trim(rsfound("DriverAddress"))%></div><%'戶籍地%>
<div id="Layer12" class="style3" style="position:absolute; left:460px; top:400px; width:202px; height:36px; z-index:5"><%
	if trim(rsfound("AgentAddress"))<>"" then
		response.write rsfound("AgentAddress")
	else
		response.write rsfound("JudeAgentAddress")
	end if
%></div><%'戶籍地%>
<div id="Layer13" class="style3" style="position:absolute; left:160px; top:550px; width:202px; height:36px; z-index:5"><%=rsfound("IllegalAddress")%></div><%'執行標的物所在地%>
<div id="Layer14" class="style3" style="position:absolute; left:550px; top:550px; width:202px; height:36px; z-index:5"><%
	if trim(rsfound("SendDate"))<>"" then
		SendDate=split(gArrDT(rsfound("SendDate")),"-")
		response.write SendDate(0)&"年"&SendDate(1)&"月"&SendDate(2)&"日"
	end if
%></div><%'執行處收案日期%>
<div id="Layer15" class="style3" style="position:absolute; left:160px; top:610px; width:350px; height:36px; z-index:5"><%
	response.write "違反道路交通管理處罰條例第"&left(trim(rsfound("Rule1")),2)&"條"
	if len(trim(rsfound("Rule1")))>7 then response.write "之"&right(trim(rsfound("Rule1")),1)
	response.write "<br>"
	response.write Mid(trim(rsfound("Rule1")),3,1)&"項"&Mid(trim(rsfound("Rule1")),4,2)&"款"&Mid(trim(rsfound("Rule1")),6,2)&"規定。"
	response.write "<br>"
	response.write "違反道路交通管理事件"
	if trim(rsfound("SendDate"))<>"" then
		IllegalDate=split(gArrDT(rsfound("IllegalDate")),"-")
		response.write IllegalDate(0)&"年"&IllegalDate(1)&"月"&IllegalDate(2)&"日"
	end if
	response.write "<br>"
	response.write "中分二交裁字第"&rsfound("JudeOGN")&"號裁決書"
%></div><%'執行處收案日期%>
<div id="Layer16" class="style3" style="position:absolute; left:550px; top:620px; width:202px; height:36px; z-index:5"><%
	if trim(rsfound("MakeSureDate"))<>"" then
		MakeSureDate=split(gArrDT(rsfound("MakeSureDate")),"-")
		response.write MakeSureDate(0)&"年"&MakeSureDate(1)&"月"&MakeSureDate(2)&"日"
	end if
%></div><%'確定日期%>
<div id="Layer17" class="style3" style="position:absolute; left:550px; top:665px; width:202px; height:36px; z-index:5"><%=rsfound("ForFeit")%></div><%'應納金額%>
<div id="Layer18" class="style3" style="position:absolute; left:190px; top:780px; width:202px; height:36px; z-index:5">依據道路交通管理處罰條例第<%=left(trim(rsfound("Rule1")),2)%>條</div><%'移送法條%>
<div id="Layer19" class="style2" style="position:absolute; left:510px; top:1035px; width:202px; height:36px; z-index:5"><%=DB_ManageMemberName%></div><%'應納金額%>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}
</script>