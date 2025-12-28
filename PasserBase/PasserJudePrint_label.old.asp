<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
fMnoth=month(now)
if fMnoth<10 then
fMnoth="0"&fMnoth
end if
fDay=day(now)
if fDay<10 then
fDay="0"&fDay
end if
fname=year(now)&fMnoth&fDay&"_裁決書.doc"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/msword; charset=MS950" 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>裁決書</title>
<style type="text/css">
<!--
.style1 {
	font-size: 16px;
	font-weight: bold;
}
.style2 {font-family: "標楷體"}
.style3 {font-size: 18px}
.style4 {font-family: "標楷體"}
.style5 {font-size: 18px}
.style6 {font-family: "標楷體"; font-size: 18px; }
.style7 {
	font-family: "標楷體";
	font-weight: bold;
	font-size: 22px;
}
.style8 {
	font-family: "標楷體";
	font-size: 16px;
}
.style9 {font-family: "標楷體"}
.style10 {font-size: 16px}
.style11 {font-size: 14px}
.style12 {
	font-size: 14px;
	font-family: "標楷體";
	font-weight: bold;
}
.style13 {font-size: 16px; font-family: "標楷體"; }
.style14 {
	font-size: 30px;
	font-family: "標楷體";
}
.style15 {font-family: "標楷體"; font-size: 22px; }
.style16 {font-family: "標楷體"; font-size: 18px; }
.style17 {font-family: "標楷體"; font-size: 23px; }
.style18 {font-family: "標楷體"; font-size: 24px; }
.style19 {font-size: 24px}
.style20 {font-size: 36px}
.style21 {font-size: 18px}
-->
</style>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body background="..\image\007.gif">
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

strState="select * from PasserJude where BillSN="&trim(request("PBillSN"))
set rsState=conn.execute(strState)
if not rsState.eof then
	UOpenGovNumber=trim(rsState("OpenGovNumber"))
	UJudeDate=split(gArrDT(rsState("JudeDate")),"-")
	UAgentName=trim(rsState("AgentName"))
	UAgentBirth=trim(rsState("AgentBirth"))
	UAgentID=trim(rsState("AgentID"))
	UAgentAddress=trim(rsState("AgentAddress"))
	UBigUnitBossName=trim(rsState("BigUnitBossName"))
	USubUnitSecBossName=trim(rsState("SubUnitSecBossName"))
	UContactTel=trim(rsState("ContactTel"))
	UForFeit=trim(rsState("ForFeit"))
	UDutyUnit=trim(rsState("DutyUnit"))
	USendAddress=trim(rsState("SendAddress"))
	UPunishmentMainBody=trim(rsState("PunishmentMainBody"))
	USimpleReson=trim(rsState("SimpleReson"))
	UNote=trim(rsState("Note"))
end if
rsState.close
set rsState=nothing
PrintDate=split(gArrDT(date),"-")
strSql="select * from PasserBase where SN="&trim(request("PBillSN"))
set rsSql=conn.execute(strSql)
if Not rsSql.eof then
	If not ifnull(Trim(rsSql("DriverID"))) Then
		If Mid(Trim(rsSql("DriverID")),2,1)="1" Then
			UAgentSex="男"
		elseif Mid(Trim(rsSql("DriverID")),2,1)="2" Then
			UAgentSex="女"
		End if
	End if
end if
'strUInfo="select * from UnitInfo where UnitID='"&trim(rsSql("BillUnitID"))&"'"
'set rsUInfo=conn.execute(strUInfo)
'if not rsUInfo.eof then
'	theSubUnitSecBossName=trim(rsUInfo("SecondManagerName"))
'	theBigUnitBossName=trim(rsUInfo("ManageMemberName"))
'	theContactTel=trim(rsUInfo("Tel"))
'	theBankAccount=trim(rsUInfo("BankAccount"))
'	thenBillUnitName=trim(rsUInfo("UnitName"))
'end if
'rsUInfo.close
'set rsUInfo=nothing

thenPasserCity="":thenUnitName=""
strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsunit=conn.execute(strSQL)
If Not rsunit.eof Then
	Sys_UnitID=trim(rsunit("UnitID"))
	Sys_UnitLevelID=trim(rsunit("UnitLevelID"))
	Sys_UnitTypeID=trim(rsunit("UnitTypeID"))
End if
rsunit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set rsUnit=conn.Execute(strSQL)
DB_UnitTel=trim(rsUnit("Tel"))
theSubUnitSecBossName=trim(rsUnit("SecondManagerName"))
theBigUnitBossName=trim(rsUnit("ManageMemberName"))
Sys_UnitID=trim(rsunit("UnitID"))
Bank_BankName=trim(rsUnit("BankName"))
bank_BankAccount=trim(rsUnit("BankAccount"))
DutyUnitName=trim(rsUnit("UnitName"))
DutyAddress=trim(rsUnit("Address"))
if Not rsunit.eof then
	for j=1 to len(trim(rsunit("UnitName")))
		if j<>1 then thenPasserUnit=thenPasserUnit&"　"
		thenPasserUnit=thenPasserUnit&Mid(trim(rsunit("UnitName")),j,1)
	next
end if
rsunit.close

strUInfo="select * from Apconfigure where ID=35"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	for j=1 to len(trim(rsUInfo("value")))
		if j<>1 then thenPasserCity=thenPasserCity&"　"
		thenPasserCity=thenPasserCity&Mid(trim(rsUInfo("value")),j,1)
		PasserCity=PasserCity&Mid(trim(rsUInfo("value")),j,1)
	next
end if
rsUInfo.close

strSql="select confiscate from PasserConfiscate where BIllSN="&trim(rsSql("SN"))
set rsfast=conn.execute(strsql)
fastring=""
while Not rsfast.eof
	if trim(fastring)<>"" then fastring=fastring&","
	fastring=fastring&rsfast("confiscate")
	rsfast.movenext
wend
rsfast.close

strMem="select MANAGEMEMBERNAME from UnitInfo where UnitID='"&Sys_UnitID&"'"
set rsUnit=conn.execute(strMem)
if Not rsUnit.eof then MemUnitName=rsUnit("MANAGEMEMBERNAME")
rsUnit.close

strSql="select ImageFilename as MemberFileName from MemberData where MemberID="&Session("User_ID")
set mem=conn.execute(strsql)
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

strSql="select ImageFilename as MemberFileName from MemberData where Chname like '%"&MemUnitName&"%'"
set mem=conn.execute(strsql)
if Not mem.eof then Sys_MemUnitFileName=trim(mem("MemberFileName"))
mem.close

strSQL="select WordNum from UnitInfo Where UnitID='"&Session("Unit_ID")&"'"
set rs=conn.execute(strSQL)
If not rs.eof Then
	If Not ifnull(rs("WordNum")) Then BillPageUnit=trim(rs("WordNum"))
end if
rs.close

%>
<table width="635" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td nowrap><div align="center" class="style14"><%=thenPasserCity%>　<%=thenPasserUnit%></div></td>
  </tr>
  <tr valign="bottom">
    <td><div align="center" class="style14">違反道路交通管理事件裁決書</div></td>
  </tr>
  <tr>
    <td><div class="style13">　　　　　　　　中　華　民　國　<%=UJudeDate(0)&UJudeDate(1)&UJudeDate(2)%>　　<%=BillPageUnit&"交裁字第"&UOpenGovNumber&"號"%></div></td>
  </tr>
</table>
<table width="665" border="1" cellspacing=0 cellpadding=0>
  <tr>
    <td width="98" nowrap><span class="style13">受處分人<br>姓　　名</span></td>
    <td><span class="style13"><%=trim(rsSql("DRIVER"))%></span></td>
    <td width="152" nowrap><span class="style13">舉發通知單<br>編　　　號</span></td>
    <td width="152"><span class="style13">第<%=trim(rsSql("BillNo"))%>號</span></td>
	<td rowspan=3 width="250" align="center" valign="top"><span class="style13">承辦人：<br><%
	If sys_City<>"澎湖縣" and sys_City<>"高雄縣" and sys_City<>"嘉義市" and sys_City<>"台東縣" Then
		if trim(Sys_MemUnitFileName)<>"" then
			response.write "<img src=""..\Member\Picture\"&Sys_MemUnitFileName&""" width=""90"" height=""30"">"
		else
			'南投竹山分局
			'南投竹山分局  從這邊可以設定 操作者與承辦人不同人
			if Session("Unit_ID") = "05FG" then 
				response.write "蘇永信"
			elseIf sys_City<>"台南市" and Session("Unit_ID") <> "F000" Then
				response.write request("Session_JudeName")
			end if
		end if
	else
		response.write "　　　　　　　　"
	end if%></span></td>
  </tr>
  <tr>
    <td><span class="style13">地址</span></td>
	<td colspan=3><span class="style13"><%=trim(rsSql("DriverZip"))&trim(rsSql("DriverAddress"))%>&nbsp;</span></td>
  </tr>
  <tr>
	<td nowrap><span class="style13">身分證字號</span></td>
    <td><span class="style13"><%=trim(rsSql("DriverID"))%></span></td>
    <td nowrap><span class="style13">出生日期</span></td>
    <td width="58"><span class="style13"><%=gInitDT(trim(rsSql("DriverBirth")))%>&nbsp;</span></td>
  </tr>
   <tr>
    <td><span class="style13">法定代理人<br>姓　　名</span></td>
	<td colspan=3><span class="style13">&nbsp;</span></td>
	<td rowspan=3 width="250" align="center" valign="top"><span class="style13"><%
		If sys_City<>"嘉義市" Then
			if Session("Unit_ID") = "05FG" then
				response.write "分局長代行："
			elseif sys_City = "高雄縣" then
				response.write "組長："
			elseif sys_City="台東縣" then
				response.write "單位主管："
			else
				response.write "單位主官："
			end if
			response.write "<br>"
			If sys_City<>"澎湖縣" and sys_City<>"台東縣" Then
				if trim(Sys_MemUnitFileName)<>"" then
					response.write "<img src=""..\Member\Picture\"&Sys_MemUnitFileName&""" width=""90"" height=""30"">"
				elseif Session("Unit_ID") <> "F000" then
					response.write theSubUnitSecBossName
				end if
			end if
		end if%>&nbsp;</span></td>
  </tr>
   <tr>
    <td><span class="style13">法定代理人<br>地　　址</span></td>
	<td colspan=3><span class="style13">&nbsp;</span></td>
  </tr>
  <tr>
    <td nowrap><span class="style13">違規日期</span></td>
    <td nowrap><span class="style13">
      <%
		if trim(rsSql("IllegalDate"))<>"" then
			IllegalDate=split(gArrDT(rsSql("IllegalDate")),"-")
			response.write IllegalDate(0)&"年"&IllegalDate(1)&"月"&IllegalDate(2)&"日"
		end if%>
	</span></td>
	<td nowrap><span class="style13">違規時間</span></td>
    <td><span class="style13">
      <%
		if trim(rsSql("IllegalDate"))<>"" then
			IllegalDate=split(gArrDT(rsSql("IllegalDate")),"-")
			response.write hour(rsSql("IllegalDate"))&"時"&minute(rsSql("IllegalDate"))&"分"
		end if%>
	</span></td>
  </tr>
  <tr>
    <td nowrap><span class="style13">違規地點</span></td>
    <td colspan=3><span class="style13"><%=trim(rsSql("IllegalAddress"))%></span></td>
	<td rowspan="4" width="250" align="center" valign="top"><span class="style13"><%
			response.write "機關首長："&left("　　　　　　　　　　",10)%>&nbsp;</span></td>
  </tr>
  <tr>
	<td nowrap><span class="style13">應到案日期</span></td>
    <td colspan="3"><span class="style13">
      <%
		if trim(rsSql("DealLineDate"))<>"" then
			DealLineDate=split(gArrDT(rsSql("DealLineDate")),"-")
			response.write DealLineDate(0)&"年"&DealLineDate(1)&"月"&DealLineDate(2)&"日前"
		end if%>
    &nbsp;</span></td>
  </tr>
  <tr>
	<td><span class="style13">違反法條</span></td>
    <td colspan="3"><span class="style13">
		道路交通管理處罰條例<%
		Sys_Rule1=trim(rsSql("Rule1"))
		response.write "第"&left(trim(Sys_Rule1),2)&"條"
		if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
		response.write "第"&Mid(trim(Sys_Rule1),3,1)&"項第"&Mid(trim(Sys_Rule1),4,2)&"款規定"
		%>
	&nbsp;</span></td>
	</tr>
  <tr>
    <td nowrap><span class="style13">違規事實</span></td>
    <td colspan="3"><span class="style13">
      <%
		if trim(rsSql("Rule1"))<>"" and not isnull(rsSql("Rule1")) then
			strRule1="select * from Law where ItemID='"&trim(rsSql("Rule1"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				response.write trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		end if	
		if trim(rsSql("Rule2"))<>"" and not isnull(rsSql("Rule2")) then
			strRule1="select * from Law where ItemID='"&trim(rsSql("Rule2"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				response.write "<br>"&trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		end if	
		if trim(rsSql("Rule3"))<>"" and not isnull(rsSql("Rule3")) then
			strRule1="select * from Law where ItemID='"&trim(rsSql("Rule3"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				response.write "<br>"&trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		end if	
		if trim(rsSql("Rule4"))<>"" and not isnull(rsSql("Rule4")) then
			strRule1="select * from Law where ItemID='"&trim(rsSql("Rule4"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				response.write "<br>"&trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		end if%>
	</span></td>
  </tr>
  <tr>
    <td nowrap><span class="style13">處罰主文</span></td>
    <td colspan="4"><span class="style13"><%=UPunishmentMainBody%></span></td>
  </tr>
   <tr>
    <td nowrap><span class="style13">裁決日期</span></td>
    <td colspan="4"><span class="style13">中華民國<%=UJudeDate(0)%>年<%=UJudeDate(1)%>月<%=UJudeDate(2)%>日</span></td>
  </tr>
  <tr>
    <td nowrap><span class="style13">簡要理由</span></td>
    <td colspan="4"><span class="style13"><%=USimpleReson%></span></td>
  </tr>
  <tr>
    <td><div><span class="style13">附　　記</span></div></td>
	<!---------- smith 20091011 修改為20天  ---->
    <td colspan="4" valign="top"><span class="style13">
	一、受處分人不服本裁決者，應以原處分機關（<%=thenPasserCity&"<br>"%>
	&nbsp;&nbsp;&nbsp;&nbsp;<%=theUnitName%>）為被告，向原告住所地、居所地、所在地、違規行為地<br>
	&nbsp;&nbsp;&nbsp;&nbsp;、或原處分機關所在地之地方法院行政訴訟庭提起訴訟；其中撤銷訴訟之<br>
	&nbsp;&nbsp;&nbsp;&nbsp;提起，應於裁決書送達後30日之不變期間內為之。<br>

    二、請依處罰主文所定期限前持本裁決書至應到案處所、郵局繳納罰鍰。<br>

    三、上開罰鍰逾期仍不繳納者，本分局將依『行政執行法』<br>

	&nbsp;&nbsp;&nbsp;&nbsp;第二章第十一條移送行政執行分署強制執行。<br>
	<%
	if sys_City="宜蘭縣" then
		Response.Write "四、罰鍰繳納方式：1.可親自或委託他人至本分局臨櫃窗口繳納。<br>"
		Response.Write "　　　　　　　　　2.購買郵政匯票（抬頭請註明本分局全銜，<br>"	
		Response.Write "　　　　　　　　　例：宜蘭縣政府警察局宜蘭分局）郵寄至本分局繳納。"
	else
		Response.Write "四、『郵局劃撥』戶名："&Bank_BankName&"　<br>"
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;劃撥帳號："&bank_BankAccount
	end if
	%></span></td>
  </tr>
</table>
<hr>
<table width="635" border="0" cellpadding="0" cellspacing="0">
  <tr valign="bottom">
    <td colspan="3"><div align="center" class="style15">違反道路交通管理事件裁決書　行政文書</div></td>
  </tr>
  <tr>
    <td colspan="3"><div class="style16"><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=DutyAddress%>(<%=DB_UnitTel%>)</div></td>
  </tr>
  <tr>
    <td><div align="right" class="style16"><br>郵遞區號：</div></td>
	<td><div align="left" class="style16"><%=trim(rsSql("DriverZip"))%></div></td>
  </tr>
  <tr>
    <td><div align="right" class="style16">地　　址：</div></td>
	<td><div align="left" class="style16"><%=trim(rsSql("DriverAddress"))%></div></td>
  </tr>
  <tr>
    <td><div align="right" class="style16">收&nbsp;件&nbsp;人：</div></td>
	<td><div align="left" class="style16"><%=trim(rsSql("Driver"))%></div></td>
  </tr>
</table>
<br>
</body>
</html>