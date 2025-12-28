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
.style13 {font-size: 14px; font-family: "標楷體"; }
.style14 {
	font-size: 30px;
	font-family: "標楷體";
}
.style15 {font-family: "標楷體"; font-size: 28px; }
.style16 {font-family: "標楷體"; font-size: 20px; }
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
	strUInfo="select * from UnitInfo where UnitID='"&trim(rsState("DutyUnit"))&"'"
	set rsUInfo=conn.execute(strUInfo)
	if not rsUInfo.eof then
		DutyUnitName=trim(rsUInfo("UnitName"))
		DutyAddress=trim(rsUInfo("Address"))
	end if
	rsUInfo.close
	set rsUInfo=nothing
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

strUInfo="select * from UnitInfo where UnitID='"&trim(rsSql("BillUnitID"))&"'"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
'	theSubUnitSecBossName=trim(rsUInfo("SecondManagerName"))
'	theBigUnitBossName=trim(rsUInfo("ManageMemberName"))
'	theContactTel=trim(rsUInfo("Tel"))
'	theBankAccount=trim(rsUInfo("BankAccount"))
	thenBillUnitName=trim(rsUInfo("UnitName"))
end if
rsUInfo.close
set rsUInfo=nothing

strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
Sys_UnitID=trim(rsUnit("UnitID"))
Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
rsUnit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set unit=conn.Execute(strSQL)
If Not unit.eof Then
	theUnitID=trim(unit("UnitID"))
	theUnitName=trim(unit("UnitName"))
	theSubUnitSecBossName=trim(unit("SecondManagerName"))
	theBigUnitBossName=trim(unit("ManageMemberName"))
	theContactTel=trim(unit("Tel"))
	theBankAccount=trim(unit("BankAccount"))
	theBankName=trim(unit("BankName"))
	theUnitAddress=trim(unit("Address"))
end if
unit.close

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
set rsunit=conn.Execute(strSQL)

if Not rsunit.eof then
	for j=1 to len(trim(rsunit("UnitName")))
		if j<>1 then thenUnitName=thenUnitName&"　"
		thenUnitName=thenUnitName&Mid(trim(rsunit("UnitName")),j,1)
	next
end if
Sys_UnitID=trim(rsunit("UnitID"))

strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
set rsbank=conn.execute(strSQL)
Bank_BankName=trim(rsbank("BankName"))
bank_BankAccount=trim(rsbank("BankAccount"))
rsbank.close

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
<table width="645" border="0">
  <tr>
    <td colspan="2"><div align="center" class="style7"><%=thenPasserCity%>　<%=thenUnitName%></div></td>
  </tr>
  <tr valign="bottom">
    <td height="26" colspan="2"><div align="center" class="style2">違反道路交通管理事件裁決書</div></td>
  </tr>
  <tr>
    <td width="312" height="26"><span class="style2"></span></td>
    <td width="323"><div align="right" class="style2"><%=BillPageUnit&"交裁字第"&UOpenGovNumber&"號"%></div></td>
  </tr>
</table>
<table width="665" border="1"  cellspacing=0 cellpadding=0>
  <tr>
    <td width="98" height="25" nowrap><span class="style2">受處分人姓名</span></td>
    <td colspan="3"><span class="style2"><%=trim(rsSql("DRIVER"))%></span></td>
    <td width="126" nowrap><span class="style2">原舉發單位通知單</span></td>
    <td width="152"><span class="style2">第<%=trim(rsSql("BillNo"))%>號</span></td>
  </tr>
  <tr>
    <td height="25" nowrap><span class="style2">出生年月日</span></td>
    <td width="58"><span class="style2"><%=gInitDT(trim(rsSql("DriverBirth")))%>&nbsp;</span></td>
    <td width="47" height="25" nowrap><span class="style2">性別</span></td>
    <td width="84"><span class="style2">
      <%=UAgentSex%>
    &nbsp;</span></td>
    <td height="25" nowrap><span class="style2">身分證統一編號</span></td>
    <td><span class="style2"><%=trim(rsSql("DriverID"))%></span></td>
  </tr>
  <tr>
    <td height="25" nowrap><span class="style2">住址</span></td>
    <td colspan="3" nowrap><span class="style2"><%=trim(rsSql("DriverZip"))&trim(rsSql("DriverAddress"))%></span></td>
    <td height="25" nowrap><span class="style2">代保管物件</span></td>
    <td><span class="style2"><%=fastring%></span></td>
  </tr>
  <tr>
    <td height="25" nowrap><span class="style2">違規時間</span></td>
    <td colspan="3"><span class="style2">
      <%
		if trim(rsSql("IllegalDate"))<>"" then
			IllegalDate=split(gArrDT(rsSql("IllegalDate")),"-")
			response.write IllegalDate(0)&"年"&IllegalDate(1)&"月"&IllegalDate(2)&"日"&hour(rsSql("IllegalDate"))&"時"&minute(rsSql("IllegalDate"))&"分"
		end if%>
	</span></td>
    <td height="25" nowrap><span class="style2">違規地點</span></td>
    <td><span class="style2"><%=trim(rsSql("IllegalAddress"))%></span></td>
  </tr>
  <tr>
    <td height="25"><span class="style2">原舉發通知單<br>
    應到案日期</span></td>
    <td colspan="3"><span class="style2">
      <%
		if trim(rsSql("DealLineDate"))<>"" then
			DealLineDate=split(gArrDT(rsSql("DealLineDate")),"-")
			response.write DealLineDate(0)&"年"&DealLineDate(1)&"月"&DealLineDate(2)&"日前"
		end if%>
    &nbsp;</span></td>
    <td height="25" nowrap><span class="style2">舉發單位</span></td>
    <td> <span class="style2"><%=thenBillUnitName%> </span></td>
  </tr>
  <tr>
    <td height="70" nowrap><span class="style2">舉發違規事實</span></td>
    <td colspan="5"><span class="style2">
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
    <td height="38" nowrap><span class="style2">違反法條</span></td>
    <td colspan="5"><span class="style2">道路交通管理處罰條例<%
		Sys_Rule1=trim(rsSql("Rule1"))
		response.write "第"&left(trim(Sys_Rule1),2)&"條"
		if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
		response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款規定"
		%></span></td>
  </tr>
  <tr>
    <td height="64" nowrap><span class="style2">處罰主文</span></td>
    <td colspan="5"><span class="style2"><%=UPunishmentMainBody%></span></td>
  </tr>
  <tr>
    <td height="115" nowrap><span class="style2">簡要理由</span></td>
    <td colspan="5"><span class="style2"><%=USimpleReson%></span></td>
  </tr>
  <tr>
    <td height="36" nowrap><span class="style2">裁決日期</span></td>
    <td colspan="5"><span class="style2">中華民國<%=UJudeDate(0)%>年<%=UJudeDate(1)%>月<%=UJudeDate(2)%>日</span></td>
  </tr>
  <tr>
    <td height="37" nowrap><span class="style2">應到案處所</span></td>
    <td colspan="5"><span class="style2"><%=PasserCity&DutyUnitName%>　<br>地址：<%=DutyAddress%>　<br>電話：<%=theContactTel%></span></td>
  </tr>
  <tr>
    <td colspan="6" height="39" nowrap><span class="style2">機關首長&nbsp;</span></td>
	<td height="39" nowrap><span class="style2"><%=MemUnitName%>&nbsp;</span></td>

  </tr>
  <tr>
    <td height="158"><div align="center"><span class="style2">附<br>
    記</span></div></td>
	<!---------- smith 20091011 修改為20天  ---->
    <td colspan="5" valign="top"><span class="style2">
	一、受處分人不服本裁決者，應以原處分機關（<%=thenPasserCity&"<br>"%>
	&nbsp;&nbsp;&nbsp;&nbsp;<%=theUnitName%>）為被告，向原告住所地、居所地、所在地、違規行為地<br>
	&nbsp;&nbsp;&nbsp;&nbsp;、或原處分機關所在地之地方法院行政訴訟庭提起訴訟；其中撤銷訴訟之<br>
	&nbsp;&nbsp;&nbsp;&nbsp;提起，應於裁決書送達後30日之不變期間內為之。<br>

    二、請依處罰主文所定期限前持本裁決書至應到案處所、郵局繳納罰鍰。<br>

    三、上開罰鍰逾30日仍不繳納者，本分局將依『行政執行法』<br>

	&nbsp;&nbsp;&nbsp;&nbsp;第二章第十一條移送行政執行分署強制執行。<br>
	<%
	if sys_City="宜蘭縣" then
		Response.Write "四、罰鍰繳納方式：1.可親自或委託他人至本分局臨櫃窗口繳納。<br>"
		Response.Write "　　　　　　　　　2.購買郵政匯票（抬頭請註明本分局全銜，<br>"	
		Response.Write "　　　　　　　　　例：宜蘭縣政府警察局宜蘭分局）郵寄至本分局繳納。"
	else
		Response.Write "四、『郵局劃撥』戶名："&theBankName&"　<br>"
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;劃撥帳號："&theBankAccount
	end if
	%></span></td>
  </tr>
</table>
<br>
</body>
</html>