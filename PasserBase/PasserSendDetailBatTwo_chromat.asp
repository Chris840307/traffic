<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>違反道路交通管理事件移送</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<%
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
'strSQL="select * from UnitInfo where UnitLevelID=1 and UnitName like '%交%隊'"
set rsUnit=conn.Execute(strSQL)
DB_UnitID=trim(rsUnit("UnitID"))
theUnitName=trim(rsUnit("UnitName"))
theSubUnitSecBossName=trim(rsUnit("SecondManagerName"))
theBigUnitBossName=trim(rsUnit("ManageMemberName"))
theContactTel=trim(rsUnit("Tel"))
rsUnit.close

strSQL="select * from UnitInfo where UnitName like '%交%隊' and UnitLevelID=1"
set unit=conn.Execute(strSQL)
theBankAccount=trim(unit("BankAccount"))
unit.close

sys_cUit=""

If sys_City = "彰化縣" then 
	
	Set UitObj = Server.CreateObject("Scripting.Dictionary")
	
	UitObj.Add "JG01","" '交通隊
	UitObj.Add "JM00","204I02" '彰化分局
	UitObj.Add "JO00","204I03" '北斗分局
	UitObj.Add "JN00","204I04" '芳苑分局
	UitObj.Add "JP00","204I05" '和美分局
	UitObj.Add "JQ00","204I06" '鹿港分局
	UitObj.Add "JR00","204I07" '溪湖分局
	UitObj.Add "JS00","204I08" '員林分局
	UitObj.Add "JT00","204I09" '田中分局

	sys_cUit=UitObj.Item(Sys_UnitTypeID)
End if 

If Not ifnull(request("Sys_SendBillSN")) Then
	Sys_SendBillSN=request("Sys_SendBillSN")
elseif Not ifnull(request("hd_BillSN")) Then
	Sys_SendBillSN=request("hd_BillSN")
else
	Sys_SendBillSN=request("BillSN")
End if

if Not ifnull(request("DB_Add")) then
	'if trim(request("Sys_SubUnitSecBossName"))<>"" or trim(request("BigUnitBossName"))<>"" then
	'	strSQL="Update UnitInfo set SecondManagerName='"&trim(request("Sys_SubUnitSecBossName"))&"',ManageMemberName='"&trim(request("Sys_BigUnitBossName"))&"' where UnitID='"&DB_UnitID&"'"
	'	conn.execute(strSQL)
	'end if

	Sys_SendDate=gOutDT(request("Sys_SendDate"))
	Sys_ForFeit=split(trim(request("Sys_ForFeit")),",")
	PBillSN=split(trim(request("PBillSN")),",")
	PBillNo=split(trim(request("PBillNo")),",")
	OpenGovNumber=split(trim(request("Sys_OpenGovNumber")),",")
	SendNumber=split(trim(request("Sys_SendNumber")),",")
	Sys_AgentAddress=split(trim(request("Sys_AgentAddress")&","),",")

	Session("Sys_SendChName")=request("Sys_SendChName")
	For i=0 to UBound(PBillSN)
		tmCnt=0
		strSQL = "select Count(1) as cnt from PasserSendDetail where BillSN="&trim(PBillSN(i))&" and SendDate in(Select max(SendDate) from PasserSend where BillSN="&trim(PBillSN(i))&")"
		set rscnt=conn.execute(strSQL)
		tmCnt=cdbl(rscnt("cnt"))
		rscnt.close

		if ifnull(Sys_ForFeit(i))="" then
			temp_ForFeit=0
		else
			temp_ForFeit=trim(Sys_ForFeit(i))
		end if

		if tmCnt=0 then
			SQLOpenGovNumber="(select OpenGovNumber from PasserSend where BillSN="&trim(PBillSN(i))&")"
			SQLSendNumber="(select SendNumber from PasserSend where BillSN="&trim(PBillSN(i))&")"
			SQLSendDate="(select SendDate from PasserSend where BillSN="&trim(PBillSN(i))&")"

			strSQL="insert into PasserSendDetail values((select nvl(max(sn),0)+1 from PasserSendDetail),"&trim(PBillSN(i))&","&SQLOpenGovNumber&","&SQLSendNumber&","&SQLSendDate&",sysdate,"&Session("User_ID")&")"
			
			conn.execute(strSQL)
		end if

		
		
strSQL = "select Count(1) as cnt from PasserSendDetail where BillSN="&trim(PBillSN(i))&" and SendDate="&funGetDate(Sys_SendDate,0)

		set rscnt=conn.execute(strSQL)
		cnt=cdbl(rscnt("cnt"))
		rscnt.close

		MakeSureDate="":LimitDate=""

		strSQL="select MakeSureDate,LimitDate from PasserSend where BillSN="&trim(PBillSN(i))&" and BillNo='"&trim(PBillNo(i))&"'"

		set rsdat=conn.execute(strSQL)
		If not rsdat.eof Then
			MakeSureDate=trim(rsdat("MakeSureDate"))
			LimitDate=trim(rsdat("LimitDate"))
		End if
		rsdat.close

		If ifnull(MakeSureDate) Then
			passerCnt=0

			MakeSureDate="":LimitDate=""
			strSQL="select ArrivedDate,ReturnResonID from PasserSendArrived where ArriveType=0 and PasserSN="&PBillSN(i)
			set rsSendArr=conn.execute(strSQL)
			If Not rsSendArr.eof Then

				If rsSendArr("ReturnResonID") = "1" Then

					MakeSureDate=DateAdd("d",50,rsSendArr("ArrivedDate"))
					LimitDate=DateAdd("d",50,rsSendArr("ArrivedDate"))
				else

					MakeSureDate=DateAdd("d",30,rsSendArr("ArrivedDate"))
					LimitDate=DateAdd("d",30,rsSendArr("ArrivedDate"))
				End if 
			else

				strSQL="select JudeDate from PasserJude where Billsn="&PBillSN(i)
				set rsjude=conn.execute(strSQL)

				If not rsjude.eof Then
				
					MakeSureDate=DateAdd("d",30,rsjude("JudeDate"))
					LimitDate=DateAdd("d",30,rsjude("JudeDate"))

				end if
				rsjude.close

			end If 
			rsSendArr.close

			strSQL="select PasserSN,ArrivedDate from PasserSendArrived where passerSN = "& PBillSN(i) &" and  ArriveType=1"

			set rszip=conn.execute(strSQL)
			
			If Not rszip.eof Then 

				If not ifnull(rszip("ArrivedDate")) Then LimitDate=DateAdd("d",15,rszip("ArrivedDate"))
			end if
			rszip.close

		end if

		if cnt=0 then
			strSQL="insert into PasserSendDetail values((select nvl(max(sn),0)+1 from PasserSendDetail),"&trim(PBillSN(i))&",'"&OpenGovNumber(i)&"','"&SendNumber(i)&"',"&funGetDate(Sys_SendDate,0)&",sysdate,"&Session("User_ID")&")"

		else
			strSQL="update PasserSendDetail set OpenGovNumber='"&OpenGovNumber(i)&"',SendNumber='"&SendNumber(i)&"',SendDate="&funGetDate(gOutDT(request("Sys_SendDate")),0)&",RecordDate=sysdate,recordMemberID="&Session("User_ID")&" where sn=(select sn from PasserSendDetail where BillSN="&trim(PBillSN(i))&" and SendDate="&funGetDate(gOutDT(request("Sys_SendDate")),0)&") and BillSN="&trim(PBillSN(i))
		end if

		conn.execute(strSQL)

		strSQL="Update PasserSend set ForFeit="&temp_ForFeit&",SendNumber='"&SendNumber(i)&"',OpenGovNumber='"&OpenGovNumber(i)&"',SendDate="&funGetDate(Sys_SendDate,0)&",BigUnitBossName='"&request("Sys_BigUnitBossName")&"',SubUnitSecBossName='"&request("Sys_SubUnitSecBossName")&"',MakeSureDate="&funGetDate(MakeSureDate,0)&",LimitDate="&funGetDate(LimitDate,0)&",AttatchJude='"&request("Sys_AttatchJude")&"',AttatchUrge='"&request("Sys_AttatchUrge")&"',AttatchFortune='"&request("Sys_AttatchFortune")&"',AttatchGround='"&request("Sys_AttatchGround")&"',AttatchRegister='"&request("Sys_AttatchRegister")&"',AttatchFileList='"&request("Sys_AttatchFileList")&"',AttatchTable='"&request("Sys_AttatchTable")&"',ATTATPOSTAGE='"&request("Sys_ATTATPOSTAGE")&"',SAFETOEXIT='"&request("Sys_SAFETOEXIT")&"',SAFEACTION='"&request("Sys_SAFEACTION")&"',SAFEASSURE='"&request("Sys_SAFEASSURE")&"',SAFEDETAIN='"&request("Sys_SAFEDETAIN")&"',SAFESHUTSHOP='"&request("Sys_SAFESHUTSHOP")&"',RecordStateID=0,RecordDate="&funGetDate(now,1)&",RecordMemberID="&Session("User_ID")&" where BillSN="&trim(PBillSN(i))&" and BillNo='"&trim(PBillNo(i))&"'"

		conn.execute(strSQL)

		If not ifnull(Sys_AgentAddress(i)) Then
			strSQL="Update PasserSend set AgentAddress='"&trim(Sys_AgentAddress(i))&"' where BillSN="&trim(PBillSN(i))

			conn.execute(strSQL)

		end if

		strSQL="select DriverAddress,DriverZip from passerBase where SN="&trim(PBillSN(i))
		set rszda=conn.execute(strSQL)
		If ifnull(rszda("DriverZip")) Then
			If isnumeric(left(rszda("DriverAddress"),1)) Then
				strSQL="Update PasserBase set DriverZip="&left(rszda("DriverAddress"),3)&" where SN="&trim(PBillSN(i))

				conn.execute(strSQL)
			else
				strSQL="Update PasserBase set DriverZip="&getZip(rszda("DriverAddress"))&" where SN="&trim(PBillSN(i))

				conn.execute(strSQL)
			End if			
		End if
		rszda.close

	next
	response.write "<script language=""JavaScript"">"
	response.write "window.opener.funSendListTwo_chromat('"&trim(request("SendType"))&"');"
	response.write "</script>"
else
'strState="select * from PasserSend where BillSN="&trim(request("PBillSN"))
'set rsState=conn.execute(strState)
'BillEof=0
'if rsState.eof then BillEof=1
'rsState.close
%>
<body onkeydown="KeyDown()">
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33">違反道路交通管理事件移送</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99">產生移送案號</td>
					<td nowrap><input name="Sys_JudeSN1" type="text" class="btn1" size="10" maxlength="12" value="">
					-
					<input name="Sys_JudeSN2" type="text" class="btn1" size="2" maxlength="5" value="" onkeyup="value=value.replace(/[^\d]/g,'')">

					<input type="button" name="btnSelt" value="產生" onclick="funJudeSN();">
					</td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99">裁罰單位</td>
					<td nowrap><%=theUnitName%></td>
					<td align="right" nowrap bgcolor="#FFFF99">移送日期</td>
					<td nowrap>
						<input name="Sys_SendDate" value="<%
							'strSQL="Select max(SendDate) SendDate from PasserSend where BillSN in("&Sys_SendBillSN&")"
							'set rsda=conn.execute(strSQL)
							'if Not ifnull(trim(rsda("SendDate"))) then
							'	response.write gInitDT(trim(rsda("SendDate")))
							'else
								response.write gInitDT(date)
							'end if
							'rsda.close
						%>" class="btn1" type="text" size="4" maxlength="10" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_SendDate');">
					</td>
					<td align="right" nowrap bgcolor="#FFFF99">承辦人</td>
					<td nowrap><input name="Sys_SendChName" class="btn1" class="btn1" type="text" size="5" maxlength="12" value="<%=Session("Ch_Name")%>"></td>
					<td nowrap bgcolor="#FFFF99" align="right">分局長</td>
					<td><input name="Sys_SubUnitSecBossName" class="btn1" class="btn1" type="text" size="10" maxlength="12" value="<%=trim(theSubUnitSecBossName)%>">
					</td>
					<td nowrap bgcolor="#FFFF99" nowrap align="right">局長</td>
					<td><input name="Sys_BigUnitBossName" class="btn1" class="btn1" type="text" size="10" maxlength="12" value="<%=trim(theBigUnitBossName)%>"></td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99">催繳情形</td>
					<td nowrap>
						<input name="radiobutton" class="btn1" type="radio" value="1">
						業經催繳
						<input name="radiobutton" class="btn1" type="radio" value="0" checked>
						未經催繳
					</td>
					<td align="right" nowrap bgcolor="#FFFF99">催繳方式</td>
					<td colspan=3 nowrap>
						<input name="Sys_UrgeTypeID" class="btn1" type="radio" value="0">
						電話
						<input name="Sys_UrgeTypeID" class="btn1" type="radio" value="1">
						信函
						<input name="Sys_UrgeTypeID" class="btn1" type="radio" value="2">
						雙掛號或裁決書
					</td>
					<td></td>
					<td></td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99">附件</td>
					<td colspan="5">
						<input class="btn1" type="checkbox" name="Sys_AttatchTable" value="1" checked>
						附表<br>          
						<input class="btn1" type="checkbox" name="Sys_AttatchJude" value="1" checked>
						處分書裁決書或義務人依法令負有義務之證明文件及送達證明文件<br>
						<input class="btn1" type="checkbox" name="Sys_AttatchUrge" value="1" checked>
						義務人經限期履行而逾期仍不履行之證明文件及送達證明文件<br>
						<input class="btn1" type="checkbox" name="Sys_AttatchFortune" value="1" checked>
						義務人之財產目錄 
						<input class="btn1" type="checkbox" name="Sys_AttatchGround" value="1">
						土地登記簿謄本<br>
						<input class="btn1" type="checkbox" name="Sys_AttatchRegister" value="1" checked>
						義務人之戶籍資料
						<input class="btn1" type="checkbox" name="Sys_AttatchFileList" value="1">
						保全措施之資料<br>          
						<input class="btn1" type="checkbox" name="Sys_ATTATPOSTAGE" value="1">
						執行（債權）憑證
					</td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99">保全措施</td>
					<td colspan="5">
						<input class="btn1" type="checkbox" name="Sys_SAFETOEXIT" value="1">
						已限制出境 
						<input class="btn1" type="checkbox" name="Sys_SAFEACTION" value="1">
						已禁止處分
						<input class="btn1" type="checkbox" name="Sys_SAFEASSURE" value="1">
						已提供擔保 
						<input class="btn1" type="checkbox" name="Sys_SAFEDETAIN" value="1">
						已假扣押
						<input class="btn1" type="checkbox" name="Sys_SAFESHUTSHOP" value="1"> 
						已勒令停業
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#FFDD77">
			<input name="btnadd" type="button" value="移送書" onclick="funAdd('1');">
<!--			<input name="btnadd" type="button" value="移送書(套印)" onclick="funAdd('');"> -->
			<input name="btnexit" type="button" value=" 關 閉 " onclick="funExt();">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="JudePrint.jpg"><font size=5 color="blue">移送書套印格式說明</font></a>
		</td>
	</tr>
</table>
<hr>
<table width="100%" border="0">
<%
tmpSN=""
strSQL="select sn,DriverID from passerbase where sn in("&Sys_SendBillSN&") order by DriverID,sn"
set rsPasser=conn.execute(strSQL)
While not rsPasser.eof
	If not ifnull(tmpSN) Then
		tmpSN=tmpSN&","
	end if

	tmpSN=tmpSN&trim(rsPasser("sn"))
	
	rsPasser.movenext
Wend
rsPasser.close

BillSN=split(tmpSN,",")
For i=0 to Ubound(BillSN)
strSql="select a.SN as BillSN,a.BillNo,a.Driver,a.DriverBirth,a.DriverID,a.DriverAddress,a.IllegalDate,a.IllegalAddress,a.DealLineDate,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillUnitID,a.RuleVer,b.OpenGovNumber as JudeOGN,b.AgentName as JudeAgentName,b.AgentSex as JudeAgentSex,b.AgentBirth as JudeAgentBirth,b.AgentID as JudeAgentID,b.AgentAddress as JudeAgentAddress,c.OpenGovNumber as UrgeOGN,c.UrgeTypeID,d.OpenGovNumber,d.BigUnitBossName,d.SubUnitSecBossName,d.SendNumber,d.SendDate,d.Agent,d.AgentBirthDate,d.AgentID,d.AgentAddress,d.ForFeit,d.MakeSureDate,d.LimitDate,d.AttatchJude,d.AttatchUrge,d.AttatchFortune,d.AttatchGround,d.AttatchRegister,d.AttatchFileList,d.AttatchTable,d.ATTATPOSTAGE,d.SafeToExit,d.SAFEACTION,d.SAFEASSURE,d.SAFEDETAIN,d.SAFESHUTSHOP from PasserBase a,PasserJude b,PasserUrge c,PasserSend d where a.SN="&trim(BillSN(i))&" and a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and a.SN=c.BillSN(+) and a.BillNo=c.BillNo(+) and a.SN=d.BillSN(+) and a.BillNo=d.BillNo(+)"

set rsfound=conn.execute(strSql)

sendNo_tmp=""
If sys_City = "彰化縣" then 
	
	sendNo_tmp=sys_cUit&rsfound("BillNo")
else
	sendNo_tmp=rsfound("BillNo")
End if 
%>
	<tr>
		<td align="right" nowrap bgcolor="#FFFF99">舉發單號</td>
		<td><%=rsfound("BillNo")%></td>
		<td><input type="Hidden" name="PBillNo" value="<%=rsfound("BillNo")%>"></td>
		<td align="right" nowrap bgcolor="#FFFF99">發文文號</td>
		<td><input name="Sys_OpenGovNumber" class="btn1" type="text" size="8" maxlength="30" value="<%=rsfound("BillNo")%>"></td>
		<td align="right" nowrap bgcolor="#FFFF99">移送分署</td>
		<td><input name="Sys_AgentAddress" value="<%=trim(rsfound("AgentAddress"))%>" class="btn1" type="text" size="8"></td>
		<td align="right" nowrap bgcolor="#FFFF99">移送案號</td>
		<td><input name="Sys_SendNumber" value="<%
		If sys_City<>"宜蘭縣" Then
			strSQL="select (count(1)) cnt from PasserSendDetail where BillSN="&trim(BillSN(i))
			set rscnt=conn.execute(strSQL)

			response.write sendNo_tmp&"-"&rscnt("cnt")

			rscnt.close
		end if
		%>" class="btn1" type="text" size="8" maxlength="18"></td>
		<td align="right" nowrap bgcolor="#FFFF99">受處分人</td>
		<td nowrap><%=rsfound("Driver")%></td>
		<td align="right" nowrap bgcolor="#FFFF99">違反法條</td>
		<td>
			<%
			if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
				response.write trim(rsfound("Rule1"))&"，"
				strRule1="select * from Law where ItemID='"&trim(rsfound("Rule1"))&"' and VerSion='"&trim(rsfound("RuleVer"))&"'"
				set rsRule1=conn.execute(strRule1)
				if not rsRule1.eof then
					response.write cint(trim(rsRule1("Level1")))
					if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
						response.write "&nbsp; ,"&cint(trim(rsRule1("Level1")))
					else
						response.write "&nbsp; ,"&cint(trim(rsRule1("Level2")))
					end if
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level3")))
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level4")))
					response.write "<br>"&trim(rsRule1("IllegalRule"))
				end if
				rsRule1.close
				set rsRule1=nothing
			end if	
			if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
				response.write "<br>"&trim(rsfound("Rule2"))&"，"
				strRule1="select * from Law where ItemID='"&trim(rsfound("Rule2"))&"' and VerSion='"&trim(rsfound("RuleVer"))&"'"
				set rsRule1=conn.execute(strRule1)
				if not rsRule1.eof then
					response.write cint(trim(rsRule1("Level1")))
					if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
						response.write "&nbsp; ,"&cint(trim(rsRule1("Level1")))
					else
						response.write "&nbsp; ,"&cint(trim(rsRule1("Level2")))
					end if
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level3")))
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level4")))
					response.write "<br>"&trim(rsRule1("IllegalRule"))
				end if
				rsRule1.close
				set rsRule1=nothing
			end if	
			if trim(rsfound("Rule3"))<>"" and not isnull(rsfound("Rule3")) then
				response.write "<br>"&trim(rsfound("Rule3"))&"，"
				strRule1="select * from Law where ItemID='"&trim(rsfound("Rule3"))&"' and VerSion='"&trim(rsfound("RuleVer"))&"'"
				set rsRule1=conn.execute(strRule1)
				if not rsRule1.eof then
					response.write cint(trim(rsRule1("Level1")))
					if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
						response.write "&nbsp; ,"&cint(trim(rsRule1("Level1")))
					else
						response.write "&nbsp; ,"&cint(trim(rsRule1("Level2")))
					end if
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level3")))
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level4")))
					response.write "<br>"&trim(rsRule1("IllegalRule"))
				end if
				rsRule1.close
				set rsRule1=nothing
			end if	
			if trim(rsfound("Rule4"))<>"" and not isnull(rsfound("Rule4")) then
				response.write "<br>"&trim(rsfound("Rule4"))&"，"
				strRule1="select * from Law where ItemID='"&trim(rsfound("Rule4"))&"' and VerSion='"&trim(rsfound("RuleVer"))&"'"
				set rsRule1=conn.execute(strRule1)
				if not rsRule1.eof then
					response.write cint(trim(rsRule1("Level1")))
					if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
						response.write "&nbsp; ,"&cint(trim(rsRule1("Level1")))
					else
						response.write "&nbsp; ,"&cint(trim(rsRule1("Level2")))
					end if
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level3")))
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level4")))
					response.write "<br>"&trim(rsRule1("IllegalRule"))
				end if
				rsRule1.close
				set rsRule1=nothing
			end if
			%>
		</td>
		<td nowrap bgcolor="#FFFF99" align="right">罰款金額</td>
		<td>
		<%
			L1ForFeit=0
			L2ForFeit=0
			L3ForFeit=0
			L4ForFeit=0
			if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
				strRule1="select * from Law where ItemID='"&trim(rsfound("Rule1"))&"' and VerSion='"&trim(rsfound("RuleVer"))&"'"
				set rsRule1=conn.execute(strRule1)
				if not rsRule1.eof then
					L1ForFeit=cint(trim(rsRule1("Level1")))
					if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
						L2ForFeit=cint(trim(rsRule1("Level1")))
					else
						L2ForFeit=cint(trim(rsRule1("Level2")))
					end if
					L3ForFeit=cint(trim(rsRule1("Level3")))
					L4ForFeit=cint(trim(rsRule1("Level4")))
				end if
				rsRule1.close
				set rsRule1=nothing
			end if
			if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
				strRule2="select * from Law where ItemID='"&trim(rsfound("Rule2"))&"' and VerSion='"&trim(rsfound("RuleVer"))&"'"
				set rsRule2=conn.execute(strRule2)
				if not rsRule2.eof then
					L1ForFeit=L1ForFeit+cint(trim(rsRule2("Level1")))
					if trim(rsRule2("Level2")="" or isnull(rsRule2("Level2"))) then
						L2ForFeit=L2ForFeit+cint(trim(rsRule2("Level1")))
					else
						L2ForFeit=L2ForFeit+cint(trim(rsRule2("Level2")))
					end if
					L3ForFeit=L3ForFeit+cint(trim(rsRule2("Level3")))
					L4ForFeit=L4ForFeit+cint(trim(rsRule2("Level4")))
				end if
				rsRule2.close
				set rsRule2=nothing
			end if	
			if trim(rsfound("Rule3"))<>"" and not isnull(rsfound("Rule3")) then
				strRule3="select * from Law where ItemID='"&trim(rsfound("Rule3"))&"' and VerSion='"&trim(rsfound("RuleVer"))&"'"
				set rsRule3=conn.execute(strRule3)
				if not rsRule3.eof then
					L1ForFeit=L1ForFeit+cint(trim(rsRule3("Level1")))
					if trim(rsRule3("Level2")="" or isnull(rsRule3("Level2"))) then
						L2ForFeit=L2ForFeit+cint(trim(rsRule3("Level1")))
					else
						L2ForFeit=L2ForFeit+cint(trim(rsRule3("Level2")))
					end if
					L3ForFeit=L3ForFeit+cint(trim(rsRule3("Level3")))
					L4ForFeit=L4ForFeit+cint(trim(rsRule3("Level4")))
				end if
				rsRule3.close
				set rsRule3=nothing
			end if
			if trim(rsfound("Rule4"))<>"" and not isnull(rsfound("Rule4")) then
				strRule4="select * from Law where ItemID='"&trim(rsfound("Rule4"))&"' and VerSion='"&trim(rsfound("RuleVer"))&"'"
				set rsRule4=conn.execute(strRule4)
				if not rsRule4.eof then
					L1ForFeit=L1ForFeit+cint(trim(rsRule4("Level1")))
					if trim(rsRule4("Level2")="" or isnull(rsRule4("Level2"))) then
						L2ForFeit=L2ForFeit+cint(trim(rsRule4("Level1")))
					else
						L2ForFeit=L2ForFeit+cint(trim(rsRule4("Level2")))
					end if
					L3ForFeit=L3ForFeit+cint(trim(rsRule4("Level3")))
					L4ForFeit=L4ForFeit+cint(trim(rsRule4("Level4")))
				end if
				rsRule4.close
				set rsRule4=nothing
			end if	
				%>  
			  <input name="Sys_ForFeit" class="btn1" class="btn1" type="text" size="12" maxlength="12" value="<%
			  if trim(rsfound("ForFeit"))<>"" then
					response.write rsfound("ForFeit")
			  else
					response.write L4ForFeit
			  end if
			  %>" onkeyup="value=value.replace(/[^\d]/g,'')">
		</td>
	</tr>
<%
rsfound.close
next
%>
</table>
<input type="Hidden" name="DB_Add" value="">
<input type="Hidden" name="SendType" value="">
<input type="Hidden" name="PBillSN" value="<%=tmpSN%>">
<input type="hidden" name="BillEof" value="<%=BillEof%>">
<input type="Hidden" name="JudeCnt" value="<%=i%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funJudeSN(){
	var strSN="";
	var strCnt="";
	var strzeo="";

	if(myForm.Sys_JudeSN2.value!=''){
		strCnt=myForm.Sys_JudeSN2.value.length;
		for(i=0;i<=myForm.JudeCnt.value-1;i++){
			strzeo="";
			strSN=eval(eval(myForm.Sys_JudeSN2.value)+eval(i));
			for(j=strSN.toString().length;j<strCnt;j++){
				strzeo=strzeo.toString()+'0';
			}
			if(myForm.JudeCnt.value==1){
				myForm.Sys_SendNumber.value=myForm.Sys_JudeSN1.value+'-'+strzeo.toString()+strSN.toString();
			}else{
				myForm.Sys_SendNumber[i].value=myForm.Sys_JudeSN1.value+'-'+strzeo.toString()+strSN.toString();
			}
			
		}
	}
}
function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}
}

function funAdd(SendType){

	var chkSendNumber='';

	if(myForm.JudeCnt.value==1){
		chkSendNumber=myForm.Sys_SendNumber.value;
	}else{
		chkSendNumber=myForm.Sys_SendNumber[0].value;
	}

	if(chkSendNumber==''){
		alert("移送案號須填寫！");
	}else{
		myForm.SendType.value=SendType;
		myForm.DB_Add.value="Save";
		myForm.submit();
	}
}
function funExt() {
	if(confirm("是否關閉維護系統?")){
		opener.myForm.submit();
		self.close();
	}
}
</script>
<%
end if
conn.close%>