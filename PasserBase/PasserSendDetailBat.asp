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


strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

showCreditor=false

if sys_City="台中市" or sys_City = "彰化縣" or sys_City = "台南市" or sys_City = "高雄市" or sys_City = "高雄縣" or sys_City="宜蘭縣" or sys_City="基隆市" or sys_City="澎湖縣" or sys_City="屏東縣" then
	showCreditor=true
end If

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
end If 

set rsUnit=conn.Execute(strSQL)
DB_UnitID=trim(rsUnit("UnitID"))
DB_UnitName=trim(rsUnit("UnitName"))
theSubUnitSecBossName=trim(rsUnit("SecondManagerName"))
rsUnit.close

strSQL="select * from UnitInfo where UnitName like '%交%隊' and UnitLevelID=1"
set unit=conn.Execute(strSQL)
theBankAccount=trim(unit("BankAccount"))
theContactTel=trim(unit("Tel"))
theBigUnitBossName=trim(unit("ManageMemberName"))
unit.close

sys_cUit=""

If sys_City = "彰化縣" then 
	
	Set UitObj = Server.CreateObject("Scripting.Dictionary")
	
	UitObj.Add "JG01","" '交通隊
	UitObj.Add "JM00","204I02" '彰化分局
	UitObj.Add "JQ00","204I03" '鹿港分局
	UitObj.Add "JP00","204I04" '和美分局
	UitObj.Add "JO00","204I05" '北斗分局
	UitObj.Add "JS00","204I06" '員林分局
	UitObj.Add "JR00","204I07" '溪湖分局
	UitObj.Add "JT00","204I08" '田中分局
	UitObj.Add "JN00","204I09" '芳苑分局

	sys_cUit=UitObj.Item(Sys_UnitTypeID)
End if 

If Not ifnull(request("Sys_SendBillSN")) Then

	sys_billsn01=request("Sys_SendBillSN")
elseif Not ifnull(request("hd_BillSN")) Then

	sys_billsn01=request("hd_BillSN")
else

	sys_billsn01=request("BillSN")
End If 

tmp_billsn=split(sys_billsn01,",")

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

if Not ifnull(request("DB_Add")) then
	If not ifnull(request("Sys_SendDate")) Then Sys_SendDate=gOutDT(request("Sys_SendDate"))	
	Sys_ForFeit=split(trim(request("Sys_ForFeit")),",")
	PBillSN=split(trim(request("PBillSN")),",")
	PBillNo=split(trim(request("PBillNo")),",")
	SendNumber=split(trim(request("Sys_SendNumber")),",")
	Sys_OpenGovNumber=split(trim(request("Sys_OpenGovNumber")),",")
	Session("Sys_SendChName")=request("Sys_SendChName")

	For i=0 to UBound(PBillSN)
		
		if ifnull(Sys_ForFeit(i))="" then
			temp_ForFeit=0
		else
			temp_ForFeit=trim(Sys_ForFeit(i))
		end if

		strSQL="Update UnitInfo set WordNum='"&trim(request("Sys_WordNum"))&"' where UnitID='"&Session("Unit_ID")&"'"
		conn.execute(strSQL)

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
		
		strSQL = "select Count(*) as cnt from PasserSend where BillSN="&trim(PBillSN(i))&" and BillNo='"&trim(PBillNo(i))&"'"
		set rscnt=conn.execute(strSQL)

		If not ifnull(request("bat_OpenGovNumber")) Then Sys_OpenGovNumber(i)=trim(request("bat_OpenGovNumber"))

		if Not Cint(rscnt("cnt"))>0 then
			strSQL="insert into PasserSend(BillSN,BillNo,OpenGovNumber,SendNumber,SendDate,ForFeit,BigUnitBossName,SubUnitSecBossName,MakeSureDate,LimitDate,AttatchJude,AttatchUrge,AttatchFortune,AttatchGround,AttatchRegister,AttatchFileList,AttatchTable,ATTATPOSTAGE,SAFETOEXIT,SAFEACTION,SAFEASSURE,SAFEDETAIN,SAFESHUTSHOP,RecordStateID,RecordDate,RecordMemberID) values("&trim(PBillSN(i))&",'"&trim(PBillNo(i))&"','"&trim(Sys_OpenGovNumber(i))&"','"&SendNumber(i)&"',"&funGetDate(Sys_SendDate,0)&","&temp_ForFeit&",'"&request("Sys_BigUnitBossName")&"','"&request("Sys_SubUnitSecBossName")&"',"&funGetDate(MakeSureDate,0)&","&funGetDate(LimitDate,0)&",'"&request("Sys_AttatchJude")&"','"&request("Sys_AttatchUrge")&"','"&request("Sys_AttatchFortune")&"','"&request("Sys_AttatchGround")&"','"&request("Sys_AttatchRegister")&"','"&request("Sys_AttatchFileList")&"','"&request("Sys_AttatchTable")&"','"&request("Sys_ATTATPOSTAGE")&"','"&request("Sys_SAFETOEXIT")&"','"&request("Sys_SAFEACTION")&"','"&request("Sys_SAFEASSURE")&"','"&request("Sys_SAFEDETAIN")&"','"&request("Sys_SAFESHUTSHOP")&"',0,"&funGetDate(now,1)&","&Session("User_ID")&")"
		
			conn.execute(strSQL)

			
			If not ifnull(request("Sys_AgentAddress")) Then
				strSQL="Update PasserSend set AgentAddress='"&trim(request("Sys_AgentAddress"))&"' where BillSN="&trim(PBillSN(i))&" and AgentAddress is null"

				conn.execute(strSQL)

			end If 

		end If 

		rscnt.close


		If showCreditor then
		
			strSQL="select count(1) cnt from PasserSendDetail where BillSN="&trim(PBillSN(i))

			set rscnt=conn.execute(strSQL)

			If cdbl(rscnt("cnt"))=0 Then
				strSQL="select OpenGovNumber,SendNumber,SendDate from PasserSend where billsn="&trim(PBillSN(i))
				set rssend=conn.execute(strSQL)

				If not rssend.eof Then

					strSQL="insert into PasserSendDetail values((select nvl(max(sn),0)+1 from PasserSendDetail),"&trim(PBillSN(i))&",'"&trim(rssend("OpenGovNumber"))&"','"&trim(rssend("SendNumber"))&"',"&funGetDate(rssend("SendDate"),0)&",sysdate,"&Session("User_ID")&")"

					conn.execute(strSQL)
				End if
				rssend.close
			End if
			rscnt.close

		end if


		strSQL="select DriverAddress,DriverZip from passerBase where SN="&trim(PBillSN(i))
		set rszda=conn.execute(strSQL)
		If ifnull(rszda("DriverZip")) Then
			If isnumeric(left(rszda("DriverAddress"),1)) Then
				strSQL="Update PasserBase set DriverZip="&left(rszda("DriverAddress"),3)&" where SN="&trim(PBillSN(i))

				conn.execute(strSQL)
			elseif not ifnull(getZip(rszda("DriverAddress"))) then
				strSQL="Update PasserBase set DriverZip="&getZip(rszda("DriverAddress"))&" where SN="&trim(PBillSN(i))

				conn.execute(strSQL)
			End if			
		End if
		rszda.close

		'strSQL="Update PasserBase set ForFeit1="&trim(Sys_ForFeit(i))&" where SN="&trim(PBillSN(i))&" and BillNo='"&trim(PBillNo(i))&"'"
		'conn.execute(strSQL)

	next
	response.write "<script language=""JavaScript"">"
	response.write "window.opener.funSendList();"
	response.write "</script>"
	Response.End
end if

'strState="select * from PasserSend where BillSN="&trim(request("PBillSN"))
'set rsState=conn.execute(strState)
'BillEof=0
'if rsState.eof then BillEof=1
'rsState.close

'if trim(request("Sys_BankAccount"))<>"" or trim(request("Sys_SubUnitSecBossName"))<>"" or trim(request("BigUnitBossName"))<>"" then
'	strSQL="Update UnitInfo set SecondManagerName='"&trim(request("Sys_SubUnitSecBossName"))&"',ManageMemberName='"&trim(request("Sys_BigUnitBossName"))&"' where UnitID='"&Session("Unit_ID")&"'"
'	conn.execute(strSQL)
'end if

SysWordNum=""
strSQL="select WordNum from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rs=conn.execute(strSQL)
If Not rs.eof Then SysWordNum=trim(rs("WordNum"))
rs.close

sys_illegaldate=""

strSQL="Select max(illegaldate) illegaldate from PasserBase where Exists(select 'Y' from "&BasSQL&" where SN=PasserBase.SN)"
set rsload=conn.execute(strSQL)
If not rsload.eof Then
	If not ifnull(rsload("illegaldate")) Then sys_illegaldate=gInitDT(rsload("illegaldate"))
End if 
rsload.close

strSQL="Select max(ArrivedDate) ArrivedDate from PassersEndArrived where Exists(select 'Y' from "&BasSQL&" where SN=PassersEndArrived.PasserSN)"
set rsload=conn.execute(strSQL)
If not rsload.eof Then
	If not ifnull(rsload("ArrivedDate")) Then sys_illegaldate=gInitDT(rsload("ArrivedDate"))
End if 
rsload.close

%>
<body onkeydown="KeyDown()">
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#1BF5FF">違反道路交通管理事件移送</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99">產生移送案號</td>
					<td nowrap><input name="Sys_JudeSN1" type="text" class="btn1" size="10" maxlength="30" value="">
					-
					<input name="Sys_JudeSN2" type="text" class="btn1" size="2" maxlength="20" value="" onkeyup="value=value.replace(/[^\d]/g,'')">

					<input type="button" name="btnSelt" value="產生" onclick="funJudeSN();">
					</td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99">舉發單位</td>
					<td nowrap><%=DB_UnitName%></td>
					<td align="right" nowrap bgcolor="#FFFF99"><font color="Red"><B>交字號</B></font></td>
					<td><input name="Sys_WordNum" type="text" class="btn1" size="10" maxlength="15" value="<%=SysWordNum%>">交字第</td>
					<td align="right" nowrap bgcolor="#FFFF99">整比發文文號</td>
					<td><input name="bat_OpenGovNumber" class="btn1" type="text" size="8" maxlength="16" value=""></td>
					<td align="right" nowrap bgcolor="#FFFF99">移送日期</td>
					<td nowrap>
						<input name="Sys_SendDate" value="<%=gInitDT(date)%>" class="btn1" type="text" size="4" maxlength="10" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_SendDate');">
					</td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99">行政執行處</td>
					<td><input name="Sys_AgentAddress" class="btn1" class="btn1" type="text" size="10" maxlength="12" value="">
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
						<input name="radiobutton" class="btn1" type="radio" value="1" checked>
						業經催繳
						<input name="radiobutton" class="btn1" type="radio" value="0">
						未經催繳
					</td>
					<td align="right" nowrap bgcolor="#FFFF99">催繳方式</td>
					<td colspan=3 nowrap>
						<input name="Sys_UrgeTypeID" class="btn1" type="radio" value="0">
						電話
						<input name="Sys_UrgeTypeID" class="btn1" type="radio" value="1">
						信函
						<input name="Sys_UrgeTypeID" class="btn1" type="radio" value="2">
						雙掛號、裁決書或員警送達
					</td>
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
		<td bgcolor="#1BF5FF">
			<input name="btnadd" type="button" value=" 確 定 " onclick="funAdd();"> 
			<input name="btnexit" type="button" value=" 關 閉 " onclick="funExt();">
		</td>
	</tr>
</table>
<hr>
<table width="100%" border="0">
<%


For i=0 to Ubound(tmp_billsn)
strSql="select a.SN as BillSN,a.BillNo,a.Driver," &_
		"a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.RuleVer," &_
		"nvl(forfeit1,0) forfeit1," &_
		"nvl(forfeit2,0) forfeit2," &_
		"(Select ForFeit from PasserSend where billsn=a.sn) ForFeit_S" &_
		" from PasserBase a where a.SN="&trim(tmp_billsn(i))

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
		<td><input name="Sys_OpenGovNumber" class="btn1" type="text" size="8" maxlength="16" value="<%=rsfound("BillNo")%>"></td>
		<td align="right" nowrap bgcolor="#FFFF99">移送案號</td>
		<td><input name="Sys_SendNumber" value="<%=sendNo_tmp%>" class="btn1" type="text" size="8" maxlength="30"></td>
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
			  if trim(rsfound("ForFeit_S"))<>"" then
					response.write rsfound("ForFeit_S")

			  elseif cdbl(rsfound("forfeit2"))>0 then
					response.write cdbl(rsfound("forfeit1"))+cdbl(rsfound("forfeit2"))

			  else
					response.write cdbl(rsfound("forfeit1"))
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
<input type="Hidden" name="PBillSN" value="<%=sys_billsn01%>">
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
	}

	for(i=0;i<=myForm.JudeCnt.value-1;i++){
		strzeo="";
		if(strCnt!=''){
			strSN=eval(eval(myForm.Sys_JudeSN2.value)+eval(i+1));
			for(j=strSN.toString().length;j<strCnt;j++){
				strzeo=strzeo.toString()+'0';
			}
		}
		myForm.Sys_SendNumber[i].value=myForm.Sys_JudeSN1.value+strzeo.toString()+strSN.toString();
	}
	
}
function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}
}

function funAdd(){
	var err='';
	var sys_illegaldate='<%=sys_illegaldate%>';

	if(eval(myForm.Sys_SendDate.value)<=eval(sys_illegaldate)){
		err="移送日必須大於送達日!!\n";
	}


	if(err!=''){
		alert(err);
	}else{
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
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}
function funPrintDetail(){
	UrlStr="PictureDetail.htm";
	newWin(UrlStr,"inputWin",1000,800,50,10,"yes","yes","yes","no");
}
</script>
<%conn.close%>