<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 800
Response.flush
%>
<%
'權限
'AuthorityCheck(234)


	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
%>

<%
RecordDate=split(gInitDT(date),"-")
	strwhere=request("SQLstr")

ExchangeTypeFlag="W"
strExchangeType="select a.ExchangeTypeID from DciLog a,BillBase f where a.BillSN=f.SN "&_
	" and f.RecordStateID=0 "&strwhere
set rsEType=conn.execute(strExchangeType)
if not rsEType.eof then
	if trim(rsEType("ExchangeTypeID"))="N" then
		ExchangeTypeFlag="N"
	else
		ExchangeTypeFlag="W"
	end if
else
	ExchangeTypeFlag="W"
end if
rsEType.close
set rsEType=nothing

If  sys_City="台南市" Then
	userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	If trim(userip) = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 

	strI="insert into Log values((select nvl(max(Sn),0)+1 from Log),360,"&Trim(Session("User_ID"))&",'"&Trim(Session("Ch_Name"))&"','"&userip&"',sysdate,'大宗掛號清冊,"&Replace(strwhere,"'","""")&"')"
	'response.write strI
	Conn.execute strI
End If 

if sys_City="基隆市" then 
	if ExchangeTypeFlag="N" then
		strSQL="select a.BatchNumber,a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.UseTool,f.RecordDate,f.Note" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailhistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and e.ExchangeTypeID=d.DCIActionID(+) and e.Status=d.DCIReturn(+)" &_
		" and e.ExchangeTypeID='W'" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and ((d.DCIreturnStatus=1 and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L','h')))" &_
		" and f.sn=g.BillSn" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate" '二分局說要改
		'" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.StoreAndSendMailNumber,f.RecordDate"	 	
	else
		strSQL="select a.BatchNumber,a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.UseTool,f.RecordDate,f.Note" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.ExchangeTypeID='W'" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and ((((d.DCIreturnStatus=1 and ((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','L','T'))) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordDate"
	end if
elseif sys_City="南投縣" then 
	if ExchangeTypeFlag="N" then
		strSQL="select a.BatchNumber,a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.UseTool,f.RecordDate,f.Note" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and e.ExchangeTypeID=d.DCIActionID(+) and e.Status=d.DCIReturn(+)" &_
		" and e.ExchangeTypeID='W'" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and ((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8)) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','h','L')))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"		
	else
		strSQL="select a.BatchNumber,a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.UseTool,f.RecordDate,f.Note" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.ExchangeTypeID='W'" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and ((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordMemberID,f.RecordDate"
	end if
else
	if ExchangeTypeFlag="N" then
		strSQL="select a.BatchNumber,a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.UseTool,f.RecordDate,f.Note" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and e.ExchangeTypeID=d.DCIActionID(+) and e.Status=d.DCIReturn(+)" &_
		" and e.ExchangeTypeID='W'" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and ((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8)) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L')))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordDate"		
	else
		strSQL="select a.BatchNumber,a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.UseTool,f.RecordDate,f.Note" &_
		",f.driveraddress,f.driverzip,f.owner,f.ownerzip,f.owneraddress" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.ExchangeTypeID='W'" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and ((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordDate"
	end if
end if
	set rs1=conn.execute(strSQL)
if sys_City="基隆市" then 
	if ExchangeTypeFlag="N" then
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and e.ExchangeTypeID=d.DCIActionID(+) and e.Status=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and ((d.DCIreturnStatus=1 and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L','h')))" &_
		" and e.ExchangeTypeID='W' and a.RecordMemberID=b.MemberID(+) "&strwhere		
	else
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and ((((d.DCIreturnStatus=1 and ((e.DciErrorCarData not in ('1','3','9','a','j','A','H','K','L','T'))) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)))" &_
		" and a.ExchangeTypeID='W' and a.RecordMemberID=b.MemberID(+) "&strwhere
	end if
elseif sys_City="南投縣" then 
	if ExchangeTypeFlag="N" then
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and e.ExchangeTypeID=d.DCIActionID(+) and e.Status=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and ((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8)) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','h','L')))" &_
		" and e.ExchangeTypeID='W' and a.RecordMemberID=b.MemberID(+) "&strwhere
	else
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and ((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)))" &_
		" and a.ExchangeTypeID='W' and a.RecordMemberID=b.MemberID(+) "&strwhere
	end if
else
	if ExchangeTypeFlag="N" then
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and e.ExchangeTypeID=d.DCIActionID(+) and e.Status=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and ((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8)) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L')))" &_
		" and e.ExchangeTypeID='W' and a.RecordMemberID=b.MemberID(+) "&strwhere
	else
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and ((((d.DCIreturnStatus=1 and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)))" &_
		" and a.ExchangeTypeID='W' and a.RecordMemberID=b.MemberID(+) "&strwhere
	end if
end if
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		DBcnt=rsCnt("Cnt")
	end if
	rsCnt.close
	set rsCnt=nothing
'response.write strSQL
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.pageprint {
  margin-left: 7mm;
  margin-right: 5.08mm;
  margin-top: 5.08mm;
  margin-bottom: 5.08mm;
}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>大宗掛號清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
</head>
<body>
<form name=myForm method="post">
<%		mailSN=0
		PageNo=1
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
		if mailSN>0 then 
			response.write "<div class=""PageNext"">&nbsp;</div>"
			PageNo=PageNo+1
		End If 
%>
	<table width="710" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td	colspan="2" align="center">中華郵政</td>
		</tr>
<%if sys_City<>"雲林縣" then %>
		<tr>
			<td align="left" colspan="2">中華民國 <%=year(now)-1911%> 年 <%=month(now)%> 月 <%=day(now)%> 日</td>
		</tr>
<%end if%>
		<tr>
			<td align="left" width="35%">
				寄件人名稱：<%
			if sys_City="基隆市" then
				strSqlSend="select * from ApConfigure where ID=27"
				set rsS=conn.execute(strSqlSend)
				if not rsS.eof then
					response.write trim(rsS("Value"))
				end if
				rsS.close
				set rsS=nothing
			elseif sys_City="高雄縣" then
				strU="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
				set rsU=conn.execute(strU)
				if not rsU.eof then
					response.write "高雄縣政府警察局"&trim(rsU("UnitName"))
				end if
				rsU.close
				set rsU=nothing
			else
				strU="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
				set rsU=conn.execute(strU)
				if not rsU.eof then
					response.write trim(rsU("UnitName"))
				end if
				rsU.close
				set rsU=nothing
			end if
				%>
			</td>
			<td align="left" width="65%">交寄大宗 &nbsp; &nbsp; &nbsp; &nbsp; 掛 &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 號 
															&nbsp; &nbsp; &nbsp; &nbsp; 函件寄存  &nbsp; 
															<%if sys_City="基隆市" then %> 
															&nbsp;  &nbsp;  &nbsp;  &nbsp;  &nbsp; &nbsp; &nbsp; &nbsp;  &nbsp;  &nbsp; 批號 <%=trim(rs1("BatchNumber"))& " - " & PageNo%>
															 <%end if%>
															 
															 </td>
		</tr>
	</table>

	<table width="710" border="1" cellpadding="3" cellspacing="0">
		<tr>
			<td width="6%" height="26" align="center">號碼</td>
			<td width="15%" align="center">掛號碼</td>
			<td width="25%" align="center">收件人姓名</td>
			<td width="44%" align="center">寄達地名(或地址)</td>
			<td width="10%" align="center">備註</td>
		</tr>
<%
			for i=1 to 30
				if rs1.eof then exit for
				mailSN=mailSN+1
%>
		<tr>
			<td align="right"><%=mailSN%></td>	
			<td align="left"><%
			if trim(request("NoteMailNo"))="1" then
				if isnull(rs1("Note")) or trim(rs1("Note"))="" then
					response.write("&nbsp;")
				else
					NoteLen=InStr(trim(rs1("Note")),"大宗:")
					if NoteLen>0 then
						response.write Mid(trim(rs1("Note")),NoteLen+3,6)
					else
						response.write trim("&nbsp;")
					end if
				end if
			else
				strSqlH="select MailNumber,StoreAndSendMailNumber,opengovreportnumber,firstbarcode from BillMailHistory where BillSN="&trim(rs1("BillSN"))
				set rsH=conn.execute(strSqlH)
				if not rsH.eof then
					if sys_City="台中市" or sys_City="雲林縣" or sys_City="基隆市" then
						if trim(rs1("ExchangeTypeID"))="W" then
							if trim(rsH("MailNumber"))<>"" and not isnull(rsH("MailNumber")) then
								response.write right("00000000" & trim(rsH("MailNumber")),6)&"&nbsp;"
							else
								response.write "&nbsp;"
							end if
						elseif trim(rs1("ExchangeTypeID"))="N" then
							if trim(rsH("StoreAndSendMailNumber"))<>"" and not isnull(rsH("StoreAndSendMailNumber")) then
								response.write right("00000000" & trim(rsH("StoreAndSendMailNumber")),6)&"&nbsp;"
							else
								response.write "&nbsp;"
							end if
						end if
					else
						if trim(rs1("ExchangeTypeID"))="W" then
							if trim(rsH("MailNumber"))<>"" and not isnull(rsH("MailNumber")) then
								response.write trim(rsH("MailNumber"))&"&nbsp;"
							else
								response.write "&nbsp;"
							end if
						elseif trim(rs1("ExchangeTypeID"))="N" then
							if trim(rsH("StoreAndSendMailNumber"))<>"" and not isnull(rsH("StoreAndSendMailNumber")) then
								response.write trim(rsH("StoreAndSendMailNumber"))&"&nbsp;"
							else
								response.write "&nbsp;"
							end if
						end if
					end if
				else
					response.write "&nbsp;"
				end if
				rsH.close
				set rsH=nothing
			end if
			%></td>	
			<td align="left"><%
			GetMailMem=""
			GetMailAddress=""
			ZipName=""
		if trim(rs1("BillTypeID"))="2" then
			if sys_City="台東縣" then
				strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress,dcierrorcardata,Nwner,NwnerZip,NwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status='Y'"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
					if instr(trim(rsD("OwnerAddress")),"(住)")>0 or instr(trim(rsD("OwnerAddress")),"(就)")>0 or instr(trim(rsD("OwnerAddress")),"（住）")>0 or instr(trim(rsD("OwnerAddress")),"（就）")>0 then
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing
		
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress")),"臺","台"),ZipName,"")
					elseif trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing
		
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress")),"臺","台"),ZipName,"")
					else
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing
		
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress")),"臺","台"),ZipName,"")
					end if
				end if
				rsD.close
				set rsD=Nothing
			elseif sys_City="宜蘭縣" Then
				strSqlD="select SN,DriverZip,DriverAddress,Owner,OwnerZip,OwnerAddress from BIllBase where BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"' and Recordstateid=0"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
					if ExchangeTypeFlag="N" then	
						GetMailMem=trim(rsD("Owner"))

						If trim(rsD("DriverAddress") &"")<>"" Then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
			
							
							GetMailAddress=trim(rsD("DriverZip"))&ZipName&replace(replace(trim(rsD("DriverAddress") &""),"臺","台"),ZipName,"")
						Else	'TITAN沒寫BILLBASE DriverAddress的話,先抓A DriverHomeAddress,再抓W DriverHomeAddress
							strSqlDciA="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where Carno in (select carno from dcilog where BillSN="&trim(rsD("SN"))&" and ExchangetypeID='A') and ExchangeTypeID='A' and Status='S'"
							set rsDciA=conn.execute(strSqlDciA)
							if not rsDciA.eof then
								if trim(rsDciA("DriverHomeAddress"))<>"" and not isnull(rsDciA("DriverHomeAddress")) Then
									strZip="select ZipName from Zip where ZipID='"&trim(rsDciA("DriverHomeZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=nothing
									GetMailAddress=trim(rsDciA("DriverHomeZip"))&ZipName&replace(replace(trim(rsDciA("DriverHomeAddress") &""),"臺","台"),ZipName,"")

								end If
							else
								strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='W' "
								set rsD2=conn.execute(strSqlD2)
								if not rsD2.eof then
									if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) then
										strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
										set rsZip=conn.execute(strZip)
										if not rsZip.eof then
											ZipName=trim(rsZip("ZipName"))
										end if
										rsZip.close
										set rsZip=nothing

										GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&replace(replace(trim(rsD2("DriverHomeAddress"))&"","臺","台"),ZipName,"")

									end if
								end if
								rsD2.close
								set rsD2=nothing
							end if
							rsDciA.close
							set rsDciA=nothing
						End If 

					else
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing
		
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress") &""),"臺","台"),ZipName,"")
						'如果Billbase有寫以billbase為主
							If Not isnull(rs1("Owner")) Then
								GetMailMem=trim(rs1("Owner"))
							End If
							If Not isnull(rs1("OwnerAddress")) Then
								strZip="select ZipName from Zip where ZipID='"&trim(rs1("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=Nothing
								
								GetMailAddress=trim(rs1("OwnerZip"))&ZipName&replace(replace(trim(rs1("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If
					End if
				end if
				rsD.close
				set rsD=Nothing
			elseif sys_City="彰化縣" then
				if ExchangeTypeFlag="N" then
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof then
						if ExchangeTypeFlag="N" then	'單退先抓W看有沒有做戶籍補正，沒有的話再抓A,再沒有就抓owner
							if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress"))  then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
								if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
									GetMailMem="&nbsp;"
								else
									GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
								end if
								GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&trim(rsD2("DriverHomeAddress"))
							else
								strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
								set rsD=conn.execute(strSqlD)
								if not rsD.eof then
									if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
										if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
											GetMailMem="&nbsp;"
										else
											GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
										end if
										GetMailAddress=trim(rsD("DriverHomeZip"))&trim(rsD("DriverHomeAddress"))
									else
										if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
											GetMailMem="&nbsp;"
										else
											GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
										end if
										GetMailAddress="(車)"&trim(rsD("OwnerZip"))&trim(rsD("OwnerAddress"))
									end if
								else
									strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=nothing
					
									if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
										GetMailMem="&nbsp;"
									else
										GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
									end if
									GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
								end if
								rsD.close
								set rsD=nothing
							end if
						else
							'入案直接抓owner
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
								if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
									GetMailMem="&nbsp;"
								else
									GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
								end if
								GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
						end if
					end if
					rsD2.close
					set rsD2=nothing
				else
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','S','n','L')"
						set rsD2=conn.execute(strSqlD2)
						if not rsD2.eof then
							if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) and ExchangeTypeFlag="N" then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&trim(rsD2("DriverHomeAddress"))
							else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								if ExchangeTypeFlag="N" then
									GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
								else
									GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
								end if
							end if
						end if
						rsD2.close
						set rsD2=nothing
				end if
			elseif sys_City="花蓮縣" then
				if ExchangeTypeFlag="N" then	'單退先抓A的driver，沒有的話再抓W的Driver,再沒有就抓W的owner
					strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
					set rsD=conn.execute(strSqlD)
					if not rsD.eof then
						if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) and ExchangeTypeFlag="N" then
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress=trim(rsD("DriverHomeZip"))&trim(rsD("DriverHomeAddress"))
						else
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress=trim(rsD("OwnerZip"))&trim(rsD("OwnerAddress"))
						end if
					else
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
						set rsD2=conn.execute(strSqlD2)
						if not rsD2.eof then
							if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) and ExchangeTypeFlag="N" then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&replace(replace(trim(rsD2("DriverHomeAddress")),"臺","台"),ZipName,"")
							else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")),"臺","台"),ZipName,"")
							end if
						end if
						rsD2.close
						set rsD2=nothing
					end if
					rsD.close
					set rsD=nothing
				else	
						'XXXXX入案先抓A的OwnerNotifyAddress 2.W driver 3.W ownerXXXX
						'入案先抓A的OwnerNotifyAddress 3.A driver 3.W owner (2021/3/16)
					BitchHL=0
					strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress,OwnerNotifyAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
					set rsD=conn.execute(strSqlD)
					if not rsD.eof then
						if trim(rsD("OwnerNotifyAddress"))<>"" and not isnull(rsD("OwnerNotifyAddress")) then
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress=trim(rsD("OwnerNotifyAddress"))
						ElseIf trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
							If trim(rsD("DriverHomeZip"))<>"" then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=Nothing
							End if
							if isnull(rsD("DriverHomeZip")) or trim(rsD("Owner"))="" then
								GetMailMem=" &nbsp;"
							else
								GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
							end if
							GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress")),"臺","台"),ZipName,"")

						Else
							BitchHL=1
						end if
					Else
						BitchHL=1
					End If 
					rsD.close
					set rsD=Nothing
					
					If BitchHL=1 Then 
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress,Driver from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
						set rsD2=conn.execute(strSqlD2)
						if not rsD2.eof Then
							If trim(rsD2("OwnerAddress"))<>"" And Not isnull(rsD2("OwnerAddress")) Then
								
								If trim(rsD2("OwnerZip"))<>"" then
									strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=Nothing
								End if
								if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
									GetMailMem="&nbsp;"
								else
									GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
								end if
								GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")),"臺","台"),ZipName,"")
							End If 
						end if
						rsD2.close
						set rsD2=Nothing
					End if
				end If
			elseif sys_City="南投縣" Then
				if ExchangeTypeFlag="N" then
					strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
					set rsD=conn.execute(strSqlD)
					if not rsD.eof then
						if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) and ExchangeTypeFlag="N" then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress")),"臺","台"),ZipName,"")
						else
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress="(車)"&trim(rsD("OwnerZip"))&replace(replace(trim(rsD("OwnerAddress")),"臺","台"),ZipName,"")
						end if
					else
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
						set rsD2=conn.execute(strSqlD2)
						if not rsD2.eof then
							if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) and ExchangeTypeFlag="N" then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&replace(replace(trim(rsD2("DriverHomeAddress")),"臺","台"),ZipName,"")
							else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")),"臺","台"),ZipName,"")
							end if
						end if
						rsD2.close
						set rsD2=nothing
					end if
					rsD.close
					set rsD=nothing
				else
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
						set rsD2=conn.execute(strSqlD2)
						if not rsD2.eof then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&" ","臺","台"),ZipName,"")
						end if
						rsD2.close
						set rsD2=nothing
						'如果Billbase有寫以billbase為主
							If Not isnull(rs1("Owner")) Then
								GetMailMem=trim(rs1("Owner"))
							End If
							If Not isnull(rs1("OwnerAddress")) Then
								strZip="select ZipName from Zip where ZipID='"&trim(rs1("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=Nothing
								
								GetMailAddress=trim(rs1("OwnerZip"))&ZipName&replace(replace(trim(rs1("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If
				end If
			'-------------------------------------------------------------------------------------------
			ElseIf sys_City="基隆市" Then
				if ExchangeTypeFlag="N" then
					strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
					set rsD=conn.execute(strSqlD)
					if not rsD.eof then
						if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) and ExchangeTypeFlag="N" then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress"))&"","臺","台"),ZipName,"")
						else
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress="(車)"&trim(rsD("OwnerZip"))&replace(replace(trim(rsD("OwnerAddress"))&"","臺","台"),ZipName,"")
						end if
					else
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
						set rsD2=conn.execute(strSqlD2)
						if not rsD2.eof then
							if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) and ExchangeTypeFlag="N" then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&replace(replace(trim(rsD2("DriverHomeAddress"))&"","臺","台"),ZipName,"")
							else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&"","臺","台"),ZipName,"")
							end if
						end if
						rsD2.close
						set rsD2=nothing
					end if
					rsD.close
					set rsD=Nothing
	
					If sys_City="基隆市" Then '如果Billbase有寫以billbase為主
						If trim(rs1("BillTypeID"))="2" Then
							If Not isnull(rs1("Owner")) Then
								GetMailMem=trim(rs1("Owner"))
							End If
							If Not isnull(rs1("DriverAddress")) Then
								GetMailAddress=trim(rs1("DriverZip"))&" "&trim(rs1("DriverAddress"))
							End If
						End If 
					End If
				Else	'入案先抓住就地,再抓查車driver,再抓入案車籍地
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof Then
						GetMailMem=trim(rsD2("Owner"))
						if instr(trim(rsD2("OwnerAddress")),"(住)")>1 or instr(trim(rsD2("OwnerAddress")),"(就)")>1 or instr(trim(rsD2("OwnerAddress")),"（住）")>1 or instr(trim(rsD2("OwnerAddress")),"（就）")>1 Or instr(trim(rsD2("OwnerAddress")),"(通)")>1 or instr(trim(rsD2("OwnerAddress")),"（通）")>1 then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
			
							
							GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")),"臺","台"),ZipName,"")
						Else
							strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
							Set rsD3=conn.execute(strSqlD)
							If Not rsD3.eof Then
								If trim(rsD3("DriverHomeAddress"))<>"" And not isnull(rsD3("DriverHomeAddress")) then
									
									GetMailAddress=trim(rsD3("DriverHomeZip"))&replace(replace(trim(rsD3("DriverHomeAddress"))&"","臺","台"),ZipName,"")&"(戶)"
								Else
									strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=nothing
									
									GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&"","臺","台"),ZipName,"")
								End If
							Else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
								
								GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If
							rsD3.close
							Set rsD3=Nothing 
						End if
					end if
					rsD2.close
					set rsD2=Nothing
					If sys_City="基隆市" Then '如果Billbase有寫以billbase為主
						If trim(rs1("BillTypeID"))="2" Then
							If Not isnull(rs1("Owner")) Then
								GetMailMem=trim(rs1("Owner"))
							End If
							If Not isnull(rs1("OwnerAddress")) Then
								strZip="select ZipName from Zip where ZipID='"&trim(rs1("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipNameBill=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								GetMailAddress=trim(rs1("OwnerZip"))&ZipNameBill&replace(replace(trim(rs1("OwnerAddress"))&"","臺","台"),ZipNameBill,"")
							End If
						End If 
					End If
				end If
			'-------------------------------------------------------------------------------------------
			elseif sys_City="台中市" or sys_City="高雄市" or sys_City="高雄縣" then
				if ExchangeTypeFlag="N" then	'單退先抓A的driver，沒有的話再抓W的Driver,再沒有就抓W的owner
					strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
					set rsD=conn.execute(strSqlD)
					if not rsD.eof then
						if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) and ExchangeTypeFlag="N" then
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress=trim(rsD("DriverHomeZip"))&trim(rsD("DriverHomeAddress"))
						else
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress="(車)"&trim(rsD("OwnerZip"))&trim(rsD("OwnerAddress"))
						end if
					else
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
						set rsD2=conn.execute(strSqlD2)
						if not rsD2.eof then
							if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) and ExchangeTypeFlag="N" then
	'							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
	'							set rsZip=conn.execute(strZip)
	'							if not rsZip.eof then
	'								ZipName=trim(rsZip("ZipName"))
	'							end if
	'							rsZip.close
	'							set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&trim(rsD2("DriverHomeAddress"))
							else
	'							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
	'							set rsZip=conn.execute(strZip)
	'							if not rsZip.eof then
	'								ZipName=trim(rsZip("ZipName"))
	'							end if
	'							rsZip.close
	'							set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
							end if
						end if
						rsD2.close
						set rsD2=nothing
					end if
					rsD.close
					set rsD=nothing
				else	'入案直接抓W的Owner
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof then
					
	'						strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
	'						set rsZip=conn.execute(strZip)
	'						if not rsZip.eof then
	'							ZipName=trim(rsZip("ZipName"))
	'						end if
	'						rsZip.close
	'						set rsZip=nothing
						if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
							GetMailMem="&nbsp;"
						else
							GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
						end if
						GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
					end if
					rsD2.close
					set rsD2=nothing
				end If
			ElseIf sys_City="嘉義市" Or sys_City="澎湖縣" Then
				if ExchangeTypeFlag="N" then	'單退先抓A的driver，沒有的話再抓W的Driver,再沒有就抓W的owner
					strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
					set rsD=conn.execute(strSqlD)
					if not rsD.eof then
						if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) and ExchangeTypeFlag="N" then
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress="(戶)"&trim(rsD("DriverHomeZip"))&trim(rsD("DriverHomeAddress"))
						else
							GetMailMem=trim(rsD("Owner"))
							GetMailAddress="(車)"&trim(rsD("OwnerZip"))&trim(rsD("OwnerAddress"))
						end if
					else
						strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
						set rsD2=conn.execute(strSqlD2)
						if not rsD2.eof then
							if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) and ExchangeTypeFlag="N" then
	'							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
	'							set rsZip=conn.execute(strZip)
	'							if not rsZip.eof then
	'								ZipName=trim(rsZip("ZipName"))
	'							end if
	'							rsZip.close
	'							set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress="(戶)"&trim(rsD2("DriverHomeZip"))&ZipName&trim(rsD2("DriverHomeAddress"))
							else
	'							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
	'							set rsZip=conn.execute(strZip)
	'							if not rsZip.eof then
	'								ZipName=trim(rsZip("ZipName"))
	'							end if
	'							rsZip.close
	'							set rsZip=nothing

								GetMailMem=trim(rsD2("Owner"))
								GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
							end if
						end if
						rsD2.close
						set rsD2=nothing
					end if
					rsD.close
					set rsD=nothing
				else	'入案先抓住就地,再抓查車driver,再抓入案車籍地
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof Then
						GetMailMem=trim(rsD2("Owner"))
						if instr(trim(rsD2("OwnerAddress")),"(住)")>1 or instr(trim(rsD2("OwnerAddress")),"(就)")>1 or instr(trim(rsD2("OwnerAddress")),"（住）")>1 or instr(trim(rsD2("OwnerAddress")),"（就）")>1 Or instr(trim(rsD2("OwnerAddress")),"(通)")>1 or instr(trim(rsD2("OwnerAddress")),"（通）")>1  then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
			
							
							GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")),"臺","台"),ZipName,"")
						Else
							strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN in(select sn from billbase where billno='"&trim(rs1("BillNo"))&"' and recordstateid=0) and ExchangetypeID='A') and ExchangetypeID='A'"
							Set rsD3=conn.execute(strSqlD)
							If Not rsD3.eof Then
								If trim(rsD3("DriverHomeAddress"))<>"" And not isnull(rsD3("DriverHomeAddress")) then
									
									GetMailAddress=trim(rsD3("DriverHomeZip"))&replace(replace(trim(rsD3("DriverHomeAddress"))&"","臺","台"),ZipName,"")&"(戶)"
								Else
									strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=nothing
									
									GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&"","臺","台"),ZipName,"")
								End If
							Else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
								
								GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If
							rsD3.close
							Set rsD3=Nothing 
						End if
					end if
					rsD2.close
					set rsD2=Nothing
					
					'如果Billbase有寫以billbase為主
							If Not isnull(rs1("Owner")) Then
								GetMailMem=trim(rs1("Owner"))
							End If
							If Not isnull(rs1("OwnerAddress")) Then
								strZip="select ZipName from Zip where ZipID='"&trim(rs1("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=Nothing
								
								GetMailAddress=trim(rs1("OwnerZip"))&ZipName&replace(replace(trim(rs1("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If
				end If
			elseif sys_City="嘉義縣" or sys_City="屏東縣" Then
				if ExchangeTypeFlag="N" then
					strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
					set rsD=conn.execute(strSqlD)
					if not rsD.eof then
						ZipName=""

						if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
							GetMailMem="&nbsp;"
						else
							GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
						end if
						GetMailAddress=trim(rsD("OwnerZip"))&ZipName&trim(rsD("OwnerAddress"))
					end if
					rsD.close
					set rsD=Nothing
				Else	'入案先抓住就地,再抓查車driver,再抓入案車籍地
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof Then
						GetMailMem=trim(rsD2("Owner"))
						if instr(trim(rsD2("OwnerAddress")),"(住)")>1 or instr(trim(rsD2("OwnerAddress")),"(就)")>1 or instr(trim(rsD2("OwnerAddress")),"（住）")>1 or instr(trim(rsD2("OwnerAddress")),"（就）")>1 Or instr(trim(rsD2("OwnerAddress")),"(通)")>1 or instr(trim(rsD2("OwnerAddress")),"（通）")>1 then
							ZipName=""			
							
							GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress")),"臺","台"),ZipName,"")
						Else
							strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn " &_
							" where CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='A' and Status='S' " &_
							" and Carno in (select carno from dcilog where BillSN="&trim(rs1("BillSN")) &_
							" and ExchangetypeID='A' and dcireturnstatusid='S')"
							Set rsD3=conn.execute(strSqlD)
							If Not rsD3.eof Then
								If trim(rsD3("DriverHomeAddress"))<>"" And not isnull(rsD3("DriverHomeAddress")) then
									
									GetMailAddress=trim(rsD3("DriverHomeZip"))&replace(replace(trim(rsD3("DriverHomeAddress"))&"","臺","台"),ZipName,"")&"(戶)"
								Else
									ZipName=""
									
									GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&"","臺","台"),ZipName,"")
								End If
							Else
								ZipName=""
								
								GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&replace(replace(trim(rsD2("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If
							rsD3.close
							Set rsD3=Nothing 
						End if
					end if
					rsD2.close
					set rsD2=Nothing
					If sys_City="屏東縣" Then '如果Billbase有寫以billbase為主
						If trim(rs1("BillTypeID"))="2" Then
							If Not isnull(rs1("Owner")) Then
								GetMailMem=trim(rs1("Owner"))
							End If
							If Not isnull(rs1("OwnerAddress")) Then
								strZip="select ZipName from Zip where ZipID='"&trim(rs1("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=Nothing
								
								GetMailAddress=trim(rs1("OwnerZip"))&ZipName&replace(replace(trim(rs1("OwnerAddress"))&"","臺","台"),ZipName,"")
							End If
						End If 
					End If

				End If 	

			else
				strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','S','n','L')"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
					if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) and ExchangeTypeFlag="N" then
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("DriverHomeZip"))&" "&ZipName&trim(rsD("DriverHomeAddress"))
					else
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("OwnerZip"))&" "&ZipName&trim(rsD("OwnerAddress"))
					end if
				end if
				rsD.close
				set rsD=nothing
			end if
			if ExchangeTypeFlag="W" then	
				If sys_City="高雄市" Or sys_City="保二總隊三大隊一中隊" Or sys_City="彰化縣" Then '如果Billbase有寫以billbase為主
					If Not isnull(rs1("Owner")) Then
						GetMailMem=trim(rs1("Owner"))
					End If
					If Not isnull(rs1("OwnerAddress")) Then
						strZip="select ZipName from Zip where ZipID='"&trim(rs1("OwnerZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							rs1ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing

						GetMailAddress=trim(rs1("OwnerZip"))&" "&rs1ZipName&trim(rs1("OwnerAddress"))
					End If
				End If
			end if 
		else	'攔停
			if sys_City="高雄縣" then
				strSqlD="select Driver,DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress,Rule1 from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
					RuleTarget=""
					strRule="select Target from Law where ItemID='"&trim(rsD("Rule1"))&"'"
					set rsRule=conn.execute(strRule)
					if not rsRule.eof then
						RuleTarget=trim(rsRule("Target"))
					end if
					rsRule.close
					set rsRule=nothing
					if RuleTarget="V" then
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("OwnerZip"))&ZipName&trim(rsD("OwnerAddress"))
					else
						if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
							if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣" or sys_City="嘉義縣" or sys_City="台南市" then
								ZipName=""
							else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
							end if
								GetMailMem=trim(rsD("Driver"))
								GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&trim(rsD("DriverHomeAddress"))
						end if
					end if
				end if
				rsD.close
				set rsD=nothing
			else
				strSqlD="select Driver,DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','S','n','L')"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
					if sys_City="台中市" or sys_City="南投縣" or sys_City="彰化縣" then
						if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								if isnull(rsD("Driver")) or trim(rsD("Driver"))="" then
									GetMailMem="&nbsp;"
								else
									GetMailMem=trim(replace(rsD("Driver")," "," &nbsp;"))
								end if
								GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&trim(rsD("DriverHomeAddress"))
						else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing

								if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
									GetMailMem="&nbsp;"
								else
									GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
								end if
		
								GetMailAddress="(車)"&trim(rsD("OwnerZip"))&ZipName&trim(rsD("OwnerAddress"))
						end if
					else
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing
						GetMailMem=trim(rsD("Driver"))
						GetMailAddress=trim(rsD("DriverHomeZip"))&" "&ZipName&trim(rsD("DriverHomeAddress"))
					end if
				end if
				rsD.close
				set rsD=nothing
			end if
		end if
			response.write funcCheckFont(GetMailMem,14,1)
			%></td>	
			<td align="left"><%=funcCheckFont(GetMailAddress,14,1)%></td>	
			<td align="left"><%=rs1("BillNO")%></td>	
		</tr>
<%			
			rs1.MoveNext
			next
%>
	</table>
<%		
		Wend
		rs1.close
		set rs1=nothing
%>			
	<table width="710" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td height="26" width="60%" align="right"><%="上開  掛號函件  共  "&mailSN&"  件照收無誤"%>
			</td>
			<td width="40%" align="right" >____________________
			</td>
		</tr>
		<tr>
			<td colspan="2" align="right">經辦員簽署
			</td>
		</tr>
	</table>
</form>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}
window.print();

</script>
<%
conn.close
set conn=nothing
%>