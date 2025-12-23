<style type="text/css">
<!--
td {font-family:新細明體;line-height:9px;font-size:9pt;}-->
</style>
<%
strSql="select LoginID from MemberData where MemberID="&session("User_ID")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_LoginID=trim(rs("LoginID"))
rs.close
set rs=nothing

strBil="select distinct BillSN,BillNo,CarNo,DCIerrorCarData from DCILOG where BillSN="&strBillSN(gyi)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_DCIerrorCarData=""
if Not rsbil.eof then Sys_DCIerrorCarData=trim(rsbil("DCIerrorCarData"))
Sys_Owner=""
Sys_OwnerZip=""
Sys_OwnerZipName=""
Sys_OwnerAddress=""
Sys_BillNo=""
Sys_CarNo=""
Sys_Rule1=""
Sys_Rule2=""
Sys_DriverID=""
Sys_BillBaseDriverID=""

If Not rsbil.eof Then
strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
if Not rs.eof then Sys_OwnerAddress=trim(rs("OwnerAddress"))
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_BillUnitID=trim(rs("BillUnitID"))
if Not rs.eof then Sys_BillBaseDriverID=trim(rs("DriverID"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close
set rs=nothing

if sys_City="基隆市" then
	strSql="select StoreAndSendFinalMailDate from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
	set rs=conn.execute(strSql)
	if Not rs.eof then Sys_StoreAndSendFinalMailDate=split(gArrDT(trim(rs("StoreAndSendFinalMailDate")&"")),"-")
	rs.close
   set rs=nothing
end If 

If ifnull(Sys_OwnerAddress) Then
	strSQL="select OwnerNotIfyAddress from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"
	
	set rsfi=conn.execute(strSql)

	if Not rsfi.eof then
		If Not ifnull(trim(rsfi("OwnerNotIfyAddress"))) Then

			notify_Zip="":notify_Addr=""

			if IsNumeric(left(trim(rsfi("OwnerNotIfyAddress")),3)) then

				notify_Zip=left(trim(rsfi("OwnerNotIfyAddress")),3)
			end If 

			notify_Addr=replace(trim(rsfi("OwnerNotIfyAddress")),notify_Zip,"")

			If instr(replace(trim(rsfi("OwnerNotIfyAddress")),"（","("),"(")<=0 then

				notify_Addr=notify_Addr&"(通)"
			end If 
			
			strSQL="update billbasedcireturn set OwnerZip='"&notify_Zip&"',OwnerAddress='"&notify_Addr&"' where exchangetypeid='W' and BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"'"

			conn.execute(strSQL)
		end if
	end If 
	rsfi.close
end If 

Sys_OwnerAddress="" 

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)
Sys_OwnerAddress=""

if sys_City="台東縣" then
	chkExchange=""
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver"))
	else
		if Not rsFound.eof then Sys_Owner=trim(rsfound("Owner"))
	End if
	if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
	if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
	if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

	If ifnull(Sys_OwnerAddress) and trim(Sys_DCIerrorCarData)<>"D" then
		chkExchange="A"
		strSql="select * from BillbaseDCIReturn where CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A'"
		set rsdata=conn.execute(strsql)
		If Sys_BillTypeID=1 Then
			if Not rsdata.eof then Sys_Owner=trim(rsdata("Driver"))
		else
			if Not rsdata.eof then Sys_Owner=trim(rsdata("Owner"))
		End if

		if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))
		if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
		if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))
		rsdata.close
	end if

	If ifnull(Sys_OwnerAddress) or len(Sys_OwnerAddress)<10 Then
		if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	end if
else
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver")&"")
		If Trim(Sys_Owner&"")="" Then If Not rsFound.eof Then Sys_Owner=trim(rsfound("Owner"))
	else
		if Not rsFound.eof then Sys_Owner=trim(rsfound("Owner")&"")
	End if
	if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))

	if Instr(request("Sys_BatchNumber"),"N")>0 Then

		Sys_OwnerAddress=""
		If sys_City="南投縣" Then
			if Sys_BillTypeID=2 Then
				strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='N'"
			else
				strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
			end if

			set rsdata=conn.execute(strsql)

			If Sys_BillTypeID=1 Then
				if Not rsdata.eof then Sys_Owner=trim(rsdata("Driver"))

				If ifnull(Sys_Owner) Then Sys_Owner=trim(rsdata("Owner"))
			else
				if Not rsdata.eof then Sys_Owner=trim(rsdata("Owner"))
			End if

			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

			rsdata.close
		End if

		If ifnull(Sys_OwnerAddress) Then			
			strSql="select * from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A') and ExchangetypeID='A'"
			set rsdata=conn.execute(strsql)
'			If Sys_BillTypeID=1 Then
'				if Not rsdata.eof then Sys_Owner=trim(rsdata("Driver")&"")
'				If Trim(Sys_Owner&"")="" Then If Not rsdata.eof Then Sys_Owner=trim(rsdata("Owner"))
'			else
'				if Not rsdata.eof then Sys_Owner=trim(rsdata("Owner")&"")
'			End if

			If Not Isnull(rsfound("DriverHomeAddress")) then
				
				Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
				Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

			elseIf Not Isnull(rsdata("DriverHomeAddress")) then

				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

			Else
				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))
			End If

			rsdata.close
		end if

		If ifnull(Sys_OwnerAddress) Then
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
			If ifnull(Sys_OwnerAddress) Then
				if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
				if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
			end If
		End If

	end if

	If ifnull(Sys_OwnerAddress) Then
		strSql="select * from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A') and ExchangetypeID='A'"
		set rsdata=conn.execute(strsql)
		If Sys_BillTypeID=1 Then
			if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver")&"")
			If Trim(Sys_Owner&"")="" Then If Not rsFound.eof Then Sys_Owner=trim(rsfound("Owner"))
		else
			if Not rsFound.eof then Sys_Owner=trim(rsfound("Owner")&"")
		End if

		if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))

		if Instr(request("Sys_BatchNumber"),"N")>0 then 

			If Not Isnull(rsfound("DriverHomeAddress")) then
				
				Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
				Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

			else

				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

			end if

		else
			If Sys_BillTypeID=1 or (Sys_BillTypeID=2 and Sys_BillBaseDriverID<>"") Then
				if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
				if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
				if Not rsfound.eof then Sys_Owner=trim(rsfound("Driver")&"")

			else
				If instr(replace(rsfound("OwnerAddress"),"（","("),"(住")>0 or instr(replace(rsfound("OwnerAddress"),"（","("),"(就")>0 or instr(replace(rsfound("OwnerAddress"),"（","("),"(通")>0 Then

					if Not rsfound.eof then Sys_OwnerAddress=rsfound("OwnerAddress")
					if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))

				elseIf sys_City="保二總隊三大隊一中隊" Then

					if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
					if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))
					

					If ifnull(Sys_OwnerAddress) Then

						if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
						if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
					End If 

				End if

				If ifnull(Sys_OwnerAddress) Then
					if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
					if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))
				End if
			End if
		end if
		rsdata.close
	end if

	If ifnull(Sys_OwnerAddress) Then
		if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner")&"")
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))

	end if
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName") & "") 
rszip.close
set rszip=nothing
If trim(chkExchange)="A" Then Sys_OwnerZipName=""

Sys_OwnerAddress=trim(replace(Sys_OwnerAddress&" ","臺","台"))

'else
'	if Not rsFound.eof then Sys_Driver=trim(rsfound("Driver") & "") 
'	if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID") & "") 
'	if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("DriverHomeAddress") & "") 
'	if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("DriverHomeZip") & "") 
'	strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
'	set rszip=conn.execute(strSQL)
'	if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName") & "") 
'	rszip.close
'	set rszip=nothing
'end if

'if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner") & "") 
'if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress") & "") 
'if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip") & "") 
'strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
'set rszip=conn.execute(strSQL)
'if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName") & "") 
'rszip.close
'set rszip=nothing

Sys_DCIReturnStation=0
Sum_Level=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation") & "") 
if Not rsfound.eof then Sys_BillNo=trim(rsfound("BillNo") & "") 
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo") & "") 
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1") & "") 
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2") & "") 
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1") & "") 
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2") & "") 
Sum_Level=Cdbl(Sys_Level1)+Cdbl(Sys_Level2)

rsfound.close
set rsfound=nothing
'Sys_Sex=""
strSql="select distinct BillFillerMemberID,DealLineDate,RECORDMEMBERID,BILLMEMID1 from BillBase where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"'"
set rssex=conn.execute(strSql)

if Not rssex.eof then Sys_RecordMemberID=trim(rssex("RECORDMEMBERID"))

if Not rssex.eof then
	Sys_DealLineDate=split(gArrDT(trim(rssex("DealLineDate"))),"-")
else
	Sys_DealLineDate=split(gArrDT(trim("")),"-")
end if

Sys_BillFillerMemberID=0
if Not rssex.eof then Sys_Billmem1ID=trim(rssex("BILLMEMID1"))

strSql="select a.LoginID,b.UnitID,b.UnitTypeID,b.UnitLevelID from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&trim(Sys_RecordMemberID)
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
mem.close
set mem=Nothing

strSql="select UnitName from UnitInfo where UnitID='"&trim(Sys_BillUnitID)&"'"
set mem=conn.execute(strsql)
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
mem.close
set mem=Nothing


'If Sys_UnitLevelID=1 Then
'	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
'else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
'end if

set unit=conn.Execute(strSQL)
If Not unit.eof Then sysunit=replace(unit("UnitName"),"保二三大一中","")
unit.close
set unit=nothing
strSql="select a.LoginID,a.ChName,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_Billmem1ID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillUnitAddress=trim(mem("Address"))
mem.close
set mem=nothing

if trim(Sys_Rule1)<>"0" and not isnull(Sys_Rule1) then
	strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
end if
rssex.close
set rssex=nothing

strSQL="select UnitName,Tel,Address from UnitInfo where UnitID=(Select UnitID from MemberData where MemberID='"&trim(Sys_RecordMemberID)&"')"
set Unit=conn.execute(strSQL)
'SysUnit=Unit("UnitName")
SysUnitTel=Unit("Tel")
SysUnitAddress=Unit("Address")
Unit.close
set Unit=nothing

If Sys_City="花蓮縣" Then

	If trim(Sys_BillUnitName) = "拖吊保管場" and Instr(request("Sys_BatchNumber"),"N")>0 Then

		strSQL="select UnitName,Tel,Address from UnitInfo where UnitID='A000'"
		set Unit=conn.execute(strSQL)

		'SysUnit=Unit("UnitName")
		SysUnitTel=Unit("Tel")
		SysUnitAddress=Unit("Address")

		Unit.close
		set Unit=nothing

	End if 
end if

Sys_IllegalRule2=""
if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then
	strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
end if

strSql="select DCISTATIONNAME,StationID from Station where DCIStationID='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_STATIONNAME=trim(rs("DCISTATIONNAME"))
if Not rs.eof then Sys_StationID=trim(rs("StationID"))
rs.close
set rs=nothing

strSql="select MailNumber,MailDate,UserMarkMemberID from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_UserMarkMemberID=trim(rs("UserMarkMemberID"))
rs.close
set rs=nothing

If sys_City="南投縣" Then
	If not ifnull(Sys_UserMarkMemberID) Then
		strSQL="select loginid from memberdata where memberid="&Sys_UserMarkMemberID
		set rsmem=conn.execute(strSQL)
		If not rsmem.eof Then Sys_BillFillerMemberID=trim(rsmem("loginid"))
		rsmem.close
	End if
	
end if

if isnull(Sys_DriverHomeZip) or trim(Sys_DriverHomeZip)="" then Sys_DriverHomeZip="001"
if isnull(Sys_OwnerZip) or trim(Sys_OwnerZip)="" then Sys_OwnerZip="001"

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
Sys_MailNumber=0
Sys_BillNo_BarCode=Sys_BillNo

If sys_City="彰化縣" or sys_City="高雄市" or sys_City="金門縣" or sys_City="苗栗縣" or sys_City="保二總隊三大隊一中隊" or sys_City="保二總隊三大隊二中隊" Then
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160,1

elseIf sys_City<>"台中縣" Then
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160

else

	Sys_BillNo_BarCode=Sys_BillNo_BarCode&"_4"
end if
if trim(Sys_DriverHomeZip)="001" then Sys_DriverHomeZip=""
if trim(Sys_OwnerZip)="001" then Sys_OwnerZip=""
end if


strSql="select MailNumber,StoreAndSendMailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))

rs.close
rsbil.close
set rsbil=nothing

tempdd=""
If sys_City="南投縣" Then 
	tempdd=right("00"&gInitDT(now),7) 
elseIf sys_City="台南市" or sys_City="台南縣" Then 
	 tempdd=right("00"&gInitDT(now),7)
elseIf sys_City="嘉義縣" Then
	tempdd=right("00"&gInitDT(Sys_MailDate),7)
elseif sys_City="基隆市" then
	 tempdd=right("00"&Sys_StoreAndSendFinalMailDate(0),3)&Sys_StoreAndSendFinalMailDate(1)&Sys_StoreAndSendFinalMailDate(2)
ElseIf Sys_City="台中市" Then 
	tempdd=Sys_UnitName
else
	tempdd=right("00"&sys_Date(0),3)&sys_Date(1)&sys_Date(2)
end If 

Data2=""
If Sys_City="台中市" Then 
	If chkStore=0 Then
		Data2=Sys_MailNumber
	else
		Data2=Sys_StoreAndSendMailNumber
	End If 
	
elseIf sys_City="嘉義縣" Then
	Data2="&nbsp;&nbsp;&nbsp;&nbsp;"

else
	Data2=cdbl(gyi+1)
End if

%>
<table width="95%" border="0" cellspacing="0">
	<tr>
		<th align="center" class="style5" rowspan="2" width="70%">
			<%=thenPasserCity&replace(sysunit,trim(thenPasserCity),"")&"送達證書" %>
		</th>
		<th align="left" class="style4">
			<%
				If sys_City="南投縣" Then 
					Response.Write "系統日期："

				elseIf sys_City="台南市" or sys_City="台南縣" Then
					Response.Write "郵寄日期："

				elseIf sys_City="嘉義縣" Then
					Response.Write "郵寄日期："

				elseif sys_City="基隆市" then
					Response.Write "郵寄日期："

				ElseIf Sys_City="台中市" Then 
					Response.Write "舉發單位："

				else
					Response.Write "填單日期："
				end If 
				
				Response.Write tempdd
			%>
		</th>
	</tr>
	<tr>
		<th align="left" class="style4">
			<%
				If sys_City="南投縣" Then 
					Response.Write "序&nbsp;&nbsp;&nbsp;&nbsp;號："

				elseIf sys_City="台南市" or sys_City="台南縣" Then
					Response.Write "序&nbsp;&nbsp;&nbsp;&nbsp;號："

				elseIf sys_City="嘉義縣" Then
					Response.Write "序&nbsp;&nbsp;&nbsp;&nbsp;號："

				elseif sys_City="基隆市" then
					Response.Write "序&nbsp;&nbsp;&nbsp;&nbsp;號："

				ElseIf Sys_City="台中市" Then 
					Response.Write "掛號號碼："

				else
					Response.Write "序&nbsp;&nbsp;&nbsp;&nbsp;號："

				end If 
				
				Response.Write Data2
			%>
		</th>
	</tr>
</table>

<table width="95%" bordercolor="#000000" border="1" cellspacing="0" class="tablestyle">
	<tr>
		<td width="40%" align="center" colspan="2" class="style6">
			受送達人名稱姓名地址
		</td>
		<td width="60%" colspan="2" class="style9">
			<table border="0" cellpadding="0" cellspacing="0">
				<tr><td class="style9" style="width:40%;"><%=funcCheckFont(Sys_Owner,14,1)%></td>
				<td><img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg"></td></tr>
			</table>
			<%=funcCheckFont(Sys_OwnerZip&" "&replace(Sys_OwnerZipName&Sys_OwnerAddress,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName),14,1)%>
		</td>
	</tr>
	<tr>
		<td width="40%" align="center" colspan="2" class="style6">
			文&nbsp;&nbsp;&nbsp;&nbsp;號
		</td>
		<td width="60%" height="25" colspan="2" class="style9">
			<%
				Response.Write "第"&Sys_BillNo&"號&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				If sys_City="嘉義市" then
					response.write Sys_CarNo
				else
					response.write left(Sys_CarNo,4)&left("*************",len(Sys_CarNo)-4)
				end if
			%>
		</td>
	</tr>
	<tr>
		<td width="40%" align="center" colspan="2" class="style6">
			送達文書（含案由）
		</td>
		<td width="60%" colspan="2" class="style9">
			<%
				Response.Write "舉發違反道路交通管理事件通知單<br>"
				Response.Write "違反法條"&Sys_Rule1
				If trim(Sys_Rule2) <> "0" Then Response.Write "   "&Sys_Rule2
			%>
		</td>
	</tr>
	<tr>
		<td width="20%" align="center" rowspan="2" class="style10">
			原郵局日戳
		</td>
		<td width="20%" align="center" rowspan="2" class="style10">
			送達郵局日戳
		</td>
		<td width="35%" class="style10">
			送達處所（由送達人填記）
		</td>
		<td width="25%" align="center" rowspan="2" class="style10">
			送達人簽章
		</td>
	</tr>
	<tr>
		<td height="40" align="left" class="style10">
			□&nbsp;同上記載地址<br>
			□&nbsp;改送：
		</td>
	</tr>
	<tr>
		<td width="20%" height="70" align="center" rowspan="2" class="style10">
			&nbsp;
		</td>
		<td width="20%" align="center" rowspan="2" class="style10">
			&nbsp;
		</td>
		<td width="35%" class="style10">
			送達時間（由送達人填記）
		</td>
		<td width="25%" align="center" rowspan="2" class="style10">
			&nbsp;
		</td>
	</tr>
	<tr>
		<td align="left" class="style10">
			中華民國
			&nbsp;&nbsp;&nbsp;&nbsp;年
			&nbsp;&nbsp;&nbsp;&nbsp;月
			&nbsp;&nbsp;&nbsp;&nbsp;日<br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;午
			&nbsp;&nbsp;&nbsp;&nbsp;時
			&nbsp;&nbsp;&nbsp;&nbsp;分
		</td>
	</tr>
	<tr>
		<td colspan="4" height="20" align="center" class="style10">
			送
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;達
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;方
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;式
		</td>
	</tr>
	<tr>
		<td colspan="4" height="20" align="center" class="style10">
			由&nbsp;&nbsp;&nbsp;&nbsp;
			送&nbsp;&nbsp;
			達&nbsp;&nbsp;
			人&nbsp;&nbsp;
			在&nbsp;&nbsp;
			□&nbsp;&nbsp;
			上&nbsp;&nbsp;
			劃&nbsp;&nbsp;
			V&nbsp;&nbsp;
			選&nbsp;&nbsp;
			記
		</td>
	</tr>
	<tr>
		<td colspan="2" height="20" align="left" class="style10">
			□&nbsp;已將文書交與應受送達人
		</td>
		<td colspan="2" align="left" class="style10">
			□&nbsp;本人
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			（簽名或蓋章）
		</td>
	</tr>
	<tr>
		<td colspan="2" height="20" align="left" class="style10">
			□&nbsp;未獲會晤本人，已將文書交與有辨別<br>
			&nbsp;&nbsp;&nbsp;事理能力之同居人、受雇人或應送達<br>
			&nbsp;&nbsp;&nbsp;處所之接收郵件人員<br>
		</td>
		<td colspan="2" align="left" class="style10">
			□&nbsp;同居人<br>
			□&nbsp;受雇人
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;（簽名或蓋章）<br>
			□&nbsp;應送達處所接收郵件人員<br>
		</td>
	</tr>
	<tr>
		<td colspan="2" height="20" align="left" class="style10">
			□&nbsp;應受送達之本人、同居人或受雇人收<br>
			&nbsp;&nbsp;&nbsp;領應，拒絕或不能簽名或蓋章者，由<br>
			&nbsp;&nbsp;&nbsp;送達人記明其事由<br>
		</td>
		<td colspan="2" align="left" class="style10">
			送達人填記：
		</td>
	</tr>
	<tr>
		<td colspan="2" height="20" align="left" class="style10">
			□&nbsp;應受送達之本人、同居人、受雇人或<br>
			&nbsp;&nbsp;&nbsp;應受送達處所接收郵件人員無正當理<br>
			&nbsp;&nbsp;&nbsp;由拒絕收領經送達人將文書留置於送<br>
			&nbsp;&nbsp;&nbsp;達處所，以為送達<br>
		</td>
		<td colspan="2" align="left" class="style10">
			□&nbsp;本人處<br>
			□&nbsp;同居人<br>
			□&nbsp;受雇人<br>
			□&nbsp;應送達處所接收郵件人員<br>
		</td>
	</tr>
	<tr>
		<td colspan="2" height="20" align="left" class="style10">
			□&nbsp;未獲會晤本人亦無受領文書之同居<br>
			&nbsp;&nbsp;&nbsp;人、受雇人或應受送達處所接收郵<br>
			&nbsp;&nbsp;&nbsp;件人員，已將該送達文書：<br>
			□&nbsp;應受送達之本人、同居人、受雇人或<br>
			&nbsp;&nbsp;&nbsp;應受送達處所接收郵件人員無正當理<br>
			&nbsp;&nbsp;&nbsp;由拒絕收領，並有難達留置情事，已<br>
			&nbsp;&nbsp;&nbsp;將該送達文書：<br>
		</td>
		<td align="left" class="style10">
			□&nbsp;寄存於&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;派出所<br>
			□&nbsp;寄存於&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;鄉（鎮、市、區）<br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;公所<br>
			□&nbsp;寄存於&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;鄉（鎮、市、區）
　　　　　　　　　　&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;公所
　　　　　　　　　　&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;村（里）辦公處<br>
			□&nbsp;寄存於&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;郵局<br>
		</td>
		<td align="left" class="style10">
			並作送達通知書二份，一<br>
			份黏貼於應受送達人住居所、事務所、營業所或其就業處所門首，一份□&nbsp;交由鄰居轉交或□&nbsp;置於該受送達處所信箱或其他適當位置，以為送達。
		</td>
	</tr>
	<tr>
		<td colspan="2" height="20" align="center" class="style10">
			送&nbsp;&nbsp;達&nbsp;&nbsp;人&nbsp;&nbsp;注&nbsp;&nbsp;意&nbsp;&nbsp;事&nbsp;&nbsp;項
		</td>
		<td colspan="2" align="left" class="style10">
			一、依上述送達方法送達者，送達人應即將本送達證書，提出<br>
			&nbsp;&nbsp;&nbsp;&nbsp;於交送達之行政機關附卷。<br>
			二、不能依上述送達方法送達者，送達人應製作記載該事由之<br>
			&nbsp;&nbsp;&nbsp;&nbsp;報告書，提出於交送達之行政機關附卷，並繳回應送達之<br>
			&nbsp;&nbsp;&nbsp;&nbsp;文書。

		</td>
	</tr>
</table>
<br>
<table width="100%" border="0" cellspacing="0">
	<tr>
		<td width="30"></td>
		<th align="left" class="style5">
			<%=thenPasserCity&replace(sysunit,trim(thenPasserCity),"")%>
		</th>
	</tr>
	<tr>
		<td></td>
		<td align="left" class="style6">
			<%
				If Sys_City<>"台中市" Then Response.Write "操作人員："&Sys_BillFillerMemberID&"<br>"
				Response.Write "應到案處所："&Sys_STATIONNAME
				If Sys_City="雲林縣" Then response.write "&nbsp;&nbsp;&nbsp;&nbsp;"&Sys_BillNo
			%>
		</td>
	</tr>
	<tr>
		<td></td>
		<td align="left" class="style6">
			本證書送回地址：
			<%
			If Sys_City="台中市" Then 
				response.write "臺中市西屯區大隆路１９２號"
			Else
				response.write SysUnitAddress
			End if
			%>
		</td>
	</tr>
</table>
