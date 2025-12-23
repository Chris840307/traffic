<%
strSQL="select * from UnitInfo where UnitID='"&DB_UnitID&"'"
set rsUnit=conn.execute(strSQL)
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
rsUnit.close

strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&strBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_Owner=""
Sys_OwnerZip=""
Sys_OwnerZipName=""
Sys_OwnerAddress=""
Sys_BillNo=""
Sys_CarNo=""
Sys_Rule1=""
Sys_Rule2=""

If Not rsbil.eof Then
strSql="select BillTypeID,Driver,DriverID,DriverAddress,DriverZip,INSURANCE,ILLEGALADDRESS,RuleVer,IllegalSpeed,RuleSpeed,Note,BillFillDate,RECORDMEMBERID from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed=""
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_DriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then Sys_BillBaseRecordMemberID=trim(rs("RECORDMEMBERID"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)
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

	If not ifnull(Sys_DriverID) and Not ifnull(trim(rsfound("DriverHomeAddress"))) Then
		if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver"))
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	end If 

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
			If Not Isnull(rsdata("DriverHomeAddress")) then
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
	'======

	else
		If Sys_BillTypeID=1 Then
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

		else
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))

		End if
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
			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

		else
			If Sys_BillTypeID=1 Then
				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

			else
				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))

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

Sys_OwnerAddress=replace(Sys_OwnerAddress,Sys_OwnerZipName,"")

Sys_DCIReturnStation=0
Sum_Level=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_BillNo=trim(rsfound("BillNo"))
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo"))
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1"))
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2"))
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=CDBL(Sys_Level1)+CDBL(Sys_Level2)
if Not rsfound.eof then Sys_DCIRETURNCARTYPE=trim(rsfound("DCIRETURNCARTYPE"))
strsql="select * from DCICODE where ID='"&Sys_DCIRETURNCARTYPE&"' and TypeID=5"
Sys_DCIRETURNCARTYPE=""
set cartype=conn.execute(strsql)
if Not cartype.eof then Sys_DCIRETURNCARTYPE=trim(cartype("Content"))
cartype.close

rsfound.close
Sys_Sex=""
strSql="select distinct BillFillerMemberID,DriverSex,DriverBirth,IllegalDate,DealLineDate,RECORDMEMBERID,IMAGEFILENAME,IMAGEPATHNAME,IMAGEFILENAMEB,BILLMEMID1 from BillBase where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"'"
set rssex=conn.execute(strSql)
if trim(Sys_BillTypeID)="1" then
	if Not rssex.eof then
		if trim(rssex("DriverSex"))="1" then
			Sys_Sex="男"
		else
			Sys_Sex="女"
		end if
	end if
end if

if Not rssex.eof then Sys_RecordMemberID=trim(rssex("RECORDMEMBERID"))

if Not rssex.eof then
	Sys_IllegalDate=split(gArrDT(trim(rssex("IllegalDate"))),"-")
else
	Sys_IllegalDate=split(gArrDT(trim("")),"-")
end if
if Not rssex.eof then
	Sys_IllegalDate_h=hour(trim(rssex("IllegalDate")))
else
	Sys_IllegalDate_h=""
end if
if Not rssex.eof then
	Sys_IllegalDate_m=minute(trim(rssex("IllegalDate")))
else
	Sys_IllegalDate_m=""
end if
if Not rssex.eof then
	Sys_DealLineDate=split(gArrDT(trim(rssex("DealLineDate"))),"-")
else
	Sys_DealLineDate=split(gArrDT(trim("")),"-")
end if
SendTime_DealLineDate=""
If SendTime>1 Then SendTime_DealLineDate=Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)

if Not rssex.eof then
	Sys_DriverBirth=split(gArrDT(trim(rssex("DriverBirth"))),"-")
else
	Sys_DriverBirth=split(gArrDT(trim("")),"-")
end if
if Not rssex.eof then Sys_IMAGEFILENAME=trim(rssex("IMAGEFILENAME"))
if Not rssex.eof then Sys_IMAGEFILENAMEB=trim(rssex("IMAGEFILENAMEB"))
if Not rssex.eof then Sys_IMAGEPATHNAME=trim(rssex("IMAGEPATHNAME"))
Sys_BillFillerMemberID=0
if Not rssex.eof then Sys_Billmem1ID=trim(rssex("BILLMEMID1"))

strSql="select a.LoginID,a.ChName,b.UnitName,b.UnitID,b.UnitTypeID,a.ImageFilename as MemberFileName,b.ImageFilename,b.UnitLevelID,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&trim(Sys_RecordMemberID)
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close
If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set unit=conn.Execute(strSQL)
If Not unit.eof Then sysunit=unit("UnitName")
unit.close

strSql="select a.LoginID,a.ChName,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_Billmem1ID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillUnitAddress=trim(mem("Address"))
'if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
'if Not mem.eof then Sys_ChName=trim(mem("ChName"))
'if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

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

strSQL="select UnitName,Tel,Address from UnitInfo where UnitID=(Select UnitID from MemberData where MemberID='"&trim(Sys_RecordMemberID)&"')"
set Unit=conn.execute(strSQL)
'SysUnit=Unit("UnitName")
SysUnitTel=Unit("Tel")
SysUnitAddress=Unit("Address")
Unit.close

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

strSql="select DCISTATIONID,DCISTATIONNAME,StationTel,StationID from Station where DCIStationID='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_DCISTATIONID=trim(rs("DCISTATIONID"))
if Not rs.eof then Sys_STATIONNAME=trim(rs("DCISTATIONNAME"))
if Not rs.eof then Sys_StationTel=trim(rs("StationTel"))
if Not rs.eof then Sys_StationID=trim(rs("StationID"))
rs.close

strSql="select MailNumber,StoreAndSendMailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))

rs.close
strSql="select b.Content from BILLFASTENERDETAIL a,DCICode b where a.FASTENERTYPEID=b.ID and b.TypeID=6 and a.BIllSN="&replace(trim(rsbil("BillSN")),"","0")&" and a.CarNo='"&trim(rsbil("CarNo"))&"'"
set rsfast=conn.execute(strsql)
fastring=""
while Not rsfast.eof
	if trim(fastring)<>"" then fastring=fastring&","
	fastring=fastring&rsfast("Content")
	rsfast.movenext
wend
rsfast.close
if isnull(Sys_DriverHomeZip) or trim(Sys_DriverHomeZip)="" then Sys_DriverHomeZip="001"
if isnull(Sys_OwnerZip) or trim(Sys_OwnerZip)="" then Sys_OwnerZip="001"

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
Sys_MailNumber=0
DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,60,160

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
end if
rsbil.close

%>
<table border="0">
	<tr>
		<td class="style2" rowspan=2 valign="top" align="right" height=320 width=0>
			<div id="idDiv" style="position:absolute; left:-8px;">
				<img src=<%="""../BarCodeImage/"&Sys_BillNo&".jpg"""%>>
			</div>
			<br><br><br><br><br><br><br><br><br><br>
			<B><%=wordporss(chstr(thenPasserCity&"送達證書"))%></B>
		</td>
		<td class="style4" valign="bottom" align="right" height=320 width=23><%
			response.write wordporss(funcCheckFont(Sys_Owner,20,4))%>
		</td>
		<td class="style4" valign="bottom" align="right" height=150 width=20><%
			strtmp=Sys_OwnerZip&" "&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,20,4)
		if len(strtmp)>19 and instr(strtmp,"img")=0 then
			response.write wordporss(chstr(mid(strtmp,1,18)))
		else
			response.write wordporss(chstr(strtmp))
		end if%>
		</td>
		<td class="style4" valign="bottom" align="right" height=150 width=15><%
			if len(strtmp)>19 then response.write wordporss(chstr(mid(strtmp,19,len(strtmp))))%>
		</td>
		<td class="style4" valign="bottom" align="right" height=320 width=20><%
			response.write wordporss(chstr(theBillNumber&"交字第"&Sys_BillNo&"號"))%>
		</td>
		<td class="style3" valign="bottom" align="right" height=320 width=25><%
			tmpstr="違反道路交通管理事件通知單"
			response.write wordporss(chstr(tmpstr))
			%>
		</td>
		<td class="style3" valign="bottom" align="right" height=320 width=10><%
			tmpstr="違反法條　"&Sys_Rule1
				response.write "<font size=1><b>"&wordporss(chstr("　"&SendTime_DealLineDate))&"</b>"&wordporss(chstr(tmpstr))&"</font>"
			%>
		</td>
		<td class="style3" valign="bottom" align="right" height=320 width=10><%
			if trim(Sys_Rule2)<>"0" then
				tmpstr="違反法條　"&Sys_Rule2
				response.write "<font size=2>"&wordporss(chstr(tmpstr))&"</font>"
			end if
			%>
		</td>
		<td class="style4" valign="bottom" align="right" width="600" width=0><%=wordporss(chstr("請繳回:"&SysUnitAddress))%></td>
	</tr>
	<tr>
		<td height=140 align="right" valign="top" class="style4"><%=wordporss(chstr("　　　　　　　　　　　　　　"))%></td><td class="style4" height=100>&nbsp;</td><td class="style4" height=100>&nbsp;</td><td class="style4" height=100>&nbsp;</td>
	</tr>
</table>

