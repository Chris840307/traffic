<%
strSQL="select * from UnitInfo where UnitID='"&DB_UnitID&"'"
set rsUnit=conn.execute(strSQL)
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
rsUnit.close

strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&strBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

strSql="select BillTypeID,Driver,DriverID,DriverAddress,DriverZip,INSURANCE,ILLEGALADDRESS,RuleVer,IllegalSpeed,RuleSpeed,Note,BillFillDate,RECORDMEMBERID from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed=""
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
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
'response.write strSql
'response.end
set rsfound=conn.execute(strSql)
'if Not rsFound.eof then
'	if trim(rsfound("OwnerCounty"))<>"" then
'		Sys_OwnerZipName=trim(rsfound("OwnerCounty"))
'	else
'		Sys_DriverHomeZip=trim(rsfound("DriverCounty"))
'	end if
'end if
if Not rsFound.eof then Sys_Driver=trim(rsfound("Driver"))
if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("DriverHomeZip"))
strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName"))
rszip.close

if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close
Sys_DCIReturnStation=0
Sum_Level=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_BillNo=trim(rsfound("BillNo"))
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo"))
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1"))
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2"))
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=Cint(Sys_Level1)+Cint(Sys_Level2)
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

strSql="select a.LoginID,a.ChName,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_BillBaseRecordMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

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
SysUnit=Unit("UnitName")
SysUnitTel=Unit("Tel")
SysUnitAddress=Unit("Address")
Unit.close

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

if trim(Sys_BillTypeID)="1" then
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160
else
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160
end if

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
rsbil.close

%>
<table border="0">
	<tr>
		<td class="style2" rowspan=2 valign="top" align="right" height=200 width=5>
			<div id=idDiv style="position:absolute;left:0px;">
				<img src=<%="""..\BarCodeImage\"&Sys_BillNo&".jpg"""%>>
			</div>
			<br><br><br><br><br><br><br><br><br><br><br><br><br>
			<B><%=wordporss(chstr(SysUnit&"送達證書"))%></B>
		</td>
		<td class="style4" valign="bottom" align="right" height=200 width=50><%
		if Sys_BillTypeID="1" then
			response.write wordporss(chstr(Sys_Driver))
		else
			response.write wordporss(chstr(Sys_Owner))
		end if%>
		</td>
		<td class="style4" valign="bottom" align="right" height=150 width=15><%
		if Sys_BillTypeID="1" then
			strtmp=Sys_DriverHomeZip&" "&Sys_DriverZipName&Sys_DriverHomeAddress
		else
			strtmp=Sys_OwnerZip&" "&Sys_OwnerZipName&Sys_OwnerAddress
		end if
		if len(strtmp)>19 then
			response.write wordporss(chstr(mid(strtmp,1,18)))
		else
			response.write wordporss(chstr(strtmp))
		end if%>
		</td>
		<td class="style4" valign="bottom" align="right" height=150 width=15><%
			if len(strtmp)>19 then response.write wordporss(chstr(mid(strtmp,19,len(strtmp))))%>
		</td>
		<td class="style4" valign="bottom" align="right" height=200 width=20><%
			response.write wordporss(chstr(theBillNumber&"交字第"&Sys_BillNo&"號"))%>
		</td>
		<td class="style3" valign="bottom" align="right" height=200 width=20><%
			tmpstr="違反道路交通管理事件通知單"
			response.write wordporss(chstr(tmpstr))
			%>
		</td>
		<td class="style3" valign="bottom" align="right" height=200 width=10><%
			tmpstr="違反法條　"&Sys_Rule1
				response.write wordporss(chstr(tmpstr))
			%>
		</td>
		<td class="style3" valign="bottom" align="right" height=200 width=10><%
			if trim(Sys_Rule2)<>"0" then
				tmpstr="違反法條　"&Sys_Rule2
				response.write wordporss(chstr(tmpstr))
			end if
			%>
		</td>
		<td class="style4" valign="bottom" align="right" width="600" width=0><%=wordporss(chstr("請繳回:"&SysUnitAddress))%></td>
	</tr>
	<tr>
		<td height=140 align="right" valign="top" class="style4"><%=wordporss(chstr("　　　　　　　　　　　　　　"))%></td><td class="style4" height=100>&nbsp;</td><td class="style4" height=100>&nbsp;</td><td class="style4" height=100>&nbsp;</td>
	</tr>
</table>

