<%

strBil="select distinct BillSN,BillNo,CarNo,BillTypeID from DCILOG where BillSN="&strBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_Owner="":Sys_Driver=""
Sys_OwnerZipName=""
Sys_BillNo=""
Sys_CarNo=""
Sys_Rule1=""
Sys_Rule2=""

Sys_OwnerAddress="":Sys_OwnerZip=""
Sys_DriverHomeAddress="":Sys_DriverHomeZip=""

If Not rsbil.eof Then
strSql="select BillNo,CarNo,Rule1,Rule2,BillTypeID,Driver,DriverID,DriverAddress,DriverZip,Owner,OwnerZip,OwnerAddress,INSURANCE,ILLEGALADDRESS,RuleVer,IllegalSpeed,RuleSpeed,Note,BillFillDate,RECORDMEMBERID from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed=""
if Not rs.eof then
	Sys_BillNo=trim(rs("BillNo"))
	Sys_CarNo=trim(rs("CarNo"))
	Sys_Rule1=trim(rs("Rule1"))
	Sys_Rule2=trim(rs("Rule2"))
	Sys_BillTypeID=trim(rs("BillTypeID"))

	Sys_Driver=trim(rs("Driver"))
	Sys_DriverID=trim(rs("DriverID"))
	Sys_DriverHomeAddress=trim(rs("DriverAddress"))
	Sys_DriverHomeZip=trim(rs("DriverZip"))
	
	'Sys_Owner=trim(rs("Owner"))
	'Sys_OwnerZip=trim(rs("OwnerZip"))
	'Sys_OwnerAddress=trim(rs("OwnerAddress"))


	Sys_INSURANCE=trim(rs("INSURANCE"))
	Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
	Sys_RuleVer=trim(rs("RuleVer"))
	Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
	Sys_RuleSpeed=trim(rs("RuleSpeed"))
	Sys_Note=trim(rs("Note"))
	Sys_BillBaseRecordMemberID=trim(rs("RECORDMEMBERID"))
end if
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

'if Instr(request("Sys_BatchNumber"),"N")>0 then
'	Sys_OwnerZip=trim(Sys_DriverHomeZip)
'	Sys_OwnerAddress=trim(Sys_DriverHomeAddress)
'end if

If ifnull(Sys_OwnerAddress) Then

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		strSql="select * from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"
	else
		strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"
	end if

	set rsfound=conn.execute(strSql)

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

	else
		If Sys_BillTypeID=1 Then
			if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

			If ifnull(Sys_OwnerZip) Then
				if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
				if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
			end if
		else
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
		End if
		
		
	end if
	rsfound.close

	if Instr(request("Sys_BatchNumber"),"N")>0 and Sys_BillTypeID=1 then
		Sys_OwnerAddress=""
	end If 

	If ifnull(Sys_OwnerAddress) Then
		if Instr(request("Sys_BatchNumber"),"N")>0 then
			strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"

		else
			strSql="select * from BillbaseDCIReturn where CarNo='"&Sys_CarNo&"' and ExchangetypeID='A'"
		end if

		set rsdata=conn.execute(strsql)

		if Instr(request("Sys_BatchNumber"),"N")>0 then
			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

		else
			If Sys_BillTypeID=1 Then
				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

				If ifnull(Sys_OwnerZip) Then
					if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
					if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))
				end if
			else
				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))
			End if
		end if
		rsdata.close
	end if
End If 


Sys_Driver=""

strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)

If ifnull(Sys_Owner) Then

	If Sys_BillTypeID=1 Then

		if Not rsfound.eof then Sys_Owner=trim(rsfound("Driver"))

		If ifnull(Sys_Owner) Then
			Sys_Owner=trim(rsfound("Owner"))
		end if

	else
		if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))

	End If 

End if

If ifnull(Sys_OwnerAddress) Then
	if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
end if

If not ifnull(Sys_OwnerAddress) Then

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		strSQL="Update Billbase set DriverZip='"&trim(Sys_OwnerZip)&"',DriverAddress='"&trim(Sys_OwnerAddress)&"' where sn="&trim(rsbil("BillSN"))&" and DriverAddress is null"
		
		conn.execute(strSQL)
	else

		strSQL="Update Billbase set Owner='"&Sys_Owner&"',OwnerZip='"&Sys_OwnerZip&"',OwnerAddress='"&Sys_OwnerAddress&"' where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is null"

		conn.execute(strSQL)

	end if

End if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

Sys_DCIReturnStation=0
Sum_Level=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=cdbl(funTnumber(Sys_Level1))+cdbl(funTnumber(Sys_Level2))
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

'if isnull(Sys_OwnerZip) or trim(Sys_OwnerZip)="" then Sys_OwnerZip="001"

DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
end if
rsbil.close

%>
<div id="L78" style="position:relative;">
<div id="Layer8" style="position:absolute; left:0px; top:0px;">
<table border="0" height=295>
	<tr>
		<td class="style1" valign="top" align="right" height=295 width=5 >
			<!--<div id="Layer5" class="style1" style="position:absolute; left:20px; z-index:5"><%'response.write wordporss(chstr(Sys_BillUnitName))%></div>-->
		</td>
		<td class="style5" valign="bottom" align="right" height=295 width=25><%
			response.write wordporss(funcCheckFont(Sys_Owner,16,4)&"　"&chstr(Sys_CarNo))%>
		</td>
		<td class="style5" valign="bottom" align="right" height=295 width=5><%
		If not ifnull(Sys_OwnerAddress) Then
			Sys_OwnerAddress=replace(Sys_OwnerAddress,"臺","台")
			strtmp=Sys_OwnerZip&Sys_OwnerZipName&funcCheckFont(replace(Sys_OwnerAddress,Sys_OwnerZipName,""),16,4)
		else
			strtmp=Sys_OwnerZip&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,16,4)
		End if
		if len(strtmp)>15 and instr(strtmp,"img")=0 then
			response.write wordporss(chstr(mid(strtmp,1,15)))
		else
			response.write wordporss(chstr(strtmp))
		end if%>
		</td>
		<td class="style5" valign="bottom" align="right" height=295 width=10><%
			if len(strtmp)>15 and instr(strtmp,"img")=0 then response.write wordporss(chstr(mid(strtmp,16,len(strtmp))))%>
		</td>

		<td class="style1" valign="bottom" align="left" height=295 width=20><%
			response.write wordporss(chstr("　　　　　　　"&Sys_BillNo))%>
		</td>　
		<td class="style1" valign="bottom" align="right" height=295 width=15><%
			tmpstr="　　　　　　　　　　　"&left(trim(Sys_Rule1),2)&"　"
			if len(trim(Sys_Rule1))>7 then tmpstr=tmpstr&"　"&right(trim(Sys_Rule1),1)
		tmpstr=tmpstr&Mid(trim(Sys_Rule1),3,1)&Mid(trim(Sys_Rule1),4,2)
				'&Mid(trim(Sys_Rule1),6,2)&"規定。" tmpstr=tmpstr&",期限內自動繳納處新台幣"&Sys_Level1&"元"
				response.write wordporss(chstr(tmpstr))
			%>
		</td>
		<td class="style1" valign="bottom" align="right" height=295 width=15><%
			if trim(Sys_Rule2)<>"0" then
				tmpstr="　　　　　　　　　　　"&left(trim(Sys_Rule2),2)&"　"
				if len(trim(Sys_Rule2))>7 then tmpstr=tmpstr&"　"&right(trim(Sys_Rule2),1)
				tmpstr=tmpstr&Mid(trim(Sys_Rule2),3,1)&Mid(trim(Sys_Rule2),4,2)
				response.write wordporss(chstr(tmpstr))
			end if
			%>
		</td>
	</tr>
	<tr>
		<td width="15" valign="top">
			<div id="idDiv" class="style1" style="position:absolute; left:5px; z-index:5">
				<img src="../BarCodeImage/<%=Sys_BillNo%>.jpg" width="100" height="30">
			</div>
		</td>
		<td class="style1" valign="top" align="right">
			<%response.write wordporss(chstr(Sys_BillUnitName))%>
		</td>
		<td class="style1" valign="top" align="right">
		</td>
		<td class="style1" valign="top" align="right"><%
			If chkStore=0 Then
				response.write wordporss(chstr(Sys_MailNumber))
			else
				response.write wordporss(chstr(Sys_StoreAndSendMailNumber))
			End if%>
		</td>
	</tr>
</table>
</div>
</div>
