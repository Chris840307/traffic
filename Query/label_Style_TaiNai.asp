<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單信封黏貼標籤</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<body>
<%

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
Sys_ExchangetypeID="W"


	PBillSN = Split(trim(request("PBillSN")),",")


for i=0 to Ubound(PBillSN)

if cint(i)>0 and i mod 5=0 then response.write "<div class=""PageNext"">&nbsp;</div>"

if i mod 5=0 then 

'---------------------------------------------------------------------------------------

strBil="select distinct BatchNumber,BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='"&Sys_ExchangetypeID&"'"
set rsbil=conn.execute(strBil)
Sys_OwnerZip=""
Sys_OwnerZipName=""
If Not rsbil.eof Then
Sys_BatchNumber=trim(rsbil("BatchNumber"))

strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed="":sys_DriverID=""
if Not rs.eof then Sys_BillNo=trim(rs("BillNo"))
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then sys_DriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

Sys_OwnerAddress="":Sys_OwnerZip="":Sys_Owner="":Sys_OwnerZipName=""

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)
if Not rsfound.eof then
	Sys_Owner=trim(rsfound("Owner"))
	if Instr(request("Sys_BatchNumber"),"N")<=0 then
		If instr(chstr(rsfound("OwnerAddress")),"（") >0 Then
			Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			Sys_OwnerZip=trim(rsfound("OwnerZip"))

		end if

		if ifnull(Sys_OwnerAddress) then
			strSql="select DriverHomeAddress,DriverHomeZip from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"

			set rsdri=conn.execute(strSQL)

			If not rsdri.eof Then

				Sys_OwnerAddress=trim(rsdri("DriverHomeAddress"))
				Sys_OwnerZip=trim(rsdri("DriverHomeZip"))
			End if

			rsdri.close
		End If 
	else
		strSql="select DriverHomeAddress,DriverHomeZip from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"

		set rsdri=conn.execute(strSQL)

		If not rsdri.eof Then

			Sys_OwnerAddress=trim(rsdri("DriverHomeAddress"))
			Sys_OwnerZip=trim(rsdri("DriverHomeZip"))
		End if

		rsdri.close

		if ifnull(Sys_OwnerAddress) then
			Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
		end if
	end If 

	If trim(Sys_BillTypeID) = "1" or sys_DriverID<>"" Then
		Sys_Owner=trim(rsfound("Driver"))
		Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	End if 
	
	if ifnull(Sys_Owner) then
		Sys_Owner=trim(rsfound("Owner"))
	end if

	if ifnull(Sys_OwnerAddress) then
		Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		Sys_OwnerZip=trim(rsfound("OwnerZip"))
	end if
end if

strSQL="select BillTypeID,Driver,DriverZip,DriverAddress,Owner,OwnerZip,OwnerAddress from billbase where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is not null"
set rs1=conn.execute(strSQL)
If not rs1.eof Then
	If rs1("BillTypeID") = "1" and trim(rs1("DriverAddress"))<>"" Then
		Sys_OwnerAddress=trim(rs1("DriverAddress"))
		Sys_OwnerZip=trim(rs1("DriverZip"))
		Sys_Owner=trim(rs1("Driver"))
	elseif Instr(request("Sys_BatchNumber"),"N")<=0 then

		Sys_OwnerAddress=trim(rs1("OwnerAddress"))
		Sys_OwnerZip=trim(rs1("OwnerZip"))
		Sys_Owner=trim(rs1("Owner"))

	elseif trim(rs1("DriverAddress"))<>"" Then

		Sys_OwnerAddress=trim(rs1("DriverAddress"))
		Sys_OwnerZip=trim(rs1("DriverZip"))
		Sys_Owner=trim(rs1("Driver"))
	
	else
		
		Sys_OwnerAddress=trim(rs1("OwnerAddress"))
		Sys_OwnerZip=trim(rs1("OwnerZip"))
		Sys_Owner=trim(rs1("Owner"))
	End if 

	if ifnull(Sys_Owner) then
		Sys_Owner=trim(rs1("Owner"))
	end If 
	
end If 

rs1.close

Sys_OwnerAddress=replace(Sys_OwnerAddress&" ","臺","台")

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

Sys_DCIReturnStation=0
Sum_Level=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo"))
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1"))
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2"))
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=cdbl(Sys_Level1)+cdbl(Sys_Level2)
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
'if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
'if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
'if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

strSql="select a.LoginID,a.ChName,a.ImageFilename as MemberFileName,b.UnitName,b.UnitID,b.UnitLevelID,b.UnitTypeID,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_Billmem1ID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillUnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
'if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
'if Not mem.eof then Sys_ChName=trim(mem("ChName"))
'if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

'20120515 by jafe 新營拖吊小隊一直抓到上面的隆田，所以加入這一段
If sys_City = "台南市" Then
	If Sys_UnitID = "07D3" or Sys_UnitID = "07D4" or Sys_UnitID = "07D2" Then 
		Sys_UnitTypeID = Sys_UnitID
	End if
End if

strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"

set unit=conn.Execute(strSQL)
If Not unit.eof Then sysunit=unit("UnitName")
if Not unit.eof then SysUnitTel=trim(unit("Tel"))
if Not unit.eof then SysUnitAddress=trim(unit("Address"))
unit.close

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

strSql="select MailNumber,StoreAndSendMailNumber,MailTypeID,MailDate,UserMarkDate,StoreAndSendFinalMailDate from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))
if Not rs.eof then Sys_StoreAndSendFinalMailDate=trim(rs("StoreAndSendFinalMailDate"))
if Not rs.eof then UserMarkDate=trim(rs("UserMarkDate"))
rs.close

strSql="select LoginID from MemberData where MemberID="&session("User_ID")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_LoginID=trim(rs("LoginID"))
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

Sys_MailNumber=0
Sys_BillNo_BarCode=Sys_BillNo

If sys_City="高雄市" Then
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160,1
else
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160
end if

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
end if
rsbil.close

%>
	
	<table border="0" width="80%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td height="100" width="60%">
				<p align="center">
				<font face="標楷體" size="5"><%
				'宜蘭停管單退要抓戶籍
				if (Sys_BillTypeID="1") and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_Driver)="" Then 				
					response.write funcCheckFont(Sys_Owner,18,1)
				  Else
					response.write funcCheckFont(Sys_Owner,18,1) 
				  End if
				else 
					response.write funcCheckFont(Sys_Owner,18,1)
				end if
				%>
				<%if sys_City="彰化縣" then
				response.write "　先生/女士"
				end if
				%>
				</font></td>
    <td>
              <p align="left">
              <img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg"><font face="標楷體" size="2"><%if sys_City="台中縣" then  response.write StationName%></font></td>
    		</tr>
	    	<tr>
			<td height="120" align="left" valign="top" colspan="2">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>""  then
				  If Trim(Sys_DriverHomeZip)="" Then 
					response.write OwnerZip 
				  Else
					response.write Sys_DriverHomeZip 
				  End if
				else 
					response.write Sys_OwnerZip  
				end if
				%>&nbsp;<%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>""  then 
				  If Trim(Sys_DriverHomeAddress)="" Then 
					response.write funcCheckFont(GetMailAddress,18,1)
				  Else
					response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
				  End if
				else 
					response.write replace(Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) ,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName)
				end if
					%></font>
         　</td>
		</tr>
	</table>
	<% 

	'---------------------------------------------------------------------------------------
	if (i+1 < Ubound(PBillSN)) or (i+1 = Ubound(PBillSN))then 

strBil="select distinct BatchNumber,BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i+1)&" and ExchangetypeID='"&Sys_ExchangetypeID&"'"
set rsbil=conn.execute(strBil)
Sys_OwnerZip=""
Sys_OwnerZipName=""
If Not rsbil.eof Then
Sys_BatchNumber=trim(rsbil("BatchNumber"))

strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed=""
if Not rs.eof then Sys_BillNo=trim(rs("BillNo"))
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

Sys_OwnerAddress="":Sys_OwnerZip="":Sys_Owner="":Sys_OwnerZipName=""

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)
if Not rsfound.eof then
	Sys_Owner=trim(rsfound("Owner"))
	if Instr(request("Sys_BatchNumber"),"N")<=0 then
		If instr(chstr(rsfound("OwnerAddress")),"（") >0 Then
			Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			Sys_OwnerZip=trim(rsfound("OwnerZip"))

		end if

		if ifnull(Sys_OwnerAddress) then
			strSql="select DriverHomeAddress,DriverHomeZip from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"

			set rsdri=conn.execute(strSQL)

			If not rsdri.eof Then

				Sys_OwnerAddress=trim(rsdri("DriverHomeAddress"))
				Sys_OwnerZip=trim(rsdri("DriverHomeZip"))
			End if

			rsdri.close
		End If 
	else
		strSql="select DriverHomeAddress,DriverHomeZip from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"

		set rsdri=conn.execute(strSQL)

		If not rsdri.eof Then

			Sys_OwnerAddress=trim(rsdri("DriverHomeAddress"))
			Sys_OwnerZip=trim(rsdri("DriverHomeZip"))
		End if

		rsdri.close

		if ifnull(Sys_OwnerAddress) then
			Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
		end if
	end If 

	If trim(Sys_BillTypeID) = "1" Then
		Sys_Owner=trim(rsfound("Driver"))
		Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	End if 
	
	if ifnull(Sys_Owner) then
		Sys_Owner=trim(rsfound("Owner"))
	end if

	if ifnull(Sys_OwnerAddress) then
		Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		Sys_OwnerZip=trim(rsfound("OwnerZip"))
	end if
end if

strSQL="select BillTypeID,Driver,DriverZip,DriverAddress,Owner,OwnerZip,OwnerAddress from billbase where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is not null"
set rs1=conn.execute(strSQL)
If not rs1.eof Then
	If rs1("BillTypeID") = "1" and trim(rs1("DriverAddress"))<>"" Then
		Sys_OwnerAddress=trim(rs1("DriverAddress"))
		Sys_OwnerZip=trim(rs1("DriverZip"))
		Sys_Owner=trim(rs1("Driver"))
	elseif Instr(request("Sys_BatchNumber"),"N")<=0 then

		Sys_OwnerAddress=trim(rs1("OwnerAddress"))
		Sys_OwnerZip=trim(rs1("OwnerZip"))
		Sys_Owner=trim(rs1("Owner"))

	elseif trim(rs1("DriverAddress"))<>"" Then

		Sys_OwnerAddress=trim(rs1("DriverAddress"))
		Sys_OwnerZip=trim(rs1("DriverZip"))
		Sys_Owner=trim(rs1("Driver"))
	
	else
		
		Sys_OwnerAddress=trim(rs1("OwnerAddress"))
		Sys_OwnerZip=trim(rs1("OwnerZip"))
		Sys_Owner=trim(rs1("Owner"))
	End if 

	if ifnull(Sys_Owner) then
		Sys_Owner=trim(rs1("Owner"))
	end If 
	
end If 

rs1.close

Sys_OwnerAddress=replace(Sys_OwnerAddress&" ","臺","台")

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

Sys_DCIReturnStation=0
Sum_Level=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo"))
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1"))
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2"))
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=cdbl(Sys_Level1)+cdbl(Sys_Level2)
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
'if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
'if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
'if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

strSql="select a.LoginID,a.ChName,a.ImageFilename as MemberFileName,b.UnitName,b.UnitID,b.UnitLevelID,b.UnitTypeID,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_Billmem1ID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillUnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
'if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
'if Not mem.eof then Sys_ChName=trim(mem("ChName"))
'if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

'20120515 by jafe 新營拖吊小隊一直抓到上面的隆田，所以加入這一段
If sys_City = "台南市" Then
	If Sys_UnitID = "07D3" or Sys_UnitID = "07D4" or Sys_UnitID = "07D2" Then 
		Sys_UnitTypeID = Sys_UnitID
	End if
End if

strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"

set unit=conn.Execute(strSQL)
If Not unit.eof Then sysunit=unit("UnitName")
if Not unit.eof then SysUnitTel=trim(unit("Tel"))
if Not unit.eof then SysUnitAddress=trim(unit("Address"))
unit.close

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

strSql="select MailNumber,StoreAndSendMailNumber,MailTypeID,MailDate,UserMarkDate,StoreAndSendFinalMailDate from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))
if Not rs.eof then Sys_StoreAndSendFinalMailDate=trim(rs("StoreAndSendFinalMailDate"))
if Not rs.eof then UserMarkDate=trim(rs("UserMarkDate"))
rs.close

strSql="select LoginID from MemberData where MemberID="&session("User_ID")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_LoginID=trim(rs("LoginID"))
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

Sys_MailNumber=0
Sys_BillNo_BarCode=Sys_BillNo

If sys_City="高雄市" Then
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160,1
else
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160
end if

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
end if
rsbil.close
'-------------------------------------------------------------------------------------
	%>
		<table border="0" width="80%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td height="100" width="60%">
				<p align="center">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1") and trim(Sys_DriverHomeAddress)<>"" Then
				  If Trim(Sys_Driver)="" Then 				
					response.write funcCheckFont(Sys_Owner,18,1) 
				  Else
					response.write funcCheckFont(Sys_Owner,18,1)  
				  End if
				else
					response.write funcCheckFont(Sys_Owner,18,1) 				
				end if
				%>				
				<%if sys_City="彰化縣" then
				response.write "　先生/女士"
				end if
				%>
				</font></td>
    <td>
              <p align="left">
              <img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg"><font face="標楷體" size="2"><%if sys_City="台中縣" then  response.write StationName%></font></td>
    		</tr>
	    	<tr>
			<td height="120" align="left" valign="top" colspan="2">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_DriverHomeZip)="" Then 
					response.write OwnerZip 
				  Else
					response.write Sys_DriverHomeZip 
				  End if
				else 
					response.write Sys_OwnerZip
				end if
				%>&nbsp;<%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_DriverHomeAddress)="" Then 
					response.write funcCheckFont(GetMailAddress,18,1) 
				  Else
					response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
				  End if
				else
'					response.write Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) 
					response.write replace(Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) ,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName)
				end if
				%></font>
         　</td>
		</tr>
	</table>
	<% 
end if	
	'---------------------------------------------------------------------------------------
	if (i+2 < Ubound(PBillSN)) or (i+2 = Ubound(PBillSN))then 

strBil="select distinct BatchNumber,BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i+2)&" and ExchangetypeID='"&Sys_ExchangetypeID&"'"
set rsbil=conn.execute(strBil)
Sys_OwnerZip=""
Sys_OwnerZipName=""
If Not rsbil.eof Then
Sys_BatchNumber=trim(rsbil("BatchNumber"))

strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed=""

if Not rs.eof then Sys_BillNo=trim(rs("BillNo"))
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

Sys_OwnerAddress="":Sys_OwnerZip="":Sys_Owner="":Sys_OwnerZipName=""

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)
if Not rsfound.eof then
	Sys_Owner=trim(rsfound("Owner"))
	if Instr(request("Sys_BatchNumber"),"N")<=0 then
		If instr(chstr(rsfound("OwnerAddress")),"（") >0 Then
			Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			Sys_OwnerZip=trim(rsfound("OwnerZip"))

		end if

		if ifnull(Sys_OwnerAddress) then
			strSql="select DriverHomeAddress,DriverHomeZip from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"

			set rsdri=conn.execute(strSQL)

			If not rsdri.eof Then

				Sys_OwnerAddress=trim(rsdri("DriverHomeAddress"))
				Sys_OwnerZip=trim(rsdri("DriverHomeZip"))
			End if

			rsdri.close
		End If 
	else
		strSql="select DriverHomeAddress,DriverHomeZip from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"

		set rsdri=conn.execute(strSQL)

		If not rsdri.eof Then

			Sys_OwnerAddress=trim(rsdri("DriverHomeAddress"))
			Sys_OwnerZip=trim(rsdri("DriverHomeZip"))
		End if

		rsdri.close

		if ifnull(Sys_OwnerAddress) then
			Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
		end if
	end If 

	If trim(Sys_BillTypeID) = "1" Then
		Sys_Owner=trim(rsfound("Driver"))
		Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	End if 
	
	if ifnull(Sys_Owner) then
		Sys_Owner=trim(rsfound("Owner"))
	end if

	if ifnull(Sys_OwnerAddress) then
		Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		Sys_OwnerZip=trim(rsfound("OwnerZip"))
	end if
end if

strSQL="select BillTypeID,Driver,DriverZip,DriverAddress,Owner,OwnerZip,OwnerAddress from billbase where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is not null"
set rs1=conn.execute(strSQL)
If not rs1.eof Then
	If rs1("BillTypeID") = "1" and trim(rs1("DriverAddress"))<>"" Then
		Sys_OwnerAddress=trim(rs1("DriverAddress"))
		Sys_OwnerZip=trim(rs1("DriverZip"))
		Sys_Owner=trim(rs1("Driver"))
	elseif Instr(request("Sys_BatchNumber"),"N")<=0 then

		Sys_OwnerAddress=trim(rs1("OwnerAddress"))
		Sys_OwnerZip=trim(rs1("OwnerZip"))
		Sys_Owner=trim(rs1("Owner"))

	elseif trim(rs1("DriverAddress"))<>"" Then

		Sys_OwnerAddress=trim(rs1("DriverAddress"))
		Sys_OwnerZip=trim(rs1("DriverZip"))
		Sys_Owner=trim(rs1("Driver"))
	
	else
		
		Sys_OwnerAddress=trim(rs1("OwnerAddress"))
		Sys_OwnerZip=trim(rs1("OwnerZip"))
		Sys_Owner=trim(rs1("Owner"))
	End if 

	if ifnull(Sys_Owner) then
		Sys_Owner=trim(rs1("Owner"))
	end If 
	
end If 

rs1.close

Sys_OwnerAddress=replace(Sys_OwnerAddress&" ","臺","台")

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

Sys_DCIReturnStation=0
Sum_Level=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo"))
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1"))
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2"))
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=cdbl(Sys_Level1)+cdbl(Sys_Level2)
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
'if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
'if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
'if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

strSql="select a.LoginID,a.ChName,a.ImageFilename as MemberFileName,b.UnitName,b.UnitID,b.UnitLevelID,b.UnitTypeID,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_Billmem1ID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillUnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
'if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
'if Not mem.eof then Sys_ChName=trim(mem("ChName"))
'if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

'20120515 by jafe 新營拖吊小隊一直抓到上面的隆田，所以加入這一段
If sys_City = "台南市" Then
	If Sys_UnitID = "07D3" or Sys_UnitID = "07D4" or Sys_UnitID = "07D2" Then 
		Sys_UnitTypeID = Sys_UnitID
	End if
End if

strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"

set unit=conn.Execute(strSQL)
If Not unit.eof Then sysunit=unit("UnitName")
if Not unit.eof then SysUnitTel=trim(unit("Tel"))
if Not unit.eof then SysUnitAddress=trim(unit("Address"))
unit.close

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

strSql="select MailNumber,StoreAndSendMailNumber,MailTypeID,MailDate,UserMarkDate,StoreAndSendFinalMailDate from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))
if Not rs.eof then Sys_StoreAndSendFinalMailDate=trim(rs("StoreAndSendFinalMailDate"))
if Not rs.eof then UserMarkDate=trim(rs("UserMarkDate"))
rs.close

strSql="select LoginID from MemberData where MemberID="&session("User_ID")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_LoginID=trim(rs("LoginID"))
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

Sys_MailNumber=0
Sys_BillNo_BarCode=Sys_BillNo

If sys_City="高雄市" Then
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160,1
else
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160
end if

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
end if
rsbil.close
'-------------------------------------------------------------------------------------
	%>
	
		<table border="0" width="80%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td height="100" width="60%">
				<p align="center">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1")  and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_Driver)="" Then 				
					response.write funcCheckFont(Sys_Owner,18,1) 
				  Else
					response.write funcCheckFont(Sys_Owner,18,1)  
				  End if
				else 
					response.write funcCheckFont(Sys_Owner,18,1)
				end if
				%>				
				<%if sys_City="彰化縣" then
				response.write "　先生/女士"
				end if
				%>
    <td>
              <p align="left">
              <img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg"><font face="標楷體" size="2"><%if sys_City="台中縣" then  response.write StationName%></font></td>
    		</tr>
	    	<tr>
			<td height="120" align="left" valign="top" colspan="2">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_DriverHomeZip)="" Then 
					response.write OwnerZip 
				  Else
					response.write Sys_DriverHomeZip 
				  End if
				else 
					response.write Sys_OwnerZip 
				end if
				%>&nbsp;<%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕"))  and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_DriverHomeAddress)="" Then 
					response.write funcCheckFont(GetMailAddress,18,1)
				  Else
					response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
				  End if
				else 
'					response.write Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) 
					response.write replace(Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) ,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName)
				end if
				%></font>
         　</td>
		</tr>
	</table>
	<% 
	'---------------------------------------------------------------------------------------
	end if
	if (i+3 < Ubound(PBillSN)) or (i+3 = Ubound(PBillSN))then 

strBil="select distinct BatchNumber,BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i+3)&" and ExchangetypeID='"&Sys_ExchangetypeID&"'"
set rsbil=conn.execute(strBil)
Sys_OwnerZip=""
Sys_OwnerZipName=""
If Not rsbil.eof Then
Sys_BatchNumber=trim(rsbil("BatchNumber"))

strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed=""

if Not rs.eof then Sys_BillNo=trim(rs("BillNo"))
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

Sys_OwnerAddress="":Sys_OwnerZip="":Sys_Owner="":Sys_OwnerZipName=""

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)
if Not rsfound.eof then
	Sys_Owner=trim(rsfound("Owner"))
	if Instr(request("Sys_BatchNumber"),"N")<=0 then
		If instr(chstr(rsfound("OwnerAddress")),"（") >0 Then
			Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			Sys_OwnerZip=trim(rsfound("OwnerZip"))

		end if

		if ifnull(Sys_OwnerAddress) then
			strSql="select DriverHomeAddress,DriverHomeZip from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"

			set rsdri=conn.execute(strSQL)

			If not rsdri.eof Then

				Sys_OwnerAddress=trim(rsdri("DriverHomeAddress"))
				Sys_OwnerZip=trim(rsdri("DriverHomeZip"))
			End if

			rsdri.close
		End If 
	else
		strSql="select DriverHomeAddress,DriverHomeZip from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"

		set rsdri=conn.execute(strSQL)

		If not rsdri.eof Then

			Sys_OwnerAddress=trim(rsdri("DriverHomeAddress"))
			Sys_OwnerZip=trim(rsdri("DriverHomeZip"))
		End if

		rsdri.close

		if ifnull(Sys_OwnerAddress) then
			Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
		end if
	end If 

	If trim(Sys_BillTypeID) = "1" Then
		Sys_Owner=trim(rsfound("Driver"))
		Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	End if 
	
	if ifnull(Sys_Owner) then
		Sys_Owner=trim(rsfound("Owner"))
	end if

	if ifnull(Sys_OwnerAddress) then
		Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		Sys_OwnerZip=trim(rsfound("OwnerZip"))
	end if
end if

strSQL="select BillTypeID,Driver,DriverZip,DriverAddress,Owner,OwnerZip,OwnerAddress from billbase where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is not null"
set rs1=conn.execute(strSQL)
If not rs1.eof Then
	If rs1("BillTypeID") = "1" and trim(rs1("DriverAddress"))<>"" Then
		Sys_OwnerAddress=trim(rs1("DriverAddress"))
		Sys_OwnerZip=trim(rs1("DriverZip"))
		Sys_Owner=trim(rs1("Driver"))
	elseif Instr(request("Sys_BatchNumber"),"N")<=0 then

		Sys_OwnerAddress=trim(rs1("OwnerAddress"))
		Sys_OwnerZip=trim(rs1("OwnerZip"))
		Sys_Owner=trim(rs1("Owner"))

	elseif trim(rs1("DriverAddress"))<>"" Then

		Sys_OwnerAddress=trim(rs1("DriverAddress"))
		Sys_OwnerZip=trim(rs1("DriverZip"))
		Sys_Owner=trim(rs1("Driver"))
	
	else
		
		Sys_OwnerAddress=trim(rs1("OwnerAddress"))
		Sys_OwnerZip=trim(rs1("OwnerZip"))
		Sys_Owner=trim(rs1("Owner"))
	End if 

	if ifnull(Sys_Owner) then
		Sys_Owner=trim(rs1("Owner"))
	end If 
	
end If 
rs1.close

Sys_OwnerAddress=replace(Sys_OwnerAddress&" ","臺","台")

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

Sys_DCIReturnStation=0
Sum_Level=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo"))
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1"))
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2"))
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=cdbl(Sys_Level1)+cdbl(Sys_Level2)
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
'if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
'if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
'if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

strSql="select a.LoginID,a.ChName,a.ImageFilename as MemberFileName,b.UnitName,b.UnitID,b.UnitLevelID,b.UnitTypeID,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_Billmem1ID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillUnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
'if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
'if Not mem.eof then Sys_ChName=trim(mem("ChName"))
'if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

'20120515 by jafe 新營拖吊小隊一直抓到上面的隆田，所以加入這一段
If sys_City = "台南市" Then
	If Sys_UnitID = "07D3" or Sys_UnitID = "07D4" or Sys_UnitID = "07D2" Then 
		Sys_UnitTypeID = Sys_UnitID
	End if
End if

strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"

set unit=conn.Execute(strSQL)
If Not unit.eof Then sysunit=unit("UnitName")
if Not unit.eof then SysUnitTel=trim(unit("Tel"))
if Not unit.eof then SysUnitAddress=trim(unit("Address"))
unit.close

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

strSql="select MailNumber,StoreAndSendMailNumber,MailTypeID,MailDate,UserMarkDate,StoreAndSendFinalMailDate from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))
if Not rs.eof then Sys_StoreAndSendFinalMailDate=trim(rs("StoreAndSendFinalMailDate"))
if Not rs.eof then UserMarkDate=trim(rs("UserMarkDate"))
rs.close

strSql="select LoginID from MemberData where MemberID="&session("User_ID")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_LoginID=trim(rs("LoginID"))
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

Sys_MailNumber=0
Sys_BillNo_BarCode=Sys_BillNo

If sys_City="高雄市" Then
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160,1
else
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160
end if

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
end if
rsbil.close
'-------------------------------------------------------------------------------------
	%>
	
		<table border="0" width="80%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td height="100" width="60%">
				<p align="center">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1") and trim(Sys_DriverHomeAddress)<>"" then
				  If Trim(Sys_Driver)="" Then 				
					response.write funcCheckFont(Sys_Owner,18,1) 
				  Else
					response.write funcCheckFont(Sys_Owner,18,1)  
				  End if
				else
					response.write funcCheckFont(Sys_Owner,18,1) 
				end if
				%>				
				<%if sys_City="彰化縣" then
				response.write "　先生/女士"
				end if
				%>
    <td>
              <p align="left">
              <img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg"><font face="標楷體" size="2"><%if sys_City="台中縣" then  response.write StationName%></font></td>
    		</tr>
	    	<tr>
			<td height="120" align="left" valign="top" colspan="2">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>"" then
				  If Trim(Sys_DriverHomeZip)="" Then 
					response.write OwnerZip 
				  Else
					response.write Sys_DriverHomeZip 
				  End if

				else 
					response.write Sys_OwnerZip  
				end if
				%>&nbsp;<%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_DriverHomeAddress)="" Then 
					response.write funcCheckFont(GetMailAddress,18,1)
				  Else
					response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
				  End if
				else 
'					response.write Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) 
					response.write replace(Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) ,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName)
				end if
				%></font>
         　</td>
		</tr>
	</table>
	<% 
	'---------------------------------------------------------------------------------------
	end if
	if (i+4 < Ubound(PBillSN)) or (i+4 = Ubound(PBillSN))then 

strBil="select distinct BatchNumber,BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i+4)&" and ExchangetypeID='"&Sys_ExchangetypeID&"'"
set rsbil=conn.execute(strBil)
Sys_OwnerZip=""
Sys_OwnerZipName=""
If Not rsbil.eof Then
Sys_BatchNumber=trim(rsbil("BatchNumber"))

strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed=""
if Not rs.eof then Sys_BillNo=trim(rs("BillNo"))
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

Sys_OwnerAddress="":Sys_OwnerZip="":Sys_Owner="":Sys_OwnerZipName=""

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)
if Not rsfound.eof then
	Sys_Owner=trim(rsfound("Owner"))
	if Instr(request("Sys_BatchNumber"),"N")<=0 then
		If instr(chstr(rsfound("OwnerAddress")),"（") >0 Then
			Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			Sys_OwnerZip=trim(rsfound("OwnerZip"))

		end if

		if ifnull(Sys_OwnerAddress) then
			strSql="select DriverHomeAddress,DriverHomeZip from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"

			set rsdri=conn.execute(strSQL)

			If not rsdri.eof Then

				Sys_OwnerAddress=trim(rsdri("DriverHomeAddress"))
				Sys_OwnerZip=trim(rsdri("DriverHomeZip"))
			End if

			rsdri.close
		End If 
	else
		strSql="select DriverHomeAddress,DriverHomeZip from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"

		set rsdri=conn.execute(strSQL)

		If not rsdri.eof Then

			Sys_OwnerAddress=trim(rsdri("DriverHomeAddress"))
			Sys_OwnerZip=trim(rsdri("DriverHomeZip"))
		End if

		rsdri.close

		if ifnull(Sys_OwnerAddress) then
			Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
		end if
	end If 

	If trim(Sys_BillTypeID) = "1" Then
		Sys_Owner=trim(rsfound("Driver"))
		Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	End if 
	
	if ifnull(Sys_Owner) then
		Sys_Owner=trim(rsfound("Owner"))
	end if

	if ifnull(Sys_OwnerAddress) then
		Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		Sys_OwnerZip=trim(rsfound("OwnerZip"))
	end if
end if

strSQL="select BillTypeID,Driver,DriverZip,DriverAddress,Owner,OwnerZip,OwnerAddress from billbase where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is not null"
set rs1=conn.execute(strSQL)
If not rs1.eof Then
	If rs1("BillTypeID") = "1" and trim(rs1("DriverAddress"))<>"" Then
		Sys_OwnerAddress=trim(rs1("DriverAddress"))
		Sys_OwnerZip=trim(rs1("DriverZip"))
		Sys_Owner=trim(rs1("Driver"))
	elseif Instr(request("Sys_BatchNumber"),"N")<=0 then

		Sys_OwnerAddress=trim(rs1("OwnerAddress"))
		Sys_OwnerZip=trim(rs1("OwnerZip"))
		Sys_Owner=trim(rs1("Owner"))

	elseif trim(rs1("DriverAddress"))<>"" Then

		Sys_OwnerAddress=trim(rs1("DriverAddress"))
		Sys_OwnerZip=trim(rs1("DriverZip"))
		Sys_Owner=trim(rs1("Driver"))
	
	else
		
		Sys_OwnerAddress=trim(rs1("OwnerAddress"))
		Sys_OwnerZip=trim(rs1("OwnerZip"))
		Sys_Owner=trim(rs1("Owner"))
	End if 

	if ifnull(Sys_Owner) then
		Sys_Owner=trim(rs1("Owner"))
	end If 
	
end If 
rs1.close

Sys_OwnerAddress=replace(Sys_OwnerAddress&" ","臺","台")

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

Sys_DCIReturnStation=0
Sum_Level=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation"))
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo"))
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1"))
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2"))
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=cdbl(Sys_Level1)+cdbl(Sys_Level2)
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
'if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
'if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
'if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

strSql="select a.LoginID,a.ChName,a.ImageFilename as MemberFileName,b.UnitName,b.UnitID,b.UnitLevelID,b.UnitTypeID,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_Billmem1ID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillUnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
'if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
'if Not mem.eof then Sys_ChName=trim(mem("ChName"))
'if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

'20120515 by jafe 新營拖吊小隊一直抓到上面的隆田，所以加入這一段
If sys_City = "台南市" Then
	If Sys_UnitID = "07D3" or Sys_UnitID = "07D4" or Sys_UnitID = "07D2" Then 
		Sys_UnitTypeID = Sys_UnitID
	End if
End if

strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"

set unit=conn.Execute(strSQL)
If Not unit.eof Then sysunit=unit("UnitName")
if Not unit.eof then SysUnitTel=trim(unit("Tel"))
if Not unit.eof then SysUnitAddress=trim(unit("Address"))
unit.close

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

strSql="select MailNumber,StoreAndSendMailNumber,MailTypeID,MailDate,UserMarkDate,StoreAndSendFinalMailDate from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))
if Not rs.eof then Sys_StoreAndSendFinalMailDate=trim(rs("StoreAndSendFinalMailDate"))
if Not rs.eof then UserMarkDate=trim(rs("UserMarkDate"))
rs.close

strSql="select LoginID from MemberData where MemberID="&session("User_ID")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_LoginID=trim(rs("LoginID"))
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

Sys_MailNumber=0
Sys_BillNo_BarCode=Sys_BillNo

If sys_City="高雄市" Then
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160,1
else
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160
end if

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
end if
rsbil.close
'-------------------------------------------------------------------------------------
	%>
	
		<table border="0" width="80%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td height="100" width="60%">
				<p align="center">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1") and trim(Sys_DriverHomeAddress)<>"" then
				  If Trim(Sys_Driver)="" Then 				
					response.write funcCheckFont(Sys_Owner,18,1) 
				  Else
					response.write funcCheckFont(Sys_Owner,18,1) 
				  End if
				else 
					response.write funcCheckFont(Sys_Owner,18,1)
				end if
				%>				
				<%if sys_City="彰化縣" then
				response.write "　先生/女士"
				end if
				%>
    <td>
              <p align="left">
              <img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg"><font face="標楷體" size="2"><%if sys_City="台中縣" then  response.write StationName%></font></td>
    		</tr>
	    	<tr>
			<td height="50" align="left" valign="top" colspan="2">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕"))  and trim(Sys_DriverHomeAddress)<>""  then 
				  If Trim(Sys_DriverHomeZip)="" Then 
					response.write OwnerZip 
				  Else
					response.write Sys_DriverHomeZip 
				  End if
				else
					response.write Sys_OwnerZip 
				end if
				%>&nbsp;<%
				if (Sys_BillTypeID="1"  or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕"))  then 
				  If Trim(Sys_DriverHomeAddress)="" Then 
					response.write funcCheckFont(GetMailAddress,18,1)
				  Else
					response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
				  End if
				else 
'					response.write Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) 
					response.write replace(Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) ,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName)
				end if
				%></font>
         　</td>
		</tr>
	</table>
	
<% 
End if
end if
%>
<%next%>
</body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="../smsx.cab#Version=6,1,432,1">
</object>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
printWindow(true,5.08,5.08,5.08,5.08);
</script></p>