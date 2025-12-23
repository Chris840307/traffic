<%
strSQL="select * from UnitInfo where UnitID='"&DB_UnitID&"'"
set rsUnit=conn.execute(strSQL)
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
rsUnit.close

strBil="select distinct BatchNumber,BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_OwnerZip=""
Sys_OwnerZipName=""
CNum=""

db_BatchNumber=""
If not ifnull(request("Sys_BatchNumber")) Then
	db_BatchNumber=trim(request("Sys_BatchNumber"))
else
	db_BatchNumber=trim(rsbil("BatchNumber"))
End if

strSQL="select cnt from (select RowNum cnt,BillSN from (select BillSN from DCILog where BatchNumber='"&trim(db_BatchNumber)&"' order by BillSN) order by BillSN) where BillSN="&PBillSN(i)

set dci=conn.execute(strSQL)
if not dci.eof then CNum=dci("cnt")
dci.close
Sys_BatChNumber=""
If not ifnull(request("Sys_BatchNumber")) Then
	Sys_BatChNumber=trim(request("Sys_BatchNumber"))&"_"&(CNum)
else
	Sys_BatChNumber=trim(rsbil("BatchNumber"))&"_"&(CNum)
End if

strSql="select BillTypeID,Driver,DriverID,DriverAddress,DriverZip,INSURANCE,ILLEGALADDRESS,RuleVer,IllegalSpeed,RuleSpeed,Note,BillFillDate,RECORDMEMBERID from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)

Sys_IllegalSpeed="":Sys_RuleSpeed=""
Sys_OwnerAddress="":Sys_OwnerZip=""

if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then Sys_BillBaseRecordMemberID=trim(rs("RECORDMEMBERID"))
if Not rs.eof then Sys_OwnerAddress=trim(rs("DriverAddress"))
if Not rs.eof then Sys_OwnerZip=trim(rs("DriverZip"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

If Sys_BillTypeID=1 Then
	Sys_OwnerAddress="":Sys_OwnerZip=""
end if

If ifnull(Sys_OwnerAddress) Then
	if Instr(request("Sys_BatchNumber"),"N")>0 then
		strSql="select * from BillbaseDCIReturn where Carno=(select carno from dcilog where billsn="&trim(rsbil("BillSN"))&" and ExchangetypeID='A') and ExchangetypeID='A'"
	else
		strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
	end if

	set rsfound=conn.execute(strSql)

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

	else
		If Sys_BillTypeID=1 Then
			if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

			If ifnull(Sys_OwnerAddress) Then
				if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
				if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
			end if
		else
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
		End if
		
		
	end if
	rsfound.close
end if

If ifnull(Sys_OwnerAddress) Then
	if Instr(request("Sys_BatchNumber"),"N")>0 then
		strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

	else
		strSql="select * from BillbaseDCIReturn where Carno=(select carno from dcilog where billsn="&trim(rsbil("BillSN"))&" and ExchangetypeID='A') and ExchangetypeID='A'"
	end if

	set rsdata=conn.execute(strsql)

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
		if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

	else
		If Sys_BillTypeID=1 Then
			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

			If ifnull(Sys_OwnerAddress) Then
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

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)

If Sys_BillTypeID=1 Then

	if Not rsfound.eof then Sys_Owner=trim(rsfound("Driver"))

	If ifnull(Sys_Owner) Then
		Sys_Owner=trim(rsfound("Owner"))
	end if

else
	if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))

End if

If ifnull(Sys_OwnerAddress) Then
	if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=replace(trim(rszip("ZipName")),"台","臺")
if Not rszip.eof then
	Sys_OwnerAddress=replace(Sys_OwnerAddress,"台","臺")

	if not isnull(Sys_OwnerAddress) then '6/25
		Sys_OwnerAddress=replace(Sys_OwnerAddress,Sys_OwnerZipName,"")
	end If 

end if
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
Sum_Level=CDBL(Sys_Level1)+CDBL(Sys_Level2)
if Not rsfound.eof then Sys_DCIRETURNCARTYPE=trim(rsfound("DCIRETURNCARTYPE"))
strsql="select * from DCICODE where ID='"&Sys_DCIRETURNCARTYPE&"' and TypeID=5"
Sys_DCIRETURNCARTYPE=""
set cartype=conn.execute(strsql)
if Not cartype.eof then Sys_DCIRETURNCARTYPE=trim(cartype("Content"))
cartype.close

rsfound.close
Sys_Sex=""
strSql="select distinct BillFillerMemberID,DriverSex,DriverBirth,IllegalDate,DealLineDate,RECORDMEMBERID,BillFillerMemberID,IMAGEFILENAME,IMAGEPATHNAME,IMAGEFILENAMEB,BILLMEMID1 from BillBase where SN="&trim(rsbil("BillSN"))
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
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
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
if Not unit.eof then Sys_UnitTel=trim(unit("Tel"))
if Not unit.eof then Sys_UnitAddress=trim(unit("Address"))
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

if Instr(request("Sys_BatchNumber"),"N")>0 then Sys_MailNumber=Sys_StoreAndSendMailNumber

If ifnull(Sys_MailNumber) Then Sys_MailNumber=0


	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,364,000,17
	
'	DelphiASPObj.CreateBarCode Sys_MailNumber&"80026336"
'	response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_OwnerZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate&",802,451,17"
'	response.end


strSql="select MAILCHKNUMBER from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
Sys_MAILCHKNUMBER=Sys_MailNumber&"36400017"
rs.close
If Sys_OwnerZip="001" then Sys_OwnerZip=""
rsbil.close
%>
<div id="L78" style="position:relative;">
<div id="Layer42" class="style2" style="position:absolute; left:60px; top:10px; z-index:5">
<%
	Response.Write "大宗郵資已付掛號函件<br>"
	Response.Write "　　第"&Sys_MailNumber&"號"
%>
</div>

<div id="Layer49" class="style2" style="position:absolute; left:450px; top:30px; z-index:5">
<%
	Response.Write Sys_BillNo
	Response.Write "　　"&funcCheckFont(Sys_Owner,20,1)
%>
</div>

<div id="Layer43" class="style2" style="position:absolute; left:270px; top:0px; z-index:5">
<%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_2.jpg""><br>"
	Response.Write "　"&Sys_MAILCHKNUMBER
%>
</div>

<div id="Layer44" class="style2" style="position:absolute; left:450px; top:70px; z-index:5">
<%
	Response.Write "<img src=""../BarCodeImage/"&Sys_BillNo&"_4.jpg"">"
%>
</div>

<div id="Layer48" class="style2" style="position:absolute; left:100px; top:70px; z-index:5">
<%
	Response.Write Sys_BatChNumber
	If Sys_RecordMemberID = 3552 Then
		Response.Write "<br>停管入案"
	End if
%>
</div>

<div id="Layer45" class="style3" style="position:absolute; left:170px; top:120px; height:13px; z-index:14">
<%=funcCheckFont(Sys_Owner,20,1)&"<br>"&Sys_OwnerZip&"　"&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress&chkaddress,20,1)%>
</div>

<div id="Layer46" class="style2" style="position:absolute; left:170px; top:155px; width:510px; height:13px; z-index:14">
<%=Sys_BillNo%>
</div>

<div id="Layer47" class="style2" style="position:absolute; left:170px; top:175px; width:510px; height:13px; z-index:14">
<%="舉發違反道路交通管理事件通知單"%>
</div>

<div id="Layer50" class="style2" style="position:absolute; left:230px; top:770px; height:13px; z-index:14">
<%=funcCheckFont(Sys_Owner,20,1)&"<br>"&Sys_OwnerZip&"　"&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress&chkaddress,20,1)%>
</div>
</Div>