<%
strSQL="select * from UnitInfo where UnitID='"&DB_UnitID&"'"
set rsUnit=conn.execute(strSQL)
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
rsUnit.close

strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_OwnerZip=""
Sys_OwnerZipName=""

If Not rsbil.eof Then
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
Sys_OwnerAddress=""

if Instr(request("Sys_BatchNumber"),"N")>0 and Sys_BillTypeID=2 Then
	strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='N'"
else
	strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
end if

set rsfound=conn.execute(strSql)

if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
if Instr(request("Sys_BatchNumber"),"N")>0 then

	if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
	if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

else
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	else
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	End if
end if

If ifnull(Sys_OwnerAddress) or ifnull(Sys_Owner) Then
	strSql="select * from BillbaseDCIReturn where Carno=(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A') and ExchangetypeID='A'"

	set rsdata=conn.execute(strsql)

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		If Sys_BillTypeID=2 Then
			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))
		end if
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
rsfound.close

Sys_Owner=""
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
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

		If ifnull(Sys_OwnerAddress) Then
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
		end if
	else
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	End if
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=replace(trim(rszip("ZipName")),"台","臺")
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

strSql="select MailNumber,StoreAndSendMailNumber,MailTypeID,MailDate,UserMarkMemberID from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))
if Not rs.eof then Sys_UserMarkMemberID=trim(rs("UserMarkMemberID"))

rs.close

Sys_UserMarName=""
If not ifnull(Sys_UserMarkMemberID) Then
	strSQL="select loginid from memberdata where memberid="&Sys_UserMarkMemberID
	set Umark=conn.execute(strSQL)
	Sys_UserMarName=Umark("LoginID")
	Umark.close
End if

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
<div id="Layer1" style="position:absolute; left:450px; top:<%=0%>px; width:150px; height:36px; z-index:5"><img src=<%="""../BarCodeImage/"&Sys_BillNo&".jpg"""%>>
</div>

<div id="Layer2" class="style2" style="position:absolute; left:240px; top:<%=50%>px; width:70px; height:36px; z-index:5"><%
reUnitName=""
for m=1 to len(sysunit)
	if Mid(sysunit,m,1)<>"分" and Mid(sysunit,m,1)<>"局" then
		reUnitName=trim(reUnitName)&Mid(sysunit,m,1)
	end if
next
response.write reUnitName%>
</div>

<div id="Layer3" class="style4" style="position:absolute; left:390px; top:<%=50%>px; width:350px; height:36px; z-index:5"><%=Sys_UnitAddress&"　　"&Sys_UserMarName%>
</div>

<div id="Layer4" class="style4" style="position:absolute; left:95px; top:<%=75%>px; width:320px; height:36px; z-index:5"><%
	response.write funcCheckFont(Sys_Owner,16,4)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;車號："
	if trim(Sys_CarNo)<>"" and not isnull(Sys_CarNo) then
		response.write left(Sys_CarNo,4)
		response.write left("*************",len(Sys_CarNo)-4)
	end if 
	response.write "<br>"&Sys_OwnerZip&"&nbsp;&nbsp;"&Sys_OwnerZipName&chstr(funcCheckFont(replace(replace(Sys_OwnerAddress,"台","臺"),Sys_OwnerZipName,""),16,4))%>
</div>

<div id="Layer5" class="style4" style="position:absolute; left:200px; top:<%=135%>px; width:150px; height:36px; z-index:5"><%=Sys_BillNo%>
</div>
<div id="Layer6" class="style4" style="position:absolute; left:225px; top:<%=150%>px; width:60px; height:36px; z-index:5"><%
	response.write left(trim(Sys_Rule1),2)'條
	if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)%>
</div>
<div id="Layer7" class="style4" style="position:absolute; left:275px; top:<%=150%>px; width:20px; height:36px; z-index:5"><%response.write Mid(trim(Sys_Rule1),3,1)'項%>
</div>
<div id="Layer8" class="style4" style="position:absolute; left:320px; top:<%=150%>px; width:20px; height:36px; z-index:5"><%response.write Mid(trim(Sys_Rule1),4,2)'款%>
</div>
<div id="Layer9" class="style4" style="position:absolute; left:360px; top:<%=150%>px; width:20px; height:36px; z-index:5"><%response.write Mid(trim(Sys_Rule2),6,2)%>
</div><%
if trim(Sys_Rule2)<>"0" then
	response.write "<div id=""Layer10"" class=""style4"" style=""position:absolute; left:225px; top:"& 170 &"px; width:60px; height:36px; z-index:5"">"&left(trim(Sys_Rule2),2)'條
	if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
	response.write "a</div>"

	response.write "<div id=""Layer11"" class=""style4"" style=""position:absolute; left:275px; top:"& 170 &"px; width:20px; height:36px; z-index:5"">"&Mid(trim(Sys_Rule2),3,1)&"b</div>"'項"

	response.write "<div id=""Layer12"" class=""style4"" style=""position:absolute; left:320px; top:"& 170 &"px; width:20px; height:36px; z-index:5"">"&Mid(trim(Sys_Rule2),4,2)&"c</div>"'款"

	response.write "<div id=""Layer13"" class=""style4"" style=""position:absolute; left:360px; top:"& 170 &"px; width:20px; height:36px; z-index:5"">"&Mid(trim(Sys_Rule2),6,2)&"d</div>"
end if%>
<div id="Layer14" class="style2" style="position:absolute; left:20px; top:<%=275%>px; width:20px; height:36px; z-index:5"><%
reUnitName=""
for m=len(Sys_UnitName) to 1 step -1
	if Mid(Sys_UnitName,m,1)<>"分" and Mid(Sys_UnitName,m,1)<>"局" then
		reUnitName=trim(reUnitName)&Mid(Sys_UnitName,m,1)
	end if
next
response.write Sys_UnitName%>
</div></Div>