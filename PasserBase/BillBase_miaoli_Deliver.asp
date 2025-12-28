<%

strSql="select * from PasserBase where SN="&trim(BillSN(i))
set rs=conn.execute(strSql)
if Not rs.eof then
	Sys_BillTypeID=trim(rs("BillTypeID"))
	Sys_BillNo=trim(rs("BillNo"))
	Sys_DOUBLECHECKSTATUS=trim(rs("DOUBLECHECKSTATUS"))
	Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
	Sys_RuleVer=trim(rs("RuleVer"))
	Sys_Note=trim(rs("Note"))
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
	Sys_Driver=trim(rs("Driver"))
	Sys_DriverID=trim(rs("DriverID"))
	Sys_DriverHomeAddress=trim(rs("DriverAddress"))
	Sys_DriverHomeZip=trim(rs("DriverZip"))
	Sys_Rule1=trim(rs("Rule1"))
	Sys_Rule2=trim(rs("Rule2"))
	Sys_Sex=""
	if Not rs.eof then
		If not ifnull(Trim(rs("DriverID"))) Then
			If Mid(Trim(rs("DriverID")),2,1)="1" Then
				Sys_Sex="男"
			elseif Mid(Trim(rs("DriverID")),2,1)="2" Then
				Sys_Sex="女"
			End if
		End if
	end if
	Sys_RecordMemberID=trim(rs("RECORDMEMBERID"))
	Sys_IllegalDate=split(gArrDT(trim(rs("IllegalDate"))),"-")
	Sys_IllegalDate_h=hour(trim(rs("IllegalDate")))
	Sys_IllegalDate_m=minute(trim(rs("IllegalDate")))
	Sys_DealLineDate=split(gArrDT(trim(rs("DealLineDate"))),"-")
	DealLineDate=trim(rs("DealLineDate"))
	Sys_DriverBirth=split(gArrDT(trim(rs("DriverBirth"))),"-")
	Sys_BillFillerMemberID=0
	Sys_Billmem1ID=trim(rs("BILLMEMID1"))
	Sys_STATIONNAME=trim(rs("MemberStation"))
end if
rs.close

Sys_UrgeDate=""
If not ifnull(request("BillUrge")) Then
	strSQL="select OpenGovNumber,UrgeDate from PasserUrge where BillSN="&trim(BillSN(i))
	set rsjude=conn.execute(strSQL)
	If not rsjude.eof Then
		Sys_OpenGovNumber=trim(rsjude("OpenGovNumber"))
		Sys_UrgeDate=split(gArrDT(trim(rsjude("UrgeDate"))),"-")
	End if
	rsjude.close

else

	strSQL="select OpenGovNumber,JudeDate from PasserJude where BillSN="&trim(BillSN(i))
	set rsjude=conn.execute(strSQL)
	If not rsjude.eof Then
		Sys_OpenGovNumber=trim(rsjude("OpenGovNumber"))
		Sys_UrgeDate=split(gArrDT(trim(rsjude("JudeDate"))),"-")
	End if
	rsjude.close
End if

If ifnull(Sys_OpenGovNumber) Then
	Sys_OpenGovNumber=trim(Sys_BillNo)
	Sys_UrgeDate=split(gArrDT(date),"-")
End if

strUnit="select * from UnitInfo where UnitID='"&Sys_STATIONNAME&"'"
set rsUnit=conn.execute(strUnit)
If Not rsUnit.eof Then
	Sys_STATIONNAME=trim(rsUnit("UnitName"))
	If Not rsUnit.eof Then sysunit=replace(rsUnit("UnitName"),"交通組","")
	if Not rsUnit.eof then Sys_UnitAddress=trim(rsUnit("Address"))
	if Not rsUnit.eof then Sys_UnitTel=trim(rsUnit("Tel"))
End if
rsUnit.close
Sys_Level1=0:Sys_Level2=0
strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VERSION=(select value from apconfigure where ID=3)"
set rsRule1=conn.execute(strRule1)
if not rsRule1.eof then
	If DateDiff("d",CDate(date),trim(DealLineDate))>-1 Then 
	  Sys_Level1=trim(rsRule1("Level1"))
	Else
	  Sys_Level1=trim(rsRule1("Level2"))
	End if
end if
rsRule1.close
set rsRule1=nothing

If Not ifnull(Sys_Rule2) Then
	strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and VERSION=(select value from apconfigure where ID=3)"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		If DateDiff("d",CDate(date),trim(DealLineDate))>-1 Then 
		  Sys_Level2=trim(rsRule1("Level1"))
		Else
		  Sys_Level2=trim(rsRule1("Level2"))
		End if
	end if
	rsRule1.close
	set rsRule1=nothing
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName"))
rszip.close

Sum_Level=cdbl(Sys_Level1)+cdbl(Sys_Level2)

strSql="select a.LoginID,a.ChName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&trim(Sys_RecordMemberID)
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close


'If not ifnull(Sys_Billmem1ID) Then
	'strSql="select a.LoginID,a.ChName,b.UnitName,a.ImageFilename as MemberFileName,b.UnitName,b.UnitID,b.UnitLevelID,b.UnitTypeID,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_Billmem1ID
'	set mem=conn.execute(strsql)
'	if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
'	if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
'	if Not mem.eof then Sys_BillUnitAddress=trim(mem("Address"))
'	if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
'	if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
'	if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
	'if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
	'if Not mem.eof then Sys_ChName=trim(mem("ChName"))
	'if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
'	mem.close
'End if

'If Sys_UnitLevelID=1 Then
'	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
'else
	'strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
'end if
'set unit=conn.Execute(strSQL)
'If Not unit.eof Then sysunit=replace(unit("UnitName"),"交通組","")
'if Not unit.eof then Sys_UnitAddress=trim(unit("Address"))
'if Not unit.eof then Sys_UnitTel=trim(unit("Tel"))
'unit.close

if trim(Sys_Rule1)<>"0" and not isnull(Sys_Rule1) then
	strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
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
if Not rs.eof then Sys_StationTel=trim(rs("StationTel"))
if Not rs.eof then Sys_StationID=trim(rs("StationID"))
rs.close

strSql="select MailNumber,StoreAndSendMailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&trim(BillSN(i))

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))

rs.close
if isnull(Sys_DriverHomeZip) or trim(Sys_DriverHomeZip)="" then Sys_DriverHomeZip="001"
if isnull(Sys_OwnerZip) or trim(Sys_OwnerZip)="" then Sys_OwnerZip="001"

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
Sys_MailNumber=0
Sys_BillNo_BarCode=Sys_OpenGovNumber
If sys_City="高雄市" or sys_City="苗栗縣" Then
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160,1

elseIf sys_City="台中縣" Then
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160

else
	DelphiASPObj.GenSendStoreBillno Sys_OpenGovNumber,0,57,160

end if

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&trim(BillSN(i))

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
%>
<div id="L78" style="position:relative;">
<div id="Layer44" class="style2" style="position:absolute; left:450px; top:50px; z-index:5">
	<img src=<%="""../BarCodeImage/"&Sys_BillNo_BarCode&".jpg"""%>>
</div>

<div id="Layer45" class="style2" style="position:absolute; left:170px; top:100px; width:550px; height:13px; z-index:14">
<%
if Sys_BillTypeID="1" and trim(Sys_DriverHomeAddress)<>"" then
	Sys_DriverHomeZip=replace(Sys_DriverHomeZip,"001","")
	response.write Sys_Driver&"　"
	response.write Sys_DriverHomeZip&Sys_DriverHomeAddress
else
	response.write Sys_Owner&"　"
	response.write Sys_OwnerAddress
end if
%>
</div>

<div id="Layer51" class="style2" style="position:absolute; left:500px; top:0px; width:510px; height:13px; z-index:14">
違警
</div>

<div id="Layer46" class="style2" style="position:absolute; left:170px; top:130px; width:510px; height:13px; z-index:14">
<%=BillPageUnit%><%=UrgeNo%><%=Sys_BillNo%>號
</div>

<div id="Layer47" class="style2" style="position:absolute; left:170px; top:150px; width:510px; height:13px; z-index:14">
違反道路交通管理事件<%=Papertype%>
</div>

<div id="Layer50" class="style2" style="position:absolute; left:230px; top:755px; width:550px; height:13px; z-index:14">
<%
if Sys_BillTypeID="1" and trim(Sys_DriverHomeAddress)<>"" then
	Sys_DriverHomeZip=replace(Sys_DriverHomeZip,"001","")
	response.write Sys_Driver&"　"
	response.write "<br>"&Sys_DriverHomeZip&Sys_DriverHomeAddress
else
	response.write Sys_Owner&"　"
	response.write "<br>"&Sys_OwnerAddress
end if
%>
</div>
</Div>