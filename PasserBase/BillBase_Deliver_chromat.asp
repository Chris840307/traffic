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

If DeliverKind=1 Then
tableName="PasserUrge"
else
tableName="PasserJude"
End if

strSQL="select OpenGovNumber from "&tableName&" where BillSN="&trim(BillSN(i))
set rsjude=conn.execute(strSQL)
If not rsjude.eof Then
	Sys_OpenGovNumber=trim(rsjude("OpenGovNumber"))
End if
rsjude.close

strUnit="select UnitName from UnitInfo where UnitID='"&Sys_STATIONNAME&"'"
set rsUnit=conn.execute(strUnit)
If Not rsUnit.eof Then
	Sys_STATIONNAME=trim(rsUnit("UnitName"))
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
Sys_BillNo_BarCode=Sys_BillNo
If sys_City<>"台中縣" Then
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160
else
	Sys_BillNo_BarCode=Sys_BillNo_BarCode&"_4"
	if trim(Sys_BillTypeID)="1" then
		DelphiASPObj.GenBillPrintBarCode BillSN(i+PrintSum),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,date
		'response.write "DelphiASPObj.GenBillPrintBarCode"& BillSN(i+PrintSum)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
		'response.end
	else
		DelphiASPObj.GenBillPrintBarCode BillSN(i+PrintSum),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,date
		'response.write "DelphiASPObj.GenBillPrintBarCode "& BillSN(i+PrintSum)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_OwnerZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
		'response.end
	end if
end if

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&trim(BillSN(i))

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
%>
<div id="Layer1" style="position:absolute; left:300px; top:110px; z-index:5"><%
	if Sys_BillTypeID="1" then
		If trim(Sys_Driver)<>"" Then response.write Sys_Driver&"　"
		If trim(Sys_DriverHomeAddress)<>"" Then response.write "<br>"&Sys_DriverHomeZip&Sys_DriverHomeAddress
	else
		If trim(Sys_Owner)<>"" Then response.write Sys_Owner&"　"
		If trim(Sys_OwnerAddress)<>"" Then response.write "<br>"&Sys_OwnerAddress
	end if%>&nbsp;</span>
</div>
<%
If DeliverKind=1 Then%>
	<div id="Layer2" style="position:absolute; left:490px; top:160px; z-index:5"><%=Sys_OpenGovNumber%></div><%
else%>
	<div id="Layer2" style="position:absolute; left:490px; top:195px; z-index:5"><%=Sys_OpenGovNumber%></div><%
End if
If DeliverKind=1 Then%>
	<div id="Layer3" style="position:absolute; left:455px; top:220px; z-index:5">1</div><%
else%>
	<div id="Layer3" style="position:absolute; left:420px; top:245px; z-index:5">1</div><%
End if%>