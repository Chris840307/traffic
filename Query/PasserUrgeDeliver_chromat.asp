<%
strState="select * from PasserUrge where BillSN="&BillSN(i)
set rsState=conn.execute(strState)
if not rsState.eof then
	UOpenGovNumber=trim(rsState("OpenGovNumber"))
	UrgeDate=split(gArrDT(rsState("UrgeDate")),"-")
end if
rsState.close
set rsState=nothing
PrintDate=split(gArrDT(date),"-")
strSql="select * from PasserBase where SN="&BillSN(i)
set rsSql=conn.execute(strSql)
if rsSql.eof then response.end
UAgentSex=trim(rsSql("DriverSex"))
strUInfo="select * from UnitInfo where UnitID='"&trim(rsSql("BillUnitID"))&"'"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	theSubUnitSecBossName=trim(rsUInfo("SecondManagerName"))
	theBigUnitBossName=trim(rsUInfo("ManageMemberName"))
	theContactTel=trim(rsUInfo("Tel"))
	theBankAccount=trim(rsUInfo("BankAccount"))
	thenBillUnitName=trim(rsUInfo("UnitName"))
end if
rsUInfo.close
set rsUInfo=nothing

strSql="select b.Content from BILLFASTENERDETAIL a,DCICode b where a.FASTENERTYPEID=b.ID and b.TypeID=6 and a.BIllSN="&trim(rsSql("SN"))
set rsfast=conn.execute(strsql)
fastring=""
while Not rsfast.eof
	if trim(fastring)<>"" then fastring=fastring&","
	fastring=fastring&rsfast("Content")
	rsfast.movenext
wend
rsfast.close

thenPasserUnit=""
strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsunit=conn.execute(strSQL)
If Not rsunit.eof Then
	Sys_UnitID=trim(rsunit("UnitID"))
	Sys_UnitLevelID=trim(rsunit("UnitLevelID"))
	Sys_UnitTypeID=trim(rsunit("UnitTypeID"))
End if
rsunit.close

If Sys_UnitLevelID=1 Then
	strSQL="select UnitName,UnitID,Tel,BANKACCOUNT from UnitInfo where UnitName like '%交通隊' and (UnitTypeID='"&Sys_UnitID&"' or UnitTYpeID='"&Sys_UnitTypeID&"' or UnitID='"&Sys_UnitID&"'or UnitID='"&Sys_UnitTypeID&"')"
else
	strSQL="select UnitName,UnitID,Tel,BANKACCOUNT from UnitInfo where UnitName like '%分局' and (UnitTypeID='"&Sys_UnitID&"' or UnitTYpeID='"&Sys_UnitTypeID&"' or UnitID='"&Sys_UnitID&"'or UnitID='"&Sys_UnitTypeID&"')"
end if
set rsunit=conn.Execute(strSQL)
Sys_UnitID=trim(rsunit("UnitID"))
Sys_Tel=trim(rsunit("Tel"))
Sys_BANKACCOUNT=trim(rsunit("BANKACCOUNT"))
if Not rsunit.eof then thenPasserUnit=trim(rsunit("UnitName"))
rsunit.close
strMem="select MANAGEMEMBERNAME from UnitInfo where UnitID='"&Sys_UnitID&"'"
set rsUnit=conn.execute(strMem)
if Not rsUnit.eof then MemUnitName=rsUnit("MANAGEMEMBERNAME")
rsUnit.close
Sys_now=split(gArrDT(date),"-")
%>
<div id="Layer1" class="style25" style="position:absolute; left:385px;; top:65px; z-index:5"><%=replace(thenPasserUnit,"分局","")%></div>

<div id="Layer2" class="style22" style="position:absolute; left:135px;; top:175px; z-index:5"><%=UrgeDate(0)&"&nbsp;&nbsp;&nbsp;"&UrgeDate(1)&"&nbsp;&nbsp;"&UrgeDate(2)%></div>

<div id="Layer3" class="style22" style="position:absolute; left:470px;; top:175px; z-index:5"><%=UOpenGovNumber%></div>

<div id="Layer4" class="style22" style="position:absolute; left:145px;; top:245px; z-index:5"><%=trim(rsSql("DRIVER"))%></div>

<div id="Layer5" class="style22" style="position:absolute; left:145px;; top:305px; z-index:5"><%=trim(rsSql("DriverAddress"))%></div>

<div id="Layer6" class="style22" style="position:absolute; left:370px;; top:520px; z-index:5"><%=trim(Sys_Tel)%></div>

<div id="Layer7" class="style22" style="position:absolute; left:130px;; top:575px; z-index:5"><%=trim(Sys_BANKACCOUNT)%></div>

<div id="Layer8" class="style25" style="position:absolute; left:195px;; top:785px; z-index:5"><%=trim(rsSql("DRIVER"))%></div>

<div id="Layer8" class="style25" style="position:absolute; left:410px;; top:1020px; z-index:5"><%=Sys_now(0)&"　　　"&Sys_now(1)&"　　　"&Sys_now(2)%></div>