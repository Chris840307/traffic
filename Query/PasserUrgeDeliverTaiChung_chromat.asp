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
if Not rsSql.eof then
	If not ifnull(Trim(rsSql("DriverID"))) Then
		If Mid(Trim(rsSql("DriverID")),2,1)="1" Then
			UAgentSex="男"
		elseif Mid(Trim(rsSql("DriverID")),2,1)="2" Then
			UAgentSex="女"
		End if
	End if
end if

strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
Sys_UnitID=trim(rsUnit("UnitID"))
Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
rsUnit.close

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where (UnitName like '%通隊' or UnitName like '%安隊') and (UnitTypeID='"&Sys_UnitID&"' or UnitTYpeID='"&Sys_UnitTypeID&"' or UnitID='"&Sys_UnitID&"'or UnitID='"&Sys_UnitTypeID&"')"
elseif sys_City<>"金門" and sys_City<>"連江" then
	strSQL="select * from UnitInfo where UnitName like '%分局' and (UnitTypeID='"&Sys_UnitID&"' or UnitTYpeID='"&Sys_UnitTypeID&"' or UnitID='"&Sys_UnitID&"'or UnitID='"&Sys_UnitTypeID&"')"
else
	strSQL="select * from UnitInfo where UnitName like '%所' and (UnitTypeID='"&Sys_UnitID&"' or UnitTYpeID='"&Sys_UnitTypeID&"' or UnitID='"&Sys_UnitID&"'or UnitID='"&Sys_UnitTypeID&"')"
end if
set rsUnit=conn.Execute(strSQL)
DB_UnitID=trim(rsUnit("UnitID"))
theSubUnitSecBossName=trim(rsUnit("SecondManagerName"))
theBigUnitBossName=trim(rsUnit("ManageMemberName"))
theContactTel=trim(rsUnit("Tel"))
theBankAccount=trim(rsUnit("BankAccount"))
thenBillUnitName=trim(rsUnit("UnitName"))
rsUnit.close

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
	strSQL="select * from UnitInfo where (UnitName like '%通隊' or UnitName like '%安隊') and (UnitTypeID='"&Sys_UnitID&"' or UnitTYpeID='"&Sys_UnitTypeID&"' or UnitID='"&Sys_UnitID&"'or UnitID='"&Sys_UnitTypeID&"')"
elseif sys_City<>"金門" and sys_City<>"連江" then
	strSQL="select * from UnitInfo where UnitName like '%分局' and (UnitTypeID='"&Sys_UnitID&"' or UnitTYpeID='"&Sys_UnitTypeID&"' or UnitID='"&Sys_UnitID&"'or UnitID='"&Sys_UnitTypeID&"')"
else
	strSQL="select * from UnitInfo where UnitName like '%所' and (UnitTypeID='"&Sys_UnitID&"' or UnitTYpeID='"&Sys_UnitTypeID&"' or UnitID='"&Sys_UnitID&"'or UnitID='"&Sys_UnitTypeID&"')"
end if
set rsunit=conn.Execute(strSQL)
Sys_UnitID=trim(rsunit("UnitID"))
Sys_Tel=trim(rsunit("Tel"))
Sys_BANKACCOUNT=trim(rsunit("BANKACCOUNT"))
if Not rsunit.eof then thenPasserUnit=trim(rsunit("UnitName"))
rsunit.close
strMem="select chName from Memberdata where MemberID="&session("User_ID")
set rsUnit=conn.execute(strMem)
if Not rsUnit.eof then MemUnitName=rsUnit("chName")
rsUnit.close
If not ifnull(Session("Ch_Name")) Then MemUnitName=Session("Ch_Name")
Sys_now=split(gArrDT(date),"-")
%>
<div id="Layer1" class="style25" style="position:absolute; left:395px;; top:65px; z-index:5"><%=replace(thenPasserUnit,"分局","")%></div>

<div id="Layer2" class="style22" style="position:absolute; left:135px;; top:180px; z-index:5"><%=UrgeDate(0)&"&nbsp;&nbsp;"&UrgeDate(1)&"&nbsp;&nbsp;"&UrgeDate(2)%></div>

<div id="Layer3" class="style22" style="position:absolute; left:460px;; top:180px; z-index:5"><%=UOpenGovNumber%></div>

<div id="Layer4" class="style22" style="position:absolute; left:135px;; top:255px; z-index:5"><%=trim(rsSql("DRIVER"))%></div>

<div id="Layer5" class="style22" style="position:absolute; left:135px;; top:325px; z-index:5"><%=trim(rsSql("DriverAddress"))%></div>

<div id="Layer6" class="style22" style="position:absolute; left:410px;; top:380px; z-index:5">1</div>

<div id="Layer7" class="style25" style="position:absolute; left:600px; top:305px; z-index:5"><%=theBigUnitBossName%></div>

<div id="Layer8" class="style25" style="position:absolute; left:600px;; top:590px; z-index:5"><%=theSubUnitSecBossName%></div>

<div id="Layer9" class="style25" style="position:absolute; left:600px;; top:845px; z-index:5"><%=MemUnitName%></div>

<div id="Layer10" class="style22" style="position:absolute; left:360px;; top:530px; z-index:5"><%=trim(Sys_Tel)%></div>

<div id="Layer11" class="style22" style="position:absolute; left:120px;; top:580px; z-index:5"><%=trim(theBankAccount)%></div>

<div id="Layer12" class="style25" style="position:absolute; left:460px;; top:575px; z-index:5"><%
If instr(Sys_GroupUnitName,"組")>0 Then
	response.write replace(right(trim(Sys_GroupUnitName),2),"組","")
else
	response.write Session("Sys_UnitGroup")
end if
%></div>

<div id="Layer13" class="style25" style="position:absolute; left:180px;; top:810px; z-index:5"><%=trim(rsSql("DRIVER"))%></div>

<div id="Layer14" class="style25" style="position:absolute; left:400px;; top:1030px; z-index:5"><%=Sys_now(0)&"　　　"&Sys_now(1)&"　　　"&Sys_now(2)%></div>