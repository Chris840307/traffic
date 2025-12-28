<%
strState="select * from PasserJude where BillSN="&BillSN(i)
set rsState=conn.execute(strState)
if not rsState.eof then
	UOpenGovNumber=trim(rsState("OpenGovNumber"))
	UJudeDate=split(gArrDT(rsState("JudeDate")),"-")
	UAgentName=trim(rsState("AgentName"))
	UAgentBirth=trim(rsState("AgentBirth"))
	UAgentID=trim(rsState("AgentID"))
	UAgentAddress=trim(rsState("AgentAddress"))
	UBigUnitBossName=trim(rsState("BigUnitBossName"))
	USubUnitSecBossName=trim(rsState("SubUnitSecBossName"))
	UContactTel=trim(rsState("ContactTel"))
	UForFeit=trim(rsState("ForFeit"))
	UDutyUnit=trim(rsState("DutyUnit"))
	USendAddress=trim(rsState("SendAddress"))
	UPunishmentMainBody=trim(rsState("PunishmentMainBody"))
	USimpleReson=trim(rsState("SimpleReson"))
	UNote=trim(rsState("Note"))
	strUInfo="select * from UnitInfo where UnitID='"&trim(rsState("DutyUnit"))&"'"
	set rsUInfo=conn.execute(strUInfo)
	if not rsUInfo.eof then
		DutyUnitName=trim(rsUInfo("UnitName"))
		DutyAddress=trim(rsUInfo("Address"))
	end if
	rsUInfo.close
	set rsUInfo=nothing
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
			UAgentSex="¨k"
		elseif Mid(Trim(rsSql("DriverID")),2,1)="2" Then
			UAgentSex="¤k"
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
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
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
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if

set rsunit=conn.Execute(strSQL)
Sys_UnitID=trim(rsunit("UnitID"))
if Not rsunit.eof then thenPasserUnit=trim(rsunit("UnitName"))
rsunit.close
strMem="select chName from Memberdata where MemberID="&session("User_ID")
set rsUnit=conn.execute(strMem)
if Not rsUnit.eof then MemUnitName=rsUnit("chName")
rsUnit.close
If not ifnull(Session("Ch_Name")) Then MemUnitName=Session("Ch_Name")
Sys_now=split(gArrDT(date),"-")
%>
<div id="Layer20" class="style25" style="position:absolute; left:390px; top:60px; z-index:1"><%=left(thenPasserUnit,2)%></div>

<div id="Layer1" class="style22" style="position:absolute; left:140px;; top:135px; z-index:5"><%=Sys_now(0)&"&nbsp;&nbsp;&nbsp;&nbsp;"&Sys_now(1)&"&nbsp;&nbsp;&nbsp;&nbsp;"&Sys_now(2)%>
</div>

<div id="Layer20" class="style22" style="position:absolute; left:415px; top:135px; z-index:1"><%=left(thenPasserUnit,1)%></div>

<div id="Layer2" class="style22" style="position:absolute; left:535px; top:135px; z-index:5"><%=UOpenGovNumber%></div>

<div id="Layer3" class="style22" style="position:absolute; left:145px;; top:220px; z-index:5"><%=trim(rsSql("DRIVER"))%></div>

<div id="Layer4" class="style22" style="position:absolute; left:325px; top:220px; z-index:5"><%=gInitDT(trim(rsSql("DriverBirth")))%></div>

<div id="Layer5" class="style22" style="position:absolute; left:445px; top:220px; z-index:5"><%=UAgentSex%></div>

<div id="Layer6" class="style22" style="position:absolute; left:500px; top:220px; z-index:5"><%=trim(rsSql("DriverID"))%></div>

<div id="Layer7" class="style22" style="position:absolute; left:660px; top:200px; z-index:5"><%=theBigUnitBossName%></div>

<div id="Layer8" class="style22" style="position:absolute; left:660px;; top:365px; z-index:5"><%=theSubUnitSecBossName%></div>

<div id="Layer9" class="style22" style="position:absolute; left:660px;; top:530px; z-index:5"><%=MemUnitName%></div>

<div id="Layer10" class="style22" style="position:absolute; left:145px;; top:260px; z-index:5"><%=trim(rsSql("DriverAddress"))%></div>

<div id="Layer11" class="style22" style="position:absolute; left:145px;; top:385px; z-index:5"><%=trim(rsSql("CarNo"))%></div>

<div id="Layer12" class="style22" style="position:absolute; left:530px; top:385px; z-index:5"><%=trim(rsSql("BillNo"))%></div>

<div id="Layer13" class="style22" style="position:absolute; left:140px;; top:420px; z-index:5"><%
	if trim(rsSql("IllegalDate"))<>"" then
		IllegalDate=split(gArrDT(rsSql("IllegalDate")),"-")
		response.write IllegalDate(0)&"&nbsp;&nbsp;&nbsp;"&IllegalDate(1)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&IllegalDate(2)
	end if
%></div>

<div id="Layer14" class="style22" style="position:absolute; left:430px; top:420px; z-index:5"><%
	if trim(rsSql("DealLineDate"))<>"" then
		DealLineDate=split(gArrDT(rsSql("DealLineDate")),"-")
		response.write DealLineDate(0)&"&nbsp;&nbsp;&nbsp;&nbsp;"&DealLineDate(1)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&DealLineDate(2)
	end if
%></div>

<div id="Layer15" class="style22" style="position:absolute; left:145px;; top:455px; z-index:5"><%=trim(rsSql("IllegalAddress"))%></div>

<div id="Layer16" class="style22" style="position:absolute; left:145px;; top:490px; z-index:5; width:400px"><%
	if trim(rsSql("Rule1"))<>"" and not isnull(rsSql("Rule1")) then
		strRule1="select * from Law where ItemID='"&trim(rsSql("Rule1"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
		set rsRule1=conn.execute(strRule1)
		if not rsRule1.eof then
			response.write trim(rsRule1("IllegalRule"))
		end if
		rsRule1.close
		set rsRule1=nothing
	end if	
	if trim(rsSql("Rule2"))<>"" and not isnull(rsSql("Rule2")) then
		strRule1="select * from Law where ItemID='"&trim(rsSql("Rule2"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
		set rsRule1=conn.execute(strRule1)
		if not rsRule1.eof then
			response.write "<br>"&trim(rsRule1("IllegalRule"))
		end if
		rsRule1.close
		set rsRule1=nothing
	end if
%></div>

<div id="Layer17" class="style22" style="position:absolute; left:145px; width:500px; top:530px; z-index:5"><%=UPunishmentMainBody%></div>

<div id="Layer18" class="style22" style="position:absolute; left:135px;; top:565px; z-index:5"><%=UJudeDate(0)&"&nbsp;&nbsp;&nbsp;&nbsp;"&UJudeDate(1)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&UJudeDate(2)%></div>

<div id="Layer19" class="style22" style="position:absolute; left:435px; top:565px; z-index:5"><%=DutyUnitName%></div>

<div id="Layer20" class="style22" style="position:absolute; left:610px; top:630px; z-index:5"><%
	Sys_Rule1=trim(rsSql("Rule1"))
	response.write left(trim(Sys_Rule1),2)&"¡@¡@¡@¡@"
	if len(trim(Sys_Rule1))>7 then response.write "¡@¡@"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)
%></div>
<div id="Layer21" class="style22" style="position:absolute; left:150px; top:655px; z-index:5"><%
	Sys_Rule1=trim(rsSql("Rule1"))
	response.write Mid(trim(Sys_Rule1),4,2)
%></div>


<div id="Layer24" class="style6" style="position:absolute; left:280px; top:760px; z-index:1"><%=left(thenPasserUnit,2)%></div>

<div id="Layer25" class="style6" style="position:absolute; left:590px; top:760px; z-index:1"><%=left(thenPasserUnit,2)%></div>

<div id="Layer22" class="style22" style="position:absolute; left:185px;; top:940px; z-index:5"><%=trim(rsSql("DRIVER"))%></div>

<div id="Layer23" class="style22" style="position:absolute; left:185px;; top:965px; z-index:5"><%=trim(rsSql("DriverAddress"))%></div>


