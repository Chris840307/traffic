<%
strSQL="select * from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
DB_ManageMemberName=trim(rsUnit("ManageMemberName"))
rsUnit.close

strSql="select a.SN as BillSN,a.BillNo,a.Driver,a.DriverBirth,a.DriverID,a.DriverAddress,a.IllegalDate,a.IllegalAddress,a.DealLineDate,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillUnitID,b.OpenGovNumber as JudeOGN,b.AgentName as JudeAgentName,b.AgentSex as JudeAgentSex,b.AgentBirth as JudeAgentBirth,b.AgentID as JudeAgentID,b.AgentAddress as JudeAgentAddress,c.OpenGovNumber as UrgeOGN,c.UrgeTypeID,d.OpenGovNumber,d.BigUnitBossName,d.SubUnitSecBossName,d.SendNumber,d.SendDate,d.Agent,d.AgentBirthDate,d.AgentID,d.AgentAddress,d.ForFeit,d.MakeSureDate,d.LimitDate,d.AttatchJude,d.AttatchUrge,d.AttatchFortune,d.AttatchGround,d.AttatchRegister,d.AttatchFileList,d.AttatchTable,d.ATTATPOSTAGE,d.SafeToExit,d.SAFEACTION,d.SAFEASSURE,d.SAFEDETAIN,d.SAFESHUTSHOP,e.ArrivedDate from PasserBase a,PasserJude b,PasserUrge c,PasserSend d,(select * from PasserSendArrived where ArriveType=0) e where a.SN="&trim(BillSN(i))&" and a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and a.SN=c.BillSN(+) and a.BillNo=c.BillNo(+) and a.SN=d.BillSN(+) and a.BillNo=d.BillNo(+) and a.SN=e.PasserSN(+)"
PrintDate=split(gArrDT(date),"-")
set rsfound=conn.execute(strSql)

If Not ifnull(rsfound("MakeSureDate")) Then
	MakeSureDate=split(gArrDT(rsfound("MakeSureDate")),"-")
	LimitDate=split(gArrDT(rsfound("LimitDate")),"-")
	
elseIf Not ifnull(rsfound("ArrivedDate")) Then
	MakeSureDate=split(gArrDT(DateAdd("d",20,rsfound("ArrivedDate"))),"-")
	LimitDate=split(gArrDT(DateAdd("d",35,rsfound("ArrivedDate"))),"-")	
End if
%>
<div id="Layer1" class="style6" style="position:absolute; left:550px; top:40px; z-index:1"><%=rsfound("SendNumber")%></div>

<div id="Layer2" class="style6" style="position:absolute; left:485px; top:105px; z-index:1"><%=PrintDate(0)&"　　"&PrintDate(1)&"　&nbsp;"&PrintDate(2)%></div>

<div id="Layer3" class="style6" style="position:absolute; left:595px; top:130px; z-index:1"><%=rsfound("OpenGovNumber")%></div>

<div id="Layer4" class="style6" style="position:absolute; left:160px; top:185px; z-index:1"><%=rsfound("Driver")%></div>

<div id="Layer5" class="style6" style="position:absolute; left:160px; top:215px; z-index:1"><%
	if trim(rsfound("DriverBirth"))<>"" then
		DriverBirth=split(gArrDT(rsfound("DriverBirth")),"-")
		response.write DriverBirth(0)&"年"&DriverBirth(1)&"月"&DriverBirth(2)&"日"
	end if
%></div>

<div id="Layer6" class="style6" style="position:absolute; left:160px; top:240px; z-index:1"> <%
	if Not rsfound.eof then
		If not ifnull(Trim(rsfound("DriverID"))) Then
			If Mid(Trim(rsfound("DriverID")),2,1)="1" Then
				Response.write "男"
			elseif Mid(Trim(rsfound("DriverID")),2,1)="2" Then
				Response.write "女"
			End if
		End if
	end if%></div>

<div id="Layer7" class="style6" style="position:absolute; left:160px; top:305px; z-index:1"><%=rsfound("DriverID")%></div>

<div id="Layer8" class="style6" style="position:absolute; left:165px; top:375px; z-index:1"><%=InstrAdd(rsfound("DriverAddress"),20)%></div>

<div id="Layer9" class="style6" style="position:absolute; left:160px; top:480px; z-index:1"><%'=rsfound("IllegalAddress")%></div>

<div id="Layer10" class="style6" style="position:absolute; left:550px; top:465px; z-index:1"><%
	'if trim(rsfound("SendDate"))<>"" then
		'SendDate=split(gArrDT(rsfound("SendDate")),"-")
		'response.write SendDate(0)&"年"&SendDate(1)&"月"&SendDate(2)&"日"
	'end if
%></div>

<div id="Layer11" class="style6" style="position:absolute; left:540px; top:500px; z-index:1"><%=MakeSureDate(0)&"　　"&MakeSureDate(1)&"　　"&MakeSureDate(2)%></div>

<div id="Layer12" class="style6" style="position:absolute; left:530px; top:555px; z-index:1"><%=LimitDate(0)&"　"&LimitDate(1)&"　　"&LimitDate(2)%></div>

<div id="Layer13" class="style6" style="position:absolute; left:580px; top:630px; z-index:1"><%=rsfound("ForFeit")%></div>

<div id="Layer14" class="style6" style="position:absolute; left:370px; top:570px; z-index:1"><%=left(trim(rsfound("Rule1")),2)%></div>

<div id="Layer15" class="style6" style="position:absolute; left:200px; top:600px; z-index:1"><%=Mid(trim(rsfound("Rule1")),3,1)%></div>

<div id="Layer16" class="style6" style="position:absolute; left:265px; top:600px; z-index:1"><%=Mid(trim(rsfound("Rule1")),4,2)%></div>

<div id="Layer17" class="style6" style="position:absolute; left:300px; top:630px; z-index:1"><%
	if trim(rsfound("SendDate"))<>"" then
		IllegalDate=split(gArrDT(rsfound("IllegalDate")),"-")
		response.write IllegalDate(0)&"&nbsp;&nbsp;"&IllegalDate(1)&"&nbsp;&nbsp;"&IllegalDate(2)
	end if
%></div>

<div id="Layer18" class="style7" style="position:absolute; left:610px; top:1005px; z-index:1"><%=theBigUnitBossName%></div>

<div id="Layer19" class="style7" style="position:absolute; left:510px; top:1030px; z-index:1"><%=theSubUnitSecBossName%></div>

<div id="Layer20" class="style7" style="position:absolute; left:140px; top:720px; z-index:1">V</div>
<div id="Layer21" class="style7" style="position:absolute; left:140px; top:760px; z-index:1">V</div>

<div id="Layer22" class="style7" style="position:absolute; left:510px; top:720px; z-index:1">V</div>

<div id="Layer23" class="style7" style="position:absolute; left:510px; top:785px; z-index:1">V</div>
<div id="Layer24" class="style7" style="position:absolute; left:510px; top:800px; z-index:1">V</div>

<div id="Layer25" class="style7" style="position:absolute; left:420px; top:825px; z-index:1">V</div>
<div id="Layer26" class="style7" style="position:absolute; left:140px; top:845px; z-index:1">V</div>

<div id="Layer27" class="style7" style="position:absolute; left:420px; top:860px; z-index:1">V</div>
<div id="Layer28" class="style7" style="position:absolute; left:140px; top:885px; z-index:1">V</div>