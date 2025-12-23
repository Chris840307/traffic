<style type="text/css">
<!--
td {font-family:標楷體;line-height:13px;font-size:9pt;}-->
</style>
<%
strSql="select LoginID from MemberData where MemberID="&session("User_ID")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_LoginID=trim(rs("LoginID"))
rs.close
set rs=nothing

strBil="select distinct BillSN,BillNo,CarNo,DCIerrorCarData from DCILOG where BillSN="&strBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_DCIerrorCarData=""
if Not rsbil.eof then Sys_DCIerrorCarData=trim(rsbil("DCIerrorCarData"))
Sys_Owner=""
Sys_OwnerZip=""
Sys_OwnerZipName=""
Sys_OwnerAddress=""
Sys_BillNo=""
Sys_CarNo=""
Sys_Rule1=""
Sys_Rule2=""

If Not rsbil.eof Then
strSql="select BillTypeID,Driver,DriverID,DriverAddress,DriverZip,RuleVer,BillFillDate from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close
set rs=nothing

if sys_City="基隆市" then
	strSql="select StoreAndSendFinalMailDate from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")
	set rs=conn.execute(strSql)
	if Not rs.eof then Sys_StoreAndSendFinalMailDate=split(gArrDT(trim(rs("StoreAndSendFinalMailDate")&"")),"-")
	rs.close
   set rs=nothing
end if

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)
if sys_City="台東縣" then
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then 
			Sys_Owner=trim(rsfound("Driver"))&""
			If Sys_Owner="" Then trim(rsfound("Owner"))&""
		End if
	else
		if Not rsFound.eof then 
			Sys_Owner=trim(rsfound("Owner"))
			If Sys_Owner="" Then trim(rsfound("Driver"))&""
		End if
	End if
	if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
	if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
	if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

	If ifnull(Sys_OwnerAddress) and trim(Sys_DCIerrorCarData)<>"D" then
		strSql="select * from BillbaseDCIReturn where CarNo in(select CarNo from DciLog where billsn="&trim(rsbil("BillSN"))&" and carno='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"
		set rsdata=conn.execute(strsql)
		If Sys_BillTypeID=1 Then
			if Not rsdata.eof then 
				Sys_Owner=trim(rsdata("Driver"))&""
				If Sys_Owner="" Then Sys_Owner=trim(rsdata("Owner"))&""
			End if
		else
			if Not rsdata.eof then 
				Sys_Owner=trim(rsdata("Owner"))
				If Sys_Owner="" Then Sys_Owner=trim(rsdata("Driver"))&""
			End if
		End if

		if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))
		if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
		if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))
		rsdata.close
	end if

	If ifnull(Sys_OwnerAddress) or len(Sys_OwnerAddress)<8 Then
		if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	end if
else
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver"))
	else
		if Not rsFound.eof then Sys_Owner=trim(rsfound("Owner"))
	End if
	if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
	if Instr(request("Sys_BatchNumber"),"N")>0 then
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
		If ifnull(Sys_OwnerAddress) Then
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
		end if
	else
		If Sys_BillTypeID=1 Then
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

		else
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))

		End if
	end if

	If ifnull(Sys_OwnerAddress) Then
		strSql="select * from BillbaseDCIReturn where CarNo in(select CarNo from DciLog where billsn="&trim(rsbil("BillSN"))&" and carno='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"
		set rsdata=conn.execute(strsql)
		If Sys_BillTypeID=1 Then
			if Not rsdata.eof then Sys_Owner=trim(rsdata("Driver"))
		else
			if Not rsdata.eof then Sys_Owner=trim(rsdata("Owner"))
		End if

		if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))

		if Instr(request("Sys_BatchNumber"),"N")>0 then
			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

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

	If ifnull(Sys_OwnerAddress) Then
		if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))

	end if


end If

					strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
					set rszip=conn.execute(strSQL)
					if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName") & "") 
					rszip.close
					set rszip=nothing

'else
'	if Not rsFound.eof then Sys_Driver=trim(rsfound("Driver") & "") 
'	if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID") & "") 
'	if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("DriverHomeAddress") & "") 
'	if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("DriverHomeZip") & "") 
'	strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
'	set rszip=conn.execute(strSQL)
'	if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName") & "") 
'	rszip.close
'	set rszip=nothing
'end if

'if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner") & "") 
'if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress") & "") 
'if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip") & "") 
'strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
'set rszip=conn.execute(strSQL)
'if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName") & "") 
'rszip.close
'set rszip=nothing

Sys_DCIReturnStation=0
Sum_Level=0
if Not rsfound.eof then Sys_DCIReturnStation=trim(rsfound("DCIReturnStation") & "") 
if Not rsfound.eof then Sys_BillNo=trim(rsfound("BillNo") & "") 
if Not rsfound.eof then Sys_CarNo=trim(rsfound("CarNo") & "") 
if Not rsfound.eof then Sys_Rule1=trim(rsfound("Rule1") & "") 
if Not rsfound.eof then Sys_Rule2=trim(rsfound("Rule2") & "") 
if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1") & "") 
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2") & "") 
Sum_Level=Cdbl(Sys_Level1)+Cdbl(Sys_Level2)

rsfound.close
set rsfound=nothing
'Sys_Sex=""
strSql="select distinct BillFillerMemberID,DealLineDate,RECORDMEMBERID,BILLMEMID1 from BillBase where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"'"
set rssex=conn.execute(strSql)

if Not rssex.eof then Sys_RecordMemberID=trim(rssex("RECORDMEMBERID"))

if Not rssex.eof then
	Sys_DealLineDate=split(gArrDT(trim(rssex("DealLineDate"))),"-")
else
	Sys_DealLineDate=split(gArrDT(trim("")),"-")
end if

Sys_BillFillerMemberID=0
if Not rssex.eof then Sys_Billmem1ID=trim(rssex("BILLMEMID1"))

strSql="select a.LoginID,b.UnitID,b.UnitTypeID,b.UnitLevelID from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&trim(Sys_RecordMemberID)
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
mem.close
set mem=nothing

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if

set unit=conn.Execute(strSQL)
If Not unit.eof Then sysunit=unit("UnitName")
unit.close
set unit=nothing
strSql="select a.LoginID,a.ChName,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_Billmem1ID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillUnitAddress=trim(mem("Address"))
mem.close
set mem=nothing

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
set rssex=nothing

strSQL="select UnitName,Tel,Address from UnitInfo where UnitID=(Select UnitID from MemberData where MemberID='"&trim(Sys_RecordMemberID)&"')"
set Unit=conn.execute(strSQL)
SysUnit=Unit("UnitName")
SysUnitTel=Unit("Tel")
SysUnitAddress=Unit("Address")
Unit.close
set Unit=nothing

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

strSql="select DCISTATIONNAME,StationID from Station where DCIStationID='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_STATIONNAME=trim(rs("DCISTATIONNAME"))
if Not rs.eof then Sys_StationID=trim(rs("StationID"))
rs.close
set rs=nothing

strSql="select MailNumber,MailDate from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))

rs.close
set rs=nothing

if isnull(Sys_DriverHomeZip) or trim(Sys_DriverHomeZip)="" then Sys_DriverHomeZip="001"
if isnull(Sys_OwnerZip) or trim(Sys_OwnerZip)="" then Sys_OwnerZip="001"

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
Sys_MailNumber=0
Sys_BillNo_BarCode=Sys_BillNo
If sys_City<>"台中縣" Then
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160
else
	Sys_BillNo_BarCode=Sys_BillNo_BarCode&"_4"
'	if trim(Sys_BillTypeID)="1" then
'		DelphiASPObj.GenBillPrintBarCode strBillSN(i+PrintSum),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,date
		'response.write "DelphiASPObj.GenBillPrintBarCode"& strBillSN(i+PrintSum)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
		'response.end
'	else
		DelphiASPObj.GenBillPrintBarCode strBillSN(i+PrintSum),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,date
		'response.write "DelphiASPObj.GenBillPrintBarCode "& strBillSN(i+PrintSum)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_OwnerZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
		'response.end
'	end if
end if

end if
rsbil.close
set rsbil=nothing


%>

<div id="L78" style="position:relative;">

<div id=idDivH style="position:absolute; left:560px;top:0px; width:189px; height:24px">
					<img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg" width="150" height="28">
				</div>
<table border="0" width="698" id="table1" height="31" cellspacing="0">
	<tr>
		<td>
			<div style="position: absolute; width: 223px; height: 14px; z-index: 2; left: 187px; top: 5px" id="layer2">
				<b><font face="標楷體" size="3"><%=sys_City%>警察局送達證書</font></b>
			</div>
		</td>
	</tr>
	<tr>
		<td>
			<div style="position: absolute; width: 500px; height: 12px; z-index: 1; left: 18px; top: 17px;" id="layer1">
				<font face="標楷體" size="3">請繳回：<%
					If Sys_UnitLevelID=1 Then 
						response.write sys_City&"警察局交通隊" 
					Else 
						response.write SysUnit
					end If 

					Response.Write "&nbsp;&nbsp;地址："&SysUnitAddress
					%>
				</font>
				
			</div>
		
		</td>
	</tr>
</table>
<table border="1" width="700" id="table1" cellspacing="0" cellpadding="0" height="268">	
	<tr>
		<td align="center" width="67" height="38"><font face="標楷體" size="1">受送達人機關姓名及地址</font></td>
		<td width="280" >

		<%
			response.write (funcCheckFont(trim(Sys_Owner),16,1))
		%>

		<%
			response.write (chstr(Sys_CarNo))
			response.write "<BR>"
			Sys_OwnerAddress=funcCheckFont(replace(Sys_OwnerAddress&"","臺","台"),16,1)
            Sys_OwnerZipName=replace(Sys_OwnerZipName,"臺","台")
			strtmp=Sys_OwnerZip&" "&replace(Sys_OwnerZipName&Sys_OwnerAddress,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName)
			response.write ((trim(strtmp&"")))
		%>
		</td>
		<td width="66" align="center" height="38"><font face="標楷體" size="1">原寄郵局日戳</font></td>
		<td width="277" colspan="3" rowspan="2">
		<div style="position: absolute; width: 382px; height: 143px; z-index: 5; left: 419px; top: 31px" id="layer7">
			<table border="1" width="281" id="table4" cellspacing="0" cellpadding="0" height="65">
				<tr>
					<td width="86">　</td>
					<td width="41" align="center">
		<font face="標楷體" size="1">送達處所</font></td>
					<td><font face="標楷體" size="1">(由送達人填記)</font><font size="1"><br>
					</font>
					<font face="標楷體" size="1">□同左記載地址</font><font size="1">
					<br></font><font face="標楷體" size="1">□改送：</font></td>
				</tr>
				<tr>
					<td width="86">　</td>
					<td width="41" align="center" height="32">
					<font face="標楷體" size="1">送達時間</font></td>
					<td ><font face="標楷體" size="1">(由送達人填記)</font><font size="1"><br>
					</font>
					<font face="標楷體" size="1">中華民國　　年　　月　　日</font><font size="1">
					<br></font><font face="標楷體" size="1">　　　　　　午　　時　　分</font></td>
				</tr>
			</table>
		</div>
		<div style="position: absolute; width: 427px; height: 278px; z-index: 4; left: 0px; top: 137px" id="layer6">
			<table border="1" width="353" id="table3" height="180" cellspacing="0" cellpadding="0">
				<tr>
					<td rowspan="5" align="center"><font face="標楷體" size="1">送達方法</font></td>
					<td rowspan="5" width="10" align="center">
					<font face="標楷體" size="1">由送達人在□上劃V號選記</font></td>
					<td width="320" height="18"><font face="標楷體" size="1">□已將文書交與應受送達人：</font></td>
				</tr>
				<tr>
					<td width="320" height="26" ><font face="標楷體" size="1">□未獲會晤本人，已將文書交與有辨別事理能力之同居人、受僱人或應送達處所之接收郵件人員
					。</font></td>
				</tr>
				<tr>
					<td width="320" height="37"><font face="標楷體" size="1">□應受送達之本人、同居人、受僱人或應送達處所之接收郵件人員收領，但拒絕或不能簽名、蓋章或按指印者，由送達人記明其事由於右欄：</font></td>
				</tr>
				<tr>
					<td width="320" height="25" ><font face="標楷體" size="1">□應受送達之本人、同居人、受僱人或應送達處所之接收郵件人員無法律上之理由拒絕收領經送達人將文書留置於送達處所，以為送達：</font></td>
				</tr>
				<tr>
					<td width="320" height="58" ><font face="標楷體" size="1">□未獲會晤本人亦無受領文書之同居人、受僱人或應送達處所之接收郵件人員，已將該送達文書：<br>
		□應受送達之本人、同居人、受僱人或應送達處所之接收郵件人員，無法律上之理由，拒絕收領，並有難達留置情事，已將該送達文書：</font></td>
				</tr>
			</table>
		</div>
　</td>
	</tr>
	<tr>
		<td align="center" width="67" height="38"><font face="標楷體" size="1">文號</font></td>
		<td width="280" height="38">
			<%=theBillNumber&"第"&(chstr(Sys_BillNo))&"號"%>
			<br>
			<%=(chstr("違反法條"&Sys_Rule1&"   "&Sys_Rule2))%>
		</td>
		<td width="66" align="center" height="38"><font face="標楷體" size="1">送達郵局日戳</font></td>
	</tr>
	<tr>
		<td align="center" width="67" height="23"><font face="標楷體" size="1">送達文書</font></td>
		<td width="280" height="23"><font face="標楷體" size="1">舉發違反道路交通管理事件通知單</font></td>
		<td width="66" height="23"><font face="標楷體" size="1">送達人簽章</font></td>
		<td width="182" height="23">　</td>
		<td width="17" rowspan="6" align="center"><font face="標楷體" size="1">送</font><p>
		<font face="標楷體" size="1">達</font></p>
		<p><font face="標楷體" size="1">人</font></p>
		<p><font face="標楷體" size="1">注</font></p>
		<p><font face="標楷體" size="1">意</font></td>
		<td width="74" rowspan="6"><font face="標楷體" size="1">
		一、依上述送達方法送達者，送達人應即將本送達證書，提出於交送達之行政機關附卷。</font><p>
		<font face="標楷體" size="1">
		二、無法依上述送達方法送達者，送達人應作記載該事由之報告書，提出於交送達之行政機關附卷，並繳回應送達之文書。</font></td>
	</tr>
	<tr>
		<td rowspan="6" width="349" colspan="2">　</td>
		<td width="250" height="18" colspan="2"><font face="標楷體" size="1">□本人　　　　　　　(簽名蓋章或按指印)</font></td>
	</tr>
	<tr>
		<td width="250" colspan="2" height="26"><font face="標楷體" size="1">□同居人□受僱人□應送達處所之接收郵件人員<br>　　　　　　　　　　(簽名蓋章或按指印)</font></td>
	</tr>
	<tr>
		<td width="250" colspan="2" align="left" valign="top" height="38">
		<font face="標楷體" size="1">送達人填記：</font></td>
	</tr>
	<tr>
		<td width="250" colspan="2" height="27"><font face="標楷體" size="1">□本人□同居人□受僱人□應送達處所之接收郵件人員<br>(拒絕收領人姓名)</font></td>
	</tr>
	<tr>
		<td width="250" align="left" valign="top" colspan="2" height="58">
		<div style="position: absolute; width: 338px; height: 158px; z-index: 3; left: 351px; top: 254px" id="layer5">
		<table border="1"  id="table2" cellspacing="0" cellpadding="0" width="254" height="63">
			<tr>
				<td align="left" valign="top"><font face="標楷體" size="1">□寄存於</font></td>
				<td width="146"><font face="標楷體" size="1">並作送達通知書二份，一份粘貼於應受送達人門首，另一份□交由鄰居轉交或□置於受送達處所信箱或適當位置，以為送達。</font></td>
			</tr>
		</table>
		</div>
		</td>
	</tr>

</table>

</Div>
				