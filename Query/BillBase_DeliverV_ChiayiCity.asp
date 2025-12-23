<style type="text/css">
<!--
td {font-family:新細明體;line-height:9px;font-size:9pt;}-->
</style>
<%
strSql="select LoginID from MemberData where MemberID="&session("User_ID")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_LoginID=trim(rs("LoginID"))
rs.close
set rs=nothing

strBil="select distinct BillSN,BillNo,CarNo,DCIerrorCarData from DCILOG where BillSN="&strBillSN(gyi)&" and ExchangetypeID='W'"
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
strSql="select BillTypeID,Driver,DriverID,DriverAddress,DriverZip,RuleVer,BillFillDate,BillUnitID from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_BillUnitID=trim(rs("BillUnitID"))
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
Sys_OwnerAddress=""

if sys_City="台東縣" then
	chkExchange=""
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver"))
	else
		if Not rsFound.eof then Sys_Owner=trim(rsfound("Owner"))
	End if
	if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
	if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
	if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

	If ifnull(Sys_OwnerAddress) and trim(Sys_DCIerrorCarData)<>"D" then
		chkExchange="A"
		strSql="select * from BillbaseDCIReturn where CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A'"
		set rsdata=conn.execute(strsql)
		If Sys_BillTypeID=1 Then
			if Not rsdata.eof then Sys_Owner=trim(rsdata("Driver"))
		else
			if Not rsdata.eof then Sys_Owner=trim(rsdata("Owner"))
		End if

		if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))
		if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
		if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))
		rsdata.close
	end if

	If ifnull(Sys_OwnerAddress) or len(Sys_OwnerAddress)<10 Then
		if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	end if
else
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver")&"")
		If Trim(Sys_Owner&"")="" Then If Not rsFound.eof Then Sys_Owner=trim(rsfound("Owner"))
	else
		if Not rsFound.eof then Sys_Owner=trim(rsfound("Owner")&"")
	End if
	if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))

	if Instr(request("Sys_BatchNumber"),"N")>0 Then

		Sys_OwnerAddress=""
		If sys_City="南投縣" Then
			if Sys_BillTypeID=2 Then
				strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='N'"
			else
				strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
			end if

			set rsdata=conn.execute(strsql)

			If Sys_BillTypeID=1 Then
				if Not rsdata.eof then Sys_Owner=trim(rsdata("Driver"))

				If ifnull(Sys_Owner) Then Sys_Owner=trim(rsdata("Owner"))
			else
				if Not rsdata.eof then Sys_Owner=trim(rsdata("Owner"))
			End if

			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

			rsdata.close
		End if

		If ifnull(Sys_OwnerAddress) Then			
			strSql="select * from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A') and ExchangetypeID='A'"
			set rsdata=conn.execute(strsql)
			If Sys_BillTypeID=1 Then
				if Not rsdata.eof then Sys_Owner=trim(rsdata("Driver")&"")
				If Trim(Sys_Owner&"")="" Then If Not rsdata.eof Then Sys_Owner=trim(rsdata("Owner"))
			else
				if Not rsdata.eof then Sys_Owner=trim(rsdata("Owner")&"")
			End if
			If Not Isnull(rsdata("DriverHomeAddress")) then
				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))
			Else
				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))
			End If

			rsdata.close
		end if

		If ifnull(Sys_OwnerAddress) Then
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
			If ifnull(Sys_OwnerAddress) Then
				if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
				if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
			end If
		End If

	end if

	If ifnull(Sys_OwnerAddress) Then
		strSql="select * from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A') and ExchangetypeID='A'"
		set rsdata=conn.execute(strsql)
		If Sys_BillTypeID=1 Then
			if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver")&"")
			If Trim(Sys_Owner&"")="" Then If Not rsFound.eof Then Sys_Owner=trim(rsfound("Owner"))
		else
			if Not rsFound.eof then Sys_Owner=trim(rsfound("Owner")&"")
		End if

		if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))

		if Instr(request("Sys_BatchNumber"),"N")>0 then
			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

		else
			If Sys_BillTypeID=1 Then
				if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
				if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

			else
				If instr(replace(rsfound("OwnerAddress"),"（","("),"(住")>0 or instr(replace(rsfound("OwnerAddress"),"（","("),"(就") Then
					if Not rsfound.eof then Sys_OwnerAddress=rsfound("OwnerAddress")
					if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
				End if

				If ifnull(Sys_OwnerAddress) Then
					if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
					if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))
				End if
			End if
		end if
		rsdata.close
	end if

	If ifnull(Sys_OwnerAddress) Then
		if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner")&"")
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))

	end if
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName") & "") 
rszip.close
set rszip=nothing
If trim(chkExchange)="A" Then Sys_OwnerZipName=""

Sys_OwnerAddress=trim(replace(Sys_OwnerAddress&" ","臺","台"))

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
set mem=Nothing

strSql="select UnitName from UnitInfo where UnitID='"&trim(Sys_BillUnitID)&"'"
set mem=conn.execute(strsql)
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
mem.close
set mem=Nothing


'If Sys_UnitLevelID=1 Then
'	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
'else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
'end if

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
'SysUnit=Unit("UnitName")
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

strSql="select MailNumber,MailDate,UserMarkMemberID from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_UserMarkMemberID=trim(rs("UserMarkMemberID"))
rs.close
set rs=nothing

If sys_City="南投縣" Then
	If not ifnull(Sys_UserMarkMemberID) Then
		strSQL="select loginid from memberdata where memberid="&Sys_UserMarkMemberID
		set rsmem=conn.execute(strSQL)
		If not rsmem.eof Then Sys_BillFillerMemberID=trim(rsmem("loginid"))
		rsmem.close
	End if
	
end if

if isnull(Sys_DriverHomeZip) or trim(Sys_DriverHomeZip)="" then Sys_DriverHomeZip="001"
if isnull(Sys_OwnerZip) or trim(Sys_OwnerZip)="" then Sys_OwnerZip="001"

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
Sys_MailNumber=0
Sys_BillNo_BarCode=Sys_BillNo
If sys_City<>"台中縣" Then
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160
else
	Sys_BillNo_BarCode=Sys_BillNo_BarCode&"_4"
end if
if trim(Sys_DriverHomeZip)="001" then Sys_DriverHomeZip=""
if trim(Sys_OwnerZip)="001" then Sys_OwnerZip=""
end if


strSql="select MailNumber,StoreAndSendMailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))

rs.close
rsbil.close
set rsbil=nothing

%>
<br>
<div id="L78" style="position:relative;">
	<div id="D78" style="position:absolute;left:-10px;">
		<table border="0" cellspacing="0">
		<td>&nbsp;</td>
		<td>&nbsp;</td>
			<th width="44" rowspan="4" align="center" valign="top" class="style4"><br>
					<table border="0">
				<%  tempdd=""

							If sys_City="南投縣" Then 
								tempdd=right("00"&gInitDT(now),7) 
							elseIf sys_City="台南市" or sys_City="台南縣" Then 
								 tempdd=right("00"&gInitDT(now),7)
							elseIf sys_City="嘉義縣" Then
								tempdd=right("00"&gInitDT(Sys_MailDate),7)
							elseif sys_City="基隆市" then
								 tempdd=right("00"&Sys_StoreAndSendFinalMailDate(0),3)&Sys_StoreAndSendFinalMailDate(1)&Sys_StoreAndSendFinalMailDate(2)
							ElseIf Sys_City="台中市" Then 
								tempdd=Sys_UnitName
							else
								tempdd=right("00"&sys_Date(0),3)&sys_Date(1)&sys_Date(2)
							end if
				%>
				<br><br><%If sys_City="嘉義縣" Then 
							response.write wordporss(chstr(tempdd)) 
						  ElseIf Sys_City="台中市" Then 
								Data2=""
								If chkStore=0 Then
									Data2=Sys_MailNumber
								else
									Data2=Sys_StoreAndSendMailNumber
								End if
							response.write wordporss2(chstr(tempdd),chstr(Data2)) 
						  else
							response.write wordporss2(chstr(tempdd),chstr(gyi+1)) 
						  End if%>
				 <%
							If sys_City="南投縣" Then 
								response.Write "<br>：：<br>期號<br>日&nbsp;&nbsp;&nbsp;&nbsp;<br>統&nbsp;&nbsp;&nbsp;&nbsp;<br>系序"
							elseIf sys_City="台南市" or sys_City="台南縣" Then
								response.Write " <br>：：<br>期號 <br>日&nbsp;&nbsp;&nbsp;&nbsp;<br>寄&nbsp;&nbsp;&nbsp;&nbsp;<br>郵序"
							elseIf sys_City="嘉義縣" Then
								response.Write "      <br>：：<br>期號 <br>日&nbsp;&nbsp;&nbsp;&nbsp;<br>寄&nbsp;&nbsp;&nbsp;&nbsp;<br>郵序"
							elseif sys_City="基隆市" then
								response.Write "<br>：：<br>期號 <br>日&nbsp;&nbsp;&nbsp;&nbsp;<br>寄&nbsp;&nbsp;&nbsp;&nbsp;<br>郵序"
							ElseIf Sys_City="台中市" Then 
								response.Write "<br>：：<br>位碼 <br>單號&nbsp;&nbsp;&nbsp;&nbsp;<br>發號&nbsp;&nbsp;&nbsp;&nbsp;<br>舉掛"
							else
								response.Write "<br>：：<br>期號 <br>日&nbsp;&nbsp;&nbsp;&nbsp;<br>單&nbsp;&nbsp;&nbsp;&nbsp;<br>填序"
							end if
				 %>
					<%If sys_City="澎湖縣" Then %>
						  <br>&nbsp;
					<%else%>
					  <br>&nbsp;
					  <br>&nbsp;
					  <br>&nbsp;
					  <br>&nbsp;
					<%End if%>

						  <br><span class="style5"><%
						  If sys_City="澎湖縣" Then 
							response.write wordporss(chstr(thenPasserCity&replace(sysunit,trim(thenPasserCity),"")&"送達證書"))
						  Else
							response.write wordporss(chstr(thenPasserCity&replace(sysunit,trim(thenPasserCity),"")))
						  End if
						  %></span>
					  </table>
				  </th>
		</table>

	</Div>
</Div>

<table width="630" bordercolor="#000000" border="1" cellspacing="0" class="tablestyle">
  <tr>
    <td width="10%" rowspan="2" class="style4">&nbsp;
		<table border="0" height="250">
			<tr>
			<td valign="top" align="left" width="-5">
				<div id=idDiv class="img1">
					<img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg" style="transform:rotate(-90deg);">
				</div>
			</td>
			<td width="6%" valign="bottom" class="style9" >
			   <%
			   If Trim(Sys_Owner)<>"" Then 
				  k=7
				  'response.write wordporss(trim(funcCheckFont(Sys_Owner,14,4)))
				  strtmp=(trim(funcCheckFont(Sys_Owner,14,4)))
			 For g=1 To Len(strtmp) 
			  If Asc(Mid(strtmp,g,1))>0 Then k=k+1
			 Next 
				response.write wordporss((trim(Mid(strtmp&"",1,k))))
				response.write "<td width=""6%"" valign=""bottom"" class=""style9"">"&wordporss((trim(Mid(strtmp&"",k+1))))&"</td>"

			   End if
			   %>
			</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td width="2%" valign="bottom" class="style10"><%
			If Trim(Sys_OwnerAddress)<>"" Then 
				Sys_OwnerAddress=funcCheckFont(replace(Sys_OwnerAddress&"","臺","台"),14,4)
			End If
			If Trim(Sys_OwnerZipName)<>"" Then 			
			   Sys_OwnerZipName=funcCheckFont(replace(Sys_OwnerZipName&"","臺","台"),14,4)
			End if
				strtmp=chstr(Sys_OwnerZip)&" "&replace(Sys_OwnerZipName&Sys_OwnerAddress,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName)

			If Trim(strtmp)<>"" Then 			
			 k=20
			 For g=1 To Len(strtmp) 
			  If Asc(Mid(strtmp,g,1))>0 Then k=k+1
			 Next 
				response.write wordporss((trim(Mid(strtmp&"",1,k))))
			End If
			
			%>
			<td width="2%" valign="bottom" class="style10"><%
				strtmp=chstr(Sys_OwnerZip)&" "&replace(Sys_OwnerZipName&Sys_OwnerAddress,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName)
			If Trim(strtmp)<>"" Then 			
				response.write wordporss((trim(Mid(strtmp&"",k+1))))
			End if
			%></td>
			</td>
			</tr>
		</table>
	</td>
    <td width="3%" rowspan="2" align="center" valign="bottom" class="style4"> 
	<%if Sys_BillTypeID="1" Then
			If Trim(Sys_CarNo)<>"" Then
				If sys_City="嘉義市" then
					response.write wordporss(chstr(Sys_CarNo))
				else
					response.write wordporss(chstr(left(Sys_CarNo,4)&left("*************",len(Sys_CarNo)-4)))
				end if
			End if
		Else
			If Trim(Sys_CarNo)<>"" Then
				If sys_City="嘉義市" then
					response.write wordporss(chstr(Sys_CarNo))
				else
					response.write wordporss(chstr(left(Sys_CarNo,4)&left("*************",len(Sys_CarNo)-4)))
				end if
			End if
		end if%><br><br><br><br><br><br><br><br>號
		<br>
		<%=wordporss(chstr(Sys_BillNo))%><br>
    第</td>
    <td width="3%" rowspan="2" align="right" valign="bottom" bordercolor="#000000" class="style4">
	<table border="0">
	<td valign="bottom" class="style4"><%=wordporss(chstr("舉發違反道路交通管理事件通知單"))%></td>
	<td valign="bottom" class="style4"><%=wordporss(chstr("違反法條"&Sys_Rule1&"   "&Sys_Rule2))%></td>
	</table>	</td>
    <td width="6%" colspan="2" align="center" valign="middle" bordercolor="#000000" class="style4">      章
    <br>簽
    <br>人
    <br>達
    <br>
    送</td>
    <td width="10%" colspan="2" align="right" valign="bottom" bordercolor="#000000" class="style4" >&nbsp;</td>
    <td width="3%" rowspan="4" align="center" valign="middle" bordercolor="#000000" class="style4"><br>
      式
      <br>&nbsp;
      <br>&nbsp;
      <br>        <br>
  &nbsp;
        <br>
      方
      <br>&nbsp;
      <br>&nbsp;
      <br> <br>
  &nbsp;
        <br>
      達
      <br>&nbsp;
      <br>&nbsp;
      <br> <br>
  &nbsp;
        <br>
    送</td>
    <td width="3%" rowspan="4" align="center" valign="middle" bordercolor="#000000" class="style4"><br>
      記
        
      <br><br>
      選
      <br> <br>
      Ｖ
      <br> <br>
      劃
      <br> <br>
      上
      <br> <br>
      □
      <br> <br>
      在
      <br> <br>
      人
      <br> <br>
      達
      <br> <br>
      送
      <br> <br>
    
    <br>由</td>
    <td width=20 rowspan="2" align="center" valign="bottom" bordercolor="#000000">︵<span  class="style4">
      <br>
      章
      <br>
      蓋
      <br>
      或
      <br>
      名
      <br>
      簽
      <br>
   </span>︶<span  class="style4">
      <br>&nbsp;
      <br>&nbsp;
      <br>&nbsp;
      <br>&nbsp;
      <br>&nbsp;
      <br>&nbsp;
      <br>&nbsp;
      <br><br>
&nbsp;
        <br>
&nbsp;
        <br>
&nbsp;
        <br>
&nbsp;
        <br>
        <br>
&nbsp;&nbsp;
      <br>
&nbsp;
      <br>
&nbsp;
      <br>
人<br>
本 <br>
□ </td>
    <td width=39 rowspan="2" align="left" valign="bottom" bordercolor="#000000" >
&nbsp;&nbsp;&nbsp;︵ <span  class="style4"><br>
　章 <br>
　蓋 <br>
　或 <br>
　名 <br>
　簽 <br>
　</span>︶ <span  class="style4"><br>
&nbsp; <br>
&nbsp; <br>
&nbsp; <br>
&nbsp; <br>
&nbsp; <br>
&nbsp; <br>
&nbsp; <br>
    　　員<br>
　　人<br>
　　件<br>
　　郵<br>
　　收<br>
　　接<br>
　　所<br>
　　處<br>
人人達<br>
居雇送<br>
同受應<br>
□□□</td>
    <td width=36 rowspan="2" align="center" valign="bottom" bordercolor="#000000" class="style4">
		：
      <br>
      記
      <br>
      填
      <br>
      人
      <br>
      達
      <br>
    送</td>
    <td width=51 rowspan="2" align="left" valign="bottom" bordercolor="#000000" class="style4">&nbsp;<br>
        　　　員<br>
        　　　人<br>
        　　　件<br>
        　　　郵<br>
        　　　收<br>
        　　　接<br>
        　　　所<br>
        處人人達<br>
        人居雇達<br>
        本同受應<br>
    □□□□</td>
    <td width=91 align="left" valign="bottom" bordercolor="#000000" class="style4">      份、處轉信送
    <br>
    一所業居所為
    <br>，居就鄰處以
    <br>份住其由達，
    <br>二人或交送置
    <br>書達所□受位
    <br>知送業份該當
    <br>通受營一於適
    <br>達應、，置他
    <br>送於所首□其
    <br>作貼務門或或。
    <br>並黏事所交箱達</td>
    <td width="5%" rowspan="2" align="left" valign="bottom" bordercolor="#000000" class="style4">      於　報書
      <br>
      出　之文
      <br>
      提　由之
      <br>
      ，　事達
      <br>
      書　該送
      <br>
      達　記回
      <br>
      送　作繳
      <br>
      本　製並
      <br>
      將　應，
      <br>
      即　人卷
      <br>
      應　達附
      <br>
      人　送關
      <br>
      達　，機
      <br>
      送　者政
      <br>
      ，。達行
      <br>
      者卷送之
      <br>
      達附法達
      <br>
      送關方送
      <br>
      法機達交
      <br>
      方政送於
      <br>
      達行述出
      <br>
      送之上提
      <br>
      述達依，
      <br>
      上送能書
      <br>
      依交不告
      <br>
      、　、
      <br>
    一　二</td>
   
  </tr>
  <tr>
    <td width="27" align="center" valign="bottom" bordercolor="#000000">︵ <span  class="style4">
      <br>
      記
    <br>填
    <br>人
    <br>達
    <br>送
    <br>由
    <br></span> ︶ <span  class="style4">
    <br>所
    <br>處
    <br>達
    <br>送</td>
    <td width="5%" align="left" valign="bottom" bordercolor="#000000" class="style4">
	地<br>址<br>載<br>記&nbsp;：<br>上&nbsp;送<br>同&nbsp;改<br>□&nbsp;□</td>
    <td width=23 align="center" valign="bottom" bordercolor="#000000">&nbsp;︵ <span  class="style4">
      <br>
      記 <br>
        填 <br>
        人 <br>
        達 <br>
        送 <br>
        由 <br>
        </span>&nbsp;︶<span  class="style4"> <br>
        間 <br>
        時 <br>
        達 <br>
    送</td>
    <td width=28 align="left" valign="bottom" bordercolor="#000000" class="style4">      日 分 
      <br>
      <br><br><br>
        月 時 
      <br><br><br><br>
        年 午 <br><br><br><br>
  &nbsp; <br>
        國 <br>
        民 <br>
        華 <br>
        中 <br>
  &nbsp; </td>
    <td width=91 align="left" valign="bottom" bordercolor="#000000">&nbsp;&nbsp;&nbsp;<font style="font-size:5pt"> ︵</font>&nbsp;&nbsp;&nbsp;<font style="font-size:5pt"> ︵</font><span  class="style4">
        <br>
&nbsp;&nbsp;區&nbsp;&nbsp;區&nbsp;&nbsp;處 <br>
&nbsp;&nbsp;、&nbsp;&nbsp;、&nbsp;&nbsp;公 <br>
&nbsp;&nbsp;市&nbsp;&nbsp;市&nbsp;&nbsp;辦 <br>
&nbsp;&nbsp;、&nbsp;&nbsp;、&nbsp;&nbsp;</span><font style="font-size:5pt"> ︵</font><span  class="style4">
        <br>
        所鎮&nbsp;&nbsp;鎮&nbsp;&nbsp;里 <br>
出</span><font style="font-size:5pt"> ︶</font><span  class="style4">所</span><font style="font-size:5pt"> ︶</font><span  class="style4">所</span><font style="font-size:5pt"> ︶</font><span  class="style4">局    
      <br>派鄉公鄉公村郵 <br>
  &nbsp;
        <br>
        <br>
        <br>
        <br>
        <br>
        <br>
  於於　於　於 <br>
        存存　存　存 <br>
        寄寄　寄　寄 <br>
    □□　□　□ </td>
  </tr>
  <tr>
    <td rowspan="2" align="center" valign="middle" bordercolor="#000000" class="style4">      址 <br>
      地 <br>
      名 <br>
      姓 <br>
      稱 <br>
      名 <br>
      人 <br>
      達 <br>
      送 <br>
    受</td>
    <td rowspan="2" align="center" bordercolor="#000000" class="style4" width="25">號<br>&nbsp;
      <br>&nbsp;
      <br>&nbsp;
      <br>&nbsp;
      <br>文</td>
    <td rowspan="2" align="center" valign="middle" bordercolor="#000000" width="46">︵<span  class="style4">
<br>由
    <br>案
    <br>含</span>
    <br>︶<span  class="style4">
	<br>書
    <br>文
    <br>達
    <br>送</td>
    <td colspan="2" align="center" bordercolor="#000000" width="61" class="style4">戳
    <br>
    日
    <br>
    局
    <br>
    郵
    <br>
    達
    <br>
    送</td>
    <td colspan="2" bordercolor="#000000">&nbsp;</td>
    <td width="3%" rowspan="2" align="center" valign="bottom" bordercolor="#000000" class="style4"><br>
      人
      <br>
      達
      <br>
      送
      <br>
      受
      <br>
      應
      <br>
      與
      <br>
      交
      <br>
      書
      <br>
      文
      <br>
      將
      <br>
      已
      <br>
    □</td>
    <td width="4%" rowspan="2" align="left" valign="bottom" bordercolor="#000000" class="style4">      別達　 
      <br>
      辨送　
      <br>
      有應　
      <br>
      與或　
      <br>
      交人　
      <br>
      書雇　
      <br>
      文受　
      <br>
      將、員
      <br>
      已人人
      <br>
      ，居件
      <br>
      人同郵
      <br>
      本之收
      <br>
      晤力接
      <br>
      會能之
      <br>
      獲理所
      <br>
      未事處
      <br>
    □　　</td>
    <td width="4%" rowspan="2" align="left" valign="bottom" bordercolor="#000000" class="style4">收由　
      <br>
      人，　
      <br>
      雇者
      <br>
      受章
      <br>
      或蓋
      <br>
      人或
      <br>
      居名
      <br>
      同簽
      <br>，能由
      <br>人不事
      <br>本或其
      <br>之絕明
      <br>達拒記
      <br>送，人
      <br>受應達
      <br>應領送
    <br>
    □　　</td>
    <td width="5%" rowspan="2" align="left" valign="bottom" bordercolor="#000000" class="style4">或理送
    　<br>
    人當於
    　<br>
    雇正置
    　<br>
    受無留
    　<br>
    、員書
    　<br>
    人人文
    　<br>
    居件將
    　<br>
    同郵人
    　<br>、收達達
    <br>人接送送
    <br>本所經為
    <br>之處領以
    <br>達達收，
    <br>送送絕所
    <br>受受拒處
    <br>應應由達
    <br>
    □　　　</td>
    <td width="9%" rowspan="2" align="left" valign="bottom" bordercolor="#000000" class="style4"> <br>     　　　
      或理將
      <br>
      居郵&nbsp;&nbsp;&nbsp;人常已<br>
        同收&nbsp;&nbsp;&nbsp;雇正事<br>
        之接&nbsp;&nbsp;&nbsp;受無情
        <br>
        書所：
        、員置<br>
        文處書
        人人留<br>
        領達文
        居件達<br>
        受送達
        同郵難<br>
        無受送
        、收有<br>
        亦應該&nbsp;人接並<br>
        人或將
        本所，：<br>
        本人已
        之處領書<br>
        晤雇，
        達達收文<br>
        會受員
        送送絕達<br>
        獲、人
        受受拒送<br>
        未人件
        應應由該<br>
    □　 　□</td>
    <td rowspan="2" align="center" valign="middle" bordercolor="#000000" class="style4"><br>
      項
    <br>&nbsp;
    <br>事
    <br>&nbsp;
    <br>意
    <br>&nbsp;
    <br>注
    <br>&nbsp;
    <br>人
    <br>&nbsp;
    <br>達
    <br>&nbsp;
    <br>送</td>
  </tr>
  <tr>
    <td colspan="2" align="center" valign="middle" bordercolor="#000000" class="style4">戳
      <br>
      日
      <br>
      局
      <br>
      郵
      <br>
      寄
      <br>
    原</td>
    <td colspan="2" bordercolor="#000000" width="55">　</td>
  </tr>
</table>
<div id="L92" style="position:relative;">
<div id="D79" style="position:absolute;left:660px;top:-350px">
<table border="0">
<th align="center" valign="bottom" bordercolor="#FFFFFF" class="style5">
    <%=wordporss(chstr(thenPasserCity&replace(sysunit,trim(thenPasserCity),"")))%></th>
    <td align="center" valign="bottom" bordercolor="#FFFFFF" class="style4"><br>
	<%If Sys_City<>"台中市" Then %>
	  <%=wordporss(chstr(Sys_BillFillerMemberID))%><br>
      ：
      <br>員
      <br>人
      <br>作
      <br>操
	<%End if%>
      <br>&nbsp;
      <br>&nbsp;
      <br>&nbsp;
	  <br><%=wordporss(chstr(Sys_STATIONNAME))%>
      <br>：
      <br>所
      <br>處
      <br>案
      <br>到      
	  <br>應
    </td>
    <td align="center" valign="bottom" bordercolor="#FFFFFF" class="style4">
	  <% If Sys_City="雲林縣" Then 
			response.write wordporss(chstr(Sys_BillNo))&"<br><br>"
		 End if%>
	  <%If Sys_City="台中市" Then 
			response.write wordporss(chstr("臺中市西屯區大隆路１９２號"))
	    Else
			response.write wordporss(chstr(SysUnitAddress))
	    End if
	  %>

	  <br>
      ：
      <br>地 <br>
        址 <br>
        回 <br>
        送 <br>
        書 <br>
        證 <br>
    本</td>
</table>
</Div></Div>