<%
'strSQL="select * from UnitInfo where UnitID='"&DB_UnitID&"'"
'set rsUnit=conn.execute(strSQL)
'DB_UnitName=trim(rsUnit("UnitName"))
'DB_UnitTel=trim(rsUnit("Tel"))
'rsUnit.close

Sys_BillBaseDriverID=""

strBil="select distinct BatchNumber,BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='"&Sys_ExchangetypeID&"'"
set rsbil=conn.execute(strBil)
Sys_OwnerZip=""
Sys_OwnerZipName=""
If Not rsbil.eof Then
Sys_BatchNumber=trim(rsbil("BatchNumber"))

strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed=""
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_BillBaseDriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close

Sys_OwnerAddress="":Sys_OwnerZip=""

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
	If Sys_BillTypeID=1 or (Sys_BillTypeID=2 and Sys_BillBaseDriverID<>"") Then
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
		if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver")&"")

		If ifnull(Sys_OwnerAddress) Then
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
		end if
	else
		If sys_City="嘉義縣" then

			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

		else
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
		end if
	End if
	
	
end if
rsfound.close

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
		If Sys_BillTypeID=1 or (Sys_BillTypeID=2 and Sys_BillBaseDriverID<>"") Then
			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

			If ifnull(Sys_OwnerAddress) Then
				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))
			end if
		else
			If sys_City="嘉義縣" then

				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

				If ifnull(Sys_OwnerAddress) Then
					if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
					if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))
				end If 

			else
				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))
			end If 
		End if
	end if
	rsdata.close
end if

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)

If Sys_BillTypeID=1 or (Sys_BillTypeID=2 and Sys_BillBaseDriverID<>"") Then

	if Not rsfound.eof then Sys_Owner=trim(rsfound("Driver"))

	If ifnull(Sys_Owner) Then
		Sys_Owner=trim(rsfound("Owner"))
	end if

else
	if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))

	If sys_City="嘉義縣" then
		If ifnull(Sys_OwnerAddress) Then
			if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
		end If 
	end If 

End if

If ifnull(Sys_OwnerAddress) Then
	if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
end If 

if Instr(request("Sys_BatchNumber"),"W")>0 then 

	strSQL="Update Billbase set Owner='"&Sys_Owner&"',OwnerZip='"&Sys_OwnerZip&"',OwnerAddress='"&Sys_OwnerAddress&"',DriverZip='"&trim(rsfound("DriverHomeZip"))&"',DriverAddress='"&trim(rsfound("DriverHomeAddress"))&"' where sn="&trim(rsbil("BillSN"))&" and OwnerAddress is null"

	conn.execute(strSQL)
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=replace(trim(rszip("ZipName")),"台","臺")
if Not rszip.eof then
	if not isnull(Sys_OwnerAddress) then '6/25
		Sys_OwnerAddress=replace(Sys_OwnerAddress,trim(rszip("ZipName")),"")
	end if
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
Sum_Level=cdbl(Sys_Level1)+cdbl(Sys_Level2)
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
'if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
'if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
'if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
if Not mem.eof then Sys_UnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

strSql="select a.LoginID,a.ChName,a.ImageFilename as MemberFileName,b.UnitName,b.UnitID,b.UnitLevelID,b.UnitTypeID,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_Billmem1ID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_BillUnitAddress=trim(mem("Address"))
if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
'if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
'if Not mem.eof then Sys_ChName=trim(mem("ChName"))
'if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

'20120515 by jafe 新營拖吊小隊一直抓到上面的隆田，所以加入這一段
If sys_City = "台南市" Then
	If Sys_UnitID = "07D3" or Sys_UnitID = "07D4" or Sys_UnitID = "07D2" or Sys_UnitID = "07C4" Then 
		Sys_UnitTypeID = Sys_UnitID
	End if
End if

strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"

set unit=conn.Execute(strSQL)
If Not unit.eof Then sysunit=unit("UnitName")
if Not unit.eof then SysUnitTel=trim(unit("Tel"))
if Not unit.eof then SysUnitAddress=trim(unit("Address"))
unit.close

If sys_City="宜蘭縣" Then 
	If Sys_RecordMemberID = 6227 or Sys_RecordMemberID = 6607 Then
		thenPasserCity="宜蘭縣政府"
		sysunit="交通處"
		SysUnitTel="(03)9251000"
		SysUnitAddress="26060 宜蘭市縣政北路1號"
	end If 

End if 


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

strSql="select MailNumber,StoreAndSendMailNumber,MailTypeID,MailDate,UserMarkDate,StoreAndSendFinalMailDate from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))
if Not rs.eof then Sys_StoreAndSendFinalMailDate=trim(rs("StoreAndSendFinalMailDate"))
if Not rs.eof then UserMarkDate=trim(rs("UserMarkDate"))
rs.close

strSql="select LoginID from MemberData where MemberID="&session("User_ID")
set rs=conn.execute(strSql)
if Not rs.eof then Sys_LoginID=trim(rs("LoginID"))
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
'if isnull(Sys_OwnerZip) or trim(Sys_OwnerZip)="" then Sys_OwnerZip="001"

Sys_MailNumber=0
Sys_BillNo_BarCode=Sys_BillNo

If sys_City="高雄市" or sys_City="金門縣" or sys_City="苗栗縣" or sys_City="保二總隊三大隊一中隊" or sys_City="保二總隊三大隊二中隊" Then
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160,1
else
	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160
end if

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
end if
rsbil.close
%>
<table width="650" border="0" cellspacing=0 cellpadding=0>
	<tr><th rowspan=2 valign="bottom" align="center" width="80%">
			<%If sys_City="台南市" then response.Write "<span class=""style2"">批號："&Sys_BatchNumber&"</span>"%>
			<strong><span class="style4">　　<%
			If sys_City="連江縣" then
				Response.Write replace(thenPasserCity,"交通隊","")&replace(sysunit,trim(thenPasserCity),"")
			else
				Response.Write thenPasserCity&replace(sysunit,trim(thenPasserCity),"")
			End if 
			%>送達證書</span>
			<br>
			<span class="style3">
			本證書送回地址：<%
				if sys_City="高雄市" and instr(Sys_BillUnitName,"分隊")>0 then
					Response.Write Sys_BillUnitAddress

				elseIf sys_City="苗栗縣" Then
					response.write "苗栗中苗郵局第260信箱"

				else
					Response.Write SysUnitAddress
				end if
			%></span>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			</strong>
		</th>
		<th valign="bottom" align="left" class="style1" nowrap><%
			If sys_City="台南市" then
				If trim(Sys_LoginID)="A043" or trim(Sys_LoginID)="A139" or trim(Sys_LoginID)="A019" or trim(Sys_LoginID)="A106" Then
					response.write "<font size=5>活動測速照相</font><br>"
				end if
				If trim(Sys_LoginID)="A011" Then
					response.write "<font size=5>現場攔停</font><br>"
				end if
			End if
			If sys_City="南投縣" Then response.Write "系統日期："&right("00"&gInitDT(now),7)&"<br>"
			If sys_City="台南市" Then
				'response.Write "郵寄日期："&right("00"&gInitDT(date),7)
				If instr(Sys_BatchNumber,"N")>0 Then
					response.Write "郵寄日期："&right("00"&gInitDT(UserMarkDate),7)
				else
					response.Write "填單日期："&right("00"&sys_Date(0),3)&sys_Date(1)&sys_Date(2)
				End if

			elseIf sys_City="嘉義縣" Then
				response.Write "郵寄日期："&right("00"&gInitDT(Sys_MailDate),7)

			elseif sys_City="基隆市" then
				response.Write "郵寄日期："&right("00"&gInitDT(Sys_StoreAndSendFinalMailDate),7)

			else
				response.Write "填單日期："&right("00"&sys_Date(0),3)&sys_Date(1)&sys_Date(2)
			end if%>
		</th>
	</tr>
	<tr><th valign="bottom" align="left" class="style1" nowrap>
			序　　號：<%If sys_City<>"嘉義縣" Then response.Write Sys_LoginID &" - "& i+1%>
		</th>
	</tr>
</table>
<table width="645" border="1" cellspacing=0 cellpadding=0>
  <tr>
    <td colspan="2" align="center"><span class="style1">受送達人名稱姓名地址</span></td>
    <td colspan="3" Valign="top"><span class="style1"><br><%
		If sys_City<>"基隆市" and sys_City<>"高雄市" and sys_City<>"保二總隊四大隊二中隊" Then 
			response.write funcCheckFont(Sys_Owner,20,1)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;車號："
			if trim(Sys_CarNo)<>"" and not isnull(Sys_CarNo) then
				'If sys_City <> "嘉義市" then
				'	response.write left(Sys_CarNo,4)
				'	response.write left("*************",len(Sys_CarNo)-4)

				'else
					Response.Write Sys_CarNo

				'End if				
			end if 
		else
			response.write funcCheckFont(Sys_Owner,20,1)
		end if
		response.write "<br>"&Sys_OwnerZip&"&nbsp;&nbsp;"&Sys_OwnerZipName
		If not ifnull(Sys_OwnerAddress) Then
			response.write funcCheckFont(replace(Sys_OwnerAddress,Sys_OwnerZipName,""),20,1)
		else
			response.write funcCheckFont(Sys_OwnerAddress,20,1)
		End if%>&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><span class="style2">文　　　　　　　　　號</span></td>
    <td colspan="3" nowrap><span class="style2"><%
	'只能宜警不能加縣
	if (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕") then
		theBillNumber="宜警"
	end if	
	response.write theBillNumber
	%>交字第<img src=<%="""../BarCodeImage/"&Sys_BillNo_BarCode&".jpg"""%>>號</span></td>
  </tr>
  <tr>
    <td colspan="2" align="center" nowrap><span class="style2">送　達　文　書</span></td>
    <td colspan="3"><span class="style3">舉發違反道路交通管理事件通知單<br>
	<%response.write "道路交通管理處理條例第"&left(trim(Sys_Rule1),2)&"條"
			if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)&" "
				If sys_City = "嘉義縣" then
					response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"
				else
					response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)&"規定"
				end if
				'response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"
			if trim(Sys_Rule2)<>"0" then
				response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
				if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)&" "
				If sys_City = "嘉義縣" then
					response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款"
				else
					response.write Mid(trim(Sys_Rule2),3,1)&"項"&Mid(trim(Sys_Rule2),4,2)&"款"&Mid(trim(Sys_Rule2),6,2)&"規定"
				end if
				'response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
			end if
			%></span></td>
  </tr>
  <tr>
    <td rowspan="2" height=100 align="center"><span class="style2">原寄郵局日戳</span></td>
    <td rowspan="2" align="center"><span class="style2">送達郵局日戳</span></td>
    <td colspan="2" align="center"><span class="style3">送達處所（由送達人填記）</span></td>
    <td rowspan="2" width="20%" align="center"><span class="style2">送達人簽章</span></td>
  </tr>
  <tr>
    <td colspan="2">
		<table border=0>
			<tr><td>□</td><td class="style2"><span class="style2">同上記載地址</span></td>
			</tr>
			<tr><td>□</td><td class="style2"><span class="style2">改送：</span></td>
			</tr>
		</table>
	</td>
  </tr>
  <tr>
    <td rowspan="2" height=100>&nbsp;</td>
    <td rowspan="2">&nbsp;</td>
    <td colspan="2" align="center"><span class="style3">送達時間（由送達人填記）</span></td>
    <td rowspan="2"><span class="style2">&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="2">
		<span class="style1">
			<table width="100%" border=0>
				<tr><td><span class="style3">中華民國</span></td>
				<td><span class="style3">　　　　年　　　　月　　　　日</span></td>
				</tr>
				<tr><td class="style3">&nbsp;</td>
				<td><span class="style3">　　　　午　　　　時　　　　分</span></td>
				</tr>
			</table>
		</span>
	</td>
  </tr>
  <tr>
    <td colspan="5" align="center">
		<span class="style1">送　　　　　　　　達　　　　　　　　方　　　　　　　　式</span>
	</td>
  </tr>
  <tr>
    <td colspan="5" align="center">
		<span class="style1">由　　送　　達　　人　　在　　□　　上　　劃　　v　　選　　記</span>
	</td>
  </tr>
  <tr>
    <td colspan="2">
		<table border=0><tr><td>□</td><td><span class="style2">已將文書交與應受送達人</span></td></tr></table>		
	</td>
    <td colspan="3">
		<table border=0>
			<tr><td>□</td><td><span class="style2">本人　　　　　　　　　　　　　　　（簽名或蓋章）</span></td></tr>
		</table>
	</td>
  </tr>
  <tr>
    <td colspan="2">
		<table border=0>
			<tr><td valign="top">□</td>
				<td>
					<span class="style2">未獲會晤本人，已將文書交與有辨別事理能力之同居人、
					受雇人或應送達處所之接收郵件人員</span>
				</td>
			</tr>
		</table>
	</td>
    <td colspan="3">
		<table border=0>
			<tr><td valign="top">□</td><td class="style3">同居人</span></td>
			</tr>
			<tr><td valign="top">□</td><td class="style3">受雇人　　　　　　　　　　　　　　　　　（簽名或蓋章）</span></td>
			</tr>
			<tr><td valign="top">□</td><td class="style3">應送達處所接收郵件人員</span></td>
			</tr>
		</table>
	</td>
  </tr>
  <tr>
    <td colspan="2">
		<table border=0>
			<tr><td valign="top">□</td>
			<td>
				<span class="style3">應受送達之本人、同居人或受雇人收領後，拒絕或不能簽名或蓋章者，
				由送達人記明其事由</span>
			</td>
			</tr>
		</table>
	</td>
    <td colspan="3"><span class="style2">送達人填記：</span></td>
  </tr>
  <tr>
    <td colspan="2">
		<table border=0>
			<tr><td valign="top">□</td>
			<td><span class="style3">應受送達之本人、同居人、受雇人或應受送達處所接收郵件人員無正當理由
				拒絕領經送達人將文書留置於送達處所，以為送達</span></td>
			</tr>
		</table>
	</td>
    <td colspan="3">
		<table border=0>
			<tr><td valign="top">□</td><td><span class="style3">本人</span></td>
			</tr>
			<tr><td valign="top">□</td><td><span class="style3">同居人　　　　　　　　　拒絕收領</span></td>
			</tr>
			<tr><td valign="top">□</td><td><span class="style3">受雇人</span></td>
			</tr>
			<tr><td valign="top">□</td><td><span class="style3">應受送達處所接收郵件人員</span></td>
			</tr>
		</table>
	</td>
  </tr>
  <tr>
    <td colspan="2">
		<table border=0>
			<tr><td valign="top">□</td><td><span class="style3">未獲會晤本人亦無受領文書之同居人、受雇人或應受送達處所接收郵件人員，
				已將該送達文書：</span></td>
			</tr>
			<tr><td valign="top">□</td><td><span class="style3">應受送達之本人、同居人、受雇人或應受送達處所接收郵件人員無正當理由
				拒絕收領，並有難達留置情事，已將該送達文書：</span></td>
			</tr>
		</table>
	</td>
    <td colspan="2">
		<table border=0 cellspacing=0 cellpadding=0>
			<tr><td valign="top">□</td><td nowrap><span class="style3">寄存於　　　　　　　　　派出所</span></td>
			</tr>
			<tr><td valign="top">□</td><td nowrap><span class="style3">寄存於　　　　　　　　　鄉（鎮、市、區）
			<br>　　　　　　　　　　　　公所</span></td>
			</tr>
			<tr><td valign="top">□</td><td nowrap><span class="style3">寄存於　　　　　　　　　鄉（鎮、市、區）
			<br>　　　　　　　　　　　　公所
			<br>　　　　　　　　　　　　村（里）辦公處</span></td>
			</tr>
			<tr><td valign="top">□</td><td nowrap><span class="style3">寄存於　　　　　　　　　郵局</span></td>
			</tr>
		</table>
	</td>
    <td><span class="style3">並作送達通知書二份，一份黏貼於應受送達人住居所、事務所、營業所或其就業處所門首，一份□交由鄰居轉交或□置於該受送達處所信箱或其他適當位置，以為送達。</span></td>
  </tr>
  <tr>
    <td colspan="2" nowrap><span class="style2">送　達　人　注　意　事　項</span></td>
    <td colspan="3">
		<table border=0>
			<tr><td valign="top"><span class="style3">一、</span></td>
			<td>
			<span class="style3">依上述送達方法送達者，送達人應即將本送達證書，
			提出於交送達之行政機關附卷。</span></td>
			</tr>
			<tr><td valign="top"><span class="style3">二、</span></td>
			<td>
			<span class="style3">不能依上述送達方法送達者，送達人應製作記載該事由之報告書，
			提出於交送達之行政機關附卷，並繳回應送達之文書。</span></td>
			</tr>
		</table>
	</td>
  </tr>
</table>
<span class="style1"><strong>※請繳回<%
	If sys_City="保二總隊四大隊二中隊" Then
		If thenPasserCity=sysunit Then
			response.write thenPasserCity
		else
			response.write thenPasserCity&sysunit
		End If 
	elseIf sys_City="苗栗縣" Then
		response.write "苗栗縣警察局交通隊"

	Else
		response.write thenPasserCity&sysunit
	End If 
	%>　操作人員：<%=Sys_BillFillerMemberID%>
	<br>應到案處所：<%=Sys_STATIONNAME%></strong>
<%
if sys_City<>"台南市" and sys_City<>"基隆市" and sys_City<>"宜蘭縣" then
	response.write "<br>寄存送達之文書，應保存3個月，如未經領取，請退還交送達機關。"
end if
%>
</span>