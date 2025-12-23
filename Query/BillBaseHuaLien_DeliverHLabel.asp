<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&strBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_OwnerZip=""
Sys_OwnerZipName=""

If Not rsbil.eof Then
strSql="select BillTypeID,Driver,DriverAddress,DriverZip,INSURANCE,ILLEGALADDRESS,RuleVer,IllegalSpeed,RuleSpeed,Note,BillFillDate,RECORDMEMBERID from Billbase where SN="&trim(rsbil("BillSN"))
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
if Not rs.eof then Sys_DriverZip=trim(rs("DriverZip"))
if Not rs.eof then Sys_DriverAddress=trim(rs("DriverAddress"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close
Sys_OwnerAddress=""
strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)
If Sys_BillTypeID=1 Then
	if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver"))
else
	if Not rsFound.eof then Sys_Owner=trim(rsfound("Owner"))
End if
if Instr(request("Sys_BatchNumber"),"N")>0 then
	if Not rsFound.eof then Sys_OwnerAddress=trim(Sys_DriverAddress)
	if Not rsFound.eof then Sys_OwnerZip=trim(Sys_DriverZip)
If ifnull(Sys_OwnerAddress) Then
	strSql="select a.Driver,Decode(b.DriverAddress,null,a.DriverHomeAddress,b.DriverAddress) DriverHomeAddress,Decode(b.DriverAddress,null,a.DriverHomeZip,b.DriverZip) DriverHomeZip,Decode(b.OwnerAddress,null,a.OwnerAddress,b.OwnerAddress) OwnerAddress,Decode(b.OwnerAddress,null,a.OwnerZip,b.OwnerZip) OwnerZip,Decode(b.OwnerAddress,null,a.Owner,b.Owner) Owner from (select CarNo,Owner,Driver,DriverHomeAddress,DriverHomeZip,OwnerAddress,OwnerZip from BillbaseDCIReturn where Carno in (select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and ExchangetypeID='A') and ExchangetypeID='A') a,(Select Owner,OwnerAddress,OwnerZip,DriverAddress,DriverZip,CarNo from Billbase where sn="&trim(rsbil("BillSN"))&")b where a.Carno=b.Carno(+)"
	set rsdata=conn.execute(strsql)
	If Sys_BillTypeID=1 Then
		if Not rsdata.eof then Sys_Owner=trim(rsdata("Driver"))
	else
		if Not rsdata.eof then Sys_Owner=trim(rsdata("Owner"))
	End if

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
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	End if
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
	if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
	if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close
Sys_OwnerAddress=replace(replace(Sys_OwnerAddress&"","臺","台"),Sys_OwnerZipName,"")
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
strSql="select distinct BillFillerMemberID,DriverSex,DriverBirth,IllegalDate,DealLineDate,RECORDMEMBERID,BillFillerMemberID,IMAGEFILENAME,IMAGEPATHNAME,IMAGEFILENAMEB,BILLMEMID1 from BillBase where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"'"
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
if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
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
if Not unit.eof then Sys_UnitTel=trim(unit("Tel"))
if Not unit.eof then Sys_UnitAddress=trim(unit("Address"))
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

strSql="select MailNumber,StoreAndSendMailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))

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
if isnull(Sys_DriverHomeZip) or trim(Sys_DriverHomeZip)="" then Sys_DriverHomeZip="001"
if isnull(Sys_OwnerZip) or trim(Sys_OwnerZip)="" then Sys_OwnerZip="001"

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
Sys_MailNumber=0
DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,30,160,0

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&replace(trim(rsbil("BillSN")),"","0")

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
end if
rsbil.close


%>

<div id="L178" style="position:relative;">
<div id="D178" style="position:absolute;">
<table border="0" width="750" id="table1" height="1030" cellspacing="0">
	<tr>
		<td height="402" align="left" valign="top" nowrap>　<p><b><font size="5">
		　　<%
				If instr(Sys_BillUnitName,"分隊")>0 Then
					Response.Write Sys_BillUnitName
				else
					Response.Write sysunit
				end if
			%></font></b></p>
		<p><b><font size="5" face="標楷體">　　　<%
			If instr(Sys_BillUnitName,"分隊")>0 Then
				Response.Write Sys_BillUnitAddress
			else
				Response.Write Sys_UnitAddress
			end if
		%></font></b></p>
		<p>　</p>
		<p>　</p>
		<p><font size="4" face="標楷體"><b>收件人：<%=funcCheckFont(Sys_Owner,24,1)%>　　台啟</b></font></p>
		<p>　</p>
		<p><font size="4" face="標楷體"><b>戶籍地：<%=Sys_OwnerZip&"　"&Sys_OwnerZipName&funcCheckFont(InstrAdd(Sys_OwnerAddress,20),24,1)%></b></font></p>
		<p></p>
		<p></p>
		<p></p>
		<p>　</p>
		<p>　</p>
		<p><b><font size="4">　</font><font size="6"  face="標楷體">行政文書郵件<%
				if Instr(request("Sys_BatchNumber"),"N")>0 then
					Response.Write "（第二次郵寄）"
				end if
				%></font></b>
		<p>　</p>
		<br>　
		</td>
	</tr>
	<tr>
		<td align="center" valign="top" height="514">
		<align="center">
		<font face="標楷體"><font size="5">高雄市政府警察局</font>　<font size="4"><%
			If instr(Sys_BillUnitName,"分隊")>0 Then
				Response.Write Sys_BillUnitName
			else
				Response.Write sysunit
			end if
			if Instr(request("Sys_BatchNumber"),"N")>0 then
				Response.Write "（二次送達）"
			end if
		%></font><font size="5">送達證書</font></font>
		<table border="0" width="750" id="table2" cellspacing="0">
			<tr>
				<td align="left" valign="top">
				<table border="1" width="748" id="table3" height="181" cellspacing="0">
					<tr>
						<td   colspan="2">
						<p align="center"><font face="標楷體">&nbsp;受送達人名稱姓名地址&nbsp;</font></td>
						<td  colspan="2"><font face="標楷體" size="2"><%=funcCheckFont(Sys_Owner,20,1)%>&nbsp;&nbsp; 
						<%=Sys_OwnerZip&"　"&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,20,1)%></font></td>
					</tr>
					<tr>
						<td  colspan="2"><font face="標楷體">&nbsp;文&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;號&nbsp;</font></td>
						<td colspan="2" height="21"><font face="標楷體">　高市警交相字第&nbsp;<%=Sys_BillNo%>&nbsp;號&nbsp;&nbsp;&nbsp;&nbsp;<%
					if sys_City<>"高雄市" then
						response.write Sys_CarNo
					end if 
						%></font></td>
					</tr>
					<tr>
						<td  colspan="2"><font face="標楷體">&nbsp;送&nbsp;達&nbsp;文&nbsp;書&nbsp;(&nbsp;含&nbsp;案&nbsp;由)&nbsp;</font></td>
						<td colspan="2"><font face="標楷體">　舉發違反道路交通管理事件通知單　(附違規採證相片)</font></td>
					</tr>
					<tr>
						<td width="14%"><font face="標楷體">原寄郵局日戳</font></td>
						<td width="14%"><font face="標楷體">送達郵局日戳</font></td>
						<td  width="57%">
						<p align="center"><font face="標楷體">&nbsp;送&nbsp;達&nbsp;處&nbsp;所&nbsp;(&nbsp;由&nbsp;送&nbsp;達&nbsp;人&nbsp;填&nbsp;記&nbsp;)</font></td>
						<td  class="style110" width="12%" height="22"><font face="標楷體">送達人簽章</font></td>
					</tr>
					<tr>
						<td  width="14%" rowspan="3">　</td>
						<td  width="14%" rowspan="3">　</td>
						<td  width="57%"><font face="標楷體">&nbsp;□同上記載地址□改送：</font></td>
						<td  width="12%" rowspan="3">　</td>
					</tr>
					<tr>
						<td width="57%"><font face="標楷體">&nbsp;送&nbsp;達&nbsp;時&nbsp;間&nbsp;（&nbsp;由&nbsp;送&nbsp;達&nbsp;人&nbsp;填&nbsp;記&nbsp;）</font></td>
					</tr>
					<tr>
						<td width="57%"><font face="標楷體">&nbsp;中華民國　　&nbsp;年&nbsp;　　月&nbsp;　日&nbsp;　　午&nbsp;　　時&nbsp;　　分&nbsp;</font></td>
					</tr>
				</table>
				</td>
			</tr>

		</font>
				<td>
				<div id="D01" style="position:absolute; left:2px;top:748px;">
				<table border="1" width="748" id="table5"  cellspacing="0">
					<tr>
						<td  class="style110" colspan="2">
						<p align="center"><font face="標楷體">送&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;達&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;方&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;式</font></p></td>
					</tr>
					<tr>
						<td  class="style110" colspan="2">
						<p align="center"><font face="標楷體">由&nbsp;&nbsp;&nbsp;&nbsp;送&nbsp;&nbsp;&nbsp;&nbsp;達&nbsp;&nbsp;&nbsp;&nbsp;人&nbsp;&nbsp;&nbsp;&nbsp;在&nbsp;&nbsp;&nbsp;&nbsp;□&nbsp;&nbsp;&nbsp;&nbsp;上&nbsp;&nbsp;&nbsp;&nbsp;劃&nbsp;&nbsp;&nbsp;&nbsp;ｖ&nbsp;&nbsp;&nbsp;&nbsp;選&nbsp;&nbsp;&nbsp;&nbsp;記</font></p></td>
					</tr>
					<tr>
						<td align="left" class="style110" width="47%" height="12"><font face="標楷體" size="2">&nbsp;□ 已將文書交與應受送達人</font></td>
						<td align="left"  class="style110" width="51%" height="12"><font face="標楷體" size="2">&nbsp;□本人
						　　　　　　　　　　　　　（簽名或蓋章）</font></td>
					</tr>
					<tr>
						<td align="left"  width="47%" class="style110"><font face="標楷體" size="2">&nbsp;□未獲會晤本人，已將文書交與有辨別事理能力之同居<br>
						&nbsp;&nbsp;&nbsp;&nbsp;人或受雇人或應送達處所之接收郵件人員</font></td>
						<td align="left" width="51%" class="style110"><font face="標楷體">□同居人□受雇人□應送達處所之接收郵件人員</font><br>
						<font face="標楷體">　　　　　　　　　　　　　　　 　（簽名或蓋章）</font></td>
					</tr>
					<tr>
						<td  class=style110 height="24" align="left" valign="top" width="47%">
						<font face="標楷體" size="2">&nbsp;□ 應受送達之本人、同居人或受雇人收領後，拒絕或不<br>
						　&nbsp; 能簽名或蓋章者，由送達人記明其事由</font></td>
						<td  class=style110 height="24" align="left" valign="top" width="51%">
						<font face="標楷體">□送達人填記：</font></td>
					</tr>
					<tr>
						<td  class=style110 align="left" valign="top" width="47%">
						<font face="標楷體" size="2">&nbsp;□ 應受送達之本人、同居人、受雇人或應送達處所接收<br>&nbsp;　 郵件人員無正當理由拒絕收領經送達人將文書留置於
						<br>&nbsp;　 送達處所，以為送達</font></td>
						<td  class=style110 align="left" valign="top" width="51%"><font face="標楷體">□本人□同居人□受雇人□應送達處所之接收郵件人員<br><br>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
						拒絕收領</font></td>
					</tr>
				</table>
				</div>
				<tr>
				</td>
			</tr>
		</table>
		<div id="D0102"  style="position:absolute; left:2px;top:902px;">
		<table border="1" width="748" id="table6"   cellspacing="0">
			<tr>
				<td  class="style110" width="294" align="left" valign="top">
				<font face="標楷體" size="2">&nbsp;□ 未獲會晤本人亦無受領文書之同居人、受雇<br>　&nbsp; 人或應送達處所接受郵件人員，已將該送達<br>&nbsp;&nbsp;&nbsp; 文書：
				<br>&nbsp;□ 應受送達之本人、同居人、受雇人或應送達<br>&nbsp;&nbsp;&nbsp; 處所接收郵件人員無正當理由拒絕收領，並
				<br>&nbsp;&nbsp;&nbsp; 有難達留置情事，已將該送達文書：　</font></td>
				<td  class="style110"  class="style110" width="266" align="left" valign="top">
				<font face="標楷體" size="2">&nbsp;□ 寄存於　　　　派出所<br>&nbsp;□ 寄存於　　　　鄉（鎮、市、區）公所
				<br>&nbsp;□ 寄存於　　　　鄉（鎮、市區）公所
				<br>　　　　　　　　&nbsp; 村（里）辦公處
				<br>&nbsp;□ 寄存於　　　　郵局</font></td>
				<td  class="style110" align="left" class="style110"><font face="標楷體" size="2">並作送達通知書二份，一
				<br>份黏貼於應受送達人住
				<br>居所、事務所、營業所或
				<br>其就業處所門首，一份□
				<br>交由鄰居轉交或□置於
				<br>該受送達處所信箱或其
				<br>他適當位置，以為送達
				</font>
				</td>
			</tr>
		</table>
		</div>		
		</td>
		<tr>
		<td>
		<div id="D03"  style="position:absolute; left:2px;top:997px;">
		<table border="1" width="748" cellspacing="0">
	<tr>
		<td><font face="標楷體">送達人注意事項</font></td>
		<td class="style110"><font face="標楷體" size="2">一、依上述送達方法送達者，送達人應即將本送達證書，提出於交送達之行政機關附卷。</font><br>
		<font face="標楷體" size="2">二、不能依上述送達方式送達者，送達人應製作記載該事由之報告書，提出於交送達之行政機關</font>
		<br><font face="標楷體" size="2">&nbsp;&nbsp;&nbsp; 附卷，並繳回應送達之文書。</font></td>
	</tr>
	</table>
	</div>

		</td>

	</tr>
</table>

<div id="Layer1" style="position:absolute; left:10px; top:1044px; height:36px; z-index:1">
<font face="標楷體"  size="2">請繳回：<%
	If instr(Sys_BillUnitName,"分隊")>0 Then
		Response.Write Sys_BillUnitAddress
	else
		Response.Write Sys_UnitAddress
	End if
%></font>
</div>

<div id="Layer1" style="position:absolute; left:450px; top:338px; height:36px; z-index:5"><img src=<%="""../BarCodeImage/"&Sys_BillNo&".jpg"""%>>
</div>

<div id="Layer1" style="position:absolute; left:450px; top:380px; height:36px; z-index:5"><font face="標楷體" size="2"><b><%=Sys_BillNo%></b></font>
</div>

<div id="Layer1" style="position:absolute; left:550px; top:1043px; height:26px; z-index:5"><img src=<%="""../BarCodeImage/"&Sys_BillNo&".jpg"""%>>
</div>

<div id="Layer5" style="position:absolute; left:450px; top:410px; height:26px; z-index:5"><table border="1" cellspacing="0"><td><font size="5"  face="標楷體"><b>車籍變更，請依規定<br>至管轄監理機關辦理異動</b></font></td></table>
</div>



</Div>
				