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
	strSql="select BillTypeID,Driver,DriverAddress,DriverZip,INSURANCE,ILLEGALADDRESS,RuleVer,IllegalSpeed,RuleSpeed,Note,BillFillDate,RECORDMEMBERID from Billbase where  SN="&trim(rsbil("BillSN"))
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

	Sys_OwnerAddress="":Sys_OwnerZip=""
	Sys_BillNo=trim(rsbil("BillNo")):Sys_CarNo=trim(rsbil("CarNo"))

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		strSql="select * from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(rsbil("BillSN"))&" and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A' and dcireturnstatusid='S') and ExchangetypeID='A'"
	else
		strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"
	end if

	set rsfound=conn.execute(strSql)

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

	else
		If Sys_BillTypeID=1 Then
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))

			If ifnull(Sys_OwnerAddress) Then
				if Not rsfound.eof then
					Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
					Sys_OwnerZip=trim(rsfound("OwnerZip"))
				end if
			end if
		else
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
		end if
	end if
	rsfound.close

	if Instr(request("Sys_BatchNumber"),"N")>0 and Sys_BillTypeID=1 then
		Sys_OwnerAddress=""
	end If 

	If ifnull(Sys_OwnerAddress) Then
		if Instr(request("Sys_BatchNumber"),"N")>0 then
			strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"

		else
			strSql="select * from BillbaseDCIReturn where CarNo='"&Sys_CarNo&"' and ExchangetypeID='A'"
		end if

		set rsdata=conn.execute(strsql)

		if Instr(request("Sys_BatchNumber"),"N")>0 then
			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

		else
			If Sys_BillTypeID=1 Then
				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

				If ifnull(Sys_OwnerAddress) Then
					if Not rsdata.eof then
						Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
						Sys_OwnerZip=trim(rsdata("OwnerZip"))
					end if
				end if
			else
				if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
				if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))
			end if
		end if
		rsdata.close
	end if

	Sys_Owner=""

	strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"

	set rsfound=conn.execute(strSql)

	If Sys_BillTypeID=1 Then

		if Not rsfound.eof then Sys_Owner=trim(rsfound("Driver"))

		If ifnull(Sys_Owner) Then
			Sys_Owner=trim(rsfound("Owner"))
		end if

	else
		if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))

	End if

	If ifnull(Sys_OwnerAddress) Then
		if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	end If 

	Sys_OwnerZipName=""

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
'	If Sys_UnitLevelID=1 Then
'		strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
'	else
		strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
'	end if
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

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		MailKindType=17
	else
		MailKindType=36
	end if

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		strSql="select StoreAndSendMailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
		set rs=conn.execute(strSql)
		if Not rs.eof then Sys_MailNumber=trim(rs("StoreAndSendMailNumber"))
		if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
		if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
	else
		strSql="select MailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
		set rs=conn.execute(strSql)
		if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
		if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
		if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
	end If 

	rs.close

	if isnull(Sys_OwnerZip) or trim(Sys_OwnerZip)="" then Sys_OwnerZip="001"

	If ifnull(Sys_MailNumber) Then Sys_MailNumber=0
	if (Sys_MailDate="" or isnull(Sys_MailDate)) then Sys_MailDate=date

	DelphiASPObj.GenBillPrintBarCode trim(rsbil("BillSN")),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,400,295,MailKindType

	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,40,160

	Sys_FirstBarCode=Sys_Rule1&"-"&Sys_BillNo
	Sys_MAILCHKNUMBER=""

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		strSql="select OpenGOVReportnumber from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
		set rs=conn.execute(strSql)
		if Not rs.eof then Sys_MAILCHKNUMBER=left(trim(rs("OpenGOVReportnumber")),6)&"&nbsp;"&Mid(trim(rs("OpenGOVReportnumber")),7,6)&"&nbsp;"&Mid(trim(rs("OpenGOVReportnumber")),13,2)&"&nbsp;"&Mid(trim(rs("OpenGOVReportnumber")),15)
		rs.close
	else
		strSql="select MAILCHKNUMBER from BillMailHistory where BILLSN="&trim(rsbil("BillSN"))
		set rs=conn.execute(strSql)
		if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
		rs.close
	end if


	If Not ifnull(request("Sys_LabelKind")) and instr(Sys_Note,"郵寄日")<=0 Then
		strSQL="select Note from BillBase where sn="&trim(rsbil("BillSN"))
		set rs=conn.execute(strSQL)
			strSQL="Update BillBase set Note='"&trim(rs("Note"))&" 郵寄日:"&gInitDT(date)&" 大宗:"&Sys_MAILCHKNUMBER&"' where sn="&trim(rsbil("BillSN"))
			conn.execute(strSQL)
			strSQL="Update BillMailHistory set StoreAndSendMailNumber=null,OpenGOVReportnumber=null where sn="&trim(rsbil("BillSN"))
			conn.execute(strSQL)
		rs.close
	end if

end if
rsbil.close


%>

<div id="L178" style="position:relative;">
<div id="D178" style="position:absolute;">
<table border="0" width="750" id="table1" cellspacing="0">
	<tr>
		<td align="left" valign="top" nowrap>
			<table width="645" height="510" border="0">
				  <tr>
					<td width="141" height="69" valign="top">&nbsp;</td>
					<td colspan="2">&nbsp;</td>
					<td rowspan="2" align="right" valign="top"><br>   	</td>
				  </tr>
				  <tr>
					<td height="41" align="left" valign="top">　　　　<img src=<%="""../BarCodeImage/"&Sys_BillNo&"_1.jpg"""%> hspace="0" vspace="0" align="top"><br>　　　　<span class="style7"><%=Sys_FirstBarCode%></span>
					</td>
					<td colspan="3" align="left" valign="top" width="300"><span class="style3"><%=Sys_OwnerZip%><br>
					<%
						If instr(Sys_DriverHomeAddress,"@") >0 Then
							Response.Write funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)
						else
							Response.Write funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,16,1)
						End if
					%></span></td>
				  </tr>
				  <tr>
					<td>&nbsp;</td>
					<td colspan="2"><span class="style3"><%=funcCheckFont(Sys_Owner,16,1)%>　台啟</span></td>
					<td width="92">&nbsp;</td>
				  </tr>
				  <tr>
					<td>&nbsp;</td>
					<td width="160" class="style4" align="center">
					  大宗郵資已付掛號函件<br>
					第<%=Sys_MailNumber%>號</td>
					<td width="23" align="center">&nbsp;</td>
					<td>&nbsp;</td>
				  </tr>
				  <tr>
					<td>&nbsp;</td>
					<td align="center"><div align="left"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_2.jpg"""%>><br>
						<%=Sys_MAILCHKNUMBER%></div></td>
					<td align="center">&nbsp;</td>
					<td align="right" nowrap><p>&nbsp;</p>
					<p class="style8"><%
						if sys_City="台中市" then
							if Instr(request("Sys_BatchNumber"),"N")>0 then response.write "行政文書"
						else
							If Not ifnull(request("Sys_LabelKind")) Then
								Sys_StationID=request("Sys_LabelKind")
							elseIf Not ifnull(request("Sys_LabelUpdate")) Then
								Sys_StationID=request("Sys_LabelUpdate")
							End if
							response.write Sys_StationID&"<br><br><br><br><br><font size=2>"&request("Sys_SendKind")&"</font>"
							response.write "<br><font size=2 color=""red"">本單如已繳納，請向監理(裁決)<br>單位查詢，以確認是否繳結。"
						end if
					%></p></td>
				  </tr>
				  <tr>
					<td height="98" valign="top" nowrap colspan="4">　　　　<span class="style7">應到案處所：<%=Sys_STATIONNAME%></span><br>
					　　　　<span class="style7">應到案處所電話：<%=Sys_StationTel%></span><br>
					　　　　<span class="style7">舉發單位：<%
																'If instr(Sys_BillUnitName,"分隊")>0 Then
																	Response.Write Sys_BillUnitName
																'else
																'	Response.Write sysunit
																'end if%>&nbsp;</span></td>
				  </tr>
			</table>
		</td>
	<tr>
	<br>
		<td align="center" valign="top">
		<align="center">
		<font face="標楷體"><font size="5">臺中市政府警察局&nbsp;交通警察大隊送達證書</font>
		<font size="2"><%
		tmpleft = 30
		tmpwidth = 710
			'If instr(Sys_BillUnitName,"分隊")>0 Then
  			  Response.Write Sys_BillUnitName
			'else
			'	Response.Write sysunit
			'end if
			if Instr(request("Sys_BatchNumber"),"N")>0 then
				Response.Write "（二次送達）"
			end if
		%></font></font>
		<table border="0" width="750" id="table2" cellspacing="0">
			<tr>
				<td align="left" valign="top">
				<div id="D01" style="position:absolute; left:<%=tmpleft%>px;top:555px;">
				<table border="1" width="<%=tmpwidth%>" id="table3" height="181" cellspacing="0">
					<tr>
						<td   colspan="2">
						<p align="center"><font face="標楷體">&nbsp;受送達人名稱姓名地址&nbsp;</font></td>
						<td  colspan="2"><font face="標楷體" size="2"><%=funcCheckFont(Sys_Owner,20,1)%>&nbsp;&nbsp; 
						<%=Sys_OwnerZip&"　"&Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,20,1)%></font></td>
					</tr>
					<tr>
						<td  colspan="2"><font face="標楷體">&nbsp;文&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;號&nbsp;</font></td>
						<td colspan="2" height="21"><font face="標楷體">　中市警交字第&nbsp;<%=Sys_BillNo%>&nbsp;號&nbsp;&nbsp;&nbsp;&nbsp;<%
					if sys_City<>"高雄市" then
						response.write Sys_CarNo & "&nbsp;&nbsp;&nbsp;&nbsp;" & Mid(Sys_MailNumber,1,6)
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
				</div>
				</td>
			</tr>

		</font>
				<td>
				<div id="D01" style="position:absolute; left:<%=tmpleft%>px;top:735px;">
				<table border="1" width="<%=tmpwidth%>" id="table5"  cellspacing="0">
					<tr>
						<td  class="style110" colspan="2">
						<p align="center"><font face="標楷體">送&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;達&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;方&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;式</font></p></td>
					</tr>
					<tr>
						<td  class="style110" colspan="2">
						<p align="center"><font face="標楷體">由&nbsp;&nbsp;&nbsp;&nbsp;送&nbsp;&nbsp;&nbsp;&nbsp;達&nbsp;&nbsp;&nbsp;&nbsp;人&nbsp;&nbsp;&nbsp;&nbsp;在&nbsp;&nbsp;&nbsp;&nbsp;□&nbsp;&nbsp;&nbsp;&nbsp;上&nbsp;&nbsp;&nbsp;&nbsp;劃&nbsp;&nbsp;&nbsp;&nbsp;ｖ&nbsp;&nbsp;&nbsp;&nbsp;選&nbsp;&nbsp;&nbsp;&nbsp;記</font></p></td>
					</tr>
					<tr>
						<td align="left" class="style110" width="47%" height="12"><font face="標楷體" size="1">&nbsp;□ 已將文書交與應受送達人</font></td>
						<td align="left"  class="style110" width="51%" height="12"><font face="標楷體" size="1">&nbsp;□本人
	　　　　　　　　　　　　　　　　　　　　（簽名或蓋章）</font></td>
					</tr>
					<tr>
						<td align="left"  width="47%" class="style110"><font face="標楷體" size="1">&nbsp;□未獲會晤本人，已將文書交與有辨別事理能力之同居<br>
						&nbsp;&nbsp;&nbsp;&nbsp;人或受雇人或應送達處所之接收郵件人員</font></td>
						<td align="left" width="51%" class="style110"><font face="標楷體" size="1">□同居人□受雇人□應送達處所之接收郵件人員</font><br>
						<font face="標楷體">　　　　　　　　　　　　　　　（簽名或蓋章）</font></td>
					</tr>
					<tr>
						<td  class=style110 height="24" align="left" valign="top" width="47%">
						<font face="標楷體" size="1">&nbsp;□ 應受送達之本人、同居人、或應送達處所之接收郵件人員收領，但拒絕或不能簽名、蓋章或按指印者，由送達人記明其事由於右欄：</font></td>
						<td  class=style110 height="24" align="left" valign="top" width="51%">
						<font face="標楷體" size="1">□送達人填記：</font></td>
					</tr>
					<tr>
						<td  class=style110 align="left" valign="top" width="47%">
						<font face="標楷體" size="1">&nbsp;□ 應受送達之本人、同居人、受雇人或應送達處所接收 郵件人員無法律上之理由拒絕收領經送達人將文書留置於
						送達處所，以為送達：</font></td>
						<td  class=style110 align="left" valign="top" width="51%"><font face="標楷體" size="1">□本人□同居人□受雇人□應送達處所之接收郵件人員<br>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
						拒絕收領</font></td>
					</tr>
				</table>
				</div>
				<tr>
				</td>
			</tr>
		</table>
		<div id="D0102"  style="position:absolute; left:<%=tmpleft%>px;top:876px;">
		<table border="1" width="<%=tmpwidth%>" id="table6"   cellspacing="0">
			<tr>
				<td  class="style110" width="294" align="left" valign="top">
				<font face="標楷體" size="1">&nbsp;□ 未獲會晤本人亦無受領文書之同居人、受雇人或應送達處所接<br>&nbsp;&nbsp;&nbsp;受郵件人員，已將該送達文書：
				<br>&nbsp;□ 應受送達之本人、同居人、受雇人或應送達處所接收郵件人員<br>&nbsp;&nbsp;&nbsp;無正當理由拒絕收領，並有難達留置情事，已將該送達文書：
				</font>
				</td>
				<td  class="style110"  class="style110" width="240" align="left" valign="top">
				<font face="標楷體" size="1">&nbsp;□ 寄存於　　　　派出所<br>&nbsp;□ 寄存於　　　　鄉（鎮、市、區）公所
				<br>&nbsp;□ 寄存於　　　　鄉（鎮、市區）公所
				<br>　　　　　　　　&nbsp; 村（里）辦公處
				<br>&nbsp;□ 寄存於　　　　郵局</font>
				</td>

				<td  class="style110" align="left" class="style110"><font face="標楷體" size="1">並作送達通知書二份，一份黏貼於應
				受送達人住居所、事務所、營業所或
				其就業處所門首，一份 □交由鄰居轉
				交或□ 置於該受送達處所信箱或其他
				適當位置，以為送達。
				</font>
				</td>
			</tr>
		</table>
		</div>		
		</td>
		<tr>
		<td>
		<div id="D03"  style="position:absolute; left:<%=tmpleft%>px;top:945px;">
		<table border="1" width="<%=tmpwidth%>" cellspacing="0">
		<td><font face="標楷體">送達人注意事項</font></td>
		<td class="style110"><font face="標楷體" size="1">一、依上述送達方法送達者，送達人應即將本送達證書，提出於交送達之行政機關附卷。</font><br>
		<font face="標楷體" size="1">二、不能依上述送達方式送達者，送達人應製作記載該事由之報告書，提出於</font>
		<br><font face="標楷體" size="1">&nbsp;&nbsp;&nbsp; 交送達之行政機關附卷，並繳回應送達之文書。</font></td>
	</tr>
	</table>
	</div>

		</td>

	</tr>
</table>

<div id="Layer1" style="position:absolute; left:<%=tmpleft%>px; top:990px; z-index:1">
<font face="標楷體"  size="2">請繳回：臺中市政府警察局交通警察大隊&nbsp;地址：407&nbsp;台中市西屯區文心路二段588號</font>
</div>

<div id="Layer1" style="position:absolute; left:575px; top:995px; z-index:5"><img  src=<%="""../BarCodeImage/"&Sys_BillNo&"_1.jpg"""%>>
</div>

</Div>
				