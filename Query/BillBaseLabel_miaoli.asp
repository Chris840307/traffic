<%
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)

Sys_DriverHomeZip=""
Sys_DriverZipName=""
Sys_DriverHomeAddress=""

strSql="select * from Billbase where SN="&trim(rsbil("BillSN"))

set rs=conn.execute(strSql)
Sys_IllegalSpeed="":Sys_RuleSpeed="":Sum_Level=0
if Not rs.eof then Sys_BillNo=trim(rs("BillNo"))
if Not rs.eof then Sys_CarNo=trim(rs("CarNo"))
if Not rs.eof then Sys_Rule1=trim(rs("Rule1"))
if Not rs.eof then Sys_Rule2=trim(rs("Rule2"))
if Not rs.eof then Sys_Level1=trim(rs("FORFEIT1"))
if Not rs.eof then Sys_Level2=trim(rs("FORFEIT2"))
Sum_Level=funTnumber(Sys_Level1)+funTnumber(Sys_Level2)
if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_IllegalSpeed=trim(rs("IllegalSpeed"))
if Not rs.eof then Sys_RuleSpeed=trim(rs("RuleSpeed"))
if Not rs.eof then Sys_Note=trim(rs("Note"))
if Not rs.eof then Sys_RecordMemberID=trim(rs("RecordMemberID"))
if Not rs.eof then
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
else
	sys_Date=split(gArrDT(trim("")),"-")
end if
rs.close
rsbil.close

Sys_DriverHomeAddress="":Sys_DriverHomeZip="":Sys_Driver=""


'if Instr(request("Sys_BatchNumber"),"N")>0 and Sys_BillTypeID=2 Then
'	strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='N'"
'else
	strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"
'end if

set rsfound=conn.execute(strSql)

Sys_Driver=trim(rsfound("Owner"))

if Instr(request("Sys_BatchNumber"),"N")>0 then
	if Not rsfound.eof then Sys_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
	if Not rsfound.eof then Sys_DriverHomeZip=trim(rsfound("DriverHomeZip"))
else
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("DriverHomeZip"))
	else
		if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("OwnerZip"))
	End if
end if
rsfound.close

If ifnull(Sys_DriverHomeAddress) or ifnull(Sys_Driver) Then

	strSql="select * from BillbaseDCIReturn where Carno in(select carno from dcilog where BillSN="&trim(PBillSN(i))&" and ExchangetypeID='A') and ExchangetypeID='A'"

	set rsdata=conn.execute(strsql)

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		If Sys_BillTypeID=2 Then
			if Not rsdata.eof then 
				Sys_DriverHomeAddress=trim(rsdata("DriverHomeAddress"))
				Sys_DriverHomeZip=trim(rsdata("DriverHomeZip"))
			end if
		end if
	else
		If Sys_BillTypeID=1 Then
			if Not rsdata.eof then Sys_DriverHomeAddress=trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then Sys_DriverHomeZip=trim(rsdata("DriverHomeZip"))
		else
			if Not rsdata.eof then Sys_DriverHomeAddress=trim(rsdata("OwnerAddress"))
			if Not rsdata.eof then Sys_DriverHomeZip=trim(rsdata("OwnerZip"))
		End if
	end if
	rsdata.close
end if

If not ifnull(Sys_DriverHomeAddress) Then
	Sys_DriverHomeAddress=replace(trim(Sys_DriverHomeAddress),"台","臺")
end if

Sys_Driver=""
strSql="select * from BillbaseDCIReturn where BillNo='"&Sys_BillNo&"' and CarNo='"&Sys_CarNo&"' and ExchangetypeID='W'"

set rsfound=conn.execute(strSql)
If Sys_BillTypeID=1 Then

	if Not rsfound.eof then Sys_Driver=trim(rsfound("Driver"))

	If ifnull(Sys_Driver) Then
		Sys_Driver=trim(rsfound("Owner"))
	end if

else
	if Not rsfound.eof then Sys_Driver=trim(rsfound("Owner"))
End if

If ifnull(Sys_DriverHomeAddress) Then
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("DriverHomeZip"))

		If ifnull(Sys_DriverHomeAddress) Then
			if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("OwnerAddress"))
			if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("OwnerZip"))
		end if
	else
		if Not rsFound.eof then Sys_DriverHomeAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_DriverHomeZip=trim(rsfound("OwnerZip"))
	End if
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then
	Sys_DriverZipName=replace(trim(rszip("ZipName")),"台","臺")
	Sys_DriverHomeAddress=replace(Sys_DriverHomeAddress,Sys_DriverZipName,"")
end if
rszip.close

strSql="select a.LoginID,a.ChName,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_RecordMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitName=replace(trim(mem("UnitName")),"台","臺")
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitAddr=trim(mem("Address"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

%>
<table width="645" border="0">
	<tr>
		<td height="120" colspan="3"></td>
	</tr>
	<tr>
		<td width="50"></td>
		<td colspan="2" align="center" class="style25">違反道路交通管理事件通知書&nbsp;&nbsp;行政文書</td>
	</tr>
	<tr>
		<td></td>
		<td colspan="2" align="center" class="style22"><%=Sys_UnitAddr%>(<%=Sys_UnitName%>)</td>
	</tr>
	<tr>
		<td colspan="3" height="30"></td>
	</tr>
	<tr>
		<td></td>
		<td align="right" class="style22">地&nbsp;&nbsp;&nbsp;&nbsp;址：</td>
		<td width="400" class="style22"><%=Sys_DriverHomeZip & funcCheckFont(Sys_DriverZipName&Sys_DriverHomeAddress,20,1)%></td>
	</tr>
	<tr>
		<td></td>
		<td align="right" class="style22">收&nbsp;件&nbsp;人：</td>
		<td class="style22"><%=funcCheckFont(Sys_Driver,20,1)%></td>
	</tr>
</table>

