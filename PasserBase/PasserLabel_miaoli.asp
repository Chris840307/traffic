<%
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
	End If 
	
	
	strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&rsSql("memberstation")&"'"
	set rsUnit=conn.execute(strSQL)
	Sys_UnitID=trim(rsUnit("UnitID"))
	Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
	Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
	rsUnit.close

	If Sys_UnitLevelID=1 Then
		strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"

		If sys_City="台南市" and trim(Sys_UnitID)="07A7" Then
			strSQL="select * from UnitInfo where UnitID='0707'"
		End if
		
	else
		strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
	end if

	set unit=conn.Execute(strSQL)
	If Not unit.eof Then
		theUnitID=trim(unit("UnitID"))
		if trim(unit("UnitName"))<>"" and not isnull(unit("UnitName")) then
			theUnitName=replace(replace(trim(unit("UnitName")),"台","臺"),"交通組","")
		end if 
		theSubUnitSecBossName=trim(unit("SecondManagerName"))
		theBigUnitBossName=trim(unit("ManageMemberName"))
		theContactTel=trim(unit("Tel"))
		theBankAccount=trim(unit("BankAccount"))
		theBankName=trim(unit("BankName"))
		theUnitAddress=trim(unit("Address"))
	end if
	unit.close 
end if

strUInfo="select * from UnitInfo where UnitID='"&trim(rsSql("BillUnitID"))&"'"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	thenBillUnitName=trim(rsUInfo("UnitName"))
end if
rsUInfo.close
set rsUInfo=nothing

strSql="select confiscate from PasserConfiscate where BIllSN="&trim(rsSql("SN"))
set rsfast=conn.execute(strsql)
fastring=""
while Not rsfast.eof
	if trim(fastring)<>"" then fastring=fastring&","
	fastring=fastring&rsfast("confiscate")
	rsfast.movenext
wend
rsfast.close

strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
Sys_UnitID=trim(rsUnit("UnitID"))
Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
rsUnit.close

chk_UnitLevelID=""
strSQL="select UnitLevelID from UnitInfo where UnitID in(select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"')"
set rsunit=conn.execute(strSQL)
If Not rsunit.eof Then
	chk_UnitLevelID=trim(rsunit("UnitLevelID"))
End if
rsunit.close

strMem="select MANAGEMEMBERNAME,secondmanagername from UnitInfo where UnitID='"&Sys_UnitID&"'"
set rsUnit=conn.execute(strMem)
if Not rsUnit.eof then 
	If sys_City="高雄市" or sys_City="宜蘭縣" or sys_City="台南市" then
		If not ifnull(rsUnit("secondmanagername")) Then
			MemUnitName="分局長 "&rsUnit("secondmanagername")
		End If 
		
	elseif sys_City="台中市" Then
		If chk_UnitLevelID = "1" Then
			MemUnitName="局長 "&rsUnit("MANAGEMEMBERNAME")
		else
			MemUnitName="分局長 "&rsUnit("secondmanagername")
		End if
	Else
		MemUnitName=rsUnit("MANAGEMEMBERNAME")
	End If 
End If 
rsUnit.close

%>
<table width="645" border="0">
	<tr>
		<td height="120" colspan="3"></td>
	</tr>
	<tr>
		<td width="50"></td>
		<td colspan="2" align="center" class="style25">違反道路交通管理事件裁決書&nbsp;&nbsp;行政文書</td>
	</tr>
	<tr>
		<td></td>
		<td colspan="2" align="center" class="style22"><%=theUnitAddress%>(<%=theUnitName%>)</td>
	</tr>
	<tr>
		<td colspan="3" height="30"></td>
	</tr>
	<tr>
		<td></td>
		<td align="right" class="style22">地&nbsp;&nbsp;&nbsp;&nbsp;址：</td>
		<td width="400" class="style22"><%=trim(rsSql("DriverZip"))&trim(rsSql("DriverAddress"))%></td>
	</tr>
	<tr>
		<td></td>
		<td align="right" class="style22">收&nbsp;件&nbsp;人：</td>
		<td class="style22"><%=trim(rsSql("Driver"))%></td>
	</tr>
</table>

