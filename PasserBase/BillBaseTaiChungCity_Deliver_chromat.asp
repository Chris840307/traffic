<%
strSql="select * from PasserBase where SN="&trim(BillSN(i))
set rs=conn.execute(strSql)
if Not rs.eof then
	Sys_BillTypeID=trim(rs("BillTypeID"))
	Sys_BillNo=trim(rs("BillNo"))
	Sys_DOUBLECHECKSTATUS=trim(rs("DOUBLECHECKSTATUS"))
	Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
	Sys_RuleVer=trim(rs("RuleVer"))
	Sys_Note=trim(rs("Note"))
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
	Sys_Driver=trim(rs("Driver"))
	Sys_DriverID=trim(rs("DriverID"))
	Sys_DriverHomeAddress=trim(rs("DriverAddress"))
	Sys_DriverHomeZip=trim(rs("DriverZip"))
	Sys_Rule1=trim(rs("Rule1"))
	Sys_Rule2=trim(rs("Rule2"))
	Sys_Sex=""
	if Not rs.eof then
		If not ifnull(Trim(rs("DriverID"))) Then
			If Mid(Trim(rs("DriverID")),2,1)="1" Then
				Sys_Sex="男"
			elseif Mid(Trim(rs("DriverID")),2,1)="2" Then
				Sys_Sex="女"
			End if
		End if
	end if
	Sys_RecordMemberID=trim(rs("RECORDMEMBERID"))
	Sys_IllegalDate=split(gArrDT(trim(rs("IllegalDate"))),"-")
	Sys_IllegalDate_h=hour(trim(rs("IllegalDate")))
	Sys_IllegalDate_m=minute(trim(rs("IllegalDate")))
	Sys_DealLineDate=split(gArrDT(trim(rs("DealLineDate"))),"-")
	DealLineDate=trim(rs("DealLineDate"))
	Sys_DriverBirth=split(gArrDT(trim(rs("DriverBirth"))),"-")
	Sys_BillFillerMemberID=0
	Sys_Billmem1ID=trim(rs("BILLMEMID1"))
	Sys_STATIONNAME=trim(rs("MemberStation"))
end if
rs.close

Sys_UrgeDate=""
If not ifnull(request("BillUrge")) Then
	strSQL="select OpenGovNumber,UrgeDate from PasserUrge where BillSN="&trim(BillSN(i))
	set rsjude=conn.execute(strSQL)
	If not rsjude.eof Then
		Sys_OpenGovNumber=trim(rsjude("OpenGovNumber"))
		Sys_UrgeDate=split(gArrDT(trim(rsjude("UrgeDate"))),"-")
	End if
	rsjude.close

else

	strSQL="select OpenGovNumber,JudeDate from PasserJude where BillSN="&trim(BillSN(i))
	set rsjude=conn.execute(strSQL)
	If not rsjude.eof Then
		Sys_OpenGovNumber=trim(rsjude("OpenGovNumber"))
		Sys_UrgeDate=split(gArrDT(trim(rsjude("JudeDate"))),"-")
	End if
	rsjude.close
End if

If ifnull(Sys_OpenGovNumber) Then
	Sys_OpenGovNumber=trim(Sys_BillNo)
	Sys_UrgeDate=split(gArrDT(date),"-")
End if

strUnit="select UnitName from UnitInfo where UnitID='"&Sys_STATIONNAME&"'"
set rsUnit=conn.execute(strUnit)
If Not rsUnit.eof Then
	Sys_STATIONNAME=trim(rsUnit("UnitName"))
End if
rsUnit.close
Sys_Level1=0:Sys_Level2=0
strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VERSION=(select value from apconfigure where ID=3)"
set rsRule1=conn.execute(strRule1)
if not rsRule1.eof then
	If DateDiff("d",CDate(date),trim(DealLineDate))>-1 Then 
	  Sys_Level1=trim(rsRule1("Level1"))
	Else
	  Sys_Level1=trim(rsRule1("Level2"))
	End if
end if
rsRule1.close
set rsRule1=nothing

If Not ifnull(Sys_Rule2) Then
	strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and VERSION=(select value from apconfigure where ID=3)"
	if not rsRule1.eof then
		If DateDiff("d",CDate(date),trim(DealLineDate))>-1 Then 
		  Sys_Level2=trim(rsRule1("Level1"))
		Else
		  Sys_Level2=trim(rsRule1("Level2"))
		End if
	end if
	rsRule1.close
	set rsRule1=nothing
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName"))
rszip.close

Sum_Level=cdbl(Sys_Level1)+cdbl(Sys_Level2)

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
if Not unit.eof then Sys_UnitAddress=trim(unit("Address"))
if Not unit.eof then Sys_UnitTel=trim(unit("Tel"))
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

if isnull(Sys_DriverHomeZip) or trim(Sys_DriverHomeZip)="" then Sys_DriverHomeZip="001"
if isnull(Sys_OwnerZip) or trim(Sys_OwnerZip)="" then Sys_OwnerZip="001"

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")

DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,57,160

%>

<table border="0" height=295>
	<tr>
		<td class="style27" valign="top" align="right" height=295 width=5 >
			<div id="Layer5" class="style27" style="position:absolute; left:20px; z-index:5"><%response.write wordporss(chstr(Sys_BillUnitName))%></div>
		</td>
		<td class="style31" valign="bottom" align="right" height=295 width=25><%
			response.write wordporss(chstr(Sys_Driver))%>
		</td>
		<td class="style31" valign="bottom" align="right" height=295 width=5><%

		strtmp=Sys_DriverHomeZip&Sys_DriverHomeAddress

		if len(strtmp)>15 then
			response.write wordporss(chstr(mid(strtmp,1,15)))
		else
			response.write wordporss(chstr(strtmp))
		end if%>
		</td>
		<td class="style31" valign="bottom" align="right" height=295 width=10><%
			if len(strtmp)>15 then response.write wordporss(chstr(mid(strtmp,16,len(strtmp))))%>
		</td>

		<td class="style27" valign="bottom" align="right" height=295 width=20><%
			response.write wordporss(chstr("　　　　　　　"&Sys_BillNo))%>
		</td>　
		<td class="style27" valign="bottom" align="right" height=295 width=15><%
			tmpstr="　　　　　　　　　　　"&left(trim(Sys_Rule1),2)&"　"
			if len(trim(Sys_Rule1))>7 then tmpstr=tmpstr&"　"&right(trim(Sys_Rule1),1)
		tmpstr=tmpstr&Mid(trim(Sys_Rule1),3,1)&Mid(trim(Sys_Rule1),4,2)
				'&Mid(trim(Sys_Rule1),6,2)&"規定。" tmpstr=tmpstr&",期限內自動繳納處新台幣"&Sys_Level1&"元"
				response.write wordporss(chstr(tmpstr))
			%>
		</td>
		<td class="style29" valign="bottom" align="right" height=295 width=10><%
			if trim(Sys_Rule2)<>"0" then
				tmpstr="　　　　　　　　　"&left(trim(Sys_Rule2),2)&"　"
				if len(trim(Sys_Rule2))>7 then tmpstr=tmpstr&"　"&right(trim(Sys_Rule2),1)
				tmpstr=tmpstr&Mid(trim(Sys_Rule2),3,1)&"　"&Mid(trim(Sys_Rule2),4,2)
				response.write wordporss(chstr(tmpstr))
			end if
			%>
		</td>
	</tr>
		<td>
			<div id="idDiv" class="style27" style="position:absolute; left:5px; z-index:5">
				<img src="../BarCodeImage/<%=Sys_BillNo%>.jpg" width="100" height="30">
			</div>
		</td>
		<td colspan=6>
			<div id="Layer6" class="style29" style="position:absolute; left:40px; z-index:5"><%
'				If chkStore=0 Then
'					response.write wordporss(chstr(Sys_MailNumber&"　　"))
'				else
'					response.write wordporss(chstr(Sys_StoreAndSendMailNumber&"　　　　　"))
'				End if%>
			</div>
		</td>
	<tr>
	</tr>
</table>

