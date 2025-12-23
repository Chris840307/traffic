<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

chkDate=trim(request("chkDate"))
strDate=split("BillFillDate,IllegalDate,RecordDate",",")
strDateName=split("填單日,違規日,建檔日",",")
UserId = Session("User_ID")
startDate_q = Trim(Request("startDate_q"))
endDate_q = Trim(Request("endDate_q"))
unit = Request("unit")
UnitID_q = Request("UnitID_q")
unitList=trim(request("unitSelectlist"))
Batchnumber_q=trim(request("Batchnumber"))
Memlist_q=trim(request("MemSelectlist"))
Server.ScriptTimeout=86400


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
elseif Sys_UnitLevelID=2 and sys_City<>"連江縣" then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
end if

set rsunit=conn.Execute(strSQL)

if Not rsunit.eof then Sys_UnitID=trim(rsunit("UnitID"))

if Not rsunit.eof then thenPasserUnit=trim(rsunit("UnitName"))
rsunit.close




strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
rsUInfo.close

tmpSql=""
'入案批號
if Batchnumber_q<>"" then
	tmpSql = tmpSql & " and SN in (select BillSn from Dcilog where BatchNumber='" & Batchnumber_q & "')"
end if
'建檔人員
if Memlist_q<>"" then
	tmpSql = tmpSql & " and RecordMemberId in (" & Memlist_q & ")"
end if
'統計日期
if startDate_q<>"" then
	tmpSql = tmpSql & " and "&strDate(chkDate)&" Between To_Date('" & gOutDT(startDate_q)&" 0:0:0" & "','YYYY/MM/DD/HH24/MI/SS') And To_Date('" & gOutDT(endDate_q)&" 23:59:59" & "','YYYY/MM/DD/HH24/MI/SS')"
end if
'舉發單號
if trim(request("startBillNo_q"))<>"" then
	tmpSql = tmpSql & " and BillNo Between '" & trim(request("startBillNo_q")) & "' And '" & trim(request("endBillNo_q")) & "'"
end if
'舉發單位
If unit="y" Then
	unitList = Split(unitList,",")
	Sys_UnitID=""
	for i=0 to UBound(unitList)
		if Sys_UnitID<>"" then Sys_UnitID=Sys_UnitID&"','"
		Sys_UnitID=Sys_UnitID&unitList(i)
	next
	UnitSql = " and BillUnitID in ('" & Sys_UnitID & "')"
End If

P_UnitName=thenPasserCity
strSQL="select UnitName from UnitInfo where UnitID='"&UnitID_q&"'"

set rsunit=conn.execute(strSQL)
If Not rsunit.eof Then P_UnitName=trim(rsunit("UnitName"))
rsunit.close
%>
<html>   
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>ExportBase</title>
<style type="text/css">
<!--
body {font-family:標楷體;font-size:12pt}
.style1 {font-family:標楷體;font-size:14pt}
-->
</style>
</head>	 
<body>  
	<!--  
	<table border="0" width="15%" cellpadding="0" cellspacing="0" style="border-collapse: collapse" >
		<tr>
			<td>
				列印時間: <%=gInitDT(Date)%> <br>
			    列印單位: <%=thenPasserUnit%> <br>
			    列印人員: <%=Session("Ch_Name")%>
			</td>
		</tr>	  
	</table>
	-->
	<table border="0" width="<%
	'if sys_City="花蓮縣" then
		response.write "680px"
	'else	
	'	response.write "100%"
	'end if
	%>" cellpadding="0" cellspacing="0" style="border-collapse: collapse" align="center" >
		<tr>				 
			<td><span class="style1"><b><center><%=thenPasserCity%><%=thenPasserUnit%></center></b></span></td>
		</tr>
		<tr>
		   <td><span class="style1"><u><b><center>郵寄未退回清冊</center></b></u></span></td>
		</tr>
		<tr>
		   <td><center><%="("&strDateName(chkDate)&")"%>統計期間: <%=startDate_q%> 至 <%=endDate_q%></center></td>
		</tr>		
	</table>
	<br>
	<%="單位名稱:" & P_UnitName & "<br>"%>
	<table border="1" width="<%
	if sys_City="台南市" then
		response.write "100%"
	else	
		response.write "680px"
	end if
	%>" cellpadding="0" cellspacing="0">	
		<tr>
			<td><B><center>序號</center></B></td>
			<td width="90"><B><center>單號</center></B></td>
			
			<%if sys_City = "南投縣" then %>
				<td><B><center>舉發單位</center></B></td>					
				<td><B><center>建檔日</center></B></td>
				<td><B><center>建檔人</center></B></td>
				<td><B><center>違規日</center></B></td>
			<% end if %>
			
			<td width="120"><B><center>違規人姓名</center></B></td>
			<td width="210"><B><center>郵寄地址</center></B></td>
			<td width="70"><B><center>郵寄日</center></B></td>
			<td width="150"><B><center>掛號碼</center></B></td>
			<%if sys_City = "台南市" then %>
				<td width="70"><B><center>建檔人</center></B></td>
				<td width="70"><B><center>舉發人</center></B></td>
				<td><B><center>違規事實</center></B></td>
			<% end if %>
		</tr><%
		filecmt=0
		BilLBase="select Sn,BillNo,CarNo,BillTypeID,BillUnitID,RecordDate,RecordMemberID,IllegalDate,Owner,OwnerAddress,OwnerZip,Rule1,Rule2,BillMem1,BillMem2,BillMem3 from BillBase where BillNo is not null "&tmpSql&UnitSql&" and recordstateid=0 and billstatus=2 and NVL(EquiPmentID,1)<>-1"	
		if sys_City="台東縣" or sys_City="南投縣" then
			BilLBase=BilLBase&"  and billstatus<>'9'"
		End if		
		'2012/05/4 南投陳淑雲說 監理單位已先入案 n 違規人已先繳結案 L ，不出來，固修改 c.Status in ('Y','S','n','L') 為  c.Status in ('Y','S') by jafe,目前只有南投有改，其他縣市未更新過去
		If sys_City="苗栗縣" Or sys_City="花蓮縣" Then
			strSQL_Plus=" and c.Status in ('Y','S','n','L') "
		elseIf sys_City="南投縣" Or sys_City="台東縣" Then	'2015/3/13 南投李疑針說要加上 L 已經入案過
			strSQL_Plus=" and c.Status in ('Y','S','L') "
		Else
			strSQL_Plus=" and c.Status in ('Y','S','L') "
		End If 
		If sys_City="苗栗縣" Or sys_City="基隆市" Or sys_City="台中市" Then
			strSQL_Order=" order by c.OwnerZip,a.billno"
		Else
			strSQL_Order=" order by a.billno"
		End If 
		strSQL="select a.BillNo,a.CarNo,a.BillTypeID,a.BillUnitID,a.RecordDate,a.RecordMemberID,a.IllegalDate,b.UnitName,c.Owner,c.OwnerAddress,c.OwnerZip,c.Driver,c.DriverHomeAddress,c.DriverHomeZip,d.mailDate,d.mailNumber,d.MailchkNumber,a.Owner as BOwner,a.OwnerAddress as BOwnerAddress,a.OwnerZip as BOwnerZip,e.billsn,a.BillMem1,a.BillMem2,a.BillMem3,a.Rule1,a.Rule2 from ("&BilLBase&") a,UnitInfo b,BillBaseDCIReturn c,BillMailHistory d ,dcilog e where a.billno=e.billno and e.exchangetypeid='W' and e.exchangetypeid=c.exchangetypeid "&strSQL_Plus&" and e.DCIErrorCarData<>'V' and a.BillUnitID=b.UnitID and a.BillNo=c.BillNo(+) and a.CarNo=c.CarNo(+) and a.SN=d.BillSN(+) " & strSQL_Order
		'response.write strSQL
		set rsfound=conn.execute(strSQL)
		While Not rsfound.eof
			filecmt=filecmt+1
			s_date=gInitDT(trim(rsfound("RecordDate")))
			s_hour=right("0"&hour(rsfound("RecordDate")),2)
			s_minute=right("0"&minute(rsfound("RecordDate")),2)
			RecordDate=s_date&"<br>"&s_hour&s_minute

			s_date=gInitDT(trim(rsfound("IllegalDate")))
			s_hour=right("0"&hour(rsfound("IllegalDate")),2)
			s_minute=right("0"&minute(rsfound("IllegalDate")),2)
			IllegalDate=s_date&"<br>"&s_hour&s_minute

			s_date=gInitDT(trim(rsfound("mailDate")))
			s_hour=right("0"&hour(rsfound("mailDate")),2)
			s_minute=right("0"&minute(rsfound("mailDate")),2)
			mailDate=s_date
			'&"<br>"&s_hour&s_minute

			response.write "<tr>"
			response.write "<td >"&filecmt&"&nbsp;</td>"
			response.write "<td >"&trim(rsfound("BillNo"))&"&nbsp;</td>"
	if sys_City = "南投縣" then
			response.write "<td >"&trim(rsfound("UnitName"))&"&nbsp;</td>"		
			response.write "<td >"&RecordDate&"&nbsp;</td>"
			response.write "<td >"
			strMem="select ChName from Memberdata where MemberID="&trim(rsfound("RecordMemberID"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("ChName"))
			end if
			rsMem.close
			set rsMem=nothing
			
			response.write "&nbsp;</td>"
			response.write "<td >"&IllegalDate&"&nbsp;</td>"
	end if
	
			
			ZipName2=""
			If trim(rsfound("BillTypeID"))="1" Then
				if trim(rsfound("DriverHomeAddress"))<>"" and not isnull(rsfound("DriverHomeAddress")) then
					strZip="select ZipName from Zip where ZipID='"&trim(rsfound("DriverHomeZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing					
					response.write "<td >"&funcCheckFont(trim(rsfound("Driver")),15,1)&"&nbsp;</td>"
					If IsNull(rsfound("DriverHomeAddress")) Then
						response.write "<td >"&trim(rsfound("DriverHomeZip"))&" "& ZipName2 & funcCheckFont(trim(rsfound("DriverHomeAddress")),15,1)&"&nbsp;</td>"
					Else 
						response.write "<td >"&trim(rsfound("DriverHomeZip"))&" "& ZipName2 & funcCheckFont(replace(replace(trim(rsfound("DriverHomeAddress")),"臺","台"),ZipName2,""),15,1)&"&nbsp;</td>"
					End If 
					
				else
					strZip="select ZipName from Zip where ZipID='"&trim(rsfound("OwnerZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing										
					response.write "<td >"&funcCheckFont(rsfound("Owner"),15,1)&"&nbsp;</td>"
					If IsNull(rsfound("OwnerAddress")) Then
						response.write "<td >"&trim(rsfound("OwnerZip"))&" "& ZipName2 &funcCheckFont(trim(rsfound("OwnerAddress")),15,1)&"&nbsp;</td>"
					else
						response.write "<td >"&trim(rsfound("OwnerZip"))&" "& ZipName2 &funcCheckFont(replace(replace(trim(rsfound("OwnerAddress")),"臺","台"),ZipName2,""),15,1)&"&nbsp;</td>"
					End If 
				end if
			else
				if sys_City="南投縣" And Trim(rsfound("BOwnerAddress"))<>"" Then
					strZip="select ZipName from Zip where ZipID='"&trim(rsfound("BOwnerZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing			
									
					response.write "<td >"&funcCheckFont(trim(rsfound("BOwner")),15,1)&"&nbsp;</td>"
					response.write "<td >"&trim(rsfound("BOwnerZip"))&" "& ZipName2 &funcCheckFont(replace(replace(trim(rsfound("BOwnerAddress")),"臺","台"),ZipName2,""),15,1)&"&nbsp;</td>"
				elseif sys_City="台東縣" Then
					Response.flush
					ZipName=""
					strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress,dcierrorcardata,Nwner,NwnerZip,NwnerAddress from BIllBaseDCIReturn where (BillNo='"&trim(rsfound("BillNo"))&"' and CarNo='"&trim(rsfound("CarNo"))&"') and ExchangeTypeID='W' and Status in('Y','S','n','L')"
					set rsD=conn.execute(strSqlD)
					if not rsD.eof Then
						if ExchangeTypeFlag="N" Then
							if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
				
								GetMailMem=trim(rsD("Owner"))
								GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress"))&"","臺","台"),ZipName,"")
							else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
				
								GetMailMem=trim(rsD("Owner"))
								GetMailAddress=trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress"))&"","臺","台"),ZipName,"")
							end If
						else
							if instr(trim(rsD("OwnerAddress")),"(住)")>1 or instr(trim(rsD("OwnerAddress")),"(就)")>1 or instr(trim(rsD("OwnerAddress")),"（住）")>1 or instr(trim(rsD("OwnerAddress")),"（就）")>1 then
								strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
				
								GetMailMem=trim(rsD("Owner"))
								GetMailAddress=trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress"))&"","臺","台"),ZipName,"")
							else
								strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where Exists (select carno from dcilog where BillSN="&trim(rsfound("BillSN"))&" and CarNo='"&trim(rsfound("CarNo"))&"' and ExchangetypeID='A' and dcireturnstatusid='S') and CarNo='"&trim(rsfound("CarNo"))&"' and ExchangeTypeID='A' and Status='S'"
								Set rsD3=conn.execute(strSqlD)
								If Not rsD3.eof Then
									If trim(rsD3("DriverHomeAddress"))<>"" And not isnull(rsD3("DriverHomeAddress")) then
										GetMailMem=trim(rsD("Owner"))

										strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
										set rsZip=conn.execute(strZip)
										if not rsZip.eof then
											ZipName=trim(rsZip("ZipName"))
										end if
										rsZip.close
										set rsZip=Nothing
										
										GetMailAddress=trim(rsD3("DriverHomeZip"))&ZipName&replace(replace(trim(rsD3("DriverHomeAddress"))&"","臺","台"),ZipName,"")
									Else
										strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
										set rsZip=conn.execute(strZip)
										if not rsZip.eof then
											ZipName=trim(rsZip("ZipName"))
										end if
										rsZip.close
										set rsZip=nothing
										GetMailMem=trim(rsD("Owner"))
										GetMailAddress=trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress"))&"","臺","台"),ZipName,"")
									End If
								Else
									strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
									set rsZip=conn.execute(strZip)
									if not rsZip.eof then
										ZipName=trim(rsZip("ZipName"))
									end if
									rsZip.close
									set rsZip=nothing
									GetMailMem=trim(rsD("Owner"))
									GetMailAddress=trim(rsD("OwnerZip"))&ZipName&replace(replace(trim(rsD("OwnerAddress"))&"","臺","台"),ZipName,"")
								End If
								rsD3.close
								Set rsD3=Nothing 
							end If
						End If 
					end if
					rsD.close
					set rsD=Nothing
					response.write "<td >"&funcCheckFont(GetMailMem,15,1)&"&nbsp;</td>"
					response.write "<td >"&funcCheckFont(GetMailAddress,15,1)&"&nbsp;</td>"
				Else
					strZip="select ZipName from Zip where ZipID='"&trim(rsfound("OwnerZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing			
									
					response.write "<td >"&funcCheckFont(trim(rsfound("Owner")),15,1)&"&nbsp;</td>"
					If IsNull(rsfound("OwnerAddress")) Then
						response.write "<td >"&trim(rsfound("OwnerZip"))&" "& ZipName2 &funcCheckFont(trim(rsfound("OwnerAddress")),15,1)&"&nbsp;</td>"
					Else 
						response.write "<td >"&trim(rsfound("OwnerZip"))&" "& ZipName2 &funcCheckFont(replace(replace(trim(rsfound("OwnerAddress")),"臺","台"),ZipName2,""),15,1)&"&nbsp;</td>"
					End If 
				End If 
			End if
			response.write "<td>"&mailDate&"&nbsp;</td>"
			if sys_City="花蓮縣" Then
				If trim(rsfound("mailchkNumber"))<>"" then
					response.write "<td>"&left(trim(rsfound("mailchkNumber")),6)&"&nbsp;</td>"
				Else
					response.write "<td>"&left(trim(rsfound("mailNumber")),6)&"&nbsp;</td>"
				End if 
			elseif sys_City="台中市" or sys_City="雲林縣" then
				response.write "<td>"&right("00000000" & trim(rsfound("mailNumber")),6)&"&nbsp;</td>"
			elseif sys_City="南投縣" Then
				if trim(rsfound("MailNumber"))<>"" and not isnull(rsfound("MailNumber")) then
					response.write "<td>"&left(right("000000000000000000" & trim(rsfound("MailNumber")),14),6)&"&nbsp;</td>"
				else
					response.write "<td>&nbsp;</td>"
				end if
			Else
				if trim(rsfound("MailNumber"))<>"" and not isnull(rsfound("MailNumber")) then
					response.write "<td>"&trim(rsfound("MailNumber"))&"&nbsp;</td>"
				else
					response.write "<td>&nbsp;</td>"
				end if
			end If

			If sys_City = "台南市" Then
				response.write "<td >"
				strMem="select ChName from Memberdata where MemberID="&trim(rsfound("RecordMemberID"))
				set rsMem=conn.execute(strMem)
				if not rsMem.eof then
					response.write trim(rsMem("ChName"))
				end if
				rsMem.close
				set rsMem=nothing
				
				response.write "&nbsp;</td>"
				response.write "<td >"&trim(rsfound("BillMem1"))
				If trim(rsfound("BillMem2"))<>"" Then
					response.write ","&trim(rsfound("BillMem2"))
				End If 
				If trim(rsfound("BillMem3"))<>"" Then
					response.write ","&trim(rsfound("BillMem3"))
				End If 
				response.write "&nbsp;</td>"

				response.write "<td >"&trim(rsfound("Rule1"))
				If trim(rsfound("Rule2"))<>"" Then
					response.write ","&trim(rsfound("Rule2"))
				End If 
				response.write "&nbsp;</td>"
			end if
			response.write "</tr>"
			rsfound.movenext
		Wend
		rsfound.close 
		Set rsfound=nothing
		%>
	</table>
<%
'fMnoth=month(now)
'if fMnoth<10 then fMnoth="0"&fMnoth
'fDay=day(now)
'if fDay<10 then	fDay="0"&fDay
'fname=year(now)&fMnoth&fDay&"_郵寄未退還清冊.xls"
'Response.AddHeader "Content-Disposition", "filename="&fname
'response.contenttype="application/x-msexcel; charset=MS950" 
%>	 
</body>
</html>