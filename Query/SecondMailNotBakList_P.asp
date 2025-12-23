<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_第二次郵寄未退回清冊.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

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
'單退_寄存上傳日
If trim(request("ReturnDateFlag"))="1" Then
	tmpDcilogSql=" and e.exchangeDate Between To_Date('" & gOutDT(trim(request("ReturnDate1")))&" 0:0:0" & "','YYYY/MM/DD/HH24/MI/SS') And To_Date('" & gOutDT(trim(request("ReturnDate2")))&" 23:59:59" & "','YYYY/MM/DD/HH24/MI/SS') and d.UserMarkResonID in ('5','6','7','T')"

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
		response.write "1040px"
	'else	
	'	response.write "100%"
	'end if
	%>" cellpadding="0" cellspacing="0" style="border-collapse: collapse" align="center" >
		<tr>				 
			<td colspan="11"><span class="style1"><b><center><%=thenPasserCity%><%=thenPasserUnit%></center></b></span></td>
		</tr>
		<tr>
		   <td colspan="11"><span class="style1"><u><b><center>二次郵寄未退回清冊</center></b></u></span></td>
		</tr>
		<tr>
		   <td colspan="11"><center><%="("&strDateName(chkDate)&")"%>統計期間: <%=startDate_q%> 至 <%=endDate_q%></center>
<%
	if trim(request("ReturnDateFlag"))="1" then
%>
			<center>(單退上傳日)統計期間: <%=trim(request("ReturnDate1"))%> 至 <%=trim(request("ReturnDate2"))%></center>
<%
	end if
%>
		   </td>
		</tr>		
	</table>
	<br>
	<%="單位名稱:" & P_UnitName & "<br>"%>
	<table border="1" width="<%
	'if sys_City="花蓮縣" then
		response.write "1040px"
	'else	
	'	response.write "100%"
	'end if
	%>" cellpadding="0" cellspacing="0">	
		<tr>
			<td><B><center>序號</center></B></td>
			<td width="43"><B><center>舉發類別</center></B></td>	
			<td width="85"><B><center>單號</center></B></td>
			<td width="85"><B><center>車號</center></B></td>
			<td width="120"><B><center>違規人姓名</center></B></td>
			<td width="210"><B><center>郵寄地址</center></B></td>
			<td width="45"><B><center>站所</center></B></td>
			<td width="80"><B><center><%=strDateName(chkDate)%></center></B></td>
			<td width="180"><B><center>掛號碼</center></B></td>
			<td width="70"><B><center>郵寄日</center></B></td>
			<td width="80"><B><center>郵局</center></B></td>
		</tr><%
		filecmt=0
		BilLBase="select Sn,BillNo,CarNo,BillTypeID,BillUnitID,BillFillDate,RecordDate,RecordMemberID,IllegalDate from BillBase where BillNo is not null "&tmpSql&UnitSql&" and recordstateid=0 and billstatus=3 and NVL(EquiPmentID,1)<>-1"										
		If sys_City="南投縣" Then
			Sqldcireturnstatusid=" and dcireturnstatusid='S'"
		elseIf sys_City="苗栗縣" Then
			Sqldcireturnstatusid=" and dcireturnstatusid in ('Y','S','n','L')"
		Else
			Sqldcireturnstatusid=" and dcireturnstatusid<>'n'"
		End If 
	
		If sys_City="苗栗縣" Then
			strSQL_Order=" order by c.OwnerZip"
		Else
			strSQL_Order=" order by a.billno"
		End If 

		strSQL="select distinct a.BillNo,a.CarNo,a.BillTypeID,a.BillUnitID,a.BillFillDate,a.RecordDate,a.RecordMemberID,a.IllegalDate,b.UnitName,c.DciReturnStation,c.Owner,c.OwnerAddress,c.OwnerZip,c.Driver,c.DriverHomeAddress,c.DriverHomeZip,d.mailDate,d.mailNumber,d.StoreAndSendMailNumber,d.StoreAndSendMailDate,d.StoreAndSendSendDate from ("&BilLBase&") a,UnitInfo b,BillBaseDCIReturn c,BillMailHistory d ,dcilog e where a.billno=e.billno and e.exchangetypeid='N' "&Sqldcireturnstatusid&" and ReturnMarkType='3' and c.exchangetypeid='W' and a.BillUnitID=b.UnitID and a.BillNo=c.BillNo(+) and a.CarNo=c.CarNo(+) and a.SN=d.BillSN(+) " & tmpDcilogSql & strSQL_Order
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

			s_date=gInitDT(trim(rsfound("StoreAndSendSendDate")))
			s_hour=right("0"&hour(rsfound("StoreAndSendSendDate")),2)
			s_minute=right("0"&minute(rsfound("StoreAndSendSendDate")),2)
			mailDate=s_date
			'&"<br>"&s_hour&s_minute

			response.write "<tr>"
			response.write "<td >"&filecmt&"&nbsp;</td>"
			If trim(rsfound("BillTypeID"))="1" Then
				response.write "<td >攔停</td>"
			elseIf trim(rsfound("BillTypeID"))="2" Then
				response.write "<td >逕舉</td>"
			End If 
			response.write "<td >"&Mid(trim(rsfound("BillNo")),1,5)&"****</td>"
			CarNo1=""
			CarNo2=""
			If trim(rsfound("CarNo"))<>"" Then
				If InStr(trim(rsfound("CarNo")),"-")>0 Then
					ArrCarNo=Split(trim(rsfound("CarNo")),"-")
					
						If Len(ArrCarNo(0))>2 then
							CarNo1=Mid(ArrCarNo(0),1,Len(ArrCarNo(0))-2) & "**"
						Else
							CarNo1=Mid(ArrCarNo(0),1,1) & "*"
						End If 
						If Len(ArrCarNo(1))>2 then
							CarNo2="**" & right(ArrCarNo(1),Len(ArrCarNo(1))-2)
						Else
							CarNo2="*" & right(ArrCarNo(1),1) 
						End If 
	
					response.write "<td >"&CarNo1&"-"&CarNo2&"</td>"
				Else 
					response.write "<td >"&Mid(trim(rsfound("CarNo")),1,2)&"****</td>"
				End If 
			Else
				response.write "<td >&nbsp;</td>"
			End If 
			
			ZipTemp=""
	
			If trim(rsfound("BillTypeID"))="1" Then
				if trim(rsfound("DriverHomeZip"))<>"" and not isnull(rsfound("DriverHomeZip")) then
					strZip="select ZipName from Zip where ZipID='"&trim(rsfound("DriverHomeZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing					
					response.write "<td >"&funcCheckFont(trim(rsfound("Driver")),15,1)&"&nbsp;</td>"
					response.write "<td >"&trim(rsfound("DriverHomeZip"))&" "& ZipName2 & funcCheckFont(trim(rsfound("DriverHomeAddress")),15,1)&"&nbsp;</td>"
				else
					strZip="select ZipName from Zip where ZipID='"&trim(rsfound("OwnerZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing										
					response.write "<td >"&funcCheckFont(trim(rsfound("Owner")),15,1)&"&nbsp;</td>"
					response.write "<td >"&trim(rsfound("OwnerZip"))&" "& ZipName2 &funcCheckFont(trim(rsfound("OwnerAddress")),15,1)&"&nbsp;</td>"
				end if
			else
				if trim(rsfound("DriverHomeAddress"))<>"" and not isnull(rsfound("DriverHomeAddress"))  then
					strZip="select ZipName from Zip where ZipID='"&trim(rsfound("DriverHomeZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing
					if isnull(rsfound("Owner")) or trim(rsfound("Owner"))="" then
						GetMailMem="&nbsp;"
					else
						GetMailMem=trim(replace(rsfound("Owner")," "," &nbsp;"))
					end if
					GetMailAddress=trim(rsfound("DriverHomeZip"))&ZipName&trim(rsfound("DriverHomeAddress"))
				else
					strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rsfound("BillNo"))&"' and CarNo='"&trim(rsfound("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rsfound("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
					set rsD=conn.execute(strSqlD)
					if not rsD.eof then
						if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
							if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
								GetMailMem="&nbsp;"
							else
								GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
							end if
							GetMailAddress=trim(rsD("DriverHomeZip"))&trim(rsD("DriverHomeAddress"))
						else
							if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
								GetMailMem="&nbsp;"
							else
								GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
							end if
							GetMailAddress="(車)"&trim(rsD("OwnerZip"))&trim(rsD("OwnerAddress"))
						end if
					else
						strZip="select ZipName from Zip where ZipID='"&trim(rsfound("OwnerZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing
		
						if isnull(rsfound("Owner")) or trim(rsfound("Owner"))="" then
							GetMailMem="&nbsp;"
						else
							GetMailMem=trim(replace(rsfound("Owner")," "," &nbsp;"))
						end if
						GetMailAddress="(車)"&trim(rsfound("OwnerZip"))&ZipName&trim(rsfound("OwnerAddress"))
					end if
					rsD.close
					set rsD=nothing
				end if		
									
				response.write "<td >"&funcCheckFont(GetMailMem,15,1)&"&nbsp;</td>"
				response.write "<td >"&funcCheckFont(GetMailAddress,15,1)&"&nbsp;</td>"
			End If
			response.write "<td >"&trim(rsfound("DciReturnStation"))&"&nbsp;</td>"
			If Trim(chkDate)="0" Then 
				response.write "<td >"&gInitDT(trim(rsfound("BillFillDate")))&"&nbsp;</td>"
			ElseIf Trim(chkDate)="1" Then
				response.write "<td >"&gInitDT(trim(rsfound("IllegalDate")))&"&nbsp;</td>"
			ElseIf Trim(chkDate)="2" Then
				response.write "<td >"&gInitDT(trim(rsfound("RecordDate")))&"&nbsp;</td>"
			End If 
			
			response.write "<td>"&trim(rsfound("StoreAndSendMailNumber"))&"&nbsp;</td>"
			response.write "<td>"&mailDate&"&nbsp;</td>"
			
			If ZipTemp<>"" Then
				response.write "<td>"
				strZS="select * from mailstation where mailareano like '"&ZipTemp&"%'"
				Set rsZS=conn.execute(strZS)
				If Not rsZS.eof Then
					response.write Left(Trim(rsZS("MainSimpleName")),2)&"郵局"
				End If
				rsZS.close
				Set rsZS=Nothing 
				
				response.write "&nbsp;</td>"
			Else
				response.write "<td>&nbsp;</td>"
			End If 

			response.write "</tr>"
			rsfound.movenext
		Wend%>
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