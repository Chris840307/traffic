<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_郵寄未退回清冊.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

function funcCheckFontML(strFont,strSize,strFILTER)
	if instr(strFont,"@@")>0 then
		arrFont=split(" "&strFont&" ","@@")
		strTmp=""
		for FontRoop=1 to ubound(arrFont)+1
			if FontRoop mod 2 =0 then
				if strFILTER="1" then	'1正常
					strTmp=strTmp&"＊"
				elseif strFILTER="2" then	'2轉90度
					strTmp=strTmp&"＊"
				elseif strFILTER="3" then	'3轉270度
					strTmp=strTmp&"＊"
				elseif strFILTER="0" then	'變空白
					strTmp=strTmp&"＊"
				elseif strFILTER="4" then	'變空白 轉V
					strTmp=strTmp&"＊"
				end if
			else
				strTmp=strTmp&trim(arrFont(FontRoop-1))
			end if
		next

		funcCheckFontML=strTmp
	else
		funcCheckFontML=replace(strFont&""," ","　")
	end if
end Function

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
		response.write "1040px"
	'else	
	'	response.write "100%"
	'end if
	%>" cellpadding="0" cellspacing="0" style="border-collapse: collapse" align="center" >
		<tr>				 
			<td colspan="11"><span class="style1"><b><center><%=thenPasserCity%><%=thenPasserUnit%></center></b></span></td>
		</tr>
		<tr>
		   <td colspan="11"><span class="style1"><u><b><center>郵寄未退回清冊</center></b></u></span></td>
		</tr>
		<tr>
		   <td colspan="11"><center><%="("&strDateName(chkDate)&")"%>統計期間: <%=startDate_q%> 至 <%=endDate_q%></center></td>
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
		BilLBase="select Sn,BillNo,CarNo,BillTypeID,BillUnitID,BillFillDate,RecordDate,RecordMemberID,IllegalDate,Owner,OwnerAddress,OwnerZip from BillBase where BillNo is not null "&tmpSql&UnitSql&" and recordstateid=0 and billstatus=2 and NVL(EquiPmentID,1)<>-1"	
		if sys_City="台東縣" or sys_City="南投縣" then
			BilLBase=BilLBase&"  and billstatus<>'9'"
		End if		
		'2012/05/4 南投陳淑雲說 監理單位已先入案 n 違規人已先繳結案 L ，不出來，固修改 c.Status in ('Y','S','n','L') 為  c.Status in ('Y','S') by jafe,目前只有南投有改，其他縣市未更新過去
		If sys_City="苗栗縣" Then
			strSQL_Plus=" and c.Status in ('Y','S','n','L') "
		elseIf sys_City="南投縣" Then	'2015/3/13 南投李疑針說要加上 L 已經入案過
			strSQL_Plus=" and c.Status in ('Y','S','L') "
		Else
			strSQL_Plus=" and c.Status in ('Y','S') "
		End If 
		If sys_City="苗栗縣" Or sys_City="基隆市" Or sys_City="台中市" Then
			strSQL_Order=" order by BOwnerZip,a.billno"
		Else
			strSQL_Order=" order by a.billno"
		End If 
		strSQL="select a.BillNo,a.CarNo,a.BillTypeID,a.BillUnitID,BillFillDate,a.RecordDate,a.RecordMemberID,a.IllegalDate,b.UnitName,c.DciReturnStation,c.Owner,c.OwnerAddress,c.OwnerZip,c.Driver,c.DriverHomeAddress,c.DriverHomeZip,d.mailDate,d.mailNumber,d.MailchkNumber,a.Owner as BOwner,a.OwnerAddress as BOwnerAddress,a.OwnerZip as BOwnerZip from ("&BilLBase&") a,UnitInfo b,BillBaseDCIReturn c,BillMailHistory d ,dcilog e where a.billno=e.billno and e.exchangetypeid='W' and e.exchangetypeid=c.exchangetypeid "&strSQL_Plus&" and e.DCIErrorCarData<>'V' and a.BillUnitID=b.UnitID and a.BillNo=c.BillNo(+) and a.CarNo=c.CarNo(+) and a.SN=d.BillSN(+)"& strSQL_Order
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
				if trim(rsfound("DriverHomeAddress"))<>"" and not isnull(rsfound("DriverHomeAddress")) Then
					ZipTemp=trim(rsfound("DriverHomeZip"))
					strZip="select ZipName from Zip where ZipID='"&trim(rsfound("DriverHomeZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing					
					response.write "<td >"&funcCheckFontML(trim(rsfound("Driver")),15,1)&"&nbsp;</td>"
					If IsNull(rsfound("DriverHomeAddress")) Then
						response.write "<td >"&trim(rsfound("DriverHomeZip"))&" "& ZipName2 & funcCheckFontML(trim(rsfound("DriverHomeAddress")),15,1)&"&nbsp;</td>"
					Else 
						response.write "<td >"&trim(rsfound("DriverHomeZip"))&" "& ZipName2 & funcCheckFontML(replace(replace(trim(rsfound("DriverHomeAddress")),"臺","台"),ZipName2,""),15,1)&"&nbsp;</td>"
					End If 
					
				Else
					ZipTemp=trim(rsfound("OwnerZip"))
					strZip="select ZipName from Zip where ZipID='"&trim(rsfound("OwnerZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing										
					response.write "<td >"&funcCheckFontML(rsfound("Owner"),15,1)&"&nbsp;</td>"
					If IsNull(rsfound("OwnerAddress")) Then
						response.write "<td >"&trim(rsfound("OwnerZip"))&" "& ZipName2 &funcCheckFontML(trim(rsfound("OwnerAddress")),15,1)&"&nbsp;</td>"
					else
						response.write "<td >"&trim(rsfound("OwnerZip"))&" "& ZipName2 &funcCheckFontML(replace(replace(trim(rsfound("OwnerAddress")),"臺","台"),ZipName2,""),15,1)&"&nbsp;</td>"
					End If 
				end if
			Else
				If Trim(rsfound("BOwnerAddress"))<>"" Then
					ZipTemp=trim(rsfound("BOwnerZip"))
					strZip="select ZipName from Zip where ZipID='"&trim(rsfound("BOwnerZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing			
									
					response.write "<td >"&funcCheckFontML(trim(rsfound("BOwner")),15,1)&"&nbsp;</td>"
					If IsNull(rsfound("BOwnerAddress")) Then
						response.write "<td >"&trim(rsfound("BOwnerZip"))&" "& ZipName2 &funcCheckFontML(trim(rsfound("BOwnerAddress")),15,1)&"&nbsp;</td>"
					Else 
						response.write "<td >"&trim(rsfound("BOwnerZip"))&" "& ZipName2 &funcCheckFontML(replace(replace(trim(rsfound("BOwnerAddress")),"臺","台"),ZipName2,""),15,1)&"&nbsp;</td>"
					End If 
				Else 
					ZipTemp=trim(rsfound("OwnerZip"))
					strZip="select ZipName from Zip where ZipID='"&trim(rsfound("OwnerZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing			
									
					response.write "<td >"&funcCheckFontML(trim(rsfound("Owner")),15,1)&"&nbsp;</td>"
					If IsNull(rsfound("OwnerAddress")) Then
						response.write "<td >"&trim(rsfound("OwnerZip"))&" "& ZipName2 &funcCheckFontML(trim(rsfound("OwnerAddress")),15,1)&"&nbsp;</td>"
					Else 
						response.write "<td >"&trim(rsfound("OwnerZip"))&" "& ZipName2 &funcCheckFontML(replace(replace(trim(rsfound("OwnerAddress")),"臺","台"),ZipName2,""),15,1)&"&nbsp;</td>"
					End If 
				End If 
			End If
			response.write "<td >"&trim(rsfound("DciReturnStation"))&"&nbsp;</td>"
			If Trim(chkDate)="0" Then 
				response.write "<td >"&gInitDT(trim(rsfound("BillFillDate")))&"&nbsp;</td>"
			ElseIf Trim(chkDate)="1" Then
				response.write "<td >"&gInitDT(trim(rsfound("IllegalDate")))&"&nbsp;</td>"
			ElseIf Trim(chkDate)="2" Then
				response.write "<td >"&gInitDT(trim(rsfound("RecordDate")))&"&nbsp;</td>"
			End If 
			
			if trim(rsfound("MailNumber"))<>"" and not isnull(rsfound("MailNumber")) then
				response.write "<td>"&Right("000000"&trim(rsfound("MailNumber")),6)&" 36 400017</td>"
			else
				response.write "<td>&nbsp;</td>"
			end if
			
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