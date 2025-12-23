<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<script type="text/javascript" src="jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="jquery-barcode-2.0.2.min.js"></script>
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

	filecmt=0
	BilLBase="select Sn,BillNo,CarNo,BillTypeID,BillUnitID,RecordDate,RecordMemberID,IllegalDate,Owner,OwnerAddress,OwnerZip from BillBase where BillNo is not null "&tmpSql&UnitSql&" and recordstateid=0 and billstatus=2 and NVL(EquiPmentID,1)<>-1"	
	if sys_City="台東縣" or sys_City="南投縣" then
		BilLBase=BilLBase&"  and billstatus<>'9'"
	End if		
	'2012/05/4 南投陳淑雲說 監理單位已先入案 n 違規人已先繳結案 L ，不出來，固修改 c.Status in ('Y','S','n','L') 為  c.Status in ('Y','S') by jafe,目前只有南投有改，其他縣市未更新過去
	If sys_City="南投縣" Then
		strSQL_Plus=" and c.Status in ('Y','S') "
	Else
		strSQL_Plus=" and c.Status in ('Y','S','n','L') "
	End If 
	If sys_City="苗栗縣" Or sys_City="台中市" Then
		strSQL_Order=" order by c.OwnerZip,a.billno"
	Else
		strSQL_Order=" order by a.billno"
	End If 
	strSQL="select a.BillNo,a.BillTypeID,a.BillUnitID,a.RecordDate,a.RecordMemberID,a.IllegalDate,b.UnitName,c.Owner,c.OwnerAddress,c.OwnerZip,c.Driver,c.DriverHomeAddress,c.DriverHomeZip,d.mailDate,d.mailNumber,d.MailchkNumber,a.Owner as BOwner,a.OwnerAddress as BOwnerAddress,a.OwnerZip as BOwnerZip,e.sn from ("&BilLBase&") a,UnitInfo b,BillBaseDCIReturn c,BillMailHistory d ,dcilog e where a.billno=e.billno and e.exchangetypeid='W' and e.exchangetypeid=c.exchangetypeid "&strSQL_Plus&" and e.DCIErrorCarData<>'V' and a.BillUnitID=b.UnitID and a.BillNo=c.BillNo(+) and a.CarNo=c.CarNo(+) and a.SN=d.BillSN(+) and not exists(select 'Y' from dcicloseclosedata where billno=a.billno) " & strSQL_Order
	'response.write strSQL

%>
<script type="text/javascript">
	$(function () {    
<%
	set rsBarcode=conn.execute(strSQL)
	While Not rsBarcode.eof
%>
	 $("#bcTarget<%=Trim(rsBarcode("Sn"))%>").barcode("<%=Trim(rsBarcode("BillNo"))%>", "code128", { barWidth: 1, barHeight: 20, fontSize: 12, showHRI: false, addQuietZone: false, bgColor: "" });
<%
	rsBarcode.movenext
	Wend
	rsBarcode.close
	Set rsBarcode=nothing
%>
	 });
</script>
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
	<table border="0" width="100%" cellpadding="0" cellspacing="0" style="border-collapse: collapse" align="center" >
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
	<table border="1" width="100%" cellpadding="0" cellspacing="0">	
		<tr>
			<td><B><center>序號</center></B></td>
			<td width="240"><B><center>單號</center></B></td>
			
			
			<td width="120"><B><center>違規人姓名</center></B></td>
			<td width="210"><B><center>郵寄地址</center></B></td>
			<td width="70"><B><center>郵寄日</center></B></td>
			<td width="150"><B><center>掛號碼</center></B></td>
		</tr><%
		
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

			response.write "<tr height='35'>"
			response.write "<td >"&filecmt&"&nbsp;</td>"
			response.write "<td >&nbsp; "
%>
<div id="bcTarget<%=trim(rsfound("SN")) %>" style= "position:absolute;   width:auto;   height:auto;   z-index:1 "> </div> 
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
<%
			response.write trim(rsfound("BillNo"))
			response.write "&nbsp;</td>"
			
			

			if sys_City="台東縣" Then
				if trim(rsfound("Driver"))<>"" and not isnull(rsfound("Driver")) then
					response.write "<td >"&funcCheckFont(trim(rsfound("Driver")),15,1)&"&nbsp;</td>"
				else
					response.write "<td >"&funcCheckFont(rsfound("Owner"),15,1)&"&nbsp;</td>"
				end if
				if trim(rsfound("DriverHomeAddress"))<>"" and not isnull(rsfound("DriverHomeAddress")) then
					strZip="select ZipName from Zip where ZipID='"&trim(rsfound("DriverHomeZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
					rsZip.close
					set rsZip=nothing					
					If IsNull(rsfound("DriverHomeAddress")) Then
						response.write "<td >"&trim(rsfound("DriverHomeZip"))&" "& ZipName2 & funcCheckFont(trim(rsfound("DriverHomeAddress")),15,1)&"&nbsp;</td>"
					else
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
					If IsNull(rsfound("OwnerAddress")) Then
						response.write "<td >"&trim(rsfound("OwnerZip"))&" "& ZipName2 &funcCheckFont(trim(rsfound("OwnerAddress")),15,1)&"&nbsp;</td>"
					else
						response.write "<td >"&trim(rsfound("OwnerZip"))&" "& ZipName2 &funcCheckFont(replace(replace(trim(rsfound("OwnerAddress")),"臺","台"),ZipName2,""),15,1)&"&nbsp;</td>"
					End If 
					
				end if
			elseIf trim(rsfound("BillTypeID"))="1" Then
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
			response.write "<td>"&trim(rsfound("mailchkNumber"))&"&nbsp;</td>"
			
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