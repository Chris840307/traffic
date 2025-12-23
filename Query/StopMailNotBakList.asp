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

Batchnumber_q=trim(request("Batchnumber"))

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
'統計日期
if startDate_q<>"" then
	tmpSql = tmpSql & " and RecordDate Between To_Date('" & gOutDT(startDate_q)&" 0:0:0" & "','YYYY/MM/DD/HH24/MI/SS') And To_Date('" & gOutDT(endDate_q)&" 23:59:59" & "','YYYY/MM/DD/HH24/MI/SS')"
end if


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
		   <td><center>統計期間: <%=startDate_q%> 至 <%=endDate_q%></center></td>
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
			<td width="120"><B><center>違規人姓名</center></B></td>
			<td width="210"><B><center>郵寄地址</center></B></td>
			<td width="70"><B><center>郵寄日</center></B></td>
			<td width="150"><B><center>掛號碼</center></B></td>
		</tr><%
AllCaseCnt=0
MailAddress_tmp=""
BillNo_Tmp=""
GetMailMem_Tmp=""
theMailNumber_Tmp=""
MailDate_Temp=""

BillNo="":CarNo="":mailnumberStr=""

strSQL="select distinct a.ImageFileNameB,a.CarNo,c.MailNumber from (select sn,carno,ImageFileNameB from BillBase where ImagePathName is not null and BillStatus not in ('0','3','7') and RecordStateId=0 and ImageFileNameB is not null and DeallineDate is not null "&tmpSql&") a,(Select * from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b,((select BillSN,CarNo,BillNo,MailNumber from StopBillMailHistory where mailnumber is not null) union all (select BillSN,CarNo,BillNo,StoreAndSendMailNumber MailNumber from StopBillMailHistory where StoreAndSendMailNumber is not null) union all (select BillSN,CarNo,BillNo,ThreeMailNumber MailNumber from StopBillMailHistory where ThreeMailNumber is not null)) c where a.SN=b.BillSN  and a.sn=c.BillSN order by a.ImageFileNameB,c.MailNumber"
'response.write strSQL
set rsbill=conn.execute(strSQL)
while Not rsbill.eof
	If trim(mailnumberStr)<>"" Then
		BillNo=BillNo&","
		CarNo=CarNo&","
		mailnumberStr=mailnumberStr&","
	end if
	BillNo=BillNo&trim(rsbill("ImageFileNameB"))
	CarNo=CarNo&trim(rsbill("CarNo"))
	mailnumberStr=mailnumberStr&trim(rsbill("MailNumber"))
	rsbill.movenext
wend
rsbill.close

PBillNo=split(trim(BillNo),",")
PCarNo=split(trim(CarNo),",")
PmailNumber=split(trim(mailnumberStr),",")

addresscnt=0:tmpBillno="":tmpMailnumber="":TypeMailNumber=""
for cmtI=0 to Ubound(PmailNumber)
	Sys_CarNo="":Sys_Owner="":Sys_OwnerAddress="":Sys_OwnerZip=""

	tmp_OwnerAddress="":tmp_OwnerZip="":arr_OwnerAddress="":arr_OwnerZip=""
	
	If Trim(PBillNo(cmtI))<>Trim(tmpBillno) And Trim(PmailNumber(cmtI))<>Trim(tmpMailnumber) Then
		addresscnt=0
		TypeMailNumber="MailNumber"
	
	ElseIf Trim(PBillNo(cmtI))=Trim(tmpBillno) And Trim(PmailNumber(cmtI))<>Trim(tmpMailnumber) Then
		addresscnt=addresscnt+1

		If addresscnt=1 Then
			TypeMailNumber="StoreAndSendMailNumber"

		ElseIf addresscnt=2 Then
			TypeMailNumber="ThreeMailNumber"
		
		End if
	End If
	
	tmpBillno=Trim(PBillNo(cmtI))
	tmpMailnumber=Trim(PmailNumber(cmtI))

	strSQL="select b.CarNo,Decode(b.Owner,null,a.Owner,b.Owner) Owner,Decode(b.OwnerAddress,null,a.OwnerAddress,b.OwnerAddress) OwnerAddress,Decode(b.DriverAddress,null,a.DriverHomeAddress,b.DriverAddress) DriverHomeAddress,Decode(b.OwnerZip,null,a.OwnerZip,b.OwnerZip) OwnerZip,Decode(b.DriverZip,null,a.DriverHomeZip,b.DriverZip) DriverHomeZip,OwnerNotifyAddress from (select CarNo,OwnerNotifyAddress,OwnerZip,OwnerAddress,DriverHomeZip,DriverHomeAddress,Owner from BillbaseDCIReturn where CarNo='"&trim(PCarNo(cmtI))&"' and ExchangetypeID='A') a,(select distinct CarNo,Owner,OwnerAddress,OwnerZip,DriverAddress,DriverZip from BillBase where ImageFileNameB='"&PBillNo(cmtI)&"') b where a.carno=b.carno"
'response.write strSQL
	set rsDci=conn.execute(strSQL)
	
	if Not rsDci.eof then
		Sys_CarNo=trim(rsDci("CarNo"))
		Sys_Owner=trim(rsDci("Owner"))

		If addresscnt=0 Then
			strSQL="update billbase set Owner='"& trim(rsDci("Owner")) &"' where ImageFileNameB='"&trim(PBillNo(cmtI))&"' and Owner is null"

			conn.execute(strSQL)

		End if

		If not ifnull(rsDci("OwnerNotifyAddress")) Then
			tmp_OwnerAddress=mid(trim(rsDci("OwnerNotifyAddress")),4)
			tmp_OwnerZip=mid(trim(rsDci("OwnerNotifyAddress")),1,3)

		end if

		If not ifnull(rsDci("OwnerAddress")) Then
			If Not ifnull(tmp_OwnerAddress) Then tmp_OwnerAddress=tmp_OwnerAddress&","
			If Not ifnull(tmp_OwnerZip) Then tmp_OwnerZip=tmp_OwnerZip&","

			tmp_OwnerAddress=tmp_OwnerAddress&trim(rsDci("OwnerAddress"))
			tmp_OwnerZip=tmp_OwnerZip&trim(rsDci("OwnerZip"))

			If addresscnt=0 Then
				strSQL="update billbase set OwnerAddress='"&trim(rsDci("OwnerAddress"))&"',OwnerZip='"&trim(rsDci("OwnerZip"))&"' where ImageFileNameB='"&trim(PBillNo(cmtI))&"' and OwnerAddress is null"

				conn.execute(strSQL)

			End if
		End if

		If not ifnull(rsDci("DriverHomeAddress")) Then
			If Not ifnull(tmp_OwnerAddress) Then tmp_OwnerAddress=tmp_OwnerAddress&","
			If Not ifnull(tmp_OwnerZip) Then tmp_OwnerZip=tmp_OwnerZip&","

			tmp_OwnerAddress=tmp_OwnerAddress&trim(rsDci("DriverHomeAddress"))
			tmp_OwnerZip=tmp_OwnerZip&trim(rsDci("DriverHomeZip"))

			If addresscnt=0 Then
				strSQL="update billbase set DriverAddress='"&trim(rsDci("DriverHomeAddress"))&"',DriverZip='"&trim(rsDci("DriverHomeZip"))&"' where ImageFileNameB='"&trim(PBillNo(cmtI))&"' and DriverAddress is null"

				conn.execute(strSQL)

			End if

		End If
		arr_OwnerAddress=Split(tmp_OwnerAddress&" ",",")
		arr_OwnerZip=Split(tmp_OwnerZip&" ",",")

		Sys_OwnerAddress=trim(arr_OwnerAddress(addresscnt))
		Sys_OwnerZip=trim(arr_OwnerZip(addresscnt))

		If not ifnull(Sys_OwnerZip) Then
			strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
			set rszip=conn.execute(strSQL)
			if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
			rszip.close
		else
			Sys_OwnerZipName=""
		End if

		Sys_Address=Sys_OwnerZip&Sys_OwnerAddress
				
	end if
	rsDci.close

	Sys_MailNumber="":Sys_MailChkNumber="":Sys_MailDate=""

	strSQL="select distinct "&TypeMailNumber&" MailNumber from StopBillMailHistory where BillNo='"&PBillNo(cmtI)&"'"
	set rsmail=conn.execute(strSQL)
	If Not rsmail.eof Then
		Sys_MailNumber=trim(rsmail("MailNumber"))&"970007 17"
		Sys_MailChkNumber=trim(rsmail("MailNumber"))&"970007 17"
	end if
	rsmail.close

	strSQL="select distinct MailDate from StopBillMailHistory where BillNo='"&PBillNo(cmtI)&"'"
	set rsmail=conn.execute(strSQL)
	If Not rsmail.eof Then
		Sys_MailDate=trim(rsmail("MailDate"))

	end if
	rsmail.close

	strSQL="select distinct CarNo,BillUnitID,DeallIneDate,ImageFileNameB from BillBase where ImageFileNameB='"&PBillNo(cmtI)&"'"
	set rsbill=conn.execute(strSQL)
	If Not rsbill.eof Then
		Sys_CarNo=trim(rsbill("CarNo"))
		Sys_BillUnitID=trim(rsbill("BillUnitID"))
		Sys_DeallIneDate=split(gArrDT(trim(rsbill("DeallIneDate"))),"-")
		Sys_ImageFileNameB=trim(rsbill("ImageFileNameB"))
	End if
	rsbill.close

	if BillNo_Tmp="" then
		BillNo_Tmp=PBillNo(cmtI)
	else
		BillNo_Tmp=BillNo_Tmp&"@!#"&PBillNo(cmtI)
	end If
	if GetMailMem_Tmp="" then
		GetMailMem_Tmp=Sys_Owner
	else
		GetMailMem_Tmp=GetMailMem_Tmp&"@!#"&Sys_Owner
	end if
	if MailAddress_tmp="" then
		MailAddress_tmp=Sys_Address
	else
		MailAddress_tmp=MailAddress_tmp&"@!#"&Sys_Address
	end If
	if theMailNumber_Tmp="" then
		theMailNumber_Tmp=Sys_MailNumber
	else
		theMailNumber_Tmp=theMailNumber_Tmp&"@!#"&Sys_MailNumber
	end If
	if MailDate_Temp="" then
		MailDate_Temp=Sys_MailDate
	else
		MailDate_Temp=MailDate_Temp&"@!#"&Sys_MailDate
	end If
	
	AllCaseCnt=AllCaseCnt+1
Next

MailAddress_Array=split(MailAddress_tmp&" ","@!#")
theMailNumber_Array=split(theMailNumber_Tmp&" ","@!#")
GetMailMem_Tmp_Array=split(GetMailMem_Tmp&" ","@!#")
BillNo_Tmp_Array=split(BillNo_Tmp&" ","@!#")
MailDate_Temp_Array=split(MailDate_Temp&" ","@!#")
CaseSN=0
mailSNTmp=0

for MAcnt=0 to ubound(MailAddress_Array)
	mailSN=mailSN+1
	CaseSN=CaseSN+1
	MailCnt=MailCnt+1

	response.write "<tr height=""23"">"
	response.write "<td align=""center"">"&mailSN&"</td>"
	
	response.write "<td align=""center"">"&BillNo_Tmp_Array(MAcnt)&"</td>"
	
	response.write "<td align=""center"">"&funcCheckFont(GetMailMem_Tmp_Array(MAcnt),14,1)&"</td>"

	response.write "<td align=""center"">"&funcCheckFont(MailAddress_Array(MAcnt),14,1)&"</td>"
	response.write "<td align=""center"">"
	If Trim(MailDate_Temp_Array(MAcnt))<>"" then
		response.write Year(Trim(MailDate_Temp_Array(MAcnt)))-1911 &"/"&month(Trim(MailDate_Temp_Array(MAcnt))) &"/"&day(Trim(MailDate_Temp_Array(MAcnt)))
	End If 
	response.write"</td>"
	response.write "<td align=""center"">"&theMailNumber_Array(MAcnt)&"</td>"
	response.write "</tr>"
Next

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