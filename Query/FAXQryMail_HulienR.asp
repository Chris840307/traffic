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
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>受理局填寫</title>

<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>
<body>
<%
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
filecmt=0
for MAcnt=0 to ubound(MailAddress_Array)
	filecmt=filecmt+1
	CaseSN=CaseSN+1
	MailCnt=MailCnt+1
			if cint(filecmt)>1 then response.write "<div class=""PageNext"">&nbsp;</div>"
           	DelphiASPObj.GenSendStoreBillno BillNo_Tmp_Array(MAcnt),0,50,160
			s_Year = year(MailDate_Temp_Array(MAcnt))-1911
			s_month = month(MailDate_Temp_Array(MAcnt))
			s_day = day(MailDate_Temp_Array(MAcnt))
%>

<div id="R1" style="position:relative;">
<table border="0" width="100" id="table1" height="625" cellspacing="0" cellpadding="0">
	<tr>
		<td>
		<table border="0" width="100" id="table2" cellspacing="0" cellpadding="0" height="625">
			<tr>
			<td colspan="3" align="right">
			<font face="標楷體" size="5">傳真查詢國內各類掛號郵件查單</font>　　　　　　　　<font face="標楷體">編列第　　　　　　　號&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
			</font>&nbsp;<tr>
				<td width="485" align="left" valign="top">
				<table border="1" width="485" id="table3" cellspacing="0" cellpadding="0" height="625">
					<tr>
						<td rowspan="3" width="16" align="center">
						<font face="標楷體">受理局填寫</font></td>
						<td rowspan="2" width="80" colspan="2" align="center">
						<font face="標楷體">原　寄<br>局　名</font></td>
						<td width="74" rowspan="2"  align="center"><font face="標楷體">
						<%if session("Unit_ID")="A000" then %>
						花蓮市府前路郵局
						<%else%>
						&nbsp;
						<%End if%></font></td>
						<td colspan="20" align="center"><font face="標楷體">條&nbsp;碼&nbsp;掛&nbsp;號&nbsp;收&nbsp;據&nbsp;之</font></td>
					</tr>
					<tr>
						<td colspan="6" align="center"><font face="標楷體">掛號號碼</font></td>
						<td  align="center" rowspan="2"><font size="2" face="標楷體">&nbsp;</font></td>
						<td  align="center" colspan="12"><font face="標楷體">&nbsp;&nbsp;原&nbsp;&nbsp;寄&nbsp;&nbsp;局&nbsp;&nbsp;碼&nbsp;&nbsp;</font></td>
					</tr>
					<tr>
						<td width="60" colspan="2" height="44" align="center">
						<font face="標楷體">掛　號<br>種　類</font></td>
						<td width="74" height="44"  align="center"><font face="標楷體">雙掛號</font></td>
						<td height="44" width="14" align="center"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),1,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),1,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="15" align="center"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),2,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),2,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="15" align="center"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),3,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),3,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="13" align="center"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),4,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),4,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="15" align="center"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),5,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),5,1) else response.write "&nbsp;"%></font></td>
						<td height="44" width="15" align="center"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),6,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),6,1) else response.write "&nbsp;"%></font></td>

						<td width="15" height="44" align="center"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),7,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),7,1) else response.write "&nbsp;"%></font></td>
						<td width="14" height="44" align="center"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),8,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),8,1) else response.write "&nbsp;"%></font></td>
						<td width="13" height="44" align="center"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),9,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),9,1) else response.write "&nbsp;"%></font></td>
						<td width="13" height="44" align="center"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),10,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),10,1) else response.write "&nbsp;"%></font></td>
						<td width="12" height="44" align="center"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),11,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),11,1) else response.write "&nbsp;"%></font></td>
						<td width="13" height="44" align="center" colspan="2"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),12,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),12,1) else response.write "&nbsp;"%></font></td>
						<td width="12" height="44" align="center" colspan="2"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),14,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),14,1) else response.write "&nbsp;"%></font></td>
						<td width="12" height="44" align="center" colspan="2"><font face="標楷體"><%if mid(theMailNumber_Array(MAcnt),15,1)<>"" then response.write mid(theMailNumber_Array(MAcnt),15,1) else response.write "&nbsp;"%></font></td>
					</tr>
					<tr>
						<td width="16" rowspan="6" align="center">
						<font face="標楷體">查</font><p><br><font face="標楷體">詢</font></p>
						<p>&nbsp;<br><font face="標楷體">人</font></p>
						<p>&nbsp;<br><font face="標楷體">填</font></p>
						<p><br><font face="標楷體">寫</font></td>
						<td width="60" colspan="2" height="36" align="center">
						<font face="標楷體">交　寄<br>日　期</font></td>
						<td colspan="21" height="36"><font face="標楷體">　<%=s_Year%>　年　<%=s_month%>　月　<%=s_day%>　日　</font><img src="../BarCodeImage/<%=BillNo_Tmp_Array(MAcnt)%>.jpg"></td>
					</tr>
					<tr>
						<td width="60" colspan="2" height="50" align="center">
						<font face="標楷體">報　值<br>保　價
						<br>金　額</font></td>
						<td width="110" height="50" colspan="4">　　</td>
						<td height="50" width="47" align="center" colspan="3"><font face="標楷體">重量</font></td>
						<td height="50" width="90" align="center" colspan="6">　</td>
						<td height="50" width="36" align="center" colspan="2"><font face="標楷體">內裝</font></td>
						<td height="50" width="96" colspan="6" align="center">　</td>
					</tr>
					<tr>
						<td width="21" rowspan="2" align="center">
						<font face="標楷體">寄件人</font></td>
						<td width="37" rowspan="2" align="center">
						<font face="標楷體">姓名住址電話</font></td>
						<td width="299" colspan="15" rowspan="2"><font face="標楷體">
							<%=thenPasserUnitName%>
							<br>
							<%=thenPasserUnitAddress%>
							<br>
							<%=thenPasserUnitTel%></font>
						</td>
						<td width="96" colspan="6" height="26" align="center">
						<font face="標楷體" size="2">寄件人FAX號碼</font></td>
					</tr>
					<tr>
						<td width="96" colspan="6" height="40">&nbsp;<%=cdbl("0"&BillNo_Tmp_Array(MAcnt))%></td>
					</tr>
					<tr>
						<td width="21" align="center" height="63"><font face="標楷體">收件人</font></td>
						<td width="37" align="center" height="63"><font face="標楷體">姓名地址電話</font></td>
						<td width="401" colspan="21" height="63"><font face="標楷體">
						&nbsp;清冊編號&nbsp;<%=filecmt%>
						<br>
							<%=funcCheckFont(GetMailMem_Tmp_Array(MAcnt),14,1)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"%>
							<%If sys_City="南投縣" Then  response.write CarNo%>
							<Br>
							<%=funcCheckFont(MailAddress_Array(MAcnt),14,1)%></font>
						</td>
					</tr>
					<tr>
						<td width="60" colspan="2" align="center">
						<font face="標楷體">查　詢<br>結　果</font></td>
						<td width="401" colspan="21"><font face="標楷體">　□電話通知　　▉傳真　　□補發回執</font></td>
					</tr>
					<tr>
						<td width="16" align="center" rowspan="2"><br><font face="標楷體">受</font><p>&nbsp;<br><font face="標楷體">理</font></p>
						<p>&nbsp;<br><font face="標楷體">局</font></p>
						<p>&nbsp;<br><font face="標楷體">填</font></p>
						<p>&nbsp;<br><font face="標楷體">寫</font></td>
						<td width="60" colspan="2" align="center">
						<font face="標楷體">投　遞<br>局　別</font></td>
						<td width="401" colspan="21"><font face="標楷體">　　　　　　　　　郵局</font></td>
					</tr>
					<tr>
						<td width="95" colspan="23"  height="82">

								<table border="0" width="457" id="table5" height="204" cellspacing="0" cellpadding="0">

									<tr>
										<td colspan="2" height="132">
										<font face="標楷體"><br>　查右列郵件，據寄件人聲稱，並未寄到，請即迅為查詢見覆。<br>
										　本局傳真號碼「&nbsp;03-8344682&nbsp;」。
										<br>　
										<br>　　　　　　　　　　　　　經辦員：
										<br>　　　　　　　　　　　　　主　管：
										<br>　中華民國　　　　年　　　　月　　　　日</font></td>
										<tr>
										<td width="433" align="right" valign="top">
										<table border="1"  id="table6" cellspacing="0" cellpadding="0" height="72" width="183">
											<tr>
												<td width="179" height="30">
												<font face="標楷體">除快捷郵件外，其他郵件應收傳真費，用郵票或郵資券粘貼於此。</font></td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
								</td>
						</table>
						</td>

				</td>
				
				<td align="left" valign="top"><font color=#ffffff>=</font></td>
				
				<td width="528" align="left" valign="top">
				<table border="1" width="528" id="table7" height="639" cellspacing="0" cellpadding="0">
					<tr>
						<td height="84" width="29" align="center">
						<font face="標楷體">投<br>遞
						<br>局
						<br>(一)</font></td>
						<td height="84" width="483">
						<font face="標楷體">該件於　　年　　月　　日隨第　　號清單第　　頁第　　格　　發<br>往　　貴局投遞(招領)請
						詳查
						<br>　　　年　　月　　日　　　　郵局　經辦員：
						<br>　　　　　　　　　　　　　　　　　主　管：</font></td>
					</tr>
					<tr>
						<td height="78" width="29" align="center">
						<font face="標楷體">投<br>遞
						<br>局
						<br>(二)</font></td>
						<td height="78" width="483"><font face="標楷體">該件於　　年　　月　　日隨第　　號清單第　　頁第　　格　　發<br>往　　貴局投遞(招領)請
						詳查
						<br>　　　年　　月　　日　　　　郵局　經辦員：
						<br>　　　　　　　　　　　　　　　　　主　管：</font></td>
					</tr>
					<tr>
						<td width="29" height="272" align="center">
						<font face="標楷體">投</font><p><font face="標楷體"><br>遞</font></p>
						<p><font face="標楷體">&nbsp;<br>局</font></p>
						<p><font face="標楷體">&nbsp;<br>(三)</font></td>
						<td height="272" width="483"><font face="標楷體">
						茲將最後查得結果說明如下（V）：</font><p><font face="標楷體">
						□一、查該件業於　　年　　月　　日妥投，妥投收據傳真如後，以為投到之據。</font></p>
						<p><font face="標楷體">□二、該件未投遞，原因如左：</font></p>
						<p><font face="標楷體">查該件</font></p>
						<p><font face="標楷體">　　　　　　　　　　　　　　　　經辦員：</font></p>
						<p><font face="標楷體">　　　　　　　　　　　　　　　　主　管：</font></p>
						<p><font face="標楷體">中華民國　　　　　年　　　　　月　　　　　日</font></td>
					</tr>
					<tr>
						<td colspan="2" align="center">
						<table border="0" width="400" id="table8" cellspacing="0" cellpadding="0">
							<tr>
								<td width="97">　</td>
								<td><font face="標楷體">妥投收據(或影本)貼此處</font><p>
								<font face="標楷體">
						一併傳真至原查詢局後，
						</font>
								<p><font face="標楷體">
						收據仍取下存檔。</font></td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<td colspan="2" height="35"><font face="標楷體">　補到回執已收訖：寄件人簽章</font></td>
					</tr>
				</table>
				　</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
</div>
<%			
response.flush
Next
		%>
</body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="../smsx.cab#Version=6,1,432,1">
</object>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
printWindow(true,7,10.08,5.08,0);
</script>
</html>