<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
	'抓縣市
	Server.ScriptTimeout=6000
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
%>
<%if sys_City<>"雲林縣" and sys_City<>"台中縣" and sys_City<>"嘉義縣" then%>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%end if%>
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>交寄大宗函件</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<%
Server.ScriptTimeout = 6800
Response.flush
'權限
'AuthorityCheck(234)
%>
<style type="text/css">
<!--

.style35 {
	font-size: 10pt;
	font-family: "標楷體";
}
.style33 {
	font-size: 9pt;
	font-family: "標楷體";
}
.style5 {
	font-size: 10pt;
	font-family: "標楷體";}
.style7 {
	font-size: 10pt;
	font-family: "標楷體";}
.style8 {
	font-size: 14pt;
	}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style11 {
	font-size: 10px;
	font-family: "標楷體";
}
.style22 {font-size: 9pt; font-family: "標楷體"; }
<%if sys_City="雲林縣" or sys_City="台中縣" or sys_City="嘉義縣" then%>
.pageprint {
  margin-left: 7mm;
  margin-right: 5.08mm;
  margin-top: 5.08mm;
  margin-bottom: 5.08mm;
}
<%end if%>
-->
</style>
</head>

<body>

<%
strwhere=request("SQLstr")

'郵資
theMailMoney=trim(request("MailMoneyValue"))
'使用者單位資料
UnitName=""
UnitAddress=""
UnitTel=""
strUnitName="select Value from ApConfigure where ID=40"
set rsUnitName=conn.execute(strUnitName)
if not rsUnitName.eof then
	TitleUnitName=trim(rsUnitName("value"))
end if
rsUnitName.close
set rsUnitName=nothing

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

ExchangeTypeFlag="W"


	strSendMailUnit="select b.UnitName,b.Address,b.Tel from MemberData a,UnitInfo b " &_
			" where a.MemberID="&trim(Session("User_ID"))&" and a.UnitID=b.UnitID"
	set rsSendMailUnit=conn.execute(strSendMailUnit)
	if not rsSendMailUnit.eof then
		
		UnitName=trim(rsSendMailUnit("UnitName"))
		UnitAddress=trim(rsSendMailUnit("Address"))
		UnitTel=trim(rsSendMailUnit("Tel"))
	end if
	rsSendMailUnit.close
	
	MDate=""

	strMailDate="select g.MailDate as MDate from DciLog b,BillBase a,StopBillMailHistory g " &_
		" where a.Sn=g.BillSn and a.Sn=b.BillSn "&strwhere

	'response.write strMailDate
	set rsMailDate=conn.execute(strMailDate)
	if not rsMailDate.eof then
		MDate=trim(rsMailDate("MDate"))
	end if
	rsMailDate.close
	set rsMailDate=nothing
	if MDate="" or isnull(MDate) then
		MDate=now
	end if

	AllCaseCnt=0
	MailAddress_tmp=""
	BillNo_Tmp=""
	GetMailMem_Tmp=""
	theMailNumber_Tmp=""
'==============================================
BillNo="":CarNo="":mailnumberStr=""

strSQL="select distinct a.ImageFileNameB,a.CarNo,c.MailNumber from (select sn,carno,ImageFileNameB from BillBase where ImagePathName is not null and BillStatus>1 and RecordStateId <> -1 and ImageFileNameB is not null and DeallineDate is not null) a,(Select * from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b,((select BillSN,CarNo,BillNo,MailNumber from StopBillMailHistory where mailnumber is not null) union all (select BillSN,CarNo,BillNo,StoreAndSendMailNumber MailNumber from StopBillMailHistory where StoreAndSendMailNumber is not null) union all (select BillSN,CarNo,BillNo,ThreeMailNumber MailNumber from StopBillMailHistory where ThreeMailNumber is not null)) c where a.SN=b.BillSN "&request("SQLstr")&" and a.sn=c.BillSN order by a.ImageFileNameB,c.MailNumber"
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
	if cmtI<>0 then response.write "<div class=""PageNext""></div>"

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

	Sys_MailNumber="":Sys_MailChkNumber=""

	strSQL="select distinct "&TypeMailNumber&" MailNumber from StopBillMailHistory where BillNo='"&PBillNo(cmtI)&"'"
	set rsmail=conn.execute(strSQL)
	If Not rsmail.eof Then
		Sys_MailNumber=trim(rsmail("MailNumber"))'&"97000717"
		Sys_MailChkNumber=trim(rsmail("MailNumber"))'&"970007 17"
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
	end if
	AllCaseCnt=AllCaseCnt+1
next
'=======================================================================================	

'	GetMailAddress=""
'	GetMailMem=""
'	OwnerAddress1=""
'	OwnerZip1=""
'	DriverAddress2=""
'	DriverZip2=""
'	strSqlD="select DriverHomeZip,nvl(DriverHomeAddress,' ') DriverHomeAddress,Owner,OwnerZip,nvl(OwnerAddress,' ') OwnerAddress,nvl(OWNERNOTIFYADDRESS,' ') OWNERNOTIFYADDRESS from BIllBaseDCIReturn where  CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='A' and Status in ('S')"
'	set rsD=conn.execute(strSqlD)
'	if not rsD.eof Then
			
			
'			if trim(rsD("OwnerAddress"))<>"" and not isnull(rsD("OwnerAddress")) then
'				OwnerAddress1=trim(rsD("OwnerAddress"))
'				OwnerZip1=trim(rsD("OwnerZip"))
'			end if
'			if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeZip")) then
'				DriverAddress2=trim(rsD("DriverHomeAddress"))
'				DriverZip2=trim(rsD("DriverHomeZip"))
'			end if
'
'			if trim(rsD("OWNERNOTIFYADDRESS"))<>"" and not isnull(rsD("OWNERNOTIFYADDRESS")) then
'				NotifyZip=""
'				strNZ="select * from Zip where ZipName like '"&left(trim(rsD("OWNERNOTIFYADDRESS")),5)&"%'"
'				set rsNZ=conn.execute(strNZ)
'				if not rsNZ.eof then
'					NotifyZip=trim(rsNZ("ZipNo"))
'				else
'					strNZ2="select * from Zip where ZipName like '"&left(trim(rsD("OWNERNOTIFYADDRESS")),3)&"%'"
'					set rsNZ2=conn.execute(strNZ2)
'					if not rsNZ2.eof then
'						NotifyZip=trim(rsNZ2("ZipNo"))
'					
'					end if
'					rsNZ2.close
'					set rsNZ2=nothing
'				end if
'				rsNZ.close
'				set rsNZ=nothing
'				if MailAddress_tmp="" then
'					MailAddress_tmp=NotifyZip&trim(rsD("OWNERNOTIFYADDRESS"))
'				else
'					MailAddress_tmp=MailAddress_tmp&"@!#"&NotifyZip&trim(rsD("OWNERNOTIFYADDRESS"))
'				end if
'			ElseIf trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) Then
'				if MailAddress_tmp="" then
'						
'					MailAddress_tmp=trim(DriverZip2)&ZipName&trim(DriverAddress2)
'				else
'					MailAddress_tmp=MailAddress_tmp&"@!#"&trim(DriverZip2)&ZipName&trim(DriverAddress2)
'				end if
'			ElseIf trim(rsD("OwnerAddress"))<>"" and not isnull(rsD("OwnerAddress")) Then
'				if MailAddress_tmp="" then
'						
'					MailAddress_tmp=trim(OwnerZip1)&ZipName&trim(OwnerAddress1)
'				else
'					MailAddress_tmp=MailAddress_tmp&"@!#"&trim(OwnerZip1)&ZipName&trim(OwnerAddress1)
'				end if
'			end if
'			if BillNo_Tmp="" then
'				BillNo_Tmp=cdbl(rs1("ImageFileNameB"))
'			else
'				BillNo_Tmp=BillNo_Tmp&"@!#"&cdbl(rs1("ImageFileNameB"))
'			end if
'			if GetMailMem_Tmp="" then
'				GetMailMem_Tmp=trim(rsD("Owner"))
'			else
'				GetMailMem_Tmp=GetMailMem_Tmp&"@!#"&trim(rsD("Owner"))
'			end if
'			if theMailNumber_Tmp="" then
'				theMailNumber_Tmp=theMailNumber1
'			else
'				theMailNumber_Tmp=theMailNumber_Tmp&"@!#"&theMailNumber1
'			end if
'			AllCaseCnt=AllCaseCnt+1
'	end if
'	rsD.close
'	set rsD=nothing



MailAddress_Array=split(MailAddress_tmp&" ","@!#")
theMailNumber_Array=split(theMailNumber_Tmp&" ","@!#")
GetMailMem_Tmp_Array=split(GetMailMem_Tmp&" ","@!#")
BillNo_Tmp_Array=split(BillNo_Tmp&" ","@!#")
CaseSN=0
mailSNTmp=0
pagecnt=fix(AllCaseCnt/20+0.9999999)

for MAcnt=0 to ubound(MailAddress_Array)
	'if MAcnt=ubound(MailAddress_Array) then exit for
	
if mailSN>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
	strList=""
	mailSN=0
	MailCnt=0
	strTmp=""
	pageNum=fix(CaseSN/20)+1
	for MAloop=1 to 20
		if MAcnt>ubound(MailAddress_Array) then exit for

		mailSN=mailSN+1
		CaseSN=CaseSN+1
		MailCnt=MailCnt+1

		strList=strList&"<tr height=""23"">"
		'順序號碼
		strList=strList&"<td align=""center"">"&mailSN&"</td>"

		
		strList=strList&"<td align=""center"">"&theMailNumber_Array(MAcnt)&"</td>"
		

		'收件人姓名
		strList=strList&"<td align=""center"" width=""100"">&nbsp;</td>"
		strList=strList&"<td align=""left"" width=""100""class=""style35"">"&funcCheckFont(GetMailMem_Tmp_Array(MAcnt),14,1)&"</td>"
		strList=strList&"<td align=""center"">&nbsp;</td>"
		'收件地址
		strList=strList&"<td align=""left"" class=""style35"" width=""400"">"&funcCheckFont(MailAddress_Array(MAcnt),14,1)&"</td>"

		strList=strList&"<td align=""center"">&nbsp;</td>"
		'郵資
		if theMailMoney<>"" then
			theMailMoneyTmp=theMailMoney
		else
			theMailMoneyTmp="&nbsp;"
		end if
		strList=strList&"<td align=""center"" width=""20"">"&theMailMoneyTmp&"</td>"
		'備考=單號
		strList=strList&"<td align=""center"">&nbsp;</td>"
		strList=strList&"<td align=""left"">"&BillNo_Tmp_Array(MAcnt)&"</td>"
		strList=strList&"</tr>"
		if MAloop<20 then
			MAcnt=MAcnt+1
		end If
		Response.flush
	next
	if mailSN<20 then
		mailSNTmp=mailSN

		for Sp=1 to 20-mailSN
			mailSNTmp=mailSNTmp+1
			strList=strList&"<tr height=""23"">"

			'順序號碼
			strList=strList&"<td align=""center"">"&mailSNTmp&"</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"</tr>"
			response.flush
		next
	end if


%>
<table width="710" align="center"  border="0">
	<tr>
		<td height="40" valign="top"><span class="style7">
<%if sys_City="澎湖縣" then %>	
		<span class="style8">□□□□□□ □□</span>
		<br>
		 &nbsp; &nbsp; &nbsp;收寄局碼&nbsp; &nbsp;郵件種類碼
		 <br>
		 &nbsp; &nbsp; &nbsp; &nbsp;(由收寄局填寫)
<%end if%>
		<br>
<%if sys_City<>"雲林縣" and sys_City<>"花蓮縣" then %>
		中華民國 <%
		response.write year(MDate)-1911
		%>年 <%
		response.write right("00"&month(MDate),2)
		%>月 <%
		response.write right("00"&day(MDate),2)
		%>日
<%end if%>
		<br>
<%if sys_City<>"澎湖縣" then %>	
		移送監理站日期 <%
			if theSendDocDate<>"" then
				if len(theSendDocDate)=6 then
					response.write left(theSendDocDate,2)
				elseif len(theSendDocDate)=7 then
					response.write left(theSendDocDate,3)
				end if
			end if
		%>年 <%
			if theSendDocDate<>"" then
				if len(theSendDocDate)=6 then
					response.write mid(theSendDocDate,3,2)
				elseif len(theSendDocDate)=7 then
					response.write mid(theSendDocDate,4,2)
				end if
			end if
		%>月 <%
			if theSendDocDate<>"" then
				if len(theSendDocDate)=6 then
					response.write mid(theSendDocDate,5,2)
				elseif len(theSendDocDate)=7 then
					response.write mid(theSendDocDate,6,2)
				end if
			end if
		%>日
		<br>
<%end if%>

		</span></td>
	</tr>

	<tr><td><span class="style7">  <% response.write UnitName %> </span> </td>
	    <td> <span class="style7"><%response.write year(now)-1911
		%>年 <%
		response.write right("00"&month(now),2)
		%>月 <%
		response.write right("00"&day(now),2)
		%>日</span> 
	  
	   <td>
		<td width="34%"><span class="style7">
		<%=pageNum%> of <%=pagecnt%>
		</span></td>	
	</tr>	
	<tr>
	</tr>

	</table>
</td>
</tr>
<tr>
<td>
	<table align="center" width="100%" border="0" cellspacing="0" cellpadding="3">
	

	<%=strList%>
	</table>
</td>
</tr>

<tr>
<td>
	<table align="center" width="710" border="0">
	<tr>
	<td width="34%" class="style5" valign="Top">
	<br>
	    掛號函件/共 
	    <%=MailCnt%> 
	    件照收無誤
 ( 
	    
	   郵資共計  
	    <%
		if theMailMoney<>"" then
			response.write theMailMoney*MailCnt
		else
			response.write "&nbsp;"
		end if
		%> 
	    元 
		)	
	  </td>
	</tr>
	</table>
</td>
</tr>

</table>

<%		
next

%>

</body>

<script language="javascript">
<%if sys_City="雲林縣" or sys_City="台中縣" or sys_City="嘉義縣" or sys_City="花蓮縣" then%>
window.print();
<%else%>
printWindow(true,7,5.08,5.08,5.08);
<%end if%>
</script>
</html>
