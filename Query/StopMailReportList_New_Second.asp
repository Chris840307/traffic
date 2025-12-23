<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
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
Server.ScriptTimeout = 800
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
	set rsSendMailUnit=nothing

	Sys_SendMarkDate1=gOutDT(request("Sys_SendMarkDate1"))&" 0:0:0"
	Sys_SendMarkDate2=gOutDT(request("Sys_SendMarkDate2"))&" 23:59:59"

	strSQL="select a.* from (select ImageFileNameB,Max(SN) as SN" &_
	" from Billbase  where RecordStateID=0 group by ImageFileNameB) k,BillBase a,stopcarsendaddress b" &_
	" where a.ImageFileNameB=b.BillNo and k.SN=a.SN" &_
	" and a.RecordStateID=0" &_
	" and b.UserMarkDate between TO_DATE('"&Sys_SendMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') " &_
	" and TO_DATE('"&Sys_SendMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')" &_
	" order by a.SN,a.ImageFileNameB,a.CarNo"
	set rs1=conn.execute(strSQL)
	MDate=""
	If Not rs1.Bof Then 
		strMailDate="select StoreAndSendMailDate as MDate from StopBillMailHistory " &_
			" where BillSn="&Trim(rs1("Sn"))

		'response.write strMailDate
		set rsMailDate=conn.execute(strMailDate)
		if not rsMailDate.eof then
			MDate=trim(rsMailDate("MDate"))
		end if
		rsMailDate.close
		set rsMailDate=Nothing
	End if
	if MDate="" or isnull(MDate) then
		MDate=now
	end if

	AllCaseCnt=0
	MailAddress_tmp=""
	BillNo_Tmp=""
	GetMailMem_Tmp=""
	theMailNumber_Tmp=""
If Not rs1.Bof Then rs1.MoveFirst 
While Not rs1.Eof
	'掛號號碼
	theMailNumber1=""
	theMailNumber2=""
	theMailNumber3=""
	'移送監理站日期
	theSendDocDate=""
	strSqlH="select MailNumber,StoreAndSendMailNumber,SendOpenGovDocToStationDate,DriverMailNumber from StopBillMailHistory where BillSN="&trim(rs1("SN"))
	set rsH=conn.execute(strSqlH)
	if not rsH.eof then
'		if trim(rsH("SendOpenGovDocToStationDate"))<>"" and not isnull(rsH("SendOpenGovDocToStationDate")) then
'			theSendDocDate=trim(rsH("SendOpenGovDocToStationDate"))
'		end if
		if trim(rsH("StoreAndSendMailNumber"))<>"" and not isnull(rsH("StoreAndSendMailNumber")) then
			theMailNumber1=trim(rsH("StoreAndSendMailNumber"))&"&nbsp;"
		else
			theMailNumber1="&nbsp;"
		end if
	else
		theMailNumber1="&nbsp;"
		theMailNumber2="&nbsp;"
		theMailNumber3="&nbsp;"
	end if
	rsH.close
	set rsH=nothing

	
	
	if MailAddress_tmp="" then
		If trim(rs1("DriverAddress"))<>"" and not isnull(rs1("DriverAddress")) Then
			MailAddress_tmp=trim(rs1("DriverZip"))&trim(rs1("DriverAddress"))
		Else 
			MailAddress_tmp="&mbsp;"
		end if
	Else
		If trim(rs1("DriverAddress"))<>"" and not isnull(rs1("DriverAddress")) Then
			MailAddress_tmp=MailAddress_tmp&"@!@"&trim(rs1("DriverZip"))&trim(rs1("DriverAddress"))
		Else 
			MailAddress_tmp=MailAddress_tmp&"@!@&mbsp;"
		end if
		
	end if
	
	if BillNo_Tmp="" then
		BillNo_Tmp=cdbl(rs1("ImageFileNameB"))
	else
		BillNo_Tmp=BillNo_Tmp&"@!@"&cdbl(rs1("ImageFileNameB"))
	end if
	if GetMailMem_Tmp="" Then
		If trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) Then
			GetMailMem_Tmp=trim(rs1("Owner"))
		Else 
			GetMailMem_Tmp="&mbsp;"
		end if
	Else
		If trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) Then
			GetMailMem_Tmp=GetMailMem_Tmp&"@!@"&trim(rs1("Owner"))
		Else 
			GetMailMem_Tmp=GetMailMem_Tmp&"@!@&mbsp;"
		end if
		
	end if
	if theMailNumber_Tmp="" then
		theMailNumber_Tmp=theMailNumber1
	else
		theMailNumber_Tmp=theMailNumber_Tmp&"@!@"&theMailNumber1
	end if
	AllCaseCnt=AllCaseCnt+1
	
	rs1.MoveNext
Wend



MailAddress_Array=split(MailAddress_tmp,"@!@")
theMailNumber_Array=split(theMailNumber_Tmp,"@!@")
GetMailMem_Tmp_Array=split(GetMailMem_Tmp,"@!@")
BillNo_Tmp_Array=split(BillNo_Tmp,"@!@")
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
		end if
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
			
		%>年 <%
			
		%>月 <%
			
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

rs1.close
set rs1=nothing

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
