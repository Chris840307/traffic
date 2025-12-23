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
strwhere=" and a.ImageFileNameB in ('" & replace(trim(request("PBillNo")),",","','") & "') "

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

	strSQL="select a.*,b.ExchangeTypeID,b.BatchNumber from (select ImageFileNameB,Max(SN) as SN" &_
	" from Billbase  where RecordStateID=0 group by ImageFileNameB) k,BillBase a,DCILog b" &_
	" where a.SN=b.BillSN and k.SN=a.SN" &_
	" and a.RecordStateID=0" &_
	" and (b.ExchangeTypeID='A' and b.DciReturnStatusID in ('S'))" &_
	" "&strwhere&" order by a.ImageFileNameB"
	set rs1=conn.execute(strSQL)

	strCnt="select count(distinct(ImageFileNameB)) as cnt" &_
	" from BillBase a,DCILog b" &_
	" where a.SN=b.BillSN" &_
	" and a.RecordStateID=0" &_
	" and (b.ExchangeTypeID='A' and b.DciReturnStatusID in ('S'))" &_
	" "&strwhere
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		if trim(rsCnt("cnt"))="0" then
			pagecnt=1
		else
			pagecnt=fix((Cint(rsCnt("cnt")))/20+0.9999999)
		end if

	end if
	rsCnt.close
	set rsCnt=nothing
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

CaseSN=0
mailSNTmp=0
If Not rs1.Bof Then rs1.MoveFirst 
While Not rs1.Eof
if mailSN>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
	BillFillDateTmp=""
	if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
		BillFillDateTmp=trim(rs1("BillFillDate"))
	end if
	strList=""
	mailSN=0
	MailCnt=0
	
	pageNum=fix(CaseSN/20)+1
	for i=1 to 20
		if rs1.eof then exit for
		MailBatchNumber=trim(rs1("BatchNumber"))
		mailSN=mailSN+1
		CaseSN=CaseSN+1
		MailCnt=MailCnt+1

		SendAddrFlag="":HomeAddress="":BillCnt1=0:BillCnt2=0
		strSendSQL="select SendAddrFlag,HomeAddress from StopCaseSendAddr where CarNo='"&trim(rs1("CarNo"))&"'"
		set rsSend=conn.execute(strSendSQL)
		If Not rsSend.eof Then
			SendAddrFlag=trim(rsSend("SendAddrFlag"))
			HomeAddress=trim(rsSend("HomeAddress"))
		end if
		rsSend.Close

		If SendAddrFlag="" or SendAddrFlag="1" Then
			BillCnt1=1:BillCnt2=1
		elseif SendAddrFlag="2" then
			BillCnt1=2:BillCnt2=2
		elseif SendAddrFlag="3" then
			BillCnt1=1:BillCnt2=2
		End if

		if strTmp<>"" and i=1 then
			strList=strList&strTmp
			i=i+1
			mailSN=mailSN+1
			CaseSN=CaseSN+1
			MailCnt=MailCnt+1
			strTmp=""
		end if

		strList=strList&"<tr height=""23"">"
		'順序號碼
		strList=strList&"<td align=""center"">"&mailSN&"</td>"

		'掛號號碼
		theMailNumber=""
		'移送監理站日期
		theSendDocDate=""
		strSqlH="select MailNumber,StoreAndSendMailNumber,SendOpenGovDocToStationDate from StopBillMailHistory where BillSN="&trim(rs1("SN"))
		set rsH=conn.execute(strSqlH)
		if not rsH.eof then
			if trim(rsH("SendOpenGovDocToStationDate"))<>"" and not isnull(rsH("SendOpenGovDocToStationDate")) then
				theSendDocDate=trim(rsH("SendOpenGovDocToStationDate"))
			end if
			'車籍抓MailNumber,駕籍抓StoreandSendMailNumber
'			if SendAddrFlag="" or SendAddrFlag="1" or SendAddrFlag="3" then
'				if trim(rsH("MailNumber"))<>"" and not isnull(rsH("MailNumber")) then
'					theMailNumber=trim(rsH("MailNumber"))&"&nbsp;"
'				else
'					theMailNumber="&nbsp;"
'				end if
'			else
				if trim(rsH("StoreandSendMailNumber"))<>"" and not isnull(rsH("StoreandSendMailNumber")) then
					theMailNumber=trim(rsH("StoreandSendMailNumber"))&"&nbsp;"
				else
					theMailNumber="&nbsp;"
				end if
'			end if
		else
			theMailNumber="&nbsp;"
		end if
		rsH.close
		set rsH=nothing
		strList=strList&"<td align=""center"">"&theMailNumber&"</td>"
		MailAddress_tmp=""
		GetMailAddress=""
		GetMailMem=""
		strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress,OWNERNOTIFYADDRESS from BIllBaseDCIReturn where  CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='A' and Status in ('S')"
		set rsD=conn.execute(strSqlD)
		if not rsD.eof then
			'if SendAddrFlag="" or SendAddrFlag="1" or SendAddrFlag="3" then
				IF trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
					strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
					set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName=trim(rsZip("ZipName"))
				
					end if
					rsZip.close
					set rsZip=nothing
					MailAddress_tmp=trim(rsD("DriverHomeAddress"))
					GetMailAddress="(戶)"&trim(rsD("DriverHomeZip"))&ZipName&replace(replace(trim(rsD("DriverHomeAddress")),"臺","台"),ZipName,"")
				end if
			'else
			'	GetMailAddress=HomeAddress
			'end if
			GetMailMem=trim(rsD("Owner"))
		end if
		rsD.close
		set rsD=nothing

		'收件人姓名
		strList=strList&"<td align=""center"" width=""100"">&nbsp;</td>"
		strList=strList&"<td align=""left"" width=""100""class=""style35"">"&GetMailMem&"</td>"
		strList=strList&"<td align=""center"">&nbsp;</td>"
		'收件地址
		strList=strList&"<td align=""left"" class=""style35"" width=""400"">"&GetMailAddress&"</td>"

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
		strList=strList&"<td align=""left"">"&cdbl(rs1("ImageFileNameB"))&"</td>"
		strList=strList&"</tr>"

		strTmp=""
		GetMailAddress=""
		isMail2=0
		strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress,OWNERNOTIFYADDRESS from BIllBaseDCIReturn where  CarNo='"&trim(rs1("CarNo"))&"' and ExchangeTypeID='A' and Status in ('S')"
		set rsD=conn.execute(strSqlD)
		if not rsD.eof then
			'strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
			'set rsZip=conn.execute(strZip)
			'if not rsZip.eof then
				'ZipName=trim(rsZip("ZipName"))
				ZipName=""
			'end if
			'rsZip.close
			'set rsZip=nothing
			if MailAddress_tmp=trim(rsD("DriverHomeAddress")) then
				isMail2=1
			end if
			GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&trim(rsD("DriverHomeAddress"))

			'GetMailAddress=HomeAddress
			GetMailMem=trim(rsD("Owner"))
		end if
		rsD.close
		set rsD=nothing
'		if trim(GetMailAddress)<>"" and isMail2=0 then
'			CaseSN=CaseSN+1
'			mailSN=mailSN+1
'			if mailSN<=20 then
'				i=i+1
'				mailSN1=mailSN
'				MailCnt=MailCnt+1
'			else
'				i=i
'				mailSN1=1
'				mailSN=mailSN
'			end if
'			strTmp=strTmp&"<tr height=""23"">"
'			'順序號碼
'			strTmp=strTmp&"<td align=""center"">"&mailSN1&"</td>"
'
'			'掛號號碼
'			theMailNumber=""
'			'移送監理站日期
'			theSendDocDate=""
'			strSqlH="select MailNumber,StoreAndSendMailNumber,SendOpenGovDocToStationDate from StopBillMailHistory where BillSN="&trim(rs1("SN"))
'			set rsH=conn.execute(strSqlH)
'			if not rsH.eof then
'				if trim(rsH("SendOpenGovDocToStationDate"))<>"" and not isnull(rsH("SendOpenGovDocToStationDate")) then
'					theSendDocDate=trim(rsH("SendOpenGovDocToStationDate"))
'				end if
'				if trim(rsH("StoreandSendMailNumber"))<>"" and not isnull(rsH("StoreandSendMailNumber")) then
'					theMailNumber=trim(rsH("StoreandSendMailNumber"))&"&nbsp;"
'				else
'					theMailNumber="&nbsp;"
'				end if
'			else
'				theMailNumber="&nbsp;"
'			end if
'			rsH.close
'			set rsH=nothing
'			strTmp=strTmp&"<td align=""center"">"&theMailNumber&"</td>"
'			
'			'收件人姓名
'			strTmp=strTmp&"<td align=""center"" width=""100"">&nbsp;</td>"
'			strTmp=strTmp&"<td align=""left"" width=""100""class=""style35"">"&GetMailMem&"</td>"
'			strTmp=strTmp&"<td align=""center"">&nbsp;</td>"
'			'收件地址
'			strTmp=strTmp&"<td align=""left"" class=""style35"" width=""400"">"&GetMailAddress&"</td>"
'
'			strTmp=strTmp&"<td align=""center"">&nbsp;</td>"
'			'郵資
'			if theMailMoney<>"" then
'				theMailMoneyTmp=theMailMoney
'			else
'				theMailMoneyTmp="&nbsp;"
'			end if
'			strTmp=strTmp&"<td align=""center"" width=""20"">"&theMailMoneyTmp&"</td>"
'			'備考=單號
'			strTmp=strTmp&"<td align=""center"">&nbsp;</td>"
'			strTmp=strTmp&"<td align=""left"">"&cdbl(rs1("ImageFileNameB"))&"</td>"
'			strTmp=strTmp&"</tr>"
'			if mailSN<=20 then
'				'response.write mailSN
'				strList=strList&strTmp
'				strTmp=""
'			end if
'		end if

		rs1.MoveNext
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
<%if sys_City="台南市" then %>	
		填單日期 <%
			if BillFillDateTmp<>"" then
				response.write year(BillFillDateTmp)-1911&"年 "
			end if
			if BillFillDateTmp<>"" then
				response.write month(BillFillDateTmp)&"月 "
			end if
			if BillFillDateTmp<>"" then
				response.write day(BillFillDateTmp)&"日"
			end if
		%>
<%elseif sys_City<>"澎湖縣" then %>	
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
	
Wend
rs1.close
set rs1=nothing

'剛好SendAddrFlag="3"在最後一筆
if strTmp<>"" then
	mailSNTmp2=1
	for Sp=1 to 19
			mailSNTmp2=mailSNTmp2+1
			strTmp=strTmp&"<tr height=""23"">"

			'順序號碼
			strTmp=strTmp&"<td align=""center"">"&mailSNTmp2&"</td>"
			strTmp=strTmp&"<td align=""center"">&nbsp;</td>"
			strTmp=strTmp&"<td align=""center"">&nbsp;</td>"
			strTmp=strTmp&"<td align=""center"">&nbsp;</td>"
			strTmp=strTmp&"<td align=""center"">&nbsp;</td>"
			strTmp=strTmp&"<td align=""center"">&nbsp;</td>"
			strTmp=strTmp&"<td align=""center"">&nbsp;</td>"
			strTmp=strTmp&"<td align=""center"">&nbsp;</td>"
			strTmp=strTmp&"<td align=""center"">&nbsp;</td>"
			strTmp=strTmp&"<td align=""center"">&nbsp;</td>"
			strTmp=strTmp&"</tr>"
	next
%>
<div class="PageNext">&nbsp;</div>
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
<%if sys_City="台南市" then %>	
		填單日期 <%
			if BillFillDateTmp<>"" then
				response.write year(BillFillDateTmp)-1911&"年 "
			end if
			if BillFillDateTmp<>"" then
				response.write month(BillFillDateTmp)&"月 "
			end if
			if BillFillDateTmp<>"" then
				response.write day(BillFillDateTmp)&"日"
			end if
		%>
<%elseif sys_City<>"澎湖縣" then %>	
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
		<%=pageNum+1%> of <%=pagecnt%>
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
	

	<%=strTmp%>
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
	    <%=1%> 
	    件照收無誤
 ( 
	    
	   郵資共計  
	    <%
		if theMailMoney<>"" then
			response.write theMailMoney
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
end if
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
