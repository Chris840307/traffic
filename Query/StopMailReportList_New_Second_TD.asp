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
<%if sys_City="新北市" then %>
<script type="text/javascript" src="../js/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../js/jquery-barcode-2.0.2.min.js"></script>
<%End If %>
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
	font-size: 9pt;
	font-family: "標楷體";
}
.style33 {
<%if sys_City="台東縣" then%>
	font-size: 8pt;
<%else%>
	font-size: 9pt;
<%end if%>
	line-height:10pt;
	font-family: "標楷體";
}
.style5 {
	font-size: 10pt;
	font-family: "標楷體";}
.style7 {
<%if sys_City="台東縣" then%>
	font-size: 9pt;
<%else%>
	font-size: 10pt;
<%end if%>
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
<%if sys_City="台東縣" then%>
	font-size: 10px;
<%else%>
	font-size: 10px;
<%end if%>
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
	" order by a.ImageFileNameB,a.CarNo"
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
		strList=strList&"<td align=""left"" width=""100""class=""style35"">"&funcCheckFont(GetMailMem_Tmp_Array(MAcnt),14,1)&"</td>"
		
		'收件地址
		strList=strList&"<td align=""left"" class=""style35"" width=""400"">"&funcCheckFont(MailAddress_Array(MAcnt),14,1)&"</td>"
		strList=strList&"<td align=""center"">&nbsp;</td>"
		strList=strList&"<td align=""center"">&nbsp;</td>"
		strList=strList&"<td align=""center"">&nbsp;</td>"
		strList=strList&"<td align=""center"">&nbsp;</td>"
		'郵資
		if theMailMoney<>"" then
			theMailMoneyTmp=theMailMoney
		else
			theMailMoneyTmp="&nbsp;"
		end if
		strList=strList&"<td align=""center"" width=""20"">"&theMailMoneyTmp&"</td>"
		'備考=單號
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

if sys_City="南投縣" or sys_City="雲林縣" or sys_City="宜蘭縣" then 
	ReportCount=3
elseif sys_City="花蓮縣" or sys_City="嘉義縣" then 
	ReportCount=1
else
	ReportCount=2
end if
if sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕" then 
	ReportCount=1
end if

%>
<%if sys_City="新北市" then %>

<script type="text/javascript">
      $(function(){
	<% for Bi=1 to ReportCount
			BarCodeName="bcTarget"&pageNum&Bi
	%>
			$("#<%=BarCodeName%>").barcode("<%=MailBatchNumber%>", "code128",{barWidth:1, barHeight:30,fontSize:12,showHRI:true,bgColor:"#FFFFFF"});
	<%next%>
      });
</script>
<%End if%>
<table width="710" align="center"  border="0">
<tr>

<td>
	<table width="100%" align="center" cellpadding="3" border="0">
<%if sys_City<>"花蓮縣" and sys_City<>"嘉義縣" and sys_City<>"台東縣" then %>
	<tr>
		<td height="25"></td>
	</tr>
<%end if%>

	<tr>
<%if sys_City<>"花蓮縣" then %>
		<td width="34%"><span class="style7">
		頁&nbsp;&nbsp;次 &nbsp;<%=pageNum%> of <%=pagecnt%>
		
		</span></td>

		<td rowspan="3" width="39%" align="center"><span class="style7">

		<table width="100%">
	
		<tr>
			<td colspan="3" height="30"><div align="center"><u><span class="style6">中 華 郵 政</span></u><%
		If sys_City="新北市" Then
			%><div id="<%
			response.write "bcTarget"&pageNum&"1"
			%>" style= "position:absolute;width:400px;height:155px;z-index:1"></div><%
		End If 
			%></div></td> 
		</tr>
		<%If sys_City="台東縣" Then %>
		<div id="num30" style="position:absolute; left:1;top:50;font-size: 36pt;line-height: 50pt;">
			<font face="標楷體"><b><%=RIGHT("000" &pageNum,3)%></b></font>
		<div>
		<%end if%>

		<tr>			
			<td width="37%" rowspan="3" align="right" class="style7">交寄大宗</td>
			<td width="26%" class="style7"><u>限時掛號</u></td>
			<td width="37%" rowspan="3" align="left" class="style7">函件執據</td>
		</tr>

		<tr>
			<td class="style7"><u>掛 &nbsp; &nbsp;號</u></td>
		</tr>
		<tr>
			<td class="style7"><u>快捷郵件</u></td>
		</tr>
<%end if%>
		</table>

	<%if sys_City<>"花蓮縣" then %>	
		</span></td>
		<td rowspan="3" width="27%"><div align="right"><img src="../Image/MailPic.JPG" width="100" height="70" /></div></td>
	<%end if%>

	</tr>

	<tr>
		<td height="40" valign="top"><span class="style7">

<%if sys_City="澎湖縣" then %>	
		<span class="style8">□□□□□□ □□</span>
		<br>
		 &nbsp; &nbsp; &nbsp;收寄局碼&nbsp; &nbsp;郵件種類碼
		 <br>
		 &nbsp; &nbsp; &nbsp; &nbsp;(由收寄局填寫)
		 <br>
<%end if%>		
<%if sys_City="台東縣" or sys_City="台南市" or sys_City="澎湖縣" then%>
		中華民國 <%
		response.write year(now)-1911
		%>年 <%
		response.write right("00"&month(now),2)
		%>月 <%
		response.write right("00"&day(now),2)
		%>日

<%elseif sys_City<>"雲林縣" and sys_City<>"花蓮縣" then %>
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
		<%
	if sys_City="南投縣" or sys_City="基隆市" or sys_City="台東縣"  then
			response.write "作業批號："&MailBatchNumber
	end if
		%>

		</span>

		</td>

	</tr>
<%if sys_City<>"花蓮縣" then %>	
	<tr>
		<td><span class="style7">
		寄件人 <%
		response.write UnitName
		%>
		</span></td>
	</tr>

	<tr>
		<td><span class="style7">
		寄件人代表 ___________
		</span></td>
		<td><span class="style7">
		詳細地址：<u><%=UnitAddress%></u>
		</span></td>
		<td><span class="style7">
		電話號碼：<u><%=UnitTel%></u>
		</span></td>
	</tr>

<%else%>
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
<%end if%>
	</table>

</td>
</tr>
<tr>
<td>
    <%if sys_City<>"花蓮縣" then%>	
	<table align="center" width="100%" border="1" cellspacing="0" cellpadding="3">
	
    <%else%>
	<table align="center" width="100%" border="0" cellspacing="0" cellpadding="3">
	
    <%end if%>
   <tr>
    <%if sys_City<>"花蓮縣" then%>
	<td width="6%" rowspan="2"><div align="center"><span class="style5">順序<br>
	  號碼</span></div></td>
   
	<td width="10%" rowspan="2"><div align="center"><span class="style5">掛號號碼</span></div></td>
	<td colspan="2"><div align="center"><span class="style5">收件人</span></div></td>

	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  回執<br>[V]</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  航空<br>[V]</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  印刷<br>[V]</span></div></td>
	<td width="3%" rowspan="2"><div align="center"><span class="style5">重量</span></div></td>

	<td width="6%" rowspan="2"><div align="center"><span class="style5">郵資</span></div></td>
	<td width="9%" rowspan="2"><div align="center"><span class="style5">備考</span></div></td>
<%end if%>
	</tr>
	<tr>
<%if sys_City<>"花蓮縣" then%>
	<td width="15%" class="style5"><div align="center">姓名</div></td>
	<td width="36%" class="style5"><div align="center">送達地名(或地址)</div></td>
<%end if%>
	</tr>
	<%=strList%>
	</table>
</td>
</tr>

<tr>
<td>
	<table align="center" width="100%" border="0">
	<tr>
<%if sys_City<>"花蓮縣" then%>
	<td width="66%" valign="top">
	  <p><span class="style11">(1) 限時掛號、掛號函件與快捷郵件不得同列一單，請將標題塗去其二。<br>
	    (2) 函件背面應註明順序號碼，並按號碼次序排齊滿二十件為一組分組交寄。<br>
	    (3) 將本埠與外埠函件分別列單交寄。
	    <br>
	    (4)如有證明郵資、重量必要者，應由寄件人自行在聯單相關欄內分別註明，並結填總郵資，交郵局</span><span class="style11">經辦員逐件核對。<br>
	    (5) 日後如須查詢，應於交寄日起六個月內檢同原件封面式樣向原寄局為之，並將本執據送驗。<br>
	    (6) 錢鈔或有價證券請利用報值或保價交寄。</span><br>
	    
	      </p>
	  </td>
<%end if%>

	<td width="34%" class="style5" valign="Top">
<%if sys_City<>"花蓮縣" then%>
	  <p>限時掛號<br>
<%else%>
	<br>
<%end if%>
	    掛號函件/共 
	    <%=mailSN%> 
	    件照收無誤
<%if sys_City<>"花蓮縣" then%>
		<br>
	    快捷郵件<br>
		<%if sys_City<>"台東縣" then%>
		<br>
		<%end if%>
<%else%>
 ( 
<%end if%>	    
	    
	   郵資共計  
	    <%
		if theMailMoney<>"" then
			response.write theMailMoney*mailSN
		else
			response.write "&nbsp;"
		end if
		%> 
	    元 
	  <%if sys_City<>"花蓮縣" then%>
		</p><p align="right">______________<br>經辦員簽署&nbsp; </p>
	  <%else%>
		)	
	  <%end if%>
	  </td>
	</tr>
	</table>
</td>
</tr>

</table>


<%if ReportCount>1 then %>
<div class="PageNext">&nbsp;</div>



<table width="710" align="center">
<tr>
<td>
	<table width="100%" align="center" cellpadding="3" border="0">
<%if sys_City<>"嘉義縣" and sys_City<>"台東縣" then%>
	<tr>
		<td height="25"></td>
	</tr>
<%end if%>
	<tr>
		<td width="34%"><span class="style7">
		頁&nbsp;&nbsp;次 &nbsp;<%=pageNum%> of <%=pagecnt%>
		</span></td>
		<td rowspan="3" width="39%" align="center"><span class="style7">
		<table width="100%">
		<tr>
			<td colspan="3" height="30"><div align="center"><u><span class="style6">中 華 郵 政</span></u><%
		If sys_City="新北市" Then
			%><div id="<%
			response.write "bcTarget"&pageNum&"2"
			%>" style= "position:absolute;width:400px;height:155px;z-index:1"></div><%
		End If 
			%></div></td> 
		</tr>
		<%If sys_City="台東縣" Then %>
		<div id="num30" style="position:absolute; left:70;top:50;font-size: 36pt;line-height: 50pt;">
			<font face="標楷體"><b><%=RIGHT("000" &pageNum,3)%></b></font>
		<div>
		<%end if%>
		<tr>
			<td width="37%" rowspan="3" align="right" class="style7">交寄大宗</td>
			<td width="26%" class="style7"><u>限時掛號</u></td>
			<td width="37%" rowspan="3" align="left" class="style7">函件存根</td>
		</tr>
		<tr>
			<td class="style7"><u>掛 &nbsp; &nbsp;號</u></td>
		</tr>
		<tr>
			<td class="style7"><u>快捷郵件</u></td>
		</tr>
		</table>
		
		</span></td>
		<td rowspan="3" width="27%"><div align="right"><img src="../Image/MailPic.JPG" width="100" height="70" /></div></td>
	</tr>
	<tr>
		<td height="40" valign="top"><span class="style7">
<%if sys_City="澎湖縣" then %>	
		<span class="style8">□□□□□□ □□</span>
		<br>
		 &nbsp; &nbsp; &nbsp;收寄局碼&nbsp; &nbsp;郵件種類碼
		 <br>
		 &nbsp; &nbsp; &nbsp; &nbsp;(由收寄局填寫)
		 <br>
<%end if%>		
<%if sys_City="台東縣" or sys_City="台南市" or sys_City="澎湖縣" then%>
		中華民國 <%
		response.write year(now)-1911
		%>年 <%
		response.write right("00"&month(now),2)
		%>月 <%
		response.write right("00"&day(now),2)
		%>日
<%elseif sys_City<>"雲林縣" and sys_City<>"花蓮縣" then %>
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
		<%
	if sys_City="南投縣"  or sys_City="基隆市" or sys_City="台東縣"  then
			response.write "作業批號："&MailBatchNumber
	end if
		%>
		</span></td>
	</tr>
	<tr>
		<td><span class="style7">
		寄件人 <%=UnitName%>
		</span></td>
	</tr>
	<tr>
		<td><span class="style7">
		寄件人代表 ___________
		</span></td>
		<td><span class="style7">
		詳細地址：<u><%=UnitAddress%></u>
		</span></td>
		<td><span class="style7">
		電話號碼：<u><%=UnitTel%></u>
		</span></td>
	</tr>
	</table>
</td>
</tr>
<tr>
<td>
	<table align="center" width="100%" border="1" cellspacing="0" cellpadding="3">
	<tr>
	<td width="6%" rowspan="2"><div align="center"><span class="style5">順序<br>
	  號碼</span></div></td>
	<td width="10%" rowspan="2"><div align="center"><span class="style5">掛號號碼</span></div></td>
	<td colspan="2"><div align="center"><span class="style5">收件人</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  回執<br>[V]</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  航空<br>[V]</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  印刷<br>[V]</span></div></td>
	<td width="3%" rowspan="2"><div align="center"><span class="style5">重量</span></div></td>
	<td width="6%" rowspan="2"><div align="center"><span class="style5">郵資</span></div></td>
	<td width="9%" rowspan="2"><div align="center"><span class="style5">備考</span></div></td>
	</tr>
	<tr>
	<td width="15%" class="style5"><div align="center">姓名</div></td>
	<td width="36%" class="style5"><div align="center">送達地名(或地址)</div></td>
	</tr>
	<%=strList%>
	</table>
</td>
</tr>
<tr>
<td>
	<table align="center" width="100%" border="0">
	<tr>
	<td width="66%" valign="top">
	  <p><span class="style11">(1) 限時掛號、掛號函件與快捷郵件不得同列一單，請將標題塗去其二。<br>
	    (2) 函件背面應註明順序號碼，並按號碼次序排齊滿二十件為一組分組交寄。<br>
	    (3) 將本埠與外埠函件分別列單交寄。
	    <br>
	    (4)如有證明郵資、重量必要者，應由寄件人自行在聯單相關欄內分別註明，並結填總郵資，交郵局</span><span class="style11">經辦員逐件核對。<br>
	    (5) 日後如須查詢，應於交寄日起六個月內檢同原件封面式樣向原寄局為之，並將本執據送驗。<br>
	    (6) 錢鈔或有價證券請利用報值或保價交寄。</span><br>
	    
	      </p>
	  </td>
	<td width="34%" class="style5" valign="Top">
	  <p>限時掛號<br>
	    掛號函件/共 
	    <%=mailSN%> 
	    件照收無誤<br>
	    快捷郵件<br>
	    
	    <%if sys_City<>"台東縣" then%>
		<br>
		<%end if%>
	    郵資共計  
	    <%
		if theMailMoney<>"" then
			response.write theMailMoney*mailSN
		else
			response.write "&nbsp;"
		end if
		%> 
	    元	  </p>
	  <p align="right">______________<br>經辦員簽署&nbsp; </p>
	  </td>
	</tr>
	</table>
</td>
</tr>
</table>
<%end if%>
<%if ReportCount=3 then %>

<div class="PageNext">&nbsp;</div>



<table width="710" align="center">
<tr>
<td>
	<table width="100%" align="center" cellpadding="3" border="0">
<%if sys_City<>"嘉義縣" and sys_City<>"台東縣" then%>
	<tr>
		<td height="25"></td>
	</tr>
<%end if%>
	<tr>
		<td width="34%"><span class="style7">
		頁&nbsp;&nbsp;次 &nbsp;<%=pageNum%> of <%=pagecnt%>
		</span></td>
		<td rowspan="3" width="39%" align="center"><span class="style7">
		<table width="100%">
		<tr>
			<td colspan="3" height="28"><div align="center"><u><span class="style6">中 華 郵 政</span></u><%
		If sys_City="新北市" Then
			%><div id="<%
			response.write "bcTarget"&pageNum&"3"
			%>" style= "position:absolute;width:400px;height:155px;z-index:1"></div><%
		End if
			%></div></td> 
		</tr>
		<tr>
			<td width="37%" rowspan="3" align="right" class="style7">交寄大宗</td>
			<td width="26%" class="style7"><u>限時掛號</u></td>
			<td width="37%" rowspan="3" align="left" class="style7">函件存根</td>
		</tr>
		<tr>
			<td class="style7"><u>掛 &nbsp; &nbsp;號</u></td>
		</tr>
		<tr>
			<td class="style7"><u>快捷郵件</u></td>
		</tr>
		</table>
		
		</span></td>
		<td rowspan="3" width="27%"><div align="right"><img src="../Image/MailPic.JPG" width="100" height="70" /></div></td>
	</tr>
	<tr>
		<td height="40" valign="top"><span class="style7">
<%if sys_City="台東縣" or sys_City="台南市" or sys_City="澎湖縣" then%>
		中華民國 <%
		response.write year(now)-1911
		%>年 <%
		response.write right("00"&month(now),2)
		%>月 <%
		response.write right("00"&day(now),2)
		%>日
<%elseif sys_City<>"雲林縣" and sys_City<>"花蓮縣" then %>
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
<%end if%>
		<br>
		<%
	if sys_City="南投縣"  or sys_City="基隆市" or sys_City="台東縣"  then
			response.write "作業批號："&MailBatchNumber
	end if
		%>
		</span></td>
	</tr>
	<tr>
		<td><span class="style7">
		寄件人 <%=UnitName%>
		</span></td>
	</tr>
	<tr>
		<td><span class="style7">
		寄件人代表 ___________
		</span></td>
		<td><span class="style7">
		詳細地址：<u><%=UnitAddress%></u>
		</span></td>
		<td><span class="style7">
		電話號碼：<u><%=UnitTel%></u>
		</span></td>
	</tr>
	</table>
</td>
</tr>
<tr>
<td>
	<table align="center" width="100%" border="1" cellspacing="0" cellpadding="3">
	<tr>
	<td width="6%" rowspan="2"><div align="center"><span class="style5">順序<br>
	  號碼</span></div></td>
	<td width="10%" rowspan="2"><div align="center"><span class="style5">掛號號碼</span></div></td>
	<td colspan="2"><div align="center"><span class="style5">收件人</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  回執<br>[V]</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  航空<br>[V]</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  印刷<br>[V]</span></div></td>
	<td width="3%" rowspan="2"><div align="center"><span class="style5">重量</span></div></td>
	<td width="6%" rowspan="2"><div align="center"><span class="style5">郵資</span></div></td>
	<td width="9%" rowspan="2"><div align="center"><span class="style5">備考</span></div></td>
	</tr>
	<tr>
	<td width="15%" class="style5"><div align="center">姓名</div></td>
	<td width="36%" class="style5"><div align="center">送達地名(或地址)</div></td>
	</tr>
	<%=strList%>
	</table>
</td>
</tr>
<tr>
<td>
	<table align="center" width="100%" border="0">
	<tr>
	<td width="66%" valign="top">
	  <p><span class="style11">(1) 限時掛號、掛號函件與快捷郵件不得同列一單，請將標題塗去其二。<br>
	    (2) 函件背面應註明順序號碼，並按號碼次序排齊滿二十件為一組分組交寄。<br>
	    (3) 將本埠與外埠函件分別列單交寄。
	    <br>
	    (4)如有證明郵資、重量必要者，應由寄件人自行在聯單相關欄內分別註明，並結填總郵資，交郵局</span><span class="style11">經辦員逐件核對。<br>
	    (5) 日後如須查詢，應於交寄日起六個月內檢同原件封面式樣向原寄局為之，並將本執據送驗。<br>
	    (6) 錢鈔或有價證券請利用報值或保價交寄。</span><br>
	    
	      </p>
	  </td>
	<td width="34%" class="style5" valign="Top">
	  <p>限時掛號<br>
	    掛號函件/共 
	    <%=mailSN%> 
	    件照收無誤<br>
	    快捷郵件<br>
	    
	    <br>
	    郵資共計  
	    <%
		if theMailMoney<>"" then
			response.write theMailMoney*mailSN
		else
			response.write "&nbsp;"
		end if
		%> 
	    元	  </p>
	  <p align="right">______________<br>經辦員簽署&nbsp; </p>
	  </td>
	</tr>
	</table>
</td>
</tr>
</table>
<%end if%>
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
