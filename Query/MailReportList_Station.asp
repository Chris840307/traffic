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
strExchangeType="select a.ExchangeTypeID,f.BillUnitID from DciLog a,BillBase f where a.BillSN=f.SN "&_
	" and f.RecordStateID=0 "&strwhere
set rsEType=conn.execute(strExchangeType)
if not rsEType.eof then
	if trim(rsEType("ExchangeTypeID"))="N" then
		ExchangeTypeFlag="N"
	else
		ExchangeTypeFlag="W"
	end if
	BillUnitIDtmp=trim(rsEType("BillUnitID"))
else
	ExchangeTypeFlag="W"
	BillUnitIDtmp=""
end if
rsEType.close
set rsEType=nothing

if sys_City="台中市" then 
	if BillUnitIDtmp="" then
		strSendMailUnit="select b.UnitName,b.Address,b.Tel from Apconfigure a,UnitInfo b " &_
				" where a.ID=49 and a.Value=b.UnitID"
		set rsSendMailUnit=conn.execute(strSendMailUnit)
		if not rsSendMailUnit.eof then
			
			if sys_City<>"花蓮縣" and sys_City<>"台中市" then 
				UnitName=TitleUnitName&trim(rsSendMailUnit("UnitName"))
			else
				UnitName=trim(rsSendMailUnit("UnitName"))
			end if
			UnitAddress=trim(rsSendMailUnit("Address"))
			UnitTel=trim(rsSendMailUnit("Tel"))
		end if
		rsSendMailUnit.close
		set rsSendMailUnit=nothing
	else
		'檢查舉發單位showorder
		strShow="select * from UnitInfo where UnitID='"&BillUnitIDtmp&"'"
		set rsShow=conn.execute(strShow)
		if not rsShow.eof then
			'showorder=0 or 1,寄件人就是舉發單位
			if trim(rsShow("ShowOrder"))="0" or trim(rsShow("ShowOrder"))="1" or trim(rsShow("UnitID"))="046A" or trim(rsShow("UnitID"))="0469" then
				UnitName=trim(rsShow("UnitName"))
				UnitAddress=trim(rsShow("Address"))
				UnitTel=trim(rsShow("Tel"))
			'showorder=2,寄件人是上層單位
			elseif trim(rsShow("ShowOrder"))="2" then
				strUnitType="select * from UnitInfo where UnitID='"&trim(rsShow("UnitTypeID"))&"'"
				set rsUnitType=conn.execute(strUnitType)
				if not rsUnitType.eof then
					UnitName=trim(rsUnitType("UnitName"))
					UnitAddress=trim(rsUnitType("Address"))
					UnitTel=trim(rsUnitType("Tel"))
				end if
				rsUnitType.close
				set rsUnitType=nothing
			end if
		else
			UnitName=""
			UnitAddress=""
			UnitTel=""
		end if
		rsShow.close
		set rsShow=nothing
	end if
else
	strSendMailUnit="select b.UnitName,b.Address,b.Tel from MemberData a,UnitInfo b " &_
			" where a.MemberID="&trim(Session("User_ID"))&" and a.UnitID=b.UnitID"
	set rsSendMailUnit=conn.execute(strSendMailUnit)
	if not rsSendMailUnit.eof then
		
		if sys_City="花蓮縣" then 
			UnitName=trim(rsSendMailUnit("UnitName"))
		elseif sys_City="屏東縣" then 
			UnitName=TitleUnitName&replace(rsSendMailUnit("UnitName"),"屏東縣政府警察局","")
		else
			UnitName=TitleUnitName&trim(rsSendMailUnit("UnitName"))
		end if
		UnitAddress=trim(rsSendMailUnit("Address"))
		UnitTel=trim(rsSendMailUnit("Tel"))
	end if
	rsSendMailUnit.close
	set rsSendMailUnit=nothing
end if



if sys_City="南投縣" then
	if ExchangeTypeFlag="N" then
		strSQL="select distinct e.DciReturnStation" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and a.ExchangeTypeID='N' and e.ExchangeTypeID='W'" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by e.DciReturnStation"
	else
		strSQL="select distinct Decode(f.BillTypeID,1,f.MemberStation,e.DciReturnStation) DciReturnStation" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','5','9','a','j','A','H','K','L','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 )) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L')) and a.ExchangeTypeID='W'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by DciReturnStation"
	end if
elseif sys_City="高雄縣" Or sys_City="保二總隊四大隊二中隊" then
	if ExchangeTypeFlag="N" then
		strSQL="select distinct e.DciReturnStation" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		
		" and (e.ExchangeTypeID='N' and d.DCIreturnStatus=1)" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by e.DciReturnStation"
	else
		strSQL="select distinct e.DciReturnStation" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','5','9','a','j','A','H','K','L','T','V')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8)) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L')) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by e.DciReturnStation"
	end if
end if
set rs1=conn.execute(strSQL)
if sys_City="南投縣" then 
	if ExchangeTypeFlag="N" then
		strCnt="select count(distinct(e.DciReturnStation)) as cnt" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=g.BillNo and a.CarNo=g.CarNO and a.BillSn=g.BillSN" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and a.ExchangeTypeID='N' and e.ExchangeTypeID='W'" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"

	else
		strCnt="select count(distinct(Decode(f.BillTypeID,1,f.MemberStation,e.DciReturnStation))) as cnt" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','5','9','a','j','A','H','K','L','T')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8)) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L')) and a.ExchangeTypeID='W'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	end if
elseif sys_City="高雄縣" Or sys_City="保二總隊四大隊二中隊" then 
	if ExchangeTypeFlag="N" then
		strCnt="select count(distinct(e.DciReturnStation)) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO" &_
		" and (e.ExchangeTypeID='N' and d.DCIreturnStatus=1)" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	else
		strCnt="select count(distinct(e.DciReturnStation)) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','5','9','a','j','A','H','K','L','T','V')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8)) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L')) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	end if
end if
set rsCnt=conn.execute(strCnt)
if not rsCnt.eof then
	if trim(rsCnt("cnt"))="0" then
		pagecnt=1
	else
		pagecnt=fix(Cint(rsCnt("cnt"))/20+0.9999999)
	end if
end if
rsCnt.close
set rsCnt=nothing
'response.write strSQL
MDate=""
if ExchangeTypeFlag="N" then
	strMailDate="select g.StoreAndSendMailDate as MDate from DciLog a,BillBase f,BillMailHistory g " &_
		" where f.Sn=g.BillSn and f.Sn=a.BillSn "&strwhere
else
	strMailDate="select g.MailDate as MDate from DciLog a,BillBase f,BillMailHistory g " &_
		" where f.Sn=g.BillSn and f.Sn=a.BillSn "&strwhere
end if
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
StataionIDTemp=""
If Not rs1.Bof Then rs1.MoveFirst 
While Not rs1.Eof
	If StataionIDTemp="" Then
		StataionIDTemp=Trim(rs1("DciReturnStation"))
	Else
		StataionIDTemp=StataionIDTemp& "#$#" & Trim(rs1("DciReturnStation"))
	End If 
	rs1.MoveNext
Wend
rs1.close
set rs1=Nothing
TaipeiFlag=0
KSFlag=0
NewTaipeiFlag=0
TCFlag=0
StationIDPrint=""
Array_StataionIDTemp=Split(StataionIDTemp,"#$#")
For i=0 To UBound(Array_StataionIDTemp)
	If instr("20,21,22,23,24,29",Array_StataionIDTemp(i))>0 Then
		TaipeiFlag=1
	ElseIf instr("30,31,32",Array_StataionIDTemp(i))>0 Then
		KSFlag=1
	ElseIf instr("40,41,46,48",Array_StataionIDTemp(i))>0 Then
		NewTaipeiFlag=1
	ElseIf instr("60,61,63,68",Array_StataionIDTemp(i))>0 Then
		TCFlag=1
	Else
		If StationIDPrint="" Then
			StationIDPrint=Trim(Array_StataionIDTemp(i))
		Else
			StationIDPrint=StationIDPrint& "#$#" & Trim(Array_StataionIDTemp(i))
		End If 
	End If 
Next 
If TCFlag=1 Then
	If StationIDPrint="" Then
		StationIDPrint="60"
	Else
		StationIDPrint="60#$#"&StationIDPrint
	End If 
End If 
If KSFlag=1 Then
	If StationIDPrint="" Then
		StationIDPrint="30"
	Else
		StationIDPrint="30#$#"&StationIDPrint
	End If 
End If 
If NewTaipeiFlag=1 Then
	If StationIDPrint="" Then
		StationIDPrint="40"
	Else
		StationIDPrint="40#$#"&StationIDPrint
	End If 
End If 
If TaipeiFlag=1 Then
	If StationIDPrint="" Then
		StationIDPrint="20"
	Else
		StationIDPrint="20#$#"&StationIDPrint
	End If 
End If 
array_StationIDPrint=Split(StationIDPrint,"#$#")
if trim(StationIDPrint)="" then
	pagecnt=1
else
	pagecnt=fix((UBound(array_StationIDPrint)+1)/20+0.9999999)
end If
AllSCnt=0
While SIDP<UBound(array_StationIDPrint)+1
if mailSN>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
	BillFillDateTmp=""
	
	strList=""
	mailSN=0
	pageNum=fix(CaseSN/20)+1
	for i=1 to 20
		if SIDP>UBound(array_StationIDPrint) then exit For
		
		ZipName=""

		mailSN=mailSN+1
		CaseSN=CaseSN+1
		if sys_City="花蓮縣"  then
			strList=strList&"<tr height=""23"">"
		else
			strList=strList&"<tr>"		
		end if
		'順序號碼

		strList=strList&"<td align=""center"">"&CaseSN&"</td>"
		'掛號號碼
		theMailNumber=""
		'移送監理站日期
		theSendDocDate=""
		
		strList=strList&"<td align=""center"">&nbsp;</td>"

		GetMailMem=""
		GetMailAddress=""
		
		'收件人姓名
		strList=strList&"<td align=""left"" class=""style33"" colspan=""2"">"
		if not isnull(array_StationIDPrint(SIDP)) and trim(array_StationIDPrint(SIDP))<>"" then
			strSqlStationName="select DCIstationName from Station where DCIstationID='"&trim(array_StationIDPrint(SIDP))&"'"
			set rsSN=conn.execute(strSqlStationName)
			if not rsSN.eof then
				strList=strList&trim(rsSN("DCIstationName"))
			end if
			rsSN.close
			set rsSN=nothing
		else
			strList=strList&"&nbsp;"
		end if
		strList=strList&"</td>"
			
		'收件地址
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
		strList=strList&"<td align=""left"">&nbsp;</td>"
		strList=strList&"</tr>"
		SIDP=SIDP+1
	next
	if mailSN<20 then
		if sys_City<>"雲林縣" and sys_City<>"台南縣" and sys_City<>"台南市" then
			mailSNTmp=mailSN
		else
			mailSNTmp=CaseSN
		end if
		for Sp=1 to 20-mailSN
			mailSNTmp=mailSNTmp+1
			if sys_City="花蓮縣"  then
				strList=strList&"<tr height=""23"">"
			else
				strList=strList&"<tr>"
			end if
			'順序號碼
			if sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕" then 
				strList=strList&"<td align=""center"">&nbsp;</td>"
			else
				strList=strList&"<td align=""center"">"&mailSNTmp&"</td>"
			end if
			strList=strList&"<td align=""center"">&nbsp;</td>"

			strList=strList&"<td align=""center"" colspan=""2"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"</tr>"
		next
	end if

if (sys_City="南投縣" And Trim(session("Unit_ID"))<>"05A7") or sys_City="雲林縣" or sys_City="台南市" or sys_City="宜蘭縣" then 
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
<table width="710" align="center"  border="0">
<tr>
<td>
	<table width="100%" align="center" cellpadding="3" border="0">
<%if sys_City<>"花蓮縣" and sys_City<>"嘉義縣" then %>
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
			<td colspan="3" height="30"><div align="center"><u><span class="style6">中 華 郵 政</span></u></div></td> 
		</tr>

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
<%end if%>
		<br>
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
		</span></td>
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
		<br>
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
<%if sys_City<>"嘉義縣" then%>
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
			<td colspan="3" height="30"><div align="center"><u><span class="style6">中 華 郵 政</span></u></div></td> 
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
<%if sys_City="澎湖縣" then %>	
		<span class="style8">□□□□□□ □□</span>
		<br>
		 &nbsp; &nbsp; &nbsp;收寄局碼&nbsp; &nbsp;郵件種類碼
		 <br>
		 &nbsp; &nbsp; &nbsp; &nbsp;(由收寄局填寫)
<%end if%>
		<br>
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
<%if ReportCount=3 then %>

<div class="PageNext">&nbsp;</div>



<table width="710" align="center">
<tr>
<td>
	<table width="100%" align="center" cellpadding="3" border="0">
<%if sys_City<>"嘉義縣" then%>
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
			<td colspan="3" height="30"><div align="center"><u><span class="style6">中 華 郵 政</span></u></div></td> 
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
	
wend
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
