<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>批次輸出系統</title>
<style type="text/css">
<!--
.style1 {font-family: "@新細明體"; font-size: 10px;}
.style2 {font-family: "@新細明體"; font-size: 16px;}
.style3 {font-family: "@新細明體"; font-size: 14px;}
.style4 {font-family: "@新細明體"; font-size: 14px;}
-->
</style>
<!--#include virtual="traffic/Common/css.txt"-->
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>
<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%
Server.ScriptTimeout=6000
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
SendTime=1
If Not ifnull(request("Sys_BatchNumber")) Then
	If instr(Ucase(request("Sys_BatchNumber")),"W")>0 Then SendTime=2
End if
'if trim(request("Sys_CityKind"))="0" then
	tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T') and f.UseTool<>8 "&request("sys_strSQL")&") or ((a.BillTypeID='1' or f.UseTool=8) and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and f.RecordStateId <> -1 "&request("sys_strSQL")&")"

	If sys_City="雲林縣" Then
		tempSQL=tempSQL&"or (a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and a.ExchangeTypeID='N' "&request("sys_strSQL")&")"
	End if

'if trim(request("PBillSN"))="" then '與dci上下查詢不同
	strSQL="select a.BillSN,a.RecordMemberID,f.RecordDate from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h "&tempSQL&" order by f.RecordDate"

	tmpSQL="select distinct a.BillSN from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h "&tempSQL
'elseif trim(request("Sys_CityKind"))="1" then
'	tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and a.BillNo=i.Billno(+) and a.CarNo=i.CarNo(+) and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T','V') "&request("sys_strSQL")&")"
'
'	If sys_City="雲林縣" Then
'		tempSQL=tempSQL&"or (a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and a.ExchangeTypeID='N' "&request("sys_strSQL")&")"
'	End if
'
''if trim(request("PBillSN"))="" then '與dci上下查詢不同
'	strSQL="select a.BillSN,a.RecordMemberID,f.RecordDate,DeCode(a.BillTypeID,'2',i.OwnerZip,'1',i.DriverHomezip) OwnerZip from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h,(select BillNo,CarNo,OwnerZip,DriverHomezip from BillBaseDCIReturn where ExchangeTypeID='W') i "&tempSQL&" order by OwnerZip"
'end if
set rssn=conn.execute(strSQL)
BillSN="":tmpBillSN=""
while Not rssn.eof
	If trim(tmpBillSN)<>trim(rssn("BillSN")) Then
		if trim(BillSN)<>"" then BillSN=trim(BillSN)&","
		BillSN=BillSN&trim(rssn("BillSN"))
		tmpBillSN=trim(rssn("BillSN"))
	end if
	rssn.movenext
wend
rssn.close
if (OptionStoreAndSendMailChk=2 or Instr(request("Sys_BatchNumber"),"N")>0) and trim(BillSN)<>"" then
	strSQL="Select BillSN from BillMailHistory where BillSN in("&tmpSQL&") order by UserMarkDate"
	set rshis=conn.execute(strSQL)
	BillSN=""
	while Not rshis.eof
		if trim(BillSN)<>"" then BillSN=trim(BillSN)&","
		BillSN=BillSN&rshis("BillSN")
		rshis.movenext
	wend
	rshis.close
	strBillSN=Split(trim(BillSN),",")
else
	strBillSN=Split(BillSN,",")
end if
thenPasserCity=""
strUInfo="select * from Apconfigure where ID=30"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	'for j=1 to len(trim(rsUInfo("value")))
		'if j<>1 then thenPasserCity=thenPasserCity&"　"
		thenPasserCity=trim(rsUInfo("value"))
	'next
end if
rsUInfo.close
strUInfo="select * from Apconfigure where ID=52"
set rsUInfo=conn.execute(strUInfo)
theBillNumber=""
if not rsUInfo.eof then
	theBillNumber=rsUinfo("Value")
end if
rsUInfo.close
set rsUInfo=nothing

Function wordporss(word)
	pro_word="":img_word=""
	min=0:max=0
	If instr(word,"<") >0 Then
		min=instr(word,"<")
		max=instr(word,">")-min+1
		img_word=mid(word,min,max)
		word=replace(word,img_word,"")
	End if
	
	for y=len(word) to 1 step -1
		If (min>0 and len(word)<min and y=len(word)) Then pro_word=img_word&"<br>"

		pro_word=pro_word&Mid(word,y,1)
		if y<>1 then
			pro_word=pro_word&"<br>"
			if (min>0 and y=min) or (min>0 and y=1) then
				If not (min>0 and len(word)<min and y=len(word)) Then pro_word=pro_word&img_word&"<br>"
			end if
		end if
	next
	wordporss=pro_word
end function

Function chstr(istr) ' 半形轉全形
Dim strtmp
min=0:max=0
If instr(istr,"<") >0 Then
	min=instr(istr,"<")
	max=instr(istr,">")-min+1
	img_word=mid(istr,min,max)
	istr=replace(istr,img_word,"")
End if
if trim(istr)<>"" then
	strtmp = Replace(istr, "(", "（") 
	strtmp = Replace(strtmp, ")", "）") 
	strtmp = Replace(strtmp, "[", "〔") 
	strtmp = Replace(strtmp, "]", "〕") 
	strtmp = Replace(strtmp, "{", "｛") 
	strtmp = Replace(strtmp, "}", "｝") 
	strtmp = Replace(strtmp, ".", "。") 
	strtmp = Replace(strtmp, ",", "，") 
	strtmp = Replace(strtmp, ";", "；") 
	strtmp = Replace(strtmp, ":", "：") 
	strtmp = Replace(strtmp, "-", "－") 
	strtmp = Replace(strtmp, "?", "？") 
	strtmp = Replace(strtmp, "!", "！") 
	strtmp = Replace(strtmp, "@", "＠") 
	strtmp = Replace(strtmp, "#", "＃") 
	strtmp = Replace(strtmp, "$", "＄") 
	strtmp = Replace(strtmp, "%", "％") 
	strtmp = Replace(strtmp, "&", "＆") 
	strtmp = Replace(strtmp, "|", "｜") 
	strtmp = Replace(strtmp, "", "＼") 
	strtmp = Replace(strtmp, "/", "／") 
	strtmp = Replace(strtmp, "+", "＋") 
	strtmp = Replace(strtmp, "=", "＝") 
	strtmp = Replace(strtmp, "*", "＊") 
	strtmp = Replace(strtmp, "0", "０") 
	strtmp = Replace(strtmp, "1", "１") 
	strtmp = Replace(strtmp, "2", "２") 
	strtmp = Replace(strtmp, "3", "３") 
	strtmp = Replace(strtmp, "4", "４") 
	strtmp = Replace(strtmp, "5", "５") 
	strtmp = Replace(strtmp, "6", "６") 
	strtmp = Replace(strtmp, "7", "７") 
	strtmp = Replace(strtmp, "8", "８") 
	strtmp = Replace(strtmp, "9", "９") 
	strtmp = Replace(strtmp, "a", "ａ") 
	strtmp = Replace(strtmp, "b", "ｂ") 
	strtmp = Replace(strtmp, "c", "ｃ") 
	strtmp = Replace(strtmp, "d", "ｄ") 
	strtmp = Replace(strtmp, "e", "ｅ") 
	strtmp = Replace(strtmp, "f", "ｆ") 
	strtmp = Replace(strtmp, "g", "ｇ") 
	strtmp = Replace(strtmp, "h", "ｈ") 
	strtmp = Replace(strtmp, "i", "ｉ") 
	strtmp = Replace(strtmp, "j", "ｊ") 
	strtmp = Replace(strtmp, "k", "ｋ") 
	strtmp = Replace(strtmp, "l", "ｌ") 
	strtmp = Replace(strtmp, "m", "ｍ") 
	strtmp = Replace(strtmp, "n", "ｎ") 
	strtmp = Replace(strtmp, "o", "ｏ") 
	strtmp = Replace(strtmp, "p", "ｐ") 
	strtmp = Replace(strtmp, "q", "ｑ") 
	strtmp = Replace(strtmp, "r", "ｒ") 
	strtmp = Replace(strtmp, "s", "ｓ") 
	strtmp = Replace(strtmp, "t", "ｔ") 
	strtmp = Replace(strtmp, "u", "ｕ") 
	strtmp = Replace(strtmp, "v", "ｖ") 
	strtmp = Replace(strtmp, "w", "ｗ") 
	strtmp = Replace(strtmp, "x", "ｘ") 
	strtmp = Replace(strtmp, "y", "ｙ") 
	strtmp = Replace(strtmp, "z", "ｚ") 
	strtmp = Replace(strtmp, "A", "Ａ") 
	strtmp = Replace(strtmp, "B", "Ｂ") 
	strtmp = Replace(strtmp, "C", "Ｃ") 
	strtmp = Replace(strtmp, "D", "Ｄ") 
	strtmp = Replace(strtmp, "E", "Ｅ") 
	strtmp = Replace(strtmp, "F", "Ｆ") 
	strtmp = Replace(strtmp, "G", "Ｇ") 
	strtmp = Replace(strtmp, "H", "Ｈ") 
	strtmp = Replace(strtmp, "I", "Ｉ") 
	strtmp = Replace(strtmp, "J", "Ｊ") 
	strtmp = Replace(strtmp, "K", "Ｋ") 
	strtmp = Replace(strtmp, "L", "Ｌ") 
	strtmp = Replace(strtmp, "M", "Ｍ") 
	strtmp = Replace(strtmp, "N", "Ｎ") 
	strtmp = Replace(strtmp, "O", "Ｏ") 
	strtmp = Replace(strtmp, "P", "Ｐ") 
	strtmp = Replace(strtmp, "Q", "Ｑ") 
	strtmp = Replace(strtmp, "R", "Ｒ") 
	strtmp = Replace(strtmp, "S", "Ｓ") 
	strtmp = Replace(strtmp, "T", "Ｔ") 
	strtmp = Replace(strtmp, "U", "Ｕ") 
	strtmp = Replace(strtmp, "V", "Ｖ") 
	strtmp = Replace(strtmp, "W", "Ｗ") 
	strtmp = Replace(strtmp, "X", "Ｘ") 
	strtmp = Replace(strtmp, "Y", "Ｙ") 
	strtmp = Replace(strtmp, "Z", "Ｚ") 
	strtmp = Replace(strtmp, " ", "　")
end if

If min >0 Then strtmp=mid(strtmp,1,min-1)&img_word&mid(strtmp,min)

chstr = strtmp 
End Function  

for i=0 to Ubound(strBillSN)
	if cint(i)<>0 then response.write "<div class=""PageNext""></div>"%>	
	<!--#include virtual="traffic/Query/BillBaseChiayiCity_Deliver.asp"--><%
Next
%>
<div id=idDiv></div>
<div id=oCodeDiv></div>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	for(i=0;i<document.all("idDiv").length-1;i++){
		document.all("idDiv")[i].style.filter="progid:DXImageTransform.Microsoft.BasicImage(rotation=3)";
	}
	window.focus();
	printWindow(true,0,0,0,0);
</script>