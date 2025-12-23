<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>批次輸出系統</title>
<style type="text/css">
<1--
.style3 {font-family: "標楷體"; font-size: 7px;}
.style4 {font-family: "標楷體"; font-size: 9px;}
.style5 {font-family: "標楷體"; font-size: 16px; line-height:16px;}
.style6 {font-family: "標楷體"; font-size: 12px; line-height:14px;}
.style9 {font-family: "標楷體"; font-size: 14px;line-height:14px;}
.style10 {font-family: "標楷體"; font-size: 10px;line-height:10px;}
.style7 {font-family: "標楷體"; font-size: 8px;}
.tablestyle{position:relative; left:30px; top:0px;}
.tablestyle2{position:relative; top:-350px;}
-->
</style>
<!--#include virtual="traffic/Common/css.txt"-->
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>
<body>
<%
Server.ScriptTimeout=60000
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

'if trim(request("Sys_CityKind"))="0" then
	If sys_City="台東縣" Then
		tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T','V') and f.UseTool<>8 "&request("sys_strSQL")&") or ((a.BillTypeID='1' or f.UseTool=8) and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&request("sys_strSQL")&")"
	elseIf sys_City="基隆市" Then
		tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T','V') and f.UseTool<>8 "&request("sys_strSQL")&") or ((a.BillTypeID='1' or f.UseTool=8) and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&request("sys_strSQL")&")"
	elseIf sys_City="南投縣" Then
		tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T','V') and a.DciReturnStatusID<>'n' "&request("sys_strSQL")&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.DciReturnStatusID<>'n' and a.ExchangeTypeID<>'E' and f.RecordStateId <> -1 "&request("sys_strSQL")&")"

	elseIf sys_City="彰化縣" Then
		tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T','V','n') and a.DciReturnStatusID<>'n' "&request("sys_strSQL")&") or ((a.BillTypeID='1' or f.UseTool=8) and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and f.RecordStateId <> -1 "&request("sys_strSQL")&")"
	elseIf sys_City="台中市" Then
		tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T') and a.DciReturnStatusID<>'n' and f.UseTool<>8 "&request("sys_strSQL")&") or ((a.BillTypeID='1' or f.UseTool=8) and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.DciReturnStatusID<>'n' and a.ExchangeTypeID<>'E' "&request("sys_strSQL")&")"

	else
		tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T') and f.UseTool<>8 "&request("sys_strSQL")&") or ((a.BillTypeID='1' or f.UseTool=8) and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&request("sys_strSQL")&")"
	End if
	
	If sys_City="雲林縣" Then
		tempSQL=tempSQL&"or (a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and a.ExchangeTypeID='N' "&request("sys_strSQL")&")"
	End if

'if trim(request("PBillSN"))="" then '與dci上下查詢不同
	strSQL="select a.BillSN,a.RecordMemberID,f.RecordDate from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h "&tempSQL&" and f.equipmentID<>'-1' order by a.RecordMemberID,f.RecordDate"
'elseif trim(request("Sys_CityKind"))="1" then
'	tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and a.BillNo=i.Billno(+) and a.CarNo=i.CarNo(+) and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T','V') "&request("sys_strSQL")&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and a.BillNo=i.Billno(+) and a.CarNo=i.CarNo(+) and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&request("sys_strSQL")&")"

'	If sys_City="雲林縣" Then
'		tempSQL=tempSQL&"or (a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and a.ExchangeTypeID='N' "&request("sys_strSQL")&")"
'	End if

'if trim(request("PBillSN"))="" then '與dci上下查詢不同
'	strSQL="select a.BillSN,a.RecordMemberID,f.RecordDate,DeCode(a.BillTypeID,'2',i.OwnerZip,'1',i.DriverHomezip) OwnerZip from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h,(select BillNo,CarNo,OwnerZip,DriverHomezip from BillBaseDCIReturn where ExchangeTypeID='W') i "&tempSQL&" order by OwnerZip"
'end if
set rssn=conn.execute(strSQL)
BillSN="":tmpBillSN=""
while Not rssn.eof
	If trim(tmpBillSN)<>trim(rssn("BillSN")) Then
		if trim(BillSN)<>"" then BillSN=trim(BillSN)&","
		BillSN=BillSN&trim(rssn("BillSN"))
	end if
	rssn.movenext
wend
rssn.close
'===by kevin 單退或入案
If Instr(request("Sys_BatchNumber"),"N")>0 Then
	chkStore=1
Else
	chkStore=0
End If 
'==========
if (OptionStoreAndSendMailChk=2 or Instr(request("Sys_BatchNumber"),"N")>0) and trim(BillSN)<>"" then
	strSQL="Select BillSN from BillMailHistory where BillSN in("&BillSN&") order by UserMarkDate"
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
thenPasserCity="":thenUnitName=""
strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsunit=conn.execute(strSQL)
if Not rsunit.eof then
	for j=1 to len(trim(rsunit("UnitName")))
		'if j<>1 then thenUnitName=thenUnitName&"　"
		thenUnitName=thenUnitName&Mid(trim(rsunit("UnitName")),j,1)
	next
end if
rsunit.close
strUInfo="select * from Apconfigure where ID=35"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	for j=1 to len(trim(rsUInfo("value")))
		'if j<>1 then thenPasserCity=thenPasserCity&"　"
		thenPasserCity=thenPasserCity&Mid(trim(rsUInfo("value")&thenUnitName),j,1)
	next
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

for gyi=0 to Ubound(strBillSN)
	if trim(strBillSN(gyi))<>"" then
		%>
		<div id="L78" style="position:relative;">
		<div id="Layer1" style="position:absolute; left:0px; top:0px; z-index:5">
		<table width="95%" border="0" cellspacing="0">
			<tr>
				<td width="49%">
					<!--#include virtual="traffic/Query/BillBaseHuaLien_DeliverV_A4.asp"-->
				</td>
				<td width="20"></td>
				<td width="49%">
					<%
						If gyi+1 <= Ubound(strBillSN) Then
							gyi=gyi+1
							%><!--#include virtual="traffic/Query/BillBaseHuaLien_DeliverV_A4.asp"--><%							
						End if 
					%>					
				</td>
			</tr>
		</table>
		</div>
		</div>
		<%
		if cint(gyi+1) mod 2 =0 and (gyi+1) <= Ubound(strBillSN) then
			response.write "<div class=""PageNext"">&nbsp;</div>"
		end if
		if (gyi mod 50)=0 then response.flush
	end if
Next
%>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">

	window.focus();
	printWindow(false,5.08,5.08,5.08,5.08);
</script>