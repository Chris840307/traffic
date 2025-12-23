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
.style1 {font-family: "@新細明體"; font-size: 16px; }
.style2 {font-family: "@新細明體"; font-size: 14px; }
.style3 {font-family: "@新細明體"; font-size: 12px; }
.style4 {font-family: "@新細明體"; font-size: 18px; }
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
	if Instr(request("Sys_BatchNumber"),"N")>0 then
		tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T','V','n') and a.DciReturnStatusID<>'n' "&request("sys_strSQL")&") or ((a.BillTypeID='1' or f.UseTool=8) and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and f.RecordStateId <> -1 "&request("sys_strSQL")&")"
	else
		tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T','V','n') and f.UseTool<>8 "&request("sys_strSQL")&") or ((a.BillTypeID='1' or f.UseTool=8) and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and f.RecordStateId <> -1 "&request("sys_strSQL")&")"
	end if

	If sys_City="雲林縣" Then
		tempSQL=tempSQL&"or (a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and a.ExchangeTypeID='N' "&request("sys_strSQL")&")"
	End if

	strSQL="select a.BillSN,a.RecordMemberID,f.RecordDate from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h "&tempSQL&" order by f.RecordDate"

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
	strSQL="Select BillSN from BillMailHistory where BillSN in("&BillSN&") order by UserMarkDate"
	set rshis=conn.execute(strSQL)
	BillSN=""
	while Not rshis.eof
		if trim(BillSN)<>"" then BillSN=trim(BillSN)&","
		BillSN=BillSN&rshis("BillSN")
		rshis.movenext
	wend
	rshis.close
	PBillSN=Split(trim(BillSN),",")
else
	PBillSN=Split(BillSN,",")
end if
thenPasserCity=""
strUInfo="select * from Apconfigure where ID=30"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	for j=1 to len(trim(rsUInfo("value")))
		if j<>1 then thenPasserCity=thenPasserCity&"　"
		thenPasserCity=thenPasserCity&Mid(trim(rsUInfo("value")),j,1)
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
for i=0 to Ubound(PBillSN)
	if cint(i)<>0 then response.write "<div class=""PageNext""></div>"%>
	<!--#include virtual="traffic/Query/BillBaseCHCG_Deliver.asp"--><%
Next
%>

</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(true,5.08,5.08,5.08,5.08);
</script>