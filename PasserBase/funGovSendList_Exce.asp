<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
%>
<%if sys_City<>"雲林縣" and sys_City<>"台中縣" and sys_City<>"嘉義縣" then%>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://10.104.10.246/traffic/smsx.cab#Version=6,1,432,1">
</object>
<%end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style3 {font-family:新細明體; color=0044ff; line-height:19px; font-size: 15px}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>公示送達清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<%
	Server.ScriptTimeout = 18000
	Response.flush
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

If Not ifnull(request("Sys_SendBillSN")) Then

	sys_billsn=request("Sys_SendBillSN")
elseif Not ifnull(request("hd_BillSN")) Then

	sys_billsn=request("hd_BillSN")
else

	sys_billsn=request("BillSN")
End If 

tmp_billsn=split(sys_billsn,",")

sys_billsn=""

For i = 0 to Ubound(tmp_billsn)

	If i >0 then

		If i mod 100 = 0 Then

			sys_billsn=sys_billsn&"@"
		elseif sys_billsn<>"" then

			sys_billsn=sys_billsn&","
		end If 
	end if

	sys_billsn=sys_billsn&tmp_billsn(i)

Next

tmpSQL=""

If Ubound(tmp_billsn) >= 100 Then

	sys_billsn=split(sys_billsn,"@")
	
	For i = 0 to Ubound(sys_billsn)
		
		If tmpSQL <>"" Then tmpSQL=tmpSQL&" union all "
		
		tmpSQL=tmpSQL&"select sn from passerbase where sn in("&sys_billsn(i)&")"
	Next

else

	tmpSQL="select sn from passerbase where sn in("&sys_billsn&")"

End if 

BasSQL="("&tmpSQL&") tmpPasser"

	strSQL="select BillNo,Driver,DriverID,Rule1,Rule2,IllegalDate," &_
	"(select Max(ArrivedDate) ArrivedDate from PassersEndArrived where ArriveType=0 and ReturnResonID='1' and PasserSN=PasserBase.sn) ArrivedDate," &_
	"(Select OpenGovNumber from PasserJude where billsn=PasserBase.sn) OpenGovNumber," &_
	"(select Max(Note) Note from PassersEndArrived where ArriveType=0 and ReturnResonID='1' and PasserSN=PasserBase.sn) Note" &_
	" from PasserBase where RecordStateID=0 and Exists(select 'Y' from "&BasSQL&" where sn=PasserBase.sn) and Exists(select 'Y' from PassersEndArrived where ArriveType=0 and ReturnResonID='1' and PasserSN=PasserBase.sn) "
	set rsdb=conn.execute(strSQL)
%>
</head>
<body>
<form name=myForm method="post">
<%
	PageContrl=0
	While not rsdb.eof
		PageContrl=PageContrl+1
		If PageContrl Mod 20 = 1 Then
			If PageContrl > 1 Then response.write "<div class=""PageNext""></div>"
			Response.Write "<center><font size=""3"">舉發違反道路交通事件通知單公示送達移送清冊</font></center>"
			Response.Write "<br>列印日期："&now
			Response.Write "<table width=""100%"" border=""1"" cellpadding=""1"" cellspacing=""0"">"
			Response.Write "<tr>"
			Response.Write "<td align=""center"">編號</td>"
			Response.Write "<td align=""center"">單號</td>"
			Response.Write "<td align=""center"">違規人姓名</td>"
'			Response.Write "<td align=""center"">違規人性別</td>"
'			Response.Write "<td align=""center"">違規人證號</td>"
			Response.Write "<td align=""center"">法條一</td>"
			Response.Write "<td align=""center"">法條二</td>"
			Response.Write "<td align=""center"">違規日期</td>"
			Response.Write "<td align=""center"">送達日期</td>"
			Response.Write "<td align=""center"">裁決文號</td>"
			Response.Write "<td align=""center"">退件原因</td>"
			Response.Write "</tr>"
		end if

		Response.Write "<tr>"
		Response.Write "<td>"&(PageContrl Mod 20)&"</td>"
		Response.Write "<td>"&trim(rsdb("Billno"))&"</td>"
		Response.Write "<td>"&trim(rsdb("Driver"))&"</td>"

'		If not ifnull(rsdb("DriverID")) Then
'			If Mid(Trim(rsdb("DriverID")),2,1)="1" Then
'				Response.Write "<td>男</td>"
'			elseif Mid(Trim(rsdb("DriverID")),2,1)="2" Then
'				Response.Write "<td>女</td>"
'			End if
'		End if

'		Response.Write "<td>"&trim(rsdb("DriverID"))&"</td>"
		Response.Write "<td>"&trim(rsdb("rule1"))&"</td>"
		Response.Write "<td>"&trim(rsdb("rule2"))&"&nbsp;</td>"
		Response.Write "<td>"&trim(gInitDT(rsdb("IllegalDate")))&"&nbsp;</td>"
		response.write "<td>"&trim(gInitDT(rsdb("ArrivedDate")))&"&nbsp;</td>"
		Response.Write "<td>"&trim(rsdb("OpenGovNumber"))&"&nbsp;</td>"
		Response.Write "<td>"&trim(rsdb("Note"))&"&nbsp;</td>"
		Response.Write "</tr>"

		If PageContrl Mod 20 = 0 Then Response.Write "</table><span class=""style5"">第"&fix(PageContrl/20+0.9999)&"頁</span>"

		rsdb.movenext
	Wend
	rsdb.close
	If PageContrl Mod 20 > 0 Then Response.Write "</table><span class=""style5"">第"&fix(PageContrl/20+0.9999)&"頁</span>"
%>

</form>
</body>
</html>
<script language="javascript">
	printWindow(true,7,5.08,5.08,5.08);
</script>
<%conn.close%>