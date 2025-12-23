<!--#include virtual="traffic/Common/cssForForm.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>催繳郵寄未退回清冊</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<%
	Server.ScriptTimeout = 800

	strRul="select Value from Apconfigure where ID=3"
	set rsRul=conn.execute(strRul)
	RuleVer=trim(rsRul("Value"))
	rsRul.Close

	strwhere=""

	If (not ifnull(request("Sys_RecordDate1"))) and (not ifnull(request("Sys_RecordDate2"))) Then

		ArgueDate1=gOutDT(request("Sys_RecordDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("Sys_RecordDate2"))&" 23:59:59"

		strwhere="and RecordDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and RecordMemberID="&session("User_ID")

	end if

	if not ifnull(Request("Sys_BatChNumber")) then

		strwhere=strwhere&" and sn in(select Billsn from DciLog where Batchnumber='"&trim(Request("Sys_BatChNumber"))&"')"
	End if
	
	


	strCnt="select count(1) as cnt from ( select a.imagefilenameb billno,a.carno,b.maildate,b.mailnumber from (select distinct imagefilenameb,carno from billbase where BillStatus=2 and imagefilenameb is not null "&strwhere&") a,(select distinct billno,carno,MailDate,mailnumber from StopBillMailHistory where billno is not null and mailnumber is not null and UserMarkResonID is null) b where a.imagefilenameb=b.billno and a.carno=b.carno)"
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		if trim(rsCnt("cnt"))="0" then
			pagecnt=1
		else
			pagecnt=fix(Cint(rsCnt("cnt"))/25+0.9999999)
		end if
	end if
	rsCnt.close
	set rsCnt=nothing

	strSQL="select a.imagefilenameb billno,a.carno,a.owner,a.owneraddress,b.maildate,b.mailnumber from (select distinct imagefilenameb,carno,owner,owneraddress from billbase where BillStatus=2 and imagefilenameb is not null "&strwhere&") a,(select distinct billno,carno,MailDate,mailnumber from StopBillMailHistory where billno is not null and mailnumber is not null and UserMarkResonID is null) b where a.imagefilenameb=b.billno and a.carno=b.carno order by billno"

	set rsfound=conn.execute(strSQL)
	tmpSQL=strwhere
	iDate1=Request("Sys_RecordDate1")
	iDate2=Request("Sys_RecordDate2")
%>

</head>
<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://10.104.10.246/traffic/smsx.cab#Version=6,1,432,1">
</object>
<%
CaseSN=0
If Not rsfound.Bof Then rsfound.MoveFirst 
While Not rsfound.Eof
%>
	<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="7" align="center" style="line-height:2;">
			<font size="3">
			台東縣警察局交通隊「路邊收費停車場停車費催繳通知單」送達證書未退回清冊
			</font>
		</td>
	</tr>
	<tr>
		<td colspan="3" align="left"><%
			If (not ifnull(request("Sys_RecordDate1"))) and (not ifnull(request("Sys_RecordDate2"))) Then

				Response.Write "查詢日期:"
				Response.Write left(iDate1,len(iDate1)-4)&"/"&Mid(iDate1,len(iDate1)-3,2)&"/"&right(iDate1,2)&" 至 "

				Response.Write left(iDate2,len(iDate2)-4)&"/"&Mid(iDate2,len(iDate2)-3,2)&"/"&right(iDate2,2)
			else
				Response.Write "查詢批號:"&trim(Request("Sys_BatChNumber"))
			end if
		%></td>
		<td colspan="4" align="right">Page <%=fix(CaseSN/25)+1%> of <%=pagecnt%> </td>
	</tr>
	<tr>
		<td colspan="7" align="left" colspan="2">登入者:<%=Session("Ch_Name")%></td>
	</tr>
	<tr><td colspan="7"><hr></td></tr>
	<tr>
		<td width="5%">序號</td>
		<td width="10%">催繳單號</td>
		<td width="6%">催繳車號</td>
		<td width="6%">郵寄日期</td>
		<td width="6%">大宗號碼</td>
		<td width="10%">收件人姓名</td>
		<td width="20%">收件人地址</td>
	</tr>
<%
for i=1 to 25
	if rsfound.eof then exit for
	CaseSN=CaseSN+1
%>
	<tr><td colspan="7"><hr></td></tr>
	<tr>
		<td><%=cdbl(CaseSN)%></td>
		<td><%=trim(rsfound("billno"))%></td>
		<td><%=trim(rsfound("carno"))%></td>
		<td><%=gInitDT(rsfound("maildate"))%></td>
		<td><%=trim(rsfound("mailnumber"))%></td>
		<td><%=funcCheckFont(trim(rsfound("owner")),20,1)%></td>
		<td><%=funcCheckFont(trim(rsfound("owneraddress")),20,1)%></td>
	</tr>
	
<%		
	rsfound.MoveNext
	next
%>
<tr><td colspan="7"><hr></td></tr>
</table>
<%
Wend
%>

<%
rsfound.close
set rsfound=nothing
%>
共計:   <%=CaseSN%>  筆

</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(true,20.08,10.08,5.08,5.08);
</script>
<%
conn.close
set conn=nothing
%>