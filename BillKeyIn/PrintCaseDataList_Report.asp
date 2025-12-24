<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://10.104.10.246/traffic/smsx.cab#Version=6,1,432,1">
</object>
<html>
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>逕舉(照片)建檔資料清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 800
Response.flush
%>
<%
'權限
'AuthorityCheck(234)

RecordDate=split(gInitDT(date),"-")
	if trim(request("CallType"))="1" then
		strwhere=" and a.BillStatus in ('0') and RecordStateID=0 and a.RecordMemberID="&session("User_ID")
	else
		strwhere=Session("PrintCarDataSQL")	
	end if
	Session.Contents.Remove("PrintCaseDataSQLxls")
	Session("PrintCaseDataSQLxls")=strwhere	

	strCnt="select count(*) as cnt from BillBase a,MemberData b where a.BillTypeID='2' and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by a.RecordDate"
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		if trim(rsCnt("cnt"))="0" then
			pagecnt=1
		else
			pagecnt=fix(Cint(rsCnt("cnt"))/50+0.9999999)
		end if
	end if
	rsCnt.close
	set rsCnt=nothing

	strSQL="select a.SN,a.BillNo,a.CarNo,a.CarSimpleID,a.IllegalDate,a.IllegalAddress,a.BillUnitID,a.DriverID,a.BillMemID1,a.BillMem1,a.Rule1,a.Rule2,a.BillFillDate,a.DeallineDate,a.MemberStation from BillBase a,MemberData b where a.BillTypeID='2' and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by a.RecordDate"
	set rsfound=conn.execute(strSQL)

	tmpSQL=strwhere
'response.write strSQL
%>

</head>
<body>
<form name=myForm method="post">
<%
CaseSN=0
If Not rsfound.Bof Then rsfound.MoveFirst 
While Not rsfound.Eof
	if CaseSN>0 then response.write "<div class=""PageNext""></div>"
%>
<table width="700" border="0" cellpadding="1" cellspacing="0">
	<tr>
		<td colspan="2" align="center">
			<font size="3">逕舉(照片)資料建檔清冊</font>
		</td>
	</tr>
	<tr>
		<td align="left"></td>
		<td align="right">Page <%=fix(CaseSN/50)+1%> of <%=pagecnt%> </td>
	</tr>
	<tr>
		<td align="left">登入者:<%=Session("Ch_Name")%></td>
		<td align="right">列印日期:<%=year(now)-1911&"/"&Right("00"&month(now),2)&"/"&Right("00"&day(now),2)%></td>
	</tr>
</table>
	<table width="700" border="1" cellpadding="0" cellspacing="0">
	<tr>
		<td>
		<table width="700" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="5%">編號</td>
			<td width="8%">登入日期</td>
			<td width="10%">車號</td>
			<td width="7%">車種</td>
			<td width="8%">違規日期</td>
			<td width="8%">違規時間</td>
			<td width="9%">法條一</td>
			<td width="9%">法條二</td>
			<td width="8%">舉發員警</td>
			<td width="28%">違規地點</td>
		</tr>
		</table>
		</td>
	</tr>
<%
for i=1 to 50
	if rsfound.eof then exit for
	CaseSN=CaseSN+1
%>
	<tr>
		<td>
	<table width="700" border="0" cellpadding="0" cellspacing="0">
		<tr>
		<td width="5%"><%=CaseSN%></td>
		<td width="8%"><%=Right("00"&year(now)-1911,2)&Right("00"&month(now),2)&Right("00"&day(now),2)%></td>
		<td width="10%"><%=trim(rsfound("CarNo"))%></td>
		<td width="7%"><%=trim(rsfound("CarSimpleID"))%></td>
		<td width="8%"><%
		if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
			response.write Right("00"&year(trim(rsfound("IllegalDate")))-1911,2)&Right("00"&month(trim(rsfound("IllegalDate"))),2)&Right("00"&day(trim(rsfound("IllegalDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td width="8%"><%
		if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
			response.write Right("00"&hour(trim(rsfound("IllegalDate"))),2)&Right("00"&minute(trim(rsfound("IllegalDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td width="9%"><%
		if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
			response.write trim(rsfound("Rule1"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td width="9%"><%
		if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
			response.write trim(rsfound("Rule2"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td width="8%"><%
		if trim(rsfound("BillMem1"))<>"" and not isnull(rsfound("BillMem1")) then
			response.write trim(rsfound("BillMem1"))
		else
			response.write "&nbsp;"
		end if		%></td>
		<td width="28%"><%
		if trim(rsfound("illegalAddress"))<>"" and not isnull(rsfound("illegalAddress")) then
			response.write trim(rsfound("illegalAddress"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		</tr>
	</table>
		<td>
	</tr>

<%		
	rsfound.MoveNext
	next
%>
	</table>
<%
Wend
rsfound.close
set rsfound=nothing
%>
共計:   <%=CaseSN%>  筆

</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
	return win;
}
function funcPrintCaseDataListExecel(){
	//UrlStr="PrintCaseDataList_Execel.asp";
	//newWin(UrlStr,"inputWin",790,550,0,0,"yes","yes","yes","no");
	location='PrintCaseDataList_Excel.asp';
}
function DP(){
	window.focus();
	window.print();
}

printWindow(true,7,5.08,5.08,5.08);
</script>
<%
conn.close
set conn=nothing
%>