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
.pageprint {
  margin-left: 0mm;
  margin-right: 0mm;
  margin-top: 0mm;
  margin-bottom: 0mm;
}
</style>
<title>行人攤販資料建檔清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
'權限
'AuthorityCheck(234)

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

	strCnt="select count(*) as cnt from "&BasSQL
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

	strSQL="select a.SN,a.BillNo,a.IllegalDate,a.Rule1,a.BillUnitID,a.Driver," &_
			"a.DriverID,a.RuleVer,a.BillMem1,a.BillMem2,a.ForFeit1," &_
			"(select UnitName from Unitinfo where UnitID=a.BillUnitID) UnitName," &_
			"(select MAX(PayDate) PayDate from PasserPay where billsn=a.sn) PayDate," &_
			"(select MAX(PayNo) PayNo from PasserPay where billsn=a.sn) PayNo," &_
			"(select Sum(NVL(PayAmount,0)) PayAmount from PasserPay where billsn=a.sn) PayAmount," &_
			"(select MAX(Decode(IsLate,'0','','1','逾期')) IsLate from PasserPay where billsn=a.sn) IsLate," &_
			"(select MAX(Note) Note from PasserPay where billsn=a.sn) Note," &_
			"(select MAX(SendDate) SendDate from PasserSend where billsn=a.sn) SendDate" &_
			" from PasserBase a where a.RecordStateID=0 and a.billstatus<>9 and Exists(select 'Y' from "&BasSQL&" where SN=a.SN)  order by IllegalDate"
	set rsfound=conn.execute(strSQL)
	tmpSQL=strwhere
%>

</head>
<body class="pageprint">
<form name=myForm method="post">
<%
CntSN=1
'If Not rsfound.Bof Then rsfound.MoveFirst
SumMemy=0
chkUnitID=""
While Not rsfound.Eof
if CntSN >1 then response.write "<div class=""PageNext""></div>"
%>
<table width="100%" border="1" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="14" align="center">
			<font size="3">每月繳納違反道路障礙罰單明細表</font>
		</td>
	</tr>
	<tr>
		<td colspan="7" align="left">列印日期:<%=year(now)-1911&"/"&Right("00"&month(now),2)&"/"&Right("00"&day(now),2)%></td>
		<td colspan="7" align="right">Page <%=fix(CntSN/20)+1%> of <%=pagecnt%> </td>
	</tr>
	<tr>
		<td colspan="14" align="left" colspan="2">登入者:<%=Session("Ch_Name")%></td>
	</tr>
	<%
		'Response.Write "<tr><td colspan=""14""><hr></td></tr>"
	%>	
	<tr>
		<td>號次</td>
		<td>舉發單位</td>
		<td>舉發人員</td>
		<td>舉發單號碼</td>
		<td>違規時間</td>
		<td>違規人姓名</td>
		<td>違規人證號</td>
		<td>舉發法條</td>
		<td>原罰款額</td>
		<td>累計原罰款額</td>
		<td>實際罰款額</td>
		<td>累計實際罰款額</td>
		<td>備註</td>
	</tr>
<%
'If chkUnitID="" Then chkUnitID=trim(rsfound("BillUnitID"))
ForFeitSum=0
for i=1 to 20
	if rsfound.eof then exit for
'	If chkUnitID<>trim(rsfound("BillUnitID")) Then
'		chkUnitID=trim(rsfound("BillUnitID"))
'		exit for
'	end if
	if IsNull(rsfound("ForFeit1")) Then
		iForFeit=0
	else
		iForFeit=rsfound("ForFeit1")
	end if
	SumMemy=SumMemy+cdbl(iForFeit)

	strRule1="select Level1 from Law where ItemID='"&trim(rsfound("Rule1"))&"' and VERSION='"&trim(rsfound("RuleVer"))&"'"
	
	ForFeit=0
	set rsRule1=conn.execute(strRule1)
	If not rsRule1.eof Then ForFeit=rsRule1("Level1")
	rsRule1.close

	ForFeitSum=ForFeitSum+cdbl(ForFeit)
	
'	chkUnitID=trim(rsfound("BillUnitID"))
	'Response.Write "<tr><td colspan=""14""><hr></td></tr>"
%>
	<tr>
		<td><%=CntSN%></td>
		<td><%=trim(rsfound("UnitName"))%></td>
		<td width="9%"><%=trim(rsfound("BillMem1"))%></td>
		<td><%=trim(rsfound("BillNo"))%></td>
		<td><%=gInitDT(trim(rsfound("IllegalDate")))&" "&right("00"&hour(rsfound("IllegalDate")),2)&":"&right("00"&minute(rsfound("IllegalDate")),2)%></td>
		<td><%=trim(rsfound("Driver"))%></td>
		<td><%=trim(rsfound("DriverID"))%></td>
		<td><%=trim(rsfound("Rule1"))%></td>
		<td><%=trim(ForFeit)%></td>
		<td><%=ForFeitSum%></td>
		<td><%=trim(rsfound("ForFeit1"))%></td>
		<td><%=SumMemy%></td>
		<td><%
		response.write trim(rsfound("IsLate"))&trim(rsfound("Note"))
		If Not ifnull(trim(rsfound("SendDate"))) Then response.write "已移送行政執行處"
		%></td>
	</tr>
	
<%		CntSN=CntSN+1
	rsfound.MoveNext
next
	%>
<tr><td colspan="14"><hr></td></tr>
</table>
<%
Wend
rsfound.close
set rsfound=nothing
%>
共計:   <%=CntSN-1%>  筆

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
</script>
<%
conn.close
set conn=nothing


fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_未繳費明細表.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>