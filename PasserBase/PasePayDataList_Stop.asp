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

	strRul="select Value from Apconfigure where ID=3"
	set rsRul=conn.execute(strRul)
	RuleVer=trim(rsRul("Value"))
	rsRul.Close

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

	strCnt="select count(1) as cnt from "&BasSQL
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		if trim(rsCnt("cnt"))="0" then
			pagecnt=1
		else
			pagecnt=fix(Cint(rsCnt("cnt"))/14+0.9999999)
		end if
	end if
	rsCnt.close
	set rsCnt=nothing

	if request("PayDate1")<>"" and request("PayDate2")<>""then
		ArgueDate1=gOutDT(request("PayDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("PayDate2"))&" 23:59:59"

		paystr=" and PayDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	strSQL="select a.SN,a.BillNo,a.DeallIneDate,a.Rule1," &_
			"(select MAX(ForFeit) ForFeit from PasserPay where billsn=a.sn) ForFeit," &_
			"(select MAX(PayDate) PayDate from PasserPay where billsn=a.sn"&paystr&") PayDate," &_
			"(select MAX(PayNo) PayNo from PasserPay where billsn=a.sn"&paystr&") PayNo," &_
			"(select Sum(NVL(PayAmount,0)) PayAmount from PasserPay where billsn=a.sn"&paystr&") PayAmount," &_
			"(select MAX(Decode(IsLate,'0','','1','逾期')) IsLate from PasserPay where billsn=a.sn"&paystr&") IsLate," &_
			"(select MAX(Note) Note from PasserPay where billsn=a.sn) Note" &_
			" from PasserBase a where a.RecordStateID=0 and Exists(select 'Y' from "&BasSQL&" where SN=a.SN) and exists(select 'Y' from PasserPay where billsn=a.sn"&paystr&") order by PayDate"

	set rsfound=conn.execute(strSQL)
	tmpSQL=strwhere

%>

</head>
<body class="pageprint">
<form name=myForm method="post">
<%
CntSN=1
'If Not rsfound.Bof Then rsfound.MoveFirst
SumMemy=0:sumLevel1=0
While Not rsfound.Eof
if CntSN >1 then response.write "<div class=""PageNext""></div>"
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="13" align="center">
			<font size="3">每月繳納違反道路障礙罰單明細表</font>
		</td>
	</tr>
	<tr>
		<td colspan="6" align="left">列印日期:<%=year(now)-1911&"/"&Right("00"&month(now),2)&"/"&Right("00"&day(now),2)%></td>
		<td colspan="6" align="right">Page <%=fix(CntSN/14)+1%> of <%=pagecnt%> </td>
	</tr>
	<tr>
		<td colspan="13" align="left" colspan="2">登入者:<%=Session("Ch_Name")%></td>
	</tr>
	<tr><td colspan="13"><hr></td></tr>
	<tr>
		<td>號次</td>
		<td>繳款日期</td>
		<td>繳款單文號</td>
		<td>舉發單號碼</td>
		<td>逾期月份</td>
		<td>舉發法條</td>
		<td>原罰款額</td>
		<td>累計原罰款額</td>
		<td>實際罰款額</td>
		<td>繳費額</td>
		<td>累計繳費額</td>
		<td>備註</td>
	</tr>
<%

for i=1 to 14
	if rsfound.eof then exit for

	if IsNull(rsfound("PayAmount").Value) Then
		PayAmount=0
	else
		PayAmount=rsfound("PayAmount")
	end if
	SumMemy=SumMemy+cdbl(PayAmount)
%>
	<tr><td colspan="13"><hr></td></tr>
	<tr>
		<td><%=CntSN%></td>
		<td>
		<%
		Response.Write "<table border=0>"
			strSQL="Select PayNo,PayDate from PasserPay where billsn="&rsfound("SN")&" "&paystr&" order by PayDate"
			set rs=conn.execute(strSQL)
			while Not rs.eof
				Response.Write "<tr><td>"&gInitDT(trim(rs("PayDate")))&"</td></tr>"
				rs.movenext
			wend
			rs.close
		Response.Write "</table>"
		%></td>
		<td width="9%">
		<%
		Response.Write "<table border=0>"
			strSQL="Select PayNo,PayDate from PasserPay where billsn="&rsfound("SN")&" "&paystr&" order by PayDate"
			set rs=conn.execute(strSQL)
			while Not rs.eof
				Response.Write "<tr><td>"&trim(rs("PayNo"))&"</td></tr>"
				rs.movenext
			wend
			rs.close
		Response.Write "</table>"
		%></td>
		<td><%=trim(rsfound("BillNo"))%></td>
		<td><%=gInitDT(trim(rsfound("DeallIneDate")))%></td>
		<td><%=trim(rsfound("Rule1"))%></td>
		<%
			strSQL="select Level1 from law where itemid='"&trim(rsfound("Rule1"))&"' and version="&RuleVer
			set rslaw=conn.execute(strSQL)
			Response.Write "<td>"
			Response.Write trim(rslaw("Level1"))
			Response.Write "</td>"

			sumLevel1=sumLevel1+cdbl(rslaw("Level1"))
			rslaw.close

			Response.Write "<td>"
			Response.Write sumLevel1
			Response.Write "</td>"
		%>
		<td><%=trim(rsfound("ForFeit"))%></td>
		<td><%

		Response.Write "<table border=0>"
			strSQL="Select PayAmount,PayDate from PasserPay where billsn="&rsfound("SN")&" "&paystr&" order by PayDate"
			set rs=conn.execute(strSQL)
			while Not rs.eof
				Response.Write "<tr><td>"&trim(rs("PayAmount"))&"</td></tr>"
				rs.movenext
			wend
			rs.close
		Response.Write "</table>"
		
		%></td>
		<td><%=SumMemy%></td>
		<td><%=trim(rsfound("IsLate"))&trim(rsfound("Note"))%></td>
	</tr>
	
<%		CntSN=CntSN+1
	rsfound.MoveNext
	next
	%>
<tr><td colspan="10"><hr></td></tr>
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
%>