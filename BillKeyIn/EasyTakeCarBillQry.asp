<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>舉發單查詢</title>
<%
if trim(request("kinds"))="BillQry" then
	SqlEx=""
	if trim(request("billno"))<>"" then
		SqlEx=" and Billno='"&trim(request("billno"))&"'"
	end if
	if trim(request("carno"))<>"" then
		SqlEx=SqlEx&" and CarNo='"&trim(request("carno"))&"'"
	end if
	strQry="select SN,BillTypeID,RecordDate,ImageFileName from BillBase where BillStatus='9' and Recordstateid=0 and RecordDate between TO_DATE('"&date&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"&SqlEx&" order by RecordDate"
	'response.write strQry
	'response.end
	set rsQry=conn.execute(strQry)
	if not rsQry.eof then
		'攔停
		if trim(rsQry("BillTypeID"))="1" then
			BillTime_StopTmp=DateAdd("s" , 1, rsQry("RecordDate"))
			Session("BillTime_Stop")=date&" "&hour(BillTime_StopTmp)&":"&minute(BillTime_StopTmp)&":"&second(BillTime_StopTmp)

			strSqlCnt="select count(*) as cnt from BillBase where BillTypeID='1' and BillStatus in ('9') and RecordStateID=0 and RecordMemberID="&trim(Session("User_ID"))&" and RecordDate between TO_DATE('"&date&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and RecordDate < TO_DATE('"&Session("BillTime_Stop")&"','YYYY/MM/DD/HH24/MI/SS')"
			set rsCnt1=conn.execute(strSqlCnt)
				Session("BillOrder_Stop")=trim(rsCnt1("cnt"))+1
			rsCnt1.close
			set rsCnt1=nothing
	%>
	<script language="JavaScript">
		opener.location="BillKeyIn_TakeCar_Back.asp?PageType=Back";
		window.close();
	</script>
	<%	'逕舉
		elseif trim(rsQry("BillTypeID"))="2" and (trim(rsQry("ImageFileName"))="" or isnull(rsQry("ImageFileName"))) then
			BillTime_ReportTmp=DateAdd("s" , 1, rsQry("RecordDate"))
			Session("BillTime_Report")=date&" "&hour(BillTime_ReportTmp)&":"&minute(BillTime_ReportTmp)&":"&second(BillTime_ReportTmp)

			strSqlCnt="select count(*) as cnt from BillBase where BillTypeID='2' and BillStatus in ('9') and RecordStateID=0 and RecordMemberID="&trim(Session("User_ID"))&" and RecordDate between TO_DATE('"&date&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and RecordDate < TO_DATE('"&Session("BillTime_Report")&"','YYYY/MM/DD/HH24/MI/SS') and ImageFileName is null"
			set rsCnt1=conn.execute(strSqlCnt)
				Session("BillOrder_Report")=trim(rsCnt1("cnt"))+1
			rsCnt1.close
			set rsCnt1=nothing
	%>
	<script language="JavaScript">
		opener.location="BillKeyIn_Report_Back.asp?PageType=Back";
		window.close();
	</script>
	<%
		'逕舉影像建檔
		elseif trim(rsQry("BillTypeID"))="2" and (trim(rsQry("ImageFileName"))<>"" and not isnull(rsQry("ImageFileName"))) then
			BillTime_ImageTmp=DateAdd("s" , 1, rsQry("RecordDate"))
			Session("BillTime_Image")=date&" "&hour(BillTime_ImageTmp)&":"&minute(BillTime_ImageTmp)&":"&second(BillTime_ImageTmp)

			strSqlCnt="select count(*) as cnt from BillBase where BillTypeID='2' and BillStatus in ('9') and RecordStateID=0 and RecordMemberID="&trim(Session("User_ID"))&" and RecordDate between TO_DATE('"&date&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and RecordDate < TO_DATE('"&Session("BillTime_Image")&"','YYYY/MM/DD/HH24/MI/SS') and ImageFileName is not null"
			set rsCnt1=conn.execute(strSqlCnt)
				Session("BillOrder_Image")=trim(rsCnt1("cnt"))+1
			rsCnt1.close
			set rsCnt1=nothing
	%>
	<script language="JavaScript">
		opener.location="BillKeyIn_TakeCar_Back.asp?PageType=Back";
		window.close();
	</script>
	<%
		end if
	else
%>
<script language="JavaScript">
	alert("此舉發單未建檔或已經入案至監理所!");
	window.close();
</script>
<%
	end if
	rsQry.close
	set rsQry=nothing
end if
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post" onsubmit="return funBillQry();">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="2">舉發單查詢(針對本日未入案案件做查詢)</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">舉發單號</td>
				<td>
					<input type="text" name="billno" value="" size="12" maxlength="9" onkeyup="value=value.toUpperCase()">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">車牌號碼</td>
				<td>
					<input type="text" name="carno" value="" size="12" maxlength="7" onkeyup="value=value.toUpperCase()">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBFBE3" align="center" colspan="2">
					<input type="submit" value="確 定">
					<input type="hidden" value="" name="kinds">
				</td>
			</tr>
		</table>		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	winopen=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	winopen.focus();
	return win;
}
function funBillQry(){
	if (myForm.billno.value=="" && myForm.carno.value==""){
		alert("請輸入舉發單號或車牌號碼任一條件!");
	}else{
		myForm.kinds.value="BillQry";
		myForm.submit();
	}
}
myForm.billno.focus();
</script>
</html>
