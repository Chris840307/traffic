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
<title>修改違規單號</title>
<%
'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1'  cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="2">修改違規單號</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">舉發單號</td>
				<td>
					<input type="text" name="billno" value="<%=Trim(request("billno"))%>" size="12" maxlength="9" onkeyup="value=value.toUpperCase()">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">車牌號碼</td>
				<td>
					<input type="text" name="carno" value="<%=Trim(request("carno"))%>" size="12" maxlength="8" onkeyup="value=value.toUpperCase()">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBFBE3" align="center" colspan="2">
					<input type="submit" value="確 定" onclick="funBillQry();">
					<input type="hidden" value="" name="kinds">
				</td>
			</tr>
		</table>	
		<br>
<%
If Trim(request("kinds"))="BillQry" Then
	%>
		
		<table width='100%' border='1'  cellpadding="1">
			<tr>
				<td>單號</td>
				<td>舉發單類別</td>
				<td>車號</td>
				<td>法條</td>
				<td>違規時間</td>
				<td>違規地點</td>
				<td>操作</td>
			</tr>
<%
	strSql=""
	If Trim(request("billno"))<>"" Then
		strSql=strSql & " and Billno='" & trim(request("billno")) & "' "
	End If 
	If Trim(request("carno"))<>"" Then
		strSql=strSql & " and carno='" & trim(request("carno")) & "' "
	End If 
	
	strBillView="select * from billbase where RecordStateID=0 " & strSql & " order by RecordDate desc"
	set rsBillView=conn.execute(strBillView)
	If Not rsBillView.Bof Then rsBillView.MoveFirst 
	While Not rsBillView.Eof
%>
		
			<tr>
				<td><%=Trim(rsBillView("BillNo"))%></td>
				<td><%
				If Trim(rsBillView("Billtypeid"))="1" Then
					response.write "攔停"
				Else
					response.write "逕舉"
				End If 
				%></td>
				<td><%=Trim(rsBillView("Carno"))%></td>
				<td><%
				If Trim(rsBillView("Rule1"))<>"" Then
					response.write Trim(rsBillView("Rule1"))
				End If 
				If Trim(rsBillView("Rule2"))<>"" Then
					response.write "/" &Trim(rsBillView("Rule2"))
				End If 
				If Trim(rsBillView("Rule3"))<>"" Then
					response.write "/" &Trim(rsBillView("Rule3"))
				End If 
				%></td>
				<td><%
				response.write year(rsBillView("IllegalDate"))-1911 & "/" & Month(rsBillView("IllegalDate"))& "/" & Day(rsBillView("IllegalDate")) &_
					 " " & Hour(rsBillView("IllegalDate")) & " : " & Minute(rsBillView("IllegalDate")) 
				%></td>
				<td><%=Trim(rsBillView("IllegalAddressID"))&" "&Trim(rsBillView("IllegalAddress"))%></td>
				<td>
					<input type="Button" value="修改違規時間" onclick="UpdateTime('<%=rsBillView("Sn")%>')">
				</td>
			</tr>
<%
	rsBillView.MoveNext
	Wend
	rsBillView.close
	set rsBillView=Nothing
	%>
		</table>
<%
End If
%>
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
function UpdateTime(BillSn){
	UrlStr="Update_report_IllegalTime2.asp?BillSn=" + BillSn;
	newWin(UrlStr,"Update_report_IllegalTime2",980,550,0,0,"yes","yes","yes","no");
}
</script>
</html>
