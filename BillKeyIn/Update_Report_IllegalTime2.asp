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
	set rsCity=Nothing
	
If Trim(request("kinds"))="BillQry" Then
	IllTime=Trim(request("IllagelDate")) & " " & Left(Trim(request("IllagelTime")),2) & ":" & right(Trim(request("IllagelTime")),2) 
	strUpd="Update Billbase set IllegalDate=to_date('" & IllTime & "','YYYY/MM/DD/HH24/MI/SS') where Sn="&Trim(request("BillSn"))
	conn.execute strUpd
	ConnExecute "強制修改違規日期"&strUpd,353
%>
<script language="javascript">
<%
if trim(request("UpdType"))="1" then
%>
	opener.myForm.IllegalTime.value="<%=Trim(request("IllagelTime"))%>";
<%
end if 
%>
	alert("儲存完成!");

</script>
<%
End If 

strSql="select * from billbase where sn="&Trim(request("BillSn"))
Set rs1=conn.execute(strSql)
If Not rs1.eof then
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
					<%=Trim(rs1("BillNo"))%>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">車牌號碼</td>
				<td>
					<%=Trim(rs1("CarNo"))%>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">違規法條</td>
				<td>
					<%
				If Trim(rs1("Rule1"))<>"" Then
					response.write Trim(rs1("Rule1"))
				End If 
				If Trim(rs1("Rule2"))<>"" Then
					response.write "/" &Trim(rs1("Rule2"))
				End If 
				If Trim(rs1("Rule3"))<>"" Then
					response.write "/" &Trim(rs1("Rule3"))
				End If 
					
					%>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">違規地點</td>
				<td>
					<%=Trim(rs1("IllegalAddressID"))&" "&Trim(rs1("IllegalAddress"))%>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">違規日期</td>
				<td>
					<%=year(rs1("IllegalDate"))-1911 & "/" & Month(rs1("IllegalDate"))& "/" & Day(rs1("IllegalDate"))%>
					<input type="hidden" name="IllagelDate" value="<%=year(rs1("IllegalDate")) & "/" & Month(rs1("IllegalDate"))& "/" & Day(rs1("IllegalDate"))%>">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">違規時間</td>
				<td>
					<input type="text" name="IllagelTime" maxlength="4" value="<%=Right("00"&Hour(rs1("IllegalDate")),2)&Right("00"&Minute(rs1("IllegalDate")),2)%>">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBFBE3" align="center" colspan="2">
					<input type="submit" value="儲 存" onclick="funBillQry();">
					<input type="hidden" value="" name="kinds">
				</td>
			</tr>
		</table>	
	</form>
<%
End If 
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
	if (myForm.IllagelTime.value=="" || myForm.IllagelTime.value.length!=4){
		alert("請輸入違規時間!");
	}else{
		myForm.kinds.value="BillQry";
		myForm.submit();
	}
}
</script>
</html>
