<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/css.txt"-->
<!--#include file="sqlDCIExchangeData.asp"-->
<title>舉發單整批撤銷送達</title>
<% Server.ScriptTimeout = 800 %>
<%
'檢查是否可進入本系統
'AuthorityCheck(237)

DelMemID=trim(Session("User_ID"))
theBatchNumber=trim(request("BatchNumber"))
'theDelType=trim(request("DelType"))	'單筆或多筆刪除

	'抓單號
'	theBillNO=""
'	theCarNO=""
'	strbillno="select BillNo,CarNo from BillBase where SN="&theBillSN
'	set rsBillno=conn.execute(strbillno)
'	if not rsBillno.eof then
'		theBillNO=trim(rsBillno("BillNo"))
'		theCarNO=trim(rsBillno("CarNo"))
'	end if
'	rsBillno.close
'	set rsBillno=nothing

if trim(request("kinds"))="DB_BillDel" then
		strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			theBatchTime=(year(now)-1911)&"N"&trim(rsSN("SN"))
		end if
		rsSN.close
		set rsSN=nothing

		strDelS="select a.SN,a.BillNo,a.BillTypeID,a.CarNo,a.BillUnitID,a.RecordDate,a.RecordMemberID,a.RecordStateID,a.BillStatus from BillBase a,DciLog b where b.BatchNumber='"&theBatchNumber&"' and a.SN=b.BillSN"
		set rsDelS=conn.execute(strDelS)
		If Not rsDelS.Bof Then rsDelS.MoveFirst 
		While Not rsDelS.Eof
			if trim(rsDelS("RecordStateID"))<>"-1" then
				
				funcStoreAndSendToGov conn,trim(rsDelS("SN")),trim(rsDelS("BillNo")),trim(rsDelS("BillTypeID")),trim(rsDelS("CarNo")),trim(rsDelS("BillUnitID")),trim(rsDelS("RecordDate")),trim(Session("User_ID")),theBatchTime
			end if
		rsDelS.MoveNext
		Wend
		rsDelS.close
		set rsDelS=nothing

%>
		<script language="JavaScript">
			alert("撤銷送達批號：<%=theBatchTime%>，請等待監理所回傳!");
			opener.myForm.submit();
			window.close();
		</script>
<%
end if

%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td ><strong>舉發單撤銷送達</strong></td>
			</tr>
			<tr>
				<td align="center" height="42" bgcolor="#EBFBE3">確定要將批號：<%=trim(request("BatchNumber"))%>，整批做撤銷送達嗎?</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" align="center">
				<input type="button" name="close" value=" 確 定 " onclick="BillDel();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
						if CheckPermission(234,4)=false or CountCaseIN>7 or CaseNotReturn=1 then
							'response.write "disabled"
						end if
						%>>
				<input type="button" name="close" value=" 取 消 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="BatchNumber" value="<%=trim(request("BatchNumber"))%>">
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
function BillDel(){
		myForm.kinds.value="DB_BillDel";
		myForm.submit();
}
</script>
</html>
