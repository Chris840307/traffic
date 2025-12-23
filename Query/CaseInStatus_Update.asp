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
<style type="text/css">
<!--
.style1 {
color:red;
font-size: 20px; 
line-height:24px;
}

.style1a {
font-size: 24px; 
line-height:28px;
color:#9933FF;
}
-->
</style>
<title>強迫入案</title>
<%
'檢查是否可進入本系統
'AuthorityCheck(237)

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

if trim(request("kinds"))="DB_Update" then
	DelMemID=trim(Session("User_ID"))
	theBillSN=trim(request("DBillSN"))

	'抓單號
	theBillNO=""
	theCarNO=""
	strbillno="select BillNo,CarNo from BillBase where SN="&theBillSN
	set rsBillno=conn.execute(strbillno)
	if not rsBillno.eof then
		theBillNO=trim(rsBillno("BillNo"))
		theCarNO=trim(rsBillno("CarNo"))
	end if
	rsBillno.close
	set rsBillno=nothing

'if sys_City="台中市" or sys_City="台中縣" or sys_City="高雄縣" or sys_City="高雄市" or sys_City="彰化縣" or sys_City="台東縣" or sys_City="花蓮縣" or sys_City="基隆市" or sys_City="嘉義縣" or sys_City="嘉義市" or sys_City="澎湖縣" or sys_City="金門縣" or sys_City="南投縣" or sys_City="台南縣" or sys_City="台南市" or sys_City="連江縣" or sys_City="宜蘭縣" or sys_City="雲林縣" or sys_City="屏東縣" then
	strDci="select * from DciLog where BillSn="&theBillSN&" and ExchangeTypeID='W'"
	set rsDci=conn.execute(strDci)
	if not rsDci.eof then
		strIns="Insert into DCISTATUSUPDATE values(" & trim(rsDci("BillSn")) & ",'" & trim(rsDci("BillNO")) & "'" &_
			",'" & trim(rsDci("CarNo")) & "','" & trim(rsDci("DciReturnStatusID")) & "'" &_
			",'" & trim(rsDci("DciErrorCarData")) & "',sysdate," & DelMemID &_
			")"
		conn.execute strIns
	end if
	rsDci.close
	set rsDci=nothing
'end if

		strUpdDelTemp="Update BillBaseDciReturn set status='Y',DCIERRORCARDATA='0'" &_
			" where BillNo='"&theBillNO&"' and CarNo='"&theCarNO&"' and ExchangeTypeID='W'"
		conn.execute strUpdDelTemp

		strUpdDel="Update DciLog set DciReturnStatusID='Y',DCIERRORCARDATA='0' " &_
			" where billSN="&theBillSN&" and ExchangeTypeID='W'"
		conn.execute strUpdDel

%>
<script language="JavaScript">
	opener.myForm.submit();
	alert("本案件尚未入案至監理站，請務必請監理站人員代為建檔!!");
	alert("強制入案完成\n再次提醒，請務必請監理站人員代為建檔!!");
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
				<td colspan="4"><strong>舉發單強制入案</strong></td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td width="15%" align="center" height="30">
				<strong>失竊註銷 (但查詢署端系統後發現車輛實際已經尋回)</strong><br>
				==>此類資料導因警政署與監理站資料更新不一致所引起，須手動移送監理站。
				<br>
				<br>

				<strong>證號不正確 (外籍人士)</strong><br>
				==>此類資料因監理站無外籍人士資料導致無法入案，須手動移送監理站。
				<br>
				<br>
				是否要將入案狀態改為正常？
				<br>&nbsp;<br/>&nbsp;
				<strong><span class="style1">＊ 重要!!! ＊</span><br><span class="style1a">強制入案監理站不會有資料</span><span class="style1"><br>，請務必手動移送至監理站</span></strong>
				<strong><span class="style1"><br>(強制入案後無法回復為未入案狀態)</span></strong>
				<br/>&nbsp;<br/>
				<input type="checkbox" name="ConfirmUpdate" value="1" onclick="ConfirmUpdate1();">我已閱讀上述事項 ( 需確認後才可按確定紐 )
				<br/>&nbsp;
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
				<input type="button" name="bUpdate1" value=" 確 定 " onclick="BillUpdate();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
'						if CheckPermission(237,4)=false then
							response.write "disabled"
'						end if
						%>>

				<input type="button" name="close" value=" 離 開 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="DBillSN" value="<%=trim(request("BillSN"))%>">
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
function BillUpdate(){
	myForm.kinds.value="DB_Update";
	myForm.submit();
}

function ConfirmUpdate1(){
	if (myForm.ConfirmUpdate.checked==true)
	{
		myForm.bUpdate1.disabled=false;
	}else{
		myForm.bUpdate1.disabled=true;
	}

}
</script>
</html>
