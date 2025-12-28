<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close


sys_City=replace(sys_City,"台中縣","台中市")
sys_City=replace(sys_City,"台南縣","台南市")

showCreditor=false

if sys_City="台中市" or sys_City = "彰化縣" or sys_City = "台南市" or sys_City = "高雄市" or sys_City = "高雄縣" or sys_City="宜蘭縣" or sys_City="基隆市" or sys_City="澎湖縣" or sys_City="屏東縣" then
	showCreditor=true
end If

if request("DB_Selt")="Save" then
	BillSN=request("BillSN")
	If not ifnull(request("Sys_PasserJude")) Then
		strSQL="delete from PasserJude where BillSN in("&BillSN&")"
		conn.execute(strSQL)
		
		strSQL="update PasserBase set forfeit1=(select max(level1) from law where version=2 and itemid=PasserBase.Rule1) where sn in("&BillSN&")"
		conn.execute(strSQL)

		strSQL="update PasserBase set forfeit2=(select max(level1) from law where version=2 and itemid=PasserBase.Rule2) where sn in("&BillSN&") and Rule2 is not null"
		conn.execute(strSQL)
	End if

	If not ifnull(request("Sys_PasserUrge")) Then
		strSQL="delete from PasserUrge where BillSN in("&BillSN&")"
		conn.execute(strSQL)
	End if

	If not ifnull(request("Sys_PasserSend")) Then
		if showCreditor then


			strSQL="delete from PasserSend where BillSN in("&BillSN&") and not exists(select 'N' from PasserCreditor where BillSN=PasserSend.BillSN)"

			conn.execute(strSQL)
			
			strSQL="delete from PasserSendDetail where BillSN in("&BillSN&") and not exists(select 'N' from PasserCreditor where BillSN=PasserSendDetail.BillSN)"

			conn.execute(strSQL)

		else

			strSQL="delete from PasserSend where BillSN in("&BillSN&")"

			conn.execute(strSQL)
		end if
		
	End if

	If not ifnull(request("Sys_Creditor")) Then

		strSQL="delete from PasserSendDetail where BillSN in("&BillSN&") and SENDDATE=(select max(SENDDATE) from PasserSendDetail ta where not exists(select 'N' from PasserCreditor where SendDetailSN=ta.SN) and BillSN=PasserSendDetail.BillSN) and not exists(select 'N' from PasserCreditor where SendDetailSN=PasserSendDetail.SN)"
		conn.execute(strSQL)

		strSQL="Update PasserSend set SendDate=(select max(SendDate) from PasserSendDetail where BillSN=PasserSend.BillSN) where BillSN in("&BillSN&")"
		conn.execute(strSQL)

		strSQL="Update PasserSend set opengovnumber=(select opengovnumber from PasserSendDetail where SendDate=PasserSend.SendDate and BillSN=PasserSend.BillSN),sendnumber=(select sendnumber from PasserSendDetail where SendDate=PasserSend.SendDate and BillSN=PasserSend.BillSN) where BillSN in("&BillSN&")"
		conn.execute(strSQL)

	End If 

	response.write "<script language=""JavaScript"">"
	response.write "alert('已取消所選取的項目');"
	response.write "window.opener.myForm.submit();"
	response.write "self.close();"
	response.write "</script>"
end if
If Not ifnull(request("Sys_SendBillSN")) Then
	Sys_SendBillSN=request("Sys_SendBillSN")
else
	Sys_SendBillSN=request("hd_BillSN")
End if
%>
<TITLE> 資料回復系統 </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<!--#include virtual="traffic/Common/css.txt"-->
</HEAD>
<BODY>
<form name=myForm method="post">
<table width="100%" border="0" bgcolor="#ffffff">
	<tr>
		<td height="27" bgcolor="#FFCC33" class="pagetitle">資料回復系統</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						欲取消項目 ( 會針對<b>勾選</b>項目進行狀態回復 )
					</td>
				</tr>
				<tr>
					<td>
						<input class="btn1" type="checkbox" name="Sys_PasserJude" value="1">
						取消裁決資料
						<input class="btn1" type="checkbox" name="Sys_PasserUrge" value="1">
						取消催告資料
						<%if sys_City <> "彰化縣" or trim(Session("Credit_ID"))="A000000000" then %>
							<input class="btn1" type="checkbox" name="Sys_PasserSend" value="1">
							取消移送資料
							<%if showCreditor then%>
								<input class="btn1" type="checkbox" name="Sys_Creditor" value="1">
								取消再次移送
							<%end if%>
						<%end If%>
						<input type="button" name="btnSelt" value="確定" onclick="funSelt();">
						<input name="Submit433222" type="button" class="style3" value=" 關 閉 " onclick="self.close();">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td height="20" bgcolor="#FFDD77">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="">
<input type="Hidden" name="BillSN" value="<%=Sys_SendBillSN%>">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
function funSelt(){
	if(myForm.BillSN.value!=''){
		myForm.DB_Selt.value="Save";
		myForm.submit();
	}
}
</script>