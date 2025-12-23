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
<title>舉發單刪除</title>
<%
'檢查是否可進入本系統
'AuthorityCheck(237)

if trim(request("kinds"))="DB_BillDel" then
	DelMemID=trim(Session("User_ID"))
	theBillSN=trim(request("BillSN"))

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

		NoteTmp=Replace(Replace(Trim(request("Note")),"'",""),"--","")
			
		'該筆紀錄的打驗資料表的 BILLSTATUS 更新為 6
		'strUpdDelTemp="Update BillBaseTmp set billstatus='6',RecordStateID=-1,DelMemberID="&Session("User_ID")&" where BillNo='"&theBillNO&"'"
		'conn.execute strUpdDelTemp

		'更新該筆紀錄的 BILLSTATUS 更新為 6
		strUpdDel="Update BillBase set BillMemID1="&Trim(Session("User_ID"))&",BillMem1='"&Trim(Session("Ch_Name"))&"',BillFillerMemberID="&Trim(Session("User_ID"))&",BillFiller='"&Trim(Session("Ch_Name"))&"',billstatus='6',RecordStateID=-1,DelMemberID="&Session("User_ID")&" where SN="&theBillSN
		conn.execute strUpdDel

		'寫入刪除原因(判斷是否已經有 有的話就update)
		strReaCheck="select * from BillDeleteReason where BillSN="&theBillSN
		set rsReaCheck=conn.execute(strReaCheck)
		if rsReaCheck.eof then
			strReaDel="Insert into BillDeleteReason(BillSN,DelDate,DelReason,Note)" &_
				" values("&theBillSN&",sysdate,'QQ','"&NoteTmp&"')" 
		else
			strReaDel="Update BillDeleteReason set DelDate=sysdate,DelReason='QQ'" &_
				",Note='"&NoteTmp&"' where BillSN="&theBillSN
		end If
		conn.execute strReaDel
		'ConnExecute "舉發單刪除 單號:"&theBillNO&" 車號:"&theCarNO&" 原因:"&DeleteReason&","&trim(NoteTmp),352
%>
<script language="JavaScript">
	
	alert("儲存完成");
<%if Trim(request("Back_flag"))="0" then%>
	opener.myForm.submit();
<%else%>
	opener.funDbMove('Back');
<%end if%>
	
	
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
				<td colspan="4"><strong>照片無效</strong></td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td width="15%" align="center" height="30">
				無效原因
				</td>
			</tr>
			<tr>
				<td>
				<select name="Reason" onchange="ReasonChange();" >
					<option value="">請選擇</option>
					<option value="車號模糊">車號模糊</option>
					<option value="車體顏色不一致">車體顏色不一致</option>
					<option value="多車入鏡">多車入鏡</option>
					<option value="公務車輛">公務車輛</option>
					<option value="重複舉發">重複舉發</option>
				</select>
				<textarea name="Note" maxlength="50" style="width:400px;height:200px;"></textarea>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
				<input type="button" name="close" value=" 確 定 " onclick="BillDel();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
'						if CheckPermission(237,4)=false then
'							response.write "disabled"
'						end if
						%>>

				<input type="button" name="close" value=" 離 開 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="DBillSN" value="<%=trim(request("DBillSN"))%>">
				<input type="hidden" name="Back_flag" value="<%=Trim(request("Back_flag"))%>">
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
	if (myForm.Note.value=="")
	{
		alert("請先輸入無效原因!");
	}else{
		myForm.kinds.value="DB_BillDel";
		myForm.submit();
	}
}

function ReasonChange(){
	myForm.Note.value=myForm.Reason.value;
}
</script>
</html>
