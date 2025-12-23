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
<title>撤銷送達</title>
<%
'檢查是否可進入本系統
'AuthorityCheck(237)

DelMemID=trim(Session("User_ID"))
theBillSN=trim(request("DBillSN"))

if trim(request("kinds"))="DB_BillDel" then
		strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			theBatchTime=(year(now)-1911)&"N"&trim(rsSN("SN"))
		end if
		rsSN.close
		set rsSN=nothing

		strDelS="select SN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate,RecordMemberID,RecordStateID,BillStatus from BillBase where SN="&theBillSN
		set rsDelS=conn.execute(strDelS)
		if not rsDelS.eof then
			if trim(rsDelS("RecordStateID"))<>"-1" then
				
				funcStoreAndSendToGov conn,trim(rsDelS("SN")),trim(rsDelS("BillNo")),trim(rsDelS("BillTypeID")),trim(rsDelS("CarNo")),trim(rsDelS("BillUnitID")),trim(rsDelS("RecordDate")),trim(Session("User_ID")),theBatchTime
			end if
		end if
		rsDelS.close
		set rsDelS=nothing
%>
<script language="JavaScript">
	alert("撤銷送達處理完成，請等待監理所回傳!");
	window.close();
</script>
<%
end if

%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="3">
			<tr bgcolor="#FFCC33">
				<td colspan="4"><strong>撤銷送達</strong></td>
			</tr>
			<tr bgcolor="#FFffff">
				<td colspan="4"><%
			'檢查是否已經有上傳撤銷，dci還沒回傳
			DciUploadFlag=0
			strChk1="select Count(*) as Cnt from DciLog where ExchangeTypeID='N' and ReturnMarkType='Y' and DciReturnStatusID is null and BillSN="&theBillSN
			set rsChk1=conn.execute(strChk1)
			if not rsChk1.eof then
				DciUploadFlag=trim(rsChk1("Cnt"))
			end if
			rsChk1.close
			set rsChk1=nothing

			'檢查最後一次上傳是否是撤銷
			DciCancelStatusFlag=0
			strChk2="select * from DciLog where BillSN="&theBillSN&" order by exchangeDate desc"
			set rsChk2=conn.execute(strChk2)
			if not rsChk2.eof then
				if trim(rsChk2("ExchangeTypeID"))="N" and trim(rsChk2("ReturnMarkType"))="Y" and trim(rsChk2("DciReturnStatusID"))="S" then
					DciCancelStatusFlag=1
				else
					DciCancelStatusFlag=0
				end if
			end if
			rsChk2.close
			set rsChk2=nothing

			if DciUploadFlag=0 then
				if DciCancelStatusFlag=0 then
					response.write "是否要撤銷送達?"
				else
					response.write "此舉發單已經上傳撤銷送達。"
				end if
			else
				response.write "此舉發單已經上傳撤銷送達至監理所，請等待監理所回傳。"
			end if
				
				%></td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
				<input type="button" name="close" value=" 確 定 " onclick="BillDel();" <%
				if DciUploadFlag<>0 or DciCancelStatusFlag<>0 then
					response.write "disabled"
				end if
				%>>
				<input type="button" name="close" value=" 離 開 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="DBillSN" value="<%=trim(request("DBillSN"))%>">
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
	//if (myForm.DelReason.value=="" && myForm.CaseInFlag.value=="Y"){
	//	alert("請選擇刪除舉發單原因！");
	//}else{
		myForm.kinds.value="DB_BillDel";
		myForm.submit();
	//}
}
</script>
</html>
