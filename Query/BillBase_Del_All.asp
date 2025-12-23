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
<title>舉發單刪除</title>
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
		BillDelFlag="N"
		strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			theBatchTime=(year(now)-1911)&"E"&trim(rsSN("SN"))
		end if
		rsSN.close
		set rsSN=nothing
		NoteTmp=replace(trim(request("Note")),",","")
		strDelS="select a.SN,a.BillNo,a.BillTypeID,a.CarNo,a.BillUnitID,a.RecordDate,a.RecordMemberID,a.RecordStateID,a.BillStatus from BillBase a,DciLog b where b.BatchNumber='"&theBatchNumber&"' and a.SN=b.BillSN"
		set rsDelS=conn.execute(strDelS)
		If Not rsDelS.Bof Then rsDelS.MoveFirst 
		While Not rsDelS.Eof
			if trim(rsDelS("RecordStateID"))<>"-1" then
				CaseInStatus=0	'入案是否成功
				strChkCaseIn="select count(*) as cnt from DciLog where BillSn="&trim(rsDelS("SN"))&" and ExchangeTypeID='W' and (DciReturnStatusID in ('Y','S','n','L') or DciReturnStatusID is null)"
				set rsChk=conn.execute(strChkCaseIn)
				if not rsChk.eof then
					if trim(rsChk("cnt"))="0" then
						CaseInStatus=0	'入案失敗
					else
						CaseInStatus=1	'入案成功
					end if
				end if
				rsChk.close
				set rsChk=nothing
				BillDelFlag=funcDCIdel(conn,trim(rsDelS("SN")),trim(rsDelS("BillNo")),trim(rsDelS("BillTypeID")),trim(rsDelS("CarNo")),trim(rsDelS("BillUnitID")),trim(rsDelS("BillStatus")),trim(rsDelS("RecordDate")),trim(rsDelS("RecordMemberID")),trim(request("DelReason")),NoteTmp,CaseInStatus,theBatchTime)

				'寫入LOG
				DeleteReason=""
				if trim(request("DelReason"))<>"" then
					strRea="select ID,Content from DCICode where TypeID=3 and ID='"&trim(request("DelReason"))&"'"
					set rsRea=conn.execute(strRea)
					if not rsRea.eof then
						DeleteReason=trim(rsRea("Content"))
					end if
					rsRea.close
					set rsRea=nothing
				else
					DeleteReason="無"
				end if
				ConnExecute "舉發單刪除 單號:"&rsDelS("BillNo")&" 車號:"&rsDelS("CarNo")&" 原因:"&DeleteReason&","&trim(NoteTmp)&","&CaseInStatus,352
			end if
		rsDelS.MoveNext
		Wend
		rsDelS.close
		set rsDelS=nothing

		if BillDelFlag="N" then
%>
		<script language="JavaScript">
			alert("本舉發單已向監理站進行刪除!");
			opener.myForm.submit();
			window.close();
		</script>
<%
		else
%>
		<script language="JavaScript">
			alert("刪除成功，批號：<%=theBatchTime%>");
			opener.myForm.submit();
			window.close();
		</script>
<%
		end if
end if
	CaseInFlagForDel="N"
	CaseInDate=""
	CaseNotReturn=0
	strCType="select a.ExchangeDate,a.DciReturnStatusID from DciLog a where a.BatchNumber='"&trim(request("BatchNumber"))&"'"
	set rsCType=conn.execute(strCType)
	if not rsCType.eof then
		'response.write trim(rsCType("DciCaseInDate"))
		CaseInDate=trim(rsCType("ExchangeDate"))
		CaseInFlagForDel="Y"
		if isnull(rsCType("DciReturnStatusID")) or trim(rsCType("DciReturnStatusID"))="" then
			CaseNotReturn=1
		end if
	end if
	rsCType.close
	set rsCType=nothing
	'計算入案幾天
	CountCaseIN=0
	if CaseInDate<>"" then
		CountCaseIN=dateDiff("d",CaseInDate,now)
	end if
	
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4"><strong>舉發單刪除</strong></td>
			</tr>
			<tr>
				<td width="25%" align="right" bgcolor="#EBFBE3">刪除原因</td>
				<td width="75%">
					<select name="DelReason">
						<option value="">請選擇</option>
<%
				strR="select ID,Content from DCICode where TypeID=3"
				set rsR=conn.execute(strR)
				If Not rsR.Bof Then rsR.MoveFirst 
				While Not rsR.Eof
%>
						<option value="<%=trim(rsR("ID"))%>"><%=trim(rsR("Content"))%></option>
<%
				rsR.MoveNext
				Wend
				rsR.close
				set rsR=nothing
%>
					</select>
				</td>
			</tr>
			<tr>
				<td width="25%" align="right" bgcolor="#EBFBE3">備註</td>
				<td width="75%">
					<input type="text" size="30" name="Note" value="">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
				<input type="button" name="close" value=" 確 定 " onclick="BillDel();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
						if CheckPermission(234,4)=false or CountCaseIN>7 or CaseNotReturn=1 then
							response.write "disabled"
						end if
						%>>
				<input type="button" name="close" value=" 離 開 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="BatchNumber" value="<%=trim(request("BatchNumber"))%>">
				<input type="hidden" name="DelType" value="<%=trim(request("DelType"))%>">
				<input type="hidden" name="CaseInFlag" value="<%=trim(CaseInFlagForDel)%>">
				<%
				if CountCaseIN>7 then
					response.write "<br><font color=""#FF0000"" size=""2"">入案已超過七天</font>"
				elseif CaseNotReturn=1 then
					response.write "<br><font color=""#FF0000"" size=""2"">( 此舉發單已上傳入案，監理站尚未回傳，<br>請等待資料回傳後再做刪除 )</font>"
				end if
				%>
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
	if (myForm.DelReason.value=="" && myForm.CaseInFlag.value=="Y"){
		alert("請選擇刪除舉發單原因！");
	}else{
		myForm.kinds.value="DB_BillDel";
		myForm.submit();
	}
}
</script>
</html>
