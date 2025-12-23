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
<%
'檢查是否可進入本系統
'AuthorityCheck(237)

DelMemID=trim(Session("User_ID"))
theBillSN=trim(request("DBillSN"))
'theDelType=trim(request("DelType"))	'單筆或多筆刪除

	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing
	
	'抓單號
	theBillNO=""
	theCarNO=""
	strbillno="select BillNo,CarNo,BillStatus from BillBase where SN="&theBillSN
	set rsBillno=conn.execute(strbillno)
	if not rsBillno.eof then
		theBillNO=trim(rsBillno("BillNo"))
		theCarNO=trim(rsBillno("CarNo"))
		theBillStatus=trim(rsBillno("BillStatus"))
	end if
	rsBillno.close
	set rsBillno=nothing
if trim(request("kinds"))="DB_BillDel2" then
		BillDelFlag="N"

		strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			theBatchTime=(year(now)-1911)&"E"&trim(rsSN("SN"))
		end if
		rsSN.close
		set rsSN=nothing
		NoteTmp=replace(trim(request("Note")),",","")
		strDelS="select SN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate,RecordMemberID,RecordStateID,BillStatus from BillBase where SN="&theBillSN
		set rsDelS=conn.execute(strDelS)
		if not rsDelS.eof then
			if trim(rsDelS("RecordStateID"))<>"-1" then
				CaseInStatus=0	'入案失敗

				BillDelFlag=funcDCIdel(conn,trim(rsDelS("SN")),trim(rsDelS("BillNo")),trim(rsDelS("BillTypeID")),trim(rsDelS("CarNo")),trim(rsDelS("BillUnitID")),trim(rsDelS("BillStatus")),trim(rsDelS("RecordDate")),trim(rsDelS("RecordMemberID")),trim(request("DelReason")),NoteTmp,CaseInStatus,theBatchTime)
			end if
		end if
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
			ConnExecute "舉發單刪除 單號:"&theBillNO&" 車號:"&theCarNO&" 原因:"&DeleteReason&","&trim(NoteTmp)&","&CaseInStatus,352
%>
		<script language="JavaScript">
			alert("刪除成功!");
			opener.myForm.submit();
			window.close();
		</script>
<%
		end if



	
end if
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
		strDelS="select SN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate,RecordMemberID,RecordStateID,BillStatus from BillBase where SN="&theBillSN
		set rsDelS=conn.execute(strDelS)
		if not rsDelS.eof then
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
			end if
		end if
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
		ConnExecute "舉發單刪除 單號:"&theBillNO&" 車號:"&theCarNO&" 原因:"&DeleteReason&","&trim(NoteTmp)&","&CaseInStatus,352
%>
		<script language="JavaScript">
			alert("刪除成功!");
			opener.myForm.submit();
			window.close();
		</script>
<%
		end if


end if
	CaseInFlagForDel="N"
	CaseInDate=""
	strCType="select a.DciCaseInDate from BillBaseDCIReturn a where a.BillNo='"&trim(theBillNO)&"' and a.CarNo='"&trim(theCarNO)&"' and ExchangeTypeID='W' and Status in ('Y','S','n','L')"
	set rsCType=conn.execute(strCType)
	if not rsCType.eof then
		'response.write trim(rsCType("DciCaseInDate"))
		if len(trim(rsCType("DciCaseInDate")))=8 then
			CaseInDate=gOutDT((left(trim(rsCType("DciCaseInDate")),4)-1911) & right(trim(rsCType("DciCaseInDate")),4))
		else
			CaseInDate=gOutDT(trim(rsCType("DciCaseInDate")))
		end if 
		CaseInFlagForDel="Y"
	end if
	rsCType.close
	set rsCType=Nothing
	If sys_City="台南市" Then
		CaseInFlagForDel="Y"
	End If 
	'計算入案幾天
	CountCaseIN=0
	if CaseInDate<>"" then
		CountCaseIN=dateDiff("d",CaseInDate,now)
	end if
'response.write CaseInDate & "vvv" & CountCaseIN
	CaseNotReturn=0
	'檢查是不是有上傳未回傳的
	strCType2="select DciReturnStatusID from DciLog a where a.BillSn='"&trim(theBillSN)&"' and ExchangeTypeID='W' Order by ExchangeDate Desc"
	set rsCType2=conn.execute(strCType2)
	if not rsCType2.eof then
		if isnull(rsCType2("DciReturnStatusID")) or trim(rsCType2("DciReturnStatusID"))="" then
			CaseNotReturn=1
		end if
	end if
	rsCType2.close
	set rsCType2=Nothing
	
	CaseDeleteError=0
	CaseDeleteBatch=""
	'檢查是不是有刪除異常的 
'	strCType2="select * from DciLog a where a.BillSn='"&trim(theBillSN)&"' and ExchangeTypeID='E' Order by ExchangeDate Desc"
'	set rsCType2=conn.execute(strCType2)
'	if not rsCType2.eof then
'		if trim(rsCType2("DciReturnStatusID"))<>"S" then
'			CaseDeleteError=1
'			CaseDeleteBatch=trim(rsCType2("Batchnumber"))
'		end if
'	end if
'	rsCType2.close
'	set rsCType2=Nothing
	
%>

<style type="text/css">
<!--
.style1 {
	color: #FF0000;
	font-weight: bold;
}
.style2 {
	color: #FF0000;
	font-weight: bold;
	line-height:23px; 
	font-size: 18px
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4"><strong>舉發單刪除</strong>
				<br/>
				<strong>單號：<font style="font-size: 38px;line-height:46px "><%=theBillNO%></font></strong>&nbsp; &nbsp; &nbsp; 
				<strong>車號：<font style="font-size: 38px;line-height:46px "><%=theCarNO%></font></strong>
				</td>
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
				<span style="color: #9900FF;font-size: 20px;line-height:23px; "><strong>(刪除案件前，請勿修改單號、車號)</strong></span>
				<br>
		<%
	If CaseDeleteError=1 Then
	%>
		<span class="style1">( 此舉發單已上傳刪除，監理站回傳異常，<br>請至上傳下載資料查詢系統確認，<br>批號：<%=CaseDeleteBatch%> )</span><br>
	<%
	elseif CaseNotReturn=1 Then
		%>
		<span class="style1">( 此舉發單已上傳入案，監理站尚未回傳，<br>請等待資料回傳後再做刪除 )</span><br>
	<%
	else
		if  CountCaseIN>7 then%>
				<input type="button" name="close" value=" 確 定 " onclick="if(confirm('是否確定要刪除此舉發單?')){BillDel2();}" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
						if CheckPermission(234,4)=false then
							response.write "disabled"
						end if
						%>>
		<%else%>
				<input type="button" name="close" value=" 確 定 " onclick="if(confirm('是否確定要刪除此舉發單?')){BillDel();}" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
						if CheckPermission(234,4)=false or CountCaseIN>7 then
							response.write "disabled"
						end if
						%>>
		<%end if
	
	end if
		%>
				<input type="button" name="close" value=" 離 開 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="DBillSN" value="<%=trim(request("DBillSN"))%>">
				<input type="hidden" name="DelType" value="<%=trim(request("DelType"))%>">
				<input type="hidden" name="CaseInFlag" value="<%=trim(CaseInFlagForDel)%>">
		<%
	if CaseNotReturn=0 and theBillStatus>0 And CaseDeleteError=0 then
		%>
			<br>
			<span class="style2">(請務必確認監理站刪除成功後才可再次上傳入案，避免資料發生錯誤)</span>
			<br>
		<%
		if  CountCaseIN>7 and trim(request("kinds"))="" then%>
			<span class="style1">(此舉發單已入案超過七天，刪除舉發單前請先請監理站手動刪除)</span>
			<script language="JavaScript">
				alert("此舉發單已入案超過七天，刪除舉發單前請先請監理站手動刪除!");
			</script>
		<%end if
	end if		
		%>
			<br>
				<span class="style2"><%
			If sys_City="台南市" Then
				response.write "請輸入刪除原因"
			Else
				response.write "未入案案件可不填刪除原因"
			End If 
				%></span>
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
// && myForm.CaseInFlag.value=="Y"
function BillDel(){
	if (myForm.DelReason.value=="" && myForm.CaseInFlag.value=="Y"){
		alert("請選擇刪除舉發單原因！");
	}else{
		if(confirm('請再次確認是否要刪除本案件，刪除後無法復原。')){
			myForm.kinds.value="DB_BillDel";
			myForm.submit();
		}
		
	}
}
function BillDel2(){
	if (myForm.DelReason.value=="" && myForm.CaseInFlag.value=="Y"){
		alert("請選擇刪除舉發單原因！");
	}else{
		if(confirm('請再次確認是否要刪除本案件，刪除後無法復原。')){
			myForm.kinds.value="DB_BillDel2";
			myForm.submit();
		}
		
	}
}
</script>
</html>
