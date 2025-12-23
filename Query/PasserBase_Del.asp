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

if trim(request("kinds"))="DB_BillDel" then
		NoteTmp=replace(trim(request("Note")),",","")
		strDel="delete from PasserDeleteReason where PasserSN="&theBillSN
		conn.execute strDel
		
		strUpd="update PasserBase set BillStatus='6',RecordStateID=-1,DelMemberID="&DelMemID&" where SN="&theBillSN
		conn.execute strUpd

		strIns="Insert into PasserDeleteReason(PasserSN,DelDate,DelReason,Note)" &_
			" values("&theBillSN&",sysdate,'"&trim(request("DelReason"))&"','"&trim(NoteTmp)&"')"
		conn.execute strIns
		
		'寫入LOG
		DeleteReason=""
		strRea="select ID,Content from DCICode where TypeID=3 and ID='"&trim(request("DelReason"))&"'"
		set rsRea=conn.execute(strRea)
		if not rsRea.eof then
			DeleteReason=trim(rsRea("Content"))
		end if
		rsRea.close
		set rsRea=nothing
		'抓單號
		theBillNo="":theCarNo="":theCarSimpleID="":theBillUnitID="":theBillTypeID=""
		strbillno="select BillNo,CarNo,BillTypeID,CarSimpleID,BillUnitID from PasserBase where SN="&theBillSN
		set rsBillno=conn.execute(strbillno)
		if not rsBillno.eof then
			theBillNo=trim(rsBillno("BillNo"))
			theCarNo=trim(rsBillno("CarNo"))
			theBillTypeID=trim(rsBillno("BillTypeID"))
			theCarSimpleID=trim(rsBillno("CarSimpleID"))
			theBillUnitID=trim(rsBillno("BillUnitID"))
		end if
		rsBillno.close
		set rsBillno=nothing
		ConnExecute "舉發單刪除 單號:"&theBillNo&" 原因:"&DeleteReason&","&trim(NoteTmp),352

		chkcnt=0
		strSQL="select count(1) cnt  " & _
			"from PasserBase a where exists(select 'Y' from PasserDCILog where billsn=a.sn and exchangetypeid='W' and dcireturnstatusid in(select dcireturn from dcireturnstatus where dciactionid like 'W%' and dcireturnstatus=1))" & _
			" and a.SN="&theBillSN
		set rscnt=conn.execute(strSQL)

		chkcnt=cdbl(rscnt("cnt"))
		rscnt.close

		If theCarSimpleID="8" and chkcnt>0 Then

			theBatchTime="":theDCISN=""

			strSN="select PASSERDCILOGBATCHNUMBER.nextval as SN from Dual"
			set rsSN=conn.execute(strSN)
			if not rsSN.eof then
				theBatchTime=(year(now)-1911)&"E"&trim(rsSN("SN"))
			end if
			rsSN.close
			set rsSN=nothing	

			strSN="select passerDCILOG_SEQ.nextval as SN from Dual"
			set rsSN=conn.execute(strSN)
			if not rsSN.eof then
				theDCISN=trim(rsSN("SN"))
			end if
			rsSN.close
			set rsSN=nothing	

				
			strInsCaseIn="insert into PASSERDCILOG(" & _
				"SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate" & _
				",RecordMemberID,ExchangeDate,ReturnMarkType,ExchangeTypeID,BatchNumber,DciUnitID)"&_
				"values("&theDCISN&","&theBillSN&",'"&theBillNo&"'"&_
				","&theBillTypeID&",'"&theCarNo&"','"&theBillUnitID&"'"&_
				",sysdate,"&Session("User_ID")&",sysdate,6,'E','"&theBatchTime&"'" &_
				",(" &_
					"select DciUnitID from UnitInfo ut where UnitID=(" &_
						"select unittypeid from UnitInfo uta where unitid='"&theBillUnitID&"'" &_
					")" &_
				")" &_
			")" 

			conn.execute strInsCaseIn

			strInsCaseIn="insert into PasserBaseDciReturn(" & _
				"DciLogSN,BillSN,BillNO,CarNo,ExchangeTypeID)"&_
				"values(" & theDCISN & ","&theBillSN&",'"&theBillNo&"','"&theCarNo&"','E')" 

			conn.execute strInsCaseIn

			sqlpasserbase="update PasserBase set DCILOGSN="&theDCISN&" where sn="&theBillSN

			conn.execute sqlpasserbase


			Response.write "<script>"
			Response.Write "alert(""刪除上傳，批號："&theBatchTime&""");"
			Response.Write "opener.myForm.submit();"
			Response.Write "window.close();"
			Response.write "</script>"
		
		else
			
			Response.write "<script>"
			Response.Write "alert(""刪除完成！"");"
			Response.Write "opener.myForm.submit();"
			Response.Write "window.close();"
			Response.write "</script>"
		End if 

else

	chkcnt=0
	strSQL="select count(1) cnt  " & _
		"from PasserBase a where CarSimpleID=8 and recordstateid=0 and DCILOGSN is not null and a.SN="&theBillSN
	set rscnt=conn.execute(strSQL)

	chkcnt=cdbl(rscnt("cnt"))
	rscnt.close

	If chkcnt > 0 Then
		
		chkcnt=0
		strSQL="select count(1) cnt  " & _
			"from PasserBase a where not exists(select 'N' from PasserDCILog where sn=a.DCILOGSN and dcireturnstatusid in(select dcireturn from dcireturnstatus where dcireturnstatus is not null))" & _
			"and recordstateid=0 and a.SN="&theBillSN
		set rscnt=conn.execute(strSQL)

		chkcnt=cdbl(rscnt("cnt"))
		rscnt.close

		If chkcnt > 0 Then
			Response.Write "該案件上傳的資料尚未回傳，請至慢車下載資料查詢確認。"
			Response.End
		end If 
	End if 
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
						<option>請選擇</option>
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
						if CheckPermission(234,4)=false then
							response.write "disabled"
						end if
						%>>

				<input type="button" name="close" value=" 離 開 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="DBillSN" value="<%=trim(request("DBillSN"))%>">
				<input type="hidden" name="DelType" value="<%=trim(request("DelType"))%>">
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
	if (myForm.DelReason.value==""){
		alert("請選擇刪除舉發單原因！");
	}else{
		myForm.kinds.value="DB_BillDel";
		myForm.submit();
	}
}
</script>
</html>
