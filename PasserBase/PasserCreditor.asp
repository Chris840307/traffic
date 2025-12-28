<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>債權憑証管理系統</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
.btn3{
   font-size:14px;
   font-family:新細明體;
   background-color:#EEEEEE;
   border-style:solid;
}
</style>
</head>
<%
	Sys_SendOpenGovNumber="":Sys_SendNumber="":Sys_SendDate=""
	Sys_PetitionDate="":Sys_OpenGovNumber="":Sys_CreditorNumber="":sys_CreditorTypeID=""
	Sys_RemainNT=""

	if trim(request("DB_Add"))="Del" then
		If not ifnull(Request("CreditorSN")) Then
			strSQL="delete from PasserCreditor where sn="&trim(Request("CreditorSN"))
			conn.execute(strSQL)
		End if
		
		'strSQL="delete from PasserSendDetail where sn="&trim(Request("SendSN"))
		'conn.execute(strSQL)

		Response.write "<script>"
		Response.Write "alert('刪除完成！');"
		Response.write "</script>"

	elseif trim(request("DB_Add"))="ADD" then

		strSQL="select count(1) cnt from PasserSendDetail where BillSN="&trim(request("BillSN"))&" and OpenGovNumber='"&trim(request("Sys_SendOpenGovNumber"))&"' and SendNumber='"&trim(request("Sys_SendNumber"))&"' and Not Exists(select 'Y' from PasserCreditor where SendDetailSN=PasserSendDetail.SN)"
		set rs=conn.execute(strSQL)

		
		dbSN="(select sn from PasserSendDetail where BillSN="&trim(request("BillSN"))&" and OpenGovNumber='"&trim(request("Sys_SendOpenGovNumber"))&"' and SendNumber='"&trim(request("Sys_SendNumber"))&"' and Not Exists(select 'Y' from PasserCreditor where SendDetailSN=PasserSendDetail.SN))"

		If cdbl(rs("cnt"))=0 Then
			strSQL="insert into PasserSendDetail values((select nvl(max(sn),0)+1 from PasserSendDetail),"&trim(request("BillSN"))&",'"&trim(request("Sys_SendOpenGovNumber"))&"','"&trim(request("Sys_SendNumber"))&"',"&funGetDate(gOutDT(request("Sys_SendDate")),0)&",sysdate,"&Session("User_ID")&")"

			conn.execute(strSQL)

			dbSN="(select nvl(max(sn),0) from PasserSendDetail)"
		End if 

		rs.close

		If not ifnull(Request("Sys_PetitionDate")) Then
			strSQL="insert into PasserCreditor(SN,BILLSN,SENDDETAILSN,OPENGOVNUMBER,CREDITORNUMBER,PETITIONDATE,CREDITORTYPEID,REMAINNT,RECORDDATE,RECORDMEMBERID,InventoryNO) values((select nvl(max(sn),0)+1 from PasserCreditor),"&trim(request("BillSN"))&","&dbSN&",'"&trim(Request("Sys_OpenGovNumber"))&"','"&trim(Request("Sys_CreditorNumber"))&"',"&funGetDate(gOutDT(request("Sys_PetitionDate")),0)&",'"&trim(Request("sys_CreditorTypeID"))&"',"&funTnumber(request("Sys_RemainNT"))&",sysdate,"&Session("User_ID")&",'"&trim(Request("sys_InventoryNO"))&"')"

			conn.execute(strSQL)
		End if

		Response.write "<script>"
		Response.Write "alert('儲存完成！');"
		Response.write "</script>"


	elseif trim(request("DB_Add"))="Update" then

		strSQL="update PasserSendDetail set OpenGovNumber='"&trim(request("Sys_SendOpenGovNumber"))&"',SendNumber='"&trim(request("Sys_SendNumber"))&"',SendDate="&funGetDate(gOutDT(request("Sys_SendDate")),0)&",RecordDate=sysdate,recordMemberID="&Session("User_ID")&" where sn="&trim(Request("SendSN"))

		conn.execute(strSQL)

		If not ifnull(Request("CreditorSN")) Then
			strSQL="update PasserCreditor set OpenGovNumber='"&trim(Request("Sys_OpenGovNumber"))&"',CreditorNumber='"&trim(Request("Sys_CreditorNumber"))&"',PetitionDate="&funGetDate(gOutDT(request("Sys_PetitionDate")),0)&",CreditorTypeID='"&trim(Request("sys_CreditorTypeID"))&"',RemainNT="&funTnumber(request("Sys_RemainNT"))&",RecordDate=sysdate,RecordMemberID="&Session("User_ID")&",InventoryNO='"&trim(Request("sys_InventoryNO"))&"' where sn="&trim(Request("CreditorSN"))&" and SendDetailSN="&trim(Request("SendSn"))

			conn.execute(strSQL)

		elseIf not ifnull(Request("Sys_PetitionDate")) Then
			strSQL="insert into PasserCreditor(SN,BILLSN,SENDDETAILSN,OPENGOVNUMBER,CREDITORNUMBER,PETITIONDATE,CREDITORTYPEID,REMAINNT,RECORDDATE,RECORDMEMBERID,InventoryNO) values((select nvl(max(sn),0)+1 from PasserCreditor),"&trim(request("BillSN"))&","&trim(Request("SendSn"))&",'"&trim(Request("Sys_OpenGovNumber"))&"','"&trim(Request("Sys_CreditorNumber"))&"',"&funGetDate(gOutDT(request("Sys_PetitionDate")),0)&",'"&trim(Request("sys_CreditorTypeID"))&"',"&funTnumber(request("Sys_RemainNT"))&",sysdate,"&Session("User_ID")&",'"&trim(Request("sys_InventoryNO"))&"')"

			conn.execute(strSQL)

		End if

		Response.write "<script>"
		Response.Write "alert('儲存完成！');"
		Response.write "</script>"
	elseif not ifnull(Request("SendSn")) then
		Sys_SendSn=trim(Request("SendSn"))
		Sys_CreditorSN=trim(Request("CreditorSN"))

		strSQL="select * from PasserSendDetail where BillSN="&trim(Request("BillSN"))&" and sn="&Sys_SendSn
		set rsSend=conn.execute(strSQL)
			Sys_SendOpenGovNumber=trim(rsSend("OpenGovNumber"))
			Sys_SendNumber=trim(rsSend("SendNumber"))
			Sys_SendDate=trim(rsSend("SendDate"))
		rsSend.close

		Sys_PetitionDate="":Sys_OpenGovNumber="":Sys_CreditorNumber=""
		sys_CreditorTypeID="":	Sys_RemainNT="":Sys_InventoryNO=""

		If not ifnull(Sys_CreditorSN) Then
			strSQL="select * from PasserCreditor where SendDetailSN="&Sys_SendSn&" and sn="&Sys_CreditorSN
			set rsSend=conn.execute(strSQL)
				Sys_PetitionDate=trim(rsSend("PetitionDate"))
				Sys_OpenGovNumber=trim(rsSend("OpenGovNumber"))
				Sys_CreditorNumber=trim(rsSend("CreditorNumber"))
				sys_CreditorTypeID=trim(rsSend("CreditorTypeID"))
				Sys_RemainNT=trim(rsSend("RemainNT"))
				Sys_InventoryNO=trim(rsSend("InventoryNO"))
			rsSend.close
		End if
	
	elseif ifnull(Request("SendSn")) then

		strSQL="select count(1) cnt from PasserSendDetail where BillSN="&trim(Request("BillSN"))

		set rscnt=conn.execute(strSQL)

		If cdbl(rscnt("cnt"))=0 Then
			strSQL="select OpenGovNumber,SendNumber,SendDate from PasserSend where billsn="&trim(Request("BillSN"))
			set rssend=conn.execute(strSQL)

			If not rssend.eof Then

				strSQL="insert into PasserSendDetail values((select nvl(max(sn),0)+1 from PasserSendDetail),"&trim(Request("BillSN"))&",'"&trim(rssend("OpenGovNumber"))&"','"&trim(rssend("SendNumber"))&"',"&funGetDate(rssend("SendDate"),0)&",sysdate,"&Session("User_ID")&")"

				conn.execute(strSQL)
			End if
			rssend.close
		End if
		rscnt.close

		strSQL="update PasserSendDetail set RECORDMEMBERID="&Session("User_ID")&",RECORDDATE=sysdate where billsn="&trim(Request("BillSN"))&" and Not Exists(select 'Y' from PasserCreditor where SendDetailSN=PasserSendDetail.SN)"

		conn.execute(strSQL)

	end if
	
%>
<body>
<form name="myForm" method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size="4">債權憑証管理系統</font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td align="right" width="80" bgcolor="#FFFF99">舉發單號</td>
					<td colspan="5"><%
						strSQL="select BillNo from passerBase where sn="&trim(Request("BillSN"))
						set rs=conn.execute(strSQL)
						Sys_BillNo=trim(rs("BillNo"))
						Response.Write rs("BillNo")
						rs.close
					%></td>
				</tr>
				<tr>
					<td align="center" colspan="6" bgcolor="FFCC33"><B>移送資料</B></td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font color="red">*</font>發文文號</td>
					<td><input name="Sys_SendOpenGovNumber" class="btn1" type="text" size="12" maxlength="30" value="<%=Sys_SendOpenGovNumber%>"></td>
					<td align="right" nowrap bgcolor="#FFFF99"><font color="red">*</font>移送案號</td>
					<td><input name="Sys_SendNumber" value="<%=Sys_SendNumber%>" class="btn1" type="text" size="12" maxlength="30"></td>
					<td align="right" nowrap bgcolor="#FFFF99"><font color="red">*</font>移送日期</td>
					<td>
						<input name="Sys_SendDate" value="<%
							if not ifnull(Sys_SendDate) then
								response.write gInitDT(Sys_SendDate)
							else
								response.write gInitDT(date)
							end if%>" class="btn1" type="text" size="4" maxlength="10" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" class="btn3" style="width:20px;height:25px;" value="..." onclick="OpenWindow('Sys_SendDate');">
					</td>
				</tr>
				<tr>
					<td align="center" colspan="6" bgcolor="FFCC33"><B>債權憑証資料</b></td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99">取得時間</td>
					<td><input name="Sys_PetitionDate" value="<%=gInitDT(Sys_PetitionDate)%>" class="btn1" type="text" size="4" maxlength="10" onkeyup="chknumber(this);">
						<input type="button" name="datestr" class="btn3" style="width:20px;height:25px;" value="..." onclick="OpenWindow('Sys_PetitionDate');">
					</td>
					<td align="right" nowrap bgcolor="#FFFF99">執行憑証編號</td>
					<td><input name="Sys_OpenGovNumber" class="btn1" type="text" size="12" maxlength="15" value="<%=Sys_OpenGovNumber%>"></td>
					<td align="right" nowrap bgcolor="#FFFF99">收文文號</td>
					<td><input name="Sys_CreditorNumber" class="btn1" type="text" size="12" value="<%=Sys_CreditorNumber%>"></td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99">執行狀態<br>查詢結果</td>
					<td><select name="sys_CreditorTypeID">

							<option value="1"<%if trim(sys_CreditorTypeID)="1" then response.write " Selected"%>>無個人財產</option>

							<option value="0"<%if trim(sys_CreditorTypeID)="0" then response.write " Selected"%>>清償中</option>

						</select>
					</td>
					<td align="right" nowrap bgcolor="#FFFF99">待執行金額</td>
					<td><input name="Sys_RemainNT" class="btn1" type="text" size="12" maxlength="12" value="<%
						if not ifnull(Sys_RemainNT) then
							response.write Sys_RemainNT
						else
							strSQL="select nvl(ForFeit1,0)+nvl(ForFeit2,0) ForFeit from PasserBase where sn="&trim(request("BillSN"))
							set rspay=conn.execute(strSQL)
							Sys_ForFeit=cdbl(rspay("ForFeit"))
							rspay.close
							
							Sys_PaySum=0
							strSQL="select nvl(sum(PayAmount),0) as PaySum from PasserPay where BillSN="&trim(request("BillSN"))
							set rspay=conn.execute(strSQL)
							if not rspay.eof then Sys_PaySum=cdbl(rspay("PaySum"))
							rspay.close

							response.write Sys_ForFeit-Sys_PaySum
						end if						
					%>"></td>

					<td align="right" nowrap bgcolor="#FFFF99">保管品<br>收據編號</td>
					<td><input name="Sys_InventoryNO" class="btn1" type="text" size="12" maxlength="15" value="<%=Sys_InventoryNO%>"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#FFDD77">
		<!--
			<input name="btnadd" type="button" value="新增" class="btn3" style="width:40px;height:20px;" onclick="funAdd();"<%If not ifnull(Sys_SendSn) Then Response.Write " disabled"%>>
		-->
			<input type="button" name="Submit" value="儲存" class="btn3" style="width:40px;height:20px;" onclick="funEdit();"<%if ifnull(Sys_SendSn) then Response.Write " disabled"%>>

			<input type="button" name="Submit2" class="btn3" style="width:40px;height:20px;" value="關閉" onclick="opener.myForm.submit();self.close();">
			
		</td>
	</tr>
</table>
<hr>
<table width="100%" height="100%" border="0" bgcolor="#E0E0E0">
		<tr>
			<td colspan="10" bgcolor="#FFCC33">歷次債權記錄</td>
		</tr>
		<tr bgcolor="#EBFBE3">
			<td width="10%" nowrap>移送日期</td>
			<td width="10%" nowrap>發文文號</td>
			<td width="10%" nowrap>移送案號</td>
			<td width="10%" nowrap>取得時間</td>
			<td width="10%" nowrap>收文文號</td>
			<td width="10%" align="center" nowrap>執行<br>憑証編號</td>
			<td width="10%" nowrap>查詢結果</td>
			<td width="10%" align="center" nowrap>待執行<br>金額</td>
			<td width="10%" align="center" nowrap>保管品<br>收據編號</td>
			<td width="8%">操作</td>
		</tr><%
		strSql="select * from (select SN SendDetialSN,SendDate,OpenGovNumber SendGovNumber,SendNumber,RecordMemberID from PasserSendDetail where BillSN="&trim(request("BillSN"))&") a,(select sn,SendDetailSN,OpenGovNumber CreditorGovNumber,CreditorNumber,PetitionDate,Decode(CreditorTypeID,1,'無個人財產','清償中') CreditorTypeName,RemainNT,Imagefilename,Imagefilename2,Imagefilename3,Imagefilename4,InventoryNO from PasserCreditor where BillSN="&trim(request("BillSN"))&")b where a.SendDetialSN=b.SendDetailSN(+) order by SendDate DESC,PetitionDate DESC"
		set rs=conn.execute(strSQL)
		If not rs.eof Then
			While Not rs.eof
				response.write "<tr align='center' bgcolor='#FFFFFF'"
				lightbarstyle 0
				response.write ">"

				response.write "<td class=""font10"">"&gInitDT(trim(rs("SendDate")))&"</td>"
				response.write "<td class=""font10"">"&trim(rs("SendGovNumber"))&"</td>"
				response.write "<td class=""font10"">"&trim(rs("SendNumber"))&"</td>"
				response.write "<td class=""font10"">"&gInitDT(trim(rs("PetitionDate")))&"</td>"
				response.write "<td class=""font10"">"&trim(rs("CreditorNumber"))&"</td>"
				response.write "<td class=""font10"">"&trim(rs("CreditorGovNumber"))&"</td>"
				response.write "<td class=""font10"">"&trim(rs("CreditorTypeName"))&"</td>"
				response.write "<td class=""font10"">"&trim(rs("RemainNT"))&"</td>"
				response.write "<td class=""font10"">"&trim(rs("InventoryNO"))&"</td>"
				response.write "<td>"

				response.write "<input type=""button"" class=""btn3"" style=""width:40px;height:20px;"" value=""修改"" onclick=""funLoadEdit('"&trim(rs("SendDetialSN"))&"','"&trim(rs("SN"))&"');"">&nbsp;"

				response.write "<input type=""button"" class=""btn3"" style=""width:40px;height:20px;"" value=""刪除"" onclick=""funDel('"&trim(rs("SendDetialSN"))&"','"&trim(rs("SN"))&"');"">&nbsp;"

				response.write "<input type=""button"" name=""btnMap"" class=""btn3"" style=""width:70px;height:20px;"" value=""債權上傳"" onclick=""funMap('"&trim(rs("SN"))&"');"">"

	
				If not ifnull(rs("Imagefilename")) Then
					Response.Write "<a href=""./Picture/"&trim(rs("Imagefilename"))&""" target=""_blank"">債權影像1</a>"
				End If 
				
				If not ifnull(rs("Imagefilename2")) Then
					Response.Write "<br><a href=""./Picture/"&trim(rs("Imagefilename2"))&""" target=""_blank"">債權影像2</a>"
				End If 

				If not ifnull(rs("Imagefilename3")) Then
					Response.Write "<br><a href=""./Picture/"&trim(rs("Imagefilename3"))&""" target=""_blank"">債權影像3</a>"
				End if 

				If not ifnull(rs("Imagefilename4")) Then
					Response.Write "<br><a href=""./Picture/"&trim(rs("Imagefilename4"))&""" target=""_blank"">債權影像4</a>"
				End if 
				
				

				Response.Write "</td></tr>"
				rs.movenext
			Wend

		End if
		rs.close%>
</table>
	<input type="Hidden" name="BilSN" value="<%=request("BilSN")%>">
	<input type="Hidden" name="SendSn" value="<%=Sys_SendSn%>">
	<input type="Hidden" name="CreditorSN" value="<%=Sys_CreditorSN%>">
	<input type="Hidden" name="DB_Add" value="">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funAdd(){
	var err=0;
	if(myForm.Sys_SendDate.value==''){
		err=1;
		alert("移送日期不可空白!!");
	}else if(!dateCheck(myForm.Sys_SendDate.value)){
		err=1;
		alert("移送日期格式錯誤!!");
	}else if(myForm.Sys_SendOpenGovNumber.value==''){
		err=1;
		alert("發文文號不可空白!!");
	}else if(myForm.Sys_SendNumber.value==''){
		err=1;
		alert("移送案號不可空白!!");
	}
	if(err==0){
		myForm.DB_Add.value="ADD";
		myForm.submit();
	}
}
function funEdit(){
	var err=0;
	if(myForm.Sys_SendDate.value==''){
		err=1;
		alert("移送日期不可空白!!");
	}else if(!dateCheck(myForm.Sys_SendDate.value)){
		err=1;
		alert("移送日期格式錯誤!!");
	}else if(myForm.Sys_SendOpenGovNumber.value==''){
		err=1;
		alert("發文文號不可空白!!");
	}else if(myForm.Sys_SendNumber.value==''){
		err=1;
		alert("移送案號不可空白!!");
	}

	if(myForm.Sys_OpenGovNumber.value!=''){
		if(myForm.Sys_PetitionDate.value==''){
			err=1;
			alert("執行時間不可空白!!");
		}
	}

	if(err==0){
		myForm.DB_Add.value="Update";
		myForm.submit();
	}
}

function funLoadEdit(SendSn,CreditorSN){
	myForm.SendSn.value=SendSn;
	myForm.CreditorSN.value=CreditorSN;
	myForm.DB_Add.value="";
	myForm.submit();
}
function funfirst(){
	myForm.CreditorSN.value="";
	myForm.DB_Add.value="";
	myForm.submit();
}
function funDel(SendSn,CreditorSN){
	myForm.SendSn.value=SendSn;
	myForm.CreditorSN.value=CreditorSN;
	myForm.DB_Add.value="Del";
	myForm.submit();
}
function funMap(SN){
	UrlStr="SendStyle_Creditor.asp?SN="+SN;
	newWin(UrlStr,"winMap",700,550,50,10,"yes","yes","yes","no");
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	PasserWin=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	PasserWin.focus();
	return win;
}
</script>
<%conn.close%>