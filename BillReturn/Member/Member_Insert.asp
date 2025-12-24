<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="..\Common\DB.ini"-->
<!-- #include file="..\Common\AllFunction.inc"-->
<%
daynow=gInitDT(date)
if request("DB_Add")="ADD" then
	StartJobDate=gOutDT(request("StartJobDate"))
	LeaveJobDate=gOutDT(request("LeaveJobDate"))
	Sys_ManagerPower="0"
	Sys_LoginID="ZZ1"
	if trim(request("Sys_LoginType"))="y" then
		strSQL="select Max(To_number(Substr(LoginID,3))) as LoginID from MemberData where LoginID like 'ZZ%'"
		set rsmax=conn.execute(strSQL)
		if Not isnull(rsmax("LoginID")) then Sys_LoginID="ZZ"&Cint(rsmax("LoginID"))+1
		rsmax.close
	else
		Sys_LoginID=request("Sys_LoginID")
	end if
	
	if trim(request("Sys_ManagerPower"))<>"" then Sys_ManagerPower=trim(request("Sys_ManagerPower"))
	strSQL="INSERT INTO MEMBERDATA(LOGINID,PASSWORD,MEMBERID,PKI,UNITID,CHNAME,JOBID,CREDITID,EMAIL,TELEPHONE,STARTJOBDATE,LEAVEJOBDATE,ACCOUNTSTATEID,MODIFYTIME,RECORDSTATEID,RECORDDATE,RECORDMEMBERID,MONEY,BANKACCOUNT,RoleID,GroupRoleID,ManagerPower,BankName,BankID) VALUES('"&Sys_LoginID&"','"&request("Sys_PassWord")&"',"&funTableSeq("MEMBERDATA","MEMBERID")&", '"&Sys_LoginID&funTableSeq("MEMBERDATA","MEMBERID")&"','"&request("Sys_UnitID")&"','"&request("Sys_ChName")&"',"&request("Sys_JOBID")&",'"&UCase(request("Sys_CREDITID"))&"','"&request("Sys_EMAIL")&"','"&request("Sys_TELEPHONE")&"',"&funGetDate(StartJobDate,0)&","&funGetDate(LeaveJobDate,0)&","&request("Sys_ACCOUNTSTATEID")&","&funGetDate(now,1)&",0,"&funGetDate(now,1)&","&session("User_ID")&","&funTnumber(request("Sys_MONEY"))&",'"&request("Sys_BANKACCOUNT")&"',"&request("Sys_RoleID")&","&request("Sys_GroupRoleID")&",'"&Sys_ManagerPower&"','"&request("Sys_BankName")&"','"&request("Sys_BankID")&"')"
	conn.execute(strSQL)
	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>人員資料新增</title>
<!-- #include file="..\Common\css.txt"-->
</head>
<body onkeydown="KeyDown()">
<form name=myForm method="post">
<table width="100%" height="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33">人員資料新增</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#dddddd">
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99"><font color="red">* </font>員警代號<br><font size="2">建檔時舉發員警代號</font> </td>
					<td>
						<input name="Sys_LoginID" class="btn1" type="text" value="" size="12" maxlength="10">
						<br><input name="Sys_LoginType" class="btn1" type="checkbox" value="y" onclick="funLoginType();">無臂章號碼
					</td>
					<td bgcolor="#FFFF99"><font color="red"> </font>系統登入密碼</td>
					<td>
						<input name="Sys_PassWord" class="btn1" type="text" value="" size="12" maxlength="12"> (請注意鍵盤大小寫)
					</td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#FFFF99"><font color="red">* </font>使用者姓名</td>
					<td>
						<input name="Sys_ChName" class="btn1" type="text" value="" size="12" maxlength="8">
					</td>
					<td bgcolor="#FFFF99"><font color="red"> </font>系統登入帳號</td>
					<td>
						<input name="Sys_CreditID" class="btn1" type="text" value="" size="12" maxlength="12">需要使用本系統才需填寫
					</td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#FFFF99"><font color="red">*</font>隸屬單位</td>
					<td>
						<%=UnSelectUnitOption("Sys_UnitID","")%>
					</td>
					<td bgcolor="#FFFF99">使用者職級</td>
					<td>
						<select name="Sys_JobID" class="btn1">
							<option value="0">請選擇</option><%
							strSQL="select Content,ID from Code where TypeID=4"
							set rs=conn.execute(strSQL)
							while Not rs.eof
								response.write "<option value="""&rs("ID")&""">"
								response.write rs("Content")
								response.write "</option>"
								rs.movenext
							wend
							rs.close%>
						</select>
					</td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#FFFF99">身分別</td>
					<td>
						<select name="Sys_RoleID" class="btn1">
							<option value="0">請選擇</option><%
							strSQL="select Content,ID from Code where TypeID=5"
							set rs=conn.execute(strSQL)
							while Not rs.eof
								response.write "<option value="""&rs("ID")&""">"
								response.write rs("Content")
								response.write "</option>"
								rs.movenext
							wend
							rs.close%>
						</select>
					</td>
					<td bgcolor="#FFFF99"><b>權限群組</b></td>
					<td>
						<select name="Sys_GroupRoleID" class="btn1">
							<option value="0">請選擇</option><%
							strSQL="select ID , content from Code where TypeID=10"
							set rs=conn.execute(strSQL)
							while Not rs.eof
								if session("Group_ID") <=rs("ID") then 
									response.write "<option value="""&rs("ID")&""">"
									response.write rs("Content")
									response.write "</option>"
								end if									
								rs.movenext								
							wend							
							rs.close%>
						</select>
						
						需要使用本系統才需填寫
					</td>
				</tr>
				<tr bgcolor="#ffffff">

					<td bgcolor="#FFFF99">主管權限</td>
					<td colspan="3">
						<input type="checkbox" name="Sys_ManagerPower" value="1">具備主管權限可以檢視同一單位其他人員資料
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">任用日期</td>
					<td>
						<input name="StartJobDate" class="btn1" type="text" value="<%=daynow%>" size="4" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('StartJobDate');">
					</td>
					<td bgcolor="#FFFF99">離職日期</td>
					<td colspan="1">
						<input name="LeaveJobDate" class="btn1" type="text" value="" size="4" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('LeaveJobDate');">
					</td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#FFFF99">聯絡電話</td>
					<td>
						<input name="Sys_Telephone" class="btn1" type="text" value="" size="12" maxlength="12">
					</td>

					<td bgcolor="#FFFF99">使用者Email</td>
					<td colspan="1">
						<input name="Sys_Email" class="btn1" type="text" value="" size="25" maxlength="25">
					</td>
				</tr>
				
				<tr  bgcolor="#ffffff">
					<td bgcolor="#FFFF99">薪資</td>
					<td>
						<input name="Sys_Money" class="btn1" type="text" value="" size="12" maxlength="12">
					</td>
					<td bgcolor="#FFFF99">銀行/郵局 局號</td>
					<td>
						<input name="Sys_BankName" class="btn1" type="text" value="" size="35" maxlength="35">
					</td>
				</tr>
				
				<tr  bgcolor="#ffffff">
					<td bgcolor="#FFFF99">銀行/郵局 代號</td>
					<td>
						<input name="Sys_BankID" class="btn1" type="text" value="" size="20" maxlength="20">
					</td>
					<td bgcolor="#FFFF99">銀行/郵局 個人帳號</td>
					<td>
						<input name="Sys_BankAccount" class="btn1" type="text" value="" size="25" maxlength="30">
					</td>
				</tr>
								
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">帳號狀態</td>
					<td colspan="3">
						<select name="Sys_ACCOUNTSTATEID" class="btn1">
							<option value="0">啟用</option>
							<option value="-1">停用</option>
						</select>
					</td>

				</tr>
		  </table>
		</td>
	</tr>
	<tr bgcolor="#ffffff" align="center">
		<td height="35" bgcolor="#FFDD77">
			<input type="button" name="save" value=" 儲 存 " onclick="funAdd();">
			<input type="button" name="exit" value=" 離 開 " onclick="funExt();">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Add" value="">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}
}
function funLoginType(){ 
	if (myForm.Sys_LoginType.checked){
		myForm.Sys_LoginID.value="";
		myForm.Sys_LoginID.disabled=true;
	}else{
		myForm.Sys_LoginID.disabled=false;
	}
}
function funAdd(){
	var err=0;
	var StartJobDate,LeaveJobDate,date_y,date_m,date_d,Sys_date;
	if(myForm.StartJobDate.value!=""){
		if(!dateCheck(myForm.StartJobDate.value)){
			err=1;
			alert("任用日輸入不正確!!");
		}
	}
	if (err==0){
		if(myForm.LeaveJobDate.value!=""){
			if(!dateCheck(myForm.LeaveJobDate.value)){
				err=1;
				alert("離職日輸入不正確!!");
			}
		}
	}
	if (err==0){
		if(myForm.StartJobDate.value!=""&&myForm.LeaveJobDate.value!=""){
			Sys_date=myForm.StartJobDate.value;
			date_y=eval(Sys_date.substr(0,eval(Sys_date.length)-4))+1911;
			date_m=Sys_date.substr(eval(Sys_date.length)-4,2);
			date_d=Sys_date.substr(eval(Sys_date.length)-2,2);
			StartJobDate= new Date(date_y+'/'+date_m+'/'+date_d);
			Sys_date=myForm.LeaveJobDate.value;
			date_y=eval(Sys_date.substr(0,eval(Sys_date.length)-4))+1911;
			date_m=Sys_date.substr(eval(Sys_date.length)-4,2);
			date_d=Sys_date.substr(eval(Sys_date.length)-2,2);
			LeaveJobDate= new Date(date_y+'/'+date_m+'/'+date_d);
			if (StartJobDate > LeaveJobDate){
				err=1;
    			alert('離職日期必需要大於任用日期');
			}else{
				myForm.Sys_ACCOUNTSTATEID.value="-1";
			}
		}
	}
	if (err==0){
		if (!myForm.Sys_LoginType.checked){
			if(myForm.Sys_LoginID.value==''){
				err=1;
				alert("使用者臂章號碼不可空白");
			}
		}
		if (err==0){
			/*if(myForm.Sys_PassWord.value==''){
				err=1;
				alert("使用者密碼不可空白");
			}else */
			if(myForm.Sys_ChName.value==''){
				err=1;
				alert("使用者姓名不可空白");
			}else if(myForm.Sys_UnitID.value==''){
				err=1;
				alert("隸屬單位必須選擇");
			}else if(myForm.Sys_JobID.value==''){
				err=1;
				alert("使用者職級必須選擇");
			}else if(myForm.Sys_RoleID.value==''){
				err=1;
				alert("身分別必須選擇");
			/*}else if(myForm.Sys_CreditID.value==''){
				err=1;
				alert("身分證必須填寫");*/
			}else{
				runServerScript("chkAddNew.asp?LoginID="+myForm.Sys_LoginID.value);
			}
		}
	}
}
function funExt() {
	if(confirm("是否關閉維護系統?")){
		opener.myForm.submit();
		self.close();
	}
}

</script>
<%conn.close%>