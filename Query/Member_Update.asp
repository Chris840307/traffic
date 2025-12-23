<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
daynow=split(gInitDT(now),"-")
if request("DB_Add")="ADD" then
	StartJobDate=gOutDT(request("StartJobDate"))
	LeaveJobDate=gOutDT(request("LeaveJobDate"))
	Sys_ManagerPower="0"

	if trim(request("Sys_ManagerPower"))<>"" then Sys_ManagerPower=trim(request("Sys_ManagerPower"))
	strSQL="Update MEMBERDATA set LOGINID='"&request("Sys_LoginID")&"',PASSWORD='"&request("Sys_PassWord")&"',JOBID="&request("Sys_JOBID")&",CREDITID='"&Ucase(request("Sys_CREDITID"))&"',EMAIL='"&request("Sys_EMAIL")&"',TELEPHONE='"&request("Sys_TELEPHONE")&"',STARTJOBDATE="&funGetDate(StartJobDate,0)&",LEAVEJOBDATE="&funGetDate(LeaveJobDate,0)&",ACCOUNTSTATEID="&request("Sys_ACCOUNTSTATEID")&",MODIFYTIME="&funGetDate(now,1)&",RECORDDATE="&funGetDate(now,1)&",RECORDMEMBERID="&session("User_ID")&",MONEY="&funTnumber(request("Sys_MONEY"))&",BANKACCOUNT='"&request("Sys_BANKACCOUNT")&"',RoleID="&request("Sys_RoleID")&",GroupRoleID="&request("Sys_GroupRoleID")&",ManagerPower='"&Sys_ManagerPower&"',BankName='"&request("Sys_BankName")&"',BankID='"&request("Sys_BankID")&"' where MemberID="&request("SN")
	
	conn.execute(strSQL)
	%>
	<script language="JavaScript">
		alert ("修改完成!!");
		opener.myForm.submit(); 
		self.close();
	</script><%
else
	strSQL="select * from MemberData where MemberID="&request("SN")
	set rs=conn.execute(strSQL)
	StartJobDate=gInitDT(rs("StartJobDate"))
	LeaveJobDate=gInitDT(rs("LeaveJobDate"))
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>人員資料修改</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
<table width="100%" height="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33">人員資料修改  
			<% 	if Ucase(rs("CreditID")) <> session("Credit_ID") or isnull(rs("CreditID")) then 			
						response.write "  ( 非本人不得修改帳號 / 密碼 ) "
					end if
		%>			
			
		
		</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#dddddd">
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99"><font color="red">* </font>臂章號碼</td>
					<td>
						<%if trim(left(rs("LoginID"),1))="Z" then%>
							無臂章號碼
							<input type="Hidden" name="Sys_LoginID" value="<%=rs("LoginID")%>">
						<%else%>
							<input name="Sys_LoginID" class="btn1" type="text" value="<%=rs("LoginID")%>" size="12" maxlength="12">
						<%end if%>
						<input type="Hidden" name="chk_LoginID" value="<%=rs("LoginID")%>">
					</td>
					<td bgcolor="#FFFF99"><font color="red">* </font>使用者密碼</td>
					<td>	
					
					<input name="Sys_PassWord" class="btn1" type="password" value="<%=rs("PassWord")%>" size="12" maxlength="12"					
					<% 	
						if Ucase(rs("CreditID")) <> session("Credit_ID") or isnull(rs("CreditID")) then 						 
					  						response.write " disabled "
						end if
					%>					 
					 >(請注意鍵盤大小寫)
					

								
					</td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#FFFF99"><font color="red">* </font>使用者姓名</td>
					<td><%=rs("ChName")%></td>
					<td bgcolor="#FFFF99"><font color="red">* </font>使用者身分證號</td>
					<td>
					
					
					<input name="Sys_CreditID" class="btn1" type="text" value="<%=Ucase(rs("CreditID"))%>" size="12" maxlength="12"
					<% 
						if Ucase(rs("CreditID")) <> session("Credit_ID") or isnull(rs("CreditID")) then						 
					  						response.write " disabled "
						end if
					%>
					>(需要使用本系統才需填寫)
						
					</td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#FFFF99">隸屬單位</td>
					<td><%
						strSQL="select UnitName,UnitID from UnitInfo where UnitID='"&trim(rs("UnitID"))&"'"
						set rs1=conn.execute(strSQL)
						response.write rs1("UnitName")
						rs1.close
						%>
					</td>
					<td bgcolor="#FFFF99">使用者職級</td>
					<td>
						<select name="Sys_JobID" class="btn1">
							<option value="0">請選擇</option><%
							strSQL="select Content,ID from Code where TypeID=4"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write "<option value="""&rs1("ID")&""""
								if trim(rs1("ID"))=trim(rs("JobID")) then response.write " selected"
								response.write ">"
								response.write rs1("Content")
								response.write "</option>"
								rs1.movenext
							wend
							rs1.close%>
						</select>
					</td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#FFFF99">身分別</td>
					<td>
						<select name="Sys_RoleID" class="btn1">
							<option value="0">請選擇</option><%
							strSQL="select Content,ID from Code where TypeID=5"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write "<option value="""&rs1("ID")&""""
								if trim(rs1("ID"))=trim(rs("RoleID")) then response.write " selected"
								response.write ">"
								response.write rs1("Content")
								response.write "</option>"
								rs1.movenext
							wend
							rs1.close%>
						</select>
					</td>
					<td bgcolor="#FFFF99">權限群組</td>
					<td>
						<select name="Sys_GroupRoleID" class="btn1">
							<option value="0">請選擇</option><%
							strSQL="select ID , content from Code where TypeID=10"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
									if session("Group_ID") <=rs1("ID") then 
										response.write "<option value="""&rs1("ID")&""""
										if trim(rs1("ID"))=trim(rs("GroupRoleID")) then response.write " selected"
										response.write ">"
										response.write rs1("Content")
										response.write "</option>"
									end if
								rs1.movenext
							wend
							rs1.close%>
						</select>需要使用本系統才需填寫
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">任用日期</td>
					<td>
						<input name="StartJobDate" class="btn1" type="text" value="<%=StartJobDate%>" size="4" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('StartJobDate');">
					</td>
					<td bgcolor="#FFFF99">主管權限</td>
					<td>
						<input type="checkbox" name="Sys_ManagerPower" value="1"<%if trim(rs("ManagerPower"))="1" then response.write " checked"%>>具備主管權限可以檢視同一單位其他人員資料 
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">離職日期</td>
					<td colspan="3">
						<input name="LeaveJobDate" class="btn1" type="text" value="<%=LeaveJobDate%>" size="4" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('LeaveJobDate');">
					</td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#FFFF99">聯絡電話</td>
					<td>
						<input name="Sys_Telephone" class="btn1" type="text" value="<%=rs("Telephone")%>" size="12" maxlength="12">
					</td>				
					<td bgcolor="#FFFF99">使用者Email</td>
					<td colspan="3">
						<input name="Sys_Email" class="btn1" type="text" value="<%=rs("Email")%>" size="20" maxlength="30">
					</td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#FFFF99">薪資</td>
					<td>
						<input name="Sys_Money" class="btn1" type="text" value="<%=rs("Money")%>" size="12" maxlength="12">
					</td>


					<td bgcolor="#FFFF99">銀行/郵局 局號</td>
						<td>
							<input name="Sys_BankName" class="btn1" type="text" value="<%=rs("BankName")%>" size="35" maxlength="35">
						</td>
					</tr>
				
					<tr  bgcolor="#ffffff">
						<td bgcolor="#FFFF99">銀行/郵局 編號</td>
						<td>
							<input name="Sys_BankID" class="btn1" type="text" value="<%=rs("BankID")%>" size="20" maxlength="20">
						</td>				
						<td bgcolor="#FFFF99">銀行/郵局 個人帳號</td>
						<td>
							<input name="Sys_BankAccount" class="btn1" type="text" value="<%=rs("BankAccount")%>" size="25" maxlength="30">
						</td>
					</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#FFFF99">帳號狀態</td>
					<td colspan="3">
						<select name="Sys_ACCOUNTSTATEID" class="btn1">
							<option value="0"<%if trim(rs("ACCOUNTSTATEID"))="0" then response.write " Selected"%>>啟用</option>
							<option value="-1"<%if trim(rs("ACCOUNTSTATEID"))="-1" then response.write " Selected"%>>停用</option>
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
<input type="Hidden" name="SN" value="<%=request("SN")%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
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
			}else{
				myForm.Sys_ACCOUNTSTATEID.value="-1";
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
			}
		}
	}
	if (err==0){
		if(myForm.Sys_LoginID.value==''){
			err=1;
			alert("使用者臂章號碼不可空白");
		/*}else if(myForm.Sys_PassWord.value==''){
			err=1;
			alert("使用者密碼不可空白");
		}else if(myForm.Sys_CreditID.value==''){
			err=1;
			alert("使用者身份證不可空白");*/
		/*}else if(myForm.Sys_ChName.value==''){
			err=1;
			alert("使用者姓名不可空白");*/
		}else if(myForm.Sys_JobID.value==''){
			err=1;
			alert("使用者職級必須選擇");
		}else if(myForm.Sys_RoleID.value==''){
			err=1;
			alert("身分別必須選擇");
		/*}else if(myForm.chk_LoginID.value!=myForm.Sys_LoginID.value){
			runServerScript("chkAddNew.asp?LoginID="+myForm.Sys_LoginID.value);*/
		}else{
			myForm.DB_Add.value='ADD';
			myForm.submit();
		}
	}
}
function funExt() {
	if(confirm("是否關閉維護系統?")){
		self.close();
	}
}

</script>
<%end if
conn.close%>