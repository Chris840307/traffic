<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
daynow=split(gInitDT(now),"-")

	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing
 

if request("DB_Add")="ADD" Then
	PassWordTemp=Trim(request("Sys_PassWord"))
	'要使用密碼加密 要記得更新 AllFunction.inc
	If sys_City="澎湖縣" Or sys_City="基隆市" Or sys_City="高雄市" Or sys_City="金門縣" Or sys_City="台東縣" Or sys_City="彰化縣" Or sys_City="台中市" Or sys_City="屏東縣" Or sys_City="嘉義縣" Or sys_City="雲林縣" Or sys_City="嘉義市" Or sys_City="新竹市" Then
		DecodePassWordTemp=decrypt(Trim(request("Sys_PassWord")))
	Else
		DecodePassWordTemp=Trim(request("Sys_PassWord"))
	End If 
	chkUp=0
	chkDown=0
	chkInt=0
	chkMark=0
	chkSingle=0
	for i=1 to Len(PassWordTemp)
		if Asc(Mid(Trim(PassWordTemp), i, 1))>=65 and Asc(Mid(Trim(PassWordTemp), i, 1))<=90 then
			chkUp=1
		end if
		if Asc(Mid(Trim(PassWordTemp), i, 1))>=97 and Asc(Mid(Trim(PassWordTemp), i, 1))<=122 then
			chkDown=1
		end if 
		if Asc(Mid(Trim(PassWordTemp), i, 1))>=48 and Asc(Mid(Trim(PassWordTemp), i, 1))<=57 then
			chkInt=1
		end if 
		if (Asc(Mid(Trim(PassWordTemp), i, 1))>=33 and Asc(Mid(Trim(PassWordTemp), i, 1))<=47) or (Asc(Mid(Trim(PassWordTemp), i, 1))>=58 and Asc(Mid(Trim(PassWordTemp), i, 1))<=64) or (Asc(Mid(Trim(PassWordTemp), i, 1))>=91 and Asc(Mid(Trim(PassWordTemp), i, 1))<=96) or (Asc(Mid(Trim(PassWordTemp), i, 1))>=123 and Asc(Mid(Trim(PassWordTemp), i, 1))<=126) then
			chkMark=1
		end If
		If Mid(Trim(PassWordTemp), i, 1)="'" Then
			chkSingle=1
		End If 
	Next
	
	PassWord3Time=0
	'檢查密碼是否用過，其他縣市要開放的話，還有下面的insert也要加
	If sys_City="基隆市" Or sys_City="澎湖縣" Or sys_City="台南市" Or sys_City="高雄市" Or sys_City="金門縣" Or sys_City="台東縣" Or sys_City="嘉義縣" Or sys_City="彰化縣" Or sys_City="雲林縣" Or sys_City="嘉義市" Or sys_City="新竹市" then
		if Trim(request("Sys_PassWord"))<> Trim(request("Sys_PassWord_Old")) then	'有修改密碼再檢查
			strChk="select count(*) as cnt from " &_
				"(select password from (select * from MemberUsePassword where MemberID="&trim(Session("User_ID"))&" order by recorddate desc) where rownum<=3)" &_
				" where password='"&DecodePassWordTemp&"'"
			Set rsChk=conn.execute(strChk)
			If Not rsChk.eof Then
				If CInt(rsChk("cnt"))>0 Then
					PassWord3Time=1
				End If 
			End If 
			rsChk.close
			Set rsChk=Nothing 
		end if
	End if

	if chkUp=0 or chkDown=0 or chkInt=0 or chkMark=0 Or Len(PassWordTemp)<8 Then
		%>
		<script language="JavaScript">
			alert('密碼長度至少為<8>碼，包含英文、數字、特殊符號及大小寫混和!!');
		</script><%
	Elseif PassWord3Time=1 Then
	%>
		<script language="JavaScript">
			alert('新密碼不可以與前三次使用過之密碼相同!!');
		</script>
	<%
	ElseIf chkSingle=1 then
		%>
		<script language="JavaScript">
			alert('密碼請勿使用單引號!!');
		</script><%
	Else
		StartJobDate=gOutDT(request("StartJobDate"))
		LeaveJobDate=gOutDT(request("LeaveJobDate"))
		Sys_ManagerPower="0"
		if sys_City="高雄市" or sys_City="台中市" or sys_City="苗栗縣" or sys_City="台南市" or sys_City="彰化縣" or sys_City="屏東縣" or sys_City="基隆市" or sys_City="保二總隊三大隊二中隊" or sys_City="嘉義市" or sys_City="金門縣" or sys_City="南投縣" then
			strADD=",MPOLICEID='"&UCase(trim(Request("Sys_MpoliceID")))&"'"
		End If 
		if trim(request("Sys_ManagerPower"))<>"" then Sys_ManagerPower=trim(request("Sys_ManagerPower"))
		strSQL="Update MEMBERDATA set PASSWORD='"&DecodePassWordTemp&"',EMAIL='"&request("Sys_EMAIL")&"',TELEPHONE='"&request("Sys_TELEPHONE")&"',MODIFYTIME="&funGetDate(now,1)&",RECORDDATE="&funGetDate(now,1)&",RECORDMEMBERID="&session("User_ID")&",BANKACCOUNT='"&request("Sys_BANKACCOUNT")&"'"&strADD&" where MemberID="&Session("User_ID")
		ConnExecute replace(strSQL,DecodePassWordTemp,"*******"),"359"
		conn.execute(strSQL)
		
		'檢查密碼是否用過
		'  CREATE TABLE TRAFFIC.MemberUsePassword
		'   (	
		'   MemberID NUMBER, 
		'   Password varchar2(200),
		'   recorddate date
		'   ) 
		If sys_City="基隆市" Or sys_City="澎湖縣" Or sys_City="台南市" Or sys_City="高雄市" Or sys_City="金門縣" Or sys_City="台東縣" Or sys_City="嘉義縣" Or sys_City="彰化縣" Or sys_City="雲林縣" Or sys_City="嘉義市" Or sys_City="新竹市" then
			strSql2="Insert Into MemberUsePassword(MemberID,Password,recorddate) " &_
				" values("&session("User_ID")&",'"&DecodePassWordTemp&"',sysdate)"
			conn.execute(strSql2)
		End If 
		%>
		<script language="JavaScript">
			alert ("修改完成，請牢記您的密碼，以免帳號輸入錯誤三次遭鎖定!!");
			//opener.location.reload(); 
			self.close();
		</script><%

	end If
		
	
End if
	strSQL="select * from MemberData where MemberID="&Session("User_ID")
	set rs=conn.execute(strSQL)
	StartJobDate=gInitDT(rs("StartJobDate"))
	LeaveJobDate=gInitDT(rs("LeaveJobDate"))
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>登入者資料修改</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
<table width="100%" height="100%" border="0">
	<tr>
		<td bgcolor="#1BF5FF">登入者資料修改</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#dddddd">
				<tr bgcolor="#ffffff">
					<td bgcolor="#EBF5FF">員警代號<br><font size="2">建檔時舉發員警代號</font></td>
					<td><%=rs("LoginID")%></td>
					<td bgcolor="#EBF5FF"><font color="red">* 系統登入密碼</font></td>
					<td width="35%">
						<input name="Sys_PassWord" class="btn1" type="password" value="<%
					If sys_City="澎湖縣" Or sys_City="基隆市" Or sys_City="高雄市" Or sys_City="金門縣" Or sys_City="台東縣" Or sys_City="彰化縣" Or sys_City="台中市" Or sys_City="屏東縣" Or sys_City="嘉義縣" Or sys_City="雲林縣" Or sys_City="嘉義市" Or sys_City="新竹市" then
						response.write encrypt(rs("PassWord"))
					Else
						response.write rs("PassWord")
					End If 
						%>" size="12" maxlength="20">
						<input name="Sys_PassWord_Old" class="btn1" type="hidden" value="<%
					If sys_City="澎湖縣" Or sys_City="基隆市" Or sys_City="高雄市" Or sys_City="金門縣" Or sys_City="台東縣" Or sys_City="彰化縣" Or sys_City="台中市" Or sys_City="屏東縣" Or sys_City="嘉義縣" Or sys_City="雲林縣" Or sys_City="嘉義市" Or sys_City="新竹市" then
						response.write encrypt(rs("PassWord"))
					Else
						response.write rs("PassWord")
					End If 
						%>" size="12" maxlength="20">
						<br> <font color="red"><strong>密碼長度至少為<8>碼，包含英文、數字、特殊符號及大小寫混和(請勿使用下列特殊符號『& , < , > , " , ' , = , --』，並請每三個月更改一次，不可以與前三次使用過之密碼相同</strong></font>
					</td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#EBF5FF">使用者姓名</td>
					<td><%=rs("ChName")%></td>
					<td bgcolor="#EBF5FF">系統登入帳號</td>
					<td><%=rs("CreditID")%></td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#EBF5FF">隸屬單位</td>
					<td><%
						strSQL="select UnitName,UnitID from UnitInfo where UnitID='"&trim(rs("UnitID"))&"'"
						set rs1=conn.execute(strSQL)
						response.write rs1("UnitName")
						rs1.close
						%>
					</td>
					<td bgcolor="#EBF5FF">使用者職級</td>
					<td><%
						if Not ifnull(rs("JobID")) then
							strSQL="select Content,ID from Code where TypeID=4 and ID="&trim(rs("JobID"))
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write rs1("Content")
								rs1.movenext
							wend
							rs1.close
						end if%>
					</td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#EBF5FF">身分別</td>
					<td><%
						if Not ifnull(rs("RoleID")) then
							strSQL="select Content,ID from Code where TypeID=5 and ID="&trim(rs("RoleID"))
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write rs1("Content")
								rs1.movenext
							wend
							rs1.close
						end if%>
					</td>
					<td bgcolor="#EBF5FF">聯絡電話</td>
					<td>
						<input name="Sys_Telephone" class="btn1" type="text" value="<%=rs("Telephone")%>" size="12" maxlength="12">
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#EBF5FF">任用日期</td>
					<td><%=StartJobDate%></td>
					<td bgcolor="#EBF5FF">主管權限</td>
					<td><%if trim(rs("ManagerPower"))="1" then
							response.write "有"
						else
							response.write "無"
						end if%>
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#EBF5FF">離職日期</td>
					<td <% 
					if sys_City<>"高雄市" and sys_City<>"台中市" and sys_City<>"苗栗縣" and sys_City<>"台南市" and sys_City<>"彰化縣" and sys_City<>"屏東縣" and sys_City<>"基隆市" and sys_City<>"保二總隊三大隊二中隊" and sys_City<>"嘉義市" and sys_City<>"金門縣"  and sys_City<>"南投縣" Then 
						response.write "colspan=3"  
					end If
					%>><%=LeaveJobDate%></td>
<% if sys_City="高雄市" or sys_City="台中市" or sys_City="苗栗縣" or sys_City="台南市" or sys_City="彰化縣" or sys_City="屏東縣" or sys_City="基隆市" or sys_City="保二總隊三大隊二中隊" or sys_City="嘉義市" or sys_City="金門縣" or sys_City="南投縣" then %>
					<td bgcolor="#EBF5FF">Mpolice帳號</td>
					<td >
						<input name="Sys_MpoliceID" class="btn1" type="text" value="<%=rs("MPOLICEID")%>" size="20" maxlength="30" onkeyup="this.value=this.value.toUpperCase()">
					</td>
<% End If %>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#EBF5FF">使用者Email</td>
					<td colspan="3">
						<input name="Sys_Email" class="btn1" type="text" value="<%=rs("Email")%>" size="20" maxlength="30">
					</td>
				</tr>
				<tr  bgcolor="#ffffff">
					<td bgcolor="#EBF5FF">薪資</td>
					<td><%=rs("Money")%></td>
					<td bgcolor="#EBF5FF">銀行帳號</td>
					<td>
						<input name="Sys_BankAccount" class="btn1" type="text" value="<%=rs("BankAccount")%>" size="25" maxlength="30">
					</td>
				</tr>
				<tr bgcolor="#ffffff">
					<td bgcolor="#EBF5FF">帳號狀態</td>
					<td><%if trim(rs("ACCOUNTSTATEID"))="0" then
								response.write "啟用"
						elseif trim(rs("ACCOUNTSTATEID"))="-1" then
							response.write "停用"
						end if%>
					</td>
					<td bgcolor="#EBF5FF">權限群組</td>
					<td><%
						if Not ifnull(rs("GroupRoleID")) then
							strSQL="select ID , content from Code where TypeID=10 and ID="&trim(rs("GroupRoleID"))
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write rs1("Content")
								rs1.movenext
							wend
							rs1.close
						end if%>
					</td>
				</tr>
		  </table>
		</td>
	</tr>
	<tr bgcolor="#ffffff" align="center">
		<td height="35" bgcolor="#1BF5FF">
			<input type="button" name="save" value=" 儲 存 " onclick="funAdd();">
			<input type="button" name="exit" value=" 離 開 " onclick="funExt();">
			<br> <font color="red"><strong>修改密碼時，請牢記您的密碼，以免帳號輸入錯誤三次遭鎖定</strong></font>
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Add" value="">
</form>
</body>
</html>
<script type="text/javascript" src="./js/date.js"></script>
<script language="javascript">
function funAdd(){
	var err=0;
	var StartJobDate,LeaveJobDate,date_y,date_m,date_d,Sys_date;
	if(myForm.Sys_PassWord.value==''){
		err=1;
		alert("使用者密碼不可空白");
	}else if(myForm.Sys_PassWord.value.indexOf('\'')>=0 || myForm.Sys_PassWord.value.indexOf('&')>=0 || myForm.Sys_PassWord.value.indexOf('<')>=0 || myForm.Sys_PassWord.value.indexOf('>')>=0 || myForm.Sys_PassWord.value.indexOf('"')>=0 || myForm.Sys_PassWord.value.indexOf('=')>=0 || myForm.Sys_PassWord.value.indexOf('--')>=0){
		err=1;
		alert("使用者密碼不可使用下列特殊符號『& , < , > , \" , ' , = , --』");
	}else{
		myForm.DB_Add.value='ADD';
		myForm.submit();
	}
}
function funExt() {
	if(confirm("是否關閉維護系統?")){
		self.close();
	}
}

</script>
<%'end if
conn.close%>