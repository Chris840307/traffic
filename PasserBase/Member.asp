<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!-- #include file="../Common/Bannernodata.asp"-->
<%
AuthorityCheck(228)
if request("DB_state")="Del" then
	strSQL="Update MemberData set RecordStateID=-1,DelMemberID="&Session("User_ID")&" where MemberID="&request("SN")
	conn.execute strSQL
end if
DB_Selt=trim(request("DB_Selt"))
if DB_Selt="Selt" then
	strwhere=""
	if request("Sys_ChName")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.ChName like '%"&request("Sys_ChName")&"%'"
		else
			strwhere=" and a.ChName like '%"&request("Sys_ChName")&"%'"
		end if
	end if
	
	if request("Sys_LoginID")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.LoginID ='"&request("Sys_LoginID")&"'"
		else
			strwhere=" and a.LoginID ='"&request("Sys_LoginID")&"'"
		end if
	end if
	if request("Sys_CreditID")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.CreditID='"&Ucase(request("Sys_CreditID"))&"'"
		else
			strwhere=" and a.CreditID='"&Ucase(request("Sys_CreditID"))&"'"
		end if
	end if
	if request("Sys_UnitID")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.UnitID in('"&request("Sys_UnitID")&"')"
		else
			strwhere=" and a.UnitID in('"&request("Sys_UnitID")&"')"
		end if
	end if
	if request("Sys_JobID")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.JobID='"&request("Sys_JobID")&"'"
		else
			strwhere=" and a.JobID='"&request("Sys_JobID")&"'"
		end if
	end if
	if request("Sys_RoleID")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.RoleID='"&request("Sys_RoleID")&"'"
		else
			strwhere=" and a.RoleID='"&request("Sys_RoleID")&"'"
		end if
	end if
	
	if request("Sys_Telephone")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.Telephone='"&request("Sys_Telephone")&"'"
		else
			strwhere=" and a.Telephone='"&request("Sys_Telephone")&"'"
		end if
	end if
	if request("Sys_ACCOUNTSTATEID")<>"" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.ACCOUNTSTATEID='"&request("Sys_ACCOUNTSTATEID")&"'"
		else
			strwhere=" and a.ACCOUNTSTATEID='"&request("Sys_ACCOUNTSTATEID")&"'"
		end if
	end if
	if request("Sys_LeaveJobDate")="Leave" then
		if strwhere<>"" then
			strwhere=strwhere&" and Not(a.LeaveJobDate is null)"
		else
			strwhere=" and Not(a.LeaveJobDate is null)"
		end if
	elseif request("Sys_LeaveJobDate")="Job" then
		if strwhere<>"" then
			strwhere=strwhere&" and a.LeaveJobDate is null"
		else
			strwhere=" and a.LeaveJobDate is null"
		end if
	end if
	if trim(strwhere)<>"" then
		strSQL="select a.LoginID,a.MemberID,a.ChName,a.Email,a.Telephone,a.AccountStateID,a.JobID,a.RoleID,a.GroupRoleID,b.Content as JobName,c.Content as RoleName,d.UnitName,e.Content as GroupRoleName from MemberData a,Code b,Code c,Code e,UnitInfo d where a.UnitID=d.UnitID(+) and a.JobID=b.ID(+) and a.RoleID=c.ID(+) and a.GroupRoleID=e.ID(+) and a.RecordStateID=0"&strwhere & " Order by a.LoginID desc "
		set rsfound=conn.execute(strSQL)
		

		strCnt="select count(*) as cnt from MemberData a,Code b,Code c,Code e,UnitInfo d where a.UnitID=d.UnitID(+) and a.JobID=b.ID(+) and a.RoleID=c.ID(+) and a.GroupRoleID=e.ID(+) and a.RecordStateID=0"&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=Dbrs("cnt")
		Dbrs.close
		tmpSQL=strwhere
	else
		DB_Selt=""
		Response.write "<script>"
		Response.Write "alert('必須有查詢條件！');"
		Response.write "</script>"
	end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>人員管理</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
<table width="100%" height="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33" height="33">人員管理<img src="space.gif" width="15" height="8"><strong> 目前登錄員警 
		
		<%
							strSQL="select count(*) as member from memberdata where accountstateid=0 and recordstateid=0"
							set rs=conn.execute(strSQL)
							if Not rs.eof then response.write rs("member")
							
							rs.close
		%>人 <img src="space.gif" width="15" height="8">( 任職狀態 )
		
		</strong> </td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table border="0" bgcolor="#FFFFFF" width="100%">
				<tr>
					<td>臂章號碼</td>
					<td>
						<input name="Sys_LoginID" class="btn1" type="text" value="<%=request("Sys_LoginID")%>" size="12" maxlength="12">
					</td>
					<td>使用者姓名</td>
					<td>
						<input name="Sys_ChName" class="btn1" type="text" value="<%=request("Sys_ChName")%>" size="12" maxlength="12">
					</td>
					<td>身分證號</td>
					<td>
						<input name="Sys_CreditID" class="btn1" type="text" value="<%=Ucase(request("Sys_CreditID"))%>" size="12" maxlength="12">
					</td>
					<td>單位</td>
					<td colspan="2">
						<%=UnSelectUnitOption("Sys_UnitID","")%>
					</td>
				</tr>
				<tr>
					<td>職級</td>
					<td>
						<select name="Sys_JobID" class="btn1">
							<option value="">請選擇</option><%
							strSQL="select Content,ID from Code where TypeID=4"
							set rs=conn.execute(strSQL)
							while Not rs.eof
								response.write "<option value="""&rs("ID")&""""
								if trim(request("Sys_JobID"))=trim(rs("ID")) then response.write " Selected"
								response.write ">"
								response.write rs("Content")
								response.write "</option>"
								rs.movenext
							wend
							rs.close%>
						</select>
					</td>
					<td>在職狀況</td>
					<td>
						<select name="Sys_LeaveJobDate" class="btn1">
							<option value="Job"<%if trim(request("Sys_LeaveJobDate"))="Job" then response.write " selected"%>>在職中</option>
							<option value="All"<%if trim(request("Sys_LeaveJobDate"))="All" then response.write " selected"%>>全部</option>
							<option value="Leave"<%if trim(request("Sys_LeaveJobDate"))="Leave" then response.write " selected"%>>已離職</option>
						</select>
					</td>

				<td>聯絡電話</td>
					<td>
						<input name="Sys_Telephone" class="btn1" type="text" value="<%=request("Sys_Telephone")%>" size="12" maxlength="12">
					</td>
					<td>帳號狀態</td>
					<td>
						<select name="Sys_ACCOUNTSTATEID" class="btn1">
							<option value="">請選擇</option>
							<option value="0"<%if trim(request("Sys_ACCOUNTSTATEID"))="0" then response.write " selected"%>>啟用</option>
							<option value="-1"<%if trim(request("Sys_ACCOUNTSTATEID"))="-1" then response.write " selected"%>>停用</option>
						</select>
					</td>
					<td>
						<input type="button" name="btnSelt" value="查詢" onClick='funSelt();'<%if Not CheckPermission(228,1) then response.write " disabled"%>>&nbsp;&nbsp;
						<input type="button" name="btnAdd" value="新增" onClick='funInsert();'<%if Not CheckPermission(228,2) then response.write " disabled"%>>&nbsp;&nbsp;
						<input type="button" name="cancel" value="清除" onClick="location='Member.asp'">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" height="33">人員管理紀錄列表<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄 ) <b> * 人員單位異動->請查詢出個人資料，再使用後方的 "調單位" 功能，避免資料錯亂 </b> </strong></td>
	</tr>
	<%if DB_Selt="Selt" then%>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
				  <th height="30">員警代號</th>
					<th height="34">使用者姓名</th>
					<th height="34">隸屬單位</th>
					<th height="34">權限群組</th>
					<th height="34">使用者職級</th>
					<!--<th height="34">身分別</th>-->
					<!--<th height="34">聯絡電話</th>-->
					<!--<th height="34">Email</th>-->
					<th height="34">帳號狀態</th>					
					<th height="34">操作</th>
				</tr><%
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsfound.eof then rsfound.move Cint(DBcnt)
					for i=DBcnt+1 to DBcnt+10
						if rsfound.eof then exit for
						response.write "<tr bgcolor='#FFFFFF' align='center' "
						lightbarstyle 0 
						response.write ">"
						if trim(rsfound("AccountStateID"))="0" then
							AccountStateName="啟用"
						else
							AccountStateName="停用"
						end if
						response.write "<td align='left'>"&rsfound("LoginID")&"</td>"
						response.write "<td align='left'>"&rsfound("ChName")&"</td>"
						response.write "<td align='left'>"&rsfound("UnitName")&"</td>"
						response.write "<td>"
						if trim(rsfound("GroupRoleID"))<>"0" then response.write rsfound("GroupRoleName")
						response.write "</td>"
						response.write "<td>"
						if trim(rsfound("JobID"))<>"0" then response.write rsfound("JobName")
						response.write "</td>"
						'response.write "<td>"
						'if trim(rsfound("RoleID"))<>"0" then response.write rsfound("RoleName")
						'response.write "</td>"
						'response.write "<td>"&rsfound("Telephone")&"</td>"
						'response.write "<td>"&rsfound("Email")&"</td>"
						response.write "<td>"&AccountStateName&"</td>"
						response.write "<td>"

						response.write "<input type=""button"" name=""Update"" value=""修改"" onclick=""funUpdate('"&rsfound("MemberID")&"');"""
						if Not CheckPermission(228,3) then response.writ " disabled"
						response.write ">"


						response.write "<img src='space.gif' width='10' height='1'>"
						response.write "<input type=""button"" name=""Update"" value=""調單位"" onclick=""funChangeUnit('"&rsfound("MemberID")&"');"""
						if Not CheckPermission(228,3) then response.writ " disabled"
						response.write ">"
						response.write "<img src='space.gif' width='10' height='1'>"
						'smith for 李桂枝
						if rsfound("ChName") <> "李桂枝" then
							response.write "<input type=""button"" name=""Del"" value=""刪除"" onclick=""funDel('"&rsfound("MemberID")&"');"""
							if Not CheckPermission(228,4) then response.writ " disabled"
							response.write ">"
						end if
						'response.write "<input type=""button"" name=""btnMap"" value=""簽章上傳"" onclick=""funMap('"&rsfound("MemberID")&"');"">"

						response.write "</td>"
						response.write "</tr>"
						rsfound.movenext
					next%>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFDD77" align="left">
			*員警代號由最大排列到最小
			<img src='space.gif' width='70' height='1'>
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=Cint(DBcnt)/10+1&"/"&fix(Cint(DBsum)/10+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
			<input type="button" name="btnExecel2" value="使用系統人員清冊" onclick="funchgExecel2();">
		</td>
	</tr>
	<%end if%>
</table>
<input type="Hidden" name="DB_Selt" value="<%=DB_Selt%>">
<input type="Hidden" name="DB_state" value="">
<input type="Hidden" name="SN" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funSelt(){
	myForm.DB_Move.value=0;
	myForm.DB_Selt.value="Selt";
	myForm.submit();
}
function funDbMove(MoveCnt){
	if (eval(MoveCnt)>0){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}else{
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}
}
function funchgExecel(){
	UrlStr="Member_Execel.asp?SQLstr=<%=tmpSQL%>";
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funchgExecel2(){
  <%
    tmpSql3=" and (a.password is not null or trim(a.password)<>'' )"
  %>
	UrlStr="Member_Execel2.asp?SQLstr=<%=tmpSQL&tmpSQL3%>";
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funInsert(){
	UrlStr="Member_Insert.asp";
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funChangeUnit(SN){
	UrlStr="Member_ChangeUnit.asp?SN="+SN;
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funUpdate(SN){
	UrlStr="Member_Update.asp?SN="+SN;
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funDel(SN){
	myForm.SN.value=SN;
	myForm.DB_state.value="Del";
	myForm.submit();
}
function funMap(SN){
	UrlStr="SendStyle.asp?MemberID="+SN;
	newWin(UrlStr,"winMap",700,150,50,10,"yes","yes","yes","no");
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}
</script>
<%conn.close%>