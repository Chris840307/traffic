<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
if request("DB_State")="Add" then
	strSQL="Insert into ArgueDetail(SN,ArgueBaseSN,AttachID,Note) values("&funTableSeq("ArgueDetail","SN")&","&request("ArgueBaseSN")&","&request("Sys_AttachID")&",'"&request("Sys_Note")&"')" 'Note Access 不充許note保留字
	conn.execute strSQL
elseif request("DB_State")="Update" then
	strSQL="Update ArgueDetail set AttachID='"&request("Sys_AttachID_Edit")&"',Note='"&request("Sys_Note_Edit")&"' where SN="&request("Update_CarSN")
	conn.execute strSQL
elseif request("DB_State")="Del" then
	strSQL="Delete from ArgueDetail where SN="&request("Update_CarSN")
	conn.execute strSQL
end if
if request("DB_Selt")="Selt" then
	str1=""
	if trim(request("Sys_AttachID"))<>"" then
		str1=" and AttachID="&trim(request("Sys_AttachID"))
	end if
	if trim(request("Sys_Note"))<>"" then
		str1=str1&" and Note Like '%"&trim(request("Sys_Note"))&"%'"
	end if
	strSQL="select a.*,b.Content as AttachName from ArgueDetail a,Code b where a.AttachID=b.ID"&str1&" and ArgueBaseSN="&request("ArgueBaseSN")
	set rs=conn.execute(strSQL)

	strCnt="select count(*) as cnt from ArgueDetail a,Code b where a.AttachID=b.ID"&str1&" and ArgueBaseSN="&request("ArgueBaseSN")

	set Dbrs=conn.execute(strCnt)
	DBsum=Cint(Dbrs("cnt"))
	Dbrs.close

	tmpSQL=strSQL
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>申訴案件</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body onkeydown="KeyDown()">
<form name=myForm method="post">
<table width="100%" height="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33">申訴案件-附件資料</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>附件物品
						<select name="Sys_AttachID" class="btn1">
							<option value="">請選擇</option><%
							strSQL="select ID,Content from Code where TypeID=14"
							set reselt=conn.execute(strSQL)
							while Not reselt.eof
								response.write "<option value="""&reselt("ID")&""">"&reselt("Content")&"</option>"
								reselt.movenext
							wend
							reselt.close
						%></select>&nbsp;&nbsp;&nbsp;&nbsp;
						備註
						<input name="Sys_Note" type="text" class="btn1" size="50" maxlength="50">&nbsp;&nbsp;&nbsp;&nbsp;

						<input type="button" name="btnSelt" value="查詢" onclick="funSelt();"<%if Not CheckPermission(223,1) then response.write " disabled"%>>&nbsp;&nbsp;
						<input type="button" name="btnAdd" value="新增" onclick="funAdd();"<%if Not CheckPermission(223,2) then response.write " disabled"%>>
					</td>
				</tr>
				<tr><td><font color="red">請於上述欄位輸入申註案件附本資料後點選新增,即可新增資料</font></td></tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33">申訴案件附件資料列表</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th bgcolor="#EBFBE3">流水號</th>
					<th bgcolor="#EBFBE3">附件物品</th>
					<th bgcolor="#EBFBE3">備註</th>
					<th bgcolor="#EBFBE3">操作</th>
				</tr>
				<%
				if request("DB_Selt")="Selt" then
					tempSN=0
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rs.eof then rs.move Cint(DBcnt)
					for i=DBcnt+1 to DBcnt+10
						if rs.eof then exit for
						tempSN=tempSN+1
						response.write "<tr align='center' bgcolor='#FFFFFF'"
						lightbarstyle 0
						response.write ">"
						if trim(request("Edit_CarSN"))=trim(rs("SN")) then
							response.write "<td height='23'>"&tempSN&"</td>"
							response.write "<td height='23'>"
							response.write "<select name=""Sys_AttachID_Edit"" class=""btn1""><option value="""">請選擇</option>"
							strSQL="select ID,Content from Code where TypeID=14"
							set reselt=conn.execute(strSQL)
							while Not reselt.eof
								response.write "<option value="""&reselt("ID")&""""
								if trim(reselt("ID"))=trim(rs("AttachID")) then response.write " selected"
								response.write ">"&reselt("Content")&"</option>"
								reselt.movenext
							wend
							reselt.close
							response.write "</select>"
							response.write "<td height='23'><input name='Sys_Note_Edit' type='text' class=""btn1"" size='21' maxlength='20' value='"&rs("Note")&"'></td>"%>
							<td>
								<input type="button" name="Update" value="確定" onclick="funUpdate('<%=rs("SN")%>');">
								<input type='button' name='Canal' value='取消' onclick="funEdit('');">
							</td><%
						else
							response.write "<td height='23'>"&tempSN&"</td>"
							response.write "<td height='23'>"&rs("AttachName")&"</td>"
							response.write "<td height='23'>"&rs("Note")&"</td>"%>
							<td>
								<input type="button" name="Edit" value="修改" onclick="funEdit('<%=rs("SN")%>');"<%if Not CheckPermission(223,3) then response.write " disabled"%>>
								<input type='button' name='Del' value='刪除' onclick="funDel('<%=rs("SN")%>');"<%if Not CheckPermission(223,4) then response.write " disabled"%>>
							</td><%
						end if
						response.write "</tr>"
						rs.movenext
					next
					rs.close
				end if
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFDD77" align="center">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=Cint(DBcnt)/10+1&"/"&fix(Cint(DBsum)/10+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<img src="space.gif" width="8" height="8">
			<img src="space.gif" width="8" height="8">
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
			<input type="button" name="exit" value=" 關 閉 " onclick="funExt();">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="DB_Edit" value="">
<input type="Hidden" name="DB_State" value="">
<input type="Hidden" name="Edit_CarSN" value="">
<input type="Hidden" name="Update_CarSN" value="">
<input type="Hidden" name="ArgueBaseSN" value="<%=request("ArgueBaseSN")%>">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/engine.js"></script>
<script language="javascript">
function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}
}
function funAdd(CarSN){
	var err=0;
	if(myForm.Sys_AttachID.value==''){
		err=1;
		alert("附件不可空白");
	}
	if(err==0){
		myForm.Edit_CarSN.value=CarSN;
		myForm.DB_Selt.value="Selt";
		myForm.DB_State.value='Add';
		myForm.submit();
	}
}
function funEdit(CarSN){
	myForm.Edit_CarSN.value=CarSN;
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
function funUpdate(CarSN){
	var err=0;
	if(myForm.Sys_AttachID_Edit.vaoue==''){
		err=1;
		alert("附件不可空白");
	}
	if(err==0){
		myForm.Update_CarSN.value=CarSN;
		myForm.DB_State.value='Update';
		myForm.submit();
	}
}
function funDel(CarSN){
	if(confirm('確定刪除此筆紀錄嗎？')){
		myForm.Update_CarSN.value=CarSN;
		myForm.DB_State.value='Del';
		myForm.submit();
	}
}
function funSelt(){
	myForm.DB_Selt.value="Selt";
	myForm.submit();
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,"otherwin","width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}
function funchgExecel(){
	UrlStr="ArgueCaseAttach_Execel.asp?SQLstr=<%=tmpSQL%>";
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funExt() {
	if(confirm("是否關閉維護系統?")){
		self.close();
	}
}
</script>
<%conn.close%>