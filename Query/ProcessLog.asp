<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!-- #include file="../Common/Bannernodata.asp"-->
<HTML>
<HEAD>
<TITLE> 伺服器狀況查詢 </TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
</HEAD>
<%
AuthorityCheck(225)
DB_Selt=request("DB_Selt")
if DB_Selt="Selt" then
	strwhere=""
	if request("Sys_Update1")<>"" and request("Sys_Update2")<>"" then
		ArgueDate1=gOutDT(request("Sys_Update1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("Sys_Update2"))&" 23:59:59"
		if strwhere<>"" then
			strwhere=strwhere&" where ""Update"" between "&funGetDate(ArgueDate1,1)&" and "&funGetDate(ArgueDate2,1)
		else
			strwhere=" where ""Update"" between "&funGetDate(ArgueDate1,1)&" and "&funGetDate(ArgueDate2,1)
		end if
	end if
	if trim(strwhere)<>"" then
		strSQL="select * from ProcessLog"&strwhere

		set rsfound=conn.execute(strSQL)
		
		strCnt="select count(*) as cnt from ProcessLog"&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=Cint(Dbrs("cnt"))
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
<BODY>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33" height="33">伺服器狀況查詢</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%"  border="0" bgcolor="#FFFFFF">
				<tr>
					<td>更新日期</td>
					<td>
						<input name="Sys_Update1" type="text" class="btn1" value="<%=request("Sys_Update1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_Update1');">
						~
						<input name="Sys_Update2" type="text" class="btn1" value="<%=request("Sys_Update2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_Update2');"> 

						<input type="button" name="btnSelt" value="查詢" onClick="funSelt();"<%if Not CheckPermission(225,1) then response.write " disabled"%>>
						&nbsp;&nbsp;
						<input type="button" name="cancel" value="清除" onClick="location='ProcessLog.asp'">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" height="33">伺服器狀況列表<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄 )</strong></td>
	</tr>
	<%if DB_Selt="Selt" then%>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th height="34">檔案名稱</th>
					<th height="34">狀態</th>
					<th height="34">使用記憶體大小(MB)</th>
					<th height="34">更新時間</th>
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
						if trim(rsfound("Status"))="1" then
							StatusName="執行"
						else
							StatusName="未執行"
						end if
						response.write "<td>"&rsfound("FileName")&"</td>"
						response.write "<td>"&StatusName&"</td>"
						response.write "<td>"&rsfound("MemorySize")&"</td>"
						response.write "<td>"&gInitDT(DateValue(rsfound("Update")))&" "&TimeValue(rsfound("Update"))&"</td>"
						response.write "</tr>"
						rsfound.movenext
					next%>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFDD77" align="center">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=Cint(DBcnt)/10+1&"/"&fix(Cint(DBsum)/10+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
		</td>
	</tr>
	<%end if%>
</table>
<input type="Hidden" name="DB_Selt" value="<%=DB_Selt%>">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funSelt(){
	var error=0;
	if(myForm.Sys_Update1.value==""){
		error=1;
		alert("日期必須填寫!!");
	}else{
		if(!dateCheck(myForm.Sys_Update1.value)){
			error=1;
			alert("日期輸入不正確!!");
		}
	}
	if (error==0){
		if(myForm.Sys_Update2.value==""){
			error=1;
			alert("日期必須填寫!!");
		}else{
			if(!dateCheck(myForm.Sys_Update2.value)){
				error=1;
				alert("日期輸入不正確!!");
			}
		}
	}
	if (error==0){
		if (error==0){
			myForm.DB_Move.value=0;
			myForm.DB_Selt.value="Selt";
			myForm.submit();
		}
	}
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
</script>
<%conn.close%>