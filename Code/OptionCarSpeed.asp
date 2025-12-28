<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->

<HTML>
<HEAD>
<TITLE>特殊車種車速設定系統</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</HEAD>

<BODY>
<%
AuthorityCheck(233)
if trim(request("DB_State"))="Update" then
	strSQL="Update CarSpeed set value="&trim(Request("Update_Value"))&" where ID='"&trim(Request("Update_CarSN"))&"'"
	conn.execute(strSQL)
	Response.write "<script>"
	Response.Write "alert('更新完成！');"
	Response.write "</script>"
end if
strSQL="select ID,Content,value from CarSpeed"

set rsfound=conn.execute(strSQL)

strCnt="select count(*) as cnt from CarSpeed"
set Dbrs=conn.execute(strCnt)
DBsum=CDbl(Dbrs("cnt"))
Dbrs.close
%>
<form name=myForm method="post">
<table width="100%" border="0">
<tr>
	<td bgcolor="#1BF5FF" height="33">稽核特殊車種車速列表<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄 )</strong></td>
</tr>
<tr>
	<td bgcolor="#E0E0E0">
		<Div style="overflow:auto;width:100%;height:360px;background:#FFFFFF">
			<table width="100%" bgcolor="#E0E0E0" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#FAFAF5" align="center">
					<td>序號</td>
					<td>車種</td>					
					<td>目前設定車速</td>
					<td>操作</td>
				</tr><%
				if Trim(request("DB_Move"))="" then
					DBcnt=0
				else
					DBcnt=request("DB_Move")
				end if
				if Not rsfound.eof then rsfound.move DBcnt
				for i=DBcnt+1 to DBcnt+10
					if rsfound.eof then exit for
					response.write "<tr bgcolor='#FFFFFF' align=""right"""
					lightbarstyle 0 
					response.write ">"
					response.write "<td>"&i&"</td>"
					response.write "<td>"&rsfound("Content")&"</td>"
					response.write "<td><input ID='Sys_Value' type='text' class=""btn1"" size='21' maxlength='20' value='"&rsfound("Value")&"'></td>"

					response.write "<td height=""23""><input type=""button"" name=""Update"" value=""確定"" onclick=""funUpdate('"&rsfound("ID")&"','"&i-DBcnt&"');""></td>"

					response.write "</tr>"
					rsfound.movenext
				next%>
			</table>
		</Div>
	</td>
</tr>
<tr>
	<td height="30" colspan="10" bgcolor="#1BF5FF" align="center">
		<a href="file:///.."></a>
		<input type="button" name="MoveFirst" value="第一頁" onclick="funDbMove(0);">
		<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
		<span class="style2"> <%=fix(CDbl(DBcnt)/(10)+1)&"/"&fix(CDbl(DBsum)/(10)+0.9)%></span>
		<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
		<input type="button" name="MoveDown" value="最後一頁" onclick="funDbMove(999);">
	</td>
</tr>
</table>
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="Hidden" name="Update_CarSN" value="">
<input type="Hidden" name="Update_Value" value="">
<input type="Hidden" name="DB_State" value="">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funDbMove(MoveCnt){
	if (eval(MoveCnt)==0){
		myForm.DB_Move.value="";
		myForm.submit();
	}else if (eval(MoveCnt)==10){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}else if(eval(MoveCnt)==-10){
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}else if(eval(MoveCnt)==999){
		if (eval(myForm.DB_Cnt.value)%(10)==0){
			myForm.DB_Move.value=(Math.floor(eval(myForm.DB_Cnt.value)/(10))-1)*(10);
		}else{
			myForm.DB_Move.value=Math.floor(eval(myForm.DB_Cnt.value)/(10))*(10);
		}
		myForm.submit();
	}
}
function funUpdate(CarSN,CarCnt){
	if(confirm('確定更新此筆紀錄嗎？')){
		myForm.Update_CarSN.value=CarSN;
		myForm.Update_Value.value=document.all.Sys_Value[CarCnt-1].value;
		myForm.DB_State.value='Update';
		myForm.submit();
	}
}
function funchgExecel(){
	myForm.action="CarSpeed_Execel.asp";
	myForm.target="DetailCar";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
</script>