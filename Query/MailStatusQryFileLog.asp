
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/bannernodata.asp"-->
<%
Function GetCnt(sn)
	tmp=""
	strSQL="select count(*) as cnt from MailResult where orgsn="&sn
	set rstmp=conn.execute(strSQL)
	If Not rstmp.eof Then tmp=rstmp("cnt")
	Getcnt=tmp
End function

strSQL="select sn,FileName,Max(RecordDate) as RecordDate,Result from MailResultProcessLog where RecordDate >="&funGetDate(DateADD("d",-7,now),1)&" Group by sn,FileName,Result order by RecordDate desc"
set rs=conn.execute(strSQL)

strCnt="select count(*) as cnt from ("&strSQL&")"

set Dbrs=conn.execute(strCnt)
DBsum=Cint(Dbrs("cnt"))
Dbrs.close

tmpSQL=strSQL
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>匯入歷程紀錄</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33">匯入歷程紀錄列表<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄 )</strong></td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th height="34">匯入檔案</th>
					<th height="34">匯入時間</th>
					<th height="34">是否成功</th>
					<th height="34">筆數</th>
				</tr><%
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rs.eof then rs.move Cint(DBcnt)
					for i=DBcnt+1 to DBcnt+10
						if rs.eof then exit for
						response.write "<tr bgcolor='#FFFFFF' align='center' "
						lightbarstyle 0 
						response.write ">" 
						response.write "<td>"&rs("Filename")&"</td>"
						response.write "<td nowrap>"&rs("RecordDate")&"</td>"
						If rs("Result")="1" Then 
							response.write "<td>成功</td>"
						Else
							response.write "<td>失敗</td>"
						End If
						response.write "<td nowrap>"&GetCnt(rs("sn"))&"</td>"


						response.write "</tr>"
						rs.movenext
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
</table>
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
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