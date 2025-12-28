
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/bannernodata.asp"-->
<%

strSQL="select FILENAME,IMPORTDATE,IMPORTMEM from CaseImport  order by IMPORTDATE desc" 'access日期時間沒有單引號
set rs=conn.execute(strSQL)
																							'and a.CarNo=b.CarNo 
strCnt="select count(*) as cnt from ("&strSQL&")"

set Dbrs=conn.execute(strCnt)
DBsum=Cint(Dbrs("cnt"))
Dbrs.close
'response.write strCnt
'response.end
tmpSQL=strSQL
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>匯入檔案歷程紀錄</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#1BF5FF">匯入檔案歷程紀錄列表<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄 )</strong></td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th height="34">匯入檔案</th>
					<th height="34">匯入時間</th>
					<th height="34" nowrap>匯入人員</th>
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
						response.write "<td>"&rs("FILENAME")&"</td>"
						response.write "<td nowrap>"&rs("IMPORTDATE")&"</td>"
						response.write "<td>"&rs("IMPORTMEM")&"</td>"
						response.write "</tr>"
						rs.movenext
					next%>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#1BF5FF" align="center">
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