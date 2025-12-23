<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>列印大宗郵件清冊</title>
<%
tmpSQL=request("SQLstr")

	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post" onsubmit="return select_street();">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td>列印大宗郵件清冊</td>
			</tr>
			<tr>
				<td bgcolor="#EBFBE3" align="center">
					<input type="button" value="依  縣  市  區  分" onclick="funMailListCity();">
					<input type="button" value="不需依縣市區分" onclick="funMailList();">
				<%if sys_City="基隆市" then%>
					<br>
					<input type="checkbox" name="NoteMailNo" value="1">附送達證書
				<%end if%>
				</td>
			</tr>
		</table>	
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	winopen=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	winopen.focus();
	return win;
}
//不須依縣市分
function funMailList(){
		var NoteMailNo="0";
		<%if sys_City="基隆市" then%>
			if (myForm.NoteMailNo.checked==true){
				NoteMailNo="1";
			}
		<%end if%>
		UrlStr="MailSendList_Excel.asp?SQLstr=<%=tmpSQL%>&NoteMailNo="+NoteMailNo;
		newWin(UrlStr,"MailSendList_1",1000,700,0,0,"yes","yes","yes","no");
		window.close();
}
//依縣市分
function funMailListCity(){
		var NoteMailNo="0";
		<%if sys_City="基隆市" then%>
			if (myForm.NoteMailNo.checked==true){
				NoteMailNo="1";
			}
		<%end if%>
		UrlStr="MailSendList_City_Excel.asp?SQLstr=<%=tmpSQL%>&NoteMailNo="+NoteMailNo;
		newWin(UrlStr,"MailSendList_2",1000,700,0,0,"yes","yes","yes","no");
		window.close();
}
</script>
</html>
