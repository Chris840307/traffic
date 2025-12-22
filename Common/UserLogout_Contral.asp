
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title></title>
</head>
<%
dim conn
Set conn = Server.CreateObject("ADODB.Connection")  
conn.Open "dsn=TFD;uid=TFD;pwd=tfduser;"
	strLog="insert into Log(MemID,LogTime,WKS,ActionID,Result,Note,UnitID)"
	strLog=strLog&" values('"&trim(request.Cookies("MemID"))&"',getdate(),'WEB系統登出',1318,'1','WEB系統登出','"&trim(request.cookies("UnitID"))&"')"
	conn.execute strLog
response.cookies("MemID")=""
response.cookies("GroupID")=""
response.cookies("UnitID")=""

Session.Contents.Remove("User_Unit_ID")
Session.Contents.Remove("User_ID")
Session.Contents.Remove("User_Logined")
Session.Contents.Remove("FuncFAX")
Session.Contents.Remove("FuncContact")

conn.close
set conn=nothing
%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="setTimeout(countDown,5000);">
<div align="center">登出成功，為安全起見，請登出後將視窗關閉<br>
  <a href="TFD_Login.asp">重新登入
  </a>
</div>
</body>
<script language="JavaScript"> 
function countDown() {   
	location="TFD_Login.asp";
}   
</Script>  
</html>
