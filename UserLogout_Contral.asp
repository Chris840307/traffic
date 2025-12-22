<%
if trim(Session("FuncID"))="" or trim(Session("FuncID"))="" then
	Response.Redirect "/traffic/Traffic_Login.asp"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>交通裁處系統</title>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
<style type="text/css">
<!--
.style1 {font-size: 12px}
.style2 {font-size: 14px}
.style3 {font-size: 16px}
-->
</style>
</head>
<%
ConnExecute Session("Credit_ID")&","&Session("Ch_Name")&" "&"登出",351 
conn.close
set conn=nothing

Session.Contents.Remove("FuncID")
Session.Contents.Remove("Unit_ID")
Session.Contents.Remove("User_ID")
Session.Contents.Remove("Ch_Name")
Session.Contents.Remove("Credit_ID")
Session.Contents.Remove("Group_ID")
Session.Contents.Remove("DoubleCheck")
Session.Contents.Remove("ManagerPower")
Session.Contents.Remove("UnitLevelID")
Session.Contents.Remove("DCIwindowName")
Session.Contents.Remove("BillIgnore_Fix")
Session.Contents.Remove("BillIgnore_Image")
%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="setTimeout(countDown,5000);">
<div align="center">
  <div id="Layer1" style="position:absolute; width:845px; height:491px; z-index:1; left: 82px; top: 52px; background-image: url(Image/login.jpg); layer-background-image: url(Image/login.jpg); border: 1px none #000000;">

  <div id="Layer2" style="position:absolute; width:415px; height:115px; z-index:1; left: 236px; top: 285px;">
    <p class="style2">登出成功，為安全起見，請在登出後將視窗關閉</p>
    <p class="style1"><br>
        <a href="Traffic_Login.asp" class="style3">重新登入智慧型交通執法管理系統
        </a></p>
  </div>
  </div>
</div>
</body>
<script language="JavaScript"> 
function countDown() {   
	location="Traffic_Login.asp";
}   
</Script>  
</html>
