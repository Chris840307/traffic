<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body onLoad="setTimeout(countDown,5000);">
    <form id="form1" runat="server">
    <div id="Layer2" style="position:absolute; width:415px; height:115px; z-index:1; left: 236px; top: 285px;">
    <p class="style2">登出成功，為安全起見，請在登出後將視窗關閉</p>
    <p class="style1"><br>
        <a href="Login.aspx" class="style3">重新登入績效獎勵金試算系統
        </a></p>
    </div>
    </form>
</body>
<script language="JavaScript"> 
function countDown() {   
	location="Login.aspx";
}   
</Script>  
</html>
