
<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>路段代碼檔維護</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {font-size: 14px}
.style3 {font-size: 15px}
-->
</style></head>
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<!-- #include file="../Common/checkFunc.inc"-->
<SCRIPT LANGUAGE=javascript>
<!--
function datacheck()
{

  if(document.all.StreetId.value=="")   
  {
    alert('請輸入【路段代碼】!!');
    document.all.StreetId.focus();
    return false;  
  }
  if(document.all.Address.value=="")   
  {
    alert('請輸入【詳細路段】!!');
    document.all.Address.focus();
    return false;  
  }  	
}
-->
</Script>
<body>
<%
if Session("Msg")<>"" then
	 Response.write "<font  color='Red' size='2'>" & Session("Msg") & "</font>"
	 Session("Msg") = ""
end if	
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
%>
<FORM NAME="addStreet" ACTION="Street_mdy.asp" METHOD="POST" onSubmit="return datacheck();">  	
	<input type="hidden" name="tag" value="<%=request("tag")%>">

<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">路段代碼檔維護</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="11%" bgcolor="#FFFFCC"><div align="right">路段代碼</div></td>
        <td width="89%">
           <input name="StreetId" type="text" size="10" maxlength="9" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
        </td>
      </tr>

          <input name="StreetSimpleName" type="hidden" size="10" maxlength="50" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">

      <tr>
        <td bgcolor="#FFFFCC"><div align="right">路段詳細</div></td>
        <td>
          <input name="Address" type="text" size="41" maxlength="50" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
        </td>
		<%If sys_City="高雄市" Then%>
			  <tr>
				<td bgcolor="#FFFFCC"><div align="right">固定桿</div></td>
				<td>
			          <input name="FixPole" type="checkbox" value="1">
				</td>
		<%End if%>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1">
    
        <input type="submit" value="確 定" name="submit001">
        <img src="space.gif" width="9" height="8">       
        <input type="button" value="關 閉" onClick="window.close();">
      </p>    
    </td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    <p>&nbsp;</p>
    <p>&nbsp;</p></td></tr>
</table>
</Form>
</body>
</html>
<!-- #include file="../Common/ClearObject.asp" -->
