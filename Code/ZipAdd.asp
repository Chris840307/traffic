
<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>郵遞區號檔維護</title>
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
  if(document.all.ZipID.value=="")   
  {
    alert('請輸入【郵遞區號】!!');
    document.all.ZipID.focus();
    return false;  
  }  	
  if(document.all.ZipName.value=="")   
  {
    alert('請輸入【郵遞區名稱】!!');
    document.all.ZipName.focus();
    return false;  
  }  	
  if(document.all.ZipNo.value=="")   
  {
    alert('請輸入【分區綑綁區號】!!');
    document.all.ZipNo.focus();
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
%>
<FORM NAME="addZip" ACTION="Zip_mdy.asp" METHOD="POST" onSubmit="return datacheck();">  	
	<input type="hidden" name="tag" value="<%=request("tag")%>">	
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">郵遞區號檔維護</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="11%" bgcolor="#FFFFCC"><div align="right" >郵遞區號</div></td>
        <td width="89%">
           <input name="ZipID" type="text" value="<%=Request("ZipID")%>" size="5" maxlength="4" onKeyDown="lockNum(this);" onKeyUp="lockNum(this);" class="btn1">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">郵遞區名稱</div></td>
        <td>
          <input name="ZipName" type="text" value="<%=Request("ZipName")%>" size="10" maxlength="9" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">分區綑綁區號</td>
        <td>
          <input name="ZipNo" value="<%=Request("ZipNo")%>" type="text" value="241" size="5" maxlength="4" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1"><a href="file:///C|/Documents%20and%20Settings/Smith/&#26700;&#38754;/&#31995;&#32113;&#35498;&#26126;/&#38936;&#21934;&#31649;&#29702;&#31995;&#32113;/sssss">
      </a>
        <input type="submit" value="確 定">
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
</body>
</html>
<!-- #include file="../Common/ClearObject.asp" -->
