
<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>監理站資料檔維護</title>
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
<body>
<%
if Session("Msg")<>"" then
	 Response.write "<font  color='Red' size='2'>" & Session("Msg") & "</font>"
	 Session("Msg") = ""
end if	
%>		
<FORM NAME="addStation" ACTION="Station_mdy.asp" METHOD="POST" >  	
	<input type="hidden" name="tag" value="<%=request("tag")%>">
	<input type="hidden" name="StationSN" value="<%=request("StationSN")%>">			
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">監理站資料檔維護</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="14%" bgcolor="#FFFFCC"><div align="right">到案監理站名稱</div></td>
        <td width="86%">
          <input name="StationName" value="<%=Request("StationName")%>" type="text" size="16" maxlength="15" class="btn1">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">代碼</div></td>
        <td>
          <input name="StationID" value="<%=Request("StationID")%>" type="text" size="16" maxlength="15" class="btn1">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">電話</div></td>
        <td>
          <input name="StationTel" value="<%=Request("StationTel")%>" type="text" size="16" maxlength="15" class="btn1">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">住址</div></td>
        <td>
          <input name="StationAddress" value="<%=Request("StationAddress")%>" type="text" size="41" maxlength="40" class="btn1">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">DCI回傳監理站名稱</div></td>
        <td>
          <input name="DCIStationName" value="<%=Request("DCIStationName")%>" type="text" size="16" maxlength="15" class="btn1">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">代碼</div></td>
        <td>
          <input name="DCIStationID" value="<%=Request("DCIStationID")%>" type="text" size="16" maxlength="15" class="btn1">
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1"><a href="file:///C|/Documents%20and%20Settings/Smith/&#26700;&#38754;/&#31995;&#32113;&#35498;&#26126;/&#38936;&#21934;&#31649;&#29702;&#31995;&#32113;/sssss">
        </a>
        <input type="submit" value="確 定">
        <img src="space.gif" width="9" height="8">
        <input type="button" onClick="window.close();" value="關 閉">
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
