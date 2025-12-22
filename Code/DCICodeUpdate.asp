
<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>DCI代碼檔維護</title>
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
<FORM NAME="addDCICode" ACTION="DCICode_mdy.asp" METHOD="POST" >  	
	<input type="hidden" name="tag" value="<%=request("tag")%>">
	<input type="hidden" name="SN" value="<%=request("SN")%>">		
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">DCI代碼檔維護</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="11%" bgcolor="#FFFFCC"><div align="right" >代碼類型</div></td>
        <td width="89%">
           <select name="TypeId">
              <option value="1" <% if Request("TypeId")="1" Then Response.Write "selected" end if%>>刪除原因</option>
              <option value="2" <% if Request("TypeId")="2" Then Response.Write "selected" end if%>>舉發單類型</option>
              <option value="4" <% if Request("TypeId")="4" Then Response.Write "selected" end if%>>車輛顏色</option>
              <option value="5" <% if Request("TypeId")="5" Then Response.Write "selected" end if%>>DCI車種代號</option>
              <option value="6" <% if Request("TypeId")="6" Then Response.Write "selected" end if%>>扣件物品</option>
              <option value="7" <% if Request("TypeId")="7" Then Response.Write "selected" end if%>>退件原因</option>
              <option value="8" <% if Request("TypeId")="8" Then Response.Write "selected" end if%>>是否有保險證</option>
            </select>
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">代碼值</div></td>
        <td>
          <input name="Id" type="text" value="<%=Request("Id")%>" size="3" maxlength="2" onchange='javascript:this.innerText=this.value.toUpperCase();' class="btn1">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">代碼內容</div></td>
        <td>
          <input name="Content" type="text" size="21" maxlength="20" value="<%=Request("Content")%>" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1"><a href="file:///C|/Documents%20and%20Settings/Smith/&#26700;&#38754;/&#31995;&#32113;&#35498;&#26126;/&#38936;&#21934;&#31649;&#29702;&#31995;&#32113;/sssss">
    </a>
        <input type="submit" value="確 定">
        <img src="space.gif" width="9" height="8">  
        <input type="submit" onClick="window.close();" value="關 閉">
</p>    </td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    <p>&nbsp;</p>
    <p>&nbsp;</p></td></tr>
</table>
</Form>
</body>
</html>
