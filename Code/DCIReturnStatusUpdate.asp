
<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>DCI回覆結果維護</title>
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
<FORM NAME="updDCIReturnStatus" ACTION="DCIReturnStatus_mdy.asp" METHOD="POST" >  	
	<input type="hidden" name="tag" value="<%=request("tag")%>">
	<input type="hidden" name="SN" value="<%=request("SN")%>">	
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">DCI回覆結果維護</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="13%" bgcolor="#FFFFCC"><div align="right">DCI資料交換類型</div></td>
        <td width="87%">
          <select name="DCIActionID">
              <option value="A" <% if Request("DCIActionID")="A" Then Response.Write "selected" end if%>>查詢車籍</option>
              <option value="W" <% if Request("DCIActionID")="W" Then Response.Write "selected" end if%>>入案</option>
              <option value="WE" <% if Request("DCIActionID")="WE" Then Response.Write "selected" end if%>>入案錯誤</option>
              <option value="N" <% if Request("DCIActionID")="N" Then Response.Write "selected" end if%>>送達註記</option>
              <option value="E" <% if Request("DCIActionID")="E" Then Response.Write "selected" end if%>>刪除資料</option>
           </select>
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">回傳狀態代碼</div></td>
        <td>
          <input name="DCIReturn" type="text" value="<%=Request("DCIReturn")%>" size="4" maxlength="3" class="btn1">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">內容</div></td>
        <td>
          <input name="StatusContent" type="text" value="<%=Request("StatusContent")%>" size="16" maxlength="15" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">結果</div></td>
        <td>
          <select name="DCIReturnStatus">
            <option value='1' <% if Request("DCIReturnStatus")="1" Then Response.Write "selected" end if%>>正常</option>
            <option value='-1' <% if Request("DCIReturnStatus")="-1" Then Response.Write "selected" end if%>>異常</option>
          </select>
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">須再處理</div></td>
        <td>
          <select name="NeedReDo">
            <option value='1' <% if Request("NeedReDo")="1" Then Response.Write "selected" end if%>>是</option>
            <option value='0' <% if Request("NeedReDo")="0" Then Response.Write "selected" end if%>>否</option>
          </select>
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">方式</div></td>
        <td>
          <input name="HowTo" type="text" value="<%=Request("HowTo")%>" size="16" maxlength="15" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
        </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1"><a href="file:///C|/Documents%20and%20Settings/Smith/&#26700;&#38754;/&#31995;&#32113;&#35498;&#26126;/&#38936;&#21934;&#31649;&#29702;&#31995;&#32113;/sssss">
       </a>
        <input type="submit" value="確 定">
        <img src="space.gif" width="9" height="8"></span>        
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
