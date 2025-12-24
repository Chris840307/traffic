<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%

sqlGroup= "Select ID,Content from Code where TypeID=10"
set RsGroup=Server.CreateObject("ADODB.RecordSet")
RsGroup.open sqlGroup,Conn,3,3

sqlSystem = "Select ID,Content from Code where TypeID=11"
set RsSystem=Server.CreateObject("ADODB.RecordSet")
RsSystem.open sqlSystem,Conn,3,3

SQL="select * from FunctionDataDetail where SN=" & Request("SN") 
set RsUpd1=Server.CreateObject("ADODB.RecordSet")
RsUpd1.open SQL,Conn,3,3

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>權限設定系統-修改</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
function datacheck()
{
	var result ;
	
  if(document.all.GroupID.value=="")   
  {
    alert('請選擇群組');
    return false;  
  }	
  
  if(document.all.SystemID.value=="")   
  {
    alert('請選擇系統');
    return false;  
  }

}
-->
</Script>

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
<FORM NAME="updFunction" ACTION="Function_mdy.asp" METHOD="POST" onSubmit="return datacheck();">  
	<input type="hidden" name="tag" value="<%=request("tag")%>"> 
	<input type="hidden" name="SN" value="<%=request("SN")%>">

<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle style3">權限設定系統-修改</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="15%" height="33" bgcolor="#FFFFCC"><div align="right" class="style3">群組</div></td>
   
        <td width="85%">
          <select name="GroupID" id="GroupID" >
            <option value="" selected>選擇群組...</option>
            <%
			p = 0
			While Not RsGroup.Eof			
		%>
		
		    <option value="<%=RsGroup("ID")%>" <%if trim(RsUpd1("GroupID"))=trim(RsGroup("ID")) then response.write " selected" end if%>><%=RsGroup("Content")%></option>
            <%
				p = p + 1
				RsGroup.MoveNext
			Wend
		%>
          </select>
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right" class="style3 style3 style3">
            <div align="right" class="style3">
              <div align="right">系統</div>
            </div>
        </div></td>
        <td>
          <select name="SystemID" id="SystemID" >
            <option value="" selected>選擇系統...</option>
            <%
			p = 0
			While Not RsSystem.Eof
			%>
			 <option value="<%=RsSystem("ID")%>" <%if trim(RsUpd1("SYSTEMID"))=trim(RsSystem("ID")) then response.write " selected" end if%>><%=RsSystem("Content")%></option>
            <%
		  		p = p + 1
		  		RsSystem.MoveNext
			Wend
		%>
          </select>
        </td>
      </tr>
      <tr>
      <td>
      </td>
      <td>
      目前系統只提供單一權限設定
      </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right" class="style3 style3 style3">
            <div align="right">新增</div>
        </div></td>
        <td>
          <input name="insert" type="radio" value="1" <% if trim(RsUpd1("InsertFlag"))=1 then response.write  " checked" %>  disabled>
    可
    <input name="insert" type="radio" value="0" <% if trim(RsUpd1("InsertFlag"))=0 then response.write  " checked" %> disabled>
    否 </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC">
          <div align="right" class="style3 style3">
            <div align="right">修改</div>
        </div></td>
        <td><input name="update" type="radio" value="1" <% if trim(RsUpd1("updateFlag"))=1 then response.write  " checked" %> disabled>
    可
      <input name="update" type="radio" value="0" <% if trim(RsUpd1("updateFlag"))=0 then response.write  " checked" %> disabled>
    否
    <!--<input name="BillStartNumber" type="text" size="10" maxlength="9" onKeyDown='lockString(this);' onKeyUp='lockString(this);'>-->
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right" class="style3">刪除</div></td>
        <td>
          <input name="delete" type="radio" value="1" <% if trim(RsUpd1("deleteFlag"))=1 then response.write  " checked" %> disabled>
    可
    <input name="delete" type="radio" value="0" <% if trim(RsUpd1("deleteFlag"))=0 then response.write  " checked" %> disabled>
    否 </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right" class="style3">查詢</div></td>
        <td>
          <input name="select" type="radio" value="1" <% if trim(RsUpd1("selectFlag"))=1 then response.write  " checked" %> disabled>
    可
    <input name="select" type="radio" value="0" <% if trim(RsUpd1("selectFlag"))=0 then response.write  " checked" %> disabled>
    否 </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right" class="style3 style3">修改人員</div></td>
        <td><%=Session("Ch_Name")%></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1">
        <input type="submit" name="Submit423" value="確 定">
        <span class="style3"><img src="space.gif" width="9" height="8"></span>        <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉">
</p>    </td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    <p>&nbsp;</p>
    <p>&nbsp;</p></td></tr>
</table>
</FORM>
</body>
</html>
<!-- #include file="../Common/ClearObject.asp" -->
