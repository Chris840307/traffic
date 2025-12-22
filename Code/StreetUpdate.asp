
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
<FORM NAME="updStreet" ACTION="Street_mdy.asp" METHOD="POST" onSubmit="return datacheck();">  	
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
           <%=Request("StreetID")%>
        </td>
      </tr>
          <input name="StreetSimpleName" type="hidden" value="<%=Request("StreetSimpleName")%>" size="10" maxlength="9" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
      <tr>
        <td bgcolor="#FFFFCC"><div align="right">路段詳細</div></td>
        <td>
          <input name="Address" type="text" size="41" maxlength="40" value="<%=Request("Address")%>" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
        </td>
		<%If sys_City="高雄市" Then%>
			  <tr>
				<td bgcolor="#FFFFCC"><div align="right">固定桿</div></td>
				<td>
			          <input name="FixPole" type="checkbox" value="1" <%If Request("FixPole")<>"" Then response.write "checked"%>>
				</td>
		<%End If%>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1"><a href="file:///C|/Documents%20and%20Settings/Smith/&#26700;&#38754;/&#31995;&#32113;&#35498;&#26126;/&#38936;&#21934;&#31649;&#29702;&#31995;&#32113;/sssss">
       </a>
        <input type="submit" name="submit001" value="確 定">
        <input type="hidden" name="StreetID" value="<%=Request("StreetID")%>">
        <img src="space.gif" width="9" height="8">       
        <input type="button" value="關 閉" onClick="window.close();">
      </p>    
    </td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    <p>　</p>
    <p>　</p></td></tr>
</table>
</body>
</html>
<!-- #include file="../Common/ClearObject.asp" -->