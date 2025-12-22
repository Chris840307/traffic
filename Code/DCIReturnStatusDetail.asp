
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
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="style3">DCI回覆結果維護</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="13%" bgcolor="#FFFFCC"><div align="right"><span class="style3">DCI資料交換類型</span></div></td>
        <td width="87%"><span class="style3">
          <select name="select2">
              <option value="A" selected>查詢車籍</option>
              <option value="W">入案</option>
              <option value="WW">入案錯誤</option>
              <option value="N">送達註記</option>
              <option value="E">刪除資料</option>
            </select>
</span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">回傳狀態代碼</span></div></td>
        <td><span class="style3">
          <input name="textfield42322" type="text" value="S" size="4" maxlength="3">
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">內容</span></div></td>
        <td><span class="style3">
          <input name="textfield423222" type="text" value="查詢成功" size="16" maxlength="15">
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">結果</span></div></td>
        <td><span class="style3">
          <select name="select3">
            <option selected>正常</option>
            <option>異常</option>
          </select>
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">須再處理</span></div></td>
        <td><span class="style3">
          <select name="select">
            <option>是</option>
            <option selected>否</option>
          </select>
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">方式</span></div></td>
        <td><span class="style3">
          <input name="textfield4232222" type="text" size="16" maxlength="15">
        </span></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1"><a href="file:///C|/Documents%20and%20Settings/Smith/&#26700;&#38754;/&#31995;&#32113;&#35498;&#26126;/&#38936;&#21934;&#31649;&#29702;&#31995;&#32113;/sssss">
    </a>
        <input type="submit" name="Submit423" value="確 定">
        <span class="style3"><img src="space.gif" width="9" height="8"></span>        <input type="submit" name="Submit4232" value="關 閉">
</p>    </td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    <p>&nbsp;</p>
    <p>&nbsp;</p></td></tr>
</table>
</body>
</html>
