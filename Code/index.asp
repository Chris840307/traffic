<!-- #include file="..\Common\db.ini" -->
<!-- #include file="..\Common\AllFunction.inc" -->
<!-- #include file="..\Common\Login_Check.asp"-->
<!-- #include file="..\Common\bannernodata.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>代碼維護</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {font-size: 14px}
.style3 {font-size: 15px}
.style4 {color: #CC9900}
-->
</style></head>
<%

If Session("User_ID")="" Then
	Response.write "系統Session值已過期,請重新登入!"
  Response.End
End If	
%>

<body>
<table width="100%" height="100%" border="0">
  <tr>
    <td height="26" bgcolor="#FFCC33"><span class="style3">代碼維護主頁</span></td>
  </tr>
  <tr>
    <td height="385" bgcolor="#E0E0E0"><table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
      <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td>　</td>      
		<td height="10"  class="style3"><img src="../Image/btn.gif"><a href="Project.asp"><span class="pagetitle">專案資料檔維護</span></a></Img></td>
      </tr>
      <!--<tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td>　</td>
        <td height="10"  class="style3"><img src="../Image/btn.gif"><a href="FixEquip.asp"><span class="pagetitle">固定桿資料維護</span></a></Img></td>
      </tr>-->
     <!-- <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td>　</td>
        <td height="10"  class="style3"><img src="../Image/btn.gif"><a href="code.asp"><span class="pagetitle">系統代碼檔維護</span></a></Img></td>
      </tr>-->
     <!--  <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td>　</td>
        <td height="10"  class="style3"><img src="../Image/btn.gif"><a href="LawOftenUse.asp"><span class="pagetitle"><span class="pagetitle">常用法條代碼檔維護</span></Img></td>
      </tr>-->
      <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td>　</td>
        <td height="10"  class="style3"><img src="../Image/btn.gif"><a href="Law.asp"><span class="pagetitle">法條代碼檔維護</span></Img></td>
      </tr>
      <!--
      <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td>　</td>
        <td height="10"  class="style3"><img src="../Image/btn.gif"><a href="Zip.asp"><span class="pagetitle">郵遞區號檔維護</span></Img></td>
      </tr>
      <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td>　</td>
        <td height="10"  class="style3"><img src="../Image/btn.gif"><a href="Station.asp"><span class="pagetitle">監理所站代碼檔</span></Img></td>
      </tr>
      -->
      <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td>　</td>
        <td height="10"  class="style3"><img src="../Image/btn.gif"><a href="Street.asp"><span class="pagetitle">縣市路段代碼檔</span></Img></td>
      </tr>
      <!--
      <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td>　</td>
        <td height="10"  class="style3"><img src="../Image/btn.gif"><a href="DCIReturnStatus.asp"><span class="pagetitle">DCI回傳代碼值</span></Img></td>
      </tr>
      <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td>　</td>
        <td height="10"  class="style3"><img src="../Image/btn.gif"><a href="DCICode.asp"><span class="pagetitle">DCI代碼表</span></Img></td>
      </tr>
      -->
	  <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td>　</td>
        <td height="10"  class="style3"><img src="../Image/btn.gif"><a href="OptionCarSpeed.asp"><span class="pagetitle">特殊車輛車速設定</span></Img></td>
      </tr>
     <% if Session("Group_ID")="200" then %>
	  <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td></td>
        <td height="10"  class="style3"><img src="../Image/btn.gif"><a href="notice.asp"><span class="pagetitle">公告訊息維護</span></td>
      </tr>
     <% end if%>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1"><a href="file:///C|/Documents%20and%20Settings/Smith/&#26700;&#38754;/&#31995;&#32113;&#35498;&#26126;/&#38936;&#21934;&#31649;&#29702;&#31995;&#32113;/sssss">
      </a></p>    </td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    <p>　</p>
    <p>　</p></td></tr>
</table>
</body>
</html>
