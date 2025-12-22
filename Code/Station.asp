
<!-- #include file="..\Common\Util.asp"-->
<!-- #include file="..\Common\DbUtil.asp"-->
<!-- #include file="..\Common\AllFunction.inc"-->
<!-- #include file="..\Common\Login_Check.asp"-->

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
.style2 {font-size: 18px}
.style3 {font-size: 15px}
-->
</style></head>
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<!-- #include file="..\Common\checkFunc.inc"-->
<!-- #include file="..\Common\bannernodata.asp" -->
<Script language="JavaScript">
<!--	
function sendQry(){
	var rtn;
	//rtn = qryCheck();
	//if (rtn!=false){
     var form_A= document.forms[0];
     form_A.action = "Station.asp";
     form_A.submit();		
	//}
}
function delStation(param){
	 var rtn;
	 rtn = window.confirm("您確定要刪除此筆資料嗎?");
	 if (rtn!=false){
     openAddGetBill(param,'Station');	
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
<body>
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">監理站資料檔維護</span><span class="style2"><span class="style3">    </span></span></td>
  </tr>
<FORM NAME="Station" ACTION="" METHOD="POST">  
	<input type="hidden" name="isQuery" value="y">	
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td>
        	到案監理站名稱
            <input name="StationName" value="<%=Request("StationName")%>" type="text" size="16" maxlength="15" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
          代碼
            <input name="StationID" value="<%=Request("StationID")%>" type="text" size="5" maxlength="4" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">      
          電話
            <input name="StationTel" value="<%=Request("StationTel")%>" type="text" size="10" maxlength="9" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
          住址
            <input name="StationAddress" value="<%=Request("StationAddress")%>" type="text" size="21" maxlength="40" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
          DCI回傳名稱
            <input name="DCIStationID" value="<%=Request("DCIStationID")%>" type="text" size="16" maxlength="15" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
          代碼
            <input name="DCIStationName" value="<%=Request("DCIStationName")%>" type="text" size="6" maxlength="5" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
          <img src="space.gif" width="9" height="8">
          <input type="submit" value="查詢" onclick="sendQry();" <%=ReturnPermission(CheckPermission(226,1))%>>
          <img src="space.gif" width="9" height="8">      
          <input type="submit" value="新增" onclick="openAddGetBill('StationAdd.asp?tag=new','StationAdd')" <%=ReturnPermission(CheckPermission(226,2))%>>
        </td>
      </tr>
    </table></td>
  </tr>
</FORM>   
  <tr>
    <td height="26" bgcolor="#FFCC33"><span class="pagetitle">監理站資料檔紀錄列表</span></td>
  </tr>
<%
if Request("isQuery") = "y" then
    SQL = "Select * From Station Where StationSN is not null "
    if Request("StationName")<>"" then
       SQL = SQL & "And StationName Like '%" & Request("StationName") & "%' "
    end if	
    if Request("StationID")<>"" then
       SQL = SQL & " And (StationID = '" & Request("StationID") & "') "
    end if	
    if Request("StationTel")<>"" then
       SQL = SQL & "And StationTel=" & Int(Request("StationTel")) & " "
    end if
    if Request("StationAddress")<>"" then
       SQL = SQL & "And StationAddress Like '%" & Request("StationAddress") & "%' "
    end if	 
    if Request("DCIStationName")<>"" then
       SQL = SQL & "And DCIStationName Like '%" & Request("DCIStationName") & "%' "
    end if
    if Request("DCIStationID")<>"" then
       SQL = SQL & "And DCIStationID=" & Int(Request("DCIStationID")) & " "
    end if    

    Session("ExcelSql") = SQL
    set Rs=Server.CreateObject("ADODB.RecordSet")
    Rs.cursorlocation = 3
    Rs.open SQL,Conn,3,3
    
    if not Rs.EOF then
    	actionPage=cint(0 & trim(request("page"))) 
    	if actionPage < 1 then actionPage=1
    	Rs.PageSize=PageSize
    	if actionPage > Rs.PageCount then actionPage=Rs.PageCount
    	Rs.AbsolutePage=actionPage 
%>   
  <tr>
    <td bgcolor="#E0E0E0"><table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
      <tr bgcolor="#EBFBE3">
        <th width="15%" height="15" nowrap>到案監理站名稱</th>
        <th width="5%" height="15" nowrap>代碼</th>
        <th width="9%" nowrap>電話</th>
        <th width="32%" nowrap>住址</th>
        <th width="9%" nowrap>DCI回傳名稱</th>
        <th width="5%" height="15" nowrap>代碼</th>
        <th width="25%" height="15" nowrap>操作</th>
      </tr>
	<%             
	for I=1 to Rs.pagesize   
	%>       
      <tr bgcolor="#FFFFFF">
        <td height="23"><div align="center"><%=Rs("StationName")%></div></td>
        <td height="23"><div align="right"><%=Rs("StationID")%></div></td>
        <td><div align="right"><%=Rs("StationTel")%></div></td>
        <td><div align="right"><%=Rs("StationAddress")%></div></td>
        <td><div align="right"><%=Rs("DCIStationName")%></div></td>
        <td height="23"><div align="right"><%=Rs("DCIStationID")%></div></td>
        <td height="23">
           <input type="button" value="修改" onclick="openAddGetBill('StationUpdate.asp?tag=upd&StationSN=<%=Rs("StationSN")%>&DCIStationID=<%=Rs("DCIStationID")%>&DCIStationName=<%=Rs("DCIStationName")%>&StationName=<%=Rs("StationName")%>&StationID=<%=Rs("StationID")%>&StationTel=<%=Rs("StationTel")%>&StationAddress=<%=Rs("StationAddress")%>','UpdateStation')" <%=ReturnPermission(CheckPermission(226,3))%>>&nbsp;&nbsp;&nbsp;&nbsp;
           <input type="button" value="刪除" onclick="delStation('Station_mdy.asp?tag=del&StationSN=<%=rs("StationSN")%>');" <%=ReturnPermission(CheckPermission(226,4))%>>
        </td>
      </tr>
	<%              
		Rs.Movenext              
		If Rs.EOF then exit for              
	next              
	%>              
         </table>
     </td>
  </tr>
	<tr>              
		<td align="center" height="35" bgcolor="#FFDD77">
<%
   urlParam = "&isQuery=" & Request("isQuery") & "&StationSN=" & Request("StationSN") & "&DCIStationID=" & Request("DCIStationID") & "&DCIStationName=" & Request("DCIStationName") & "&StationName=" & Request("StationName") & "&StationID=" & Request("StationID") & "&StationTel=" & Request("StationTel") & "&StationAddress=" & Request("StationAddress")
%>			
			<font size="2"><%ShowPageLink actionPage,Rs.PageCount,"Station.asp",urlParam%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" name="SaveAs" value="轉換成Excel" onclick="exportExcel('StationExcel.asp','StationExcel')">
		  <input type="button" value="回到前一頁" onClick="window.location.href='index.asp'">
		</td>              
	</tr>  

<% else %>    
  <tr>
  	 <td align="center" >        
	      <center><font  color="Red" size="2">              
	<%              
	Response.Write "目前查無任何資料 ..."              
	%>              
	      </font></center><br> 
	   </td>
	</tr>             
<%              
   end if              
   Rs.close              
   set Rs = nothing   
end if           
%>   
  <tr>
    <td>    
    	 <p align="center">　</p>    <p>　</p>
       <p>　</p>
    </td>
  </tr>
</table>
</body>
</html>
<!-- #include file="../Common/ClearObject.asp" -->
