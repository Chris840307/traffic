
<!-- #include file="..\Common\Util.asp"-->
<!-- #include file="..\Common\DbUtil.asp"-->
<!-- #include file="..\Common\AllFunction.inc"-->
<!-- #include file="..\Common\Login_Check.asp"-->
<!-- #include file="..\Common\bannernodata.asp" -->
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
.style2 {font-size: 18px}
.style3 {font-size: 15px}
-->
</style></head>
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<!-- #include file="..\Common\checkFunc.inc"-->
<Script language="JavaScript">
<!--	
function qryCheck()
{
	var form_A= document.forms[0];
	if ((form_A.ZipID.value=="") && (form_A.ZipID.value=="") && (form_A.ZipNo.value=="")){
	   alert("您尚未輸入任何查詢條件!!");
	   return false;	
	}
}
function sendQry(){
	var rtn;
	//rtn = qryCheck();
	//if (rtn!=false){
     var form_A= document.forms[0];
     form_A.action = "Zip.asp";
     form_A.submit();		
	//}
}
function delStreet(param){
	 var rtn;
	 rtn = window.confirm("您確定要刪除此筆資料嗎?");
	 if (rtn!=false){
     openAddGetBill(param,'Zip');	
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
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">郵遞區號管理</td>
  </tr>
<FORM NAME="Zip" ACTION="" METHOD="POST">    
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td>
        	郵遞區號
          <input name="ZipID" type="text" value="<%=Request("ZipID")%>" size="5" maxlength="4" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
          郵遞區名稱
          <input name="ZipName" type="text" value="<%=Request("ZipName")%>" size="10" maxlength="9" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">      
          分區綑綁區號
          <input name="ZipNo" type="text" value="<%=Request("ZipNo")%>" size="5" maxlength="4" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
          <input type="button" value="查詢" onclick="sendQry();" <%=ReturnPermission(CheckPermission(226,1))%>>
          <img src="space.gif" width="9" height="8">      
          <input type="button" value="新增" onclick="openAddGetBill('ZipAdd.asp?tag=new','ZipAdd')" <%=ReturnPermission(CheckPermission(226,2))%>>
        </td>
      </tr>
    </table></td>
  </tr>
</FORM>    
  <tr>
    <td height="26" bgcolor="#FFCC33"><span class="pagetitle">郵遞區號檔紀錄列表</span></td>
  </tr>
<%
   SQL = "Select * From Zip Where ZipID is not null "
   if Request("ZipID")<>"" then
      SQL = SQL & "And ZipID='" & Request("ZipID") & "' "
   end if
   if Request("ZipName")<>"" then
      SQL = SQL & " And (ZipName Like '%" & Request("ZipName") & "%') "
   end if	   
   if Request("ZipNo")<>"" then
      SQL = SQL & "And ZipNo='" & Request("ZipNo") & "' "
   end if	

   SQL = SQL & "Order By ZipID"
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
        <th width="15%" height="15" nowrap>郵遞區號</th>
        <th width="27%" height="15" nowrap>郵遞區名稱</th>
        <th width="35%" nowrap>分區綑綁區號</th>
        <th width="23%" height="15" nowrap>操作</th>
      </tr>
	<%             
	for I=1 to Rs.pagesize   
	%>       
      <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td ><div align="center"><%=Rs("ZipId")%></div></td>
        <td ><div align="center"><%=Rs("ZipName")%></div></td>
        <td><div align="center"><%=Rs("ZipNo")%></div></td>
        <td >
          <input type="button" value="修改" onclick="openAddGetBill('ZipUpdate.asp?tag=upd&ZipID=<%=Rs("ZipID")%>&ZipName=<%=Rs("ZipName")%>&ZipNo=<%=Rs("ZipNo")%>','UpdateZip')" <%=ReturnPermission(CheckPermission(226,3))%>>&nbsp;&nbsp;&nbsp;&nbsp;
          <input type="button" value="刪除" onclick="delStreet('Zip_mdy.asp?tag=del&ZipID=<%=rs("ZipID")%>');" <%=ReturnPermission(CheckPermission(226,4))%>>
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
   urlParam = "&ZipID=" & Request("ZipID") & "&ZipName=" & Request("ZipName") & "&ZipNo=" & Request("ZipNo")
%>			
			<font size="2"><%ShowPageLink actionPage,Rs.PageCount,"Zip.asp",urlParam%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" name="SaveAs" value="轉換成Excel" onclick="exportExcel('ZipExcel.asp','ZipExcel')">
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