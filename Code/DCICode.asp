<!-- #include file="..\Common\Util.asp"-->
<!-- #include file="..\Common\DbUtil.asp"-->
<!-- #include file="..\Common\AllFunction.inc"-->
<!-- #include file="..\Common\Login_Check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; chaRset=big5">
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
function qryCheck()
{
	var form_A= document.forms[0];
	if ((form_A.TypeId.value=="") && (form_A.Id.value=="") && (form_A.Content.value=="")){
	   alert("您尚未選擇任何查詢條件!!");
	   return false;	
	}
}	
function sendQry(){
	var rtn;
	//rtn = qryCheck();
	//if (rtn!=false){
     var form_A= document.forms[0];
     form_A.action = "DCICode.asp";
     form_A.submit();		
	//}
}
function delDCICode(param){
	 var rtn;
	 rtn = window.confirm("您確定要刪除此筆資料嗎?");
	 if (rtn!=false){
     openAddGetBill(param,'DelDCICode');	
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
<table width="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">DCI代碼檔維護</span></td>
  </tr>
<FORM NAME="DCICode" ACTION="" METHOD="POST">   
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td>代碼類型
            <select name="TypeId">
                <option value="">全部</option>          	
                <option value="1" <% if Request("TypeId")="1" Then Response.Write "selected" end if%>>刪除原因</option>
                <option value="2" <% if Request("TypeId")="2" Then Response.Write "selected" end if%> >舉發單類型</option>
                <option value="4" <% if Request("TypeId")="4" Then Response.Write "selected" end if%>>車輛顏色</option>
                <option value="5" <% if Request("TypeId")="5" Then Response.Write "selected" end if%>>DCI車種代號</option>
                <option value="6" <% if Request("TypeId")="6" Then Response.Write "selected" end if%>>扣件物品</option>
                <option value="7" <% if Request("TypeId")="7" Then Response.Write "selected" end if%>>退件原因</option>
                <option value="8" <% if Request("TypeId")="8" Then Response.Write "selected" end if%>>是否有保險證</option>
             </select>      
             代碼值&nbsp;<input name="Id" type="text" size="3" maxlength="2" value="<%=Request("Id")%>" onchange='javascript:this.innerText=this.value.toUpperCase();' class="btn1">
             代碼內容&nbsp;<input name="Content" type="text" value="<%=Request("Content")%>" size="21" maxlength="20" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
             <input type="button" onclick="sendQry();" name="Submit" value="查詢" <%=ReturnPermission(CheckPermission(226,1))%>>
             <img src="space.gif" width="9" height="8">      
             <input type="button" name="Submit2" value="新增" onclick="openAddGetBill('DCICodeAdd.asp?tag=new','AddDCICode')" <%=ReturnPermission(CheckPermission(226,2))%>>
        </td>
      </tr>
    </table></td>
  </tr>
</FORM>  
  <tr>
    <td height="26" bgcolor="#FFCC33"><span class="pagetitle">DCI代碼檔紀錄列表</span></td>
  </tr>
<%
SQL = "Select * From DciCode Where Sn is not null "
if Request("TypeId")<>"" then
   SQL = SQL & "And TypeId='" & Request("TypeId") & "' "
end if
if Request("Id")<>"" then
   SQL = SQL & "And Id='" & Request("Id") & "' "
end if	
if Request("Content")<>"" then
   SQL = SQL & "And Content Like '%" & Request("Content") & "%' "
end if  
SQL = SQL & "Order By Sn"

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
        <th width="9%" height="15" nowrap><span class="style3">代碼類型</span></th>
        <th width="7%" height="15" nowrap><span class="style3">代碼值</span></th>
        <th width="53%" height="15" nowrap><span class="style3">代碼內容</span></th>
        <th width="31%" height="15" nowrap><span class="style3">操作</span></th>
      </tr>
	<%             
	for I=1 to Rs.pagesize   
	   TypeDesc = GetDciTypeById (Rs("TypeId"))
	%>       
      <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
         <td nowrap><div align="right"><%=TypeDesc%></div></td>
         <td nowrap><div align="right"><%=Rs("Id")%></div></td>
         <td nowrap><div align="left"></div>           
           <div align="left"><%=Rs("Content")%></div></td>
         <td>
         	  <input type="button" value="修改" onclick="openAddGetBill('DCICodeUpdate.asp?tag=upd&SN=<%=Rs("SN")%>&TypeId=<%=Rs("TypeId")%>&Id=<%=Rs("Id")%>&Content=<%=Rs("Content")%>','UpdateDCICode')" <%=ReturnPermission(CheckPermission(226,3))%>>&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value="刪除" onclick="delDCICode('DCICode_mdy.asp?tag=del&SN=<%=rs("SN")%>');" <%=ReturnPermission(CheckPermission(226,4))%>>
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
   urlParam = "&TypeId=" & Request("TypeId") & "&Id=" & Request("Id") & "&Content=" & Request("Content")
%>			
			<font size="2"><%ShowPageLink actionPage,Rs.PageCount,"DCICode.asp",urlParam%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" name="SaveAs" value="轉換成Excel" onclick="exportExcel('DCICodeExcel.asp','DCICodeExcel')">
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
</table>
</body>
</html>
<!-- #include file="../Common/ClearObject.asp" -->
