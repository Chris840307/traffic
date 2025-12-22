<!-- #include file="..\Common\Util.asp"-->
<!-- #include file="..\Common\DbUtil.asp"-->
<!-- #include file="..\Common\AllFunction.inc"-->
<!-- #include file="..\Common\Login_Check.asp"-->

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
     form_A.action = "DCIReturnStatus.asp";
     form_A.submit();		
	//}
}
function delDCIReturnStatus(param){
	 var rtn;
	 rtn = window.confirm("您確定要刪除此筆資料嗎?");
	 if (rtn!=false){
     openAddGetBill(param,'DCIReturnStatus');	
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
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">DCI回覆結果維護</span></td>
  </tr>
<FORM NAME="DCIReturnStatus" ACTION="" METHOD="POST">   
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td>DCI資料交換類型
            <select name="DCIActionID">          	
              <option value="">全部</option>
              <option value="A" <% if Request("DCIActionID")="A" Then Response.Write "selected" end if%>>查詢車籍</option>
              <option value="W" <% if Request("DCIActionID")="W" Then Response.Write "selected" end if%>>入案</option>
              <option value="WE" <% if Request("DCIActionID")="WE" Then Response.Write "selected" end if%>>入案錯誤</option>
              <option value="N" <% if Request("DCIActionID")="N" Then Response.Write "selected" end if%>>送達註記</option>
              <option value="E" <% if Request("DCIActionID")="E" Then Response.Write "selected" end if%>>刪除資料</option>
            </select>      
            回傳狀態代碼
            <input name="DCIReturn" type="text" value="<%=Request("DCIReturn")%>" size="4" maxlength="3" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
            內容
            <input name="StatusContent" type="text" value="<%=Request("StatusContent")%>" size="11" maxlength="10" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">            
            結果
            <select name="DCIReturnStatus">
              <option value="">全部</option>            	
              <option value='1' <% if Request("DCIReturnStatus")="1" Then Response.Write "selected" end if%>>正常</option>
              <option value='-1' <% if Request("DCIReturnStatus")="-1" Then Response.Write "selected" end if%>>異常</option>
            </select>
            須再處理
            <select name="NeedReDo">
              <option value="">全部</option>            	
              <option value='1' <% if Request("NeedReDo")="1" Then Response.Write "selected" end if%>>是</option>
              <option value='0' <% if Request("NeedReDo")="0" Then Response.Write "selected" end if%>>否</option>
            </select>
            方式
             <input name="HowTo" type="text" size="16" maxlength="15" value="<%=Request("HowTo")%>" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
             <input type="button" onclick="sendQry();" value="查詢" <%=ReturnPermission(CheckPermission(226,1))%>>
             <img src="space.gif" width="9" height="8">      
             <input type="button" value="新增" onclick="openAddGetBill('DCIReturnStatusAdd.asp?tag=new','DCIReturnStatusAdd')" <%=ReturnPermission(CheckPermission(226,2))%>>
        </td>
      </tr>
    </table></td>
  </tr>
</FORM>  
  <tr>
    <td height="26" bgcolor="#FFCC33"><span class="pagetitle">DCI回覆結果檔紀錄列表</span></td>
  </tr>
<%
SQL = "Select SN,DCIActionID,DCIActionName,DCIReturn,DCIReturnStatus,StatusContent,NeedReDo,HowTo," & _
      "Decode(DCIReturnStatus,1,'正常',-1,'異常') As DCIReturnStatusDesc,Decode(NeedReDo,0,'是',1,'否') As NeedReDoDesc " & _
      "From DciReturnStatus Where Sn is not null "
if Request("DCIActionID")<>"" then
   SQL = SQL & "And DCIActionID='" & Request("DCIActionID") & "' "
end if
if Request("DCIReturn")<>"" then
   SQL = SQL & "And DCIReturn='" & Request("DCIReturn") & "' "
end if	
if Request("StatusContent")<>"" then
   SQL = SQL & " And (StatusContent Like '%" & Request("StatusContent") & "%') "
end if	
if Request("DCIReturnStatus")<>"" then
   SQL = SQL & "And DCIReturnStatus=" & Int(Request("DCIReturnStatus")) & " "
end if
if Request("NeedReDo")<>"" then
   SQL = SQL & "And NeedReDo=" & Int(Request("NeedReDo")) & " "
end if	 
if Request("HowTo")<>"" then
   SQL = SQL & " And (HowTo Like '%" & Request("HowTo") & "%') "
end if

SQL = SQL & "Order By DCIActionID"
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
        <th width="8%" height="15" nowrap>資料交換類型</th>
        <th width="4%" height="15" nowrap>狀態代碼</th>
        <th width="25%" nowrap>內容</th>
        <th width="4%" nowrap>結果</th>
        <th width="4%" nowrap>須再處理</th>
        <th width="35%" height="15" nowrap>方式</th>
        <th width="20%" height="15" nowrap>操作</th>
      </tr>
	<%             
	for I=1 to Rs.pagesize   
	%>       
      <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td>
          <div align="right"><%=Rs("DCIActionName")%></div>
        </td>
        <td>
        	<div align="right"><%=Rs("DCIReturn")%></div>
        </td>
        <td>
        	<div align="right"><%=Rs("StatusContent")%></div>
        </td>
        <td>
        	<div align="right"><%=Rs("DCIReturnStatusDesc")%></div>
        </td>
        <td>
          <div align="right"><%=Rs("NeedReDoDesc")%></div>
        </td>
        <td>
        	<div align="center"><%=Rs("HowTo")%></div>
        </td>
        <td>
              <input type="button" value="修改" onclick="openAddGetBill('DCIReturnStatusUpdate.asp?tag=upd&SN=<%=Rs("SN")%>&DCIActionID=<%=Rs("DCIActionID")%>&DCIReturn=<%=Rs("DCIReturn")%>&StatusContent=<%=Rs("StatusContent")%>&DCIReturnStatus=<%=Rs("DCIReturnStatus")%>&NeedReDo=<%=Rs("NeedReDo")%>&HowTo=<%=Rs("HowTo")%>','UpdateDCICode')" <%=ReturnPermission(CheckPermission(226,3))%>>&nbsp;&nbsp;&nbsp;&nbsp;
              <input type="button" value="刪除" onclick="delDCIReturnStatus('DCIReturnStatus_mdy.asp?tag=del&SN=<%=rs("SN")%>');" <%=ReturnPermission(CheckPermission(226,4))%>>
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
   urlParam = "&DCIActionID=" & Request("DCIActionID") & "&DCIReturn=" & Request("DCIReturn") & "&StatusContent=" & Request("StatusContent") & "&DCIReturnStatus=" & Request("DCIReturnStatus") & "&NeedReDo=" & Request("NeedReDo") & "&HowTo=" & Request("HowTo")
%>			
			<font size="2"><%ShowPageLink actionPage,Rs.PageCount,"DCIReturnStatus.asp",urlParam%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" name="SaveAs" value="轉換成Excel" onclick="exportExcel('DCIReturnStatusExcel.asp','DCIReturnStatusExcel')">
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
