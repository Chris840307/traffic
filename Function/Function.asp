<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!-- #include file="../Common/Bannernodata.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%

sqlGroup= "Select ID,Content from Code where TypeID=10 order by ShowOrder"
set RsGroup=Server.CreateObject("ADODB.RecordSet")
RsGroup.cursorlocation = 3
RsGroup.open sqlGroup,Conn,3,3

sqlSystem = "Select ID,Content from Code where TypeID=11 order by Content"
set RsSystem=Server.CreateObject("ADODB.RecordSet")
RsSystem.cursorlocation = 3
RsSystem.open sqlSystem,Conn,3,3

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>權限設定系統</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style3 {font-size: 15px}
.style5 {font-size: 15px; font-weight: bold; }
-->
</style></head>
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<!-- #include file="../Common/checkFunc.inc"-->
<Script language="JavaScript">
<!--	
function qryCheck()
{
	result=true; 
}	

function sendQry(){
	var rtn;
	rtn = qryCheck();
	if (rtn!=false){
     var form_A= document.forms[0];
     form_A.action = "Function.asp";
     form_A.submit();		
	}
}

function delFunction(param){
	 var rtn;
	 rtn = window.confirm("您確定要刪除此筆資料嗎?");
	 if (rtn!=false){
     openAddWindow(param,'delFunction');	
   }
}
-->
</Script>

<body>
<%
if Session("Msg")<>"" then
	 Response.write "<font  color='Red' size='3'>" & Session("Msg") & "</font>"
	 Session("Msg") = ""
end if	
%> 
<table width="100%" border="0">
  <tr>
    <td height="27" bgcolor="#1BF5FF"><span class="style3">權限設定系統</span></td>
  </tr>
<FORM NAME="Function" ACTION="" METHOD="POST">       
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td>
群組
      <select name="GroupID"  >
        <option value="" selected>選擇群組...</option>
		<%
			p = 0
			While Not RsGroup.Eof	
		%>
		
		 <option value="<%=RsGroup("ID")%>" <%
		 if trim(request("GroupID"))=trim(RsGroup("ID")) then
			response.write "selected"
		 end if
		 %>><%=RsGroup("Content")%></option>
         		<%
				p = p + 1
				RsGroup.MoveNext
			Wend
		%>          
      </select>
      系統
      <select name="SystemID"  >
        <option value="" selected>選擇系統...</option>
        <%
			p = 0
			While Not RsSystem.Eof
			%>
			 <option value="<%=RsSystem("ID")%>" <%
		 if trim(request("SystemID"))=trim(RsSystem("ID")) then
			response.write "selected"
		 end if
		 %>><%=RsSystem("Content")%></option>
			<%
		  		p = p + 1
		  		RsSystem.MoveNext
			Wend
		%>
       </select>  
      <img src="space.gif" width="13" height="8">      
      <input type="button" name="Submit" onclick="sendQry();" value="查詢"
          <% if CheckPermission(229,1)=false then  response.write "disabled"  end if %>
		>
     
      <input type="button" name="Submit2" value="群組權限設定" onclick="openAddWindow('FunctionAdd.asp?tag=new','AddFunction')"
	    <% if CheckPermission(229,2)=false then  response.write "disabled"  end if %>>
      <input type="button" name="Submit3" value="新增群組名稱" onclick="openAddWindow('FunctionGroupAdd.asp?tag=new','AddGroupFunction')">
        </span>      <span class="style3">
       
      </span>            </td>
      </tr>
    </table></td>
  </tr>
</FORM>
  <tr>
     <td height="26" bgcolor="#1BF5FF"><span class="pagetitle style3"><span class="style3">權限設定</span>紀錄列表</span></td>
  </tr>
<%
qryType = 0
SQL="Select a.* ,  b.CONTENT GROUPNAME , c.CONTENT SYSTEMNAME FROM FunctionDataDetail a, Code b , Code c " &_
		" Where a.GroupID=b.ID and a.SYStemID=c.ID "

	  if Request("SN")<>"" then
	     SQL = SQL & " And a.SN='" & Request("SN") & "' "
	     qryType = 1
	  end if
	  if Request("GroupID")<>"" then
	     SQL = SQL & " And a.GroupID='" & Request("GroupID") & "' "
	     qryType = 1
	  end if
	  if Request("SystemID")<>"" then
	     SQL = SQL & " And a.SystemID='" & Request("SystemID") & "' "
	     qryType = 1
	  end if	  
     SQL = SQL & " Order by GroupID , SystemID "  

Session("ExcelSql") = SQL
set Rs=Server.CreateObject("ADODB.RecordSet")
rs.cursorlocation = 3
rs.open SQL,Conn,3,3

if not rs.EOF then
	actionPage=cint(0 & trim(request("page"))) 
	if actionPage < 1 then actionPage=1
	rs.PageSize=PageSize
	if actionPage > rs.PageCount then actionPage=rs.PageCount
	rs.AbsolutePage=actionPage 
%>  
  <tr>
     <td height="80" bgcolor="#E0E0E0">
     	   <table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
              <tr bgcolor="#FAFAF5">
                <th width="20%" height="15" nowrap><span class="style3 style3 style3">群組名稱</span></th>
                <th width="40%" height="15" nowrap><span class="style3 style3 style3">系統名稱</span></th>
                <!--
                <td width="3" height="15" nowrap><span class="style5">新增</span></td>
                <td width="3" nowrap><span class="style3"><strong>修改</strong></span></td>
                <td width="3" height="15" nowrap><span class="style3"><strong>刪除</strong></span></td>
                <td width="3" nowrap><span class="style3"><strong>查詢</strong></span></td>
                -->
                <th width="20%" height="15" nowrap><span class="style3 style3 style3">操作</span></th>
              </tr>
	<%             
	for I=1 to rs.pagesize   
	%>              
              <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
                <td height="29" ><%=rs("GROUPNAME")%></td>
                <td ><%=rs("SYSTEMNAME")%></td>
                <!--
                <td width="3" nowrap ><%=rs("InsertFlag")%></td>
                <td width="3" nowrap ><%=rs("UpdateFlag")%></td>
                <td width="3" nowrap ><%=rs("DeleteFlag")%></td>
                <td width="3" nowrap ><%=rs("SelectFlag")%></td>				
                -->
				<td nowrap >
                    <input type="button" name="Submit433" value="修改"  <% if CheckPermission(229,3)=false then  response.write "disabled"  end if %>  onclick="openAddWindow('FunctionUpdate.asp?tag=upd&SN=<%=rs("SN")%>','updFunction')">
                  <input type="button" name="Submit3" value="刪除" <% if CheckPermission(229,4)=false then  response.write "disabled"  end if %>  onclick="delFunction('Function_mdy.asp?tag=del&SN=<%=rs("SN")%>&GroupID=<%=rs("GroupID")%>&SystemID=<%=rs("SystemID")%>')">
                </td>
              </tr>
	<%              
		rs.Movenext              
		If rs.EOF then exit for              
	next              
	%>              
         </table>
     </td>
  </tr>
	<tr>              
		<td align="center" height="35" bgcolor="#1BF5FF">
<%urlParam = "&SN=" & Request("SN") & "&GroupID=" & trim(request("GroupID")) & "&SystemID=" & trim(request("SystemID"))%>			
			<font size="2"><%ShowPageLink actionPage,rs.PageCount,"Function.asp",urlParam%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" name="SaveAs" value="轉換成Excel" onclick="openAddWindow('saveExcel.asp','saveExcel')"></td>              
	</tr>  

<% else %>    
  <tr>
  	 <td align="center" >        
	      <center><font  color="Red" size="2">              
	<%              
	Response.Write "目前查無任何資料 ..."              
	%>              
	      </font></center><br> 
    	  <p align="center">&nbsp;</p>    	  <p>&nbsp;</p>   	    <p>&nbsp;</p></td>
  </tr>             
<%              
end if              
rs.close              
set rs = nothing              
%>   
</table>
</body>

</html>
<!-- #include file="../Common/ClearObject.asp" -->