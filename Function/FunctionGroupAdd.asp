<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!-- #include file="../Common/Bannernodata.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%

if Trim(request("DB_Selt"))="DB_Insert" then 

  sqlGroup= "Select ID,Content from Code where TypeID=10 and Content='"&trim(request("GroupName"))&"'"
set Rs=Server.CreateObject("ADODB.RecordSet")
Rs.open sqlGroup,Conn,3,3
if Rs.eof then 

  strSQL="Select max(id)+1 as S1 from code"
  set rssysinfo=conn.execute(strSQL)
  S1=rssysinfo("S1") 
  set rssysinfo=nothing

  strSQL="select (max(showorder)+1) as S2 from code where typeid=10"
  set rssysinfo=conn.execute(strSQL)
  S2=rssysinfo("S2") 
  set rssysinfo=nothing

          strIns="insert into Code values(" & S1 & ",10,'"&trim(request("GroupName"))&"'," & S2 & ",0,0,0,0)"

		  conn.execute strIns

  else
    response.write "<script>"
	response.write "alert(""該群組名稱已建立"");"
	response.write "</script>"
  end if
Rs.close
set rs=nothing 
end if

  if Trim(request("DB_Selt"))="DB_Delete" then 

          strIns="Delete Code where ID="& request("DB_ID")

		  conn.execute strIns
  end if

sqlGroup= "Select ID,Content from Code where TypeID=10 order by ShowOrder"
set Rs=Server.CreateObject("ADODB.RecordSet")
Rs.cursorlocation = 3
Rs.open sqlGroup,Conn,3,3



%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>群組設定系統</title>
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

function DataAdd()
{

 if (confirm("是否確定新增"))
 {
  myForm.DB_Selt.value="DB_Insert";
  myForm.submit();
  }
}

function DataDelete(ID)
{

 if (confirm("是否確定刪除"))
 {
  myForm.DB_Selt.value="DB_Delete";
  myForm.DB_ID.value=ID;
  myForm.submit();
  }
}


-->
</Script>

<body>

<table width="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="style3">群組設定系統</span></td>
  </tr>
<form name="myForm" method="post"> 
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td>

      <img src="space.gif" width="13" height="8">

      <input type="TEXT" Name="GroupName" value="">
      <input type="button" name="Submit3" value="新增群組名稱" onclick="DataAdd();">
   
      </span></td>
      </tr>
    </table></td>
  </tr>
  <input type="Hidden" name="DB_Selt" value="">
  <input type="Hidden" name="DB_ID" value="">
  <tr>
     <td height="26" bgcolor="#FFCC33"><span class="pagetitle style3"><span class="style3">群組設定</span>紀錄列表</span></td>
  </tr>
<%

if not Rs.EOF then
	actionPage=cint(0 & trim(request("page"))) 
	if actionPage < 1 then actionPage=1
	Rs.PageSize=PageSize
	if actionPage > Rs.PageCount then actionPage=Rs.PageCount
	Rs.AbsolutePage=actionPage 
%>  
  <tr>
     <td height="80" bgcolor="#E0E0E0">
     	   <table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
              <tr bgcolor="#EBFBE3">
                <th width="50%" height="15" nowrap><span class="style3 style3 style3">群組名稱</span></th>
				<td width="50%" height="15" nowrap><span class="style5 style3 style3">操作</span></td>
              </tr>
	<%
	for I=1 to Rs.pagesize
	%>              
              <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">

                <td ><%=Rs("Content")%></td>
                   <td> 
   <input type="button" name="Submit3" value="刪除" onclick="DataDelete(<%=Rs("ID")%>);">
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
<%urlParam = "&SN=" & Request("SN") & "&GroupID=" & trim(request("GroupID")) & "&SystemID=" & trim(request("SystemID"))%>			
			<font size="2"><%ShowPageLink actionPage,Rs.PageCount,"FunctionGroupAdd.asp",urlParam%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</tr>  
<% else %>
  <tr>
  	 <td align="center">
	      <center><font  color="Red" size="2">
	<%              
	Response.Write "目前查無任何資料 ..."              
	%>              
	      </font></center><br> 
    	  <p align="center">&nbsp;</p>    	  <p>&nbsp;</p>   	    <p>&nbsp;</p></td>
  </tr>             
<%              
end if              
Rs.close              
set Rs = nothing              
%>   
</FORM>
</table>
</body>

</html>
<!-- #include file="../Common/ClearObject.asp" -->