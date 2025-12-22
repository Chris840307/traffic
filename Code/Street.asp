<!-- #include file="..\Common\DbUtil.asp"-->
<!-- #include file="..\Common\Util.asp"-->
<!-- #include file="..\Common\AllFunction.inc"-->
<!-- #include file="..\Common\Login_Check.asp"-->
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
	if ((form_A.StreetID.value=="") && (form_A.StreetSimpleName.value=="") && (form_A.Address.value=="")){
	   alert("您尚未輸入任何查詢條件!!");
	   return false;	
	}
}
function sendQry(){
	var rtn;
	//rtn = qryCheck();
	//if (rtn!=false){
     var form_A= document.forms[0];
     form_A.action = "Street.asp";
     form_A.submit();		
	//}
}
function delStreet(param){
	 var rtn;
	 rtn = window.confirm("您確定要刪除此筆資料嗎?");
	 if (rtn!=false){
     openAddGetBill(param,'Street');	
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
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">路段代碼檔維護  (為確保歷史案件資料查詢及統計正確性,路段代碼不可修改)</span></td>
  </tr>
<FORM NAME="Street" ACTION="" METHOD="POST">   
	<input type="hidden" name="isQuery" value="y">
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td>
        	路段代碼
          <input name="StreetID" type="text" value="<%=Request("StreetID")%>" size="10" maxlength="9" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
           路段
          <input name="Address" type="text" value="<%=Request("Address")%>" size="41" maxlength="40" onKeyDown="lockSpecialCharr(this);" onKeyUp="lockSpecialCharr(this);" class="btn1">
          <img src="space.gif" width="13" height="8">
			
		  <%If sys_City="高雄市" Then %>
           固定桿
          <input name="FixPole" type="checkbox" value="0" <%If request("FixPole")<>"" Then response.write "checked"%>>
          <img src="space.gif" width="13" height="8">
		  <%End if%>

          <input type="button" value="查詢" onclick="sendQry();" <%=ReturnPermission(CheckPermission(226,1))%>>
          <img src="space.gif" width="9" height="8">      
          <input type="button" value="新增" onclick="openAddGetBill('StreetAdd.asp?tag=new','StreetAdd')" <%=ReturnPermission(CheckPermission(226,2))%>>
       
		 <font size="2">路段代碼 與 路段 可使用關鍵字查詢</font>
		 <br>
		 <font size="2">您可以輸入單位專用 路段代碼第一碼 進行查詢</font>
		
<%

	if sys_City="雲林縣" then
%>
		<br><font color="red"><strong>路段代碼規則：英文字母(本局=A、B、C，斗六=D，斗南=E，虎尾=F，西螺=G，北港=H，
		台西=I)+ 3碼流水號，例如A001、A002</strong></font>
<%
	end if
%>
        </td>
      </tr>
    </table></td>
  </tr>
</FORM>    
  <tr>
    <td height="26" bgcolor="#FFCC33"><span class="pagetitle">路段代碼檔紀錄列表</span></td>
  </tr>
<%
if Request("isQuery") = "y" then
   SQL = "Select a.*,b.UnitName From Street a,UnitInfo b Where a.StreetID is not null and a.UnitID=b.UnitID(+)"
   if Request("StreetID")<>"" then
      SQL = SQL & "And a.StreetID like '" & Request("StreetID") & "%' "
   end if
   if Request("StreetSimpleName")<>"" then
      SQL = SQL & "And a.StreetSimpleName='" & Request("StreetSimpleName") & "' "
   end if	
   if Request("Address")<>"" then
      SQL = SQL & " And (a.Address Like '%" & Request("Address") & "%') "
   end if	

If sys_City="高雄市" Then
   if Request("FixPole")<>"" then
      SQL = SQL & " And (a.FixPole = '1') "
   end if	
End if

   SQL = SQL & "Order By StreetID"
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
		<th width="15%" height="15" nowrap>單位名稱</th>
        <th width="10%" height="15" nowrap>路段代碼</th>
        <th width="26%" height="15" nowrap>路段</th>
        <th width="51%" height="15" nowrap>操作</th>
      </tr>
	<%             
	for I=1 to Rs.pagesize   
	%>        
      <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
		<td height="23"><div align="left"><%=Rs("UnitName")%></div></td>
        <td height="23"><div align="left"><%=Rs("StreetID")%></div></td>
        <td height="23" width="60%" nowarp><div align="left"><%=Rs("Address")%></div></td>
        <td height="23">
			<%If sys_City="高雄市" Then%>
           <input type="button" value="修改" onclick="openAddGetBill('StreetUpdate.asp?tag=upd&StreetID=<%=Rs("StreetID")%>&StreetSimpleName=<%=Rs("StreetSimpleName")%>&Address=<%=Rs("Address")%>&FixPole=<%=Rs("FixPole")%>','UpdateStreet')" <%=ReturnPermission(CheckPermission(226,3))%>>&nbsp;&nbsp;&nbsp;&nbsp;
           <!-- <input type="button" value="刪除" onclick="delStreet('Street_mdy.asp?tag=del&StreetID= --><%
		   'response.write rs("StreetID")
		   %><!-- ');"  --><%
		   'response.write ReturnPermission(CheckPermission(226,4))
		   %><!-- > -->
		   <%else%>
			<input type="button" value="修改" onclick="openAddGetBill('StreetUpdate.asp?tag=upd&StreetID=<%=Rs("StreetID")%>&StreetSimpleName=<%=Rs("StreetSimpleName")%>&Address=<%=Rs("Address")%>','UpdateStreet')" <%=ReturnPermission(CheckPermission(226,3))%>>&nbsp;&nbsp;&nbsp;&nbsp;
		   <%End if%>
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
   urlParam = "&isQuery=" & Request("isQuery") & "&StreetID=" & Request("StreetID") & "&StreetSimpleName=" & Request("StreetSimpleName") & "&Address=" & Request("Address")
%>			
			<font size="2"><%ShowPageLink actionPage,Rs.PageCount,"Street.asp",urlParam%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" name="SaveAs" value="轉換成Excel" onclick="exportExcel('StreetExcel.asp','StreetExcel')">
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
<tr>




</tr>	
<%              
   end if              
   Rs.close              
   set Rs = nothing  
end if            
%>   
<tr>
<td>
<%

if sys_City="高雄市" then 	
	response.write "<b>各路段代碼最大值 , 新增 路段代碼 請再加1 </b><br>"
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like '0%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like '1%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID & " "
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like '2%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID 
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like '3%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID 
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like '4%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID 
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like '5%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID 
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like '6%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID 
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like '7%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID 
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like '8%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID 
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like '9%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID 
	
%>
	<Br>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'A%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write maxStreetID
	
%>

<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'B%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>

<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'E%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'F%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'G%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'H%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'I%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'J%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'K%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'L%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
 <br>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'M%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'N%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'O%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'P%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'Q%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'R%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>

<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'T%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>

<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'V%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'W%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'X%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'Y%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
	
%>
<%
	strMaxStreetID="select max(Streetid) as maxstreetid from Street where Streetid like 'Z%'"
	set rsCity=conn.execute(strMaxStreetID)
	maxStreetID=trim(rsCity("maxstreetid"))
	rsCity.close
	set rsCity=nothing
	response.write " 　" & maxStreetID
   end if
%>
</td>
</tr>

<br>  
</table>
</body>
</html>
<!-- #include file="../Common/ClearObject.asp" -->