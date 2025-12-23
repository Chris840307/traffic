<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>違規地點代碼查詢</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<script LANGUAGE=javascript>
<!--	
function setItemValue(obj,str){
	var tmpCounter;
	if (obj.checked==true){
		tmpCounter = eval(document.sendStreet.checkedNum.value) + 1 ;
		document.sendStreet.checkedNum.value = tmpCounter ;
	  obj.value = str;
	}else {
		tmpCounter = eval(document.sendStreet.checkedNum.value) - 1 ;
		document.sendStreet.checkedNum.value = tmpCounter ;
	  obj.value = "";		  
	}
}	

function Load2Parent(indx,strText,strValue){
  var tmpindex=window.opener.document.forms[0].elements["streetSelect"].length;
  var chk=false;
  for(i=0;i<=tmpindex-1;i++){
	  var tmpvalue=window.opener.document.forms[0].elements["streetSelect"].options[i].value;
	  if (tmpvalue==strValue){chk=true;}
  }
  if(chk==false){
	  window.opener.document.forms[0].elements["streetSelect"].length=tmpindex+1;
	  window.opener.document.forms[0].elements["streetSelect"].options[tmpindex].text = strText;
	  window.opener.document.forms[0].elements["streetSelect"].options[tmpindex].value = strValue;
	  
  }
}	
//-->
</script>
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
.style5 {
	font-size: 14px;
	font-weight: bold;
	color: #FF0000;
}
.style9 {font-family: "標楷體"}
.style10 {font-family: "標楷體"; font-size: 15px; }
.style9 {font-family: "標楷體"; font-weight: bold;}
-->
</style></head>

<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>

<%
Sys_Address = Request("Sys_Address")
Sys_AddID = Request("Sys_AddID")
qryType = Trim(Request("qryType"))
totalCnt = Trim(Request("totalCnt"))
reportId = Request("reportId")

Select Case qryType
   Case "2" :
      sqlStreet = "Select StreetId,Address From Street where 1=1"
      If Sys_Address <> "" Then 
      	 sqlStreet = sqlStreet & " and Address Like '%" & Sys_Address & "%'"
      End If
	  If Sys_AddID <> "" Then 
      	 sqlStreet = sqlStreet & " and streetid = '" & Sys_AddID & "'"
      End If	
      sqlStreet = sqlStreet & " Order By StreetId"
      Set RsStreet=Server.CreateObject("ADODB.RecordSet")

      RsStreet.open sqlStreet,Conn,3,3
   Case "3" :
      ListLength = Int(Request("checkedNum"))
      'Response.write "<script>window.opener.document.forms(0).elements(""streetSelect"").length=" & ListLength & ";</script>"
      j = 0
      For i = 1 To Int(totalCnt)
         fldName = "item_" & i
         strValue = Request(fldName) 
         if strValue <> "" then         	  
         	  strText = Request("text_" & i)
            Response.Write "<script>Load2Parent(" & j & ",'" & strText & "','" & strValue & "');</script>"	
            j = j + 1
         end if
      Next
      Response.Write "<script>window.close();</script>"
End Select   

%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >	
	<form name="qryStreet" method="post" action="QueryStreet.asp"> 
		<table width='100%' border='1' align="left" cellpadding="1">
			
			<tr bgcolor="#FFCC33">
				<td colspan="4">違規地點代碼查詢</td>
			</tr> 
				<INPUT TYPE="HIDDEN" NAME="qryType" VALUE="2">
			<tr>
				<td colspan="4">路段代碼：<input type="text" name="Sys_AddID" value="<%=request("Sys_AddID")%>"></td>
			</tr>
			<tr>
				<td colspan="4">路段名稱：<input type="text"  name="Sys_Address" value="<%=request("Sys_Address")%>"></td>
			</tr>
			<tr>
				<td colspan="4">
				<input type="button" value="查詢" onclick='funQry();'>
				<input type="button" value="關閉視窗" onclick="window.close();">
				</td>
			</tr>
			</form>	
			
			<form name="sendStreet" method="post" action="QueryStreet.asp"> 
				<INPUT TYPE="HIDDEN" NAME="qryType" VALUE="3">	
				<input type="hidden" name="checkedNum" value="0">		
			<tr bgcolor="#FFCC33">
				<td colspan="7">違規地點代碼列表&nbsp;&nbsp;&nbsp;&nbsp; 
					<%If qryType="2" Then Response.Write "<input type='submit' value='加入' >" End If %>
				</td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td width="25%" align="center">代碼</td>
				<td width="75%" align="center">路段</td>
			</tr>

<%
p = 1
IF qryType = "2" THEN
   While Not RsStreet.Eof      
%>
      
			<tr onMouseOver="this.style.backgroundColor='#FF99FF'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
				<td bgcolor="#FFFFCC" align="center">
					<input type="hidden" name="text_<%=p%>" value="<%=RsStreet("Address")%>">					
					<input type="checkbox" name="item_<%=p%>" onClick="setItemValue(this,'<%=RsStreet("StreetId")%>');" >
					<%=RsStreet("StreetId")%>
				</td>
				<td><%=RsStreet("Address")%>&nbsp;</td>
			</tr>	
<%
     p = p + 1
     RsStreet.MoveNext
   Wend
   RsStreet.close
END IF   

%>  
        <INPUT TYPE="HIDDEN" NAME="totalCnt" VALUE="<%=p%>">
		  </form>	
		</table>		
	</form>
<script>
function funQry(){
	if (qryStreet.Sys_AddID.value=="" && qryStreet.Sys_Address.value==""){
		alert("需輸入路段代碼或名稱!");
	}else{
		qryStreet.submit();
	}
}
</script>
</body>		
</html>		
<!-- #include file="../Common/ClearObject.asp" -->