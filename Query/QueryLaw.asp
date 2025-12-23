<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>法條查詢</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<script LANGUAGE=javascript>
<!--	
function setItemValue(obj,str){
	var tmpCounter;
	if (obj.checked==true){
		tmpCounter = eval(document.sendLaw.checkedNum.value) + 1 ;
		document.sendLaw.checkedNum.value = tmpCounter ;
	  //obj.value = str;
	}else {
		tmpCounter = eval(document.sendLaw.checkedNum.value) - 1 ;
		document.sendLaw.checkedNum.value = tmpCounter ;
	 // obj.value = "";		  
	}
}	

function funAllclick(){
	var obj=document.getElementsByTagName("input");
	for(i=0;i<obj.length;i++){
		obj[i].checked=true;
		setItemValue(obj[i]);
	}
}

function Load2Parent(indx,strText,strValue){
  var tmpindex=window.opener.document.forms(0).elements("select3").length;
  var chk=false;
  for(i=0;i<=tmpindex-1;i++){
	  var tmpvalue=window.opener.document.forms(0).elements("select3").options[i].value;
	  if (tmpvalue==strValue){chk=true;}
  }
  if(chk==false){
	  window.opener.document.forms(0).elements("select3").length=tmpindex+1;
	  window.opener.document.forms(0).elements("select3").options[tmpindex].text = strText;
	  window.opener.document.forms(0).elements("select3").options[tmpindex].value = strValue;
	  
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
ITEMID = Trim(Request("ITEMID"))
qryType = Trim(Request("qryType"))
totalCnt = Trim(Request("totalCnt"))
reportId = Request("reportId")

SQL = "select VALUE from Apconfigure where ID=3"
set RsTemp = Server.CreateObject("ADODB.RecordSet")
Set RsTemp = Conn.Execute(SQL)

Select Case qryType
   Case "1" :
      'sqlLaw = "Select ITEMID,CARSIMPLEID,ILLEGALRULE,DECODE(CARSIMPLEID,'1','汽車','2','拖車','3','重機','4','輕機','') CARSIMPLE_DESC, " & _
      '         "LEVEL1,LEVEL2,LEVEL3,LEVEL4, (Case WHEN (ITEMID,CARSIMPLEID) IN (Select ItemId, CarimpleId From UserLawInfo) THEN 'Y' ELSE '' END ) ISSELECT " & _
      '         "From Law Where Version=" & RsTemp("VALUE") & " " & _
      'sqlLaw = sqlLaw & " Order By ITEMID"
      'Set RsLaw=Server.CreateObject("ADODB.RecordSet")
      'RsLaw.open sqlLaw,Conn,3,3                  
   Case "2" :
      sqlLaw = "Select Distinct ITEMID,ILLEGALRULE From Law Where Version=" & RsTemp("VALUE") & " "
      If ITEMID<>"" Then 
      	 sqlLaw = sqlLaw & "And ITEMID LIKE '" & ITEMID & "%' "
      End If	

      sqlLaw = sqlLaw & " Order By ITEMID"
      Set RsLaw=Server.CreateObject("ADODB.RecordSet")
      RsLaw.open sqlLaw,Conn,3,3
   Case "3" :
      ListLength = Int(Request("checkedNum"))
      'Response.write "<script>window.opener.document.forms(0).elements(""select3"").length=" & ListLength & ";</script>"
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
		<table width='100%' border='1' align="left" cellpadding="1">
			
			<tr bgcolor="#FFCC33">
				<td colspan="7">法條查詢</td>
			</tr>
			<form name="qryLaw" method="post" action="QueryLaw.asp">  
				<INPUT TYPE="HIDDEN" NAME="qryType" VALUE="2">
			<tr>
				<td colspan="7">
				代碼：<input type="text" name="ITEMID" value="">&nbsp;&nbsp;&nbsp;
				<input type="submit" value="查詢" >
				<input type="button" value="關閉視窗" onclick="window.close();">
				</td>
			<tr>
			</form>	
			
			<form name="sendLaw" method="post" action="QueryLaw.asp"> 
				<INPUT TYPE="HIDDEN" NAME="qryType" VALUE="3">
				<input type="hidden" name="checkedNum" value="0">
			<tr bgcolor="#FFCC33">
				<td colspan="2">法條列表&nbsp;&nbsp;&nbsp;&nbsp; 
					<%If qryType="2" Then
						Response.Write "<input type='submit' value='加入' >&nbsp;&nbsp;"
						Response.Write "<input type='button' value='全選' onclick='funAllclick();'>"
					End If %>
				</td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td width="20%" align="center">法條代碼</td>
				<td width="80%" align="center">法條內容</td>
			</tr>

<%
p = 1
IF qryType = "2" THEN
   While Not RsLaw.Eof      
%>
      
			<tr onMouseOver="this.style.backgroundColor='#FF99FF'" onMouseOut="this.style.backgroundColor='#FFFFFF'">
				<td bgcolor="#FFFFCC" >
					<input type="hidden" name="text_<%=p%>" value="<%=RsLaw("ILLEGALRULE")%>">					
					<input type="checkbox" name="item_<%=p%>" value="<%=RsLaw("ITEMID")%>" onClick="setItemValue(this);" >
					<%=RsLaw("ITEMID")%>
				</td>
				<td><%=RsLaw("ILLEGALRULE")%></td>
			</tr>	
<%
     p = p + 1
     RsLaw.MoveNext
   Wend
END IF   
%>  
        <INPUT TYPE="HIDDEN" NAME="totalCnt" VALUE="<%=p%>">
		  </form>	
		</table>		
	</form>

</body>		
</html>		
<!-- #include file="../Common/ClearObject.asp" -->