
<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
sql = Session("ExcelSql")
set RsTemp=Server.CreateObject("ADODB.RecordSet")
RsTemp.open sql,Conn,3,3
%>
<html>   
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>ExportBase</title>
</head>	 
<body>    
 <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#808080" >    
   <tr>
     <td bgcolor="#EBFBE3"><B><center>到案監理站名稱</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>代碼</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>電話</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>住址</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>DCI回傳名稱</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>代碼</center></B></td>     
   </tr>                     
<%
While Not RsTemp.Eof
%>                             
   <tr>
     <td ><%=RsTemp("StationName")%></td>
     <td ><%=RsTemp("StationID")%></td>
     <td ><%=RsTemp("StationTel")%></td>     
     <td ><%=RsTemp("StationAddress")%></td>
     <td ><%=RsTemp("DCIStationName")%></td>
     <td ><%=RsTemp("DCIStationID")%></td>                             
   </tr>   
<%
   RsTemp.MoveNext
Wend
%>     
 </table>    
 </body>
<!-- #include file="../Common/ClearObject.asp" --> 
 <SCRIPT>document.execCommand('SaveAs', true, '監理站資料檔紀錄列表.xls');window.close();</SCRIPT>    
 </html>