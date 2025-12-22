
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
     <td bgcolor="#EBFBE3"><B><center>郵遞區號</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>郵遞區名稱</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>分區綑綁區號</center></B></td>
   </tr>                     
<%
While Not RsTemp.Eof
%>                             
   <tr>
     <td ><%=RsTemp("ZipID")%></td>
     <td ><%=RsTemp("ZipName")%></td>
     <td ><%=RsTemp("ZipNo")%></td>                            
   </tr>   
<%
   RsTemp.MoveNext
Wend
%>     
 </table>    
 </body>
<!-- #include file="../Common/ClearObject.asp" --> 
 <SCRIPT>document.execCommand('SaveAs', true, '郵遞區號檔紀錄列表.xls');window.close();</SCRIPT>    
 </html>