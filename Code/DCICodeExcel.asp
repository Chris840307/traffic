
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
     <td bgcolor="#EBFBE3"><B><center>代碼類型</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>代碼值</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>代碼內容</center></B></td>
   </tr>                     
<%
While Not RsTemp.Eof
   TypeDesc = GetDciTypeById (RsTemp("TypeId"))
%>                             
   <tr>
     <td ><%=TypeDesc%></td>
     <td ><%=RsTemp("Id")%></td>
     <td ><%=RsTemp("Content")%></td>                            
   </tr>   
<%
   RsTemp.MoveNext
Wend
%>     
 </table>    
 </body>
<!-- #include file="../Common/ClearObject.asp" --> 
 <SCRIPT>document.execCommand('SaveAs', true, 'DCI代碼檔紀錄列表.xls');window.close();</SCRIPT>    
 </html>