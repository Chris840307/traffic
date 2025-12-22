
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
     <td bgcolor="#EBFBE3"><B><center>資料交換類型</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>狀態代碼</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>內容</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>結果</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>須再處理</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>方式</center></B></td>
   </tr>                     
<%
While Not RsTemp.Eof
%>                             
   <tr>
     <td ><%=RsTemp("DCIActionName")%></td>
     <td ><%=RsTemp("DCIReturn")%></td>
     <td ><%=RsTemp("StatusContent")%></td> 
     <td ><%=RsTemp("DCIReturnStatusDesc")%></td>
     <td ><%=RsTemp("NeedReDoDesc")%></td>
     <td ><%=RsTemp("HowTo")%></td>                           
   </tr>   
<%
   RsTemp.MoveNext
Wend
%>     
 </table>    
 </body>
<!-- #include file="../Common/ClearObject.asp" --> 
 <SCRIPT>document.execCommand('SaveAs', true, 'DCI回覆結果檔紀錄列表.xls');window.close();</SCRIPT>    
 </html>