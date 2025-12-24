<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
AuthorityCheck(229)
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
              <tr bgcolor="#EBFBE3">
                <th width="15%"  nowrap><span class="style3 style3 style3">群組名稱</span></th>
                <th width="30%" nowrap><span class="style3 style3 style3">系統名稱</span></th>
                <td width="3" nowrap><span class="style5">新增</span></td>
                <td width="3" nowrap><span class="style3"><strong>修改</strong></span></td>
                <td width="3" nowrap><span class="style3"><strong>刪除</strong></span></td>
                <td width="3" nowrap><span class="style3"><strong>查詢</strong></span></td>
              </tr>                    
<%
While Not RsTemp.Eof
%>                             
   <tr>
     <td ><%=RsTemp("GroupName")%></td>
     <td ><%=RsTemp("SYSTEMName")%></td>
     <td ><%=RsTemp("InsertFlag")%></td>
     <td ><%=RsTemp("UpdateFlag")%></td>
     <td ><%=RsTemp("DeleteFlag")%></td>
     <td ><%=RsTemp("SelectFlag")%></td>                     
   </tr>   
<%
   RsTemp.MoveNext
Wend
%>     
 </table>    
 </body>
<!-- #include file="../Common/ClearObject.asp" --> 
 <SCRIPT>document.execCommand('SaveAs', true, '權限管理列表.xls');window.close();</SCRIPT>    
 </html>