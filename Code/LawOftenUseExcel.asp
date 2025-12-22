<!-- #include file="..\Common\db.ini" -->
<!-- #include file="..\Common\AllFunction.inc" -->
<%
 Response.AddHeader "content-disposition","filename=常用法條檔檔紀錄列表.xls"
 Response.ContentType = "application/vnd.ms-excel"

Set gadoRS = Server.CreateObject("ADODB.Recordset")

	gsql=Session("ExcelSql")
	gadoRS.open gsql,conn,3,1

%>



<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

<title>常用法條檔維護</title>

<body>

<table width="100%" height="100%" border="0">
  <tr>
    <td height="26" bgcolor="#FFCC33"><span class="style2">常用法條檔檔紀錄列表 </span></td>
  </tr>
  <tr>
    <td height="335" bgcolor="#E0E0E0"><table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
      <tr bgcolor="#EBFBE3">
        <th width="8%" height="15" nowrap><span class="style3">法條代碼</span></th>
        <th width="9%" height="15" nowrap><span class="style3">舉發類型</span></th>
        <th width="11%" height="15" nowrap><span class="style3">顯示次序</span></th>
      </tr>
<%


	
		if not gadoRS.eof then 

				

		while not gadoRS.eof 	

%>      
      <tr >
        <td height="23"><div align="center" class="style3"><%=gadoRS("ItemID")%></div></td>
        <td height="23"><div align="center" class="style3"><%=gadoRS("Content")%></div></td>
        <td height="23"><div align="center" class="style3"><%=gadoRS("ShowOrder")%></div></td>
      </tr>
	<%

				gadoRS.movenext
			wend
	

		end if	

	gadoRS.close
	set gadoRS=nothing		
		
	%>      
    </table></td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    <p>　</p>
    <p>　</p></td></tr>
</table>

</body>
</html>