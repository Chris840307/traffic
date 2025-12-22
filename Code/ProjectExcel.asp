<!-- #include file="..\Common\db.ini" -->
<!-- #include file="..\Common\AllFunction.inc" -->
<%

 Response.AddHeader "content-disposition","filename=專案資料檔紀錄列表.xls"
 Response.ContentType = "application/vnd.ms-excel"

Set gadoRS = Server.CreateObject("ADODB.Recordset")

	gsql=Session("ExcelSql")
	gadoRS.open gsql,conn,3,1


	
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

<title>專案資料檔維護</title>
<body>

<table width="100%" height="100%" border="0">
    <td height="26" bgcolor="#FFCC33"><span class="style2">專案資料檔紀錄列表 </span></td>
  </tr>
  <tr>
    <td height="335" bgcolor="#E0E0E0"><table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
      <tr bgcolor="#EBFBE3">
        <th width="8%" height="15" nowrap><span class="style3">專案代碼</span></th>
        <th width="14%" height="15" nowrap><span class="style3">專案名稱</span></th>
        <th width="14%" height="15" nowrap><span class="style3">專案施行期間</span></th>
        <th width="6%" nowrap><span class="style3">狀態</span></th>

      </tr>
	<%	
	
		if not gadoRS.eof then 



		while not gadoRS.eof

 %>      
      <tr  bgcolor="#FFFFFF">
        <td height="23"><div align="center" class="style3"><%=gadoRS("ProjectID")%></div></td>
        <td height="23"><div align="center" class="style3"><%=gadoRS("Name")%></div></td>
        <td height="23"><div align="center" class="style3"><%=gInitDT(gadoRS("StartDate"))%> ~ <%=gInitDT(gadoRS("EndDate"))%> </div></td>
        <td><div align="center" class="style3"><%=gadoRS("RecordStateID")%></div></td>

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
</table>
</form>
</body>
<script>window.close</script>
</html>


