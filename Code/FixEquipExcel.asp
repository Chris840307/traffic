<!-- #include file="..\Common\db.ini" -->
<!-- #include file="..\Common\AllFunction.inc" -->
<%
 Response.AddHeader "content-disposition","filename=固定桿資料檔紀錄列表.xls"
 Response.ContentType = "application/vnd.ms-excel"

Set gadoRS = Server.CreateObject("ADODB.Recordset")

	gsql=Session("ExcelSql")
	gadoRS.open gsql,conn,3,1


%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>固定桿資料維護</title>

<body>

<table width="100%" height="100%" border="0">

  <tr><%'=gsql%>
    <td height="26" bgcolor="#FFCC33"><span class="style2">固定桿資料檔紀錄列表</span></td>
  </tr>
  <tr>
    <td height="335" bgcolor="#E0E0E0"><table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
      <tr bgcolor="#EBFBE3">
        <th width="7%" height="15" nowrap><span class="style3">固定桿代碼</span></th>
        <th width="13%" height="15" nowrap><span class="style3">地點</span></th>
        <th width="9%" nowrap><span class="style3">違規影像位置</span></th>
        <th width="9%" nowrap><span class="style3">即時影像位置</span></th>
        <th width="8%" nowrap><span class="style3">OC位置</span></th>
        <th width="6%" height="15" nowrap><span class="style3">類型</span></th>
        <th width="6%" height="15" nowrap><span class="style3">使用狀態</span></th>        

      </tr>
	<%	
	
		if not gadoRS.eof then 


		while not gadoRS.eof 

 %>            
      <tr  bgcolor="#FFFFFF">
        <td height="23"><div align="right"><%=gadoRS("EquipmentID")%></div></td>
        <td height="23"><div align="center" class="style3">
          <div align="right"><%=gadoRS("Address")%></div>
        </div></td>
        <td ><div align="right"><span class="style3"><%=gadoRS("ImageIP")%></span></div></td>
        <td><div align="right"><span class="style3"><%=gadoRS("VideoIP")%></span></div></td>
        <td><div align="right"><span class="style3"><%=gadoRS("OCIP")%></span></div></td>
        <td height="23"><div align="center" class="style3">
          <div align="right"><%=gadoRS("content")%></div>
        </div></td>
        <td><div align="right"><span class="style3"><%=gadoRS("StateDesc")%></span></div></td>        

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
</form>
</body>
</html>
