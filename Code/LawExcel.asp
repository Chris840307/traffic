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

<title>法條檔維護</title>

<body>

<table width="100%" height="100%" border="0">
  <tr>
    <td height="26" bgcolor="#FFCC33"><span class="style2">法條檔檔紀錄列表</span></td>
  </tr>
  <tr>
    <td height="335" bgcolor="#E0E0E0"><table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
      <tr bgcolor="#EBFBE3">
        <th width="6%" height="15" nowrap><span class="style3">法條代碼</span></th>
        <th width="6%" height="15" nowrap><span class="style3">簡式<br>
          車種</span></th>
        <th width="21%" nowrap><span class="style3">違規事實</span></th>
        <th width="9%" nowrap><span class="style3">期限內 </span></th>
        <th width="4%" nowrap><span class="style3">歸責<br>
          對象</span></th>
        <th width="4%" nowrap><span class="style3">記點<br>
          記次</span></th>
        <th width="5%" nowrap><span class="style3">吊扣<br>
          吊銷</span></th>
        <th width="3%" nowrap><span class="style3">禁考<br>
          註記</span></th>
        <th width="3%" nowrap><span class="style3">保留</span></th>
        <th width="3%" nowrap><span class="style3">特殊<br>
          處罰</span></th>

      </tr>
<%


	
		if not gadoRS.eof then 

     	

		while not gadoRS.eof 

%>            
      <tr bgcolor="#FFFFFF" >
        <td height="23"><div align="center" class="style3"><%=gadoRS("ItemID")%></div></td>
        <td height="23"><div align="center" class="style3"><%=gadoRS("CarsImpleID")&gadoRS("CarType")%></div></td>
        <td><%=replace(trim(gadoRS("IllegalRule")&" "),trim(Request("IllegalRule")),"<b>"&trim(Request("IllegalRule"))&"</b>")%>　</td>
        <td>&nbsp;<%=gadoRS("Level1")%>,<%=gadoRS("Level2")%>,<%=gadoRS("Level3")%>,<%=gadoRS("Level4")%></td>
        <% if gadoRS("Target")="V" then 
        			starget="歸車"
        	 ElseIf gadoRS("Target")="0" then        	 
        	 		starget="駕駛人"
					 else
					 		starget=""        	 		
        	 end if
        %>
        <td><%=starget%></td>
        <td><%=gadoRS("RecordPoint")%>　</td>
        <td><%=gadoRS("RevokePoint")%>　</td>
        <td><%=gadoRS("Notest")%>　</td>
        <td><%=gadoRS("Retain")%>　</td>
        <td><%=gadoRS("SpecPunish")%>　</td>
        <td height="23"><span class="style3">
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
