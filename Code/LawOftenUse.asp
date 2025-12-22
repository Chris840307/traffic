<!-- #include file="..\Common\db.ini" -->
<!-- #include file="..\Common\AllFunction.inc" -->
<!-- #include file="..\Common\Login_Check.asp" -->
<%
		'權限
           '1:查詢 ,2:新增 ,3:修改 ,4:刪除

           if CheckPermission(226,1)=false then
              iPermission1= "disabled"
           else
              iPermission1= " "
		   end if

           if CheckPermission(226,2)=false then
              iPermission2= "disabled"
           else
              iPermission2= " "
		   end if

           if CheckPermission(226,3)=false then
              iPermission3= "disabled"
           else
              iPermission3= " "
		   end if

           if CheckPermission(226,4)=false then
              iPermission4= "disabled"
           else
              iPermission4= " "
		   end if

%>
<!-- #include file="..\Common\bannernodata.asp" -->
<%
If Session("User_ID")="" Then
	Response.write "系統Session值已過期,請重新登入!"
  Response.End
End If	
Set gadoRS = Server.CreateObject("ADODB.Recordset")

iTag= Request.querystring("tag")


gsql="select a.SN,a.ItemID,a.BillTypeID ,b.Content ,a.ShowOrder from LawOftenUse a ,(select id,content  from DCICODE where TypeID ='2') b  where  a.BillTypeID =b.id (+) "
	if iTag="search" then
		if Request("ItemID")<>"" then 
			gsql=gsql & " and ItemID like '%"&Request("ItemID")&"%' "
		end if
	
		if Request("BillTypeID")<>"" then 
			gsql=gsql & " and a.BillTypeID ='"&Request("BillTypeID")&"' "
		end if
		if Request("ShowOrder")<>"" then 
			gsql=gsql & " and a.ShowOrder='"&Request("ShowOrder")&"' "
		end if
		
	end if

gsql=gsql & "and 1=1 order by a.ShowOrder "
	
'response.write gsql 
'response.end	
	gadoRS.cursorlocation = 3
	gadoRS.open gsql,conn,3,1
	Session("ExcelSql") = gsql
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<title>常用法條檔維護</title>

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
.style4 {color: #666666}
.style5 {font-size: 15px; color: #666666; }
-->
</style></head>
<script language="javascript">
function openLawOftenUse(OpenFileStr,frmName)
{

window.open(OpenFileStr,frmName,"scrollbars=yes,menubar=no,width=1200,height=900,resizable=yes,left=0,top=0,status=no");	

//window.open(OpenFileStr,frmName,"scrollbars=yes,menubar=no,width=800,height=450,resizable=yes,left=0,top=0,status=no");	
}
function DelLawOftenUse(param){
	 var rtn;
	 rtn = window.confirm("您確定要刪除此筆資料嗎?");
	 if (rtn!=false){
	 //alert(param)
     openLawOftenUse(param,'DelLawOftenUse');	
   }
}
function openExcel(OpenFileStr,frmName)
{

window.open(OpenFileStr,frmName);	

}


</script>
<body>
<form name="LawOftenUse" method="post" action="LawOftenUse.asp?tag=search">
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="style2">常用法條檔檔維護</span><span class="style2"><span class="style3">    </span></span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td><span class="style3">      法條代碼
          <input name="ItemID" type="text" value="<%=Request("ItemID")%>" size="10" maxlength="9"  class="btn1">
      舉發類型
      <%
Set iadoRS = Server.CreateObject("ADODB.Recordset")
gsql="select ID,Content from DCICODE where TypeID ='2'"
iadoRS.open gsql,conn,3,1
%>            
            <select name="BillTypeID">
             <option value=''></option>
            <%if not iadoRS.eof then
            	while not iadoRS.eof 
            	if trim(Request("BillTypeID"))=trim(iadoRS("ID")) then
            		iselected =" selected"
            	else
            		iselected =" "            	
            	end if
            	%>
              <option value='<%=iadoRS("ID")%>' <%=iselected%>><%=iadoRS("Content")%></option>
              
             <%
             		iadoRS.movenext
             	wend
              end if
              iadoRS.close
              set iadoRS =nothing
             %>
           </select>
      顯示次序
      <input name="ShowOrder" type="text" value="<%=Request("ShowOrder")%>" size="3" maxlength="2"  class="btn1">      
      <img src="space.gif" width="13" height="8">
      <input type="submit" name="Submit" value="查詢" <%=iPermission1%>>
      <img src="space.gif" width="9" height="8">      
      <input type="button" name="addLawOftenUse" value="新增" onclick="openLawOftenUse('LawOftenUseDetail.asp?tag=new','AddLawOftenUse')" <%=iPermission2%>>
      </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#FFCC33"><span class="style2">常用法條檔檔紀錄列表 </span></td>
  </tr>
  <tr>
    <td height="25" bgcolor="#E0E0E0"><table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
      <tr bgcolor="#EBFBE3">
        <th width="8%" height="15" nowrap><span class="style3">法條代碼</span></th>
        <th width="9%" height="15" nowrap><span class="style3">舉發類型</span></th>
        <th width="11%" height="15" nowrap><span class="style3">顯示次序</span></th>
        <th width="72%" height="15" nowrap><span class="style3">操作</span></th>
      </tr>
<%


	
		if not gadoRS.eof then 

				gadoRS.PageSize=PageSize
			  	IF trim(Request("Page")&" ")="" then 
				   page = 1			  	
				ElseIf Cint(Request("Page")) < 1 Then
				   page = 1
				else 
				   page =Cint(Request("Page"))
				End If

				If page>gadoRS.PageCount then
				   page=gadoRS.PageCount
				end if        	
		
		   		gadoRS.AbsolutePage = page   ' 將資料錄移至 page 頁
          	
				i=0          	

		while not gadoRS.eof and i<gadoRS.PageSize		

%>      
      <tr bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
        <td height="23"><div align="center" class="style3"><%=gadoRS("ItemID")%></div></td>
        <td height="23"><div align="center" class="style3"><%=gadoRS("Content")%></div></td>
        <td height="23"><div align="center" class="style3"><%=gadoRS("ShowOrder")%></div></td>
        <td height="23"><span class="style3">
		  <input type="button" name="btnmodify" value="修改" onclick="javascript: document.location.href='LawOftenUseDetail.asp?tag=mdy&SN=<%=gadoRS("SN")%>' " <%=iPermission3%> >&nbsp;
          <input type="button" name="btnDel" value="刪除" onclick="DelLawOftenUse('LawOftenUseDetail.asp?tag=del&SN=<%=gadoRS("SN")%>');" <%=iPermission4%>>        
</span></td>
      </tr>
	<%
				i=i+1
				gadoRS.movenext
			wend
		iPageCount=gadoRS.PageCount
	else
	%>
		<tr><td rowspan =4 align=center>	<font color='red'>查無資料 ....</font></td>
		
		</tr>
	
	<%
					
		end if	

	gadoRS.close
	set gadoRS=nothing		
		
	%>      
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1">
      <input type=submit name="Submit422" value="上一頁" onclick="javascript:document.LawOftenUse.Page.value='<%=cint(Page)-1%>' ;">
      <input type=hidden name='Page' value='' >
      <%=Page%>/<%=iPageCount%>
      <input type=submit name="Submit42" value="下一頁" onclick="javascript:document.LawOftenUse.Page.value='<%=cint(Page)+1%>' ;">

        <span class="style3"><img src="space.gif" width="13" height="8"></span>       
			<input type="button" name="SaveAs" value="轉換成Excel" onclick="openExcel('LawOftenUseExcel.asp','LawOftenUseExcel')">
        <span class="style3"><img src="space.gif" width="13" height="8"></span>        <span class="style2"><span class="style3">
        <input type="button" name="goback" value="回到前一頁" onclick="javascript: document.location.href='index.asp'">&nbsp;   </span></span></p>    </td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    <p>　</p>
    <p>　</p></td></tr>
</table>
</form>
</body>
<%

	Conn.close
	Set Conn=nothing


%>
</html>