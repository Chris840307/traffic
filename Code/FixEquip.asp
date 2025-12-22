<!-- #include file="..\Common\db.ini" -->
<!-- #include file="..\Common\AllFunction.inc" -->
<!-- #include file="..\Common\Login_Check.asp" -->

<%
If Session("User_ID")="" Then
	Response.write "系統Session值已過期,請重新登入!"
  Response.End
End If	
Set gadoRS = Server.CreateObject("ADODB.Recordset")

iTag= Request.querystring("tag")

%>
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>固定桿資料維護</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
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
-->
</style></head>
<script language="javascript">
function openFixEquip(OpenFileStr,frmName)
{

window.open(OpenFileStr,frmName,"scrollbars=yes,menubar=no,width=1200,height=900,resizable=yes,left=0,top=0,status=no");	

//window.open(OpenFileStr,frmName,"scrollbars=yes,menubar=no,width=800,height=450,resizable=yes,left=0,top=0,status=no");	
}
function DelFixEquip(param){
	 var rtn;
	 rtn = window.confirm("您確定要刪除此筆資料嗎?");
	 if (rtn!=false){
	 //alert(param)
     openFixEquip(param,'DelProject');	
   }
}

function openExcel(OpenFileStr,frmName)
{

window.open(OpenFileStr,frmName);	

}

function getAddress(obj){
	if (obj.options[obj.selectedIndex].value!=""){
		document.FixEquip.Address.value=obj.options[obj.selectedIndex].text ;
		document.FixEquip.Address1.value=obj.options[obj.selectedIndex].text ;
	}
	else
	{
		document.FixEquip.Address.value="";
		document.FixEquip.Address1.value="";
	
	}
}
</script>

<body>
<form name="FixEquip" method="post" action="FixEquip.asp?tag=search">
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="style2">固定桿資料檔維護</span><span class="style2"><span class="style3">    </span></span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td><span class="style3">      固定桿代碼
            <input name="EquipmentID" type="text" value="<%=Request("EquipmentID")%>" size="7" maxlength="6" class="btn1">      
<%

Set iadoRS = Server.CreateObject("ADODB.Recordset")

	gsql=" select StreetID,Address from Street order by Address "
	iadoRS.open gsql,conn,3,1

%>            
            
            
           路段代碼<select name='StreeID' onchange="getAddress(this)">
           				<option value=''>ALL</option>           
           			<%if not iadoRS.eof then
           				while not iadoRS.eof
           				if trim(Request("StreeID"))=trim(iadoRS(0)) then 
           					iselected =" selected"
           				else
           					iselected =" "
           				end if
           				%>
           				<option value='<%=iadoRS(0)%>' <%=iselected %>><%=iadoRS(1)%></option>
           				<%
           					iadoRS.movenext
           				wend 
           			  end if
           			  iadoRS.close
           				%>
           			</select>            
            地點
            <input name="Address1" type="text" value="<%=Request("Address")%>" size="21" maxlength="20"  class="btn1" disabled>
            <input name="Address" type="hidden" value="<%=Request("Address")%>" size="21" maxlength="20"  class="btn1">            
            違規影像位置
            <input name="ImageIP" type="text" value="<%=Request("ImageIP")%>" size="9" maxlength="12"  class="btn1">
            即時影像位置
            <input name="VideoIP" type="text" value="<%=Request("VideoIP")%>" size="9" maxlength="20"  class="btn1">
            OC位置
            <input name="OCIP" type="text" value="<%=Request("OCIP")%>" size="9" maxlength="12"  class="btn1">
            類型
<%

gsql="select ID ,Content from Code where TypeID ='18' "
iadoRS.open gsql,conn,3,1
%>            
            <select name="TypeID">
             <option value=''></option>
            <%if not iadoRS.eof then
            	while not iadoRS.eof 
            	if trim(Request("TypeID"))=trim(iadoRS("ID")) then
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

      <img src="space.gif" width="13" height="8">
      <input type="submit" name="Submit" value="查詢" <%=iPermission1%>>
      <img src="space.gif" width="9" height="8">      
      <input type="button" name="addFixEquip" value="新增" onclick="openFixEquip('FixEquipDetail.asp?tag=new','AddFixEquip')" <%=iPermission2%>>
  </span></td>
      </tr>
    </table></td>
  </tr>
<%
	gsql="select a.EquipmentID , a.TypeID,b.Content ,a.Address,a.ImageIP,a.VideoIP,a.OCIP ,(Decode(State,'0','空桿','1','使用中','')) as StateDesc from FixEquip a , "
	gsql=gsql & " (select ID , Content from Code Where TypeID =18) b "
	gsql=gsql & " where a.TypeID =b.ID and nvl(a.RecordStateID,0)=0 "
	if iTag="search" then
		if trim(Request("EquipmentID"))<>"" then
			gsql=gsql & " and a.EquipmentID like '"&trim(Request("EquipmentID"))&"%' "	
		end if
		if trim(Request("Address"))<>"" then
			gsql=gsql & " and a.Address  like '%"&trim(Request("Address"))&"%' "	
		end if
		if trim(Request("ImageIP"))<>"" then
			gsql=gsql & " and a.ImageIP  like '"&trim(Request("ImageIP"))&"%' "	
		end if
		if trim(Request("VideoIP"))<>"" then
			gsql=gsql & " and a.VideoIP like '"&trim(Request("VideoIP"))&"%' "	
		end if

		if trim(Request("OCIP"))<>"" then
			gsql=gsql & " and a.OCIP like '"&trim(Request("OCIP"))&"%' "	
		end if

		if trim(Request("TypeID"))<>"" then
			gsql=gsql & " and a.TypeID = '"&trim(Request("TypeID"))&"' "	
		end if
		
	end if
	gsql=gsql & " order by a.State desc "
	gadoRS.cursorlocation = 3
	gadoRS.open gsql,conn,3,1
	Session("ExcelSql") = gsql

%>
  <tr><%'=gsql%>
    <td height="26" bgcolor="#FFCC33"><span class="style2">固定桿資料檔紀錄列表</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#E0E0E0"><table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
      <tr bgcolor="#EBFBE3">
        <th width="7%" height="15" nowrap><span class="style3">固定桿代碼</span></th>
        <th width="13%" height="15" nowrap><span class="style3">地點</span></th>
        <th width="9%" nowrap><span class="style3">違規影像位置</span></th>
        <th width="9%" nowrap><span class="style3">即時影像位置</span></th>
        <th width="8%" nowrap><span class="style3">OC位置</span></th>
        <th width="6%" height="15" nowrap><span class="style3">類型</span></th>
        <th width="6%" height="15" nowrap><span class="style3">使用狀態</span></th>        
        <th width="48%" height="15" nowrap><span class="style3">操作</span></th>
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
      <tr  bgcolor="#FFFFFF" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
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
        <td height="23"><span class="style3">
		  <input type="button" name="btnmodify" value="修改" onclick="javascript: document.location.href='FixEquipDetail.asp?tag=mdy&EquipmentID=<%=gadoRS("EquipmentID")%>' " <%=iPermission3%>>&nbsp;
          <input type="button" name="btnDel" value="刪除" onclick="DelFixEquip('FixEquipDetail.asp?tag=del&EquipmentID=<%=gadoRS("EquipmentID")%>');" <%=iPermission4%> >        
        
        </td>

      </tr>
	<%
				i=i+1
				gadoRS.movenext
			wend
		iPageCount=gadoRS.PageCount
	else
	%>
		<tr><td rowspan =8 align=center>	<font color='red'>查無資料 ....</font></td>
		
		</tr>
	
	<%
					
		end if	

	gadoRS.close
	set gadoRS=nothing		
		
	Conn.close
	Set Conn=nothing
		
		
	%>
    </table></td>    
    
  </tr>
  <tr>
    <td height="26" bgcolor="#FFDD77"><p align="center" class="style1">
      <input type=submit name="Submit422" value="上一頁" onclick="javascript:document.FixEquip.Page.value='<%=cint(Page)-1%>' ;">
      <input type=hidden name='Page' value='' >
      <%=Page%>/<%=iPageCount%>
      <input type=submit name="Submit42" value="下一頁" onclick="javascript:document.FixEquip.Page.value='<%=cint(Page)+1%>' ;">
        <span class="style3"><img src="space.gif" width="13" height="8"></span>        
			<input type="button" name="SaveAs" value="轉換成Excel" onclick="openExcel('FixEquipExcel.asp','FixEquipExcel')">
        <span class="style3"><img src="space.gif" width="13" height="8"></span>        <span class="style2"><span class="style3">
        <input type="button" name="goback" value="回到前一頁" onclick="javascript: document.location.href='index.asp'">
</span></span></p>    </td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    <p>　</p>
    <p>　</p></td></tr>
</table>
</form>
</body>
</html>
