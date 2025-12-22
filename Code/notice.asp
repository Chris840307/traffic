\<!-- #include file="..\Common\db.ini" -->
<!-- #include file="..\Common\AllFunction.inc" -->
<!-- #include file="..\Common\Login_Check.asp"-->
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


<%
If Session("User_ID")="" Then
	Response.write "系統Session值已過期,請重新登入!"
  Response.End
End If	
Set gadoRS = Server.CreateObject("ADODB.Recordset")
iTag= Request.querystring("tag")

	gsql="select ID,noticeData ,StartDate,EndDate ,decode (RecordStateID,'0','開始','停止') RecordStateID,RecordDate,RecordMemID  "
	gsql=gsql & " from notice where 1=1 "	
	if iTag="search" then
		if Request("noticeData ")<>"" then
			gsql=gsql &" and noticeData like '"&Request("noticeData ")&"%' "		
		end if

		if Request("StartDate")<>"" then
			gsql=gsql & " and to_char(StartDate,'yyyy/mm/dd') >='"&gOutDT(Request("StartDate"))&"' "
		end if
		if Request("EndDate")<>"" then
			gsql=gsql & " and to_char(EndDate,'yyyy/mm/dd') <='"&gOutDT(Request("EndDate"))&"' "
		end if
		if Request("RecordStateID")<>"" then
			gsql=gsql &" and RecordStateID = '"&Request("RecordStateID")&"' "		
		end if		
	end if

	gsql=gsql & " Order by RecordDate desc"
	'response.write gsql
	'response.end
	gadoRS.cursorlocation = 3
	gadoRS.open gsql,conn,3,1

	Session("ExcelSql") = gsql
	
%>
<!-- #include file="..\Common\bannernodata.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<title>公告訊息維護</title>
<style type="text/css">
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
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<script language="javascript">
function openNotice(OpenFileStr,frmName)
{

window.open(OpenFileStr,frmName,"scrollbars=yes,menubar=no,width=1200,height=900,resizable=yes,left=0,top=0,status=no");	

//window.open(OpenFileStr,frmName,"scrollbars=yes,menubar=no,width=800,height=450,resizable=yes,left=0,top=0,status=no");	
}

function DelNotice(param){
	 var rtn;
	 rtn = window.confirm("您確定要刪除此筆資料嗎?");
	 if (rtn!=false){
	 //alert(param)
     openNotice(param,'DelProject');	
   }
}

function openExcel(OpenFileStr,frmName)
{

//window.open(OpenFileStr,frmName);	
	newWin(OpenFileStr,frmName,900,550,50,10,"yes","yes","yes","no");

}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}

</script>
<body>
<form name="Notice" method="post" action="Notice.asp?tag=search">
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">公告訊息維護</span><span class="style2"><span class="style3">    </span></span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td><span class="style3"> 
      公告訊息
      <input name="NoticeData" type="text" value="<%=Request("NoticeData")%>" size="10" maxlength="9"  class="btn1">      
      公告訊息期間
<input type='text' size='10' id='StartDate'  class="btn1" name='StartDate' value="<%=Request("StartDate")%>" readonly onclick="OpenWindow('StartDate')"><input type=button value="..." name='btnDateS'  onclick="OpenWindow('StartDate')">      ~
<input type='text' size='10' id='EndDate'  class="btn1" name='EndDate' value="<%=Request("EndDate")%>" readonly onclick="OpenWindow('EndDate')"><input type=button value="..." name='btnDateE'  onclick="OpenWindow('EndDate')"> 
狀態
<select name="RecordStateID">
	<%if Request("RecordStateID")="-1" then%>
  <option value='-1'>停止</option>
  <option value ='0'>開始</option>  	
  	<%else%>
  <option value ='0'>開始</option>
  <option value='-1'>停止</option>


  		<%end if%>
      </select>
      <img src="space.gif" width="13" height="8">
      <input type="submit" name="Submit" value="查詢" <%=iPermission1%>>
      <img src="space.gif" width="9" height="8">      
      <input type=button name="Submit2" value="新增"  onclick="openNotice('NoticeDetail.asp?tag=new','AddNotice') " <%=iPermission2%>>
     </span></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#FFCC33"><span class="style2">公告訊息紀錄列表 </span></td>
  </tr>
  <tr>
    <td height="25" bgcolor="#E0E0E0"><table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
      <tr bgcolor="#EBFBE3">
        <th width="14%" height="15" nowrap><span class="style3">公告訊息</span></th>
        <th width="14%" height="15" nowrap><span class="style3">公告訊息期間</span></th>
        <th width="6%" nowrap><span class="style3">狀態</span></th>
        <th width="58%" height="15" nowrap><span class="style3">操作</span></th>
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
        <td height="23"><div align="center" class="style3"><%=gadoRS("NoticeData")%></div></td>
        <td height="23"><div align="center" class="style3"><%=gInitDT(gadoRS("StartDate"))%> ~ <%=gInitDT(gadoRS("EndDate"))%> </div></td>
        <td><div align="center" class="style3"><%=gadoRS("RecordStateID")%></div></td>
        <td height="25"><span class="style3">
<input type="button" name="btnmodify" value="修改" onclick="javascript: document.location.href='NoticeDetail.asp?tag=mdy&ID=<%=gadoRS("ID")%>' " <%=iPermission3%>>&nbsp;
          <input type="button" name="btnDel" value="刪除" onclick="DelNotice('NoticeDetail.asp?tag=del&ID=<%=gadoRS("ID")%>');" <%=iPermission4%> >
	</span></td>
      </tr>
	<%
				i=i+1
				gadoRS.movenext
			wend
		iPageCount=gadoRS.PageCount

	else
%>
  <tr>
  	 <td align="center" >        
	      <center><font  color="Red" size="2">              
	<%              
	Response.Write "目前查無任何資料 ..."              
	%>              
	      </font></center><br> 
	   </td>
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
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1">
      <input type=submit name="Submit422" value="上一頁" onclick="javascript:document.Notice.Page.value='<%=cint(Page)-1%>' ;">
      <input type=hidden name='Page' value='' >
      <%=Page%>/<%=iPageCount%>
      <input type=submit name="Submit42" value="下一頁" onclick="javascript:document.Notice.Page.value='<%=cint(Page)+1%>' ;">
        <span class="style3"><img src="space.gif" width="13" height="8"></span>       
			
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
