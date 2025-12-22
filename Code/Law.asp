<!-- #include file="..\Common\db.ini" -->
<!-- #include file="..\Common\AllFunction.inc" -->
<!-- #include file="..\Common\Login_Check.asp" -->
<!-- #include file="..\Common\bannernodata.asp" -->
<%
If Session("User_ID")="" Then
	Response.write "系統Session值已過期,請重新登入!"
  Response.End
End If	
Set gadoRS = Server.CreateObject("ADODB.Recordset")

iTag= Request.querystring("tag")


gsql=" select ItemID,CarSimpleID,decode(CarsImpleID,'0','不分車種','1','自用汽車','2','營業車','3','機車','4','汽車','5','小型車','6','大型車','7','大客','8','營大客','') as CarType,IllegalRule,Level1,Level2,Level3,Level4,Target "
gsql=gsql & " ,RecordPoint,RevokePoint,NoTest,Retain,SpecPunish "
gsql=gsql & " from Law  where 1=1 and nvl(RecordStateID,0)=0  "
	if iTag="search" then
		if Request("ItemID")<>"" then 
			gsql=gsql & " and ItemID like '%"&Request("ItemID")&"%' "
		end if
	

		if Request("IllegalRule")<>"" then 
			gsql=gsql & " and IllegalRule like '%"&Request("IllegalRule")&"%' "
		end if

		if Request("CarSimpleID")<>"" then 						'&"' or CarSimpleID='0' ) "
			gsql=gsql & " and ( CarSimpleID='"&Request("CarSimpleID") & "') "
		end if

		if Request("Level1")<>"" then 
			gsql=gsql & " and Level1='"&Request("Level1")&"' "
		end if

		if Request("Level2")<>"" then 
			gsql=gsql & " and Level2='"&Request("Level2")&"' "
		end if
		if Request("Level3")<>"" then 
			gsql=gsql & " and Level3='"&Request("Level3")&"' "
		end if

		if Request("Level4")<>"" then 
			gsql=gsql & " and Level4='"&Request("Level4")&"' "
		end if

		if Request("Target")<>"" then 
			gsql=gsql & " and Target='"&Request("Target")&"' "
		end if

		if Request("RecordPoint")<>"" then 
			gsql=gsql & " and RecordPoint='"&Request("RecordPoint")&"' "
		end if
		if Request("RevokePoint")<>"" then 
			gsql=gsql & " and RevokePoint='"&Request("RevokePoint")&"' "
		end if
		if Request("NoTest")<>"" then 
			gsql=gsql & " and NoTest='"&Request("NoTest")&"' "
		end if

		if Request("Retain")<>"" then 
			gsql=gsql & " and Retain='"&Request("Retain")&"' "
		end if
		if Request("SpecPunish")<>"" then 
			gsql=gsql & " and SpecPunish='"&Request("SpecPunish")&"' "
		end if
		if Request("version")<>"" then 
			gsql=gsql & " and version='"&Request("version")&"' "
		end if	
	end if
gsql=gsql & "  and 2=2 order by ItemID "
	
'response.write gsql 
'response.end	
	gadoRS.cursorlocation = 3
	gadoRS.open gsql,conn,3,1
	Session("ExcelSql") = gsql
	'response.write gsql
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


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<title>法條檔維護</title>

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
function openLaw(OpenFileStr,frmName)
{

window.open(OpenFileStr,frmName,"scrollbars=yes,menubar=no,width=1200,height=900,resizable=yes,left=0,top=0,status=no");	

//window.open(OpenFileStr,frmName,"scrollbars=yes,menubar=no,width=800,height=450,resizable=yes,left=0,top=0,status=no");	
}
function DelLaw(param){
	 var rtn;
	 rtn = window.confirm("您確定要刪除此筆資料嗎?");
	 if (rtn!=false){
	 //alert(param)
     openLaw(param,'DelLaw');	
   }
}
function openExcel(OpenFileStr,frmName)
{

window.open(OpenFileStr,frmName);	

}

</script>
<body>
<form name="Law" method="post" action="Law.asp?tag=search">
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="style2">法條檔檔維護</span><span class="style2"><span class="style3">    </span></span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td><span class="style3">      法條代碼
          <input name="ItemID" type="text" value="<%=Request("ItemID")%>" size="10" maxlength="9" class="btn1">
          簡式車種
          <select name="CarSimpleID">
          	<%
			select case Request("CarSimpleID")			
				case "1"
					iselect1 =" selected"
				case"2"
					iselect2 =" selected"				
				case "3"
					iselect3 =" selected"
				case "4"
					iselect4 =" selected"			
				case "5"
					iselect5 =" selected"
				case "6"
					iselect6 =" selected"				
				case "7"
					iselect7 =" selected"
				case "8"
					iselect8 =" selected"							
			end select			
			%>
		   <option value=""></option>
            <option value="1" <%=iselect1 %>>自用汽車</option>
            <option value="2" <%=iselect2 %>>營業車</option>
            <option value="3" <%=iselect3 %>>機車</option>
            <option value="4" <%=iselect4 %>>汽車</option>
            <option value="5" <%=iselect5 %>>小型車</option>
            <option value="6" <%=iselect6 %>>大型車</option>
            <option value="7" <%=iselect7 %>>大客車</option>
            <option value="8" <%=iselect8 %>>營大客</option>

          </select>          
          違規事實
          <input name="IllegalRule" type="text" value="<%=Request("IllegalRule")%>" size="60" maxlength="100" class="btn1">
          <br>          
          期限內
          <input name="Level1" type="text" value="<%=Request("Level1")%>" size="6" maxlength="5"  class="btn1"> 
          15天內罰款
          <input name="Level2" type="text" value="<%=Request("Level2")%>" size="6" maxlength="5"  class="btn1">
15~30天
<input name="Level3" type="text" value="<%=Request("Level3")%>" size="6" maxlength="5"  class="btn1"> 
超過30天以上
<input name="Level4" type="text" value="<%=Request("Level4")%>" size="6" maxlength="5"  class="btn1"> 
          歸責對象
          <select name="Target">
            <option value=""></option>          
			<%
			select case Request("Target")			
				case "V"
					iselect11 =" selected"
				case"0"
					iselect21 =" selected"				
			end select			
			%>               
            <option value="V" <%=iselect11%>>車主</option>
            <option value="0" <%=iselect21%>>駕駛人</option>
            </select>
          版本
<%
Set iadoRS = Server.CreateObject("ADODB.Recordset")
gsql="select distinct(version) from Law order by version desc "
iadoRS.open gsql,conn,3,1
%>
          <select name="version" id="version">
             <option value=''></option>
            <%if not iadoRS.eof then
		
            	while not iadoRS.eof 
            	%>
              <option value='<%=iadoRS("Version")%>' 
              <%if Request("version")=iadoRS("Version") then response.write " Selected"%>
              >
              <%=iadoRS("Version")%></option>
              
             <%
             		iadoRS.movenext
             	wend
              end if
              iadoRS.close
              set iadoRS =nothing
             %>
              </select>		  

          <br>
          記點記次
          <input name="RecordPoint" type="text" size="4" maxlength="3" value="<%=Request("RecordPoint")%>"  class="btn1">
          吊扣吊銷
          <input name="RevokePoint" type="text" size="4" maxlength="3"  value="<%=Request("RevokePoint")%>"  class="btn1">
          禁考註記
          <input name="NoTest" type="text" size="4" maxlength="3"  value="<%=Request("NoTest")%>"  class="btn1">
保留
<input name="Retain" type="text" size="4" maxlength="3"  value="<%=Request("Retain")%>"  class="btn1">
特殊處罰
<input name="SpecPunish" type="text" size="4" maxlength="3" value="<%=Request("SpecPunish")%>"  class="btn1">
          <img src="space.gif" width="13" height="8">
      <input type="submit" name="Submit" value="查詢" >
      <img src="space.gif" width="9" height="8">      
    <!--  <input type="button" name="addLaw" value="新增" onclick="openLaw('LawDetail.asp?tag=new','AddLaw')"<%=iPermission2%>> -->
        </span></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#FFCC33"><span class="style2">法條檔檔紀錄列表</span> <img src="space.gif" width="22" height="8"><b>歸責對象</b>( V歸車 0駕駛人 ) </td>
  </tr>
  <tr>
    <td height="25" bgcolor="#E0E0E0"><table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
      <tr bgcolor="#EBFBE3">
        <th width="6%" height="15" nowrap><span class="style3">法條代碼</span></th>
        <th width="5%" height="15" nowrap><span class="style3">簡式<br>
          車種</span></th>
        <th width="31%" nowrap><span class="style3">違規事實</span></th>
        <th width="18%" nowrap>罰款</th>
        <th width="5%" nowrap><span class="style3">歸責<br>
          對象</span></th>
        <th width="4%" nowrap><span class="style3">記點<br>
          記次</span></th>
        <th width="4%" nowrap><span class="style3">吊扣<br>
          吊銷</span></th>
        <th width="4%" nowrap><span class="style3">禁考<br>
          註記</span></th>
        <th width="4%" nowrap><span class="style3">保留</span></th>
        <th width="4%" nowrap><span class="style3">特殊<br>
          處罰</span></th>
           <!--   <th width="15%" height="15" nowrap><span class="style3">操作</span></th>-->
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
        <td height="23"  nowrap><div align="left" class="style3" ><%=gadoRS("CarsImpleID")&gadoRS("CarType")%></div></td>
        <td><%=replace(trim(gadoRS("IllegalRule")&" "),trim(Request("IllegalRule")),"<b>"&trim(Request("IllegalRule"))&"</b>")%>　</td>
        <td><%=gadoRS("Level1")%>,<%=gadoRS("Level2")%>,<%=gadoRS("Level3")%>,<%=gadoRS("Level4")%></td>
        <td><%=gadoRS("Target")%></td>
        <td><%=gadoRS("RecordPoint")%>　</td>
        <td><%=gadoRS("RevokePoint")%>　</td>
        <td><%=gadoRS("Notest")%>　</td>
        <td><%=gadoRS("Retain")%>　</td>
        <td><%=gadoRS("SpecPunish")%>　</td>
    <!--    <td height="23"><span class="style3"> -->
<!-- <input type=button name="mdyLaw" value="修改"  onclick="javascript: document.location.href='LawDetail.asp?tag=mdy&ItemID=<%=gadoRS("ItemID")%>&CarSimpleID=<%=gadoRS("CarSimpleID")%>' " <%=iPermission3%>>-->
          
   <!--       <input type="button" name="btnDel" value="刪除" onclick="DelLaw('LawDetail.asp?tag=del&ItemID=<%=gadoRS("ItemID")%>&CarSimpleID=<%=gadoRS("CarSimpleID")%>');" <%=iPermission4%>> -->
    <!--  </span></td>-->
      </tr>
	<%
				i=i+1
				gadoRS.movenext
			wend
		iPageCount=gadoRS.PageCount
	else
	%>
		<tr><td rowspan =11 align=center>	<font color='red'>查無資料 ....</font></td>
		
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
      <input type=submit name="Submit422" value="上一頁" onclick="javascript:document.Law.Page.value='<%=cint(Page)-1%>' ;">
      <input type=hidden name='Page' value='' >
      <%=Page%>/<%=iPageCount%>
      <input type=submit name="Submit42" value="下一頁" onclick="javascript:document.Law.Page.value='<%=cint(Page)+1%>' ;">
        <span class="style3"><img src="space.gif" width="13" height="8"></span>       
         <input type="button" name="SaveAs" value="轉換成Excel" onclick="openExcel('LawExcel.asp','LawExcel')">
        <span class="style3"><img src="space.gif" width="13" height="8"></span>        <span class="style2"><span class="style3">
				<input type="button" name="goback" value="回到前一頁" onclick="javascript: document.location.href='index.asp'">&nbsp;   </span></span></p>  
 </td>
 
  </tr>
  <tr>
  <td>
 <b>歸責對象:</b><br>V 歸車   0駕駛人 <br>
 <b>吊扣吊銷:</b><br>
  0無吊扣銷 <img src="space.gif" width="10" height="8">
  2吊銷<img src="space.gif" width="10" height="8">
  3註銷<img src="space.gif" width="10" height="8">
  4逕行註銷<img src="space.gif" width="10" height="8">
  <br>
  A吊扣一月 <img src="space.gif" width="10" height="8">
  B吊扣二月 <img src="space.gif" width="10" height="8">
  C吊扣三月 <img src="space.gif" width="10" height="8">
  D吊扣四月 <img src="space.gif" width="10" height="8">
  E吊扣五月 <img src="space.gif" width="10" height="8">
  F吊扣六月 <img src="space.gif" width="10" height="8"><br>
  G吊扣七月 <img src="space.gif" width="10" height="8">
  H吊扣八月 <img src="space.gif" width="10" height="8">
  I吊扣九月 <img src="space.gif" width="10" height="8">
  J吊扣10月 <img src="space.gif" width="10" height="8">
  K吊扣11月 <img src="space.gif" width="10" height="8">
  L吊扣12月 <img src="space.gif" width="10" height="8">
  M吊扣13月 <img src="space.gif" width="10" height="8"><br>
  N吊扣14月 <img src="space.gif" width="10" height="8">
  O吊扣15月 <img src="space.gif" width="10" height="8">
  X吊扣24個月 <br>
	<b> 禁考註記 </b> <br>
	 0 無禁考 ; <img src="space.gif" width="10" height="8">
	1 禁考1年 ;<img src="space.gif" width="10" height="8">
	2 禁考2年; <img src="space.gif" width="10" height="8">
	3 禁考3年 ; <img src="space.gif" width="10" height="8">
	4 禁考4年 ; <img src="space.gif" width="10" height="8">
	9 永久禁考<br>
<b>特殊處罰</b> <br>
 	0 無;<img src="space.gif" width="10" height="8">
	1 車輛沒入 <img src="space.gif" width="10" height="8">
	2 測速雷達感應器沒入<img src="space.gif" width="10" height="8">
	3 噪音器物沒入<img src="space.gif" width="10" height="8">
	4 責令檢驗; <img src="space.gif" width="10" height="8">
	5 責令臨時檢驗; <img src="space.gif" width="10" height="8">
	6 施以道安講習 ;<img src="space.gif" width="10" height="8">
	<br>
	7 扣繳駕照<img src="space.gif" width="10" height="8">
	8 扣繳牌照<img src="space.gif" width="10" height="8">
	9 吊扣牌照至檢驗合格發還<img src="space.gif" width="10" height="8">
	a 吊扣駕照至參加講習後發還<img src="space.gif" width="10" height="8">
	b 記點2點+記1次<img src="space.gif" width="10" height="8">
	c 執業登記證    <br>
    
  </td>
  </tr>
  <tr>
  
    <td>    <p align="center">&nbsp;
      </p>    <p>　</p>
    <p>　</p></td></tr>
    <td>
</table>
</form>

</body>
</html>
