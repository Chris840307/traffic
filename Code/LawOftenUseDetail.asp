<!-- #include file="..\Common\db.ini" -->
<!-- #include file="..\Common\AllFunction.inc" -->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
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
iSN= Request.querystring("SN")


	if iTag="del" then
	
		gsql="delete LawOftenUse where SN ='"&iSN&"' "
		conn.execute gsql

				response.redirect "LawOftenUse.asp"	
	end if

	if Request.querystring("Save")="Y" then
	
		if iTag="new" then



				gsql="insert into LawOftenUse(SN,ItemID,BillTypeID,ShowOrder,ModifyDate)"
				gsql=gsql & " values(LawOftenUse_SN.NEXTVAL,'"&trim(Request("ItemID"))&"','"&trim(Request("BillTypeID"))&"','"
				gsql=gsql & trim(Request("ShowOrder"))&"',sysdate) "

				'response.write gsql
				'response.end
				conn.execute gsql
			
				response.redirect "LawOftenUse.asp"
			
			


		else
				gsql="update LawOftenUse set ItemID='"&trim(Request("ItemID"))&"' "
				gsql=gsql & " , BillTypeID='"&trim(Request("BillTypeID"))&"' "
				gsql=gsql & " , ShowOrder='"&trim(Request("ShowOrder"))&"' "
				gsql=gsql & " where SN ='"&trim(Request("SN"))&"' "
				
				conn.execute (gsql)
				response.redirect "LawOftenUse.asp"
				

		end if
	end if

	if itag="mdy" and iSN <>"" and Request.querystring("Save")="" then

		gsql="select SN,ItemID,BillTypeID,ShowOrder  "
		gsql=gsql & " from LawOftenUse where SN='"&iSN&"'"
		set gadoRS=conn.execute(gsql)
		if not gadoRS.eof then
			iItemID=gadoRS("ItemID")
			iBillTypeID=gadoRS("BillTypeID")
			iShowOrder=gadoRS("ShowOrder")

		end if
		gadoRS.close
	end if	



%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>常用法條檔維護</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {font-size: 14px}
.style3 {font-size: 15px}
-->
</style></head>
<script language="javascript">

function  datacheck(){

	if (document.LawOftenUseDetail.ItemID.value==""){
		alert ('法條代碼...!!')
		return false
	}
	
	var obj =document.LawOftenUseDetail.BillTypeID ;	
	
	if (obj.options[obj.selectedIndex].value==""){
		alert ('舉發單類型...!!')
		return false
	}


	if (document.LawOftenUseDetail.ShowOrder.value==""){
		alert ('顯示次序...!!')
		return false
	}



	if(isNaN(document.LawOftenUseDetail.ShowOrder.value)){
		alert(document.LawOftenUseDetail.ShowOrder.value+"不是數字")	
		return false
	}
}


</script>
<body>
<FORM NAME="LawOftenUseDetail" ACTION="LawOftenUseDetail.asp?tag=<%=iTag%>&Save=Y" METHOD="POST" onSubmit="return datacheck();">
<table width="100%" height="100%" border="0">
<input type =hidden name="SN" type="text" value="<%=iSN%>" size="10" maxlength="9">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="style3">常用法條檔維護</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="11%" bgcolor="#FFFFCC"><div align="right" class="style3">法條代碼</div></td>
        <td width="89%"><span class="style3">
<input name="ItemID" type="text" value="<%=iItemID%>" size="10" maxlength="9"  class="btn1">
</span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">舉發單類型 </span></div></td>
        <td><span class="style3">
      <%
Set iadoRS = Server.CreateObject("ADODB.Recordset")
gsql="select ID,Content from DCICODE where TypeID ='2'"
iadoRS.open gsql,conn,3,1
%>            
            <select name="BillTypeID">

            <%if not iadoRS.eof then
            	while not iadoRS.eof 
            	if trim(iBillTypeID)=trim(iadoRS("ID")) then
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
</td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3"> 顯示次序</span></div></td>
        <td><span class="style3">
          <input name="ShowOrder" type="text" value="<%=iShowOrder%>" size="3" maxlength="2"  class="btn1">
        </span></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1">
    </a>
        <input type="submit" name="btnSubmit" value="確 定" <%=iPermission2%>>
        <span class="style3"><img src="space.gif" width="9" height="8"></span>        <input type="button" name="btnBack" value="關 閉" onclick="javascript:document.location.href='LawOftenUse.asp'"></p>    </td></p>    </td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    <p>　</p>
    <p>　</p></td></tr>
</table>
</form>
<%

	Conn.close
	Set Conn=nothing

%>
</body>
</html>
