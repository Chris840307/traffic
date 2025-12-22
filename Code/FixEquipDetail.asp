<!-- #include file="..\Common\db.ini" -->
<!-- #include file="..\Common\AllFunction.inc" -->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
If Session("User_ID")="" Then
	Response.write "系統Session值已過期,請重新登入!"
  Response.End
End If	
Set gadoRS = Server.CreateObject("ADODB.Recordset")
Set iadoRS = Server.CreateObject("ADODB.Recordset")

iTag= Request.querystring("tag")
iEquipmentID= Request.querystring("EquipmentID")

	if iTag="del" then


					
			gsql="update FixEquip set "
			gsql=gsql & " RecordStateID='-1' "
			gsql=gsql & " ,RecordDate =sysdate "
			gsql=gsql & " ,RecordMemberID ='"&Session("User_ID")&"' "		
			gsql=gsql & " ,DelMemberID ='"&Session("User_ID")&"' "
			gsql=gsql & " where EquipmentID ='"&trim(iEquipmentID )&"'"

			
		
			conn.execute (gsql)
				'response.write gsql
				'response.end
				response.redirect "FixEquip.asp"		
	end if


	if Request.querystring("Save")="Y" then
	
		if iTag="new" then
			gsql="select EquipmentID from FixEquip where EquipmentID ='"&trim(Request("EquipmentID"))&"' "
			set gadoRS=conn.execute(gsql)
			if gadoRS.eof then


				gsql="insert into FixEquip (EquipmentID,TypeID ,Address,State ,ImageIP,VideoIP,OCIP,RecordStateID,RecordDate,RecordMemberID,StreetID)"
				gsql=gsql & " values('"&trim(Request("EquipmentID"))&"','"&trim(Request("TypeID"))&"','"&replace(trim(Request("Address")),"'","''")&"','"
				gsql=gsql & trim(Request("State"))&"','"&trim(Request("ImageIP"))&"','"&trim(Request("VideoIP"))&"','"&trim(Request("OCIP"))&"','0',sysdate,'"&Session("User_ID")&"','"&trim(Request("StreetID"))&"') "

				'response.write gsql
				'response.end
				conn.execute gsql
			
				response.redirect "FixEquip.asp"
			
			
			else
				   response.write "<script>alert('專案代碼已存在');history.back() ;</script>"
				   response.write "<script>history.back() ;</script>"				   
				 gadoRS.close  
			
			end if

		else
				gsql="update FixEquip set TypeID ='"&trim(Request("TypeID"))&"' "
				gsql=gsql & " , Address ='"&replace(trim(Request("Address")),"'","''")&"' "
				gsql=gsql & " , State='"&trim(Request("State"))&"' "
				gsql=gsql & " , ImageIP='"&trim(Request("ImageIP"))&"' "
				gsql=gsql & " , VideoIP='"&trim(Request("VideoIP"))&"' "
				gsql=gsql & " , OCIP='"&trim(Request("OCIP"))&"' "
				gsql=gsql & " , StreetID='"&trim(Request("StreetID"))&"' "				
				gsql=gsql & " ,RecordDate =sysdate "
				gsql=gsql & " ,RecordMemberID ='"&Session("User_ID")&"' "
				gsql=gsql & " where EquipmentID ='"&trim(Request("mdyEquipmentID"))&"' "
				
				conn.execute (gsql)
				response.redirect "FixEquip.asp"
				

		end if
	end if
	
	if itag="mdy" and iEquipmentID <>"" and Request.querystring("Save")="" then

		gsql="select EquipmentID,TypeID ,Address,State ,ImageIP,VideoIP,OCIP,StreetID  "
		gsql=gsql & " from FixEquip where EquipmentID='"&iEquipmentID&"'"
		set gadoRS=conn.execute(gsql)
		if not gadoRS.eof then
			iTypeID=gadoRS("TypeID")
			iAddress=gadoRS("Address")
			iState=gadoRS("State")
			iImageIP=gadoRS("ImageIP")
			iVideoIP=gadoRS("VideoIP")
			iOCIP=gadoRS("OCIP")			
			iStreetID=gadoRS("StreetID")			

		end if
		gadoRS.close
	end if	
	
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<title>固定桿資料維護</title>

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


<%
gsql="select EquipmentID from FixEquip "
gadoRS.open gsql,conn,3,1



%>

<script language="javascript">

var id = new Array(<%=gadoRS.recordcount%>); 


<%
	if not gadoRS.eof then
		k=0
		while not gadoRS.eof
%>
		id[<%=k%>]="<%=gadoRS("EquipmentID")%>"; 
<%
			k=k+1
			gadoRS.movenext
		wend
	end if
	gadoRS.close
%>
function  datacheck(){
<%if itag<>"mdy" then %>
	if (document.FixEquipDetail.EquipmentID.value==""){
		alert ('固定桿代碼...!!')
		return false
	}
	
	if (document.FixEquipDetail.StreetID.value==""){
		alert ('路段代碼...!!')
		return false
	}

	
	
	if (document.FixEquipDetail.EquipmentID.value!=""){
           for( j=0; j<id.length-1; j++ )
            {		
				if (id[j]==document.FixEquipDetail.EquipmentID.value){
				alert ('固定桿代碼已存在...!!')				
				return false ;
				}
			}
	}

<%end if%>

	if (document.FixEquipDetail.Address.value==""){
		alert ('請輸入地點...!!')
		return false
	}

}

function checkEquipID(obj){

   for( j=0; j<id.length; j++ )
    {		
		if (id[j]==obj.value){
		alert ('固定桿代碼已存在...!!')				
		return false ;
		}
	}


}
function getAddress(obj){

	document.FixEquipDetail.Address.value=obj.options[obj.selectedIndex].text ;
	document.FixEquipDetail.Address1.value=obj.options[obj.selectedIndex].text ;
}

</script>
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

<body>
<FORM NAME="FixEquipDetail" ACTION="FixEquipDetail.asp?tag=<%=iTag%>&Save=Y" METHOD="POST" onSubmit="return datacheck();">
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="style3">固定桿資料維護</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="13%" bgcolor="#FFFFCC"><div align="right" class="style3">
          <div align="right">固定桿代碼</div>
        </div></td>
        <td width="87%"><span class="style3">
<input name="EquipmentID" type="text" value="<%=iEquipmentID%>" size="13" maxlength="12" onchange="checkEquipID(this)"  class="btn1">
<input name="mdyEquipmentID" type="hidden" value="<%=iEquipmentID%>" size="10"  >          
        </span></td>
      </tr>
<%

Set iadoRS = Server.CreateObject("ADODB.Recordset")

	gsql=" select StreetID,Address from Street order by Address "
	iadoRS.open gsql,conn,3,1

%>      
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">路段代碼</span></div></td>
        <td><span class="style3">
          <select name='StreetID' onchange="getAddress(this)">
           				<option value=''>選擇路段代碼</option>           
           			<%if not iadoRS.eof then
           				while not iadoRS.eof
           				if trim(iStreetID)=trim(iadoRS(0)) then 
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
           			</select>                </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">地點</span></div></td>
        <td><span class="style3">
          <input name="Address1" type="text" value="<%=iAddress%>" size="41" maxlength="40"  class="btn1" disabled>
          <input name="Address" type="hidden" value="<%=iAddress%>" size="41" maxlength="40"  class="btn1">
        </span></td>
      </tr>
<%


gsql="select ID ,Content from Code where TypeID ='18' "
iadoRS.open gsql,conn,3,1

%>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">違規影像位置</span></div></td>
        <td><span class="style3">
          <input name="ImageIP" type="text" value="<%=iImageIP%>" size="13" maxlength="12"  class="btn1">
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">即時影像位置</span></div></td>
        <td><span class="style3">
          <input name="VideoIP" type="text" value="<%=iVideoIP%>" size="13" maxlength="12"  class="btn1">
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">OC位置</span></div></td>
        <td><span class="style3">
          <input name="OCIP" type="text" value="<%=iOCIP%>" size="13" maxlength="12"  class="btn1">
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">類型</span></div></td>
        <td><span class="style3">
          <select name="TypeID">
<%
		if not iadoRS.eof then
			while not iadoRS.eof
				if trim(iTypeID)=trim(iadoRS("id")) then
					iselected =" selected"
				else
					iselected =" "
				end if
%>          
            <option value ='<%=iadoRS("id")%>' <%=iselected%>><%=iadoRS("Content")%></option>
            
<%
				iadoRS.movenext
			wend
		end if

%>            
          </select>
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">使用狀態</span></div></td>
        <td><span class="style3">
          <select name="State">
          <%if iState="0" then %>
            <option value ='0'>空桿</option>
            <option value ='1'>使用中</option>            
		  <%else%>

            <option value ='1'>使用中</option>            
            <option value ='0'>空桿</option>
		  <%end if%>
          </select>
        </span></td>
      </tr>
      
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1">
    </a>
        <input type="submit" name="Submit423" value="確 定" <%=iPermission2%>>
        <span class="style3"><img src="space.gif" width="9" height="8"></span>       <input type="button" name="btnBack" value="關 閉" onclick="javascript:document.location.href='FixEquip.asp'"></p>    </td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    <p>　</p>
    <p>　</p></td></tr>
</table>
</body>

<%
	iadoRS.close
	Set iadoRS=nothing


	Set gadoRS=nothing
	Conn.close
	Set Conn=nothing


%></html>