<!-- #include file="..\Common\db.ini" -->
<!-- #include file="..\Common\AllFunction.inc" -->
<!--#include virtual="traffic/Common/Login_Check.asp"-->

<%
If Session("User_ID")="" Then
	Response.write "系統Session值已過期,請重新登入!"
  Response.End
End If	
Set gadoRS = Server.CreateObject("ADODB.Recordset")

iTag= Request.querystring("tag")
iItemID= Request.querystring("ItemID")
iCarSimpleID = Request.querystring("CarSimpleID")


	if iTag="del" then
	
		gsql="update Law set "
		gsql=gsql & " RecordStateID='-1' "
		gsql=gsql & " ,ModifyTime=sysdate "
		gsql=gsql & " ,RecordMemberID ='"&Session("User_ID")&"' "		
		gsql=gsql & " ,DelMemberID ='"&Session("User_ID")&"' "

		gsql=gsql & " where ItemID='"&iItemID&"' and CarSimpleID  ='"&iCarSimpleID &"'"
		conn.execute gsql

				response.redirect "Law.asp"	
	end if

	if Request.querystring("Save")="Y" then
	
		if iTag="new" then



				gsql="insert into Law(ItemID,CarSimpleID ,IllegalRule,Level1,Level2,Level3,Level4,Target "
				gsql=gsql & " ,RecordPoint,RevokePoint,NoTest,Retain,SpecPunish,ModifyTime,RecordStateID,RecordDate,RecordMemberID) "
				gsql=gsql & " values('"&trim(Request("ItemID"))&"','"&trim(Request("CarSimpleID"))&"','"&replace(trim(Request("IllegalRule")&" "),"'","''")&"','"
				gsql=gsql & trim(Request("Level1"))&"','"&trim(Request("Level2"))&"','"
				gsql=gsql & trim(Request("Level3"))&"','"&trim(Request("Level4"))&"','"
				gsql=gsql & trim(Request("Target"))&"','"&trim(Request("RecordPoint"))&"','"
				gsql=gsql & trim(Request("RevokePoint"))&"','"&trim(Request("NoTest"))&"','"								
				gsql=gsql & trim(Request("Retain"))&"','"&trim(Request("SpecPunish"))&"',sysdate,'0',sysdate,'"&Session("User_ID")&"') "												
				'response.write gsql
				'response.end
				conn.execute gsql
			
				response.redirect "Law.asp"
			
			


		else
				gsql="update Law set IllegalRule='"&trim(Request("IllegalRule"))&"' "
				gsql=gsql & " , Level1='"&trim(Request("Level1"))&"' "
				gsql=gsql & " , Level2='"&trim(Request("Level2"))&"' "
				gsql=gsql & " , Level3='"&trim(Request("Level3"))&"' "
				gsql=gsql & " , Level4='"&trim(Request("Level4"))&"' "
				gsql=gsql & " , Target='"&trim(Request("Target"))&"' "
				gsql=gsql & " , RecordPoint='"&trim(Request("RecordPoint"))&"' "
				gsql=gsql & " , RevokePoint='"&trim(Request("RevokePoint"))&"' "
				gsql=gsql & " , NoTest='"&trim(Request("NoTest"))&"' "
				gsql=gsql & " , Retain='"&trim(Request("Retain"))&"' "
				gsql=gsql & " , SpecPunish='"&trim(Request("SpecPunish"))&"' "
				gsql=gsql & " , ModifyTime=sysdate "				
				gsql=gsql & " where ItemID='"&trim(Request("mdyItemID"))&"' "
				gsql=gsql & " and CarSimpleID= '"&trim(Request("mdyCarSimpleID"))&"' "
																
				conn.execute (gsql)
				'response.write gsql
				'response.end
				response.redirect "Law.asp"
				

		end if
	end if

	if itag="mdy" and iItemID <>"" and iCarSimpleID<>"" and Request.querystring("Save")="" then

		gsql="select IllegalRule,Level1,Level2,Level2,Level3,Level4,Target "
		gsql=gsql & " ,RecordPoint,RevokePoint,NoTest,Retain,SpecPunish  "
		gsql=gsql & " from Law where ItemID='"&iItemID&"' and CarSimpleID ='"&iCarSimpleID&"' "
		'response.write gsql 
		set gadoRS=conn.execute(gsql)
		if not gadoRS.eof then

			iIllegalRule=gadoRS("IllegalRule")
			iLevel1=gadoRS("Level1")
			iLevel2=gadoRS("Level2")
			iLevel3=gadoRS("Level3")						
			iLevel4=gadoRS("Level4")
			iTarget=gadoRS("Target")						
			iRecordPoint=gadoRS("RecordPoint")
			iRevokePoint=gadoRS("RevokePoint")						
			iNoTest=gadoRS("NoTest")
			iRetain=gadoRS("Retain")					
			iNoTest=gadoRS("NoTest")
			iSpecPunish=gadoRS("SpecPunish")		

		end if
		gadoRS.close
	end if	



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
.style3 {font-size: 15px}
-->
</style></head>
<script language="javascript">

<%
Set iadoRS = Server.CreateObject("ADODB.Recordset")
  gsql="select ItemID,CarSimpleID from law "
  iadoRS.open gsql,conn,3,1

%>


var LawID = new Array(<%=iadoRS.recordcount%>); 


<%
	if not iadoRS.eof then
		k=0
		while not iadoRS.eof

		response.write "LawID["&k&"]='"&iadoRS("ItemID")&"-"&iadoRS("CarSimpleID")&"'; "

			k=k+1
			iadoRS.movenext
		wend
	end if
	iadoRS.close
	set iadoRS=nothing
%>


function  datacheck(){
<%if itag="new" then %>
	if (document.Law.ItemID.value==""){
		alert ('法條代碼...!!')
		return false
	}
	
	var obj1 =document.Law.ItemID ;
	var obj2 =document.Law.CarSimpleID ;
	
	var KeyID ="" ;
	KeyID =obj1+'-'+obj2 ;
	//alert (KeyID) ;
   for( j=0; j<LawID.length; j++ )
    {		
		if (LawID[j]==KeyID){
		alert ('法條代碼+簡式車種已存在...!!')				
		return false ;
		}
	}	
<%end if%>	



}

function checkKey(obj1,obj2){
	var KeyID ="" ;
	KeyID =obj1.value+'-'+obj2.value ;
	//alert (KeyID) ;
	
	var iIndex=document.Law.OldCarSimpleID.value ;
   for( j=0; j<LawID.length; j++ )
    {		
		if (LawID[j]==KeyID){
		alert ('法條代碼+簡式車種已存在...!!')				
		obj2.options[iIndex].selected=true ;
		return false ;
		}
	}
	document.Law.OldCarSimpleID.value=obj2.options[obj2.selectedIndex].value
}


</script>
<body>
<FORM NAME="Law" ACTION="LawDetail.asp?tag=<%=iTag%>&Save=Y" METHOD="POST" onSubmit="return datacheck();">
<!--
<input type=button name='xxx' value='test' onclick="checkKey(document.Law.ItemID,document.Law.CarSimpleID)">
-->
<input type =hidden name="mdyItemID" type="text" value="<%=iItemID%>" size="10" maxlength="9">
<input type =hidden name="mdyCarSimpleID" type="text" value="<%=iCarSimpleID %>" size="10" maxlength="9">
<table width="100%" height="100%" border="0">



  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="style3">法條檔維護</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
    
<%if itag <>"new" then

	iDisabled =" disabled"


  end if
%>    
      <tr>
        <td width="11%" bgcolor="#FFFFCC"><div align="right" class="style3">法條代碼</div></td>
        <td width="89%"><span class="style3">
          <input name="ItemID" type="text" value="<%=iItemID%>" size="10" maxlength="9" <%=iDisabled%>  class="btn1">
</span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">簡式車種</span></div></td>
        
        <td><span class="style3" >
          <select name="CarSimpleID" onchange="checkKey(document.Law.ItemID,document.Law.CarSimpleID)" <%=iDisabled%>>
          
          	<%
          	select case iCarSimpleID
          		case "0"
          			iselect0 =" selected"
          		case "1"
          			iselect1 =" selected" 
          		case "2"
          			iselect2 =" selected"           		
          		case "3"
          			iselect3 =" selected"           		
          		case "4"
          			iselect4 =" selected"           		
          		case else
          			iCarSimpleID="0"
          			iselect0 =" selected"
          	end select
          	%>
              <option value="0" <%=iselect0%>></option>          
              <option value="1" <%=iselect1%>>汽車</option>
              <option value="2" <%=iselect2%>>拖車</option>
              <option value="3" <%=iselect3%>>重機</option>
              <option value="4" <%=iselect4%>>輕機</option>
            </select>
          </span>
          <input type=hidden name='OldCarSimpleID' value='<%=iCarSimpleID%>'  class="btn1">
          
          </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">違規事實</span></div></td>
        <td><span class="style3">
          <input name="IllegalRule" type="text" value="<%=iIllegalRule%>" size="60" maxlength="100"  class="btn1">
            </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">期限內</span></div></td>
        <td><span class="style3">
          <input name="Level1" type="text" size="6" maxlength="5"value='<%=iLevel1%>' class="btn1" >
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">15天內罰款</span></div></td>
        <td><span class="style3">
          <input name="Level2" type="text" size="6" maxlength="5" value='<%=iLevel2%>'  class="btn1">
</span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">15~30天</span></div></td>
        <td><span class="style3"> 
          <input name="Level3" type="text"  size="6" maxlength="5" value='<%=iLevel3%>'  class="btn1">
</span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">超過30天以上</span></div></td>
        <td><span class="style3">
          <input name="Level4" type="text"  size="6" maxlength="5" value='<%=iLevel4%>'  class="btn1">
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">歸責對象</span></div></td>
        <td><span class="style3">
          <select name="Target">
              <option value="0" >車主</option>          
              <option value="V" >駕駛人</option>
            </select>
            </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">記點記次</span></div></td>
        <td><span class="style3"><input name="RecordPoint" type="text" size="4" maxlength="3" value='<%=iRecordPoint%>'  class="btn1">
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">吊扣吊銷</span></div></td>
        <td><span class="style3">
          <input name="RevokePoint" type="text" size="4" maxlength="3" value='<%=iRevokePoint%>'  class="btn1">
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">禁考註記</span></div></td>
        <td><span class="style3">
          <input name="NoTest" type="text" size="4" maxlength="3" value='<%=iNoTest%>'  class="btn1">
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">保留</span></div></td>
        <td><span class="style3">
          <input name="Retain" type="text" size="4" maxlength="3" value='<%=iRetain%>'  class="btn1">
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3">特殊處罰</span></div></td>
        <td><span class="style3">
          <input name="SpecPunish" type="text" size="4" maxlength="3" value='<%=iSpecPunish%>'  class="btn1">
        </span></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1">
    </a>
        <input type="submit" name="Submit423" value="確 定" <%=iPermission2%>>
        <span class="style3"><img src="space.gif" width="9" height="8"></span>        <input type="button" name="btnBack" value="關 閉" onclick="javascript:document.location.href='Law.asp'"></p>    </td>
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
