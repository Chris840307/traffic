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

If Session("User_ID")="" Then
	Response.write "系統Session值已過期,請重新登入!"
  Response.End
End If	

iTag= Request.querystring("tag")
iID= Request.querystring("ID")
Set gadoRS = Server.CreateObject("ADODB.Recordset")

	if iTag="del" then

			iDelMember="null"
					
			gsql="update Notice set "
			gsql=gsql & " RecordStateID='-1' "
			gsql=gsql & " ,RecordDate=sysdate "
			gsql=gsql & " ,RecordMemID='"&Session("User_ID")&"' "
			gsql=gsql & " ,DelMemID ="&iDelMember
			gsql=gsql & " where ID ='"&trim(iID)&"'"
		
			conn.execute (gsql)

				response.redirect "Notice.asp"		
	end if
	
	if Request.querystring("Save")="Y" then
	
		if iTag="new" then
			if trim(Request("RecordStateID"))="0" then
				iDelMember="null"
			else
				iDelMember="'"&Session("User_ID")&"'"
			end if
			
			gsql="select nvl(max(ID)+1,1) as ID from Notice"
			set gadoRS=conn.execute(gsql)
			  ID=gadoRS("ID")
	     	 gadoRS.close  

				gsql="insert into Notice( ID , NoticeData, StartDate , EndDate , RecordStateID, RecordDate , RecordMemID , DelMemID)"
				gsql=gsql & " values (" & ID & " ,'"&replace(trim(Request("NoticeData")),"'","''")&"',"
				gsql=gsql & funGetDate(gOutDT(trim(Request("StartDate"))),0) &","&funGetDate(gOutDT(trim(Request("EndDate"))),0)&","&trim(Request("RecordStateID"))&",sysdate,'"&Session("User_ID")&"',"&iDelMember&") "
		'---------------------------------------------------------------------------------------		
		'response.write gsql
		'response.end
				conn.execute gsql
			
				response.redirect "Notice.asp"
	
		else
		
			if trim(Request("RecordStateID"))="0" then
				iDelMember="null"
			else
				iDelMember="'"&Session("User_ID")&"'"
			end if
					
			gsql="update Notice set NoticeData ='"&replace(trim(Request("NoticeData")),"'","''")&"' "
			gsql=gsql & " , StartDate ="&funGetDate(gOutDT(trim(Request("StartDate"))),0)
			gsql=gsql & " , EndDate ="&funGetDate(gOutDT(trim(Request("EndDate"))),0)
			gsql=gsql & " ,RecordStateID='"&trim(Request("RecordStateID"))&"' "
			gsql=gsql & " ,RecordDate=sysdate "
			gsql=gsql & " ,RecordMemID='"&Session("User_ID")&"' "
			gsql=gsql & " ,DelMemID ="&iDelMember
			gsql=gsql & " where ID ='"&trim(Request("mdyID"))&"'"
		'response.write gsql
			conn.execute (gsql)
				'
				'response.end
				response.redirect "Notice.asp"		
		end if
	
	end if
		
	if itag="mdy" and iID <>"" and Request.querystring("Save")="" then

		gsql="select NoticeData, StartDate,EndDate,RecordStateID ,(decode (RecordStateID,'0','開始','停止') ) as RecordStateDesc,RecordDate "
		gsql=gsql & " from Notice where ID="&iID
		
		set gadoRS=conn.execute(gsql)
		if not gadoRS.eof then
			iNoticeData=gadoRS("NoticeData")
			iStartDate=gadoRS("StartDate")
			iRecordStateDesc=gadoRS("RecordStateDesc")			
			iEndDate=gadoRS("EndDate")
			iRecordStateID=gadoRS("RecordStateID")

		end if
		gadoRS.close
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>公告訊息維護</title>
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


<%
gsql="select * from Notice "
gadoRS.open gsql,conn,3,1


%>
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<script language="javascript">


function  datacheck(){

	if (document.NoticeData.name.value==""){
		alert ('請輸入公告訊息...!!');
		return false
	}


	if (document.Notice.StartDate.value==""){
		alert ('請輸入專案施行期間...!!');
		return false
	}

	if (document.Notice.EndDate.value==""){
		alert ('請輸入專案施行期間...!!')
		return false
	}	

var DateS =new String(document.Notice.StartDate.value)	;
var DateE =new String(document.Notice.EndDate.value) ;
DateS =new Date(DateS)
DateE =new Date(DateE)
	if (DateS > DateE){	
		alert ('起始日期大於結束日期...!!');
		return false
	}	

}

</script>

<body>
<FORM NAME="Notice" ACTION="NoticeDetail.asp?tag=<%=iTag%>&Save=Y" METHOD="POST" onSubmit="return datacheck();">
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="style3">公告訊息維護</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC">
    <table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF" height="127">
      <tr>
      	<%	if iID<>"" then 
      			idisabled=" disabled"
      	
      		end if
      	%>
        <td width="11%" bgcolor="#FFFFCC" height="21"><div align="right" class="style3">
			公告訊息內容</div></td>
        <td width="89%" height="21"><span class="style3">
          <input name="NoticeData" type="text" value="<%=iNoticeData%>" size="117" maxlength="100" class="btn1">
          <input name="mdyID" type="hidden" value="<%=iID%>" size="10"   class="btn1">          
          </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC" height="41"><div align="right"><span class="style3">公告訊息期間</span></div></td>
        <td height="41"><span class="style3">
<input type='text' size='10' id='StartDate' class="btn1" name='StartDate' value="<%=gInitDT(iStartDate)%>" readonly onclick="OpenWindow('StartDate')"><input type=button value="..." name='btnDateS'  onclick="OpenWindow('StartDate')">      ~
<input type='text' size='10' id='EndDate'  class="btn1"name='EndDate' value="<%=gInitDT(iEndDate)%>" readonly onclick="OpenWindow('EndDate')"><input type=button value="..." name='btnDateE'  onclick="OpenWindow('EndDate')">
        </span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC" height="31"><div align="right"><span class="style3">狀態</span></div></td>
        <td height="31"><span class="style3">
          <select name="RecordStateID">
          	<%if trim(iRecordStateID&" ")="-1" then %>
              <option value='-1'>停止</option>            
              <option value='0'>開始</option>            

            <%else%>  
              <option value='0'>開始</option>
              <option value='-1'>停止</option>
            
            <%end if%>  
            </select>
        </span></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1"><a href="file:///C|/Documents%20and%20Settings/Smith/&#26700;&#38754;/&#31995;&#32113;&#35498;&#26126;/&#38936;&#21934;&#31649;&#29702;&#31995;&#32113;/sssss">
    </a>
        <input type="submit" name="Submit423" value="確 定" <%=iPermission2%>>

        <span class="style3"><img src="space.gif" width="9" height="8"></span>        <input type="button" name="btnBack" value="關 閉" onclick="javascript:document.location.href='Notice.asp'">
</p>    </td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    </td></tr>
</table>

<%
	'gadoRS.close
	Set gadoRS=nothing
	Conn.close
	Set Conn=nothing

%>
</body>
</html>