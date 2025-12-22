
<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
tag = UCase(Request("tag"))
SN = Request("SN") 
TypeId = Request("TypeId") 
Id = Trim(Request("Id"))
Content = Trim(Request("Content")) 

Select Case tag
	Case "NEW" :
	  sql = "select SN From DciCode Where TypeId=" & TypeId & " And Id='" & Id & "'" 
    Set RsChk = Conn.Execute(sql)	
    if not RsChk.eof then
    	 TypeDesc = GetDciTypeById (TypeId)
    	 Session("Msg") = "代碼類型：【" & TypeDesc & "】且代碼值：【" & Id & "】已經存在，請重新輸入!!"
	  else 	
	     sql = "select DCICODE_SN.NEXTVAL as DCICODE_SN from DUAL"
	     set RsTemp = Server.CreateObject("ADODB.RecordSet")
	     Set RsTemp = Conn.Execute(sql)
	     DCICODE_SN = RsTemp("DCICODE_SN")
	     
	     sql = "Insert Into DciCode (SN,TypeId,Id,Content) Values (" & DCICODE_SN & "," & Int(TypeId) & ",'" & Id & "'," & _
	           "'" & Content & "')"
	     Conn.Execute(sql)
    end if
	Case "UPD" :
	  'sql = "select SN From DciCode Where TypeId=" & TypeId & " And Id='" & Id & "'" 
    'Set RsChk = Conn.Execute(sql)		
    'if not RsChk.eof then
    '	 Session("Msg") = "代碼值：【" & Id & "】已經存在,請重新輸入!!"	  
	  'else
	     sql = "Update DciCode Set TypeId=" & TypeId & ",Id='" & Id & "' ,Content='" & Content & "' Where SN=" & SN
	     Conn.Execute(sql) 
	  'end if
	Case "DEL" :
	   sql = "Delete From DciCode Where SN=" & SN
	   Conn.Execute(sql)
End Select

IF( err.number<>0) THEN
	 Session("Msg") = Session("Msg") & "<br>新增/修改失敗,錯誤訊息:" & err.description
end if		   
%>
<!-- #include file="../Common/ClearObject.asp" -->
<%
if Session("msg") <> "" then
	if tag="NEW" then
		  response.contenttype = "text/html;charset=big5"			
%>
				<form name=DCICodeAdd method=post action=DCICodeAdd.asp>
					<input type=hidden name="tag" value="new">
					<input type=hidden name="SN" value="<%=request("SN")%>">
					<input type=hidden name="TypeId" value="<%=request("TypeId")%>">
					<input type=hidden name="Id" value="<%=request("Id")%>">
					<input type=hidden name="Content" value="<%=request("Content")%>">
				</form>
			  <Script Language=JavaScript>window.DCICodeAdd.submit();</Script>	
<%
   elseif tag="UPD" then
%>
				<form name=DCICodeUpdate method=post action=DCICodeUpdate.asp>
					<input type=hidden name=tag value="upd">
					<input type=hidden name="SN" value="<%=request("SN")%>">
					<input type=hidden name="TypeId" value="<%=request("TypeId")%>">
					<input type=hidden name="Id" value="<%=request("Id")%>">
					<input type=hidden name="Content" value="<%=request("Content")%>">
				</form>
				<Script Language=JavaScript>window.DCICodeUpdate.submit();</Script>
<%      
   end if
end if

IF( Session("Msg") = "") THEN
	   response.write "<script>window.opener.location.reload();window.close();</script>"
END If
%>
