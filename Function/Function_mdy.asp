<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%

tag = UCase(Request("tag"))
SN = Request("SN")
GroupID = Request("GroupID")
SystemID = Request("SystemID")

InsertFlag=Request("Insert")
UpdateFlag= Request("Update")
DeleteFlag= Request("Delete")
SelectFlag=Request("Select")
 
Select Case tag
	Case "NEW" :
	  sql = "select GroupID,SystemID from FunctionDataDetail where GroupID=" & GroupID & " And SystemID= " & SystemID 
	  set Rs = Server.CreateObject("ADODB.RecordSet")
	  set Rs = Conn.Execute(sql)
      if not Rs.eof then
    	 Session("Msg") = "...該筆權限設定已經存在"
	     Rs.Close
	     Set Rs = Nothing
	  else  
		  sql = "select GroupID,SystemID from FunctionData where GroupID=" & GroupID & " And SystemID= "& SystemID 
		  set Rs = Server.CreateObject("ADODB.RecordSet")
		  Set Rs = Conn.Execute(sql)
		  if Rs.eof then 
			//	
			sql = "Insert into FunctionData (GroupID,SYSTEMID,Function) values(" & GroupID & "," & SystemID  & ", 1)"
			Conn.Execute(sql)			 
		  end if	  
			 Rs.Close
			 Set Rs = Nothing					  
		  '----------------------------------------------------------------------
	      sql = "Insert into FunctionDataDetail (SN,GROUPID,SYSTEMID,INSERTFLAG,UPDATEFLAG,DELETEFLAG,SELECTFLAG) " & _
	            "values (FUNCTIONDATADETAIL_SN.NEXTVAL," & GROUPID & "," & SYSTEMID & ",'" & INSERTFLAG & "','" & UPDATEFLAG & "','" & DELETEFLAG & "','"& SELECTFLAG &"')"  
		  Conn.Execute(sql)			
	  end if
	Case "UPD":	
	  sql = "Update FunctionDataDetail Set GROUPID="& GROUPID & ",SYSTEMID=" & SYSTEMID &  _
			",INSERTFLAG='" & INSERTFLAG & "',UPDATEFLAG='" & UPDATEFLAG & "',DELETEFLAG='" & DELETEFLAG & "'" & _
			",SELECTFLAG='" & SELECTFLAG &  "' Where SN=" & SN 
	  Conn.Execute(sql)
  Case "DEL":  
		  'sql = "select SN from FunctionDataDetail Where GroupID=" & GroupID & " and SystemID ="& SystemID
		  'set Rs = Server.CreateObject("ADODB.RecordSet")
		  'Set Rs = Conn.Execute(sql)
		  'if Rs.recordcount=1 then 
		  	sql1 = "Delete From FunctionData Where GroupID=" & GroupID & " and SystemID ="& SystemID
				Conn.Execute(sql1)				  
		  'end if  
		  'Rs.Close
	    'Set Rs = Nothing
    	sql1 = "Delete From FunctionDataDetail Where GroupID=" & GroupID & " and SystemID ="& SystemID
    	Conn.Execute(sql1)
	
End Select	
IF( err.number<>0) THEN
	 Session("Msg") = Session("Msg") & "<br>????" & err.description
end if	
%>
<!-- #include file="../Common/ClearObject.asp" -->
<%
if Session("msg") <> "" then
	if tag="NEW" then
		  response.contenttype = "text/html;charset=big5"			
%>
				<form name=FunctionAdd method=post action=FunctionAdd.asp>
					<input type=hidden name=tag value="new">
					<input type=hidden name=SN value="<%=request("SN")%>">
				</form>
			  <Script Language=JavaScript>window.FunctionAdd.submit();</Script>	
<%
   elseif tag="UPD" then
%>
				<form name=FunctionUpdate method=post action=FunctionUpdate.asp>
					<input type=hidden name=tag value="UPD">
					<input type=hidden name=SN value="<%=request("SN")%>">
				</form>
				<Script Language=JavaScript>window.FunctionUpdate.submit();</Script>
<%      
   end if
end if


IF( Session("Msg") = "") THEN
	  response.write "<script>window.opener.location.reload();window.close();</script>"
END If

%>
