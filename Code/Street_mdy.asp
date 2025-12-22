
<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
tag = UCase(Request("tag"))
StreetId = Trim(Request("StreetId")) 
StreetSimpleName = Trim(Request("StreetSimpleName")) 
Address = Trim(Request("Address"))
UnitID = Trim(Session("Unit_ID"))

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing
	
Select Case tag
	Case "NEW" :
	  sql = "select StreetID From Street Where StreetID='" & StreetID & "'" 
    Set RsChk = Conn.Execute(sql)	
    if not RsChk.eof then
    	 Session("Msg") = "路段代碼：【" & StreetID & "】已經存在，請重新輸入!!"
	  else 	
		If sys_City="高雄市" Then

			 FixPole=request("FixPole")
			 If Trim(FixPole)="" Then
					 FixPole="null"
			 Else
					 FixPole="1"
			 End if
	     sql = "Insert Into Street (StreetID,Address,UnitID,FixPole) " & _
	           "Values ('" & trim(StreetId) & "','" & Address & "','" & UnitID & "',"&FixPole&")"
		else
	     sql = "Insert Into Street (StreetID,Address,UnitID) " & _
	           "Values ('" & trim(StreetId) & "','" & Address & "','" & UnitID & "')"
	    End if
	     'response.write sql

             Conn.Execute(sql)
    end if
	Case "UPD" :
		If sys_City="高雄市" Then

			 FixPole=request("FixPole")
			 If Trim(FixPole)="" Then
					 FixPole="null"
			 Else
					 FixPole="1"
			 End If
	     sql = "Update Street Set Address='" & Address & "',UnitID='" & UnitID & "',FixPole=" & FixPole & " Where StreetID='" & trim(StreetID) & "'"
		else
	     sql = "Update Street Set Address='" & Address & "',UnitID='" & UnitID & "' Where StreetID='" & trim(StreetID) & "'"
		End if
	     Conn.Execute(sql) 
	Case "DEL" :
	   sql = "Delete From Street Where StreetId='" & trim(StreetId) & "'"
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
				<form name=StreetAdd method=post action=StreetAdd.asp>
					<input type=hidden name="tag" value="new">
					<input type=hidden name="StreetID" value="<%=request("StreetID")%>">
					<input type=hidden name="StreetSimpleName" value="<%=request("StreetSimpleName")%>">
					<input type=hidden name="Address" value="<%=request("Address")%>">
				</form>
			  <Script Language=JavaScript>window.StreetAdd.submit();</Script>	
<%
   elseif tag="UPD" then
%>
				<form name=StreetUpdate method=post action=StreetUpdate.asp>
					<input type=hidden name=tag value="upd">
					<input type=hidden name="StreetID" value="<%=request("StreetID")%>">
					<input type=hidden name="StreetSimpleName" value="<%=request("StreetSimpleName")%>">
					<input type=hidden name="Address" value="<%=request("Address")%>">
				</form>
				<Script Language=JavaScript>window.StreetUpdate.submit();</Script>
<%      
   end if
end if

IF( Session("Msg") = "") THEN
	   response.write "<script language='javascript'>alert('修改完成');window.opener.Street.submit();window.close();</script>"
END If
%>