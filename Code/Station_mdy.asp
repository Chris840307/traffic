
<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
tag = UCase(Request("tag"))
StationSN = Request("StationSN") 
DCIStationID = Trim(Request("DCIStationID")) 
DCIStationName = Trim(Request("DCIStationName"))
StationName = Trim(Request("StationName")) 
StationID = Trim(Request("StationID")) 
StationTel = Trim(Request("StationTel"))
StationAddress = Trim(Request("StationAddress")) 

Select Case tag
	Case "NEW" :	
	  sql = "select Max(StationSN) as StationSN from Station"
	  'set RsTemp = Server.CreateObject("ADODB.RecordSet")
	  Set RsTemp = Conn.Execute(sql)
	  STATION_SN = Cint(RsTemp("StationSN"))+1
	  RsTemp.close

	  sql = "Insert Into Station (StationSN,DCIStationID,DCIStationName,StationName,StationID,StationTel,StationAddress) " & _
	        "Values ("&STATION_SN&",'" & DCIStationID & "','" & DCIStationName & "','" & StationName & "','" & StationID & "'," & _
	        "'" & StationTel & "','" & StationAddress & "')"
	  Conn.Execute(sql)

	Case "UPD" : 
	   sql = "Update Station Set DCIStationID='" & DCIStationID & "',DCIStationName='" & DCIStationName & "',StationName='" & StationName & "' ," & _
	         "StationID='" & StationID & "',StationTel='" & StationTel & "',StationAddress='" & StationAddress & "' Where StationSN=" & StationSN
	   Conn.Execute(sql) 

	Case "DEL" :
	   sql = "Delete From Station Where StationSN=" & StationSN
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
				<form name=StationAdd method=post action=StationAdd.asp>
					<input type=hidden name="tag" value="new">
					<input type=hidden name="StationSN" value="<%=request("StationSN")%>">
					<input type=hidden name="DCIStationID" value="<%=request("DCIStationID")%>">
					<input type=hidden name="DCIStationName" value="<%=request("DCIStationName")%>">
					<input type=hidden name="StationName" value="<%=request("StationName")%>">
					<input type=hidden name="StationID" value="<%=request("StationID")%>">
					<input type=hidden name="StationID" value="<%=request("StationID")%>">
					<input type=hidden name="StationAddress" value="<%=request("StationAddress")%>">
				</form>
			  <Script Language=JavaScript>window.StationAdd.submit();</Script>	
<%
   elseif tag="UPD" then
%>
				<form name=StationUpdate method=post action=StationUpdate.asp>
					<input type=hidden name="tag" value="new">
					<input type=hidden name="StationSN" value="<%=request("StationSN")%>">
					<input type=hidden name="DCIStationID" value="<%=request("DCIStationID")%>">
					<input type=hidden name="DCIStationName" value="<%=request("DCIStationName")%>">
					<input type=hidden name="StationName" value="<%=request("StationName")%>">
					<input type=hidden name="StationID" value="<%=request("StationID")%>">
					<input type=hidden name="StationID" value="<%=request("StationID")%>">
					<input type=hidden name="StationAddress" value="<%=request("StationAddress")%>">
				</form>
				<Script Language=JavaScript>window.StationUpdate.submit();</Script>
<%      
   end if
end if

IF( Session("Msg") = "") THEN
	   response.write "<script>window.opener.location.reload();window.close();</script>"
END If
%>
