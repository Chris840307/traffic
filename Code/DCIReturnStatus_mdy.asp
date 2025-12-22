
<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
tag = UCase(Request("tag"))
SN = Request("SN") 
DCIActionID = Request("DCIActionID") 
DCIReturn = Trim(Request("DCIReturn"))
StatusContent = Trim(Request("StatusContent")) 
DCIReturnStatus = Request("DCIReturnStatus") 
NeedReDo = Trim(Request("NeedReDo"))
HowTo = Trim(Request("HowTo")) 

Select Case tag
	Case "NEW" :
	  sql = "select SN From DciReturnStatus Where DCIActionID='" & DCIActionID & "' And DCIReturn='" & DCIReturn & "'" 
    Set RsChk = Conn.Execute(sql)	
    if not RsChk.eof then
    	 Session("Msg") = "DCI資料交換類型：【" & DCIActionID & "】且回傳狀態代碼：【" & DCIReturn & "】已經存在，請重新輸入!!"
	  else 	
	     sql = "select DCIRETURNSTATUS_SN.NEXTVAL as DCIRETURNSTATUS_SN from DUAL"
	     set RsTemp = Server.CreateObject("ADODB.RecordSet")
	     Set RsTemp = Conn.Execute(sql)
	     DCIRETURNSTATUS_SN = RsTemp("DCIRETURNSTATUS_SN")
       DCIActionName = GetDCIActionNameById (DCIActionID)	     
	     sql = "Insert Into DciReturnStatus (SN,DCIActionID,DCIActionName,DCIReturn,DCIReturnStatus,StatusContent,NeedReDo,HowTo) " & _
	           "Values (" & DCIRETURNSTATUS_SN & ",'" & DCIActionID & "','" & DCIActionName & "','" & DCIReturn & "'," & Int(DCIReturnStatus) & "," & _
	           "'" & StatusContent & "'," & Int(NeedReDo) & ",'" & HowTo & "')"
	     Conn.Execute(sql)
    end if
	Case "UPD" :
	   DCIActionName = GetDCIActionNameById (DCIActionID)	   
	  'sql = "select SN From DciReturnStatus Where DCIActionID='" & DCIActionID & "' And DCIReturn='" & DCIReturn & "'" 
    'Set RsChk = Conn.Execute(sql)		
    'if not RsChk.eof then
    '	 Session("Msg") = "DCI資料交換類型：【" & DCIActionID & "】且回傳狀態代碼：【" & DCIReturn & "】已經存在，請重新輸入!!"
	  'else
	     sql = "Update DciReturnStatus Set DCIActionID='" & DCIActionID & "',DCIActionName='" & DCIActionName & "',DCIReturn='" & DCIReturn & "' ," & _
	           "StatusContent='" & StatusContent & "',DCIReturnStatus=" & DCIReturnStatus & "," & _
	           "NeedReDo=" & NeedReDo & ",HowTo='" & HowTo & "' Where SN=" & SN
	     Conn.Execute(sql) 
	  'end if
	Case "DEL" :
	   sql = "Delete From DciReturnStatus Where SN=" & SN
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
				<form name=DCIReturnStatusAdd method=post action=DCIReturnStatusAdd.asp>
					<input type=hidden name="tag" value="new">
					<input type=hidden name="SN" value="<%=request("SN")%>">
					<input type=hidden name="DCIActionID" value="<%=request("DCIActionID")%>">
					<input type=hidden name="DCIReturn" value="<%=request("DCIReturn")%>">
					<input type=hidden name="StatusContent" value="<%=request("StatusContent")%>">
					<input type=hidden name="NeedReDo" value="<%=request("NeedReDo")%>">
					<input type=hidden name="DCIReturnStatus" value="<%=request("DCIReturnStatus")%>">
					<input type=hidden name="HowTo" value="<%=request("HowTo")%>">
				</form>
			  <Script Language=JavaScript>window.DCIReturnStatusAdd.submit();</Script>	
<%
   elseif tag="UPD" then
%>
				<form name=DCIReturnStatusUpdate method=post action=DCIReturnStatusUpdate.asp>
					<input type=hidden name=tag value="upd">
					<input type=hidden name="SN" value="<%=request("SN")%>">
					<input type=hidden name="DCIActionID" value="<%=request("DCIActionID")%>">
					<input type=hidden name="DCIReturn" value="<%=request("DCIReturn")%>">
					<input type=hidden name="StatusContent" value="<%=request("StatusContent")%>">
					<input type=hidden name="NeedReDo" value="<%=request("NeedReDo")%>">
					<input type=hidden name="DCIReturnStatus" value="<%=request("DCIReturnStatus")%>">
					<input type=hidden name="HowTo" value="<%=request("HowTo")%>">
				</form>
				<Script Language=JavaScript>window.DCIReturnStatusUpdate.submit();</Script>
<%      
   end if
end if

IF( Session("Msg") = "") THEN
	   response.write "<script>window.opener.location.reload();window.close();</script>"
END If
%>
