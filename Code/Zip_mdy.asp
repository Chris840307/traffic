
<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
tag = UCase(Request("tag"))
ZipID = Trim(Request("ZipID")) 
ZipName = Trim(Request("ZipName")) 
ZipNo = Trim(Request("ZipNo"))

Select Case tag
	Case "NEW" :
	  sql = "select ZipID From Zip Where ZipID='" & ZipID & "'" 
    Set RsChk = Conn.Execute(sql)	
    if not RsChk.eof then
    	 Session("Msg") = "郵遞區號：【" & ZipID & "】已經存在，請重新輸入!!"
	  else 	
	     sql = "Insert Into Zip (ZipID,ZipName,ZipNo) " & _
	           "Values ('" & ZipID & "','" & ZipName & "','" & ZipNo & "')"
	     Conn.Execute(sql)
    end if
	Case "UPD" :
	     sql = "Update Zip Set ZipName='" & ZipName & "',ZipNo='" & ZipNo & "' Where ZipID='" & ZipID & "'"
	     Conn.Execute(sql) 
	Case "DEL" :
	   sql = "Delete From Zip Where ZipID='" & ZipID & "'"
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
				<form name=ZipAdd method=post action=ZipAdd.asp>
					<input type=hidden name="tag" value="new">
					<input type=hidden name="ZipID" value="<%=request("ZipID")%>">
					<input type=hidden name="ZipName" value="<%=request("ZipName")%>">
					<input type=hidden name="ZipName" value="<%=request("ZipName")%>">
				</form>
			  <Script Language=JavaScript>window.ZipAdd.submit();</Script>	
<%
   elseif tag="UPD" then
%>
				<form name=ZipUpdate method=post action=ZipUpdate.asp>
					<input type=hidden name=tag value="upd">
					<input type=hidden name="ZipID" value="<%=request("ZipID")%>">
					<input type=hidden name="ZipName" value="<%=request("ZipName")%>">
					<input type=hidden name="ZipName" value="<%=request("ZipName")%>">
				</form>
				<Script Language=JavaScript>window.ZipUpdate.submit();</Script>
<%      
   end if
end if

IF( Session("Msg") = "") THEN
	   response.write "<script>window.opener.location.reload();window.close();</script>"
END If
%>
