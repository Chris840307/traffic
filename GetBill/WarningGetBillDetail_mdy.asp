<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
If Session("User_ID")="" Then
	Response.write "系統Session值已過期,請重新登入!"
  Response.End
End If
if trim(Request("getBillSn"))<>"" then
	getBillSn = Request("getBillSn")
	billNo = Request("billNo")
	noteContent = Request("noteContent")
	BillStateId = Int(Request("BillStateId"))
	todayTemp = Right("0"&year(Date()),4) &"/" & Right("0"&month(Date()),2) &"/" & Right("0"&day(Date()),2)

	sql = "Update WarningGetBillDetail Set noteContent='" & noteContent & "' ,BillStateId=" & BillStateId & _
		  ",RecordDate=to_date('" & todayTemp & "','YYYY/MM/DD'),RecordMemberId=" & Session("User_ID") & _
		  " Where billNo='" & billNo & "' And getBillSn=" & getBillSn
	Conn.Execute(sql)
else
	billstartnumber=Request("billstartnumber")
	billendnumber=Request("billendnumber")
	noteContent = Request("noteContent")
	BillStateId = Int(Request("BillStateId"))
	todayTemp = Right("0"&year(Date()),4) &"/" & Right("0"&month(Date()),2) &"/" & Right("0"&day(Date()),2)

'	for i=1 to len(billstartnumber)
'		if IsNumeric(mid(billstartnumber,i,1)) then
'			Sno1=MID(billstartnumber,1,i-1)
'			Tno1=MID(billstartnumber,i,len(billstartnumber))
'			exit for
'		end if
'	next
'
'	for i=1 to len(billendnumber)
'		if IsNumeric(mid(billendnumber,i,1)) then
'			Sno2=MID(billendnumber,1,i-1)
'			Tno2=MID(billendnumber,i,len(billendnumber))
'			exit for
'		end if
'	next

	sql = "Update WarningGetBillDetail Set noteContent='" & noteContent & "' ,BillStateId=" & BillStateId & _
		  ",RecordDate=to_date('" & todayTemp & "','YYYY/MM/DD'),RecordMemberId=" & Session("User_ID") & _
		  " Where BillNo between '"&billstartnumber&"' and '"&billendnumber&"'"
	Conn.Execute(sql)
end if

IF( err.number<>0) THEN
	 Session("Msg") = Session("Msg") & "<br>修改失敗,錯誤訊息:" & err.description
end if	
%>
<!-- #include file="../Common/ClearObject.asp" -->
<%
IF( Session("Msg") = "") THEN
	   response.write "<script>window.opener.location.reload();window.close();</script>"
END If
%>
