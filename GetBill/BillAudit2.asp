<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
BillStartNumber_q = Request("BillStartNumber_q")
BillEndNumber_q = Request("BillEndNumber_q")
BillStartNumber = Request("BillStartNumber")
BillEndNumber = Request("BillEndNumber")

if BillStartNumber_q <> "" then
	 BillStart = BillStartNumber_q
	 BillEnd = BillEndNumber_q
else
	 BillStart = BillStartNumber
	 BillEnd = BillEndNumber
end if

sql = "Select /*+ INDEX(GetBillDetail GETBILLDETAIL_PK) */ * from GetBillDetail Where billno between '" & BillStart & "' and '" & BillEnd & "' " & _
      "And BillNo Not in (Select nvl(BillNo,' ') From BillBase)"
Set RsLoss=Server.CreateObject("ADODB.RecordSet")
		RsLoss.cursorlocation = 3
		RsLoss.open SQL,Conn,3,1   
While Not RsLoss.Eof
     sqlUpd = "Update GetBIllDetail Set BillStateId=464 Where GetBillSn=" & RsLoss("GetBillSn") & _ 
           " And BillNo='" & RsLoss("BillNo") & "'"
     Conn.Execute(sqlUpd)
   RsLoss.MoveNext
Wend 
%>
<!-- #include file="../Common/ClearObject.asp" -->
<%
if err.number = 0 then
   Session("Msg") = BillStart & " ~ " & BillEnd & "漏號稽核完成"
else
   Session("Msg") = "錯誤訊息 : " & Err.description
end if	
Response.Write "<script>window.opener.location.reload();window.close();</script>"
%>