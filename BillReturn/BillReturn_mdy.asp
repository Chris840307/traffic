<!-- #include file="../Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
tag = UCase(Request("tag"))
BillSN=Request("BIllSN")
BillNo=Request("BillNo")
CarNo=Request("CarNo")
MailDate=Request("MailDate")
MailReturnDate=Request("MailReturnDate")
ReturnResonID=Request("ReturnResonID")
if MailReturnDate <> "" then MailReturnDate=funGetDate(gOutDt(MailReturnDate),0)
StoreAndSendMailDate=Request("StoreAndSendMailDate")

if Request("firstisstoreandsend")="yes" then
    StoreAndSendReturnResonID = Request("ReturnResonID")
	StoreAndSendMailReturnDate = Request("MailReturnDate")
else
	StoreAndSendMailReturnDate=Request("StoreAndSendMailReturnDate")
	StoreAndSendReturnResonID=Request("StoreAndSendReturnResonID")
end if
if StoreAndSendMailReturnDate <> "" then StoreAndSendMailReturnDate=funGetDate(gOutDt(StoreAndSendMailReturnDate),0)


Select Case tag
	Case "NEW" :
		'不應進到這邊. 因為郵寄的時候已經寫入資料
	  'sql = "Insert into BillMailHistory (BillSN,BillNo,CarNo,MailReturnDate,ReturnResonID ) " & _
			'"values (" & BillSN & ",'" & BillNo & "','" & CarNo & ","& MailReturnDate & "," & ReturnResonID & ")"  
	  'Conn.Execute(sql)					  
		Session(msg)=" 該案件沒有郵寄資料. "
        	Response.write "<script>"
			Response.Write "alert('該案件沒有郵寄資料 無法退件！');"
			Response.write "self.close();"
			Response.write "</script>"
	Case "UPD":
 
	'	if ReturnResonID <> "" then
			sResonSQL=",ReturnResonID='" & ReturnResonID & "'"
	'	end if
        
	'	if StoreAndSendReturnResonID <> "" then
			sStoreAndSendResonSQL=",StoreAndSendReturnResonID='" & StoreAndSendReturnResonID & "'"
	'	end if

		if StoreAndSendMailReturnDate <> "" then
			sStoreAndSendMailReturnDateSQL=",StoreAndSendMailReturnDate=" & StoreAndSendMailReturnDate
		else
			sStoreAndSendMailReturnDateSQL=",StoreAndSendMailReturnDate=null"
		end if
	
		if MailReturnDate <> "" then
			MailReturnDateSQL=" MailReturnDate=" & MailReturnDate
		else
			MailReturnDateSQL=" MailReturnDate=null"
		end if
		
	  	sql = "Update BillMailHistory Set " & MailReturnDateSQL & sResonSQL & _
              sStoreAndSendMailReturnDateSQL &  sStoreAndSendResonSQL & _
              " ,ReturnRECORDMemberID=" & session("user_id") &_
            	" ,RETURNRECORDDATE=sysdate " & _
	    	  	" Where BillSN=" & BillSN

	  	Conn.Execute(sql)


        
		'退件註記 if request("wantDCIReturn") = "dcireturn" then
	  	sql2 = " Update BillBase Set BillStatus=3 Where SN= " & BillSN
	  	Conn.Execute(sql2)
		'end if 
 

End Select

IF( err.number<>0) THEN
	 Session("Msg") = Session("Msg") & "新增/修改過程有誤:" & err.description
end if	
%>
<!-- #include file="../Common/ClearObject.asp" -->
<%
if Session("msg") <> "" then
	if tag="NEW" then
		  response.contenttype = "text/html;charset=big5"			
%>
				<form name=BillReturnUpdate method=post action=BillReturnUpdate.asp>
					<input type=hidden name=tag value="NEW">
					<input type=hidden name=BillSn value="<%=request("BillSN")%>">
				</form>
			  <Script Language=JavaScript>window.BillReturnUpdate.submit();</Script>	
<%
   elseif tag="UPD" then
%>
				<form name=BillReturnUpdate method=post action=BillReturnUpdate.asp>
					<input type=hidden name=tag value="UPD">
					<input type=hidden name=BillSn value="<%=request("BillSN")%>">
				</form>
				<Script Language=JavaScript>window.BillReturnUpdate.submit();</Script>
<%      
   end if
end if


IF( Session("Msg") = "") THEN
	   response.write "<script>window.opener.location.reload();window.close();</script>"
END If

%>

