<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Query/sqlDCIExchangeData.asp"-->
<%
  typeid=""
	If Trim(request("TypeQry"))="PicQry" Then typeid="1"
	If Trim(request("TypeQry"))="BookQry" Then typeid="0"
	If Trim(request("TypeQry"))="BookReturnQry" Then typeid="2"
	If Trim(request("TypeQry"))="BookBillnoReturnQry" Then typeid="3"
	If Trim(request("TypeQry"))="BookOtherQry" Then typeid="4"


    strUpdate="select BillNo from BillAttatchImage where recordstateid='0' and typeid='"&typeid&"' and BillNo='"&trim(request("BillNo"))&"'"
	Set rs=conn.execute(strUpdate)
	If Not rs.eof Then 
		If typeid<>"4" and typeid<>"0" Then 
			response.write "alert('單號已輸入過');"
		Else
			response.write "Qry();"
		End if
	Else
	   strUpdate="Update BillAttatchImage set BillNo='"&trim(request("BillNo"))&"' where SN=" & request("SN")
	   Conn.execute strUpdate	


		If request("AcceptMark")="1" Then 
		'更新受收註記狀態
			strSQL="Update BillMailHistory set SignResonID='A',SignDate=sysdate,SignRecordMemberID="&Session("User_ID")&",ReturnReCordDate=sysdate,UserMarkMemberID="&Session("User_ID")&",UserMarkDate=sysdate,			UserMarkResonID='A',UserMarkReturnDate=sysdate,	mailStation='',signman='' where BillNo='"&trim(request("BillNo"))&"'"
			conn.execute(strSQL)
			
			strSQL="update billbase set billstatus=7 where billno='"&trim(request("BillNo"))&"' and recordstateid<>-1 "
			conn.execute(strSQL)			

		'上傳監理站
'			strReturn="select a.SN,a.IllegalDate,a.BillTypeID" &_
'		",a.BillNo,a.CarNo" &_
'		",a.BillUnitID,a.BillStatus,a.RecordDate" &_
'		",a.RecordMemberID,c.UserMarkResonID,c.StoreAndSendReturnResonID from BillBase a" &_
'		",MemberData b,BillMailHistory c where a.RecordStateID<>-1" &_
'		" and a.RecordMemberID=b.MemberID(+) and c.BillSN=a.SN and a.BillNo='"&trim(request("BillNo"))&"' order by c.UserMarkDate"
'		set rsReturn=conn.execute(strReturn)
'			If Not rsReturn.eof Then 
'				funcBillGet conn,trim(rsReturn("SN")),trim(rsReturn("BillNo")),trim(rsReturn("BillTypeID")),trim(rsReturn("CarNo")),trim(rsReturn("BillUnitID")),trim(rsReturn("RecordDate")),trim(rsReturn("RecordMemberID")),Right("0"&Year(date)-1911,3)&Right("0"&Month(date),2)&Right("0"&day(date),2)
'			End if
		End if
       response.write "Qry();"
    End if	
    

%>

