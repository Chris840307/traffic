<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
If Session("User_ID")="" Then
	Response.write "系統Session值已過期,請重新登入!"
  Response.End
End If	

Server.ScriptTimeOut=99999

tag = UCase(Request("tag"))
DispatchMemberID = Session("User_ID") 
GetBillDate = Request("GetBillDate")
GetBillMemberID = Request("GetBillMemberID")
BillStartNumber = Request("BillStartNumber")
BillEndNumber = Request("BillEndNumber")
CounterfoiReturn = Request("CounterfoiReturn")
Note = Request("Note")
GETBILLSN = Request("GETBILLSN")
BillCount = Request("BillCount")
isBillIn  = Request("BillIn")

If IsNumeric(BillStartNumber) Then
	BillStartNumber="NO"&BillStartNumber
	BillEndNumber="NO"&BillEndNumber
End if


sys_title=0

for i=len(trim(BillStartNumber)) to 1 step -1
	if not IsNumeric(mid(BillStartNumber,i,1)) then
		sys_title=i
		exit for
	end if
next

If Not ifnull(BillCount) Then
	Sno=MID(BillStartNumber,1,sys_title)
	Tno=MID(BillStartNumber,sys_title+1,len(BillStartNumber))
	Tno2=Right("00000000000" & CDbl(Tno)+(BillCount), len(BillStartNumber)-len(Sno))
	BillEndNumber=Sno&(Right("00000000000" & CDbl(Tno)+(BillCount), len(BillStartNumber)-len(Sno)))
end if

startHead = Mid(BillStartNumber,1,sys_title)
startTail = Mid(BillStartNumber,sys_title+1,len(BillStartNumber))

Select Case tag
	Case "NEW" :
	  sql = "select /*+ INDEX(getbilldetail GETBILLDETAIL_PK) */ BILLNO from Warninggetbilldetail where " & _
	        "length(billno)=length('"&trim(BillStartNumber)&"') and BillNo like '"&Sno&"%' and to_Number(SUBSTR(BillNo,"&len(Sno)+1&")) between "&Tno&" and "&Tno2
    Set Rs = Conn.Execute(sql)
    if not Rs.eof and isBillIn=0 then
    	 Session("Msg") = Rs("BILLNO") & "...等舉發單號已經存在,請重新輸入!!"
	     Rs.Close
	     Set Rs = Nothing
	  else    
	      sql = "select nvl(max(getbillsn),0)+1 as GETBILLSN from WARNINGGETBILLBASE"
	      set RsTemp = Server.CreateObject("ADODB.RecordSet")
	      Set RsTemp = Conn.Execute(sql)
	      GETBILLSN = RsTemp("GETBILLSN")
        RsTemp.Close
    	  	
	      
	      numStart = FormatNumber(startTail,0)
	      
	      TempNo = Int(numStart) - 1
	      sql = "Insert into WarningGetBillBase (GETBILLSN,DISPATCHMEMBERID,GETBILLDATE,GETBILLMEMBERID,BILLSTARTNUMBER,BILLENDNUMBER,COUNTERFOIRETURN,RecordDate,RecordMemberID,note,BillIn,RECORDSTATEID) " & _
	            "values (" & GETBILLSN & "," & Session("User_ID") & "," & funGetDate(gOutDT(GetBillDate),1) &"," & GetBillMemberID & ",'" & BILLSTARTNUMBER & "','" & BILLENDNUMBER & _
	            "'," & COUNTERFOIRETURN & ",sysdate," & Session("User_ID") & ",'" & Note & "'," & isBillIn & ",0)"  
						'to_date('" & GetBillDate & "','yyyy/mm/dd')
        Conn.BeginTrans
        Conn.Execute(sql)
		' 	入庫的紀錄不用新增GetBillDetail 避免後續領單使用新增會重覆
       if isBillIn = 0 then    
          For i=0 To BillCount
             'TempNo = TempNo + 1
             'FormatNum = Right("00000" & TempNo, 6)
             BillNo = Sno&(Right("00000000000" & CDbl(Tno)+(i), len(BillStartNumber)-len(Sno)))
             sqlTemp = "Insert into WarningGetBillDetail (GetBillSN,BillNo,BillStateID) Values (" & GETBILLSN & ",'" & BillNo & "'," & 463 & ")"
			
             Conn.Execute(sqlTemp)
             if err.number <> 0 then
             	  Exit For
             end if
          Next
		end if
        if err.number = 0 then
        	 Conn.CommitTrans
        else    	
           Conn.RollbackTrans
        end if
    end if
	Case "UPD":
	  sql = "Update WarningGetBillBase Set GetBillDate=" & funGetDate(gOutDT(GetBillDate),1)  & ",GetBillMemberID=" & GetBillMemberID & _
	        ",BillStartNumber='" & BillStartNumber & "',BillEndNumber='" & BillEndNumber & "'" & _
	        ",COUNTERFOIRETURN=" & COUNTERFOIRETURN & ",note='" & note & "' Where GETBILLSN=" & GETBILLSN
	  Conn.Execute(sql)
	Case "CHANGE":
		sql = "select BILLNO from Warninggetbilldetail  Where GETBILLSN=" & GETBILLSN &" and SUBSTR(BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"'"
		set Rs = Server.CreateObject("ADODB.RecordSet")
		Set Rs = Conn.Execute(sql)
		if not Rs.eof then
			rs.close
			sql = "select BILLSN.NEXTVAL as GETBILLSN from DUAL"
			set RsTemp = Server.CreateObject("ADODB.RecordSet")
			Set RsTemp = Conn.Execute(sql)
			new_GETBILLSN = RsTemp("GETBILLSN")
			RsTemp.Close

			numStart = FormatNumber(startTail,0)
			TempNo = Int(numStart) - 1

			sql = "Insert into WarningGetBillBase (GETBILLSN,DISPATCHMEMBERID,GETBILLDATE,GETBILLMEMBERID,BILLSTARTNUMBER,BILLENDNUMBER,COUNTERFOIRETURN,RecordDate,RecordMemberID,note,BillIn,RECORDSTATEID) " & _
			"values (" & new_GETBILLSN & "," & Session("User_ID") & "," & funGetDate(gOutDT(GetBillDate),1) &"," & GetBillMemberID & ",'" & BILLSTARTNUMBER & "','" & BILLENDNUMBER & _
			"'," & COUNTERFOIRETURN & ",sysdate," & Session("User_ID") & ",'" & Note & "',0,0)"
			'to_date('" & GetBillDate & "','yyyy/mm/dd')
			Conn.BeginTrans
			Conn.Execute(sql)

		' 	入庫的紀錄不用新增GetBillDetail 避免後續領單使用新增會重覆

			tempEndNum=Trim(Sno)&right("000000000"&trim(int(Tno)-1),len(Tno))
			if trim(BillStartNumber)=trim(tempEndNum) then
				strSQL = "delete from WarningGetBillBase Where GETBILLSN=" & GETBILLSN
				Conn.Execute(strSQL)
			else
				strSQL = "Update WarningGetBillBase Set BillEndNumber='" & tempEndNum & "' Where GETBILLSN=" & GETBILLSN
				Conn.Execute(strSQL)
			end if

			strSQL="delete from WarningGetBillDetail where GETBILLSN=" & GETBILLSN & " and SUBSTR(BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"'"
			Conn.Execute(strSQL)

			For i=int(Tno) To int(Tno2)
				BillNo = Trim(Sno)&right("000000000"&trim(int(i)),len(Tno))
				sqlTemp = "Insert into WarningGetBillDetail (GetBillSN,BillNo,BillStateID) Values (" & new_GETBILLSN & ",'" & BillNo & "'," & 463 & ")"

				Conn.Execute(sqlTemp)
				if err.number <> 0 then
					Exit For
				end if
			Next

			if err.number = 0 then
				 Conn.CommitTrans
			else    	
			   Conn.RollbackTrans
			end if
		else
			Session("Msg") = Rs("BILLNO") & "舉發單並無領單記錄!!"
			rs.close
			Set Rs = Nothing
		end if

  Case "DEL":
    sql1 = "Delete From WarningGetBillDetail Where GetBillSN=" & GetBillSN
    sql2 = "Delete From WarningGetBillBase Where GetBillSN=" & GetBillSN
    Conn.BeginTrans
       Conn.Execute(sql1)
       Conn.Execute(sql2)
    if err.number = 0 then
    	 Conn.CommitTrans
    else    	
       Conn.RollbackTrans
    end if    
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
				<form name=WarningGetBillAdd method=post action=WarningGetBillAdd.asp>
					<input type=hidden name=tag value="new">
					<input type=hidden name=UnitID value="<%=request("UnitID")%>">
					<input type=hidden name=GetBillDate value="<%=GetBillDate%>">
					<input type=hidden name=GetBillMemberID value="<%=GetBillMemberID%>">
					<input type=hidden name=BillStartNumber value="<%=BillStartNumber%>">
					<input type=hidden name=BillEndNumber value="<%=BillEndNumber%>">
					<input type=hidden name=CounterfoiReturn value="<%=CounterfoiReturn%>">
					<input type=hidden name=BillCount value="<%=Request("BillCount")%>">
					<input type=hidden name=Note value="<%=Note%>">
				</form>
			  <Script Language=JavaScript>window.WarningGetBillAdd.submit();</Script>	
<%
   elseif tag="UPD" then
%>
				<form name=WarningGetBillUpdate method=post action=WarningGetBillUpdate.asp>
					<input type=hidden name=tag value="upd">
					<input type=hidden name=GETBILLSN value="<%=request("GETBILLSN")%>">
					<input type=hidden name=CounterfoiReturn value="<%=CounterfoiReturn%>">
					<input type=hidden name=BillCount value="<%=Request("BillCount")%>">
					<input type=hidden name=Note value="<%=Note%>">
				</form>
				<Script Language=JavaScript>window.WarningGetBillUpdate.submit();</Script>
<%      
   elseif tag="CHANGE" then
%>
				<form name=WarningGetBillChange method=post action=WarningGetBillChange.asp>
					<input type=hidden name=tag value="upd">
					<input type=hidden name=GETBILLSN value="<%=request("GETBILLSN")%>">
					<input type=hidden name=CounterfoiReturn value="<%=CounterfoiReturn%>">
					<input type=hidden name=BillCount value="<%=Request("BillCount")%>">
					<input type=hidden name=Note value="<%=Note%>">
				</form>
				<Script Language=JavaScript>window.WarningGetBillChange.submit();</Script>
<%      
   end if
end if

IF( Session("Msg") = "") THEN
	if tag="NEW" then
		response.write "<script>alert('儲存完成!!');location='WarningGetBillAdd.asp';</script>"
	else
		response.write "<script>window.opener.location.reload();window.close();</script>"
	end if
END If
%>
