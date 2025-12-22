<!-- #include file="../Common/dbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
SQL = Session("ExcelSql")
SQL = SQL & " And gb.counterfoireturn=1 and gb.BillIn=0 "
SQL= SQL & " order by gb.BillStartNumber "
'response.write SQL
'response.end

Set rs=Server.CreateObject("ADODB.RecordSet")
rs.cursorlocation = 3
rs.open SQL,Conn,3,1
' 有可能查詢到多筆已經使用完畢的主領單紀錄
i = 1
While Not rs.Eof
	 'getBillBase 一筆
   '只有一筆使用完畢的領單紀錄.
   if rs.RecordCount=1 then
   	  BillStartNumber_q = rs("BillStartNumber")
   	  BillEndNumber_q = rs("BillEndNumber")
   	  GETBILLSN = rs("GETBILLSN")   	  
   else
   		'以第一筆的startnumber 作為起始號碼
      if i=1 then
         BillStartNumber_q = rs("BillStartNumber")
      elseif i=rs.RecordCount then
      	  BillEndNumber_q = rs("BillEndNumber")
      end if
   	  
   end if
   '-------------------------------------------------------------------------------------------------
   '把billno 再 Billbase 理的紀錄奇billstateid 為漏號的先更正為正常.
   '避免之前漏號的後來補登狀態還是漏號
   sql = "Select * From GetBillDetail Where getbillsn=" & rs("GetBillSn") & " And (billno between " & _
         " '" & rs("BillStartNumber") & "' And '" & rs("BillEndNumber") & "') " & _
         " And BillNo in (Select nvl(BillNo,' ') From BillBase)"

   Set RsLoss=Server.CreateObject("ADODB.RecordSet")
		RsLoss.cursorlocation = 3
		RsLoss.open sql,Conn,3,1    
   While Not RsLoss.Eof
        sqlUpd = "Update GetBIllDetail Set BillStateId=463 Where GetBillSn=" & rs("GetBillSn") & _ 
        			" and BillStateID =464 " & _ 
              " And BillNo='" & RsLoss("BillNo") & "'"

        Conn.Execute(sqlUpd)
       
      RsLoss.MoveNext
   Wend  
	'------------------------------------------------------------------------------------------------------
   
   sql = "Select * From GetBillDetail Where getbillsn=" & rs("GetBillSn") & " and BillStateID =463 " & _ 
   				" And (billno between " & _
         " '" & rs("BillStartNumber") & "' And '" & rs("BillEndNumber") & "') " & _
         " And BillNo Not in (Select nvl(BillNo,' ') From BillBase)"

   Set RsLoss=Server.CreateObject("ADODB.RecordSet")
	 RsLoss.cursorlocation = 3
	 RsLoss.open sql,Conn,3,1    
   While Not RsLoss.Eof
        sqlUpd = "Update GetBIllDetail Set BillStateId=464 Where GetBillSn=" & rs("GetBillSn") & _ 
              " And BillNo='" & RsLoss("BillNo") & "'"

        Conn.Execute(sqlUpd)
       
      RsLoss.MoveNext
   Wend  
   i = i + 1
   rs.MoveNext
Wend 

%>

<%
if err.number = 0 then
 
   if Trim(BillStartNumber_q)<>"" then
   	  rs.movefirst
   	  detailPara = "GetBillDetail.asp?GETBILLSN=" & rs("GETBILLSN") & "&getbilldate=" & getbilldateTemp & "&chname=" & rs("ChName") & "&billstartnumber=" & billstartnumber & "&billendnumber=" & billendnumber & "&qryType=1"
      Response.Redirect "GetBillDetail.asp?BillStartNumber_q=" & trim(BillStartNumber_q) & " &BillEndNumber_q=" & trim(BillEndNumber_q) & "&qryType=2"
   else
   	  Session("Msg") = "查無相關漏號資料..." 
   	  Response.Write "<script>window.opener.location.reload();window.close();</script>"
   end if
else
   	  Session("Msg") = "錯誤訊息 : " & Err.description
   	  Response.Write "<script>window.opener.location.reload();window.close();</script>"	
end if	
%>
<!-- #include file="../Common/ClearObject.asp" -->