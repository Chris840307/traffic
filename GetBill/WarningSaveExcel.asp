<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
If not ifnull(request("Sys_Order")) Then
	strOrder=",gb."&request("Sys_Order")
End if

SQL="select " & _
    "gb.GETBILLSN,ut.UNITNAME,md.Loginid,md.ChName,gb.GetBillDate,gb.BillStartNumber,gb.BILLENDNUMBER, " & _
    "decode(gb.CounterfoiReturn,0,'使用中',1,'使用完畢') as CounterfoiReturnDesc,gb.BILLIN , gb.note, gb.RecordMemberID,gb.RecordDate " & _
    "from warninggetbillbase gb,memberdata md,unitinfo ut " & _
    "where gb.GetBillMemberID=md.MemberID and gb.RecordStateID <> -1 " & _
    "and md.UnitID=ut.UnitID "&Request("Sys_strWhere")&  " order by md.UnitID"&strOrder

set RsTemp=Server.CreateObject("ADODB.RecordSet")
RsTemp.open sql,Conn,3,3
%>
<html>   
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>ExportBase</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style3 {font-size: 15px; mso-number-format:"\@";}
.style5 {
	font-size: 11px;
	color: #666666;
}
-->
</style>
</head>	 
<body>    
 <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#808080" >    
   <tr>
     <td bgcolor="#EBFBE3"><B><center>領單單位</center></B></td>
	 <td bgcolor="#EBFBE3"><B><center>領單代碼</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>領單人員</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>領單日期</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>舉發單起始碼~單截止碼</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>數量</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>存根聯</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>備註</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>漏號</center></B></td>
   </tr>                     
<%
While Not RsTemp.Eof
	   billstartnumber = RsTemp("billstartnumber")
	   billendnumber = RsTemp("billendnumber")

	for i=len(billstartnumber) to 1 step -1
		if not IsNumeric(mid(billstartnumber,i,1)) then
			startTail=MID(billstartnumber,1,i)
			numStart=MID(billstartnumber,i+1,len(billstartnumber))
			exit for
		end if
	next

	for i=len(billendnumber) to 1 step -1
		if not IsNumeric(mid(billstartnumber,i,1)) then
			endTail=MID(billstartnumber,1,i)
			numEnd=MID(billstartnumber,i+1,len(billstartnumber))
			exit for
		end if
	next
'     startTail = Mid(billstartnumber,4,6)
'     endTail = Mid(billendnumber,4,6)
'     numStart = FormatNumber(startTail,0)
'     intStart = Int(numStart)
'     numEnd = FormatNumber(endTail,0)
'     intEnd = Int(numEnd)	
	   billAmount = (intEnd - intStart)+1
	   getbilldateTemp = gInitDT(RsTemp("getbilldate")) 'Right("00" & Year(RsTemp("getbilldate"))-1911, 3) & "-" & Right("0" & Month(RsTemp("getbilldate")), 2) & "-" & Right("0" & Day(RsTemp("getbilldate")), 2)
%>                             
   <tr>
     <td ><%=RsTemp("unitname")%></td>
	 <td class="style3"><%=RsTemp("LoginID")%></td>
     <td ><%=RsTemp("chname")%></td>
     <td ><%=getbilldateTemp%></td>
     <td ><%=billstartnumber%>~<%=billendnumber%></td>
     <td align="right"><%=billAmount%></td>
     <td ><%=RsTemp("CounterfoiReturnDesc")%></td>
     <td ><%
		if RsTemp("BILLIN")=1 then
			response.write "入庫 .  "&RsTemp("note")
		elseif RsTemp("BILLIN")=2 then
			response.write "出庫 .  "&RsTemp("note")
		else
			response.write "領取 .  "&RsTemp("note")
		end if
	 %></td>
     <td >&nbsp;</td>                              
   </tr>   
<%
   RsTemp.MoveNext
Wend
%>     
 </table>    
 </body>
<!-- #include file="../Common/ClearObject.asp" --> 
 <%
	fMnoth=month(now)
	if fMnoth<10 then fMnoth="0"&fMnoth
	fDay=day(now)
	if fDay<10 then	fDay="0"&fDay
	fname=year(now)&fMnoth&fDay&"_領件紀錄列表.xls"
	Response.AddHeader "Content-Disposition", "filename="&fname
	response.contenttype="application/x-msexcel; charset=MS950"
 %>
 </html>