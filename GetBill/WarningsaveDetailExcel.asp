<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
sql = Session("DetailSQL") 
set RsTemp=Server.CreateObject("ADODB.RecordSet")
RsTemp.cursorlocation = 3
RsTemp.open sql,Conn,3,1
'response.write sql & "<br>"

fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_領取警告單列表.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950"
%>
<html>   
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>ExportDetail</title>
</head>	 
<body>    
 <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#808080" >    
   <tr>
     <td bgcolor="#EBFBE3"><B><center>舉發單號</center></B></td>
      <th bgcolor="#EBFBE3" width="12%" nowrap><span class="font12">領取單位</span></th>
     <th bgcolor="#EBFBE3" width="10%" nowrap><span class="style3">領單人員</span></th>     
     <td bgcolor="#EBFBE3"><B><center>特殊狀態設定</center></B></td>
     <!--<td bgcolor="#EBFBE3"><B><center>特殊說明紀錄人員</center></B></td>-->
     <td bgcolor="#EBFBE3"><B><center>紀錄時間</center></B></td>
     <td bgcolor="#EBFBE3" width="20%"><B><center>特殊說明</center></B></td>
   </tr>                     
<%
  sqlTemp = "Select id,content From Code Where TypeId=17"  
  Set RsTemp2=Server.CreateObject("ADODB.RecordSet")
  RsTemp2.open sqlTemp,Conn,3,3 
  Set dicObj = Server.CreateObject("Scripting.Dictionary")
  dicObj.RemoveAll
  While Not RsTemp2.Eof
     idStr = Cstr(RsTemp2("id"))
     contentStr = RsTemp2("content")
     dicObj.Add idStr,contentStr
     RsTemp2.MoveNext
  Wend
  if RsTemp2.state then  RsTemp2.close
While Not RsTemp.Eof
	   RecordDateTemp = ""
	   ChName = ""
	   BillState = ""
	   if (RsTemp("RecordDate") & "") <> "" Then
	      RecordDateTemp = gInitDT(RsTemp("RecordDate"))
	   end if
	   'RecordMemberID = RsTemp("RecordMemberID") & ""
	   'if RecordMemberID <> "" Then
	   '   sqlTemp2 = "Select chname From MemberData Where MemberID=" & Int(RecordMemberID)	   
     '   RsTemp2.open sqlTemp2,Conn,3,3
     '   if Not RsTemp2.Eof Then
     '      ChName = RsTemp2("chname")
     '   end if
	   'End If
	   
	   BillStateId = Cstr(RsTemp("BillStateId")) & ""
	   if BillStateId <> "" then   
        BillState = dicObj.Item(BillStateId)
     end if
%>                             
   <tr>
     <td ><%=RsTemp("BillNo")%></td>
             <td><%=RsTemp("UnitName")%></td>
     <td ><%=RsTemp("GetBillChName")%></td>     
     <td ><%=BillState%></td>
     <!-- <td ><%=ChName%></td> -->
     <td ><%=RecordDateTemp%></td>
     <td ><%=RsTemp("NoteContent")%></td>
                            
   </tr>   
<%
   RsTemp.MoveNext
Wend
Set dicObj = Nothing
%>     
 </table>    
 </body>
<!-- #include file="../Common/ClearObject.asp" --> 
 </html>