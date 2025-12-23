
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
hasDate = True
UserId = Session("User_ID")
ReportId = "report0010"
rptHead1 = Trim(Request("rptHead1"))
rptHead2 = Trim(Request("rptHead2"))
startDate_q = Trim(Request("startDate_q"))
endDate_q = Trim(Request("endDate_q"))
IllegalDate_start=Trim(Request("IllegalDate_start"))
IllegalDate_end=Trim(Request("IllegalDate_end"))
ListOrder=Trim(Request("ListOrder"))
sumDate_q=Trim(Request("sumDate_q"))
unit = Request("unit")
UnitID_q = Request("UnitID_q")
ReportName=request("rptHead2")
startDate_Name=request("startDate_Name")
endDate_Name=request("endDate_Name")
Server.ScriptTimeout=6000
'smiht  新增逕舉或攔停 or 所有
BillBaseType=Trim(Request("BillBaseType"))
overtype=Trim(Request("overtype"))

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

thenPasserUnit=""
strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsunit=conn.execute(strSQL)
If Not rsunit.eof Then
	Sys_UnitID=trim(rsunit("UnitID"))
	Sys_UnitLevelID=trim(rsunit("UnitLevelID"))
	Sys_UnitTypeID=trim(rsunit("UnitTypeID"))
End if
rsunit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
elseif sys_City<>"金門縣" and sys_City<>"連江縣" then
	strSQL="select * from UnitInfo where UnitName like '%分局' and (UnitTypeID='"&Sys_UnitID&"' or UnitTYpeID='"&Sys_UnitTypeID&"' or UnitID='"&Sys_UnitID&"'or UnitID='"&Sys_UnitTypeID&"')"
else
	strSQL="select * from UnitInfo where UnitName like '%所' and (UnitTypeID='"&Sys_UnitID&"' or UnitTYpeID='"&Sys_UnitTypeID&"' or UnitID='"&Sys_UnitID&"'or UnitID='"&Sys_UnitTypeID&"')"
end if
set rsunit=conn.Execute(strSQL)
Sys_UnitID=trim(rsunit("UnitID"))
if Not rsunit.eof then thenPasserUnit=trim(rsunit("UnitName"))
rsunit.close

strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
rsUInfo.close

strRul="select Value from Apconfigure where ID=3"
set rsRul=conn.execute(strRul)
RuleVer=trim(rsRul("Value"))
rsRul.Close
'--------------------------------------------------------smith 把相關的選項塞入 userrptinfo . 自訂報表. --------------------------------------------------------
Conn.BeginTrans
	sqlDel = "Delete From UserRptInfo Where UserId=" & UserId & " And ReportId='" & ReportId & "' "   
	Conn.Execute(sqlDel)
	sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType) Values (" & UserId & ",'" & ReportId & "','rptHead1','" & rptHead1 & "','TEXT')"
	Conn.Execute(sqlAdd)
	sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType) Values (" & UserId & ",'" & ReportId & "','rptHead2','" & rptHead2 & "','TEXT')"
	Conn.Execute(sqlAdd)
	sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType) Values (" & UserId & ",'" & ReportId & "','IllegalDate_start','" & IllegalDate_start & "','TEXT')"
	Conn.Execute(sqlAdd)
	sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType) Values (" & UserId & ",'" & ReportId & "','IllegalDate_end','" & IllegalDate_end & "','TEXT')"
	Conn.Execute(sqlAdd)
	sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType) Values (" & UserId & ",'" & ReportId & "','sumDate_q','" & sumDate_q & "','TEXT')"
	Conn.Execute(sqlAdd)
	sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType) Values (" & UserId & ",'" & ReportId & "','ListOrder','" & ListOrder & "','SELECT')"
	Conn.Execute(sqlAdd)
	'smith 新增 案件類型 以及 日期差距	
     sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType,ReportName) Values (" & UserId & ",'" & ReportId & "','BillBaseType','" & BillBaseType & "','RADIO','"&ReportName&"')"
     Conn.Execute(sqlAdd)
	 sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType,ReportName) Values (" & UserId & ",'" & ReportId & "','overtype','" & overtype & "','RADIO','"&ReportName&"')"
     Conn.Execute(sqlAdd)	

	If startDate_q <> "" And endDate_q <> "" Then
		hasDate = True
		sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType) Values (" & UserId & ",'" & ReportId & "','startDate_q','" & startDate_q & "','SELECT')"
		Conn.Execute(sqlAdd)   
		sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType) Values (" & UserId & ",'" & ReportId & "','endDate_q','" & endDate_q & "','SELECT')"
		Conn.Execute(sqlAdd)         	
	End If  

	If unit="y" Then
		sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType) Values (" & UserId & ",'" & ReportId & "','unit','" & unit & "','CHECKBOX')"
		Conn.Execute(sqlAdd)    	
		sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType) Values (" & UserId & ",'" & ReportId & "','UnitID_q','" & UnitID_q & "','SELECT')"
		Conn.Execute(sqlAdd)     	
	End If

if err.number = 0 then
	Conn.CommitTrans
else    	
	Conn.RollbackTrans
end if  
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	Set RsTemp = Server.CreateObject("ADODB.RecordSet")
	'smith 檢查是否為 行人慢車案件的檢核. 是的話有些條件無法統計 如: 郵寄日期. 簽收日期等等.  只能做建檔. 違規. 應到案日等的稽核
	'str_DayID="RecordDate,IllegalDate,DeallIneDate,BillFillDate,ExChangeDate,MailDate,MailAcceptDate,MailReturnDate,StoreAndSendMailReturnDate,OpenGovMailReturnDate"
	'str_DayName="建檔日期,違規日期,應到案日期,填單日期,入案日期,郵寄日期,收受日期,單退日期,寄存日期,公示日期"
	if BillBaseType=9 then 
		if (startDate_q="ExChangeDate" or  startDate_q="MailDate" or startDate_q ="MailAcceptDate" or startDate_q="MailReturnDate" or startDate_q="StoreAndSendMailReturnDate" or startDate_q="OpenGovMailReturnDate" ) or (endDate_q="ExChangeDate" or  endDate_q="MailDate" or endDate_q ="MailAcceptDate" or endDate_q="MailReturnDate" or endDate_q="StoreAndSendMailReturnDate" or endDate_q="OpenGovMailReturnDate" ) then 
			response.write startDate_q & " " & endDate_q
			response.write "<font size='5'> <br><br>行人慢車道路障礙 案件，只能稽核 下列條件 :<br><br> 建檔日期,違規日期,應到案日期,填單日期 <br><br> 請重新設定</font>"
			response.end
		end if
	end if
	
	'稽核日期的sql 設定
	tmpSql = ""
	If hasDate Then
		'smith 加上  大於 或是 小於  overtype = 1  or overtype=2     1大於  2小於
		if overtype="1" then 
			tmpSql = tmpSql & " where (floor("&startDate_q&" - "&endDate_q&") > "&sumDate_q&" or floor("&startDate_q&" - "&endDate_q&") < -"&sumDate_q&")"
		end if
		if overtype="2" then 
			tmpSql = tmpSql & " where (floor("&startDate_q&" - "&endDate_q&") < "&sumDate_q&"  and floor("&startDate_q&" - "&endDate_q&") > -"&sumDate_q&")"
									 '(floor("&startDate_q&" - "&endDate_q&") < "&sumDate_q&" or
		end if
		
		If trim(startDate_q)="MailDate" and trim(endDate_q)="MailDate" and BillBaseType<>9 Then tmpSql = tmpSql & " and BillTypeID=2"
	End If
	'smith  是否選擇 單位
	If unit="y" Then
		'行人慢車找billunitid
		if BillBaseType=9 then 
			tmpSql = tmpSql & " And BillUnitID='" & UnitID_q & "' "
		'1-69條找view中的unitid
		else
			tmpSql = tmpSql & " And UnitID='" & UnitID_q & "' "
		end if
	End If
	'smith  設定統計區間
	If IllegalDate_start <> "" And IllegalDate_end <> "" Then
		tmpSql = tmpSql & " And IllegalDate Between To_Date('" & gOutDT(IllegalDate_start)&" 0:0:0" & "','YYYY/MM/DD/HH24/MI/SS') And To_Date('" & gOutDT(IllegalDate_end)&" 23:59:59" & "','YYYY/MM/DD/HH24/MI/SS')"
	end if
	'smiht 有郵寄日期稽核的話 就不用再加入billtype , 因為上方有加入限制
	If trim(startDate_q)<>"MailDate" and trim(endDate_q)<>"MailDate"  Then 
		'smith稽核 所有案件 or 行人攤販按鍵就不用加上 billtypeid    '1欄  2逕  9行人慢車
		if BillBaseType<>9 then 
			If BillBaseType<>0 then tmpSql = tmpSql & " And BillTypeId="&BillBaseType	
		end if
	end if
	
	'smith 加入 建檔人員的查詢條件   BILLBASEAUDITVIEW 裡面沒有 
	if trim(request("RecordMemID"))<>"" then
		tmpSql = tmpSql & " and RecordMemberID in (select memberid from memberdata where Loginid='"&trim(request("RecordMemID"))&"')"
	end if

	if BillBaseType<>9 then
	   if ListOrder="Billmem1" then ListOrder="ChName" 
		tmpSql =tmpSql & " order by "&ListOrder
		'smith 20091005 南投加入 相距天數
		if sys_City="南投縣" or sys_City="雲林縣" then	'南投縣要加入建檔人 but  BILLBASEAUDITVIEW 裡面沒有 
			strSQL="select billno,illegaldate,deallinedate,billfilldate,billtypeid,exchangedate,Recordmemberid,maildate,mailacceptdate,mailreturndate,storeandsendmailreturndate,opengovmailreturndate,memberid,chname,unitid,unitname, abs(to_date(" & endDate_q & ")  - to_date(" & startDate_q & ")) as daysbetween from (" &_
			"select a.BillNo,a.IllegalDate,a.DeallIneDate,a.BillFillDate,a.BillTypeID,a.Recordmemberid" &_
			",b.ExChangeDate,d.MailDate,c.MailAcceptDate,d.MailReturnDate,d.StoreAndSendMailReturnDate" &_
			",d.OpenGovMailReturnDate,e.MemberID,e.ChName,f.UnitID,f.UnitName from " &_
			"(select * from BillBase where RecordStateID=0) a," &_
			"(select BillSN,BillNo,CarNo,ExChangeDate from DCILog where ExChangeTypeID='W' " &_
			" and DciReturnStatusID='Y') b,(Select b.BillSN,b.BillNo,b.CarNo,b.MailReturnDate as MailAcceptDate " &_
			" from BillBase a,BillMailHistory b where a.SN=b.BillSN and a.BillNo=b.BillNo " &_
			" and a.CarNo=b.CarNo and a.BillStatus='7') c,(Select b.* from BillBase a,BillMailHistory b " &_
			" where a.SN=b.BillSN and a.BillNo=b.BillNo and a.CarNo=b.CarNo and a.BillStatus<>'7') d, " &_
			"MemberData e,UnitInfo f where a.BillNo=b.BillNo(+) and a.CarNo=b.CarNo(+) " &_
			" and a.SN=c.BillSN(+) and a.BillNo=c.BillNo(+) and a.CarNo=c.CarNo(+) and a.SN=d.BillSN(+) " &_
			" and a.BillNo=d.BillNo(+) and a.CarNo=d.CarNo(+) and a.BillMemID1=e.MemberID(+) " &_
			" and a.BillUnitID=f.UnitID(+) " &_
			") " & tmpSql
		else
			strSQL="select billno,illegaldate,deallinedate,billfilldate,billtypeid,exchangedate,maildate,mailacceptdate,mailreturndate,storeandsendmailreturndate,opengovmailreturndate,memberid,chname,unitid,unitname, abs(to_date(" & endDate_q & ")  - to_date(" & startDate_q & ")) as daysbetween from BILLBASEAUDITVIEW " & tmpSql
		end if
	else 
		
		tmpSql =tmpSql & " order by "&ListOrder
		strSQL="select * from PasserBase " & tmpSql
	end if
	
'response.write strSQL
'response.end

'response.write startDate_q & "," 
'response.write endDate_q
'response.end

	set rsdata=conn.execute(strSQL)
   
set RSSystem=Server.CreateObject("ADODB.RecordSet")
sql = "select UnitName from UnitInfo where UnitID= '" & Session("Unit_ID") & "'"
Set RSSystem = Conn.Execute(sql)
if Not RSSystem.Eof Then
	printUnit = RSSystem("UnitName")
End If	

selectUnit = ""
If unit = "y" Then
   sql = "Select UnitName , UnitID from UnitInfo Where UnitID= '" & UnitID_q & "'"
   Set RSSystem = Conn.Execute(sql)
   if Not RSSystem.Eof Then
   	  selectUnit = RSSystem("UnitName")
   End If
End If 
%>
<html>   
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>ExportBase</title>
<style type="text/css">
<!--
body {font-family:標楷體;font-size:12pt}
.style1 {font-family:標楷體;font-size:14pt}
-->
</style>
</head>	 
<body>    
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" >
		<tr>
			<td colspan=2>
				  列印時間: <%=gInitDT(date)%> <br>
			    列印單位: <%=printUnit%> <br>
			    列印人員: <%=Session("Ch_Name")%>
			</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>

		</tr>	  
	</table>
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" align="center" >
			<tr>				 
				<td colspan=12><span class="style1"><b><center><%=rptHead1%></center></b></span></td>
			</tr>			
			<tr>
			   <td colspan=12><span class="style1"><b><center><%=rptHead2%></center></b></span></td>
			</tr>			
			<tr>
			   <td colspan=12><center>統計方式: <%=startDate_Name%> 與 <%=endDate_Name%> 
			   相差 
			   <%  
			   'smith 新增大於 或是 小於
				if overtype=1 then 
					response.write " 大於 "
				else
					response.write " 小於 "	
				end if

			   %>  <%=sumDate_q%>  天</center></td>
			</tr>						
	</table>

	<table border="1" width="100%" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#808080" >	
<%
response.write "<tr>"
response.write "<td>單號</td>"
response.write "<td>類型</td>"
response.write "<td>單位</td>"
response.write "<td>舉發人員</td>"
response.write "<td>舉發日</td>"
response.write "<td>填單日</td>"
response.write "<td>入案日</td>"
response.write "<td>應到案日</td>"
response.write "<td>郵寄日</td>"
response.write "<td>收受日</td>"
response.write "<td>單退日</td>"
response.write "<td>寄存送達日</td>"
response.write "<td>公示送達日</td>"
response.write "<td>相差天數<br><font size='2'>翌日起算</font></td>"

response.write "</tr>"
while Not rsdata.eof
	BillTypeName="攔停"
	if trim(rsdata("BillTypeID"))="2" then BillTypeName="逕舉"
	response.write "<td>"&rsdata("BillNo")&"</td>"
	response.write "<td>"&BillTypeName&"</td>"
	response.write "<td>"&rsdata("UnitName")&"</td>"	
	if BillBaseType<>9 then
		response.write "<td>"&rsdata("ChName")&"</td>"
	else 
		response.write "<td>"&rsdata("BillMem1")&"</td>"
	end if
	response.write "<td>"&gInitDT(rsdata("IllegalDate"))&"</td>"
	response.write "<td>"&gInitDT(rsdata("BillFillDate"))&"</td>"
	if BillBaseType<>9 then
		response.write "<td>"&gInitDT(rsdata("ExChangeDate"))&"</td>"
	else 
		response.write "<td></td>"
	end if
	response.write "<td>"&gInitDT(rsdata("DeallIneDate"))&"</td>"
	if BillBaseType<>9 then	
		response.write "<td>"&gInitDT(rsdata("MailDate"))&"</td>"
		response.write "<td>"&gInitDT(rsdata("MailAcceptDate"))&"</td>"
		response.write "<td>"&gInitDT(rsdata("MailReturnDate"))&"</td>"
		response.write "<td>"&gInitDT(rsdata("StoreAndSendMailReturnDate"))&"</td>"
		response.write "<td>"&gInitDT(rsdata("OpenGovMailReturnDate"))&"</td>"
		'startDate_q , endDate_q   days between
		response.write "<td>"&(rsdata("daysbetween"))&"</td>"
	else 
		response.write "<td></td>"
		response.write "<td></td>"
		response.write "<td></td>"
		response.write "<td></td>"
		response.write "<td></td>"
		response.write "<td></td>"
	end if	
	response.write "</tr>"
	rsdata.movenext
wend
rsdata.close
response.write "</table> "

fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay

fname=year(now)&fMnoth&fDay&"_"&ReportName&".xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%> 
 </body>
<!-- #include file="../Common/ClearObject.asp" --> 
</html>