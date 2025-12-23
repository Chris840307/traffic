<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DCIURL.ini"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>郵寄狀態查詢紀錄</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 16px; color:#ff0000; }
.btn3{
   font-size:12px;
   font-family:新細明體;
   background-color:#EEEEEE;
   border-style:solid;
}
-->
</style>
</head>
<%

	strsql="select max(ReCordDate) ReCordDate from mailresultprocesslog"
	Set rs=conn.execute(strsql)
	LastUpdateTime=rs("ReCordDate")

Function GetCDate(IDate)
	If IDate="" Or IsNull(IDate) Then 
		GetCDate=""
	Else
		GetCDate=Year(IDate)-1911&"-"&Right("0"&Month(IDate),2)&"-"&right("0"&day(IDate),2)
	End if

End Function

Function GetCDateTime(IDate)
	If IDate="" Or IsNull(IDate) Then 
		GetCDateTime=""
	Else
		GetCDateTime=Year(IDate)-1911&"-"&Right("0"&Month(IDate),2)&"-"&right("0"&day(IDate),2)&" "&right("0"&Hour(IDate),2)&":"&right("0"&minute(IDate),2)&":"&right("0"&second(IDate),2)
	End if		
End Function

Function GetCDate2(IDate)
	If IDate="" Or IsNull(IDate) Then 
		GetCDate2=""
	Else
		GetCDate2=CDbl(Left(IDate,4))-1911&Right(IDate,Len(IDate)-4)
	End if
End function

Function GetResonName(ID)
	tmp=""
	strsql="select Content from dcicode where typeid=7 and ID='"&ID&"'"
	Set rstmp=conn.execute(strsql)
	If Not rstmp.eof Then tmp=rstmp("Content")
	Set rstmp=Nothing
	GetResonName=tmp
End function
if request("DB_Selt")="Selt" then
	strwhere=""

	tmpBillno=""
	Billno=""
	if trim(request("Sys_BillNo"))<>"" Then
		tmpBillno=Split(request("Sys_BillNo"),",")
		For i=0 To UBound(tmpBillno)
			Billno=Billno&tmpBillNo(i)&"','"
		next
		
			strwhere = strwhere & " and BillNo in ('"&Billno&"')"
	end If

	tmpBatchNumber=""
	BatchNumber=""
	if trim(request("Sys_BatchNumber"))<>"" Then
		tmpBatchNumber=Split(request("Sys_BatchNumber"),",")
		For i=0 To UBound(tmpBatchNumber)
			BatchNumber=BatchNumber&tmpBatchNumber(i)&"','"
		next
		
			strwhere = strwhere & " and BillNo in (select Billno from dcilog where BatchNumber in ('"&BatchNumber&"')) "
	end If

	
	if trim(strwhere)<>"" Then
		strsql="select * from ("
		strsql=strsql&" select a.Billno,a.Carno,a.MailDate,a.MailReturnDate,a.ReturnResonID,"
		strsql=strsql&" b.Mailnumber,b.ProcDate,b.ProcTime,b.MailStatus,b.HandleBrueau,to_date(b.procdate,'yyyy/mm/dd')-a.MailDate as TDay "
		strsql=strsql&" from BILLMAILHISTORY a INNER JOIN MailResult b ON replace(a.mailchknumber,' ','')=b.Mailnumber"
		strsql=strsql&" ) c where c.Tday>0 "		
'	response.write strsql
		Set rsfound=conn.execute(strsql&strwhere&" order by Billno,ProcDate,ProcTime")

		strsql2="select count(*) as cnt from ("&strsql&strwhere&")"
		Set rscnt=conn.execute(strsql2)
			DBSUM=rscnt("cnt")
		Set rscnt=Nothing
		DB_Display="show"
	else
		DB_Display=""
		Response.write "<script>"
		Response.Write "alert('必須有查詢條件！');"
		Response.write "</script>"
	end if
end if
%>

<body>

<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><span class="style3">郵寄狀態查詢&nbsp;&nbsp;&nbsp;&nbsp;最後郵寄狀態匯入時間【<%=LastUpdateTime%>】</span></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						作業批號 
						<input name="Sys_BatchNumber" type="text" class="btn1" value="<%=UCase(request("Sys_BatchNumber"))%>" size="50" onkeyup="value=value.toUpperCase()">
						&nbsp;&nbsp;&nbsp;&nbsp;<a href="MailStatusQryFileLog.asp" target="_blank"><font  class="font10">-> 查詢匯入歷程紀錄</font></a>
						<br>
						舉發單號
						<input name="Sys_BillNo" type="text" class="btn1" value="<%=request("Sys_BillNo")%>" size="50" onkeyup="value=value.toUpperCase()">
						<img src="space.gif" width="3" height="10">
						<input type="button" name="btnSelt" value="查詢" class="btn3" style="width:60px; height:20px;" onclick="funSelt('Selt');">
						<input type="button" name="cancel" value="清除" class="btn3" style="width:40px; height:20px;" onClick="location='MailStatusQry.asp'">
						<br>
						<font  class="font10"> (多個單號、批號處理，用,隔開。如：96W161,96W162或S01688195,S01688189,S01688233）</font>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">
		郵寄狀態查詢紀錄
		每頁<select name="sys_MoveCnt" onchange="repage();">
				<option value="0"<%if trim(request("sys_MoveCnt"))="0" then response.write " Selected"%>>10</option>
				<option value="10"<%if trim(request("sys_MoveCnt"))="10" then response.write " Selected"%>>20</option>
				<option value="20"<%if trim(request("sys_MoveCnt"))="20" then response.write " Selected"%>>30</option>
				<option value="30"<%if trim(request("sys_MoveCnt"))="30" then response.write " Selected"%>>40</option>
				<option value="40"<%if trim(request("sys_MoveCnt"))="40" then response.write " Selected"%>>50</option>
				<option value="50"<%if trim(request("sys_MoveCnt"))="50" then response.write " Selected"%>>60</option>
				<option value="60"<%if trim(request("sys_MoveCnt"))="60" then response.write " Selected"%>>70</option>
				<option value="70"<%if trim(request("sys_MoveCnt"))="70" then response.write " Selected"%>>80</option>
				<option value="80"<%if trim(request("sys_MoveCnt"))="80" then response.write " Selected"%>>90</option>
				<option value="90"<%if trim(request("sys_MoveCnt"))="90" then response.write " Selected"%>>100</option>
			</select>筆<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄 </strong>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th class="font10">單號</th>
					<th class="font10">車號</th>
					<th class="font10">郵件號碼</th>
					<th class="font10">郵寄日期</th>
					<th class="font10">郵寄退回日期</th>
					<th class="font10">退件原因</th>
					<th class="font10">處理日期</th>
					<th class="font10">郵件狀態</th>
					<th class="font10">處理郵局</th>
				</tr>
				<%
				if request("DB_Selt")="Selt" Then
					If DB_Display="show" Then 
						if Trim(request("DB_Move"))="" then
							DBcnt=0
						else
							DBcnt=request("DB_Move")
						end if
						if Not rsfound.eof then rsfound.move DBcnt
						
						for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
							if rsfound.eof then exit for
							response.write "<tr bgcolor='#FFFFFF'"
							lightbarstyle 0 
							response.write ">"
							response.write "<td class=""font10"">&nbsp;"&rsfound("Billno")&"</td>"     
							response.write "<td class=""font10"">&nbsp;"&rsfound("CarNo")&"</td>"     
							response.write "<td class=""font10"">&nbsp;"&rsfound("Mailnumber")&"</td>"     
							response.write "<td class=""font10"">&nbsp;"&GetCDateTime(rsfound("Maildate"))&"</td>"     
							response.write "<td class=""font10"">&nbsp;"&GetCDate(rsfound("MailReturndate"))&"</td>"     
							response.write "<td class=""font10"">&nbsp;"&GetResonName(rsfound("ReturnResonID"))&"</td>"     
							response.write "<td class=""font10"">&nbsp;"&GetCDate2(rsfound("ProcDate"))&" "&trim(rsfound("ProcTime"))&"</td>"
							response.write "<td class=""font10"">&nbsp;"&trim(rsfound("MailStatus"))&"</td>"
							response.write "<td class=""font10"">&nbsp;"&trim(rsfound("HandleBrueau"))&"</td>"
							response.write "</tr>"
							response.flush
							rsfound.movenext
						Next
					End if
				End IF
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td height="30" bgcolor="#FFDD77" align="center">
			<a href="file:///.."></a>
			<input type="button" name="MoveFirst" value="第一頁" class="btn3" style="width:50px; height:20px;" onclick="funDbMove(0);">
			<input type="button" name="MoveUp" value="上一頁" class="btn3" style="width:50px; height:20px;" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(CDbl(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(CDbl(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" class="btn3" style="width:50px; height:20px;" onclick="funDbMove(10);">
			<input type="button" name="MoveDown" value="最後一頁" class="btn3" style="width:60px; height:20px;" onclick="funDbMove(999);">
			<br>
			<%If CDbl("0"&DBsum)<>"0" Then %>
			<input type="button" name="btnExecel" value="轉換成Excel" class="btn3" style="width:70px; height:25px;" onclick="funchgExecel();">
			<%End if%>
</table>
<br><b>
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="DB_Display" value="<%=DB_Display%>">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
<form Name=CarForm method="post">
<input type="Hidden" name="TempSQL" value="<%=strsql&strwhere%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
var winopen;
var sys_City='<%=sys_City%>';

function funSelt(DBKind){
	var error=0;
	if(DBKind=='Selt'){

				myForm.DB_Move.value="";
				myForm.DB_Selt.value=DBKind;
				myForm.submit();
		}
}
function fnBatchNumber(){
	myForm.Sys_BatchNumber.value=myForm.Selt_BatchNumber.value;
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	winopen=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	winopen.focus();
	return win;
}


function funchgExecel(){
	CarForm.action="MailStatusQry_Execel.asp";
	CarForm.target="inputWin";
	CarForm.submit();
	CarForm.action="";
	CarForm.target="";
}
function funDbMove(MoveCnt){
	if (eval(MoveCnt)==0){
		myForm.DB_Move.value="";
		myForm.submit();
	}else if (eval(MoveCnt)==10){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10-eval(myForm.sys_MoveCnt.value)){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt+eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else if(eval(MoveCnt)==-10){
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt-eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else if(eval(MoveCnt)==999){
		if (eval(myForm.DB_Cnt.value)%(10+eval(myForm.sys_MoveCnt.value))==0){
			myForm.DB_Move.value=(Math.floor(eval(myForm.DB_Cnt.value)/(10+eval(myForm.sys_MoveCnt.value)))-1)*(10+eval(myForm.sys_MoveCnt.value));
		}else{
			myForm.DB_Move.value=Math.floor(eval(myForm.DB_Cnt.value)/(10+eval(myForm.sys_MoveCnt.value)))*(10+eval(myForm.sys_MoveCnt.value));
		}
		myForm.submit();
	}
}
function repage(){
	myForm.DB_Move.value=0;
	myForm.submit();
}
</script>
<%conn.close%>                       