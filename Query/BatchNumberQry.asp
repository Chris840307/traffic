<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title></title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<!--#include file="sqlDCIExchangeData.asp"-->

<%


%>
<html>
<style type="text/css">
<!--
.style5 {
	font-size: 10pt;
}
.style7 {
	font-size: 10pt;
	font-family: "標楷體";}
.style8 {
	font-size: 14pt;
	}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style11 {
	font-size: 10px;
	font-family: "標楷體";
}
.style22 {font-size: 9pt; font-family: "標楷體"; }
-->
</style>
</head>
<body>
<form name="myForm" method="post">
<table width='1000' border='1' align="center" cellpadding="1">
	<tr>
		<td bgcolor="#FFCC99">建檔日期
		</td>
		<td colspan="6">
			<input type="text" name="RecordDate1" value="<%
		If Trim(request("RecordDate1"))="" Then
			response.write Year(now)-1911 & Right("00"&Month(now),2) & Right("00"&day(now),2)
		Else
			response.write Trim(request("RecordDate1"))
		End If 
			%>" maxlength="7" onKeyup="value=value.replace(/[^\d]/g,'')"> ~
			<input type="text" name="RecordDate2" value="<%
		If Trim(request("RecordDate2"))="" Then
			response.write Year(now)-1911 & Right("00"&Month(now),2) & Right("00"&day(now),2)
		Else
			response.write Trim(request("RecordDate2"))
		End If 
			%>" maxlength="7" onKeyup="value=value.replace(/[^\d]/g,'')">&nbsp; 
			<input type="button" value="查 詢" onclick="BatchNumberQry();">
		</td>
		<input type="hidden" value="" name="kinds">
	</tr>
	<tr>
		<td colspan="7" bgcolor="#CCFFFF">
			逕舉
		</td>
	</tr>
	<tr bgcolor="#FFFFCC">
		<td width="15%">舉發單位</td>
		<td width="15%">建檔日期</td>
		<td width="20%">車籍查詢</td>
		<td width="20%">入案</td>
		<td width="10%">列印人/移送清冊</td>
		<td width="10%">列印人/舉發單</td>
		<td width="10%">列印人/大宗單</td>
	</tr>
<%
If Trim(request("kinds"))="BatchNumberQry" Then
	RecordDate1=gOutDT(request("RecordDate1"))&" 0:0:0"
	RecordDate2=gOutDT(request("RecordDate2"))&" 23:59:59"
	strwhere=" and RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
Else
	RecordDate1=date&" 0:0:0"
	RecordDate2=date&" 23:59:59"
	strwhere=" and RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
End If 

strSql="select distinct Batchnumber from Dcilog where (BillTypeID='2' and ExchangeTypeID='A') "&strwhere&" order by length(Batchnumber) desc ,Batchnumber desc"
set rs1=conn.execute(strSql)
If Not rs1.Bof Then rs1.MoveFirst 
While Not rs1.Eof
	ExchangeDate=""
	DciReturnStatusID=""
	QryDciRecordMember=""
	strB1="select ExchangeDate,DciReturnStatusID,(select Chname from Memberdata where MemberID=dcilog.RecordMemberID) as RecordMember from dcilog where batchnumber='"&trim(rs1("Batchnumber"))&"' and rownum<=1"
	set rsb1=conn.execute(strB1)
	if not rsb1.eof then
		ExchangeDate=trim(rsb1("ExchangeDate"))
		DciReturnStatusID=trim(rsb1("DciReturnStatusID"))
		QryDciRecordMember=trim(rsb1("RecordMember"))
	end if
	rsb1.close
	set rsb1=nothing 
	DciCnt=0
	strB1="select count(*) as cnt from dcilog where batchnumber='"&trim(rs1("Batchnumber"))&"'"
	set rsb1=conn.execute(strB1)
	if not rsb1.eof then
		DciCnt=trim(rsb1("cnt"))
	end if
	rsb1.close
	set rsb1=nothing 

	BillUnit=""
	strc2="select distinct b.UnitTypeID from BillBase a,UnitInfo b where a.RecordStateID=0 " &_
		" and a.BillUnitID=b.UnitID and a.SN in (select BillSn from dcilog where batchnumber='"&trim(rs1("Batchnumber"))&"')" &_
		" order by UnitTypeID"
	set rsc2=conn.execute(strc2)
	If Not rsc2.Bof Then rsc2.MoveFirst 
	While Not rsc2.Eof
		
		strU="select UnitName from UnitInfo where UnitID='"&trim(rsc2("UnitTypeID"))&"'"
		set rsU=conn.execute(strU)
		if not rsU.eof then
			if BillUnit="" then
				BillUnit=trim(rsU("UnitName"))
			else
				BillUnit=BillUnit&"<br>"&trim(rsU("UnitName"))
			end if 
		end if
		rsU.close
		set rsU=nothing 
	
	rsc2.MoveNext
	Wend
	rsc2.close
	set rsc2=nothing

	RecordDateTmp=""
	RecordMember=""
	strc2="select RecordDate,(select Chname from Memberdata where MemberID=BillBase.RecordMemberID) as RecordMember from BillBase " &_
		" where RecordStateID=0 and rownum<=1" &_
		" and SN in (select BillSn from dcilog where batchnumber='"&trim(rs1("Batchnumber"))&"')" 
	set rsc2=conn.execute(strc2)
	If Not rsc2.Bof Then 
		RecordDateTmp=Trim(rsc2("RecordDate"))
		RecordMember=Trim(rsc2("RecordMember"))
	End If 
	rsc2.close
	set rsc2=nothing
%>
	<tr>
		<td><%=BillUnit%></td>
		<td><%=RecordDateTmp&"<br>"&RecordMember%></td>
		<td>
		<strong><%=trim(rs1("Batchnumber"))%></strong>
		<br>
		<%=trim(ExchangeDate)&"<br>"&QryDciRecordMember%>
		<br>
		<%="共 "&DciCnt&" 筆  ."%>
		<%
		if DciReturnStatusID="" then 
			response.write "<font color=""red"">未處理</font>"
		else
			response.write "<font color=""green"">已處理</font>"
		end if 
		%>
		</td>
		<td>
<%
	CaseInBatchnumber=""
	strc1="select distinct Batchnumber from Dcilog where BillTypeID=2 and exchangeTypeID='W' and BillSn in (select BillSn from Dcilog where Batchnumber='"&trim(rs1("Batchnumber"))&"')"
	set rsc1=conn.execute(strc1)
	if not rsc1.eof then
		CaseInBatchnumber=Trim(rsc1("Batchnumber"))
		ExchangeDate2=""
		DciReturnStatusID2=""
		DciCnt2=0
		CaseInDciRecordMember=""
		strc2="select ExchangeDate,DciReturnStatusID,(select Chname from Memberdata where MemberID=dcilog.RecordMemberID) as RecordMember from Dcilog where BillTypeID=2 and exchangeTypeID='W' and BillSn in (select BillSn from Dcilog where Batchnumber='"&trim(rs1("Batchnumber"))&"')"
		set rsc2=conn.execute(strc2)
		if not rsc2.eof then
			ExchangeDate2=trim(rsc2("ExchangeDate"))
			DciReturnStatusID2=trim(rsc2("DciReturnStatusID"))
			CaseInDciRecordMember=trim(rsc2("RecordMember"))
		end if
		rsc2.close
		set rsc2=nothing

		strc2="select Count(*) as cnt from Dcilog where BillTypeID=2 and exchangeTypeID='W' and BillSn in (select BillSn from Dcilog where Batchnumber='"&trim(rs1("Batchnumber"))&"')"
		set rsc2=conn.execute(strc2)
		if not rsc2.eof then
			DciCnt2=trim(rsc2("cnt"))
		end if
		rsc2.close
		set rsc2=nothing
%>
		<strong><%=trim(rsc1("Batchnumber"))%></strong>
		<br>
		<%=trim(ExchangeDate2)&"<br>"&CaseInDciRecordMember%>
		<br>
		<%="共 "&DciCnt2&" 筆  ."%>
		<%
		if DciReturnStatusID2="" then 
			response.write "<font color=""red"">未處理</font>"
		else
			response.write "<font color=""green"">已處理</font>"
		end if 
		%>
<%	Else
		response.write "&nbsp;"
	end if
	rsc1.close
	set rsc1=nothing
%>
		</td>
		<td><%
	strJ1="select (select ChName from memberdata where MemberID=BatchnumberJob.MemberID) as RecMem,RecordDate from BatchnumberJob where UPPER(batchnumber)='"&Trim(CaseInBatchnumber)&"' and PrintTypeID=0"
	Set rsJ1=conn.execute(strJ1)
	If Not rsJ1.eof Then
		response.write Trim(rsJ1("RecordDate"))&"<br>"&Trim(rsJ1("RecMem"))
	Else
		response.write "&nbsp;"
	End If
	rsJ1.close
	Set rsJ1=Nothing 
		%></td>
		<td><%
	strJ1="select (select ChName from memberdata where MemberID=BatchnumberJob.MemberID) as RecMem,RecordDate from BatchnumberJob where UPPER(batchnumber)='"&Trim(CaseInBatchnumber)&"' and PrintTypeID=1"
	Set rsJ1=conn.execute(strJ1)
	If Not rsJ1.eof Then
		response.write Trim(rsJ1("RecordDate"))&"<br>"&Trim(rsJ1("RecMem"))
	Else
		response.write "&nbsp;"
	End If
	rsJ1.close
	Set rsJ1=Nothing 
		%></td>
		<td><%
	strJ1="select (select ChName from memberdata where MemberID=BatchnumberJob.MemberID) as RecMem,RecordDate from BatchnumberJob where UPPER(batchnumber)='"&Trim(CaseInBatchnumber)&"' and PrintTypeID=2"
	Set rsJ1=conn.execute(strJ1)
	If Not rsJ1.eof Then
		response.write Trim(rsJ1("RecordDate"))&"<br>"&Trim(rsJ1("RecMem"))
	Else
		response.write "&nbsp;"
	End If
	rsJ1.close
	Set rsJ1=Nothing 
		%></td>
	</tr>
<%
	rs1.MoveNext
Wend
rs1.close
set rs1=nothing
%>
	<tr>
		<td colspan="7" bgcolor="#CCFFFF">
			攔停
		</td>
	</tr>
	<tr bgcolor="#FFFFCC">
		<td >舉發單位</td>
		<td >建檔日期</td>
		<td colspan="2">入案</td>
		<td >列印人/移送清冊</td>
		<td >列印人/舉發單</td>
		<td >列印人/大宗單</td>
	</tr>
<%
strSql="select distinct Batchnumber from Dcilog where (BillTypeID='1' and ExchangeTypeID='W') "&strwhere&" order by length(Batchnumber) desc ,Batchnumber desc"
set rs1=conn.execute(strSql)
If Not rs1.Bof Then rs1.MoveFirst 
While Not rs1.Eof
	ExchangeDate=""
	DciReturnStatusID=""
	QryDciRecordMember=""
	strB1="select ExchangeDate,DciReturnStatusID,(select Chname from Memberdata where MemberID=dcilog.RecordMemberID) as RecordMember from dcilog where batchnumber='"&trim(rs1("Batchnumber"))&"' and rownum<=1"
	set rsb1=conn.execute(strB1)
	if not rsb1.eof then
		ExchangeDate=trim(rsb1("ExchangeDate"))
		DciReturnStatusID=trim(rsb1("DciReturnStatusID"))
		QryDciRecordMember=trim(rsb1("RecordMember"))
	end if
	rsb1.close
	set rsb1=nothing 
	DciCnt=0
	strB1="select count(*) as cnt from dcilog where batchnumber='"&trim(rs1("Batchnumber"))&"'"
	set rsb1=conn.execute(strB1)
	if not rsb1.eof then
		DciCnt=trim(rsb1("cnt"))
	end if
	rsb1.close
	set rsb1=nothing 

	BillUnit=""
	strc2="select distinct b.UnitTypeID from BillBase a,UnitInfo b where a.RecordStateID=0 " &_
		" and a.BillUnitID=b.UnitID and a.SN in (select BillSn from dcilog where batchnumber='"&trim(rs1("Batchnumber"))&"')" &_
		" order by UnitTypeID"
	set rsc2=conn.execute(strc2)
	If Not rsc2.Bof Then rsc2.MoveFirst 
	While Not rsc2.Eof
		
		strU="select UnitName from UnitInfo where UnitID='"&trim(rsc2("UnitTypeID"))&"'"
		set rsU=conn.execute(strU)
		if not rsU.eof then
			if BillUnit="" then
				BillUnit=trim(rsU("UnitName"))
			else
				BillUnit=BillUnit&"<br>"&trim(rsU("UnitName"))
			end if 
		end if
		rsU.close
		set rsU=nothing 
	
	rsc2.MoveNext
	Wend
	rsc2.close
	set rsc2=nothing

	RecordDateTmp=""
	RecordMember=""
	strc2="select RecordDate,(select Chname from Memberdata where MemberID=BillBase.RecordMemberID) as RecordMember from BillBase " &_
		" where RecordStateID=0 and rownum<=1" &_
		" and SN in (select BillSn from dcilog where batchnumber='"&trim(rs1("Batchnumber"))&"')" 
	set rsc2=conn.execute(strc2)
	If Not rsc2.Bof Then 
		RecordDateTmp=Trim(rsc2("RecordDate"))
		RecordMember=Trim(rsc2("RecordMember"))
	End If 
	rsc2.close
	set rsc2=nothing
%>
	<tr>
		<td><%=BillUnit%></td>
		<td><%=RecordDateTmp&"<br>"&RecordMember%></td>
		<td colspan="2">
		<strong><%=trim(rs1("Batchnumber"))%></strong>
		<br>
		<%=trim(ExchangeDate)&"<br>"&QryDciRecordMember%>
		<br>
		<%="共 "&DciCnt&" 筆  ."%>
		<%
		if DciReturnStatusID="" then 
			response.write "<font color=""red"">未處理</font>"
		else
			response.write "<font color=""green"">已處理</font>"
		end if 
		%>
		</td>
		<td><%
	strJ1="select (select ChName from memberdata where MemberID=BatchnumberJob.MemberID) as RecMem,RecordDate from BatchnumberJob where UPPER(batchnumber)='"&Trim(rs1("Batchnumber"))&"' and PrintTypeID=0"
	Set rsJ1=conn.execute(strJ1)
	If Not rsJ1.eof Then
		response.write Trim(rsJ1("RecordDate"))&"<br>"&Trim(rsJ1("RecMem"))
	Else
		response.write "&nbsp;"
	End If
	rsJ1.close
	Set rsJ1=Nothing 
		%></td>
		<td><%
	strJ1="select (select ChName from memberdata where MemberID=BatchnumberJob.MemberID) as RecMem,RecordDate from BatchnumberJob where UPPER(batchnumber)='"&Trim(rs1("Batchnumber"))&"' and PrintTypeID=1"
	Set rsJ1=conn.execute(strJ1)
	If Not rsJ1.eof Then
		response.write Trim(rsJ1("RecordDate"))&"<br>"&Trim(rsJ1("RecMem"))
	Else
		response.write "&nbsp;"
	End If
	rsJ1.close
	Set rsJ1=Nothing 
		%></td>
		<td><%
	strJ1="select (select ChName from memberdata where MemberID=BatchnumberJob.MemberID) as RecMem,RecordDate from BatchnumberJob where UPPER(batchnumber)='"&Trim(rs1("Batchnumber"))&"' and PrintTypeID=2"
	Set rsJ1=conn.execute(strJ1)
	If Not rsJ1.eof Then
		response.write Trim(rsJ1("RecordDate"))&"<br>"&Trim(rsJ1("RecMem"))
	Else
		response.write "&nbsp;"
	End If
	rsJ1.close
	Set rsJ1=Nothing 
		%></td>
	</tr>
<%
	rs1.MoveNext
Wend
rs1.close
set rs1=nothing
%>
</table>
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function BatchNumberQry(){
	var error=0;
		var errorString="";
		if(myForm.RecordDate1.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入建檔日期!!";
		}else if(myForm.RecordDate1.value!=""){
			if(!dateCheck(myForm.RecordDate1.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}
		}
		if(myForm.RecordDate2.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入建檔日期!!";
		}else if(myForm.RecordDate2.value!=""){
			if(!dateCheck(myForm.RecordDate2.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：建檔日期輸入不正確!!";
			}
		}

		if (error>0){
			alert(errorString);
		}else{
			myForm.kinds.value="BatchNumberQry";
			myForm.submit();
		}
}
</script>
