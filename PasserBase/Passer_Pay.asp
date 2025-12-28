<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>繳款登錄</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
'response.write request("PBillSN")
'檢查是否可進入本系統
AuthorityCheck(224)
memID=Session("User_ID")

Function ChkNum(strValue)
	if ISNull(strValue) or trim(strValue)="" or IsEmpty(strValue) then
		ChkNum="null"
	else
		ChkNum=strValue
	end if
End Function

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

if trim(request("kinds"))="UpMoney" then

	if trim(request("IsLate"))="1" then 
		strUpd="Update PasserBase set ForFeit1="&trim(request("L2ForFeit"))&" where SN="&trim(request("PBillSN"))
		conn.execute strUpd
		ConnExecute strUpd,353

		If not ifnull(trim(request("Rule2_L2ForFeit"))) Then

			strUpd="Update PasserBase set ForFeit2="&trim(request("Rule2_L2ForFeit"))&" where rule2 is not null and SN="&trim(request("PBillSN"))
			conn.execute strUpd
			ConnExecute strUpd,353

		End if 
	else

		strUpd="Update PasserBase set ForFeit1="&trim(request("L1ForFeit"))&" where SN="&trim(request("PBillSN"))
		conn.execute strUpd
		ConnExecute strUpd,353

		If not ifnull(trim(request("Rule2_L1ForFeit"))) Then

			strUpd="Update PasserBase set ForFeit2="&trim(request("Rule2_L1ForFeit"))&" where rule2 is not null and SN="&trim(request("PBillSN"))
			conn.execute strUpd
			ConnExecute strUpd,353

		End if 
	end if 
	%>
	<script language="JavaScript">
		alert("裁罰金額修改完成");
	</script>
	<%
end if

if trim(request("kinds"))="Del" then
	strSQL="Delete from PasserPay where BillSN="&trim(request("BillSN"))&" and PayTimes="&trim(request("BillTime"))
	conn.execute(strSQL)
	%>
	<script language="JavaScript">
		alert("刪除完成");
	</script>
	<%
end if

if trim(request("kinds"))="db_insert" then
	'繳費次數
	MaxTime=1
	strTime="select max(PayTimes) as MaxTime from PasserPay where BillSN="&trim(request("BillSN"))
	set rsTime=conn.execute(strTime)
	if not rsTime.eof then
		if trim(rsTime("MaxTime"))="" or isnull(rsTime("MaxTime")) then
			MaxTime=1
		else
			MaxTime=cint(trim(rsTime("MaxTime")))+1
		end if
	end if
	rsTime.close
	set rsTime=nothing
	'繳費日期
	'if trim(request("PayTypeID"))="1" then
		'thePayDate=date
	'else
		thePayDate=gOutDT(request("PayDate"))
	'end if

	'結案

	if not ifnull(request("MemberStation")) then
		
		strUp="update PasserConfisCate set DCISTATIONID='"&trim(request("MemberStation"))&"' where BILLSN="&trim(request("BillSN"))
		
		conn.execute strUp
	End if 

	if trim(request("CaseClose"))="1" then

		theCaseClose=1
		strUpd="Update PasserBase set BillStatus='9',ForFeit1="&trim(request("ForFeit1"))&" where SN="&trim(request("BillSN"))
		conn.execute strUpd
		ConnExecute strUpd,353
		
		If not ifnull(trim(request("ForFeit2"))) Then

			strUpd="Update PasserBase set ForFeit2="&trim(request("ForFeit2"))&" where SN="&trim(request("BillSN"))
			conn.execute strUpd
			ConnExecute strUpd,353

		End if 
	else
		theCaseClose=0
		strUpd="Update PasserBase set BillStatus='0',ForFeit1="&trim(request("ForFeit1"))&" where SN="&trim(request("BillSN"))
		conn.execute strUpd
		ConnExecute strUpd,353

		If not ifnull(trim(request("ForFeit2"))) Then

			strUpd="Update PasserBase set ForFeit2="&trim(request("ForFeit2"))&" where rule2 is not null and SN="&trim(request("BillSN"))
			conn.execute strUpd
			ConnExecute strUpd,353

		End if 
	end If 

	if trim(request("Sys_PasserNote"))<>"" then

		strUpd="Update PasserBase set note=note||'"&trim(request("Sys_PasserNote"))&"' where SN="&trim(request("BillSN"))

		conn.execute strUpd

	end if

	if trim(request("PayMIDDLEMONEY"))="" then
		PayMIDDLEMONEY=0
	else
		PayMIDDLEMONEY=trim(request("PayMIDDLEMONEY"))
	end if
	if (trim(request("PayAmount"))<>"" and trim(request("PayNo"))<>"") or trim(request("Note"))<>"" then
		session("cache_PayTypeID")=trim(request("PayTypeID"))
		strIns="insert into PasserPay(BillSN,BillNo,PayNo,PayTimes,PayTypeID,PayDate,Payer,ForFeit" &_
			",PayAmount,CaseClose,RecordStateID,RecordDate,RecordMemberID,Note,IsLate,MIDDLEMONEY)" &_
			"values("&trim(request("BillSN"))&",'"&trim(request("BillNo"))&"'" &_
			",'"&trim(request("PayNo"))&"',"&MaxTime&","&trim(request("PayTypeID")) &_
			",TO_DATE('"&thePayDate&"','YYYY/MM/DD'),'"&trim(request("DRIVER"))&"'" &_
			","&trim(request("ForFeit"))&","&ChkNum(request("PayAmount"))&","&theCaseClose &_
			",0,sysdate,'"&memID&"','"&trim(request("Note"))&"','"&trim(request("IsLate"))&"'" &_
			","&PayMIDDLEMONEY&_
			")" 
		conn.execute strIns		
		ConnExecute strIns,353
		if theCaseClose=1 then
			strIns="Update PasserPay set CaseCloseDate=TO_DATE('"&gOutDT(request("CaseCloseDate"))&"','YYYY/MM/DD') where BillSN="&trim(request("BillSN"))
			conn.execute strIns
		else
			strIns="Update PasserPay set CaseClose=0,CaseCloseDate=null where BillSN="&trim(request("BillSN"))
			conn.execute strIns
		end if
		%>
		<script language="JavaScript">
			alert("新增完成");
		</script>
		<%
	else
		strUpdate="Update PasserPay set ForFeit="&trim(request("ForFeit"))&" where BillSN="&trim(request("BillSN"))
		conn.execute(strUpdate)

		if theCaseClose=1 then
			strIns="Update PasserPay set CaseCloseDate=TO_DATE('"&gOutDT(request("CaseCloseDate"))&"','YYYY/MM/DD') where BillSN="&trim(request("BillSN"))
			conn.execute strIns
		else
			strIns="Update PasserPay set CaseClose=0,CaseCloseDate=null where BillSN="&trim(request("BillSN"))
			conn.execute strIns
		end if
		
		%>
		<script language="JavaScript">
			alert("修改完成");
			opener.myForm.submit(); 
			self.close();
		</script>
		<%
	end if
end if

	strSql="select a.*,b.ForFeit,b.CaseClose from PasserBase a,PasserPay b where a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and SN="&trim(request("PBillSN"))
	set rsSql=conn.execute(strSql)
%>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">
<form name=myForm method="post">
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="style5">繳款登錄</span></td>
  </tr>
  <tr>
    <td height="26"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="13%" nowrap bgcolor="#FFFF99"><div align="right"><span class="style3">舉發單號</span></div></td>
        <td width="30%"><%
		theBillNo=""
		if trim(rsSql("BILLNO"))<>"" and not isnull(rsSql("BILLNO")) then
			response.write trim(rsSql("BILLNO"))
			theBillNo=trim(rsSql("BILLNO"))
		end if
		%>
		<input type="hidden" name="BillNo" value="<%=theBillNo%>">
		<input type="hidden" name="BillSN" value="<%=trim(request("PBillSN"))%>">
		</td>
        <td nowrap bgcolor="#FFFF99" width="13%"><div align="right">違規人</div></td>
        <td width="44%"><%
		theDRIVER=""
		if trim(rsSql("DRIVER"))<>"" and not isnull(rsSql("DRIVER")) then
			response.write trim(rsSql("DRIVER"))
			theDRIVER=trim(rsSql("DRIVER"))
		end if
		%>
		<input type="hidden" name="DRIVER" value="<%=theDRIVER%>">
		</td>
      </tr>
      <tr>
        <td nowrap bgcolor="#FFFF99"><div align="right">應到案日期</div></td>
        <td width="19%"><%
		if trim(rsSql("DealLineDate"))<>"" and not isnull(rsSql("DealLineDate")) then
			response.write gInitDT(trim(rsSql("DealLineDate")))
		end if			  
		  %></td>
        <td width="9%" nowrap bgcolor="#FFFF99"><div align="right">違規法條</div></td>
        <td width="59%"><%
		if trim(rsSql("Rule1"))<>"" and not isnull(rsSql("Rule1")) then
			response.write trim(rsSql("Rule1"))&"，"
			strRule1="select * from Law where ItemID='"&trim(rsSql("Rule1"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				L1ForFeit=cint(trim(rsRule1("Level1")))
				L2ForFeit=cint(trim(rsRule1("Level2")))
				L3ForFeit=cint(trim(rsRule1("Level3")))
				L4ForFeit=cint(trim(rsRule1("Level4")))

				response.write cint(trim(rsRule1("Level1")))
				if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level1")))
				else
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level2")))
				end if
				response.write "&nbsp; ,"&cint(trim(rsRule1("Level3")))
				response.write "&nbsp; ,"&cint(trim(rsRule1("Level4")))
				response.write "<br>"&trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		end if	
		Rule2_L1ForFeit=0:Rule2_L2ForFeit=0:Rule2_L3ForFeit=0:Rule2_L4ForFeit=0
		if trim(rsSql("Rule2"))<>"" and not isnull(rsSql("Rule2")) then
			response.write "<br>"&trim(rsSql("Rule2"))&"，"
			strRule1="select * from Law where ItemID='"&trim(rsSql("Rule2"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				Rule2_L1ForFeit=cint(trim(rsRule1("Level1")))
				Rule2_L2ForFeit=cint(trim(rsRule1("Level2")))
				Rule2_L3ForFeit=cint(trim(rsRule1("Level3")))
				Rule2_L4ForFeit=cint(trim(rsRule1("Level4")))

				response.write cint(trim(rsRule1("Level1")))
				if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level1")))
				else
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level2")))
				end if
				response.write "&nbsp; ,"&cint(trim(rsRule1("Level3")))
				response.write "&nbsp; ,"&cint(trim(rsRule1("Level4")))
				response.write "<br>"&trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		end if	
		if trim(rsSql("Rule3"))<>"" and not isnull(rsSql("Rule3")) then
			response.write "<br>"&trim(rsSql("Rule3"))&"，"
			strRule1="select * from Law where ItemID='"&trim(rsSql("Rule3"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				response.write cint(trim(rsRule1("Level1")))
				if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level1")))
				else
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level2")))
				end if
				response.write "&nbsp; ,"&cint(trim(rsRule1("Level3")))
				response.write "&nbsp; ,"&cint(trim(rsRule1("Level4")))
				response.write "<br>"&trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		end if	
		if trim(rsSql("Rule4"))<>"" and not isnull(rsSql("Rule4")) then
			response.write "<br>"&trim(rsSql("Rule4"))&"，"
			strRule1="select * from Law where ItemID='"&trim(rsSql("Rule4"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				response.write cint(trim(rsRule1("Level1")))
				if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level1")))
				else
					response.write "&nbsp; ,"&cint(trim(rsRule1("Level2")))
				end if
				response.write "&nbsp; ,"&cint(trim(rsRule1("Level3")))
				response.write "&nbsp; ,"&cint(trim(rsRule1("Level4")))
				response.write "<br>"&trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		end if	
		%></td>
      </tr>
      <tr>
        <td nowrap bgcolor="#FFFF99" height="27"><div align="right"><span class="style3">已繳金額</span></div></td>
        <td width="19%"><%
			strPay="select sum(PayAmount) as PaySum from PasserPay where BillSN="&trim(request("PBillSN"))
			set rsPay=conn.execute(strPay)
			if trim(rsPay("PaySum"))="" or isnull(rsPay("PaySum")) then
				response.write "0"
			else
				response.write trim(rsPay("PaySum"))
			end if
			rsPay.close
			set rsPay=nothing
			%></td>
			<td nowrap bgcolor="#FFFF99"><div align="right" class="style3">裁罰金額</div></td>
			<td><%
			theForFeit1="":theForFeit2=""
			if sys_City = "宜蘭縣" then
				if datediff("d",trim(rsSql("DealLineDate")),now)=<0 then

					theForFeit1=L1ForFeit:theForFeit2=Rule2_L1ForFeit

				elseif datediff("d",trim(rsSql("DealLineDate")),now)>0 and datediff("d",trim(rsSql("DealLineDate")),now)<=15 then

					theForFeit1=L2ForFeit:theForFeit2=Rule2_L2ForFeit

				elseif datediff("d",trim(rsSql("DealLineDate")),now)>15 and datediff("d",trim(rsSql("DealLineDate")),now)<=30 then

					theForFeit1=L3ForFeit:theForFeit2=Rule2_L3ForFeit

				elseif datediff("d",trim(rsSql("DealLineDate")),now)>30 then

					theForFeit1=L4ForFeit:theForFeit2=Rule2_L4ForFeit

				end If 
			else

				theForFeit1=trim(rsSql("ForFeit1"))

				if trim(rsSql("Rule2"))<>"" and not isnull(rsSql("Rule2")) then theForFeit2=trim(rsSql("ForFeit2"))

			end If 

		%>
		法條1<input type="text" name="ForFeit1" value="<%=theForFeit1%>" onkeyup="<%
			Response.Write "if(!myForm.ForFeit2.value){"
			Response.Write "myForm.ForFeit.value=eval(myForm.ForFeit1.value);"
			Response.Write "}else{"
			Response.Write "myForm.ForFeit.value=eval(myForm.ForFeit1.value)+eval(myForm.ForFeit2.value);"
			Response.Write "}"
		%>" <%
			if trim(Session("Credit_ID"))<>"A000000000" Then Response.Write " ReadOnly"	
		%>><br>
		法條2<input type="text" name="ForFeit2" value="<%=theForFeit2%>" onkeyup="<%
			Response.Write "if(!myForm.ForFeit2.value){"
			Response.Write "myForm.ForFeit.value=eval(myForm.ForFeit1.value);"
			Response.Write "}else{"
			Response.Write "myForm.ForFeit.value=eval(myForm.ForFeit1.value)+eval(myForm.ForFeit2.value);"
			Response.Write "}"
		%>" <%
			if trim(Session("Credit_ID"))<>"A000000000" Then Response.Write " ReadOnly"	
		%>><br>
		合計:<input type="text" name="ForFeit" Readonly value="<%
			tmpForfeit=0

			If not ifnull(theForFeit2) Then

				tmpForfeit=cdbl(theForFeit1)+cdbl(theForFeit2)
			else

				tmpForfeit=cdbl(theForFeit1)
			End if 

			strSQL="select nvl(sum(PayAmount),0) as PaySum from PasserPay where BillSN="&trim(request("PBillSN"))
			set rspay=conn.execute(strSQL)
			if not rspay.eof then Sys_PaySum=cdbl(rspay("PaySum"))
			rspay.close

			totalMoney=tmpForfeit-Sys_PaySum
			
			If totalMoney < 0 Then totalMoney=0

			response.write totalMoney
		%>">		
		</td>
      </tr>
      <tr>
        <td nowrap bgcolor="#FFFF99"><div align="right">繳費方式</div></td>
        <td>
			<input type="radio" value="2" name="PayTypeID"<%if trim(session("cache_PayTypeID"))="2" then response.write " checked"%>>郵撥
			<input type="radio" value="1" name="PayTypeID"<%
				if trim(session("cache_PayTypeID"))="1" then
					response.write " checked"
				elseif trim(session("cache_PayTypeID"))="" then
					response.write " checked"
				end if%>>窗口
			<input type="radio" value="3" name="PayTypeID"<%if trim(session("cache_PayTypeID"))="3" then response.write " checked"%>>其他代收單位
		</td>
		<td width="9%" nowrap bgcolor="#FFFF99"><div align="right" class="style3">手續費</div></td>
        <td>
			<input type="text" name="PayMIDDLEMONEY" value="" maxlength="6" onkeyup="value=value.replace(/[^\d]/g,'')">
		</td>
      </tr>

      <tr>
        <td nowrap bgcolor="#FFFF99"><div align="right"><span class="style3">逾期與否</span></div></td>
        <td width="19%"><%
			tmpIsLate="0":Sys_JudeDate=""
			strSQL="select JudeDate from PasserJude where BillSn="&trim(request("PBillSN"))
			set rsjude=conn.execute(strSQL)
			If Not rsjude.eof Then
				Sys_JudeDate="裁決日："&gInitDT(rsjude("JudeDate"))
				tmpIsLate="1"
			else
				Sys_JudeDate="無裁決"
				tmpIsLate="0"
			End if
			If not ifnull(Request("IsLate")) Then tmpIsLate=trim(Request("IsLate"))

			Response.Write "<select name=""IsLate"" onchange=""funUpMoney('"&trim(request("PBillSN"))&"');"">"

			Response.Write "<option value=""0"""
			If tmpIsLate = "0" Then Response.Write " selected"
			Response.Write ">如期繳納</option>"

			Response.Write "<option value=""1"""
			If tmpIsLate = "1" Then Response.Write " selected"
			Response.Write ">逾期繳納</option>"
			Response.Write "</select>"
			Response.Write "　　"&Sys_JudeDate
		%></td>
        <td width="9%" nowrap bgcolor="#FFFF99"><div align="right" class="style3">繳費金額</div></td>
        <td>
			<input type="text" name="PayAmount" value="<%

				If not ifnull(theForFeit2) Then

					theForFeit=cdbl(theForFeit1)+cdbl(theForFeit2)
				else

					theForFeit=cdbl(theForFeit1)
				End if 	
				Sys_PaySum=0
				strSQL="select nvl(sum(PayAmount),0) as PaySum from PasserPay where BillSN="&trim(request("PBillSN"))
				set rspay=conn.execute(strSQL)
				if not rspay.eof then Sys_PaySum=cdbl(rspay("PaySum"))
				rspay.close
				totalMoney=theForFeit-Sys_PaySum
				response.write totalMoney

			%>" maxlength="6" onkeyup="funChkMoney(this,'<%=Sys_PaySum%>');">
			<span><font color="red" size="3"><%
			If sys_City = "彰化縣" Then
				strSQL="select to_char(JudeDate,'YYYY') JudeDate from PasserJude where BillSN="&trim(request("PBillSN"))
				set rsjude=conn.execute(strSQL)
				If not rsjude.eof Then
					If cdbl(rsjude("JudeDate"))=< cdbl(Year(date)) Then
						Response.Write "保留至"&(cdbl(rsjude("JudeDate"))-1911)&"年度"
					end if
				end if
				rsjude.close
			end if
		%></font></span>
		</td>
      </tr>
      <tr>
        <td nowrap bgcolor="#FFFF99"><div align="right"><span class="style3">繳費日期</span></div></td>
        <td width="19%">
		<input type="text" name="PayDate" size="6" maxlength="7" value="<%=gInitDT(date)%>" onkeyup="value=value.replace(/[^\d]/g,'')">
		<input type="button" name="datestr" value="..." onclick="OpenWindow('PayDate');">
		</td>
        <td nowrap bgcolor="#FFFF99"><div align="right"><span class="style3">收據字號</span></div></td>
        <td>
          <input name="PayNo" type="text" value="" size="31" maxlength="50">
          <br><font size="2">如民眾繳費，請必填 收據字號</font>
        </td>
      </tr>
      <tr>
		<td nowrap bgcolor="#FFFF99"><div align="right" class="style3">結案日期/狀態</div></td>
        <td nowrap bgcolor="#FFFFff">
			<input type="text" name="CaseCloseDate" size="6" maxlength="7" value="<%=gInitDT(date)%>" onkeyup="value=value.replace(/[^\d]/g,'')">
			<input type="button" name="datestr" value="..." onclick="OpenWindow('CaseCloseDate');">
			<strong><font color="red">結案(預設結案)</font></strong>
			<input type="checkbox" name="CaseClose" value="1" <%
				If Not ifnull(rsSql("CaseClose")) Then
					If trim(rsSql("CaseClose"))="1" Then response.write "checked"
				else
					response.write "checked"
				End if
			%>>
		</td>
        <td nowrap bgcolor="#FFFF99"><div align="right" class="style3">無法處理原因</div></td>
        <td><span class="style3">
          <input name="Note" type="text" size="31" maxlength="30" value="">
        </span></td>
      </tr>
	  <tr>
		<td nowrap bgcolor="#FFFF99"><div align="right" class="style3">備註</div></td>
		<td nowrap bgcolor="#FFFFff">
			<input type="text" name="Sys_PasserNote" size="50">
		</td>

		<td nowrap bgcolor="#FFFF99"><div align="right" class="style3">扣件移送監理站</div></td>
		<td nowrap bgcolor="#FFFFff">
			<table border=0>
				<tr>
					<td>
						<%
							DCISTATIONNAME="":DCISTATIONID=""
							strUp="select (select DCISTATIONNAME from Station where StationID=PasserConfisCate.DCISTATIONID) DCISTATIONNAME,DCISTATIONID from PasserConfisCate where BILLSN="&trim(request("PBillSN"))

							set rsStion=conn.execute(strUp)

							If not rsStion.eof Then
								DCISTATIONID=trim(rsStion("DCISTATIONID"))
								DCISTATIONNAME=trim(rsStion("DCISTATIONNAME"))
							End if 
							rsStion.close
						%>
						
						<input type="text" name="MemberStation" onkeyup="getStation();" size="10" value="<%=DCISTATIONID%>">
						
						<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("../BillKeyIn/Query_Station.asp","WebPage1","left=0,top=0,location=0,width=760,height=575,resizable=yes,scrollbars=yes")'>

					</td>
					<td>

						<div id="Layer5" style="position:absolute ; width:120px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;">
						<%=DCISTATIONNAME%>
						</div>
					</td>
				</tr>
			</table>
		</td>
      </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77" colspan="4"><p align="center" class="style1">
        <img src="space.gif" width="14" height="8">
        <input name="subPayCLose" type="button" class="style3" value="儲存" onclick="db_insert();">
		<img src="space.gif" width="20" height="5"> 
       <!--  <input name="Submit433223" type="button" class="style3" value="結案">
        <img src="space.gif" width="20" height="5">  -->
		<input name="Submit433222" type="button" class="style3" value="關閉" onclick="funExit();">
        <br><font size="3"><font color="red"><strong> 分期繳款</strong></font> <font size="2">功能 . 輸入該期繳費金額 以及 收據字號 。 取消 結案 後點選 確定 即可記錄每次繳款。 最後一期繳款完成後再勾選 結案即可。 </font></font>
        <img src="space.gif" width="20" height="5">
		<!--<input type="reset" value="重置">-->
        <input type="hidden" name="kinds" value="">
        </p></td>
  </tr>
  <tr>
	<td colspan="4" bgcolor="#FFFFCC">
	<table width="100%" height="100%" border="0" bgcolor="#E0E0E0">
		<tr>
			<td colspan="9" bgcolor="#FFCC33">歷次繳費記錄</td>
		</tr>
		<tr bgcolor="#EBFBE3">
			<td width="10%" nowrap>繳費日期</td>
			<td width="10%" nowrap>繳費方式</td>
			<td width="10%" nowrap>繳費金額</td>
			<td width="10%" nowrap>結案日期</td>
			<td width="15%">逾期與否</td>
			<td width="20%">收據字號</td>
			<td width="20%">手續費</td>
			<td width="27%" nowrap>無法處理原因</td>
			<td width="8%">修改</td>
		</tr>
	<%
	strPayHis="select * from PasserPay where BillSN="&trim(request("PBillSN"))&" order by PayDate desc"
	set rsPayHis=conn.execute(strPayHis)
	If Not rsPayHis.Bof Then rsPayHis.MoveFirst 
	While Not rsPayHis.Eof
%>		<tr bgcolor="#FFFFFF">
			<td><%=gInitDT(trim(rsPayHis("PayDate")))%></td>
			<td><%
			if trim(rsPayHis("PayTypeID"))="1" then
				response.write "窗口"
			else
				response.write "郵撥"
			end if
			%></td>
			<td><%
			response.write trim(rsPayHis("PayAmount"))
			%></td>
			<td><%
			If Not ifnull(rsPayHis("CaseCloseDate")) Then response.write gInitDT(trim(rsPayHis("CaseCloseDate")))
			%></td>
			<td><%
			if trim(rsPayHis("IsLate"))="1" then
				response.write "逾期繳納"
			else
				response.write "如期繳納"
			end if
			%></td>
			<td><%
			response.write trim(rsPayHis("PayNo"))
			%></td>
			<td><%
			response.write trim(rsPayHis("MIDDLEMONEY"))
			%></td>
			<td><%
			if trim(rsPayHis("Note"))<>"" and not isnull(rsPayHis("Note")) then
				response.write trim(rsPayHis("Note"))
			end if
			%></td>
			<td nowrap>
			<%if sys_City = "彰化縣" Then%>
				<input type="button" name="PrintImg" value="列印收據" 
				onclick="funPrintImage('<%=trim(rsPayHis("BillSN"))%>','<%=trim(rsPayHis("PayTimes"))%>');">
			<%end if%>
			<%if trim(rsPayHis("RecordMemberID"))=trim(Session("User_ID")) or trim(Session("Credit_ID"))="A000000000" then%>
			<input type="button" name="<%=gInitDT(trim(rsPayHis("PayDate")))%>" value="修改" onclick="location='Passer_Pay_Update.asp?PBillSN=<%=trim(rsPayHis("BillSN"))%>&PTime=<%=trim(rsPayHis("PayTimes"))%>'">

			<input type="button" name="<%=gInitDT(trim(rsPayHis("PayDate")))%>" value="刪除" onclick="funDel('<%=trim(rsPayHis("BillSN"))%>','<%=trim(rsPayHis("PayTimes"))%>');">
			<%end if%>&nbsp;
			</td>
		</tr>
<%	rsPayHis.MoveNext
	Wend
	rsPayHis.close
	set rsPayHis=nothing
	%>
	</table>
	</td>
  </tr>
</table>
<%
rsSql.close
set rsSql=nothing
%>
<input type="Hidden" name="BillTime" value="">
<input type="Hidden" name="totalMoney" value="<%=totalMoney%>">
<input type="Hidden" name="PBillSN" value="<%=trim(request("PBillSN"))%>">
<input type="Hidden" name="PayTimes" value="">
</form>

<form name="updFrom" method="post">	
	<input type="Hidden" name="kinds" value="">
	<input type="Hidden" name="PBillSN" value="">
	<input type="Hidden" name="IsLate" value="">
	<input type="Hidden" name="L1ForFeit" value="<%=L1ForFeit%>">
	<input type="Hidden" name="L2ForFeit" value="<%=L2ForFeit%>">
	<input type="Hidden" name="Rule2_L1ForFeit" value="<%=Rule2_L1ForFeit%>">
	<input type="Hidden" name="Rule2_L2ForFeit" value="<%=Rule2_L2ForFeit%>">
</form>
</body>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
var TDStationErrorLog=0;

function funPrintImage(PBillSN,PayTimes){

	myForm.PBillSN.value=PBillSN;
	myForm.PayTimes.value=PayTimes;

	UrlStr="Passer_Pay_chromat_1060921.asp";
	myForm.action=UrlStr;
	myForm.target="PrintImage";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function getStation(){

	myForm.MemberStation.value=myForm.MemberStation.value.replace(/[^\d]/g,'');

	if (myForm.MemberStation.value.length > 1){
		var StationNum=myForm.MemberStation.value;
		runServerScript("../BillKeyIn/getMemberStation.asp?StationID="+StationNum);
	}else{
		Layer5.innerHTML=" ";
		TDStationErrorLog=1;
	}
}

function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}
}

function funChkMoney(Obj,Sys_PaySum){
	Obj.value=Obj.value.replace(/[^\d]/g,'');

	myForm.CaseClose.checked=true;
	myForm.CaseClose.disabled=false;

	if(eval(Obj.value) > 0){

		if(eval(Obj.value)+eval(Sys_PaySum) < eval(myForm.ForFeit.value)){

			myForm.CaseClose.checked=false;
			myForm.CaseClose.disabled=true;

		}else if(eval(Obj.value)+eval(Sys_PaySum) > eval(myForm.ForFeit.value)){

			myForm.CaseClose.checked=false;
			myForm.CaseClose.disabled=true;
		}else{

			myForm.CaseClose.checked=true;
			myForm.CaseClose.disabled=false;
		}
	}
}

function db_insert(){
	var errorString="";error=0;
	/*
	if (myForm.PayAmount.value==""||myForm.PayNo.value==""){
		if(myForm.CaseClose.checked){
			if(!confirm("是否確定要結案?")){
				error=error+1;
				errorString=error+"：請輸入繳費金額及收據字號";
			}
		}else{
			if(!confirm("是否確定不要結案?")){
				error=error+1;
				errorString=error+"：請輸入繳費金額及收據字號";
			}
		}
	}
	*/
	if(myForm.CaseClose.checked){
		if(myForm.Note.value==""){
			if(myForm.Note.value==""){
				if (myForm.PayAmount.value==""||myForm.PayAmount.value=="0"){
					error=error+1;
					errorString=error+"：請輸入繳費金額，或輸入無法處理原因。";
				}
				if (myForm.PayNo.value==""){
					error=error+1;
					errorString=errorString+"\n"+error+"：請輸入收據字號，或輸入無法處理原因。";
				}
			}

			
			if (myForm.CaseCloseDate.value==""){
				error=error+1;
				errorString=errorString+"\n"+error+"：請輸入結案日期。";
			}
		}
	}

	if (myForm.PayAmount.value!=""&&myForm.PayNo.value!=""){
		if(myForm.PayTypeID[0].checked){
			if(myForm.PayDate.value!=""){
				if(!dateCheck(myForm.PayDate.value)){
					error=error+1;
					errorString=errorString+"\n"+error+"：執行劃撥日期輸入錯誤。";
				}
			}else{
				error=error+1;
				errorString=errorString+"\n"+error+"：執行劃撥日期未輸入。";
			}
		}
	}
	if (myForm.PayDate.value!=""){
		if(!dateCheck(myForm.PayDate.value)){
			error=error+1;
			errorString=errorString+"\n"+error+"：繳費日期輸入錯誤。";
		}
	}
	if (myForm.CaseCloseDate.value!=""){
		if(!dateCheck(myForm.CaseCloseDate.value)){
			error=error+1;
			errorString=errorString+"\n"+error+"：結案日期輸入錯誤。";
		}
	}
	if (error==0){
		myForm.subPayCLose.disabled=true;
		myForm.kinds.value="db_insert";
		myForm.submit();
	}else{
		alert(errorString);
	}

}
function funDel(BillSN,BillTime){
	myForm.BillSN.value=BillSN;
	myForm.BillTime.value=BillTime;
	myForm.kinds.value="Del";
	myForm.submit();
}

function funUpMoney(BillSN){
	updFrom.PBillSN.value=BillSN;
	updFrom.IsLate.value=myForm.IsLate.value;
	updFrom.kinds.value="UpMoney";
	updFrom.submit();
}

function funExit(){
	opener.myForm.submit(); 
	self.close();
}
</script>
</html>
