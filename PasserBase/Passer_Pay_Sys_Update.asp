<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>繳款記錄</title>
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
'response.write request("PBillSN")
'檢查是否可進入本系統
AuthorityCheck(224)
memID=Session("User_ID")

if trim(request("kinds"))="Del" then

	strSQL="insert into PASSERPAYDEL select BILLSN,BILLNO,PAYNO,PAYTIMES,PAYTYPEID,PAYDATE,PAYER,FORFEIT,PAYAMOUNT,CASECLOSE,RECORDSTATEID,RECORDDATE,RECORDMEMBERID,sysdate DELDATE,"&memID&" DELMEMBERID,NOTE,ISLATE,MIDDLEMONEY,CASECLOSEDATE,'"&Request("DelNote")&"' DELNOTE,CREDITORSENDNO from passerpay where BillSN="&trim(request("BillSN"))&" and PayTimes="&trim(request("BillTime"))&" and PAYNO is not null"
	conn.execute(strSQL)

	strSQL="Delete from PasserPay where BillSN="&trim(request("BillSN"))&" and PayTimes="&trim(request("BillTime"))
	conn.execute(strSQL)
	
	strSQL="update passerbase set billstatus=0 where (select count(1) cnt from passerpay where billsn="&trim(request("BillSN"))&")=0 and sn="&trim(request("BillSN"))
	conn.execute(strSQL)
	%>
	<script language="JavaScript">
		alert("刪除完成");
	</script>
	<%
end if

if trim(request("kinds"))="db_Update" then
	'繳費日期
	'if trim(request("PayTypeID"))="1" then
		'thePayDate=date
	'else
		thePayDate=gOutDT(request("PayDate"))
	'end if
	'結案
	if trim(request("CaseClose"))="1" then
		theCaseClose=1
		strUpd="Update PasserBase set BillStatus='9' where SN="&trim(request("BillSN"))
		conn.execute strUpd
	else
		strUpd="Update PasserBase set BillStatus='0' where SN="&trim(request("BillSN"))
		conn.execute strUpd
		theCaseClose=0
	end if
	Sys_PayMIDDLEMONEY=0
	if trim(request("PayMIDDLEMONEY"))<>"" then Sys_PayMIDDLEMONEY=trim(request("PayMIDDLEMONEY"))
	Sys_PayAmount=0
	if trim(request("PayAmount"))<>"" then Sys_PayAmount=trim(request("PayAmount"))
	strUpd="update PasserPay set PayTypeID="&trim(request("PayTypeID")) &_
		",PayAmount="&Sys_PayAmount&",IsLate='"&trim(request("IsLate"))&"'" &_
		",PayDate=TO_DATE('"&thePayDate&"','YYYY/MM/DD')" &_
		",Note='"&trim(request("Note"))&"'" &_
		",CaseClose='"&theCaseClose&"',MIDDLEMONEY="&Sys_PayMIDDLEMONEY &_
		",CreditorSendNo='"&trim(request("Sys_CreditorSendNo"))&"'" &_
		" where BillSN="&trim(request("BillSN"))&" and PayTimes="&trim(request("PTime"))
	conn.execute strUpd

	strUpdate="Update PasserPay set ForFeit="&trim(request("ForFeit"))&" where BillSN="&trim(request("BillSN"))&" and PayTimes="&trim(request("PTime"))
	conn.execute(strUpdate)

	strUpd="Update PasserBase set ForFeit1="&trim(request("ForFeit1"))&" where SN="&trim(request("BillSN"))
	conn.execute strUpd
	
	If not ifnull(trim(request("ForFeit2"))) Then

		strUpd="Update PasserBase set ForFeit2="&trim(request("ForFeit2"))&" where SN="&trim(request("BillSN"))
		conn.execute strUpd

	End if 

	if theCaseClose=1 then
		strIns="Update PasserPay set CaseCloseDate=TO_DATE('"&gOutDT(request("CaseCloseDate"))&"','YYYY/MM/DD') where BillSN="&trim(request("BillSN"))&" and PayTimes="&trim(request("PTime"))
		conn.execute strIns
	else
		strIns="Update PasserPay set CaseCloseDate=null where BillSN="&trim(request("BillSN"))&" and PayTimes="&trim(request("PTime"))
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

	strSql="select * from PasserBase where SN="&trim(request("PBillSN"))
	set rsSql=conn.execute(strSql)

	strSql2="select * from PasserPay where BillSN="&trim(request("PBillSN"))&" and PayTimes="&trim(request("PTime"))
	set rs2=conn.execute(strSql2)
%>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">
<form name=myForm method="post">
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="style5">繳款記錄</span></td>
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
		if trim(rsSql("Rule2"))<>"" and not isnull(rsSql("Rule2")) then
			response.write "<br>"&trim(rsSql("Rule2"))&"，"
			strRule1="select * from Law where ItemID='"&trim(rsSql("Rule2"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
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

		theForFeit1="":theForFeit2=""
			
		if sys_City = "宜蘭縣" Then

			if datediff("d",trim(rsSql("DealLineDate")),now)=<0 then

				theForFeit1=L1ForFeit:theForFeit2=Rule2_L1ForFeit

			elseif datediff("d",trim(rsSql("DealLineDate")),now)>0 and datediff("d",trim(rsSql("DealLineDate")),now)<=15 then

				theForFeit1=L2ForFeit:theForFeit2=Rule2_L2ForFeit

			elseif datediff("d",trim(rsSql("DealLineDate")),now)>15 and datediff("d",trim(rsSql("DealLineDate")),now)<=30 then

				theForFeit1=L3ForFeit:theForFeit2=Rule2_L3ForFeit

			elseif datediff("d",trim(rsSql("DealLineDate")),now)>30 then

				theForFeit1=L4ForFeit:theForFeit2=Rule2_L4ForFeit

			end if

		else

			theForFeit1=cdbl(rsSql("ForFeit1"))
			if trim(rsSql("Rule2"))<>"" and not isnull(rsSql("Rule2")) then theForFeit2=cdbl(rsSql("ForFeit2"))
		end if
		%>
		法條1<input type="text" name="ForFeit1" value="<%=theForFeit1%>"  onkeyup="<%
			Response.Write "if(!myForm.ForFeit2.value){"
			Response.Write "myForm.ForFeit.value=eval(myForm.ForFeit1.value);"
			Response.Write "}else{"
			Response.Write "myForm.ForFeit.value=eval(myForm.ForFeit1.value)+eval(myForm.ForFeit2.value);"
			Response.Write "}"
		%>"><br>
		法條2<input type="text" name="ForFeit2" value="<%=theForFeit2%>"  onkeyup="<%
			Response.Write "if(!myForm.ForFeit2.value){"
			Response.Write "myForm.ForFeit.value=eval(myForm.ForFeit1.value);"
			Response.Write "}else{"
			Response.Write "myForm.ForFeit.value=eval(myForm.ForFeit1.value)+eval(myForm.ForFeit2.value);"
			Response.Write "}"
		%>"><br>
		合計:<input type="text" name="ForFeit" Readonly value="<%=trim(rs2("ForFeit"))%>">
		</td>
      </tr>
      <tr>
        <td nowrap bgcolor="#FFFF99"><div align="right">繳費方式</div></td>
        <td>
			<input class="btn1" type="radio" value="2" name="PayTypeID" <%
			if trim(rs2("PayTypeID"))="2" then
				response.write "checked"
			end if
			%>>郵撥
			<input class="btn1" type="radio" value="1" name="PayTypeID" <%
			if trim(rs2("PayTypeID"))="1" then
				response.write "checked"
			end if
			%>>窗口
			<input class="btn1" type="radio" value="3" name="PayTypeID" <%
			if trim(rs2("PayTypeID"))="3" then
				response.write "checked"
			end if
			%>>其他代收單位
		</td>
        <td width="9%" nowrap bgcolor="#FFFF99"><div align="right" class="style3">手續費</div></td>
        <td>
			<input type="text" name="PayMIDDLEMONEY" value="<%=trim(rs2("MIDDLEMONEY"))%>">
		</td>
      </tr>

      <tr>
        <td nowrap bgcolor="#FFFF99"><div align="right"><span class="style3">逾期與否</span></div></td>
        <td width="19%">
			<select name="IsLate">
				<option value="0" <%if trim(rs2("IsLate"))="0" then response.write"selected"%>>如期繳納</option>
				<option value="1" <%if trim(rs2("IsLate"))="1" then response.write"selected"%>>逾期繳納</option>
			</select>
		</td>

		<td width="9%" nowrap bgcolor="#FFFF99"><div align="right" class="style3">強制執行<br>移送案號</div></td>
        <td>
			<input type="text" name="Sys_CreditorSendNo" value="<%=trim(rs2("CreditorSendNo"))%>">
		</td>
      </tr>
      <tr>
        <td nowrap bgcolor="#FFFF99"><div align="right"><span class="style3">繳費日期</span></div></td>
        <td width="19%">
		<input class="btn1" type="text" name="PayDate" size="6" maxlength="8" value="<%=gInitDT(rs2("PayDate"))%>" onkeyup="value=value.replace(/[^\d]/g,'')">
		<input type="button" name="datestr" value="..." onclick="OpenWindow('PayDate');">
		</td>
		
        <td width="9%" nowrap bgcolor="#FFFF99"><div align="right" class="style3">繳費金額</div></td>
        <td>
			<input class="btn1" type="text" name="PayAmount" value="<%
				response.write trim(rs2("PayAmount"))
			%>">
		</td>
      </tr>
      <tr>
		<td nowrap bgcolor="#FFFF99"><div align="right" class="style3">結案日期/狀態</div></td>
        <td nowrap bgcolor="#FFFFff"><div>
			<input type="text" name="CaseCloseDate" size="6" maxlength="8" value="<%=gInitDT(trim(rs2("CaseCloseDate")))%>" onkeyup="value=value.replace(/[^\d]/g,'')">
			<input type="button" name="datestr" value="..." onclick="OpenWindow('CaseCloseDate');">
			結案<input class="btn1" type="checkbox" name="CaseClose" value="1" <%
			if trim(rsSql("BillStatus"))="9" then
				response.write "checked"
			end if
			%>>
		</div></td>
		
        <td nowrap bgcolor="#FFFF99"><div align="right"><span class="style3">收據字號</span></div></td>
        <td>
          <input name="PayNo" class="btn1" type="text" value="<%
		  if trim(rs2("PayNo"))<>"" and not isnull(rs2("PayNo")) then
			response.write trim(rs2("PayNo"))
		  end if
		  %>" size="31" maxlength="50" readonly>
        </td>
      </tr>
	  <tr>
		<td bgcolor="#FFFF99"><div align="right" class="style3">作癈已結案件備註</div></td>
        <td nowrap bgcolor="#FFFFff">
			<input type="text" name="DelNote" size="31" maxlength="30" value="">
			
		</td>

        <td nowrap bgcolor="#FFFF99"><div align="right" class="style3">無法處理原因</div></td>
        <td><span class="style3">
          <input name="Note" class="btn1" type="text" size="31" maxlength="30" value="<%
		  if trim(rs2("Note"))<>"" and not isnull(rs2("Note")) then
			response.write trim(rs2("Note"))
		  end if
		  %>">
        </span></td>
      </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77" colspan="4"><p align="center" class="style1">
        <input name="Submit43322" type="button" class="style3" value="儲存" onclick="db_Update();">
		<img src="space.gif" width="20" height="5"> 
       <!--  <input name="Submit433223" type="button" class="style3" value="結案">
        <img src="space.gif" width="20" height="5">  -->
		<input name="Submit433222" type="button" class="style3" value="關閉" onclick="funExit();">
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
			response.write gInitDT(trim(rsPayHis("CaseCloseDate")))
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
			<%if trim(rsPayHis("RecordMemberID"))=trim(Session("User_ID")) then%>
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
rs2.close
set rs2=nothing
%>
<input type="Hidden" name="BillTime" value="">
</form>
</body>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}
}

function db_Update(){
	error=0;
	var errorString="";

	if(myForm.CaseClose.checked){
		if(myForm.Note.value==""){
			if (myForm.PayAmount.value==""){
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
	
	if(myForm.PayDate.value!=""){
		if(!dateCheck(myForm.PayDate.value)){
			error=error+1;
			errorString=errorString+"\n"+error+"：執行劃撥日期輸入錯誤。";
		}
	}else{
		if(myForm.PayTypeID[0].checked){
			error=error+1;
			errorString=errorString+"\n"+error+"：執行劃撥日期未輸入。";
		}
	}

	if (myForm.CaseCloseDate.value!=""){
		if(!dateCheck(myForm.CaseCloseDate.value)){
			error=error+1;
			errorString=errorString+"\n"+error+"：結案日期輸入錯誤。";
		}
	}

	if (error==0){
		myForm.kinds.value="db_Update";
		myForm.submit();
	}else{
		alert(errorString);
	}

}
function db_insert(){
	error=0;
	var errorString="";
	if (myForm.PayAmount.value==""){
		error=error+1;
		errorString=error+"：請輸入繳費金額。";
	}
	if (myForm.PayNo.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入收據字號。";
	}
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
	if (error==0){
		myForm.kinds.value="db_insert";
		myForm.submit();
	}else{
		alert(errorString);
	}

}
function funDel(BillSN,BillTime){
	
	if(myForm.DelNote.value==""){
		alert("作癈結案紀錄須填寫備註！！");
	}else{
		myForm.BillSN.value=BillSN;
		myForm.BillTime.value=BillTime;
		myForm.kinds.value="Del";
		myForm.submit();
	}
}
function funExit(){
	opener.myForm.submit(); 
	self.close();
}
</script>
</html>
