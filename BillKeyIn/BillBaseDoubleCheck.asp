<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!-- #include file="../Common/Bannernoimage.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/css.txt"-->
<title>打驗校對作業</title>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'檢查是否可進入本系統
'AuthorityCheck(236)
'==========POST=========
'單號
if trim(request("billno"))="" then
	theBillno=""
else
	theBillno=trim(request("billno"))
end if
'==========cookie==========
'填單人代碼
theRecordMemberID=trim(Session("User_ID"))
'==========================
	'BillBase
	strBill1="select * from BillBase where BillNo='"&theBillno&"'"
	set rs1=conn.execute(strBill1)
	if rs1.eof then
%>
	<script language="JavaScript">
		alert("此單號尚未建檔！");
		window.close();
	</script>
<%	end if
	
	strBill2="select * from BillBaseTmp where BillNo='"&theBillno&"'"
	set rs2=conn.execute(strBill2)
	if rs1.eof then
%>
	<script language="JavaScript">
		alert("此單號尚未打驗！");
		window.close();
	</script>
<%	end if
%>
<style type="text/css">
<!--
.style1 {font-size: 14px}
.style3 {font-size: 15px}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='760' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4">打驗校對表</td>
			</tr>
			<tr>
				<td bgcolor="#FFFF99" width="15%" align="right">單號</td>
				<td width="35%"><%=theBillno%></td>
				<td bgcolor="#FFFF99" width="15%" align="right">告發類別</td>
				<td width="35%"><%
				strBillType="select Content from DciCode where TypeID=2 and ID='"&trim(rs1("BillTypeID"))&"'"
				set rsBType=conn.execute(strBillType)
				if not rsBType.eof then
					response.write trim(rsBType("Content"))
				end if
				rsBType.close
				set rsBType=nothing
				%></td>
			</tr>
			<tr>
				<td colspan="2" align="center" bgcolor="#FFFFCC">第一次建檔</td>
				<td colspan="2" align="center" bgcolor="#EAFDE1">第二次建檔</td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("CarNo")) and not isnull(rs2("CarNo")) then
				if trim(rs1("CarNo"))<>trim(rs2("CarNo")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">車號</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
					response.write trim(rs1("CarNo"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">車號</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("CarNo"))<>"" and not isnull(rs2("CarNo")) then
					response.write trim(rs2("CarNo"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("Driver")) and not isnull(rs2("Driver")) then
				if trim(rs1("Driver"))<>trim(rs2("Driver")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">違規人姓名</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
					response.write trim(rs1("Driver"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">違規人姓名</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("Driver"))<>"" and not isnull(rs2("Driver")) then
					response.write trim(rs2("Driver"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("DriverBirth")) and not isnull(rs2("DriverBirth")) then
				if trim(rs1("DriverBirth"))<>trim(rs2("DriverBirth")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">違規人出生日</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("DriverBirth"))<>"" and not isnull(rs1("DriverBirth")) then
					response.write gArrDT(trim(rs1("DriverBirth")))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">違規人出生日</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("DriverBirth"))<>"" and not isnull(rs2("DriverBirth")) then
					response.write gArrDT(trim(rs2("DriverBirth")))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("DriverSex")) and not isnull(rs2("DriverSex")) then
				if trim(rs1("DriverSex"))<>trim(rs2("DriverSex")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">違規人性別</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("DriverSex"))<>"" and not isnull(rs1("DriverSex")) then
					if  trim(rs1("DriverSex"))="1" then
						response.write "男"
					else
						response.write "女"
					end if
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">違規人性別</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("DriverSex"))<>"" and not isnull(rs2("DriverSex")) then
					if  trim(rs2("DriverSex"))="1" then
						response.write "男"
					else
						response.write "女"
					end if
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("DriverZip")) and not isnull(rs2("DriverZip")) then
				if trim(rs1("DriverZip"))<>trim(rs2("DriverZip")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">違規人郵遞區號</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("DriverZip"))<>"" and not isnull(rs1("DriverZip")) then
					response.write trim(rs1("DriverZip"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">違規人郵遞區號</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("DriverZip"))<>"" and not isnull(rs2("DriverZip")) then
					response.write trim(rs2("DriverZip"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("DriverAddress")) and not isnull(rs2("DriverAddress")) then
				if trim(rs1("DriverAddress"))<>trim(rs2("DriverAddress")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">違規人地址</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("DriverAddress"))<>"" and not isnull(rs1("DriverAddress")) then
					response.write trim(rs1("DriverAddress"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">違規人地址</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("DriverAddress"))<>"" and not isnull(rs2("DriverAddress")) then
					response.write trim(rs2("DriverAddress"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("CarSimpleID")) and not isnull(rs2("CarSimpleID")) then
				if trim(rs1("CarSimpleID"))<>trim(rs2("CarSimpleID")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">簡式車種</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
					if trim(rs1("CarSimpleID"))="1" then
						response.write "汽車"
					elseif trim(rs1("CarSimpleID"))="2" then
						response.write "拖車"
					elseif trim(rs1("CarSimpleID"))="3" then
						response.write "重機"
					elseif trim(rs1("CarSimpleID"))="4" then
						response.write "輕機"
					end if
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">簡式車種</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("CarSimpleID"))<>"" and not isnull(rs2("CarSimpleID")) then
					if trim(rs2("CarSimpleID"))="1" then
						response.write "汽車"
					elseif trim(rs2("CarSimpleID"))="2" then
						response.write "拖車"
					elseif trim(rs2("CarSimpleID"))="3" then
						response.write "重機"
					elseif trim(rs2("CarSimpleID"))="4" then
						response.write "輕機"
					end if
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("CarAddID")) and not isnull(rs2("CarAddID")) then
				if trim(rs1("CarAddID"))<>trim(rs2("CarAddID")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">簡式車種</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("CarAddID"))<>"" and not isnull(rs1("CarAddID")) then
					if trim(rs1("CarAddID"))="1" then
						response.write "大貨車"
					elseif trim(rs1("CarAddID"))="2" then
						response.write "大客車"
					elseif trim(rs1("CarAddID"))="3" then
						response.write "砂石車"
					elseif trim(rs1("CarAddID"))="4" then
						response.write "土方車"
					elseif trim(rs1("CarAddID"))="5" then
						response.write "動力機"
					elseif trim(rs1("CarAddID"))="6" then
						response.write "貨櫃"
					end if
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">簡式車種</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("CarAddID"))<>"" and not isnull(rs2("CarAddID")) then
					if trim(rs2("CarAddID"))="1" then
						response.write "大貨車"
					elseif trim(rs2("CarAddID"))="2" then
						response.write "大客車"
					elseif trim(rs2("CarAddID"))="3" then
						response.write "砂石車"
					elseif trim(rs2("CarAddID"))="4" then
						response.write "土方車"
					elseif trim(rs2("CarAddID"))="5" then
						response.write "動力機"
					elseif trim(rs2("CarAddID"))="6" then
						response.write "貨櫃"
					end if
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("IllegalDate")) and not isnull(rs2("IllegalDate")) then
				if trim(rs1("IllegalDate"))<>trim(rs2("IllegalDate")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">違規日期時間</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
					response.write gArrDT(trim(rs1("IllegalDate")))&" "
					response.write Right("00"&hour(rs1("IllegalDate")),2)&":"
					response.write Right("00"&minute(rs1("IllegalDate")),2)
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">違規日期時間</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("IllegalDate"))<>"" and not isnull(rs2("IllegalDate")) then
					response.write gArrDT(trim(rs2("IllegalDate")))&" "
					response.write Right("00"&hour(rs2("IllegalDate")),2)&":"
					response.write Right("00"&minute(rs2("IllegalDate")),2)
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("IllegalAddressID")) and not isnull(rs2("IllegalAddressID")) then
				if trim(rs1("IllegalAddressID"))<>trim(rs2("IllegalAddressID")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">違規地點代碼</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("IllegalAddressID"))<>"" and not isnull(rs1("IllegalAddressID")) then
					response.write trim(rs1("IllegalAddressID"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">違規地點代碼</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("IllegalAddressID"))<>"" and not isnull(rs2("IllegalAddressID")) then
					response.write trim(rs2("IllegalAddressID"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("IllegalAddress")) and not isnull(rs2("IllegalAddress")) then
				if trim(rs1("IllegalAddress"))<>trim(rs2("IllegalAddress")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">違規地點</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
					response.write trim(rs1("IllegalAddress"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">違規地點</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("IllegalAddress"))<>"" and not isnull(rs2("IllegalAddress")) then
					response.write trim(rs2("IllegalAddress"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			'先將法條全抓出來再比對
			RuleStr1=""
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				RuleStr1=trim(rs1("Rule1"))
			end if
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				if RuleStr1="" then
					RuleStr1=trim(rs1("Rule2"))
				else
					RuleStr1=RuleStr1&","&trim(rs1("Rule2"))
				end if
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				if RuleStr1="" then
					RuleStr1=trim(rs1("Rule3"))
				else
					RuleStr1=RuleStr1&","&trim(rs1("Rule3"))
				end if
			end if
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
				if RuleStr1="" then
					RuleStr1=trim(rs1("Rule4"))
				else
					RuleStr1=RuleStr1&","&trim(rs1("Rule4"))
				end if
			end if
			RuleStr2=""
			if trim(rs2("Rule1"))<>"" and not isnull(rs2("Rule1")) then
				RuleStr2=trim(rs2("Rule1"))
			end if
			if trim(rs2("Rule2"))<>"" and not isnull(rs2("Rule2")) then
				if RuleStr2="" then
					RuleStr2=trim(rs2("Rule2"))
				else
					RuleStr2=RuleStr2&","&trim(rs2("Rule2"))
				end if
			end if
			if trim(rs2("Rule3"))<>"" and not isnull(rs2("Rule3")) then
				if RuleStr2="" then
					RuleStr2=trim(rs2("Rule3"))
				else
					RuleStr2=RuleStr2&","&trim(rs2("Rule3"))
				end if
			end if
			if trim(rs2("Rule4"))<>"" and not isnull(rs2("Rule4")) then
				if RuleStr2="" then
					RuleStr2=trim(rs2("Rule4"))
				else
					RuleStr2=RuleStr2&","&trim(rs2("Rule4"))
				end if
			end if
			
			if RuleStr1<>"" and RuleStr2<>"" then
				RuleArray1=split(RuleStr1,",")
				RuleArray2=split(RuleStr2,",")
				RuleStrTmp=""
				for RA1=0 to ubound(RuleArray1)
					for RA2=0 to ubound(RuleArray2)
						if RuleArray1(RA1)=RuleArray2(RA2) then
							if RuleStrTmp="" then
								RuleStrTmp=RuleArray2(RA2)
							else
								RuleStrTmp=RuleStrTmp&","&RuleArray2(RA2)
							end if
							exit for
						end if
					next
				next
				if RuleStr1<>RuleStrTmp then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">違規法條</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
					response.write RuleStr1
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">違規法條</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("IllegalAddress"))<>"" and not isnull(rs2("IllegalAddress")) then
					response.write RuleStr2
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("RuleSpeed")) and not isnull(rs2("RuleSpeed")) then
				if trim(rs1("RuleSpeed"))<>trim(rs2("RuleSpeed")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">限速、限重</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
					response.write trim(rs1("RuleSpeed"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">限速、限重</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("RuleSpeed"))<>"" and not isnull(rs2("RuleSpeed")) then
					response.write trim(rs2("RuleSpeed"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("IllegalSpeed")) and not isnull(rs2("IllegalSpeed")) then
				if trim(rs1("IllegalSpeed"))<>trim(rs2("IllegalSpeed")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">車速</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) then
					response.write trim(rs1("IllegalSpeed"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">車速</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("IllegalSpeed"))<>"" and not isnull(rs2("IllegalSpeed")) then
					response.write trim(rs2("IllegalSpeed"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("BillFillDate")) and not isnull(rs2("BillFillDate")) then
				if trim(rs1("BillFillDate"))<>trim(rs2("BillFillDate")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">填單日期</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
					response.write gArrDT(trim(rs1("BillFillDate")))&" "
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">填單日期</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("BillFillDate"))<>"" and not isnull(rs2("BillFillDate")) then
					response.write gArrDT(trim(rs2("BillFillDate")))&" "
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("DealLineDate")) and not isnull(rs2("DealLineDate")) then
				if trim(rs1("DealLineDate"))<>trim(rs2("DealLineDate")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">應到案日期</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
					response.write gArrDT(trim(rs1("DealLineDate")))&" "
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">應到案日期</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("DealLineDate"))<>"" and not isnull(rs2("DealLineDate")) then
					response.write gArrDT(trim(rs2("DealLineDate")))&" "
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("MemberStation")) and not isnull(rs2("MemberStation")) then
				if trim(rs1("MemberStation"))<>trim(rs2("MemberStation")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">應到案處所</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("MemberStation"))<>"" and not isnull(rs1("MemberStation")) then
					strStation="select StationName from Station where StationID='"&trim(rs1("MemberStation"))&"'"
					set rsStaion=conn.execute(strStation)
					if not rsStaion.eof then
						response.write trim(rsStaion("StationName"))
					end if
					rsStaion.close
					set rsStaion=nothing
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">應到案處所</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("MemberStation"))<>"" and not isnull(rs2("MemberStation")) then
					strStation="select StationName from Station where StationID='"&trim(rs2("MemberStation"))&"'"
					set rsStaion=conn.execute(strStation)
					if not rsStaion.eof then
						response.write trim(rsStaion("StationName"))
					end if
					rsStaion.close
					set rsStaion=nothing
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("BillUnitID")) and not isnull(rs2("BillUnitID")) then
				if trim(rs1("BillUnitID"))<>trim(rs2("BillUnitID")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">舉發單位</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
					strStation="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
					set rsStaion=conn.execute(strStation)
					if not rsStaion.eof then
						response.write trim(rsStaion("UnitName"))
					end if
					rsStaion.close
					set rsStaion=nothing
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">舉發單位</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("BillUnitID"))<>"" and not isnull(rs2("BillUnitID")) then
					strStation="select UnitName from UnitInfo where UnitID='"&trim(rs2("BillUnitID"))&"'"
					set rsStaion=conn.execute(strStation)
					if not rsStaion.eof then
						response.write trim(rsStaion("UnitName"))
					end if
					rsStaion.close
					set rsStaion=nothing
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			'先將舉發人全取出
			BillMemID1=""
			if trim(rs1("BillMemID1"))<>"" and not isnull(rs1("BillMemID1")) then
				BillMemID1=trim(rs1("BillMemID1"))
			end if
			if trim(rs1("BillMemID2"))<>"" and not isnull(rs1("BillMemID2")) then
				if BillMemID1="" then
					BillMemID1=trim(rs1("BillMemID2"))
				else
					BillMemID1=BillMemID1&","&trim(rs1("BillMemID2"))
				end if
			end if
			if trim(rs1("BillMemID3"))<>"" and not isnull(rs1("BillMemID3")) then
				if BillMemID1="" then
					BillMemID1=trim(rs1("BillMemID3"))
				else
					BillMemID1=BillMemID1&","&trim(rs1("BillMemID3"))
				end if
			end if
			
			BillMemID2=""
			if trim(rs2("BillMemID1"))<>"" and not isnull(rs2("BillMemID1")) then
				BillMemID2=trim(rs2("BillMemID1"))
			end if
			if trim(rs2("BillMemID2"))<>"" and not isnull(rs2("BillMemID2")) then
				if BillMemID2="" then
					BillMemID2=trim(rs2("BillMemID2"))
				else
					BillMemID2=BillMemID2&","&trim(rs2("BillMemID2"))
				end if
			end if
			if trim(rs2("BillMemID3"))<>"" and not isnull(rs2("BillMemID3")) then
				if BillMemID2="" then
					BillMemID2=trim(rs2("BillMemID3"))
				else
					BillMemID2=BillMemID2&","&trim(rs2("BillMemID3"))
				end if
			end if

			if BillMemID1<>"" and BillMemID2<>"" then
				BillMemIDArray1=split(BillMemID1,",")
				BillMemIDArray2=split(BillMemID2,",")
				BillMemIDStrTmp=""
				for BA1=0 to ubound(BillMemIDArray1)
					for BA2=0 to ubound(BillMemIDArray2)
						if BillMemIDArray1(BA1)=BillMemIDArray2(BA2) then
							if BillMemIDStrTmp="" then
								BillMemIDStrTmp=BillMemIDArray2(BA2)
							else
								BillMemIDStrTmp=BillMemIDStrTmp&","&BillMemIDArray2(BA2)
							end if
							exit for
						end if
					next
				next
				if BillMemID1<>BillMemIDStrTmp then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">舉發人</td>
				<td><%=FontType1%>
				<%
				BillMemName1=""
				if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
					BillMemName1=trim(rs1("BillMem1"))
				end if
				if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
					if BillMemName1="" then
						BillMemName1=trim(rs1("BillMem2"))
					else
						BillMemName1=BillMemName1&","&trim(rs1("BillMem2"))
					end if
				end if
				if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
					if BillMemName1="" then
						BillMemName1=trim(rs1("BillMem3"))
					else
						BillMemName1=BillMemName1&","&trim(rs1("BillMem3"))
					end if
				end if
				if BillMemName1<>"" then
					response.write BillMemName1
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">舉發人</td>
				<td><%=FontType1%>
				<%
				BillMemName2=""
				if trim(rs2("BillMem1"))<>"" and not isnull(rs2("BillMem1")) then
					BillMemName2=trim(rs2("BillMem1"))
				end if
				if trim(rs2("BillMem2"))<>"" and not isnull(rs2("BillMem2")) then
					if BillMemName2="" then
						BillMemName2=trim(rs2("BillMem2"))
					else
						BillMemName2=BillMemName2&","&trim(rs2("BillMem2"))
					end if
				end if
				if trim(rs2("BillMem3"))<>"" and not isnull(rs2("BillMem3")) then
					if BillMemName2="" then
						BillMemName2=trim(rs2("BillMem3"))
					else
						BillMemName2=BillMemName2&","&trim(rs2("BillMem3"))
					end if
				end if
				if BillMemName2<>"" then
					response.write BillMemName2
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			'先將代保管物全取出
			FastenerStr1=""
			strFastener1="select b.Content from BillFastenerDetail a,DciCode b" &_
				" where a.BillSn="&trim(rs1("Sn"))&" and b.TypeID=6 and a.FastenerTypeID=b.ID"
			set rsFa=conn.execute(strFastener1)
			If Not rsFa.Bof Then rsFa.MoveFirst 
			While Not rsFa.Eof
				if FastenerStr1="" then
					FastenerStr1=trim(rsFa("Content"))
				else
					FastenerStr1=FastenerStr1&","&trim(rsFa("Content"))
				end if
			rsFa.MoveNext
			Wend
			rsFa.close
			set rsFa=nothing

			FastenerStr2=""
			strFastener2="select b.Content from BillFastenerDetailTemp a,DciCode b" &_
				" where a.BillSn="&trim(rs2("Sn"))&" and b.TypeID=6 and a.FastenerTypeID=b.ID"
			set rsFa=conn.execute(strFastener2)
			If Not rsFa.Bof Then rsFa.MoveFirst 
			While Not rsFa.Eof
				if FastenerStr2="" then
					FastenerStr2=trim(rsFa("Content"))
				else
					FastenerStr2=FastenerStr2&","&trim(rsFa("Content"))
				end if
			rsFa.MoveNext
			Wend
			rsFa.close
			set rsFa=nothing
			

			if FastenerStr1<>"" and FastenerStr2<>"" then
				FastenerArray1=split(FastenerStr1,",")
				FastenerArray2=split(FastenerStr2,",")
				FastenerStrTmp=""
				for FA1=0 to ubound(FastenerArray1)
					for FA2=0 to ubound(FastenerArray2)
						if FastenerArray1(FA1)=FastenerArray2(FA2) then
							if FastenerStrTmp="" then
								FastenerStrTmp=FastenerArray2(FA2)
							else
								FastenerStrTmp=FastenerStrTmp&","&FastenerArray2(FA2)
							end if
							exit for
						end if
					next
				next
				if FastenerStr1<>FastenerStrTmp then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">代保管物</td>
				<td><%=FontType1%>
				<%
				if FastenerStr1<>"" then
					response.write FastenerStr1
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">代保管物</td>
				<td><%=FontType1%>
				<%
				if FastenerStr2<>"" then
					response.write FastenerStr2
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("ProjectID")) and not isnull(rs2("ProjectID")) then
				if trim(rs1("ProjectID"))<>trim(rs2("ProjectID")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">專案</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("ProjectID"))<>"" and not isnull(rs1("ProjectID")) then
					strProject1="select Name from Project where ProjectID='"&trim(rs1("ProjectID"))&"'"
					set rsPro1=conn.execute(strProject1)
					if not rsPro1.eof then
						response.write trim(rsPro1("Name"))
					end if
					rsPro1.close
					set rsPro1=nothing
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">專案</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("ProjectID"))<>"" and not isnull(rs2("ProjectID")) then
					strProject="select Name from Project where ProjectID='"&trim(rs2("ProjectID"))&"'"
					set rsPro2=conn.execute(strProject)
					if not rsPro2.eof then
						response.write trim(rsPro2("Name"))
					end if
					rsPro2.close
					set rsPro2=nothing
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("Insurance")) and not isnull(rs2("Insurance")) then
				if trim(rs1("Insurance"))<>trim(rs2("Insurance")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">保險證</td>
				<td><%=FontType1%>
				<%'(0:有出示/1:未出示/2:肇事且未出示/3:逾期或未保險/4:肇事且逾期或未保險) 
				if trim(rs1("Insurance"))<>"" and not isnull(rs1("Insurance")) then
					if trim(rs1("Insurance"))="0" then
						response.write "有出示"
					elseif trim(rs1("Insurance"))="1" then
						response.write "未出示"
					elseif trim(rs1("Insurance"))="2" then
						response.write "肇事且未出示"
					elseif trim(rs1("Insurance"))="3" then
						response.write "逾期或未保險"
					elseif trim(rs1("Insurance"))="4" then
						response.write "肇事且逾期或未保險"
					end if
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">保險證</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("Insurance"))<>"" and not isnull(rs2("Insurance")) then
					if trim(rs2("Insurance"))="0" then
						response.write "有出示"
					elseif trim(rs2("Insurance"))="1" then
						response.write "未出示"
					elseif trim(rs2("Insurance"))="2" then
						response.write "肇事且未出示"
					elseif trim(rs2("Insurance"))="3" then
						response.write "逾期或未保險"
					elseif trim(rs2("Insurance"))="4" then
						response.write "肇事且逾期或未保險"
					end if
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("TrafficAccidentNo")) and not isnull(rs2("TrafficAccidentNo")) then
				if trim(rs1("TrafficAccidentNo"))<>trim(rs2("TrafficAccidentNo")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">交通事故案號</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("TrafficAccidentNo"))<>"" and not isnull(rs1("TrafficAccidentNo")) then
					response.write trim(rs1("TrafficAccidentNo"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">交通事故案號</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("TrafficAccidentNo"))<>"" and not isnull(rs2("TrafficAccidentNo")) then
					response.write trim(rs2("TrafficAccidentNo"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("TrafficAccidentType")) and not isnull(rs2("TrafficAccidentType")) then
				if trim(rs1("TrafficAccidentType"))<>trim(rs2("TrafficAccidentType")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">交通事故種類</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("TrafficAccidentType"))<>"" and not isnull(rs1("TrafficAccidentType")) then
					response.write trim(rs1("TrafficAccidentType"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">交通事故種類</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("TrafficAccidentType"))<>"" and not isnull(rs2("TrafficAccidentType")) then
					response.write trim(rs2("TrafficAccidentType"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
			<%
			if not isnull(rs1("Note")) and not isnull(rs2("Note")) then
				if trim(rs1("Note"))<>trim(rs2("Note")) then
					FontType1="<font color='#FF0000'><strong>"
					FontType2="</strong></font>"
				else
					FontType1="<font color='#000000'>"
					FontType2="</font>"
				end if
			else
				FontType1="<font color='#FF0000'><strong>"
				FontType2="</strong></font>"
			end if
			%>
				<td bgcolor="#FFFFCC" align="right">備註</td>
				<td><%=FontType1%>
				<%
				if trim(rs1("Note"))<>"" and not isnull(rs1("Note")) then
					response.write trim(rs1("Note"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
				<td bgcolor="#EAFDE1" align="right">備註</td>
				<td><%=FontType1%>
				<%
				if trim(rs2("Note"))<>"" and not isnull(rs2("Note")) then
					response.write trim(rs2("Note"))
				else
					response.write "&nbsp;"
				end if
				%>
				<%=FontType2%></td>
			</tr>
			<tr>
				<td colspan="4" align="center">
					<input type="button" value="關  閉" name="closeWin" onclick="window.close();">
				</td>
			</tr>
		</table>		
	</form>
<%
rs1.close
set rs1=nothing
rs2.close
set rs2=nothing
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">

</script>
</html>
