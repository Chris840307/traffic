<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!-- #include file="../Common/Bannernodata.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/css.txt"-->
<title>舉發單綜合查詢</title>
<script type="text/javascript" src="../js/Print.js"></script>
<script type="text/javascript" src="../js/date.js"></script>
<%	
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close

	sys_City=replace(sys_City,"台中縣","台中市")
	sys_City=replace(sys_City,"台南縣","台南市")

	strSQLTemp=""
	if trim(request("BillNo"))<>"" then
		strSQLTemp=" where BillNO='"&trim(request("BillNo"))&"'"
	end if
	if trim(request("illFID"))<>"" then
		if strSQLTemp<>"" then
			strSQLTemp=strSQLTemp&" and DriverID='"&trim(request("illFID"))&"'"
		else
			strSQLTemp=" where DriverID='"&trim(request("illFID"))&"'"
		end if
	end if
	if trim(request("illName"))<>"" then
		if strSQLTemp<>"" then
			strSQLTemp=strSQLTemp&" and Driver='"&trim(request("illName"))&"'"
		else
			strSQLTemp=" where Driver='"&trim(request("illName"))&"'"
		end if
	end if
	if trim(request("BillSn"))<>"" then
		if strSQLTemp<>"" then
			strSQLTemp=strSQLTemp&" and SN='"&trim(request("BillSn"))&"'"
		else
			strSQLTemp=" where SN='"&trim(request("BillSn"))&"'"
		end if
	end if
	strSQL="select * from PasserBase"&strSQLTemp
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
	set rs1=conn.execute(strSQL)
	If Not rs1.Bof Then
		rs1.MoveFirst 
	else
%>
<script language="JavaScript">
	alert("查無資料！");
	window.close();
</script>	
<%
	end if
	While Not rs1.Eof

	Sys_Forfeit=cdbl(rs1("Forfeit1"))
	If not ifnull(rs1("Forfeit2")) Then Sys_Forfeit=Sys_Forfeit+cdbl(rs1("Forfeit2"))

	strSQL="select MANAGEMEMBERNAME,SECONDMANAGERNAME from Unitinfo where UnitID in(select Unittypeid from Unitinfo where Unitid='"&trim(rs1("BillUnitID"))&"')"
%>
	<table width='100%' border='1' cellpadding="2">
		<tr bgcolor="#1BF5FF">
			<td colspan="6"><strong>舉發單資料</strong>
				<a href="#<%=trim(rs1("SN"))%>1">
				
			</td>
		</tr>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>1"></a>
				<strong>舉發單基本資料</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>2">行人攤販裁決書</a>‧
				<a href="#<%=trim(rs1("SN"))%>3">行人攤販移送書</a>‧
				<a href="#<%=trim(rs1("SN"))%>4">行人攤販催告書</a>‧
				<a href="#<%=trim(rs1("SN"))%>5">行人攤販繳費記錄</a>‧
				<a href="#<%=trim(rs1("SN"))%>7">行人攤販送達紀錄</a>
			</td>
		</tr>
		<tr>
			<td bgcolor="#EBE5FF" width="13%" align="right"><strong>單號</strong></td>
			<td align="left" width="20%"><%
			'單號
			if trim(rs1("BillNo"))<>"" and not isnull(rs1("BillNo")) then
				response.write trim(rs1("BillNo"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#EBE5FF" width="13%" align="right"><strong>舉發類別</strong></td>
			<td align="left" ><%
			'舉發類別
			if trim(rs1("BillTypeID"))<>"" and not isnull(rs1("BillTypeID")) then
				response.write "慢車 – "
				if trim(rs1("BillTypeID"))="1" then
					response.write "攔停"
				elseif trim(rs1("BillTypeID"))="2" then
					response.write "逕舉"
				end if

			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#EBE5FF" width="13%" align="right"><strong>違規人性別</strong></td>
			<td align="left"><%
			'違規人性別
			if trim(rs1("DriverSex"))<>"" and not isnull(rs1("DriverSex")) then
				if trim(rs1("DriverSex"))="1" then
					response.write "男"
				else
					response.write "女"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td bgcolor="#EBE5FF" width="13%" align="right"><strong>違規人姓名</strong></td>
			<td align="left" width="20%"><%
			'違規人姓名
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write trim(rs1("Driver"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#EBE5FF" align="right"><strong>違規人身份證</strong></td>
			<td align="left" width="20%"><%
			'違規人身分証
			if trim(rs1("DriverID"))<>"" and not isnull(rs1("DriverID")) then
				response.write trim(rs1("DriverID"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#EBE5FF" align="right"><strong>違規人生日</strong></td>
			<td align="left" width="20%"><%
			'違規人生日
			if trim(rs1("DriverBirth"))<>"" and not isnull(rs1("DriverBirth")) then
				response.write gArrDT(trim(rs1("DriverBirth")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>			
			<td align="right" bgcolor="#EBE5FF"><strong>違規人地址</strong></td>
			<td align="left" colspan="3"><%
			'違規人地址
			if trim(rs1("DriverZip"))<>"" and not isnull(rs1("DriverZip")) then
				response.write trim(rs1("DriverZip"))&"&nbsp;"
			end if
			if trim(rs1("DriverAddress"))<>"" and not isnull(rs1("DriverAddress")) then
				response.write trim(rs1("DriverAddress"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>違規日期</strong></td>
			<td align="left"><%
			'違規日期
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gArrDT(trim(rs1("IllegalDate")))&"&nbsp;&nbsp;"
				response.write Right("00"&hour(rs1("IllegalDate")),2)&":"
				response.write Right("00"&minute(rs1("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>違規地點</strong></td>
			<td align="left" colspan="5"><%
			'違規地點
			if trim(rs1("IllegalAddressID"))<>"" and not isnull(rs1("IllegalAddressID")) then
				response.write trim(rs1("IllegalAddressID"))&" "
			end if
			if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
				response.write trim(rs1("IllegalAddress"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>違規法條</strong></td>
			<td align="left" colspan="5"><%
			'違規法條
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				'chRule=rs1("Rule1")
				strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule1"))&"' and Version='"&trim(rs1("RuleVer"))&"' order by CarSimpleID Desc"
				set rsR1=conn.execute(strR1)
				if not rsR1.eof then 
					chRule=trim(rs1("Rule1"))&" "&trim(rsR1("IllegalRule"))
				end if
				rsR1.close
				set rsR1=nothing
			end if
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				'chRule=chRule&"<br>"&rs1("Rule2")
				strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule2"))&"' and Version='"&trim(rs1("RuleVer"))&"' order by CarSimpleID Desc"
				set rsR1=conn.execute(strR1)
				if not rsR1.eof then 
					chRule=chRule&"<br>"&rs1("Rule2")&" "&trim(rsR1("IllegalRule"))
				end if
				rsR1.close
				set rsR1=nothing
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				'chRule=chRule&"<br>"&rs1("Rule3")
				strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule3"))&"' and Version='"&trim(rs1("RuleVer"))&"' order by CarSimpleID Desc"
				set rsR1=conn.execute(strR1)
				if not rsR1.eof then 
					chRule=chRule&"<br>"&rs1("Rule3")&" "&trim(rsR1("IllegalRule"))
				end if
				rsR1.close
				set rsR1=nothing
			end if
			if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then 
				'chRule=chRule&"<br>"&rs1("Rule4")
				strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule4"))&"' and Version='"&trim(rs1("RuleVer"))&"' order by CarSimpleID Desc"
				set rsR1=conn.execute(strR1)
				if not rsR1.eof then 
					chRule=chRule&"<br>"&rs1("Rule4")&" "&trim(rsR1("IllegalRule"))
				end if
				rsR1.close
				set rsR1=nothing
			end if
			response.write chRule
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>車號</strong></td>
			<td align="left"><%
			'車號
			if trim(rs1("CarNO"))<>"" and not isnull(rs1("CarNO")) then
				response.write rs1("CarNO")
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>簡式車種</strong></td>
			<td align="left"><%
			'簡式車種
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				if trim(rs1("CarSimpleID"))="8" then
					response.write "微電車"
				end if 
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>保險證</strong></td>
			<td align="left"><%
			'保險證
			'第三責任險(0:有出示/1:未出示/2:肇事且未出示/3:逾期或未保險/4:肇事且逾期或未保險) *欄停才顯示
			if trim(rs1("Insurance"))<>"" and not isnull(rs1("Insurance")) and rs1("BillTypeID")="1" then
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
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>限速</strong></td>
			<td align="left"><%
			'限速
			if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
				response.write trim(rs1("RuleSpeed"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>車速</strong></td>
			<td align="left" colspan="3"><%
			'車速
			if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) then
				response.write trim(rs1("IllegalSpeed"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			
			
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>填單日期</strong></td>
			<td align="left"><%
			'填單日期
			if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
				response.write gArrDT(trim(rs1("BillFillDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>應到案日期</strong></td>
			<td align="left"><%
			'應到案日期
			if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
				response.write gArrDT(trim(rs1("DealLineDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>應到案處所</strong></td>
			<td align="left"><%
			'應到案處所
			if trim(rs1("MemberStation"))<>"" and not isnull(rs1("MemberStation")) then
				strMStation="select UnitName from UnitInfo where UnitID='"&trim(rs1("MemberStation"))&"'"
				set rsMStation=conn.execute(strMStation)
				if not rsMStation.eof then
					response.write trim(rsMStation("UnitName"))
				end if
				rsMStation.close
				set rsMStation=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>舉發單位</strong></td>
			<td align="left"><%
			'舉發單位
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				strUName="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsUName=conn.execute(strUName)
				if not rsUName.eof then
					response.write trim(rsUName("UnitName"))
				end if
				rsUName.close
				set rsUName=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>舉發人</strong></td>
			<td align="left"><%
			'舉發人
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write trim(rs1("BillMem1"))
			else
				response.write "&nbsp;"
			end if
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write "、"&trim(rs1("BillMem2"))
			end if
			if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
				response.write "、"&trim(rs1("BillMem3"))
			end if
			if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
				response.write "、"&trim(rs1("BillMem4"))
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>代保管物</strong></td>
			<td align="left"><%
			'代保管物
			FastenerDetail=""
			strFas="select Confiscate from PasserConfiscate where BillSN="&trim(rs1("SN"))
			set rsFas=conn.execute(strFas)
			If Not rsFas.Bof Then
				rsFas.MoveFirst 
			else
				response.write "&nbsp;"
			end if
			While Not rsFas.Eof
				if FastenerDetail="" then
					FastenerDetail=trim(rsFas("Confiscate"))
				else
					FastenerDetail=FastenerDetail&"、"&trim(rsFas("Confiscate"))
				end if
			rsFas.MoveNext
			Wend
			rsFas.close
			set rsFas=nothing
				response.write FastenerDetail
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>專案</strong></td>
			<td align="left"><%
			'專案
			if trim(rs1("ProjectID"))<>"" and not isnull(rs1("ProjectID")) then
				strProj="select Name from Project where ProjectID='"&trim(rs1("ProjectID"))&"'"
				set rsProj=conn.execute(strProj)
				if not rsProj.eof then
					response.write trim(rsProj("Name"))
				end if
				rsProj.close
				set rsProj=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>備註</strong></td>
			<td align="left" colspan="3"><%
			'備註
			if trim(rs1("Note"))<>"" and not isnull(rs1("Note")) then
				response.write trim(rs1("Note"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>

			<td align="right" bgcolor="#EBE5FF"><strong>填單人</strong></td>
			<td align="left"><%
			'填單人
			if trim(rs1("BillFiller"))<>"" and not isnull(rs1("BillFiller")) then
				response.write trim(rs1("BillFiller"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>建檔人</strong></td>
			<td align="left"><%
			'建檔人
			if trim(rs1("RecordMemberID"))<>"" and not isnull(rs1("RecordMemberID")) then
				strRMem="select ChName from MemberData where MemberID="&trim(rs1("RecordMemberID"))
				set rsRMem=conn.execute(strRMem)
				if not rsRMem.eof then
					response.write trim(rsRMem("ChName"))
				end if
				rsRMem.close
				set rsRMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>建檔日期</strong></td>
			<td align="left"><%
			'建檔日期
			if trim(rs1("RecordDate"))<>"" and not isnull(rs1("RecordDate")) then
				response.write gArrDT(trim(rs1("RecordDate")))&" "
				response.write Right("00"&hour(rs1("RecordDate")),2)&":"
				response.write Right("00"&minute(rs1("RecordDate")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>是否應聽講習</strong></td>
			<td align="left"><%
			'是否講習
			if trim(rs1("ISLECTURE"))<>"" and not isnull(rs1("ISLECTURE")) then
				if trim(rs1("ISLECTURE"))="0" then
					response.write "否"
				else
					response.write "是"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>法條版本</strong></td>
			<td align="left"><%
			'法條版本
			if trim(rs1("RuleVer"))<>"" and not isnull(rs1("RuleVer")) then
				response.write trim(rs1("RuleVer"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>刪除原因</strong></td>
			<td align="left"><%
			'刪除原因
			strDelRea="select b.Content from PasserDeleteReason a,DciCode b where a.PasserSn="&trim(rs1("SN"))&" and b.TypeID=3 and a.DelReason=b.ID"
			set rsDelRea=conn.execute(strDelRea)
			if not rsDelRea.eof then
				response.write trim(rsDelRea("Content"))
			else
				response.write "&nbsp;"
			end if
			rsDelRea.close
			set rsDelRea=nothing
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>刪除人</strong></td>
			<td align="left" colspan="5"><%
			'刪除人
			if trim(rs1("DelMemberID"))<>"" and not isnull(rs1("DelMemberID")) then
				strDMem="select ChName from MemberData where MemberID="&trim(rs1("DelMemberID"))
				set rsDMem=conn.execute(strDMem)
				if not rsDMem.eof then
					response.write trim(rsDMem("ChName"))
				end if
				rsDMem.close
				set rsDMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
<%'If trim(Session("Credit_ID"))="A000000000" then%>
		<tr bgcolor="#FAFAF5">
			<td align="center" colspan="6"><strong>違規影像</strong></td>
		</tr>
		<tr>	
			<td colspan="6" align="center">&nbsp;
<%
		strImg="select * from PasserIllegalImage where billsn="&trim(rs1("SN"))
		set rsImgKS=conn.execute(strImg)
		if not rsImgKS.eof then

			if trim(rsImgKS("ImageFileNameA"))<>"" then
		%>
				<img src="<%=trim(rsImgKS("IISImagePath"))&trim(rsImgKS("ImageFileNameA"))%>" name="imgB1" width="700" alt="" >

		<%
			end if
			if trim(rsImgKS("ImageFileNameB"))<>"" then
		%>
				<img src="<%=trim(rsImgKS("IISImagePath"))&trim(rsImgKS("ImageFileNameB"))%>" name="imgB2" width="700" alt="" >
			
		<%
			end if
			if trim(rsImgKS("ImageFileNameC"))<>"" then
		%>
				<img src="<%=trim(rsImgKS("IISImagePath"))&trim(rsImgKS("ImageFileNameC"))%>" name="imgB3" width="700" >
			
		<%
			end If
			
		end if
		rsImgKS.close
		set rsImgKS=Nothing
%>
			</td>
		</tr>
<%'end if%>
		<tr bgcolor="#FAFAF5">
			<td align="center" colspan="6"><strong>相關文件掃描檔</strong></td>
		</tr>
		<tr>	
			<td colspan="6">&nbsp;
			<%
			strScan2="select * from BillAttatchImage where BillNo='"&trim(rs1("BillNo"))&"' and Recordstateid=0"
			set rsScan2=conn.execute(strScan2)
			while Not rsScan2.eof
			%>
				<a title="開啟相關文件掃描檔.." href="<%=replace(trim(rsScan2("FileName")),"/img/","/scanimg/")%>" target="_blank" <%lightbarstyle 1 %>><u>開啟相關文件掃描檔</u></a><br>
				<%
			rsScan2.movenext
			wend
			rsScan2.close
			set rsScan2=nothing

			if sys_City="彰化縣" or sys_City="基隆市" then 
				strSQL="select ImageFileName from PasserImage where billsn="&trim(rs1("SN"))
				set rsScan2=conn.execute(strSQL)
				filecnt=0
				while Not rsScan2.eof

					filecnt=filecnt+1

					Response.Write "<a href=""../PasserBase/PasserImage/"&trim(rsScan2("ImageFileName"))&""" target=""_blank"">掃描檔"&filecnt&"</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"

					rsScan2.movenext
				wend
				rsScan2.close
				set rsScan2=nothing
			end if
			%></td>
		</tr>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>2"></a>
				<strong>行人攤販裁決書</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>3">行人攤販移送書</a>‧
				<a href="#<%=trim(rs1("SN"))%>4">行人攤販催告書</a>‧
				<a href="#<%=trim(rs1("SN"))%>5">行人攤販繳費記錄</a>‧
				<a href="#<%=trim(rs1("SN"))%>7">行人攤販送達紀錄</a>
			</td>
		</tr>
<%
	strJude="select * from PasserJude where BillSn="&trim(rs1("SN"))
	set rsJude=conn.execute(strJude)
	if not rsJude.eof then
%>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>單號</strong></td>
			<td align="left"><%
			'單號
			if trim(rsJude("BILLNO"))<>"" and not isnull(rsJude("BILLNO")) then
				response.write trim(rsJude("BILLNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>發文字號</strong></td>
			<td align="left" colspan="3"><%
			'發文字號
			if trim(rsJude("OPENGOVNUMBER"))<>"" and not isnull(rsJude("OPENGOVNUMBER")) then
				response.write trim(rsJude("OPENGOVNUMBER"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>裁決日期</strong></td>
			<td align="left"><%
			'裁決日期
			if trim(rsJude("JUDEDATE"))<>"" and not isnull(rsJude("JUDEDATE")) then
				response.write gArrDT(trim(rsJude("JUDEDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>應到案處所</strong></td>
			<td align="left"><%
			'應到案處所
			if trim(rsJude("DUTYUNIT"))<>"" and not isnull(rsJude("DUTYUNIT")) then
				response.write trim(rsJude("DUTYUNIT"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>罰款金額</strong></td>
			<td align="left"><%
			'罰款金額
			if trim(rsJude("FORFEIT"))<>"" and not isnull(rsJude("FORFEIT")) then
				response.write trim(rsJude("FORFEIT"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>處罰主文</strong></td>
			<td align="left" colspan="5"><%
			'處罰主文
			if trim(rsJude("PUNISHMENTMAINBODY"))<>"" and not isnull(rsJude("PUNISHMENTMAINBODY")) then
				response.write trim(rsJude("PUNISHMENTMAINBODY"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td bgcolor="#EBE5FF" align="right"><strong>簡要理由</strong></td>
			<td align="left" colspan="5"><%
			'簡要理由
			if trim(rsJude("SIMPLERESON"))<>"" and not isnull(rsJude("SIMPLERESON")) then
				response.write trim(rsJude("SIMPLERESON"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>局長</strong></td>
			<td align="left"><%
			'局長
			if trim(rsJude("BIGUNITBOSSNAME"))<>"" and not isnull(rsJude("BIGUNITBOSSNAME")) then
				response.write trim(rsJude("BIGUNITBOSSNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>分局長</strong></td>
			<td align="left"><%
			'分局長
			if trim(rsJude("SUBUNITSECBOSSNAME"))<>"" and not isnull(rsJude("SUBUNITSECBOSSNAME")) then
				response.write trim(rsJude("SUBUNITSECBOSSNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>聯絡電話</strong></td>
			<td align="left"><%
			'聯絡電話
			if trim(rsJude("CONTACTTEL"))<>"" and not isnull(rsJude("CONTACTTEL")) then
				response.write trim(rsJude("CONTACTTEL"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>法定代理人姓名</strong></td>
			<td align="left"><%
			'法定代理人姓名
			if trim(rsJude("AGENTNAME"))<>"" and not isnull(rsJude("AGENTNAME")) then
				response.write trim(rsJude("AGENTNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>法定代理人生日</strong></td>
			<td align="left"><%
			'法定代理人生日
			if trim(rsJude("AGENTBIRTH"))<>"" and not isnull(rsJude("AGENTBIRTH")) then
				response.write gArrDT(trim(rsJude("AGENTBIRTH")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>法定代理人身分證字號</strong></td>
			<td align="left"><%
			'法定代理人身分證字號
			if trim(rsJude("AGENTID"))<>"" and not isnull(rsJude("AGENTID")) then
				response.write trim(rsJude("AGENTID"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>法定代理人性別</strong></td>
			<td align="left"><%
			'法定代理人性別
			if trim(rsJude("AGENTSEX"))<>"" and not isnull(rsJude("AGENTSEX")) then
				if trim(rsJude("AGENTSEX"))="0" then
					response.write "女"
				elseif trim(rsJude("AGENTSEX"))="1" then
					response.write "男"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>法定代理人住址</strong></td>
			<td align="left" colspan="3"><%
			'法定代理人住址
			if trim(rsJude("AGENTADDRESS"))<>"" and not isnull(rsJude("AGENTADDRESS")) then
				response.write trim(rsJude("AGENTADDRESS"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>紀錄狀態</strong></td>
			<td align="left"><%
			'紀錄狀態
			if trim(rsJude("RECORDSTATEID"))<>"" and not isnull(rsJude("RECORDSTATEID")) then
				if trim(rsJude("RECORDSTATEID"))="0" then
					response.write "正常"
				else
					response.write "刪除"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>紀錄時間</strong></td>
			<td align="left"><%
			'紀錄時間
			if trim(rsJude("RECORDDATE"))<>"" and not isnull(rsJude("RECORDDATE")) then
				response.write gArrDT(trim(rsJude("RECORDDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>紀錄人員</strong></td>
			<td align="left"><%
			'紀錄人員
			if trim(rsJude("RECORDMEMBERID"))<>"" and not isnull(rsJude("RECORDMEMBERID")) then
				strRecordMem="select chName from MemberData where MemberId="&trim(rsJude("RECORDMEMBERID"))
				set rsRecMem=conn.execute(strRecordMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>送信地址</strong></td>
			<td align="left" colspan="3"><%
			'送信地址
			if trim(rsJude("SENDADDRESS"))<>"" and not isnull(rsJude("SENDADDRESS")) then
				response.write trim(rsJude("SENDADDRESS"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>刪除人員</strong></td>
			<td align="left"><%
			'刪除人員
			if trim(rsJude("DELMEMBERID"))<>"" and not isnull(rsJude("DELMEMBERID")) then
				strRecordMem="select chName from MemberData where MemberId="&trim(rsJude("DELMEMBERID"))
				set rsRecMem=conn.execute(strRecordMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>備註</strong></td>
			<td align="left" colspan="5"><%
			'備註
			if trim(rsJude("NOTE"))<>"" and not isnull(rsJude("NOTE")) then
				response.write trim(rsJude("NOTE"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
<%	end if
	rsJude.close
	set rsJude=nothing
%>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>3"></a>
				<strong>行人攤販移送書</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>2">行人攤販裁決書</a>‧
				<a href="#<%=trim(rs1("SN"))%>4">行人攤販催告書</a>‧
				<a href="#<%=trim(rs1("SN"))%>5">行人攤販繳費記錄</a>‧
				<a href="#<%=trim(rs1("SN"))%>7">行人攤販送達紀錄</a>
			</td>
		</tr>
<%
	strPasserSend="select * from PasserSend where BillSn="&trim(rs1("SN"))
	set rsSend=conn.execute(strPasserSend)
	if not rsSend.eof then
%>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>單號</strong></td>
			<td align="left"><%
			'單號
			if trim(rsSend("BILLNO"))<>"" and not isnull(rsSend("BILLNO")) then
				response.write trim(rsSend("BILLNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>發文文號</strong></td>
			<td align="left"><%
			'發文文號
			if trim(rsSend("OPENGOVNUMBER"))<>"" and not isnull(rsSend("OPENGOVNUMBER")) then
				response.write trim(rsSend("OPENGOVNUMBER"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>移送字號</strong></td>
			<td align="left"><%
			'移送字號
			if trim(rsSend("SENDNUMBER"))<>"" and not isnull(rsSend("SENDNUMBER")) then
				response.write trim(rsSend("SENDNUMBER"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>移送日期</strong></td>
			<td align="left"><%
			'移送日期
			if trim(rsSend("SENDDATE"))<>"" and not isnull(rsSend("SENDDATE")) then
				response.write gArrDT(trim(rsSend("SENDDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>法定代理人</strong></td>
			<td align="left"><%
			'法定代理人
			if trim(rsSend("AGENT"))<>"" and not isnull(rsSend("AGENT")) then
				response.write trim(rsSend("AGENT"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>法定代理人生日</strong></td>
			<td align="left"><%
			'法定代理人生日
			if trim(rsSend("AGENTBIRTHDATE"))<>"" and not isnull(rsSend("AGENTBIRTHDATE")) then
				response.write gArrDT(trim(rsSend("AGENTBIRTHDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>法定代理人證號</strong></td>
			<td align="left"><%
			'法定代理人證號
			if trim(rsSend("AGENTID"))<>"" and not isnull(rsSend("AGENTID")) then
				response.write trim(rsSend("AGENTID"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>法定代理人住址</strong></td>
			<td align="left" colspan="3"><%
			'法定代理人住址
			if trim(rsSend("AGENTADDRESS"))<>"" and not isnull(rsSend("AGENTADDRESS")) then
				response.write trim(rsSend("AGENTADDRESS"))
			else
				response.write "&nbsp;"
			end if
			%></td>

		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>罰款金額</strong></td>
			<td align="left"><%
			'罰款金額
			if trim(Sys_Forfeit)<>"" then
				response.write Sys_Forfeit
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>局長</strong></td>
			<td align="left"><%
			'局長
			if trim(rsSend("BIGUNITBOSSNAME"))<>"" and not isnull(rsSend("BIGUNITBOSSNAME")) then
				response.write trim(rsSend("BIGUNITBOSSNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>分局長</strong></td>
			<td align="left"><%
			'分局長
			if trim(rsSend("SUBUNITSECBOSSNAME"))<>"" and not isnull(rsSend("SUBUNITSECBOSSNAME")) then
				response.write trim(rsSend("SUBUNITSECBOSSNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>確定日期</strong></td>
			<td align="left"><%
			'確定日期
			if trim(rsSend("MAKESUREDATE"))<>"" and not isnull(rsSend("MAKESUREDATE")) then
				response.write gArrDT(trim(rsSend("MAKESUREDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>限繳日期</strong></td>
			<td align="left"><%
			'現繳日期
			if trim(rsSend("LIMITDATE"))<>"" and not isnull(rsSend("LIMITDATE")) then
				response.write gArrDT(trim(rsSend("LIMITDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>執行處回文日期</strong></td>
			<td align="left"><%
			'執行處回文日期
			if trim(rsSend("EXECUTERETURNDATE"))<>"" and not isnull(rsSend("EXECUTERETURNDATE")) then
				response.write gArrDT(trim(rsSend("EXECUTERETURNDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>執行處回文文號</strong></td>
			<td align="left"><%
			'執行處回文文號
			if trim(rsSend("EXECUTERETURNNUMBER"))<>"" and not isnull(rsSend("EXECUTERETURNNUMBER")) then
				response.write trim(rsSend("EXECUTERETURNNUMBER"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>執行處回文登記</strong></td>
			<td align="left"><%
			'執行處回文登記
			if trim(rsSend("EXECUTERETURNNOTE"))<>"" and not isnull(rsSend("EXECUTERETURNNOTE")) then
				response.write trim(rsSend("EXECUTERETURNNOTE"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>紀錄狀態</strong></td>
			<td align="left"><%
			'紀錄狀態
			if trim(rsSend("RECORDSTATEID"))<>"" and not isnull(rsSend("RECORDSTATEID")) then
				if trim(rsSend("RECORDSTATEID"))="0" then
					response.write "正常"
				elseif trim(rsSend("RECORDSTATEID"))="-1" then
					response.write "刪除"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>紀錄時間</strong></td>
			<td align="left"><%
			'紀錄時間
			if trim(rsSend("RECORDDATE"))<>"" and not isnull(rsSend("RECORDDATE")) then
				response.write gArrDT(trim(rsSend("RECORDDATE")))&" "
				response.write Right("00"&hour(rsSend("RECORDDATE")),2)&":"
				response.write Right("00"&minute(rsSend("RECORDDATE")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>紀錄人員</strong></td>
			<td align="left"><%
			'紀錄人員
			if trim(rsSend("RECORDMEMBERID"))<>"" and not isnull(rsSend("RECORDMEMBERID")) then
				strSRecMem="select chName from MemberData where MemberID="&trim(rsSend("RECORDMEMBERID"))
				set rsSRmem=conn.execute(strSRecMem)
				if not rsSRmem.eof then
					response.write trim(rsSRmem("chName"))
				end if
				rsSRmem.close
				set rsSRmem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>刪除人員</strong></td>
			<td align="left"><%
			'刪除人員
			if trim(rsSend("DELMEMBERID"))<>"" and not isnull(rsSend("DELMEMBERID")) then
				strSRecMem="select chName from MemberData where MemberID="&trim(rsSend("DELMEMBERID"))
				set rsSRmem=conn.execute(strSRecMem)
				if not rsSRmem.eof then
					response.write trim(rsSRmem("chName"))
				end if
				rsSRmem.close
				set rsSRmem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>保全措施</strong></td>
			<td align="left" colspan="5"><%
			'保全措施
			SafeWay=""
			if trim(rsSend("SAFETOEXIT"))<>"" and not isnull(rsSend("SAFETOEXIT")) then
				if trim(rsSend("SAFETOEXIT"))="1" then
					SafeWay="已限制出境"
				end if
			end if
			if trim(rsSend("SAFEACTION"))<>"" and not isnull(rsSend("SAFEACTION")) then
				if trim(rsSend("SAFEACTION"))="1" then
					if SafeWay="" then
						SafeWay="已禁止處分"
					else
						SafeWay=SafeWay&"、已禁止處分"
					end if
				end if
			end if
			if trim(rsSend("SAFEASSURE"))<>"" and not isnull(rsSend("SAFEASSURE")) then
				if trim(rsSend("SAFEASSURE"))="1" then
					if SafeWay="" then
						SafeWay="已提供擔保"
					else
						SafeWay=SafeWay&"、已提供擔保"
					end if
				end if
			end if
			if trim(rsSend("SAFEDETAIN"))<>"" and not isnull(rsSend("SAFEDETAIN")) then
				if trim(rsSend("SAFEDETAIN"))="1" then
					if SafeWay="" then
						SafeWay="已假扣押"
					else
						SafeWay=SafeWay&"、已假扣押"
					end if
				end if
			end if
			if trim(rsSend("SAFESHUTSHOP"))<>"" and not isnull(rsSend("SAFESHUTSHOP")) then
				if trim(rsSend("SAFESHUTSHOP"))="1" then
					if SafeWay="" then
						SafeWay="已勒令停業"
					else
						SafeWay=SafeWay&"、已勒令停業"
					end if
				end if
			end if
			if SafeWay="" then
				response.write "&nbsp;"
			else
				response.write SafeWay
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>附件</strong></td>
			<td align="left" colspan="5"><%
			'附件
			SubObject=""
			if trim(rsSend("ATTATCHTABLE"))<>"" and not isnull(rsSend("ATTATCHTABLE")) then
				if trim(rsSend("ATTATCHTABLE"))="1" then
					SubObject="附表"
				end if
			end if
			if trim(rsSend("ATTATCHJUDE"))<>"" and not isnull(rsSend("ATTATCHJUDE")) then
				if trim(rsSend("ATTATCHJUDE"))="1" then
					if SubObject="" then
						SubObject="處分書裁決書或義務人依法令負有義務之證明文件及送達證明文件"
					else
						SubObject=SubObject&"、處分書裁決書或義務人依法令負有義務之證明文件及送達證明文件"
					end if
				end if
			end if
			if trim(rsSend("ATTATCHURGE"))<>"" and not isnull(rsSend("ATTATCHURGE")) then
				if trim(rsSend("ATTATCHURGE"))="1" then
					if SubObject="" then
						SubObject="義務人經限期履行而逾期能不履行之證明文件及送達證明文件"
					else
						SubObject=SubObject&"、義務人經限期履行而逾期能不履行之證明文件及送達證明文件"
					end if
				end if
			end if
			if trim(rsSend("ATTATCHFORTUNE"))<>"" and not isnull(rsSend("ATTATCHFORTUNE")) then
				if trim(rsSend("ATTATCHFORTUNE"))="1" then
					if SubObject="" then
						SubObject="義務人之財產目錄 "
					else
						SubObject=SubObject&"、義務人之財產目錄 "
					end if
				end if
			end if
			if trim(rsSend("ATTATCHGROUND"))<>"" and not isnull(rsSend("ATTATCHGROUND")) then
				if trim(rsSend("ATTATCHGROUND"))="1" then
					if SubObject="" then
						SubObject="土地登記部謄本"
					else
						SubObject=SubObject&"、土地登記部謄本"
					end if
				end if
			end if
			if trim(rsSend("ATTATCHREGISTER"))<>"" and not isnull(rsSend("ATTATCHREGISTER")) then
				if trim(rsSend("ATTATCHREGISTER"))="1" then
					if SubObject="" then
						SubObject="義務人之戶籍資料"
					else
						SubObject=SubObject&"、義務人之戶籍資料"
					end if
				end if
			end if
			if trim(rsSend("ATTATCHFILELIST"))<>"" and not isnull(rsSend("ATTATCHFILELIST")) then
				if trim(rsSend("ATTATCHFILELIST"))="1" then
					if SubObject="" then
						SubObject="磁片電子檔清單"
					else
						SubObject=SubObject&"、磁片電子檔清單"
					end if
				end if
			end if
			if trim(rsSend("ATTATPOSTAGE"))<>"" and not isnull(rsSend("ATTATPOSTAGE")) then
				if trim(rsSend("ATTATPOSTAGE"))="1" then
					if SubObject="" then
						SubObject="郵資"
					else
						SubObject=SubObject&"、郵資"
					end if
				end if
			end if
			if SubObject="" then
				response.write "&nbsp;"
			else
				response.write SubObject
			end if
			%></td>
		</tr><%

		if sys_City="台中市" or sys_City = "彰化縣" or sys_City = "台南市" or sys_City = "高雄市" or sys_City = "高雄縣" or sys_City="宜蘭縣" or sys_City="基隆市" or sys_City="屏東縣" then
			Response.Write "<tr>"
				Response.Write "<td align=""center"" bgcolor=""#EBE5FF"" colspan=""6"">"
				Response.Write "<strong>移送歷程與債權記錄</strong>"
				Response.Write "</td>"
			Response.Write "</tr>"

			strSql="select * from (select SN SendDetialSN,SendDate,OpenGovNumber SendGovNumber,SendNumber,RecordMemberID from PasserSendDetail where BillSN="&trim(rs1("SN"))&") a,(select sn,SendDetailSN,OpenGovNumber CreditorGovNumber,CreditorNumber,PetitionDate,Decode(CreditorTypeID,1,'無個人財產','清償中') CreditorTypeName,RemainNT from PasserCreditor where BillSN="&trim(rs1("SN"))&")b where a.SendDetialSN=b.SendDetailSN(+) order by SendDate DESC,PetitionDate DESC"

			set rs=conn.execute(strSQL)
			While Not rs.eof
				Response.Write "<tr>"
					Response.Write "<td align=""right"" bgcolor=""#EBE5FF"">"
					Response.Write "<strong>移送日期</strong>"
					Response.Write "</td>"
					Response.Write "<td>"
					Response.Write gInitDT(trim(rs("SendDate")))
					Response.Write "</td>"

					Response.Write "<td align=""right"" bgcolor=""#EBE5FF"">"
					Response.Write "<strong>發文文號</strong>"
					Response.Write "</td>"
					Response.Write "<td>"
					Response.Write trim(rs("SendGovNumber"))
					Response.Write "</td>"

					Response.Write "<td align=""right"" bgcolor=""#EBE5FF"">"
					Response.Write "<strong>移送案號</strong>"
					Response.Write "</td>"
					Response.Write "<td>"
					Response.Write trim(rs("SendNumber"))
					Response.Write "</td>"
				Response.Write "</tr>"
				Response.Write "<tr>"
					Response.Write "<td align=""right"" bgcolor=""#EBE5FF"">"
					Response.Write "<strong>申請時間</strong>"
					Response.Write "</td>"
					Response.Write "<td>"
					Response.Write gInitDT(trim(rs("PetitionDate")))
					Response.Write "</td>"

					Response.Write "<td align=""right"" bgcolor=""#EBE5FF"">"
					Response.Write "<strong>憑証編號</strong>"
					Response.Write "</td>"
					Response.Write "<td>"
					Response.Write trim(rs("CreditorGovNumber"))
					Response.Write "</td>"

					Response.Write "<td align=""right"" bgcolor=""#EBE5FF"">"
					Response.Write "<strong>執行案號</strong>"
					Response.Write "</td>"
					Response.Write "<td>"
					Response.Write trim(rs("CreditorNumber"))
					Response.Write "</td>"
				Response.Write "</tr>"
				Response.Write "<tr>"
					Response.Write "<td align=""right"" bgcolor=""#EBE5FF"">"
					Response.Write "<strong>執行狀態</strong>"
					Response.Write "</td>"
					Response.Write "<td>"
					Response.Write trim(rs("CreditorTypeName"))
					Response.Write "</td>"

					Response.Write "<td align=""right"" bgcolor=""#EBE5FF"">"
					Response.Write "<strong>待執行金額</strong>"
					Response.Write "</td>"
					Response.Write "<td>"
					Response.Write trim(Sys_Forfeit)
					Response.Write "</td>"
				Response.Write "</tr>"

				rs.movenext
			Wend
		end if
	end if
	rsSend.close
	set rsSend=nothing
%>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>4"></a>
				<strong>行人攤販催告書</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>2">行人攤販裁決書</a>‧
				<a href="#<%=trim(rs1("SN"))%>3">行人攤販移送書</a>‧
				<a href="#<%=trim(rs1("SN"))%>5">行人攤販繳費記錄</a>‧
				<a href="#<%=trim(rs1("SN"))%>7">行人攤販送達紀錄</a>
			</td>
		</tr>
<%
	strPasserUrge="select * from PasserUrge where BillSn="&trim(rs1("Sn"))
	set rsUrge=conn.execute(strPasserUrge)
	if not rsUrge.eof then
%>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>舉發單號</strong></td>
			<td align="left"><%
			'舉發單號
			if trim(rsUrge("BILLNO"))<>"" and not isnull(rsUrge("BILLNO")) then
				response.write trim(rsUrge("BILLNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>發文字號</strong></td>
			<td align="left" colspan="3"><%
			'發文字號
			if trim(rsUrge("OPENGOVNUMBER"))<>"" and not isnull(rsUrge("OPENGOVNUMBER")) then
				response.write trim(rsUrge("OPENGOVNUMBER"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>催告日期</strong></td>
			<td align="left"><%
			'催告日期
			if trim(rsUrge("URGEDATE"))<>"" and not isnull(rsUrge("URGEDATE")) then
				response.write gArrDT(trim(rsUrge("URGEDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>催繳方式</strong></td>
			<td align="left"><%
			'催繳方式
			if trim(rsUrge("URGETYPEID"))<>"" and not isnull(rsUrge("URGETYPEID")) then
				if trim(rsUrge("URGETYPEID"))="0" then
					response.write "電話"
				elseif trim(rsUrge("URGETYPEID"))="1" then
					response.write "信函"
				elseif trim(rsUrge("URGETYPEID"))="2" then
					response.write "雙掛號或裁決書"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>罰款金額</strong></td>
			<td align="left"><%
			'罰款金額
			if trim(rsUrge("FORFEIT"))<>"" and not isnull(rsUrge("FORFEIT")) then
				response.write trim(rsUrge("FORFEIT"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>寄送地址</strong></td>
			<td align="left" colspan="3"><%
			'寄送地址
			if trim(rsUrge("SENDADDRESS"))<>"" and not isnull(rsUrge("SENDADDRESS")) then
				response.write trim(rsUrge("SENDADDRESS"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>局長</strong></td>
			<td align="left"><%
			'局長
			if trim(rsUrge("BIGUNITBOSSNAME"))<>"" and not isnull(rsUrge("BIGUNITBOSSNAME")) then
				response.write trim(rsUrge("BIGUNITBOSSNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>分局長</strong></td>
			<td align="left"><%
			'分局長
			if trim(rsUrge("SUBUNITSECBOSSNAME"))<>"" and not isnull(rsUrge("SUBUNITSECBOSSNAME")) then
				response.write trim(rsUrge("SUBUNITSECBOSSNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>聯絡電話</strong></td>
			<td align="left"><%
			'聯絡電話
			if trim(rsUrge("CONTACTTEL"))<>"" and not isnull(rsUrge("CONTACTTEL")) then
				response.write trim(rsUrge("CONTACTTEL"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>紀錄狀態</strong></td>
			<td align="left"><%
			'紀錄狀態
			if trim(rsUrge("RECORDSTATEID"))<>"" and not isnull(rsUrge("RECORDSTATEID")) then
				if trim(rsUrge("RECORDSTATEID"))="0" then
					response.write "正常"
				else
					response.write "刪除"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#EBE5FF"><strong>紀錄時間</strong></td>
			<td align="left"><%
			'紀錄時間
			if trim(rsUrge("RECORDDATE"))<>"" and not isnull(rsUrge("RECORDDATE")) then
				response.write gArrDT(trim(rsUrge("RECORDDATE")))&" "
				response.write Right("00"&hour(rsUrge("RECORDDATE")),2)&":"
				response.write Right("00"&minute(rsUrge("RECORDDATE")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>紀錄人員</strong></td>
			<td align="left"><%
			'紀錄人員
			if trim(rsUrge("RECORDMEMBERID"))<>"" and not isnull(rsUrge("RECORDMEMBERID")) then
				strRecMem="select chName from MemberData where MemberID="&trim(rsUrge("RECORDMEMBERID"))
				set rsRecMem=conn.execute(strRecMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#EBE5FF"><strong>刪除人員</strong></td>
			<td align="left"><%
			'刪除人員
			if trim(rsUrge("DELMEMBERID"))<>"" and not isnull(rsUrge("DELMEMBERID")) then
				strRecMem="select chName from MemberData where MemberID="&trim(rsUrge("DELMEMBERID"))
				set rsRecMem=conn.execute(strRecMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
<%	end if
	rsUrge.close
	set rsUrge=nothing
%>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>5"></a>
				<strong>行人攤販繳費記錄</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>2">行人攤販裁決書</a>‧
				<a href="#<%=trim(rs1("SN"))%>3">行人攤販移送書</a>‧
				<a href="#<%=trim(rs1("SN"))%>4">舉行人攤販催告書</a>‧
				<a href="#<%=trim(rs1("SN"))%>7">行人攤販送達紀錄</a>
			</td>
		</tr>
<%	i=0
	strPasserPay="select * from PasserPay where BillSn="&trim(rs1("SN"))&" order by RECORDDATE desc"
	set rsPay=conn.execute(strPasserPay)
	If Not rsPay.Bof Then rsPay.MoveFirst 
	While Not rsPay.Eof
		if i=0 then
			i=i+1
			TRcolor="#EBE5FF"
		else
			i=i-1
			TRcolor="#AAF2A2"
		end if
%>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>單號</strong></td>
			<td align="left"><%
			'單號
			if trim(rsPay("BILLNO"))<>"" and not isnull(rsPay("BILLNO")) then
				response.write trim(rsPay("BILLNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>收據號碼</strong></td>
			<td align="left"><%
			'收據號碼
			if trim(rsPay("PAYNO"))<>"" and not isnull(rsPay("PAYNO")) then
				response.write trim(rsPay("PAYNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>繳費次數</strong></td>
			<td align="left"><%
			'繳費次數
			if trim(rsPay("PAYTIMES"))<>"" and not isnull(rsPay("PAYTIMES")) then
				response.write trim(rsPay("PAYTIMES"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>繳費方式</strong></td>
			<td align="left"><%
			'繳費方式
			if trim(rsPay("PAYTYPEID"))<>"" and not isnull(rsPay("PAYTYPEID")) then
				if trim(rsPay("PAYTYPEID"))="1" then 
					response.write "窗口"
				else
					response.write "郵撥"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>繳費時間</strong></td>
			<td align="left"><%
			'繳費時間
			if trim(rsPay("PAYDATE"))<>"" and not isnull(rsPay("PAYDATE")) then
				response.write gArrDT(trim(rsPay("PAYDATE")))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>繳費人姓名</strong></td>
			<td align="left"><%
			'繳費人姓名
			if trim(rsPay("PAYER"))<>"" and not isnull(rsPay("PAYER")) then
				response.write trim(rsPay("PAYER"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>應繳金額</strong></td>
			<td align="left"><%
			'應繳金額
			if trim(rsPay("FORFEIT"))<>"" and not isnull(rsPay("FORFEIT")) then
				response.write trim(rsPay("FORFEIT"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>繳費金額</strong></td>
			<td align="left"><%
			'繳費金額
			if trim(rsPay("PAYAMOUNT"))<>"" and not isnull(rsPay("PAYAMOUNT")) then
				response.write trim(rsPay("PAYAMOUNT"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>是否繳清結案</strong></td>
			<td align="left"><%
			'是否繳清結案
			if trim(rsPay("CASECLOSE"))<>"" and not isnull(rsPay("CASECLOSE")) then
				if trim(rsPay("CASECLOSE"))="1" then
					response.write "結案"
				else
					response.write "未結案"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>紀錄狀態</strong></td>
			<td align="left"><%
			'紀錄狀態
			if trim(rsPay("RECORDSTATEID"))<>"" and not isnull(rsPay("RECORDSTATEID")) then
				if trim(rsPay("RECORDSTATEID"))="0" then
					response.write "正常"
				else
					response.write "刪除"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>紀錄時間</strong></td>
			<td align="left"><%
			'紀錄時間
			if trim(rsPay("RECORDDATE"))<>"" and not isnull(rsPay("RECORDDATE")) then
				response.write gArrDT(trim(rsPay("RECORDDATE")))&" "
				response.write Right("00"&hour(rsPay("RECORDDATE")),2)&":"
				response.write Right("00"&minute(rsPay("RECORDDATE")),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>紀錄人員</strong></td>
			<td align="left"><%
			'紀錄人員
			if trim(rsPay("RECORDMEMBERID"))<>"" and not isnull(rsPay("RECORDMEMBERID")) then
				strRecMem="select chName from MemberData where MemberID="&trim(rsPay("RECORDMEMBERID"))
				set rsRecMem=conn.execute(strRecMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>刪除人員</strong></td>
			<td align="left"><%
			'刪除人員
			if trim(rsPay("DELMEMBERID"))<>"" and not isnull(rsPay("DELMEMBERID")) then
				strRecMem="select chName from MemberData where MemberID="&trim(rsPay("DELMEMBERID"))
				set rsRecMem=conn.execute(strRecMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>備註</strong></td>
			<td align="left"><%
			'備註
			if trim(rsPay("NOTE"))<>"" and not isnull(rsPay("NOTE")) then
				response.write trim(rsPay("NOTE"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>是否遲繳</strong></td>
			<td align="left"><%
			'是否遲繳
			if trim(rsPay("ISLATE"))<>"" and not isnull(rsPay("ISLATE")) then
				if trim(rsPay("ISLATE"))="0" then
					response.write "如期繳納"
				else
					response.write "逾期繳納"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td></td>
		</tr>
<%	rsPay.MoveNext
	Wend
	rsPay.close
	set rsPay=nothing
%>
		<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35">
				<a name="#<%=trim(rs1("SN"))%>7"></a>
				<strong>行人攤販送達記錄</strong>&nbsp;&nbsp;&nbsp;&nbsp;
				>><a href="#<%=trim(rs1("SN"))%>1">舉發單基本資料</a>‧
				<a href="#<%=trim(rs1("SN"))%>2">行人攤販裁決書</a>‧
				<a href="#<%=trim(rs1("SN"))%>3">行人攤販移送書</a>‧
				<a href="#<%=trim(rs1("SN"))%>4">舉行人攤販催告書</a>‧
				<a href="#<%=trim(rs1("SN"))%>5">行人攤販繳費記錄</a>
			</td>
		</tr>
<%	i=0
	strPasserArrived="select * from PasserSendArrived where PasserSn="&trim(rs1("SN"))&" order by SN"
	set rsArrived=conn.execute(strPasserArrived)
	If Not rsArrived.Bof Then rsArrived.MoveFirst 
	While Not rsArrived.Eof
		if i=0 then
			i=i+1
			TRcolor="#EBE5FF"
		else
			i=i-1
			TRcolor="#AAF2A2"
		end if
%>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>送達流水號</strong></td>
			<td align="left"><%
			'送達流水號
			if trim(rsArrived("SN"))<>"" and not isnull(rsArrived("SN")) then
				response.write trim(rsArrived("SN"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>登錄人</strong></td>
			<td align="left"><%
			'登錄人
			if trim(rsArrived("RecordMemberID"))<>"" and not isnull(rsArrived("RecordMemberID")) then
				strRecMem="select chName from MemberData where MemberID="&trim(rsArrived("RecordMemberID"))
				set rsRecMem=conn.execute(strRecMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>送達日期</strong></td>
			<td align="left"><%
			'送達日期
			if trim(rsArrived("ArrivedDate"))<>"" and not isnull(rsArrived("ArrivedDate")) then
				response.write gArrDT(trim(rsArrived("ArrivedDate")))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>送達人員</strong></td>
			<td align="left"><%
			'送達人員
			if trim(rsArrived("SenderMemID"))<>"" and not isnull(rsArrived("SenderMemID")) then
				strRecMem="select chName from MemberData where MemberID="&trim(rsArrived("SenderMemID"))
				set rsRecMem=conn.execute(strRecMem)
				if not rsRecMem.eof then
					response.write trim(rsRecMem("chName"))
				end if
				rsRecMem.close
				set rsRecMem=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="<%=TRcolor%>"><strong>送達圖片</strong></td>
			<td colspan="3"><%
				If not ifnull(rsArrived("Imagefilename")) Then 
					Response.Write "<a href='"&replace(rsArrived("IMAGEDIRECTORYNAME")&rsArrived("IMAGEFILENAME"),"./Picture/","../PasserBase/Picture/")&"' TARGET ='_blank'>"
					Response.Write "<Span "& replace(ShowMap(trim(rsArrived("SN"))),"./Picture/","../PasserBase/Picture/")&">送達圖片</span></a>"
				end If 
			%>
			</td>
		</tr>
		<tr>
			<td></td>
		</tr>
<%	rsArrived.MoveNext
	Wend
	rsArrived.close
	set rsArrived=nothing
%>
	</table>
	<br>
<%	rs1.MoveNext
	Wend
	rs1.close
	set rs1=nothing
%>
<%
conn.close
set conn=nothing
%>
<center>
<input type="button" value="列印" onclick="DP();">
<br>
(若無列印鈕，可按下滑鼠右鍵選擇列印功能)
</center>
</body>
<script language="JavaScript">
function DP(){
	window.focus();
	window.print();
}
</script>
</html>
