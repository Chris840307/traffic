<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style3 {font-family:新細明體; color:0044ff; line-height:19px; font-size: 15px}
.style4 {font-family:新細明體;  line-height:19px;font-size: 12pt}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>退件清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 800
Response.flush
%>
<%
'權限
'AuthorityCheck(234)

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
%>
<%
	strwhere=request("SQLstr")
%>

</head>
<body>
<form name=myForm method="post">
<%	
	TitleValue=""
	strUnitName2="select UnitName from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
	set rsUnitName2=conn.execute(strUnitName2)
	if not rsUnitName2.eof then
		TitleUnitName2=trim(rsUnitName2("UnitName"))
		
	end if
	rsUnitName2.close
	set rsUnitName2=nothing

	strTitle="select Value from Apconfigure where ID=31"
	set rsTitle=conn.execute(strTitle)
	if not rsTitle.eof then
		TitleValue=rsTitle("Value")
		TitleValue=Replace(TitleValue,"台","臺")
	end if
	rsTitle.close
	set rsTitle=nothing

	PrintSN=0
%>
<%		
		strwhere=""
		if UCase(request("Sys_BatchNumber"))<>"" then
			tmp_BatchNumber=split(UCase(request("Sys_BatchNumber")),",")
			for i=0 to Ubound(tmp_BatchNumber)
				if i>0 then Sys_BatchNumber=trim(Sys_BatchNumber)&","
				if i=0 then
					Sys_BatchNumber=trim(Sys_BatchNumber)&UCase(tmp_BatchNumber(i))
				else
					Sys_BatchNumber=trim(Sys_BatchNumber)&"'"&UCase(tmp_BatchNumber(i))
				end if
				if i<Ubound(tmp_BatchNumber) then Sys_BatchNumber=trim(UCase(Sys_BatchNumber))&"'"
			next
			dciStr=" and BatchNumber in ('"&Sys_BatchNumber&"')"
		end if

		if trim(request("Sys_ImageFileNameB1"))<>"" and trim(request("Sys_ImageFileNameB2"))<>"" then
			Sys_BillNo1=right("00000000000000000"&trim(request("Sys_ImageFileNameB1")),16)
			Sys_BillNo2=right("00000000000000000"&trim(request("Sys_ImageFileNameB2")),16)

			strwhere=strwhere&" and a.ImageFileNameB between '"&Sys_BillNo1&"' and '"&Sys_BillNo2&"'"

		elseif trim(request("Sys_ImageFileNameB1"))<>"" then
			Sys_BillNo1=right("00000000000000000"&trim(request("Sys_ImageFileNameB1")),16)

			strwhere=strwhere&" and a.ImageFileNameB between '"&Sys_BillNo1&"' and '"&Sys_BillNo1&"'"

		elseif trim(request("Sys_ImageFileNameB2"))<>"" then
			Sys_BillNo2=right("00000000000000000"&trim(request("Sys_ImageFileNameB2")),16)

			strwhere=strwhere&" and a.ImageFileNameB between '"&Sys_BillNo2&"' and '"&Sys_BillNo2&"'"

		end if

		UserMarkDate1=gOutDT(request("Sys_UserMarkDate1"))&" 0:0:0"
		UserMarkDate2=gOutDT(request("Sys_UserMarkDate2"))&" 23:59:59"
		strwhere=strwhere&" and g.UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')"

		strCnt="select count(*) as cnt" &_
		" from BillBase a,(select billsn from DciLog where exchangetypeid='A'"&dciStr&") b,(select CarNo,Owner from billbasedcireturn where exchangetypeid='A') e,StopBillMailHistory g" &_
		" where a.Sn=b.BillSn and a.CarNo=e.CarNo and a.Sn=g.BillSn" &_
		" and a.RecordStateID=0 and g.UserMarkResonID in ('5','6','7','T')" &strwhere

		set rsCnt=conn.execute(strCnt)
		if not rsCnt.eof then
			GetCnt=cint(rsCnt("Cnt"))
		end if
		rsCnt.close
		set rsCnt=nothing

		MaxDate=now
		MinDate=now
		strDate="select Max(IllegalDate) as MaxDate,Min(IllegalDate) as MinDate" &_
		" from BillBase a,(select billsn from DciLog where exchangetypeid='A'"&dciStr&") b,(select CarNo,Owner from billbasedcireturn where exchangetypeid='A') e,StopBillMailHistory g" &_
		" where a.Sn=b.BillSn and a.CarNo=e.CarNo and a.Sn=g.BillSn" &_
		" and a.RecordStateID=0 and g.UserMarkResonID in ('5','6','7','T')" &strwhere

		set rsDate=conn.execute(strDate)
		if not rsDate.eof then
			MaxDate=rsDate("MaxDate")
			MinDate=rsDate("MinDate")
		end if
		rsDate.close
		set rsDate=nothing

		strSQL="select a.SN,a.CarNO,e.Owner,a.IllegalDate" &_
		",a.ImageFileNameB,g.UserMarkResonID,g.StoreAndSendEffectDate,g.UserMarkDate" &_
		" from BillBase a,(select billsn from DciLog where exchangetypeid='A'"&dciStr&") b,(select CarNo,Owner from billbasedcireturn where exchangetypeid='A') e,StopBillMailHistory g" &_
		" where a.Sn=b.BillSn and a.CarNo=e.CarNo and a.Sn=g.BillSn" &_
		" and a.RecordStateID=0 and g.UserMarkResonID in ('5','6','7','T')" &strwhere &_
		" order by g.UserMarkDate"

		set rs1=conn.execute(strSQL)
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
		if PrintSN>0 then response.write "<div class=""PageNext""></div>"
%>		
	<table width="100%" border="1" cellpadding="0" cellspacing="0">
	<tr>
		<td align="center" height="55" colspan="7"><span class="style4"><%
			response.write TitleValue&"公有路邊收費停車場停車費催繳單郵寄無法送達清冊"
			response.write "<br>"
			response.write "<div align=""right"">停車日期："&year(MinDate)-1911&"/"&month(MinDate)&"/"&Day(MinDate)&"~"
			response.write year(MaxDate)-1911&"/"&month(MaxDate)&"/"&Day(MaxDate)&"&nbsp; &nbsp;</div>"
			if GetCnt="0" then
				pagecnt=1
			else
				pagecnt=fix(GetCnt/20+0.9999999)
			end if
		%></span></td>
	</tr>
	<tr>
		<td align="center" height="40" width="5%"><span class="style4">編號</span></td>
		<td align="center" width="8%"><span class="style4">違規日期</span></td>
		<td align="center" width="8%"><span class="style4">寄存日期</span></td> 
		<td align="center" width="19%"><span class="style4">催繳單號</span></td>
		<td align="center" width="9%"><span class="style4">車號</span></td>
		<td align="center" width="17%"><span class="style4">車主姓名</span></td>
		<% If sys_City<>"台東縣" then %>
		<td align="center" width="28%"><span class="style4">車主地址</span></td>
		<% end If %>
		<td align="center" width="17%"><span class="style4">退件原因</span></td>
	</tr>
<%		for i=1 to 20
			if rs1.eof then exit for
			PrintSN=PrintSN+1
%>
	<tr>
		<td align="center" height="40"><span class="style4"><%
		'編號
		response.write PrintSN
		%></span></td>
		<td align="center"><span class="style4"><%
		'違規日期
		response.write gArrDT(trim(rs1("IllegalDate")))
		%></span></span></td>
		<td align="center"><span class="style4"><%
		'寄存日期
		response.write gArrDT(trim(rs1("StoreAndSendEffectDate")))
		%></span></span></td>
		<td align="center"><span class="style4"><%
		'催繳單號
		response.write trim(rs1("ImageFileNameB"))
		%></span></td>
		<td align="center"><span class="style4"><%
		'車號
		response.write trim(rs1("CarNo"))
		%></span></td>
		<td align="center"><span class="style4"><%
		'車主姓名
		response.write funcCheckFont(trim(rs1("Owner")),24,1)
		%></span></td>

		<% If sys_City<>"台東縣" then %>
		<td align="center"><span class="style4"><%
		'車主地址
		if trim(rs1("ImageFileNameB"))<>"" or not isnull(rs1("ImageFileNameB")) then
			strAddr1="select * from Billbase where imagefilenameb='"&trim(rs1("ImageFileNameB"))&"'"
			Set rsAddr1=conn.execute(strAddr1)
			If Not rsAddr1.eof Then
				If Trim(rsAddr1("OwnerAddress"))<>"" and Not IsNull(rsAddr1("OwnerAddress")) Then
					response.write trim(rsAddr1("OwnerZip"))&funcCheckFont(trim(rsAddr1("OwnerAddress")),18,1)
				Else
					strAddr2="select * from BillBaseDcireturn where CarNo='"&trim(rs1("CarNo"))&"' and Exchangetypeid='A'"
					Set rsAddr2=conn.execute(strAddr2)
					If Not rsAddr2.eof Then
						If Trim(rsAddr2("OWNERNOTIFYADDRESS"))<>"" and Not IsNull(rsAddr2("OWNERNOTIFYADDRESS")) Then
							NotifyZip=""
							strNZ="select * from Zip where ZipName like '"&left(trim(rsAddr2("OWNERNOTIFYADDRESS")),5)&"%'"
							set rsNZ=conn.execute(strNZ)
							if not rsNZ.eof then
								NotifyZip=trim(rsNZ("ZipNo"))
							else
								strNZ2="select * from Zip where ZipName like '"&left(trim(rsAddr2("OWNERNOTIFYADDRESS")),3)&"%'"
								set rsNZ2=conn.execute(strNZ2)
								if not rsNZ2.eof then
									NotifyZip=trim(rsNZ2("ZipNo"))
								
								end if
								rsNZ2.close
								set rsNZ2=nothing
							end if
							rsNZ.close
							set rsNZ=Nothing
							response.write NotifyZip&funcCheckFont(trim(rsAddr2("OWNERNOTIFYADDRESS")),18,1)
						Else	
							response.write trim(rsAddr2("OwnerZip"))&funcCheckFont(trim(rsAddr2("OwnerAddress")),18,1)
						End if
					End If
					rsAddr2.close
					Set rsAddr2=Nothing 
				End if
			End If
			rsAddr1.close
			Set rsAddr1=nothing
		else
			response.write "&nbsp;"
		end if	
		%></span></td>
		<% end If %>
		<td align="center"><span class="style4"><%
		'退件原因
		strCode="select Content from DCIcode where TypeID=7 and ID='"&trim(rs1("UserMarkResonID"))&"'"
		set rsCode=conn.execute(strCode)
		if not rsCode.eof then
			response.write trim(rsCode("Content"))
		end if
		rsCode.close
		set rsCode=nothing
		%></span></td>
	</tr>
<%			
		rs1.MoveNext
		next
%>
	</table>
<%
		Wend
		rs1.close
		set rs1=nothing
%>		

</form>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}

printWindow(true,8,8.08,8.08,8.08);
</script>
<%conn.close%>