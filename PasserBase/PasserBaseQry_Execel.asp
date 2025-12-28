<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
Server.ScriptTimeout=12000

fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_慢車行人道路障礙舉發單.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

sys_City=replace(sys_City,"台中縣","台中市")
sys_City=replace(sys_City,"台南縣","台南市")

showCreditor=false
if sys_City="台中市" or sys_City = "彰化縣" or sys_City = "台南市" or sys_City = "高雄市" or sys_City = "高雄縣" or sys_City="宜蘭縣" or sys_City="基隆市" or sys_City="澎湖縣" or sys_City="屏東縣" then
	showCreditor=true
end If 

If Request("DB_Cnt") < 5000 Then
	
	If Not ifnull(request("Sys_SendBillSN")) Then

		sys_billsn=request("Sys_SendBillSN")
	elseif Not ifnull(request("hd_BillSN")) Then

		sys_billsn=request("hd_BillSN")
	else

		sys_billsn=request("BillSN")
	End If 

	tmp_billsn=split(sys_billsn,",")

	sys_billsn=""

	For i = 0 to Ubound(tmp_billsn)

		If i >0 then

			If i mod 100 = 0 Then

				sys_billsn=sys_billsn&"@"
			elseif sys_billsn<>"" then

				sys_billsn=sys_billsn&","
			end If 
		end if

		sys_billsn=sys_billsn&tmp_billsn(i)

	Next

	tmpSQL=""

	If Ubound(tmp_billsn) >= 100 Then

		sys_billsn=split(sys_billsn,"@")
		
		For i = 0 to Ubound(sys_billsn)
			
			If tmpSQL <>"" Then tmpSQL=tmpSQL&" union all "
			
			tmpSQL=tmpSQL&"select sn from passerbase where sn in("&sys_billsn(i)&")"
		Next

	else

		tmpSQL="select sn from passerbase where sn in("&sys_billsn&")"

	End if 

	BasSQL="("&tmpSQL&") tmpPasser"
	theWhere=" and Exists(select 'Y' from "&BasSQL&" where sn=a.sn)"&Request("orderstr")
else
	theWhere=Request("ExportSQL")
End if 

 

'檢查是否可進入本系統
	strSQLTemp="select a.SN,a.IllegalDate,a.BillNo,a.Driver,a.DriverBirth,a.DriverID,a.DriverAddress," &_
	"a.IllegalAddress,a.BillFillDate,a.DeallIneDate,a.Memberstation,a.Rule1,a.Rule2,a.Rule3," &_
	"a.Rule4,a.RuleVer,a.FORFEIT1,a.FORFEIT2,a.FORFEIT3,a.FORFEIT4,a.BILLSTATUS," &_
	"(select nvl(min(level1),0) from law where version=2 and itemid=a.rule1) nForfeit1," &_
	"(select nvl(min(level1),0) from law where version=2 and itemid=a.rule2) nForfeit2," &_
	"(select nvl(min(level1),0) from law where version=2 and itemid=a.rule3) nForfeit3," &_
	"(select nvl(min(level1),0) from law where version=2 and itemid=a.rule4) nForfeit4," &_
	"a.BillMem1,a.DoubleCheckStatus,a.RecordDate," &_
	"(Select name from Project where projectID=a.projectID) ProjectName," &_
	"(Select JudeDate from PasserJude where billsn=a.sn) JUDEDATE," &_
	"(Select SendDate from PasserSend where billsn=a.sn) SENDDATE," &_
	"(Select UrgeDate from PasserUrge where billsn=a.sn) URGEDATE," &_
	"(select UnitName from Unitinfo where UnitID=a.billUnitID) UnitName," &_
	"(select MAX(PayDate) PayDate from PasserPay where billsn=a.sn) PayDate"&showFiled &_
	" from PasserBase a where a.RecordStateID=0"&theWhere

	'If sys_City="台南市" Then ConnExecute "慢車匯出："&strSQLTemp ,360	

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style2 {font-size: 14px;mso-style-parent:style0;mso-number-format:"\@";}
-->
</style>
<title>慢車行人道路障礙舉發單查詢</title>
</head>
<body>
<table width="100%" border="1">
	<tr>
		<td height="26" align="center"><strong>舉發單紀錄列表</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="1">
				<tr>
					<th>違規日期</th>
					<th>違規時間</th>
					<th>舉發單號</th>
					<th>舉發單位</th>
					<th>車種(專案)</th>
					<th>舉發人</th>
					<th>違規人</th>
					<th>違規人出生日</th>
					<th>違規人身份證字號</th>
					<th>違規人地址</th>
					<th>違規地點</th>
					<th>法條</th>
					<th>原罰款金額</th>
					<th>裁決金額</th>
					<th>填單日</th>
					<th>應到案日期</th>
					<th>應到案處所</th>
					<th>催告日</th>
					<th>裁決日</th>
					<th>移送日</th>
					<th>送達日</th>
					<th>付費日</th>
					<th>結案日</th>
					<th>繳費字號</th>
					<th>已繳金額</th>
					<th>結案狀態</th>
					<th>建檔日</th>
					<%
					if showCreditor then
						Response.Write "<th>債權取得日</th>"
						Response.Write "<th>債權收文文號</th>"
					end if
					%>
					
				</tr>
				<%
				fileCnt=0
				set rsfound=conn.execute(strSQLTemp)
				while Not rsfound.eof
					fileCnt=fileCnt+1

					response.write "<tr>"
					response.write "<td>"&gInitDT(trim(rsfound("IllegalDate")))& "</td>"
					response.write "<td>"&hour(rsfound("IllegalDate"))&":"&Minute(rsfound("IllegalDate"))&"</td>"
					response.write "<td>"&trim(rsfound("BillNo"))&"</td>"
					response.write "<td>"&trim(rsfound("UnitName"))&"</td>"
					response.write "<td>"&trim(rsfound("ProjectName"))&"</td>"
					response.write "<td>"&trim(rsfound("BillMem1"))&"</td>"
					response.write "<td>"&trim(rsfound("Driver"))&"</td>"
					response.write "<td>"&gInitDT(trim(rsfound("DriverBirth")))&"</td>"
					response.write "<td>"&trim(rsfound("DriverID"))&"</td>"
					response.write "<td>"&trim(rsfound("DriverAddress"))&"</td>"
					response.write "<td>"&trim(rsfound("IllegalAddress"))&"</td>"

					chRule=trim(rsfound("Rule1"))
					if rsfound("Rule2")<>"" then chRule=chRule&"/"&rsfound("Rule2")
					if rsfound("Rule3")<>"" then chRule=chRule&"/"&rsfound("Rule3")

					nforfeit=cdbl(rsfound("nForfeit1"))+cdbl(rsfound("nForfeit2"))+cdbl(rsfound("nForfeit3"))+cdbl(rsfound("nForfeit4"))

					FORFEIT=cdbl("0"&rsfound("FORFEIT1"))

					if rsfound("FORFEIT2")<>"" then
						FORFEIT=FORFEIT+cdbl("0"&rsfound("FORFEIT2"))
					end if

					if rsfound("FORFEIT3")<>"" then
						FORFEIT=FORFEIT+cdbl("0"&rsfound("FORFEIT3"))
					end if

					if rsfound("FORFEIT4")<>"" then
						FORFEIT=FORFEIT+cdbl("0"&rsfound("FORFEIT4"))
					end If 

					if sys_City="基隆市" then
						If not ifnull(rsfound("JUDEDATE")) Then

							strRule1="select * from Law where ItemID='"&trim(rsfound("Rule1"))&"' and VERSION='2'"
							set rsRule1=conn.execute(strRule1)
							if not rsRule1.eof then

								FORFEIT=cint(trim(rsRule1("Level2")))
							end if
							rsRule1.close
							set rsRule1=nothing

							if rsfound("FORFEIT2")<>"" then

								strRule1="select * from Law where ItemID='"&trim(rsfound("Rule2"))&"' and VERSION='2'"
								set rsRule1=conn.execute(strRule1)
								if not rsRule1.eof then

									FORFEIT=FORFEIT+cdbl("0"&trim(rsRule1("Level2")))
								end if
								rsRule1.close
								set rsRule1=nothing
							end if

							if rsfound("FORFEIT3")<>"" then
								
								strRule1="select * from Law where ItemID='"&trim(rsfound("Rule3"))&"' and VERSION='2'"
								set rsRule1=conn.execute(strRule1)
								if not rsRule1.eof then

									FORFEIT=FORFEIT+cdbl("0"&trim(rsRule1("Level2")))
								end if
								rsRule1.close
								set rsRule1=nothing
							end if

							if rsfound("FORFEIT4")<>"" then
								
								strRule1="select * from Law where ItemID='"&trim(rsfound("Rule4"))&"' and VERSION='2'"
								set rsRule1=conn.execute(strRule1)
								if not rsRule1.eof then

									FORFEIT=FORFEIT+cdbl("0"&trim(rsRule1("Level2")))
								end if
								rsRule1.close
								set rsRule1=nothing
							end If 
						
						End if 
					end If 

					response.write "<td>"&chRule&"</td>"
					response.write "<td>"&nforfeit&"</td>"
					response.write "<td>"&FORFEIT&"</td>"
					response.write "<td>"&gInitDT(trim(rsfound("BillFillDate")))&"</td>"
					response.write "<td>"&gInitDT(trim(rsfound("DeallIneDate")))&"</td>"
					response.write "<td>"
					If trim(rsfound("MemberStation"))<>"" Then
						strMS="select * from UnitInfo where UnitID='"&trim(rsfound("MemberStation"))&"'"
						Set rsMs=conn.execute(strMS)
						If Not rsMs.eof Then
							response.write Trim(rsMs("UnitName"))
						End if
						rsMs.close
						Set rsMs=nothing
					End If 
					response.write "</td>"
					response.write "<td>"&trim(gInitDT(rsfound("URGEDATE")))&"</td>"
					response.write "<td>"&trim(gInitDT(rsfound("JUDEDATE")))&"</td>"

					SENDDATEtmp=""
					if showCreditor then
						strSD="select max(SendDate) as SendDate from PasserSendDetail where BillSn="&Trim(rsfound("sn"))&" and SendDate is not null"
						Set rsSD=conn.execute(strSD)
						If Not rsSD.eof then	
							SENDDATEtmp=Trim(rsSD("SendDate"))
						End If
						rsSD.close
						Set rsSD=Nothing 
					end if
					
					If SENDDATEtmp="" Or IsNull(SENDDATEtmp) Then
						strSD2="select * from PasserSend where BillSn="&Trim(rsfound("sn"))
						Set rsSD2=conn.execute(strSD2)
						If Not rsSD2.eof Then
							SENDDATEtmp=Trim(rsSD2("SendDate"))
						End If
						rsSD2.close
						Set rsSD2=Nothing 
					End If 
					response.write "<td>"&trim(gInitDT(SENDDATEtmp))&"</td>"
					strSQL="Select PASSERSN,Max(ARRIVEDDATE) as ARRIVEDDATE from PassersEndArrived where PasserSN="&rsfound("SN")&" group by PASSERSN"
					set rssend=conn.execute(strSQL)
					ARRIVEDDATE=""
					if Not rssend.eof then ARRIVEDDATE=trim(gInitDT(rssend("ARRIVEDDATE")))
					rssend.close
					response.write "<td>"&ARRIVEDDATE&"</td>"
					
					tmpPayNo="":tmpPayDate=""
					strSD2="select PayNo,PayDate from PasserPay where BillSN="&Trim(rsfound("sn"))
					Set rsSD2=conn.execute(strSD2)

					while Not rsSD2.eof
						If tmpPayDate <>"" Then
							tmpPayNo=tmpPayNo&"\"
							tmpPayDate=tmpPayDate&"\"
						end If 
						
						tmpPayDate=tmpPayDate&trim(gInitDT(rsSD2("PayDate")))
						tmpPayNo=tmpPayNo&trim(rsSD2("PayNo"))


						rsSD2.movenext
					wend
					rsSD2.close

					response.write "<td>"&tmpPayDate&"</td>"

					tmpPayDate=""
					strSD2="select CaseCloseDate from PasserPay where BillSN="&Trim(rsfound("sn"))
					Set rsSD2=conn.execute(strSD2)

					while Not rsSD2.eof
						If tmpPayDate <>"" Then
							tmpPayDate=tmpPayDate&"\"
						end If 
						
						tmpPayDate=tmpPayDate&trim(gInitDT(rsSD2("CaseCloseDate")))

						rsSD2.movenext
					wend
					rsSD2.close

					response.write "<td>"&tmpPayDate&"</td>"


					response.write "<td class=""style2"">"&tmpPayNo&"</td>"

					Sys_Payamount=0
					strSQL="select sum(Payamount) as Sys_Payamount from PasserPay where BillSN="&rsfound("SN")
					set rspay=conn.execute(strSQL)
					if not rspay.eof then Sys_Payamount=rspay("Sys_Payamount")
					rspay.close

					If ifnull(Sys_Payamount) Then Sys_Payamount=0					

					response.write "<td>"&cdbl(Sys_Payamount)&"</td>"

					if trim(rsfound("BILLSTATUS"))="9" then
						response.write "<td>已結案</td>"
					else
						response.write "<td>未結案</td>"
					end If 
					
					response.write "<td>"&gInitDT(trim(rsfound("RecordDate")))&"</td>"
					
					if showCreditor then
						PetitionDatetmp=""

						strSD="select max(PetitionDate) as PetitionDate from PasserCreditor where BillSn="&Trim(rsfound("sn"))

						Set rsSD=conn.execute(strSD)
						If Not rsSD.eof then	
							PetitionDatetmp=trim(gInitDT(rsSD("PetitionDate")))
						End If
						rsSD.close
						Set rsSD=Nothing 

						response.write "<td>"&PetitionDatetmp&"</td>"

						CreditorNumbertmp=""
						strSD="select CreditorNumber from PasserCreditor where BillSn="&Trim(rsfound("sn"))&" and PetitionDate=(select min(PetitionDate) from PasserCreditor pc where BillSn="&Trim(rsfound("sn"))&")"

						Set rsSD=conn.execute(strSD)
						If Not rsSD.eof then	
							CreditorNumbertmp=trim(rsSD("CreditorNumber"))
						End If
						rsSD.close
						Set rsSD=Nothing 

						response.write "<td class=""style2"">"&CreditorNumbertmp&"</td>"
					end if

					response.write "</tr>"

					if (fileCnt mod 10)=0 then response.flush

					rsfound.MoveNext
				wend
				rsfound.close
				set rsfound=nothing
				%>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%
conn.close
set conn=nothing
%>