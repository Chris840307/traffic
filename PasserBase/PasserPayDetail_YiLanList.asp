<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%

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

max_year="":min_Year="":sysdate1="":sysdate2=""

If not (ifnull(Request("IllegalDate1")) and ifnull(Request("IllegalDate2"))) Then

	min_Year=split(gArrMT(Request("IllegalDate1")),"-")

	max_year=split(gArrMT(Request("IllegalDate2")),"-")

	sysdate1=gOutDT(Request("IllegalDate1"))

	sysdate2=gOutDT(Request("IllegalDate2"))

elseIf not (ifnull(Request("BillFillDate1")) and ifnull(Request("BillFillDate2"))) Then

	min_Year=split(gArrMT(Request("BillFillDate1")),"-")

	max_year=split(gArrMT(Request("BillFillDate2")),"-")

	sysdate1=gOutDT(Request("BillFillDate1"))

	sysdate2=gOutDT(Request("BillFillDate2"))

elseIf not (ifnull(Request("RecordDate1")) and ifnull(Request("RecordDate2"))) Then

	min_Year=split(gArrMT(Request("RecordDate1")),"-")

	max_year=split(gArrMT(Request("RecordDate2")),"-")
	
	sysdate1=gOutDT(Request("RecordDate1"))

	sysdate2=gOutDT(Request("RecordDate2"))

elseIf not (ifnull(Request("SendDate1")) and ifnull(Request("SendDate2"))) Then

	min_Year=split(gArrMT(Request("SendDate1")),"-")

	max_year=split(gArrMT(Request("SendDate2")),"-")
	
	sysdate1=gOutDT(Request("SendDate1"))

	sysdate2=gOutDT(Request("SendDate2"))

elseIf not (ifnull(Request("JudeDate1")) and ifnull(Request("JudeDate2"))) Then

	min_Year=split(gArrMT(Request("JudeDate1")),"-")

	max_year=split(gArrMT(Request("JudeDate2")),"-")
	
	sysdate1=gOutDT(Request("JudeDate1"))

	sysdate2=gOutDT(Request("JudeDate2"))

elseIf not (ifnull(Request("PayDate1")) and ifnull(Request("PayDate2"))) Then

	min_Year=split(gArrMT(Request("PayDate1")),"-")

	max_year=split(gArrMT(Request("PayDate2")),"-")
	
	sysdate1=gOutDT(Request("PayDate1"))

	sysdate2=gOutDT(Request("PayDate2"))

elseIf not (ifnull(Request("Sys_SendDetailDate1")) and ifnull(Request("Sys_SendDetailDate2"))) Then

	min_Year=split(gArrMT(Request("Sys_SendDetailDate1")),"-")

	max_year=split(gArrMT(Request("Sys_SendDetailDate2")),"-")
	
	sysdate1=gOutDT(Request("Sys_SendDetailDate1"))

	sysdate2=gOutDT(Request("Sys_SendDetailDate2"))

elseIf not (ifnull(Request("MakeSureDate1")) and ifnull(Request("MakeSureDate2"))) Then

	min_Year=split(gArrMT(Request("MakeSureDate1")),"-")

	max_year=split(gArrMT(Request("MakeSureDate2")),"-")
	
	sysdate1=gOutDT(Request("MakeSureDate1"))

	sysdate2=gOutDT(Request("MakeSureDate2"))

elseIf not (ifnull(Request("Sys_PetitionDate1")) and ifnull(Request("Sys_PetitionDate2"))) Then

	min_Year=split(gArrMT(Request("Sys_PetitionDate1")),"-")

	max_year=split(gArrMT(Request("Sys_PetitionDate2")),"-")
	
	sysdate1=gOutDT(Request("Sys_PetitionDate1"))

	sysdate2=gOutDT(Request("Sys_PetitionDate2"))

elseIf not (ifnull(Request("DeallIneDate1")) and ifnull(Request("DeallIneDate2"))) Then

	min_Year=split(gArrMT(Request("DeallIneDate1")),"-")

	max_year=split(gArrMT(Request("DeallIneDate2")),"-")
	
	sysdate1=gOutDT(Request("DeallIneDate1"))

	sysdate2=gOutDT(Request("DeallIneDate2"))

else
	strSQL="select (to_number(to_char(max(IllegalDate),'YYYY'))-1911) nowYear," & _
		"max(IllegalDate) max_year,min(IllegalDate) min_Year" & _
		" from passerBase pb where Exists(select 'Y' from "&BasSQL&" where SN=pb.SN)"

	
	set rs=conn.execute(strSQL)
	If not rs.eof Then

		max_year=split(gArrDT(trim(rs("max_year"))),"-")

		min_Year=split(gArrDT(trim(rs("min_Year"))),"-")

		
	
		sysdate1=year(rs("min_Year"))&"/"&month(rs("min_Year"))&"/"&day(rs("min_Year"))

		sysdate2=year(rs("max_year"))&"/"&month(rs("max_year"))&"/"&day(rs("max_year"))
	End if 
	rs.close

End if 



chkUit="''"
chkUitOrder="''"

If not Ifnull(Request("Sys_MemberStation")) Then
	chkUit="(select UnitName from Unitinfo where Unitid=pb.memberstation)"
	chkUitOrder="(select UnitOrder from Unitinfo where Unitid=pb.memberstation)"
end if


strSQL="select UitName,UnitOrder,IllegalDate_year,sum(sroCnt) sroCnt,sum(Forfeit) Forfeit,sum(CloseCnt) CloseCnt,sum(CloseForFeit) CloseForFeit," & _
		"sum(normalCnt) normalCnt,sum(normalForFeit) normalForFeit,sum(notSendCnt) notSendCnt,sum(notSendForFeit) notSendForFeit," & _
		"sum(SendCnt) SendCnt,sum(SendForFeit) SendForFeit,sum(CreditorCnt) CreditorCnt,sum(CreditorForFeit) CreditorForFeit" & _
		" from (" & _
		"select "&chkUit&" UitName," & _
		""&chkUitOrder&" UnitOrder," & _
		"to_number(to_char(IllegalDate,'YYYY'))-1911 IllegalDate_year," & _
		"1 sroCnt,(forfeit1+nvl(forfeit2,0)) Forfeit," & _
		"(case when billstatus=9 and (select count(1) cnt from PasserPay where CaseCloseDate<=to_date('"&sysdate2&"','YYYY/MM/DD') and billsn=pb.sn)>0 then 1 else 0 end) CloseCnt," & _
		"(case when billstatus=9 and (select count(1) cnt from PasserPay where CaseCloseDate<=to_date('"&sysdate2&"','YYYY/MM/DD') and billsn=pb.sn)>0 then (select nvl(sum(payAmount),0) payAmount from passerpay where billsn=pb.sn) else 0 end) CloseForFeit," & _
		"(case when (billstatus<>9 or (select count(1) cnt from PasserPay where CaseCloseDate<=to_date('"&sysdate2&"','YYYY/MM/DD') and billsn=pb.sn)=0) and DEALLINEDATE >= to_date('"&sysdate2&"','YYYY/MM/DD') then 1 else 0 end) normalCnt," & _
		"(case when (billstatus<>9 or (select count(1) cnt from PasserPay where CaseCloseDate<=to_date('"&sysdate2&"','YYYY/MM/DD') and billsn=pb.sn)=0) and DEALLINEDATE >= to_date('"&sysdate2&"','YYYY/MM/DD') then (forfeit1+nvl(forfeit2,0)) else 0 end) normalForFeit," & _
		"(case when (select count(1) cnt from PasserPay where CaseCloseDate<=to_date('"&sysdate2&"','YYYY/MM/DD') and billsn=pb.sn)=0 and DEALLINEDATE < to_date('"&sysdate2&"','YYYY/MM/DD')" & _
		" and (select count(1) cnta from passersend where billsn=pb.sn) = 0 then 1 else 0 end) notSendCnt," & _
		"(case when (select count(1) cnt from PasserPay where CaseCloseDate<=to_date('"&sysdate2&"','YYYY/MM/DD') and billsn=pb.sn)=0 and DEALLINEDATE < to_date('"&sysdate2&"','YYYY/MM/DD')" & _
		" and (select count(1) cnta from passersend where billsn=pb.sn) = 0 then (forfeit1+nvl(forfeit2,0)) else 0 end) notSendForFeit," & _
		"(case when (select count(1) cnt from PasserPay where CaseCloseDate<=to_date('"&sysdate2&"','YYYY/MM/DD') and billsn=pb.sn)=0 and DEALLINEDATE < to_date('"&sysdate2&"','YYYY/MM/DD') and (select count(1) cnta from PasserCreditor where billsn=pb.sn) = 0" & _
		" and (select count(1) cnta from passersend where billsn=pb.sn) > 0 then 1 else 0 end) SendCnt," & _
		"(case when (select count(1) cnt from PasserPay where CaseCloseDate<=to_date('"&sysdate2&"','YYYY/MM/DD') and billsn=pb.sn)=0 and DEALLINEDATE < to_date('"&sysdate2&"','YYYY/MM/DD') and (select count(1) cnta from PasserCreditor where billsn=pb.sn) = 0" & _
		" and (select count(1) cnta from passersend where billsn=pb.sn) > 0 then (forfeit1+nvl(forfeit2,0)) else 0 end) SendForFeit," & _
		"(case when (select count(1) cnt from PasserPay where CaseCloseDate<=to_date('"&sysdate2&"','YYYY/MM/DD') and billsn=pb.sn)=0 and DEALLINEDATE < to_date('"&sysdate2&"','YYYY/MM/DD') and (select count(1) cnta from PasserCreditor where billsn=pb.sn) > 0 then 1 else 0 end) CreditorCnt," & _
		"(case when (select count(1) cnt from PasserPay where CaseCloseDate<=to_date('"&sysdate2&"','YYYY/MM/DD') and billsn=pb.sn)=0 and DEALLINEDATE < to_date('"&sysdate2&"','YYYY/MM/DD') and (select count(1) cnta from PasserCreditor where billsn=pb.sn) > 0 then (forfeit1+nvl(forfeit2,0)) else 0 end) CreditorForFeit" & _
		" from passerBase pb where Exists(select 'Y' from "&BasSQL&" where SN=pb.SN)" & _
		") sumTable group by UitName,UnitOrder,IllegalDate_year" & _
		" order by UnitOrder,UitName,IllegalDate_year"

set rs=conn.execute(strSQL)

thenPasserCity=""
strUInfo="select * from Apconfigure where ID=31"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then 
	if trim(rsUInfo("value"))<>"" and not isnull(rsUInfo("value")) then
		thenPasserCity=replace(trim(rsUInfo("value")),"台","臺")
	end if
end if 
rsUInfo.close
set rsUInfo=nothing

nowpage=0
chkUit=""

Uit_sroCnt=0:Uit_Forfeit=0:Uit_CloseCnt=0:Uit_CloseForFeit=0:Uit_normalCnt=0:Uit_normalForFeit=0:Uit_notSendCnt=0
Uit_notSendForFeit=0:Uit_SendCnt=0:Uit_SendForFeit=0:Uit_CreditorCnt=0:Uit_CreditorForFeit=0

total_sroCnt=0:total_Forfeit=0:total_CloseCnt=0:total_CloseForFeit=0:total_normalCnt=0:total_normalForFeit=0
total_notSendCnt=0:total_notSendForFeit=0:total_SendCnt=0:total_SendForFeit=0:total_CreditorCnt=0:total_CreditorForFeit=0

If not rs.eof Then chkUit=rs("UitName")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>行政罰鍰收繳情形明細表</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
.font3{
   font-size:10px;
   font-family:新細明體;
   font-weight:900;
}
</style>
</head>
<body>
<%
While not rs.eof
	nowpage=nowpage+1

	If nowpage > 1 Then response.write "<div class=""PageNext"">&nbsp;</div>"
%>
<table width="100%" border="0">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td align="center"><strong><%=thenPasserCity%>政府暨所屬機關<%=max_year(0)%>年度行政罰鍰收繳情形明細表</strong></td>
			</tr>
			<tr>
				<td align="center"><strong>自<%=min_Year(0)%>年<%=min_Year(1)%>月<%=min_Year(2)%>日起至<%=max_year(0)%>年<%=max_year(1)%>月<%=max_year(2)%>日止</strong></td>
			</tr>
			<tr>
				<td class="font3" align="right">單&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;位：件、千元</td>
			</tr>
			<tr>
				<td class="font3" align="right">印表日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td class="font3" align="right">頁
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								數：<%=right("0000"&nowpage,2)%>
				</td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="0">
				<tr>
					<td class="font3" align="center" rowspan=3>機關別</td>
					<td class="font3" align="center" rowspan=3 nowrap>年度</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>年度裁處罰鍰</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>已收繳及註銷數</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>期未餘額</td>
					<td class="font3" align="center" colspan=4 nowrap>執行率</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>未逾繳款期限</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>已逾繳款期限未移<br>送執行分署執行</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>已逾繳款期限已移<br>送執行分署執行</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>取得債權憑證數</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>行政救濟未繳數</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>其它數</td>
				</tr>
				<tr>
					<td class="font3" align="center" colspan=2 nowrap>清理比率</td>
					<td class="font3" align="center" colspan=2 nowrap>已移送執行比率</td>
				</tr>
				<tr>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數％</td>
					<td class="font3" nowrap>金額％</td>
					<td class="font3" nowrap>件數％</td>
					<td class="font3" nowrap>金額％</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
				</tr>
				<%
				For i = 1 to 10
					If rs.eof Then exit For 

					If chkUit <> rs("UitName") Then exit For 

					Uit_sroCnt=Uit_sroCnt+cdbl(rs("sroCnt"))
					Uit_Forfeit=Uit_Forfeit+cdbl(rs("Forfeit"))
					Uit_CloseCnt=Uit_CloseCnt+cdbl(rs("CloseCnt"))
					Uit_CloseForFeit=Uit_CloseForFeit+cdbl(rs("CloseForFeit"))
					Uit_normalCnt=Uit_normalCnt+cdbl(rs("normalCnt"))
					Uit_normalForFeit=Uit_normalForFeit+cdbl(rs("normalForFeit"))
					Uit_notSendCnt=Uit_notSendCnt+cdbl(rs("notSendCnt"))
					Uit_notSendForFeit=Uit_notSendForFeit+cdbl(rs("notSendForFeit"))
					Uit_SendCnt=Uit_SendCnt+cdbl(rs("SendCnt"))
					Uit_SendForFeit=Uit_SendForFeit+cdbl(rs("SendForFeit"))
					Uit_CreditorCnt=Uit_CreditorCnt+cdbl(rs("CreditorCnt"))
					Uit_CreditorForFeit=Uit_CreditorForFeit+cdbl(rs("CreditorForFeit"))

					total_sroCnt=total_sroCnt+cdbl(rs("sroCnt"))
					total_Forfeit=total_Forfeit+cdbl(rs("Forfeit"))
					total_CloseCnt=total_CloseCnt+cdbl(rs("CloseCnt"))
					total_CloseForFeit=total_CloseForFeit+cdbl(rs("CloseForFeit"))
					total_normalCnt=total_normalCnt+cdbl(rs("normalCnt"))
					total_normalForFeit=total_normalForFeit+cdbl(rs("normalForFeit"))
					total_notSendCnt=total_notSendCnt+cdbl(rs("notSendCnt"))
					total_notSendForFeit=total_notSendForFeit+cdbl(rs("notSendForFeit"))
					total_SendCnt=total_SendCnt+cdbl(rs("SendCnt"))
					total_SendForFeit=total_SendForFeit+cdbl(rs("SendForFeit"))
					total_CreditorCnt=total_CreditorCnt+cdbl(rs("CreditorCnt"))
					total_CreditorForFeit=total_CreditorForFeit+cdbl(rs("CreditorForFeit"))

					Response.Write "<tr>"
					Response.Write "<td class=""font3"" nowrap>"
					Response.Write rs("UitName")
					Response.Write "&nbsp;</td>"

					Response.Write "<td class=""font3"">"
					Response.Write rs("IllegalDate_year")
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write rs("sroCnt")
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write rs("Forfeit")
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write rs("CloseCnt")
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write rs("CloseForFeit")
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write cdbl(rs("sroCnt"))-cdbl(rs("CloseCnt"))
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write cdbl(rs("Forfeit"))-cdbl(rs("CloseForFeit"))
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write fix(cdbl(rs("CloseCnt"))/cdbl(rs("sroCnt"))*100)&"%"
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write fix(cdbl(rs("CloseForFeit"))/cdbl(rs("Forfeit"))*100)&"%"
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write fix((cdbl(rs("CloseCnt"))+cdbl(rs("normalCnt"))+cdbl(rs("SendCnt"))+cdbl(rs("CreditorCnt")))/cdbl(rs("sroCnt"))*100)&"%"
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write fix((cdbl(rs("CloseForFeit"))+cdbl(rs("normalForFeit"))+cdbl(rs("SendForFeit"))+cdbl(rs("CreditorForFeit")))/cdbl(rs("Forfeit"))*100)&"%"
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write rs("normalCnt")
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write rs("normalForFeit")
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write rs("notSendCnt")
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write rs("notSendForFeit")
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write rs("SendCnt")
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write rs("SendForFeit")
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write rs("CreditorCnt")
					Response.Write "</td>"

					Response.Write "<td class=""font3"">"
					Response.Write rs("CreditorForFeit")
					Response.Write "</td>"


					Response.Write "<td>&nbsp;</td>"
					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"
					Response.Write "<td>&nbsp;</td>"


					Response.Write "</tr>"


					rs.movenext
				Next

				If not rs.eof Then
					If chkUit <> rs("UitName") Then
						chkUit=rs("UitName")

						Response.Write "<tr>"
						Response.Write "<td class=""font3"" colspan=2>小　　計</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_sroCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_Forfeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_CloseCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_CloseForFeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write cdbl(Uit_sroCnt)-cdbl(Uit_CloseCnt)
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write cdbl(Uit_Forfeit)-cdbl(Uit_CloseForFeit)
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write fix(cdbl(Uit_CloseCnt)/cdbl(Uit_sroCnt)*100)&"%"
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write fix(cdbl(Uit_CloseForFeit)/cdbl(Uit_Forfeit)*100)&"%"
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write fix((cdbl(Uit_CloseCnt)+cdbl(Uit_normalCnt)+cdbl(Uit_SendCnt)+cdbl(Uit_CreditorCnt))/cdbl(Uit_sroCnt)*100)&"%"
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write fix((cdbl(Uit_CloseForFeit)+cdbl(Uit_normalForFeit)+cdbl(Uit_SendForFeit)+cdbl(Uit_CreditorForFeit))/cdbl(Uit_Forfeit)*100)&"%"
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_normalCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_normalForFeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_notSendCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_notSendForFeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_SendCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_SendForFeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_CreditorCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_CreditorForFeit
						Response.Write "</td>"


						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td>&nbsp;</td>"

						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "</tr>"

						
						Uit_sroCnt=0:Uit_Forfeit=0:Uit_CloseCnt=0:Uit_CloseForFeit=0:Uit_normalCnt=0:Uit_normalForFeit=0:Uit_notSendCnt=0
						Uit_notSendForFeit=0:Uit_SendCnt=0:Uit_SendForFeit=0:Uit_CreditorCnt=0:Uit_CreditorForFeit=0
						
					end If 
				else
					Response.Write "<tr>"
						Response.Write "<td class=""font3"" colspan=2>小　　計</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_sroCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_Forfeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_CloseCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_CloseForFeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write cdbl(Uit_sroCnt)-cdbl(Uit_CloseCnt)
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write cdbl(Uit_Forfeit)-cdbl(Uit_CloseForFeit)
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write fix(cdbl(Uit_CloseCnt)/cdbl(Uit_sroCnt)*100)&"%"
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write fix(cdbl(Uit_CloseForFeit)/cdbl(Uit_Forfeit)*100)&"%"
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write fix((cdbl(Uit_CloseCnt)+cdbl(Uit_normalCnt)+cdbl(Uit_SendCnt)+cdbl(Uit_CreditorCnt))/cdbl(Uit_sroCnt)*100)&"%"
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write fix((cdbl(Uit_CloseForFeit)+cdbl(Uit_normalForFeit)+cdbl(Uit_SendForFeit)+cdbl(Uit_CreditorForFeit))/cdbl(Uit_Forfeit)*100)&"%"
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_normalCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_normalForFeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_notSendCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_notSendForFeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_SendCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_SendForFeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_CreditorCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write Uit_CreditorForFeit
						Response.Write "</td>"


						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td>&nbsp;</td>"

						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "</tr>"

						Response.Write "<tr>"
						Response.Write "<td class=""font3"" colspan=2>總　　計</td>"

						Response.Write "<td class=""font3"">"
						Response.Write total_sroCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write total_Forfeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write total_CloseCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write total_CloseForFeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write cdbl(total_sroCnt)-cdbl(total_CloseCnt)
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write cdbl(total_Forfeit)-cdbl(total_CloseForFeit)
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write fix(cdbl(total_CloseCnt)/cdbl(total_sroCnt)*100)&"%"
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write fix(cdbl(total_CloseForFeit)/cdbl(total_Forfeit)*100)&"%"
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write fix((cdbl(total_CloseCnt)+cdbl(total_normalCnt)+cdbl(total_SendCnt)+cdbl(total_CreditorCnt))/cdbl(total_sroCnt)*100)&"%"
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write fix((cdbl(total_CloseForFeit)+cdbl(total_normalForFeit)+cdbl(total_SendForFeit)+cdbl(total_CreditorForFeit))/cdbl(total_Forfeit)*100)&"%"
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write total_normalCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write total_normalForFeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write total_notSendCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write total_notSendForFeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write total_SendCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write total_SendForFeit
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write total_CreditorCnt
						Response.Write "</td>"

						Response.Write "<td class=""font3"">"
						Response.Write total_CreditorForFeit
						Response.Write "</td>"


						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td>&nbsp;</td>"

						Response.Write "<td>&nbsp;</td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "<tr>"
				End if 
			%>
			</table>
		</td>
	</tr>
	<tr>
		<td class="font3" style="font-size:12px;">
			承辦人：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%If chkUitOrder="''" Then
				Response.Write "隊長：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			else
				Response.Write "組長：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			End if %>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			會計：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			機關首長：
		</td>
	</tr>
	<tr>
		<td class="font3">
			備註
		</td>
	</tr>
	<tr>
		<td class="font3">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			1.清理比率=已收繳數級註銷數&nbsp;÷&nbsp;年度處罰鍰<br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			2.已移送執行比率=（已收繳數及註銷數+未逾繳款期限+已逾繳款期限已移送執行分署執行+取得債權憑證數+行政救濟未繳數+其他數)&nbsp;÷&nbsp;年度處罰鍰
		</td>
	</tr>
</table>
<%
wend
rs.close

If nowpage = 0 Then

strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsuit=conn.execute(strSQL)

chkUit=rsuit("UnitName")

max_year=split(gArrDT(date),"-")

min_Year=split(gArrDT(date),"-")

rsuit.close
%>
<table width="100%" border="0">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td align="center"><strong><%=thenPasserCity%>政府暨所屬機關<%=nowYear%>年度行政罰鍰收繳情形明細表</strong></td>
			</tr>
			<tr>
				<td align="center"><strong>自<%=min_Year(0)%>年<%=min_Year(1)%>月<%=min_Year(2)%>日起至<%=max_year(0)%>年<%=max_year(1)%>月<%=max_year(2)%>日止</strong></td>
			</tr>
			<tr>
				<td class="font3" align="right">單&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;位：件、千元</td>
			</tr>
			<tr>
				<td class="font3" align="right">印表日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td class="font3" align="right">頁
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								數：<%=right("0000"&nowpage,2)%>
				</td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="0">
				<tr>
					<td class="font3" align="center" rowspan=3>機關別</td>
					<td class="font3" align="center" rowspan=3 nowrap>年度</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>年度裁處罰鍰</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>已收繳及註銷數</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>期未餘額</td>
					<td class="font3" align="center" colspan=4 nowrap>執行率</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>未逾繳款期限</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>已逾繳款期限未移<br>送執行分署執行</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>已逾繳款期限已移<br>送執行分署執行</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>取得債權憑證數</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>行政救濟未繳數</td>
					<td class="font3" align="center" rowspan=2 colspan=2 nowrap>其它數</td>
				</tr>
				<tr>
					<td class="font3" align="center" colspan=2 nowrap>清理比率</td>
					<td class="font3" align="center" colspan=2 nowrap>已移送執行比率</td>
				</tr>
				<tr>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數％</td>
					<td class="font3" nowrap>金額％</td>
					<td class="font3" nowrap>件數％</td>
					<td class="font3" nowrap>金額％</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
					<td class="font3" nowrap>件數</td>
					<td class="font3" nowrap>金額</td>
				</tr>
				<%
				For i = 1 to 3

					Response.Write "<tr>"
					Response.Write "<td class=""font3"" nowrap>"
					Response.Write chkUit
					Response.Write "</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"
					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"
					Response.Write "<td>&nbsp;</td>"


					Response.Write "</tr>"


				Next

				
					Response.Write "<tr>"
					Response.Write "<td class=""font3"" colspan=2>小　　計</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"


					Response.Write "<td>&nbsp;</td>"
					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"
					Response.Write "<td>&nbsp;</td>"
					Response.Write "</tr>"

					Response.Write "<tr>"
					Response.Write "<td class=""font3"" colspan=2>總　　計</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"


					Response.Write "<td>&nbsp;</td>"
					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"
					Response.Write "<td>&nbsp;</td>"
					Response.Write "<tr>"

			%>
			</table>
		</td>
	</tr>
	<tr>
		<td class="font3" style="font-size:12px;">
			承辦人：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%If chkUitOrder="''" Then
				Response.Write "隊長：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			else
				Response.Write "組長：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			End if %>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			會計：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			機關首長：
		</td>
	</tr>
	<tr>
		<td class="font3">
			備註
		</td>
	</tr>
	<tr>
		<td class="font3">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			1.清理比率=已收繳數級註銷數&nbsp;÷&nbsp;年度處罰鍰<br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			2.已移送執行比率=（已收繳數及註銷數+未逾繳款期限+已逾繳款期限已移送執行分署執行+取得債權憑證數+行政救濟未繳數+其他數)&nbsp;÷&nbsp;年度處罰鍰
		</td>
	</tr>
</table>
<%
End if 
%>
</body>
</html>
<%
conn.close
set conn=nothing
%>