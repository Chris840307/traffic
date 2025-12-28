<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%

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

strSQL="select count(1) cnt" & _
		" from passerBase pb where Exists(select 'Y' from "&BasSQL&" where SN=pb.SN)" & _
		" and Exists(select 'Y' from passersendDetail where billsn=pb.sn" & _
		" and not Exists(select 'Y' from PasserCreditor where senddetailsn=passersendDetail.sn))"


set rs=conn.execute(strSQL)

pageCnt=fix(cdbl(rs("cnt"))/5+0.99999999999999)
rs.close

strSQL="select (select UnitName from Unitinfo where Unitid=pb.memberstation) UitName," & _
		"(select UnitOrder from Unitinfo where Unitid=pb.memberstation) UnitOrder," & _
		"(select ChName from MemberData where MemberID=pb.RecordMemberID) UitChName," & _
		"(select opengovNumber from PasserSend where billsn=pb.SN) opengovNumber," & _
		"BillNo,DriverID,BILLFILLDATE,DEALLINEDATE,Driver,DriverAddress,Rule1,Rule2,(forfeit1+nvl(forfeit2,0)) Forfeit1" & _
		" from passerBase pb where Exists(select 'Y' from "&BasSQL&" where SN=pb.SN)" & _
		" and Exists(select 'Y' from passersendDetail where billsn=pb.sn" & _
		" and not Exists(select 'Y' from PasserCreditor where senddetailsn=passersendDetail.sn))" & _
		" order by UnitOrder,UitName"

set rs=conn.execute(strSQL)

nowpage=0
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>債權憑證準備再移送清冊</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>
<body>
<%
While not rs.eof
	nowpage=nowpage+1

	chkUit=rs("UitName")

	rs_UitChName="":rs_BillNo="":rs_DriverID="":rs_BILLFILLDATE="":rs_DEALLINEDATE=""
	rs_Driver="":rs_DriverAddress="":rs_Forfeit1="":rs_opengovNumber=""

	If nowpage > 1 Then response.write "<div class=""PageNext"">&nbsp;</div>"
%>
<table width="100%" border="0">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td colspan="2" align="center"><strong><%=thenPasserCity%>政府暨所屬機關</strong></td>
			</tr>
			<tr>
				<td colspan="2" align="center"><strong>債權憑證準備再移送清冊（第一聯）</strong></td>
			</tr>
			<tr>
				<td align="left">機關名稱：<%=chkUit%></td>
				<td align="right">列印日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td colspan="2" align="right">頁&nbsp;&nbsp;&nbsp;&nbsp;數：<%=nowpage%></td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="0">
				<tr>
					<td>承辦人</td>
					<td>罰鍰單號<br>身分證號</td>
					<td>開立日期<br>繳款期限</td>
					<td>受處罰人<br>地址</td>
					<td>罰單金額</td>
					<td>執行名義文號</td>
				</tr>
				<%
				For i = 1 to 5
					If rs.eof Then exit For 

					If chkUit <> rs("UitName") Then exit For 

					If not ifnull(rs_UitChName) Then
						rs_UitChName=rs_UitChName&"@"
						rs_BillNo=rs_BillNo&"@"
						rs_DriverID=rs_DriverID&"@"
						rs_BILLFILLDATE=rs_BILLFILLDATE&"@"
						rs_DEALLINEDATE=rs_DEALLINEDATE&"@"
						rs_Driver=rs_Driver&"@"
						rs_DriverAddress=rs_DriverAddress&"@"
						rs_Forfeit1=rs_Forfeit1&"@"
						rs_opengovNumber=rs_opengovNumber&"@"
						
					end If 

					rs_UitChName=rs_UitChName&rs("UitChName")
					rs_BillNo=rs_BillNo&rs("BillNo")
					rs_DriverID=rs_DriverID&rs("DriverID")
					rs_BILLFILLDATE=rs_BILLFILLDATE&rs("BILLFILLDATE")
					rs_DEALLINEDATE=rs_DEALLINEDATE&rs("DEALLINEDATE")
					rs_Driver=rs_Driver&rs("Driver")
					rs_DriverAddress=rs_DriverAddress&rs("DriverAddress")
					rs_Forfeit1=rs_Forfeit1&rs("Forfeit1")
					rs_opengovNumber=rs_opengovNumber&rs("opengovNumber")
					
					Response.Write "<tr>"
					Response.Write "<td>"
					Response.Write rs("UitChName")
					Response.Write "</td>"

					Response.Write "<td>"
					Response.Write rs("BillNo")
					If not ifnull(rs("DriverID")) Then Response.Write "<br>"&rs("DriverID")
					Response.Write "</td>"

					Response.Write "<td>"
					Response.Write gInitDT(rs("BILLFILLDATE"))
					If not ifnull(rs("DEALLINEDATE")) Then Response.Write "<br>"&gInitDT(rs("DEALLINEDATE"))
					Response.Write "</td>"

					Response.Write "<td>"
					Response.Write rs("Driver")
					If not ifnull(rs("DriverAddress")) Then Response.Write "<br>"&rs("DriverAddress")
					Response.Write "</td>"

					Response.Write "<td>"
					Response.Write rs("Forfeit1")
					Response.Write "</td>"

					Response.Write "<td>"
					Response.Write rs("opengovNumber")
					Response.Write "&nbsp;</td>"

					Response.Write "</tr>"


					rs.movenext
				Next
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td style="font-size:12px;">
			備註
		</td>
	</tr>
	<tr>
		<td style="font-size:12px;">
			1.本案業經○年○月○日簽奉 縣長核可在案。
		</td>
	</tr>
	<tr>
		<td style="font-size:12px;">
			2.本表設計二聯
		</td>
	</tr>
	<tr>
		<td style="font-size:12px;">
			&nbsp;&nbsp;&nbsp;&nbsp;（1）第一聯：主要是提供給各單位由系統自動產生後，本府各單位逕送為民服務中心，各機關則函送地方稅務局，以統一表格簡化行政作業程序。
		</td>
	</tr>
	<tr>
		<td style="font-size:12px;">
			&nbsp;&nbsp;&nbsp;&nbsp;（2）第二聯：主要是經由地方稅務局查明確認並蓋章後回復各查詢單位，各承辦人員於接獲地方稅務局回復資料，檢視各義務人的財產所得資料是
	</tr>
	<tr>
		<td style="font-size:12px;">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			否有供再執行的可能性，分析結果應送該單位業務主管及機關首長的核可。
		</td>
	</tr>
	<tr>
		<td style="font-size:12px;">
			主辦業務人員：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			業務主管：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			機關首長：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
</table>
<%
	If not ifnull(rs_UitChName) then

		response.write "<div class=""PageNext"">&nbsp;</div>"

		rs_UitChName=split(rs_UitChName,"@")
		rs_BillNo=split(rs_BillNo,"@")
		rs_DriverID=split(rs_DriverID,"@")
		rs_BILLFILLDATE=split(rs_BILLFILLDATE,"@")
		rs_DEALLINEDATE=split(rs_DEALLINEDATE,"@")
		rs_Driver=split(rs_Driver,"@")
		rs_DriverAddress=split(rs_DriverAddress,"@")
		rs_Forfeit1=split(rs_Forfeit1,"@")
		rs_opengovNumber=split(rs_opengovNumber&" ","@")	
	
	%>
		<table width="100%" border="0">
			<tr><td>
				<table width="100%" border="0" cellpadding="4" cellspacing="1">
					<tr>
						<td colspan="2" align="center"><strong><%=thenPasserCity%>政府暨所屬機關</strong></td>
					</tr>
					<tr>
						<td colspan="2" align="center"><strong>債權憑證準備再移送清冊（第二聯）</strong></td>
					</tr>
					<tr>
						<td align="left">機關名稱：<%=chkUit%></td>
						<td align="right">列印日期：<%=gInitDt(date)%></td>
					</tr>
					<tr>
						<td colspan="2" align="right">頁&nbsp;&nbsp;&nbsp;&nbsp;數：<%=nowpage%></td>
					</tr>
				</table>
			</td></tr>
			<tr>
				<td>
					<table width="100%" border="1" cellpadding="4" cellspacing="0">
						<tr>
							<td style="font-size:14px;" rowspan=3>承辦人</td>
							<td style="font-size:14px;" rowspan=3 nowrap>罰鍰單號<br>身分證號</td>
							<td style="font-size:14px;" rowspan=3 nowrap>開立日期<br>繳款期限</td>
							<td style="font-size:14px;" rowspan=3>受處罰人<br>地址</td>
							<td style="font-size:14px;" rowspan=3 nowrap>罰單金額</td>
							<td style="font-size:14px;" rowspan=3 nowrap>執行名義文號</td>
							<td style="font-size:14px;" colspan=5 align="center">查調及辦理情形</td>
						</tr>
						<tr>
							<td style="font-size:14px;" colspan=2>有無財產、所得資料</td>
							<td style="font-size:14px;" colspan=2 align="center">辦理情形</td>
							<td style="font-size:14px;" align="center" rowspan=2 nowrap>承辦人核章</td>
						</tr>
						<tr>
							<td style="font-size:14px;">無</td>
							<td style="font-size:14px;">有</td>
							<td style="font-size:14px;">無移送價值繼續列管</td>
							<td style="font-size:14px;">擬辦理移送執行</td>
						</tr>
						<%
						For j = 0 to Ubound(rs_UitChName)

							Response.Write "<tr>"
							Response.Write "<td nowrap>"
							Response.Write rs_UitChName(j)
							Response.Write "</td>"

							Response.Write "<td>"
							Response.Write rs_BillNo(j)
							If not ifnull(rs_DriverID(j)) Then Response.Write "<br>"&rs_DriverID(j)
							Response.Write "</td>"

							Response.Write "<td>"
							Response.Write gInitDT(rs_BILLFILLDATE(j))
							If not ifnull(rs_DEALLINEDATE(j)) Then Response.Write "<br>"&gInitDT(rs_DEALLINEDATE(j))
							Response.Write "</td>"

							Response.Write "<td>"
							Response.Write rs_Driver(j)
							If not ifnull(rs_DriverAddress(j)) Then Response.Write "<br>"&rs_DriverAddress(j)
							Response.Write "</td>"

							Response.Write "<td>"
							Response.Write rs_Forfeit1(j)
							Response.Write "</td>"

							Response.Write "<td>"
							Response.Write rs_opengovNumber(j)
							Response.Write "&nbsp;</td>"

							Response.Write "<td>&nbsp;</td>"
							Response.Write "<td>&nbsp;</td>"
							Response.Write "<td>&nbsp;</td>"
							Response.Write "<td>&nbsp;</td>"
							Response.Write "<td>&nbsp;</td>"

							Response.Write "</tr>"

						Next
						%>
					<tr>
						<td colspan=6>地方稅務局核章</td>
						<td colspan=4>
							填表人：<br><br>
							業務主管：<br><br>
							機關首長：<br>
						</td>
						<td>&nbsp;</td>
					</tr>
					</table>
				</td>
			</tr>			
		</table>
<%
	end if
wend
rs.close

If nowpage = 0 Then

strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsuit=conn.execute(strSQL)

chkUit=rsuit("UnitName")

rsuit.close
%>
<table width="100%" border="0">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td colspan="2" align="center"><strong><%=thenPasserCity%>政府暨所屬機關</strong></td>
			</tr>
			<tr>
				<td colspan="2" align="center"><strong>債權憑證準備再移送清冊（第一聯）</strong></td>
			</tr>
			<tr>
				<td align="left">機關名稱：<%=chkUit%></td>
				<td align="right">列印日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td colspan="2" align="right">頁&nbsp;&nbsp;&nbsp;&nbsp;數：<%=nowpage%></td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="0">
				<tr>
					<td>承辦人</td>
					<td>罰鍰單號<br>身分證號</td>
					<td>開立日期<br>繳款期限</td>
					<td>受處罰人<br>地址</td>
					<td>罰單金額</td>
					<td>執行名義文號</td>
				</tr>
				<%
				For i = 0 to 2
					Response.Write "<tr>"
					Response.Write "<td>"
					Response.Write chkUit
					Response.Write "</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "<td>&nbsp;</td>"

					Response.Write "</tr>"
				Next
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td style="font-size:12px;">
			備註
		</td>
	</tr>
	<tr>
		<td style="font-size:12px;">
			1.本案業經○年○月○日簽奉 縣長核可在案。
		</td>
	</tr>
	<tr>
		<td style="font-size:12px;">
			2.本表設計二聯
		</td>
	</tr>
	<tr>
		<td style="font-size:12px;">
			&nbsp;&nbsp;&nbsp;&nbsp;（1）第一聯：主要是提供給各單位由系統自動產生後，本府各單位逕送為民服務中心，各機關則函送地方稅務局，以統一表格簡化行政作業程序。
		</td>
	</tr>
	<tr>
		<td style="font-size:12px;">
			&nbsp;&nbsp;&nbsp;&nbsp;（2）第二聯：主要是經由地方稅務局查明確認並蓋章後回復各查詢單位，各承辦人員於接獲地方稅務局回復資料，檢視各義務人的財產所得資料是
	</tr>
	<tr>
		<td style="font-size:12px;">
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			否有供再執行的可能性，分析結果應送該單位業務主管及機關首長的核可。
		</td>
	</tr>
	<tr>
		<td style="font-size:12px;">
			主辦業務人員：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			業務主管：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			機關首長：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		</td>
	</tr>
</table>

<div class="PageNext">&nbsp;</div>

<table width="100%" border="0">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td colspan="2" align="center"><strong><%=thenPasserCity%>政府暨所屬機關</strong></td>
			</tr>
			<tr>
				<td colspan="2" align="center"><strong>債權憑證準備再移送清冊（第二聯）</strong></td>
			</tr>
			<tr>
				<td align="left">機關名稱：<%=chkUit%></td>
				<td align="right">列印日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td colspan="2" align="right">頁&nbsp;&nbsp;&nbsp;&nbsp;數：<%=nowpage%></td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="0">
				<tr>
					<td style="font-size:14px;" rowspan=3>承辦人</td>
					<td style="font-size:14px;" rowspan=3 nowrap>罰鍰單號<br>身分證號</td>
					<td style="font-size:14px;" rowspan=3 nowrap>開立日期<br>繳款期限</td>
					<td style="font-size:14px;" rowspan=3>受處罰人<br>地址</td>
					<td style="font-size:14px;" rowspan=3 nowrap>罰單金額</td>
					<td style="font-size:14px;" rowspan=3 nowrap>執行名義文號</td>
					<td style="font-size:14px;" colspan=5 align="center">查調及辦理情形</td>
				</tr>
				<tr>
					<td style="font-size:14px;" colspan=2>有無財產、所得資料</td>
					<td style="font-size:14px;" colspan=2 align="center">辦理情形</td>
					<td style="font-size:14px;" align="center" rowspan=2 nowrap>承辦人核章</td>
				</tr>
				<tr>
					<td style="font-size:14px;">無</td>
					<td style="font-size:14px;">有</td>
					<td style="font-size:14px;">無移送價值繼續列管</td>
					<td style="font-size:14px;">擬辦理移送執行</td>
				</tr>
				<%
				For j = 0 to 2

					Response.Write "<tr>"
					Response.Write "<td nowrap>"
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

					Response.Write "</tr>"

				Next
				%>
			<tr>
				<td colspan=6>地方稅務局核章</td>
				<td colspan=4>
					填表人：<br><br>
					業務主管：<br><br>
					機關首長：<br>
				</td>
				<td>&nbsp;</td>
			</tr>
			</table>
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