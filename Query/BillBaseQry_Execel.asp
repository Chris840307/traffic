<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_資料交換紀錄.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
	'檢查是否可進入本系統
	'AuthorityCheck(234)
Server.ScriptTimeout = 16800
Response.flush
	strwhere=Session("PrintCarDataSQL")
if trim(request("WorkType"))="1" then
	strSQL="select a.SN,a.IllegalDate,a.CarSimpleID,a.BillMem1,a.BillMem2,a.BillMem3,a.BillMem4,b.ChName,a.BillTypeID,a.BillNo,a.CarNo,a.Driver,a.DriverID,a.IllegalAddress,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.ForFeit1,a.ForFeit2,a.ForFeit3,a.ForFeit4,a.BillUnitID,a.BillStatus,a.RecordStateID,a.RecordDate,a.RecordMemberID from BillBase a,MemberData b where a.RecordMemberID=b.MemberID(+) "&strwhere&" order by a.RecordDate"
elseif trim(request("WorkType"))="2" then
	strSQL="select a.SN,a.IllegalDate,a.BillMem1,a.BillMem2,a.BillMem3,a.BillMem4,b.ChName,a.BillTypeID" &_
			",a.BillNo,a.CarNo,a.CarSimpleID,a.Driver,a.DriverID,a.IllegalAddress,a.Rule1,a.Rule2,a.Rule3" &_
			",a.Rule4,a.ForFeit1,a.ForFeit2,a.ForFeit3,a.ForFeit4,a.BillUnitID,a.BillStatus" &_
			",a.RecordStateID,a.RecordDate,a.RecordMemberID,c.UserMarkDate from BillBase a,MemberData b" &_
			",BillMailHistory c where a.RecordMemberID=b.MemberID(+) and c.BillSN=a.SN "&strwhere&" order by c.UserMarkDate"
elseif trim(request("WorkType"))="3" then
	strSQL="select a.SN,a.IllegalDate,a.CarSimpleID,a.BillMem1,a.BillMem2,a.BillMem3,a.BillMem4,b.ChName,a.BillTypeID,a.BillNo,a.CarNo,a.Driver,a.DriverID,a.IllegalAddress,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.ForFeit1,a.ForFeit2,a.ForFeit3,a.ForFeit4,a.BillUnitID,a.BillStatus,a.RecordStateID,a.RecordDate,a.RecordMemberID,c.UserMarkResonID,c.StoreAndSendReturnResonID,c.UserMarkDate,c.StoreAndSendMailReturnDate from BillBase a,MemberData b,BillMailHistory c where a.RecordMemberID=b.MemberID and c.BillSN=a.SN and c.UserMarkResonID in ('5','6','7','T')"&strwhere&" order by c.UserMarkDate"
elseif trim(request("WorkType"))="4" then
	strwhere=Session("DciOpenGovToExcelSession")
	strSQL="select a.SN,a.IllegalDate,a.CarSimpleID,a.BillMem1,a.BillMem2,a.BillMem3,a.BillMem4,b.ChName,a.BillTypeID,a.BillNo,a.CarNo,a.Driver,a.DriverID,a.IllegalAddress,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.ForFeit1,a.ForFeit2,a.ForFeit3,a.ForFeit4,a.BillUnitID,a.BillStatus,a.RecordStateID,a.RecordDate,a.RecordMemberID,c.UserMarkResonID,c.StoreAndSendReturnResonID,c.UserMarkDate,c.StoreAndSendMailReturnDate from BillBase a,MemberData b,BillMailHistory c where a.RecordMemberID=b.MemberID and c.BillSN=a.SN and c.UserMarkResonID in ('1','2','3','4','8','M','K','L','O','P','Q')"&strwhere&" order by c.UserMarkDate"
elseif trim(request("WorkType"))="5" then
	strSQL="select a.SN,a.IllegalDate,a.CarSimpleID,a.BillMem1,a.BillMem2,a.BillMem3,a.BillMem4,b.ChName,a.BillTypeID" &_
			",a.BillNo,a.CarNo,a.Driver,a.DriverID,a.IllegalAddress,a.Rule1,a.Rule2,a.Rule3" &_
			",a.Rule4,a.ForFeit1,a.ForFeit2,a.ForFeit3,a.ForFeit4,a.BillUnitID,a.BillStatus" &_
			",a.RecordStateID,a.RecordDate,a.RecordMemberID from BillBase a,MemberData b" &_
			",BillMailHistory c where a.RecordMemberID=b.MemberID(+) and c.BillSN=a.SN "&strwhere&" order by c.UserMarkDate"
end if
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)

	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單查詢</title>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
</head>
<body>
<table width="100%" border="1">
	<tr>
		<td align="center"><strong>舉發單紀錄</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="1">
				<tr>
					<th width="10%">違規日期</th>
					<th width="8%">舉發員警</th>
					<th width="8%">舉發單號</th>
					<th width="7%">車號</th>
					<th width="7%">簡式車種</th>
					<th width="7%">詳細車種</th>
					<th width="7%">類別</th>
					<th width="8%">駕駛人</th>
					<th width="14%">違規地點</th>
					<th width="10%">法條</th>
					<!-- <th width="6%">罰款</th> -->
					<th width="8%">DCI</th>
				</tr>
				<%
					If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
						chname="":chRule="":ForFeit=""
						if rsfound("BillMem1")<>"" then	chname=rsfound("BillMem1")
						if rsfound("BillMem2")<>"" then	chname=chname&"/"&rsfound("BillMem2")
						if rsfound("BillMem3")<>"" then	chname=chname&"/"&rsfound("BillMem3")
						if rsfound("BillMem4")<>"" then	chname=chname&"/"&rsfound("BillMem4")
						if rsfound("Rule1")<>"" then chRule=rsfound("Rule1")
						if rsfound("Rule2")<>"" then chRule=chRule&"/"&rsfound("Rule2")
						if rsfound("Rule3")<>"" then chRule=chRule&"/"&rsfound("Rule3")
						if rsfound("Rule4")<>"" then chRule=chRule&"/"&rsfound("Rule4")
						if rsfound("ForFeit1")<>"" then ForFeit=rsfound("ForFeit1")
						if rsfound("ForFeit2")<>"" then ForFeit=ForFeit&"/"&rsfound("ForFeit2")
						if rsfound("ForFeit3")<>"" then ForFeit=ForFeit&"/"&rsfound("ForFeit3")
						if rsfound("ForFeit4")<>"" then ForFeit=ForFeit&"/"&rsfound("ForFeit4")

						response.write "<tr bgcolor='#FFFFFF' align='center' "
						response.write ">"
						response.write "<td width='10%'>"&gInitDT(rsfound("IllegalDate"))&"</td>"
						response.write "<td width='8%'>"&chname&"</td>"
						response.write "<td width='8%'>"&rsfound("BillNo")&"</td>"
						response.write "<td width='7%'>"&rsfound("CarNo")&"</td>"
						response.write "<td width='7%'>"
						If Trim(rsfound("CarSimpleID"))="1" Then
							response.write "1汽車"
						elseIf Trim(rsfound("CarSimpleID"))="2" Then
							response.write "2拖車"
						elseIf Trim(rsfound("CarSimpleID"))="3" Then
							response.write "3重機"
						ElseIf Trim(rsfound("CarSimpleID"))="4" Then
							response.write "4輕機"
						Else
							response.write "6臨時車牌"
						End if
						response.write "</td>"
						CarType=""
						strCType="select b.Content from BillBaseDCIReturn a,DCIcode b where b.TypeID=5 and b.ID=a.DCIReturnCarType and ((a.BillNo='"&trim(rsfound("BillNo"))&"' and a.CarNo='"&trim(rsfound("CarNo"))&"') or (a.BillNo is null and a.CarNo='"&trim(rsfound("CarNo"))&"'))"
						set rsCType=conn.execute(strCType)
						if not rsCType.eof then
							CarType=trim(rsCType("Content"))
						end if
						rsCType.close
						set rsCType=nothing
						response.write "<td width='5%'>"&CarType&"</td>"
						response.write "<td width='7%'>"
					strBTypeVal="select Content from DCIcode where TypeID=2 and ID='"&trim(rsfound("BillTypeID"))&"'"
					set rsBTypeVal=conn.execute(strBTypeVal)
					if not rsBTypeVal.eof then
						response.write rsBTypeVal("Content")
					end if
					rsBTypeVal.close
					set rsBTypeVal=nothing
						response.write "</td>"
						response.write "<td width='8%'>"&rsfound("Driver")&"</td>"
						response.write "<td width='14%'>"&rsfound("IllegalAddress")&"</td>"
						response.write "<td width='10%'>"&chRule&"</td>"
						'response.write "<td width='6%'>"&ForFeit&"</td>"
						response.write "<td width='8%'>"
						if trim(rsfound("BillStatus"))="0" then
							response.write "未處理"
						elseif trim(rsfound("BillStatus"))="1" then
							response.write "車籍查詢"
						elseif trim(rsfound("BillStatus"))="2" then
							response.write "入案"
						elseif trim(rsfound("BillStatus"))="3" then
							response.write "單退"
						elseif trim(rsfound("BillStatus"))="4" then
							response.write "寄存"
						elseif trim(rsfound("BillStatus"))="5" then
							response.write "公示"
						elseif trim(rsfound("BillStatus"))="6" then
							response.write "刪除"
						elseif trim(rsfound("BillStatus"))="7" then
							response.write "收受註記,收受日期:"
							strMail1="select UserMarkResonID,UserMarkDate,SignDate from BillMailHistory where BillSN="&trim(rsfound("SN"))
							set rsMail1=conn.execute(strMail1)
							if not rsMail1.eof then
								if not isnull(rsMail1("SignDate")) and rsMail1("SignDate")<>"" then
									response.write gInitDT(rsMail1("SignDate"))
								end if
							end if
							rsMail1.close
							set rsMail1=nothing
						end if
						response.write "</td>"
					rsfound.MoveNext
					Wend
					rsfound.close
					set rsfound=nothing
				%>
				</tr>
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