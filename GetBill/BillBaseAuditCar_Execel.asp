<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
	fMnoth=month(now)
	if fMnoth<10 then fMnoth="0"&fMnoth
	fDay=day(now)
	if fDay<10 then	fDay="0"&fDay
	fname=year(now)&fMnoth&fDay&"_標示單漏號稽核紀錄.xls"
	Response.AddHeader "Content-Disposition", "filename="&fname
	response.contenttype="application/x-msexcel; charset=MS950"

	Server.ScriptTimeout = 6800
	Response.flush

	BillCreate=0:BillQuery=0:BillKeyin=0:BillAccept=0:BillReturn=0:BillSend=0:BillOpen=0:BillDel=0:Billclose=0
	Billnormal=0:Billcancel=0:Billerr=0:BillLose=0:Billnever=0:Billstained=0:BillOther=0
	changeBillno=0:changeWaring=0:takeCar=0:notFinal=0

	Sno="":Tno=0:Tno2=0:BillStartNumber="":BillEndNumber="":Type_strSQL=""

	BillStartNumber = trim(Request("BillStartNumber"))
	BillEndNumber = trim(Request("BillEndNumber"))

	for i=len(BillStartNumber) to 1 step -1
		if not IsNumeric(mid(BillStartNumber,i,1)) then
			Sno=MID(BillStartNumber,1,i)
			Tno=MID(BillStartNumber,i+1,len(BillStartNumber))
			exit for
		end if
	next

	for i=len(BillEndNumber) to 1 step -1
		if not IsNumeric(mid(BillEndNumber,i,1)) then
			Tno2=MID(BillEndNumber,i+1,len(BillEndNumber))
			exit for
		end if
	next

	DB_Selt="Selt":strBillWhere=""
	if Not ifnull(Sno) then
		Sno=Ucase(trim(Sno))

		whereRepor=whereRepor&" and BillStartNumber like '"&Sno&"%'"
	end if

	If not ifnull(BillStartNumber) Then
		Tno=trim(Tno):Tno2=trim(Tno2)

		whereRepor=whereRepor&" and SUBSTR(BillStartNumber,"&len(Sno)+1&") <= '"&Tno&"' and SUBSTR(BillEndNumber,"&len(Sno)+1&") >='"&Tno2&"'"

		whereDet=" and SUBSTR(BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"'"
	End if
	
		
	if Not ifnull(request("fGetBillDate_q")) then
		RecordDate1=gOutDT(request("fGetBillDate_q"))&" 0:0:0"
		RecordDate2=gOutDT(request("tGetBillDate_q"))&" 23:59:59"
		
		If (not ifnull(Request("chkIllegalData"))) then
			strBillWhere=strBillWhere&" and IllegalDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		else

			whereRepor=whereRepor&" and GetBillDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		end if
	end if

	if not ifnull(request("GetBillMemberID")) then
		whereRepor=whereRepor&" and GetBillMemberID="&request("GetBillMemberID")

	elseif not ifnull(request("UnitID")) then
		whereRepor=whereRepor&" and GetBillMemberID in(select MemberID from MemberData where UnitID in('"&request("UnitID")&"'))"
	end if

	'BillBaseView="select BillNo,BillStateID,NoteContent from WarningGetBillDetail where GetBillSN in(select GetBillsn from WarningGetBillBase where BillIn=0"&whereRepor&")"

	chkData="":Type_strSQL=""
	if trim(request("Sys_Audit"))="" then
		BillBaseView="select BillRep.BillSN,WaringNo.ReportNo,WaringNo.BillStateID from (select BillNo ReportNo,BillStateID from WarningGetBillDetail where GetBillSN in(select GetBillsn from WarningGetBillBase where BillIn=0"&whereRepor&")"&whereDet&") WaringNo,BillReportNo BillRep where WaringNo.ReportNo=BillRep.ReportNo(+)"

		Type_strSQL="select a.sn,b.reportNo,a.BillNo,a.BillStatus,c.unitname,d.chname billunitname,e.chname recordName,a.BillFillDate,a.Note,a.BillBaseTypeID from (select sn,BillNo,BillStatus,BillUnitID,BillMemID1,RecordMemberID,BillBaseTypeID,BillFillDate,Note from Billbase where BillTypeID=2) a,("&BillBaseView&") b,UnitInfo c,memberdata d,memberdata e where a.sn(+)=b.BillSN and a.billunitid=c.unitid(+) and a.billmemid1=d.memberid(+) and a.recordmemberid=e.memberid(+) order by b.reportNo"

		strSQL="select BillStatus,count(1) as cnt from (select a.sn,a.BillStatus from (select sn,BillStatus from Billbase where BillTypeID=2) a,("&BillBaseView&") b where a.sn=b.BillSN) group by BillStatus order by BillStatus"

		set rscnt=conn.execute(strSQL)
		
		If not rscnt.eof Then
			while Not rscnt.eof
				if rscnt("BillStatus")=0 then BillCreate=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=1 then BillQuery=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=2 then BillKeyin=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=3 then BillReturn=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=4 then BillSend=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=5 then BillOpen=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=6 then BillDel=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=7 then BillAccept=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=9 then Billclose=cdbl(rscnt("cnt"))
				DBsum=int(DBsum)+cdbl(rscnt("cnt"))
				rscnt.movenext
			wend
			UseBill=DBsum
			chkData="1"
		end if
		rscnt.close

		strSQL="select a.BillStateID,count(*) as cnt from ("&BillBaseView&") a where a.BillSN is null group by a.BillStateID order by a.BillStateID"
		set rscnt=conn.execute(strSQL)
		If not rscnt.eof Then
			while Not rscnt.eof
				if rscnt("BillStateID")=463 then Billnormal=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=461 then Billcancel=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=462 then Billerr=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=460 then BillLose=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=464 then Billnever=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=459 then Billstained=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=555 then BillOther=cdbl(rscnt("cnt"))

				if rscnt("BillStateID")=556 then changeBillno=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=557 then changeWaring=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=558 then takeCar=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=559 then notFinal=cdbl(rscnt("cnt"))

				DrawBill=int(DrawBill)+cdbl(rscnt("cnt"))
				rscnt.movenext
			wend
			chkData="1"
		end if
		rscnt.close

		DBsum=DrawBill+UseBill
		BillNotUse=0

	elseif trim(request("Sys_Audit"))="1" then

		Sys_BillNo=""
		If trim(request("strBillNo"))="" Then

			If trim(Request("chkBillBase"))="1" Then
				Type_chkBillBase=" and NoteContent is null"

			elseIf trim(Request("chkBillBase"))="2" Then
				Type_chkBillBase=" and NoteContent is not null"

			end If
			
			strSQL="Select distinct GetBillSN from WarningGetBillDetail where 1=1"&whereDet&" and Exists(select 'Y' from BillReportNo where ReportNo=WarningGetBillDetail.BillNo) and Exists (select 'Y' from WarningGetBillBase where BillIn=0"&whereRepor&" and GetBillSN=WarningGetBillDetail.GetBillSN)"

			set rsfound=conn.execute(strSQL)
			While Not rsfound.eof
				strSQL="select BillStartNumber,BillEndNumber from WarningGetBillBase where GetBillSN="&rsfound("GetBillSN")
				set rsda=conn.execute(strSQL)

				BillStartNumber = trim(rsda("BillStartNumber"))
				BillEndNumber = trim(rsda("BillEndNumber"))
				for i=len(BillStartNumber) to 1 step -1
					if not IsNumeric(mid(BillStartNumber,i,1)) then
						Sno=MID(BillStartNumber,1,i)
						Tno=MID(BillStartNumber,i+1,len(BillStartNumber))
						exit for
					end if
				next

				for i=len(BillEndNumber) to 1 step -1
					if not IsNumeric(mid(BillEndNumber,i,1)) then
						Tno2=MID(BillEndNumber,i+1,len(BillEndNumber))
						exit for
					end if
				next
				rsda.close

				strSQL="select Max(SubStr(BillNo,1,"&len(Sno)&")) Sno,Min(SubStr(BillNo,"&len(Sno)+1&")) Tno1,Max(SubStr(BillNo,"&len(Sno)+1&")) Tno2 from (select BillNo from WarningGetBillDetail where GetBillSN="&rsfound("GetBillSN")&" and Exists(select 'Y' from BillReportNo where ReportNo=WarningGetBillDetail.BillNo))"
				set rsbillno=conn.execute(strSQL)
				If Not rsbillno.eof Then
					tmp_Sno=trim(rsbillno("Sno"))
					tmp_Tno1=trim(rsbillno("Tno1"))
					tmp_Tno2=trim(rsbillno("Tno2"))
				End if
				rsbillno.close

				If Not ifnull(tmp_Tno2) Then

					strSQL="select BillNo from WarningGetBillDetail where GetBillSN="&rsfound("GetBillSN")&" and SubStr(BillNo,1,"&len(tmp_Sno)&")='"&tmp_Sno&"' and SubStr(BillNo,"&len(tmp_Sno)+1&")<="&tmp_Tno2&whereDet&Type_chkBillBase&" and Not Exists(select 'Y' from BillReportNo where ReportNo=WarningGetBillDetail.BillNo) order by BillNo"

					set rsloss=conn.execute(strSQL)
					While Not rsloss.eof
						If instr(Sys_BillNo,rsloss("BillNo"))=0 Then
							If trim(Sys_BillNo)<>"" Then Sys_BillNo=Sys_BillNo&","
							Sys_BillNo=Sys_BillNo&rsloss("BillNo")
						End if
						rsloss.movenext
					Wend
					rsloss.close
				end if

				rsfound.movenext
			Wend
			rsfound.close
		else
			Sys_BillNo=trim(request("strBillNo"))
		end if

		arrBillNo=split(Sys_BillNo,",")
		If not ifnull(Sys_BillNo) Then
			Billnormal=cdbl(Ubound(arrBillNo))+1
			DBsum=cdbl(Ubound(arrBillNo))+1
			DrawBill=cdbl(Ubound(arrBillNo))+1
			chkData="1"
		end if
	elseif trim(request("Sys_Audit"))="2" then
		BillBaseView="select BillSN,ReportNo from BillReportNo where Exists(select 'Y' from WarningGetBillDetail where GetBillSN in(select GetBillsn from WarningGetBillBase where BillIn=0"&whereRepor&")"&whereDet&" and BillNo=BillReportNo.ReportNo)"

		Type_strSQL="select a.sn,b.reportNo,a.BillNo,a.BillStatus,c.unitname,d.chname billunitname,e.chname recordName,a.BillFillDate,a.Note,a.BillBaseTypeID from (select sn,BillNo,BillStatus,BillUnitID,BillMemID1,RecordMemberID,BillFillDate,Note,BillBaseTypeID from Billbase where BillTypeID=2 and billstatus=6 and Exists(select 'Y' from BillReportNo where Exists(select 'Y' from WarningGetBillDetail where GetBillSN in(select GetBillsn from WarningGetBillBase where BillIn=0"&whereRepor&")"&whereDet&" and BillNo=BillReportNo.ReportNo) and BillSN=BillBase.SN)) a,("&BillBaseView&") b,UnitInfo c,memberdata d,memberdata e where a.sn=b.BillSN and a.billunitid=c.unitid and a.billmemid1=d.memberid and a.recordmemberid=e.memberid order by b.reportNo"



		strSQL="select BillStatus,count(1) as cnt from (select a.sn,a.BillStatus from (select sn,BillNo,BillStatus,BillUnitID,BillMemID1,RecordMemberID,BillFillDate,Note,BillBaseTypeID from Billbase where BillTypeID=2 and billstatus=6 and Exists(select 'Y' from BillReportNo where Exists(select 'Y' from WarningGetBillDetail where GetBillSN in(select GetBillsn from WarningGetBillBase where BillIn=0"&whereRepor&")"&whereDet&" and BillNo=BillReportNo.ReportNo) and BillSN=BillBase.SN)) a,("&BillBaseView&") b,UnitInfo c,memberdata d,memberdata e where a.sn=b.BillSN and a.billunitid=c.unitid and a.billmemid1=d.memberid and a.recordmemberid=e.memberid) group by BillStatus order by BillStatus"
		set rscnt=conn.execute(strSQL)
		
		If not rscnt.eof Then
			while Not rscnt.eof
				if rscnt("BillStatus")=0 then BillCreate=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=1 then BillQuery=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=2 then BillKeyin=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=3 then BillReturn=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=4 then BillSend=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=5 then BillOpen=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=6 then BillDel=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=7 then BillAccept=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=9 then Billclose=cdbl(rscnt("cnt"))
				DBsum=int(DBsum)+cdbl(rscnt("cnt"))
				rscnt.movenext
			wend
			UseBill=DBsum
			chkData="1"
		end if
		rscnt.close
	elseif trim(request("Sys_Audit"))="3" then

		BillBaseView="select BillRep.BillSN,WaringNo.ReportNo,WaringNo.BillStateID from (select BillNo ReportNo,BillStateID from WarningGetBillDetail where GetBillSN in(select GetBillsn from WarningGetBillBase where BillIn=0"&whereRepor&")"&whereDet&") WaringNo,BillReportNo BillRep where WaringNo.ReportNo=BillRep.ReportNo"

		Type_strSQL="select a.sn,b.reportNo,a.BillNo,a.BillStatus,c.unitname,d.chname billunitname,e.chname recordName,a.BillFillDate,a.Note,a.BillBaseTypeID from (select sn,BillNo,BillStatus,BillUnitID,BillMemID1,RecordMemberID,BillBaseTypeID,BillFillDate,Note from Billbase where BillTypeID=2"&strBillWhere&") a,("&BillBaseView&") b,UnitInfo c,memberdata d,memberdata e where a.sn=b.BillSN and a.billunitid=c.unitid and a.billmemid1=d.memberid and a.recordmemberid=e.memberid order by b.reportNo"

		strSQL="select BillStatus,count(1) as cnt from (select a.sn,a.BillStatus from (select sn,BillStatus from Billbase where BillTypeID=2"&strBillWhere&") a,("&BillBaseView&") b where a.sn=b.BillSN) group by BillStatus order by BillStatus"

		set rscnt=conn.execute(strSQL)
		
		If not rscnt.eof Then
			while Not rscnt.eof
				if rscnt("BillStatus")=0 then BillCreate=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=1 then BillQuery=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=2 then BillKeyin=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=3 then BillReturn=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=4 then BillSend=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=5 then BillOpen=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=6 then BillDel=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=7 then BillAccept=cdbl(rscnt("cnt"))
				if rscnt("BillStatus")=9 then Billclose=cdbl(rscnt("cnt"))
				DBsum=int(DBsum)+cdbl(rscnt("cnt"))
				rscnt.movenext
			wend
			UseBill=DBsum
			chkData="1"
		end if
		rscnt.close

		strSQL="select a.BillStateID,count(*) as cnt from ("&BillBaseView&") a,(select sn from Billbase where BillTypeID=2"&strBillWhere&") b where a.BillSN=b.SN group by a.BillStateID order by a.BillStateID"
		set rscnt=conn.execute(strSQL)
		If not rscnt.eof Then
			while Not rscnt.eof
				if rscnt("BillStateID")=463 then Billnormal=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=461 then Billcancel=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=462 then Billerr=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=460 then BillLose=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=464 then Billnever=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=459 then Billstained=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=555 then BillOther=cdbl(rscnt("cnt"))

				if rscnt("BillStateID")=556 then changeBillno=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=557 then changeWaring=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=558 then takeCar=cdbl(rscnt("cnt"))
				if rscnt("BillStateID")=559 then notFinal=cdbl(rscnt("cnt"))

				DrawBill=int(DrawBill)+cdbl(rscnt("cnt"))
				rscnt.movenext
			wend
			chkData="1"
		end if
		rscnt.close

		DBsum=UseBill
		BillNotUse=0
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>標示單漏號稽核</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="100%" height="100%" border="1">
	<tr>
		<td height="33">標示單漏號稽核紀錄列表</td>
	</tr>
	<tr>
		<td>
			<table width="100%" height="100%" border="1" cellpadding="4" cellspacing="1">
				<tr align="center">
					<th height="25">標示單</th>
					<th height="25">單號</th>
					<th height="25">舉發單位</th>
					<th height="25">舉發員警</th>
					<th height="25">填單日期</th>
					<th height="25">建檔人</th>
					<th height="25">舉發狀態</th>
					<th height="25">備註</th>
				</tr><%
					if trim(request("Sys_Audit"))="1" then
						for i=0 to Ubound(arrBillNo)
							Type_strSQL="select '' sn,a.BillNo ReportNo,'' BillNo,c.UnitName unitname,d.ChName billunitname,'' recordName,'' BillFillDate,a.BILLSTATUS,'' BillBaseTypeID,a.NoteContent Note from (Select GetBillSN,BillNo,(Select content From Code Where TypeId=17 and ID=WarningGetBillDetail.BillStateID) BILLSTATUS,NoteContent from WarningGetBillDetail where BillNo='"&arrBillNo(i)&"') a,(select GetBillsn,GetBillMemberID from WarningGetBillBase where BillIn=0"&whereRepor&") b,UnitInfo c,MemberData d where a.GetBillSN=b.GetBillSN and b.GetBillMemberID=d.MemberID and d.UnitID=c.UnitID"	
								
							set rs=conn.execute(Type_strSQL)
							response.write "<tr bgcolor='#FFFFFF' align='center' "
							lightbarstyle 0 
							response.write ">"
							if Not rs.eof then
								response.write "<td>"&rs("ReportNo")&"</td>"
								response.write "<td>"&rs("BillNo")&"</td>"
								response.write "<td>"&rs("UnitName")&"</td>"
								response.write "<td>"&rs("billunitname")&"</td>"
								response.write "<td>"&gInitDT(rs("BillFillDate"))&"</td>"
								response.write "<td>"&rs("recordName")&"</td>"
								response.write "<td>"

								if IsNumeric(rs("BILLSTATUS")) then
									response.write BillStatusTmp(rs("BILLSTATUS"))
								else
									response.write "<strong>"&rs("BILLSTATUS")&"</strong>"
								end if
								response.write "</td>"

								Response.Write "<td>"

								if trim(request("Sys_Audit"))="1" then
									strSQL="select NoteContent from WarningGetBillDetail where billno='"&rs("ReportNo")&"'"

									set rsrep=conn.execute(strSQL)
									If not rsrep.eof Then
										response.write rsrep("NoteContent")
									else
										Response.Write "無領用標示單紀錄"
									End if
									rsrep.close
								else
									strSQL="select NoteContent from WarningGetBillDetail where billno='"&rs("ReportNo")&"'"
									set rsrep=conn.execute(strSQL)
									If not rsrep.eof Then
										If not ifnull(rsrep("NoteContent")) Then Response.Write rsrep("NoteContent")&"||"
										
									end if
									rsrep.close
									Response.Write rs("Note")
								end if
								
								Response.Write "</td>"
							end if
							response.write "</tr>"
							rs.close
						next
					else
						If not ifnull(Type_strSQL) Then			
							BillStatusTmp=split("建檔,車籍查詢,入案,單退,寄存,公示,刪除,收受,,結案",",")
							set rs=conn.execute(Type_strSQL)
							While not rs.eof

								response.write "<tr bgcolor='#FFFFFF' align='center' "
								lightbarstyle 0 
								response.write ">"
								if Not rs.eof then
									response.write "<td>"&rs("ReportNo")&"</td>"
									response.write "<td>"&rs("BillNo")&"</td>"
									response.write "<td>"&rs("UnitName")&"</td>"
									response.write "<td>"&rs("billunitname")&"</td>"
									response.write "<td>"&gInitDT(rs("BillFillDate"))&"</td>"
									response.write "<td>"&rs("recordName")&"</td>"
									response.write "<td>"

									if IsNumeric(rs("BILLSTATUS")) then
										response.write BillStatusTmp(rs("BILLSTATUS"))
									else
										response.write "<strong>"&rs("BILLSTATUS")&"</strong>"
									end if
									response.write "</td>"

									Response.Write "<td>"

									if trim(request("Sys_Audit"))="1" then
										strSQL="select NoteContent from WarningGetBillDetail where billno='"&rs("ReportNo")&"'"

										set rsrep=conn.execute(strSQL)
										If not rsrep.eof Then
											response.write rsrep("NoteContent")
										else
											Response.Write "無領用標示單紀錄"
										End if
										rsrep.close
									else
										strSQL="select NoteContent from WarningGetBillDetail where billno='"&rs("ReportNo")&"'"
										set rsrep=conn.execute(strSQL)
										If not rsrep.eof Then
											If not ifnull(rsrep("NoteContent")) Then Response.Write rsrep("NoteContent")&"||"
											
										end if
										rsrep.close
										Response.Write rs("Note")
									end if
									
									Response.Write "</td>"
								end if
								response.write "</tr>"
								rs.movenext
							wend
							rs.close
						end if
					end if
					%>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%conn.close%>