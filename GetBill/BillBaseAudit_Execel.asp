<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
	fMnoth=month(now)
	if fMnoth<10 then fMnoth="0"&fMnoth
	fDay=day(now)
	if fDay<10 then	fDay="0"&fDay
	fname=year(now)&fMnoth&fDay&"_漏號稽核紀錄.xls"
	Response.AddHeader "Content-Disposition", "filename="&fname
	response.contenttype="application/x-msexcel; charset=MS950"

	Server.ScriptTimeout = 6800
	Response.flush

	BillCreate=0:BillQuery=0:BillKeyin=0:BillAccept=0:BillReturn=0:BillSend=0:BillOpen=0:BillDel=0:Billclose=0
	Billnormal=0:Billcancel=0:Billerr=0:BillLose=0:Billnever=0:Billstained=0:BillOther=0
	UseBill=0:DrawBill=0:BillNotUse=0:Sno="":Tno=0:Tno2=0:BillStartNumber="":BillEndNumber=""

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

	DB_Selt="Selt"
	if Not ifnull(Sno) then
		Sno=Ucase(trim(Sno)):Tno=trim(Tno):Tno2=trim(Tno2)

		whereSql=whereSql&" and SUBSTR(billD.BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(billD.BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"'"
	end if
		
	if Not ifnull(request("fGetBillDate_q")) then
		RecordDate1=gOutDT(request("fGetBillDate_q"))&" 0:0:0"
		RecordDate2=gOutDT(request("tGetBillDate_q"))&" 23:59:59"

		whereSql=whereSql&" and bill.GetBillDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	if not ifnull(request("GetBillMemberID")) then
		whereSql=whereSql&" and bill.GetBillMemberID="&request("GetBillMemberID")
	end if

	if not ifnull(request("UnitID")) then
		strSQL="select count(1) cnt from Unitinfo where Unitid in('"&request("UnitID")&"') and unitName like '%分局%' and unitlevelid=2"

		set rsuit=conn.execute(strSQL)

		If cdbl(rsuit("cnt"))=0 Then
			whereSql=whereSql&" and bill.GetBillMemberID in(select MemberID from MemberData where UnitID in('"&request("UnitID")&"'))"
		else
			whereSql=whereSql&" and bill.GetBillMemberID in(select MemberID from MemberData where UnitID in(select UnitID from Unitinfo where Unittypeid in(select distinct UnitTypeid from Unitinfo where UnitID in('"&request("UnitID")&"'))))"
		End if

		rsuit.close
	end if

	StartSQL="select bill.GetBillDate,bill.GetBillMemberID,billD.GetBillSN,billD.BillNo,billD.BillStateID from GetBillBase bill,GetBillDetail billD where bill.getBillSN=billD.getBillSN and bill.RecordStateID<>-1"&whereSql
	if trim(request("Sys_Audit"))="2" then
		BillBaseView="select a.Sn,a.BillNo,BillStatus,billunitid,BillMem1,Note,BillFillDate,BillBaseTypeID,RecordMemberID from Billbase a,(select max(sn) sn,BillNo from Billbase group by BillNo) b where a.sn=b.sn union all select a.Sn,a.BillNo,BillStatus,billunitid,BillMem1,Note,BillFillDate,BillBaseTypeID,RecordMemberID from Passerbase a,(select max(sn) sn,BillNo from PasserBase group by BillNo) b where a.sn=b.sn"
	else
		BillBaseView="select Sn,BillNo,BillStatus,billunitid,BillMem1,Note,BillFillDate,BillBaseTypeID,RecordMemberID from Billbase union all select Sn,BillNo,BillStatus,billunitid,BillMem1,Note,BillFillDate,BillBaseTypeID,RecordMemberID from Passerbase"
	end if
	chkData="":Type_strSQL=""
	if trim(request("Sys_Audit"))="" then
		Type_strSQL="select a.BillNo,a.BillStateID,b.BillStatus from ("&StartSQL&") a,("&BillBaseView&") b where a.BillNo=b.BillNo(+) order by a.BillNo"

		strSQL="select b.BillStatus,count(*) as cnt from ("&StartSQL&") a,("&BillBaseView&") b where a.BillNo=b.BillNo group by b.BillStatus order by b.BillStatus"
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
				UseBill=int(UseBill)+cdbl(rscnt("cnt"))
				rscnt.movenext
			wend
		end if
		rscnt.close
		strSQL="select a.BillStateID,count(*) as cnt from ("&StartSQL&") a,("&BillBaseView&") b where a.billNo=b.billNo(+) and b.BillNo is null group by a.BillStateID order by a.BillStateID"
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
		If Not ifnull(request("chkLossBillBase")) Then
			if trim(request("strBillNo"))="" then
				Sys_BillNo=""
				strSQL="select distinct a.GetBillSN from ("&StartSQL&" and billD.BillStateID in(463)) a,("&BillBaseView&") b where a.BillNo=b.BillNo(+) and b.BillNo is null order by GetBillSN"

				set rsfound=conn.execute(strSQL)
				While Not rsfound.eof
					strSQL="select BillStartNumber,BillEndNumber from GetBillBase where GetBillSN="&rsfound("GetBillSN")
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

					strSQL="select Max(SubStr(a.BillNo,1,"&len(Sno)&")) Sno,Min(SubStr(a.BillNo,"&len(Sno)+1&")) Tno1,Max(SubStr(a.BillNo,"&len(Sno)+1&")) Tno2 from (select BillNo from GetBillDetail where GetBillSN="&rsfound("GetBillSN")&") a,("&BillBaseView&") b where a.BillNo=b.BillNo(+) and b.BillNo is not null"
					set rsbillno=conn.execute(strSQL)
					If Not rsbillno.eof Then
						tmp_Sno=trim(rsbillno("Sno"))
						tmp_Tno1=trim(rsbillno("Tno1"))
						tmp_Tno2=trim(rsbillno("Tno2"))
					End if
					rsbillno.close

					If Not ifnull(tmp_Tno2) Then
						strSQL="select distinct a.BillNo from ("&StartSQL&" and billD.GetBillSN="&rsfound("GetBillSN")&" and billD.BillStateID in(463) and SubStr(billD.BillNo,1,"&len(tmp_Sno)&")='"&tmp_Sno&"' and SubStr(billD.BillNo,"&len(tmp_Sno)+1&")<="&tmp_Tno2&") a,("&BillBaseView&") b where a.BillNo=b.BillNo(+) and b.BillNo is null order by a.BillNo"

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
		else
			Type_strSQL="select a.BillNo from ("&StartSQL&" and billD.BillStateID in(463)) a,("&BillBaseView&") b where a.BillNo=b.BillNo(+) and b.BillNo is null order by BillNo"

			strSQL="select count(*) as cnt from ("&StartSQL&" and billD.BillStateID in(463)) a,("&BillBaseView&") b where a.BillNo=b.BillNo(+) and b.BillNo is null"
			set rscnt=conn.execute(strSQL)
			if not rscnt.eof then
				Billnormal=cdbl(rscnt("cnt"))
				DBsum=cdbl(rscnt("cnt"))
				DrawBill=cdbl(rscnt("cnt"))
				chkData="1"
			end if
			rscnt.close
		End if

	elseif trim(request("Sys_Audit"))="2" then
		If not ifnull(request("chkDelBillBase")) Then delBill="b.BillStatus in(6) or"

		Type_strSQL="select a.BillNo from ("&StartSQL&") a,("&BillBaseView&") b where a.BillNo=b.BillNo(+) and ("&delBill&" a.BillStateID in(461,462,460,459,555)) order by BillNo"


		If not ifnull(request("chkDelBillBase")) Then
			strSQL="select count(*) as cnt from ("&StartSQL&") a,("&BillBaseView&") b where a.BillNo=b.BillNo and b.BillStatus in(6)"
			set rscnt=conn.execute(strSQL)
			if Not rscnt.eof then BillDel=cdbl(rscnt("cnt"))
			rscnt.close
		else
			BillDel=0
		end if

		strSQL="select BillStateID,count(*) as cnt from ("&StartSQL&") a,("&BillBaseView&") b where a.BillNo=b.BillNo(+) and ("&delBill&" a.BillStateID in(461,462,460,459,555)) group by a.BillStateID order by a.BillStateID"

		set rscnt=conn.execute(strSQL)

		while Not rscnt.eof
			if rscnt("BillStateID")=463 then Billnormal=cdbl(rscnt("cnt"))
			if rscnt("BillStateID")=461 then Billcancel=cdbl(rscnt("cnt"))
			if rscnt("BillStateID")=462 then Billerr=cdbl(rscnt("cnt"))
			if rscnt("BillStateID")=460 then BillLose=cdbl(rscnt("cnt"))
			if rscnt("BillStateID")=464 then Billnever=cdbl(rscnt("cnt"))
			if rscnt("BillStateID")=459 then Billstained=cdbl(rscnt("cnt"))
			if rscnt("BillStateID")=555 then BillOther=cdbl(rscnt("cnt"))
			DrawBill=cdbl(DrawBill)+cdbl(rscnt("cnt"))
			chkData="1"
			rscnt.movenext
		wend
		rscnt.close
		DBsum=DrawBill+BillDel
		DrawBill=DrawBill+BillDel
		Billnormal=Billnormal+BillDel
		UseBill=BillDel
	end If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>漏號稽核</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="100%" height="100%" border="1">
	<tr>
		<td height="33">漏號稽核紀錄列表</td>
	</tr>
	<tr>
		<td>
			<table width="100%" height="100%" border="1" cellpadding="4" cellspacing="1">
				<tr align="center">
					<td width="30">單號</td>
					<td width="30">領單單位</td>
					<td width="30">員警代碼</td>
					<td width="30">舉發/領單員警</td>
					<td width="30">填單/領用日期</td>
					<td width="30">建檔人</td>
					<td width="60">舉發狀態 / 領單狀態</td>
					<td width="80">備註</td>
				</tr><%
					if trim(request("Sys_Audit"))="" then
						set rs=conn.execute(Type_strSQL)
						While Not rs.eof
							strSQL="select a.SN,a.BillNo,a.BILLMEM1,a.Note,a.BILLFILLDATE,a.BILLSTATUS,a.BillBaseTypeID,d.UnitName,b.LoginID,b.ChName,c.BillNo as chkNo,c.NoteContent,c.BillStateID from ("&BillBaseView&") a,MemberData b,GetBillDetail c,UnitInfo d where a.BillNo='"&trim(rs("BillNo"))&"' and a.RECORDMEMBERID=b.MemberID and a.BillUnitID=d.UnitID(+) and a.BillNo=c.BillNo(+)"
							if Not ifnull(rs("BillStatus")) then
								strSQL=strSQL&" and a.BillStatus="&trim(rs("BillStatus"))
							end if
							set rsfound=conn.execute(strSQL)
							BillStatusTmp=split("建檔,車籍查詢,入案,單退,寄存,公示,刪除,收受,,結案",",")
							response.write "<tr>"
							if Not rsfound.eof then
								response.write "<td>"&rsfound("BillNo")&"</td>"
								response.write "<td>"&rsfound("UnitName")&"</td>"
								response.write "<td>"&rsfound("LoginID")&"</td>"
								response.write "<td>"&rsfound("BILLMEM1")&"</td>"
								response.write "<td>"&gInitDT(rsfound("BILLFILLDATE"))&"</td>"
								response.write "<td>"&rsfound("ChName")&"</td>"
								response.write "<td>"
								'if trim(rsfound("TrafficAccidentType"))<>"" then response.write "A"&trim(rsfound("TrafficAccidentType"))&"&nbsp;&nbsp;"

								if trim(rsfound("BILLSTATUS"))<>"" and int(rsfound("BILLSTATUS"))<100 then
									response.write BillStatusTmp(rsfound("BILLSTATUS"))
								else
									response.write "<strong>末使用</strong>"
								end if
								response.write "&nbsp;&nbsp;&nbsp;&nbsp;"

								if trim(rsfound("BillStateID"))<>"" and int(rsfound("BillStateID"))>100 then
									strCode="select Content,ID from Code where TypeID=17 and ID='"&trim(rsfound("BillStateID"))&"'"
									set rscode=conn.execute(strCode)
									response.write trim(rscode("Content"))
									rscode.close
								end if
								response.write "</td>"

								response.write "<td>"
								'if trim(rsfound("chkNo"))<>"" then
									response.write rsfound("NoteContent")

								'elseif trim(rsfound("BillNo"))<>"" then
									response.write rsfound("Note")

								'end if
								response.write "</td>"
								response.write "</tr>"
							else
								rsfound.close
							
								strSQL="select BillNo,UnitName,LoginID,BILLMEM1,BILLFILLDATE,BILLSTATUS,NoteContent from (select b.BillNo,d.UnitName,c.LoginID,c.ChName as BillMem1,a.getbilldate as BILLFILLDATE,b.BillStateID as BILLSTATUS,b.NoteContent from GetBillBase a,GetBillDetail b,MemberData c,UnitInfo d where a.GetBillSN=b.GetBillSN and a.GETBILLMEMBERID=c.MemberID and c.UnitID=d.UnitID) where BillNo='"&trim(rs("BillNo"))&"'"
								set rsfound=conn.execute(strSQL)
								if Not rsfound.eof then
									response.write "<td>"&rsfound("BillNo")&"</td>"
									response.write "<td>"&rsfound("UnitName")&"</td>"
									response.write "<td>"&rsfound("LoginID")&"</td>"
									response.write "<td>"&rsfound("BILLMEM1")&"</td>"
									response.write "<td>"&gInitDT(rsfound("BILLFILLDATE"))&"</td>"
									response.write "<td></td>"
									response.write "<td>"
									if trim(rsfound("BILLSTATUS"))<>"" and int(rsfound("BILLSTATUS"))>100 then
										strCode="select Content,ID from Code where TypeID=17 and ID='"&trim(rsfound("BILLSTATUS"))&"'"
										set rscode=conn.execute(strCode)
										response.write trim(rscode("Content"))
										rscode.close
									else
										response.write "<strong>末使用</strong>"
									end if
									response.write "</td>"

									response.write "<td>"
									response.write rsfound("NoteContent")
									response.write "</td>"
									rsfound.close
								else
									rsfound.close
									response.write "<td><strong>"&trim(rs("BillNo"))&"</strong></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td><strong>末使用</strong></td>"
								end if
							end if
							response.write "</tr>"
							rs.movenext
						wend
						rs.close
					elseif trim(request("Sys_Audit"))="1" then
						if (Not ifnull(request("chkLossBillBase"))) and (Not ifnull(Sys_BillNo)) then
							for i=0 to Ubound(arrBillNo)
								strSQL="select BillNo,UnitName,LoginID,BILLMEM1,BILLFILLDATE,BILLSTATUS,NoteContent from (select b.BillNo,d.UnitName,c.LoginID,c.ChName as BillMem1,a.getbilldate as BILLFILLDATE,b.BillStateID as BILLSTATUS,b.NoteContent from GetBillBase a,GetBillDetail b,MemberData c,UnitInfo d where a.GetBillSN=b.GetBillSN and a.GETBILLMEMBERID=c.MemberID and c.UnitID=d.UnitID) where BillNo='"&arrBillNo(i)&"'"
								set rsfound=conn.execute(strSQL)
								response.write "<tr>"
								if Not rsfound.eof then
									response.write "<td>"&rsfound("BillNo")&"</td>"
									response.write "<td>"&rsfound("UnitName")&"</td>"
									response.write "<td>"&rsfound("LoginID")&"</td>"
									response.write "<td>"&rsfound("BILLMEM1")&"</td>"
									response.write "<td>"&gInitDT(rsfound("BILLFILLDATE"))&"</td>"
									response.write "<td></td>"
									
									response.write "<td>"
									if trim(rsfound("BILLSTATUS"))<>"" and int(rsfound("BILLSTATUS"))>100 then
										strCode="select Content,ID from Code where TypeID=17 and ID='"&trim(rsfound("BILLSTATUS"))&"'"
										set rscode=conn.execute(strCode)
										response.write trim(rscode("Content"))
										rscode.close
									else
										response.write "<strong>末使用</strong>"
									end if
									response.write "</td>"

									response.write "<td>"
									response.write rsfound("NoteContent")
									response.write "</td>"
									rsfound.close
								else
									rsfound.close
									response.write "<td><strong>"&arrBillNo(i)&"</strong></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td><strong>末使用</strong></td>"
								end if
								response.write "</tr>"
							next
						else
							set rs=conn.execute(Type_strSQL)
							While Not rs.eof
								strSQL="select BillNo,UnitName,LoginID,BILLMEM1,BILLFILLDATE,BILLSTATUS,NoteContent from (select b.BillNo,d.UnitName,c.LoginID,c.ChName as BillMem1,a.getbilldate as BILLFILLDATE,b.BillStateID as BILLSTATUS,b.NoteContent from GetBillBase a,GetBillDetail b,MemberData c,UnitInfo d where a.GetBillSN=b.GetBillSN and a.GETBILLMEMBERID=c.MemberID and c.UnitID=d.UnitID) where BillNo='"&trim(rs("BillNo"))&"'"
								set rsfound=conn.execute(strSQL)
								response.write "<tr>"
								if Not rsfound.eof then
									response.write "<td>"&rsfound("BillNo")&"</td>"
									response.write "<td>"&rsfound("UnitName")&"</td>"
									response.write "<td>"&rsfound("LoginID")&"</td>"
									response.write "<td>"&rsfound("BILLMEM1")&"</td>"
									response.write "<td>"&gInitDT(rsfound("BILLFILLDATE"))&"</td>"
									response.write "<td></td>"
									
									response.write "<td>"
									if trim(rsfound("BILLSTATUS"))<>"" and int(rsfound("BILLSTATUS"))>100 then
										strCode="select Content,ID from Code where TypeID=17 and ID='"&trim(rsfound("BILLSTATUS"))&"'"
										set rscode=conn.execute(strCode)
										response.write trim(rscode("Content"))
										rscode.close
									else
										response.write "<strong>末使用</strong>"
									end if
									response.write "</td>"

									response.write "<td>"
									response.write rsfound("NoteContent")
									response.write "</td>"
									rsfound.close
								else
									rsfound.close
									response.write "<td><strong>"&rs("BillNo")&"</strong></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td><strong>末使用</strong></td>"
								end if
								response.write "</tr>"
								rs.movenext
							wend
							rs.close
						end if
					elseif trim(request("Sys_Audit"))="2" then
						BillStatusTmp=split("建檔,車籍查詢,入案,單退,寄存,公示,刪除,收受,,結案",",")
						set rs=conn.execute(Type_strSQL)
						
						While Not rs.eof
							strSQL="select SN,BillNo,UnitName,BillBaseTypeID,UnitName,LoginID,BILLMEM1,BILLFILLDATE,BillStateID,BILLSTATUS,NoteContent,delson,Note from (select d.SN,d.BillBaseTypeID,b.BillNo,g.UnitName,c.LoginID,c.ChName as BillMem1,a.getbilldate as BILLFILLDATE,b.BillStateID,d.BILLSTATUS,b.NoteContent,d.Note,f.content as delson from GetBillBase a,GetBillDetail b,MemberData c,(select * from ("&BillBaseView&") where BillStatus=6) d,BillDeleteReason e,(select id,content from DCICode where typeid=3) f,UnitInfo g where a.GetBillSN=b.GetBillSN and a.GETBILLMEMBERID=c.MemberID and b.BillNo=d.BillNo(+) and d.SN=e.BillSN(+) and e.Delreason=f.id(+) and c.UnitID=g.UnitID) where BillNo='"&trim(rs("BillNo"))&"'"

							set rsfound=conn.execute(strSQL)
							response.write "<tr>"
							if Not rsfound.eof then
								response.write "<td>"&rsfound("BillNo")&"</td>"
								response.write "<td>"&rsfound("UnitName")&"</td>"
								response.write "<td>"&rsfound("LoginID")&"</td>"
								response.write "<td>"&rsfound("BILLMEM1")&"</td>"
								response.write "<td>"&gInitDT(rsfound("BILLFILLDATE"))&"</td>"
								response.write "<td></td>"
								
								response.write "<td>"
								if trim(rsfound("BillStateID"))<>"" and int(rsfound("BillStateID"))>100 then
									strCode="select Content,ID from Code where TypeID=17 and ID='"&trim(rsfound("BillStateID"))&"'"
									set rscode=conn.execute(strCode)
									response.write trim(rscode("Content"))
									rscode.close
								else
									response.write "<strong>末使用</strong>"
								end if

								if trim(rsfound("BILLSTATUS"))<>"" and int(rsfound("BILLSTATUS"))<100 then
									response.write " / " & BillStatusTmp(rsfound("BILLSTATUS"))
								end if
								response.write "</td>"

								response.write "<td>"
								response.write rsfound("delson")&rsfound("Note")&rsfound("NoteContent")
								response.write "</td>"
								rsfound.close
							else
								rsfound.close
								response.write "<td><strong>"&trim(rs("BillNo"))&"</strong></td>"
								response.write "<td></td>"
								response.write "<td></td>"
								response.write "<td></td>"
								response.write "<td></td>"
								response.write "<td></td>"
								response.write "<td></td>"
								response.write "<td><strong>末使用</strong></td>"
							end if
							response.write "</tr>"
							rs.movenext
						wend
						rs.close
					end if
					%>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%conn.close%>