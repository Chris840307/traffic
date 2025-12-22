<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!-- #include file="../Common/Bannernodata.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>漏號稽核</title>
<%
Server.ScriptTimeout=9000
if request("DB_state")="UpdateDetail" then
	strSQL="Update GetBillDetail set BillStateID="&request("Sys_BillStateID_"&request("SN"))&",NoteContent='"&request("Sys_NoteContent_"&request("SN"))&"' where BillNo='"&request("SN")&"'"
	conn.execute(strSQL)
	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
end if
if request("DB_state")="UpdateBill" then
	strSQL="Update BillBase set Note='"&request("Sys_NoteContent_"&request("SN"))&"' where BillNo='"&request("BillNo")&"' and SN="&request("SN")
	conn.execute(strSQL)
	strSQL="Update PasserBase set Note='"&request("Sys_NoteContent_"&request("SN"))&"' where BillNo='"&request("BillNo")&"' and SN="&request("SN")
	conn.execute(strSQL)
	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
end if
if request("DB_Selt")="Selt" then
	DB_Selt="Selt"
	strwhere=" BillNo between '"&Ucase(trim(request("Sys_Sno")))&trim(request("Sys_Tno"))&"' and '"&Ucase(trim(request("Sys_Sno")))&trim(request("Sys_Tno2"))&"'"

	BillCreate=0:BillQuery=0:BillKeyin=0:BillReturn=0:BillSend=0:BillOpen=0:BillDel=0:Billclose=0
	Billnormal=0:Billcancel=0:Billerr=0:BillLose=0:Billnever=0:Billstained=0:BillOther=0
	UseBill=0:DrawBill=0:BillNotUse=0
	if Not ifnull(request("Sys_Sno")) then
		Sno=Ucase(trim(request("Sys_Sno"))):Tno=trim(request("Sys_Tno")):Tno2=trim(request("Sys_Tno2"))
	else
		Sno="":Tno=0:Tno2=0:BillStartNumber="":BillEndNumber=""
		RecordDate1=gOutDT(request("fGetBillDate_q"))&" 0:0:0"
		RecordDate2=gOutDT(request("tGetBillDate_q"))&" 23:59:59"

		StartSQL="select * from (select BillStartNumber from GetBillBase where GetBillDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS') and RecordStateID<>-1 order by BillStartNumber) where rownum = 1"
		set rsStart=conn.execute(StartSQL)
		if Not rsStart.eof then
			if Not ifNull(rsStart("BillStartNumber")) then
				BillStartNumber=trim(rsStart("BillStartNumber"))
				for i=1 to len(BillStartNumber)
					if IsNumeric(mid(BillStartNumber,i,1)) then
						Sno=MID(BillStartNumber,1,i-1)
						Tno=MID(BillStartNumber,i,len(BillStartNumber))
						exit for
					end if
				next
			end if
		end if

		EndSQL="select * from (select BillEndNumber from GetBillBase where GetBillDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS') and RecordStateID<>-1 order by BillEndNumber DESC) where rownum = 1"

		set rsEnd=conn.execute(EndSQL)
		if Not rsStart.eof then
			if Not ifNull(rsEnd("BillEndNumber")) then
				BillEndNumber=trim(rsEnd("BillEndNumber"))
				for i=1 to len(BillEndNumber)
					if IsNumeric(mid(BillEndNumber,i,1)) then
						Tno2=MID(BillEndNumber,i,len(BillEndNumber))
						exit for
					end if
				next
			end if
		end if
	end if
	if trim(request("Sys_Audit"))="" then
		strSQL="select BILLSTATUS,count(*) as cnt from BillBaseView where SUBSTR(BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"' group by BillStatus order by BillStatus"
		set rscnt=conn.execute(strSQL)

		while Not rscnt.eof
			if rscnt("BillStatus")=0 then BillCreate=cdbl(rscnt("cnt"))
			if rscnt("BillStatus")=1 then BillQuery=cdbl(rscnt("cnt"))
			if rscnt("BillStatus")=2 then BillKeyin=cdbl(rscnt("cnt"))
			if rscnt("BillStatus")=3 then BillReturn=cdbl(rscnt("cnt"))
			if rscnt("BillStatus")=4 then BillSend=cdbl(rscnt("cnt"))
			if rscnt("BillStatus")=5 then BillOpen=cdbl(rscnt("cnt"))
			if rscnt("BillStatus")=6 then BillDel=cdbl(rscnt("cnt"))
			if rscnt("BillStatus")=9 then Billclose=cdbl(rscnt("cnt"))
			UseBill=int(UseBill)+cdbl(rscnt("cnt"))
			rscnt.movenext
		wend
		rscnt.close

		strSQL="select BillStateID,count(*) as cnt from GetBillDetail where SUBSTR(BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"' group by BillStateID order by BillStateID"

		set rscnt=conn.execute(strSQL)

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
		rscnt.close

		DBsum=Tno2-Tno+1
		BillNotUse=DBsum-DrawBill
	elseif trim(request("Sys_Audit"))="1" then
		BillSN=""
		BillBaseCnt=0
		if trim(request("BillSN"))="" then
			for j=int(Tno) to int(Tno2)
				Sys_BilLNo=Ucase(trim(Sno))&right("000000000"&j,len(Tno))
				strSQL="select count(*) cnt from BillBaseView where BillNo='"&Sys_BilLNo&"'"
				set rsfound=conn.execute(strSQL)
				if Cint(rsfound("cnt"))=0 then
					rsfound.close
					
					strSQL="select count(*) cnt from (select b.BillNo,c.ChName as BillMem1,a.RECORDDATE as BILLFILLDATE,b.BillStateID as BILLSTATUS from GetBillBase a,GetBillDetail b,MemberData c where a.GetBillSN=b.GetBillSN and a.DispatchMemberID=c.MemberID) where BillNo='"&Sys_BilLNo&"' and BillStatus not in(463)"
					set rsfound=conn.execute(strSQL)
					if Cint(rsfound("cnt"))=0 then
						if trim(BillSN)<>"" then BillSN=trim(BillSN)&","
						BillSN=BillSN&Sys_BilLNo
					end if
				else
					BillBaseCnt=BillBaseCnt+1
				end if
				rsfound.close
			next
		else
			BillSN=trim(request("BillSN"))
		end if
		strSQL="select count(*) as cnt from GetBillDetail where SUBSTR(BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"' and BillStateID in(463)"

		set rscnt=conn.execute(strSQL)
		
		if not rscnt.eof then Billnormal=cdbl(rscnt("cnt"))-BillBaseCnt
		rscnt.close

		if trim(BillSN)<>"" then
			arr_BillSN=split(BillSN,",")
			DBsum=Ubound(arr_BillSN)+1
			BillNotUse=DBsum-Billnormal
		else
			DBsum=0
		end if
	elseif trim(request("Sys_Audit"))="2" then
		BillSN=""
		BillBaseCnt=0
		if trim(request("BillSN"))="" then
			for j=int(Tno) to int(Tno2)
				Sys_BilLNo=Ucase(trim(Sno))&right("000000000"&j,len(Tno))
				strSQL="select BillNo,BILLMEM1,BILLFILLDATE,BILLSTATUS from (select b.BillNo,c.ChName as BillMem1,a.RECORDDATE as BILLFILLDATE,b.BillStateID as BILLSTATUS from GetBillBase a,GetBillDetail b,MemberData c where a.GetBillSN=b.GetBillSN and a.DispatchMemberID=c.MemberID) where BillNo='"&Sys_BilLNo&"' and BillStatus not in(463)"
				set rsfound=conn.execute(strSQL)
				if Not rsfound.eof then
					if trim(BillSN)<>"" then BillSN=trim(BillSN)&","
					BillSN=BillSN&Sys_BilLNo
					BillBaseCnt=BillBaseCnt+1
				end if
				rsfound.close
			next
		else
			BillSN=trim(request("BillSN"))
		end if

		strSQL="select BillStateID,count(*) as cnt from GetBillDetail where BillStateID not in(463) and SUBSTR(BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"' group by BillStateID order by BillStateID"

		set rscnt=conn.execute(strSQL)

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
		rscnt.close

		DBsum=BillBaseCnt
		if trim(BillSN)<>"" then
			arr_BillSN=split(BillSN,",")
			DBsum=Ubound(arr_BillSN)+1
		else
			DBsum=0
		end if
	end if
end if
%>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
<table width="100%" height="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33" height="33">漏號稽核</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table border="0" bgcolor="#FFFFFF" width="100%">
				<tr>
					<td>
					舉發單號
						<input name="Sys_Sno" class="btn1" type="text" value="<%=Ucase(trim(request("Sys_Sno")))%>" size="3" onkeyup="funSnokey(this);">
						<input name="Sys_Tno" class="btn1" type="text" value="<%=trim(request("Sys_Tno"))%>" size="10" onkeyup="funTnokey(this);">
						～
						<input name="Sys_Sno2" class="btn1" type="text" value="<%=Ucase(trim(request("Sys_Sno2")))%>" size="3" disabled>
						<input name="Sys_Tno2" class="btn1" type="text" value="<%=trim(request("Sys_Tno2"))%>" size="10"  onkeyup="value=value.replace(/[^\d]/g,'')">
					稽核案件
						<select name="Sys_Audit" class="btn1">
							<option value="">全部</option>
							<option value="1"<%if request("Sys_Audit")="1" then response.write " selected"%>>未使用</option>
							<option value="2"<%if request("Sys_Audit")="2" then response.write " selected"%>>異常</option>
						</select>
					<br>
						<span class="font12">領單日期</span>
						<input class="btn1" type='text' size='7' id='fGetBillDate_q' name='fGetBillDate_q' value='<%=trim(request("fGetBillDate_q"))%>'>
						<input type="button" name="datestra" value="..." onclick="OpenWindow('fGetBillDate_q');">
						~
						<input class="btn1" type='text' size='7' id='tGetBillDate_q' name='tGetBillDate_q' value='<%=trim(request("tGetBillDate_q"))%>'>
						<input type="button" name="datestrb" value="..." onclick="OpenWindow('tGetBillDate_q');">
						<input type="button" name="btnSelt" value="確定" onClick='funSelt();'>&nbsp;&nbsp;
						<input type="button" name="cancel" value="清除" onClick="location='BillBaseAudit.asp'">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" height="33">漏號稽核紀錄列表<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄）<br>
		<font size="2">
		&nbsp;&nbsp;
		已領用並開單：<%=UseBill%>筆（<%=BillCreate%>筆建檔，<%=BillQuery%>筆車籍查詢，<%=BillKeyin%>筆入案，<%=BillReturn%>筆單退，<%=BillSend%>筆寄存，<%=BillOpen%>筆公示，<%=BillDel%>筆刪除，<%=Billclose%>筆結案）</strong><br>
		&nbsp;&nbsp;&nbsp;
		<strong>已領用未開單：<%=DrawBill-UseBill%>筆（<%=Billnormal-UseBill%>筆未開單，<%=Billcancel%>筆註銷，<%=Billerr%>筆誤寫，<%=BillLose%>筆遺失，
		<%=Billnever%>筆未開單，<%=Billstained%>筆污損，<%=BillOther%>筆其他原因 )</font><br>
		&nbsp;
		<%=BillNotUse%>筆未領單</strong>
		</font>
		</td>
	</tr>
	<%if DB_Selt="Selt" then%>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th height="34">單號</th>
					<th height="34">舉發/領單員警</th>
					<th height="34">入案/領用日期</th>
					<th height="34">建檔人</th>
					<th height="34">狀態</th>
					<th height="34">備註</th>
				</tr><%
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if trim(request("Sys_Audit"))="" then
						for i=DBcnt+1 to DBcnt+10
							Sys_BilLNo=trim(Sno)&right("000000000"&trim(Tno+int(i)-1),len(Tno))
							strSQL="select a.SN,a.BillNo,a.BILLMEM1,a.Note,a.BILLFILLDATE,a.BILLSTATUS,b.ChName,c.BillNo as chkNo,c.NoteContent,c.BillStateID from BillBaseView a,MemberData b,GetBillDetail c where a.BillNo='"&Sys_BilLNo&"' and a.RECORDMEMBERID=b.MemberID(+) and a.BillNo=c.BillNo(+)"
							set rsfound=conn.execute(strSQL)
							BillStatusTmp=split("建檔,車籍查詢,入案,單退,寄存,公示,刪除,收受,,結案",",")
							response.write "<tr bgcolor='#FFFFFF' align='center' "
							lightbarstyle 0 
							response.write ">"
							if Not rsfound.eof then
								response.write "<td>"&rsfound("BillNo")&"</td>"
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
									strCode="select Content,ID from Code where TypeID=17 order by ID"
									set rscode=conn.execute(strCode)
									response.write "<select name=""Sys_BillStateID_"&rsfound("BillNo")&""">"
									while Not rsCode.eof
										response.write "<option value="""&rscode("ID")&""""
										if trim(rsfound("BillStateID"))=trim(rscode("ID")) then response.write " selected"
										response.write ">"&trim(rscode("Content"))&"</option>"
										rsCode.movenext
									wend
									response.write "</select>"
									rscode.close
								end if
								response.write "</td>"

								response.write "<td>"
								if trim(rsfound("chkNo"))<>"" then
									response.write "<input name=""Sys_NoteContent_"&rsfound("BillNo")&""" class=""btn1"" type=""text"" value="""&rsfound("NoteContent")&""" size=""15"">"
									response.write "<input type=""button"" name=""btnNote"" value=""確定"" onClick=""funSubmitDetailNote('"&rsfound("BillNo")&"');"">"
								elseif trim(rsfound("BillNo"))<>"" then
									response.write "<input name=""Sys_NoteContent_"&rsfound("SN")&""" class=""btn1"" type=""text"" value="""&rsfound("Note")&""" size=""15"">"
									response.write "<input type=""button"" name=""btnNote"" value=""確定"" onClick=""funSubmitBillNote('"&rsfound("SN")&"','"&rsfound("BillNo")&"');"">"
								end if
								response.write "</td>"
								response.write "</tr>"
							else
								rsfound.close

								strSQL="select BillNo,BILLMEM1,BILLFILLDATE,BILLSTATUS,NoteContent from (select b.BillNo,c.ChName as BillMem1,a.RECORDDATE as BILLFILLDATE,b.BillStateID as BILLSTATUS,b.NoteContent from GetBillBase a,GetBillDetail b,MemberData c where a.GetBillSN=b.GetBillSN and a.GETBILLMEMBERID=c.MemberID) where BillNo='"&Sys_BilLNo&"'"
								set rsfound=conn.execute(strSQL)
								if Not rsfound.eof then
									response.write "<td>"&rsfound("BillNo")&"</td>"
									response.write "<td>"&rsfound("BILLMEM1")&"</td>"
									response.write "<td>"&gInitDT(rsfound("BILLFILLDATE"))&"</td>"
									response.write "<td></td>"
									response.write "<td>"
									if trim(rsfound("BILLSTATUS"))<>"" and int(rsfound("BILLSTATUS"))>100 then
										strCode="select Content,ID from Code where TypeID=17 order by ID"
										set rscode=conn.execute(strCode)
										response.write "<select name=""Sys_BillStateID_"&rsfound("BillNo")&""">"
										while Not rsCode.eof
											response.write "<option value="""&rscode("ID")&""""
											if trim(rsfound("BILLSTATUS"))=trim(rscode("ID")) then response.write " selected"
											response.write ">"&trim(rscode("Content"))&"</option>"
											rsCode.movenext
										wend
										response.write "</select>"
										rscode.close
									else
										response.write "<strong>末使用</strong>"
									end if
									response.write "</td>"

									response.write "<td>"
									response.write "<input name=""Sys_NoteContent_"&rsfound("BillNo")&""" class=""btn1"" type=""text"" value="""&rsfound("NoteContent")&""" size=""15"">"
									response.write "<input type=""button"" name=""btnNote"" value=""確定"" onClick=""funSubmitDetailNote('"&rsfound("BillNo")&"');"">"
									response.write "</td>"
									rsfound.close
								else
									rsfound.close
									response.write "<td><strong>"&Sys_BilLNo&"</strong></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td><strong>末使用</strong></td>"
								end if
							end if
							response.write "</tr>"
							if trim(int(Tno)+int(i)-1)=trim(int(Tno2)) then exit for
						next
					elseif trim(request("Sys_Audit"))="1" then
						for i=DBcnt+1 to DBcnt+10
							if trim(BillSN)<>"" then
								Sys_BilLNo=arr_BillSN(i-1)
								strSQL="select BillNo,BILLMEM1,BILLFILLDATE,BILLSTATUS,NoteContent from (select b.BillNo,c.ChName as BillMem1,a.RECORDDATE as BILLFILLDATE,b.BillStateID as BILLSTATUS,b.NoteContent from GetBillBase a,GetBillDetail b,MemberData c where a.GetBillSN=b.GetBillSN and a.GETBILLMEMBERID=c.MemberID) where BillNo='"&Sys_BilLNo&"'"
								set rsfound=conn.execute(strSQL)
								response.write "<tr bgcolor='#FFFFFF' align='center' "
								lightbarstyle 0 
								response.write ">"
								if Not rsfound.eof then
									response.write "<td>"&rsfound("BillNo")&"</td>"
									response.write "<td>"&rsfound("BILLMEM1")&"</td>"
									response.write "<td>"&gInitDT(rsfound("BILLFILLDATE"))&"</td>"
									response.write "<td></td>"
									
									response.write "<td>"
									if trim(rsfound("BILLSTATUS"))<>"" and int(rsfound("BILLSTATUS"))>100 then
										strCode="select Content,ID from Code where TypeID=17 order by ID"
										set rscode=conn.execute(strCode)
										response.write "<select name=""Sys_BillStateID_"&rsfound("BillNo")&""">"
										while Not rsCode.eof
											response.write "<option value="""&rscode("ID")&""""
											if trim(rsfound("BILLSTATUS"))=trim(rscode("ID")) then response.write " selected"
											response.write ">"&trim(rscode("Content"))&"</option>"
											rsCode.movenext
										wend
										response.write "</select>"
										rscode.close
									else
										response.write "<strong>末使用</strong>"
									end if
									response.write "</td>"

									response.write "<td>"
									response.write "<input name=""Sys_NoteContent_"&rsfound("BillNo")&""" class=""btn1"" type=""text"" value="""&rsfound("NoteContent")&""" size=""15"">"
									response.write "<input type=""button"" name=""btnNote"" value=""確定"" onClick=""funSubmitDetailNote('"&rsfound("BillNo")&"');"">"
									response.write "</td>"
									rsfound.close
								else
									rsfound.close
									response.write "<td><strong>"&Sys_BilLNo&"</strong></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td><strong>末使用</strong></td>"
								end if
								response.write "</tr>"
							end if
							if i>=DBsum then exit for
						next
					elseif trim(request("Sys_Audit"))="2" then
						
						for i=DBcnt+1 to DBcnt+10
							if trim(BillSN)<>"" then								
								Sys_BilLNo=arr_BillSN(i-1)
								strSQL="select BillNo,BILLMEM1,BILLFILLDATE,BILLSTATUS,NoteContent from (select b.BillNo,c.ChName as BillMem1,a.RECORDDATE as BILLFILLDATE,b.BillStateID as BILLSTATUS,b.NoteContent from GetBillBase a,GetBillDetail b,MemberData c where a.GetBillSN=b.GetBillSN and a.GETBILLMEMBERID=c.MemberID) where BillNo='"&Sys_BilLNo&"'"
								set rsfound=conn.execute(strSQL)
								response.write "<tr bgcolor='#FFFFFF' align='center' "
								lightbarstyle 0 
								response.write ">"
								if Not rsfound.eof then
									response.write "<td>"&rsfound("BillNo")&"</td>"
									response.write "<td>"&rsfound("BILLMEM1")&"</td>"
									response.write "<td>"&gInitDT(rsfound("BILLFILLDATE"))&"</td>"
									response.write "<td></td>"
									
									response.write "<td>"
									if trim(rsfound("BILLSTATUS"))<>"" and int(rsfound("BILLSTATUS"))>100 then
										strCode="select Content,ID from Code where TypeID=17 order by ID"
										set rscode=conn.execute(strCode)
										response.write "<select name=""Sys_BillStateID_"&rsfound("BillNo")&""">"
										while Not rsCode.eof
											response.write "<option value="""&rscode("ID")&""""
											if trim(rsfound("BILLSTATUS"))=trim(rscode("ID")) then response.write " selected"
											response.write ">"&trim(rscode("Content"))&"</option>"
											rsCode.movenext
										wend
										response.write "</select>"
										rscode.close
									else
										response.write "<strong>末使用</strong>"
									end if
									response.write "</td>"

									response.write "<td>"
									response.write "<input name=""Sys_NoteContent_"&rsfound("BillNo")&""" class=""btn1"" type=""text"" value="""&rsfound("NoteContent")&""" size=""15"">"
									response.write "<input type=""button"" name=""btnNote"" value=""確定"" onClick=""funSubmitDetailNote('"&rsfound("BillNo")&"');"">"
									response.write "</td>"
									rsfound.close
								else
									rsfound.close
									response.write "<td><strong>"&Sys_BilLNo&"</strong></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td><strong>末使用</strong></td>"
								end if
								response.write "</tr>"
							end if
							if i>=DBsum then exit for
						next
					end if
					%>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFDD77" align="center">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=int(DBcnt)/10+1&"/"&fix(int(DBsum)/10+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
		</td>
	</tr>
	<%end if%>
</table>
<input type="Hidden" name="DB_Selt" value="<%=DB_Selt%>">
<input type="Hidden" name="DB_state" value="">
<input type="Hidden" name="SN" value="">
<input type="Hidden" name="BillNo" value="">
<input type="Hidden" name="BillSN" value="<%=BillSN%>">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funSnokey(obj){
	obj.maxLength=9-myForm.Sys_Tno.value.length;
	obj.value=obj.value.replace(/[^\A-Za-z]/g,'').toUpperCase();
	myForm.Sys_Sno2.value=obj.value;
	myForm.Sys_Tno.maxLength=9-myForm.Sys_Sno.value.length;
	myForm.Sys_Tno2.maxLength=9-myForm.Sys_Sno.value.length;
}
function funTnokey(obj){
	obj.maxLength=9-myForm.Sys_Sno.value.length;
	obj.value=obj.value.replace(/[^\d]/g,'');
	myForm.Sys_Sno.maxLength=9-myForm.Sys_Tno.value.length;
}
function funSelt(){
	if (myForm.Sys_Tno.value==""&&myForm.Sys_Tno2.value==""&&myForm.fGetBillDate_q.value==""&&myForm.tGetBillDate_q.value==""){
		alert("必須填寫單號範圍！！");
	}else if(myForm.Sys_Tno.value==""&&myForm.Sys_Tno2.value==""&&eval(myForm.Sys_Tno.value)>eval(myForm.Sys_Tno2.value)){
		alert("單號範圍填寫錯誤!!");
	}else{
		myForm.DB_Move.value=0;
		myForm.BillSN.value="";
		myForm.DB_Selt.value="Selt";
		myForm.submit();
	}
}
function funDbMove(MoveCnt){
	if (eval(MoveCnt)>0){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}else{
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}
}
function funchgExecel(){
	UrlStr="BillBaseAudit_Execel.asp?Sys_Sno=<%=Ucase(Sno)%>&Sys_Tno=<%=Tno%>&Sys_Sno2=<%=Ucase(Sno)%>&Sys_Tno2=<%=Tno2%>&Sys_Audit=<%=request("Sys_Audit")%>";
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funInsert(){
	UrlStr="Member_Insert.asp";
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funChangeUnit(SN){
	UrlStr="Member_ChangeUnit.asp?SN="+SN;
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funSubmitDetailNote(SN){
	myForm.SN.value=SN;
	myForm.DB_state.value="UpdateDetail";
	myForm.submit();
}
function funSubmitBillNote(SN){
	myForm.SN.value=SN;
	myForm.DB_state.value="UpdateBill";
	myForm.submit();
}
function funUpdate(SN){
	UrlStr="Member_Update.asp?SN="+SN;
	newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funDel(SN){
	myForm.SN.value=SN;
	myForm.DB_state.value="Del";
	myForm.submit();
}
function funMap(SN){
	UrlStr="SendStyle.asp?MemberID="+SN;
	newWin(UrlStr,"winMap",700,150,50,10,"yes","yes","yes","no");
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}
</script>
<%conn.close%>