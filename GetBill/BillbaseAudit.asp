<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>漏號稽核</title>
<%
Server.ScriptTimeout=9000
todayTemp = Right("0"&year(Date()),4) &"/" & Right("0"&month(Date()),2) &"/" & Right("0"&day(Date()),2)
if request("DB_state")="UpdateDetail" then
	
	strSQL="Update GetBillDetail set BillStateID="&request("Sys_BillStateID_"&request("SN"))&",NoteContent='"&request("Sys_NoteContent_"&request("SN"))&"',RecordDate=to_date('" & todayTemp & "','YYYY/MM/DD'),RecordMemberId=" & Session("User_ID") & " where BillNo='"&request("SN")&"'"

	conn.execute(strSQL)
	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
end if
if request("DB_state")="UpdateBill" then
	strSQL="Update GetBillDetail set BillStateID="&request("Sys_BillStateID_"&request("BillNo"))&",NoteContent='"&request("Sys_NoteContent_"&request("SN"))&"',RecordDate=to_date('" & todayTemp & "','YYYY/MM/DD'),RecordMemberId=" & Session("User_ID") & " where BillNo='"&request("BillNo")&"'"

	conn.execute(strSQL)

	strSQL="Update BillBase set Note='"&request("Sys_NoteContent_"&request("SN"))&"' where BillNo='"&request("BillNo")&"'"

	conn.execute(strSQL)

	strSQL="Update PasserBase set Note='"&request("Sys_NoteContent_"&request("SN"))&"' where BillNo='"&request("BillNo")&"'"

	conn.execute(strSQL)
	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
end if
SQLBillNo="":whereSql=""
if request("DB_Selt")="Selt" then
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
			chkData="1"
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
				strSQL="select distinct a.GetBillSN from ("&StartSQL&" and billD.BillStateID in(463)) a,("&BillBaseView&") b where  a.BillNo=b.BillNo(+) and b.BillNo is null order by GetBillSN"

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
	end if
	If ifnull(chkData) Then
		DB_Selt=""
		Response.write "<script>"
		Response.Write "alert('查無資料！');"
		'Response.write "location='BillBaseAudit.asp';"
		Response.write "</script>"
	End if
end if
%>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
<table width="100%" height="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33" height="23">漏號稽核&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;稽核案件:可下拉選擇要稽核的舉發單種類。</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table border="0" bgcolor="#FFFFFF" width="100%">
				<tr>
					<td>
					舉發單號
						<input name="BillStartNumber" class="btn1" type="text" value="<%=trim(request("BillStartNumber"))%>" size="10" maxlength="9" onkeydown="keyFunction();">
						～
						<input name="BillEndNumber" class="btn1" type="text" value="<%=trim(request("BillEndNumber"))%>" size="10" maxlength="9">
					稽核案件
						<select name="Sys_Audit" class="btn1" onchange="funCheckBox();">
							<option value="">全部</option>
							<option value="1"<%if request("Sys_Audit")="1" then response.write " selected"%>>註記&nbsp;未使用&nbsp;之舉發單</option>
							<option value="2"<%if request("Sys_Audit")="2" then response.write " selected"%>>註記&nbsp;異常&nbsp;之舉發單</option>
						</select>
						<input class="btn1" type="checkbox" name="chkDelBillBase" value="true"<%
							If Not ifnull(request("chkDelBillBase")) Then response.write " checked"%> disabled>稽核刪除案件&nbsp;&nbsp;
						<input class="btn1" type="checkbox" name="chkLossBillBase" value="true"<%
							If Not ifnull(request("chkLossBillBase")) Then response.write " checked"%> disabled>只稽核跳號案件
					<br>
						<span class="font12">領單日期</span>
						<input class="btn1" type='text' size='7' id='fGetBillDate_q' name='fGetBillDate_q' value='<%=trim(request("fGetBillDate_q"))%>'>
						<input type="button" name="datestra" value="..." onclick="OpenWindow('fGetBillDate_q');">
						~
						<input class="btn1" type='text' size='7' id='tGetBillDate_q' name='tGetBillDate_q' value='<%=trim(request("tGetBillDate_q"))%>'>
						<input type="button" name="datestrb" value="..." onclick="OpenWindow('tGetBillDate_q');">
						<span class="font12">領單單位</span>
						<%=UnSelectUnitOption("UnitID","GetBillMemberID")%>
						<span class="font12">領單人員</span>
						<%=UnSelectMemberOption("UnitID","GetBillMemberID")%>
						<br>
							&nbsp;&nbsp;&nbsp;	&nbsp;&nbsp;&nbsp;	&nbsp;&nbsp;&nbsp;	&nbsp;&nbsp;&nbsp;	
							<input type="button" name="btnSelt" value="確定" onClick='funSelt();'>&nbsp;&nbsp;
						<input type="button" name="cancel" value="清除" onClick="location='BillBaseAudit.asp'">
                        
                        &nbsp;&nbsp;<font size="2" color="red">* 針對單位內所有派出所 或 單位進行漏號稽核，領單單位 選項請選擇 "所有單位"</font>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" height="33">漏號稽核紀錄列表<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄）<br>
		<font size="2">
		&nbsp;&nbsp;
		已領用並開單：<%=UseBill%>筆（<%=BillCreate%>筆建檔，<%=BillQuery%>筆車籍查詢，<%=BillKeyin%>筆入案，<%=BillAccept%>筆收受，<%=BillReturn%>筆單退，<%=BillSend%>筆寄存，<%=BillOpen%>筆公示，<%=BillDel%>筆刪除，<%=Billclose%>筆結案）</strong><br>
		&nbsp;&nbsp;&nbsp;
		<strong>已領用未開單：<%=DrawBill%>筆（<%=Billnormal%>筆使用中，<%=Billcancel%>筆註銷，<%=Billerr%>筆誤寫，<%=BillLose%>筆遺失，
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
					<th height="25">單號</th>
					<th height="25">領單單位</th>
					<th height="25">舉發/領單員警</th>
					<th height="25">填單/領用日期</th>
					<th height="25">建檔人</th>
					<th height="25">舉發狀態 / 領單狀態</th>
					<th height="25">備註</th>
				</tr><%
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if trim(request("Sys_Audit"))="" then
						set rs=conn.execute(Type_strSQL)
						if Not rs.eof then rs.move Cint(DBcnt)
						for i=DBcnt+1 to DBcnt+10
							if rs.eof then exit for
							strSQL="select a.SN,a.BillNo,a.billunitid,a.BILLMEM1,a.Note,a.BILLFILLDATE,a.BILLSTATUS,a.BillBaseTypeID,b.ChName,c.BillNo as chkNo,c.NoteContent,c.BillStateID,d.UnitName from ("&BillBaseView&") a,MemberData b,GetBillDetail c,UnitInfo d where a.BillNo='"&trim(rs("BillNo"))&"' and a.RECORDMEMBERID=b.MemberID(+) and a.BillNo=c.BillNo(+) and a.billUnitid=d.unitid"
							if Not ifnull(rs("BillStatus")) then
								strSQL=strSQL&" and a.BillStatus="&trim(rs("BillStatus"))
							end if
							set rsfound=conn.execute(strSQL)
							BillStatusTmp=split("建檔,車籍查詢,入案,單退,寄存,公示,刪除,收受,,結案",",")
							response.write "<tr bgcolor='#FFFFFF' align='center' "
							lightbarstyle 0 
							response.write ">"
							if Not rsfound.eof then
								response.write "<td>"&rsfound("BillNo")&"</td>"
								response.write "<td>"&rsfound("UnitName")&"</td>"
								response.write "<td>"&rsfound("BILLMEM1")&"</td>"
								response.write "<td>"&gInitDT(rsfound("BILLFILLDATE"))&"</td>"
								response.write "<td>"&rsfound("ChName")&"</td>"
								response.write "<td>"
								'if trim(rsfound("TrafficAccidentType"))<>"" then response.write "A"&trim(rsfound("TrafficAccidentType"))&"&nbsp;&nbsp;"

								if trim(rsfound("BILLSTATUS"))<>"" and int(rsfound("BILLSTATUS"))<100 then
									response.write BillStatusTmp(rsfound("BILLSTATUS"))
								else
									response.write "<strong>未使用</strong>"
								end if
								response.write "&nbsp;&nbsp;&nbsp;&nbsp;"

								if trim(rsfound("BillStateID"))<>"" and int(rsfound("BillStateID"))>100 then
									strCode="select Content,ID from Code where TypeId=17 and ID in(555,463,461,462,460,464,459) order by showorder"
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
								
								if trim(rsfound("BillNo"))<>"" then
									If trim(rsfound("NoteContent")) <>"" Then
										response.write "<input name=""Sys_NoteContent_"&rsfound("SN")&""" class=""btn1"" type=""text"" value="""&rsfound("NoteContent")&""" size=""15"">"

									else
										response.write "<input name=""Sys_NoteContent_"&rsfound("SN")&""" class=""btn1"" type=""text"" value="""&rsfound("Note")&""" size=""15"">"

									End if 
									
									response.write "<input type=""button"" name=""btnNote"" value=""確定"" onClick=""funSubmitBillNote('"&rsfound("SN")&"','"&rsfound("BillNo")&"');"">"
								elseif trim(rsfound("chkNo"))<>"" then
									response.write "<input name=""Sys_NoteContent_"&rsfound("BillNo")&""" class=""btn1"" type=""text"" value="""&rsfound("NoteContent")&""" size=""15"">"
									response.write "<input type=""button"" name=""btnNote"" value=""確定"" onClick=""funSubmitDetailNote('"&rsfound("BillNo")&"');"">"
								end if
								if trim(rsfound("BillBaseTypeID"))="0" then%>	
									<input type="button" name="b1" value="詳細" onclick='window.open("../Query/BillBaseData_Detail.asp?BillSN=<%=trim(rsfound("SN"))%>&BillType=0","WebPage2","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")' <%
									'1:查詢 ,2:新增 ,3:修改 ,4:刪除
									if CheckPermission(234,1)=false then response.write "disabled"
									%> style="font-size: 10pt; width: 40px; height:26px;"><%
								else%>	
									<input type="button" name="b1" value="詳細" onclick='window.open("../Query/ViewBillBaseData_People.asp?BillSN=<%=trim(rsfound("SN"))%>&BillType=1","WebPage2","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")' <%
									'1:查詢 ,2:新增 ,3:修改 ,4:刪除
									if CheckPermission(234,1)=false then
										response.write "disabled"
									end if
									%> style="font-size: 10pt; width: 40px; height:26px;"><%
								end if
								response.write "</td>"
								response.write "</tr>"
							else
								rsfound.close
							
								strSQL="select BillNo,BILLMEM1,BILLFILLDATE,BILLSTATUS,NoteContent,UnitName from (select b.BillNo,c.ChName as BillMem1,decode(b.recorddate,null,a.getbilldate,b.recorddate) as BILLFILLDATE,b.BillStateID as BILLSTATUS,b.NoteContent,d.UnitName from GetBillBase a,GetBillDetail b,MemberData c,UnitInfo d where  a.GetBillSN=b.GetBillSN and a.GETBILLMEMBERID=c.MemberID and c.UnitID=d.UnitID) where BillNo='"&trim(rs("BillNo"))&"'"
								set rsfound=conn.execute(strSQL)
								if Not rsfound.eof then
									response.write "<td>"&rsfound("BillNo")&"</td>"
									response.write "<td>"&rsfound("UnitName")&"</td>"
									response.write "<td>"&rsfound("BILLMEM1")&"</td>"
									response.write "<td>"&gInitDT(rsfound("BILLFILLDATE"))&"</td>"
									response.write "<td></td>"
									response.write "<td>"
									if trim(rsfound("BILLSTATUS"))<>"" and int(rsfound("BILLSTATUS"))>100 then
										strCode="select Content,ID from Code where TypeId=17 and ID in(555,463,461,462,460,464,459) order by showorder"
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
										response.write "<strong>未使用</strong>"
									end if
									response.write "</td>"

									response.write "<td>"
									response.write "<input name=""Sys_NoteContent_"&rsfound("BillNo")&""" class=""btn1"" type=""text"" value="""&rsfound("NoteContent")&""" size=""15"">"
									response.write "<input type=""button"" name=""btnNote"" value=""確定"" onClick=""funSubmitDetailNote('"&rsfound("BillNo")&"');"">"
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
									response.write "<td><strong>未使用</strong></td>"
								end if
							end if
							response.write "</tr>"
							rs.movenext
						next
						rs.close

					elseif trim(request("Sys_Audit"))="1" then

						if (Not ifnull(request("chkLossBillBase"))) and (Not ifnull(Sys_BillNo)) then
							for i=DBcnt+1 to DBcnt+10
								strSQL="select BillNo,BILLMEM1,BILLFILLDATE,BILLSTATUS,NoteContent,UnitName from (select b.BillNo,c.ChName as BillMem1,decode(b.recorddate,null,a.getbilldate,b.recorddate) as BILLFILLDATE,b.BillStateID as BILLSTATUS,b.NoteContent,d.UnitName from GetBillBase a,GetBillDetail b,MemberData c,UnitInfo d where  a.GetBillSN=b.GetBillSN and a.GETBILLMEMBERID=c.MemberID and c.unitid=d.unitid) where BillNo='"&arrBillNo(i-1)&"'"
								set rsfound=conn.execute(strSQL)

								response.write "<tr bgcolor='#FFFFFF' align='center' "
								lightbarstyle 0 
								response.write ">"
								if Not rsfound.eof then
									response.write "<td>"&rsfound("BillNo")&"</td>"
									response.write "<td>"&rsfound("UnitName")&"</td>"
									response.write "<td>"&rsfound("BILLMEM1")&"</td>"
									response.write "<td>"&gInitDT(rsfound("BILLFILLDATE"))&"</td>"
									response.write "<td></td>"
									
									response.write "<td>"
									if trim(rsfound("BILLSTATUS"))<>"" and int(rsfound("BILLSTATUS"))>100 then
										strCode="select Content,ID from Code where TypeId=17 and ID in(555,463,461,462,460,464,459) order by showorder"
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
										response.write "<strong>未使用</strong>"
									end if
									response.write "</td>"

									response.write "<td>"
									response.write "<input name=""Sys_NoteContent_"&rsfound("BillNo")&""" class=""btn1"" type=""text"" value="""&rsfound("NoteContent")&""" size=""15"">"
									response.write "<input type=""button"" name=""btnNote"" value=""確定"" onClick=""funSubmitDetailNote('"&rsfound("BillNo")&"');"">"
									response.write "</td>"
									rsfound.close
								else
									rsfound.close
									response.write "<td><strong>"&arrBillNo(i-1)&"</strong></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td></td>"
									response.write "<td><strong>未使用</strong></td>"
								end if
								response.write "</tr>"

								if i>=DBsum then exit for
							next
						else
							set rs=conn.execute(Type_strSQL)
							if Not rs.eof then rs.move Cint(DBcnt)
							for i=DBcnt+1 to DBcnt+10
								if rs.eof then exit for
								strSQL="select BillNo,BILLMEM1,BILLFILLDATE,BILLSTATUS,NoteContent,UnitName from (select b.BillNo,c.ChName as BillMem1,decode(b.recorddate,null,a.getbilldate,b.recorddate) as BILLFILLDATE,b.BillStateID as BILLSTATUS,b.NoteContent,d.UnitName from GetBillBase a,GetBillDetail b,MemberData c,UnitInfo d where a.GetBillSN=b.GetBillSN and a.GETBILLMEMBERID=c.MemberID and c.unitid=d.unitid) where BillNo='"&trim(rs("BillNo"))&"'"
								set rsfound=conn.execute(strSQL)
								response.write "<tr bgcolor='#FFFFFF' align='center' "
								lightbarstyle 0 
								response.write ">"
								if Not rsfound.eof then
									response.write "<td>"&rsfound("BillNo")&"</td>"
									response.write "<td>"&rsfound("UnitName")&"</td>"
									response.write "<td>"&rsfound("BILLMEM1")&"</td>"
									response.write "<td>"&gInitDT(rsfound("BILLFILLDATE"))&"</td>"
									response.write "<td></td>"
									
									response.write "<td>"
									if trim(rsfound("BILLSTATUS"))<>"" and int(rsfound("BILLSTATUS"))>100 then
										strCode="select Content,ID from Code where TypeId=17 and ID in(555,463,461,462,460,464,459) order by showorder"
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
										response.write "<strong>未使用</strong>"
									end if
									response.write "</td>"

									response.write "<td>"
									response.write "<input name=""Sys_NoteContent_"&rsfound("BillNo")&""" class=""btn1"" type=""text"" value="""&rsfound("NoteContent")&""" size=""15"">"
									response.write "<input type=""button"" name=""btnNote"" value=""確定"" onClick=""funSubmitDetailNote('"&rsfound("BillNo")&"');"">"
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
									response.write "<td><strong>未使用</strong></td>"
								end if
								response.write "</tr>"
								rs.movenext
							next
							rs.close
						end if
					elseif trim(request("Sys_Audit"))="2" then
						BillStatusTmp=split("建檔,車籍查詢,入案,單退,寄存,公示,刪除,收受,,結案",",")
						set rs=conn.execute(Type_strSQL)
						if Not rs.eof then rs.move Cint(DBcnt)
						for i=DBcnt+1 to DBcnt+10
							if rs.eof then exit for
							strSQL="select SN,BillNo,BillBaseTypeID,BILLMEM1,BILLFILLDATE,BillStateID,BILLSTATUS,NoteContent,delson,Note,UnitName from (select d.SN,d.BillBaseTypeID,b.BillNo,c.ChName as BillMem1,decode(b.recorddate,null,a.getbilldate,b.recorddate) as BILLFILLDATE,b.BillStateID,d.BILLSTATUS,b.NoteContent,d.Note,f.content as delson,g.UnitName from GetBillBase a,GetBillDetail b,MemberData c,(select * from ("&BillBaseView&") where BillStatus=6) d,BillDeleteReason e,(select id,content from DCICode where typeid=3) f,UnitInfo g where a.GetBillSN=b.GetBillSN and a.GETBILLMEMBERID=c.MemberID and c.unitid=g.unitid and b.BillNo=d.BillNo(+) and d.SN=e.BillSN(+) and e.Delreason=f.id(+)) where BillNo='"&trim(rs("BillNo"))&"'"

							set rsfound=conn.execute(strSQL)
							response.write "<tr bgcolor='#FFFFFF' align='center' "
							lightbarstyle 0 
							response.write ">"
							if Not rsfound.eof then
								response.write "<td>"&rsfound("BillNo")&"</td>"
								response.write "<td>"&rsfound("UnitName")&"</td>"
								response.write "<td>"&rsfound("BILLMEM1")&"</td>"
								response.write "<td>"&gInitDT(rsfound("BILLFILLDATE"))&"</td>"
								response.write "<td></td>"
								
								response.write "<td>"
								if trim(rsfound("BillStateID"))<>"" and int(rsfound("BillStateID"))>100 then
									strCode="select Content,ID from Code where TypeId=17 and ID in(555,463,461,462,460,464,459) order by showorder"
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
								else
									response.write "<strong>未使用</strong>"
								end if

								if trim(rsfound("BILLSTATUS"))<>"" and int(rsfound("BILLSTATUS"))<100 then
									response.write " / " & BillStatusTmp(rsfound("BILLSTATUS"))
								end if
								response.write "</td>"

								response.write "<td>"
								response.write "<input name=""Sys_NoteContent_"&rsfound("BillNo")&""" class=""btn1"" type=""text"" value="""&rsfound("delson")&rsfound("Note")&rsfound("NoteContent")&""" size=""15"">"
								response.write "<input type=""button"" name=""btnNote"" value=""確定"" onClick=""funSubmitDetailNote('"&rsfound("BillNo")&"');"">"

								If Not ifnull(trim(rsfound("SN"))) Then
									if trim(rsfound("BillBaseTypeID"))="0" then%>	
										<input type="button" name="b1" value="詳細" onclick='window.open("../Query/BillBaseData_Detail.asp?BillSN=<%=trim(rsfound("SN"))%>&BillType=0","WebPage2","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")' <%
										'1:查詢 ,2:新增 ,3:修改 ,4:刪除
										if CheckPermission(234,1)=false then response.write "disabled"
										%> style="font-size: 10pt; width: 40px; height:26px;"><%
									else%>	
										<input type="button" name="b1" value="詳細" onclick='window.open("../Query/ViewBillBaseData_People.asp?BillSN=<%=trim(rsfound("SN"))%>&BillType=1","WebPage2","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")' <%
										'1:查詢 ,2:新增 ,3:修改 ,4:刪除
										if CheckPermission(234,1)=false then
											response.write "disabled"
										end if
										%> style="font-size: 10pt; width: 40px; height:26px;"><%
									end if
								end if
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
								response.write "<td><strong>未使用</strong></td>"
							end if
							response.write "</tr>"
							rs.movenext
						next
						rs.close
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
<input type="Hidden" name="strBillNo" value="<%=Sys_BillNo%>">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
<%response.write "UnitMan('UnitID','GetBillMemberID','"&request("GetBillMemberID")&"');"%>
function funSelt(){
	if (myForm.BillStartNumber.value==""&&myForm.BillEndNumber.value==""&&myForm.fGetBillDate_q.value==""&&myForm.tGetBillDate_q.value==""&&myForm.GetBillMemberID.value==""){
		alert("必須填寫單號範圍！！");
	}else{
		myForm.DB_Move.value=0;
		myForm.strBillNo.value="";
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
	UrlStr="BillBaseAudit_Execel.asp";
	myForm.action=UrlStr;
	myForm.target="inputWin";
	myForm.submit();
	myForm.action="";
	myForm.target="";
	//UrlStr="BillBaseAudit_Execel.asp?Sys_Sno=<%=Ucase(Sno)%>&Sys_Tno=<%=Tno%>&Sys_Sno2=<%=Ucase(Sno)%>&Sys_Tno2=<%=Tno2%>&Sys_Audit=<%=request("Sys_Audit")%>";
	//newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
}
function funCheckBox(){
	if(myForm.Sys_Audit.value=='1'){

		myForm.chkLossBillBase.disabled=false;
		myForm.chkDelBillBase.disabled=true;

	}else if(myForm.Sys_Audit.value=='2'){

		myForm.chkDelBillBase.disabled=false;
		myForm.chkLossBillBase.disabled=true;

	}else{

		myForm.chkLossBillBase.disabled=true;
		myForm.chkDelBillBase.disabled=true;
	}
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
function funSubmitBillNote(SN,BillNo){
	myForm.SN.value=SN;
	myForm.BillNo.value=BillNo;
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
function keyFunction() {
	
	if (event.keyCode==13||event.keyCode==9) {
		if (myForm.BillStartNumber.value.length==9) {
			runServerScript("chkAudit.asp?BillStartNumber="+myForm.BillStartNumber.value);
		}else{
			alert("單號長度必須為9碼!!"  );
		}
	}
}
funCheckBox();
</script>
<%conn.close%>