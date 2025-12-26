<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DCIURL.ini"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>DCI 資料交換紀錄</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 16px; color:#ff0000; }
.btn3{
   font-size:12px;
   font-family:新細明體;
   background-color:#EEEEEE;
   border-style:solid;
}
-->
</style>
</head>
<%
'檢查是否可進入本系統

If isEmpty(request("DB_Display")) Then
	Sys_Now=DateAdd("d",-2,date)&" "&hour(time)&":"&Minute(time)&":"&Second(time)

	Sys_Now2=DateAdd("d",-10,date)&" "&hour(time)&":"&Minute(time)&":"&Second(time)
	strSQL="select distinct a.batchnumber from PasserDCILog a where (select count(1) cnt from DCIReturnStatus where a.ExchangeTypeID=DCIActionID and a.DCIReturnStatusID=DCIReturn and DCIRETURNSTATUS in('1','-1')) = 0 and a.ExchangeDate between TO_DATE('"&Sys_Now2&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&Sys_Now1&"','YYYY/MM/DD/HH24/MI/SS') and substr(a.batchnumber,1,1)<>'A' and a.RecordMemberID ="&Session("User_ID")

	chkbat=""

	set rschk=conn.execute(strSQL)
	while not rschk.eof
		If Not ifnull(chkbat) then chkbat=chkbat&"\n"
		chkbat=chkbat&rschk("batchnumber")
		rschk.movenext
	wend
	rschk.close
	If not ifnull(chkbat) Then
		Response.write "<script>"
		Response.Write "alert('您下列批號尚未回傳，請盡速確認！\n"&chkbat&"');"
		Response.write "</script>"
	End if
End if

Dim RecordDate,RecordDate1,strwhere,tmp_BatchNumber,Sys_BatchNumber,DB_Display

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

RecordDate=split(gInitDT(date),"-")
strwhere="":tmp_BatchNumber="":Sys_BatchNumber=""
if UCase(trim(request("Sys_BatchNumber")))<>"" then
	tmp_BatchNumber=split(UCase(trim(request("Sys_BatchNumber"))),",")
	for i=0 to Ubound(tmp_BatchNumber)
		if i>0 then Sys_BatchNumber=trim(Sys_BatchNumber)&","
		if i=0 then
			Sys_BatchNumber=trim(Sys_BatchNumber)&tmp_BatchNumber(i)
		else
			Sys_BatchNumber=trim(Sys_BatchNumber)&"'"&tmp_BatchNumber(i)
		end if
		if i<Ubound(tmp_BatchNumber) then Sys_BatchNumber=trim(Sys_BatchNumber)&"'"
	next
	'strwhere=" and a.BatchNumber in('"&Sys_BatchNumber&"')"
end if

if request("DB_state")="ReSend" then
	sql1 = "Update PasserDCILog set FileName='',SeqNo='',dcireturnstatusid='',dcierrorcardata=null,dcierroriddata=null Where SN in (select a.SN from PasserDCILog a where (select count(1) cnt from DCIReturnStatus where a.ExchangeTypeID=DCIActionID and a.DCIReturnStatusID=DCIReturn and DCIRETURNSTATUS='1') = 0 and a.BatchNumber in('"&Sys_BatchNumber&"'))"

	if instr(Sys_BatchNumber,"W")>0 then

		sql1=sql1&" and exists(select 'Y' from passerbase where sn=PasserDCILog.billsn and recordstateid=0)"
	end If 

	Conn.Execute(sql1)

	if instr(Sys_BatchNumber,"E")>0 then
		sql2 = "Update passerbase set BillStatus=6,RecordStateID=-1 Where SN in (select billsn from PasserDCILog where BatchNumber in('"&Sys_BatchNumber&"'))"
		Conn.Execute(sql2)
	end if 

	Response.write "<script>"
	Response.Write "alert('重送完成！');"
	Response.write "</script>"
end if
if request("DB_Selt")="BatchSelt" then
	strwhere=""
	if UCase(request("Sys_BatchNumber"))<>"" then
		strwhere=" and a.BatchNumber in('"&Sys_BatchNumber&"')"
	end if

	if request("Sys_DCIReturnStatus_Batch")<>"" then

		If request("Sys_DCIReturnStatus_Batch") = "is null" Then

			
			strwhere=strwhere&" and a.DCIReturnStatusID is null"

		else

			
			strwhere=strwhere&" and (select count(1) cnt from DCIReturnStatus where a.ExchangeTypeID=DCIActionID and a.DCIReturnStatusID=DCIReturn and DCIReturnStatus "&request("Sys_DCIReturnStatus_Batch")&" ) > 0"		
		End if 

	end if
	orderwhere=" order by a.Batchnumber,a.RecordDate"
end if
if request("DB_Selt")="Selt" then
	strwhere=""

	if request("RecordDate")<>"" and request("RecordDate1")<>""then
		strwhere=" and a.RecordDate between TO_DATE('"&RecordDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&RecordDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	if request("ExchangeDate")<>"" and request("ExchangeDate1")<>""then
		ExchangeDate1=gOutDT(request("ExchangeDate"))&" 0:0:0"
		ExchangeDate2=gOutDT(request("ExchangeDate1"))&" 23:59:59"

		strwhere=strwhere&" and ExchangeDate between TO_DATE('"&ExchangeDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ExchangeDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	if trim(request("ExchangeDate_h"))<>"" or trim(request("ExchangeDate1_h"))<>"" then

		strwhere=strwhere&" and to_char(ExchangeDate,'hh') between "&trim(request("ExchangeDate_h"))&" and "&trim(request("ExchangeDate1_h"))
	end if

	if request("Sys_BillUnitID")<>"" and trim(request("Sys_BillNo"))="" then
		if request("Sys_BillMem")<>"" then

			strwhere=strwhere&" and billsn in(select sn from passerbase where billmemid1 ="&request("Sys_BillMem")&")"
		else

			strwhere=strwhere&" and billsn in(select sn from passerbase where billunitid in('"&request("Sys_BillUnitID")&"'))"
		end if
	end if
	if request("Sys_BillTypeID")<>"" then

		strwhere=strwhere&" and BillTypeID='"&request("Sys_BillTypeID")&"'"
	end if

	if request("Sys_DCIReturnStatus")<>"" then
		
		If request("Sys_DCIReturnStatus") = "is null" Then

			
			strwhere=strwhere&" and a.DCIReturnStatusID is null"

		else

			
			strwhere=strwhere&" and (select count(1) cnt from DCIReturnStatus where a.ExchangeTypeID=DCIActionID and a.DCIReturnStatusID=DCIReturn and DCIReturnStatus "&request("Sys_DCIReturnStatus")&" ) > 0"		
		End if 
	end if

	if trim(request("Sys_BillNo"))<>"" then

		strwhere=" and BillSN in(select sn from passerbase where BillNo='"&Ucase(request("Sys_BillNo"))&"')"
	end if

	if request("Sys_CarNo")<>"" then

		strwhere=strwhere&" and CarNo='"&Ucase(request("Sys_CarNo"))&"'"
	end if

	if request("Sys_ExchangeTypeID")<>"" then
		'Sys_ExchangeTypeID=split(trim(request("Sys_ExchangeTypeID")),"_")
		If trim(request("Sys_ExchangeTypeID"))="3" or trim(request("Sys_ExchangeTypeID"))="4" or trim(request("Sys_ExchangeTypeID"))="5" or trim(request("Sys_ExchangeTypeID"))="7" or trim(request("Sys_ExchangeTypeID"))="Y" Then

			strwhere=strwhere&" and ExchangeTypeID='N' and ReturnMarkType='"&trim(request("Sys_ExchangeTypeID"))&"'"
		else

			strwhere=strwhere&" and ExchangeTypeID='"&trim(request("Sys_ExchangeTypeID"))&"'"
		end if
	end If 

	orderwhere=" order by a.ExchangeDate"
end if
DB_Display=request("DB_Display")
if DB_Display="show" then
	if trim(strwhere&strDCISQL)<>"" then
		strwhereToPrintCarData=strwhere

		strSQL="select a.SN,a.BillSN,a.RecordDate,a.ReturnMarkType,a.FileName,a.DCIReturnStatusID,a.ExchangeTypeID,a.DciErrorCarData,a.DCIErrorIDdata	,a.BillNo,a.CarNo,a.BillTypeID,a.EXCHANGEDATE,a.RecordMemberID,a.seqNo,a.BatchNumber," &_
		"(select chname from memberdata where a.RecordMemberID=MemberID) ChName," &_
		"(select max(Content) from DCIcode where TypeID=2 and ID=a.BillTypeID) BillTypeName," &_
		"(select max(DCIReturn) from DCIReturnStatus where a.ExchangeTypeID=DCIActionID and a.DCIReturnStatusID=DCIReturn) DCIReturn," &_
		"(select max(StatusContent) from DCIReturnStatus where a.ExchangeTypeID=DCIActionID and a.DCIReturnStatusID=DCIReturn) StatusContent," &_
		"(select max(DCIRETURNSTATUS) from DCIReturnStatus where a.ExchangeTypeID=DCIActionID and a.DCIReturnStatusID=DCIReturn) DCIRETURNSTATUS," &_
		"(select max(DCIActionName) from DCIReturnStatus where a.ExchangeTypeID=DCIActionID) DCIActionName," &_
		"(select max(DCIreturn) from DciReturnStatus where DciActionID='WE' and a.DCIERRORCARDATA=DciReturn) CarErrorSN," &_
		"(select max(StatusContent) from DciReturnStatus where DciActionID='WE' and a.DCIERRORCARDATA=DciReturn) CarErrorContent," &_
		"(select max(DCIreturn) from DciReturnStatus where DciActionID='WE' and a.DCIERRORIDDATA=DciReturn) DCIErrorSN," &_
		"(select max(StatusContent) from DciReturnStatus where DciActionID='WE' and a.DCIERRORIDDATA=DciReturn) DCIErrorContent " &_
		"from PasserDCILog a where a.sn=a.sn "&strwhere & orderwhere

		set rsfound=conn.execute(strSQL)

		BillSN=""
		If (instr(strwhere,"BatchNumber") >0 and instr(strwhere,(year(date)-1911)&"W")>0) or (instr(strwhere,"BatchNumber") >0 and instr(strwhere,(year(date)-1912)&"W")>0) or (instr(strwhere,"ExchangeTypeID='W'")>0) Then

			tmpSQL="select a.BillSN,a.Billno,a.RecordDate,a.EXCHANGEDATE,a.BatchNumber " &_
			" from PasserDCILog a where a.sn=a.sn and billtypeid=2 and (select count(1) cnt from DCIReturnStatus where a.ExchangeTypeID=DCIActionID and a.DCIReturnStatusID=DCIReturn and DCIReturnStatus =1 ) > 0 "&strwhere & " order by Billno"
			set rspg=conn.execute(tmpSQL)

			While not rspg.eof

				If BillSN<>"" then BillSN=BillSN&","

				BillSN=BillSN&rspg("BillSN")

				pbSQL="select SN,DriverAddress from passerbase where sn="&rspg("BillSN")&" and recordstateid=0"

				set rssn=conn.execute(pbSQL)
				If not rssn.eof Then
					
					If ifnull(rssn("DriverAddress")) Then

						chkOwnerAddr="":Sys_OwnerZipName=""

						carSQL="select Owner,ownerID,OwnerZip,Owneraddress from PasserBaseDciReturn where billsn="&trim(rssn("SN"))&" and ExchangetypeID='A' and (select count(1) cnt from PasserDCILog where ExchangetypeID='A' and billsn="&trim(rssn("SN"))&" and dcireturnstatusid in(select dcireturn from dcireturnstatus where dcireturnstatus=1))>0"
			
						set rsfi=conn.execute(carSQL)

						if Not rsfi.eof then
							If Not ifnull(trim(rsfi("Owneraddress"))) Then
								chkOwnerAddr=trim(rsfi("Owneraddress"))&"(車)"

								strSQL="select ZipName from Zip where ZipID='"&trim(rsfi("OwnerZip"))&"'"
								set rszip=conn.execute(strSQL)
								if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
								rszip.close

								chkOwnerAddr=Sys_OwnerZipName&replace(replace(chkOwnerAddr,"臺","台"),Sys_OwnerZipName,"")

								strSQL="update passerbase set" & _
								" Driver='"&trim(rsfi("Owner"))&"',DriverZip='"&trim(rsfi("OwnerZip"))&"'" & _
								",DriverID='"&trim(rsfi("ownerID"))&"',DriverAddress='"&chkOwnerAddr&"'"& _
								" where SN="&trim(rssn("SN"))

								conn.execute(strSQL)
							end if
						end If 
						rsfi.close

						If ifnull(chkOwnerAddr) Then					

							carSQL="select Owner,ownerID,OwnerZip,Owneraddress from PasserBaseDciReturn where billsn="&trim(rssn("SN"))&" and ExchangetypeID='W' and (select count(1) cnt from PasserDCILog where ExchangetypeID='W' and billsn="&trim(rssn("SN"))&" and dcireturnstatusid in(select dcireturn from dcireturnstatus where dcireturnstatus=1))>0"
				
							set rsfi=conn.execute(carSQL)

							if Not rsfi.eof then
								If Not ifnull(trim(rsfi("Owneraddress"))) Then

									chkOwnerAddr=trim(rsfi("Owneraddress"))&"(車)"

									strSQL="select ZipName from Zip where ZipID='"&trim(rsfi("OwnerZip"))&"'"
									set rszip=conn.execute(strSQL)
									if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
									rszip.close

									chkOwnerAddr=Sys_OwnerZipName&replace(replace(chkOwnerAddr,"臺","台"),Sys_OwnerZipName,"")
									
									strSQL="update passerbase set" & _
									" Driver='"&trim(rsfi("Owner"))&"',DriverZip='"&trim(rsfi("OwnerZip"))&"'" & _
									",DriverID='"&trim(rsfi("ownerID"))&"',DriverAddress='"&chkOwnerAddr&"'" & _
									" where SN="&trim(rssn("SN"))

									conn.execute(strSQL)
								end if
							end If 
							rsfi.close
						End if 
					end If 
				End if 
				rssn.close

				rspg.movenext
			Wend
			rspg.close
		End if 

		
		strSQL="select sum(cnt) cnt from (select count(*) as cnt from PasserDCILog a where a.BillTypeID='2' and (select count(1) cnt from DCIReturnStatus where a.ExchangeTypeID=DCIActionID and a.DCIReturnStatusID=DCIReturn and DCIRETURNSTATUS='1') > 0 and a.ExchangeTypeID<>'E' and nvl(a.DciErrorCarData,0) Not in ('1','3','9','a','j','A','H','K','T','n') "&strwhere &_
		" union all " &_
		"select count(*) as cnt from PasserDCILog a where a.BillTypeID='1' and (select count(1) cnt from DCIReturnStatus where a.ExchangeTypeID=DCIActionID and a.DCIReturnStatusID=DCIReturn and DCIRETURNSTATUS='1') > 0 and a.ExchangeTypeID<>'E' "&strwhere&")"


		set chksuess=conn.execute(strSQL)

		filsuess=CDbl(chksuess("cnt"))
		chksuess.close

		strSQL="select sum(cnt) cnt from (select count(*) as cnt from PasserDCILog a where  a.ExchangeTypeID='E' and a.DCIReturnStatusID='n'" & strwhere &_
		" union all " &_
		" select count(*) as cnt from PasserDCILog a where a.ExchangeTypeID='W' and a.DCIReturnStatusID in ('S','d','e') " & strwhere &_
		" union all " &_
		"select count(*) as cnt from PasserDCILog a where a.ExchangeTypeID='N' and a.DCIReturnStatusID in('n','h') "&strwhere&")"
		set chksuess=conn.execute(strSQL)

		filClose=cdbl(chksuess("cnt"))
		chksuess.close

		strSQL="select count(*) as cnt from PasserDCILog a where (select count(1) cnt from DCIReturnStatus where a.ExchangeTypeID=DCIActionID and a.DCIReturnStatusID=DCIReturn and DCIRETURNSTATUS='-1') > 0 "&strwhere
		set chksuess=conn.execute(strSQL)

		fildel=CDbl(chksuess("cnt"))
		chksuess.close

		strCnt="select count(*) as cnt from PasserDCILog a where a.sn=a.sn"&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=CDbl(Dbrs("cnt"))
		Dbrs.close

		strCnt="select count(*) as cnt from PasserDCILog a where a.ExchangeTypeID='E' and (select count(1) cnt from DCIReturnStatus where a.ExchangeTypeID=DCIActionID and a.DCIReturnStatusID=DCIReturn and DCIRETURNSTATUS='1') > 0"&strwhere
		set Dbrs=conn.execute(strCnt)
		deldata=CDbl(Dbrs("cnt"))
		Dbrs.close

		strCnt="select count(*) as cnt from PasserDCILog a where a.BillTypeID='2' and a.DciErrorCarData in ('1','3','9','a','j','A','H','K','T','n') and (select count(1) cnt from DCIReturnStatus where a.ExchangeTypeID=DCIActionID and a.DCIReturnStatusID=DCIReturn and DCIRETURNSTATUS='1') > 0"&strwhere

		set Dbrs=conn.execute(strCnt)
		errCatCnt=CDbl(Dbrs("cnt"))
		Dbrs.close

	else
		DB_Display=""
		Response.write "<script>"
		Response.Write "alert('必須有查詢條件！');"
		Response.write "</script>"
	end if
end if
%>

<body>

<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#1BF5FF"><span class="style3">DCI 資料交換紀錄</span>(逕舉手開單入案後，請確認自動帶回之應到案處所是否與舉發單上相同)</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						作業批號 
						<Select Name="Selt_BatchNumber" onchange="fnBatchNumber();">
							<option value="">請點選</option><%
							
							'這裡設定設定DCI Log 哪些縣市 批號要顯示幾天
							if sys_City="雲林縣" then
								nowdate=-2
							elseif sys_City="基隆市" then
								nowdate=-3
							else
								nowdate=-5
							end if
							strSQL="select Max(ExchangeDate) ExchangeDate,BatchNumber from PasserDCILog where RecordMemberID="&Session("User_ID")&" and ExchangeDate between TO_DATE('"&DateAdd("d",nowdate, date)&" 00:00','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59"&"','YYYY/MM/DD/HH24/MI/SS') group by BatchNumber order by ExchangeDate DESC"

										
							set rs=conn.execute(strSQL)
							cut=0
							while Not rs.eof
								ExchangeDate=gInitDT(trim(rs("ExchangeDate")))
		
								response.write "<option value="""&trim(rs("BatchNumber"))&""">"
								response.write ExchangeDate& " - "&cut&"　"&trim(rs("BatchNumber"))
								response.write "</option>"


								cut=cut+1
								rs.movenext
							wend
							rs.close
						%>
						</select>
						<input name="Sys_BatchNumber" type="text" class="btn1" value="<%=UCase(trim(request("Sys_BatchNumber")))%>" size="20" maxlength="25">
						　結果
						<select name="Sys_DCIReturnStatus_Batch" class="btn1">
							<option value="">全部</option>
							<option value="is null"<%if trim(request("Sys_DCIReturnStatus_Batch"))="is null" then response.write " Selected"%>>未處理</option>
							<option value="=1"<%if trim(request("Sys_DCIReturnStatus_Batch"))="=1" then response.write " Selected"%>>正常</option>
							<option value="=-1"<%if trim(request("Sys_DCIReturnStatus_Batch"))="=-1" then response.write " Selected"%>>異常</option>
						</select>　
						<input type="button" name="btnSelt" value="查詢" class="btn3" style="width:40px;height:20px;" onclick="funSelt('BatchSelt');">&nbsp;<%

							Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
							response.write "<input type=""button"" name=""ReSend"" value=""該批異常或未處理資料再次上傳"" class=""btn3"" style=""width:240px;height:20px;"" onclick=""funReSend('ReSend');"""
							if DBsum>0 and trim(request("Sys_BatchNumber"))<>"" then
								if Not ifnull(rsfound("RecordMemberID")) then
									if trim(rsfound("RecordMemberID"))<>trim(Session("User_ID")) and trim(Session("Credit_ID"))<>"A000000000" and trim(Session("Credit_ID"))<>"TIFFANY" and trim(Session("Credit_ID"))<>"19870107" then response.write " disabled"
								else
									response.write " disabled"
								end if
							else
								response.write " disabled"
							end if
							response.write ">"
						%>　
						
						<a href="javascript:void(0)" onclick="window.open('PasserDCICreateFileLog.asp','WebPage4','left=130,top=30,location=0,width=800,height=500,resizable=yes,scrollbars=yes');"><font  class="font10">-> 查詢上傳下載歷程紀錄</font></a>
						<!--
							<img src="space.gif" width="15" height="8"></img><a href="uploadtime.htm" target="_blank"><font  class="font12"> ** 查詢系統上傳檔案時間點 ** </font></a>
						-->
					</td>
				</tr>
				<tr>
					<td>
						<hr>
					</td>
				</tr>
				<tr>
					<td nowrap>
						上傳日期
						<input name="ExchangeDate" type="text" class="btn1" value="<%
							if DB_Display="show" then
								response.write trim(request("ExchangeDate"))
							else
								response.write gInitDT(DateAdd("d",-5, date))
							end if%>" size="5" maxlength="8" onkeyup="chknumber(this);">
						<input type="button" name="datestr" value="..." class="btn3" style="width:15px; height:20px;" onclick="OpenWindow('ExchangeDate');">
						~
						<input name="ExchangeDate1" type="text" class="btn1" value="<%
							if DB_Display="show" then
								response.write trim(request("ExchangeDate1"))
							else
								response.write gInitDT(date)
							end if%>" size="5" maxlength="8" onkeyup="chknumber(this);">
						<input type="button" name="datestr" value="..." class="btn3" style="width:15px; height:20px;" onclick="OpenWindow('ExchangeDate1');">
						時段
						<input name="ExchangeDate_h" type="text" class="btn1" value="<%=request("ExchangeDate_h")%>" size="1" maxlength="2" onkeyup="value=value.replace(/[^\d]/g,'')">
						時 ~ 
						<input name="ExchangeDate1_h" type="text" class="btn1" value="<%=request("ExchangeDate1_h")%>" size="1" maxlength="2" onkeyup="value=value.replace(/[^\d]/g,'')">時
						<img src="space.gif" width="5" height="10">

						舉發單類別
						<select name="Sys_BillTypeID" class="btn1">
							<option value="">全部</option>
							<%strSQL="select * from DCIcode where TypeID=2 order by ID"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write "<option value="""&rs1("ID")&""""
								if trim(request("Sys_BillTypeID"))=trim(rs1("ID")) then response.write " selected"
								response.write ">"&rs1("Content")&"</option>"
								rs1.movenext
							wend
							rs1.close%>
						</select>
						DCI作業
						<select name="Sys_ExchangeTypeID" class="btn1">
							<option value="">全部</option>
							<%strSQL="select distinct DCIActionID,DCIActionName from DCIReturnStatus where DCIRETURNSTATUS=1"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								if trim(rs1("DCIActionID"))="N" then
									response.write "<option value=""3"""
									if trim(request("Sys_ExchangeTypeID"))="3" then response.write " selected"
									response.write ">單退</option>"

									response.write "<option value=""4"""
									if trim(request("Sys_ExchangeTypeID"))="4" then response.write " selected"
									response.write ">寄存</option>"

									response.write "<option value=""5"""
									if trim(request("Sys_ExchangeTypeID"))="5" then response.write " selected"
									response.write ">公示</option>"

									response.write "<option value=""Y"""
									if trim(request("Sys_ExchangeTypeID"))="Y" then response.write " selected"
									response.write ">撤銷</option>"

									response.write "<option value=""7"""
									if trim(request("Sys_ExchangeTypeID"))="7" then response.write " selected"
									response.write ">收受</option>"
								else
									response.write "<option value="""&rs1("DCIActionID")&""""
									if trim(request("Sys_ExchangeTypeID"))=trim(rs1("DCIActionID")) then response.write " selected"
									response.write ">"&rs1("DCIActionName")&"</option>"
								end if
								rs1.movenext
							wend
							rs1.close%>
						</select>
						<img src="space.gif" width="3" height="10">
						結果
						<select name="Sys_DCIReturnStatus" class="btn1">
							<option value="">全部</option>
							<option value="is null"<%if trim(request("Sys_DCIReturnStatus"))="is null" then response.write " Selected"%>>未處理</option>
							<option value="=1"<%if trim(request("Sys_DCIReturnStatus"))="=1" then response.write " Selected"%>>正常</option>
							<option value="=-1"<%if trim(request("Sys_DCIReturnStatus"))="=-1" then response.write " Selected"%>>異常</option>
						</select>
						<br>
						舉發單位
						<%						
						Response.Write "<select name=""Sys_BillUnitID"" id=""Sys_BillUnitID"" class=""btn1""  onchange=""UnitMan('Sys_BillUnitID','Sys_BillMem')"";>"

							If Session("UnitLevelID") = 1 Then
								
								strSQL="select UnitID,UnitName from Unitinfo order by UnitOrder,UnitTypeID,UnitID"

								Response.Write "<option value="""">全部</option>"
							elseIf Session("UnitLevelID") = 2 Then
								
								strSQL="select UnitID,UnitName from Unitinfo where UnitTypeID in(select UnitTypeID from Unitinfo uit where UnitID='"&Session("Unit_ID")&"') order by UnitOrder,UnitTypeID,UnitID"
								
								UitType=""
								set rs=conn.execute(strSQL)
								while Not rs.eof

									If UitType <>"" Then UitType=UitType&"','"

									UitType=UitType&trim(rs("UnitID"))

									rs.movenext
								wend
								rs.close

								Response.Write "<option value="""&UitType&""">全部</option>"
							elseIf Session("UnitLevelID") = 3 Then
								
								strSQL="select UnitID,UnitName from Unitinfo where UnitID='"&Session("Unit_ID")&"' order by UnitOrder,UnitTypeID,UnitID"
							End if 
							

							set rs=conn.execute(strSQL)
							while Not rs.eof
		
								response.write "<option value="""&trim(rs("UnitID"))&""""
								If not isEmpty(Request("Sys_BillUnitID")) Then

									If trim(Request("Sys_BillUnitID")) = trim(rs("UnitID")) Then Response.Write " selected"
								else

									If trim(Session("Unit_ID")) = trim(rs("UnitID")) Then Response.Write " selected"
								End if 
								

								Response.Write ">"
								response.write trim(rs("UnitName"))
								response.write "</option>"

								rs.movenext
							wend
							rs.close
							Response.Write "</select>"
						%>
						<img src="space.gif" width="3" height="10">
						舉發人
						<%

							Response.Write "<select name=""Sys_BillMem"" id=""Sys_BillMem"" class=""btn1"">"
							Response.Write "<option value="""">全部</option>"
							
							strSQL="select memberid,chname from memberdata where UnitID='"&Session("Unit_ID")&"' and recordstateid=0 and ACCOUNTSTATEID=0 order by chname"

							If not ifnull(Request("Sys_BillUnitID")) Then

								strSQL="select memberid,chname from memberdata where UnitID in('"&trim(Request("Sys_BillUnitID"))&"') and recordstateid=0 and ACCOUNTSTATEID=0 order by chname"

							End if 
							
							set rs=conn.execute(strSQL)
							while Not rs.eof
		
								response.write "<option value="""&trim(rs("memberid"))&""""

								If trim(Session("User_ID")) = trim(rs("memberid")) Then Response.Write " Selected"

								Response.Write ">"
								response.write trim(rs("chname"))
								response.write "</option>"

								rs.movenext
							wend
							rs.close
							Response.Write "</select>"
						%>
						<img src="space.gif" width="5" height="10">
						舉發單號
						<input name="Sys_BillNo" type="text" class="btn1" value="<%=request("Sys_BillNo")%>" size="10" maxlength="9">
						<img src="space.gif" width="3" height="10">
						車號
						<input name="Sys_CarNo" type="text" class="btn1" value="<%=request("Sys_CarNo")%>" size="8">
						<img src="space.gif" width="5" height="9">
						<input type="button" name="btnSelt" value="進階查詢" class="btn3" style="width:60px; height:20px;" onclick="funSelt('Selt');">
						<input type="button" name="cancel" value="清除" class="btn3" style="width:40px; height:20px;" onClick="location='PasserDCIExchangeQry.asp'">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#1BF5FF">
		資料交換紀錄
		每頁<select name="sys_MoveCnt" onchange="repage();">
				<option value="0"<%if trim(request("sys_MoveCnt"))="0" then response.write " Selected"%>>10</option>
				<option value="10"<%if trim(request("sys_MoveCnt"))="10" then response.write " Selected"%>>20</option>
				<option value="20"<%if trim(request("sys_MoveCnt"))="20" then response.write " Selected"%>>30</option>
				<option value="30"<%if trim(request("sys_MoveCnt"))="30" then response.write " Selected"%>>40</option>
				<option value="40"<%if trim(request("sys_MoveCnt"))="40" then response.write " Selected"%>>50</option>
				<option value="50"<%if trim(request("sys_MoveCnt"))="50" then response.write " Selected"%>>60</option>
				<option value="60"<%if trim(request("sys_MoveCnt"))="60" then response.write " Selected"%>>70</option>
				<option value="70"<%if trim(request("sys_MoveCnt"))="70" then response.write " Selected"%>>80</option>
				<option value="80"<%if trim(request("sys_MoveCnt"))="80" then response.write " Selected"%>>90</option>
				<option value="90"<%if trim(request("sys_MoveCnt"))="90" then response.write " Selected"%>>100</option>
			</select>筆<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄 , <%=filsuess%>筆成功(<%=filClose%>筆結案) , <%=errCatCnt%> 筆無效  ,  <%=fildel%> 筆失敗 , <%=deldata%> 筆刪除  ,  <%=DBsum-CDbl(filsuess)-CDbl(fildel)-CDbl(deldata)-CDbl(errCatCnt)%>筆未處理. )</strong>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th class="font10">批號</th>
					
					<th class="font10" width="3%" nowrap>上傳日期</th>
					<th class="font10" nowrap>作業</th>
					
					<th class="font10" nowrap>結果</th>
					<th class="font10">訊息</th>					
					<th class="font10" width="3%" nowrap>上傳人員</th>
					<th class="font10" nowrap>類別</th>
					<th class="font10">舉發單號</th>
					<th class="font10">車號</th>
					<th class="font10">應到案處所</th>
					<th class="font10" width="3%" nowrap>廠牌.顏色<br>車藉狀況</th>
					
					<th class="font10">操作</th>
					<th  width="5%"><font size="1">上下載檔案.序號</font></th>
				</tr>
				<%
				if DB_Display="show" then
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end If 
					if Not rsfound.eof then rsfound.move DBcnt
					ReturnMarkType=split("2,3,4,5,6,7,8,9",",")
					ReturnMarkName=Split("入案,單退,寄存,公示,刪除,收受,撤消,結案",",")
					chkTypeID=0:chkBillNo=""
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound.eof then exit for
						response.write "<tr bgcolor='#FFFFFF'"
						lightbarstyle 0 
						response.write ">"

						CNum=""

						if instr(rsfound("BatchNumber"),"N")>0 then

							strSQL="select cnt from (select RowNum cnt,billsn,UserMarkDate from (select billsn,UserMarkDate from billmailhistory where billsn in (select BillSN from PasserDCILog where BatchNumber='"&trim(rsfound("BatchNumber"))&"') order by UserMarkDate) order by UserMarkDate) where BillSN="&rsfound("BillSN")

							set dci=conn.execute(strSQL)
							if not dci.eof then CNum=dci("cnt")
							dci.close

						else
							strSQL="select cnt from (select RowNum cnt,BillSN from (select BillSN from PasserDCILog where BatchNumber='"&trim(rsfound("BatchNumber"))&"' order by BillSN) order by BillSN) where BillSN="&rsfound("BillSN")

							set dci=conn.execute(strSQL)
							if not dci.eof then CNum=dci("cnt")
							dci.close
						end if

						response.write "<td class=""font10"" >"&rsfound("BatchNumber")&"&nbsp("&CNum&")"&"</td>"

						response.write "<td class=""font10"" nowrap>"&gInitDT(trim(rsfound("ExchangeDate")))&"</td>"
						
						if trim(rsfound("ExchangeTypeID"))="N" then
							response.write "<td class=""font10"" align=""center"">"
							for arr=0 to Ubound(ReturnMarkType)
								if trim(ReturnMarkType(arr))=trim(rsfound("ReturnMarkType")) then
									response.write ReturnMarkName(arr)
									exit for
								end if
							next
							if arr>Ubound(ReturnMarkType) then response.write "送達註記"
							response.write "&nbsp;</td>"
						else
							response.write "<td class=""font10"" align=""center"" >"&rsfound("DCIActionName")&"</td>"
						end if
						
						if trim(rsfound("DCIRETURNSTATUS"))="1" then
							response.write "<td class=""font10"" nowrap>正常</td>"
						elseif trim(rsfound("DCIRETURNSTATUS"))="-1" then
							response.write "<td class=""font10"" nowrap><font color=""red"">異常</font></td>"
						else
							response.write "<td class=""font10"" nowrap>未處理</td>"
						end If 

						DCIerror="":dciSQL=""
						if trim(rsfound("DCIReturnStatusID"))="00" then
							if trim(rsfound("DciErrorCarData"))<>"" then
								dciSQL="'"&rsfound("DciErrorCarData")&"'"
							end if
							if trim(rsfound("DCIErrorIDdata"))<>"" then
								if trim(dciSQL)<>"" then
									dciSQL=dciSQL&",'"&rsfound("DCIErrorIDdata")&"'"
								else
									dciSQL="'"&rsfound("DCIErrorIDdata")&"'"
								end if
							end if
							if trim(dciSQL)<>"" then
								strSQL="select DCIReturn,StatusContent from DCIReturnStatus where DCIActionID='"&rsfound("ExchangeTypeID")&"E' and DCIReturn in("&dciSQL&")"
								set rsdci=conn.execute(strSQL)
								while Not rsdci.eof
									if trim(DCIerror)<>"" then DCIerror=trim(DCIerror)&","
									DCIerror=trim(DCIerror)&rsdci("DCIReturn")&". "&rsdci("StatusContent")
									rsdci.movenext
								wend
								rsdci.close
							end if
						end if
						if trim(rsfound("BillTypeID"))="2" then
							strSQL="select ID,Content from DCICode where TypeID=10 and ID in(Select Rule4 from BillBaseDCIReturn where BillNo='"&rsfound("BillNo")&"' and CarNo='"&rsfound("CarNo")&"')"

							set rsdci=conn.execute(strSQL)
							while Not rsdci.eof
								if trim(DCIerror)<>"" then DCIerror=trim(DCIerror)&","
								DCIerror=trim(DCIerror)&rsdci("ID")&". "&rsdci("Content")
								rsdci.movenext
							wend
							rsdci.close
						end if

						Message=rsfound("DCIReturn")&". "&rsfound("StatusContent")
						'if trim(DCIerror)<>"" then Message=Message&"<br>"&DCIerror
						if trim(rsfound("CarErrorSN"))<>"" then Message=Message&"<br>"&rsfound("CarErrorSN")&". "&rsfound("CarErrorContent")
						if trim(rsfound("DCIErrorSN"))<>"" then Message=Message&"<br>"&rsfound("DCIErrorSN")&". "&rsfound("DCIErrorContent")

						response.write "<td class=""font10"" nowrap>"
						response.write Message
						response.write "</td>"
												
						'--------------------------------------------------------------
						response.write "<td class=""font10"" >"&rsfound("ChName")&"</td>"
						response.write "<td class=""font10"" >"&rsfound("BillTypeName")&"</td>"
						response.write "<td class=""font10"" >"&rsfound("BillNo")&"</td>"

						If i = (DBcnt+1) Then
							If not ifnull(rsfound("BillNo")) Then
								strSQL="select BillTypeID from PasserBase where billno='"&trim(rsfound("BillNo"))&"'"
								set chktype=conn.execute(strSQL)
								If not chktype.eof Then
									chkTypeID=cdbl(chktype("BillTypeID"))
									chkBillNo=trim(rsfound("BillNo"))
								End if
								chktype.close
							End if
						End if
						
						response.write "<td class=""font10""  nowrap>"&rsfound("CarNo")&"</td>"


						StrBass="select a.A_Name,a.DciReturnCarColor,c.ID as CarStatusID,c.Content as CarStatusName,d.ID as Rule4,d.Content as Rule4Name,e.DCIStationName from (select * from BillBaseDCIReturn where EXCHANGETYPEID='A'  and CarNo='"&rsfound("CarNo")&"') a,(select ID,Content from DCICode where TypeID=10) c,(select ID,Content from DCICode where TypeID=10) d,Station e where a.DCIReturnCarStatus=c.ID(+) and a.Rule4=d.ID(+) and a.DCIReturnStation=e.DCIStationID(+)"
						set rsCarType=conn.execute(strBass)
						Sys_DciReturnCarColor="":Sys_DCIStationName="":Sys_A_Name="":Sys_CarStatusID="":Sys_CarStatusName="":Sys_Rule4="":Sys_Rule4Name="":Sys_CarColorID="":Sys_CarColorName=""
						if not rsCarType.eof then
							Sys_DciReturnCarColor=trim(rsCarType("DciReturnCarColor"))
							Sys_DCIStationName=trim(rsCarType("DCIStationName"))
							Sys_A_Name=trim(rsCarType("A_Name"))
							Sys_CarStatusID=trim(rsCarType("CarStatusID"))
							Sys_CarStatusName=trim(rsCarType("CarStatusName"))
							Sys_Rule4=trim(rsCarType("Rule4"))
							Sys_Rule4Name=trim(rsCarType("Rule4Name"))
						end if
						rsCarType.close

						StrBass="select a.DciReturnCarColor,b.DCIStationName from (select * from BillBaseDCIReturn where EXCHANGETYPEID='W' and CarNo='"&trim(rsfound("CarNo"))&"' and BillNo='"&trim(rsfound("BillNo"))&"') a,Station b where a.DCIReturnStation=b.DCIStationID(+)"

						set rsCarType=conn.execute(strBass)
						if not rsCarType.eof then
							Sys_DciReturnCarColor=trim(rsCarType("DciReturnCarColor"))
							If trim(rsfound("ExchangeTypeID"))<>"A" then Sys_DCIStationName=trim(rsCarType("DCIStationName"))
						end if
						rsCarType.close

						if len(Sys_DciReturnCarColor)>1 then Sys_DciReturnCarColor=left(Sys_DciReturnCarColor,1)&","&right(Sys_DciReturnCarColor,1)
						if ifnull(Sys_DciReturnCarColor) then Sys_DciReturnCarColor=""
						Sys_CarColorID=split(Sys_DciReturnCarColor,",")
						for y=0 to ubound(Sys_CarColorID)
							strColor="select Content from DCICode where TypeID=4 and ID='"&Sys_CarColorID(y)&"'"
							set rscolor=conn.execute(strColor)
							if not rscolor.eof then
								if trim(Sys_CarColorName)<>"" then Sys_CarColorName=Sys_CarColorName&","
								Sys_CarColorName=Sys_CarColorName&trim(rscolor("Content"))
							end if
							rscolor.close
						next
						response.write "<td class=""font10""  nowrap>"&Sys_DCIStationName&"</td>"
						response.write "<td class=""font10"" nowrap > "&Sys_A_Name
						if trim(Sys_A_Name)<>"" then response.write ". "
						response.write Sys_CarColorName
						response.write "<br>"

						'response.write "<td class=""font10"" nowrap>"&rsCarType("CarColor")&"</td>"
						if not ifnull(Sys_CarStatusID) then response.write Sys_CarStatusID&"_"&Sys_CarStatusName

						if not ifnull(Sys_CarStatusID) and Not ifnull(Sys_Rule4) then response.write "<br>"
						if not ifnull(Sys_Rule4) then response.write Sys_Rule4&"_"&Sys_Rule4Name
						response.write "</td>"
					
						'--------------------------------------------------------------
						response.write "<td class=""font10"">"
												' 
						if (trim(rsfound("RecordMemberID"))=trim(Session("User_ID")) and trim(rsfound("DCIRETURNSTATUS"))="-1") or trim(Session("Credit_ID"))="A000000000" or trim(Session("Credit_ID"))="A01" Then
						%>
							<input type="button" name="b1" value="修改" class="btn3" style="width:40px; height:20px;" onclick='window.open("../BillKeyIn/BillKeyIn_People.asp?BillSN=<%=trim(rsfound("BillSN"))%>","WebPage2_Update","left=0,top=0,location=0,width=1000,height=650,resizable=yes,scrollbars=yes")' <%
								'1:查詢 ,2:新增 ,3:修改 ,4:刪除
								if CheckPermission(234,3)=false then
									response.write "disabled"
								end if
							%> style="font-size: 12pt; width: 45px; height:26px;"><%
						End If 

						response.write "&nbsp;</td>"
						'--------------------------------------------------------------
						response.write "<td font size=""1"" nowrap>"
						if trim(rsfound("DCIReturnStatusID"))<>"" then
							response.write "<a href='DCIfile.asp?DCIfile=/UP/"&trim(rsfound("FileName"))&"' target='_blank'><font size='1'>"&trim(rsfound("FileName"))&"</font>&nbsp;<font size='1' color=""Red"">"&trim(rsfound("seqNo"))&"</font></a><br>"
							response.write "<a href='DCIfile.asp?DCIfile=/Down/"&trim(rsfound("FileName"))&".big' target='_blank'><font size='1'>"&trim(rsfound("FileName"))&".big </font>&nbsp;<font size=""1"" color=""Red"">"&trim(rsfound("seqNo"))&"</font></a>"
						else
							response.write "<a href='DCIfile.asp?DCIfile=/UP/"&trim(rsfound("FileName"))&"' target='_blank'><font size='1'>"&trim(rsfound("FileName"))&"</font>&nbsp;<font size=""1""  color=""Red"">"&trim(rsfound("seqNo"))&"</font></a><br>"
							response.write "<font size='1'>" & trim(rsfound("FileName"))& "&nbsp;"&trim(rsfound("seqNo"))&"</font>"
						end if
						response.write "</td>"
						response.write "</tr>"
						response.flush
						rsfound.movenext
					next
				end if
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td height="30" bgcolor="#1BF5FF" align="center">
			<a href="file:///.."></a>
			<input type="button" name="MoveFirst" value="第一頁" class="btn3" style="width:50px; height:20px;" onclick="funDbMove(0);">
			<input type="button" name="MoveUp" value="上一頁" class="btn3" style="width:50px; height:20px;" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(CDbl(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(CDbl(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" class="btn3" style="width:50px; height:20px;" onclick="funDbMove(10);">
			<input type="button" name="MoveDown" value="最後一頁" class="btn3" style="width:60px; height:20px;" onclick="funDbMove(999);">
			<br>
			<input name="btnexit" type="button" value=" 產生舉發單 " class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funPrintLegal();">
			&nbsp;&nbsp;&nbsp;
			<input name="btnexit" type="button" value=" 產生送達證書 " class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funBillDeliver();">
			&nbsp;&nbsp;&nbsp;
			<input type="button" name="btnExecel" value="郵局大宗函件" class="btn3" style="width:120px;height:30px;font-size:14px;" onclick="funPasserMailMoney();">
		</td>
		
	  </tr>
</table>
<br><b>
* 系統上傳DCI 時間 :</b> 0850 ,  1050 ,  1250 , 1450 ,  1620 , 1850 。 請於各梯次 <b>前5分鐘</b> 上傳監理所。
<br>
<b>* DCI 抓取檔案時間  :</b> 0900 ,  1100 ,  1300 , 1500 ,  1630 , 1900 。 </b> 
<br>

<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="DB_Display" value="<%=DB_Display%>">
<input type="Hidden" name="DB_state" value="">
<input type="Hidden" name="SN" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="Hidden" name="BillSN" value="<%=BillSN%>">
<input type="Hidden" name="MailMoneyValue" value="">
</form>

</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
var winopen;
var Sys_City='<%=sys_City%>';
<%
if trim(session("ManagerPower"))="1" and sys_City="苗栗縣" then
	response.write "UnitMan('Sys_BillUnitID','Sys_BillMem','"&request("Sys_BillMem")&"');"
elseif Sys_City<>"苗栗縣" then
	response.write "UnitMan('Sys_BillUnitID','Sys_BillMem','"&request("Sys_BillMem")&"');"
end if
%>
function funSelt(DBKind){
	var error=0;
	if(DBKind=='Selt'){
		/*if(myForm.RecordDate.value!=""){
			if(!dateCheck(myForm.RecordDate.value)){
				error=1;
				alert("建檔日輸入不正確!!");
			}
		}
		if (error==0){
			if(myForm.RecordDate1.value!=""){
				if(!dateCheck(myForm.RecordDate1.value)){
					error=1;
					alert("建檔日輸入不正確!!");
				}
			}
		}*/
		if (error==0){
			if(myForm.ExchangeDate.value!=""){
				if(!dateCheck(myForm.ExchangeDate.value)){
					error=1;
					alert("上傳日輸入不正確!!");
				}
			}
		}
		if (error==0){
			if(myForm.ExchangeDate1.value!=""){
				if(!dateCheck(myForm.ExchangeDate1.value)){
					error=1;
					alert("上傳日輸入不正確!!");
				}
			}
			if (error==0){
				myForm.BillSN.value="";
				//CarForm.BillSN.value="";
				myForm.DB_Move.value="";
				myForm.DB_Selt.value=DBKind;
				myForm.DB_Display.value='show';
				myForm.submit();
			}
		}
	}else if(DBKind=='BatchSelt'){
		myForm.BillSN.value="";
		//CarForm.BillSN.value="";
		myForm.DB_Move.value="";
		myForm.DB_Selt.value=DBKind;
		myForm.DB_Display.value='show';
		myForm.submit();
	}
}

function funReSend(SN){
	if(SN=='ReSend'&&myForm.DB_Display.value!=""){
		if(myForm.Sys_BatchNumber.value!=""){
			if(confirm("是否確定要再次上傳?")){
				myForm.SN.value="";
				myForm.DB_state.value="ReSend";
				myForm.submit();
			}
		}else{
			alert("請先進行批號查詢!!");
		}
	}else{
		alert("請先進行批號查詢!!");
	}
}

function fnBatchNumber(){
	myForm.Sys_BatchNumber.value=myForm.Selt_BatchNumber.value;
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	winopen=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	winopen.focus();
	return win;
}

function funDbMove(MoveCnt){
	if (eval(MoveCnt)==0){
		myForm.DB_Move.value="";
		myForm.submit();
	}else if (eval(MoveCnt)==10){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10-eval(myForm.sys_MoveCnt.value)){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt+eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else if(eval(MoveCnt)==-10){
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt-eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else if(eval(MoveCnt)==999){
		if (eval(myForm.DB_Cnt.value)%(10+eval(myForm.sys_MoveCnt.value))==0){
			myForm.DB_Move.value=(Math.floor(eval(myForm.DB_Cnt.value)/(10+eval(myForm.sys_MoveCnt.value)))-1)*(10+eval(myForm.sys_MoveCnt.value));
		}else{
			myForm.DB_Move.value=Math.floor(eval(myForm.DB_Cnt.value)/(10+eval(myForm.sys_MoveCnt.value)))*(10+eval(myForm.sys_MoveCnt.value));
		}
		myForm.submit();
	}
}

function funPrintLegal(){	

	var UrlStr="";	
	
	if(Sys_City=='花蓮縣'){
		UrlStr="../PasserQuery/BillPrintLegal_YiLan_chromat_1110817.asp";
		//UrlStr="../PasserQuery/BillPrintLegal_CHCG_1110817.asp";

	}else if(Sys_City=='嘉義縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_Chiayi_1110817.asp";

	}else if(Sys_City=='嘉義市'){
		
		UrlStr="../PasserQuery/BillPrints_ChiayiCity_a4_1110817.asp";

	}else if(Sys_City=='高雄市'){
		
		UrlStr="../PasserQuery/BillPrintLegal_KaoHsiungCity_1110817.asp";
	}else if(Sys_City=='基隆市'){
		
		UrlStr="../PasserQuery/BillPrintLegal_KeeLung_1110817.asp";
	}else if(Sys_City=='金門縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_KMA_1110817.asp";
	}else if(Sys_City=='苗栗縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_miaoli_1110817.asp";
	}else if(Sys_City=='南投縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_NanTou_1110817.asp";
	}else if(Sys_City=='屏東縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_PingTung_1110817.asp";
	}else if(Sys_City=='台南市'){
		
		UrlStr="../PasserQuery/BillPrintLegal_TaiNanCity_1110817.asp";
	}else if(Sys_City=='宜蘭縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_YiLan_chromat_1110817.asp";
	}else if(Sys_City=='雲林縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_Yunlin_1110817.asp";
	}else if(Sys_City=='彰化縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_CHCG_1110817.asp";
	}else if(Sys_City=='臺東縣'){
		
		UrlStr="../PasserQuery/BillPrintsTaiTung_chromat_1110817.asp";
	}else if(Sys_City=='台中市'){
		
		UrlStr="../PasserQuery/BillPrints_TaiChungCity_1110817.asp";
	}else if(Sys_City=='連江縣'){
		
		UrlStr="../PasserQuery/BillPrints_lattice_MU.asp";
	}else if(Sys_City=='澎湖縣'){
		
		UrlStr="../PasserQuery/BillPrints_a4_penghu1120118.asp";
	}

	if(myForm.BillSN.value!=""){
		myForm.action=UrlStr;
		myForm.target="HuaLien";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行逕舉入案查詢");
	}
}

function funBillDeliver(){

	var UrlStr="../PasserBase/BillBase_Deliver_Word.asp";

	if(myForm.BillSN.value!=""){
		myForm.action=UrlStr;
		myForm.target="HuaLien";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行逕舉入案查詢");
	}
}

function funPasserMailMoney(){
	if(myForm.BillSN.value!=""){
		newWin("../PasserBase/PasserMailMoneyList_Select.asp","PasserMailMoneyList",400,200,50,10,"yes","yes","yes","no");
	}else{
		alert("請先進行逕舉入案查詢");
	}
}

function funPasserReportList(){

	UrlStr="../PasserBase/PasserReportList.asp";

	if(myForm.BillSN.value!=""){
		myForm.action=UrlStr;
		myForm.target="PasserReportList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行逕舉入案查詢");
	}
}

function repage(){
	myForm.DB_Move.value=0;
	myForm.submit();
}

</script>
<%conn.close%>