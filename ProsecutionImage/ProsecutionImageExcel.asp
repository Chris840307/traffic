<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include File="../Common/DbUtil.asp"-->
<!--#include File="../Common/Allfunction.inc"-->
<%
'AuthorityCheck(270)
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_數位相片舉發.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

if Conn.state then
	j = 1
	sFileName = ""
	sCarNo = ""
	sType = ""
	sImgWebPath = "/Img/finish/"
	
	'日期判斷
	if Request("StartDate") <> "" and Request("EndDate")<>"" then
		formStartDate = gOutDT(request("StartDate"))&" 0:0:0"
		formEndDate = gOutDT(request("EndDate"))&" 23:59:59"
		
		formYear = Request.Form("Year")
		formStartMonth = Request.Form("StartMonth")
		formStartDay = Request.Form("StartDay")
		formEndMonth = Request.Form("EndMonth")
		formEndDay = Request.Form("EndDay")

		sWhere = sWhere + " and PROSECUTIONTIME between TO_DATE('"&formStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&formEndDate&"','YYYY/MM/DD/HH24/MI/SS')"
	end if
	'舉發路段判斷
	if Request("Location") <> "" then
		sWhere = sWhere + " and LOCATION like '%" + Request("Location") + "%'"
		formLocation = Request("Location")
	end if
	
	'類型判斷
	if Request("ProsecutionType") <> "" then
		sWhere = sWhere + " and PROSECUTIONTYPEID = '" + Request("ProsecutionType") + "'"
		formProsecutionType = Request("ProsecutionType")
	end if
		
	'有效性判斷
	if Request("VerifyResult") <> "" then
		if Request("VerifyResult") <> "NO" then
			if Request("VerifyResult")="2" then
				sWhere = sWhere + " and VERIFYRESULTID = 0 and BillSn is null"
				formVerifyResult = Request("VerifyResult")
			else
				sWhere = sWhere + " and VERIFYRESULTID = " + Request("VerifyResult")
				formVerifyResult = Request("VerifyResult")
			end if
		end if 
	else
		sWhere = sWhere + " and VERIFYRESULTID = 1 "
		formVerifyResult = Request("VerifyResult")
	end if
	
	'車牌號碼判斷
	if Request("CarNo") <> "" then
		sWhere = sWhere + " and (b.CARNO like '%" + UCASE(Request("CarNo")) + "%' OR REALCARNO like '%" + UCASE(Request("CarNo")) + "%') "
		formCarNo = UCASE(Request("CarNo"))
	end if
	'舉發人判斷 by kevin
	if Request("OperatorID") <> "" then
		sWhere = sWhere + " and (a.OperatorA = '" + Request("OperatorID") + "' or a.OperatorB = '" + Request("OperatorID") + "')" 
		formOperator = Request("OperatorID")
	end if
	'新增 SpecCar 資料 2006/10/25
	sSQL = "select OVERSPEED,LIMITSPEED,LINE,a.FILENAME, case when FIXEQUIPTYPE = 1 then 'Type1' when FIXEQUIPTYPE = 2 then 'Type2' else 'Type3' end FIXEQUIPTYPE, a.DIRECTORYNAME," + _
			" a.PROSECUTIONTIME, a.LOCATION, a.PROSECUTIONTYPEID, a.IMAGEFILENAMEA, a.IMAGEFILENAMEB, a.VIDEOFILENAME, b.VERIFYRESULTID, b.SN, " + _
			" b.CARNO, b.MEMBERID,b.REALCARNO, c.CHNAME, d.CARSN " + _
			"from ProsecutionImage a, ProsecutionImageDetail b, MEMBERDATA c, SpecCar d " + _
			"where a.FILENAME = b.FILENAME and b.MEMBERID = c.MEMBERID(+) and b.CARNO = d.CARNO(+) " + sWhere + "  order by FIXEQUIPTYPE desc,Location,PROSECUTIONTypeID,PROSECUTIONTIME desc"
			
			'response.write sSQL
'response.end
	set oRST = Conn.execute(sSQL)
	if not oRST.EOF then
		do while not oRST.EOF					
			'車牌號碼
			sCarNo = ""
			if oRST("VerifyResultID").value = "1" or isnull(oRST("VerifyResultID").value) then
				sMemberName = ""
									
				sCarNo = oRST("CarNo").value & "&nbsp;"
			else
				sMemberName = oRST("CHNAME").value
				
				'車牌號碼
				sCarNo = oRST("RealCarNo").value		
			end if
			
			'類型
			if trim(oRST("FIXEQUIPTYPE").value) = "Type1" then
				sFIXEQUIPTYPE = "數位桿"
			elseif trim(oRST("FIXEQUIPTYPE").value) = "Type2" then
				sFIXEQUIPTYPE = "升級桿"
			else 
				sFIXEQUIPTYPE = "相機"
			end if
			
			'審核狀態
			if oRST("VerifyResultID").value = "1" or isnull(oRST("VerifyResultID").value) then
				sVerifyResultID = "未審核"
			else
				if oRST("VerifyResultID").value = "0" then
					sVerifyResultID = "有效"
				else
					sVerifyResultID = "無效"
				end if
			end if
			
			'違規類型
			sTypeID = oRST("ProsecutionTypeID").value
			select case(sTypeID)
				case "R":	sTypeID = "闖紅燈"
				case "S":	sTypeID = "超速"
							sLIMITSPEED = oRST("LIMITSPEED").value
							sOVERSPEED = oRST("OVERSPEED").value
							sLINE = oRST("LINE").value
				case "L":	sTypeID = "越線"
				case "T":	sTypeID = "緊跟著前車駕駛"
				case "U":	sTypeID = "左轉"
				defaule:	sTypeID = "未知類型"
			end select
			if not isnull(oRST("ProsecutionTime")) then 
				sProsecutionDate= gInitDT(oRST("ProsecutionTime").value)
				sProsecutionTime= TimeValue(oRST("ProsecutionTime").value)
			else 
				sProsecutionDate=""
				sProsecutionTime=""
			end if
			sTR = sTR + "<tr bgcolor=""#FFFFFF"" " & sMylightbarstyle & ">" + _
						"	<td height=""25"" nowrap>" & sProsecutionDate & " " & sProsecutionTime & "</td>" + _
						"	<td>" & oRST("Location").value & "</td>" + _
						"	<td>" & sTypeID & "</td>" + _
						"	<td>" & sLIMITSPEED & "</td>" + _
						"	<td>" & sOVERSPEED & "</td>" + _
						"	<td>" & sLINE & "</td>" + _
						"</tr>"	
						'"	<td>" & sFIXEQUIPTYPE & "</td>" + _
						'"	<td>@@CarNo@@</td>" + _
						'"	<td>@@VerifyResult@@</td>" + _
						
			j = j + 1	
		if sCarNo<>"" and not isnull(sCarNo) then
			sTR = replace(sTR, "@@CarNo@@", sCarNo)
		end if
			sTR = replace(sTR, "@@VerifyResult@@", sVerifyResultID)
			
			oRST.movenext
		loop		
	end if
	oRST.close
	set oRST = nothing
end if
	
if isObject(Conn) then
    if Conn.state then
        Conn.close
    end if
    Set Conn = Nothing
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>數位相片舉發-資料列表</title>
</head>
<body>
<table width="100%" border="1">
	<tr>
		<td height="26" align="center"><strong>數位相片舉發-資料列表</strong></td>
	</tr>
	<tr>
		<td>
			<table border="1" cellpadding="4" cellspacing="1">
		      <tr bgcolor="#EBFBE3">
					        <th width="120" height="15" nowrap><span class="style3">違規日期</span></th>
					        <th width="120" nowrap><span class="style3">路段</span></th>
					    <!--<th width="40" nowrap><span class="style3">來源</span></th>
					        <th width="80" nowrap><span class="style3">車牌</span></th>-->
					        <th width="60" nowrap><span class="style3">類型</span></th>
					        <th width="40" nowrap>限速</th>
					        <th width="40" nowrap>車速</th>
					        <th width="40" nowrap>車道</th>
					    <!--<th width="40" nowrap>審核</th>-->
		      </tr>
				<%=sTR%>
			</table>
		</td>
	</tr>
</table>
</body>
</html>

