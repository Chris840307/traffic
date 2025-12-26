<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include File="../Common/DbUtil.asp"-->
<!--#include File="../Common/Allfunction.inc"-->
<!--#include file="../Common/Banner.asp"-->
<%
AuthorityCheck(270)
Dim sMessage:sumOverSpeed=0
sMessage = ""

'固定測試人員
gUser_ID = Session("User_ID")
gCh_Name = session("CH_Name")
gUnit_ID = Session("Unit_ID")

'gUser_ID = "1"
'gCh_Name = "Max"
'gUnit_ID = "Group"

'***********************************************************************************
'傳送資料處理
'***********************************************************************************
	if Request.Form("SubmitType") <> "" then
		'============更改審核狀況時==============
		if Request.Form("SubmitType") = "UpdateVerifyResult" then
				'將單一筆變成有效(只更新 REALCARNO,保留 CARNO,以方便以後驗證用)
				sSQL ="update ProsecutionImageDetail set MEMBERID = " & gUser_ID & ", VerifyResultID = " & Request.Form("SelValue") & ",REALCARNO = '" & UCASE(Request.Form("SelCarNo")) & "' where FileName = '" & Request.Form("SelFileName") & "' and SN = '" & Request.Form("SelSN") & "' " + vbcrlf
				Conn.execute(sSQL)
				
				'response.write sSQL
				
				sMessage = "審核完成!!"
		end if

		'============逕舉建檔=================
		if Request.Form("SubmitType") = "InsertBillBase" then
			'***********************************************************************
			'BillBase 資料處理
			'***********************************************************************
			Dim bOK,sMaxSN, sValue, piPROSECUTIONTIME,piLOCATION,piTRIGGERSPEED,piLIMITSPEED,piSITECODE,piIMAGEFILENAMEA,piIMAGEPATHNAME,sCarSimpleID,sLawItemID
			
			bOK = true
			'抓 ApConfigure id=3 Value
			sSQL = "SELECT VALUE FROM ApConfigure WHERE ID=3"
			set oRST = Conn.execute(sSQL)
			if not oRST.EOF then
				sValue = oRST(0).value
			else
				sMessage = sMessage + "無法取得 ApConfigure 資料!!" + vbcrlf
				bOK = false
			end if
			oRST.close
			set oRST = nothing
			'抓
			sSqlDetail="select CarSimpleID,LawItemID from ProsecutionImageDetail where FileName='"&Request.Form("SelFileName")&"' and SN='" & Request.Form("SelSN") & "'"
			set rsDetail=Conn.execute(sSqlDetail)
			if not rsDetail.eof then
				sCarSimpleID=trim(rsDetail("CarSimpleID"))
				sLawItemID=trim(rsDetail("LawItemID"))
			end if
			rsDetail.close
			set rsDetail=nothing

			'抓 ProsecutionImage 資料
			sSQL = "select PROSECUTIONTIME,LOCATION,TRIGGERSPEED,LIMITSPEED,SITECODE,IMAGEFILENAMEA, FILENAME,case when FIXEQUIPTYPE = 1 then 'Type1' when  FIXEQUIPTYPE = 2 then 'Type2' when  FIXEQUIPTYPE = 5 then 'Type5' else 'Type3' end FIXEQUIPTYPE from ProsecutionImage where FileName = '" & Request.Form("SelFileName") & "'"
			set oRST = Conn.execute(sSQL)
			if not oRST.EOF then
				piPROSECUTIONTIME = oRST(0).value
				piLOCATION = oRST(1).value
				piTRIGGERSPEED = oRST(2).value
				piLIMITSPEED = oRST(3).value
				piSITECODE = oRST(4).value
		
				if trim(oRST("FIXEQUIPTYPE").value)="Type1" or trim(oRST("FIXEQUIPTYPE").value)="Type2" then
					piIMAGEPATHNAME = "\" & oRST("FIXEQUIPTYPE").value & "\" & oRST("FILENAME").value & "\"
					piIMAGEFILENAMEA = oRST(5).value
				else
					piIMAGEPATHNAME = "\" & oRST("FIXEQUIPTYPE").value & "\" & left(trim(oRST("FILENAME").value),14) & "\"
					piIMAGEFILENAMEA = right(trim(oRST(5).value),8)
				end if
				if piPROSECUTIONTIME="" or isnull(piPROSECUTIONTIME) then
					piPROSECUTIONTIME = "null"
				else
					piPROSECUTIONTIME = FormatDateTime(piPROSECUTIONTIME, 2) + " " + FormatDateTime(piPROSECUTIONTIME, 4)
					piPROSECUTIONTIME = "TO_DATE('" & FormatDateTime(piPROSECUTIONTIME, 2) + " " + FormatDateTime(piPROSECUTIONTIME, 4) & "', 'YYYY/MM/DD HH24:MI')"
				end if
				if piTRIGGERSPEED="" or isnull(piTRIGGERSPEED) then
					piTRIGGERSPEED="null"
				end if
				if piLIMITSPEED="" or isnull(piLIMITSPEED) then
					piLIMITSPEED="null"
				end if
			else
				sMessage = sMessage + "無法取得 ProsecutionImage 資料!!" + vbcrlf
				bOK = false
			end if
			oRST.close
			set oRST = nothing
			'抓取使用者臂章號碼
			if gUser_ID<>"" and not isnull(gUser_ID) then
				strUser="select LoginID from MemberData where MemberID="&gUser_ID
				set rsUser=conn.execute(strUser)
				if not rsUser.eof then
					gLogin_ID=trim(rsUser("LoginID"))
				end if
				rsUser.close
				set rsUser=nothing
			end if
			'***********************************************************************
			
			if bOK = true then
				'抓最大值
				sSQL = "select BillBase_seq.nextval as SN from Dual"
				set oRST = Conn.execute(sSQL)
				if not oRST.EOF then
					sMaxSN = oRST("SN")
				end if
				oRST.close
				set oRST = nothing
				'更新PID的BILLSN
				strUpdate="Update ProsecutionImageDetail set BillSn="&sMaxSN&" where FileName='"&Request.Form("SelFileName")&"' and SN='" & Request.Form("SelSN") & "'"
				Conn.execute strUpdate

				sSQL = "INSERT INTO BillBase (SN, BILLTYPEID,CARNO,ILLEGALDATE,ILLEGALADDRESS,ILLEGALSPEED,RULESPEED,USETOOL,BILLUNITID,BILLMEMID1,BILLMEM1,BILLFILLERMEMBERID,BILLFILLER,BILLFILLDATE,BILLSTATUS,RECORDSTATEID,RECORDDATE,RECORDMEMBERID,EQUIPMENTID,RULEVER,IMAGEFILENAME,IMAGEPATHNAME,CarSimpleID,Rule1) VALUES " + _
					"(" & sMaxSN & ",2,'" & UCASE(Request.Form("SelCarNo")) & "'," & piPROSECUTIONTIME & ",'" & piLOCATION & "'," & piTRIGGERSPEED & "," & piLIMITSPEED & ",1,'" & gUnit_ID & "','" & gLogin_ID & "','" & gCh_Name & "'," & gLogin_ID & ",'" & gCh_Name & "',SYSDATE,-1,0,SYSDATE," & gUser_ID & ",'" & piSITECODE & "','" & sValue & "','" & piIMAGEFILENAMEA & "', '" & piIMAGEPATHNAME & "','" & sCarSimpleID & "','" & sLawItemID & "') "
				Conn.execute(sSQL)
				'response.write sSQL
				'response.end
%>
<script language="javascript">
	window.open('../BillKeyIn/BillKeyIn_Car_Report_Update.asp?BillSN='+<%=sMaxSN%>,'MyFromCaseIn','left=0,top=0,width=1000,height=615,resizable=yes,scrollbars=yes');
</script>
<%
			else
				sMessage = "建檔發生錯誤如下,請通知系統管理員:" + vbcrlf + sMessage
			end if
		end if 
	end if
'***********************************************************************************
'主資料處理
'***********************************************************************************
Dim i
Dim j	'機算需要補多少空白
Dim oRST
Dim sTypeID	'ProsecutionTypeID
Dim sFileName, sVerifyResult
Dim sSQL, sOption, sRoadOption, sTypeOption
Dim sMylightbarstyle
Dim sMenuButton	
Dim sWhere	
Dim formLocation, formFIXEQUIPTYPE, formVerifyResult, formCarNo, formStartDate, formEndDate
Dim formYear, formStartMonth, formStartDay, formEndMonth, formEndDay
Dim formOperator

'if Request.Form("SubmitType") = "DataSearch" then
	j = 1
	sFileName = ""
	sCarNo = ""
	sType = ""

	'分業處理
	sAllPage = 1
	sNowPage = 1
	if Request.QueryString("page") <> "" then
		sNowPage = Request.QueryString("page")
	end if	
	'日期判斷 by kevin
	if Request("StartDate") <> "" and Request("EndDate") <> "" then
		formStartDate =gOutDT(request("StartDate"))&" 0:0:0"
		formEndDate =gOutDT(request("EndDate"))&" 23:59:59"
		
		formYear = Request("Year")
		formStartMonth = Request("StartMonth")
		formStartDay = Request("StartDay")
		formEndMonth = Request("EndMonth")
		formEndDay = Request("EndDay")

		sWhere = sWhere + " and PROSECUTIONTIME between TO_DATE('"&formStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&formEndDate&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	'舉發路段判斷
	if Request("Location") <> "" then
		sWhere = sWhere + " and LOCATION like '%" + Request("Location") + "%'"
		formLocation = Request("Location")
	end if
	
	'類型判斷
'	if Request("FIXEQUIPTYPE") <> "" then
'		sWhere = sWhere + " and FIXEQUIPTYPE = '" + Request("FIXEQUIPTYPE") + "'"
'		formFIXEQUIPTYPE = Request("FIXEQUIPTYPE")
'	end if
		
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
	if Request("sys_CarNo") <> "" then
		sWhere = sWhere + " and (b.CARNO like '%" + UCASE(Request("sys_CarNo")) + "%' OR REALCARNO like '%" + UCASE(Request("sys_CarNo")) + "%') "
		formCarNo = UCASE(Request("sys_CarNo"))
	end if
	
	'舉發人判斷 by kevin
	if Request("OperatorID") <> "" then
		sWhere = sWhere + " and (OperatorA = '" + Request("OperatorID") + "' or OperatorB = '" + Request("OperatorID") + "')" 
		formOperator = Request("OperatorID")
	end if

	if Request("Sys_ProsecutionTypeID") <> "" then
		sWhere = sWhere + " and a.ProsecutionTypeID = '" + Request("Sys_ProsecutionTypeID") + "'"
	end if
	sMylightbarstyle = "onMouseOver=""this.style.backgroundColor='#FF99FF'"" onMouseOut=""this.style.backgroundColor='#FFFFFF'"""
	
	'新增 SpecCar 資料 2006/10/25
	sSQL = "select OVERSPEED,LIMITSPEED,LINE,a.FILENAME, case when FIXEQUIPTYPE = 1 then 'Type1' when FIXEQUIPTYPE = 2 then 'Type2' when FIXEQUIPTYPE = 5 then 'Type5' when FIXEQUIPTYPE = 3 then 'Type3' else 'Type10' end  FIXEQUIPTYPE, a.DIRECTORYNAME," + _
			" a.PROSECUTIONTIME, a.LOCATION, a.PROSECUTIONTYPEID, a.IMAGEFILENAMEA, a.IMAGEFILENAMEB, a.VIDEOFILENAME, a.RejectReason, b.VERIFYRESULTID, b.SN, " + _
			" b.CARNO, b.MEMBERID,b.REALCARNO,b.BillSN, c.CHNAME, d.CARSN " + _
			"from ProsecutionImage a, ProsecutionImageDetail b, MEMBERDATA c, SpecCar d " + _
			"where a.FILENAME = b.FILENAME and b.MEMBERID = c.MEMBERID(+) and b.CARNO = d.CARNO(+) " + sWhere + "  order by FIXEQUIPTYPE desc,Location,PROSECUTIONTypeID,PROSECUTIONTIME desc"
			
			'response.write sSQL
			'response.end

	'set oRST = Conn.execute(sSQL)
	set oRST = CreateObject("ADODB.Recordset")
	oRST.cursorlocation = 3	
	oRST.Open sSQL, Conn, 3,1

	if not oRST.eof Then
		sImgWebPath = toImageDir(oRST("PROSECUTIONTIME"))
		oRST.pagesize=PageSize
		sAllPage = oRST.pagecount
		
		if Cint(sNowPage)>Cint(sAllPage) then
			sNowPage = sAllPage
		end if		
		oRST.absolutepage=sNowPage

	'if not oRST.EOF then
		for i=1 to oRST.pagesize
		'do while not oRST.EOF							
			'審核狀態
			sCarNo = ""
			if oRST("VerifyResultID").value =  "1" or isnull(oRST("VerifyResultID").value) then
				sMemberName = ""
				sVerifyResultID = "未處理"
				'"<button onClick=""UpdateData('select_" & j & "', '" & oRST("FileName").value & "', '0', '" & oRST("SN").value & "')"">有效</button>" + _
				'"<button onClick=""UpdateData('select_" & j & "', '" & oRST("FileName").value & "', '-1', '" & oRST("SN").value & "')"">無效</button>"
									
				'車牌號碼 (含 SpaceCar)
				if oRST("CARSN").value = "" or isnull(oRST("CARSN").value) then
				 sCarNo = "<input id=""select_" & j & """ type=""text"" size=9 maxlength=8 value=""" & oRST("CarNo").value & """  onkeyup='value=value.toUpperCase()'/>"
				else
				 sCarNo = "<b><font color=red>*</font></b> <input id=""select_" & j & """ type=""text"" size=9 maxlength=8 value=""" & oRST("CarNo").value & """  onkeyup='value=value.toUpperCase()'/>"
				end if
			else
				sMemberName = oRST("CHNAME").value
				
				'車牌號碼
				if trim(oRST("RealCarNo").value)<>"" and not isnull(oRST("RealCarNo").value) then 
					sCarNo = oRST("RealCarNo").value
				else
					sCarNo=" "
				end if
			
				if oRST("VerifyResultID").value = "0" then
					sVerifyResultID = "有效"
				else
					sVerifyResultID = "無效"
				end if
			end if
			'無效就不出現建檔鈕
			'建檔按鈕(BillSN有值的話直接進入修改頁面 不做新增)
			if oRST("VerifyResultID").value = "0" then
				if trim(oRST("BillSN").value)<>"" and not isnull(oRST("BillSN").value) then
					sCaseInButton="<input type='button' value='舉發單' onClick=""OpenCaseData('" & trim(oRST("BillSN").value) & "')"">"
				else
					sCaseInButton=""
'					sCaseInButton="<input type='button' value='建檔' onClick=""funcInCaseData('" & trim(oRST("FileName").value) & "','" & trim(oRST("SN").value) & "','" & trim(oRST("REALCARNO").value) & "')"">"
				end if
			else
				sCaseInButton=""
			end if

			'類型
			if trim(oRST("FIXEQUIPTYPE").value) = "Type1" then
				sFIXEQUIPTYPE = "數位桿"
			elseif trim(oRST("FIXEQUIPTYPE").value) = "Type2" then
				sFIXEQUIPTYPE = "升級桿"
			else 
				sFIXEQUIPTYPE = "相機"
			end if
			
			'按紐功能
			sMenuButton = ""
			sOVERSPEED= ""
			'違規類型
'		if trim(oRST("FIXEQUIPTYPE").value)="Type1" or trim(oRST("FIXEQUIPTYPE").value)="Type2" then
			sSmallPicFileName = replace(sImgWebPath & replace(oRST("DIRECTORYNAME"),"\","/") & lcase(oRST("IMAGEFILENAMEA").value),"//","/")
'		elseif trim(oRST("FIXEQUIPTYPE").value)="Type5" then
'			sSmallPicFileName = sImgWebPath & oRST("DIRECTORYNAME") & "/" & lcase(oRST("IMAGEFILENAMEA").value)
'		else
'			sSmallPicFileName = sImgWebPath & oRST("DIRECTORYNAME") & "/" & right(oRST("FileName").value,4) & ".jpg"
'		end if
			sTypeID = oRST("ProsecutionTypeID").value
			sLIMITSPEED=""
			select case(sTypeID)
				case "R":	sTypeID = "闖紅燈"
							sLIMITSPEED = oRST("LIMITSPEED").value
							sOVERSPEED = "<b><font color=red>" & oRST("OVERSPEED").value & "</font></b>"
							sLINE = oRST("LINE").value

							if oRST("VideoFileName").value <> "" then
								sMenuButton = "<input type=""button"" onclick=""OpenPic2('" & sImgWebPath & oRST("FIXEQUIPTYPE").value & "/" & oRST("FileName").value & "/" & oRST("VideoFileName").value & "')"" value=""動態"">"
							end if
							'************員警上傳影像kevin******************
							if trim(oRST("FIXEQUIPTYPE"))="Type3" then	
								bPicWebPath=""
								RealFileName=right(oRST("FileName").value,4)
								WebPicPathTmp=left(oRST("FileName").value,14)
								bPicWebPath=sImgWebPath & "upload/" & WebPicPathTmp & "/" & RealFileName & ".jpg"
								sMenuButton = sMenuButton + "<input type=""button"" onClick=""OpenPic('" & bPicWebPath & "')"" value=""圖一"">"

								if oRST("IMAGEFILENAMEB").value <> "" then
									sPicWebPath=""
									sRealFileName=right(replace(oRST("ImageFileNameB"),".jpg",""),4)
									sWebPicPathTmp=left(replace(oRST("ImageFileNameB"),".jpg",""),14)

									sPicWebPath=sImgWebPath & "upload/" & sWebPicPathTmp & "/" & sRealFileName & ".jpg"
									sMenuButton = sMenuButton + "<input type=""button"" onClick=""OpenPic('" & sPicWebPath & "')"" value=""圖二"">"
								end if
							else
								sMenuButton = sMenuButton + "<input type=""button"" onClick=""OpenPic('" & sSmallPicFileName & "')"" value=""圖一"">"
								if oRST("IMAGEFILENAMEB").value <> "" then
									sMenuButton = sMenuButton + "<input type=""button"" onClick=""OpenPic('" & replace(sImgWebPath & replace(oRST("DIRECTORYNAME").value,"\","/") & lcase(oRST("IMAGEFILENAMEB").value),"//","/") & "')"" value=""圖二"">"
								end if
							end if
							'***********************************************
							'目前並沒有檢視原圖,因為縮圖就已經很清楚了,而且檔案又小.
							'"<input type=""button"" onClick=""OpenPic('" & sImgWebPath & oRST("DirectoryName").value & oRST("FileName").value & "')"" value=""原圖"">" + _
							sMenuButton = sMenuButton + "<input type=""button"" onClick=""OpenDetail('" & oRST("FileName").value & "','" & oRST("SN").value & "')"" value=""詳細"">"
				case "S":	sTypeID = "超速"
							sLIMITSPEED = oRST("LIMITSPEED").value
							sOVERSPEED = "<b><font color=red>" & oRST("OVERSPEED").value & "</font></b>"
							sLINE = oRST("LINE").value
							
							sumOverSpeed=sumOverSpeed+cdbl(oRST("OVERSPEED").value)
							if oRST("VideoFileName").value <> "" then
								sMenuButton = "<input type=""button"" onclick=""OpenPic2('" & sImgWebPath & oRST("FIXEQUIPTYPE").value & "/" & oRST("FileName").value & "/" & oRST("VideoFileName").value & "')"" value=""動態"">"
							end if
							'************員警上傳影像kevin******************
							if trim(oRST("FIXEQUIPTYPE"))="Type3" then	
								bPicWebPath=""
								RealFileName=right(oRST("FileName").value,4)
								WebPicPathTmp=left(oRST("FileName").value,14)

								bPicWebPath=sImgWebPath & "upload/" & WebPicPathTmp & "/" & RealFileName & ".jpg"
								sMenuButton = sMenuButton + "<input type=""button"" onClick=""OpenPic('" & bPicWebPath & "')"" value=""圖一"">"
							else
								If lcase(oRST("IMAGEFILENAMEA").value)<>"" Then
									sMenuButton = sMenuButton + "<input type=""button"" onClick=""OpenPic('" & sSmallPicFileName & "')"" value=""圖一"">"
								end if
							end if
							'***********************************************
							'目前並沒有檢視原圖,因為縮圖就已經很清楚了,而且檔案又小.
							'"<input type=""button"" onClick=""OpenPic('" & sImgWebPath & oRST("DirectoryName").value & oRST("FileName").value & "')"" value=""原圖"">" + _
							sMenuButton = sMenuButton + "<input type=""button"" onClick=""OpenDetail('" & oRST("FileName").value & "','" & oRST("SN").value & "')"" value=""詳細"">"
							'response.write sSmallPicFileName&"  qqq"
				case "SR":	sTypeID = "超速、闖紅燈"
							sLIMITSPEED = oRST("LIMITSPEED").value
							sOVERSPEED = "<b><font color=red>" & oRST("OVERSPEED").value & "</font></b>"
							sLINE = oRST("LINE").value

							If lcase(oRST("IMAGEFILENAMEA").value)<>"" Then
								sMenuButton = sMenuButton + "<input type=""button"" onClick=""OpenPic('" & sSmallPicFileName & "')"" value=""圖一"">"
							end if
							if oRST("IMAGEFILENAMEB").value <> "" then
								sMenuButton = sMenuButton + "<input type=""button"" onClick=""OpenPic('" & sImgWebPath & oRST("DIRECTORYNAME").value & "/" & lcase(oRST("IMAGEFILENAMEB").value) & "')"" value=""圖二"">"
							end if

							if oRST("VideoFileName").value <> "" then
								sMenuButton = "<input type=""button"" onclick=""OpenPic2('" & sImgWebPath & oRST("FIXEQUIPTYPE").value & "/" & oRST("FileName").value & "/" & oRST("VideoFileName").value & "')"" value=""動態"">"
							end if
							
							'目前並沒有檢視原圖,因為縮圖就已經很清楚了,而且檔案又小.
							'"<input type=""button"" onClick=""OpenPic('" & sImgWebPath & oRST("DirectoryName").value & oRST("FileName").value & "')"" value=""原圖"">" + _
							sMenuButton = sMenuButton + "<input type=""button"" onClick=""OpenDetail('" & oRST("FileName").value & "','" & oRST("SN").value & "')"" value=""詳細"">"
				case "L":	sTypeID = "越線"
				case "T":	sTypeID = "緊跟著前車駕駛"
				case "U":	sTypeID = "左轉"
				case "E":	sTypeID = "錯行"
				'defaule:	sTypeID = "未知類型"
			end select
			if not isnull(oRST("ProsecutionTime")) then 
				sProsecutionDate= gInitDT(oRST("ProsecutionTime").value)
				sProsecutionTime= TimeValue(oRST("ProsecutionTime").value)
			else 
				sProsecutionDate=""
				sProsecutionTime=""
			end if
			if not isnull(oRST("RejectReason")) and trim(oRST("RejectReason"))<>"" then 
				sRejectReason= trim(oRST("RejectReason").value)
			else 
				sRejectReason=""
			end if
			sTR = sTR + "<tr bgcolor=""#FFFFFF"" " & sMylightbarstyle & ">" + _
						"	<td height=""25"" nowrap><font size=""2"">" &  sProsecutionDate & sProsecutionTime & "</font></td>" + _
						"	<td><span class=""font11"">" & oRST("Location").value & "</span></td>" + _
						"	<td><font size=""2"">" & sTypeID & "</font></td>" + _
						"	<td><span class=""font9"">" & sLIMITSPEED & "</span></td>" + _
						"	<td><span class=""font9"">" & sOVERSPEED & "</span></td>" + _
						"	<td><font size=""2"">" & sRejectReason & "</font></td>" + _
						"	<td>" & sMenuButton & "</td>" + _
						"</tr>"		
						
			j = j + 1		
			'"	<td><span class=""font12"">@@CarNo@@</span></td>"
			'"	<td><span class=""font9""> @@VerifyResult@@ " & sCaseInButton & "</span></td>"
			if sCarNo <> "" then
				sTR = replace(sTR, "@@CarNo@@", sCarNo)
				sTR = replace(sTR, "@@VerifyResult@@", sVerifyResultID)
			end if
			
			oRST.movenext
		'loop		
		if oRST.eof then exit for
		next
		
	'end if
	
	end if
	oRST.close
	set oRST = nothing
'end if

'補空白
for i = j to 10
	sTR = sTR + "<tr bgcolor=""#FFFFFF"">" + _
				"	<td height=""25""></td>" + _
				"	<td></td>" + _
				"	<td></td>" + _
				"	<td></td>" + _
				"	<td></td>" + _
				"	<td></td>" + _
				"	<td></td>" + _
				"</tr>"
next
'***********************************************************************************
'舉發路段資料處理
'***********************************************************************************
	sRoadOption = ""
	sSQL = "select distinct(ProsecutionImage.LOCATION) as location from ProsecutionImage"

	set oRST = Conn.execute(sSQL)
	if not oRST.EOF then
		do while not oRST.EOF
			sRoadOptionSelected=""
			if trim(oRST("location").value)=trim(request("Loaction")) then
				sRoadOptionSelected=" selected"
			end if	
			sRoadOption = sRoadOption + "<option value=""" & oRST("location").value & """" & sRoadOptionSelected & ">" & oRST("location").value & "</option>"
			oRST.movenext
		loop
	end if
	oRST.close
	set oRST = nothing
'***********************************************************************************
'類型資料處理
'***********************************************************************************
	sTypeOption = ""
	sID = ""
	sSQL = "select Content,ID from Code where TypeID=18"

	set oRST = Conn.execute(sSQL)
	if not oRST.EOF then
		do while not oRST.EOF
			sID = oRST("ID").value
			if sID = "472" then
				sID = "1"
			end if
			
			if sID = "471" then
				sID = "2"
			end if
			if sID = "469" then
				sID = "3"
			end if
			sTypeOptionSelected=""
			if trim(sID)=trim(request("FIXEQUIPTYPE")) then
				sTypeOptionSelected=" selected"
			end if
			sTypeOption = sTypeOption + "<option value=""" & sID & """" & sTypeOptionSelected & ">" & oRST("Content").value & "</option>"
			oRST.movenext
		loop
	end if
	oRST.close
	set oRST = nothing
'***********************************************************************************


%>


<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5">
	<title>數位固定桿處理系統</title>
	<!--#include File="../Common/css.txt"-->
<style type="text/css">
<!--
.btn2 { FONT-SIZE: 9pt; FONT-FAMILY: Arial;}
-->
</style>
</head>

<body topmargin="0" leftmargin="0" onfocus="ClosePic();">
<form method="post" action="ProsecutionImage.asp">
<input type="hidden" name="SubmitType" value="">
<input type="hidden" name="SelValue" value="">
<input type="hidden" name="SelFileName" value="">
<input type="hidden" name="SelCarNo" value="">
<input type="hidden" name="SelSN" value="">
	  
	  
<table width="100%" border="0">
  <tr class="pagetitle">
    <td height="20" bgcolor="#1BF5FF" class="font12">數位固定桿處理系統 </td>
  </tr>
  <tr>
    <td height="20" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td><span class="font12">違規日期        
			<input name="StartDate" type="text" value="<%
				StartDateTmp=trim(request("StartDate"))
			response.write StartDateTmp
			%>" size="8" maxlength="7" class="btn1">
			<input type="button" name="datestr" value="..." onclick="OpenWindow('StartDate');">
			~
			<input name="EndDate" type="text" value="<%
				EndDateTmp=trim(request("EndDate"))
			response.write EndDateTmp
			%>" size="8" maxlength="7" class="btn1">
			<input type="button" name="datestr" value="..." onclick="OpenWindow('EndDate');">
     車牌號碼
		<input type="text" name="sys_CarNo" value="<%=trim(request("sys_CarNo"))%>" size="8" onkeyup="CarNoFormat();">
      路段              
<select name="Location">
  <option selected value="">選擇舉發路段...</option>
  <%=sRoadOption%>
</select>
<!-- 舉發來源        
<select name="FIXEQUIPTYPE">
  <option selected value="">選擇舉發來源...</option> -->
  <%'=sTypeOption%>
<!-- </select> -->

舉發類型
<select name="Sys_ProsecutionTypeID">
  <option value="">全部</option>
  <option value="R"<%If trim(request("Sys_ProsecutionTypeID"))="R" Then response.write " selected"%>>闖紅燈</option>
  <option value="S"<%If trim(request("Sys_ProsecutionTypeID"))="S" Then response.write " selected"%>>超速</option>
  <option value="L"<%If trim(request("Sys_ProsecutionTypeID"))="L" Then response.write " selected"%>>越線</option>
  <option value="T"<%If trim(request("Sys_ProsecutionTypeID"))="T" Then response.write " selected"%>>緊跟著前車駕駛</option>
  <option value="U"<%If trim(request("Sys_ProsecutionTypeID"))="U" Then response.write " selected"%>>左轉</option>
</select>

 有效性        
<select name="VerifyResult">
	<option value="NO" <%if trim(request("VerifyResult"))="NO" then response.write "selected"%>>選擇有效性</option>
  <option value="1" <%if trim(request("VerifyResult"))="1" or (trim(request("SubmitType"))="" and trim(request("VerifyResult"))="") then response.write "selected"%>>未處理</option>
  <option value="0" <%if trim(request("VerifyResult"))="0" then response.write "selected"%>>有效</option>
  <option value="2" <%if trim(request("VerifyResult"))="2" then response.write "selected"%>>有效未建檔</option>
  <option value="-1" <%if trim(request("VerifyResult"))="-1" then response.write "selected"%>>無效</option>
</select>


<!-- 舉發人
 -->	<!-- <select name="OperatorID" class="btn1">
		<option value="">請選擇</option> -->
<%
	'strMember="select MemberID,chName from MemberData where UnitID='"&trim(Session("Unit_ID"))&"' order by chName"
	'set rsMember=conn.execute(strMember)
	'If Not rsMember.Bof Then rsMember.MoveFirst 
	'While Not rsMember.Eof
%>
			<%
			'	formOperatorTmp=trim(formOperator)
			'if trim(rsMember("chName"))= trim(formOperatorTmp) then
			'	response.write "selected"
			'end if
			%>
<%
	'rsMember.MoveNext
	'Wend
	'rsMember.close
	'set rsMember=nothing
%>
	<!-- </select>
	車牌號碼        
	<input name="CarNo" type="text" value="" size="7" maxlength="8"> -->

      <input type="button" name="Submit" value="查詢" onclick='funcDataSearch();'>     
	          </span>  
			  <br>
		<strong>( 闖紅燈 <%
		strRed = "select count(*) as cnt " + _
			"from ProsecutionImage a, ProsecutionImageDetail b, MEMBERDATA c, SpecCar d " + _
			"where a.FILENAME = b.FILENAME and b.MEMBERID = c.MEMBERID(+) and b.CARNO = d.CARNO(+) and a.ProsecutionTypeID like '%R%' " + sWhere
		set rsRed=Conn.execute(strRed)
		if not rsRed.eof then
			response.write trim(rsRed("cnt"))
		end if
		rsRed.close
		set rsRed=nothing
		%> 件，超速 <%
		cntspeend=0
		strS = "select count(*) as cnt " + _
			"from ProsecutionImage a, ProsecutionImageDetail b, MEMBERDATA c, SpecCar d " + _
			"where a.FILENAME = b.FILENAME and b.MEMBERID = c.MEMBERID(+) and b.CARNO = d.CARNO(+) and a.ProsecutionTypeID like '%S%' " + sWhere
		set rsS=Conn.execute(strS)
		if not rsS.eof then
			response.write trim(rsS("cnt"))
			cntspeend=trim(rsS("cnt"))
		end if
		rsS.close
		set rsS=nothing
		%> 件，平均時速 <%
		SpeedAvg=0
		SpeedCount=0
		strOver = "select a.OverSpeed " + _
			"from ProsecutionImage a, ProsecutionImageDetail b, MEMBERDATA c, SpecCar d " + _
			"where a.FILENAME = b.FILENAME and b.MEMBERID = c.MEMBERID(+) and b.CARNO = d.CARNO(+) " + sWhere
		set rsOver=Conn.execute(strOver)
		If Not rsOver.Bof Then rsOver.MoveFirst 
		While Not rsOver.Eof
			if trim(rsOver("OverSpeed"))<>"" and not isnull(rsOver("OverSpeed")) then
				SpeedAvg=SpeedAvg+cdbl(rsOver("OverSpeed"))
				SpeedCount=SpeedCount+1
			end if
		rsOver.MoveNext
		Wend
		rsOver.close
		set rsOver=nothing	
		if SpeedAvg<>0 and SpeedCount<>0 then
			response.write formatNumber(SpeedAvg/SpeedCount,2)
		else
			response.write "0"
		end if
'		If sumOverSpeed<>0 and cntspeend<>0 Then
'			response.write Cint(sumOverSpeed/cntspeend)
'		else
'			response.write 0
'		end if
		%> Km )</strong>
<%
if isObject(Conn) then
    if Conn.state then
        Conn.close
    end if
    Set Conn = Nothing
end if
%>
	</td>
      </tr>
    </table></td>
  </tr>
  <tr class="listtitle">
    <td height="26" bgcolor="#1BF5FF"><span class="font12">數位相片舉發-資料列表</span></td>
  </tr>
  <tr>
    <td height="335" bgcolor="#E0E0E0"><span id="DataList"><table id="DataList1" width="100%" height="100%" border="0" cellpadding="2" cellspacing="1">
      <tr bgcolor="#EBFBE3">
        <th width="120" height="15" nowrap><span class="font12">違規日期</span></th>
        <!-- <th width="40" nowrap><span class="font12">來源</span></th> -->      
        <th width="120" nowrap><span class="font12">路段</span></th>

       <!--  <th width="80" nowrap><span class="font12">車牌</span></th> -->
        <th width="50" nowrap><span class="font12">類型</span></th>
        <th width="40" nowrap>限速</th>
        <th width="40" nowrap>車速</th>
        <!-- <th width="40" nowrap>車道</th> -->
		<th width="80" nowrap>異常原因</th>
        <!-- <th width="40" nowrap>審核</th> -->
        <th width="100" nowrap></th>
      </tr>
	<%=sTR%>
    </table></span></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#1BF5FF"><p align="center" class="style1">
    	<%
    		ShowPageLink sNowPage, sAllPage, "ProsecutionImage.asp", "&VerifyResult="&trim(request("VerifyResult"))&"&StartDate="&trim(request("StartDate"))&"&EndDate="&trim(request("EndDate"))&"&Location="&trim(request("Location"))&"&FIXEQUIPTYPE="&trim(request("FIXEQUIPTYPE"))&"&OperatorID="&trim(request("OperatorID"))&"&CarNo="&trim(request("CarNo"))
    	%>
      <img src="space.gif" width="20" height="5"> 
      <input type="button" onclick="OpenExcel()" value="轉換成Excel"> 
</p>    </td>
  </tr>
</table>
</form>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
	var OldObj = '';
	var sformLocation = '<%=formLocation%>';
	var formFIXEQUIPTYPE = '<%=formFIXEQUIPTYPE%>';
	var sformVerifyResult = '<%=formVerifyResult%>';
	var sformCarNo = '<%=formCarNo%>';
	var sMessage = '<%=sMessage%>';
	
	if(sMessage!=''){
		alert(sMessage);
	}
	
	if(sformCarNo!=''){
		document.getElementById('CarNo').value=sformCarNo;
	}
	
	if(sformLocation!=''){
		for(var i=0;i<document.getElementById('Location').length;i++){
			if(document.getElementById('Location')[i].value==sformLocation){
				document.getElementById('Location')[i].selected=true;
			}
		}
	}
	
	if(formFIXEQUIPTYPE!=''){
		for(var i=0;i<document.getElementById('FIXEQUIPTYPE').length;i++){
			if(document.getElementById('FIXEQUIPTYPE')[i].value==formFIXEQUIPTYPE){
				document.getElementById('FIXEQUIPTYPE')[i].selected=true;
			}
		}
	}
	
	if(sformVerifyResult!=''){
		for(var i=0;i<document.getElementById('VerifyResult').length;i++){
			if(document.getElementById('VerifyResult')[i].value==sformVerifyResult){
				document.getElementById('VerifyResult')[i].selected=true;
			}
		}
	}
	
	//置中 function
	function centerPos(size, type) {
	    switch(type) {
	        case 0:   //Top position
	            return (parseInt(window.screen.height) - size) / 2;
	            break;
	        case 1:   //Left position
	            return (parseInt(window.screen.width) - size) / 2;
	            break;
	        default:
	            alert('centerPos() : Type value error!!');
	    }
	}
	
	function OpenExcel(){
		window.open('ProsecutionImageExcel.asp?' +
					'Location='+escape(document.all('Location').value) +
					'&VerifyResult='+document.all('VerifyResult').value +
//					'&CarNo='+escape(document.all('CarNo').value) +
					'&StartDate='+document.all('StartDate').value +
					'&EndDate='+document.all('EndDate').value 
					//'&FIXEQUIPTYPE='+document.all('FIXEQUIPTYPE').value 
					//'&OperatorID='+document.all('OperatorID').value
					,'',',scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=yes,toolbar=no');
	}
	
	function NewWindow(Width, Height, URL, WinName){
	    var nWidth = Width;
	    var nHeight = Height;
	    var sURL = URL;
	    var nTop = centerPos(nHeight,0);
	    var nLeft = centerPos(nWidth,1);
	    var sWinSize = "left=" + nLeft.toString(10) + ",top=" + nTop.toString(10) + ",width=" + nWidth.toString(10) + ",height=" + nHeight.toString(10);
	    var sWinStatus = "menubar=0,toolbar=0,scrollbars=1,resizable=1,status=0";
	    var sWinName = WinName;
	    OldObj = window.open(sURL,sWinName,sWinSize + ",left=0,top=0," + sWinStatus);
	}

	//開啟檢視圖
	function OpenPic(FileName){
		//alert(FileName);
		NewWindow(1000, 700, 'ShowMap.asp?PicName=' + FileName.replace(/\+/g, '@2@'), 'MyPic');
	}

	//開啟檢視圖
	function OpenPic2(FileName){
		NewWindow(1000, 700, FileName, 'MyPic');
	}

	//關閉檢是圖
	function ClosePic(){
		try{
			if(OldObj!=''){
				OldObj.close();
				OldObj = '';
			}
		}
		catch(e){
		}
	}
	
	//開啟詳細資料
	function OpenDetail(FileName, SN){
		//+ URL 傳送時會不見,所以置換,到Server Side 再換回來
		NewWindow(1000, 600, 'ProsecutionImageDetail.asp?FileName=' + FileName.replace(/\+/g, '@2@') + '&SN='+SN, 'MyDetail');
	}
	
	//資料審核
	function UpdateData(CarNoObj, FileName, Value, SN){
		//抓取汽車號碼
		var CarNo = document.getElementById(CarNoObj).value;

		if(CarNo!='' || (CarNo=='' && Value=='-1')){
			var sType = '';
			if(Value=='0'){
				sType = '有效';
			}
			else{
				sType = '無效';
			}
	
			var bOK = confirm('確定要將 ' + CarNo + ' 資料審核成 \'' + sType + '\' 嗎?');
			
			if(bOK){
				document.getElementsByName('SubmitType')[0].value = 'UpdateVerifyResult';
				document.getElementsByName('SelFileName')[0].value = FileName;
				document.getElementsByName('SelValue')[0].value = Value;
				document.getElementsByName('SelCarNo')[0].value = CarNo;
				document.getElementsByName('SelSN')[0].value = SN;
				
				document.forms[0].submit();
			}
		}
		else if (CarNo=='' && Value=='0'){
			alert('請輸入車牌號碼!!');
		}
	}

	//進入舉發單修改頁面
	function OpenCaseData(BillSN){
		NewWindow(1000, 600, '../BillKeyIn/BillKeyIn_Car_Report_Update.asp?BillSN=' + BillSN, 'MyDetail');
	}
	//逕舉資料建檔
	function funcInCaseData(FileName,SN,CarNo){
		document.getElementsByName('SubmitType')[0].value = 'InsertBillBase';
		document.getElementsByName('SelFileName')[0].value = FileName;
		document.getElementsByName('SelCarNo')[0].value = CarNo;
		document.getElementsByName('SelSN')[0].value = SN;

		document.forms[0].submit();
	}
	function funcDataSearch(){
		document.getElementsByName('SubmitType')[0].value = 'DataSearch';
		document.forms[0].submit();
	}

	//
	function CarNoFormat(){
		document.all('sys_CarNo').value=document.all('sys_CarNo').value.toUpperCase();
		document.all('sys_CarNo').value=document.all('sys_CarNo').value.replace(/[\s　]+/g, "");
	}
<%if trim(request("SubmitType"))="" then%>
	//funcDataSearch();
<%end if%>
</script>
</body>
</html>