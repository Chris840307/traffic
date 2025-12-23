<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/css.txt"-->
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'on error resume next
'檢查是否可進入本系統
'AuthorityCheck(223)
'==========POST=========
'單號
if trim(request("billno"))="" then
	theBillno=""
else
	theBillno=trim(request("billno"))
end if
' 告發類別
' theBilltype=1  1 攔停  2 逕舉
if trim(request("Billtype"))="" then
	theBilltype=""
else
	theBilltype=trim(request("Billtype"))
end if
'==========cookie==========
'填單人代碼
theRecordMemberID=trim(Session("User_ID"))
gCh_Name = session("CH_Name")
gUnit_ID = Session("Unit_ID")
'=========fucntion=========
function DateFormatChange(changeDate)
	DateFormatChange=funGetDate(gOutDT(changeDate),1)
	'smith remark
	'if left(trim(changeDate),1)="0" then
	'	theFormatDate=cint(mid(trim(changeDate),2,2))+1911&"/"&mid(trim(changeDate),4,2)&"/"&mid(trim(changeDate),6,2)
	'else
	'	theFormatDate=cint(left(trim(changeDate),3))+1911&"/"&mid(trim(changeDate),4,2)&"/"&mid(trim(changeDate),6,2)
	'end if
	'DateFormatChange=theFormatDate
end function
'==========================
	'要到ApConfigure抓法條版本
	strRuleVer="select Value from ApConfigure where ID=3"
	set rsRuleVer=conn.execute(strRuleVer)
	if not rsRuleVer.eof then
		theRuleVer=trim(rsRuleVer("Value"))
	end if
	rsRuleVer.close
	set rsRuleVer=nothing
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	if Trim(request("DB_Move"))="" then
		DBcnt=0
	else
		DBcnt=request("DB_Move")
	end if

	strwhere=trim(request("SQLstr"))
	strDCISQL=trim(request("strDCISQL"))

	'刪除
	if trim(request("DB_Kind"))="DB_Delete" then
		strDel="Update BillBase set RecordStateID=-1,BillStatus='6',DelMemberID="&theRecordMemberID &_
			" where SN="&trim(request("BillSN"))
		conn.execute strDel
		
		strDelP="Update ProsecutionImageDetail set BillSn=null,VERIFYRESULTID=1 where BillSn="&trim(request("BillSN"))
		conn.execute strDelP

		if Trim(request("DB_Move"))="" or Trim(request("DB_Move"))="0" then
			DBcnt=0
		else
			DBcnt=cint(Trim(request("DB_Move")))-1
		end if
	end if

	'無效
	if trim(request("DB_Kind"))="DB_Verify" then
		strDel="Update BillBase set RecordStateID=-1,BillStatus='6',DelMemberID="&theRecordMemberID &_
			" where SN="&trim(request("BillSN"))
		conn.execute strDel
		
		strDelP="Update ProsecutionImageDetail set MEMBERID="&theRecordMemberID&",VERIFYRESULTID=-1 where BillSn="&trim(request("BillSN"))
		conn.execute strDelP

		if Trim(request("DB_Move"))="" or Trim(request("DB_Move"))="0" then
			DBcnt=0
		else
			DBcnt=cint(Trim(request("DB_Move")))-1
		end if
	end If
	
	'通過DB_Update
	if trim(request("DB_Kind"))="DB_Update" Then
		If sys_City="苗栗縣" Then
			
		Else 
			strDel="Update BillBase set BillMemID1="&Trim(Session("User_ID"))&",BillMem1='"&Trim(Session("Ch_Name"))&"',BillFillerMemberID="&Trim(Session("User_ID"))&",BillFiller='"&Trim(Session("Ch_Name"))&"'" &_
			" where SN="&trim(request("BillSN"))
			conn.execute strDel
		End if
		
		
		
		
	end If

	if trim(request("DB_Kind"))="DB_Verify" or trim(request("DB_Kind"))="DB_Delete" then
		'存下載沖洗照片的資料夾
		strDownFolder="select Value from Apconfigure where id=50"
		set rsDownFolder=conn.execute(strDownFolder)
		if not rsDownFolder.eof then
			DownFolder=trim(rsDownFolder("Value"))
		end if
		rsDownFolder.close
		set rsDownFolder=nothing
		
		'日期資料夾名稱
		TodayFolder=""
		strFileDate="select RecordDate from billbase where sn="&trim(request("BillSN"))
		set rsFileDate=conn.execute(strFileDate)
		if not rsFileDate.eof then
			TodayFolder=year(rsFileDate("RecordDate"))-1911&right("00"&month(rsFileDate("RecordDate")),2)&right("00"&day(rsFileDate("RecordDate")),2)
		end if
		rsFileDate.close
		set rsFileDate=nothing

		dim fso 
		set fso=Server.CreateObject("Scripting.FileSystemObject")
		'檔案名稱
		thePicImageFileA=""
		thePicImageFileB=""
		strFile="select * from BILLILLEGALIMAGE where BillSn="&trim(request("BillSN"))
		set rsFile=conn.execute(strFile)
		IF not rsFile.eof then
			if trim(rsFile("ImageFileNameA"))<>"" and not isnull(rsFile("ImageFileNameA")) then
				thePicImageFileA=trim(rsFile("ImageFileNameA"))
				if (fso.FileExists(DownFolder&Session("User_ID")&"\"&TodayFolder&"\"&thePicImageFileA))=true then
					fso.DeleteFile DownFolder&Session("User_ID")&"\"&TodayFolder&"\"&thePicImageFileA
				end if
			end if
			if trim(rsFile("ImageFileNameB"))<>"" and not isnull(rsFile("ImageFileNameB")) then
				thePicImageFileB=trim(rsFile("ImageFileNameB"))
				if (fso.FileExists(DownFolder&Session("User_ID")&"\"&TodayFolder&"\"&thePicImageFileB))=true then
					fso.DeleteFile DownFolder&Session("User_ID")&"\"&TodayFolder&"\"&thePicImageFileB
				end if	
			end if
		end if
		rsFile.close
		set rsFile=nothing
	end if
	'總共幾筆
	'Session.Contents.Remove("BillCnt_Image")
	'Session.Contents.Remove("BillOrder_Image")
	strSqlCnt="select count(*) as cnt from (select distinct c.SN,c.CarSimpleID,c.Rule1,c.Rule2,c.Rule3,c.Rule4,c.IllegalAddress,c.RuleSpeed,c.IllegalSpeed,c.BillStatus,c.RecordStateID,c.RecordDate,c.RecordMemberID,c.BillNo,c.BillTypeID,c.RuleVer,c.IllegalDate,c.imagefilename,c.imagefilenameb,c.Note,e.CarNo,e.DCIReturnCarType,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerAddress,e.OwnerZip,e.DCIReturnCarStatus from (select * from DCILog "&strDCISQL&") a,MemberData b,BillBase c,DCIReturnStatus d,BillBaseDCIReturn e where a.BillSN=c.SN and e.ExchangeTypeID='A' and e.Status='S' and c.CarNo=e.CarNo (+) and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and c.RecordStateID=0 "&strwhere&")"
	set rsCnt1=conn.execute(strSqlCnt)
		'Session("BillCnt_Image")=trim(rsCnt1("cnt"))
		'Session("BillOrder_Image")=trim(rsCnt1("cnt"))+1
		DBsum=CDbl(rsCnt1("cnt"))
	rsCnt1.close
	set rsCnt1=nothing

	strSQL="select distinct c.SN,c.CarSimpleID,c.Rule1,c.Rule2,c.Rule3,c.Rule4,c.IllegalAddress,c.RuleSpeed,c.IllegalSpeed,c.BillStatus,c.RecordStateID,c.RecordDate,c.RecordMemberID,c.BillNo,c.BillTypeID,c.RuleVer,c.IllegalDate,c.imagefilename,c.imagefilenameb,c.Note,e.CarNo,e.DCIReturnCarType,e.A_Name,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerAddress,e.OwnerZip,e.Nwner,e.NwnerID,e.NwnerAddress,e.NwnerZip,e.DCIReturnCarStatus from (select * from DCILog "&strDCISQL&") a,MemberData b,BillBase c,DCIReturnStatus d,BillBaseDCIReturn e where a.BillSN=c.SN and e.ExchangeTypeID='A' and e.Status='S' and c.CarNo=e.CarNo (+) and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and c.RecordStateID=0 "&strwhere&" order by c.RecordDate"

	set rsfound=conn.execute(strSQL)

	if rsfound.Bof then
%>
<script language="JavaScript">
	alert("查無車籍資料!");
	window.close();
</script>
<%		response.end
	end if
	If Not rsfound.Bof Then rsfound.move DBcnt
%>
<title>數位固定桿違規影像車籍資料</title>
<style type="text/css">
<!--
.style2 {font-size: 12px}
.style3 {
font-size: 12px ;
color: #FF0000}
.style4 {
font-size: 12px ;
}
.style5 {
color: #0000FF;
font-size: 13px ;
}
.style6 {
color: #FF0000;
font-size: 13px ;
}
.style10 {
	line-height:20px;
	font-size: 12pt
}
.style11 {
	line-height:28px;
	font-size: 18pt
}
.btn2 {font-size: 13px}
.Text1{
font-weight:bold;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">

<form name="myForm" method="post">  
<table width='1200' border='1' align="left" cellpadding="0">
	<tr>
		<td width="24%" height="250" valign="top">

		<span class="style4">共 <%=DBsum%> 筆</span>
	
		<br>
	<%
	If sys_City="高雄市" Then
	
		response.write "<span class='style11'>"&Trim(rsfound("BillNo"))&"</span>"
	End If
	%>
	<%If sys_City<>"高雄市" then%>
		<table width='239' border='1' align="left" cellpadding="3">
			<tr bgcolor="#FFFFCC">
			<td align="center" width="40%"><span class="style2">車號</span></td>
			<td align="center" width="60%"><span class="style2">違規時間</span></td>
			</tr>
		<%	
			set rs2=conn.execute(strSQL)
			If Not rs2.Bof Then rs2.move DBcnt
			for i=1 to 9 
				if rs2.eof then exit for
		%>	
			<tr <%
			if i=1 then
				response.write "bgcolor='#CCFFFF'"
			end if
			%>>
				<td><span class="style2">
		        <%
				if trim(rs2("CarNo"))<>"" and not isnull(rs2("CarNo")) then 
					response.write trim(rs2("CarNo"))
				else
					response.write "&nbsp;"
				end if				
				%>
				  </span></td>
				<td><span class="style2">
		        <%
				if trim(rs2("IllegalDate"))<>"" and not isnull(rs2("IllegalDate")) then 
					response.write ginitdt(trim(rs2("IllegalDate")))&"&nbsp; "&right("00"&hour(trim(rs2("IllegalDate"))),2)&right("00"&minute(trim(rs2("IllegalDate"))),2)
				else
					response.write "&nbsp;"
				end if				
				%>
				  </span></td>
			</tr>
		<%
				rs2.MoveNext
			next
			rs2.close
			set rs2=nothing
		%>
		</table>
	<%End if%>
		</td>
	<%
		theImageFileNameA=""
		theImageFileNameB=""
		theIISImagePath=""
		if sys_City="台東縣" Then
			If (Trim(rsfound("Rule1"))="5620001" Or Trim(rsfound("Rule1"))="5630001") And Not IsNull(rsfound("ImageFileName")) Then
				theImageFileNameA="\traffic\StopCarPicture\"&Trim(rsfound("ImageFileName"))
			Else
				strImage="select * from BillIllegalImage where BillSn="&trim(rsfound("SN"))
				set rsImage=conn.execute(strImage)
				if not rsImage.eof then
					if trim(rsImage("ImageFileNameA"))<>"" and not isnull(rsImage("ImageFileNameA")) then
						theImageFileNameA=trim(rsImage("ImageFileNameA"))
					end if
					if trim(rsImage("ImageFileNameB"))<>"" and not isnull(rsImage("ImageFileNameB")) then
						theImageFileNameB=trim(rsImage("ImageFileNameB"))
					end if
					if trim(rsImage("IISImagePath"))<>"" and not isnull(rsImage("IISImagePath")) then
						theIISImagePath=trim(rsImage("IISImagePath"))
					end if
				end if
				rsImage.close
				set rsImage=Nothing
			End If 
		Else
				strImage="select * from BillIllegalImage where BillSn="&trim(rsfound("SN"))
				set rsImage=conn.execute(strImage)
				if not rsImage.eof then
					if trim(rsImage("ImageFileNameA"))<>"" and not isnull(rsImage("ImageFileNameA")) then
						theImageFileNameA=trim(rsImage("ImageFileNameA"))
					end if
					if trim(rsImage("ImageFileNameB"))<>"" and not isnull(rsImage("ImageFileNameB")) then
						theImageFileNameB=trim(rsImage("ImageFileNameB"))
					end if
					if trim(rsImage("IISImagePath"))<>"" and not isnull(rsImage("IISImagePath")) then
						theIISImagePath=trim(rsImage("IISImagePath"))
					end if
				end if
				rsImage.close
				set rsImage=Nothing
		End If 

		bPicWebPath = ""
		if trim(theImageFileNameA)<>"" then
			bPicWebPath=theIISImagePath&theImageFileNameA
		end if
	%>
		<td rowspan="2" valign="top" height="490">
		<!-- 影像大圖 -->
		
		<%if bPicWebPath<>"" then%>
		<img src="<%=bPicWebPath%>" border=1 height="490" onmousemove="show(this, '<%=bPicWebPath%>')" onmousedown="show(this, '<%=bPicWebPath%>')" id="imgSource" src="<%=bPicWebPath%>" >

		<div id="div1" style="position:absolute; overflow:hidden; width:<%
				If sys_City="高雄市" Then
					response.write "810"
				Else
					response.write "210"
				End If 
			%>px; height:90px; left:<%
			if trim(request("divX"))="" Then
				If sys_City="台中市" Then
					response.write "1005"
				Else
					response.write "248"
				End If 
				
			else
				response.write trim(request("divX"))
			end if
			%>px; top:<%
			if trim(request("divY"))="" Then
				If sys_City="台中市" Then
					response.write "10"
				Else
					response.write "405"
				End If 
				
			else
				response.write trim(request("divY"))
			end if
			%>px; z-index:1;border-right: white thin ridge; border-top: white thin ridge; border-left: white thin ridge; border-bottom: white thin ridge" onMousedown="initializedragie( )">
				<img id="BigImg" style='position:relative' src="<%=bPicWebPath%>">
			</div>
		<%end if%>
		<%	if not rsfound.eof Then
				If sys_City="苗栗縣" then
					
					response.write "<br><span class=""style11""><strong>"
					response.write "&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; "
					response.write trim(rsfound("CarNo"))&"&nbsp; &nbsp; &nbsp; &nbsp; "&Year(rsfound("IllegalDate"))-1911&Right("00"&Month(rsfound("IllegalDate")),2)&Right("00"&Day(rsfound("IllegalDate")),2)& " " &Right("00"&Hour(rsfound("IllegalDate")),2)&Right("00"&Minute(rsfound("IllegalDate")),2) 
					response.write "</strong></span>"
				end if 
			End if
				%>
		</td>
	</tr>
	<tr>
		<td height="120" align="center">
		<!-- 影像小圖 -->
		<%
		sPicWebPath=""
		if trim(theImageFileNameB)<>"" then
			sPicWebPath=theIISImagePath&theImageFileNameB
		elseif bPicWebPath<>"" then
			sPicWebPath=bPicWebPath
		end if
		%>
		<%if sPicWebPath<>"" then%>
		<img src="<%=sPicWebPath%>" border=1 width="210" id="SmallImg" ondblclick="ChangeImg()">
		<%end if%>
		<br>
		<%if bPicWebPath<>"" then%>
			<input type="button" onClick="OpenPic('<%=replace(bPicWebPath,"\","/")%>')" value="大圖一" class="style4">
		<%end if%>
		<%if trim(theImageFileNameB)<>"" then%>
			<input type="button" onClick="OpenPic('<%=replace(sPicWebPath,"\","/")%>')" value="大圖二" class="style4">
		<%end if%>
		<%
	If sys_City="XXX市" Then
		strPro="select * from ProsecutionImage a,ProsecutionImageDetail b where a.FileName=b.FileName and b.BillSN="&trim(rsfound("SN"))
		set rsPro=conn.execute(strPro)
		if not rsPro.eof then
			if trim(rsPro("VideoFileName"))<>"" and not isnull(rsPro("VideoFileName")) then
			VideoFilePath=theIISImagePath & rsPro("VideoFileName")
		%>
			<input type="button" onClick="OpenPic2('<%=VideoFilePath%>')" value="錄影" class="style4">
		<%	end if%>

		
			<input type="button" onClick="OpenDetail('<%=trim(rsPro("FileName"))%>','<%=trim(rsPro("SN"))%>')" value="詳細" class="style4">
			<input type="hidden" name="SelFileName" value="<%=trim(rsPro("FileName"))%>">
			<input type="hidden" name="SelSN" value="<%=trim(rsPro("SN"))%>">
		<%
		end if
		rsPro.close
		set rsPro=Nothing
	End if
		%>
		</td>
	</tr>

	<tr>
		<td height="50" colspan="2" valign="top">
		<%if not rsfound.eof then%>
		<table width='100%' border='1' align="left" cellpadding="0">
			<tr height="25">
				<td bgcolor="#FFFFCC" width="8%"><span class="style10">車號</span></td>
				<td width="17%"><span class="style11"><strong>
				<%=trim(rsfound("CarNo"))%></strong></span>
				</td>
				<td bgcolor="#FFFFCC" width="8%"><span class="style10">詳細車種</td>
				<td width="17%"><span class="style11"><strong>
				<%
				if trim(rsfound("DCIReturnCarType"))<>"" and not isnull(rsfound("DCIReturnCarType")) then
					strCType="select * from DCIcode where TypeID=5 and ID='"&trim(rsfound("DCIReturnCarType"))&"'"
					set rsCType=conn.execute(strCType)
					if not rsCType.eof then
						response.write trim(rsCType("Content"))
					end if
					rsCType.close
					set rsCType=nothing
				end if								
				%></strong></span>
				</td>
				<td bgcolor="#FFFFCC" width="8%"><span class="style10">廠牌</span></td>
				<td width="17%"><span class="style11"><strong>
				<%
				if (trim(rsfound("A_Name"))<>"" and not isnull(rsfound("A_Name"))) then
					response.write trim(rsfound("A_Name"))
				else
					response.write "&nbsp;"
				end if
				%></strong></span>
				</td>
				<td bgcolor="#FFFFCC" width="8%"><span class="style10">顏色</span></td>
				<td width="17%"><span class="style11"><strong>
				<%
				if trim(rsfound("DCIReturnCarColor"))<>"" and not isnull(rsfound("DCIReturnCarColor")) then
					ColorLen=cint(Len(rsfound("DCIReturnCarColor")))
					for Clen=1 to ColorLen
						colorID=mid(rsfound("DCIReturnCarColor"),Clen,1)
						strColor="select * from DCIcode where TypeID=4 and ID='"&trim(colorID)&"'"
						set rsColor=conn.execute(strColor)
						if not rsColor.eof then
							response.write trim(rsColor("Content"))
						end if
						rsColor.close
						set rsColor=nothing
					next
				else
					response.write "&nbsp;"
				end if
				%></strong></span>
				</td>
			</tr>
			<tr height="25">
		<%If sys_City="高雄市" Then%>
				<td bgcolor="#FFFFCC" ><span class="style10">違規時間</span></td>
				<td ><span class="style10"><%
				if trim(rsfound("IllegalDate"))<>"" then
					response.write Year(rsfound("IllegalDate"))-1911 & " / " & Right("00" & month(rsfound("IllegalDate")),2 ) & " / " & Right("00" & day(rsfound("IllegalDate")),2 ) & "&nbsp; " & Right("00" & Hour(rsfound("IllegalDate")),2 ) & " : " & Right("00" & minute(rsfound("IllegalDate")),2 )
				end if
				%></span>
				</td>
		<%End if%>
				<td bgcolor="#FFFFCC" ><span class="style10">車主姓名</span></td>
				<td colspan="<%
			if sys_City="花蓮縣" Then
				response.write "2"
			elseif sys_City="高雄市" Then
				response.write "0"
			Else
				response.write "3"
			End If 
				%>"><span class="style10"><%
				if trim(rsfound("Owner"))<>"" then
					response.write funcCheckFont(rsfound("Owner"),20,1)
				end if
				%></span>
				</td>
		
				<td>
		<%If sys_City<>"台中市" Then%>
				<%
			If Trim(rsfound("DCIReturnCarStatus"))<>"" Then
				response.write Trim(rsfound("DCIReturnCarStatus"))&"_"
				strDCode="select * from dcicode where typeid=10 and id='"&Trim(rsfound("DCIReturnCarStatus"))&"'"
				Set rsDCode=conn.execute(strDCode)
				If Not rsDCode.eof Then
					response.write Trim(rsDCode("Content"))
				End If
				rsDCode.close
				Set rsDCode=Nothing 
			Else
				response.write "&nbsp;"
			End If 
		end if
				%>
				</td>
				<td bgcolor="#FFFFCC" ><span class="style10">車主地址</span></td>
				<td colspan="3"><span class="style10"><%
				if (trim(rsfound("OwnerAddress"))<>"" and not isnull(rsfound("OwnerAddress"))) then
					response.write trim(rsfound("OwnerZip"))&funcCheckFont(rsfound("OwnerAddress"),20,1)
				else
					response.write "&nbsp;"
				end if
				%></span></td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" ><span class="style10">違規地點</span></td>
				<td ><span class="style10"><%
				if trim(rsfound("illegaladdress"))<>"" then
					response.write funcCheckFont(rsfound("illegaladdress"),20,1)
				else
					response.write "&nbsp;"
				end if
				%></span>
				</td>
				<td bgcolor="#FFFFCC" ><span class="style10">違規事實一</span></td>
				<td colspan="2"><span class="style10"><%
				if trim(rsfound("Rule1"))<>"" Then
					strR1="select * from Law where ItemID='"&trim(rsfound("Rule1"))&"' and Version='"&trim(rsfound("RuleVer"))&"'"
					Set rsR1=conn.execute(strR1)
					If Not rsR1.eof then
						response.write trim(rsR1("IllegalRule"))
					End If
					rsR1.close
					Set rsR1=Nothing
					if trim(rsfound("Rule4") & " ")<>"" then
						response.write "(" & trim(rsfound("Rule4") & " ") & ")"
					end if 
					if sys_City="南投縣" Or sys_City="花蓮縣" Or sys_City="嘉義縣" Or sys_City="台中市" Or sys_City="高雄市" Or sys_City="苗栗縣" Or Trim(request("QueryType"))="ML" Then
						If Left(trim(rsfound("Rule1")),2)="40" Or Left(trim(rsfound("Rule1")),5)="33101" Or Left(trim(rsfound("Rule1")),5)="43102" then
							response.write "<br>限速 <strong>"&trim(rsfound("RuleSpeed"))&"</strong> 公里,"
							If sys_City="台中市" Or sys_City="高雄市" then
								response.write "實際車速"
							Else
								response.write "實速"
							End If 
							response.write " <strong>"&trim(rsfound("illegalSpeed"))&"</strong> 公里" 
							response.write ",超速 <strong>"&(cdbl(rsfound("illegalSpeed"))-cdbl(rsfound("RuleSpeed")))&"</strong> 公里"
						End If 
					End if
				else
					response.write "&nbsp;"
				end if
				%></span>
				</td>
				<td bgcolor="#FFFFCC" ><span class="style10">違規事實二</span></td>
				<td colspan="2"><span class="style10"><%
				if trim(rsfound("Rule2"))<>"" Then
					strR2="select * from Law where ItemID='"&trim(rsfound("Rule2"))&"' and Version='"&trim(rsfound("RuleVer"))&"'"
					Set rsR2=conn.execute(strR2)
					If Not rsR2.eof then
						response.write trim(rsR2("IllegalRule"))
					End If
					rsR2.close
					Set rsR2=nothing
				else
					response.write "&nbsp;"
				end if
				%></span>
				</td>
			</tr>
		<%
			DeleteOk_flag=0
			VerifyOk_flag=0

			If sys_City="台中市" Then
		%>
			<tr>
				<td bgcolor="#FFFFCC" ><span class="style10">舉發員警</span></td>
				<td colspan="4"><span class="style11"><strong>
				<%
				BillMemID_Temp=""
				strBill="select billMemId1,(select chName from memberdata where memberID=billbase.BillMemID1) as BillchName1 from billbase where sn="&trim(rsfound("Sn"))
				Set rsBill=conn.execute(strBill)
				If Not rsBill.eof Then
					BillMemID_Temp=Trim(rsBill("billMemId1"))
					response.write Trim(rsBill("BillchName1"))
				End If
				rsBill.close
				Set rsBill=Nothing 
				

				%></strong></span>
				</td>
				<td bgcolor="#FFFFCC" ><span class="style10">牌照狀態</span></td>
				<td colspan="2">
				<%
			If Trim(rsfound("DCIReturnCarStatus"))<>"" Then
				response.write Trim(rsfound("DCIReturnCarStatus"))&"_"
				strDCode="select * from dcicode where typeid=10 and id='"&Trim(rsfound("DCIReturnCarStatus"))&"'"
				Set rsDCode=conn.execute(strDCode)
				If Not rsDCode.eof Then
					response.write Trim(rsDCode("Content"))
				End If
				rsDCode.close
				Set rsDCode=Nothing 
			Else
				response.write "&nbsp;"
			End If 
				%>
				</td>
			</tr>
		<%
			End If 
		%>
		</table>
		<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
		<%end if%>
		</td>
	</tr>
	<tr bgcolor="#FFCC33">
		<td height="28" colspan="2" align="center">
				<%
				if DBcnt=0 then
					BackDisabled="disabled"
				else
					BackDisabled=""
				end if
				
				if DBcnt+1=DBsum then
					NextDisabled="disabled"
				else
					NextDisabled=""
				end if
				%>
				<%
			If sys_City="台中市" Or Trim(request("QueryType"))="ML" Then
					
				%>
				<input type="button" name="Submit2932" onClick="funUpdate();" value="通 過 F2" style="font-size: 10pt; width: 70px; height: 27px" <%
					
					if trim(rsfound("BillStatus"))="1" Then
						'If BillMemID_Temp=Trim(session("User_ID")) then
							VerifyOk_flag=1
						'End If 
					End If
					If VerifyOk_flag=0 Then
						response.write " disabled"
					End If 
				%>>

				<input type="button" name="Submit2912" onClick="funVerifyResult_TC();" value="無 效 F6" style="font-size: 10pt; width: 70px; height: 27px" <%
					
					if trim(rsfound("BillStatus"))="1" Then
						'If BillMemID_Temp=Trim(session("User_ID")) then
							VerifyOk_flag=1
						'End If 
					End If
					If VerifyOk_flag=0 Then
						response.write " disabled"
					End If 
				%>>

				<input type="button" name="Submit2232"  onClick="funReturnList();" value="無效清冊" style="font-size: 10pt; width: 70px; height: 27px" >
				&nbsp; &nbsp; 
				<%
					
			End If 

				%>

				<input type="button" name="SubmitGo" onClick="funMoveGoCnt();" value="跳至" >
				第 <input type="text" name="MoveGoCnt" value="<%=Trim(request("MoveGoCnt"))%>" size="5" onkeyup="value=value.replace(/[^\d]/g,'')"> 筆 &nbsp; &nbsp; &nbsp; &nbsp; 
				<input type="button" name="SubmitBack2" onClick="funDbMove('First');" value="<< 第一筆 Home" style="font-size: 9pt; width: 100px; height: 27px" <%=BackDisabled%>>
				
				<input type="button" name="SubmitBack" onClick="funDbMove('Back');" value="< 上一筆 PgUp" style="font-size: 9pt; width: 100px; height: 27px" <%=BackDisabled%>>
				<%
				response.write DBcnt+1 &" / "&DBsum
				%>
				<input type="button" name="SubmitNext" onClick="funDbMove('Next');" value="下一筆 PgDn >" style="font-size: 9pt; width: 100px; height: 27px" <%=NextDisabled%>>

				<input type="button" name="SubmitBack" onClick="funDbMove('Last');" value="最後一筆 End >>" style="font-size: 9pt; width: 100px; height: 27px" <%=NextDisabled%>>
			<%
		If sys_City="高雄市" Then
			%>

				<input type="button" name="Submit29g2" onClick="funVerifyResultKS();" value="註 銷 F4" style="font-size: 10pt; width: 70px; height: 27px" <%
					
					if trim(rsfound("BillStatus"))="1" Or trim(rsfound("BillStatus"))="2" Then
						If Trim(rsfound("RecordMemberID"))=Trim(session("User_ID")) then
							VerifyOk_flag=1
							DeleteOk_flag=1
						End If 
					End If
					If VerifyOk_flag=0 Then
						response.write " disabled"
					End If 
				%>>
			<%
		else
			if trim(rsfound("BillStatus"))="1" and sys_City<>"苗栗縣" then
				if (checkIsAllowDel(sys_City,trim(rsfound("BillTypeID")))=true) or (trim(rsfound("imagefilenameb"))<>"")  or ( (Instr(rsfound("Rule1"),"56")>0) and (Instr(rsfound("Note"),"txt")>0) and (sys_City="花蓮縣") ) Then
				DeleteOk_flag=1
				VerifyOk_flag=1
			%>
				<input type="button" name="Submit2932" onClick="funVerifyResult();" value="刪 除 F4" style="font-size: 10pt; width: 70px; height: 27px">
				<input type="button" name="Submit2f32" onClick="funVerifyResult2();" value="照片無效 F6" style="font-size: 10pt; width: 80px; height: 27px">
			<%
				end if
			end If
		End if
			%>
				
				 <!-- <input type="button" name="Submit4234222" value="車籍清冊" onclick="funchgCarDataList()" style="font-size: 10pt; width: 70px; height: 27px"> -->
                 <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉 F8" style="font-size: 10pt; width: 70px; height: 27px">
				  &nbsp; &nbsp; &nbsp; &nbsp; 
				<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
				<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
				<input type="Hidden" name="DB_Kind" value="">
				<input type="Hidden" name="BillSN" value="<%=trim(rsfound("SN"))%>">
				<!-- 逕舉類別 -->
				<input type="hidden" size="3" maxlength="1" value="2" name="BillType" readonly>
				<!-- 應到案日期 -->
				<input type="hidden" size="12" maxlength="7" name="DealLineDate">
				<!-- 應到案處所 -->
				<input type="hidden" size="10" value="" name="MemberStation">
				<!-- <input type="button" value="？" name="StationSelect" onclick='window.open("Query_Station.asp","WebPage1","left=0,top=0,location=0,width=660,height=375,resizable=yes,scrollbars=yes")'> -->
				<div id="Layer5" style="position:absolute ; width:221px; height:24px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000; visibility :hidden;"></div>
				<input type="hidden" name="SessionFlag" value="1">
				<!--浮動視窗座標-->
				<input type="hidden" name="divX" value="<%
				if trim(request("divX"))="" Then
					If sys_City="台中市" Then
						response.write "1005"
					Else
						response.write "248"
					End If 
				else
					response.write trim(request("divX"))
				end if
				%>">
				<input type="hidden" name="divY" value="<%
				if trim(request("divY"))="" Then
					If sys_City="台中市" Then
						response.write "10"
					Else
						response.write "405"
					End If 
				else
					response.write trim(request("divY"))
				end if
				%>">
		</td>
	</tr>
</table>
</form>
</body>
<script language="JavaScript">
function funDbMove(MoveCnt){
	if (MoveCnt=="First"){
		myForm.DB_Move.value="0";
		myForm.submit();
	}else if (MoveCnt=="Back"){
		//if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10-eval(myForm.sys_MoveCnt.value)){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)-1;
			myForm.submit();
		//}
	}else if(MoveCnt=="Next"){
		//if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+1;
			myForm.submit();
		//}
	}else if(MoveCnt=="Last"){

		myForm.DB_Move.value=eval(myForm.DB_Cnt.value)-1;
		myForm.submit();
	}
}

function NewWindow(Width, Height, URL, WinName){
	var nWidth = Width;
	var nHeight = Height;
	var sURL = URL;
	var nTop = 0;
	var nLeft = 0;
	var sWinSize = "left=" + nLeft + ",top=" + nTop + ",width=" + nWidth.toString(10) + ",height=" + nHeight.toString(10);
	var sWinStatus = "menubar=0,toolbar=0,scrollbars=1,resizable=1,status=0";
	var sWinName = WinName;
	OldObj = window.open(sURL,sWinName,sWinSize + "," + sWinStatus);
}

//開啟檢視圖
function OpenPic(FileName){
//alert(FileName);
	NewWindow(1000, 700, '../ProsecutionImage/ShowMap.asp?PicName=' + FileName.replace(/\+/g, '@2@'), 'MyPic');
}
//開啟詳細資料
function OpenDetail(FileName, SN){
	//+ URL 傳送時會不見,所以置換,到Server Side 再換回來
	NewWindow(1000, 600, '../ProsecutionImage/ProsecutionImageDetail.asp?FileName=' + FileName.replace(/\+/g, '@2@') + '&SN='+SN, 'MyDetail');
}
//開啟檢視圖
function OpenPic2(FileName){
	NewWindow(1000, 700, FileName, 'MyPic');
}
//略過
function funIgnore(){
	myForm.kinds.value="BillIgnore";
	myForm.submit();
}

//=====放大鏡=======================================
var iDivHeight = 90; //放大?示?域?度
var iDivWidth = <%
	If sys_City="高雄市" Then
		response.write "810"
	Else
		response.write "210"
	End If 
%>;//放大?示?域高度
var iMultiple = 3; //放大倍?

//?示放大?，鼠?移?事件和鼠???事件都??用本事件
//??：src代表?略?，sFileName放大?片名?
//原理：依据鼠????略?左上角（0，0）上的位置控制放大?左上角???示?域左上角（0，0）的位置
function show(src, sFileName)
{
//判?鼠?事件?生?是否同?按下了
if ((event.button == 1) && (event.ctrlKey == true))
  iMultiple -= 1;
else
  if (event.button == 1)
  iMultiple += 1;
if (iMultiple < 2) iMultiple = 2;

if (iMultiple > 14) iMultiple = 14;
  
var iPosX, iPosY; //放大????示?域左上角的坐?
var iMouseX = event.offsetX; //鼠????略?左上角的?坐?
var iMouseY = event.offsetY; //鼠????略?左上角的?坐?
var iBigImgWidth = src.clientWidth * iMultiple;  //放大??度，是?略?的?度乘以放大倍?
var iBigImgHeight = src.clientHeight * iMultiple; //放大?高度，是?略?的高度乘以放大倍?

if (iBigImgWidth <= iDivWidth)
{
  iPosX = (iDivWidth - iBigImgWidth) / 2;
}
else
{
  if ((iMouseX * iMultiple) <= (iDivWidth / 2))
  {
  iPosX = 0;
  }
  else
  {
  if (((src.clientWidth - iMouseX) * iMultiple) <= (iDivWidth / 2))
  {
    iPosX = -(iBigImgWidth - iDivWidth);
  }
  else
  {
    iPosX = -(iMouseX * iMultiple - iDivWidth / 2);
  }
  }
}

if (iBigImgHeight <= iDivHeight)
{
  iPosY = (iDivHeight - iBigImgHeight) / 2;
}
else
{
  if ((iMouseY * iMultiple) <= (iDivHeight / 2))
  {
	iPosY = 0;
  }
  else
  {
	  if (((src.clientHeight - iMouseY) * iMultiple) <= (iDivHeight / 2))
	  {
		iPosY = -(iBigImgHeight - iDivHeight);
	  }
	  else
	  {
		iPosY = -(iMouseY * iMultiple - iDivHeight / 2);
	  }
  }
}
div1.style.height = iDivHeight;
div1.style.width = iDivWidth;

myForm.BigImg.width = iBigImgWidth;
myForm.BigImg.height = iBigImgHeight;
myForm.BigImg.style.top = iPosY;
myForm.BigImg.style.left = iPosX;
}
//============================================================

function ChangeImg(){
	oBigImg=myForm.imgSource.src;
	oSmallImg=myForm.SmallImg.src;

	myForm.SmallImg.src=oBigImg;
	myForm.imgSource.src=oSmallImg;
	myForm.BigImg.src=oSmallImg;
}
//-------------浮動視窗------------------
var dragswitch=0 ;
var nsx ;

function drag_dropns(name){ 
temp=eval(name) 
temp.captureEvents(Event.MOUSEDOWN | Event.MOUSEUP) 
temp.onmousedown=gons 
temp.onmousemove=dragns 
temp.onmouseup=stopns 
} 

function gons(e){ 
temp.captureEvents(Event.MOUSEMOVE) 
nsx=e.x 
nsy=e.y 
} 
function dragns(e){ 
if (dragswitch==1){ 
temp.moveBy(e.x-nsx,e.y-nsy) 
return false 
} 
} 

function stopns(){ 
temp.releaseEvents(Event.MOUSEMOVE) 
}

var dragapproved=false ;

function drag_dropie(){ 
if (dragapproved==true){ 
myForm.divX.value=tempx+event.clientX-iex
myForm.divY.value=tempy+event.clientY-iey 
document.all.div1.style.pixelLeft=tempx+event.clientX-iex 
document.all.div1.style.pixelTop=tempy+event.clientY-iey 
return false 
} 
} 

function initializedragie(){ 
iex=event.clientX 
iey=event.clientY 
tempx=div1.style.pixelLeft 
tempy=div1.style.pixelTop 
dragapproved=true 
document.onmousemove=drag_dropie 
} 

if (document.all){ 
document.onmouseup=new Function("dragapproved=false") 
} 
//------------------------------
//刪除
function funVerifyResult(){
	if(confirm('確定要刪除此筆舉發單？')){
		myForm.DB_Kind.value="DB_Delete";
		myForm.submit();
	}
	
}
//無效
function funVerifyResult2(){
	if(confirm('確定要將此筆舉發單設為無效？')){
		myForm.DB_Kind.value="DB_Verify";
		myForm.submit();
	}
} 

//註銷 高雄
function funVerifyResultKS(){
	window.open("BillBase_Del_DCI.asp?DBillSN=<%=trim(rsfound("SN"))%>","WebPage_Del_Bill","left=300,top=200,location=0,width=400,height=200,resizable=yes,scrollbars=yes")

}

//車籍清冊
function funchgCarDataList(){
	UrlStr="DciPrintCarDataList.asp?SQLstr=<%=strwhere%>&strDCISQL=<%=strDCISQL%>";
	window.open(UrlStr,"WebPage_cardataimg","left=30,top=30,location=0,width=1000,height=650,resizable=yes,scrollbars=yes,menubar=no")

}

function funUpdate(){
<%if NextDisabled="" then%>
	myForm.DB_Move.value=eval(myForm.DB_Move.value)+1;
<%end if%>
	myForm.DB_Kind.value="DB_Update";
	myForm.submit();
}

function funVerifyResult_TC(){
<%if BackDisabled="" then%>
	var Back_flag="1";
<%else%>
	var Back_flag="0";
<%end if%>
	UrlStr="DciCarDataList_Delete.asp?BillSN=<%=trim(rsfound("SN"))%>&Back_flag="+Back_flag;
	window.open(UrlStr,"WebPage_cardataimg","left=250,top=100,location=0,width=520,height=350,resizable=yes,scrollbars=yes,menubar=no")
}

function funReturnList(){
	UrlStr="DciCarDataList_Return.asp?strDCISQL=<%=strDCISQL%>";
	window.open(UrlStr,"WebPage_cardataimg","left=50,top=10,location=0,width=920,height=650,resizable=yes,scrollbars=yes,menubar=yes")
}

function funMoveGoCnt(){
	var MaxCnt=<%=DBsum%>;

	if (myForm.MoveGoCnt.value=="")
	{
		alert("請輸入 1 ~ " + MaxCnt + " 的數字!");
	}
	else if (myForm.MoveGoCnt.value > MaxCnt)
	{
		alert("請勿輸入大於 " + MaxCnt + " 的數字!");
	}else if (myForm.MoveGoCnt.value == "0"){
		alert("請輸入大於 0 的數字!");
	}else{
		myForm.DB_Move.value=eval(myForm.MoveGoCnt.value)-1;
		myForm.submit();
	}

}

function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}else if (event.keyCode==115){ //F4刪除
		event.keyCode=0;   
		event.returnValue=false;  
	<%if DeleteOk_flag=1 then%>
		funVerifyResult();
	<%end if%>
		//myForm.DB_Kind.value="DB_Delete";
		//myForm.submit();
	}else if (event.keyCode==117){ //F6無效
		event.keyCode=0;   
		event.returnValue=false;  
	<%if VerifyOk_flag=1 then%>
		funVerifyResult2();
		//myForm.DB_Kind.value="DB_Verify";
		//myForm.submit();
	<%end if%>
	}else if (event.keyCode==119){ //F8關閉
		event.keyCode=0;   
		event.returnValue=false;  
		window.close();
	}else if (event.keyCode==33){ //上一筆PageUp
		event.keyCode=0;   
		event.returnValue=false; 
		if (myForm.DB_Move.value!="0"){
			funDbMove('Back');
		}
	}else if (event.keyCode==34){ //下一筆PageDn
		event.keyCode=0;   
		event.returnValue=false; 
		if (myForm.DB_Move.value!=eval(myForm.DB_Cnt.value)-1){
			funDbMove('Next');
		}
	}else if (event.keyCode==36){ //第一筆Home
		event.keyCode=0; 
		event.returnValue=false; 
		if (myForm.DB_Move.value!="0"){
			funDbMove('First');
		}
	}else if (event.keyCode==35){ //最後一筆End
		event.keyCode=0;   
		event.returnValue=false; 
		if (myForm.DB_Move.value!=eval(myForm.DB_Cnt.value)-1){
			funDbMove('Last');
		}
	}
<%If sys_City="台中市" or trim(request("QueryType"))="ML" Then%>
	else if (event.keyCode==113){ //F2通過
		event.keyCode=0;   
		event.returnValue=false;  
	<%if VerifyOk_flag=1 then%>
		funUpdate();
	<%end if%>
	}
<%end if%>
}
<%
rsfound.close
set rsfound=nothing
conn.close
set conn=nothing
%>
</script>
</html>
