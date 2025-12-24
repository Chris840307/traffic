<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html onkeydown="KeyDown()">

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

end function

	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
'==========================
'是否要放大鏡功能(Y/N)
if sys_City="台東縣" then
	isBig="N" 
else
	isBig="Y" 
end if

'============================================
	'要到ApConfigure抓法條版本
	strRuleVer="select Value from ApConfigure where ID=3"
	set rsRuleVer=conn.execute(strRuleVer)
	if not rsRuleVer.eof then
		theRuleVer=trim(rsRuleVer("Value"))
	end if
	rsRuleVer.close
	set rsRuleVer=nothing

'新增告發單
if trim(request("kinds"))="DB_insert" Then
	If Trim(request("RuleSpeed"))<>"" And Trim(request("IllegalSpeed"))<>"" Then
		If Trim(request("RuleSpeed"))>300 Or Trim(request("IllegalSpeed"))>300 Then
			chkIsSpeedTooOver=1
		Else
			chkIsSpeedTooOver=0
		End If 
	Else
		chkIsSpeedTooOver=0
	End If 

	chkIllegalDateAndCar_KS=0
	chkAlertString=""
	If sys_City="高雄市" Then
		illegalDate1=gOutDT(request("IllegalDate"))&" 0:0:0"
		illegalDate2=gOutDT(request("IllegalDate"))&" 23:59:59"
		strIllDate=" and IllegalDate between TO_DATE('"&illegalDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&illegalDate2&"','YYYY/MM/DD/HH24/MI/SS')"
		strChk="select (select UnitName from UnitInfo where UnitID=BillUnitID) as UnitName,BillNo,Rule1,Rule2 " &_
			" from Billbase where CarNo='"&UCase(trim(request("CarNo")))&"' and RecordStateID=0 " &_
			" " & strIllDate
		Set rsChk=conn.execute(strChk)
		If Not rsChk.eof Then
			chkIllegalDateAndCar_KS=1
			chkAlertString="此車號在此違規日有違規紀錄，舉發單位:"&Trim(rsChk("UnitName"))&"，單號:"&Trim(rsChk("BillNo"))&"，法條:"&Trim(rsChk("Rule1"))
			If Trim(rsChk("Rule2"))<>"" Then
				chkAlertString=chkAlertString & "/" & Trim(rsChk("Rule2"))
			End If 
		End If 
		rsChk.close
		Set rsChk=Nothing 
	End If
	
If chkIsSpeedTooOver=0 then

	strSqlA="select * from BillBaseTemp2 where Sn=" & Trim(request("CheckSn"))
	set rsA=conn.execute(strSqlA)
	If Not rsA.eof then
	ReportSn=trim(rsA("ReportSn"))
	'SN抓最大值
	sSQL = "select BillBase_seq.nextval as SN from Dual"
	set oRST = Conn.execute(sSQL)
	if not oRST.EOF then
		sMaxSN = oRST("SN")
	end if
	oRST.close
	set oRST = Nothing
	
	theCarSimpleID="null"
	If trim(rsA("CarSimpleID"))<>"" Then
		theCarSimpleID=trim(rsA("CarSimpleID"))
	End If 
	theCarAddID="null"
	If trim(rsA("CarAddID"))<>"" Then
		theCarAddID=trim(rsA("CarAddID"))
	End If 
	theIllegalDate="null"
	If trim(rsA("IllegalDate"))<>"" Then
		theIllegalDate="to_date('"&Year(rsA("IllegalDate"))&"/"&month(rsA("IllegalDate"))&"/"&day(rsA("IllegalDate"))&" "&Hour(rsA("IllegalDate"))&":"&Minute(rsA("IllegalDate"))&":"&Second(rsA("IllegalDate"))&"','YYYY/MM/DD/HH24/MI/SS')"
	End If 
	theIllegalSpeed="null"
	If Trim(rsA("IllegalSpeed"))<>"" Then
		theIllegalSpeed=trim(rsA("IllegalSpeed"))
	End If 
	theRuleSpeed="null"
	If Trim(rsA("RuleSpeed"))<>"" Then
		theRuleSpeed=trim(rsA("RuleSpeed"))
	End If 
	theForFeit1="null"
	If Trim(rsA("ForFeit1"))<>"" Then
		theForFeit1=trim(rsA("ForFeit1"))
	End If 
	theForFeit2="null"
	If Trim(rsA("ForFeit2"))<>"" Then
		theForFeit2=trim(rsA("ForFeit2"))
	End If 
	theForFeit3="null"
	If Trim(rsA("ForFeit3"))<>"" Then
		theForFeit3=trim(rsA("ForFeit3"))
	End If 
	theForFeit4="null"
	If Trim(rsA("ForFeit4"))<>"" Then
		theForFeit4=trim(rsA("ForFeit4"))
	End If 
	theInsurance="null"
	If Trim(rsA("Insurance"))<>"" Then
		theInsurance=trim(rsA("Insurance"))
	End If
	theUseTool="null"
	If Trim(rsA("UseTool"))<>"" Then
		theUseTool=trim(rsA("UseTool"))
	End If
	theBillFillDate="null"
	If trim(rsA("BillFillDate"))<>"" Then
		theBillFillDate="to_date('"&Year(rsA("BillFillDate"))&"/"&month(rsA("BillFillDate"))&"/"&day(rsA("BillFillDate"))&" "&Hour(rsA("BillFillDate"))&":"&Minute(rsA("BillFillDate"))&":"&Second(rsA("BillFillDate"))&"','YYYY/MM/DD/HH24/MI/SS')"
	End If 
	theDealLineDate="null"
	If trim(rsA("DealLineDate"))<>"" Then
		theDealLineDate="to_date('"&Year(rsA("DealLineDate"))&"/"&month(rsA("DealLineDate"))&"/"&day(rsA("DealLineDate"))&" "&Hour(rsA("DealLineDate"))&":"&Minute(rsA("DealLineDate"))&":"&Second(rsA("DealLineDate"))&"','YYYY/MM/DD/HH24/MI/SS')"
	End If 
	theRecordDate="null"
	If trim(rsA("RecordDate"))<>"" Then
		theRecordDate="to_date('"&Year(rsA("RecordDate"))&"/"&month(rsA("RecordDate"))&"/"&day(rsA("RecordDate"))&" "&Hour(rsA("RecordDate"))&":"&Minute(rsA("RecordDate"))&":"&Second(rsA("RecordDate"))&"','YYYY/MM/DD/HH24/MI/SS')"
	End If 
	theJurgeDay="null"
	If trim(rsA("JurgeDay"))<>"" Then
		theJurgeDay="to_date('"&Year(rsA("JurgeDay"))&"/"&month(rsA("JurgeDay"))&"/"&day(rsA("JurgeDay"))&" "&Hour(rsA("JurgeDay"))&":"&Minute(rsA("JurgeDay"))&":"&Second(rsA("JurgeDay"))&"','YYYY/MM/DD/HH24/MI/SS')"
	End If 
	'BillBase
	'If sys_City="高雄市" Then
		ColAdd=",IllegalZip"
		valueAdd=",'"&trim(rsA("IllegalZip"))&"'"
	'End if	
	strInsert="insert into BillBase(SN,BillTypeID,BillNo,CarNo,CarSimpleID,CarAddID,IllegalDate" & _
				",IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed,ForFeit1" &_
				",Rule2,ForFeit2,Rule3,ForFeit3,Rule4,ForFeit4,Insurance,UseTool,ProjectID" &_
				",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
				",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
				",BillMemID4,BillMem4,BillMemID2,BillMem2,BillMemID3,BillMem3" &_
				",BillFillerMemberID,BillFiller" &_
				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
				",Note,EquipmentID,RuleVer,DriverSex,ImageFileName"&ColAdd&",JurgeDay" &_
				")" &_
				" values("&sMaxSN&",'"&trim(rsA("BillTypeId"))&"','"&UCase(trim(rsA("Billno")))&"'" &_
				",'"&UCase(trim(rsA("CarNo")))&"',"&theCarSimpleID &_						          
				","&theCarAddID&","&theIllegalDate&",'"&trim(rsA("IllegalAddressID"))&"'" &_
				",'"&trim(rsA("IllegalAddress"))&"','"&trim(rsA("Rule1"))&"',"&theIllegalSpeed &_
				","&theRuleSpeed&","&theForFeit1&",'"&trim(rsA("Rule2"))&"'" &_
				","&theForFeit2&",'"&trim(rsA("Rule3"))&"',"&theForFeit3&",'"&trim(rsA("Rule4"))&"'" &_
				","&theForFeit4&","&theInsurance&","&theUseTool&",'"&trim(rsA("ProjectID"))&"'" &_
				",'',null,''" &_
				",'','','"&trim(rsA("MemberStation"))&"'" &_
				",'"&trim(rsA("BillUnitID"))&"','"&trim(rsA("BillMemID1"))&"','"&trim(rsA("BillMem1"))&"'" &_
				",'"&trim(rsA("BillMemID4"))&"','"&trim(rsA("BillMem4"))&"'" &_
				",'"&trim(rsA("BillMemID2"))&"','"&trim(rsA("BillMem2"))&"'" &_
				",'"&trim(rsA("BillMemID3"))&"','"&trim(rsA("BillMem3"))&"'" &_
				",'"&trim(rsA("BillFillerMemberID"))&"','"&trim(rsA("BillFiller"))&"'" &_
				","&theBillFillDate&","&theDealLineDate&",'0',0,"&theRecordDate&",'" & trim(rsA("RecordMemberID")) &"'" &_
				",'"&trim(rsA("Note"))&"','1','"&trim(rsA("RuleVer"))&"'" &_
				",'"&trim(rsA("DriverSex"))&"','"&trim(rsA("ImageFileName"))&"'" &_
				""&valueAdd&"," & theJurgeDay &_
				")"
				'response.write strInsert
				'response.end
				conn.execute strInsert  

		ConnExecute "民眾檢舉審核通過:"&strInsert,371
	End If
	rsA.close
	Set rsA=Nothing 
	'寫入BILLILLEGALIMAGE
	strSqlB="select * from BILLILLEGALIMAGETemp2 where BillSn=" & Trim(request("CheckSn"))
	set rsB=conn.execute(strSqlB)
	If Not rsB.eof Then
		'只將有效照片寫到舉發資料
		fileTemp1=""
		fileTemp2=""
		fileTemp3=""
		fileTemp4=""
		If Trim(request("chkImgNoUseA"))="1" Then
			If trim(request("ImageFileNameA"))<>"" Then
				fileTemp1=trim(request("ImageFileNameA"))
			End If 
		End If 
		If Trim(request("chkImgNoUseB"))="1" Then
			If trim(request("ImageFileNameB"))<>"" Then
				If fileTemp1="" Then
					fileTemp1=trim(request("ImageFileNameB"))
				Else
					fileTemp2=trim(request("ImageFileNameB"))
				End If 				
			End If 
		End If 
		If Trim(request("chkImgNoUseC"))="1" Then
			If trim(request("ImageFileNameC"))<>"" Then
				If fileTemp1="" Then
					fileTemp1=trim(request("ImageFileNameC"))
				ElseIf fileTemp2="" Then
					fileTemp2=trim(request("ImageFileNameC"))
				Else 
					fileTemp3=trim(request("ImageFileNameC"))
				End If 				
			End If 
		End If
		If Trim(request("chkImgNoUseD"))="1" Then
			If trim(request("ImageFileNameD"))<>"" Then
				If fileTemp1="" Then
					fileTemp1=trim(request("ImageFileNameD"))
				ElseIf fileTemp2="" Then
					fileTemp2=trim(request("ImageFileNameD"))
				ElseIf fileTemp3="" Then
					fileTemp3=trim(request("ImageFileNameD"))
				Else 
					fileTemp4=trim(request("ImageFileNameD"))
				End If 				
			End If 
		End If
		strBillImage="Insert Into BILLILLEGALIMAGE(BillSn,BillNo,ImageFileNameA,ImageFileNameB,ImageFileNameC," &_
		"ImageFileNameD,IISImagePath) " &_
		"values("&sMaxSN&",'"&UCase(trim(rsB("Billno")))&"','"&fileTemp1&"'" &_
		",'"&fileTemp2&"','"&fileTemp3&"'" &_
		",'"&fileTemp4&"','"&trim(rsB("IISImagePath"))&"')"

		conn.execute strBillImage  
		
		'將審核無效照片設為-1
		strfileFlag=""
		FileNameArray=Array(trim(rsB("ImageFileNameA")),trim(rsB("ImageFileNameB")),trim(rsB("ImageFileNameC")),trim(rsB("ImageFileNameD")))
		ColArray=Array("ImageFlagA","ImageFlagB","ImageFlagC","ImageFlagD")
		If Trim(request("chkImgNoUseA"))="-1" Then
			For i=0 To UBound(FileNameArray)
				If Trim(FileNameArray(i))=Trim(request("ImageFileNameA")) Then
					strfileFlag=Trim(ColArray(i))&"='-1'"
					Exit for
				End If 
			Next
		End If 
		If Trim(request("chkImgNoUseB"))="-1" Then
			For i=0 To UBound(FileNameArray)
				If Trim(FileNameArray(i))=Trim(request("ImageFileNameB")) Then
					If strfileFlag="" Then
						strfileFlag=Trim(ColArray(i))&"='-1'"
					Else
						strfileFlag=strfileFlag&","&Trim(ColArray(i))&"='-1'"
					End If 
					
					Exit for
				End If 
			Next
		End If 
		If Trim(request("chkImgNoUseC"))="-1" Then
			For i=0 To UBound(FileNameArray)
				If Trim(FileNameArray(i))=Trim(request("ImageFileNameC")) Then
					If strfileFlag="" Then
						strfileFlag=Trim(ColArray(i))&"='-1'"
					Else
						strfileFlag=strfileFlag&","&Trim(ColArray(i))&"='-1'"
					End If 
					Exit for
				End If 
			Next
		End If 
		If Trim(request("chkImgNoUseD"))="-1" Then
			For i=0 To UBound(FileNameArray)
				If Trim(FileNameArray(i))=Trim(request("ImageFileNameD")) Then
					If strfileFlag="" Then
						strfileFlag=Trim(ColArray(i))&"='-1'"
					Else
						strfileFlag=strfileFlag&","&Trim(ColArray(i))&"='-1'"
					End If 
					Exit for
				End If 
			Next
		End If 
	End If
	rsB.close
	Set rsB=Nothing 
	If strfileFlag<>"" Then
		strImgUpd="Update BILLILLEGALIMAGETemp2 set "&strfileFlag&" where BillSn=" & Trim(request("CheckSn"))
		conn.execute strImgUpd
	End If 
	'將舉發BILL SN寫回檢舉資料billbaseTmp
	strUpd1="Update billbaseTmp set BillSn="&sMaxSN  &_
		" where Sn=" & ReportSn
	conn.execute strUpd1
	'將BillBaseTemp2改為已審核
	strUpd2="Update BillBaseTemp2 set CheckFlag='1'"  &_
		" where Sn=" & Trim(request("CheckSn"))
	conn.execute strUpd2
	'寫入審核紀錄
	strIns2="Insert into ReportCaseCheckRecord(Sn,ReportSN,BillTempSN,CheckFlag,RecordMemberID" &_
		",RecordDate,Note)" &_
		" values((select nvl(max(Sn),0)+1 from ReportCaseCheckRecord),"&ReportSn &_
		",'"&Trim(request("CheckSn"))&"','1',"&Trim(session("User_ID"))&"" &_
		",sysdate,''" &_
		")"
	conn.execute strIns2
%>
<script language="JavaScript">

	//alert("儲存完成!");
	//opener.myForm.submit();
	//window.close();
</script>
<%
Else
	%>
	<script language="JavaScript">
		alert("限速或實速超過300Km，請確認是否正確！！");
	</script>
	<%
End If
	If chkIllegalDateAndCar_KS=1 Then
%>
	<script language="JavaScript">
		alert("<%=chkAlertString%>");
	</script>
<%
	End If 
end if
'無效
if trim(request("kinds"))="VerifyResultNull" then
	strUpd="Update billbaseTmp set BillStatus='6'" &_
		" where Sn=" & Trim(request("CheckSn"))
	conn.execute strUpd

	ConnExecute "民眾檢舉無效案件:"&strUpd,372
%>
<script language="JavaScript">
	
	alert("儲存完成!");
	opener.myForm.submit();
	window.close();
</script>
<%
end if

'response.write Session("ReportCaseCheckSn")
FirstSn=""
UpSn=""
DownSn=""
LastSn=""
AllSn=0
If Trim(Session("ReportCaseCheckSn"))<>"" Then
	ThisSn=-1
	ArrayReportCaseCheckSn=Split(Trim(Session("ReportCaseCheckSn")),",")
	For i=0 To UBound(ArrayReportCaseCheckSn)
		If Trim(ArrayReportCaseCheckSn(i))=Trim(request("CheckSn")) Then
			ThisSn=i
			Exit for
		End If 
	Next 
	FirstSn=Trim(ArrayReportCaseCheckSn(0))
	If ThisSn>0 Then
		UpSn=Trim(ArrayReportCaseCheckSn(ThisSn-1))
	End If 
	If ThisSn<UBound(ArrayReportCaseCheckSn) Then
		DownSn=Trim(ArrayReportCaseCheckSn(ThisSn+1))
	End If 
	LastSn=Trim(ArrayReportCaseCheckSn(UBound(ArrayReportCaseCheckSn)))
	AllSn=UBound(ArrayReportCaseCheckSn)+1
End If 
'response.write "<br>" & UpSn & "/" & DownSn & " " & FirstSn & "/" &LastSn
PicturePath="/ReportCaseImage"

strSql1="select * from BillBaseTemp2 where Sn=" & Trim(request("CheckSn"))
set rs1=conn.execute(strSql1)

%>
<title>數位固定桿違規影像審核</title>
<style type="text/css">
<!--
.style1 {font-size: 12px}
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
.style66 {
color: #FF0000;
font-size: 12px ;
}
.style67 {
color: #0033CC;
font-size: 11px ;
}
.btn2 {font-size: 13px}
.Text1{
font-weight:bold;
}
.Text2{
line-height:23px ;
font-size: 20px ;
font-weight:bold;
}
.styleA2 {font-size: 28px; line-height:100%;}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="myForm" method="post"> 
<table width='1150' border='1' align="left" cellpadding="0">
	<tr>
		<td rowspan="3" valign="top" >
		<!-- 影像大圖 -->
	<%if not rs1.eof Then
		file1=""
		file2=""
		file3=""
		file4=""
		
		strImgFile="select * from BILLILLEGALIMAGETemp2 where billSn="&Trim(rs1("SN"))
		Set rsImgFile=conn.execute(strImgFile)
		If Not rsImgFile.eof Then
			If Trim(rsImgFile("IMAGEFILENAMEA"))<>"" Then
				file1= Trim(rsImgFile("IISIMAGEPATH"))&Trim(rsImgFile("IMAGEFILENAMEA"))
			End If 
			If Trim(rsImgFile("IMAGEFILENAMEB"))<>"" Then
				file2= Trim(rsImgFile("IISIMAGEPATH"))&Trim(rsImgFile("IMAGEFILENAMEB"))
			End If
			If Trim(rsImgFile("IMAGEFILENAMEC"))<>"" Then
				file3= Trim(rsImgFile("IISIMAGEPATH"))&Trim(rsImgFile("IMAGEFILENAMEC"))
			End If
			If Trim(rsImgFile("IMAGEFILENAMED"))<>"" Then
				file4= Trim(rsImgFile("IISIMAGEPATH"))&Trim(rsImgFile("IMAGEFILENAMED"))
			End If
		End If 
		rsImgFile.close
		Set rsImgFile=Nothing 

	%>
		<input type="hidden" name="ImageFileNameA" value="<%
		if file1<>"" Then
			ImageFileNameAArray=Split(file1,"/")
			response.write ImageFileNameAArray(UBound(ImageFileNameAArray))
			ImageFileNameATemp=ImageFileNameAArray(UBound(ImageFileNameAArray))
			ImageFileNameTemp="/ReportCaseImage" & Replace(file1,ImageFileNameAArray(UBound(ImageFileNameAArray)),"")
		End if
		%>">
		<input type="hidden" name="ImagePathName" value="<%=ImageFileNameTemp%>">

		<%if file1<>"" then%>
		<%
		If UCase(Right(file1,3))="BMP" Or UCase(Right(file1,3))="PNG" Or UCase(Right(file1,3))="JPG" Or UCase(Right(file1,4))="JPEG" Or UCase(Right(file1,3))="GIF" Then
			IsPicture="1"
		Else
			IsPicture="0"
		End If 
		
		bPicWebPath=file1
		If IsPicture="1" then
			%>
			<img src="<%=bPicWebPath%>" border=1 height="<%
			response.write "460"
			%>" <%
			'放大鏡功能
			if isBig="Y"  then
			%>onmousemove="show(this, '<%=bPicWebPath%>')" onmousedown="show(this, '<%=bPicWebPath%>')"<%
			end if
			%> id="imgSource"> 
			<input type="button" name="btnImgNoUseA" value="相片無效" onclick="setImageNotUse('A');">
			<input type="hidden" name="chkImgNoUseA" value="1">
			
		<%else%>
			<a href="<%=bPicWebPath%>" target="_blank" style="font-size: 18px;">開啟檔案 <%
			response.write "..."&Right(ImageFileNameATemp,14)
			
			%></a>
		<%End If %>
			<div id="div1" style="position:absolute; overflow:hidden; width:<%
			'If sys_City=ApconfigureCityName Then
				response.write "230"
			'Else
			'	response.write "210"
			'End If 
			%>px; height:<%
			'If sys_City=ApconfigureCityName Then
				response.write "110"
			'Else
			'	response.write "90"
			'End If 
			%>px; left:<%
			if trim(request("divX"))="" then
				response.write "640"
			else
				response.write trim(request("divX"))
			end if
			%>px; top:<%
			if trim(request("divY"))="" Then
				response.write "160"
			else
				response.write trim(request("divY"))
			end if
			%>px; z-index:1;border-right: white thin ridge; border-top: white thin ridge; border-left: white thin ridge; border-bottom: white thin ridge <%
		'放大鏡功能
		if isBig="N" Or IsPicture="0" then
		%> ;visibility: hidden;<%
		end if
		%>" onMousedown="initializedragie( )">
				<img id="BigImg" style='position:relative' src="<%=bPicWebPath%>">
			
		<%end if%>
	<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
	<%end if%>
		</td>
		<td height="100" width="23%" align="center">
	<%if not rs1.eof Then
		if file2<>"" Then
	%>

		<input type="hidden" name="ImageFileNameB" value="<%
			ImageFileNameBarray=Split(file2,"/")
			response.write ImageFileNameBarray(UBound(ImageFileNameBarray))
			ImageFileNameBTemp=ImageFileNameBarray(UBound(ImageFileNameBarray))
		%>">
		<!-- 影像小圖 B-->
		<%
			If UCase(Right(file2,3))="BMP" Or UCase(Right(file2,3))="PNG" Or UCase(Right(file2,3))="JPG" Or UCase(Right(file2,4))="JPEG" Then
				IsPictureB="1"
			Else
				IsPictureB="0"
			End If 
			sPicWebPath2=file2

			If IsPictureB="1" then
		%>
		<img src="<%=sPicWebPath2%>" border=1 id="SmallImgB" width="<%
			response.write "230"
		%>" height="<%
			response.write "130"
		%>" <%
			response.write "ondblclick=""ChangeImgB()"""
			'response.write "ondblclick=""OpenPic('"&sPicWebPath2&"')"""
		%>>
			<input type="button" name="btnImgNoUseB" value="相片無效" onclick="setImageNotUse('B');">
			<input type="hidden" name="chkImgNoUseB" value="1">
			<%else%>
			<a href="<%=sPicWebPath2%>" target="_blank" style="font-size: 18px;">開啟檔案 <%
			response.write "..."&Right(ImageFileNameBTemp,14)
			%></a>
			<%end if%>
		<%else%>
		&nbsp;
		<%end if%>
	<%end if%>
		</td>
		
	</tr>
	<tr>
		<td height="100" align="center">
	<%if not rs1.eof Then
		if file3<>"" Then
	%>
		<input type="hidden" name="ImageFileNameC" value="<%
			ImageFileNameCarray=Split(file3,"/")
			response.write ImageFileNameCarray(UBound(ImageFileNameCarray))
			ImageFileNameCTemp=ImageFileNameCarray(UBound(ImageFileNameCarray))
		%>">
		<!-- 影像小圖 C-->
		<%
			If UCase(Right(file3,3))="BMP" Or UCase(Right(file3,3))="PNG" Or UCase(Right(file3,3))="JPG" Or UCase(Right(file3,4))="JPEG" Then
				IsPictureC="1"
			Else
				IsPictureC="0"
			End If 

			sPicWebPath=file3
			If IsPictureC="1" then
		%>
		<img src="<%=sPicWebPath%>" border=1 id="SmallImgC" width="<%
			response.write "230"
		%>" height="<%
			response.write "130"
		%>"  <%
			response.write "ondblclick=""ChangeImgC()"""
			'response.write "ondblclick=""OpenPic('"&sPicWebPath&"')"""
		%>>
			<input type="button" name="btnImgNoUseC" value="相片無效" onclick="setImageNotUse('C');">
			<input type="hidden" name="chkImgNoUseC" value="1">
			<%else%>
			<a href="<%=sPicWebPath%>" target="_blank" style="font-size: 18px;">開啟檔案 <%
			response.write "..."&Right(ImageFileNameCTemp,14)
			%></a>
			<%end if%>
		<%else%>
		&nbsp;
		<%end if%>
	<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
	<%end if%>
		</td>
	</tr>
	<tr>
		<td height="100" align="center">
	<%if not rs1.eof Then
		if file4<>"" Then
	%>
		<input type="hidden" name="ImageFileNameD" value="<%
			ImageFileNameDarray=Split(file4,"/")
			response.write ImageFileNameDarray(UBound(ImageFileNameDarray))
			ImageFileNameDTemp=ImageFileNameDarray(UBound(ImageFileNameDarray))
		%>">
		<!-- 影像小圖 D-->
		<%
			If UCase(Right(file4,3))="BMP" Or UCase(Right(file4,3))="PNG" Or UCase(Right(file4,3))="JPG" Or UCase(Right(file4,4))="JPEG" Then
				IsPictureD="1"
			Else
				IsPictureD="0"
			End If 

			sPicWebPath3=file4

			If IsPictureD="1" then
		%>
		<img src="<%=sPicWebPath3%>" border=1 id="SmallImgD" width="<%
			response.write "230"
		%>" height="<%
			response.write "130"
		%>" <%
'		If (sys_City="宜蘭縣" And Trim(Session("Unit_ID"))="TQ00") Or sys_City="高雄市" Then
'			response.write "ondblclick=""ChangeImg()"""
'		Else
			response.write "ondblclick=""ChangeImgD()"""
			'response.write "ondblclick=""OpenPic('"&sPicWebPath3&"')"""
'		End If 
		%>>
			<input type="button" name="btnImgNoUseD" value="相片無效" onclick="setImageNotUse('D');">
			<input type="hidden" name="chkImgNoUseD" value="1">
			<%else%>
			<a href="<%=sPicWebPath3%>" target="_blank" style="font-size: 18px;">開啟檔案 <%
			response.write "..."&Right(ImageFileNameDTemp,14)
			%></a>
			<%end if%>
		<%else%>
		&nbsp;
		<%end if%>
	<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
	<%end if%>
		</td>
	</tr>
	<tr>
		<td height="100" colspan="3" valign="top">
		<%if not rs1.eof then%>
		<table width='100%' border='1' align="left" cellpadding="0">
			<tr>
				<td bgcolor="#FFFFCC" width="6%"><div align="right"> <span class="style3">＊</span>車號&nbsp;</div></td>
				<td width="12%">
				<input type="text" size="9" name="CarNo" onBlur="getVIPCar();" value="<%
				if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
					response.write trim(rs1("CarNo"))
				end if
				%>" style=ime-mode:disabled maxlength="8" class="Text2" onkeydown="funTextControl(this);">
				<span class="style6">
			    <div id="Layer7" style="position:absolute; width:70px; height:24px; z-index:0;  border: 1px none #000000; color: #FF0000; font-weight: bold;"><%
			if trim(Session("SpecUser"))="1" then
				strSC="select count(*) as cnt from SpecCar where CarNo='"&trim(rs1("CarNo"))&"' and RecordStateID<>-1"
				set rsSC=conn.execute(strSC)
				if not rsSC.eof then
					if trim(rsSC("cnt"))<>"0" then
						response.write "＊業管車輛"
					end if
				end if
				rsSC.close
				set rsSC=nothing
			end if
				%></div>
				</span>
				
				</td>
				<td bgcolor="#FFFFCC" width="8%"><div align="right"><span class="style3">＊</span>車種&nbsp;</div>
				</td>
				<td colspan="3" >
                    <!-- 簡式車種 -->
                    <input type="text" maxlength="1" size="2" value="<%
                    if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
                    	response.write trim(rs1("CarSimpleID"))
                    end if
                    %>" name="CarSimpleID" onBlur="getRuleAll();" style=ime-mode:disabled onkeydown="funTextControl(this);">
                    <div id="Layer012" style="display: inline; width:300px; height:27px; z-index:1; visibility: visible;">
                    <font color="#0000FF" size="2">&nbsp;1汽車 / 2拖車 / 3重機/ 4輕機/5動力機械/6臨時車牌</font></div>
				</td>
				<td bgcolor="#FFFFCC" width="7%"><div align="right"><span class="style3">＊</span>違規時間</div></td>
				<td width="13%" colspan="3">
							<!-- 違規日期 -->
					<input type="text" size="6" maxlength="7" name="IllegalDate" class='Text1' value="<%
					if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then 
						response.write gInitDT(rs1("IllegalDate"))
					end If
					%>" onBlur="getBillFillDate()" style=ime-mode:disabled onkeydown="funTextControl(this);" onkeyup="IllegalDateKeyUP()" >&nbsp;
							<!-- 違規時間 -->
					<input type="text" size="3" maxlength="4" name="IllegalTime" class='Text1' value="<%
					if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then 
						response.write Right("00"&hour(rs1("IllegalDate")),2)&Right("00"&minute(rs1("IllegalDate")),2)
					end if
					%>" onBlur="value=value.replace(/[^\d]/g,'')" style=ime-mode:disabled onkeydown="funTextControl(this);" onKeyUP="IllegalTimeKeyUP()">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span>地點&nbsp;</div></td>
				<td colspan="3">
					<input type="text" size="4" value="<%
					response.write Trim(rs1("IllegalAddressID"))
					%>" name="IllegalAddressID" onKeyUp="getillStreet();" onblur="funGetSpeedRule()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<%'if sys_City="高雄市" then %>
						區號
						<input type="text" class="btn5" size="3" value="<%=Trim(rs1("IllegalZip"))%>" name="IllegalZip" onKeyUp="getIllZip();" onkeydown="funTextControl(this);" maxlength="3">
						<Input type="hidden" name="OldIllegalZip" value="<%=Trim(request("IllegalZip"))%>">
						
						<img src="../Image/BillkeyInButtonsmall.jpg" onclick="QryIllegalZip();">
						<div id="LayerIllZip" style="display: inline; width:160px; height:30; z-index:0;  border: 1px none #000000;""><%
					if Trim(rs1("IllegalZip"))<>"" then
						strZip1="select ZipName from Zip where ZipNo='"&Trim(rs1("IllegalZip"))&"'"
						set rsZip1=conn.execute(strZip1)
						if not rsZip1.eof then
							response.write trim(rsZip1("ZipName"))
						end if
						rsZip1.close
						set rsZip1=nothing
					end if
					%></div><br>
					<%'end if%>
					<input type="text" size="40" value="<%
					if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
						response.write trim(rs1("IllegalAddress"))
					end If
					%>" name="IllegalAddress" style=ime-mode:active onblur="funGetSpeedRule()" onkeyup="AutoGetIllStreet();" onkeydown="funTextControl(this);">
					<input type="checkbox" name="chkHighRoad" value="1" <%if trim(request("chkHighRoad"))="1" then response.write "checked"%> onclick="setIllegalRule()" <%if sys_City="南投縣" then response.write "disabled"%>>
					<div id="Layerert45" style="display: inline; width:30px; height:28px; z-index:0;  layer-background-color: #FFFFFF; border: 1px none #000000;"><span class="style1">快速道路</span></div>
                </td>
				<td bgcolor="#FFFFCC" ><div align="right"><span class="style3">＊</span>法條一</div></td>
				<td colspan="5">
					<input type="text" maxlength="9" size="7" value="<%
					if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
						response.write trim(rs1("Rule1"))
					end If
					%>" name="Rule1" onKeyUp="getRuleData1();" style=ime-mode:disabled onkeydown="funTextControl(this);" >
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_Law.asp?LawOrder=1&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<img src="../Image/BillLawPlusButton_Small.JPG" onclick="Add_LawPlus()" alt="附加說明">
					實際
					<input type="text" size="2" maxlength="3" name="IllegalSpeed" class='Text1' value="<%
					if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) then
						response.write trim(rs1("IllegalSpeed"))
					end If
					%>" onkeyup="IllegalSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					限制
					<input type="text" size="2" name="RuleSpeed" maxlength="3" class='Text1' value="<%
					if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
						response.write trim(rs1("RuleSpeed"))
					end If
					%>" onBlur="RuleSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					&nbsp;
					<span class="style5">
					<div id="Layer1" style="position:absolute ; width:230px; height:28px; z-index:0;  layer-background-color: #FFFFFF; border: 1px none #000000;"><%
					strR1="select * from Law where itemid='"&trim(rs1("Rule1"))&"' and Version=2"
					Set rsR1=conn.execute(strR1)
					If Not rsR1.eof Then
						response.write rsR1("IllegalRule")
					End If 
					rsR1.close
					Set rsR1=Nothing 
					%></div></span>
					<input type="hidden" name="ForFeit1" value="<%

					%>">
					
				</td>
		    </tr>
			<tr>
				<td bgcolor="#FFFFCC" ><div align="right">法條二</div></td>
				<td colspan="3">
					<input type="text" maxlength="9" size="7" value="<%
					if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
						response.write trim(rs1("Rule2"))
					end If
					%>" name="Rule2" onkeyup="getRuleData2();" onkeydown="funTextControl(this);" style=ime-mode:disabled >
					<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_Law.asp?LawOrder=2&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer2" style="position:absolute ; width:260px; height:28px; z-index:0; border: 1px none #000000;"><%
					strR1="select * from Law where itemid='"&trim(rs1("Rule2"))&"' and Version=2"
					Set rsR1=conn.execute(strR1)
					If Not rsR1.eof Then
						response.write rsR1("IllegalRule")
					End If 
					rsR1.close
					Set rsR1=Nothing 
					%></div>
					</span>
					<input type="hidden" name="ForFeit2" value="<%
			
					%>">

				</td>
				<td bgcolor="#FFFFCC" height="30"><div align="right"><span class="style3">＊</span>舉發人&nbsp;</div></td>
		  		<td colspan="<%
				response.write "3"

				%>">
					<input type="text" size="9" name="BillMem1" value="<%
				If Trim(rs1("BillMemID1"))<>"" Then
					strMem="select * from Memberdata where MemberID="&Trim(rs1("BillMemID1"))
					Set rsMem=conn.execute(strMem)
					If Not rsMem.eof Then
						response.write Trim(rsMem("LoginID"))
					End If
					rsMem.close
					Set rsMem=nothing 
				End If 
				%>" onKeyUp="getBillMemID1();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_MemID.asp?MemOrder=1","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer12" style="display: inline; width:60px; height:30;  z-index:0;  border: 1px none #00000; "><%
				If Trim(rs1("BillMem1"))<>"" Then
					response.write Trim(rs1("BillMem1"))
				End If 
					%></div>
					</span>
					<input type="hidden" value="<%%>" name="BillMemID1">
					<input type="hidden" value="<%
						
					%>" name="BillMemName1">

					<td bgcolor="#FFFFCC" height="30"><div align="right" style="font-size: 12px ;">舉發人二</div></td>
					<td >
						
						<input type="text" size="7" name="BillMem2" value="<%
				If Trim(rs1("BillMemID2"))<>"" Then
					strMem="select * from Memberdata where MemberID="&Trim(rs1("BillMemID2"))
					Set rsMem=conn.execute(strMem)
					If Not rsMem.eof Then
						response.write Trim(rsMem("LoginID"))
					End If
					rsMem.close
					Set rsMem=nothing 
				End If 
					%>" onKeyUp="getBillMemID2();" style=ime-mode:disabled onkeydown="funTextControl(this);">
						<img src="../Image/BillkeyInButtonsmall.jpg" onclick='window.open("Query_MemID.asp?MemOrder=2","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
						<span class="style5">
						<div id="Layer13" style="display: inline; width:60px; height:30;  z-index:0;  border: 1px none #000000; "><%
				If Trim(rs1("BillMem2"))<>"" Then
					response.write Trim(rs1("BillMem2"))
				End If 
						%></div>
						</span>
						<input type="hidden" value="<%=BillRecordID2%>" name="BillMemID2">
						<input type="hidden" value="<%
			
						%>" name="BillMemName2">
					</td>
			
			</tr>
			<tr>

				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span><span class="style4">舉發單位</span></div></td>
				<td colspan="3">
					<input type="text" size="4" name="BillUnitID" value="<%=Trim(rs1("BillUnitID"))%>" onKeyUp="getUnit();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_Unit.asp?SType=U","WebPage2","left=0,top=0,location=0,width=800,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer6" style="display: inline; width:200px; height:30px; z-index:0;  border: 1px none #000000; "><%
					if Trim(rs1("BillUnitID"))<>"" then
						strUnitName="select UnitName from UnitInfo where UnitID='"&Trim(rs1("BillUnitID"))&"'"
						set rsUnitName=conn.execute(strUnitName)
						if not rsUnitName.eof then
							response.write trim(rsUnitName("UnitName"))
						end if
						rsUnitName.close
						set rsUnitName=nothing
					end if
					%></div>
					</span>
					&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <span class="style4">民眾檢舉時間</span>
					<input type="text" name="JurgeDay" value="<%
					if trim(rs1("JurgeDay"))<>"" and not isnull(rs1("JurgeDay")) then 
						response.write gInitDT(rs1("JurgeDay"))
					end If
					%>" size="10" maxlength="7" style=ime-mode:disabled onkeydown="funTextControl(this);" onblur="this.value=this.value.replace(/[^\d]/g,'');">
				</td>
				<td bgcolor="#FFFFCC" width="8%">

				<div id="Layer110" style="position:absolute; width:265px; height:27px; z-index:1; background-color: #FFCCCC; visibility: hidden;">
				<font color="#0000FF" size="2">1大貨/2大客/3砂石/4土方/5動力/6貨櫃/7大型重機/8拖吊/9(550cc)重機 /10計程車/ 11危險物品 </font>
				</div>

				<div align="right"><span class="style3">＊</span>填單日期</div></td>
				<td width="9%">
				
				&nbsp;<input type="text" size="6" value="<%=ginitdt(date)%>" maxlength="7" name="BillFillDate" onBlur="getDealLineDate()" style=ime-mode:disabled onkeydown="funTextControl(this);">

				<input type="hidden" name="SelSN" value="<%=trim(rs1("SN"))%>">

				</td>

				<td bgcolor="#FFFFCC" align="right" width="8%">輔助車種&nbsp;</td>
				<td width="6%">
                &nbsp;<input type="text" maxlength="2" size="4" value="<%
				if trim(rs1("CarAddID"))<>"" and not isnull(rs1("CarAddID")) then 
					response.write rs1("CarAddID")
				end If
				%>" name="CarAddID" onBlur="getAddID();" style=ime-mode:disabled onfocus="Layer110.style.visibility='visible';" onkeydown="funTextControl(this);">
                
				</td>

				<td bgcolor="#FFFFCC" width="8%">
		
				<div align="right">專案代碼&nbsp;</div></td>
				<td width="12%">
					&nbsp;<input type="text" size="5" value="<%
				if trim(rs1("ProjectID"))<>"" and not isnull(rs1("ProjectID")) then 
					response.write rs1("ProjectID")
				end If
				%>" name="ProjectID" style=ime-mode:disabled onkeyup="ProjectF5()" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButtonsmall.jpg"  onClick='window.open("Query_Project.asp","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer001" style="position:absolute ; width:180px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000; visibility: hidden;"></div>

					<!-- <div id="Layer5012" style="position:absolute; width:125px; height:27px; z-index:1; visibility: visible;">
                    <font color="#0000FF" size="2">&nbsp;1檢舉達人<br>&nbsp;9拖吊</font></div> -->

					<!-- 採証工具 -->
					<input maxlength="1" size="4" value="1" name="UseTool"  onkeyup="getFixID();" type='hidden' style=ime-mode:disabled> 
			        <div id="Layer11" style="position:absolute; width:275px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold; visibility: hidden;"> <font color="#0000FF">&nbsp;&nbsp;<font color="#000000">固定桿編號：</font></font>
                    <input type='text' size='6' name='FixID' value='<%
					'if trim(rs1("SiteCode"))<>"" and not isnull(rs1("SiteCode")) then
					'	response.write trim(rs1("SiteCode"))
					'end if
					%>' onBlur="setFixEquip();" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButtonsmall.jpg"  onclick='window.open("Query_FixEquip.asp","WebPageFix","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					</div>
					<!-- <font color="#ff000" size="2"> 1固定桿/ 2雷達三腳架/ 3相機</font> -->
			    <!-- 備註 -->
					<input type="hidden" size="29" value="<%
					if trim(rs1("Note"))<>"" and not isnull(rs1("Note")) then
						response.write trim(rs1("Note"))
					end If
					if sys_City="花蓮縣" then	
						if trim(rs1("SiteCode"))<>"" and not isnull(rs1("SiteCode")) then
							response.write trim(rs1("SiteCode"))
						end If
					End If 
					%>" name="Note" style=ime-mode:active>
				</td>

			</tr>
				
		</table>
		<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
		<%end if%>
		</td>
	</tr>
	<tr bgcolor="#FFCC33">
		<td height="28" colspan="3" align="center">


			<input type="button" value="審核通過 F2" onclick="InsertBillVase();"  <%
		if rs1.eof then
			response.write "disabled"
		ElseIf Trim(rs1("CheckFlag"))<>"0" Then
			response.write "disabled"
		end if
			%> style="font-size: 10pt; width: 100px; height: 27px">
			
			<input type="button" name="Submit2932" onClick="funVerifyResult();" value="審核無效 F9" <%
		if rs1.eof then
			response.write "disabled"
		ElseIf Trim(rs1("CheckFlag"))<>"0" Then
			response.write "disabled"
		end if
			%> style="font-size: 10pt; width: 100px; height: 27px">
			<img src="/image/space.gif" width="29" height="8">
			<input type="hidden" name="kinds" value="">
			
			<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Image_ReportCase_Check.asp?CheckSn=<%=FirstSn%>'" value="<< 第一筆 Home" style="font-size: 9pt; width: 90px; height: 27px" <%
			If UpSn="" Then
				response.write "Disabled"
			End If 
			%>>
			<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Image_ReportCase_Check.asp?CheckSn=<%=UpSn%>'" value="< 上一筆 PgUp" style="font-size: 9pt; width: 90px; height: 27px"  <%
			If UpSn="" Then
				response.write "Disabled"
			End If 
			%>>
			<%=ThisSn+1 & " / " & AllSN%>
			<input type="button" name="SubmitNext" onClick="location='BillKeyIn_Image_ReportCase_Check.asp?CheckSn=<%=DownSn%>'" value="下一筆 PgDn >" style="font-size: 9pt; width: 90px; height: 27px"  <%
			If DownSn="" Then
				response.write "Disabled"
			End If 
			%>>
			<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Image_ReportCase_Check.asp?CheckSn=<%=LastSn%>'" value="最後一筆 End >>" style="font-size: 9pt; width: 90px; height: 27px" <%
			If DownSn="" Then
				response.write "Disabled"
			End If 
			%>>


			<img src="/image/space.gif" width="29" height="8">
			<input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉 F8" style="font-size: 10pt; width: 70px; height: 27px">
			
			
			
			<img src="/image/space.gif" width="29" height="8">


             <input type="hidden" name="Tmp_Order" value="<%=Session("BillCnt_Image")%>">
				<input type="hidden" name="CheckSn" value="<%=Trim(request("CheckSn"))%>">
				<!-- 逕舉類別 -->
				<input type="hidden" size="3" maxlength="1" value="2" name="BillType" readonly>
				<!-- 應到案日期 -->
				<input type="hidden" size="12" maxlength="7" name="DealLineDate">
				<!-- 應到案處所 -->
				<input type="hidden" size="10" value="" name="MemberStation">
				<input type="hidden" value="" name="Rule3">
				<input type="hidden" name="ForFeit3" value="">
				<input type="hidden" value="" name="Rule4">
				<input type="hidden" name="ForFeit4" value="">
				<input type="hidden" value="" name="Billno1">
				<input type="hidden" value="" name="Insurance">
				<input type="hidden" value="" name="BillMemID3">
				<input type="hidden" value="" name="BillMemID4">
				<input type="hidden" value="" name="BillMemName3">
				<input type="hidden" value="" name="BillMemName4">
				<!-- <input type="button" value="？" name="StationSelect" onclick='window.open("Query_Station.asp","WebPage1","left=0,top=0,location=0,width=660,height=375,resizable=yes,scrollbars=yes")'> -->
				<div id="Layer5" style="position:absolute ; width:221px; height:24px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000; visibility :hidden;"></div>
				<input type="hidden" name="SessionFlag" value="1">
				<!--浮動視窗座標-->
				<input type="hidden" name="divX" value="<%
				if trim(request("divX"))="" then
					If sys_City=ApconfigureCityName Then
						response.write "650"
					elseIf sys_City="苗栗縣" Then
						response.write "1210"
					Else
						response.write "540"
					End If 
				else
					response.write trim(request("divX"))
				end if
				%>">
				<input type="hidden" name="divY" value="<%
				if trim(request("divY"))="" Then
					If sys_City=ApconfigureCityName Then
						response.write "490"
					elseIf sys_City="苗栗縣" Then
						response.write "40"
					Else
						response.write "400"
					End If 
				else
					response.write trim(request("divY"))
				end if
				%>">
				
		</td>
	</tr>
<%If sys_City="宜蘭縣" then%>
	<tr>
	<td colspan="2">
	<a href="逕舉相片建檔.doc" target="_blank"><font  class="styleA2">使用說明下載</font></a>
	</td>
	</tr>
<%End if%>
</table>

</form>
</body>
<script language="JavaScript">
var TDLawNum=0;
var TDLawErrorLog1=0;
var TDLawErrorLog2=0;
var TDLawErrorLog3=0;
var TDLawErrorLog4=0;
var TDStationErrorLog=0;
var TDUnitErrorLog=0;
var TDFastenerErrorLog1=0;
var TDFastenerErrorLog2=0;
var TDFastenerErrorLog3=0;
var TDMemErrorLog1=0;
var TDMemErrorLog2=0;
var TDMemErrorLog3=0;
var TDMemErrorLog4=0;
var TDIllZipErrorLog=0;
var TDProjectIDErrorLog=0;
var TDVipCarErrorLog=0;
var SpeedError=0;
var TodayDate=<%=ginitdt(date)%>;

var InsertFlag=0;
<%if sys_City="宜蘭縣" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID");
<%elseif sys_City="南投縣" Or sys_City="屏東縣" or sys_City="花蓮縣" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,RuleSpeed,IllegalSpeed||Rule2,BillMem1||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID");
<%elseif sys_City="高雄市" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,BillMem1,BillMem2||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID");
<%elseif sys_City="苗栗縣" then%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalSpeed,RuleSpeed,Rule1,Rule2||IllegalAddressID,IllegalAddress,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,JurgeDay,BillFillDate,ProjectID,CarAddID");
<%else%>
MoveTextVar("CarNo,CarSimpleID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed||Rule2,BillMem1||BillUnitID,JurgeDay,BillFillDate,CarAddID,ProjectID");
<%end if%>

//新增告發單
function InsertBillVase(){
	var error=0;
	var errorString="";


	if (error==0){
		if (InsertFlag==0){
			InsertFlag=1;
			getChkCarIllegalDate();
		}
	}else{
		alert(errorString);
	}
}

//檢查同車號同法條在同一天違規日期及違規時間前後兩小時內
function getChkCarIllegalDate(){
	NewIllDate=myForm.IllegalDate.value;
	NewIllTime=myForm.IllegalTime.value;
	NewIllRule1=myForm.Rule1.value;
	NewIllRule2="";
	NewCarNo=myForm.CarNo.value;
	NewCarSimpleID=myForm.CarSimpleID.value;
	NewBillUnitID=myForm.BillUnitID.value;
	NewIllegalAddress=myForm.IllegalAddress.value;
	runServerScript("getChkCarIllegalDate.asp?CarID="+NewCarNo+"&IllDate="+NewIllDate+"&IllTime="+NewIllTime+"&IllRule1="+NewIllRule1+"&IllRule2="+NewIllRule2+"&CarSimpleID="+NewCarSimpleID+"&BillUnitID="+NewBillUnitID+"&IllegalAddress="+NewIllegalAddress);

	//window.open("getChkCarIllegalDate.asp?CarID="+NewCarNo+"&IllDate="+NewIllDate+"&IllTime="+NewIllTime+"&IllRule1="+NewIllRule1+"&IllRule2="+NewIllRule2+"&CarSimpleID="+NewCarSimpleID+"&BillUnitID="+NewBillUnitID+"&IllegalAddress="+NewIllegalAddress,"WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
}
//檢查同車號同法條在同一天違規日期及違規時間前後兩小時內
function setChkCarIllegalDate(CarCnt,Illdate,RuleDetail)
{
	var ErrorStringChkCarIllegal="";
	if (CarCnt=="1"){
		ChkCarIlldateFlag="1";
	}else{
		ChkCarIlldateFlag="0";
	}
	if (RuleDetail==2){
		alert("舉發單位代號輸入錯誤。");
		InsertFlag=0;
<%if sys_City="高雄市" then%>
	}else if (RuleDetail==3 || RuleDetail==4){
		alert("此車號為業管車輛。");
		InsertFlag=0;
<%end if%>
<%if sys_City="南投縣" then%>
	}else if (RuleDetail==5){
		alert("此車號在違規時間6分鐘內有相同違規法條，請確認是否正確。");
		InsertFlag=0;
<%elseif sys_City="宜蘭縣" then%>
	}else if (RuleDetail==5){
		alert("此車號在違規時間同一日內有違規，請確認是否正確，如須建檔請洽交通隊張良相先生。");
		InsertFlag=0;
<%end if%>
<%if sys_City="台中市" then%>
	}else if (RuleDetail==6){
		alert("此車號在同一違規時間，有相同違規法條，請確認是否正確。");
<%elseif sys_City<>"台東縣" then%>
	}else if (RuleDetail==6){
		alert("此車號在同一違規時間、違規地點，有相同違規法條，請確認是否正確。");
		InsertFlag=0;
<%end if%>
	}else{
		if (RuleDetail==1 || RuleDetail==3){
			ErrorStringChkCarIllegal='違規事實與簡式車種不符，請確認是否正確。\n';
		}
		if (ChkCarIlldateFlag=="1"){
		<%if sys_City="宜蘭縣" Or sys_City="基隆市" Or sys_City="台南市" then%>
			ErrorStringChkCarIllegal=ErrorStringChkCarIllegal+'此車號於'+Illdate+'，有違規舉發記錄，請確認有無連續開單。\n';
		<%else%>
			ErrorStringChkCarIllegal=ErrorStringChkCarIllegal+'此車號於'+Illdate+'，有相同違規舉發，請確認有無連續開單。\n';
		<%end if%>
		}
		<%if sys_City="高雄市" then%>
		if ((myForm.IllegalAddressID.value=='00212' || myForm.IllegalAddressID.value=='00213') && myForm.chkHighRoad.checked==false){
			ErrorStringChkCarIllegal=ErrorStringChkCarIllegal+'此違規地點為快速道路，請確認是否勾選快速道路。\n';
		}
		<%end if%>
		if (ErrorStringChkCarIllegal != ""){
			if(confirm(ErrorStringChkCarIllegal + '\n是否確定要存檔？')){
				myForm.kinds.value="DB_insert";
				myForm.submit();
			}else{
				InsertFlag=0;
			}
		}else{
			myForm.kinds.value="DB_insert";
			myForm.submit();
		}
	}
}
//是否為特殊用車
function getVIPCar(){
	myForm.CarNo.value=myForm.CarNo.value.toUpperCase();
	myForm.CarNo.value=myForm.CarNo.value.replace(" ", "");

}

//檢查輔助車種
function getAddID(){
	//myForm.CarAddID.value=myForm.CarAddID.value.replace(/[^\d]/g,'');
	Layer110.style.visibility='hidden';
	if (myForm.CarAddID.value.length>0){
		if (myForm.CarAddID.value != "1" && myForm.CarAddID.value != "2" && myForm.CarAddID.value != "3" && myForm.CarAddID.value != "4" && myForm.CarAddID.value != "5" && myForm.CarAddID.value != "6" && myForm.CarAddID.value != "7" && myForm.CarAddID.value != "8" && myForm.CarAddID.value != "9" && myForm.CarAddID.value != "10" && myForm.CarAddID.value != "11"){
			alert("輔助車種填寫錯誤!");
			//myForm.CarAddID.value = "";
			myForm.CarAddID.focus();
		}
	}
}
//檢查簡式車種
function getRuleAll(){
	//myForm.CarSimpleID.value=myForm.CarSimpleID.value.replace(/[^\d]/g,'');
	//Layer012.style.visibility='hidden';
	if (myForm.CarSimpleID.value.length>0){
		if (myForm.CarSimpleID.value != "1" && myForm.CarSimpleID.value != "2" && myForm.CarSimpleID.value != "3" && myForm.CarSimpleID.value != "4" && myForm.CarSimpleID.value != "5" && myForm.CarSimpleID.value != "6" && myForm.CarSimpleID.value != "7"){
			alert("簡式車種填寫錯誤!");
			myForm.CarSimpleID.focus();
			myForm.CarSimpleID.value = "";
		}
	}
}
//違規事實1(ajax)
function getRuleData1(flag){
	if (myForm.Rule1.value.length > 6){
		var Rule1Num=myForm.Rule1.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail_forLawPlus.asp?RuleOrder=1&RuleID="+Rule1Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo+"&nowTime=<%=now%>");
	<%if not rs1.eof then%>
		<%'if trim(rs1("ProsecutionTypeID"))<>"R" then%>
		CallChkLaw1();
		<%'end if%>
	<%end if%>
		if (event){
			if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106)){
				if (myForm.Rule1.value.length=="7"){
					if ((myForm.Rule1.value.substr(0,2))!="29" && ((myForm.Rule1.value.substr(0,5))!="33101" && (myForm.Rule1.value.substr(0,2))!="40" && (myForm.Rule1.value.substr(0,5))!="43102")){
						myForm.Rule2.select();
						myForm.IllegalSpeed.value="";
						myForm.RuleSpeed.value="";
					}else{
						if (flag!="NoSelect"){
						<%if sys_City="屏東縣" then%>
							if (myForm.RuleSpeed.value==""){
								myForm.RuleSpeed.select();
							}else{
								myForm.IllegalSpeed.select();
							}
						<%else%>
							myForm.IllegalSpeed.select();
						<%end if %>
						}
					}
				}
			}
		}

	}else if (myForm.Rule1.value.length <= 6 && myForm.Rule1.value.length > 0){
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=1;
	}else{
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=0;
	}
	//AutoGetRuleID(1);
}
function getRuleData2(){
	if (myForm.Rule2.value.length > 6){
		var Rule2Num=myForm.Rule2.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail.asp?RuleOrder=2&RuleID="+Rule2Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo+"&nowTime=<%=now%>");
	<%if not rs1.eof then%>
		CallChkLaw2();
	<%end if%>
		if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106)){
			if (myForm.Rule2.value.length=="7"){
			<%if sys_City="苗栗縣" then%>
				myForm.IllegalAddressID.select();
			<%else%>
				myForm.BillMem1.select();
			<%end if %>
			}
		}
	}else if (myForm.Rule2.value.length <= 6 && myForm.Rule2.value.length > 0){
		Layer2.innerHTML=" ";
		myForm.ForFeit2.value="";
		TDLawErrorLog2=1;
	}else{
		Layer2.innerHTML=" ";
		myForm.ForFeit2.value="";
		TDLawErrorLog2=0;
	}
	//AutoGetRuleID(1);
}
//function TabFocus(){
	//建檔時除了超重超速時游標才跳至限速限量欄位，其它法條則游標不跳至超重超速
//	Rule1tmp=myForm.Rule1.value;
//		if ((Rule1tmp.substr(0,2))!="33" && (Rule1tmp.substr(0,2))!="40" && (Rule1tmp.substr(0,2))!="43" && (Rule1tmp.substr(0,2))!="29"){
			//myForm.BillMem1.focus();
//		}
//}
//到案處所(ajax)
function getStation(){
	if (myForm.MemberStation.value.length > 1){
		var StationNum=myForm.MemberStation.value;
		runServerScript("getMemberStation.asp?StationID="+StationNum);
	}else{
		Layer5.innerHTML=" ";
		TDStationErrorLog=1;
	}
}
//舉發單位(ajax)
function getUnit(){
	myForm.BillUnitID.value=myForm.BillUnitID.value.toUpperCase();
	if (event.keyCode==<%
if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then
		response.write "116"
else
		response.write "117"
end if
	%>){	
		event.keyCode=0;
		event.returnValue=false;
		window.open("Query_Unit.asp?SType=U","WebPage_Station12","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillUnitID.value.length > 1){
		var BillUnitNum=myForm.BillUnitID.value;
		runServerScript("getBillUnitID.asp?BillUnitID="+BillUnitNum);
	}else{
		Layer6.innerHTML=" ";
		TDUnitErrorLog=1;
	}
}

function UserInputBillType(){

}
//逕舉不一定要輸入固定桿編號. 除了是下方選擇使用固定桿
function getFixID(){
	if (myForm.UseTool.value.length == "1"){
		if (myForm.UseTool.value != "1" && myForm.UseTool.value != "2" && myForm.UseTool.value != "3"){
			alert("採証工具填寫錯誤!");
			myForm.UseTool.focus();
			myForm.UseTool.value = "";
		}else if (myForm.UseTool.value == "1"){
			//Layer11.style.visibility = "visible"; 
		}else{
			//Layer11.style.visibility = "hidden"; 
		}
	}
}
//違規地點代碼(ajax)
function getillStreet(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91) || event.keyCode==<%
	if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then 
		response.write "116"
	else
		response.write "117"
	end if 
		%>){
		myForm.IllegalAddressID.value=myForm.IllegalAddressID.value.toUpperCase();
		if (event.keyCode==<%
	if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then 
		response.write "116"
	else
		response.write "117"
	end if 
		%>){	
			event.keyCode=0;
			event.returnValue=false;
			OstreetID=myForm.IllegalAddressID.value;
			window.open("Query_Street.asp?OstreetID="+OstreetID,"WebPage_Street_People2","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
		}
		if (myForm.IllegalAddressID.value.length > 2){
			var illAddrNum=myForm.IllegalAddressID.value;
			runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
		}
	
		if (myForm.IllegalAddressID.value.length == 6){
		<%if sys_City="苗栗縣" then %>
			myForm.IllegalAddress.select();
		<%else%>
			myForm.Rule1.select();
		<%end if%>
		}
	}
}
//舉發人一(ajax)
function getBillMemID1(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMem1.value=myForm.BillMem1.value.toUpperCase();
	}
	if (event.keyCode==<%
if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then
		response.write "116"
else
		response.write "117"
end if
	%>){	
		event.keyCode=0;
		event.returnValue=false;
		window.open("Query_MemID.asp?MemOrder=1","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem1.value.length > 2){
		var BillMemNum=myForm.BillMem1.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=1&MemID="+BillMemNum);
	}else if (myForm.BillMem1.value.length <= 2 && myForm.BillMem1.value.length > 0){
		Layer12.innerHTML=" ";
		myForm.BillMemID1.value="";
		myForm.BillMemName1.value="";
		TDMemErrorLog1=1;
	}else{
		Layer12.innerHTML=" ";
		myForm.BillMemID1.value="";
		myForm.BillMemName1.value="";
		TDMemErrorLog1=0;
	}
}
//舉發人2(ajax)
function getBillMemID2(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMem2.value=myForm.BillMem2.value.toUpperCase();
	}
	if (event.keyCode==117){	
		event.keyCode=0;
		event.returnValue=false;
		window.open("Query_MemID.asp?MemOrder=2","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem2.value.length > 2){
		var BillMemNum=myForm.BillMem2.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=2&MemID="+BillMemNum);
	}else if (myForm.BillMem2.value.length <= 2 && myForm.BillMem2.value.length > 0){
		Layer13.innerHTML=" ";
		myForm.BillMemID2.value="";
		myForm.BillMemName2.value="";
		TDMemErrorLog2=1;
	}else{
		Layer13.innerHTML=" ";
		myForm.BillMemID2.value="";
		myForm.BillMemName2.value="";
		TDMemErrorLog2=0;
	}
}
//舉發人3(ajax)
function getBillMemID3(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMem3.value=myForm.BillMem3.value.toUpperCase();
	}
	if (event.keyCode==117){	
		event.keyCode=0;
		event.returnValue=false;
		window.open("Query_MemID.asp?MemOrder=3","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem3.value.length > 2){
		var BillMemNum=myForm.BillMem3.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=3&MemID="+BillMemNum);
	}else if (myForm.BillMem3.value.length <= 2 && myForm.BillMem3.value.length > 0){
		Layer14.innerHTML=" ";
		myForm.BillMemID3.value="";
		myForm.BillMemName3.value="";
		TDMemErrorLog3=1;
	}else{
		Layer14.innerHTML=" ";
		myForm.BillMemID3.value="";
		myForm.BillMemName3.value="";
		TDMemErrorLog3=0;
	}
}
//舉發人4(ajax)
function getBillMemID4(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91)){
		myForm.BillMem4.value=myForm.BillMem4.value.toUpperCase();
	}
	if (event.keyCode==117){	
		event.keyCode=0;
		event.returnValue=false;
		window.open("Query_MemID.asp?MemOrder=4","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem4.value.length > 2){
		var BillMemNum=myForm.BillMem4.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=4&MemID="+BillMemNum);
	}else if (myForm.BillMem4.value.length <= 2 && myForm.BillMem4.value.length > 0){
		Layer17.innerHTML=" ";
		myForm.BillMemID4.value="";
		myForm.BillMemName4.value="";
		TDMemErrorLog4=1;
	}else{
		Layer17.innerHTML=" ";
		myForm.BillMemID4.value="";
		myForm.BillMemName4.value="";
		TDMemErrorLog4=0;
	}
}
function getBillFillDate(){
	myForm.IllegalDate.value=myForm.IllegalDate.value.replace(/[^\d]/g,'');
	if(eval(TodayDate) < eval(myForm.IllegalDate.value)){
		alert("違規日期不得大於今天!!");
		myForm.IllegalDate.select();
	}

//	if (myForm.IllegalDate.value.length >= 6 ){
//		myForm.BillFillDate.value=myForm.IllegalDate.value;
//		getDealLineDate();
//	}
}
//逕舉由填單日期帶入應到案日期
function getDealLineDate(){
	getDealDateValue=<%=getReportDealDateValue%>;	//要加幾天
	myForm.BillFillDate.value=myForm.BillFillDate.value.replace(/[^\d]/g,'');
	BFillDateTemp=myForm.BillFillDate.value;
	if (BFillDateTemp.length >= 6 && myForm.BillType.value=="2"){
		Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
		Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
		Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
		var BFillDate=new Date(Byear,Bmonth-1,Bday);
		var DLineDate=new Date()
		DLineDate=DateAdd("d",getDealDateValue,BFillDate);
		Dyear=parseInt(DLineDate.getYear())-1911;
		Dmonth=parseInt(DLineDate.getMonth())+1;
		Dday=DLineDate.getDate();
		Dyear=Dyear.toString();
		if (Dmonth < 10){
			Dmonth="0"+Dmonth;
		}
		if (Dday < 10){
			Dday="0"+Dday;
		}
		myForm.DealLineDate.value=Dyear+Dmonth+Dday;
	}
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
	return win;
}
//用固定桿編號抓出違規地點
function setFixEquip(){
	if (myForm.FixID.value.length > 2){
		var FixNum=myForm.FixID.value;
		runServerScript("getFixIDAddress.asp?FixNum="+FixNum);
	}
}
function RuleSpeedforLaw(){
	//myForm.RuleSpeed.value=myForm.RuleSpeed.value.replace(/[^\d]/g,'');
	CallChkLaw1();
	CallChkLaw2();
	var IntError=0;
	var StrError="";
	if (myForm.IllegalSpeed.value > <%
	if sys_City="雲林縣" or sys_City="高雄市" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>){
		IntError=IntError+1;
		StrError=StrError+"\n"+IntError+"：車速、車重超過<%
	if sys_City="雲林縣" or sys_City="高雄市" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>。";
	}
	if((myForm.Rule1.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= <%
	if sys_City="雲林縣" then 
		response.write "41"
	else
		response.write "41"
	end if
	%>){
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：車速超過限速<%
	if sys_City="雲林縣" then 
		response.write "40"
	else
		response.write "40"
	end if
	%>公里以上。";
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：超過最高限速40公里以上需另單舉發法條4340068(處車主)!!\n(112/6/30前案件須超過60公里以上另單舉發法條4340044)";
			}
		}
	}
	if (IntError!=0){
		alert(StrError+"\n\n請確認是否正確!");
	}
<%if sys_City="高雄市" then%>
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
		if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) > 100 && (myForm.IllegalSpeed.value - myForm.RuleSpeed.value) < 150)
		{
			SpeedError=1;
			window.open("../BillKeyIn/ChkSpeedPW.asp","ChkSpeedPW","left=300,top=100,width=350,height=200,resizable=yes,scrollbars=no");
		}else{
			SpeedError=0;
		}
	}
<%end if%>
	<%if UpdateIllegalRuleFlag=1 then		'是否用車速判斷超速法條
	%>
	setIllegalRule();
	<%end if%>
}
function funGetSpeedRule(){
	<%if UpdateIllegalRuleFlag=1 then		'是否用車速判斷超速法條
	%>
	setIllegalRule();
	<%end if%>
}
function IllegalSpeedforLaw(){
	myForm.IllegalSpeed.value=myForm.IllegalSpeed.value.replace(/^[^\d]+|[^\d.]|,+$/g,'');
	<%if not rs1.eof then%>
		<%'if trim(rs1("ProsecutionTypeID"))<>"R" then%>
		CallChkLaw1();
		<%'end if%>
		CallChkLaw2();
	<%end if%>
	var IntError=0;
	var StrError="";
	if (myForm.IllegalSpeed.value > <%
	if sys_City="雲林縣" or sys_City="高雄市" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>){
		IntError=IntError+1;
		StrError=StrError+"\n"+IntError+"：車速、車重超過<%
	if sys_City="雲林縣" or sys_City="高雄市" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>。";
	}
	if((myForm.Rule1.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= <%
	if sys_City="雲林縣" then 
		response.write "41"
	else
		response.write "41"
	end if
	%>){
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：車速超過限速<%
	if sys_City="雲林縣" then 
		response.write "40"
	else
		response.write "40"
	end if
	%>公里以上。";
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：超過最高限速40公里以上需另單舉發法條4340068(處車主)!!\n(112/6/30前案件須超過60公里以上另單舉發法條4340044)";
			}
		}
	}
	if (IntError!=0){
		alert(StrError+"\n\n請確認是否正確!");
	}
<%if sys_City="高雄市" then%>
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
		if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) > 100 && (myForm.IllegalSpeed.value - myForm.RuleSpeed.value) < 150)
		{
			SpeedError=1;
			window.open("../BillKeyIn/ChkSpeedPW.asp","ChkSpeedPW","left=300,top=100,width=350,height=200,resizable=yes,scrollbars=no");
		}else{
			SpeedError=0;
		}
	}
<%end if%>
	<%if UpdateIllegalRuleFlag=1 then		'是否用車速判斷超速法條
	%>
	setIllegalRule("NoSelect");
	<%end if%>
}
function CallChkLaw1(){
	if (!funcChkLaw(myForm.Rule1.value) && !funcChkLaw(myForm.Rule2.value) && myForm.Rule1.value!="" && myForm.Rule2.value!=""){
		if (!funcChkLaw(myForm.Rule1.value)){
			alert("請確認法條一是否填寫正確");
		}
	}else if (!funcChkLaw(myForm.Rule1.value) && myForm.Rule2.value==""){
		alert("請確認法條一是否填寫正確");
	}
}
function CallChkLaw2(){
	if (!funcChkLaw(myForm.Rule1.value) && !funcChkLaw(myForm.Rule2.value) && myForm.Rule1.value!="" && myForm.Rule2.value!=""){
		if (!funcChkLaw(myForm.Rule2.value)){
			alert("請確認法條二是否填寫正確");
		}
	}else if (!funcChkLaw(myForm.Rule2.value) && myForm.Rule1.value==""){
		alert("請確認法條二是否填寫正確");
	}
}
//法律條文建檔檢查
function funcChkLaw(thisLaw){
	if (thisLaw.length>=2){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			//當有打速限及車速時 法條一定落在33XXXX,40XXXX,43XXXX
			if ((thisLaw.substr(0,2))!="33" && (thisLaw.substr(0,2))!="40" && (thisLaw.substr(0,2))!="43" && (thisLaw.substr(0,2))!="29"){
				return false;
			}else{
				//違規地點含有"快速道路"判斷法條是否選33XXX而非選40XXX
				if ((myForm.IllegalAddress.value.indexOf("快速道路",0)) != -1){
					if ((thisLaw.substr(0,2))=="40"){
						return false;
					}else{
						return true;
					}
				}else{
					return true;
				}
			}
		}else{
			return true;
		}
	}else{
		return true;
	}
}


//審核無效
function funVerifyResult(){
//	if(confirm('確定要將此筆檢舉案件設為無效？')){
//		myForm.kinds.value="VerifyResultNull";
//		myForm.submit();
//	}
	UrlStr="../ReportCase/ReportCase_Verify.asp?CheckType=0&CheckSn=<%=trim(request("CheckSn"))%>&ReportCaseSn=<%=trim(rs1("ReportSn"))%>";
	newWin(UrlStr,"ReportCase_Verify",800,450,0,0,"yes","yes","yes","no");
}


function KeyDown(){ 
	if (event.keyCode==<%
if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then 
		response.write "117"
else
		response.write "116"
end if 
	%>){	//F5查詢
		event.keyCode=0;   
		event.returnValue=false;   
<%if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then %>
	}else if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
<%end if %>
	}else if (event.keyCode==113){ //F2存檔
		event.keyCode=0;   
<%
	if not rs1.eof then
		if trim(rs1("CheckFlag"))="0" then
%>
		InsertBillVase();
<%
		end if 
	end if
%>
	}else if (event.keyCode==115){ //F4清除
		event.keyCode=0;   
		event.returnValue=false;  
	//}else if (event.keyCode==117){ //F6查詢
	//	event.keyCode=0;   
	//	event.returnValue=false;  
	//	funcOpenBillQry();
	}else if (event.keyCode==119){ //F8關閉
		event.keyCode=0;   
		event.returnValue=false;  
		window.close();
	}else if (event.keyCode==120){ //F9審核無效
		event.keyCode=0;   
		event.returnValue=false;  
<%
	if not rs1.eof then
		if trim(rs1("CheckFlag"))="0" then
%>
		funVerifyResult();
<%		end if 
	end if
%>
	}else if (event.keyCode==121){ //F10查詢未建檔
		event.keyCode=0;   
		event.returnValue=false;  

	}else if (event.keyCode==122){ //F11略過
		event.keyCode=0;   
		event.returnValue=false;  

	}else if (event.keyCode==33){ //上一筆PageUp
		event.keyCode=0;   
		event.returnValue=false; 
	<%if UpSn<>"" then%>
		location='BillKeyIn_Image_ReportCase_Check.asp?CheckSn=<%=UpSn%>'
	<%end if %>
	}else if (event.keyCode==36){ //第一筆Home
		event.keyCode=0;   
		event.returnValue=false; 
	<%if UpSn<>"" then%>
		location='BillKeyIn_Image_ReportCase_Check.asp?CheckSn=<%=FirstSn%>'
	<%end if %>
	}else if (event.keyCode==34){ //下一筆PageDn
		event.keyCode=0;   
		event.returnValue=false; 
	<%if DownSn<>"" then%>
		location='BillKeyIn_Image_ReportCase_Check.asp?CheckSn=<%=DownSn%>'
	<%end if %>
	}else if (event.keyCode==35){ //最後一筆End
		event.keyCode=0;   
		event.returnValue=false; 
	<%if DownSn<>"" then%>
		location='BillKeyIn_Image_ReportCase_Check.asp?CheckSn=<%=LastSn%>'
	<%end if %>
	}
}

function AutoGetIllStreet(){	//按F6可以直接顯示相關路段
	if (event.keyCode==<%
if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then 
		response.write "116"
else
		response.write "117"
end if 
	%>){	
		event.keyCode=0;
		event.returnValue=false;
		Ostreet=myForm.IllegalAddress.value;
		window.open("Query_Street.asp?OStreet="+Ostreet,"WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
}
function AutoGetRuleID(LawOrder){	//按F6可以直接顯示相關法條
	//if (event.keyCode==117){	
//		event.keyCode=0;
		if (LawOrder==1){
			ORuleID=myForm.Rule1.value;
		}else{
			ORuleID=myForm.Rule2.value;
		}
		window.open("Query_Law.asp?LawOrder="+LawOrder+"&RuleVer=<%=theRuleVer%>&ORuleID="+ORuleID,"WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	//}
}
function ProjectF5(){
	if (event.keyCode==<%
if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then
		response.write "116"
else
		response.write "117"
end if
	%>){	
		event.keyCode=0;
		event.returnValue=false;
		window.open("Query_Project.asp","WebPage_Street_People","left=0,top=0,location=0,width=800,height=460,resizable=yes,scrollbars=yes");
	}
	if (myForm.ProjectID.value.length > 0){
		var BillProjectID=myForm.ProjectID.value;
		runServerScript("getProjectID.asp?BillProjectID="+BillProjectID);
<%if sys_City="苗栗縣" then%>
		if (myForm.ProjectID.value=="9"){
			myForm.CarAddID.value="8";
		}
<%end if%>
	}else{
		Layer001.innerHTML="";
		TDProjectIDErrorLog=0;
	}
}
//用地點、車速抓違規法條
function setIllegalRule(flag){
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!="" && myForm.IllegalAddress.value!=""){
	<%if not rs1.eof then%>
		if ((myForm.Rule1.value.substr(0,2))!="29"){
			IllegalRule=getIllegalRule(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
			if (IllegalRule!="Null"){
				myForm.Rule1.value=IllegalRule;
				getRuleData1(flag);
			}
		}
		if ((myForm.Rule2.value.substr(0,2))!="29" && ((myForm.Rule1.value.substr(0,5))!="33101" && (myForm.Rule1.value.substr(0,2))!="40" && (myForm.Rule1.value.substr(0,5))!="43102")){
			IllegalRule2=getIllegalRule(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
			if (IllegalRule2!="Null"){
				myForm.Rule2.value=IllegalRule2;
				getRuleData2();
			}
		}
	<%end if%>
	}else{
//		if ((myForm.Rule1.value.substr(0,2))!="29" && ProsecutionTypeID=="R"){
//			IllegalRule=getIllegalRule(myForm.IllegalAddress.value,"0","0",ProsecutionTypeID,myForm.chkHighRoad.checked);
//			if (IllegalRule!="Null"){
//				myForm.Rule1.value=IllegalRule;
//				getRuleData1();
//			}
//		}
	}
}
//附加說明
function Add_LawPlus(){
	if (myForm.Rule1.value==""){
		alert("請先輸入違規法條一!!");
	}else{
	RuleID=myForm.Rule1.value;
	window.open("Query_LawPlus.asp?RuleID="+RuleID+"&theRuleVer=<%=theRuleVer%>","WebPage1","left=20,top=10,location=0,width=500,height=455,resizable=yes,scrollbars=yes");
	}
}
function changeStreet(){
	//if (myForm.getStreetName.value!=""){
		myForm.kinds.value="getStreet";
		myForm.submit();
	//}
}
<%'if sys_City="高雄市" then%>
var sys_City="<%=sys_City%>";
function QryIllegalZip(){
	window.open("Query_Zip.asp?ZipCity="+sys_City+"&IllegalZip="+myForm.IllegalZip.value+"&ObjName=IllegalZip","WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes,status=yes");

}
function getIllZip(){
	runServerScript("getZipNameForCar.asp?ZipID="+myForm.IllegalZip.value);
}
<%'end if %>
function funcUpdSaveLocation(){
		myForm.kinds.value="";
		myForm.submit();
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

	//-----------上下左右-------------
	function funTextControl(obj){
		if (event.keyCode==13){ //Enter換欄
			event.keyCode=0;
			event.returnValue=false;
			
			//if (obj==myForm.CarNo && myForm.CarNo.value!=""){
				//myForm.IllegalDate.select();
			//}else{
				CodeEnter(obj.name);
			//}
		}else if (event.keyCode==38){ //上換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeMoveLeft(obj.name);
		}else if (event.keyCode==40){ //下換欄
			event.keyCode=0;
			event.returnValue=false;
			
			//if (obj==myForm.CarNo && myForm.CarNo.value!=""){
			//	myForm.IllegalDate.select();
			//}else{
				CodeMoveRight(obj.name);
			//}
		}else if (event.keyCode==<%
if sys_City="南投縣" or sys_City="屏東縣" or sys_City="嘉義縣" then 
		response.write "116"
else
		response.write "117"
end if 
	%>){ 
			event.keyCode=0;
			event.returnValue=false;
			if (obj==myForm.Rule1){
				AutoGetRuleID(1);
			}else if (obj==myForm.Rule2){
				AutoGetRuleID(2);
			}
		}else if (event.keyCode==9){ //tab
			event.keyCode=0;
			event.returnValue=false;
			
			if (obj==myForm.CarNo && myForm.CarNo.value!=""){
				myForm.IllegalDate.select();
			}else{
				CodeEnter(obj.name);
			}
		}
	}
	//------------------------------

function IllegalDateKeyUP(){
	//打數字才會跳下攔
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106)){
		if (myForm.IllegalDate.value.substr(0,1)=="1"){
			if (myForm.IllegalDate.value.length=="7"){
				myForm.IllegalTime.select();
			}
		}else{
			if (myForm.IllegalDate.value.length=="6"){
				myForm.IllegalTime.select();
			}
		}
	}
}

function IllegalTimeKeyUP(){
	//打數字才會跳下攔
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106)){
<%if sys_City="苗栗縣" then%>
		if (myForm.IllegalTime.value.length=="4"){
			myForm.IllegalSpeed.select();
		}
<%else%>
		if (myForm.IllegalTime.value.length=="4"){
			if (myForm.IllegalAddressID.value==""){
				myForm.IllegalAddressID.select();
			}else if (myForm.IllegalAddress.value==""){
				myForm.IllegalAddress.select();
			}else{
				myForm.Rule1.select();
			}
		}
<%end if %>
	}
}

//開啟檢視圖
function OpenPic(FileName){
//alert(FileName);
	window.open("../Query/ShowIllegalImage.asp?FileName="+FileName,"UploadFile","left=0,top=0,location=0,width=910,height=705,resizable=yes,status=yes,scrollbars=yes,menubar=yes");
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


//=====放大鏡=======================================
var iDivHeight = <%
	'If sys_City=ApconfigureCityName Then
		response.write "110"
	'Else
	'	response.write "90"
	'End If 
			%>; //放大?示?域?度
var iDivWidth = <%
	'If sys_City=ApconfigureCityName Then
		response.write "230"
	'Else
	'	response.write "210"
	'End If 
			%>;//放大?示?域高度
var iMultiple = 4; //放大倍?

//?示放大?，鼠?移?事件和鼠???事件都??用本事件
//??：src代表?略?，sFileName放大?片名?
//原理：依据鼠????略?左上角（0，0）上的位置控制放大?左上角???示?域左上角（0，0）的位置
function show(src, sFileName)
{
//判?鼠?事件?生?是否同?按下了
if ((event.button == 1) && (event.ctrlKey == true)){
  iMultiple -= 1;
  //myForm.CarNo.focus();
}else
  if (event.button == 1){
  iMultiple += 1;
   //myForm.CarNo.focus();
  }
if (iMultiple < 3) iMultiple = 3;

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
var ImageFileNameTemp;
function ChangeImgB(){
<%if sPicWebPath2<>"" then%>
	oBigImg=myForm.imgSource.src;
	oSmallImg=myForm.SmallImgB.src;

	myForm.SmallImgB.src=oBigImg;
	myForm.imgSource.src=oSmallImg;
	myForm.BigImg.src=oSmallImg;

	ImageFileNameTemp=myForm.ImageFileNameA.value;
	myForm.ImageFileNameA.value=myForm.ImageFileNameB.value;
	myForm.ImageFileNameB.value=ImageFileNameTemp;
<%end if%>
}

function ChangeImgC(){
<%if sPicWebPath<>"" then%>
	oBigImg=myForm.imgSource.src;
	oSmallImg=myForm.SmallImgC.src;

	myForm.SmallImgC.src=oBigImg;
	myForm.imgSource.src=oSmallImg;
	myForm.BigImg.src=oSmallImg;

	ImageFileNameTemp=myForm.ImageFileNameA.value;
	myForm.ImageFileNameA.value=myForm.ImageFileNameC.value;
	myForm.ImageFileNameC.value=ImageFileNameTemp;
<%end if%>
}

function ChangeImgD(){
<%if sPicWebPath3<>"" then%>
	oBigImg=myForm.imgSource.src;
	oSmallImg=myForm.SmallImgD.src;

	myForm.SmallImgD.src=oBigImg;
	myForm.imgSource.src=oSmallImg;
	myForm.BigImg.src=oSmallImg;

	ImageFileNameTemp=myForm.ImageFileNameA.value;
	myForm.ImageFileNameA.value=myForm.ImageFileNameD.value;
	myForm.ImageFileNameD.value=ImageFileNameTemp;
<%end if%>
}

function setImageNotUse(ImgID){
<%if bPicWebPath<>"" then%>
	if (ImgID=="A")
	{
		if (myForm.chkImgNoUseA.value=="-1")
		{
			myForm.chkImgNoUseA.value="1";
			myForm.btnImgNoUseA.style.backgroundColor='';
			
		}else{
			myForm.chkImgNoUseA.value="-1";
			myForm.btnImgNoUseA.style.backgroundColor='red';
		}		
	}
<%end if %>
<%if sPicWebPath2<>"" then%>
	if (ImgID=="B")
	{
		if (myForm.chkImgNoUseB.value=="-1")
		{
			myForm.chkImgNoUseB.value="1";
			myForm.btnImgNoUseB.style.backgroundColor='';
			
		}else{
			myForm.chkImgNoUseB.value="-1";
			myForm.btnImgNoUseB.style.backgroundColor='red';
		}		
	}
<%end if %>
<%if sPicWebPath<>"" then%>
	if (ImgID=="C")
	{
		if (myForm.chkImgNoUseC.value=="-1")
		{
			myForm.chkImgNoUseC.value="1";
			myForm.btnImgNoUseC.style.backgroundColor='';
			
		}else{
			myForm.chkImgNoUseC.value="-1";
			myForm.btnImgNoUseC.style.backgroundColor='red';
		}		
	}
<%end if %>
<%if sPicWebPath3<>"" then%>
	if (ImgID=="D")
	{
		if (myForm.chkImgNoUseD.value=="-1")
		{
			myForm.chkImgNoUseD.value="1";
			myForm.btnImgNoUseD.style.backgroundColor='';
			
		}else{
			myForm.chkImgNoUseD.value="-1";
			myForm.btnImgNoUseD.style.backgroundColor='red';
		}		
	}
<%end if %>
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
//------------------------------------------------
<%
if not rs1.eof then
%>
//myForm.CarNo.select();
getBillFillDate();
getDealLineDate();
setIllegalRule();
<%
	if trim(rs1("CarSimpleID"))="" or isnull(rs1("CarSimpleID")) or trim(rs1("CarSimpleID"))="0" then
		if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
%>
<%if sys_City<>"高雄市" then%>
		CarType=chkCarNoFormat(myForm.CarNo.value);
		if (CarType!=0){
			myForm.CarSimpleID.value=CarType;
		}
<%end if%>
		
<%
		end if
	end if
end if
rs1.close
set rs1=nothing
conn.close
set conn=nothing
%>
</script>

</html>
