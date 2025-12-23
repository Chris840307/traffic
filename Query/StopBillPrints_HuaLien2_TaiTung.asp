<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/CreateChkCode.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單列印-Legal Size</title>
<style type="text/css">
<!--
.style1 {font-family:"標楷體"; font-size: 20px; line-height:2;}
.style2 {font-family:"標楷體"; font-size: 16px;}
.style3 {font-family:"標楷體"; font-size: 14px;}
.style4 {font-family:"標楷體"; font-size: 22px; line-height:2;}
.style5 {font-family:"標楷體"; font-size: 13px;}
.style6 {font-family:"標楷體"; font-size: 22px; line-height:1;}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>
<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<%
response.Buffer=true
BillNo="":CarNo=""

strSQL="select distinct a.ImageFileNameB,a.CarNo from (select sn,carno,ImageFileNameB from BillBase where ImagePathName is not null and BillStatus>1 and RecordStateId <> -1 and ImageFileNameB is not null and DeallineDate is not null) a,(Select * from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b where a.SN=b.BillSN "&request("SQLstr")&" order by a.CarNo"
set rsbill=conn.execute(strSQL)
while Not rsbill.eof
	If trim(BillNo)<>"" Then
		BillNo=BillNo&","
		CarNo=CarNo&","
	end if
	BillNo=BillNo&trim(rsbill("ImageFileNameB"))
	CarNo=CarNo&trim(rsbill("CarNo"))
	rsbill.movenext
wend
rsbill.close

PBillNo=split(trim(BillNo),",")
PCarNo=split(trim(CarNo),",")
Server.ScriptTimeout=6000
PageCnt=0
tmpdate=split(gArrDT(trim(date)),"-")
SysDate=tmpdate(1)&tmpdate(2)
'on Error Resume Next
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
PrintCnt=0
toppx=-5

for cmtI=0 to Ubound(PBillNo)
	if cmtI<>0 then response.write "<div class=""PageNext""></div>"
	Sys_Address="":strOwnerZip="":Sys_OwnerZip=""
		
	strSQL="select CarNo,DeCode(OwnerNotifyAddress,null,DeCode(DriverHomeAddress,null,OwnerAddress,DriverHomeAddress),OwnerNotifyAddress) OwnerAddress,DeCode(OwnerNotifyAddress,null,DeCode(DriverHomeZip,null,OwnerZip,DriverHomeZip),null) OwnerZip,Owner from BillbaseDCIReturn where CarNo='"&trim(PCarNo(cmtI))&"' and ExchangetypeID='A'"

	set rsDci=conn.execute(strSQL)
	if Not rsDci.eof then
		Sys_CarNo="":Sys_Owner="":Sys_OwnerAddress="":Sys_OwnerZip=""

		Sys_CarNo=trim(rsDci("CarNo"))
		Sys_Owner=trim(rsDci("Owner"))
		Sys_OwnerAddress=trim(rsDci("OwnerAddress"))
		Sys_OwnerZip=trim(rsDci("OwnerZip"))

		strSQL="select distinct Owner,OwnerZip,OwnerAddress from billbase where ImageFileNameB='"&trim(PBillNo(cmtI))&"'"

		set rsbill=conn.execute(strSQL)

		If not ifnull(rsbill("OwnerAddress")) and not ifnull(rsbill("Owner")) Then
			Sys_Owner=trim(rsbill("Owner"))
			Sys_OwnerZip=trim(rsbill("OwnerZip"))
			Sys_OwnerAddress=trim(rsbill("OwnerAddress"))
		else
			strSQL="update billbase set Owner='"&Sys_Owner&"',OwnerAddress='"&Sys_OwnerAddress&"',OwnerZip='"&Sys_OwnerZip&"' where ImageFileNameB='"&trim(PBillNo(cmtI))&"'"

			conn.execute(strSQL)
		End if
		rsbill.close

		If not ifnull(Sys_OwnerZip) Then
			strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
			set rszip=conn.execute(strSQL)
			if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
			rszip.close
		else
			Sys_OwnerZipName=""
		End if

		Sys_Address=Sys_OwnerZip&Sys_OwnerAddress
				
	end if
	rsDci.close

	Sys_MailNumber="":Sys_MailChkNumber=""

	strSQL="select distinct MailNumber from StopBillMailHistory where BillNo='"&PBillNo(cmtI)&"'"
	set rsmail=conn.execute(strSQL)
	If Not rsmail.eof Then
		Sys_MailNumber=trim(rsmail("MailNumber"))&"95000017"
		Sys_MailChkNumber=trim(rsmail("MailNumber"))&" 950000 17"
	end if
	rsmail.close

	strSQL="select distinct CarNo,BillUnitID,DeallIneDate,ImageFileNameB from BillBase where ImageFileNameB='"&PBillNo(cmtI)&"'"
	set rsbill=conn.execute(strSQL)
	If Not rsbill.eof Then
		Sys_CarNo=trim(rsbill("CarNo"))
		Sys_BillUnitID=trim(rsbill("BillUnitID"))
		Sys_DeallIneDate=split(gArrDT(trim(rsbill("DeallIneDate"))),"-")
		Sys_ImageFileNameB=trim(rsbill("ImageFileNameB"))
	End if
	rsbill.close
%>
<div id="L78" style="position:relative;">

<div id="Layer1" class="style4" style="position:absolute; left:60px; top:<%=110+toppx%>px; z-index:1"><%=funcCheckFont(Sys_Address,20,1)%></div>

<div id="Layer2" class="style4" style="position:absolute; left:60px; top:<%=190+toppx%>px; z-index:1"><%="收件人："&funcCheckFont(Sys_Owner,20,1)%></div>

<div id="Layer18" class="style4" style="position:absolute; font-size:14px; left:500px; top:<%=170+toppx%>px; z-index:1"><%
	DelphiASPObj.CreateBarCode Sys_MailNumber,128,25,260
	response.write "<img src=""../BarCodeImage/"&Sys_MailNumber&".jpg"" width=""220"" height=""90""><br>　　"&Sys_MailChkNumber
%></div>
<!--
<div id="Layer3" class="style4" style="position:absolute; left:600px; top:<%=200+toppx%>px; z-index:1">台啟</div>
-->
<div id="Layer4" class="style3" style="position:absolute; left:200px; top:<%=250+toppx%>px; z-index:1"><%=Sys_ImageFileNameB%></div>

<div id="Layer5" class="style2" style="position:absolute; left:50px; top:<%=380+toppx%>px; z-index:1"><%="車號："&Sys_CarNo&"　車主地址："&funcCheckFont(Sys_Address,20,1)%></div>

<div id="Layer6" style="position:absolute; left:50px; top:<%=400+toppx%>px; z-index:1">
	<table border="0" width="80%" cellspacing=0 cellpadding=0>
		<tr>
			<td class="style3">日期</td>
			<td class="style3">時間</td>
			<td class="style3">地點</td>
			<td class="style3">繳費通知單</td>
			<td class="style3">停車費</td>
		</tr><%
			ForFeitSum=0:Cntsum=0
			strSQL="select ForFeit1,IllegalDate,IllegalAddress,ImagePathName from BillBase where ImageFileNameB='"&Sys_ImageFileNameB&"' order by IllegalDate"
			set rst=conn.execute(strSQL)
			While Not rst.eof
				Sys_IllegalDate=split(gArrDT(rst("IllegalDate")),"-")
				response.write "<tr>"
				response.write "<td class=""style3"">"&Sys_IllegalDate(0)&"/"&Sys_IllegalDate(1)&"/"&Sys_IllegalDate(2)&"</td>"
				response.write "<td class=""style3"">"&hour(rst("IllegalDate"))&":"&minute(rst("IllegalDate"))&"</td>"
				response.write "<td class=""style3"">"&trim(rst("IllegalAddress"))&"</td>"
				response.write "<td class=""style3"">"&trim(rst("ImagePathName"))&"</td>"
				response.write "<td class=""style3"">"&trim(rst("ForFeit1"))&"</td>"
				response.write "</tr>"
				ForFeitSum=ForFeitSum+rst("ForFeit1")
				Cntsum=Cntsum+1
				rst.movenext
			Wend
			rst.close
		%>
	</table>
</div>
<div id="Layer7" class="style4" style="position:absolute; left:50px; top:<%=545+toppx%>px; z-index:1"><B><%
	response.write "催繳單號："&Sys_ImageFileNameB&"　　　繳費期限"&Sys_DeallIneDate(0)&"/"&Sys_DeallIneDate(1)&"/"&Sys_DeallIneDate(2)&"日止"
%></B>
</div>
<div id="Layer8" class="style4" style="position:absolute; left:50px; top:<%=570+toppx%>px; z-index:1"><B><%
	response.write "共計催繳："&Cntsum&"筆，停車費："&ForFeitSum&"元、工本費34元，總金額："&(ForFeitSum+34)&"元"
%></B>
</div>

<div id="Layer9" class="style5" style="position:absolute; left:50px; top:<%=605+toppx%>px; z-index:1">
請於收到本催繳通知單繳費期限內繳納停車費，逾期仍未繳納，依違反道路交通管理處罰條例第56條3項舉發。<br>
繳費方式：請持本催繳通知單至統一7-11、全家便利超商全省各門市繳納或受委託台灣國際開發事業有限公司<br>
(臺東市新生路191號旁公共造產停五停車場，停車收費時間繳納)，並保留本收據聯6個月，如已繳交停車欠費<br>
請勿重覆繳費，以維護您的權益。<br>
<span class="style6"><b>
	台端查詢停車紀錄及任何疑問：請電(089)349867</b>
</span><br>
或網址：臺東縣政府 www.taitung.gov.tw上網查詢。
</div>

<div id="Layer10" class="style2" style="position:absolute; left:80px; top:<%=775+toppx%>px; z-index:1"><%
	response.write "繳費期限及單位代碼："&Sys_DeallIneDate(0)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&"29A"
	response.write "　　&nbsp;&nbsp;"
	DelphiASPObj.CreateBarCode right(Sys_DealLineDate(0),2)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&"29A",0,35,160
	response.write "<img src=""../BarCodeImage/"&right(Sys_DealLineDate(0),2)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&"29A.jpg"">"
'	response.write "<span class=""style6"">*adfsdfsd*</span>"
'	response.write "<script>haiwaocde """&Sys_DeallIneDate(0)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&"241"",""popobj1""</script>"
'	response.write haiwaocde(Sys_DeallIneDate(0)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&"241")
%>
</div>
<div id="Layer11" class="style2" style="position:absolute; left:80px; top:<%=825+toppx%>px; z-index:1"><%
	response.write "催繳單號："&Sys_ImageFileNameB
	response.write "　　　　"
	DelphiASPObj.CreateBarCode Sys_ImageFileNameB,0,35,260
	response.write "<img src=""../BarCodeImage/"&Sys_ImageFileNameB&".jpg"">"
'	response.write "<span class=""style6"">"&Sys_ImageFileNameB&"*</span>"
'	response.write "<span id=""popobj2""></span>"
'	response.write "<script>haiwaocde """&Sys_ImageFileNameB&""",""popobj2""</script>"
'	response.write haiwaocde(Sys_ImageFileNameB)
%>
</div>
<div id="Layer12" class="style2" style="position:absolute; left:80px; top:<%=875+toppx%>px; z-index:1"><%
	tmpDeallIneDate=right(Sys_DealLineDate(0),2)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)
	tempChkMemony=right("0000000000"&(ForFeitSum+34),9)
	SysChkNumber=CreateChkCode(tmpDeallIneDate,"29A",Sys_ImageFileNameB,SysDate,tempChkMemony)
	show_Memony="("&SysDate&SysChkNumber&")"&tempChkMemony
	barCodeMemony=SysDate&SysChkNumber&tempChkMemony

	response.write "繳費金額："&show_Memony
	response.write "　　　&nbsp"
	DelphiASPObj.CreateBarCode barCodeMemony,0,35,260
	response.write "<img src=""../BarCodeImage/"&barCodeMemony&".jpg"">"
'	response.write "<span class=""style6"">*"&barCodeMemony&"*</span>"
'	response.write "<span id=""popobj3""></span>"
'	response.write "<script>haiwaocde """&barCodeMemony&""",""popobj3""</script>"
'	response.write haiwaocde(barCodeMemony)
%>
</div>

<div id="Layer13" class="style3" style="position:absolute; left:630px; top:<%=870+toppx%>px; z-index:1"><%
	response.write Sys_CarNo&"<br>"
	response.write "經收人蓋章："
%>
</div>

<div id="Layer14" class="style2" style="position:absolute; left:300px; top:<%=920+toppx%>px; z-index:1"><%="車號："&Sys_CarNo%></div>

<div id="Layer15" class="style3" style="position:absolute; width:200px; left:130px; top:<%=975+toppx%>px; z-index:1"><%
	response.write funcCheckFont(Sys_Owner&"　"&Sys_CarNo,18,1)&"<br>"
	response.write funcCheckFont(Sys_Address,18,1)
%>
</div>
<div id="Layer16" class="style3" style="position:absolute; left:130px; top:<%=1030+toppx%>px; z-index:1"><%
	'response.write Sys_ImageFileNameB
	response.write "<img src=""../BarCodeImage/"&Sys_MailNumber&".jpg"" width=""140"" height=""20""><br>　"
	response.write Sys_MailNumber
	
%>
</div>
<div id="Layer17" class="style3" style="position:absolute; left:555px; top:<%=1050+toppx%>px; z-index:1"><%
	response.write "<img src=""../BarCodeImage/"&Sys_ImageFileNameB&".jpg"" width=""180"" height=""40"">"
	Response.Write "<br>　"&Sys_ImageFileNameB
'	response.write "<img src=""../BarCodeImage/"&Sys_MailNumber&".jpg""><br>　　　　　"&Sys_MailChkNumber
'	response.write "<img src=""../BarCodeImage/"&Sys_ImageFileNameB&".jpg"">"
'	response.write "<span class=""style6"">"&Sys_ImageFileNameB&"</span>"
'	response.write "<span id=""popobj4""></span>"
'	response.write "<script>haiwaocde """&Sys_ImageFileNameB&""",""popobj4""</script>"
'	response.write haiwaocde(Sys_ImageFileNameB)
%>
</div>
<div id="Layer16" class="style3" style="position:absolute; left:335px; top:<%=1190+toppx%>px; z-index:1"><%
	'response.write "<img src=""../BarCodeImage/"&Sys_MailNumber&".jpg"" width=""140"" height=""30""><br>　"
	'response.write Sys_MailNumber
%>
</div>
</div>
<%
response.flush
next%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(true,5.08,5.08,5.08,5.08);
</script>