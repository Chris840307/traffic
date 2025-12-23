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

UserMarkDate1=gOutDT(request("Sys_SendMarkDate1"))&" 0:0:0"
UserMarkDate2=gOutDT(request("Sys_SendMarkDate2"))&" 23:59:59"

strSQL="select BillNo,CarNo from StopCarSendAddress where UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS')"
set rsbill=conn.execute(strSQL)
while Not rsbill.eof
	If trim(BillNo)<>"" Then
		BillNo=BillNo&","
		CarNo=CarNo&","
	end if
	BillNo=BillNo&trim(rsbill("BillNo"))
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

newCode="24A"
sys_Company="國雲公寓大廈管理維護股份有限公司"
sys_CompanyAddress="花蓮市中福路230號"
sys_CompanyTEL="(03)8312920"
sys_CompanyUrl="www.tidch.tw/p/"


If trim(Request("newCode"))="1" Then
	newCode="245"
	sys_Company="建鈺有限公司"
	sys_CompanyAddress="花蓮市中福路229號"
	sys_CompanyTEL="(03)8315477"
	sys_CompanyUrl="www.chien-yu.tw"

End if

for cmtI=0 to Ubound(PBillNo)
	if cmtI<>0 then response.write "<div class=""PageNext""></div>"
	Sys_Address="":Sys_OwnerZip=""

	strSQL="select distinct CarNo,Owner,DriverZip,DriverAddress from billbase where ImageFileNameB='"&trim(PBillNo(cmtI))&"'"
		
	set rsDci=conn.execute(strSQL)
	if Not rsDci.eof then
		Sys_CarNo="":Sys_Owner="":Sys_OwnerAddress="":Sys_OwnerZip=""
		Sys_CarNo=trim(rsDci("CarNo"))
		Sys_Owner=trim(rsDci("Owner"))
		Sys_OwnerAddress=trim(rsDci("DriverAddress"))
		Sys_OwnerZip=trim(rsDci("DriverZip"))

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

	strSQL="select distinct StoreAndSendMailNumber from StopBillMailHistory where BillNo='"&PBillNo(cmtI)&"'"
	set rsmail=conn.execute(strSQL)
	If Not rsmail.eof Then
		Sys_MailNumber=trim(rsmail("StoreAndSendMailNumber"))&"97300717"
		Sys_MailChkNumber=trim(rsmail("StoreAndSendMailNumber"))&"973007 17"
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

<div id="Layer1" class="style4" style="position:absolute; left:60px; top:110px; z-index:1"><%=funcCheckFont(Sys_Address,20,1)%></div>

<div id="Layer2" class="style4" style="position:absolute; left:60px; top:190px; z-index:1"><%="收件人："&funcCheckFont(Sys_Owner,20,1)%></div>

<div id="Layer18" class="style4" style="position:absolute; font-size:14px; left:200px; top:230px; z-index:1"><%
	DelphiASPObj.CreateBarCode Sys_MailNumber,128,25,260
	response.write "<img src=""../BarCodeImage/"&Sys_MailNumber&".jpg""><br>　　　　　"&Sys_MailChkNumber
%></div>

<div id="Layer3" class="style4" style="position:absolute; left:600px; top:200px; z-index:1">台啟</div>

<div id="Layer4" class="style3" style="position:absolute; left:540px; top:240px; z-index:1"><%=Sys_ImageFileNameB%></div>

<div id="Layer5" class="style2" style="position:absolute; left:50px; top:380px; z-index:1"><%="車號："&Sys_CarNo&"　車主地址："&funcCheckFont(Sys_Address,20,1)%></div>

<div id="Layer6" style="position:absolute; left:50px; top:400px; z-index:1">
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
<div id="Layer7" class="style4" style="position:absolute; left:50px; top:560px; z-index:1"><B><%
	response.write "催繳單號："&Sys_ImageFileNameB&"　　　繳費期限"&Sys_DeallIneDate(0)&"/"&Sys_DeallIneDate(1)&"/"&Sys_DeallIneDate(2)&"日止"
%></B>
</div>
<div id="Layer8" class="style4" style="position:absolute; left:50px; top:585px; z-index:1"><B><%
	response.write "共計催繳："&Cntsum&"筆，停車費："&ForFeitSum&"元、工本費43元，總金額："&(ForFeitSum+43)&"元"
%></B>
</div>
<div id="Layer9" class="style5" style="position:absolute; left:50px; top:620px; z-index:1">
請於收到本催繳通知單繳費期限內繳納停車費，逾期仍未繳納，依違反道路交通管理處罰條例第56條2項舉發。<br>
繳費方式：請持本催繳通知單至「三大超商」統一7-11、萊爾富、OK&nbsp;便利超商全省各門市繳納或受委<br>
　　　　　託<%=Sys_Company%>(<%=Sys_CompanyAddress%>，10：00至18：00停車收費時間繳納)，<br>
　　　　　並保留本收據聯6個月，如已繳交停車欠費請勿重覆繳費，以維護您的權益。<br>
查詢停車紀錄：請電<%=Sys_CompanyTel%>或網址：www.hl.gov.tw或<%=Sys_CompanyUrl%>上網查詢；台端任何疑問請電<br>
　　　　　　　(03)8239164(花蓮縣警察局交通隊)洽詢。
</div>
<div id="Layer10" class="style2" style="position:absolute; left:80px; top:765px; z-index:1"><%
	response.write "繳費期限及單位代碼："&Sys_DeallIneDate(0)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&newCode
	response.write "　　&nbsp;&nbsp;"
	DelphiASPObj.CreateBarCode right(Sys_DealLineDate(0),2)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&newCode,0,35,160
	response.write "<img src=""../BarCodeImage/"&right(Sys_DealLineDate(0),2)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&newCode&".jpg"">"
'	response.write "<span class=""style6"">*adfsdfsd*</span>"
'	response.write "<script>haiwaocde """&Sys_DeallIneDate(0)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&"241"",""popobj1""</script>"
'	response.write haiwaocde(Sys_DeallIneDate(0)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&"241")
%>
</div>
<div id="Layer11" class="style2" style="position:absolute; left:80px; top:815px; z-index:1"><%
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
<div id="Layer12" class="style2" style="position:absolute; left:80px; top:865px; z-index:1"><%
	tmpDeallIneDate=right(Sys_DealLineDate(0),2)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)
	tempChkMemony=right("0000000000"&(ForFeitSum+43),9)
	SysChkNumber=CreateChkCode(tmpDeallIneDate,newCode,Sys_ImageFileNameB,SysDate,tempChkMemony)
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

<div id="Layer13" class="style3" style="position:absolute; left:630px; top:870px; z-index:1"><%
	response.write Sys_CarNo&"<br>"
	response.write "經收人蓋章："
%>
</div>

<div id="Layer14" class="style2" style="position:absolute; left:300px; top:925px; z-index:1"><%="車號："&Sys_CarNo%></div>

<div id="Layer15" class="style3" style="position:absolute; left:130px; top:980px; z-index:1"><%
	response.write funcCheckFont(Sys_Owner,18,1)&"<br>"
	response.write funcCheckFont(Sys_Address,18,1)
%>
</div>
<div id="Layer16" class="style3" style="position:absolute; left:130px; top:1080px; z-index:1"><%
	response.write Sys_ImageFileNameB
%>
</div>
<div id="Layer16" class="style3" style="position:absolute; left:124px; top:1030px; z-index:1"><%
	DelphiASPObj.CreateBarCode cdbl(Sys_ImageFileNameB),39,35,160
	response.write "<img src=""../BarCodeImage/"&(cdbl(Sys_ImageFileNameB))&".jpg"">"
%>
</div>
<div id="Layer17" class="style3" style="position:absolute; left:410px; top:1055px; z-index:1"><%
	response.write "<img src=""../BarCodeImage/"&Sys_MailNumber&".jpg""><br>　　　　　"&Sys_MailChkNumber
'	response.write "<img src=""../BarCodeImage/"&Sys_ImageFileNameB&".jpg"">"
'	response.write "<span class=""style6"">"&Sys_ImageFileNameB&"</span>"
'	response.write "<span id=""popobj4""></span>"
'	response.write "<script>haiwaocde """&Sys_ImageFileNameB&""",""popobj4""</script>"
'	response.write haiwaocde(Sys_ImageFileNameB)
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