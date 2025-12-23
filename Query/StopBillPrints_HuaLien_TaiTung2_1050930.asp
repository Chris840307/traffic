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
.style7 {font-family:"標楷體"; font-size: 16px; line-height:2;}
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
errBillno=""
Server.ScriptTimeout=6000
PageCnt=0
tmpdate=split(gArrDT(trim(date)),"-")
SysDate=tmpdate(1)&tmpdate(2)
'on Error Resume Next
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
PrintCnt=0
toppx=15

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

		strSQL="select distinct DeCode(Driver,null,Owner,Driver) Driver,DriverZip,DriverAddress from billbase where ImageFileNameB='"&trim(PBillNo(cmtI))&"'"

		set rsbill=conn.execute(strSQL)

		If not ifnull(rsbill("DriverAddress")) and not ifnull(rsbill("Driver")) Then
			Sys_Owner=trim(rsbill("Driver"))
			Sys_OwnerZip=trim(rsbill("DriverZip"))
			Sys_OwnerAddress=trim(rsbill("DriverAddress"))
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

	strSQL="select distinct StoreAndSendMailNumber from StopBillMailHistory where BillNo='"&PBillNo(cmtI)&"'"
	set rsmail=conn.execute(strSQL)
	If Not rsmail.eof Then
		Sys_MailNumber=trim(rsmail("StoreAndSendMailNumber"))&"95100017"
		Sys_MailChkNumber=trim(rsmail("StoreAndSendMailNumber"))&" 951000 17"
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

	If ifnull(Sys_Owner) Then errBillno=errBillno&Sys_ImageFileNameB&"\n"
%>
<div id="L78" style="position:relative;">

<div id="Layer5" class="style3" style="position:absolute; width:200px; left:80px; top:0px; z-index:1">
	請繳回：臺東縣警察局交通隊
</div>

<div id="Layer5" class="style3" style="position:absolute; width:200px; left:520px; top:0px; z-index:1">
	地址：95051臺東市更生路11號
</div>

<div id="Layer15" class="style3" style="position:absolute; width:200px; left:130px; top:<%=5+toppx%>px; z-index:1"><%
	response.write funcCheckFont(Sys_Owner&"　"&Sys_CarNo,18,1)&"<br>"
	response.write funcCheckFont(Sys_Address,18,1)
%>
</div>

<div id="Layer070" style="position:absolute; left:420px; top:<%=0+toppx%>px; z-index:8"><%
	Response.Write "<img src=""../Image/BillNoPage.gif"" width=""80"">"
	%>
</div>

<div id="Layer071" style="position:absolute; left:435px; top:<%=33+toppx%>px; font-size: 12px; z-index:9"><%
	Response.Write replace(gArrDT(date),"-",".")
	%>
</div>

<div id="Layer17" class="style3" style="position:absolute; left:415px; top:<%=85+toppx%>px; z-index:1"><%
	Response.Write Sys_ImageFileNameB&"&nbsp;&nbsp;二次"
%>
</div>

<div id="Layer17" class="style3" style="position:absolute; left:540px; top:<%=80+toppx%>px; z-index:1"><%
	response.write "<img src=""../BarCodeImage/"&Sys_ImageFileNameB&".jpg"" height=""25"">"
%>
</div>


<div id="Layer1" class="style7" style="position:absolute; left:60px; top:<%=240+toppx%>px; width:400px; z-index:1">寄件人：臺東縣警察局交通隊</div>

<div id="Layer2" class="style7" style="position:absolute; left:60px; top:<%=260+toppx%>px; z-index:1">收件人：<%=funcCheckFont(Sys_Owner,20,1)%></div>

<div id="Layer1" class="style7" style="position:absolute; left:60px; top:<%=280+toppx%>px; width:400px; z-index:1">地址：<%=funcCheckFont(Sys_Address,20,1)%></div>

<div id="Layer13" class="style3" style="position:absolute; left:200px; top:<%=340+toppx%>px; z-index:1"><%
	response.write Sys_CarNo
%>
</div>

<div id="Layer17" class="style3" style="position:absolute; left:280px; top:<%=325+toppx%>px; z-index:1"><%
	response.write "<img src=""../BarCodeImage/"&Sys_ImageFileNameB&".jpg"" height=""25"">"
	Response.Write "<br>　"&Sys_ImageFileNameB
%>
</div>

<div id="Layer18" class="style4" style="position:absolute; font-size:14px; left:500px; top:<%=275+toppx%>px; z-index:5"><%
	DelphiASPObj.CreateBarCode Sys_MailNumber,128,25,260
	response.write "<img src=""../BarCodeImage/"&Sys_MailNumber&".jpg"" width=""220"" height=""70""><br>　　　　"&Sys_MailChkNumber
%></div>

<div id="Layer10" class="style2" style="position:absolute; left:80px; top:<%=540+toppx%>px; z-index:1"><%
	response.write "繳費期限及單位代碼："&Sys_DeallIneDate(0)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&"277"
	response.write "　　&nbsp;&nbsp;"
	DelphiASPObj.CreateBarCode right(Sys_DealLineDate(0),2)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&"277",0,35,160
	response.write "<img src=""../BarCodeImage/"&right(Sys_DealLineDate(0),2)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&"277.jpg"">"
'	response.write "<span class=""style6"">*adfsdfsd*</span>"
'	response.write "<script>haiwaocde """&Sys_DeallIneDate(0)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&"241"",""popobj1""</script>"
'	response.write haiwaocde(Sys_DeallIneDate(0)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)&"241")
%>
</div>
<div id="Layer11" class="style2" style="position:absolute; left:80px; top:<%=590+toppx%>px; z-index:1"><%
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
<div id="Layer12" class="style2" style="position:absolute; left:80px; top:<%=640+toppx%>px; z-index:1"><%
	ForFeitSum=0
	strSQL="select sum(ForFeit1) ForFeit1 from BillBase where ImageFileNameB='"&Sys_ImageFileNameB&"' order by IllegalDate"
	set rst=conn.execute(strSQL)
	While Not rst.eof
		
		ForFeitSum=cdbl(rst("ForFeit1"))
		rst.movenext
	Wend
	rst.close

	tmpDeallIneDate=right(Sys_DealLineDate(0),2)&Sys_DeallIneDate(1)&Sys_DeallIneDate(2)
	tempChkMemony=right("0000000000"&(ForFeitSum+43),9)
	SysChkNumber=CreateChkCode(tmpDeallIneDate,"277",Sys_ImageFileNameB,SysDate,tempChkMemony)
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

<div id="Layer13" class="style3" style="position:absolute; left:630px; top:<%=640+toppx%>px; z-index:1"><%
	response.write Sys_CarNo
%>
</div>

<div id="Layer14" class="style2" style="position:absolute; left:300px; top:<%=700+toppx%>px; z-index:1"><%="車號："&Sys_CarNo%></div>

<div id="Layer5" class="style2" style="position:absolute; left:50px; top:<%=805+toppx%>px; z-index:1"><%="車號："&Sys_CarNo&"　車主地址："&funcCheckFont(Sys_Address,20,1)%></div>

<div id="Layer6" style="position:absolute; left:50px; top:<%=835+toppx%>px; z-index:1">
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
				response.write "<td class=""style3"">"&right("00"&hour(rst("IllegalDate")),2)&":"&right("00"&minute(rst("IllegalDate")),2)&"</td>"
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
<div id="Layer7" class="style4" style="position:absolute; left:50px; top:<%=1050+toppx%>px; z-index:1"><B><%
	response.write "催繳單號："&Sys_ImageFileNameB&"　　　繳費期限"&Sys_DeallIneDate(0)&"/"&Sys_DeallIneDate(1)&"/"&Sys_DeallIneDate(2)&"日止"
%></B>
</div>
<div id="Layer8" class="style4" style="position:absolute; left:50px; top:<%=1075+toppx%>px; z-index:1"><B><%
	response.write "共計催繳："&Cntsum&"筆，停車費："&ForFeitSum&"元、工本費43元，總金額："&(ForFeitSum+43)&"元"
%></B>
</div>

<div id="Layer9" class="style3" style="position:absolute; left:50px; top:<%=1125+toppx%>px; z-index:1">
請於收到本催繳通知單於繳費期限內繳納費用，逾期仍未繳納者，將依違反道路交通管理處罰條例第56條3項<br>
逕行舉發。<br>
繳費方式：請持本催繳通知單至臺東縣公有路邊停車場收費管理中心國雲科技(股)公司(臺東市武昌街120號)<br>
或至全省統一超商7-ELEVEN、全家便利商店各門市繳納，繳費後請保留本收據聯6個月，如已有繳交停車欠費者<br>
請勿重覆繳費，以維護您的權益。另有提供停車費預付自動扣繳方案，可避免停車單逾期或遺失造成催繳及違反<br>
道路交通管理處罰條例遭警察局逕行舉發。<br>
<span class="style6"><b>
	台端查詢停車紀錄及任何疑問：請電(089)342-500</b>
</span><br>
或網址：http://tt.guoyun.com.tw/ParkQuery/上網查詢。
</div>

</div>
<%
response.flush
next%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	var errBillno='<%=trim(errBillno)%>';
	if(errBillno!=''){alert("以下單號姓名是空的請確認：\n"+errBillno);}
	window.focus();
	printWindow(true,5.08,5.08,5.08,5.08);
</script>