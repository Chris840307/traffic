<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>無標題文件</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>

<body>
<%
on Error Resume Next
if trim(request("printStyle"))<>"" then
PBillSN=split(trim(request("hd_BillSN")),",")
for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext""></div>"
Sys_Sex=""
Sys_DCIRETURNCARTYPE=""
Sys_IMAGEFILENAME=""
Sys_IMAGEPATHNAME=""
Sys_BillFillerMemberID=0
Sys_BillTypeID=1

strSql="select * from PasserBase where SN="&PBillSN(i)
set rs=conn.execute(strSql)
if Not rs.eof then Sys_BillNo=trim(rs("BillNo"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_RuleVer=trim(rs("RuleVer"))
if Not rs.eof then Sys_Driver=trim(rs("Driver"))
if Not rs.eof then Sys_DriverID=trim(rs("DriverID"))
if Not rs.eof then Sys_DriverAddress=trim(rs("DriverAddress"))
if Not rs.eof then Sys_DriverZip=trim(rs("DriverZip"))
if Not rs.eof then Sys_Owner=trim(rs("Owner"))
if Not rs.eof then Sys_OwnerID=trim(rs("OwnerID"))
if Not rs.eof then Sys_OwnerAddress=trim(rs("OwnerAddress"))
if Not rs.eof then Sys_OwnerZip=trim(rs("OwnerZip"))
if Not rs.eof then Sys_Rule1=trim(rs("Rule1"))
if Not rs.eof then
	If not ifnull(Trim(rs("DriverID"))) Then
		If Mid(Trim(rs("DriverID")),2,1)="1" Then
			Sys_Sex="男"
		elseif Mid(Trim(rs("DriverID")),2,1)="2" Then
			Sys_Sex="女"
		End if
	End if
end if

if Not rs.eof then
	Sys_IllegalDate=split(gArrDT(trim(rs("IllegalDate"))),"-")
else
	Sys_IllegalDate=split(gArrDT(trim("")),"-")
end if
if Not rs.eof then
	Sys_IllegalDate_h=hour(trim(rs("IllegalDate")))
else
	Sys_IllegalDate_h=""
end if
if Not rs.eof then
	Sys_IllegalDate_m=minute(trim(rs("IllegalDate")))
else
	Sys_IllegalDate_m=""
end if
if Not rs.eof then
	Sys_DealLineDate=split(gArrDT(trim(rs("DealLineDate"))),"-")
else
	Sys_DealLineDate=split(gArrDT(trim("")),"-")
end if
if Not rs.eof then
	Sys_DriverBirth=split(gArrDT(trim(rs("DriverBirth"))),"-")
else
	Sys_DriverBirth=split(gArrDT(trim("")),"-")
end if
if Not rs.eof then Sys_BillFillerMemberID=trim(rs("BillFillerMemberID"))
if Not rs.eof then Sys_BillUnitID=trim(rs("BillUnitID"))
rs.close

strSQL="select ZipName from Zip where ZipID='"&Sys_DriverZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName"))
rszip.close

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close

strSql="select a.LoginID,a.ChName,b.UnitName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_BillFillerMemberID
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitName=trim(mem("UnitName"))
if Not mem.eof then Sys_UnitTel=trim(mem("Tel"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close

if trim(Sys_Rule1)<>"" and not isnull(Sys_Rule1) then
	strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		Sys_Level1=trim(rsRule1("Level1"))
		Sys_IllegalRule=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
end if
rs.close
strSql="select MailNumber,MailTypeID,MailDate,MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&PBillSN(i)

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close

strSql="select * from UnitInfo where UnitID='"&Sys_BillUnitID&"'"
set rs=conn.execute(strSql)

if Not rs.eof then Sys_STATIONNAME=trim(rs("UnitName"))
if Not rs.eof then Sys_StationTel=trim(rs("Tel"))
if Not rs.eof then Sys_StationID=trim(rs("UnitID"))
rs.close
sys_Date=split(gArrDT(date),"-")
fastring=""

if trim(Sys_BillTypeID)="1" then
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sys_Level1,0,True,False,Sys_MailDate
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sys_Level1&",0,True,False,"&Sys_MailDate
	'response.end
else
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sys_Level1,0,True,False,Sys_MailDate
end if
%>
<!--<div id="Layer1" style="position:absolute; left:70px; top:<%=5+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>-->
<div id="Layer2" style="position:absolute; left:70px; top:<%=35+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer3" style="position:absolute; left:185px; top:<%=5+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
<div id="Layer4" style="position:absolute; left:185px; top:<%=25+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<!--<div id="Layer5" style="position:absolute; left:185px; top:<%=45+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>-->
<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer6" style="position:absolute; left:665px; top:<%=10+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer7" style="position:absolute; left:665px; top:<%=25+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer8" style="position:absolute; left:665px; top:<%=40+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<div id="Layer9" style="position:absolute; left:35px; top:<%=55+1085*i%>px; width:202px; height:36px; z-index:5"><img src=<%="""..\BarCodeImage\"&Sys_BillNo&"_3.jpg"""%>></div>
<div id="Layer10" style="position:absolute; left:550px; top:<%=55+1085*i%>px; width:233px; height:32px; z-index:6"><img src=<%="""..\BarCodeImage\"&Sys_BillNo&"_4.jpg"""%>></div>
<div id="Layer11" style="position:absolute; left:515px; top:<%=100+1085*i%>px; width:230px; height:12px; z-index:7"><font size=1>北縣警交字第<%=Sys_BillNo%>號</font></div>
<div id="Layer12" style="position:absolute; left:130px; top:<%=115+1085*i%>px; width:150px; height:11px; z-index:3">逕行舉發<br><font size=2>附採證照片</font></div>
<div id="Layer13" style="position:absolute; left:285px; top:<%=115+1085*i%>px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" style="position:absolute; left:385px; top:<%=115+1085*i%>px; width:324px; height:10px; z-index:4"><font size=2><%=Sys_DriverZipName&Sys_DriverHomeAddress%>test</font></div>
<div id="Layer15" style="position:absolute; left:290px; top:<%=135+1085*i%>px; width:100px; height:10px; z-index:8"><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&Sys_DriverBirth(1)&"月"&Sys_DriverBirth(2)&"日"%></div>
<div id="Layer16" style="position:absolute; left:460px; top:<%=135+1085*i%>px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" style="position:absolute; left:655px; top:<%=135+1085*i%>px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" style="position:absolute; left:130px; top:<%=165+1085*i%>px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer19" style="position:absolute; left:295px; top:<%=165+1085*i%>px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" style="position:absolute; left:520px; top:<%=165+1085*i%>px; width:201px; height:17px; z-index:13"><%=Sys_Owner%></div>
<div id="Layer21" style="position:absolute; left:140px; top:<%=185+1085*i%>px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZipName&Sys_OwnerAddress%></div>

<div id="Layer22" style="position:absolute; left:130px; top:<%=205+1085*i%>px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%>年</div>
<div id="Layer23" style="position:absolute; left:185px; top:<%=205+1085*i%>px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%>月</div>
<div id="Layer24" style="position:absolute; left:225px; top:<%=205+1085*i%>px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%>日</div>
<div id="Layer25" style="position:absolute; left:265px; top:<%=205+1085*i%>px; width:40px; height:16px; z-index:18"><%=Sys_IllegalDate_h%>時</div>
<div id="Layer26" style="position:absolute; left:310px; top:<%=205+1085*i%>px; width:40px; height:13px; z-index:19"><%=Sys_IllegalDate_m%>分</div>
<div id="Layer27" style="position:absolute; left:445px; top:<%=220+1085*i%>px; width:600px; height:21px; z-index:20"><%=Sys_IllegalRule%></div>
<div id="Layer28" style="position:absolute; left:130px; top:<%=235+1085*i%>px; width:217px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer29" style="position:absolute; left:130px; top:<%=270+1085*i%>px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%>年</div>
<div id="Layer30" style="position:absolute; left:185px; top:<%=270+1085*i%>px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%>月</div>
<div id="Layer31" style="position:absolute; left:225px; top:<%=270+1085*i%>px; width:32px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%>日</div>
<div id="Layer32" style="position:absolute; left:445px; top:<%=280+1085*i%>px; width:400px; height:49px; z-index:29"><%response.write "違反道路交通管理處罰條例第"&left(trim(Sys_Rule1),2)&"條"
			if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
				response.write "<br>"
				response.write Mid(trim(Sys_Rule1),3,2)&"項"&Mid(trim(Sys_Rule1),5,2)&"款規定"%></div>

<div id="Layer33" style="position:absolute; left:445px; top:<%=345+1085*i%>px; width:100px; height:30px; z-index:28"><font size=2><%=Sys_STATIONNAME%></font></div>
<div id="Layer34" style="position:absolute; left:525px; top:<%=315+1085*i%>px; width:400px; height:30px; z-index:28"><img src=<%="""..\BarCodeImage\"&Sys_BillNo&"_5.jpg"""%>></div>
<div id="Layer35" style="position:absolute; left:455px; top:<%=420+1085*i%>px; width:100px; height:49px; z-index:29"><%=Sys_UnitName%></div>
<div id="Layer36" style="position:absolute; left:580px; top:<%=420+1085*i%>px; width:100px; height:43px; z-index:30">主管</div>
<div id="Layer37" style="position:absolute; left:690px; top:<%=420+1085*i%>px; width:200px; height:46px; z-index:31"><%=Sys_ChName%></div>
<div id="Layer38" style="position:absolute; left:175px; top:<%=470+1085*i%>px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%>年</div>
<div id="Layer39" style="position:absolute; left:330px; top:<%=470+1085*i%>px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%>月</div>
<div id="Layer40" style="position:absolute; left:480px; top:<%=470+1085*i%>px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%>日</div>
<div id="Layer41" style="position:absolute; left:580px; top:<%=470+1085*i%>px; width:80px; height:12px; z-index:36">填單　<%=Sys_BillFillerMemberID%></div>



<!--<div id="Layer42" style="position:absolute; left:65px; top:<%=530+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>-->
<div id="Layer43" style="position:absolute; left:65px; top:<%=555+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>

<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer44" style="position:absolute; left:180px; top:<%=535+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer45" style="position:absolute; left:180px; top:<%=550+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>
<!--<div id="Layer46" style="position:absolute; left:180px; top:<%=565+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>-->

<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer47" style="position:absolute; left:660px; top:<%=535+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer48" style="position:absolute; left:660px; top:<%=550+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer49" style="position:absolute; left:660px; top:<%=565+1085*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>

<div id="Layer50" style="position:absolute; left:35px; top:<%=585+1085*i%>px; width:202px; height:36px; z-index:5"><img src=<%="""..\BarCodeImage\"&Sys_BillNo&"_3.jpg"""%>></div>
<div id="Layer51" style="position:absolute; left:550px; top:<%=580+1085*i%>px; width:233px; height:32px; z-index:6"><img src=<%="""..\BarCodeImage\"&Sys_BillNo&"_4.jpg"""%>></div>
<div id="Layer52" style="position:absolute; left:515px; top:<%=625+1085*i%>px; width:230px; height:12px; z-index:7"><font size=1>北縣警交字第<%=Sys_BillNo%>號</font></div>
<div id="Layer53" style="position:absolute; left:130px; top:<%=640+1085*i%>px; width:150px; height:11px; z-index:3">逕行舉發<br><font size=2>附採證照片</font></div>
<div id="Layer54" style="position:absolute; left:295px; top:<%=640+1085*i%>px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer55" style="position:absolute; left:395px; top:<%=640+1085*i%>px; width:324px; height:10px; z-index:4"><font size=2><%=Sys_DriverZipName&Sys_DriverHomeAddress%>test</font></div>
<div id="Layer56" style="position:absolute; left:290px; top:<%=665+1085*i%>px; width:100px; height:10px; z-index:8"><%=Sys_DriverBirth(0)%>年<%=Sys_DriverBirth(1)%>月<%=Sys_DriverBirth(2)%>日</div>
<div id="Layer57" style="position:absolute; left:460px; top:<%=665+1085*i%>px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer58" style="position:absolute; left:655px; top:<%=665+1085*i%>px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer59" style="position:absolute; left:130px; top:<%=685+1085*i%>px; width:100px; height:14px; z-index:11"><%=Sys_CarNo%></div>
<div id="Layer60" style="position:absolute; left:295px; top:<%=685+1085*i%>px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer61" style="position:absolute; left:520px; top:<%=685+1085*i%>px; width:201px; height:17px; z-index:13"><%=Sys_Owner%></div>
<div id="Layer62" style="position:absolute; left:130px; top:<%=710+1085*i%>px; width:507px; height:13px; z-index:14"><%=Sys_OwnerZipName&Sys_OwnerAddress%></div>

<div id="Layer63" style="position:absolute; left:130px; top:<%=735+1085*i%>px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%>年</div>
<div id="Layer64" style="position:absolute; left:185px; top:<%=735+1085*i%>px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%>月</div>
<div id="Layer65" style="position:absolute; left:225px; top:<%=735+1085*i%>px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%>日</div>
<div id="Layer66" style="position:absolute; left:265px; top:<%=735+1085*i%>px; width:40px; height:16px; z-index:18"><%=Sys_IllegalDate_h%>時</div>
<div id="Layer67" style="position:absolute; left:310px; top:<%=735+1085*i%>px; width:40px; height:13px; z-index:19"><%=Sys_IllegalDate_m%>分</div>
<div id="Layer68" style="position:absolute; left:445px; top:<%=750+1085*i%>px; width:600px; height:21px; z-index:20"><%=Sys_IllegalRule%></div>
<div id="Layer69" style="position:absolute; left:130px; top:<%=765+1085*i%>px; width:217px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer70" style="position:absolute; left:130px; top:<%=800+1085*i%>px; width:34px; height:11px; z-index:22"><%=Sys_IllegalDate(0)%>年</div>
<div id="Layer71" style="position:absolute; left:185px; top:<%=800+1085*i%>px; width:35px; height:13px; z-index:23"><%=Sys_IllegalDate(1)%>月</div>
<div id="Layer72" style="position:absolute; left:225px; top:<%=800+1085*i%>px; width:32px; height:15px; z-index:24"><%=Sys_IllegalDate(2)%>日</div>
<div id="Layer73" style="position:absolute; left:445px; top:<%=805+1085*i%>px; width:400px; height:49px; z-index:29"><%response.write "違反道路交通管理處罰條例第"&left(trim(Sys_Rule1),2)&"條"
			if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
				response.write "<br>"
				response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)&"規定"%></div>

<div id="Layer74" style="position:absolute; left:445px; top:<%=875+1085*i%>px; width:100px; height:30px; z-index:28"><font size=2><%=Sys_STATIONNAME%></font></div>
<div id="Layer75" style="position:absolute; left:520px; top:<%=840+1085*i%>px; width:400px; height:30px; z-index:28"><img src=<%="""..\BarCodeImage\"&Sys_BillNo&"_5.jpg"""%>></div>
<div id="Layer76" style="position:absolute; left:455px; top:<%=950+1085*i%>px; width:100px; height:49px; z-index:29"><%=Sys_UnitName%></div>
<div id="Layer77" style="position:absolute; left:630px; top:<%=950+1085*i%>px; width:200px; height:46px; z-index:31"><%=Sys_ChName%></div>
<div id="Layer78" style="position:absolute; left:175px; top:<%=995+1085*i%>px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%>年</div>
<div id="Layer79" style="position:absolute; left:330px; top:<%=995+1085*i%>px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%>月</div>
<div id="Layer80" style="position:absolute; left:480px; top:<%=995+1085*i%>px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%>日</div>
<div id="Layer81" style="position:absolute; left:580px; top:<%=995+1085*i%>px; width:80px; height:12px; z-index:36">填單　<%=Sys_BillFillerMemberID%></div>
<%next
end if%></body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}
</script>