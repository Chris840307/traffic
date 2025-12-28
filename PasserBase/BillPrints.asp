<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>無標題文件</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 9px}
.style2 {font-size: 10px}
.style3 {font-size: 12px}
.style4 {font-size: 12px}
.style7 {font-size: 13px}
.style8 {font-size: 36px}
.style11 {font-size: 14px}
-->
</style>
</head>

<body>
<%
on Error Resume Next
PBillSN=split(trim(request("hd_BillSN")),",")

strSQL="select * from UnitInfo where UnitID='"&DB_UnitID&"'"
set rsUnit=conn.execute(strSQL)
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
rsUnit.close

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
'if Not rs.eof then Sys_BillTypeID=trim(rs("BillTypeID"))
'if Not rs.eof then Sys_Driver=trim(rs("Driver"))
if Not rs.eof then Sys_INSURANCE=trim(rs("INSURANCE"))
if Not rs.eof then Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
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
if Not rs.eof then Sys_Note=trim(rs("Note"))
'trim(Sys_BillTypeID)="1" 是欄停
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

'strSql="select b.Content from BILLFASTENERDETAIL a,DCICode b where a.FASTENERTYPEID=b.ID and b.TypeID=6 and a.BIllSN="&replace(trim(rsbil("BillSN")),"","0")&" and a.CarNo='"&trim(rsbil("CarNo"))&"'"
'set rsfast=conn.execute(strsql)
fastring=""
'while Not rsfast.eof
'	if trim(fastring)<>"" then fastring=fastring&","
'	fastring=fastring&rsfast("Content")
'	rsfast.movenext
'wend
'rsfast.close

if trim(Sys_BillTypeID)="1" then
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sys_Level1,0,True,False,Sys_MailDate
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sys_Level1&",0,True,False,"&Sys_MailDate
	'response.end
else
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
	DelphiASPObj.GenBillPrintBarCode PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_OwnerZip,Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sys_Level1,0,True,False,Sys_MailDate
	'response.write "DelphiASPObj.GenBillPrintBarCode "& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_OwnerZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sys_Level1&",0,True,False,"&Sys_MailDate
	'response.end
end if
if err.Number<>0 then msgBox("資料讀取錯誤"&Cstr(Err.Number)&Err.Description)
err.Clear
%>
<div style="position:absolute; left:10px; top:<%=10+1306*i%>px;">
<table width="645" height="393" border="0">
  <tr>
    <td width="141" height="69" valign="top">&nbsp;</td>
    <td colspan="2">&nbsp;</td>
    <td rowspan="2" align="right" valign="top"><br>   	</td>
  </tr>
  <tr>
    <td height="41" align="center" valign="top">　　　<!--<img src=<%="""../BarCodeImage/"&Sys_BillNo&"_1.jpg"""%> hspace="0" vspace="0" align="top">--><br>　　　<span class="style7"><%=Sys_FirstBarCode%></span>
	</td>
	<%if trim(Sys_BillTypeID)="1" then%>
    <td colspan="2" align="left" valign="top" nowrap><span class="style7"><%=Sys_DriverZip%><br>
    <%=Sys_DriverZipName&Sys_DriverAddress%></span></td>
	<%elseif trim(Sys_BillTypeID)="2" then%>
	<td align="left" valign="top" nowrap><span class="style7"><%=Sys_OwnerZip%><br>
    <%=Sys_OwnerZipName&Sys_OwnerAddress%></span></td>
	<%end if%>
  </tr>
  <tr>
    <td>&nbsp;</td>
	<%if trim(Sys_BillTypeID)="1" then%>
    <td colspan="2"><span class="style7"><%=Sys_Driver%>　台啟</span></td>
	<%elseif trim(Sys_BillTypeID)="2" then%>
	<td width="222"><span class="style7"><%=Sys_Owner%>　台啟</span></td>
	<%end if%>
    <td width="92">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td width="145" align="center"><p class="style4">&nbsp;
   	    </p>
      <p class="style4">大宗郵資已付掛號函件<br>
    第<%=Sys_MailNumber%>號  </p>    </td>
    <td width="23" align="center">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td align="center"><div align="left"><!--<img src=<%="""../BarCodeImage/"&Sys_BillNo&"_2.jpg"""%>>--><br>
        <%=Sys_MAILCHKNUMBER%></div></td>
    <td align="center">&nbsp;</td>
    <td align="right"><p>&nbsp;</p>
    <p class="style8"><%=Sys_StationID%></p></td>
  </tr>
  <tr>
    <td height="98" valign="top" nowrap><p>　　　　<span class="style7">應到案處所：<%=Sys_STATIONNAME%></span><br>
   	　　　　<span class="style7">應到案處所電話：<%=Sys_StationTel%></span></p>
    <p>&nbsp;</p></td>
    <td colspan="2">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</div>
<%if trim(Sys_IMAGEFILENAME)<>"" then%>
<div style="position:absolute; left:370px; top:<%=485+1306*i%>px;"><img src=<%="""/Img"&Sys_IMAGEPATHNAME&Sys_IMAGEFILENAME&""""%> width="365" height="265"></DIV>
<%end if%>

<div id="Layer6" style="position:absolute; left:40px; top:<%=775+1306*i%>px; width:400px; height:36px; z-index:5"><span class="style7">查詢電話：<%=DB_UnitTel%>（<%=DB_UnitName%>）</span></div>
<!--<div id="Layer1" style="position:absolute; left:50px; top:<%=810+1306*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>-->
<!--<div id="Layer2" style="position:absolute; left:50px; top:<%=840+1306*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>-->
<%'if trim(Sys_BillTypeID)="1" then%>
<!--	<div id="Layer3" style="position:absolute; left:165px; top:<%=815+1306*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%'else%>
<div id="Layer4" style="position:absolute; left:165px; top:<%=830+1306*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%'end if%>
<div id="Layer5" style="position:absolute; left:165px; top:<%=845+1306*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>-->
<%if trim(Sys_BillTypeID)="1" then%>
	<%if trim(Sys_INSURANCE)="0" then%>
		<div id="Layer6" style="position:absolute; left:645px; top:<%=810+1306*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%elseif trim(Sys_INSURANCE)="1" then%>
		<div id="Layer7" style="position:absolute; left:645px; top:<%=825+1306*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%else%>
		<div id="Layer8" style="position:absolute; left:645px; top:<%=840+1306*i%>px; width:202px; height:36px; z-index:5">Ｖ</div>
	<%end if%>
<%end if%>
<!--<div id="Layer9" style="position:absolute; left:10px; top:<%=860+1306*i%>px; width:202px; height:36px; z-index:5"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_3.jpg"""%>></div>
<div id="Layer10" style="position:absolute; left:525px; top:<%=855+1306*i%>px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<div id="Layer11" style="position:absolute; left:590px; top:<%=895+1306*i%>px; width:230px; height:12px; z-index:7"><font size=1><%=Sys_BillNo%></font></div>-->
<%if trim(Sys_BillTypeID)="1" then%>
<div id="Layer12" style="position:absolute; left:125px; top:<%=930+1306*i%>px; width:150px; height:11px; z-index:3"><span class="style7"><%=Sys_Driver%></span></div>
<%end if
if trim(Sys_BillTypeID)="1" then%>
<div id="Layer13" style="position:absolute; left:275px; top:<%=915+1306*i%>px; width:28px; height:11px; z-index:3"><span class="style7"><%=Sys_Sex%></span></div>
<div id="Layer14" style="position:absolute; left:370px; top:<%=915+1306*i%>px; width:324px; height:10px; z-index:4"><span class="style7"><%=Sys_DriverZipName&Sys_DriverAddress%></Span></div>
<div id="Layer15" style="position:absolute; left:275px; top:<%=940+1306*i%>px; width:100px; height:10px; z-index:8"><span class="style7"><%=Sys_DriverBirth(0)%></span></div>
<div id="Layer142" style="position:absolute; left:305px; top:<%=940+1306*i%>px; width:100px; height:10px; z-index:8"><span class="style7"><%=Sys_DriverBirth(1)%></span></div>
<div id="Layer143" style="position:absolute; left:335px; top:<%=940+1306*i%>px; width:100px; height:10px; z-index:8"><span class="style7"><%=Sys_DriverBirth(2)%></span></div>
<div id="Layer16" style="position:absolute; left:440px; top:<%=940+1306*i%>px; width:106px; height:13px; z-index:9"><span class="style7"><%=Sys_DriverID%></span></div>
<div id="Layer17" style="position:absolute; left:640px; top:<%=940+1306*i%>px; width:99px; height:12px; z-index:10"><span class="style7"><%=fastring%></span></div>
<%end if%>
<div id="Layer18" style="position:absolute; left:125px; top:<%=965+1306*i%>px; width:100px; height:14px; z-index:11"><span class="style7"><%=Sys_CarNo%></span></div>
<div id="Layer19" style="position:absolute; left:275px; top:<%=965+1306*i%>px; width:117px; height:20px; z-index:12"><span class="style7"><%=Sys_DCIRETURNCARTYPE%></span></div>
<div id="Layer20" style="position:absolute; left:510px; top:<%=965+1306*i%>px; width:201px; height:17px; z-index:13"><span class="style7"><%=Sys_Owner%></span></div>
<div id="Layer21" style="position:absolute; left:125px; top:<%=990+1306*i%>px; width:507px; height:13px; z-index:14"><span class="style7"><%=Sys_OwnerZipName&Sys_OwnerAddress%></span></div>

<div id="Layer22" style="position:absolute; left:130px; top:<%=1010+1306*i%>px; width:40px; height:13px; z-index:15"><span class="style7"><%=Sys_IllegalDate(0)%></span></div>
<div id="Layer23" style="position:absolute; left:180px; top:<%=1010+1306*i%>px; width:40px; height:17px; z-index:16"><span class="style7"><%=Sys_IllegalDate(1)%></span></div>
<div id="Layer24" style="position:absolute; left:230px; top:<%=1010+1306*i%>px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" style="position:absolute; left:275px; top:<%=1010+1306*i%>px; width:40px; height:16px; z-index:18"><span class="style7"><%=Sys_IllegalDate_h%></span></div>
<div id="Layer26" style="position:absolute; left:325px; top:<%=1010+1306*i%>px; width:40px; height:13px; z-index:19"><span class="style7"><%=Sys_IllegalDate_m%></span></div>
<div id="Layer27" style="position:absolute; left:415px; top:<%=1015+1306*i%>px; width:600px; height:21px; z-index:20"><span class="style7"><%response.write Sys_IllegalRule
	if trim(Sys_Note)<>"" then response.write "("&Sys_Note&")"
%></span></div>
<div id="Layer28" style="position:absolute; left:125px; top:<%=1035+1306*i%>px; width:217px; height:15px; z-index:21"><span class="style7"><%=Sys_ILLEGALADDRESS%></span></div>
<div id="Layer29" style="position:absolute; left:130px; top:<%=1065+1306*i%>px; width:34px; height:11px; z-index:22"><span class="style7"><%=Sys_DealLineDate(0)%></span></div>
<div id="Layer30" style="position:absolute; left:220px; top:<%=1065+1306*i%>px; width:35px; height:13px; z-index:23"><span class="style7"><%=Sys_DealLineDate(1)%></span></div>
<div id="Layer31" style="position:absolute; left:300px; top:<%=1065+1306*i%>px; width:32px; height:15px; z-index:24"><span class="style7"><%=Sys_DealLineDate(2)%></span></div>
<div id="Layer32" style="position:absolute; left:405px; top:<%=1075+1306*i%>px; width:400px; height:49px; z-index:29"><span class="style7"><%response.write "第"&left(trim(Sys_Rule1),2)&"條"
			if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
				response.write Mid(trim(Sys_Rule1),3,1)&"項"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)&"規定"
				response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"%></span></div>

<div id="Layer33" style="position:absolute; left:405px; top:<%=1120+1306*i%>px; width:80px; height:30px; z-index:28"><span class="style7"><%=Sys_STATIONNAME%></span></font></div>
<!--<div id="Layer34" style="position:absolute; left:485px; top:<%=1100+1306*i%>px; width:400px; height:30px; z-index:28"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_5.jpg"""%>></div>-->
<div id="Layer35" style="position:absolute; left:370px; top:<%=1165+1306*i%>px; width:100px; height:49px; z-index:29"><%
	'if billprintuseimage=1 then
		'response.write "<img src=""../UnitInfo/Picture/"&Sys_UnitFilename&""" width=""70"" height=""70"">"
	'else
		'response.write Sys_UnitName
	'end if%></div>
<div id="Layer36" style="position:absolute; left:510px; top:<%=1180+1306*i%>px; width:100px; height:43px; z-index:30"><%'=主管%></div>
<div id="Layer37" style="position:absolute; left:625px; top:<%=1180+1306*i%>px; width:200px; height:46px; z-index:31"><%
	if trim(Sys_MemberFilename)<>"" then
		response.write "<img src=""../Member/Picture/"&Sys_MemberFilename&""" width=""80"" height=""25"">"
	else
		response.write Sys_ChName
	end if%></div>
<div id="Layer38" style="position:absolute; left:210px; top:<%=1250+1306*i%>px; width:60px; height:10px; z-index:32"><span class="style7"><%=sys_Date(0)%></span></div>
<div id="Layer39" style="position:absolute; left:365px; top:<%=1250+1306*i%>px; width:60px; height:13px; z-index:33"><span class="style7"><%=sys_Date(1)%></span></div>
<div id="Layer40" style="position:absolute; left:515px; top:<%=1250+1306*i%>px; width:60px; height:11px; z-index:34"><span class="style7"><%=sys_Date(2)%></span></div>
<div id="Layer41" style="position:absolute; left:680px; top:<%=1250+1306*i%>px; width:80px; height:12px; z-index:36"><span class="style7"><%=Sys_BillFillerMemberID%></span></div>
<%next%>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}
</script>