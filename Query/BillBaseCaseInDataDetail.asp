<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_入案資料審核清冊 .xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 

	if request("Sys_BatchNumber")<>"" then

		strwhere=" sn in (select BillSn from DciLog where batchnumber='"&request("Sys_BatchNumber")&"' and dcireturnstatusid in('S','Y','n'))"
	end If 

	strSQL="select * from (" & _
	"select Sn,BillTypeID,BillNo,CarNo,DriverID,IllegalDate,RecordDate,IllegalAddress" & _
	",DeCode(CarSimpleID,1,'汽車',2,'拖車',3,'重機',4,'輕機',6,'臨時車牌') CarSimpleID" & _
	",(select DciReturnCarType from BillbaseDCIReturn where BillNo=BillBase.BillNo and CarNo=BillBase.CarNo and ExchangetypeID='W') DciReturnCarType" & _
	",(select DciReturnCarColor from BillBaseDCIReturn where BillNo=BillBase.BillNo and CarNo=BillBase.CarNo and ExchangetypeID='W') DciReturnCarColor" & _
	",DeCode(MemberStation,null,(select DCIReturnStation from BillBaseDCIReturn where BillNo=BillBase.BillNo and CarNo=BillBase.CarNo and ExchangetypeID='W'),MemberStation) DCIReturnStation" & _
	",Rule1||DeCode(Rule2,null,null,'/'||Rule2)||DeCode(Rule3,null,null,'/'||Rule3) Rule" & _
	",BillFillDate,DeallineDate" & _
	",(Select UnitName from Unitinfo where UnitID=BillBase.BillUnitID) BillUnitName" & _
	",BillMem1||DeCode(BillMem2,null,null,'/'||BillMem2)||DeCode(BillMem3,null,null,'/'||BillMem3)||DeCode(BillMem4,null,null,'/'||BillMem4) BillMem" & _
	",DeCode(SignType,'A','簽收','U','拒簽收','2','拒簽已收','3','已簽拒收','5','補開單') SignType" & _
	",RuleSpeed,IllegalSpeed" & _
	",(select Driver from BillbaseDCIReturn where BillNo=BillBase.BillNo and CarNo=BillBase.CarNo and ExchangetypeID='W') Driver" & _
	",(select (select zipName from zip where zipID=BillBaseDcireturn.DriverHomeZip) from BillbaseDCIReturn where BillNo=BillBase.BillNo and CarNo=BillBase.CarNo and ExchangetypeID='W') DriverZipName" & _
	",(select DriverHomeAddress from BillbaseDCIReturn where BillNo=BillBase.BillNo and CarNo=BillBase.CarNo and ExchangetypeID='W') DriverAddress" & _
	",(select Owner from BillbaseDCIReturn where BillNo=BillBase.BillNo and CarNo=BillBase.CarNo and ExchangetypeID='W') Owner" & _
	",(select (select zipName from zip where zipID=BillBaseDcireturn.OwnerZip) from BillbaseDCIReturn where BillNo=BillBase.BillNo and CarNo=BillBase.CarNo and ExchangetypeID='W') OwnerZipName" & _
	",(select OwnerAddress from BillbaseDCIReturn where BillNo=BillBase.BillNo and CarNo=BillBase.CarNo and ExchangetypeID='W') OwnerAddress" & _
	",Driver DriverFix,(select zipName from zip where zipID=BillBase.DriverZip) DriverZipNameFix,DriverAddress DriverAddressFix" & _
	",Owner OwnerFix,(select zipName from zip where zipID=BillBase.OwnerZip) OwnerZipNameFix,OwnerAddress OwnerAddressFix " & _
	" from BillBase where " & strwhere & _
	") bill order by BillTypeID,RecordDate"
'	" Union All " & _
'	"select Sn,'3' BillTypeID,BillNo,'' CarNo,DriverID,IllegalDate,RecordDate,IllegalAddress" & _
'	",'行人慢車' CarSimpleID" & _
'	",'' DciReturnCarType" & _
'	",'' DciReturnCarColor" & _
'	",(select UnitName from Unitinfo where UnitID=PasserBase.MemberStation) DCIReturnStation" & _
'	",Rule1||DeCode(Rule2,null,null,'/'||Rule2)||DeCode(Rule3,null,null,'/'||Rule3) Rule" & _
'	",BillFillDate,DeallineDate" & _
'	",(Select UnitName from Unitinfo where UnitID=PasserBase.BillUnitID) BillUnitName" & _
'	",BillMem1||DeCode(BillMem2,null,null,'/'||BillMem2)||DeCode(BillMem3,null,null,'/'||BillMem3)||DeCode(BillMem4,null,null,'/'||BillMem4) BillMem" & _
'	",DeCode(SignType,'A','簽收','U','拒簽收','2','拒簽已收','3','已簽拒收','5','補開單') SignType" & _
'	",null RuleSpeed,null IllegalSpeed" & _
'	",Driver" & _
'	",DriverAddress" & _
'	",'' Owner" & _
'	",'' OwnerAddress" & _
'	" from PasserBase where RecordDate" & strwhere & _
'	") bill order by BillTypeID,RecordDate"

	set rs=conn.execute(strSQL)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發入案資料審核清冊</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="100%" border="0">
	<tr>
		<td align="center"><strong>入案資料審核清冊</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" border="2" cellpadding="0" cellspacing="0">
				<tr>
					<td align="right">單號</td>
					<td align="right">入案訊息</td>
					<td align="right">證號</td>
					<td align="right">車號</td>
					<td align="right">簡式車種</td>
					<td align="right"><B>詳細車種</B></td>
<!--					<td align="right">顏色</td>-->
					<td align="right">駕駛人</td>
					<td align="right">駕駛地址</td>
					<td align="right">車主</td>
					<td align="right">車主地址</td>
					<td align="right">駕駛人(補正)</td>
					<td align="right">駕駛地址(補正)</td>
					<td align="right">車主(補正)</td>
					<td align="right">車主地址(補正)</td>
					<td align="right">違規日期</td>
					<td align="right">違規時間</td>
					<td align="right">違規地點</td>
					<td align="right">違規法條</td>
					<td align="right">填單日</td>
					<td align="right">應到案日</td>
					<td align="right">應到案處所</td>					
					<td align="right">舉發單位</td>
					<td align="right">舉發人</td>
					<td align="right">代保管物件</td>
					<td align="right">簽收狀態</td>
					<td align="right">限速(限重)</td>
					<td align="right">實速(實重)</td>
				</tr>
				<%
					filecnt=0
					while Not rs.eof
						Message="":Sys_DCIRETURNCARTYPE="":fastring="":Sys_CarColor="":Sys_DCIReturnStation=rs("DCIReturnStation")
						if rs("BillTypeID")<>"3" then
							strSQL="select * from DciLog where BillSN="&rs("SN")&" and ExChangetypeID='W'"
							set rsfound=conn.execute(strSQL)
							if not rsfound.eof then

								strSQL="select DCIReturn,StatusContent from DCIReturnStatus where DCIReturn='"&rsfound("DCIReturnStatusID")&"'"
								set rsdci=conn.execute(strSQL)
								if Not rsdci.eof then
									if trim(Message)<>"" then Message=trim(Message)&"<br>"
									Message=trim(Message)&rsdci("DCIReturn")&". "&rsdci("StatusContent")
								end if
								rsdci.close

								strSQL="select DCIReturn,StatusContent from DCIReturnStatus where DciActionID='WE' and DCIReturn='"&rsfound("DCIERRORCARDATA")&"'"
								set rsdci=conn.execute(strSQL)
								if Not rsdci.eof then
									if trim(Message)<>"" then Message=trim(Message)&"<br>"
									Message=trim(Message)&rsdci("DCIReturn")&". "&rsdci("StatusContent")
								end if
								rsdci.close

								strSQL="select DCIReturn,StatusContent from DCIReturnStatus where DciActionID='WE' and DCIReturn='"&rsfound("DCIERRORIDDATA")&"'"
								set rsdci=conn.execute(strSQL)
								if Not rsdci.eof then
									if trim(Message)<>"" then Message=trim(Message)&"<br>"
									Message=trim(Message)&rsdci("DCIReturn")&". "&rsdci("StatusContent")
								end if
								rsdci.close
							end if
							rsfound.close

							strsql="select * from DCICODE where ID='"&rs("DCIRETURNCARTYPE")&"' and TypeID=5"
							
							set cartype=conn.execute(strsql)
							if Not cartype.eof then Sys_DCIRETURNCARTYPE=trim(cartype("Content"))
							cartype.close

							Sys_DciReturnCarColor=trim(rs("DciReturnCarColor"))

							if len(trim(rs("DciReturnCarColor")))>1 then Sys_DciReturnCarColor=left(trim(rs("DciReturnCarColor")),1)&","&right(trim(rs("DciReturnCarColor")),1)

							Sys_CarColorID=split(Sys_DciReturnCarColor&"",",")
							for y=0 to ubound(Sys_CarColorID)
								if trim(Sys_CarColor)<>"" then Sys_CarColor=Sys_CarColor&","

								if trim(Sys_CarColorID(y))<>"" and not isnull(Sys_CarColorID(y)) then
									strColor="select Content from DCICode where TypeID=4 and ID='"&Sys_CarColorID(y)&"'"
									set rscolor=conn.execute(strColor)
									if not rscolor.eof then
										Sys_CarColor=Sys_CarColor&trim(rscolor("Content"))
									end if
									rscolor.close
								end if
							next

							strSql="select DCISTATIONID,DCISTATIONNAME,StationTel,StationID from Station where DCIStationID='"&rs("DCIReturnStation")&"'"
							set rssta=conn.execute(strSql)
							if Not rssta.eof then Sys_DCIReturnStation=trim(rssta("DCISTATIONNAME"))
							rssta.close

						end if

						strSql="select b.Content from BILLFASTENERDETAIL a,DCICode b where a.FASTENERTYPEID=b.ID and b.TypeID=6 and a.BIllSN="&trim(rs("SN"))&" and a.CarNo='"&trim(rs("CarNo"))&"'"
						set rsfast=conn.execute(strsql)
						fastring=""
						while Not rsfast.eof
							if trim(fastring)<>"" then fastring=fastring&","
							fastring=fastring&rsfast("Content")
							rsfast.movenext
						wend
						rsfast.close
						
						response.write "<tr>"
						response.write "<td align=""right"">"
						Response.Write rs("BillNo")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write Message
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("DriverID")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("CarNo")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("CarSimpleID")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right""><B>"
						Response.Write Sys_DCIRETURNCARTYPE
						Response.Write "&nbsp;</B></td>"

'						response.write "<td align=""right"">"
'						Response.Write Sys_CarColor
'						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("Driver")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("DriverZipName")&replace(rs("DriverAddress")&"",rs("DriverZipName")&"","")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("Owner")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("OwnerZipName")&replace(rs("OwnerAddress")&"",rs("OwnerZipName")&"","")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("DriverFix")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("DriverZipNameFix")&replace(rs("DriverAddressFix")&"",rs("DriverZipNameFix")&"","")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("OwnerFix")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("OwnerZipNameFix")&replace(rs("OwnerAddressFix")&"",rs("OwnerZipNameFix")&"","")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write gInitDT(trim(rs("IllegalDate")))
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write right("00"&hour(rs("IllegalDate")),2)&":"&right("00"&Minute(rs("IllegalDate")),2)
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("IllegalAddress")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("Rule")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write gInitDT(trim(rs("BillFillDate")))
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write gInitDT(trim(rs("DeallineDate")))
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write Sys_DCIReturnStation
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("BillUnitName")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("BillMem")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write fastring
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("SignType")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("RuleSpeed")
						Response.Write "&nbsp;</td>"

						response.write "<td align=""right"">"
						Response.Write rs("IllegalSpeed")
						Response.Write "&nbsp;</td>"

						response.write "</tr>"
						rs.movenext
					wend
				rs.close
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td align="center">
			前揭違規案件業已建檔傳送資料庫，請查核無誤後，於603表蓋章，本件留存備查。
		</td>
	</tr>
</table>
</body>
</html>
<%conn.close%>