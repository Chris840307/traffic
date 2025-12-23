<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include file="sqlDCIExchangeData.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style1 {
	font-size:11pt; 
	font-weight: bold;
	font-family: "標楷體";
}
.style2 {
	font-size:11pt; 
}
.style3 {
	font-size:11pt; 
	font-weight: bold;
}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:20px;
	font-family: "標楷體";
}
-->
</style>
<title>舉發單綜合查詢</title>
<% Server.ScriptTimeout = 6800 %>
<%
function getDciCode(code)
	  if code="00" then
		getDciCode="入案  未寫入資料庫"
	  elseif code="Y" Then
		getDciCode="入案  寫入資料庫"
	  elseif code="N" then
		getDciCode="入案	未寫入資料庫"
	  elseif code="S" then
		getDciCode="結案	違規人已經繳費"
	  elseif code="L" then
		getDciCode="入案	已經入案過"
	  elseif code="n" then
		getDciCode="入案	監理單位已經入案"
	  elseif code="0" then
		getDciCode="入案	正常"
	  elseif code="1" then
		getDciCode="入案錯誤車號不全"
	  elseif code="2" then
		getDciCode="入案錯誤	扣件不符"
	  elseif code="3" then
		getDciCode="入案錯誤	車號不正確"
	  elseif code="4" then
		getDciCode="入案錯誤	證號不正確"
	  elseif code="5" then
		getDciCode="入案錯誤	處所不明"
	  elseif code="6" then
		getDciCode="入案錯誤	違警案件"
	  elseif code="7" then
		getDciCode="入案錯誤	未簽名"
	  elseif code="8" then
		getDciCode="入案錯誤	未通知違規人"
	  elseif code="9" then
		getDciCode="入案錯誤	時間不明確"
	  elseif code="a" then
		getDciCode="入案錯誤	事實不明確"
	  elseif code="b" then
		getDciCode="入案錯誤	欠缺駕照證號"
	  elseif code="c" then
		getDciCode="入案錯誤	公文移轉"
	  elseif code="d" then
		getDciCode="入案錯誤	碰結案"
	  elseif code="e" then
		getDciCode="入案錯誤	非管轄碰結案"
	  elseif code="f" then
		getDciCode="入案錯誤	碰未結"
	  elseif code="g" then
		getDciCode="入案錯誤	移出"
	  elseif code="h" then
		getDciCode="入案錯誤	舉類錯誤"
	  elseif code="i" then
		getDciCode="入案錯誤	攔停需指定所站"
	  elseif code="j" then
		getDciCode="入案錯誤	單號不足9位"
	  elseif code="k" then
		getDciCode="入案錯誤	攔停車駕條款一起"
	  elseif code="l" then
		getDciCode="入案錯誤	無此單號剔退"
	  elseif code="m" then
		getDciCode="入案錯誤	條款與車別不符"
	  elseif code="z" then
		getDciCode="入案錯誤	道安已完成記點"
	  elseif code="A" then
		getDciCode="入案錯誤	條款錯誤"
	  elseif code="B" then
		getDciCode="入案錯誤	車籍無記錄"
	  elseif code="C" then
		getDciCode="入案錯誤	駕籍無紀錄"
	  elseif code="D" then
		getDciCode="入案錯誤	過戶前案"
	  elseif code="E" then
		getDciCode="入案錯誤	繳註銷前案"
	  elseif code="F" then
		getDciCode="入案錯誤	繳註銷後案"
	  elseif code="G" then
		getDciCode="入案錯誤	吊扣銷中案"
	  elseif code="H" then
		getDciCode="入案錯誤	證號重號剔退"
	  elseif code="I" then
		getDciCode="入案錯誤	達記點吊扣"
	  elseif code="J" then
		getDciCode="入案錯誤	達記點吊銷"
	  elseif code="K" then
		getDciCode="入案錯誤	單號+車號重覆"
	  elseif code="L" then
		getDciCode="入案錯誤	重覆入銷案剔退"
	  elseif code="M" then
		getDciCode="入案錯誤	無照駕駛"
	  elseif code="N" then
		getDciCode="入案錯誤	未知,找不到"
	  elseif code="O" then
		getDciCode="入案錯誤車駕非管轄"
	  elseif code="P" then
		getDciCode="入案錯誤	照類不符"
	  elseif code="Q" then
		getDciCode="入案錯誤	前車號違規"
	  elseif code="q" then
		getDciCode="入案錯誤	已定期換牌"
	  elseif code="S" then
		getDciCode="入案錯誤	非管轄"
	  elseif code="T" then
		getDciCode="入案錯誤	問題車牌"
	  elseif code="U" then
		getDciCode="入案錯誤	未異動"
	  elseif code="V" then
		getDciCode="入案錯誤	失竊註銷"
	  elseif code="X" then
		getDciCode="入案錯誤	未新增"
	  elseif code="Y" then
		getDciCode="入案錯誤	資料庫錯誤"
	  elseif code="Z" then
		getDciCode="入案錯誤	道安已完成開單"
	  elseif code="*" then
		getDciCode="入案錯誤	刪除不入案"
	  elseif code="y" then
		getDciCode="入案錯誤	行照過期"
	  elseif code="x" then
		getDciCode="入案錯誤	駕照過期"
	  elseif code="R" then
		getDciCode="入案錯誤	非管轄"
	  else
		getDciCode="&nbsp;"
	  end if

	end Function
	
	function GetDate(tDate)
		if len(tDate)=7 then
			GetDate=left(tDate,3)&"年"& mid(tDate,4,2)&"月"& Right(tDate,2)&"日"
		else
			GetDate="&nbsp;"
		end if
	end function

	function GetTime(ttime)
	  W=""
	  H=""
	  H=left(ttime,2)
	  N=right(ttime,2)
	  if cdbl(H)=12 then
		W="中午"
	  elseif cdbl(H)<6  then
		W="凌晨"
	  elseif cdbl(H)>5 and cdbl(H)<12 then
		W="早上"
	  elseif cdbl(H)>12 and cdbl(H)<18 then
		W="下午"
	  elseif cdbl(H)>17 then
		W="晚上"
	  end if

	  SH=0

	  if H>12 then SH=cdbl(H)-12 else SH=H

		if len(ttime)=4 then
			GetTime=W&" "&right("00"&SH,2)&"點"&N&"分"
		else
			GetTime="&nbsp;"
		end if
	end Function

	function getDciCodeN(code)
	  if code="S" then
		getDciCodeN="送達註記	成功註記"
	  elseif code="N" then
		getDciCodeN="送達註記	找不到此筆資料"
	  elseif code="n" then
		getDciCodeN="送達註記	已經結案"
	  elseif code="k" then
		getDciCodeN="送達註記	已送達不可做未送達註記"
	  elseif code="h" then
		getDciCodeN="送達註記	已開裁決書"
	  elseif code="B" then
		getDciCodeN="送達註記	無此車號/無此證號"
	  elseif code="E" then
		getDciCodeN="送達註記	日期錯誤"
	  else
		getDciCodeN="&nbsp;"
	  end if
	end Function
	'=====
	function chkBillType(BillTypeID)
		if trim(BillTypeID) <> "" then
			Select Case  trim(BillTypeID)
				Case "1" chkBillType="攔停"
				Case "2" chkBillType="逕舉"
				Case "3" chkBillType="逕舉手開單" 
				Case "4" chkBillType="拖吊" 
				Case "5" chkBillType="慢車行人"   
				Case "6" chkBillType="肇事"   
				Case "7" chkBillType="掌-攔停"   
				Case "8" chkBillType="掌-行人"   
				Case "9" chkBillType="掌電拖吊"   
				Case "H" chkBillType="人工移送"   
				Case "M" chkBillType="郵寄處理"   
				Case "N" chkBillType="攔停逕行(未開單)"   
				Case "D" chkBillType="註銷"   
				Case "R" chkBillType="單退"   
				Case "V" chkBillType="掌電拖吊(補開單)"   
			end select       
		end if 
	end Function
            
	'查詢車種 
	function GetCarType(CarTypeID)
		if trim(CarTypeID) <> "" then
			Select Case  trim(CarTypeID)
				Case "1" GetCarType="自大客車"
				Case "2" GetCarType="自大貨車"
				Case "3" GetCarType="自小客(貨)" 
				Case "4" GetCarType="營大客車" 
				Case "5" GetCarType="營大貨車"   
				Case "6" GetCarType="營小貨車"   
				Case "7" GetCarType="營小客車"   
				Case "8" GetCarType="租賃小客"   
				Case "9" GetCarType="遊覽客車"   
				Case "A" GetCarType="營交通車"   
				Case "B" GetCarType="貨櫃曳引"   
				Case "C" GetCarType="自用拖車"   
				Case "D" GetCarType="營業拖車"   
				Case "E" GetCarType="外賓小客"   
				Case "F" GetCarType="外賓大客"   
				Case "H" GetCarType="普通重機"    
				Case "L" GetCarType="輕機"   
				Case "p" GetCarType="併裝車"   
				Case "x" GetCarType="動力機械"      
				Case "Y" GetCarType="租賃小貨車"
				Case "W" GetCarType="自小客"
				Case "V" GetCarType="自小貨"
				Case "G" GetCarType="大型重機250CC"
				Case "Q" GetCarType="大型重機550CC"    
			end select       
		end if 
	end Function
            
	'檢查簽收情形
	Function chksigner(value)
		 if trim(value)="0" then
			chksigner="正常"
		 elseif  trim(value)="1" then
			chksigner="拒簽"
		 elseif  trim(value)="2" then
			chksigner="拒收"
		 elseif  trim(value)="3" then
			chksigner="拒簽拒收"
		 end if     
	end function  
	
	'檢查保險證
	Function chkissure(value)
		 if trim(value)="0" then
			chkissure="正常"
		 elseif  trim(value)="1" then
			chkissure="未帶"
		 elseif  trim(value)="2" then
			chkissure="肇事且未帶"
		 elseif  trim(value)="3" then
			chkissure="逾期且未保"
		 elseif  trim(value)="4" then
			chkissure="肇事且逾期或未帶"
		 end if     
	end Function
            
	'組地址字串
	function composeAddress(Address,lane,alley,No,Dash)
		composeAddress=""
		composeAddress=Address
		if  trim(lane) <> "" then
			composeAddress=composeAddress & lane & "巷"
		end if
		if  trim(alley) <> "" then
			composeAddress=composeAddress & alley & "弄"
		end if
		if  trim(No) <> "" then
			composeAddress=composeAddress & No & "號"
		end if
		if  trim(Dash) <> "" then
			composeAddress=composeAddress & "之" & Dash
		end if
	end Function

%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
<%
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	strSQLTemp=""
	strQry="[快速查詢]"
	if trim(request("BillNo"))<>"" Then
		If strQry="[快速查詢]" then
			strQry=strQry&"BillNo="&Trim(request("BillNo"))
		Else
			strQry=strQry&",BillNo="&Trim(request("BillNo"))
		End if
		strSQLTemp=strSQLTemp&" and a.BillNo='"&trim(request("BillNo"))&"'"
	end if
	if trim(request("CarNo"))<>"" Then
		If strQry="[快速查詢]" then
			strQry=strQry&"CarNo="&Trim(request("CarNo"))
		Else
			strQry=strQry&",CarNo="&Trim(request("CarNo"))
		End if
		strSQLTemp=strSQLTemp&" and a.CarNo='"&trim(request("CarNo"))&"'"
	end if
	if trim(request("IllegalName"))<>"" Then
		If strQry="[快速查詢]" then
			strQry=strQry&"IllegalName="&Trim(request("IllegalName"))
		Else
			strQry=strQry&",IllegalName="&Trim(request("IllegalName"))
		End If
		BillNoTemp=""
		strName="select a.BillNo from Billbase a,BillBaseDciReturn b where a.BillNo=b.Billno and a.CarNo=b.CarNo " &_
			" and (b.Owner='"&trim(request("IllegalName"))&"' or b.Driver='"&trim(request("IllegalName"))&"') " &_
			" and b.ExchangeTypeid='W'"
		Set rsName=conn.execute(strName)
		While Not rsName.Eof
			If BillNoTemp="" Then
				BillNoTemp="'"&Trim(rsName("BillNo"))&"'"
			Else
				BillNoTemp=BillNoTemp&",'"&Trim(rsName("BillNo"))&"'"
			End If 

			rsName.movenext
		wend
		rsName.close
		Set rsName=Nothing 

		strName2="select a.BillNo from traffic2.Billbase a,traffic2.BillBaseDciReturn b " &_
			" where a.BillNo=b.Billno and a.CarNo=b.CarNo " &_
			" and (b.Owner='"&trim(request("IllegalName"))&"' or b.Driver='"&trim(request("IllegalName"))&"') "
		Set rsName2=conn.execute(strName2)
		While Not rsName2.Eof
			If BillNoTemp="" Then
				BillNoTemp="'"&Trim(rsName2("BillNo"))&"'"
			Else
				BillNoTemp=BillNoTemp&",'"&Trim(rsName2("BillNo"))&"'"
			End If 

			rsName2.movenext
		wend
		rsName2.close
		Set rsName2=Nothing 

		strName3="select FSEQ from traffic3.FMASTER where " &_
			" IName='"&trim(request("IllegalName"))&"' "
		Set rsName3=conn.execute(strName3)
		While Not rsName3.Eof
			If BillNoTemp="" Then
				BillNoTemp="'"&Trim(rsName3("FSEQ"))&"'"
			Else
				BillNoTemp=BillNoTemp&",'"&Trim(rsName3("FSEQ"))&"'"
			End If 

			rsName3.movenext
		wend
		rsName3.close
		Set rsName3=Nothing 

		strName4="select FSEQ from traffic4.FMASTER where " &_
			" IName='"&trim(request("IllegalName"))&"' "
		Set rsName4=conn.execute(strName4)
		While Not rsName4.Eof
			If BillNoTemp="" Then
				BillNoTemp="'"&Trim(rsName4("FSEQ"))&"'"
			Else
				BillNoTemp=BillNoTemp&",'"&Trim(rsName4("FSEQ"))&"'"
			End If 

			rsName4.movenext
		wend
		rsName4.close
		Set rsName4=Nothing 

		if BillNoTemp="" then
			BillNoTemp="'NoDate'"
		end if 
		strSQLTemp=strSQLTemp&" and (a.BillNo in ("&BillNoTemp&"))"
	end if
	if trim(request("IllegalID"))<>"" Then
		If strQry="[快速查詢]" then
			strQry=strQry&"IllegalID="&Trim(request("IllegalID"))
		Else
			strQry=strQry&",IllegalID="&Trim(request("IllegalID"))
		End if
		strSQLTemp=strSQLTemp&" and (a.DriverID='"&trim(request("IllegalID"))&"')"
	end If

	if trim(request("BillSn"))<>"" then
		strSQLTemp=strSQLTemp&" and a.SN='"&trim(request("BillSn"))&"'"
	end If
	'response.write BillNoTemp
	'response.end
	strOld1="select * from BILLBASEVIEW_OLD2 a where 1=1 " &strSQLTemp
	'response.write strOld1
	'response.end
	Cnt=0
	NoCase=0
	NewCase=0
	'on error resume next  
	ConnExecute strQry,356
	set rsOld1=conn.execute(strOld1)
	If Not rsOld1.Bof Then
		rsOld1.MoveFirst 
	End if
	While Not rsOld1.Eof
		
		if Trim(rsOld1("CityFlag"))="1" Or Trim(rsOld1("CityFlag"))="2" Or isnull(rsOld1("CityFlag")) then
			DBUser="traffic"
		elseif Trim(rsOld1("CityFlag"))="3" Or Trim(rsOld1("CityFlag"))="4" then
			DBUser="traffic2"
		elseif trim(rsOld1("CityFlag"))="5" then
			DBUser="traffic3"
		elseif trim(rsOld1("CityFlag"))="6" then
			DBUser="traffic4"
		Else
			DBUser="traffic4"
		end If
		'response.write trim(rsOld1("CityFlag"))&"//"&DBUser
		'台南縣市系統==========================================================================================
		If Trim(rsOld1("CityFlag"))="1" Or Trim(rsOld1("CityFlag"))="2" Or Trim(rsOld1("CityFlag"))="3" Or Trim(rsOld1("CityFlag"))="4" Or isnull(rsOld1("CityFlag")) Then
			strSQLTemp1=" and a.sn="&Trim(rsOld1("Sn"))
			strSQL="Select a.BillNo,a.Sn,a.CarNo,a.BillTypeID,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.MemberStation,a.EquipMentID" &_
			",a.Recorddate,a.RecordMemberID,a.RecordStateID,a.IllegalDate,a.BillMemID1,a.BillMem1" &_
			",a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.RuleSpeed,a.IllegalSpeed" &_
			",a.BillMemID4,a.BillMem4,a.RuleVer,a.IllegalAddressID,a.IllegalAddress,a.BillFillDate,a.BillUnitID,a.SignType" &_
			",a.driveraddress,a.driverzip,a.owner,a.ownerzip,a.owneraddress" &_
			",a.DealLineDate,a.Note,a.CarSimpleID,a.CarAddID,a.ImageFileName from "&DBUser&".BillBase a,"&DBUser&".BillBaseDciReturn b" &_
			" where ((a.RecordStateID<>-1 and a.BillStatus='0')" &_
			" or a.BillStatus<>'0') and a.BillNo=b.BillNo and a.CarNo=b.CarNo and b.ExChangeTypeID='W'" &_
			" and b.Status in ('Y','S','n','L') "&strSQLTemp1
			'response.write strSQL
			set rs1=conn.execute(strSQL)
			If Not rs1.Bof  then
				if BillSnTmp="" then
				BillSnTmp=trim(rs1("Sn"))
			else
				BillSnTmp=BillSnTmp&","&trim(rs1("Sn"))
			end if
			If Not rs1.eof Then
			
			StationNameBillBase=trim(rs1("MemberStation"))
			'--------------------------------------BILLBASEDCIRETURN------------------------------------
		'先查有沒有車籍查尋的資料 沒有的話再用入案資料
			StationName=""	'到案處所
			IllegalMemID=""	'違規人證號
			IllegalMem=""	'違規人姓名
			IllegalAddress=""	'違規人地址
			OwnerName=""	'車主姓名
			OwnerAddress=""	'車主地址
			DciCarTypeID=""	'詳細車種代碼
			DciCarType=""	'詳細車種
			strDciB="select a.* from "&DBUser&".BillBaseDciReturn a,"&DBUser&".DciReturnStatus b" &_
			" where a.ExchangeTypeID=b.DciActionID and a.Status=b.DciReturn" &_
			" and (a.BillNo='"&trim(rs1("BillNo"))&"' or a.BillNo is Null)" &_
			" and a.CarNo='"&trim(rs1("CarNo"))&"'" &_
			" and b.DciReturnStatus=1 and ExchangeTypeID='W'"
			set rsDciB=conn.execute(strDciB)
			if not rsDciB.eof then

	'			if sys_City<>"台中市" then
	'				OwnerZipName=""
	'				DriverZipName=""
	'			else
					strOZip="select ZipName from "&DBUser&".Zip where ZipID='"&trim(rsDciB("OwnerZip"))&"'"
					set rsOZip=conn.execute(strOZip)
					if not rsOZip.eof Then
						If CDbl(Year(rs1("IllegalDate")))<2011 then
							OwnerZipName=ChangeOldCity(trim(rsDciB("OwnerZip")),trim(rsOZip("ZipName")))
						Else
							OwnerZipName=trim(rsOZip("ZipName"))
						End If					
					end if
					rsOZip.close
					set rsOZip=nothing

					strDZip="select ZipName from "&DBUser&".Zip where ZipID='"&trim(rsDciB("DriverHomeZip"))&"'"
					set rsDZip=conn.execute(strDZip)
					if not rsDZip.eof Then
						If CDbl(Year(rs1("IllegalDate")))<2011 then
							DriverZipName=ChangeOldCity(trim(rsDciB("DriverHomeZip")),trim(rsDZip("ZipName")))
						Else
							DriverZipName=trim(rsDZip("ZipName"))
						End If								
					end if
					rsDZip.close
					set rsDZip=nothing
	'			end if
				if trim(rs1("BillTypeID"))="2" then
					StationName=trim(rsDciB("DciReturnStation"))
				else
					StationName=trim(rs1("MemberStation"))
				end if
				OwnerName=trim(rsDciB("Owner"))
				OwnerAddress=trim(rsDciB("OwnerZip"))&" "&OwnerZipName&trim(rsDciB("OwnerAddress"))
				DciCarTypeID=trim(rsDciB("DciReturnCarType"))
				if trim(rs1("BillTypeID"))="1" then
					IllegalMemID=trim(rsDciB("DriverID"))
					IllegalMem=trim(rsDciB("Driver"))
					IllegalAddress=trim(rsDciB("DriverHomeZip"))&DriverZipName&" "&trim(rsDciB("DriverHomeAddress"))
				else
					if sys_City="台中市" then
						IllegalMemID=""
						IllegalMem=""
						IllegalAddress=""
					else
						IllegalMemID=trim(rsDciB("OwnerID"))
						IllegalMem=trim(rsDciB("Owner"))
						IllegalAddress=trim(rsDciB("OwnerZip"))&" "&OwnerZipName&trim(rsDciB("OwnerAddress"))
					end if
				end If
			else
				if (sys_City="高雄市" or sys_City=ApconfigureCityName) and trim(rs1("BillTypeID"))="1" then
					strDciA1="select a.* from "&DBUser&".BillBaseDciReturn a,"&DBUser&".DciReturnStatus b" &_
					" where a.ExchangeTypeID=b.DciActionID and a.Status=b.DciReturn" &_
					" and (a.BillNo='"&trim(rs1("BillNo"))&"' or a.BillNo is Null)" &_
					" and a.CarNo='"&trim(rs1("CarNo"))&"'" &_
					" and b.DciReturnStatus=1 and ExchangeTypeID='A'"
					set rsDciA1=conn.execute(strDciA1)
					if not rsDciA1.eof then
						strOZip1="select ZipName from "&DBUser&".Zip where ZipID='"&trim(rsDciA1("OwnerZip"))&"'"
						set rsOZip1=conn.execute(strOZip1)
						if not rsOZip1.eof Then
							If CDbl(Year(rs1("IllegalDate")))<2011 then
								OwnerZipName=ChangeOldCity(trim(rsDciA1("OwnerZip")),trim(rsOZip1("ZipName")))
							Else
								OwnerZipName=trim(rsOZip1("ZipName"))
							End If								
						end if
						rsOZip1.close
						set rsOZip1=nothing

						strDZip1="select ZipName from "&DBUser&".Zip where ZipID='"&trim(rsDciA1("DriverHomeZip"))&"'"
						set rsDZip1=conn.execute(strDZip1)
						if not rsDZip1.eof then		
							If CDbl(Year(rs1("IllegalDate")))<2011 then
								DriverZipName=ChangeOldCity(trim(rsDciA1("DriverHomeZip")),trim(rsDZip1("ZipName")))
							Else
								DriverZipName=trim(rsDZip1("ZipName"))
							End If								
						end if
						rsDZip1.close
						set rsDZip1=nothing

						OwnerName=trim(rsDciA1("Owner"))
						If Not IsNull(rsDciA1("OwnerAddress")) Then
							OwnerAddress=trim(rsDciA1("OwnerZip"))&" "&OwnerZipName&replace(replace(trim(rsDciA1("OwnerAddress")),"臺","台"),OwnerZipName,"")
						Else
							OwnerAddress=trim(rsDciA1("OwnerZip"))&" "&OwnerZipName&trim(rsDciA1("OwnerAddress"))
						End If 
						DciCarTypeID=trim(rsDciA1("DciReturnCarType"))
						IllegalMemID=trim(rsDciA1("DriverID"))
						IllegalMem=trim(rsDciA1("Driver"))
						If Not IsNull(rsDciA1("DriverHomeAddress")) Then
							IllegalAddress=trim(rsDciA1("DriverHomeZip"))&DriverZipName&" "&replace(replace(trim(rsDciA1("DriverHomeAddress")),"臺","台"),DriverZipName,"")
						Else
							IllegalAddress=trim(rsDciA1("DriverHomeZip"))&DriverZipName&" "&trim(rsDciA1("DriverHomeAddress"))
						End If 
					end if
					rsDciA1.close
					set rsDciA1=nothing
				end if
			end if
			rsDciB.close
			set rsDciB=Nothing

			strCarType="select Content from "&DBUser&".DciCode where TypeID=5 and ID='"&DciCarTypeID&"'"
			set rsCarType=conn.execute(strCarType)
			if not rsCarType.eof then
				DciCarType=trim(rsCarType("Content"))
			end if
			rsCarType.close
			set rsCarType=nothing

			DciA_Name=""	'廠牌
			DciColor=""		'顏色
			DciDriverHomeAddress=""	'車主戶籍地址
			DciIDstatus="" '行駕照狀態
			'if sys_City="台東縣" Or sys_City="高雄市" Or sys_City="高雄縣" then
				strDciA="select * from "&DBUser&".BillBaseDciReturn where (BillNo='"&trim(rs1("BillNo"))&"' or BillNo is Null)" &_
						" and CarNo='"&trim(rs1("CarNo"))&"'" &_
						" and ExchangeTypeID='A' and Status='S'"
				set rsDciA=conn.execute(strDciA)
				if not rsDciA.eof then
					OwnerCID=trim(rsDciA("OwnerID"))
					If trim(rsDciA("A_Name"))<>"" And Not IsNull(rsDciA("A_Name")) then
						DciA_Name=trim(rsDciA("A_Name"))
					End If
					if trim(rsDciA("DCIReturnCarColor"))<>"" and not isnull(rsDciA("DCIReturnCarColor")) then
						ColorLen=cint(Len(rsDciA("DCIReturnCarColor")))
						for Clen=1 to ColorLen
							colorID=mid(rsDciA("DCIReturnCarColor"),Clen,1)
							strColor="select * from "&DBUser&".DCIcode where TypeID=4 and ID='"&trim(colorID)&"'"
							set rsColor=conn.execute(strColor)
							if not rsColor.eof then
								DciColor=DciColor & trim(rsColor("Content"))
							end if
							rsColor.close
							set rsColor=nothing
						next
					end If
					If trim(rsDciA("DriverHomeAddress"))<>"" And Not isnull(rsDciA("DriverHomeAddress")) then
						DciDriverHomeAddress=trim(rsDciA("DriverHomeZip"))&trim(rsDciA("DriverHomeAddress"))
					End If
					If trim(rsDciA("DciCounterID"))<>"" And Not isnull(rsDciA("DciCounterID")) then
						If trim(rsDciA("DciCounterID"))="x" Then
							DciIDstatus="駕照過期"
						ElseIf trim(rsDciA("DciCounterID"))="y" Then
							DciIDstatus="行照過期"
						ElseIf trim(rsDciA("DciCounterID"))="v" Then
							DciIDstatus="行駕照過期"
						End If 
					End If
				end if
				rsDciA.close
				set rsDciA=Nothing
				
				If sys_City="高雄市" Then '如果Billbase有寫以billbase為主
				If trim(rs1("BillTypeID"))="2" Then
					If Not isnull(rs1("driveraddress")) then
						DciDriverHomeAddress=trim(rs1("DriverZip"))&" "&trim(rs1("driveraddress"))
					End If
				End If 
			End If

			
			CaseInDate=""	'入案日期
			CaseStatus=""	'入案狀態
			DciFileName=""	'入案檔名
			DciBatchNumber=""	'入案批號
			DciForfeit1=""	'罰金1
			DciForfeit2=""	'罰金2
			DciForfeit3=""	'罰金3

			strCaseIn="select a.*,c.* from "&DBUser&".BillBaseDciReturn a,"&DBUser&".DciReturnStatus b,"&DBUser&".DciLog c" &_
					" where a.ExchangeTypeID=b.DciActionID and a.Status=b.DciReturn" &_
					" and a.ExchangeTypeID=c.ExchangeTypeID and a.Status=c.DciReturnStatusID" &_
					" and a.BillNo=c.BillNo and a.CarNo=c.CarNo" &_
					" and c.BillSN='"&trim(rs1("SN"))&"' " &_
					" and a.ExchangeTypeID='W'" &_
					" order by c.ExchangeDate Desc"
			set rsCaseIn=conn.execute(strCaseIn)
			if not rsCaseIn.eof then
				CaseInDate=trim(rsCaseIn("DciCaseInDate"))
				if trim(rsCaseIn("STATUS"))<>"" and not isnull(rsCaseIn("STATUS")) then
					strStuts="select StatusContent from "&DBUser&".DciReturnStatus where DciActionID='W' and DciReturn='"&trim(rsCaseIn("STATUS"))&"'"
					set rsStuts=conn.execute(strStuts)
					if not rsStuts.eof then
						CaseStatus=trim(rsStuts("StatusContent"))
					end if
					rsStuts.close
					set rsStuts=nothing
				else
					CaseStatus="未處理"
				end if
				DciFileName=trim(rsCaseIn("FileName"))
				DciBatchNumber=trim(rsCaseIn("BatchNumber"))
				If Trim(rsCaseIn("Forfeit1"))<>"0" Then
					DciForfeit1=Trim(rsCaseIn("Forfeit1"))
				End If
				If Trim(rsCaseIn("Forfeit2"))<>"0" Then
					DciForfeit2=Trim(rsCaseIn("Forfeit2"))
				End If
				If Trim(rsCaseIn("Forfeit3"))<>"0" Then
					DciForfeit3=Trim(rsCaseIn("Forfeit3"))
				End if
			else
				if sys_City<>"台中市" then
					CaseStatus="未上傳"
				else
					CaseStatus="&nbsp;"
				end if 
			end if
			rsCaseIn.close
			set rsCaseIn=nothing

		'-----------------------------------BillMailHistory-------------------------------------
			StoreAndSendFlag=0	'是否做過寄存

			MailDate=""	'郵寄日期
			MailNumber=""	'郵寄序號
			MailStation=""	'寄存郵局
			GetFileName=""	'收受檔案
			GetBatchNumber=""	'收受批號
			GetStatus=""	'收受上傳狀態
			GetMailDate=""	'收受日期
			GetMailReason=""	'收受原因
			ReturnMailDate=""	'退回日期
			ReturnReason=""	'退件原因
			ReturnSendDate=""	'移送日期
			ReturnMailNumber=""	'退件郵寄序號
			ReturnSendMailDate=""	'退件郵寄日期
			StoreAndSendGovNumber=""	'寄存送達書號
			StoreAndSendEffectDate=""	'寄存送達日
			StoreAndSendEndDate=""	'寄存送達生效(完成)日
			OpenGovGovNumber=""	'公示送達書號
			OpenGovEffectDate=""	'公示送達生效日
			StoreAndSendDate=""	'二次送達日期
			StoreAndSendReason=""	'二次送達原因
			BillMailNo=""	'郵寄序號
			ReturnMailNo=""	'退件郵寄序號
			MailCheckNumber="" '郵局查詢號
			MailReturnCheckNumber="" '單退後投遞郵局查詢號
			StoreAndSendFinalMailDate=""	'送達證書郵寄日期
			SignMan=""	'簽收人
			'-------------------------------------------------------
			CancalSendDate=""   '撤銷送達日
			strCaseIn="select * from "&DBUser&".dcilog where " &_
								" BillSn=" & trim(rs1("SN")) & " and BillNo='"&trim(rs1("BillNo")) & "' and ExchangeTypeID='N' and ReturnMarkType='Y' and DCIRETURNSTATUSID='S'" 
			set rsCaseIn=conn.execute(strCaseIn)
			'response.write strCaseIn
			if not rsCaseIn.eof then
				CancalSendDate=gArrDT(trim(rsCaseIn("Exchangedate")))	
			end if
			rsCaseIn.close
			set rsCaseIn=nothing	
			'-------------------------------------------------------

			'檢查是單退還是收受
			strCheck="select count(*) as cnt from "&DBUser&".Dcilog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='7'"
			set rsCheck=conn.execute(strCheck)
			if not rsCheck.eof then
				if rsCheck("cnt")="0" then
					CheckFlag=0	'單退
				else
					CheckFlag=1	'收受
				end if
			end if
			rsCheck.close
			set rsCheck=nothing

			strMail="select * from "&DBUser&".BillMailHistory where BillSN='"&trim(rs1("SN"))&"'"
			set rsMail=conn.execute(strMail)
			if not rsMail.eof then
				if trim(rs1("BillTypeID"))="2" or (trim(rs1("BillTypeID"))="1" and trim(rs1("EquipMentID"))="1") then
					if trim(rsMail("MailDate"))<>"" and not isnull(rsMail("MailDate")) then
						MailDate=gArrDT(trim(rsMail("MailDate")))
					end if
				end if
				If sys_City="高雄市" Then
					MailNumber=trim(rsMail("MailChkNumber"))
				Else
					MailNumber=trim(rsMail("MailNumber"))
				End If 
				if CheckFlag=0 then
					if trim(rsMail("MAILRETURNDATE"))<>"" and not isnull(rsMail("MAILRETURNDATE")) then
						ReturnMailDate=gArrDT(trim(rsMail("MAILRETURNDATE")))
					end if
					if trim(rsMail("OpenGovMailReturnDate"))<>"" and not isnull(rsMail("OpenGovMailReturnDate")) then
						ReturnMailDate="&nbsp;"&gArrDT(trim(rsMail("OpenGovMailReturnDate")))
					end if
					GetMailDate=""
				else 
					'如果是收受. 這邊應該改成日期再 6/30前的要讀舊的欄位,支後讀這些欄未
					'台中市 6/30開始轉換 ,或是說如果signdate是空的就讀舊的. 						
					if trim(rsMail("SIGNDATE"))<>"" and not isnull(rsMail("SIGNDATE")) then
						GetMailDate=gArrDT(trim(rsMail("SIGNDATE")))
					else
						if trim(rsMail("MAILRETURNDATE"))<>"" and not isnull(rsMail("MAILRETURNDATE")) then
							GetMailDate=gArrDT(trim(rsMail("MAILRETURNDATE")))
						end if
					end if
					ReturnMailDate=""
				end if
				'退件or收受原因
				if CheckFlag=0 then
					'smith 20080626 暫時把收受注記誤寫入的部份排除掉
					if trim(rsMail("RETURNRESONID"))<>"" and not isnull(rsMail("RETURNRESONID")) and (rsMail("RETURNRESONID") <> "A") and (rsMail("RETURNRESONID") <> "B")and (rsMail("RETURNRESONID") <> "C")and (rsMail("RETURNRESONID") <> "D")  then
						strReturnReason="select Content from "&DBUser&".DciCode where TypeID=7 and ID='"&trim(rsMail("RETURNRESONID"))&"'"
						set rsRR=conn.execute(strReturnReason)
						if not rsRR.eof then
							ReturnReason=trim(rsRR("Content"))
						end if
						rsRR.close			
						set rsRR=nothing
					end if

					GetMailReason=""
					GetFileName=""
					GetBatchNumber=""
					if sys_City<>"台中市" then
						GetStatus="未上傳"
					else
						GetStatus="&nbsp;"
					end if 
				else
					'如果是收受. 這邊應該改成日期再 6/30前的要讀舊的欄位,支後讀這些欄未
					'台中市 6/30開始轉換 ,或是說如果SIGNRESONID是空的就讀舊的. 				
					if trim(rsMail("SIGNRESONID"))<>"" and not isnull(rsMail("SIGNRESONID")) then
						strReturnReason="select Content from "&DBUser&".DciCode where TypeID=7 and ID='"&trim(rsMail("SIGNRESONID"))&"'"
						set rsRR=conn.execute(strReturnReason)
						if not rsRR.eof then
							GetMailReason=trim(rsRR("Content"))
						end if
						rsRR.close
						set rsRR=nothing
					else
						if trim(rsMail("RETURNRESONID"))<>"" and not isnull(rsMail("RETURNRESONID")) then
							strReturnReason="select Content from "&DBUser&".DciCode where TypeID=7 and ID='"&trim(rsMail("RETURNRESONID"))&"'"
							set rsRR=conn.execute(strReturnReason)
							if not rsRR.eof then
								GetMailReason=trim(rsRR("Content"))
							end if
							rsRR.close
							set rsRR=nothing
						end if				
					end if
					ReturnReason=""
					if trim(rsMail("SignMan"))<>"" and not isnull(rsMail("SignMan")) then
						SignMan=trim(rsMail("SignMan"))
					end if
					strGet="select * from "&DBUser&".Dcilog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='7' order by ExchangeDate desc"
					set rsGet=conn.execute(strGet)
					if not rsGet.eof then
						GetFileName=trim(rsGet("FileName"))
						GetBatchNumber=trim(rsGet("BatchNumber"))
						if trim(rsGet("DciReturnStatusID"))<>"" and not isnull(rsGet("DciReturnStatusID")) then
							strGStuts="select StatusContent from "&DBUser&".DciReturnStatus where DciActionID='N' and DciReturn='"&trim(rsGet("DciReturnStatusID"))&"'"
							set rsGStuts=conn.execute(strGStuts)
							if not rsGStuts.eof then
								GetStatus=trim(rsGStuts("StatusContent"))
							end if
							rsGStuts.close
							set rsGStuts=nothing
						else
							GetStatus="未處理"
						end if
					end if
					rsGet.close
					set rsGet=nothing
				end if

					if trim(rsMail("OPENGOVRESONID"))<>"" and not isnull(rsMail("OPENGOVRESONID")) and trim(ReturnReason) = "" then				
						
							strReturnReason="select Content from "&DBUser&".DciCode where TypeID=7 and ID='"&trim(rsMail("OPENGOVRESONID"))&"'"
							set rsRR=conn.execute(strReturnReason)
							if not rsRR.eof then	
								ReturnReason=trim(rsRR("Content"))			
							end if				
						rsRR.close			
						set rsRR=nothing			
					end if		

				if trim(rsMail("MailStation"))<>"" and not isnull(rsMail("MailStation")) then
					MailStation=trim(rsMail("MailStation"))
				end if
				if trim(rsMail("SendOpenGovDocToStationDate"))<>"" and not isnull(rsMail("SendOpenGovDocToStationDate")) then
					ReturnSendDate=left(trim(rsMail("SendOpenGovDocToStationDate")),len(trim(rsMail("SendOpenGovDocToStationDate")))-4)&"-"&mid(trim(rsMail("SendOpenGovDocToStationDate")),len(trim(rsMail("SendOpenGovDocToStationDate")))-3,2)&"-"&mid(trim(rsMail("SendOpenGovDocToStationDate")),len(trim(rsMail("SendOpenGovDocToStationDate")))-1,2)
				end if
				ReturnMailNumber=trim(rsMail("StoreAndSendMailNumber"))
				if trim(rsMail("StoreAndSendSendDate"))<>"" and not isnull(rsMail("StoreAndSendSendDate")) then
					ReturnSendMailDate=gArrDT(trim(rsMail("StoreAndSendSendDate")))
				end if
				if trim(rsMail("STOREANDSENDGOVNUMBER"))<>"" and not isnull(rsMail("STOREANDSENDGOVNUMBER")) then
					StoreAndSendGovNumber=trim(rsMail("STOREANDSENDGOVNUMBER"))
				end if
				if trim(rsMail("STOREANDSENDEFFECTDATE"))<>"" and not isnull(rsMail("STOREANDSENDEFFECTDATE")) then
					StoreAndSendEffectDate=gArrDT(trim(rsMail("STOREANDSENDEFFECTDATE")))
				end if
				if trim(rsMail("StoreAndSendMailDate"))<>"" and not isnull(rsMail("StoreAndSendMailDate")) then
					StoreAndSendEndDate=gArrDT(trim(rsMail("StoreAndSendMailDate")))
				end if
				if trim(rsMail("OPENGOVNUMBER"))<>"" and not isnull(rsMail("OPENGOVNUMBER")) then
					OpenGovGovNumber=trim(rsMail("OPENGOVNUMBER"))
				end if
				if trim(rsMail("OPENGOVDATE"))<>"" and not isnull(rsMail("OPENGOVDATE")) then
					OpenGovEffectDate=gArrDT(trim(rsMail("OPENGOVDATE")))
					KSOpenGovEffDate=gArrDT(DateAdd("d",rsMail("OPENGOVDATE"),35))
				end if
				if trim(rsMail("OPENGOVRESONID"))<>"" and not isnull(rsMail("OPENGOVRESONID")) then
					strSReason="select Content from "&DBUser&".DciCode where TypeID=7 and ID='"&trim(rsMail("OPENGOVRESONID"))&"'"
					set rsSR=conn.execute(strSReason)
					if not rsSR.eof then
						OPENGOVRESONID=trim(rsSR("Content"))
					end if
					rsSR.close
					set rsSR=nothing
				end if	
				if trim(rsMail("STOREANDSENDMAILRETURNDATE"))<>"" and not isnull(rsMail("STOREANDSENDMAILRETURNDATE")) then
					StoreAndSendDate=gArrDT(trim(rsMail("STOREANDSENDMAILRETURNDATE")))
				end if
				if trim(rsMail("STOREANDSENDRETURNRESONID"))<>"" and not isnull(rsMail("STOREANDSENDRETURNRESONID")) then
					strSReason="select Content from "&DBUser&".DciCode where TypeID=7 and ID='"&trim(rsMail("STOREANDSENDRETURNRESONID"))&"'"
					set rsSR=conn.execute(strSReason)
					if not rsSR.eof then
						StoreAndSendReason=trim(rsSR("Content"))
					end if
					rsSR.close
					set rsSR=nothing
				end if
				if trim(rsMail("MailSeqNo1"))<>"" and not isnull(rsMail("MailSeqNo1")) then
					BillMailNo=trim(rsMail("MailSeqNo1"))
				end if
				if trim(rsMail("MailSeqNo2"))<>"" and not isnull(rsMail("MailSeqNo2")) then
					ReturnMailNo=trim(rsMail("MailSeqNo2"))
				end if
				if trim(rsMail("ReturnResonID"))<>"" and not isnull(rsMail("ReturnResonID")) then
					if trim(rsMail("ReturnResonID"))="5" or trim(rsMail("ReturnResonID"))="6" or trim(rsMail("ReturnResonID"))="7" or trim(rsMail("ReturnResonID"))="T" then
						StoreAndSendFlag=1
					end if
				end if
				'不管公示寄存都要寄第二次
				if sys_City="基隆市" then
					if trim(rsMail("UserMarkResonID"))<>"" and not isnull(rsMail("UserMarkResonID")) and trim(rsMail("UserMarkResonID"))<>"A" and trim(rsMail("UserMarkResonID"))<>"B" and trim(rsMail("UserMarkResonID"))<>"C" and trim(rsMail("UserMarkResonID"))<>"D" then
						StoreAndSendFlag=1
					end if
				end if
				if trim(rsMail("MailChkNumber"))<>"" and not isnull(rsMail("MailChkNumber")) then
					MailCheckNumber=trim(rsMail("MailChkNumber"))
				end if
				if trim(rsMail("OpenGovReportNumber"))<>"" and not isnull(rsMail("OpenGovReportNumber")) then
					MailReturnCheckNumber=trim(rsMail("OpenGovReportNumber"))
				end if
				'送達證書郵寄日期
				if sys_City="基隆市" then
					if trim(rsMail("StoreAndSendFinalMailDate"))<>"" and not isnull(rsMail("StoreAndSendFinalMailDate")) then
						StoreAndSendFinalMailDate=gArrDT(trim(rsMail("StoreAndSendFinalMailDate")))
					end if
				end if
			end if
			rsMail.close
			set rsMail=nothing

		'-----------------------------DciLog退件-----------------------------
			ReturnFileName=""	'退件上傳檔名
			ReturnBatchNumber=""	'退件批號
			ReturnStatus=""	'退件上傳狀態
			ReturnIsClose=0 '單退是否結案
			strReturn="select * from "&DBUser&".DciLog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='3'" &_
				" order by ExchangeDate desc"
			set rsReturn=conn.execute(strReturn)
			if not rsReturn.eof then
				ReturnFileName=trim(rsReturn("FileName"))
				ReturnBatchNumber=trim(rsReturn("BatchNumber"))
				if trim(rsReturn("DciReturnStatusID"))="n" then
					ReturnIsClose=1
				end if
				if trim(rsReturn("DciReturnStatusID"))<>"" and not isnull(rsReturn("DciReturnStatusID")) then
					strRStuts="select StatusContent from "&DBUser&".DciReturnStatus where DciActionID='N' and DciReturn='"&trim(rsReturn("DciReturnStatusID"))&"'"
					set rsRStuts=conn.execute(strRStuts)
					if not rsRStuts.eof then
						ReturnStatus=trim(rsRStuts("StatusContent"))
					end if
					rsRStuts.close
					set rsRStuts=nothing
				else
					ReturnStatus="未處理"
				end if
			else
				if sys_City<>"台中市" then
					ReturnStatus="未上傳"
				else
					ReturnStatus="&nbsp;"
				end if 
			end if
			rsReturn.close
			set rsReturn=nothing

		'-----------------------DciLog寄存--------------------------------
			StoreAndSendFileName=""	'寄存上傳檔名
			StoreAndSendBatchNumber=""	'寄存檔名
			StoreAndSendStatus=""	'寄存上傳狀態
			strSAndS="select * from "&DBUser&".DciLog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='4'" &_
				" order by ExchangeDate desc"
			set rsSAndS=conn.execute(strSAndS)
			if not rsSAndS.eof then
				StoreAndSendFileName=trim(rsSAndS("FileName"))
				StoreAndSendBatchNumber=trim(rsSAndS("BatchNumber"))
				if trim(rsSAndS("DciReturnStatusID"))<>"" and not isnull(rsSAndS("DciReturnStatusID")) then
					strSStuts="select StatusContent from "&DBUser&".DciReturnStatus where DciActionID='N' and DciReturn='"&trim(rsSAndS("DciReturnStatusID"))&"'"
					set rsSStuts=conn.execute(strSStuts)
					if not rsSStuts.eof then
						StoreAndSendStatus=trim(rsSStuts("StatusContent"))
					end if
					rsSStuts.close
					set rsSStuts=nothing
				else
					StoreAndSendStatus="未處理"
				end if
			else
				if sys_City<>"台中市" then
					StoreAndSendStatus="未上傳"
				else
					StoreAndSendStatus="&nbsp;"
				end if 
			end if
			rsSAndS.close
			set rsSAndS=nothing
		'-----------------------DciLog公示--------------------------------
			OpenGovFileName=""	'公示上傳檔名
			OpenGovBatchNumber=""	'公示檔名
			OpenGovStatus=""	'公示上傳狀態
			strOpenGov="select * from "&DBUser&".DciLog where BillSn="&trim(rs1("SN"))&" and ExchangeTypeID='N' and ReturnMarkType='5'" &_
				" order by ExchangeDate desc"
			set rsOpenGov=conn.execute(strOpenGov)
			if not rsOpenGov.eof then
				OpenGovFileName=trim(rsOpenGov("FileName"))
				OpenGovBatchNumber=trim(rsOpenGov("BatchNumber"))
				if trim(rsOpenGov("DciReturnStatusID"))<>"" and not isnull(rsOpenGov("DciReturnStatusID")) then
					strOStuts="select StatusContent from "&DBUser&".DciReturnStatus where DciActionID='N' and DciReturn='"&trim(rsOpenGov("DciReturnStatusID"))&"'"
					set rsOStuts=conn.execute(strOStuts)
					if not rsOStuts.eof then
						OpenGovStatus=trim(rsOStuts("StatusContent"))
					end if
					rsOStuts.close
					set rsOStuts=nothing
				else
					OpenGovStatus="未處理"
				end if
			else
				if sys_City<>"台中市" then
					OpenGovStatus="未上傳"
				else
					OpenGovStatus="&nbsp;"
				end if 
			end if
			rsOpenGov.close
			set rsOpenGov=Nothing
	if Cnt>0 then
%>
		<div class="PageNext">&nbsp;</div>
<%	end If
	Cnt=Cnt+1
	NewCase=NewCase+1
%>
	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td align="center">
				<span class="style6">舉發違反交通管理事件通知單</span>
			</td>
		</tr>
		<tr>
			<td><span class="style2">製表單位：</span><span class="style1"><%
			strUnit="select UnitName from "&DBUser&".UnitInfo where UnitID='"&trim(session("Unit_ID"))&"'"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				response.write trim(rsUnit("UnitName"))
			end if
			rsUnit.close
			set rsUnit=nothing
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">操作人：</span><span class="style1"><%
			strMem="select ChName from "&DBUser&".MemberData where MemberID='"&trim(session("User_ID"))&"'"
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("ChName"))
			end if
			rsMem.close
			set rsMem=nothing
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">製表時間：</span><span class="style3"><%=now%></span></td>
		</tr>
	</table>
	<hr>
	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td width="25%"><span class="style2">告發單號：</span><span class="style1"><%
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNO"))
			end if
			%></span></td>
			<td width="27%"><span class="style2">到案處所：</span><span class="style1"><%
			if trim(rs1("BillTypeID"))<>"2" then
			'	StationName=StationNameDci
			'else
				StationName=StationNameBillBase
			end if
			strStation="select * from "&DBUser&".Station where DciStationID='"&StationName&"'"
			set rsStation=conn.execute(strStation)
			if not rsStation.eof then
				response.write trim(rsStation("DCIStationName"))
			end if
			rsStation.close
			set rsStation=nothing
			%></span></td>
			<td width="23%"><span class="style2">告發類別：</span><span class="style1"><%
			if trim(rs1("BillTypeID"))="2" then
				response.write "逕舉"
			else
				response.write "攔停"
			end if
			%></span></td>
			<td width="25%"><span class="style2">舉發單狀態：</span><span class="style1"><%
			if trim(rs1("RecordStateID"))="-1" then
				response.write "<font color=""red"">已刪除</font>"
			else
				response.write "正常"
			end if
			'刪除原因
			if trim(rs1("RecordStateID"))="-1" or sys_City="台中市" or trim(Session("Credit_ID"))="A000000000" then
				strDelRea="select b.Content from "&DBUser&".BillDeleteReason a,DciCode b where a.BillSn="&trim(rs1("Sn"))&" and b.TypeID=3 and a.DelReason=b.ID"
				set rsDelRea=conn.execute(strDelRea)
				if not rsDelRea.eof then
					response.write "<font color=""red"">." & trim(rsDelRea("Content")) & "</font>"
				else
					response.write "&nbsp;"
				end if
				rsDelRea.close
				set rsDelRea=nothing
			end if
			if trim(rs1("RecordStateID"))="-1" and (sys_City="高雄市" or sys_City=ApconfigureCityName) then
				strDelTime="select * from "&DBUser&".log where typeid=352 and ActionContent like '%單號:"&trim(rs1("BillNo"))&"%' and ActionContent like '%車號:"&trim(rs1("CarNo"))&"%' and rownum<=1 order by ActionDate Desc"
				set rsDelTime=conn.execute(strDelTime)
				if not rsDelTime.eof then
					response.write "<font color=""red"">."&year(rsDelTime("ActionDate"))-1911&"/"&month(rsDelTime("ActionDate"))&"/"&day(rsDelTime("ActionDate"))&" "&hour(rsDelTime("ActionDate"))&":"&minute(rsDelTime("ActionDate"))&"</font>"
				end if
				rsDelTime.close
				set rsDelTime=nothing
			end if
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">入案日期：</span><span class="style1"><%
			if CaseInDate<>"" and not isnull(CaseInDate) then
				response.write left(CaseInDate,len(CaseInDate)-4)&"-"&mid(CaseInDate,len(CaseInDate)-3,2)&"-"&mid(CaseInDate,len(CaseInDate)-1,2)
			end if
			%></span></td>
			<td><span class="style2">違規時間：</span><span class="style1"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gArrDT(trim(rs1("IllegalDate")))&"&nbsp;"
				response.write Right("00"&hour(rs1("IllegalDate")),2)&":"
				response.write Right("00"&minute(rs1("IllegalDate")),2)
			end if		
			%></span></td>
			<td colspan="2"><span class="style2">舉發員警：</span><span class="style1"><%
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write trim(rs1("BillMem1"))
				strMem1="select LoginID from "&DBUser&".MemberData where memberId="&trim(rs1("BillMemID1"))
				set rsMem1=conn.execute(strMem1)
				if not rsMem1.eof then
					response.write "("&trim(rsMem1("LoginID"))&")"
				end if
				rsMem1.close
				set rsMem1=nothing
			end if	
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write "/&nbsp;"&trim(rs1("BillMem2"))
				strMem2="select LoginID from "&DBUser&".MemberData where memberId="&trim(rs1("BillMemID2"))
				set rsMem2=conn.execute(strMem2)
				if not rsMem2.eof then
					response.write "("&trim(rsMem2("LoginID"))&")"
				end if
				rsMem2.close
				set rsMem2=nothing
			end if	
			if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
				response.write "/&nbsp;"&trim(rs1("BillMem3"))
				strMem3="select LoginID from "&DBUser&".MemberData where memberId="&trim(rs1("BillMemID3"))
				set rsMem3=conn.execute(strMem3)
				if not rsMem3.eof then
					response.write "("&trim(rsMem3("LoginID"))&")"
				end if
				rsMem3.close
				set rsMem3=nothing
			end if	
			if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
				response.write "/&nbsp;"&trim(rs1("BillMem4"))
				strMem4="select LoginID from "&DBUser&".MemberData where memberId="&trim(rs1("BillMemID4"))
				set rsMem4=conn.execute(strMem4)
				if not rsMem4.eof then
					response.write "("&trim(rsMem4("LoginID"))&")"
				end if
				rsMem4.close
				set rsMem4=nothing
			end if	
			%></span></td>
		</tr>
<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if (left(trim(rs1("Rule1")),2)="40" or (int(rs1("Rule1"))>4310200 and int(rs1("Rule1"))<4310209) or (int(rs1("Rule1"))>3310100 and int(rs1("Rule1"))<3310111)) and sys_City="基隆市" then
				response.write trim(rs1("Rule1"))&" "&"該路段限速"&trim(rs1("RuleSpeed"))&"公里、經雷達測速為"&trim(rs1("IllegalSpeed"))&"公里、超速"&cint(rs1("IllegalSpeed"))-cint(rs1("RuleSpeed"))&"公里"
			else
				if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
					if left(trim(rs1("Rule1")),4)="2110" or trim(rs1("Rule1"))="4310102" or trim(rs1("Rule1"))="4310103" or trim(rs1("Rule1"))="4310104" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple=" and CarSimpleID in ('3','0')"
						else
							strCarImple=""
						end if
					end if
					strR1="select IllegalRule,Level1 from "&DBUser&".Law where ItemID='"&trim(rs1("Rule1"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple&" order by CarSimpleID Desc"
					set rsR1=conn.execute(strR1)
					if not rsR1.eof then 
						response.write trim(rs1("Rule1"))&" "&trim(rsR1("IllegalRule"))
					end if
					rsR1.close
					set rsR1=nothing

					if trim(rs1("BillTypeID"))="2" and trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
						response.write "&nbsp;"&trim(rs1("Rule4"))
					end if
				end if	
			end If
			If DciForfeit1<>"" And (sys_City="高雄市" or sys_City=ApconfigureCityName) Then
				response.write " &nbsp; 處新台幣 "&DciForfeit1&" 元"
			End if
			%></span></td>
		</tr>
<%if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then%>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if (left(trim(rs1("Rule2")),2)="40" or (int(rs1("Rule2"))>4310200 and int(rs1("Rule2"))<4310209) or (int(rs1("Rule2"))>3310100 and int(rs1("Rule2"))<3310111)) and sys_City="基隆市" then
				response.write trim(rs1("Rule2"))&" "&"該路段限速"&trim(rs1("RuleSpeed"))&"公里、經雷達測速為"&trim(rs1("IllegalSpeed"))&"公里、超速"&cint(rs1("IllegalSpeed"))-cint(rs1("RuleSpeed"))&"公里"
			else
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					if left(trim(rs1("Rule2")),4)="2110" or trim(rs1("Rule2"))="4310102" or trim(rs1("Rule2"))="4310103" or trim(rs1("Rule2"))="4310104" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple2=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple2=" and CarSimpleID in ('3','0')"
						else
							strCarImple2=""
						end if
					end if
					strR2="select IllegalRule,Level1 from "&DBUser&".Law where ItemID='"&trim(rs1("Rule2"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2&" order by CarSimpleID Desc"
					set rsR2=conn.execute(strR2)
					if not rsR2.eof then 
						response.write trim(rs1("Rule2"))&" "&trim(rsR2("IllegalRule"))
					end if
					rsR2.close
					set rsR2=nothing
				end if
			end If
			If DciForfeit2<>"" And (sys_City="高雄市" or sys_City=ApconfigureCityName) Then
				response.write " &nbsp; 處新台幣 "&DciForfeit2&" 元"
			End if
			%></span></td>
		</tr>
<%end if%>
<%if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then%>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if (left(trim(rs1("Rule3")),2)="40" or (int(rs1("Rule3"))>4310200 and int(rs1("Rule3"))<4310209) or (int(rs1("Rule3"))>3310100 and int(rs1("Rule3"))<3310111)) and sys_City="基隆市" then
				response.write trim(rs1("Rule3"))&" "&"該路段限速"&trim(rs1("RuleSpeed"))&"公里、經雷達測速為"&trim(rs1("IllegalSpeed"))&"公里、超速"&cint(rs1("IllegalSpeed"))-cint(rs1("RuleSpeed"))&"公里"
			else
				if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
					if left(trim(rs1("Rule3")),4)="2110" or trim(rs1("Rule3"))="4310102" or trim(rs1("Rule3"))="4310103" or trim(rs1("Rule3"))="4310104" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple2=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple2=" and CarSimpleID in ('3','0')"
						else
							strCarImple2=""
						end if
					end if
					strR2="select IllegalRule,Level1 from "&DBUser&".Law where ItemID='"&trim(rs1("Rule3"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2&" order by CarSimpleID Desc"
					set rsR2=conn.execute(strR2)
					if not rsR2.eof then 
						response.write trim(rs1("Rule3"))&" "&trim(rsR2("IllegalRule"))
					end if
					rsR2.close
					set rsR2=nothing
				end if	
			end If
			If DciForfeit3<>"" And (sys_City="高雄市" or sys_City=ApconfigureCityName) Then
				response.write " &nbsp; 處新台幣 "&DciForfeit3&" 元"
			End if
			%></span></td>
		</tr>
<%end if%>
<%if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) and trim(rs1("BillTypeID"))<>"2" then%>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if (left(trim(rs1("Rule4")),2)="40" or (int(rs1("Rule4"))>4310200 and int(rs1("Rule4"))<4310209) or (int(rs1("Rule4"))>3310100 and int(rs1("Rule4"))<3310111)) and sys_City="基隆市" then
				response.write trim(rs1("Rule4"))&" "&"該路段限速"&trim(rs1("RuleSpeed"))&"公里、經雷達測速為"&trim(rs1("IllegalSpeed"))&"公里、超速"&cint(rs1("IllegalSpeed"))-cint(rs1("RuleSpeed"))&"公里"
			else
				if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
					if left(trim(rs1("Rule4")),4)="2110" or trim(rs1("Rule4"))="4310102" or trim(rs1("Rule4"))="4310103" or trim(rs1("Rule4"))="4310104" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple2=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple2=" and CarSimpleID in ('3','0')"
						else
							strCarImple2=""
						end if
					end if
					strR2="select IllegalRule,Level1 from "&DBUser&".Law where ItemID='"&trim(rs1("Rule4"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2&" order by CarSimpleID Desc"
					set rsR2=conn.execute(strR2)
					if not rsR2.eof then 
						response.write trim(rs1("Rule4"))&" "&trim(rsR2("IllegalRule"))
					end if
					rsR2.close
					set rsR2=nothing
				end if	
			end If
			'If DciForfeit4<>"" And sys_City="高雄市" Then
			'	response.write " &nbsp; 處新台幣 "&DciForfeit4&" 元"
			'End if
			%></span></td>
		</tr>
<%end if%>
<tr>
			<td colspan="<%
			If sys_City="台東縣" Then 
				response.write "2"
			Else
				response.write "3"
			End If 
			%>"><span class="style2">違規路段：</span><span class="style1"><%
			response.write trim(rs1("IllegalAddressID"))&" "&trim(rs1("IllegalAddress"))
			%></span></td>
<%	If sys_City="台東縣" Then
		If trim(rs1("Rule1"))="5620001" And not isnull(rs1("imagefilename")) Then 
%>
			<td><span class="style2">停車時間：</span><span class="style1"><%
			PFileArr=Split(trim(rs1("imagefilename")),"\")
			If UBound(PFileArr)>0 Then 
				PFile=Replace(PFileArr(1),".jpg","")
			End If 
			strPTime="select DealLineDate,IllegalDate from "&DBUser&".billbase " &_
				" where CarNo='"&Trim(rs1("CarNo"))&"' and billno is null " &_
				" and ImageFileNameB is not null and imagepathname='" & PFile & "'" &_
				" and recordstateid=0 "
			Set rsPTime=conn.execute(strPTime)
			If Not rsPTime.eof Then
				if trim(rsPTime("IllegalDate"))<>"" and not isnull(rsPTime("IllegalDate")) then
					response.write gArrDT(trim(rsPTime("IllegalDate")))&"&nbsp;"
					response.write Right("00"&hour(rsPTime("IllegalDate")),2)&":"
					response.write Right("00"&minute(rsPTime("IllegalDate")),2)
				end if	
			End If
			rsPTime.close
			Set rsPTime=Nothing 
			
			%></span></td>
<%		End If 
	End If 
%>
			<td><span class="style2">是否郵寄：</span><span class="style1"><%
			if trim(rs1("EquipMentID"))<>"" and not isnull(rs1("EquipMentID")) then
				if trim(rs1("EquipMentID"))="1" then
					response.write "是"
				else
					response.write "否"
				end if
			end if	
			%></span></td>
		</tr>

		<tr>
			<td><span class="style2">郵寄日期：</span><span class="style1"><%=MailDate%></span></td>
			<td><span class="style2">郵寄序號：</span><span class="style1"><%
			if sys_City<>"台南縣" and sys_City<>"台南市" then
				response.write MailNumber
			else
				response.write BillMailNo
			end if
			%></span></td>
			<td><span class="style2">簽收狀況：</span><span class="style1">
			<%
				'可參考google doc "攔停 簽收 狀況 "
				if trim(rs1("SignType"))<>"" and not isnull(rs1("SignType")) then
					if rs1("SignType")="A" then response.write "簽收"
					if rs1("SignType")="U" then 
						strR2="select SignStateID from "&DBUser&".BillUserSignDate where billsn=" & trim(rs1("sn"))
						set rsR2=conn.execute(strR2)
						if not rsR2.eof then 
							if rsR2("SignStateID")="2" then response.write "拒簽已收"
							if rsR2("SignStateID")="3" then response.write "已簽拒收"							
						else 
							response.write "拒簽收"
						end if
						rsR2.close
						set rsR2=nothing																
					end if				
				else
						strR2="select SignStateID from "&DBUser&".BillUserSignDate where billsn=" & trim(rs1("sn"))
						set rsR2=conn.execute(strR2)
						if not rsR2.eof then 
							if rsR2("SignStateID")="5" then response.write "補開單"
						end if
						rsR2.close
						set rsR2=nothing															
				end if
			%>			
			</span></td>

		<% if sys_City="台東縣" then %>
			<td>
				車主證號(查車) : 
				<%
					strReturn="select OWNERID from "&DBUser&".BillBaseDCIReturn where BillNo is null and CarNo='"&trim(rs1("CarNo"))&"' order by DCICASEINDATE desc"	
					set rsReturn=conn.execute(strReturn)
					If Not rsReturn.eof Then 
						response.write "<span class='style1'>" & rsReturn("OWNERID") & "</span>"
					else
						response.write ""	
					end if
					rsReturn.close
					set rsReturn=nothing
	
				%>
			</td>
	  <% end if %>			
			
			
			
		</tr>
		<tr>
			<td><span class="style2">違規人證號：</span><span class="style1"><%=IllegalMemID%></span></td>
			<td><span class="style2">違規人姓名：</span><span class="style1"><%=funcCheckFont(IllegalMem,20,1)%></span></td>
			<td colspan="3"><span class="style2">違規人住址：</span><span class="style1"><%=funcCheckFont(IllegalAddress,20,1)%></span></td>
		</tr>
		<tr>
			<td><span class="style2">車號：</span><span class="style1"><font size="4"><%=trim(rs1("CarNo"))%></font></span></td>
			<td><span class="style2">車主姓名：</span><span class="style1"><%=funcCheckFont(OwnerName,20,1)%></span></td>
			<td colspan="3"><span class="style2">車主住址：</span><span class="style1"><%=funcCheckFont(OwnerAddress,20,1)%></span></td>
		</tr>
		<tr>
			<td><span class="style2">填單日期：</span><span class="style1"><%
			if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
				response.write gArrDT(trim(rs1("BillFillDate")))
			end if	
			%></span></td>
			<td><span class="style2">詳細車種：</span><span class="style1"><%=DciCarType%></span></td>
			<td colspan="3"><span class="style2">舉發單位：</span><span class="style1"><%
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				response.write trim(rs1("BillUnitID"))&"&nbsp;"
				strBillUnit="select UnitName from "&DBUser&".UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsBillUnit=conn.execute(strBillUnit)
				if not rsBillUnit.eof then
					response.write trim(rsBillUnit("UnitName"))
				end if
				rsBillUnit.close
				set rsBillUnit=nothing
			end if	
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">到案日期：</span><span class="style1"><%
			if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
				response.write gArrDT(trim(rs1("DealLineDate")))
			end if	
			%></span></td>
			<td><span class="style2">簡式車種：</span><span class="style1"><%
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				if trim(rs1("CarSimpleID"))="1" then
					response.write "汽車"
				elseif trim(rs1("CarSimpleID"))="2" then
					response.write "拖車"
				elseif trim(rs1("CarSimpleID"))="3" then
					response.write "重機"
				elseif trim(rs1("CarSimpleID"))="4" then
					response.write "輕機"
				elseif trim(rs1("CarSimpleID"))="6" then
					response.write "臨時車牌"
				elseif trim(rs1("CarSimpleID"))="7" then
					response.write "試車牌"
				end if
			end if	
			%></span></td>
			<td><span class="style2">建檔日期：</span><span class="style1"><%
			if trim(rs1("RecordDate"))<>"" and not isnull(rs1("RecordDate")) then
				response.write gArrDT(trim(rs1("RecordDate")))
			end if	
			%></span></td>
			<td><span class="style2">操作人員：</span><span class="style1"><%
			strRecMem="select ChName from "&DBUser&".MemberData where MemberID='"&trim(rs1("RecordMemberID"))&"'"
			set rsRecMem=conn.execute(strRecMem)
			if not rsRecMem.eof then
				response.write trim(rsRecMem("ChName"))
			end if
			rsRecMem.close
			set rsRecMem=nothing
			%></span></td>
		</tr>
			<%if sys_City<>"台中市" then%>	
				<tr>
				<%if sys_City="台南市" then%>	
					<td ><span class="style2">輔助車種：</span><span class="style1"><%
					If Trim(rs1("CarAddID"))="1" Then
						response.write "1大貨"
					ElseIf Trim(rs1("CarAddID"))="2" Then
						response.write "2大客"
					ElseIf Trim(rs1("CarAddID"))="3" Then
						response.write "3砂石"
					ElseIf Trim(rs1("CarAddID"))="4" Then
						response.write "4土方"
					ElseIf Trim(rs1("CarAddID"))="5" Then
						response.write "5動力"
					ElseIf Trim(rs1("CarAddID"))="6" Then
						response.write "6貨櫃"
					ElseIf Trim(rs1("CarAddID"))="7" Then
						response.write "7大型重機"
					ElseIf Trim(rs1("CarAddID"))="8" Then
						response.write "8拖吊"
					ElseIf Trim(rs1("CarAddID"))="9" Then
						response.write "9(550cc)重機"
					ElseIf Trim(rs1("CarAddID"))="10" Then
						response.write "10計程車"
					ElseIf Trim(rs1("CarAddID"))="11" Then
						response.write "11危險物品"
					End If 
					%></span></td>
				<%End If %>
					<td colspan="3"><span class="style2">行駕照狀態：</span><span class="style1"><%=DciIDstatus%></span></td>
				</tr>
			<%End if%>
<%
		strDSupd="select * from "&DBUser&".DCISTATUSUPDATE where Billsn="&Trim(rs1("Sn"))
		Set rsDSupd=conn.execute(strDSupd)
		If Not rsDSupd.eof Then
		%>
				<tr>
				<td colspan="3">
					<span class="style2">強制入案前狀態：</span><span class="style1"><%
				strDS1="select * from "&DBUser&".Dcireturnstatus where DciActionID='W' " &_
					" and DciReturn='"&Trim(rsDSupd("StatUS"))&"'"
				Set rsDS1=conn.execute(strDS1)
				If Not rsDS1.eof Then
					response.write rsDS1("StatusContent")
				End If
				rsDS1.close
				Set rsDS1=Nothing
				strDS2="select * from "&DBUser&".Dcireturnstatus where DciActionID='WE' " &_
					" and DciReturn='"&Trim(rsDSupd("DciErrorCarData"))&"'"
				Set rsDS2=conn.execute(strDS2)
				If Not rsDS2.eof Then
					response.write " "&rsDS2("StatusContent")
				End If
				rsDS2.close
				Set rsDS2=Nothing
				response.write " "&rsDSupd("RecordDate")
					%></span>
				</td>
				</tr>
		<%
		End If
		rsDSupd.close
		Set rsDSupd=nothing
		%>			

	</table>
	<hr>
	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td width="35%"><span class="style2">入案檔名：</span><span class="style1"><%=DciFileName%></span></td>
			<td width="25%">
			<span class="style2">入案批號：</span><span class="style1"><%=DciBatchNumber%></span>
			</td>
			<td width="40%"><span class="style2">入案狀態：</span><span class="style1"><%
			if CaseStatus<>"" and not isnull(CaseStatus) then
				response.write CaseStatus
			end if
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2"><%
			if sys_City<>"高雄縣" then
				response.write "簽收/寄存 上傳檔名："
			else
				response.write "簽收 上傳檔名："
			end if
			%></span><span class="style1"><%
			response.write GetFileName
			%></span></td>
			<td>
			<span class="style2"><%
			if sys_City<>"高雄縣" then
				response.write "簽收/寄存批號："
			else
				response.write "簽收批號："
			end if
			%></span><span class="style1"><%
			response.write GetBatchNumber
			if sys_City="台南市" Then
				If GetBatchNumber<>"" Then
					response.write "("
					strBseq="select sn-(select sn from (select * from dcilog where batchnumber='" & GetBatchNumber & "' order by sn) where rownum=1)+1 as BatchSeq from dcilog where batchnumber='" & GetBatchNumber & "' and billSn=" & Trim(rs1("Sn"))
					Set rsBseq=conn.execute(strBseq)
					If not rsBseq.eof Then
						response.write rsBseq("BatchSeq")
					End If 
					rsBseq.close
					Set rsBseq=Nothing 
					response.write ")"
				End If 
			End If 
			%></span>
			</td>
			<td><span class="style2"><%
			if sys_City<>"高雄縣" then
				response.write "簽收/寄存 上傳狀態："
			else
				response.write "簽收 上傳狀態："
			end if
			%></span><span class="style1"><%
			response.write GetStatus
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2"><%
			if sys_City<>"高雄縣" then
				response.write "簽收/寄存 日期："
			else
				response.write "簽收 日期："
			end if
			%></span><span class="style1"><%
			response.write GetMailDate
			%></span></td>
			<td ><span class="style2">簽收人：</span><span class="style1"><%
			response.write SignMan
			%></span></td>
			<td><span class="style2"><%
			if sys_City<>"高雄縣" then
				response.write "簽收/寄存 原因："
			else
				response.write "簽收 原因："
			end if
			%></span><span class="style1"><%
			response.write GetMailReason
			%></span></td>
		</tr>
		<tr>
			<td colspan="2"><span class="style2"><font color="clred">撤銷送達 日期：</font></span><span class="style1"><%
			response.write CancalSendDate
			%></span>
			</td>
			<td><span class="style2"><%
			if sys_City="花蓮縣" then
			%>查證文號：<%
			else
			%>寄存郵局：<%
			end if
			%></span><span class="style1"><%
			response.write MailStation
			%></span></td>
		</tr>
		<tr>
			<td ><span class="style2">退件上傳檔名：</span><span class="style1"><%=ReturnFileName%></span></td>
			<td>
			<span class="style2">退件批號：</span><span class="style1"><%=ReturnBatchNumber%></span>
			</td>
			<td ><span class="style2">退件上傳狀態：</span><span class="style1"><%=ReturnStatus%></span></td>
		</tr>
		<tr>
			<%if sys_City<>"台東縣" then %>
				<td colspan="2"><span class="style2">退件郵寄日期：</span><span class="style1"><%
				if sys_City="南投縣" then
					if ReturnIsClose=0 then
						response.write ReturnSendMailDate
					end if
				else
					response.write ReturnSendMailDate
				end if
				%></span></td>
			<%else%>
				<td colspan="2"><span class="style2">退件郵寄日期：</span><span class="style1"><%=MailReturnDCIDate%></span></td>
			<%end if%>
			<td><span class="style2">退件郵寄序號：</span><span class="style1"><%
			if sys_City<>"台南縣" and sys_City<>"台南市" then
				response.write ReturnMailNumber
			else
				response.write ReturnMailNo
			end if
			%></span></td>
		</tr>
		<tr>
			<td colspan="2"><span class="style2">退回日期：</span><span class="style1"><%=ReturnMailDate%></span></td>
			<td><span class="style2">退件原因：</span><span class="style1"><%
			response.write ReturnReason
			if ReturnReason<>"" and sys_City="南投縣" then
				if instr(trim(rs1("Note")),"退回原因：")>0 then
					response.write "("&mid(trim(rs1("Note")),instr(trim(rs1("Note")),"退回原因：")+5,4)&")"
				end if
			end if
			%></span></td>
		</tr>
		<tr>
			<td colspan="3"><span class="style2">二次郵寄地址：</span><span class="style1"><%
	if sys_City="基隆市" then
		ShowSecondAddress=0
		strRCnt="select count(*) as cnt from Dcilog where BillSn="&trim(rs1("Sn"))&" and ExchangeTypeID='N' and ReturnMarkType='3'"
		set rsRCnt=conn.execute(strRCnt)
		if not rsRCnt.eof then
			if cint(rsRCnt("cnt"))>1 then
				ShowSecondAddress=1
			end if
		end if
		rsRCnt.close

	end if
	if ((StoreAndSendFlag=1 or sys_City="彰化縣") and sys_City<>"基隆市") or (ShowSecondAddress=1 and sys_City="基隆市") then
		if trim(rs1("BillTypeID"))="1" then	
			'攔停抓戶籍地址
			strSqlD="select Driver,DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status='Y'"
			set rsD=conn.execute(strSqlD)
			if not rsD.eof then
			if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣" or sys_City="嘉義縣"  then
				ZipName=""
			else
				strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof Then
					If CDbl(Year(rs1("IllegalDate")))<2011 then
						ZipName=ChangeOldCity(trim(rsD("DriverHomeZip")),trim(rsZip("ZipName")))
					Else
						ZipName=trim(rsZip("ZipName"))
					End If						
				end if
				rsZip.close
				set rsZip=nothing
			end if
				GetMailMem=trim(rsD("Driver"))
				GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&trim(rsD("DriverHomeAddress"))
			end if
			rsD.close
			set rsD=nothing
			response.write funcCheckFont(GetMailMem,20,1)&"--"&funcCheckFont(GetMailAddress,20,1)
		else
			if sys_City="台東縣" then
				
			else
				strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
				set rsD=conn.execute(strSqlD)
				if not rsD.eof then
					if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("DriverHomeZip"))&trim(rsD("DriverHomeAddress"))
					else
						GetMailMem=trim(rsD("Owner"))
						GetMailAddress=trim(rsD("OwnerZip"))&trim(rsD("OwnerAddress"))
					end if
				else
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status='Y'"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof then
						if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof Then
								If CDbl(Year(rs1("IllegalDate")))<2011 then
									ZipName=ChangeOldCity(trim(rsD2("DriverHomeZip")),trim(rsZip("ZipName")))
								Else
									ZipName=trim(rsZip("ZipName"))
								End If									
							end if
							rsZip.close
							set rsZip=nothing

							GetMailMem=trim(rsD2("Owner"))
							GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&trim(rsD2("DriverHomeAddress"))
						else
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof Then
								If CDbl(Year(rs1("IllegalDate")))<2011 then
									ZipName=ChangeOldCity(trim(rsD2("OwnerZip")),trim(rsZip("ZipName")))
								Else
									ZipName=trim(rsZip("ZipName"))
								End If									
							end if
							rsZip.close
							set rsZip=nothing

							GetMailMem=trim(rsD2("Owner"))
							GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
						end if
					end if
					rsD2.close
					set rsD2=nothing
				end if
				rsD.close
				set rsD=nothing
			end If
			response.write funcCheckFont(GetMailMem,20,1)&"--"&funcCheckFont(GetMailAddress,20,1)
		end if
	end if
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">寄存送達上傳檔名：</span><span class="style1"><%=StoreAndSendFileName%></span></td>
			<td>
			<span class="style2">寄存送達批號：</span>
			<span class="style1"><%=StoreAndSendBatchNumber%></span>
			</td>
			<td><span class="style2">寄存送達上傳狀態：</span><span class="style1"><%=StoreAndSendStatus%></span></td>
		</tr>
		<tr>
			<td colspan="2"><span class="style2">寄存送達書號：</span><span class="style1"><%=StoreAndSendGovNumber%></span></td>
			<td><span class="style2">寄存送達日：</span><span class="style1"><%=StoreAndSendEffectDate%></span></td>
		</tr>
		<tr>
			<td colspan="2"><span class="style2"><% if sys_City="台東縣" then %>
					寄存送達生效
				<%elseif sys_City="高雄縣" then %>
					寄存送達期滿
				<%else%>
					寄存送達生效(完成)
				<% end if%>	日：</span><span class="style1"><%=StoreAndSendEndDate%></span></td>
			<td><span class="style2">寄存送達 退件原因：</span><span class="style1"><%=StoreAndSendReason%></span></td>
		</tr>
		<tr>
			<%	
				'smith 加入寄存期滿退回日顯示 
				if sys_City<>"花蓮縣" then
			%>	
				<% if sys_City<>"台東縣" then %>
					<td colspan="3"><span class="style2">寄存送達 退回日期：</span><span class="style1"><%=StoreAndSendDate%></span></td>			
				<% end if %>					
			<% else %>
					<td colspan="2"><span class="style2">寄存送達 退回日期：</span><span class="style1"><%=StoreAndSendDate%></span></td>
					<td><span class="style2">寄存送達期滿 退回日期：</span><span class="style1"><%=MailStationReturnDate%></span></td>
			<% end if%>
		</tr>
		<tr>
			<td><span class="style2">公示送達上傳檔名：</span><span class="style1"><%=OpenGovFileName%></span></td>
			<td>
			<span class="style2">公示送達批號：</span><span class="style1"><%=OpenGovBatchNumber%></span>
			</td>
			<td><span class="style2">公示送達上傳狀態：</span><span class="style1"><%=OpenGovStatus%></span></td>
		</tr>
		<tr>
			<td colspan="2"><span class="style2">公示送達書號：</span><span class="style1"><%=OpenGovGovNumber%></span></td>
			<td><span class="style2"><%
			if sys_City="高雄縣" then
				response.write "公示送達公告日："
			else
				response.write "公示送達生效日："
			end if
			%></span><span class="style1"><%=OpenGovEffectDate%></span></td>
		</tr>

		<tr>
			<td colspan="2"><span class="style2">發文監理站日期：</span><span class="style1"><%=ReturnSendDate%></span></td>
			<td><span class="style2">公示送達原因：</span>
				<span class="style1">
					<% 							
						if  trim(OpenGovBatchNumber) <> "" then 
							response.write OPENGOVRESONID
						 end if
					%>
				</span>
			</td>
		</tr>
		<tr>
			<td colspan="<%
			if sys_City="高雄縣" then
				response.write "2"
			else
				response.write "3"
			end if
			%>"><span class="style2">備註：</span><span class="style1"><%=trim(rs1("Note"))%></span></td>
			<%if sys_City="高雄縣" then%>
			<td><span class="style2">公示送達生效日：</span>
				<span class="style1">
					<% 							
						if  trim(KSOpenGovEffDate) <> "" then 
							response.write KSOpenGovEffDate
						 end if
					%>
				</span>
			</td>
			<%end if%>
		</tr>
	<%if sys_City="基隆市" or sys_City="台中市" then%>
		<tr>
			<td colspan="3"><span class="style2">第一次投遞郵局查詢號：</span><span class="style1"><%=MailCheckNumber%></span></td>
		</tr>
		<tr>
			<td colspan="3"><span class="style2">單退後投遞郵局查詢號：</span><span class="style1"><%=MailReturnCheckNumber%></span></td>
		</tr>
	<%end if%>
	<%if sys_City="基隆市" then%>
		<tr>
			<td colspan="3"><span class="style2">送達證書郵寄日期：</span><span class="style1"><%=StoreAndSendFinalMailDate%></span></td>
		</tr>
	<%end if%>
		
	</table>
	<%
	If sys_City="台南市" or sys_City="花蓮縣" Then
		'送達証書
		strScan="select * from BillAttatchImage where BillNo='"&trim(rs1("BillNo"))&"' and TypeID=0 and Recordstateid=0 order by RecordDate"
		set rsScan=conn.execute(strScan)
		while Not rsScan.eof
		%>
			<div class='PageNext'>&nbsp;</div> <strong>送達証書&nbsp;<%
			'掃描日期
			response.write year(rsScan("RecordDate"))&"/"&month(rsScan("RecordDate"))&"/"&day(rsScan("RecordDate"))&" "&hour(rsScan("RecordDate"))&":"&minute(rsScan("RecordDate"))
			'掃瞄人
			strSMem="select Chname from Memberdata where memberid="&trim(rsScan("RecordMemberiD"))
			set rsSMem=conn.execute(strSMem)
			if not rsSMem.eof then
				response.write "&nbsp;"&rsSMem("Chname")
			end if
			rsSMem.close
			set rsSMem=nothing
			%></strong><br>
		<%If sys_City="台南市" then%>
			<img src='<%=replace(trim(rsScan("FileName")),"/img/","/scanimg/")%>' name='imgB1' width='800' onclick="OpenPic('<%=replace(trim(rsScan("FileName")),"/img/","/scanimg/")%>')">
		<%else%>
			<img src='<%=replace(trim(rsScan("FileName")),"/img/","/img/")%>' name='imgB1' width='800' onclick="OpenPic('<%=replace(trim(rsScan("FileName")),"/img/","/img/")%>')">
		<%End If %>
		<%
			rsScan.movenext
		wend
		rsScan.close
		set rsScan=nothing
		
		'公告公文
		strScan2="select * from BillAttatchImage where BillNo='"&trim(OpenGovBatchNumber)&"' and TypeID=4 and Recordstateid=0"
		set rsScan2=conn.execute(strScan2)
		while Not rsScan2.eof
		%>
		<%If sys_City="台南市" then%>
			<div class='PageNext'>&nbsp;</div><img src='<%=replace(trim(rsScan2("FileName")),"/img/","/scanimg/")%>' name='imgB1' width='750' onclick="OpenPic('<%=replace(trim(rsScan2("FileName")),"/img/","/scanimg/")%>')">
		<%else%>
			<div class='PageNext'>&nbsp;</div><img src='<%=replace(trim(rsScan2("FileName")),"/img/","/img/")%>' name='imgB1' width='750' onclick="OpenPic('<%=replace(trim(rsScan2("FileName")),"/img/","/img/")%>')">
		<%End If %>
		<%
		rsScan2.movenext
		wend
		rsScan2.close
		set rsScan2=nothing
	End If %>
<Div id="Layer112" style="width:1041px; height:24px;">
  <div align="center">
	<%if (sys_City<>"高雄市" and sys_City<>ApconfigureCityName) or trim(Session("Credit_ID"))="A000000000" then%>
  <input type="button" value="完整詳細資料" onclick='window.open("<%
  If DBUser="traffic" Then
	response.write "ViewBillBaseData_Car.asp"
  Else
	response.write "../OldBillData/ViewBillBaseData_Car_NT_OLD.asp"
  End If 
  %>?BillSN=<%=trim(rs1("SN"))%>&BillType=0","WebPage_Detail2","left=0,top=0,location=0,width=980,height=575,resizable=yes,scrollbars=yes,menubar=yes")'>
  <%end if%>
  </div>
</Div>
<%
			End if
		End If
		rs1.close
		Set rs1=Nothing 
		Else
		'惠嚨系統==============================================================================================
			str2="select * from "&DBUser&".FMASTER where FSEQ='"&trim(rsOld1("BillNo"))&"'"
			'response.write str2
			Set rs2=conn.execute(str2)
			If Not rs2.eof Then
			
	if Cnt>0 then
%>
		<div class="PageNext">&nbsp;</div>
<%	end If
	Cnt=Cnt+1
%>
	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td align="center">
				<span class="style6">舉發違反交通管理事件通知單</span>
			</td>
		</tr>
		<tr>
			<td><span class="style2">製表單位：</span><span class="style1"><%
			strUnit="select UnitName from traffic.UnitInfo where UnitID='"&trim(session("Unit_ID"))&"'"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				response.write trim(rsUnit("UnitName"))
			end if
			rsUnit.close
			set rsUnit=nothing
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">操作人：</span><span class="style1"><%
			strMem="select ChName from traffic.MemberData where MemberID='"&trim(session("User_ID"))&"'"
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("ChName"))
			end if
			rsMem.close
			set rsMem=nothing
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">製表時間：</span><span class="style3"><%=now%></span></td>
		</tr>
	</table>
	<hr>
	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td width="25%"><span class="style2">告發單號：</span><span class="style1"><%
			if trim(rs2("fseq"))<>"" and not isnull(rs2("fseq")) then
				response.write trim(rs2("fseq"))
			end if
			%></span></td>
			<td width="27%"><span class="style2">入案狀態：</span><span class="style1"><%
			if trim(rs2("FStatus"))<>"" and not isnull(rs2("FStatus")) then
				response.write trim(rs2("FStatus"))&"&nbsp;"&getDciCode(trim(rs2("FStatus")))
			else
				response.write "&nbsp;"
			end if
			%></span></td>
			<td width="23%"><span class="style2">填單日期：</span><span class="style1"><%
				if trim(rs2("RBDate"))<>"" and not isnull(rs2("RBDate")) then
					response.write GetDate(rs2("RBDate"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td width="25%"><span class="style2">員警：</span><span class="style1"><%
				if trim(rs2("PCode1"))<>"" and not isnull(rs2("PCode1")) then
					response.write trim(rs2("PCode1"))

					sql="select PNAME from "&DBUser&".POLICE where PCODE ='"&Trim(rs2("PCode1"))&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write " "&rsPCode("PName")
					set rsPcode=Nothing
				End If 

				 if trim(rs2("PCode2"))<>"" and not isnull(rs2("PCode2")) then
					response.write ","&trim(rs2("PCode2"))
					sql="select PNAME from "&DBUser&".Police where PCODE ='"&Trim(rs2("PCode2"))&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write " "&rsPCode("PName")
					set rsPcode=nothing
                end if

				if trim(rs2("PCode3"))<>"" and not isnull(rs2("PCode3")) then
					response.write ","&trim(rs2("PCode3"))
					sql="select PNAME from "&DBUser&".Police where PCODE ='"&Trim(rs2("PCode3"))&"'"
	                set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write " "&rsPCode("PName")
					set rsPcode=nothing
				end if

				if trim(rs2("PCode4"))<>"" and not isnull(rs2("PCode4")) then
					response.write ","&trim(rs1("PCode4"))
					sql="select PNAME from "&DBUser&".Police where PCODE ='"&Trim(rs2("PCode4"))&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write " "&rsPCode("PName")
					set rsPcode=nothing
				end if
			%>
			</span></td>
		</tr>
		<tr>
			<td ><span class="style2">違規車號：</span><span class="style1"><%
				if trim(rs2("CarNo"))<>"" and not isnull(rs2("CarNo")) then
					response.write trim(rs2("CarNo"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">違規日期：</span><span class="style1"><%
				if trim(rs2("IDate"))<>"" and not isnull(rs2("IDate")) then
					response.write GetDate(rs2("IDate"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td colspan="2"><span class="style2">違規時間：</span><span class="style1"><%
				if trim(rs2("ITime"))<>"" and not isnull(rs2("ITime")) then
					response.write GetTime(rs2("ITime")) &" = "&left(rs2("ITime"),2)&":"&right(rs2("ITime"),2)
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			
		</tr>
		<tr>
			<td colspan="2" ><span class="style2">違規地點：</span><span class="style1"><%
				if trim(rs2("IRCODE"))<>"" and not isnull(rs2("IRCODE")) then
		    		response.write trim(rs2("IRCODE"))
		    	else
		    		response.write "&nbsp;"
		    	end if
		    	if trim(rs2("IRNAME"))<>"" and not isnull(rs2("IRNAME")) then
		    		response.write " "&trim(rs2("IRNAME"))
		    	else
		    		response.write "&nbsp;"
		    	end if
			%></span></td>
			<td colspan="2"><span class="style2">簡式車種代碼：</span><span class="style1"><%
				if trim(rs2("CDKind"))<>"" and not isnull(rs2("CDKind")) then
						if rs2("CDKIND")="1" then
							CDKIND="汽車"
						elseif rs2("CDKIND")="2" then
							CDKIND="拖車"
						elseif rs2("CDKIND")="3" then
							CDKIND="重機"
						elseif rs2("CDKIND")="4" then
							CDKIND="輕機"
						end if
		    		response.write trim(rs2("CDKind")) & " " & CDKIND
		    	else
		    		response.write "&nbsp;"
		    	end if
			%></span></td>
		</tr>
		<tr>
			<td ><span class="style2">保險證狀態：</span><span class="style1"><%
				if trim(rs2("INSUCERT"))<>"" and not isnull(rs2("INSUCERT")) then
					if rs2("INSUCERT")="0" then
						INSUCERT="正常"
					elseif rs2("INSUCERT")="1" then
						INSUCERT="未帶"
					elseif rs2("INSUCERT")="2" then
						INSUCERT="肇事且未帶"
					elseif rs2("INSUCERT")="3" then
						INSUCERT="過期或未保"
					elseif rs2("INSUCERT")="4" then
						INSUCERT="肇事且過期或未保"
					end if
		    		response.write trim(rs2("INSUCERT")) & " " & INSUCERT
		    	else
		    		response.write "&nbsp;"
		    	end if
			%></span></td>
			<td ><span class="style2">告發類別：</span><span class="style1"><%
				if trim(rs2("AccUSeCode"))<>"" and not isnull(rs2("AccUSeCode")) then
						if rs2("AccUSeCode")="1" then 
						  AccUSeCode="攔停"
						elseif rs2("AccUSeCode")="2" then 
						  AccUSeCode="逕舉"
						elseif rs2("AccUSeCode")="8" then 
						  AccUSeCode="行人攤販"
						elseif rs2("AccUSeCode")="3" then 
						  AccUSeCode="肇事"
						elseif rs2("AccUSeCode")="4" then 
						  AccUSeCode="拖吊"
						elseif rs2("AccUSeCode")="5" then 
						  AccUSeCode="戴運砂石土方"
						elseif rs2("AccUSeCode")="A" then 
						  AccUSeCode="違規營業"
						elseif rs2("AccUSeCode")="B" then 
						  AccUSeCode="違規重標"
						elseif rs2("AccUSeCode")="N" then 
						  AccUSeCode="未知"
						end if 
		    		response.write trim(rs2("AccUSeCode")) & " " &AccUSeCode
		    	else
		    		response.write "&nbsp;"
		    	end if
			%></span></td>
			<td ><span class="style2">應到案日期：</span><span class="style1"><%
				if trim(rs2("ARVDATE"))<>"" and not isnull(rs2("ARVDATE")) then
					response.write GetDate(rs2("ARVDATE"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">到案處所：</span><span class="style1"><%
				response.write trim(rs2("SPRVSNNO"))&" "
				if trim(rs2("SPRVSNNO"))<>"" and not isnull(rs2("SPRVSNNO")) then
					strStation="select SPNAME from traffic4.SPRVSN where SPRVSNNO='"&trim(rs2("SPRVSNNO"))&"'"
					set rsStation=conn.execute(strStation)
					if not rsStation.eof then
						response.write trim(rsStation("SPNAME"))
					end if
					rsStation.close
					set rsStation=Nothing
				End if
			%>
			</span></td>
		</tr>
		<tr>
			<td colspan="2"><span class="style2">代保管物件：</span><span class="style1"><%
				if trim(rs2("HOLDCODE1"))<>"" and not isnull(rs2("HOLDCODE1")) and trim(rs2("HOLDCODE1"))<>"0" then
					response.write trim(rs2("HOLDCODE1"))

					rstemp="select HOLDName from "&DBUser&".HOLD where HOLDCode='"&trim(rs2("HOLDCODE1"))&"'"
					set rstemp=conn.execute(rstemp)
					If Not rstemp.eof Then  
						response.write   "&nbsp;"&trim(rstemp("HOLDName")&"")
					Else
						response.write   "&nbsp;"
					End if
					set rstemp=nothing

					if trim(rs2("HOLDCODE2"))<>"" and not isnull(rs2("HOLDCODE2")) and trim(rs2("HOLDCODE2"))<>"0" then
						response.write "<br>"&trim(rs2("HOLDCODE2"))
						rstemp="select HOLDName from "&DBUser&".HOLD where HOLDCode='"&trim(rs2("HOLDCODE2"))&"'"
						set rstemp=conn.execute(rstemp)
						If Not rstemp.eof Then  
							response.write   "&nbsp;"&trim(rstemp("HOLDName")&"")
						Else
							response.write   "&nbsp;"
						End if
						set rstemp=nothing
					end if

					if trim(rs2("HOLDCODE3"))<>"" and not isnull(rs2("HOLDCODE3")) and trim(rs2("HOLDCODE3"))<>"0" then
						response.write "<br>"&trim(rs2("HOLDCODE3"))
						rstemp="select HOLDName from "&DBUser&".HOLD where HOLDCode='"&trim(rs2("HOLDCODE3"))&"'"
						set rstemp=conn.execute(rstemp)
						If Not rstemp.eof Then  
							response.write   "&nbsp;"&trim(rstemp("HOLDName")&"")
						Else
							response.write   "&nbsp;"
						End if
						set rstemp=nothing
					end if
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">違規人姓名：</span><span class="style1"><%
				if trim(rs2("IName"))<>"" and not isnull(rs2("IName")) then
					response.write trim(rs2("IName"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">身份證字號：</span><span class="style1"><%
				if trim(rs2("IIDNO"))<>"" and not isnull(rs2("IIDNO")) then
					response.write trim(rs2("IIDNO"))
				else
					response.write "&nbsp;"
				end if
			%>
			</span></td>
		</tr>
		<tr>
			<td ><span class="style2">出生日期：</span><span class="style1"><%
				if trim(rs2("IBIRTH"))<>"" and not isnull(rs2("IBIRTH")) then
					response.write GetDate(rs2("IBIRTH"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td colspan="3"><span class="style2">違規人地址：</span><span class="style1"><%
				if trim(rs2("IADDR"))<>"" and not isnull(rs2("IADDR")) then
					response.write trim(rs2("IZIP"))& " " &trim(rs2("IADDR"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
	
		</tr>
		<tr>
			<td colspan="4"><span class="style2">違規法條一：</span><span class="style1"><%
				if trim(rs2("RULEF1"))<>"" and not isnull(rs2("RULEF1")) then
					response.write trim(rs2("RULEF1")) 

					sql="select RULENAME from traffic4.RULEF where RULECODE ='"&rs2("RULEF1")&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("RULENAME")&"&nbsp;"&trim(rs2("FACTG1"))
					set rsPcode=nothing
				else
					response.write "&nbsp;"
				end if
			%></span></td>
		</tr>
<%if trim(rs2("RULEF2"))<>"" and not isnull(rs2("RULEF2")) then%>
		<tr>
			<td colspan="4"><span class="style2">違規法條二：</span><span class="style1"><%
				if trim(rs2("RULEF2"))<>"" and not isnull(rs2("RULEF2")) then
					response.write trim(rs2("RULEF2")) 

					sql="select RULENAME from traffic4.RULEF where RULECODE ='"&rs2("RULEF2")&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("RULENAME")&"&nbsp;"&trim(rs2("FACTG2"))
					set rsPcode=nothing
				else
					response.write "&nbsp;"
				end if
			%></span></td>
		</tr>
<%End If %>
<%if trim(rs2("RULEF3"))<>"" and not isnull(rs2("RULEF3")) then%>
		<tr>
			<td colspan="4"><span class="style2">違規法條三：</span><span class="style1"><%
				if trim(rs2("RULEF3"))<>"" and not isnull(rs2("RULEF3")) then
					response.write trim(rs2("RULEF3")) 

					sql="select RULENAME from traffic4.RULEF where RULECODE ='"&rs2("RULEF3")&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("RULENAME")&"&nbsp;"&trim(rs2("FACTG3"))
					set rsPcode=nothing
				else
					response.write "&nbsp;"
				end if
			%></span></td>
		</tr>
<%End If %>
<%if trim(rs2("RULEF4"))<>"" and not isnull(rs2("RULEF4")) then%>
		<tr>
			<td colspan="4"><span class="style2">違規法條四：</span><span class="style1"><%
				if trim(rs2("RULEF4"))<>"" and not isnull(rs2("RULEF4")) then
					response.write trim(rs2("RULEF4")) 

					sql="select RULENAME from traffic4.RULEF where RULECODE ='"&rs2("RULEF4")&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("RULENAME")&"&nbsp;"&trim(rs2("FACTG4"))
					set rsPcode=nothing
				else
					response.write "&nbsp;"
				end if
			%></span></td>
		</tr>
<%End If %>
		<tr>	
			<td colspan="4" ><hr></td>
		<tr>
		<tr>
			<td ><span class="style2">車主姓名：</span><span class="style1"><%
				if trim(rs2("OWNAME"))<>"" and not isnull(rs2("OWNAME")) then
					response.write trim(rs2("OWNAME"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td colspan="3"><span class="style2">車主地址：</span><span class="style1"><%
				if trim(rs2("OWADDR"))<>"" and not isnull(rs2("OWADDR")) then
					response.write trim(rs2("OWZIP"))&"  " & trim(rs2("OWADDR"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
		</tr>
		<tr>
			<td ><span class="style2">詳細車種代碼：</span><span class="style1"><%
				if trim(rs2("CDType"))<>"" and not isnull(rs2("CDType")) then
					response.write trim(rs2("CDType")) 

					sql="select CDName from traffic4.CARKIND where CDType ='"&rs2("CDType")&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("CDName")
					set rsPcode=nothing

				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td colspan="2"><span class="style2">單位代碼：</span><span class="style1"><%
				if trim(rs2("PBCode"))<>"" and not isnull(rs2("PBCode")) then
					response.write trim(rs2("PBCode")) 

					sql="select PGName from "&DBUser&".PGList where PGCode ='"&rs2("PBCode")&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("PGName")
					set rsPcode=nothing


				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">告發單種類：</span><span class="style1"><%
				if trim(rs2("RBType"))<>"" and not isnull(rs2("RBType")) then	
					response.write trim(rs2("RBType")) 
					if trim(rs2("RBType")) ="1" then 
					response.write "&nbsp;電腦製單"
					elseif trim(rs2("RBType")) ="2" then 
					response.write "&nbsp;手開單"
					end if
				else
					response.write "&nbsp;"
				end if
			%>
			</span></td>
		</tr>
		<tr>
			<td ><span class="style2">操作人員：</span><span class="style1"><%
				if trim(rs2("OPCODE"))<>"" and not isnull(rs2("OPCODE")) then
					response.write trim(rs2("OPCODE")) 

					sql="select OPName from "&DBUser&".OPER where OPCode ='"&rs2("OPCODE")&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("OPName")
					set rsPcode=nothing

				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">入案日期：</span><span class="style1"><%
				if trim(rs2("FinDate"))<>"" and not isnull(rs2("FinDate")) then
					response.write GetDate(rs2("FinDate"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">操作日期：</span><span class="style1"><%
				if trim(rs2("OPDate"))<>"" and not isnull(rs2("OPDate")) then
					response.write GetDate(rs2("OPDate"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			
		</tr>
		<tr>
			<td colspan="2"><span class="style2">入案檔名：</span><span class="style1"><%
				if trim(rs2("batChNo"))<>"" and not isnull(rs2("batChNo")) then
					response.write rs2("batChNo") 
				else
					response.write "&nbsp;"
				end if
			%>
			</span></td>
			<td ><span class="style2">郵局日期：</span><span class="style1"><%
				if trim(rs2("MailDate"))<>"" and not isnull(rs2("MailDate")) then
					response.write GetDate(rs2("MailDate"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">車籍狀態：</span><span class="style1"><%
				if trim(rs2("Errst_SV"))<>"" and not isnull(rs2("Errst_SV")) then
					response.write rs2("Errst_SV") 

					sql="select ERRName from "&DBUser&".ErrCode where ErrCode ='"&trim(rs2("Errst_SV"))&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("ERRName")
					set rsPcode=nothing

				else
					response.write "&nbsp;"
				end if
			%></span></td>
			
		</tr>
		<tr>
			<td ><span class="style2">掛號號碼：</span><span class="style1"><%
				if trim(rs2("MailSEQNO"))<>"" and not isnull(rs2("MailSEQNO")) then
					response.write rs2("MailSEQNO") 
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">駕籍狀態：</span><span class="style1"><%
				if trim(rs2("ErrST_SD"))<>"" and not isnull(rs2("ErrST_SD")) then
					response.write rs2("ErrST_SD") 
					sql="select ERRName from "&DBUser&".ErrCode where ErrCode ='"&rs2("ErrST_SD")&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("ERRName")
					set rsPcode=nothing
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">移送日期：</span><span class="style1"><%
				if trim(rs2("SendPrnDate"))<>"" and not isnull(rs2("SendPrnDate")) then
					response.write GetDate(rs2("SendPrnDate"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			
		</tr>
		<tr>
			<td ><span class="style2">法條金額１：</span><span class="style1"><%
				if trim(rs2("AMT1_Fin"))<>"" and not isnull(rs2("AMT1_Fin")) and trim(rs2("AMT1_Fin"))<>"0" then
					response.write rs2("AMT1_Fin") 
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">法條金額２：</span><span class="style1"><%
				if trim(rs2("AMT2_Fin"))<>"" and not isnull(rs2("AMT2_Fin")) and trim(rs2("AMT2_Fin"))<>"0" then
					response.write rs2("AMT2_Fin") 
				else
					response.write "&nbsp;"
				end if
			%>
			</span></td>
			<td ><span class="style2">法條金額３：</span><span class="style1"><%
				if trim(rs2("AMT3_Fin"))<>"" and not isnull(rs2("AMT3_Fin")) and trim(rs2("AMT3_Fin"))<>"0" then
					response.write rs2("AMT3_Fin") 
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">法條金額４：</span><span class="style1"><%
				if trim(rs2("AMT4_Fin"))<>"" and not isnull(rs2("AMT4_Fin")) and trim(rs2("AMT4_Fin"))<>"0" then
					response.write rs2("AMT4_Fin") 
				else
					response.write "&nbsp;"
				end if
			%>	
			</span></td>
		</tr>
		<tr>
			<td colspan="2"><span class="style2">違反牌照稅註記：</span><span class="style1"><%
				if trim(rs2("TAXFLAG"))<>"" and not isnull(rs2("TAXFLAG")) then
					response.write rs2("TAXFLAG") 
					if rs2("TAXFLAG") ="0" then
					  response.write "&nbsp;正常(無違反)"
					elseif rs2("TAXFLAG") ="1" then
					  response.write "&nbsp;違反牌照稅註記"
					end if
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td colspan="2"><span class="style2">送達註記：</span><span class="style1"><%
				if trim(rs2("PUBFLAG"))<>"" and not isnull(rs2("PUBFLAG")) then
					response.write rs2("PUBFLAG") 
					if rs2("PUBFLAG") ="1" then
					  response.write "&nbsp;公示送達"
					elseif rs2("PUBFLAG") ="2" then
					  response.write "&nbsp;寄存送達"
					elseif rs2("PUBFLAG") ="3" then
					  response.write "&nbsp;留置送達"
					end if
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			
		</tr>
<%		'刪除
		strD="select * from "&DBUser&".FinDel where FSEQ='"&Trim(rs2("FSEQ"))&"' order by Del_OPDate"
		delint=0
		Set rsDel=conn.execute(strD)
		while Not rsDel.eof
			delint=delint+1
%>
		<tr>
			<td colspan="4"><hr><br><strong>刪除資料(<%=delint%>)</strong></td>
		</tr>
		<tr>
			<td ><span class="style2">刪除狀態：</span><span class="style1"><%
				if trim(rsDel("Del_FStatus"))<>"" and not isnull(rsDel("Del_FStatus")) then
					response.write trim(rsDel("Del_FStatus"))
					if rsDel("Del_FStatus") ="0" then
					  response.write "&nbsp;未上傳"
					elseif rsDel("Del_FStatus") ="2" then
					  response.write "&nbsp;已上傳"
					elseif rsDel("Del_FStatus") ="S" then
					  response.write "&nbsp;上傳成功"
					elseif rsDel("Del_FStatus") ="N" then
					  response.write "&nbsp;無此資料"
					elseif rsDel("Del_FStatus") ="n" then
					  response.write "&nbsp;已結案不做刪除"
					elseif rsDel("Del_FStatus") ="B" then
					  response.write "&nbsp;無此車號/無此證號"
					elseif rsDel("Del_FStatus") ="Z" then
					  response.write "&nbsp;不可刪除"
					end if
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">刪除人員：</span><span class="style1"><%
				if trim(rsDel("Del_OPCode"))<>"" and not isnull(rsDel("Del_OPCode")) then
					response.write trim(rsDel("Del_OPCode"))

					sql="select OPName from "&DBUser&".OPER where OPCode ='"&rsDel("Del_OPCode")&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("OPName")
					set rsPcode=nothing
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">刪除日期：</span><span class="style1"><%
				if trim(rsDel("Del_OPDate"))<>"" and not isnull(rsDel("Del_OPDate")) then
					response.write GetDate(rsDel("Del_OPDate"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">刪除上傳檔名：</span><span class="style1"><%
				if trim(rsDel("Del_BatChNo"))<>"" and not isnull(rsDel("Del_BatChNo")) then
					response.write trim(rsDel("Del_BatChNo"))
				else
					response.write "&nbsp;"
				end if
			%>
			</span></td>
		</tr>
		<%
		rsDel.movenext
		wend
		rsDel.close
		Set rsDel=Nothing 
	If DBUser="Xtraffic3" Then
		
	else
		'traffic單退
		strB="select FSEQ,BackDate,MailDate,MailNo,MailSEQNo,SendPrnDate,MailDate2,BackCode,BackDate2,MailNo2,ShowDate,CloseDate,BackCode2,PubType,Opdate,FStatus,BatChNo,FStatus_P,BatChNo_P,SEQNO,PubDate from "&DBUser&".FinBack where FSEQ='"&Trim(rs2("FSEQ"))&"' "
		Set rsBack=conn.execute(strB)
		while Not rsBack.eof
%>
		<tr>
			<td colspan="4"><hr><br><strong>退件資料</strong><br>第一次退件資料</td>
		</tr>
		<tr>
			<td ><span class="style2">退件日期：</span><span class="style1"><%
				if trim(rsBack("BackDate"))<>"" and not isnull(rsBack("BackDate")) then
					response.write GetDate(rsBack("BackDate"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">郵寄日期：</span><span class="style1"><%
				if trim(rsBack("MailDate"))<>"" and not isnull(rsBack("MailDate")) then
					response.write GetDate(rsBack("MailDate"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">貼條號碼：</span><span class="style1"><%
				if trim(rsBack("MailNo"))<>"" and not isnull(rsBack("MailNo")) then
					response.write trim(rsBack("MailNo"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">郵寄序號：</span><span class="style1"><%
				if trim(rsBack("MailSEQNo"))<>"" and not isnull(rsBack("MailSEQNo")) then
					response.write trim(rsBack("MailSEQNo"))
				else
					response.write "&nbsp;"
				end if
			%>
			</span></td>
		</tr>
		<tr>
			<td ><span class="style2">移送日期：</span><span class="style1"><%
				if trim(rsBack("SendPrnDate"))<>"" and not isnull(rsBack("SendPrnDate")) then
					response.write GetDate(rsBack("SendPrnDate"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">二次郵寄日期：</span><span class="style1"><%
				if trim(rsBack("MailDate2"))<>"" and not isnull(rsBack("MailDate2")) then
					response.write GetDate(rsBack("MailDate2"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">退件原因：</span><span class="style1"><%
				if trim(rsBack("BackCode"))<>"" and not isnull(rsBack("BackCode")) then
					response.write trim(rsBack("BackCode"))
					sql="select BACKName from "&DBUser&".BACKCODE where BACKCODE ='"&rsBack("BackCode")&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("BACKName")
					set rsPcode=nothing
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">：</span><span class="style1"><%
				
			%>
			</span></td>
		</tr>
		<tr>
			<td colspan="4"><span class="style2">第二次退件資料</span></td>
		</tr>
		<tr>
			<td ><span class="style2">退件日期：</span><span class="style1"><%
				if trim(rsBack("BackDate2"))<>"" and not isnull(rsBack("BackDate2")) then
					response.write GetDate(rsBack("BackDate2"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">貼條號碼：</span><span class="style1"><%
				if trim(rsBack("MailNo2"))<>"" and not isnull(rsBack("MailNo2")) then
					response.write trim(rsBack("MailNo2"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">送達日期：</span><span class="style1"><%
				if trim(rsBack("ShowDate"))<>"" and not isnull(rsBack("ShowDate")) then
					response.write GetDate(rsBack("ShowDate"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">送達完成日期：</span><span class="style1"><%
				if trim(rsBack("CloseDate"))<>"" and not isnull(rsBack("CloseDate")) then
					response.write GetDate(rsBack("CloseDate"))
				else
					response.write "&nbsp;"
				end If
			%></span></td>
		</tr>
		<tr>
			<td colspan="6"><span class="style2">退件原因：</span><span class="style1"><%
				if trim(rsBack("BackCode2"))<>"" and not isnull(rsBack("BackCode2")) then
					response.write trim(rsBack("BackCode2"))
					sql="select BACKName from BACKCODE where BACKCODE ='"&rsBack("BackCode2")&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("BACKName")
					set rsPcode=nothing
				else
					response.write "&nbsp;"
				end if
			%></span></td>
		</tr>
		<tr>
			<td colspan="4"><span class="style2">退件／送達上傳資料</span></td>
		</tr>
		<tr>
			<td ><span class="style2">資料類別：</span><span class="style1"><%
				if trim(rsBack("PubType"))<>"" and not isnull(rsBack("PubType")) then
					response.write trim(rsBack("PubType"))
					if rsBack("PubType") ="1" then
					  response.write "&nbsp;公示送達"
					elseif rsBack("PubType") ="2" then
					  response.write "&nbsp;寄存送達"
					end if
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">操作日期：</span><span class="style1"><%
				if trim(rsBack("Opdate"))<>"" and not isnull(rsBack("Opdate")) then
					response.write GetDate(rsBack("Opdate"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td colspan="2"><span class="style2">退件上傳註記：</span><span class="style1"><%
				if trim(rsBack("FStatus"))<>"" and not isnull(rsBack("FStatus")) then
					response.write trim(rsBack("FStatus"))&"&nbsp;"&getDciCodeN(rsBack("FStatus"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
		
		</tr>
		<tr>
			<td colspan="2"><span class="style2">退件上傳檔名：</span><span class="style1"><%
				if trim(rsBack("BatChNo"))<>"" and not isnull(rsBack("BatChNo")) then
					response.write trim(rsBack("BatChNo"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td colspan="2"><span class="style2">送達上傳註記：</span><span class="style1"><%
				if trim(rsBack("FStatus_P"))<>"" and not isnull(rsBack("FStatus_P")) then
					response.write trim(rsBack("FStatus_P"))&"&nbsp;"&getDciCodeN(rsBack("FStatus_P"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
		
		</tr>
		<tr>
			<td colspan="2"><span class="style2">送達上傳檔名：</span><span class="style1"><%
				if trim(rsBack("BatChNo_P"))<>"" and not isnull(rsBack("BatChNo_P")) then
					response.write trim(rsBack("BatChNo_P"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">送達書號：</span><span class="style1"><%
				if trim(rsBack("SEQNO"))<>"" and not isnull(rsBack("SEQNO")) then
					response.write trim(rsBack("SEQNO"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>
			<td ><span class="style2">送達生效日：</span><span class="style1"><%
				if trim(rsBack("PubDate"))<>"" and not isnull(rsBack("PubDate")) then
					response.write GetDate(rsBack("PubDate"))
				else
					response.write "&nbsp;"
				end if
			%></span></td>

		</tr>

<%
			rsBack.movenext
		wend
		rsBack.close
		Set rsBack=Nothing 
	End If 
%>
	</table>
<%
			End If
			rs2.close
			Set rs2=Nothing 
			'===============================================================================================
		End If 
		
	rsOld1.MoveNext
	Wend
	rsOld1.close
	set rsOld1=nothing
%>
<Div id="Layer111" style="width:1041px; height:24px; ">
  <div align="center">
  <input type="hidden" value="" name="IsShow">
  <input type="button" value="列印" onclick="DP();">
  <br>
   <%if (sys_City<>"高雄市" and sys_City<>ApconfigureCityName) then%>
    (若無列印鈕，可按下滑鼠右鍵選擇列印功能，格式為A4橫印)
	<%end if%>
  </div>
</Div>

	</form>
<%
conn.close
set conn=nothing
%>
</body>

<script language="JavaScript">
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
		var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
		win.focus();
		return win;
}
function OpenImageWin(ImgSN,illdate){
	urlstr='../ProsecutionImage/ShowIllImage.asp?ImgSN='+ImgSN+'&illdate='+illdate;
	newWin(urlstr,'MyDetail',1000,600,0,0,"yes","no","yes","no");
}

function showMailHistory(){
	
	myForm.IsShow.value="1";
	myForm.submit();
}

function hiddenMailHistory(){
	
	myForm.IsShow.value="0";
	myForm.submit();
}

function DP(){
<%if (sys_City="高雄市" or sys_City=ApconfigureCityName) and NoCase=0 then%>
	urlstr='BillBaseData_Detail_Print_Set.asp?BillSnTmp=<%=BillSnTmp%>';
	newWin(urlstr,'Billprint',350,400,300,150,"no","no","yes","no");
<%else%>
	window.focus();
	<%if NewCase>=1 then%>
	Layer112.style.visibility="hidden";
	<%end if%>
	Layer111.style.visibility="hidden";
	window.print();
	window.close();
<%end if%>
}
//開啟檢視圖
function OpenPic(FileName){
//alert(FileName);
	window.open("ShowIllegalImage.asp?FileName="+FileName,"UploadFile","left=0,top=0,location=0,width=910,height=705,resizable=yes,status=yes,scrollbars=yes,menubar=yes")
}
function OpenImageWinUserUpload(ImgFileName){
	urlstr=ImgFileName;
	newWin(urlstr,'MyDetail',1000,600,0,0,"yes","no","yes","no");
}

</script>
</html>
