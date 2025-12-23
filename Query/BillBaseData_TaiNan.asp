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
<script type="text/javascript" src="../js/Print.js"></script>
<script type="text/javascript" src="../js/date.js"></script>
<%	
	Server.ScriptTimeout = 65000
	Response.flush

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
<%
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing
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
	'ConnExecute strQry,356
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
		<!-- <div class="PageNext">&nbsp;</div> -->
<%	end If
	Cnt=Cnt+1
	NewCase=NewCase+1
%>
	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td width="25%"><span class="style2">告發單號：</span><span class="style1"><%
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNO"))
			end if
			%></span></td>
			<td><span class="style2">車號：</span><span class="style1"><%=trim(rs1("CarNo"))%></span></td>
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
			<td colspan="2"><span class="style2">違規時間：</span><span class="style1"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gArrDT(trim(rs1("IllegalDate")))&"&nbsp;"
				response.write Right("00"&hour(rs1("IllegalDate")),2)&":"
				response.write Right("00"&minute(rs1("IllegalDate")),2)
			end if		
			%></span></td>
			<td colspan="2"><span class="style2">舉發單位：</span><span class="style1"><%
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
			<td><span class="style2">是否郵寄：</span><span class="style1"><%
			if trim(rs1("EquipMentID"))<>"" and not isnull(rs1("EquipMentID")) then
				if trim(rs1("EquipMentID"))="1" then
					response.write "是"
				else
					response.write "否"
				end if
			end if	
			%></span></td>
			<td colspan="3"><span class="style2">詳細車種：</span><span class="style1"><%=DciCarType%></span></td>
		</tr>

	</table>
	<hr>
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
		<!-- <div class="PageNext">&nbsp;</div> -->
<%	end If
	Cnt=Cnt+1
%>
	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td width="25%"><span class="style2">告發單號：</span><span class="style1"><%
			if trim(rs2("fseq"))<>"" and not isnull(rs2("fseq")) then
				response.write trim(rs2("fseq"))
			end if
			%></span></td>
			<td width="25%"><span class="style2">車號：</span><span class="style1"><%
			if trim(rs2("CarNo"))<>"" and not isnull(rs2("CarNo")) then
				response.write trim(rs2("CarNo"))
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
			<td ><span class="style2">單位代碼：</span><span class="style1"><%
				if trim(rs2("PBCode"))<>"" and not isnull(rs2("PBCode")) then
					response.write trim(rs2("PBCode")) 

					sql="select UnitName from "&DBUser&".Unitid where PBCode ='"&rs2("PBCode")&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("UnitName")
					set rsPcode=nothing


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
			<td ><span class="style2">詳細車種代碼：</span><span class="style1"><%
				if trim(rs2("CDType"))<>"" and not isnull(rs2("CDType")) then
					response.write trim(rs2("CDType")) 

					sql="select CDName from "&DBUser&".CARKIND where CDType ='"&rs2("CDType")&"'"
					set rsPCode=conn.execute(sql)
					if not rspcode.eof then response.write "&nbsp;"&rsPCode("CDName")
					set rsPcode=nothing

				else
					response.write "&nbsp;"
				end if
			%></span></td>
		</tr>
		
		<tr>	
			<td colspan="4" ><hr></td>
		<tr>

		<%

	
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
function OpenImageWin(ImgFileName){
	urlstr='../ProsecutionImage/ProsecutionImageDetail.asp?FileName='+ImgFileName.replace(/\+/g,'@2@')+'&SN=1';
	newWin(urlstr,'MyDetail',1000,600,0,0,"yes","no","yes","no");
}
function DP(){
	window.focus();
<%if Cnt=1 then%>
	Layer112.style.visibility="hidden";
<%end if%>
	Layer111.style.visibility="hidden";
	window.print();
	window.close();
}
</script>
</html>
