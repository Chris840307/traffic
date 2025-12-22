<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--

.TitleSet {
	background-image: url(Image/title.jpg);
	background-repeat: no-repeat;
}
#Layer1 {
	position:absolute;
	width:209px;
	height:38px;
	z-index:2;
	top: 13px;
}
.style1 {font-size: 14px}
.style2 {font-size: 13px; }
#Layer2 {
	position:absolute;
	width:566px;
	height:38px;
	z-index:3;
	top: 15px;
}
#LayerTime {
	position:absolute;
	width:311px;
	height:34px;
	z-index:1;
}
#Layer151 {
	position:absolute;
	width:209px;
	height:38px;
	z-index:2;
}
#D1 {
	BACKGROUND-COLOR: #99FFFF; 
	BORDER-BOTTOM: white 2px outset; 
	BORDER-LEFT: white 2px outset; 
	BORDER-RIGHT: white 2px outset; 
	BORDER-TOP: white 2px outset; 
	LEFT: 0px; POSITION: absolute; 
	TOP: 0px; VISIBILITY: hidden; 
	WIDTH: 150px; 
	layer-background-color: #99FFFF;
	
	z-index:5;
}
#Banner-Submenu {
	height: 21px;
	/*width: 840px;*/
	width: 960px;
	background: url(Image/menu_bg.gif) repeat-x;
	position: absolute;
	z-index: 3;
}
/*主選單*/
#Banner-Menu {
	height: 50px;
}
#Banner-Menu ul {
	padding: 0px;
	margin: 0px 15px 0px 0px;
	list-style-type: none;
}
#Banner-Menu ul li {
	font-family: "標楷體";
	font-size: 24px;
	font-weight: bold;
	color: #FFFF00;
	text-align: center;
	line-height: 50px;
	height: 50px;
	width: 214px;
	float: left;
	display: block;
	background: url(Image/title1.jpg) no-repeat;
}
#Banner-Menu ul li a {
	color: #005BB7;
	display: block;
}
#Banner-Menu ul li a:hover {
	color: #000000;
	text-decoration: none;	
}
#Banner-Menu ul li.current1 {
	background: url(Image/title1.jpg) no-repeat;
}
#Banner-Menu ul li.current2 {
	background: url(Image/title2.jpg) no-repeat;
}
#Banner-Menu ul li.current3 {
	background: url(Image/title1-2.jpg) no-repeat;
}
#Banner-Menu ul li.current4 {
	background: url(Image/title2-2.jpg) no-repeat;
}
#Banner-Menu ul li.current1 a {
	color: #CC3300;	
}


-->
</style>
<head>
<!--#include virtual="traffic/Common/css.txt"-->
<title>智慧型交通執法系統</title>
<%
memName=Session("Ch_Name")
GroupID=Session("Group_ID")
UnitNo=Session("Unit_ID")

Function IDNumber(strUID)
	If strUID <> "" And Len(strUID) = 10 Then
		strCheckID = true
		'如果第一位非半形26個英文字母
		If Asc(Left(strUID, 1)) <= 64 Or Asc(Left(strUID, 1)) >= 91 Then
			strCheckID = False
		'如果第二位非1或2 
		ElseIf Mid(strUID, 2, 1) <> "1" And Mid(strUID, 2, 1) <> "2" Then
			strCheckID = False
		Else
			'如果第3-10位非半形阿拉伯數字
			For i=3 To 10
				If Asc(Mid(strUID, i, 1)) <= 47 Or Asc(Mid(strUID, i, 1)) >= 58 Then
					strCheckID = False
				End If 
			Next
			If strCheckID = True Then
				'檢查台灣身分證字號排列規則
				ID_ABC_Data = "A10B11C12D13E14F15G16H17I34J18K19L20M21N22O35P23Q24R25S26T27U28V29W32X30Y31Z33" 
				strUID = Mid(ID_ABC_Data, InStr(ID_ABC_Data, Left(strUID, 1)) + 1, 2) & Mid(strUID, 2) 
				GetNo = 2 
				SUM = CInt(Left(strUID, 1)) + CInt(Right(strUID, 1))
				For i=9 To 1 Step -1 
					SUM = SUM + Mid(strUID, GetNo, 1) * i 
					GetNo = GetNo + 1 
				Next
				If SUM Mod 10 = 0 Then 
					strCheckID = True
				Else
					strCheckID = False
				End If
			End If 			
		End If
	Else
		'不是10碼
		strCheckID = False
	End If	
	IDNumber=strCheckID
End Function


 	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	if sys_City<>"台中縣" then
		ArgueDate1=DateAdd("d",-10,date) & " 0:0:0"
		ArgueDate2=date & " 23:29:59" 
		strDelErr="select * from Dcilog where ExchangeTypeID='E' and (DciReturnStatusID<>'S')" &_
			" and ExchangeDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS')" &_
			" and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and RecordMemberID="&Session("User_ID")
		set rsDelErr=conn.execute(strDelErr)
		If Not rsDelErr.Bof Then rsDelErr.MoveFirst 
		While Not rsDelErr.Eof
			if (trim(rsDelErr("DciReturnStatusID"))<>"S" and not isnull(rsDelErr("DciReturnStatusID"))) Then
				ISDel=1
				strCheckErr="select * from (select * from Dcilog where ExchangeTypeID='E' and billsn="&trim(rsDelErr("BillSn"))&" order by ExchangeDate Desc) where Rownum<=1 "
				Set rsCE=conn.execute(strCheckErr)
				If Not rsCE.eof Then 
					If trim(rsCE("DciReturnStatusID"))="S" or isnull(rsCE("DciReturnStatusID")) Then 
						ISDel=0
					End If 
				End If
				rsCE.close
				Set rsCE=Nothing 
				If ISDel=1 Then 
					if trim(rsDelErr("DciReturnStatusID"))="n" then
						strUpd="Update BillBase set RecordStateID=0,BillStatus='2' where RecordStateID<>0 and Sn="&trim(rsDelErr("BillSn"))
						conn.execute strUpd
					else
						strUpd="Update BillBase set RecordStateID=0,BillStatus='2' where RecordStateID<>0 and Sn="&trim(rsDelErr("BillSn"))
						conn.execute strUpd
					end If
				End If 
			end if
		rsDelErr.MoveNext
		Wend
		rsDelErr.close
		set rsDelErr=nothing
	end if
	

If Trim(request("SystemType"))="" Then
	

'1110801=====================================================================
	strChkL2="select * from Law where itemid ='3110009' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		strInsL2="insert into law values('3110009','0','汽車行駛於一般道路上營業大客車駕駛人未依規定繫安全帶',2000,2000,2000,2000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('3120023','0','汽車行駛於高速公路未依規定繫安全帶(二人以上)—營業大客車、計程車或租賃車輛有代僱駕駛人',6000,6000,6000,6000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('3120024','0','汽車行駛於快速公路未依規定繫安全帶(二人以上)—營業大客車、計程車或租賃車輛有代僱駕駛人',6000,6000,6000,6000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('3120025','0','營業大客車行駛於高速公路上其四歲以上乘客經告知仍未繫安全帶(罰乘客)',3000,3300,3900,4500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('3120026','0','營業大客車行駛於快速公路上其四歲以上乘客經告知仍未繫安全帶(罰乘客)',3000,3300,3900,4500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('18100051','1','汽車未依規定裝設防止捲入裝置',12000,13000,15000,18000,'V','0','0','0','0','5',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('18100051','2','汽車未依規定裝設防止捲入裝置',14000,15000,18000,21000,'V','0','0','0','0','5',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('18100061','0','二種以上設備同時違反第18條之1第1項規定',16000,17000,19000,22000,'V','0','0','0','0','5',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('18200051','0','汽車裝設之防止捲入裝置無法正常運作，未於行車前改善，仍繼續行駛',9000,9900,11000,13000,'V','0','0','0','0','5',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('18200061','0','二種以上設備同時違反第18條之1第2項規定',12000,13000,15000,18000,'V','0','0','0','0','5',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('8519903','0','違規處罰，以主要駕駛人為被通知人',0,0,0,0,'+','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('8519904','0','違規處罰，以主要駕駛人為被通知人，不記點',0,0,0,0,'+','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('8519905','0','違規處罰，以長租車租用人為被通知人',0,0,0,0,'+','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
	End if
	rsChkL2.close
	Set rsChkL2=Nothing
'1111130微電車=====================================================================
	strChkL2="select * from Law where itemid ='32000101' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		strInsL2="insert into law values('6920003','0','人力行駛車輛，未依規定辦理登記，領取證照即行駛道路',300,300,300,300,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6920004','0','獸力行駛車輛，未依規定辦理登記，領取證照即行駛道路',300,300,300,300,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6950001','0','個人行動器具未依直轄市、縣（市）政府所定規格、指定行駛路段、時間、速度限制、安全注意及其他管理事項規定行駛',1200,1600,1600,1600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6950002','0','個人行動器具未依直轄市、縣（市）政府所定規格、指定行駛路段、時間、速度限制、安全注意及其他管理事項規定行駛，肇事致人受傷',2400,3200,3200,3200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6950003','0','個人行動器具未依直轄市、縣（市）政府所定規格、指定行駛路段、時間、速度限制、安全注意及其他管理事項規定行駛，肇事致人重傷或死亡',3600,3600,3600,3600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6950004','0','個人行動器具違反道路交通管理處罰條例慢車章節規定',1200,1600,1600,1600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6950005','0','個人行動器具違反道路交通管理處罰條例慢車章節規定，肇事致人受傷',2400,3200,3200,3200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6950006','0','個人行動器具違反道路交通管理處罰條例慢車章節規定，肇事致人重傷或死亡',3600,3600,3600,3600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7000002','0','微型電動二輪車，經依規定淘汰並公告禁止行駛後仍行駛',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7000003','0','微型電動二輪車以外其他慢車，經依規定淘汰並公告禁止行駛後仍行駛',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7110002','0','經型式審驗合格，電動輔助自行車，未黏貼審驗合格標章，於道路行駛',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7120002','0','未經型式審驗合格，電動輔助自行車，於道路行駛',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210003','0','微型電動二輪車，未經核准，擅自變更裝置',1200,1200,1200,1200,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210004','0','微型電動二輪車以外其他慢車，未經核准，擅自變更裝置',300,500,500,500,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210005','0','微型電動二輪車，不依規定保持煞車之良好與完整',1200,1200,1200,1200,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210006','0','微型電動二輪車以外其他慢車，不依規定保持煞車之良好與完整',300,500,500,500,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210007','0','微型電動二輪車，不依規定保持鈴號之良好與完整',1200,1200,1200,1200,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210008','0','微型電動二輪車以外其他慢車，不依規定保持鈴號之良好與完整',300,500,500,500,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210009','0','微型電動二輪車，不依規定保持燈光之良好與完整',1200,1200,1200,1200,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210010','0','微型電動二輪車以外其他慢車，不依規定保持燈光之良好與完整',300,500,500,500,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210011','0','微型電動二輪車，不依規定保持反光裝置之良好與完整',1200,1200,1200,1200,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210012','0','微型電動二輪車以外其他慢車，不依規定保持反光裝置之良好與完整',300,500,500,500,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7220003','0','微型電動二輪車，於道路行駛或使用，擅自增、減、變更行駛速率以外之電子控制裝置或原有規格',2500,2700,2700,2700,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7220004','0','電動輔助自行車，於道路行駛或使用，擅自增、減、變更行駛速率以外之電子控制裝置或原有規格',1800,2000,2000,2000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7220005','0','微型電動二輪車，於道路行駛或使用，擅自增、減、變更與行駛速率相關之電子控制裝置或原有規格',5400,5400,5400,5400,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7220006','0','電動輔助自行車，於道路行駛或使用，擅自增、減、變更與行駛速率相關之電子控制裝置或原有規格',5400,5400,5400,5400,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310105','0','微型電動二輪車，不在劃設之慢車道通行',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310106','0','微型電動二輪車以外其他慢車，不在劃設之慢車道通行',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310107','0','微型電動二輪車，無正當理由在未劃設慢車道之道路不靠右側路邊行駛',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310108','0','微型電動二輪車以外其他慢車，無正當理由在未劃設慢車道之道路不靠右側路邊行駛',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310203','0','微型電動二輪車，不在規定之地區路線行駛',400,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310204','0','微型電動二輪車以外其他慢車，不在規定之地區路線行駛',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310205','0','微型電動二輪車，不在規定時間內行駛',400,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310206','0','微型電動二輪車以外其他慢車，不在規定時間內行駛',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310303','0','微型電動二輪車，不依規定轉彎',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310304','0','微型電動二輪車以外其他慢車，不依規定轉彎',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310305','0','微型電動二輪車，不依規定超車',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310306','0','微型電動二輪車以外其他慢車，不依規定超車',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310307','0','微型電動二輪車，不依規定停車',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310308','0','微型電動二輪車以外其他慢車，不依規定停車',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310309','0','微型電動二輪車，不依規定通過交岔路口',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310310','0','微型電動二輪車以外其他慢車，不依規定通過交岔路口',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310403','0','微型電動二輪車，在道路上爭先、爭道',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310404','0','微型電動二輪車以外其他慢車，在道路上爭先、爭道',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310405','0','微型電動二輪車，在道路上以其他危險方式駕車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310406','0','微型電動二輪車以外其他慢車，在道路上以其他危險方式駕車',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310503','0','微型電動二輪車，在夜間行車未開啟燈光',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310504','0','微型電動二輪車以外其他慢車，在夜間行車未開啟燈光',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310607','0','微型電動二輪車，以手持方式使用行動電話，進行撥接、通話、數據通訊',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310608','0','微型電動二輪車以外其他慢車，以手持方式使用行動電話，進行撥接、通話、數據通訊',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310609','0','微型電動二輪車，以手持方式使用行動電話有礙駕駛安全之行為',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310610','0','微型電動二輪車以外其他慢車，以手持方式使用行動電話有礙駕駛安全之行為',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310611','0','微型電動二輪車，以手持方式使用電腦，進行撥接、通話、數據通訊',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310612','0','微型電動二輪車以外其他慢車，以手持方式使用電腦，進行撥接、通話、數據通訊',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310613','0','微型電動二輪車，以手持方式使用電腦有礙駕駛安全之行為',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310614','0','微型電動二輪車以外其他慢車，以手持方式使用電腦有礙駕駛安全之行為',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310615','0','微型電動二輪車，以手持方式使用其他相類功能裝置進行撥接、通話、數據通訊',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310616','0','微型電動二輪車以外其他慢車，以手持方式使用其他相類功能裝置進行撥接、通話、數據通訊',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310617','0','微型電動二輪車，以手持方式使用其他相類功能裝置有礙駕駛安全之行為',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310618','0','微型電動二輪車以外其他慢車，以手持方式使用其他相類功能裝置有礙駕駛安全之行為',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320005','0','微型電動二輪車，駕駛人吐氣酒精濃度達每公升0.15毫克以上，未滿0.25毫克或血液中酒精濃度達百分之0.03以上，未滿0.05',1600,1800,1800,1800,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320006','0','微型電動二輪車以外其他慢車，駕駛人吐氣酒精濃度達每公升0.15毫克以上，未滿0.25毫克或血液中酒精濃度達百分之0.03以上，未滿0.05',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320007','0','微型電動二輪車，駕駛人吐氣酒精濃度達每公升0.25毫克以上或血液中酒精濃度達百分之0.05以上',2400,2400,2400,2400,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320008','0','微型電動二輪車以外其他慢車，駕駛人吐氣酒精濃度達每公升0.25毫克以上或血液中酒精濃度達百分之0.05以上',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320009','0','微型電動二輪車，經測試檢定，有吸食毒品',2400,2400,2400,2400,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320010','0','微型電動二輪車以外其他慢車，經測試檢定，有吸食毒品',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320011','0','微型電動二輪車，經測試檢定，有吸食迷幻藥',2400,2400,2400,2400,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320012','0','微型電動二輪車以外其他慢車，經測試檢定，有吸食迷幻藥',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320013','0','微型電動二輪車，經測試檢定，有吸食麻醉藥品',2400,2400,2400,2400,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320014','0','微型電動二輪車以外其他慢車，經測試檢定，有吸食麻醉藥品',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320015','0','微型電動二輪車，經測試檢定，有吸食管制藥品',2400,2400,2400,2400,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320016','0','微型電動二輪車以外其他慢車，經測試檢定，有吸食管制藥品',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7330002','0','微型電動二輪車，駕駛人拒絕接受酒精濃度測試之檢定',4800,4800,4800,4800,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7330003','0','微型電動二輪車以外其他慢車，駕駛人拒絕接受酒精濃度測試之檢定',4800,4800,4800,4800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7340002','0','微型電動二輪車，駕駛人未依規定戴安全帽',300,300,300,300,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410102','0','微型電動二輪車，不服從執行交通勤務警察之指揮',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410103','0','微型電動二輪車以外其他慢車，不服從執行交通勤務警察之指揮',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410104','0','微型電動二輪車，不依標誌之指示',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410105','0','微型電動二輪車以外其他慢車，不依標誌之指示',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410106','0','微型電動二輪車，不依標線之指示',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410107','0','微型電動二輪車以外其他慢車，不依標線之指示',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410108','0','微型電動二輪車，不依號誌之指示',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410109','0','微型電動二輪車以外其他慢車，不依號誌之指示',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410202','0','微型電動二輪車，在同一慢車道上，不按遵行之方向行駛',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410203','0','微型電動二輪車以外其他慢車，在同一慢車道上，不按遵行之方向行駛',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410302','0','微型電動二輪車，不依規定，擅自穿越快車道',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410303','0','微型電動二輪車以外其他慢車，不依規定，擅自穿越快車道',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410402','0','微型電動二輪車，不依規定停放車輛',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410403','0','微型電動二輪車以外其他慢車，不依規定停放車輛',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410502','0','微型電動二輪車，違規行駛人行道',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410503','0','微型電動二輪車以外其他慢車，違規行駛人行道',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410504','0','微型電動二輪車，在快車道行駛',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410505','0','微型電動二輪車以外其他慢車，在快車道行駛',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410603','0','微型電動二輪車，聞消防車、警備車、救護車、工程救險車、毒性化學物質災害事故應變車之警號不立即避讓',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410604','0','微型電動二輪車以外其他慢車，聞消防車、警備車、救護車、工程救險車、毒性化學物質災害事故應變車之警號不立即避讓',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410703','0','微型電動二輪車，行經行人穿越道有行人穿越時，未讓行人優先通行',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410704','0','微型電動二輪車以外其他慢車，行經行人穿越道有行人穿越時，未讓行人優先通行',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410705','0','微型電動二輪車，行駛至交岔路口轉彎時，未讓行人優先通行',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410706','0','微型電動二輪車以外其他慢車，行駛至交岔路口轉彎時，未讓行人優先通行',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410803','0','微型電動二輪車，於設置有必要之標誌或標線供慢車行駛之人行道上，未讓行人優先通行',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410804','0','微型電動二輪車以外其他慢車，於設置有必要之標誌或標線供慢車行駛之人行道上，未讓行人優先通行',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410903','0','微型電動二輪車，聞或見大眾捷運系統車輛之聲號或燈光，不依規定避讓',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410904','0','微型電動二輪車以外其他慢車，聞或見大眾捷運系統車輛之聲號或燈光，不依規定避讓',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410905','0','微型電動二輪車，聞或見大眾捷運系統車輛之聲號或燈光，在後跟隨迫近',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410906','0','微型電動二輪車以外其他慢車，聞或見大眾捷運系統車輛之聲號或燈光，在後跟隨迫近',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7420002','0','微型電動二輪車，行近行人穿越道，遇有攜帶白手杖或導盲犬之視覺功能障礙者時，不暫停讓視覺功能障礙者先行通過',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7420003','0','微型電動二輪車以外其他慢車，行近行人穿越道，遇有攜帶白手杖或導盲犬之視覺功能障礙者時，不暫停讓視覺功能障礙者先行通過',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430003','0','微型電動二輪車，違規行駛人行道，導致視覺功能障礙者受傷',1600,1800,1800,1800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430004','0','微型電動二輪車以外其他慢車，違規行駛人行道，導致視覺功能障礙者受傷',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430005','0','微型電動二輪車，行駛快車道，導致視覺功能障礙者受傷',1600,1800,1800,1800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430006','0','微型電動二輪車以外其他慢車，行駛快車道，導致視覺功能障礙者受傷',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430007','0','微型電動二輪車，於設置有必要之標誌或標線供慢車行駛之人行道上，未讓行人優先通行，導致視覺功能障礙者受傷',1600,1800,1800,1800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430008','0','微型電動二輪車以外其他慢車，於設置有必要之標誌或標線供慢車行駛之人行道上，未讓行人優先通行，導致視覺功能障礙者受傷',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430009','0','微型電動二輪車，違規行駛人行道，導致視覺功能障礙者死亡',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430010','0','微型電動二輪車以外其他慢車，違規行駛人行道，導致視覺功能障礙者死亡',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430011','0','微型電動二輪車，行駛快車道，導致視覺功能障礙者死亡',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430012','0','微型電動二輪車以外其他慢車，行駛快車道，導致視覺功能障礙者死亡',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430013','0','微型電動二輪車，於設置有必要之標誌或標線供慢車行駛之人行道上，未讓行人優先通行，導致視覺功能障礙者死亡',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430014','0','微型電動二輪車以外其他慢車，於設置有必要之標誌或標線供慢車行駛之人行道上，未讓行人優先通行，導致視覺功能障礙者死亡',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500001','0','微型電動二輪車，在鐵路平交道，不遵看守人員指示，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500002','0','微型電動二輪車以外其他慢車，在鐵路平交道，不遵看守人員指示，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500003','0','微型電動二輪車，在鐵路平交道，警鈴已響、閃光號誌已顯示，或遮斷器開始放下，仍強行闖越，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500004','0','微型電動二輪車以外其他慢車，在鐵路平交道，警鈴已響、閃光號誌已顯示，或遮斷器開始放下，仍強行闖越，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500005','0','微型電動二輪車，在無看守人員管理或無遮斷器、警鈴及閃光號誌設備之鐵路平交道，設有警告標誌或跳動路面，不依規定暫停，逕行通過，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500006','0','微型電動二輪車以外其他慢車，在無看守人員管理或無遮斷器、警鈴及閃光號誌設備之鐵路平交道，設有警告標誌或跳動路面，不依規定暫停，逕行通過，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500007','0','微型電動二輪車，在鐵路平交道超車，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500008','0','微型電動二輪車以外其他慢車，在鐵路平交道超車，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500009','0','微型電動二輪車，在鐵路平交道迴車，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500010','0','微型電動二輪車以外其他慢車，在鐵路平交道迴車，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500011','0','微型電動二輪車，在鐵路平交道倒車，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500012','0','微型電動二輪車以外其他慢車，在鐵路平交道倒車，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500013','0','微型電動二輪車，在鐵路平交道臨時停車，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500014','0','微型電動二輪車以外其他慢車，在鐵路平交道臨時停車，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500015','0','微型電動二輪車，在鐵路平交道停車，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500016','0','微型電動二輪車以外其他慢車，在鐵路平交道停車，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500017','0','微型電動二輪車，在鐵路平交道，不遵看守人員指示，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500018','0','微型電動二輪車以外其他慢車，在鐵路平交道，不遵看守人員指示，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500019','0','微型電動二輪車，在鐵路平交道，警鈴已響、閃光號誌已顯示，或遮斷器開始放下，仍強行闖越，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500020','0','微型電動二輪車以外其他慢車，在鐵路平交道，警鈴已響、閃光號誌已顯示，或遮斷器開始放下，仍強行闖越，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500021','0','微型電動二輪車，在無看守人員管理或無遮斷器、警鈴及閃光號誌設備之鐵路平交道，設有警告標誌或跳動路面，不依規定暫停，逕行通過，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500022','0','微型電動二輪車以外其他慢車，在無看守人員管理或無遮斷器、警鈴及閃光號誌設備之鐵路平交道，設有警告標誌或跳動路面，不依規定暫停，逕行通過，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500023','0','微型電動二輪車，在鐵路平交道超車，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500024','0','微型電動二輪車以外其他慢車，在鐵路平交道超車，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500025','0','微型電動二輪車，在鐵路平交道迴車，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500026','0','微型電動二輪車以外其他慢車，在鐵路平交道迴車，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500027','0','微型電動二輪車，在鐵路平交道倒車，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500028','0','微型電動二輪車以外其他慢車，在鐵路平交道倒車，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500029','0','微型電動二輪車，在鐵路平交道臨時停車，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500030','0','微型電動二輪車以外其他慢車，在鐵路平交道臨時停車，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500031','0','微型電動二輪車，在鐵路平交道停車，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500032','0','微型電動二輪車以外其他慢車，在鐵路平交道停車，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610102','0','微型電動二輪車，慢車乘坐人數超過規定數額',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610103','0','微型電動二輪車以外其他慢車，乘坐人數超過規定數額',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610202','0','微型電動二輪車，裝載貨物超過規定重量',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610203','0','微型電動二輪車以外其他慢車，裝載貨物超過規定重量',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610204','0','微型電動二輪車，裝載貨物超出車身一定限制',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610205','0','微型電動二輪車以外其他慢車，裝載貨物超出車身一定限制',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610302','0','微型電動二輪車，裝載容易滲漏、飛散、有惡臭氣味貨物',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610303','0','微型電動二輪車以外其他慢車，裝載容易滲漏、飛散、有惡臭氣味貨物',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610304','0','微型電動二輪車，裝載危險性貨物不嚴密封固或不為適當之裝置',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610305','0','微型電動二輪車以外其他慢車，裝載危險性貨物不嚴密封固或不為適當之裝置',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610402','0','微型電動二輪車，裝載禽、畜重疊',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610403','0','微型電動二輪車以外其他慢車，裝載禽、畜重疊',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610404','0','微型電動二輪車，裝載禽、畜倒置',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610405','0','微型電動二輪車以外其他慢車，裝載禽、畜倒置',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610502','0','微型電動二輪車，裝載貨物不捆紮結實',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610503','0','微型電動二輪車以外其他慢車，裝載貨物不捆紮結實',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610602','0','微型電動二輪車，上、下乘客不緊靠路邊妨礙交通',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610603','0','微型電動二輪車以外其他慢車，上、下乘客不緊靠路邊妨礙交通',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610604','0','微型電動二輪車，裝卸貨物不緊靠路邊妨礙交通',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610605','0','微型電動二輪車以外其他慢車，裝卸貨物不緊靠路邊妨礙交通',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610703','0','微型電動二輪車，牽引其他車輛',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610704','0','微型電動二輪車以外其他慢車，牽引其他車輛',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610705','0','微型電動二輪車，攀附車輛隨行',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610706','0','微型電動二輪車以外其他慢車，攀附車輛隨行',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620102','0','腳踏自行車，附載幼童，駕駛人未滿18歲',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620103','0','電動輔助自行車，附載幼童，駕駛人未滿18歲',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620202','0','腳踏自行車，附載之幼童年齡超過規定',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620203','0','電動輔助自行車，附載之幼童年齡超過規定',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620204','0','腳踏自行車，附載之幼童體重超過規定',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620205','0','電動輔助自行車，附載之幼童體重超過規定',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620302','0','腳踏自行車或電動輔助自行車，附載幼童，不依規定使用合格之兒童座椅',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620303','0','附載幼童，不依規定使用合格腳踏自行車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620304','0','附載幼童，不依規定使用合格之電動輔助自行車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620402','0','腳踏自行車，附載幼童，違反第76條第2項第1款至第3款以外附載幼童之規定',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620403','0','電動輔助自行車，附載幼童，違反第76條第2項第1款至第4款以外附載幼童之規定',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7810102','0','行人不依標誌標線號誌之指示或警察指揮',500,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7810203','0','行人不在劃設之人行道通行',500,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7810204','0','無正當理由在未劃設人行道之道路不靠邊通行',500,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7810302','0','行人不依規定擅自穿越車道',500,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7810402','0','行人於交通頻繁之道路或鐵路平交道附近阻礙交通',500,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000101','0','非屬汽車、動力機械及個人行動器具範圍之動力載具於快車道以外之道路範圍行駛或使用',1200,1300,1400,1500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000111','0','非屬汽車、動力機械及個人行動器具範圍之動力運動休閒器材於快車道以外之道路範圍行駛或使用',1200,1300,1400,1500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000121','0','非屬汽車、動力機械及個人行動器具範圍之其他相類之動力器具於快車道以外之道路範圍行駛或使用',1200,1300,1400,1500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000131','0','非屬汽車、動力機械及個人行動器具範圍之動力載具於快車道行駛或使用',2000,2200,2400,2600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000141','0','非屬汽車、動力機械及個人行動器具範圍之動力運動休閒器材於快車道行駛或使用',2000,2200,2400,2600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000151','0','非屬汽車、動力機械及個人行動器具範圍之其他相類之動力器具於快車道行駛或使用',2000,2200,2400,2600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000161','0','非屬汽車、動力機械及個人行動器具範圍之動力載具於道路上行駛或使用因而肇事',2800,3000,3300,3600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000171','0','非屬汽車、動力機械及個人行動器具範圍之動力運動休閒器材於道路上行駛或使用因而肇事',2800,3000,3300,3600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000181','0','非屬汽車、動力機械及個人行動器具範圍之其他相類之動力器具於道路上行駛或使用因而肇事',2800,3000,3300,3600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71101011','0','經型式審驗合格，微型電動二輪車，未依規定領用牌照，於道路行駛',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71101021','0','未經型式審驗合格，微型電動二輪車，於道路行駛',3600,3600,3600,3600,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71102011','0','經型式審驗合格，微型電動二輪車，使用偽造或變造之牌照，於道路行駛',3600,3600,3600,3600,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71102021','0','未經型式審驗合格，微型電動二輪車，使用偽造或變造之牌照，於道路行駛',3600,3600,3600,3600,'V','0','0','0','0','8、1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71103011','0','經型式審驗合格，微型電動二輪車，牌照借供他車使用，於道路行駛',3600,3600,3600,3600,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71103021','0','經型式審驗合格，微型電動二輪車，使用他車牌照，於道路行駛',3600,3600,3600,3600,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71103031','0','未經型式審驗合格，微型電動二輪車，使用他車牌照，於道路行駛',3600,3600,3600,3600,'V','0','0','0','0','8、1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71104011','0','經型式審驗合格，微型電動二輪車，已領有牌照而未懸掛，於道路行駛',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71104021','0','經型式審驗合格，微型電動二輪車，已領有牌照而不依指定位置懸掛，於道路行駛',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71105011','0','經型式審驗合格，微型電動二輪車，牌照業經註銷，仍懸掛該註銷牌照行駛道路',1800,2000,2000,2000,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71105021','0','經型式審驗合格，微型電動二輪車，牌照業經註銷，無牌照行駛道路',1800,2000,2000,2000,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71106011','0','經型式審驗合格，微型電動二輪車，牌照遺失不報請該管主管機關補發，經舉發後仍不辦理而行駛道路',1200,1400,1400,1400,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300011','0','經型式審驗合格，微型電動二輪車，未依規定領用牌照，於道路停車',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300021','0','未經型式審驗合格，微型電動二輪車，於道路停車',3600,3600,3600,3600,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300031','0','經型式審驗合格，微型電動二輪車，使用偽造或變造之牌照，於道路停車',3600,3600,3600,3600,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300041','0','未經型式審驗合格，微型電動二輪車，使用偽造或變造之牌照，於道路停車',3600,3600,3600,3600,'V','0','0','0','0','8、1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300051','0','經型式審驗合格，微型電動二輪車，使用他車牌照，於道路停車',3600,3600,3600,3600,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300061','0','未經型式審驗合格，微型電動二輪車，使用他車牌照，於道路停車',3600,3600,3600,3600,'V','0','0','0','0','8、1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300071','0','經型式審驗合格，微型電動二輪車，已領有牌照而未懸掛，於道路停車',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300081','0','經型式審驗合格，微型電動二輪車，已領有牌照而不依指定位置懸掛，於道路停車',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300091','0','經型式審驗合格，微型電動二輪車，牌照業經註銷，仍懸掛該註銷牌照於道路停車',1500,1700,1700,1700,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300101','0','經型式審驗合格，微型電動二輪車，牌照業經註銷，無牌照於道路停車',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300111','0','經型式審驗合格，微型電動二輪車，牌照遺失不報請該管主管機關補發，經舉發後仍不辦理，於道路停車',1200,1400,1400,1400,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71400011','0','經型式審驗合格並黏貼審驗合格標章，微型電動二輪車，未於本條例111年4月19日修正施行後2年內依規定登記、領用、懸掛牌照，於道路行駛',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71100012','0','微型電動二輪車，損毀牌照，使不能辨認其牌號',900,1800,1800,1800,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71100022','0','微型電動二輪車，變造牌照，使不能辨認其牌號',900,1800,1800,1800,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71100032','0','微型電動二輪車，塗抹污損牌照，使不能辨認其牌號',900,1800,1800,1800,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71100042','0','微型電動二輪車，安裝其他器具之方式，使不能辨認其牌號',900,1800,1800,1800,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71201012','0','微型電動二輪車，牌照遺失，不報請補發',300,300,300,300,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71201022','0','微型電動二輪車，牌照破損，不報請換發或重新申請',300,300,300,300,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71202012','0','微型電動二輪車，牌照污穢，不洗刷清楚，非行車途中因遇雨、雪道路泥濘所致',150,300,300,300,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71202022','0','微型電動二輪車，牌照為他物遮蔽',150,300,300,300,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72000041','0','微型電動二輪車，於道路行駛或使用，行駛速率超過每小時25公里，未超過35公里',900,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72000051','0','微型電動二輪車，於道路行駛或使用，行駛速率超過每小時35公里，未超過45公里',1200,1500,1500,1500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72000061','0','微型電動二輪車，於道路行駛或使用，行駛速率超過每小時45公里',1500,1800,1800,1800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72100012','0','未滿14歲之人，駕駛微型電動二輪車',1000,1200,1200,1200,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72100022','0','未滿14歲之人，駕駛個人行動器具',600,800,800,800,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72200012','0','微型電動二輪車租賃業者，未於租借予駕駛人前，教導駕駛人車輛操作方法及道路行駛規定',800,1200,1200,1200,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72200022','0','個人行動器具租賃業者，未於租借予駕駛人前，教導駕駛人車輛操作方法及道路行駛規定',600,800,800,800,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2		



	End if
	rsChkL2.close
	Set rsChkL2=Nothing

'1120331=====================================================================
	strChkL2="select * from Law where itemid ='4420003' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		strInsL2="insert into law values('4420003','3','駕駛汽車行經行人穿越道有行人穿越時，不暫停讓行人先行通過',1200,1300,1400,1500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4420003','4','駕駛汽車行經行人穿越道有行人穿越時，不暫停讓行人先行通過',3600,3600,3600,3600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4430003','3','汽車駕駛人駕駛汽車行近行人穿越道遇有攜帶白手杖之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',2400,2600,2800,3100,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4430003','5','汽車駕駛人駕駛汽車行近行人穿越道遇有攜帶白手杖之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',4800,5200,6400,7200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4430003','6','汽車駕駛人駕駛汽車行近行人穿越道遇有攜帶白手杖之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',7200,7200,7200,7200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4430004','3','汽車駕駛人駕駛汽車行近行人穿越道遇有攜帶導盲犬之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',2400,2600,2800,3100,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4430004','5','汽車駕駛人駕駛汽車行近行人穿越道遇有攜帶導盲犬之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',4800,5200,6400,7200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4430004','6','汽車駕駛人駕駛汽車行近行人穿越道遇有攜帶導盲犬之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',7200,7200,7200,7200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510904','3','支線道車不讓幹線道車先行',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510904','5','支線道車不讓幹線道車先行',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510904','6','支線道車不讓幹線道車先行',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510905','3','少線道車不讓多線道車先行',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510905','5','少線道車不讓多線道車先行',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510905','6','少線道車不讓多線道車先行',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510906','3','車道數相同時，左方車不讓右方車先行',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510906','5','車道數相同時，左方車不讓右方車先行',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510906','6','車道數相同時，左方車不讓右方車先行',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511507','3','行經無號誌交岔路口不依規定',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511507','5','行經無號誌交岔路口不依規定',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511507','6','行經無號誌交岔路口不依規定',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511508','3','行經無號誌交岔路口不依標誌指示',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511508','5','行經無號誌交岔路口不依標誌指示',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511508','6','行經無號誌交岔路口不依標誌指示',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511509','3','行經無號誌交岔路口不依標線指示',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511509','5','行經無號誌交岔路口不依標線指示',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511509','6','行經無號誌交岔路口不依標線指示',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511510','3','行經巷道不依規定',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511510','5','行經巷道不依規定',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511510','6','行經巷道不依規定',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511511','3','行經巷道不依標誌指示',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511511','4','行經巷道不依標誌指示',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511511','5','行經巷道不依標誌指示',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511512','3','行經巷道不依標線指示',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511512','5','行經巷道不依標線指示',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511512','6','行經巷道不依標線指示',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4820003','3','汽車駕駛人轉彎時，除禁止行人穿越路段外，不暫停讓行人優先通行',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4820003','4','汽車駕駛人轉彎時，除禁止行人穿越路段外，不暫停讓行人優先通行',3600,3600,3600,3600,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4830003','3','汽車駕駛人轉彎時除禁止行人穿越路段外行近攜帶白手杖之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',2400,2600,2800,3100,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4830003','5','汽車駕駛人轉彎時除禁止行人穿越路段外行近攜帶白手杖之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',4800,5200,6400,7200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4830003','6','汽車駕駛人轉彎時除禁止行人穿越路段外行近攜帶白手杖之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',7200,7200,7200,7200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4830004','3','汽車駕駛人轉彎時除禁止行人穿越路段外行近攜帶導盲犬之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',2400,2600,2800,3100,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4830004','5','汽車駕駛人轉彎時除禁止行人穿越路段外行近攜帶導盲犬之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',4800,5200,6400,7200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4830004','6','汽車駕駛人轉彎時除禁止行人穿越路段外行近攜帶導盲犬之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',7200,7200,7200,7200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('6020304','3','不遵守道路交通號誌之指示(遇閃光紅燈未停車再開)',1200,1300,1400,1500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('6020304','5','不遵守道路交通號誌之指示(遇閃光紅燈未停車再開)',1500,1600,1700,1800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('6020304','6','不遵守道路交通號誌之指示(遇閃光紅燈未停車再開)',1800,1800,1800,1800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('6020305','0','不遵守道路交通號誌之指示(其他)',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

	End if
	rsChkL2.close
	Set rsChkL2=Nothing
	'1130630=====================================================================
	if now>="2024/6/30" then
	strChkL2="select * from Law where itemid ='5000205' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		
		strInsL2="insert into law values('5000205','0','倒車前未顯示倒車燈光',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5000206','0','倒車時不注意其他車輛或行人',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5000305','6','大型汽車無人在後指引時，不先測明車後有足夠之地位',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5000306','6','大型汽車無人在後指引時，不促使行人避讓',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510115','3','在橋樑臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510115','4','在橋樑臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510116','3','在隧道臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510116','4','在隧道臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510117','3','在圓環臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510117','4','在圓環臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510118','3','在障礙物對面臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510118','4','在障礙物對面臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510119','3','在快車道臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510119','4','在快車道臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510120','3','在騎樓以外之人行道臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510120','4','在騎樓以外之人行道臨時停車',600,600,600,600,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510121','3','在騎樓臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510121','4','在騎樓臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510122','3','在行人穿越道臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510122','4','在行人穿越道臨時停車',600,600,600,600,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510211','3','在消防車出入口五公尺內臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510211','4','在消防車出入口五公尺內臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510212','3','在交岔路口十公尺內臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510212','4','在交岔路口十公尺內臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510213','3','在公共汽車招呼站十公尺內臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510213','4','在公共汽車招呼站十公尺內臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510407','3','併排臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510407','4','併排臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610118','3','在公共汽車招呼站十公尺內停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610118','4','在公共汽車招呼站十公尺內停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610119','3','在橋樑停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610119','5','在橋樑停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610119','6','在橋樑停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610120','3','在隧道停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610120','5','在隧道停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610120','6','在隧道停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610121','3','在圓環停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610121','5','在圓環停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610121','6','在圓環停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610122','3','在障礙物對面停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610122','5','在障礙物對面停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610122','6','在障礙物對面停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610123','3','在行人穿越道停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610123','5','在行人穿越道停車',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610123','6','在行人穿越道停車',1200,1200,1200,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610124','3','在快車道停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610124','5','在快車道停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610124','6','在快車道停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610125','3','在交岔路口十公尺內停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610125','5','在交岔路口十公尺內停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610125','6','在交岔路口十公尺內停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610126','3','在消防車出入口五公尺內停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610126','5','在消防車出入口五公尺內停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610126','6','在消防車出入口五公尺內停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610127','3','在騎樓以外之人行道停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610127','5','在騎樓以外之人行道停車',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610127','6','在騎樓以外之人行道停車',1200,1200,1200,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610312','3','在消防栓之前停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610312','5','在消防栓之前停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610312','6','在消防栓之前停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200271','3','機車駕駛人行駛道路以手持方式使用行動電話進行撥接',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200281','3','機車駕駛人行駛道路以手持方式使用行動電話進行通話',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200291','3','機車駕駛人行駛道路以手持方式使用行動電話進行數據通訊',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200301','3','機車駕駛人行駛道路以手持方式使用行動電話進行有礙駕駛安全之行為',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200311','3','機車駕駛人行駛道路以手持方式使用電腦進行撥接',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200321','3','機車駕駛人行駛道路以手持方式使用電腦進行通話',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200331','3','機車駕駛人行駛道路以手持方式使用電腦進行數據通訊',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200341','3','機車駕駛人行駛道路以手持方式使用電腦進行有礙駕駛安全之行為',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200351','3','機車駕駛人行駛道路以手持方式使用相類功能裝置進行撥接',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200361','3','機車駕駛人行駛道路以手持方式使用相類功能裝置進行通話',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200371','3','機車駕駛人行駛道路以手持方式使用相類功能裝置進行數據通訊',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200381','3','機車駕駛人行駛道路以手持方式使用相類功能裝置進行有礙駕駛安全之行為',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2


		strInsL2="update law set illegalrule='駕駛汽車行近行人穿越道有行人穿越時，不暫停讓行人先行通過' where itemid='4420004'"
		conn.execute strInsL2

		strInsL2="update law set recordstateid=-1 where itemid in ('1610301','1610401','1610402','1610403','1610404','1610501','1610502','2110101','2110102','2110201','2110301','2110302','2110303','2110304','2110305','2110306','2110401','2110402','2110501','2130001','2130002','2130003','2130004','2130005','2130006','2130007','2130008','2150001','2150002','2150003','2150004','2150005','2150006','2150007','2150008','2150009','2150010','2150011','2150012','2150013','2150014','2150015','2210501','2230001','2230002','2230003','2230004','2230005','2230006','2230007','2230008','2230009','2230010','2230011','2230012','2230013','2230014','2230015','2300201','2440001','3310101','3310102','3310103','3310104','3310105','3310106','3310107','3310108','3310109','3310110','3310111','3310112','3310113','3310114','3310115','3310116','3310117','3310118','3310119','3310120','3310121','3310122','3310123','3310124','3310125','3310126','3310127','3310128','3310129','3310130','3310131','3310132','3310201','3310202','3310203','3310204','3310401','3310402','3310403','3310404','3310607','3310701','3310702','3310703','3310704','3310705','3310706','3310707','3310708','3310709','3310710','3310711','3310712','3310713','3310714','3310715','3310901','3310902','3311105','3311106','3311601','3311602','3311603','3311604','3311605','3311606','3311701','3311702','3400001','3400002','3400003','3400004','3400005','3400006','3400013','3400014','3400015','3400016','3400017','3400018','3400019','3810001','3810002','3810003','4200001','4200002','4200003','4310105','4310106','4310107','4310108','4310113','4310114','4310210','4310211','4310212','4310213','4310214','4310215','4310216','4310217','4310218','4310219','4310220','4310221','4310222','4310223','4310224','4310225','4310226','4310227','4310228','4310229','4310230','4310231','4310232','4310233','4310234','4310235','4310236','4310237','4310238','4310239','4310307','4310308','4310309','4310310','4310311','4310312','4310313','4310314','4310315','4310316','4310317','4310318','4310319','4310320','4310321','4310322','4310323','4310324','4310401','4310402','4310403','4310404','4310405','4310406','4310407','4310408','4310409','4310410','4310411','4310412','4310413','4310414','4310415','4310416','4310417','4310418','4330003','4330008','4330013','4330018','4330029','4330030','4330031','4330032','4340011','4340014','4340035','4340044','4340045','4340056','4340057','4420003','4430003','4430004','4700101','4700102','4700103','4700104','4700105','4700106','4700107','4700108','4700109','4700110','4700111','4700112','4700201','4700202','4700203','4700204','4700205','4700206','4700207','4700208','4700209','4700210','4700211','4700212','4700301','4700302','4700303','4700304','4700305','4700306','4700401','4700402','4700403','4700404','4700501','4700502','4700503','4700504','4700505','4700506','4700507','4700508','4820003','4830003','4830004','5000201','5000202','5000301','5000302','5320001','5510101','5510102','5510103','5510104','5510105','5510106','5510107','5510201','5510202','5510203','5510401','5510404','5610102','5610103','5610310','5620002','6020304','6020305','6110401','6110403','6130001','6130002','6330001','6330002','6330003','6330004','6720011','6820004','7430003','7430005','7430007','7430009','7430010','7430011','7430012','7430013','7430014','8230101','8230102','8230103','8230104','8310101','8310201','8310301','8410101','8540001','8540002','8540003','8540004','21101021','21102021','21103021','21104021','21104041','21105021','21105041','21106021','21106041','21106061','21107081','30101001','30101002','30101003','30101004','30107001','30107002','30107003','30107004','31100011','31100021','31100031','31100041','31100051','31100061','31100071','31100081','31100091','31100101','31100111','31100121','31100131','31100141','31200011','31200021','31200031','31200041','31200051','31200061','31200071','31200081','31200091','31200101','31200111','31200121','31200131','31200141','56000011','56000021','56000031','56000041','56000051','56000061','56000071','56000081','56000091','56000101','56000111','56000121','56000131','56000141','56000151','56000161','56000171','56000181','56000191','56000201','56000211','56000221','56000231','56000241','63000011','63100011','351000031','352000021','353000011','4410201','4000007','4000010','4000013','4000016','4810103','4810104','4810105','4810113','4810114','4810115','4810201','4810301','4810401','4810402','4810501','4810502','4810601','4810602','4810701','4810702','4810703','5610116','5510204','5510205','5610110')"
		conn.execute strInsL2
	End if
	rsChkL2.close
	Set rsChkL2=Nothing
	end if
'==================================================================


End If 
%>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</head>
<body leftmargin="25" topmargin="5" marginwidth="0" marginheight="0" <%
	if sys_City<>"台中市" and sys_City<>"台中縣" then
%>
		onLoad="init()"
<%
	end if
%>>
<div id="D1" style="width:350px">
<%
	if sys_City<>"台中市" and sys_City<>"台中縣" then
%>
  <table border="0" width="350">
    <TBODY> 
	<tr>
		<!-- 處理進度 -->
		<td valign="top"> 
			<table width="100%" border="1">
				<tr bgcolor="#FFFF99">
					<td >處理進度(1~68條)</td>
				</tr>
				<tr bgColor="#FFFFFF">
					<td id="YestodayLayer">

					</td>
				</tr>
				<tr bgColor="#FFFFFF">
					<td id="TodayLayer">
					
					</td>
				</tr>
			</table>
		</td>
		<td align="middle" bgColor="#FF0000" rowSpan="2" width="20" >
			<font color="#FFFFFF"> 
			個<br>人<br>資<br>料
			</font>
		</td>
    </tr>
    <tr>
		<!-- 上傳紀錄 -->
		<td valign="top" id="UpLoadLayer"> 

		</td>
    </tr>

    </TBODY> 
	<tr>
		<td height="5"></td>
	</tr>
	<tr>
	<td><iframe frameborder="0" width="320px" height="200px" src="chat/onlinelist.asp"></iframe></td>
		<td align="middle" bgColor="#FF0000" rowSpan="2" width="20" >
			<font color="#FFFFFF"> 
			線<br>上<br>人<br>員
			</font>
		</td>
  </table>
<%
	end if
%>
			
</div>

<font color="#000000"><b>
<%

strUnit="select * from UnitInfo where UnitID='"&UnitNo&"'"
set rsUnit=conn.execute(strUnit)
If Not rsUnit.eof then
UnitName=rsUnit("UnitName")
End if
rsUnit.close
set rsUnit=nothing

%>
</b>
</font>
<table width='1000' border='0' align="center" class="TitleSet">
	<tr>
		<td height="95">
		<form name="myForm" method="post">
		<input type="hidden" name="SystemType" value="">
		</form>
		</td>
	</tr>
	<tr>
		<td colspan="5" height="40">


<br><br>

	<div id='Banner-Menu' class='floatL'>
      <ul>
        <li <%
	If Trim(request("SystemType"))="0" Then	
		response.write "class='current3'"
	Else
		response.write "class='current1'"
	End If 
		%>onclick="SelFunc('0');"></li>
        <li <%
	If Trim(request("SystemType"))="1" Or Trim(request("SystemType"))="" Then	
		response.write "class='current4'"
	Else
		response.write "class='current2'"
	End If 
		%> onclick="SelFunc('1');"></li>
      </ul>

	  <div>
	  <table width="550" border="0">
		<tr>
			<td rowspan="2" width="45%" valign="middle">&nbsp;<font color="#FF0000"><strong>
				<%=UnitName%>
				<img src="image/space.gif" alt="" width="5" height="2" border="0" align="baseline">
				<%=memName%><%
			strMpID="select MpoliceID from memberdata where memberid="& Trim(session("User_ID"))
			Set rsMpID=conn.execute(strMpID)
			If Not rsMpID.eof Then
				response.write "&nbsp; "&Trim(rsMpID("MpoliceID"))
			End If 
			rsMpID.close
			Set rsMpID=Nothing 
				%></strong></font>
				<br>&nbsp; 登入IP&nbsp;<%
				ServerIp999 = Request.ServerVariables("Local_ADDR") 
				response.write ServerIp999
				%>
				<!-- <br>&nbsp;上次登入&nbsp; -->
				<%
'				strMD="select max(Actiondate) as MaxDate FROM Log where typeid=350 and Actiondate > TO_DATE('2021/6/1 0:0:0','YYYY/MM/DD/HH24/MI/SS') and actionmemberid="&Trim(session("User_ID"))
'				Set rsMD=conn.execute(strMD)
'				If Not rsMD.eof Then
'					If Not IsNull(rsMD("MaxDate")) Then 
'						response.write rsMD("MaxDate")
'					End If 
'				End If
'				rsMD.close
'				Set rsMD=Nothing 
				%>
			</td>
			<td width="55%" >
				<img src="Image/dot.gif" >
					<font class="style2">客服 : (02) 2790-0989</font>
				<img src="Image/space.gif" width="5" height="1" >
				<img src="Image/dot.gif" ondblclick="location='UserLogout_Contral.asp'">
					<font class="style2">傳真 : (02) 2790-3616</font>	
			</td>
		</tr>
		<tr>
			<td><img src="Image/dot.gif" >
			<font class="style2">信箱<b>  178hyndai@gmail.com </b></font>
			</td>
		</tr>
		
	  </table>
	  </div>
	</div>

	  <td>
	</tr>
	
<%'-------------------------------公告區---------------------------------------------
If Trim(request("SystemType"))="0" then%>
	<tr>
		<td>
		<iframe frameborder="0" width="970px" height="200px" src="ArgueCase/CaseNotifyPublic.asp"></iframe>
		</td>
	</tr>
<%if sys_City="台南市" And Trim(Session("Group_ID"))="200" then%>




	<tr>
		<td >
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr>
			    <td bgcolor="#FFCC33">
				<font size="4"><strong>系統警示</strong>
				</font>	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			    </td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td >
		<table width="100%" align="left" border="0">
			<tr>
				<td width="40%">
				十天內建檔後超過四天內未入案之案件共 <strong><%
			LimitDate=DateAdd("d",-4,date)
			TenDate=DateAdd("d",-10,date)
			strA="select count(*) as cnt from billbase where billstatus in ('0','1') and recordstateid=0 and recorddate between to_date('"&TenDate&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and to_date('"&LimitDate&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and SN not in (select BillSN from ALERTCHECK where TypeID='1')"
			Set rsA=conn.execute(strA)
			If Not rsA.eof Then
				response.write rsA("cnt")
			End If
			rsA.close
			Set rsA=Nothing 
				%></strong> 筆
				</td>
				<td width="60%">
				<input type="button" value="檢視" onclick='window.open("Check_NotCaseIn.asp","Check_NotCaseIn","left=100,top=50,location=0,width=860,height=580,resizable=yes,scrollbars=yes,status=yes")'>
				</td>
			</tr>
			<tr>
				<td>
				十天內建檔後僅做車籍查詢就刪除之案件共 <strong><%
			strA="select count(*) as cnt from billbase where billno is null and sn in (select billsn from dcilog where exchangetypeid='A') and recordstateid=-1 and Recorddate > to_date('"&TenDate&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and SN not in (select BillSN from ALERTCHECK where TypeID='2')"
			Set rsA=conn.execute(strA)
			If Not rsA.eof Then
				response.write rsA("cnt")
			End If
			rsA.close
			Set rsA=Nothing 
				%></strong> 筆 
				</td>
				<td>
				<input type="button" value="檢視" onclick='window.open("Check_QryCar.asp","Check_QryCar","left=100,top=50,location=0,width=860,height=580,resizable=yes,scrollbars=yes,status=yes")'>
				</td>
			</tr>
			<tr>
				<td>
				十天內密碼輸入錯誤三次後系統鎖定之帳號共 <strong><%
			strA="select count(*) as cnt from Memberdata where AccountStateID=-1 and DelMemberID=99999 and LeaveJOBDate > to_date('"&TenDate&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
			Set rsA=conn.execute(strA)
			If Not rsA.eof Then
				response.write rsA("cnt")
			End If
			rsA.close
			Set rsA=Nothing 
				%></strong> 筆 
				</td>
				<td>
				<input type="button" value="檢視" onclick='window.open("Check_UserLock.asp","Check_UserLock","left=100,top=50,location=0,width=860,height=580,resizable=yes,scrollbars=yes,status=yes")'>
				</td>
			</tr>
		</table>
		</td>
	</tr>
<%End if%>		
	<tr>
		
		<td >
	
		<!--
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr>
			    <td bgcolor="#FFCC33">
				<font size="4"><strong>公告訊息</strong>
				</font>	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			    </td>
			</tr>
		</table>
		-->
<%
set fs=Server.CreateObject("Scripting.FileSystemObject")

tdate ="select to_char(sysdate,'yymmdd') as tdate from dual"
set rstdate =conn.execute(tdate )
tdate =trim(rstdate ("tdate"))
rstdate.close
'response.write "note"& tdate &".txt"
FileName=Server.MapPath(fs.GetFileName("note"& tdate &".txt"))

	    if fs.fileExists(FileName)=true then
           set txtStream = fs.opentextfile(FileName) 
              txtline = txtStream.readAll
              response.write "<font color=red size=""4"">"&txtline&"</font>"
     	end if

 set txtStream = nothing
 set fs = nothing 
%>
<%	

	strDelErr="select * from Dcilog where ExchangeTypeID='E' and (DciReturnStatusID<>'S' or DciReturnStatusID is null)" &_
		" and ExchangeDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS')" &_
		" and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and RecordMemberID="&Session("User_ID")
	set rsDelErr=conn.execute(strDelErr)
	if not rsDelErr.eof then
%>
	<table border="1" width="100%" id="table3">
		<tr bgcolor="#FFCC33">
			<td colspan="4">十日內刪除未處理、異常</td>
		</tr>
		<tr bgcolor="#FFCC33">
			<td width="20%">上傳日期</td>
			<td width="20%">批號</td>
			<td width="20%">單號</td>
			<td width="40%">訊息</td>
		</tr>
<%
	end if
	If Not rsDelErr.Bof Then rsDelErr.MoveFirst 
	While Not rsDelErr.Eof
%>
		<tr>
			<td>
			<%=year(rsDelErr("ExchangeDate"))-1911&right("00"&month(rsDelErr("ExchangeDate")),2)&right("00"&day(rsDelErr("ExchangeDate")),2)%>
			</td>
			<td>
			<%=rsDelErr("Batchnumber")%>
			</td>
			<td>
			<%=rsDelErr("BillNo")%>
			</td>
			<td>
			<%
			if trim(rsDelErr("DciReturnStatusID"))="" or isnull(rsDelErr("DciReturnStatusID")) then
				response.write "未處理"
			else
				strErr="select StatusContent from DciReturnStatus where DciActionID='E'" &_
					" and DciReturn='"&trim(rsDelErr("DciReturnStatusID"))&"' " 
				set rsErr=conn.execute(strErr)
				if not rsErr.eof then
					response.write rsErr("StatusContent")
				end if
				rsErr.close
				set rsErr=nothing
			end if
			%>
			</td>
		</tr>
<%
	rsDelErr.MoveNext
	Wend
	if not rsDelErr.eof then
%>
	</table>
<%
	end if
	rsDelErr.close
	set rsDelErr=nothing
%>
<%
	strNotUpload="select * from billbase where recorddate<(sysdate-10) and billstatus in ('0','1') " &_
		" and RecordMemberID="& Session("User_ID") & " and Recordstateid=0 order by recorddate desc"
	set rsNotUpload=conn.execute(strNotUpload)
	if not rsNotUpload.eof then
%>
	<table border="1" width="800" id="table3">
		<tr bgcolor="#FF3300">
			<td colspan="4">超過十日未上傳舉發單</td>
		</tr>
		<tr bgcolor="#FFFF99">
			<td>類別</td>
			<td>單號</td>
			<td>車號</td>
		</tr>
<%
	
	If Not rsNotUpload.Bof Then rsNotUpload.MoveFirst  
	While Not rsNotUpload.Eof
%>
		<tr>
			<td>
			<%
			if trim(rsNotUpload("BillTypeid"))="1" then
				response.write "攔停"
			else
				response.write "逕舉"
			end if
			%>
			</td>
			<td>
			<%
			if trim(rsNotUpload("BillNo"))<>"" then
				response.write trim(rsNotUpload("BillNo"))
			else
				response.write "&nbsp;"
			end if
			%>
			</td>
			<td>
			<%
			if trim(rsNotUpload("CarNo"))<>"" then
				response.write trim(rsNotUpload("CarNo"))
			else
				response.write "&nbsp;"
			end if
			%>
			</td>
		</tr>
<%
	rsNotUpload.MoveNext
	Wend
%>

	</table>

<%
	end if
	rsNotUpload.close
	set rsNotUpload=Nothing
	%>
		</td>
	</tr>
		

	
<%'---------------------------作業區--------------------------------------
elseIf Trim(request("SystemType"))="1" Or Trim(request("SystemType"))="" then%>
	<tr>
		<td colspan="5" class="style3">
			<font style="font-size:12pt;line-height:18px;font-weight:bold;"><div style="text-align:right;width:100%">電話客服時間 週一~週五(國定假日除外) 上午09:00~12:00 下午01:00~05:00</div></font>
			<br>
			<font color="red"><strong>快速查詢</strong></font>
			<input type="hidden" name="creditidhidden" value="@@@<%=Session("Credit_ID")%>^^^" >
			<img src="image/space.gif" alt="" width="12" height="2" border="0" align="baseline">
			&nbsp;<strong>舉發單號</strong>&nbsp;<input type="text" name="BillNo" size="10" maxlength="9" onkeyup="EnterBillQry();">
			&nbsp;<strong>車號</strong>&nbsp;<input type="text" name="CarNo" size="8" maxlength="9" onkeyup="EnterBillQry();">
			<!-- 花蓮拖吊場限制查詢條件 -->
			<% if Session("Credit_ID")="0000" AND sys_City="花蓮縣" then %>
				&nbsp;<strong></strong>&nbsp;<input type="hidden" name="IllegalName" size="12" maxlength="30">
				&nbsp;<strong></strong>&nbsp;<input type="hidden" name="IllegalID" size="10" maxlength="10">	

                       <% else %>
				&nbsp;<strong>姓名</strong>&nbsp;<input type="text" name="IllegalName" size="12" maxlength="30"  onkeyup="EnterBillQry();">
				&nbsp;<strong>身份證號</strong>&nbsp;<input type="text" name="IllegalID" size="10" maxlength="10" onkeyup="EnterBillQry();">	


			 <% end if %>
			 <strong>查詢事由</strong>
			<select name="QryReason" >
				<option value="" >請選擇</option>
				<option value="資料檢核" <%If Trim(request("QryReason"))="資料檢核" Then response.write "selected" End if%>>資料檢核</option>
				<option value="執行業務" <%If Trim(request("QryReason"))="執行業務" Then response.write "selected" End if%>>執行業務</option>
				<option value="民眾申訴" <%If Trim(request("QryReason"))="民眾申訴" Then response.write "selected" End if%>>民眾申訴</option>
				<option value="事故處理" <%If Trim(request("QryReason"))="事故處理" Then response.write "selected" End if%>>事故處理</option>
				<option value="偵查刑案" <%If Trim(request("QryReason"))="偵查刑案" Then response.write "selected" End if%>>偵查刑案</option>
			</select>
			<input type="button" value="查詢" onclick='getBillData()'>
			<span class='style2'><b>可跨分局查詢</b></span>
			<br>
			<font color="red"><strong>&nbsp; &nbsp; &nbsp;未入案案件及入案異常案件，無法使用快速查詢，請至『舉發單資料維護系統』查詢案件
			<br>
			&nbsp; &nbsp; &nbsp;如要查詢案件入案異常原因，請至『上傳下載資料查詢系統』查詢
			</strong></font>
			<%If ((Trim(Session("Credit_ID"))="N121946889" Or Trim(Session("Credit_ID"))="P120936942" Or Trim(Session("Credit_ID"))="E121544459" Or Trim(Session("Credit_ID"))="M121135787" Or Trim(Session("Credit_ID"))="Z016") And sys_City="台中市") Or Session("Credit_ID")="A000000000" Or (sys_City="宜蘭縣" And Trim(Session("Credit_ID"))="G121048936") then%>
			<a href="BillKeyIn/Update_Report_IllegalTime.asp" target="_blank" class="style2">** 強制修改違規時間 ** </a>	
			<%End If %>
			<br>			
			<a href="入案系統使用說明V8.pdf" target="_blank" style="font-size: 24px;">交通執法系統使用手冊</a>
			<%
			if sys_City="台中市" Then%>
			<br>			
			<a href="BillUnitMem.asp" target="_blank" class="style2">** 各單位舉發單職名章主管人員 ** </a>
			&nbsp;&nbsp;
			<a href="MailNotBack.doc" target="_blank" class="style2">** 郵寄未退回清冊使用手冊 ** </a>	
			&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
			<strong>告示單號</strong>&nbsp;<input type="text" name="ReportNo" size="15" onkeyup="EnterBillQry();" >
			<%End if%>
			<%If sys_City="南投縣" And (Trim(session("Credit_ID"))="M120783510" Or Trim(session("Credit_ID"))="A000000000" Or Trim(session("Credit_ID"))="M220555223") then%>
				<input type="button" value="建檔同車號案件檢查" onclick="window.open('BillKeyIn/setCheckCarRule.asp','winother1','width=500,height=350,left=100,top=50,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=yes,toolbar=no')" style="font-size: 9pt; width: 160px; height: 27px">
			<%End if%>
			<%if (Session("Credit_ID")<>"0000" AND sys_City="花蓮縣") Or sys_City="台東縣" then%>
			<br>
			催繳單號
			<input type="text" class="btn1" size="16" maxlength="16" value="" name="StopBillNo" onkeyup="EnterBillQry_Stop();">
			催繳車號
			<input name="StopCarNo" type="text" value="" size="8" maxlength="7" class="btn1" onkeyup="EnterBillQry_Stop();">
			<input type="button" name="btStopBill" value="催繳單查詢" onclick="Selt_Stop();">
			<%end if%>
			
			<%if Session("Credit_ID")="A000000000" then%>

							<font color="red">v$flash_recovery_area_usage</font>
							<%
								strDB="select PERCENT_SPACE_USED from v$flash_recovery_area_usage  where File_Type='ARCHIVELOG'"
								set rsDB=conn.execute(strDB)
								if not rsDB.eof Then
									if cint(rsDB("PERCENT_SPACE_USED"))>"80" Then
										response.write "<font color='red' style='line-height:48px;font-size:40pt;'>"&cint(rsDB("PERCENT_SPACE_USED"))&"%</font>"
									Else
										response.write "<font color='red' style='line-height:28px;font-size:20pt;'>"&cint(rsDB("PERCENT_SPACE_USED"))&"%</font>"
									End if
									
									if cint(rsDB("PERCENT_SPACE_USED"))>80 and Session("Credit_ID")="A000000000" then
							%>
									<script language="JavaScript">
										alert("Oracle 暫存區快滿了，請趕快清一清");
									</script>
							<%
									end if
								end if
								rsDB.close
								set rsDB=Nothing
								'response.write "<font color='red' style='font-size:40pt;'>ooo</font>"
			  			%>	
					
				&nbsp; &nbsp; <font color="red"> v$session</font>
							<%
								strDB="Select count(*) as cnt from v$session"
								set rsDB=conn.execute(strDB)
								if not rsDB.eof Then
										response.write "<font color='red' style='line-height:28px;font-size:20pt;'>"&cint(rsDB("cnt"))&"</font>"

								end if
								rsDB.close
								set rsDB=Nothing
								'response.write "<font color='red' style='font-size:40pt;'>ooo</font>"
			  			%>		
							<input type="button" value="資料庫Session" onclick="window.open('trafficDBCheck.asp','winother','width=900,height=550,left=0,top=0,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=yes,toolbar=no')" style="font-size: 9pt; width: 100px; height: 27px">
			<%end if%>
			<br>
			<input type="button" name="b10010" value="修改個人資料" onclick="funMember();" style="font-size: 9pt; width: 90px; height:23px;">
		</td>
	</tr>

<%
strFuncGroup="select * from Code where TypeID=19 order by ShowOrder"
set rsFuncGroup=conn.execute(strFuncGroup)
If Not rsFuncGroup.Bof Then rsFuncGroup.MoveFirst 
While Not rsFuncGroup.Eof
%>
	<tr><td width="20%"></td><td width="20%"></td><td width="20%"></td><td width="20%"></td><td width="20%"></td><tr>
	<tr>
		<td colspan="5">
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr>
			    <td bgcolor="#FFCC33">
				<font size="4"><strong><%=trim(rsFuncGroup("Content"))%></strong>
				<%
				If trim(rsFuncGroup("ID"))="514" Then
					response.write "&nbsp;<a href='UserDataEdit.asp' target='_blank' ><font size='4' color='red'>修改個人密碼</font></a>"
response.write "&nbsp;<a href='拖吊系統_系統功能操作手冊0601.doc' target='_blank' ><font size='4' color='red'>拖吊系統_系統功能操作手冊</font></a>"
				End If 
				%>
				</font>	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<% if sys_city="高雄市" then
						ShowTimeGroupID=516
					else
						ShowTimeGroupID=512
					end if
					if Trim(rsFuncGroup("ID"))=Trim(ShowTimeGroupID) then
				%>
					<font size="3"><strong>系統時間 <div id="LayerTime" ondblclick="NowDownProcess();"></div></strong></font>		
				<%  end if
				
					if rsFuncGroup("ID")="509" then
				%>
					<br>
					<font size="5">
					<strong>
					<img src="Image/dot.gif"></img>
					<a href="send.doc" target="_blank" >民眾簽收 簡要說明.doc</a>
					<img src="Image/dot.gif"></img>
					<a href="opengov.doc" target="_blank" >公示送達 簡要說明.doc</a>
					<img src="Image/dot.gif"></img><!--
					<a href="storeandsend.doc" target="_blank" >寄存送達 簡要說明.doc</a>-->
					</strong>
					
					</font>		
				<%  end if				
							
				%>
			    </td>
			</tr>
		</table>
		</td>
	</tr>
<%
	s=1
	strFunc="select a.* from FunctionPageData a,FunctionData b where a.SystemID=b.SystemID and b.GroupID='"&trim(GroupID)&"' and a.SystemGroupID="&trim(rsFuncGroup("ID"))&" and b.Function='1' order by ShowOrder"
	set rsFunc=conn.execute(strFunc)
	If Not rsFunc.Bof Then rsFunc.MoveFirst 
	While Not rsFunc.Eof

	if s=1 then
		response.write "<tr>"
	end if
%>	<td width="190" height="180">
<div id="<%=rsFunc("SystemID")%>" style="width:155px; height:170px; z-index:1 ;">  
		<table id="<%="table"&rsFunc("SystemID")%>" width='100%' border='0' align="center" >
		<tr><td id="td1" align="center">
		<a onclick="OpenSystem('<%=rsFunc("URLLocation")%>','<%=rsFunc("SystemID")%>');" onMouseOver="DivColorChange('<%="table"&rsFunc("SystemID")%>');" onMouseOut="DivColorChange2('<%="table"&rsFunc("SystemID")%>');">
		<%
		if trim(rsFunc("ImageLocation"))="" or isnull(rsFunc("ImageLocation")) then
			picName="tmp.jpg"
		else
			picName=rsFunc("ImageLocation")
		end if
		%>
		<img src="image/<%=picName%>" alt="" width="128" height="128" border="0" align="baseline">
		  <br><font size="4"><%
		  strCode="select * from Code where ID="&trim(rsFunc("SystemID"))
		  set rsCode=conn.execute(strCode)
		  if not rsCode.eof then
			response.write trim(rsCode("Content"))
		  end if
		  rsCode.close
		  set rsCode=nothing
		  %></font></a>
		</td></tr>
		</table>
</div>
	</td>
<%	
	if s=5 then
		response.write "</tr>"
		s=1
	else
		s=s+1
	end if

	rsFunc.MoveNext
	Wend

rsFuncGroup.MoveNext
Wend
rsFuncGroup.close
set rsFuncGroup=nothing
%>
  <%
	rsFunc.close
	set rsFunc=nothing
	strSQL="select TO_Char(sysdate,'YYYY/MM/DD HH24:MI:SS') tmpDate from dual"
	set rstime=conn.execute(strSQL)
	SysDate=rstime("tmpDate")
	rstime.close

%>
<%End If %>

</table>
</body>
<script type="text/javascript" src="./js/date.js"></script>
<script language="JavaScript">
<%
'if sys_City="金門縣" Then

	Modifydate=""
	PassWordTemp=""
	showUpdateMemberdateFlag=0
	strMem="select * from memberdata where memberid=" & Trim(session("User_ID")) & " and recordstateid=0 and accountstateid=0"
	Set rsMem=conn.execute(strMem)
	If Not rsMem.eof Then
		If Trim(rsMem("ModifyTime"))="" Then
			Modifydate=Trim(rsMem("ModifyTime"))
		Else
			Modifydate=Trim(rsMem("RecordDate"))
		End If 
		PassWordTemp=Trim(rsMem("PassWord"))
	End If 
	rsMem.close
	Set rsMem=Nothing 
	If DateDiff("d",Modifydate,now)>90 Then
		response.write "alert('您的密碼已超過三個月( " & DateDiff("d",Modifydate,now) & " 天)未更換，請儘速修改您的密碼!!');"
		showUpdateMemberdateFlag=1
	elseif Len(PassWordTemp)<8 then
		response.write "alert('密碼長度至少為<8>碼，請儘速修改您的密碼!!');"
		showUpdateMemberdateFlag=1
	else
		chkUp=0
		chkDown=0
		chkInt=0
		chkMark=0
		for i=1 to Len(PassWordTemp)
			if Asc(Mid(Trim(PassWordTemp), i, 1))>=65 and Asc(Mid(Trim(PassWordTemp), i, 1))<=90 then
				chkUp=1
			end if
			if Asc(Mid(Trim(PassWordTemp), i, 1))>=97 and Asc(Mid(Trim(PassWordTemp), i, 1))<=122 then
				chkDown=1
			end if 
			if Asc(Mid(Trim(PassWordTemp), i, 1))>=48 and Asc(Mid(Trim(PassWordTemp), i, 1))<=57 then
				chkInt=1
			end if 
			if (Asc(Mid(Trim(PassWordTemp), i, 1))>=33 and Asc(Mid(Trim(PassWordTemp), i, 1))<=47) or (Asc(Mid(Trim(PassWordTemp), i, 1))>=58 and Asc(Mid(Trim(PassWordTemp), i, 1))<=64) or (Asc(Mid(Trim(PassWordTemp), i, 1))>=91 and Asc(Mid(Trim(PassWordTemp), i, 1))<=96) or (Asc(Mid(Trim(PassWordTemp), i, 1))>=123 and Asc(Mid(Trim(PassWordTemp), i, 1))<=126) then
				chkMark=1
			end if
		next
		if chkUp=0 or chkDown=0 or chkInt=0 or chkMark=0 then
			response.write "alert('密碼需包含英文、數字、特殊符號及大小寫混和，請儘速修改您的密碼!!');"
			showUpdateMemberdateFlag=1
		end if
	End If 
'End If 

'conn.close
'set conn=nothing
%>

function EnterBillQry(){
	document.all.BillNo.value=document.all.BillNo.value.toUpperCase();
	document.all.CarNo.value=document.all.CarNo.value.toUpperCase();
	document.all.IllegalID.value=document.all.IllegalID.value.toUpperCase();
<%if sys_City="台中市" then%>
	document.all.ReportNo.value=document.all.ReportNo.value.toUpperCase();
<%end if%>
	if (event.keyCode==13){
		getBillData();
	}
}

function getBillData(){
	
	if (document.all.BillNo.value.length < 9 && document.all.BillNo.value!=""){
		alert("舉發單號小於九碼！");
<%if sys_City="台中市" then%>
	}else if (document.all.CarNo.value=="" && document.all.BillNo.value=="" && document.all.IllegalName.value=="" && document.all.IllegalID.value=="" && document.all.ReportNo.value==""){
		alert("必須填入單號或車號或違規人或告示單號！");
<%else%>
	}else if (document.all.CarNo.value=="" && document.all.BillNo.value=="" && document.all.IllegalName.value=="" && document.all.IllegalID.value=="" ){
		alert("必須填入單號或車號或違規人！");
<%end if %>	
	}else if (document.all.QryReason.value==""){
		alert("因資安審查規定，查詢必須選擇查詢事由！");
		
	}else{
<%if sys_City="台中市" then%>
		UrlStr="../traffic/Query/BillBaseData_Detail_Main.asp?BillNo="+document.all.BillNo.value+"&CarNo="+document.all.CarNo.value+"&IllegalName="+document.all.IllegalName.value+"&IllegalID="+document.all.IllegalID.value+"&ReportNo="+document.all.ReportNo.value;
		newWin(UrlStr,"winMap",800,550,50,10,"yes","yes","yes","no");
<%elseif sys_City="台南市QQ" then %>
		UrlStr="../traffic/Query/BillBaseData_Detail_Main_TN.asp?BillNo="+document.all.BillNo.value+"&CarNo="+document.all.CarNo.value+"&IllegalName="+document.all.IllegalName.value+"&IllegalID="+document.all.IllegalID.value;
		newWin(UrlStr,"winMap",950,650,50,10,"yes","yes","yes","no");
<%else%>
		UrlStr="/traffic/Query/BillBaseData_Detail_Main.asp?BillNo="+document.all.BillNo.value+"&CarNo="+document.all.CarNo.value+"&IllegalName="+document.all.IllegalName.value+"&IllegalID="+document.all.IllegalID.value+"&QryReason="+document.all.QryReason.value;
		newWin(UrlStr,"winMap",800,550,50,10,"yes","yes","yes","no");
<%end if %>
	}
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
}

function OpenSystem(PageUrl,Sn){
	SCheight=screen.availHeight;
	SCWidth=screen.availWidth;
	UrlStr=PageUrl;
	newWin(UrlStr,'',SCWidth,SCheight,0,0,"yes","no","yes","no");
}
function DivColorChange(DivNo){
	eval(DivNo).border="1";
}
function DivColorChange2(DivNo){
	eval(DivNo).border="0";
}
function funMember(){
	UrlStr="UserDataEdit.asp";
	newWin(UrlStr,"winMap",800,450,50,10,"yes","no","yes","no");
}
function change_Time(){
//	var time=new Date();
//	t_Hour=time.getHours();
//	t_Minute=time.getMinutes();
//	t_Second=time.getSeconds();
	//runServerScript("getServerTime.asp");

	//LayerTime.innerHTML="目前時間  "+t_Hour+"："+t_Minute;
	//setTimeout(change_Time,60000);
}
function funOpenWindow(){
	//跳出視窗的網址
	UrlStr="NOTICEMain.asp";
	newWin(UrlStr,"winMap111",600,450,50,10,"yes","no","yes","no");
<%if sys_City="台中市" or (sys_City="彰化縣" and trim(Session("Group_ID"))=199) then%>
	newWin("WeekDeleteCase.asp","winMap181",760,600,50,10,"yes","no","yes","no");
<%end if%>

}
<%
	if showUpdateMemberdateFlag=1 then
		response.write "funMember();"
	end if 
%>
//登入就跳視窗
//funOpenWindow();

function menuIn() //隱藏
{
        if(n4) {
                clearTimeout(out_ID)
                if( menu.left > menuH*-1+30+10 ) {  
                        menu.left -= 14
                        in_ID = setTimeout("menuIn()", 1)
                }
                else if( menu.left > menuH*-1+30 ) {
                        menu.left--
                        in_ID = setTimeout("menuIn()", 1)
                }
        }
        else { 
                clearTimeout(out_ID)
                if( menu.pixelLeft > menuH*-1+30+10 ) {
                        menu.pixelLeft -= 14
                        in_ID = setTimeout("menuIn()", 1) 
                }
                else if( menu.pixelLeft > menuH*-1+30 ) {
                        menu.pixelLeft--
                        in_ID = setTimeout("menuIn()", 1)
                }
        }
}
function menuOut() //顯示
{
        if(n4) {
                clearTimeout(in_ID)
                if( menu.left < -10) { 
                        menu.left += 4
                        out_ID = setTimeout("menuOut()", 1)
                }
                else if( menu.left < 0) { 
                        menu.left++
                        out_ID = setTimeout("menuOut()", 1)
                }
                
        }
        else { 
                clearTimeout(in_ID)
                if( menu.pixelLeft < -10) {
                        menu.pixelLeft += 2
                        out_ID = setTimeout("menuOut()", 1)
                }
                else if( menu.pixelLeft < 0 ) {
                        menu.pixelLeft++
                        out_ID = setTimeout("menuOut()", 1)
                }
        }
}
function fireOver() { 
        clearTimeout(F_out)	       
        F_over = setTimeout("menuOut()", 1) 
}
function fireOut() { 
        clearTimeout(F_over)
         F_out = setTimeout("menuIn()", 1)
}
function init() {
        if(n4) {
                menu = document.D1
                menuH = menu.document.width
                menu.left = menu.document.width*-1+30 
                menu.onmouseover = menuOut
                menu.onmouseout = menuIn
				menu.visibility = "visible"
        }
        else if(e4) {
                menu = D1.style
                menuH = D1.offsetWidth
                //D1.style.pixelLeft = D1.offsetWidth*-1+20
                D1.onmouseover = fireOver
                D1.onmouseout = fireOut
				D1.onclick = fireOut
                D1.style.visibility = "visible"
        }
		UpdateLayer();
}
function UpdateLayer(){
	//UpLoadLayer.innerHTML="";
	runServerScript("UpdateMainLayer.asp");
	setTimeout(UpdateLayer,1200000);
	//alert("1");
}
F_over=F_out=in_ID=out_ID=null
n4 = 0;
e4 = 1;
var procesID='<%=Session("Credit_ID")%>'

function DownProcess(nowtime){
	<%if sys_City="台南市" or sys_City="花蓮縣" or sys_City="台東縣" or sys_City="雲林縣" or sys_City="台中市" or sys_City="高雄市" then%>
		runServerScript("/traffic/BillReturn/SystemDownloadFile.asp?nowTime="+nowtime);
	<%end if%>
}

function NowDownProcess(){
	<%if sys_City="台南市" or sys_City="花蓮縣" or sys_City="台東縣" or sys_City="雲林縣" or sys_City="台中市" or sys_City="高雄市" then%>
		var nowtime="<%=year(SysDate)&"/"&Month(SysDate)&"/"&Day(SysDate)&" "&hour(SysDate)&"："&minute(SysDate)&"："&second(SysDate)%>";
		runServerScript("/traffic/BillReturn/SystemDownloadFile.asp?nowTime="+nowtime);
		alert("已重新處理");
	<%end if%>
}

function SQLDownProcess(nowtime){
	<%if sys_City="台中縣" or sys_City="雲林縣" or sys_City="屏東縣"then%>
//		if(procesID=='A000000000'){
			//alert("/traffic/BillReturn/T-SQL.asp?nowTime="+nowtime);
			runServerScript("/traffic/BillReturn/T-SQL.asp?nowTime="+nowtime);
//		}
	<%end if%>
}

function ExportDB(){
	//runServerScript("OracleReturn.asp");
	newWin("/traffic/BillReturn/OracleReturn.aspx","winMap181",760,600,50,10,"yes","no","yes","no");
}

function funcUpdate(){
	newWin("/traffic/Update.asp","winMapUpdate",760,600,50,10,"yes","no","yes","no");
}
<%if (Session("Credit_ID")<>"0000" AND sys_City="花蓮縣") Or sys_City="台東縣" then%>
function EnterBillQry_Stop(){
	StopBillNo.value=StopBillNo.value.toUpperCase();
	StopCarNo.value=StopCarNo.value.toUpperCase();
	if (event.keyCode==13){
		Selt_Stop();
	}
}

function Selt_Stop(){
	if (StopCarNo.value=="" && StopBillNo.value==""){
		alert("必須填入單號或車號或違規人！");
	}else{
		window.open("../traffic/Query/<%
		if (sys_City="台東縣") then
			response.write "StopBillBaseData_Detail_TaiDung.asp"
		else
			response.write "StopBillBaseData_Detail.asp"
		end if 
		%>?BillNo="+StopBillNo.value+"&CarNo="+StopCarNo.value,"WebPage2","left=0,top=0,location=0,width=980,height=555,resizable=yes,scrollbars=yes,menubar=yes,status=yes");
	}
	
}
<%end if%>

function SelFunc(SystemType){
	myForm.SystemType.value=SystemType;
	myForm.submit();
}

newWin("SystemBulletin.asp","",1000,700,50,0,"yes","no","yes","no");
</script>
<%If Trim(request("SystemType"))="1" or Trim(request("SystemType"))="" then%>
<script language="vbscript"> 
	<%randomize%>
	Dim secondDiff:SecordNow=0:RndSecond=<%=fix(rnd*60)%>:RndMinute=<%=fix(rnd*16)%>
	Sub UpdateTime()
		SecordNow=SecordNow+1
		Sys_nowTime=DateAdd("s", SecordNow,secondDiff)
		LayerTime.innerText = TimeValue(Sys_nowTime)

		If (minute(Sys_nowTime) mod 20 = 0) and second(Sys_nowTime)=0 Then DownProcess(FormatDateTime(Sys_nowTime))

		If (hour(Sys_nowTime) mod 1 = 0) and (minute(Sys_nowTime) = 10+RndMinute) and second(Sys_nowTime)=RndSecond Then SQLDownProcess(FormatDateTime(Sys_nowTime,4))

		'If second(Sys_nowTime) mod (10+cdbl(RndMinute)) = 0 Then SQLDownProcess(FormatDateTime(Sys_nowTime,4))
	End Sub

	Sub SetTime(serverDateTime)  		
		secondDiff  = serverDateTime
		oInterval = setInterval("UpdateTime()", 1000)
	End Sub

	SetTime("<%=year(SysDate)&"/"&Month(SysDate)&"/"&Day(SysDate)&" "&hour(SysDate)&"："&minute(SysDate)&"："&second(SysDate)%>")
</script>
<%End if
conn.close
set conn=nothing
%>
</html>
