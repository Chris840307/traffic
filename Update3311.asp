<html>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->

<body>
<%
Server.ScriptTimeout = 65000

If request("DB_Selt")="UpdateData" Then 
	strchk1="select * from law where itemid='3311301' and illegalrule='行駛高速公路未依標誌指示行車' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311301' or rule2='3311301' or rule3='3311301') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311201' where rule1='3311301' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311201' where rule2='3311301' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311201' where rule3='3311301' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311301'"
		conn.execute strdel

		strIns="insert into law values('3311301','5','汽車行駛高速公路進入禁止通行之路段',3000,3300,3900,4500" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

		strIns="insert into law values('3311301','6','汽車行駛高速公路進入禁止通行之路段',4000,4400,	5200,	6000" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns
	End If 
	rschk1.close
	Set rschk1=Nothing 

	strchk1="select * from law where itemid='3311302' and illegalrule='行駛高速公路未依標線指示行車' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311302' or rule2='3311302' or rule3='3311302') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311202' where rule1='3311302' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311202' where rule2='3311302' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311202' where rule3='3311302' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311302'"
		conn.execute strdel

		strIns="insert into law values('3311302','0','載運危險物品車輛行駛高速公路進入禁止通行之路段',5000,	5500,	6000,	6000 " &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing 

	strchk1="select * from law where itemid='3311303' and illegalrule='行駛高速公路未依號誌指示行車' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311303' or rule2='3311303' or rule3='3311303') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311203' where rule1='3311303' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311203' where rule2='3311303' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311203' where rule3='3311303' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311303'"
		conn.execute strdel

		strIns="insert into law values('3311303','5','汽車行駛快速公路進入禁止通行之路段',3000,	3300,	3900,	4500 " &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

		strIns="insert into law values('3311303','6','汽車行駛快速公路進入禁止通行之路段',4000,	4400,	5200,	6000 " &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing
	'3311304
	strchk1="select * from law where itemid='3311304' and illegalrule='行駛快速公路未依標誌指示行車' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311304' or rule2='3311304' or rule3='3311304') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311204' where rule1='3311304' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311204' where rule2='3311304' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311204' where rule3='3311304' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311304'"
		conn.execute strdel

		strIns="insert into law values('3311304','0','載運危險物品車輛行駛快速公路進入禁止通行之路段',5000,	5500,	6000	,6000 " &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311305
	strchk1="select * from law where itemid='3311305' and illegalrule='行駛快速公路未依標線指示行車' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311305' or rule2='3311305' or rule3='3311305') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311205' where rule1='3311305' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311205' where rule2='3311305' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311205' where rule3='3311305' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311305'"
		conn.execute strdel

		strIns="insert into law values('3311305','5','汽車行駛高速公路行駛禁止通行之路段',3000,	3300,	3900,	4500 " &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

		strIns="insert into law values('3311305','6','汽車行駛高速公路行駛禁止通行之路段',4000,	4400,	5200,	6000 " &_
			",'0','1','0','0','0','0' " &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311306
	strchk1="select * from law where itemid='3311306' and illegalrule='行駛快速公路未依號誌指示行車' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311306' or rule2='3311306' or rule3='3311306') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311206' where rule1='3311306' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311206' where rule2='3311306' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311206' where rule3='3311306' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311306'"
		conn.execute strdel

		strIns="insert into law values('3311306','0','載運危險物品車輛行駛高速公路行駛禁止通行之路段',5000,	5500,	6000,	6000" &_
			",'0','1','0','0','0','0' " &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311401
	strchk1="select * from law where itemid='3311401' and illegalrule='汽車行駛高速公路進入禁止通行之路段' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311401' or rule2='3311401' or rule3='3311401') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311301' where rule1='3311401' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311301' where rule2='3311401' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311301' where rule3='3311401' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311401'"
		conn.execute strdel

		strIns="insert into law values('3311401','5','汽車行駛高速公路連續密集按鳴喇叭迫使前車讓道',3000,	3300,	3900,	4500 " &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

		strIns="insert into law values('3311401','6','汽車行駛高速公路連續密集按鳴喇叭迫使前車讓道',4000,	4400,	5200,	6000 " &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing
	'3311402
	strchk1="select * from law where itemid='3311402' and illegalrule='載運危險物品車輛行駛高速公路進入禁止通行之路段' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311402' or rule2='3311402' or rule3='3311402') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311302' where rule1='3311402' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311302' where rule2='3311402' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311302' where rule3='3311402' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311402'"
		conn.execute strdel

		strIns="insert into law values('3311402','5','汽車行駛高速公路連續變換燈光迫使前車讓道',3000,	3300,	3900,	4500" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

		strIns="insert into law values('3311402','6','汽車行駛高速公路連續變換燈光迫使前車讓道',4000,	4400,	5200,	6000" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing
	'3311403
	strchk1="select * from law where itemid='3311403' and illegalrule='汽車行駛快速公路進入禁止通行之路段' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311403' or rule2='3311403' or rule3='3311403') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311303' where rule1='3311403' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311303' where rule2='3311403' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311303' where rule3='3311403' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311403'"
		conn.execute strdel

		strIns="insert into law values('3311403','5','汽車行駛高速公路以其他方式迫使前車讓道',3000,	3300,	3900,	4500" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

		strIns="insert into law values('3311403','6','汽車行駛高速公路以其他方式迫使前車讓道',4000,	4400,	5200,	6000" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311404
	strchk1="select * from law where itemid='3311404' and illegalrule='載運危險物品車輛行駛快速公路進入禁止通行之路段' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311404' or rule2='3311404' or rule3='3311404') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311304' where rule1='3311404' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311304' where rule2='3311404' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311304' where rule3='3311404' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311404'"
		conn.execute strdel

		strIns="insert into law values('3311404','5','汽車行駛快速公路連續密集按鳴喇叭迫使前車讓道',3000,	3300,	3900,	4500" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

		strIns="insert into law values('3311404','6','汽車行駛快速公路連續密集按鳴喇叭迫使前車讓道',4000,	4400,	5200,	6000" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311405
	strchk1="select * from law where itemid='3311405' and illegalrule='汽車行駛高速公路行駛禁止通行之路段' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311405' or rule2='3311405' or rule3='3311405') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311305' where rule1='3311405' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311305' where rule2='3311405' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311305' where rule3='3311405' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311405'"
		conn.execute strdel

		strIns="insert into law values('3311405','5','汽車行駛快速公路連續變換燈光迫使前車讓道',3000,	3300,	3900,	4500" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

		strIns="insert into law values('3311405','6','汽車行駛快速公路連續變換燈光迫使前車讓道',4000,	4400,	5200,	6000" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311406
	strchk1="select * from law where itemid='3311406' and illegalrule='載運危險物品車輛行駛高速公路行駛禁止通行之路段' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311406' or rule2='3311406' or rule3='3311406') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311306' where rule1='3311406' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311306' where rule2='3311406' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311306' where rule3='3311406' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311406'"
		conn.execute strdel

		strIns="insert into law values('3311406','5','汽車行駛快速公路以其他方式迫使前車讓道',3000,	3300,	3900,	4500" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

		strIns="insert into law values('3311406','6','汽車行駛快速公路以其他方式迫使前車讓道',4000,	4400,	5200,	6000" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns


	End If 
	rschk1.close
	Set rschk1=Nothing
	'3311407
	strchk1="select * from law where itemid='3311407' and illegalrule='汽車行駛快速公路行駛禁止通行之路段' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311407' or rule2='3311407' or rule3='3311407') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311307' where rule1='3311407' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311307' where rule2='3311407' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311307' where rule3='3311407' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311407'"
		conn.execute strdel

	End If 
	rschk1.close
	Set rschk1=Nothing
	'3311408
	strchk1="select * from law where itemid='3311408' and illegalrule='載運危險物品車輛行駛快速公路行駛禁止通行之路段' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311408' or rule2='3311408' or rule3='3311408') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311308' where rule1='3311408' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311308' where rule2='3311408' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311308' where rule3='3311408' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311408'"
		conn.execute strdel

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311501
	strchk1="select * from law where itemid='3311501' and illegalrule='汽車行駛高速公路連續密集按鳴喇叭迫使前車讓道' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311501' or rule2='3311501' or rule3='3311501') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311401' where rule1='3311501' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311401' where rule2='3311501' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311401' where rule3='3311501' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311501'"
		conn.execute strdel
		
		strIns="insert into law values('3311501','0','汽車行駛於高速公路向車外丟棄物品',3000,	3300,	3900,	4500" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311502
	strchk1="select * from law where itemid='3311502' and illegalrule='汽車行駛高速公路連續變換燈光迫使前車讓道' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311502' or rule2='3311502' or rule3='3311502') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311402' where rule1='3311502' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311402' where rule2='3311502' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311402' where rule3='3311502' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311502'"
		conn.execute strdel
		
		strIns="insert into law values('3311502','0','汽車行駛於快速公路向車外丟棄物品',3000,	3300,	3900,	4500" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311503
	strchk1="select * from law where itemid='3311503' and illegalrule='汽車行駛高速公路以其他方式迫使前車讓道' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311503' or rule2='3311503' or rule3='3311503') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311403' where rule1='3311503' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311403' where rule2='3311503' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311403' where rule3='3311503' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311503'"
		conn.execute strdel

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311504
	strchk1="select * from law where itemid='3311504' and illegalrule='汽車行駛快速公路連續密集按鳴喇叭迫使前車讓道' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311504' or rule2='3311504' or rule3='3311504') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311404' where rule1='3311504' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311404' where rule2='3311504' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311404' where rule3='3311504' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311504'"
		conn.execute strdel

		strIns="insert into law values('3311504','0','汽車行駛於高速公路，行駛中向車外丟棄廢棄物',3000,	3300,	3900,	4500" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311505
	strchk1="select * from law where itemid='3311505' and illegalrule='汽車行駛快速公路連續變換燈光迫使前車讓道' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311505' or rule2='3311505' or rule3='3311505') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311405' where rule1='3311505' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311405' where rule2='3311505' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311405' where rule3='3311505' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311505'"
		conn.execute strdel

		strIns="insert into law values('3311505','0','汽車行駛於快速公路，行駛中向車外丟棄廢棄物',3000,	3300,	3900,	4500" &_
			",'0','1','0','0','0','0'" &_
			",to_date('2014/3/31','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strIns

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311506
	strchk1="select * from law where itemid='3311506' and illegalrule='汽車行駛快速公路以其他方式迫使前車讓道' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311506' or rule2='3311506' or rule3='3311506') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311406' where rule1='3311506' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311406' where rule2='3311506' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311406' where rule3='3311506' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311506'"
		conn.execute strdel

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311601
	strchk1="select * from law where itemid='3311601' and illegalrule='汽車行駛於高速公路向車外丟棄物品' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311601' or rule2='3311601' or rule3='3311601') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311501' where rule1='3311601' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311501' where rule2='3311601' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311501' where rule3='3311601' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311601'"
		conn.execute strdel

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311602
	strchk1="select * from law where itemid='3311602' and illegalrule='汽車行駛於快速公路，行駛中向車外丟棄物品' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311602' or rule2='3311602' or rule3='3311602') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311502' where rule1='3311602' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311502' where rule2='3311602' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311502' where rule3='3311602' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311602'"
		conn.execute strdel

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311604
	strchk1="select * from law where itemid='3311604' and illegalrule='汽車行駛於高速公路，行駛中向車外丟棄廢棄物' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311604' or rule2='3311604' or rule3='3311604') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311504' where rule1='3311604' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311504' where rule2='3311604' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311504' where rule3='3311604' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311604'"
		conn.execute strdel

	End If 
	rschk1.close
	Set rschk1=Nothing

	'3311605
	strchk1="select * from law where itemid='3311605' and illegalrule='汽車行駛於快速公路，行駛中向車外丟棄廢棄物' and version=2"
	Set rschk1=conn.execute(strchk1)
	If Not rschk1.eof Then
		str1="select count(*) as cnt from billbase where (rule1='3311605' or rule2='3311605' or rule3='3311605') and recordstateid=0"
		Set rs1=conn.execute(str1)
		If not rs1.eof Then
			If CDbl(rs1("cnt"))>0 Then
				response.write rs1("cnt")
				strUpd1="update billbase set rule1='3311505' where rule1='3311605' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule2='3311505' where rule2='3311605' and recordstateid=0"
				conn.execute strUpd1

				strUpd1="update billbase set rule3='3311505' where rule3='3311605' and recordstateid=0"
				conn.execute strUpd1
			End If 
		End If 
		rs1.close
		Set rs1=Nothing 

		strdel="delete from law where version=2 and itemid='3311605'"
		conn.execute strdel

	End If 
	rschk1.close
	Set rschk1=Nothing
End If 
%>
<form name="myForm">
<input type="button" name="t1" value="檢查" onclick="GetData()">

<input type="button" name="t1" value="更新" onclick="UpdateData()">

<input type="hidden" name="DB_Selt" value="">

<%
If request("DB_Selt")="Selt" Then 

	str1="select count(*) as cnt from billbase where (rule1='3311301' or rule2='3311301' or rule3='3311301') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311301: " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing 

	str1="select count(*) as cnt from billbase where (rule1='3311302' or rule2='3311302' or rule3='3311302') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311302: " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311303' or rule2='3311303' or rule3='3311303') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311303: " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311304' or rule2='3311304' or rule3='3311304') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311304: " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311305' or rule2='3311305' or rule3='3311305') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311305: " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing


	str1="select count(*) as cnt from billbase where (rule1='3311306' or rule2='3311306' or rule3='3311306') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311306 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	'===

	str1="select count(*) as cnt from billbase where (rule1='3311401' or rule2='3311401' or rule3='3311401') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311401 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311402' or rule2='3311402' or rule3='3311402') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311402 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311403' or rule2='3311403' or rule3='3311403') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311403 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311404' or rule2='3311404' or rule3='3311404') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311404 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311405' or rule2='3311405' or rule3='3311405') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311405 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311406' or rule2='3311406' or rule3='3311406') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311406 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311407' or rule2='3311407' or rule3='3311407') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311407 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311408' or rule2='3311408' or rule3='3311408') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311408 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311501' or rule2='3311501' or rule3='3311501') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311501 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311502' or rule2='3311502' or rule3='3311502') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311502 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311504' or rule2='3311504' or rule3='3311504') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311504 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311505' or rule2='3311505' or rule3='3311505') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311505 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311506' or rule2='3311506' or rule3='3311506') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311506 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311601' or rule2='3311601' or rule3='3311601') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311601 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311602' or rule2='3311602' or rule3='3311602') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311602 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311604' or rule2='3311604' or rule3='3311604') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311604 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing

	str1="select count(*) as cnt from billbase where (rule1='3311605' or rule2='3311605' or rule3='3311605') and recordstateid=0"
	Set rs1=conn.execute(str1)
	If not rs1.eof Then
		If CDbl(rs1("cnt"))>0 then
		response.write "3311605 " & ": " & rs1("cnt") &"<br>"
		End If 
	End If 
	rs1.close
	Set rs1=Nothing
End if
%>
</form>
</body>

<script>
function GetData()
{
myForm.DB_Selt.value="Selt";
myForm.submit();

}
function UpdateData()
{
myForm.DB_Selt.value="UpdateData";
myForm.submit();

}
</script>
<%
conn.close
Set conn=nothing
%>
</html>