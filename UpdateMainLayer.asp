<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： UpdateMainLayer.asp
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	'顯示dcilog上傳記錄
	strUpdDci="<table border='1' width='100%'><tr bgcolor='#FFFF99'><td colspan='2'>"
	if sys_City="雲林縣" then
		strUpdDci=strUpdDci&"監理站資料交換進度(一天內)"
	else
		strUpdDci=strUpdDci&"監理站資料交換進度(三天內)"
	end if
	strUpdDci=strUpdDci&"</td></tr><tr bgcolor='#99FF99'><td align='center' width='50%'>批號</td><td align='center' width='50%'>狀態</td></tr>"
	if sys_City="雲林縣" then
		strUp="select distinct BatchNumber from Dcilog where ExchangeDate between " &_
			" TO_DATE('"&Date&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') " &_
			" and TO_DATE('"&Date&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') " &_
			" and RecordMemberID="&trim(Session("User_ID"))&" order by BatchNumber desc"
	else
		strUp="select distinct BatchNumber from Dcilog where ExchangeDate between " &_
			" TO_DATE('"&DateAdd("d",-3,Date)&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') " &_
			" and TO_DATE('"&Date&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') " &_
			" and RecordMemberID="&trim(Session("User_ID"))&" order by BatchNumber desc"
	end if
	set rsUp=conn.execute(strUp)
	If Not rsUp.Bof Then rsUp.MoveFirst 
	While Not rsUp.Eof
			

			strStatus="select DciReturnStatusID from Dcilog where batchNumber='"&trim(rsUp("BatchNumber"))&"' and rownum=1"
			set rsStatus=conn.execute(strStatus)
			if not rsStatus.eof then
				if not isnull(rsStatus("DciReturnStatusID")) Then
					strUpdDci=strUpdDci&"<tr bgColor='#CCFFCC'><td align='center' width='50%'>"&trim(rsUp("BatchNumber"))&"</td>"
					strUpdDci=strUpdDci&"<td align='center' width='50%'>"
					strUpdDci=strUpdDci&"已回傳"
					strUpdDci=strUpdDci&"</td></tr>"

				Else
					strUpdDci=strUpdDci&"<tr bgColor='#FFCCCC'><td align='center' width='50%'>"&trim(rsUp("BatchNumber"))&"</td>"
					strUpdDci=strUpdDci&"<td align='center' width='50%'>"
					strUpdDci=strUpdDci&"未回傳"
					strUpdDci=strUpdDci&"</td></tr>"

				end if
			end if
			rsStatus.close
			set rsStatus=nothing
			
			
	rsUp.MoveNext
	Wend
	rsUp.close
	set rsUp=nothing
	strUpdDci=strUpdDci&"</table>"
'========================================================
	'今日建檔
	sStartDate=Date & " 00:00:00 "
	sEndDate=Date & " 23:59:59 "
	strSQL="Select count(*) billcount from BillBase Where RecordDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS')" &" and RecordMemberID="& Session("User_ID") & " and RecordStateID <> -1 "

	set rssysinfo=conn.execute(strSQL)
	billcount=rssysinfo("billcount") 
	set rssysinfo=nothing

	'入案
	strSQL="select count(ExchangeTypeID) as Num from DCILog Where ExchangeDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS')" &" and RecordMemberID="& Session("User_ID")  & " and ExchangeTypeID='W' "
	set rssysinfo=conn.execute(strSQL)
	Sendcount=rssysinfo("Num") 
	set rssysinfo=nothing

	'成功
	strwhere=" and a.RecordMemberID="& Session("User_ID") & " and a.ExchangeDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS') and a.ExchangeTypeID='W'"
		if sys_City="基隆市" then
			strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','F','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','9','a','j','A','F','H','K','L','T','V') and (a.BillNo in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') or usetool=8) "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		else
			strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','9','a','j','A','H','K','L','T','V') and usetool=8 "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		end if


		set chksuess=conn.execute(strSQL)
		filsuess=CDbl(chksuess("cnt"))
		chksuess.close
		
		'異常
		strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='-1' "&strwhere

		set chksuess=conn.execute(strSQL)
		fildel=CDbl(chksuess("cnt"))
		chksuess.close

		'無效
		if sys_City="基隆市" then
			strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and a.DciErrorCarData in ('1','3','9','a','j','A','F','H','K','L','T','V') and a.BillNo not in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') and usetool<>8 and d.DCIRETURNSTATUS='1'"&strwhere
		else
			strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and a.DciErrorCarData in ('1','3','9','a','j','A','H','K','L','T','V') and usetool<>8 and d.DCIRETURNSTATUS='1'"&strwhere
		end if
		set Dbrs=conn.execute(strCnt)
		errCatCnt=CDbl(Dbrs("cnt"))
		Dbrs.close

		'全部
		strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) "&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum=CDbl(Dbrs("cnt"))
		Dbrs.close

		'刪除
		strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and a.ExchangeTypeID='E' and d.DCIRETURNSTATUS='1'"&strwhere
		set Dbrs=conn.execute(strCnt)
		deldata=CDbl(Dbrs("cnt"))
		Dbrs.close
'========================================================
	'昨日建檔
	sStartDate=DateAdd("d",-1,Date) & " 00:00:00 "
	sEndDate=DateAdd("d",-1,Date) & " 23:59:59 "
	strSQL="Select count(*) billcount2 from BillBase Where RecordDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS')" &" and RecordMemberID="& Session("User_ID") & " and RecordStateID <> -1 "

	set rssysinfo=conn.execute(strSQL)
	billcount2=rssysinfo("billcount2") 
	set rssysinfo=nothing

	'入案
	strSQL="select count(ExchangeTypeID) as Num from DCILog Where ExchangeDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS')" &" and RecordMemberID="& Session("User_ID") & " and ExchangeTypeID='W' "
	set rssysinfo=conn.execute(strSQL)
	Sendcount2=rssysinfo("Num") 
	set rssysinfo=nothing

	'成功
	strwhere="and a.RecordMemberID="& Session("User_ID") & " and a.ExchangeDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS') and a.ExchangeTypeID='W'"
	if sys_City="基隆市" then
			strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','F','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','9','a','j','A','F','H','K','L','T','V') and (a.BillNo in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') or usetool=8) "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		else
			strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T','V') "&strwhere&") or (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData in ('1','3','9','a','j','A','H','K','L','T','V') and usetool=8 "&strwhere&") or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&")"
		end if


		set chksuess=conn.execute(strSQL)
		filsuess2=CDbl(chksuess("cnt"))
		chksuess.close
	
	'異常
	strSQL="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and d.DCIRETURNSTATUS='-1' "&strwhere

		set chksuess=conn.execute(strSQL)
		fildel2=CDbl(chksuess("cnt"))
		chksuess.close
	
	'無效
	if sys_City="基隆市" then
			strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and a.DciErrorCarData in ('1','3','9','a','j','A','F','H','K','L','T','V') and a.BillNo not in (select distinct BillNo from BillBaseDCIReturn where EXCHANGETYPEID='W' and Rule4='2607') and usetool<>8 and d.DCIRETURNSTATUS='1'"&strwhere
		else
			strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and a.DciErrorCarData in ('1','3','9','a','j','A','H','K','L','T','V') and usetool<>8 and d.DCIRETURNSTATUS='1'"&strwhere
		end if
		set Dbrs=conn.execute(strCnt)
		errCatCnt2=CDbl(Dbrs("cnt"))
		Dbrs.close
	
	'刪除
	strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) and a.ExchangeTypeID='E' and d.DCIRETURNSTATUS='1'"&strwhere
		set Dbrs=conn.execute(strCnt)
		deldata2=CDbl(Dbrs("cnt"))
		Dbrs.close	

	'全部
	strCnt="select count(*) as cnt from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN(+) and a.BillNo=h.BillNo(+) "&strwhere
		set Dbrs=conn.execute(strCnt)
		DBsum2=CDbl(Dbrs("cnt"))
		Dbrs.close
	'===============================================
	strUpdYDay="<昨日><br>建檔 : "&billcount2&" &nbsp; &nbsp;  入案 : "&Sendcount2&" &nbsp; &nbsp;  成功 : "&filsuess2&" <br>異常 : "&fildel2+errCatCnt2&" &nbsp; &nbsp;  未處理 : "&DBsum2-CDbl(filsuess2)-CDbl(fildel2)-CDbl(deldata2)-CDbl(errCatCnt2)

	strUpdToDay="<今日><br>建檔 : "&billcount&" &nbsp; &nbsp;  入案 : "&Sendcount&" &nbsp; &nbsp;  成功 : "&filsuess&" <br>異常 : "&fildel+errCatCnt&" &nbsp; &nbsp;  未處理 : "&DBsum-CDbl(filsuess)-CDbl(fildel)-CDbl(deldata)-CDbl(errCatCnt)
%>
			

UpLoadLayer.innerHTML="<%=strUpdDci%>";
YestodayLayer.innerHTML="<%=strUpdYDay%>";
TodayLayer.innerHTML="<%=strUpdToDay%>";
<%
conn.close
set conn=nothing
%>
