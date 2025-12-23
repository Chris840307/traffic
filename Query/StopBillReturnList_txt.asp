<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
	fMnoth=month(now)
	if fMnoth<10 then fMnoth="0"&fMnoth
	fDay=day(now)
	if fDay<10 then	fDay="0"&fDay

	if len(year(now)-1911)=2 then
		sYear = "0" & year(now)-1911
	else
		sYear = year(now)-1911
	end if
	fname= sYear & fMnoth & fDay & "_NotReturn.bak"
	Response.AddHeader "Content-Disposition","attachment;filename=" &fname
	Response.ContentType = "application/vnd.ms-txt"


	If (not ifnull(request("Sys_RecordDate1"))) and (not ifnull(request("Sys_RecordDate2"))) Then

		ArgueDate1=gOutDT(request("Sys_RecordDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("Sys_RecordDate2"))&" 23:59:59"

		strwhere="and RecordDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and RecordMemberID="&session("User_ID")

	end if

	if not ifnull(Request("Sys_BatChNumber")) then

		strwhere=strwhere&" and sn in(select Billsn from DciLog where Batchnumber='"&trim(Request("Sys_BatChNumber"))&"')"
	End if


	strSQL="select distinct a.IllegalDate,RPAD(a.CarNo,9,' ') as sCarNo,RPAD(a.ForFeit1,6,' ') as ForFeit1,RPAD(a.CarSimpleid,2,' ') as CarSimpleid,RPAD(a.illegaladdress,41,' ') as illegaladdress,RPAD(a.imagepathname,17,' ') as imagepathname,RPAD(a.imagefilenameb,17,' ') as imagefilenameb,a.RecordDate from (select IllegalDate,carno,ForFeit1,CarSimpleid,illegaladdress,imagepathname,imagefilenameb,RecordDate from billbase where BillStatus=2 and imagefilenameb is not null "&strwhere&") a,(select distinct billno,carno,MailDate,mailnumber from StopBillMailHistory where billno is not null and mailnumber is not null and UserMarkResonID is null) b where a.imagefilenameb=b.billno and a.carno=b.carno order by imagefilenameb"

	set rsfound=conn.execute(strSQL)
	While Not rsfound.Eof
		'車號
		response.write rsfound("sCarNo")
			
		'違規日期
		if Not ifnull(rsfound("IllegalDate")) then
			sDate=gInitDT(rsfound("IllegalDate"))
			if Len(sDate)>"6" then							
				response.write sDate &  " "
			else
				response.write "0" & sDate &  " "	
			end if	
		else
			response.write ""
		end if

		'違規時間
		if Not ifnull(rsfound("IllegalDate")) then
			response.write  Right("00"&hour(trim(rsfound("IllegalDate"))),2)&Right("00"&minute(trim(rsfound("IllegalDate"))),2) 
		else
			response.write  ""
		end if
		response.write " "

		'停車費
		response.write rsfound("ForFeit1")

		'簡式車種1汽車 / 2拖車/ 3重機/ 4輕機 
		response.write rsfound("CarSimpleID")

		'違規地點						
		response.write rsfound("IllegalAddress")
								
		'停車單號
		response.write rsfound("imagepathname")

		'催繳單號
		response.write rsfound("imagefilenameb")&vbnewline
		rsfound.MoveNext
	Wend
	rsfound.close
	set rsfound=nothing
	conn.close
	set conn=nothing
%>