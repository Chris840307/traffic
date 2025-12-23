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
fname= sYear & fMnoth & fDay & "30Store.bak"
Response.AddHeader "Content-Disposition","attachment;filename=" &fname
Response.ContentType = "application/vnd.ms-txt"

strwhere=trim(request("SQLstr"))

strQuery="select distinct a.IllegalDate,a.CarNo as sCarNo,a.ForFeit1 as ForFeit1,a.CarSimpleid as CarSimpleid,a.illegaladdress as illegaladdress,a.imagepathname as imagepathname,a.imagefilenameb as imagefilenameb,a.RecordDate,c.UserMarkDate from (select * from BillBase where ImagePathName is not null and RecordStateId <> -1) a,(Select distinct BillSN,BatchNumber from DCILog where ExchangeTypeID='A' and DCIReturnStatusID='S') b,(select * from StopBillMailHistory where UserMarkResonID in('5','6','7','T')) c where a.SN=b.BillSN and a.SN=c.BillSn and a.imagefilenameb=c.BillNo "&strwhere&" order by c.UserMarkDate"
set rsfound=conn.execute(strQuery)
While Not rsfound.Eof
'車號
	response.write rsfound("sCarNo")&  ","
	
'違規日期
	sDate=gInitDT(rsfound("IllegalDate"))
	response.write sDate &  ","

'違規時間
	response.write  Right("00"&hour(trim(rsfound("IllegalDate"))),2)&Right("00"&minute(trim(rsfound("IllegalDate"))),2) 
	response.write ","
'停車費
	response.write rsfound("ForFeit1")&  ","


'簡式車種1汽車 / 2拖車/ 3重機/ 4輕機 
	response.write rsfound("CarSimpleID")&  ","
	
'違規地點						
	response.write rsfound("IllegalAddress")&  ","
							
'停車單號
	response.write rsfound("imagepathname")&  ","
	
'催繳單號
	
	response.write rsfound("imagefilenameb")&vbnewline
	rsfound.MoveNext
Wend
rsfound.close
set rsfound=nothing
conn.close
set conn=nothing
%>