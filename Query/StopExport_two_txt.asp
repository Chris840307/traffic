
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

fname= sYear & fMnoth & fDay & "30.bak"
Response.Buffer = true
Response.AddHeader "Content-Disposition","attachment;filename=" &fname
Response.ContentType = "application/vnd.ms-txt"
Server.ScriptTimeout = 800
Response.flush

	UserMarkDate1=gOutDT(request("Sys_SendMarkDate1"))&" 0:0:0"
	UserMarkDate2=gOutDT(request("Sys_SendMarkDate2"))&" 23:59:59"

	strSQL=""
	
	strQuery="select distinct a.*,RPAD(a.CarNo,9,' ') as sCarNo  , RPAD(a.ForFeit1,6,' ') as ForFeit1 , RPAD(a.CarSimpleid,2,' ') as CarSimpleid , RPAD(a.illegaladdress,41,' ') as illegaladdress, RPAD(a.imagepathname,17,' ') as imagepathname, RPAD(a.imagefilenameb,17,' ') as imagefilenameb ,e.DciReturnCarColor,e.Owner,e.OwnerAddress,e.OwnerZip,e.A_Name from (select * from billbase where imagefilenameb in(select BillNo from StopCarSendAddress where UserMarkDate between TO_DATE('"&UserMarkDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&UserMarkDate2&"','YYYY/MM/DD/HH24/MI/SS'))) a,DciLog b,BillBaseDciReturn e where a.CarNo=e.CarNo and e.ExchangeTypeID='A' and e.Status='S' and a.Sn=b.BillSn and a.RecordStateID=0 order by a.RecordDate"
	
	'response.write strQuery
	'response.end
	set rsfound=conn.execute(strQuery)
	

					If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
					'車號
						response.write rsfound("sCarNo")
						
					'違規日期
						if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
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
						if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
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
						
						response.write rsfound("imagefilenameb")
												
					
					rsfound.MoveNext
					response.write(vbCrLf)
					Wend
					rsfound.close
					set rsfound=nothing
				
conn.close
set conn=nothing

%>