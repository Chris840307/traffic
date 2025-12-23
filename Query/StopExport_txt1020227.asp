
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

	'檢查是否可進入本系統
	'AuthorityCheck(234)
	strwhere=trim(request("SQLstr"))

	strQuery="select distinct a.*,a.CarNo as sCarNo , a.ForFeit1 as ForFeit1 , a.CarSimpleid as CarSimpleid , a.illegaladdress as illegaladdress, a.imagepathname as imagepathname, a.imagefilenameb as imagefilenameb ,e.DciReturnCarColor,e.Owner,e.OwnerAddress,e.OwnerZip,e.A_Name from BillBase a,DciLog b,BillBaseDciReturn e where a.CarNo=e.CarNo and e.ExchangeTypeID='A' and e.Status='S' and a.Sn=b.BillSn and a.RecordStateID=0 "&strwhere&" order by a.RecordDate"
	
	'response.write strQuery
	'response.end
	set rsfound=conn.execute(strQuery)
	

					If Not rsfound.Bof Then rsfound.MoveFirst 
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
						
						response.write rsfound("imagefilenameb")
													
						
						rsfound.MoveNext
						response.write(vbCrLf)
					Wend
					rsfound.close
					set rsfound=nothing
				
conn.close
set conn=nothing

%>