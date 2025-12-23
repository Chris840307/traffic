<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"(台南+金門)建檔日_舉發單資料.txt"
Response.AddHeader "Content-Disposition", "attachment;filename=" &fname
response.contenttype="application/vnd.ms-txt" 

Server.ScriptTimeout = 650000

%>
<%
	'檢查是否可進入本系統
	'AuthorityCheck(234)

	strSQL="select a.sn,a.billno,a.Carno,a.BillTypeID,a.IllegalDate,a.BillFillDate,a.RecordStateID,a.Recorddate,b.DciCaseInDate,b.Status,c.FileName,b.dcireturnstation from Billbase a,BillBaseDciReturn b,DciLog c where a.RecordDate between to_date('"&gOutDT(request("date1"))&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&gOutDT(request("date2"))&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
	" and a.sn=c.BillSn and b.ExChangeTypeID=c.ExChangeTypeID and b.Status=c.DciReturnStatusID" &_
	" and a.BillNo=b.BillNo and a.CarNo=b.CarNo and b.ExChangeTypeID='W' " &_
	" and b.dcireturnstation in ('74','36')" &_
	" and a.Recordstateid=0 order by b.dcireturnstation,a.Recorddate "
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)


%>監理站,單號,建檔日,上傳日,上傳是否成功, 上傳檔名
<%  
					If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
					Response.flush
				
					CaseInDate=trim(rsfound("DciCaseInDate"))
					If trim(rsfound("dcireturnstation"))="74" then
						response.write "台南監理站"
					Else
						response.write "金門監理站"
					End if
					%>,<%
					'告發單號11
					response.write rsfound("BillNo")
					%>,<%'建檔時間12
					if trim(rsfound("RecordDate"))<>"" and not isnull(rsfound("RecordDate")) then
						response.write year(trim(rsfound("RecordDate")))-1911&right("00"&month(trim(rsfound("RecordDate"))),2)&right("00"&day(trim(rsfound("RecordDate"))),2)
					end if
					%>,<%'入案日期9
					response.write CaseInDate
					%>,<%'上傳是否成功
					strS="select * from dcireturnstatus where DciActionID='W' and DciReturn='"&trim(rsfound("Status"))&"'"
					set rsS=conn.execute(strS)
					if not rsS.eof then
						response.write trim(rsS("StatusContent"))
					end if
					rsS.close
					set rsS=nothing
					%>,<%'上傳檔名
					response.write rsfound("FileName")
					

				response.write vbCrLf
				rsfound.MoveNext
				Wend
				rsfound.close
				set rsfound=nothing
				%>
				

<%
conn.close
set conn=nothing
%>