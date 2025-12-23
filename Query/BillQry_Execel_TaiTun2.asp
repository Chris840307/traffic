<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"(台東)違規日_舉發單資料69-84.txt"
Response.AddHeader "Content-Disposition", "attachment;filename=" &fname
response.contenttype="application/vnd.ms-txt" 

Server.ScriptTimeout = 650000

date1="2011/01/01"
date2="2011/11/30"

Function GetDataName(tName,tTable,Twhere,tID)
tmp=""
tmpsql="select "&tName&" from "&tTable&" where "&Twhere&"='"&tID&"'"
	set rstmp=conn.execute(tmpsql)
	If Not rstmp.eof Then tmp=rstmp(0)
GetDataName=tmp
End Function

Function GetDataName2(Billno,Carno)
tmp=""
tmpsql="select DciReturnStation from billbasedcireturn where billno='"&Billno&"' and carno='"&Carno&"' and exchangetypeid='W'"
	set rstmp=conn.execute(tmpsql)
	If Not rstmp.eof Then tmp=rstmp(0)
GetDataName2=tmp
End Function

Function GetOwner(Billno,Carno)
tmp=""
tmpsql="select owner from billbasedcireturn where billno='"&Billno&"' and carno='"&Carno&"' and exchangetypeid='W'"
	set rstmp=conn.execute(tmpsql)
	If Not rstmp.eof Then tmp=rstmp(0)
GetOwner=tmp
End Function

Function GetDriver(Billno,Carno)
tmp=""
tmpsql="select Driver from billbasedcireturn where billno='"&Billno&"' and carno='"&Carno&"' and exchangetypeid='W'"
	set rstmp=conn.execute(tmpsql)
	If Not rstmp.eof Then tmp=rstmp(0)
GetDriver=tmp
End Function

Function GetCaseInDate(Billno,Carno)
tmp=""
tmpsql="select DCICaseInDate from billbasedcireturn where billno='"&Billno&"' and carno='"&Carno&"' and exchangetypeid='W'"
	set rstmp=conn.execute(tmpsql)
	If Not rstmp.eof Then tmp=rstmp(0)
GetCaseInDate=tmp
End Function

Function GetCdate(tdate)
	If tdate<>"" Then 
	GetCdate=Year(tdate)-1911&"/"&Right("0"&Month(tdate),2)&"/"&Right("0"&day(tdate),2)
	Else
	GetCdate=""
	End if
End Function

Function GetDataName3(tName,tTable,Twhere,tID)
tmp=""
tmpsql="select "&tName&" from "&tTable&" where ArriveType=0 and "&Twhere&"='"&tID&"'"
	set rstmp=conn.execute(tmpsql)
	If Not rstmp.eof Then tmp=rstmp(0)
GetDataName3=tmp
End Function

%>
裁決單位,到案日期,違規人證號,違規法條一,告發單單號,違規日期,裁決日,送達註記,列管狀態,繳結狀態
<%  
	strSQL="select sn,MemberStation,DeallineDate,DriverID,Rule1,BillNO,Illegaldate,BillStatus from passerbase where RecordStateID=0 and IllegalDate between to_date('"&date1&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&date2&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')	"
	set rsfound=conn.execute(strSQL)

If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
					Response.flush

					'裁決單位
					response.write GetDataName("UnitName","UnitInfo","UnitID",rsfound("Memberstation"))
					%>,<%
					'應到案日期 
				    response.write GetCdate(rsfound("DealLineDate"))
					%>,<%
					'證號
					    response.write rsfound("DriverID")
					%>,<%
					'違規法條一 
				    response.write (rsfound("Rule1"))
					%>,<%
					'告發單單號 
				    response.write rsfound("BillNO")
					%>,<%
					'違規日期 
				    response.write GetCdate(rsfound("Illegaldate"))
					%>,<%
					'裁決日 
				    response.write GetCdate(GetDataName("JudeDate","PasserJude","Billsn",rsfound("sn")))
					%>,<%
					'送達註記 
				    response.write GetCdate(GetDataName3("ArrivedDate","PassersEndArrived","passersn",rsfound("sn")))
					%>,<%
					'列管狀態 
				    response.write "無"
					%>,<%
					'繳結狀態 
					If rsfound("BillStatus")="9" Then 
					    response.write "已結"
					Else
					    response.write "未結"
					End if


				response.write vbCrLf
				rsfound.MoveNext
				Wend
				rsfound.close
				set rsfound=Nothing
				%>
				

<%
conn.close
set conn=nothing
%>