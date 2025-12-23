<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"(台東)違規日_舉發單資料.txt"
Response.AddHeader "Content-Disposition", "attachment;filename=" &fname
response.contenttype="application/vnd.ms-txt" 

Server.ScriptTimeout = 650000

date1=gOutDT(request("date1"))
date2=gOutDT(request("date2"))

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
	GetCdate=Year(tdate)-1911&"/"&Right("0"&Month(tdate),2)&"/"&Right("0"&day(tdate),2)
End function

%>
舉發類型,舉發單位,到案地點,應到案日期,證號,資料建檔日期,車號,舉發員警,入案日期,違規法條一,單號,違規日
<%  
	'檢查是否可進入本系統
	'AuthorityCheck(234)

	strSQL="select Memberstation,BillTypeID,Billunitid,DealLineDate,driverid,BillfillDate,CarNo,BillMem1,Rule1,Billno,Illegaldate,recorddate from passerbase where RecordStateID=0 and recorddate between to_date('"&date1&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&date2&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')	"
	set rsfound=conn.execute(strSQL)

If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
					Response.flush
					'舉發類型
					If trim(rsfound("BillTypeID"))="1" Then 
						response.write "慢車行人攤販"
					Else
						response.write "慢車行人攤販"
					End if
					%>,<%
					'舉發單位
					response.write GetDataName("UnitName","UnitInfo","UnitID",rsfound("Billunitid"))
					%>,<%
					'到案地點 Memberstation
					response.write GetDataName("UnitName","UnitInfo","UnitID",rsfound("Memberstation"))
					%>,<%
					'應到案日期 
				    response.write GetCdate(rsfound("DealLineDate"))
					%>,<%
					'證號
					    response.write rsfound("driverid")
					%>,<%
					'資料建檔日期 
				    response.write GetCdate(rsfound("BillfillDate"))
					%>,<%
					'車號 
				    response.write rsfound("CarNo")
					%>,<%
					'舉發員警 
				    response.write rsfound("BillMem1")
					%>,<%
					'入案日期 
				    response.write GetCdate(rsfound("recorddate"))
					%>,<%
					'違規法條一 
				    response.write rsfound("Rule1")
					%>,<%
					'單號 
				    response.write rsfound("Billno")
					%>,<%
					'違規日 
				    response.write GetCdate(rsfound("illegaldate"))


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