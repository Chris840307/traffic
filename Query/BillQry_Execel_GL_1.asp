<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"(基隆)_攔停逕舉舉發單資料.txt"
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
	rstmp.close
	Set rstmp=Nothing 
End Function

Function GetOwner(Billno,Carno)
tmp=""
tmpsql="select owner from billbasedcireturn where billno='"&Billno&"' and carno='"&Carno&"' and exchangetypeid='W'"
	set rstmp=conn.execute(tmpsql)
	If Not rstmp.eof Then tmp=rstmp(0)
		GetOwner=tmp
	rstmp.close
	Set rstmp=Nothing 
End Function

Function GetDriver(Billno,Carno)
tmp=""
tmpsql="select Driver from billbasedcireturn where billno='"&Billno&"' and carno='"&Carno&"' and exchangetypeid='W'"
	set rstmp=conn.execute(tmpsql)
	If Not rstmp.eof Then tmp=rstmp(0)
		GetDriver=tmp
	rstmp.close
	Set rstmp=Nothing 
End Function

Function GetCaseInDate(Billno,Carno)
tmp=""
tmpsql="select DCICaseInDate from billbasedcireturn where billno='"&Billno&"' and carno='"&Carno&"' and exchangetypeid='W'"
	set rstmp=conn.execute(tmpsql)
	If Not rstmp.eof Then tmp=rstmp(0)
		GetCaseInDate=gOutDT(tmp)
	rstmp.close
	Set rstmp=Nothing 
End Function

Function GetCdate(tdate)
	GetCdate=Year(tdate)-1911&"/"&Right("0"&Month(tdate),2)&"/"&Right("0"&day(tdate),2)
End function

	'檢查是否可進入本系統
	'AuthorityCheck(234)
	'欄停逕舉
	strSQL="select a.BillTypeID,a.Billunitid,a.DealLineDate,a.BillfillDate,a.CarNo,a.BillMem1,a.Rule1,a.Billno,a.Illegaldate,b.DCICaseInDate,b.Driver,b.DriverID,b.Owner,b.OwnerID from BillBase a,BillbaseDcireturn b where a.RecordStateID=0 and a.recorddate between to_date('"&date1&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&date2&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS') and a.Billno=b.billno and a.carno=b.carno and b.exchangetypeid='W'	"
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)


%>舉發類型,舉發單位,違規人證號,違規車號,入案日期,車主/違規人,違規法條一,單號,違規日
<%  
					If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
					Response.flush
					'舉發類型
					If trim(rsfound("BillTypeID"))="1" Then 
						response.write "攔停"
					Else
						response.write "逕舉"
					End if
					%>,<%
					'舉發單位
					response.write GetDataName("UnitName","UnitInfo","UnitID",rsfound("Billunitid"))
					%>,<%
					'證號
					If trim(rsfound("BillTypeID"))="1" Then 
					    response.write trim(rsfound("DriverID"))
					Else
					    response.write trim(rsfound("OwnerID"))
					End If
					%>,<%
					'車號 
				    response.write rsfound("CarNo")
					%>,<%
					'入案日期 
					If Trim(rsfound("DCICaseInDate"))<>"" then
						response.write gOutDT(Trim(rsfound("DCICaseInDate")))
					End If 
					%>,<%
					'車主/違規人
				    If trim(rsfound("BillTypeID"))="1" Then 
						response.write trim(rsfound("Driver"))
					Else
						response.write trim(rsfound("Owner"))
					End if
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