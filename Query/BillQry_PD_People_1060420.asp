<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"違規日_慢車行人舉發單資料.txt"
Response.AddHeader "Content-Disposition", "attachment;filename=" &fname
response.contenttype="application/vnd.ms-txt" 

Server.ScriptTimeout = 650000

	'檢查是否可進入本系統
	'AuthorityCheck(234)

	strSQL="select a.Rule1,a.Rule2,a.billstatus,a.sn,a.billno,a.Carno,a.BillTypeID,a.IllegalDate,a.IllegalAddress,a.MemberStation,a.BillFillDate,a.BillUnitiD,a.DeallineDate,a.CarSimpleID,a.SignType,a.RecordStateID,a.Recorddate,a.RecordMemberID,a.note from PasserBase a where a.IllegalDate between to_date('2013/01/01 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('2013/1/10 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
	" and Recordstateid=0 and billno is not null order by IllegalDate"
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)


%>違規日期;違規時間;舉發單號;舉發單位;裁決日;結案狀態;應到案日期;車號;入案日期;應到案處所;舉發類別;填單日期;法條1;法條2
<%  
					If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
					Response.flush

					DciReturnStation=""
					CaseInDate=""
					IllegalMemID=""
					IllegalMem=""
					IllegalAddress=""
					OwnerName=""
					OwnerAddress=""
					DciCarTypeID=""
					SecondAddress=""
					Rule1=""
					Rule2=""
					ForFeit1=""
					ForFeit2=""

										
					
					'違歸日期
					if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
						response.write trim(Year(rsfound("IllegalDate"))-1911)&Right("00"&Month(rsfound("IllegalDate")),2)&Right("00"&day(rsfound("IllegalDate")),2) 
					End If
					response.write ";"
					'違歸時間
					if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
						response.write Right("00"&hour(rsfound("IllegalDate")),2)&Right("00"&minute(rsfound("IllegalDate")),2)
					End If
					response.write ";"
					'舉發單號
					response.write rsfound("Billno")
					response.write ";"
					'舉發單位
					If trim(rsfound("BillUnitID"))<>"" Then
						strBUnit="select UnitName from UnitInfo where UnitID='"&Trim(trim(rsfound("BillUnitID")))&"'"
						Set rsBunit=conn.execute(strBUnit)
						If Not rsBunit.eof Then
							response.write trim(rsBunit("UnitName"))
						End If
						rsBunit.close
						Set rsBunit=Nothing
					End If 
					response.write ";"
					'裁決日
						strBUnit="select JudeDate from PasserJude where billsn='"&Trim(trim(rsfound("sn")))&"'"
						Set rsBunit=conn.execute(strBUnit)
						If Not rsBunit.eof Then
							response.write trim(Year(rsBunit("JudeDate"))-1911)&Right("00"&Month(rsBunit("JudeDate")),2)&Right("00"&day(rsBunit("JudeDate")),2)
						End If
						rsBunit.close
						Set rsBunit=Nothing

					response.write ";"
					'結案狀態
					If rsfound("billstatus")="9" then
						response.write "結案"
					Else
						response.write "正常"
					End if
					response.write ";"

					'應到案日期
					if trim(rsfound("DeallineDate"))<>"" and not isnull(rsfound("DeallineDate")) then
						response.write trim(Year(rsfound("DeallineDate"))-1911)&Right("00"&Month(rsfound("DeallineDate")),2)&Right("00"&day(rsfound("DeallineDate")),2)
					end if	
					response.write ";"					
					response.write rsfound("CarNo")
					response.write ";"					
					If CaseInDate<>"" then
 						response.write CaseInDate
					End If 
					response.write ";"					
					'應到案處所
					If trim(rsfound("BillUnitID"))<>"" Then
						strBUnit="select UnitName from UnitInfo where UnitID='"&Trim(trim(rsfound("MemberStation")))&"'"
						Set rsBunit=conn.execute(strBUnit)
						If Not rsBunit.eof Then
							response.write trim(rsBunit("UnitName"))
						End If
						rsBunit.close
						Set rsBunit=Nothing
					End If 

					'舉發類別
					response.write ";"					
					if trim(rsfound("BillTypeID"))="2" then
						response.write "逕舉"
					else
						response.write "攔停"
					end If

					response.write ";"					

					if trim(rsfound("BillFilldate"))<>"" and not isnull(rsfound("BillFilldate")) then
						response.write trim(Year(rsfound("BillFilldate"))-1911)&Right("00"&Month(rsfound("BillFilldate")),2)&Right("00"&day(rsfound("BillFilldate")),2)
					end if	
					response.write ";"					
					response.write rsfound("Rule1")
				
					response.write ";"					
					response.write rsfound("Rule2")

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