<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"違規日_舉發單資料.txt"
Response.AddHeader "Content-Disposition", "attachment;filename=" &fname
response.contenttype="application/vnd.ms-txt" 

Server.ScriptTimeout = 650000

%>
<%
	'檢查是否可進入本系統
	'AuthorityCheck(234)

	strSQL="select a.sn,a.billno,a.Carno,a.BillTypeID,a.IllegalDate,a.BillMemID1,a.BillMem1,a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.BillMemID4,a.BillMem4,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.DriverID,a.IllegalAddress,a.MemberStation,a.BillFillDate,a.EquipMentID,a.BillUnitiD,a.DeallineDate,a.CarSimpleID,a.SignType,a.RecordStateID,a.Recorddate,a.RecordMemberID,a.note from Billbase a where a.Recorddate between to_date('"&gOutDT(request("date1"))&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&gOutDT(request("date2"))&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
	" and BillStatus>'1' and Recordstateid=0 and billno is not null order by Recorddate"
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)


%>簡式車種,違規事實一,違規事實二,違規日期,違規時間,違規地點
<%  
					If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
					Response.flush

					
					'簡式車種
					If rsfound("CarSimpleID")="1" Then
						response.write "汽車"
					ElseIf rsfound("CarSimpleID")="2" Then
						response.write "拖車"
					ElseIf rsfound("CarSimpleID")="3" Then
						response.write "重機"
					ElseIf rsfound("CarSimpleID")="4" Then
						response.write "輕機"
					ElseIf rsfound("CarSimpleID")="6" Then
						response.write "臨時車牌"
					End If
					%>,<%'違規事實一
					If trim(rsfound("Rule1"))<>"" Then
						strBUnit="select * from law where itemid='"&Trim(trim(rsfound("Rule1")))&"' and version='2'"
						Set rsBunit=conn.execute(strBUnit)
						If Not rsBunit.eof Then
							response.write trim(rsBunit("illegalrule"))
						End If
						rsBunit.close
						Set rsBunit=Nothing
					End If 
					%>,<%'違規事實二
					If trim(rsfound("Rule2"))<>"" Then
						strBUnit="select * from law where itemid='"&Trim(trim(rsfound("Rule2")))&"' and version='2'"
						Set rsBunit=conn.execute(strBUnit)
						If Not rsBunit.eof Then
							response.write trim(rsBunit("illegalrule"))
						End If
						rsBunit.close
						Set rsBunit=Nothing
					End If 
					%>,<%'違歸時間
					if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
						response.write trim(Year(rsfound("IllegalDate"))-1911)&"/"&Right("00"&Month(rsfound("IllegalDate")),2)&"/"&Right("00"&day(rsfound("IllegalDate")),2)
						response.write " " & Right("00"&hour(rsfound("IllegalDate")),2)&":"&Right("00"&minute(rsfound("IllegalDate")),2)
					end if	
					%>,<%'違規地點
					If trim(rsfound("IllegalAddress"))<>"" Then
						response.write trim(rsfound("IllegalAddress"))
					End If 
				

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