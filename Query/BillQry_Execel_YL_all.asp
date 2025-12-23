<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_"&Trim(request("date1"))&"違規日_舉發單資料.txt"

'Response.AddHeader "Content-Disposition", "filename="&fname
'response.contenttype="application/x-msexcel; charset=MS950" 
Response.AddHeader "Content-Disposition", "attachment;filename=" &fname
response.contenttype="application/vnd.ms-txt" 

Server.ScriptTimeout = 650000


	'檢查是否可進入本系統
	'AuthorityCheck(234)

	strSQL="select a.sn,a.billno,a.Carno,a.BillTypeID,a.IllegalDate,a.BillMemID1,a.BillMem1,a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.BillMemID4,a.BillMem4,a.Rule1,a.Rule2,a.IllegalAddress,a.MemberStation,a.BillFillDate,a.EquipMentID,a.BillUnitiD,a.DeallineDate,a.CarSimpleID,a.SignType,a.RecordStateID,a.Recorddate,a.RecordMemberID,a.note from Billbase a where a.Illegaldate between to_date('"&gOutDT(request("date1"))&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&gOutDT(request("date2"))&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
	" and BillStatus>'1' and billno is not null and RecordStateid=0 order by Illegaldate"
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)


%>
告發單號,違規日,入案日,填單日,應到案日
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
					
				
					'告發單號
					response.write rsfound("BillNo")
					%>,<%'違歸日期
					if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
						response.write trim(Year(rsfound("IllegalDate")))&"/"&Month(rsfound("IllegalDate"))&"/"&day(rsfound("IllegalDate"))
					end if	
					%>,<%
					strU="select DciCaseInDate from Billbasedcireturn where BillNo='"&Trim(rsfound("BillNO"))&"' and exchangetypeid='W'"
					Set rsU=conn.execute(strU)
					If Not rsU.eof Then
						If Trim(rsU("DciCaseInDate"))<>"" And Len(Trim(rsU("DciCaseInDate")))>5 then
							response.write gOutDT(rsU("DciCaseInDate"))
						End If 
					End If 
					rsU.close
					Set rsU=Nothing 
					%>,<%'填單日期
					if trim(rsfound("BillFilldate"))<>"" and not isnull(rsfound("BillFilldate")) then
						response.write trim(Year(rsfound("BillFilldate")))&"/"&Month(rsfound("BillFilldate"))&"/"&day(rsfound("BillFilldate"))
					end if	
					%>,<%'應到案日期
					if trim(rsfound("DeallineDate"))<>"" and not isnull(rsfound("DeallineDate")) then
						response.write trim(Year(rsfound("DeallineDate")))&"/"&Month(rsfound("DeallineDate"))&"/"&day(rsfound("DeallineDate"))
					end if	
					
				
				response.write vbCrLf
				rsfound.MoveNext
				Wend
				rsfound.close
				set rsfound=Nothing
				

	
	strSQL="select a.sn,a.billno,a.Carno,a.BillTypeID,a.IllegalDate,a.BillMemID1,a.BillMem1,a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.BillMemID4,a.BillMem4,a.Rule1,a.Rule2,a.IllegalAddress,a.MemberStation,a.BillFillDate,a.BillUnitiD,a.DeallineDate,a.CarSimpleID,a.SignType,a.RecordStateID,a.Recorddate,a.RecordMemberID,a.note,a.Driver,a.DriverID from Passerbase a where a.Illegaldate between to_date('"&gOutDT(request("date1"))&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&gOutDT(request("date2"))&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
	" and billno is not null and RecordStateid=0 order by Illegaldate"
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)

	If Not rsfound.Bof Then rsfound.MoveFirst 
	While Not rsfound.Eof
		Response.flush
	
			'告發單號
			response.write rsfound("BillNo")
			%>,<%'違歸日期
			if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
				response.write trim(Year(rsfound("IllegalDate")))&"/"&Month(rsfound("IllegalDate"))&"/"&day(rsfound("IllegalDate"))
			end if	
			%>,<%
			 
			%>,<%'填單日期
			if trim(rsfound("BillFilldate"))<>"" and not isnull(rsfound("BillFilldate")) then
				response.write trim(Year(rsfound("BillFilldate")))&"/"&Month(rsfound("BillFilldate"))&"/"&day(rsfound("BillFilldate"))
			end if	
			%>,<%'應到案日期
			if trim(rsfound("DeallineDate"))<>"" and not isnull(rsfound("DeallineDate")) then
				response.write trim(Year(rsfound("DeallineDate")))&"/"&Month(rsfound("DeallineDate"))&"/"&day(rsfound("DeallineDate"))
			end if	

			response.write vbCrLf
	rsfound.MoveNext
	Wend
	rsfound.close
	set rsfound=nothing

conn.close
set conn=nothing
%>