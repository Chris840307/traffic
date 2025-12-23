<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"填單日_舉發單資料.txt"
Response.AddHeader "Content-Disposition", "attachment;filename=" &fname
response.contenttype="application/vnd.ms-txt" 

Server.ScriptTimeout = 650000

%>
<%
	'檢查是否可進入本系統
	'AuthorityCheck(234)

	strSQL="select a.sn,a.billno,a.Carno,a.BillTypeID,a.IllegalDate,a.BillMemID1,a.BillMem1,a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.BillMemID4,a.BillMem4,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.DriverID,a.IllegalAddress,a.MemberStation,a.BillFillDate,a.EquipMentID,a.BillUnitiD,a.DeallineDate,a.CarSimpleID,a.SignType,a.RecordStateID,a.Recorddate,a.RecordMemberID,a.note from Billbase a where a.IllegalDate between to_date('"&gOutDT(request("date1"))&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&gOutDT(request("date2"))&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
	" and BillStatus>'1' and Recordstateid=0 and billno is not null order by IllegalDate"
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)


%>舉發類型,舉發單位,到案地點,應到案日期,證號,資料建檔日期,車號,舉發員警,入案日期,違規法條一,單號,違規日期,違規時間,填單日期
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
					strsql3="select * from Billbasedcireturn where billno='"&trim(rsfound("Billno"))&"' " &_
						" and carno='"&trim(rsfound("carno"))&"' and exchangetypeid='W'"
					set rs3=conn.execute(strsql3)
					if not rs3.eof then
						DciReturnStation=trim(rs3("DciReturnStation"))
						CaseInDate=trim(rs3("DciCaseInDate"))
						if trim(rsfound("BillTypeID"))="1" then
							IllegalMemID=trim(rs3("DriverID"))
							IllegalMem=trim(rs3("Driver"))
							IllegalAddress=trim(rs3("DriverHomeZip"))&" "&trim(rs3("DriverHomeAddress"))
						end if
						OwnerName=trim(rs3("Owner"))
						OwnerAddress=trim(rs3("OwnerZip"))&" "&trim(rs3("OwnerAddress"))
						SecondAddress=trim(rs3("DriverHomeZip"))&" "&trim(rs3("DriverHomeAddress"))
						DciCarTypeID=trim(rs3("DciReturnCarType"))
						Rule1=trim(rs3("Rule1"))
						Rule2=trim(rs3("Rule2"))
						ForFeit1=trim(rs3("ForFeit1"))
						ForFeit2=trim(rs3("ForFeit2"))
					end if
					rs3.close
					set rs3=nothing
					
					'告發類別
					if trim(rsfound("BillTypeID"))="2" then
						response.write "逕舉"
					else
						response.write "攔停"
					end if
					%>,<%'舉發單位
					If trim(rsfound("BillUnitID"))<>"" Then
						strBUnit="select * from UnitInfo where UnitID='"&Trim(trim(rsfound("BillUnitID")))&"'"
						Set rsBunit=conn.execute(strBUnit)
						If Not rsBunit.eof Then
							response.write trim(rsBunit("UnitName"))
						End If
						rsBunit.close
						Set rsBunit=Nothing
					End If 
					%>,<%'應到案處所
					If DciReturnStation<>"" then
 						strDRS="select * from station where DciStationID='"&DciReturnStation&"'"
						Set rsDRS=conn.execute(strDRS)
						If Not rsDRS.eof Then
							response.write Trim(rsDRS("DciStationName"))
						End If
						rsDRS.close
						Set rsDRS=nothing
					End If 
					%>,<%'應到案日期
					if trim(rsfound("DeallineDate"))<>"" and not isnull(rsfound("DeallineDate")) then
						response.write trim(Year(rsfound("DeallineDate"))-1911)&Right("00"&Month(rsfound("DeallineDate")),2)&Right("00"&day(rsfound("DeallineDate")),2)
					end if	
					%>,<%'駕駛人id
					response.write rsfound("DriverID")
					%>,<%'建檔日期
					if trim(rsfound("Recorddate"))<>"" and not isnull(rsfound("Recorddate")) then
						response.write trim(Year(rsfound("Recorddate"))-1911)&Right("00"&Month(rsfound("Recorddate")),2)&Right("00"&day(rsfound("Recorddate")),2)
					end if	
					%>,<%'車號
					response.write rsfound("CarNo")
					%>,<%'舉發員警
					response.write trim(rsfound("BillMem1"))
					%>,<%'入案日期
					If CaseInDate<>"" then
 						response.write CaseInDate
					End If 
					%>,<%'法條一
					response.write Rule1
					%>,<%'舉發單號
					response.write rsfound("Billno")
					%>,<%'違歸日期
					if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
						response.write trim(Year(rsfound("IllegalDate"))-1911)&Right("00"&Month(rsfound("IllegalDate")),2)&Right("00"&day(rsfound("IllegalDate")),2)
					end if	
					%>,<%'違歸時間
					if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
						response.write Right("00"&hour(rsfound("IllegalDate")),2)&Right("00"&minute(rsfound("IllegalDate")),2)
					end if	
					%>,<%'填單日期
					if trim(rsfound("BillFilldate"))<>"" and not isnull(rsfound("BillFilldate")) then
						response.write trim(Year(rsfound("BillFilldate"))-1911)&Right("00"&Month(rsfound("BillFilldate")),2)&Right("00"&day(rsfound("BillFilldate")),2)
					end if	
					
				

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