<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"(彰化縣)違規日_舉發單資料.txt"
Response.AddHeader "Content-Disposition", "attachment;filename=" &fname
response.contenttype="application/vnd.ms-txt" 

Server.ScriptTimeout = 650000

%>
<%
	'檢查是否可進入本系統
	'AuthorityCheck(234)

	strSQL="select a.sn,a.billno,a.Carno,a.BillTypeID,a.IllegalDate,a.BillMemID1,a.BillMem1,a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.BillMemID4,a.BillMem4,a.Rule1,a.Rule2,a.IllegalAddress,a.MemberStation,a.BillFillDate,a.EquipMentID,a.BillUnitiD,a.DeallineDate,a.CarSimpleID,a.SignType,a.RecordStateID,a.Recorddate,a.RecordMemberID,a.note from Billbase a where a.illegaldate between to_date('"&gOutDT(request("date1"))&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&gOutDT(request("date2"))&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
	" and BillStatus>'1' and Recordstateid=0 order by illegaldate "
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)


%>告發類別,舉發單位,到案處所,到案日期,違規人證號,建檔日期,車號,舉發員警,入案日期,違反法條,告發單號,違規時間
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
						else
							strsql65="select OwnerID from Billbasedcireturn where " &_
								" carno='"&trim(rsfound("carno"))&"' and exchangetypeid='A'"
							set rs65=conn.execute(strsql65)
							if not rs65.eof then
								IllegalMemID=trim(rs65("OwnerID"))
							end if
							rs65.close
							set rs65=nothing
						end if
						OwnerName=trim(rs3("Owner"))
						OwnerAddress=trim(rs3("OwnerZip"))&" "&trim(rs3("OwnerAddress"))
						SecondAddress=trim(rs3("DriverHomeZip"))&" "&trim(rs3("DriverHomeAddress"))
						DciCarTypeID=trim(rs3("DciReturnCarType"))
					end if
					rs3.close
					set rs3=nothing

					'告發類別1
					if trim(rsfound("BillTypeID"))="2" then
						response.write "逕舉"
					else
						response.write "攔停"
					end if
					%>,<%'舉發單位2
					if trim(rsfound("BillUnitiD"))<>"" then
						strUid="select * from UnitInfo where UnitID='"&trim(rsfound("BillUnitiD"))&"'"
						set rsUid=conn.execute(strUid)
						if not rsUid.eof then
							response.write trim(rsUid("UnitName"))
						end if
						rsUid.close
						set rsUid=nothing
					end if
					%>,<%
					'到案處所3
					if trim(rsfound("BillTypeID"))="2" then
						StationName=DciReturnStation
					else
						StationName=trim(rsfound("MemberStation"))
					end if
					if StationName<>"" then
						strStation="select * from station where DciStationID='"&StationName&"'"
						set rsStation=conn.execute(strStation)
						if not rsStation.eof then
							response.write trim(rsStation("DciStationName"))
						end if
						rsStation.close
						set rsStation=nothing
					end if
					%>,<%'到案日期4
					if trim(rsfound("DeallineDate"))<>"" and not isnull(rsfound("DeallineDate")) then
						response.write trim(rsfound("DeallineDate"))
					end if	
					%>,<%'違規人證號5
					
					response.write IllegalMemID
					%>,<%'建檔日期6
					if trim(rsfound("Recorddate"))<>"" and not isnull(rsfound("Recorddate")) then
						response.write year(trim(rsfound("Recorddate")))&"/"&month(trim(rsfound("Recorddate")))&"/"&day(trim(rsfound("Recorddate")))
					end if	
					%>,<%'車號7
					response.write trim(rsfound("Carno"))
					
					%>,<%'舉發員警8
					response.write trim(rsfound("BillMem1"))
					if trim(rsfound("BillMem2"))<>"" then
					response.write "/"&trim(rsfound("BillMem2"))
					end if
					if trim(rsfound("BillMem3"))<>"" then
					response.write "/"& trim(rsfound("BillMem3"))
					end if
					if trim(rsfound("BillMem4"))<>"" then
					response.write "/"& trim(rsfound("BillMem4"))
					end if
					%>,<%'入案日期9
					response.write CaseInDate
					%>,<%'違反法條10
					response.write trim(rsfound("Rule1"))
					%>,<%'告發單號11
					response.write rsfound("BillNo")
					%>,<%'違歸時間12
					if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
						response.write year(trim(rsfound("IllegalDate")))&"/"&month(trim(rsfound("IllegalDate")))&"/"&day(trim(rsfound("IllegalDate")))
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