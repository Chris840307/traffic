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

	strSQL="select a.sn,a.billno,a.Carno,a.BillTypeID,a.IllegalDate,a.BillMemID1,a.BillMem1,a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.BillMemID4,a.BillMem4,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.DriverID,a.IllegalAddress,a.MemberStation,a.BillFillDate,a.EquipMentID,a.BillUnitiD,a.DeallineDate,a.CarSimpleID,a.SignType,a.RecordStateID,a.Recorddate,a.RecordMemberID,a.note from Billbase a where a.illegaldate between to_date('"&gOutDT(request("date1"))&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&gOutDT(request("date2"))&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
	" and BillStatus>'1' and Recordstateid=0 and billno is not null order by BillFillDate"
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)


%>類別,違規日期,違規時間,舉發單號,車號,簡式車種,駕駛人ID,駕駛人姓名,車主姓名,詳細車種,舉發員警,舉發單位,違規地點,法條一,法條二,罰款一,罰款二,填單日期,應到案日期,應到案處所,建檔日期,入案日期,代保管物件,操作人員
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
					%>,<%'違歸日期
					if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
						response.write trim(Year(rsfound("IllegalDate"))-1911)&Right("00"&Month(rsfound("IllegalDate")),2)&Right("00"&day(rsfound("IllegalDate")),2)
					end if	
					%>,<%'違歸時間
					if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
						response.write Right("00"&hour(rsfound("IllegalDate")),2)&Right("00"&minute(rsfound("IllegalDate")),2)
					end if	
					%>,<%'舉發單號
					response.write rsfound("Billno")
					%>,<%'車號
					response.write rsfound("CarNo")
					%>,<%'簡式車種
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
					End if
					%>,<%'駕駛人id
					response.write rsfound("DriverID")
					%>,<%'駕駛人姓名
					response.write IllegalMem
					%>,<%'車主姓名
					response.write OwnerName
					%>,<%'祥細車種
					If DciCarTypeID<>"" Then
						strDCT="select * from DciCode where TypeID='5' and ID='"&DciCarTypeID&"'"
						Set rsDCT=conn.execute(strDCT)
						If Not rsDCT.eof Then
							response.write Trim(rsDCT("Content"))
						End If
						rsDCT.close
						Set rsDCT=nothing
					End if
					%>,<%'舉發員警
					response.write trim(rsfound("BillMem1"))
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
					%>,<%'違規地點
					If trim(rsfound("IllegalAddress"))<>"" Then
						response.write trim(rsfound("IllegalAddress"))
					End If 
					%>,<%'法條一
					response.write Rule1
					%>,<%'法條二
					If Rule2<>"0" And Rule2<>"" then
						response.write Rule2
					End If 
					%>,<%'罰款一
					response.write ForFeit1
					%>,<%'罰款二
					If ForFeit2<>"0" And ForFeit2<>"" then
 						response.write ForFeit2
					End If 
					%>,<%'填單日期
					if trim(rsfound("BillFilldate"))<>"" and not isnull(rsfound("BillFilldate")) then
						response.write trim(Year(rsfound("BillFilldate"))-1911)&Right("00"&Month(rsfound("BillFilldate")),2)&Right("00"&day(rsfound("BillFilldate")),2)
					end if	
					%>,<%'應到案日期
					if trim(rsfound("DeallineDate"))<>"" and not isnull(rsfound("DeallineDate")) then
						response.write trim(Year(rsfound("DeallineDate"))-1911)&Right("00"&Month(rsfound("DeallineDate")),2)&Right("00"&day(rsfound("DeallineDate")),2)
					end if	
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
					%>,<%'建檔日期
					if trim(rsfound("Recorddate"))<>"" and not isnull(rsfound("Recorddate")) then
						response.write trim(Year(rsfound("Recorddate"))-1911)&Right("00"&Month(rsfound("Recorddate")),2)&Right("00"&day(rsfound("Recorddate")),2)
					end if	
					%>,<%'入案日期
					If CaseInDate<>"" then
 						response.write CaseInDate
					End If 
					%>,<%'代保管物
					BillFastenerDetail=""
					strBFD="select * from BillFastenerDetail where BillSn="&trim(rsfound("Sn"))
					Set rsBFD=conn.execute(strBFD)
					If Not rsBFD.Bof Then rsBFD.MoveFirst 
					While Not rsBFD.Eof
						If BillFastenerDetail<>"" Then
							BillFastenerDetail=BillFastenerDetail&"、"
						End If
						If Trim(rsBFD("FastenerTypeID"))<>"" And Not IsNull(rsBFD("FastenerTypeID")) Then
							strFT="select * from DciCode where TypeID='6' and ID='"&Trim(rsBFD("FastenerTypeID"))&"'"
							Set rsFT=conn.execute(strFT)
							If Not rsFT.eof Then
								BillFastenerDetail=BillFastenerDetail&Trim(rsFT("Content"))
							End If
							rsFT.close
							Set rsFT=nothing
						End If
						
						rsBFD.MoveNext
					Wend
					rsBFD.close
					set rsBFD=Nothing
					response.write BillFastenerDetail
					%>,<%'操作人員
					if trim(rsfound("RecordMemberID"))<>"" and not isnull(rsfound("RecordMemberID")) Then
						strRMem="select * from Memberdata where memberid="&trim(rsfound("RecordMemberID"))
						Set rsRMem=conn.execute(strRMem)
						If Not rsRMem.eof Then
							response.write Trim(rsRMem("ChName"))
						End If
						rsRMem.close
						Set rsRMem=Nothing 
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