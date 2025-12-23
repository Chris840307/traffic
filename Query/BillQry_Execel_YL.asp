<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"舉發單資料.txt"
Response.AddHeader "Content-Disposition", "attachment;filename=" &fname
response.contenttype="application/vnd.ms-txt" 

Server.ScriptTimeout = 650000

%>
<%
	'檢查是否可進入本系統
	'AuthorityCheck(234)

	strSQL="select a.sn,a.billno,a.Carno,a.BillTypeID,a.IllegalDate,a.BillMemID1,a.BillMem1,a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.BillMemID4,a.BillMem4,a.Rule1,a.Rule2,a.IllegalAddress,a.MemberStation,a.BillFillDate,a.EquipMentID,a.BillUnitiD,a.DeallineDate,a.CarSimpleID,a.SignType,a.RecordStateID,a.Recorddate,a.RecordMemberID,a.note from Billbase a where a.illegaldate between to_date('"&gOutDT(request("date1"))&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&gOutDT(request("date2"))&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
	" and Recordstateid=0 and billno is not null order by illegaldate "
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)


%>單號,入案日期,監理站接收日期,入案狀態,結案狀態 
<%  
					If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
					Response.flush

					DciCaseIn=0
					DciReturnStation=""
					CaseInDate=""
					IllegalMemID=""
					IllegalMem=""
					IllegalAddress=""
					OwnerName=""
					OwnerAddress=""
					DciCarTypeID=""
					SecondAddress=""
					DciStatus=""
					strsql3="select * from Billbasedcireturn where billno='"&trim(rsfound("Billno"))&"' " &_
						" and carno='"&trim(rsfound("carno"))&"' and exchangetypeid='W'"
					set rs3=conn.execute(strsql3)
					if not rs3.eof then
						DciCaseIn=1
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
						DciStatus=trim(rs3("Status"))
						BILLCLOSEIDtmp=trim(rs3("BILLCLOSEID"))
					end if
					rs3.close
					set rs3=nothing

					
					'單號
					response.write trim(rsfound("BillNo"))
					
					%>,<%'入案日
					response.write CaseInDate
					
					%>,<%'入案日
					response.write CaseInDate
					
					%>,<%'入案狀態
					strSt="select * from dcireturnstatus where DciActionID='W' and Dcireturn='"&DciStatus&"'"
					set rsSt=conn.execute(strSt)
					if not rsSt.eof then
						response.write trim(rsSt("StatusContent"))
					end if
					rsSt.close
					set rsSt=nothing
					%>,<%'結案狀態
					BILLCLOSEIDflag=""
					strClose="select * from billbasedcireturn where billno='"&trim(rsfound("BillNo"))&"'" &_
						" and CarNo='"&trim(rsfound("CarNo"))&"' and exchangetypeid='N'" 
					Set rsClose=conn.execute(strClose)
					If Not rsClose.eof Then
						If trim(rsClose("BILLCLOSEID"))="" Or isnull(rsClose("BILLCLOSEID")) Then
							BILLCLOSEIDflag=BILLCLOSEIDtmp
						else
							BILLCLOSEIDflag=trim(rsClose("BILLCLOSEID"))
						End if
					Else
						BILLCLOSEIDflag=BILLCLOSEIDtmp
					End If
					rsClose.close
					Set rsClose=nothing
					strClose2="select * from DciCode where TypeID=9 and ID='"&BILLCLOSEIDflag&"'"
					set rsClose2=conn.execute(strClose2)
					if not rsClose2.eof then
						response.write trim(rsClose2("Content"))
					Else
						strClose3="select * from DciCode where TypeID=9 and ID='"&BILLCLOSEIDtmp&"'"
						set rsClose3=conn.execute(strClose3)
						if not rsClose3.eof then
							response.write trim(rsClose3("Content"))
						end if
						rsClose3.close
						set rsClose3=nothing		
					end if
					rsClose2.close
					set rsClose2=nothing										

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