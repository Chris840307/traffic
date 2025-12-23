<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"(台南縣)違規日_逕舉_舉發單資料.txt"
Response.AddHeader "Content-Disposition", "attachment;filename=" &fname
response.contenttype="application/vnd.ms-txt" 

Server.ScriptTimeout = 650000

%>
<%
	'檢查是否可進入本系統
	'AuthorityCheck(234)

	strSQL="select a.sn,a.billno,a.Carno,a.BillTypeID,a.IllegalDate,a.BillMemID1,a.BillMem1,a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.BillMemID4,a.BillMem4,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.IllegalAddress,a.MemberStation,a.BillFillDate,a.EquipMentID,a.BillUnitiD,a.DeallineDate,a.CarSimpleID,a.SignType,a.RecordStateID,a.Recorddate,a.RecordMemberID,a.note from Billbase a where a.illegaldate between to_date('"&gOutDT(request("date1"))&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&gOutDT(request("date2"))&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
	" and a.billtypeid='2' and BillStatus>'1' and Recordstateid=0 and billno is not null order by illegaldate"
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)


%>告發單號,告發類別,入案日期,違規時間,法條一,法條二,法條三,是否郵寄,郵寄日期,違規人姓名,車主姓名,舉發單位,建檔日期,簽收/寄存上傳狀態,簽收/寄存日期,簽收/寄存原因
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
						end if
						OwnerName=trim(rs3("Owner"))
						OwnerAddress=trim(rs3("OwnerZip"))&" "&trim(rs3("OwnerAddress"))
						SecondAddress=trim(rs3("DriverHomeZip"))&" "&trim(rs3("DriverHomeAddress"))
						DciCarTypeID=trim(rs3("DciReturnCarType"))
					end if
					rs3.close
					set rs3=nothing
					'告發單號
					response.write rsfound("BillNo")
					%>,<%'告發類別
					if trim(rsfound("BillTypeID"))="2" then
						response.write "逕舉"
					else
						response.write "攔停"
					end if
					%>,<%'入案日期
					response.write CaseInDate
					%>,<%'違歸時間
					if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
						response.write trim(rsfound("IllegalDate"))
					end if	
					%>,<%'違反法條一
					response.write trim(rsfound("Rule1"))
					%>,<%'法條二
					response.write rsfound("rule2")
					%>,<%'法條3
					response.write rsfound("rule3")
					%>,<%'是否郵寄
					if trim(rsfound("EquipMentID"))<>"-1" then
						response.write "是"
					else
						response.write "否"
					end if
					%>,<%'郵寄日期
					MailDate=""
					MailNumber=""
					SignDate=""
					SignMem=""
					SignReson=""
					MailStation=""
					ReturnSendMailDate=""
					ReturnMailNumber=""
					ReturnMailDate=""
					ReturnReason=""
					StoreAndSendGovNumber=""
					StoreAndSendEffectDate=""
					StoreAndSendEndDate=""
					StoreAndSendReason=""
					StoreAndSendDate=""
					OpenGovGovNumber=""
					OpenGovEffectDate=""
					ReturnSendDate=""
					ogReturnReason=""
					if trim(rsfound("EquipMentID"))<>"-1" then
						strMail="select * from billmailhistory where billsn="&trim(rsfound("Sn"))
						set rsM=conn.execute(strMail)
						if not rsM.eof then
							if trim(rsM("MailDate"))<>"" and not isnull(rsM("MailDate")) then
								MailDate=trim(rsM("MailDate"))
							end if
							MailNumber=trim(rsM("MailNumber"))
							SignDate=trim(rsM("SignDate"))
							SignMem=trim(rsM("SignMan"))
							SignReson=trim(rsM("SignResonID"))
							MailStation=(rsM("MailStation"))
							ReturnSendMailDate=trim(rsM("StoreAndSendSendDate"))
							ReturnMailNumber=trim(rsM("StoreAndSendMailNumber"))
							ReturnMailDate=trim(rsM("MAILRETURNDATE"))
							ReturnReason=trim(rsM("RETURNRESONID"))
							StoreAndSendGovNumber=trim(rsM("STOREANDSENDGOVNUMBER"))
							StoreAndSendEffectDate=trim(rsM("STOREANDSENDEFFECTDATE"))
							StoreAndSendEndDate=trim(rsM("StoreAndSendMailDate"))
							StoreAndSendReason=trim(rsM("STOREANDSENDRETURNRESONID"))
							StoreAndSendDate=trim(rsM("STOREANDSENDMAILRETURNDATE"))
							OpenGovGovNumber=trim(rsM("OPENGOVNUMBER"))
							OpenGovEffectDate=trim(rsM("OPENGOVDATE"))
							ReturnSendDate=trim(rsM("SendOpenGovDocToStationDate"))
							ogReturnReason=trim(rsM("OPENGOVRESONID"))
						end if
						rsM.close
						set rsM=nothing
						response.write MailDate
					end if
					%>,<%'違規人姓名
					response.write IllegalMem
					%>,<%'車主姓名
					response.write OwnerName
					%>,<%'舉發單位
					response.write trim(rsfound("BillUnitiD"))
					%>,<%'建檔日期
					if trim(rsfound("Recorddate"))<>"" and not isnull(rsfound("Recorddate")) then
						response.write trim(rsfound("Recorddate"))
					end if	
					%>,<%'簽收/寄存上傳狀態
					response.write StatusID7
					%>,<%'簽收/寄存日期
					response.write SignDate
					%>,<%'簽收/寄存原因
					if trim(SignReson)<>"" and not isnull(SignReson) then
						strReturnReason="select Content from DciCode where TypeID=7 and ID='"&trim(SignReson)&"'"
						set rsRR=conn.execute(strReturnReason)
						if not rsRR.eof then
							response.write trim(rsRR("Content"))
						end if
						rsRR.close
						set rsRR=nothing
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