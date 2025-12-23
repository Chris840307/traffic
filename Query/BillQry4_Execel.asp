<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"違規日_逕舉_舉發單資料.txt"
Response.AddHeader "Content-Disposition", "attachment;filename=" &fname
response.contenttype="application/vnd.ms-txt" 

Server.ScriptTimeout = 650000
Response.flush
%>
<%
	'檢查是否可進入本系統
	'AuthorityCheck(234)

	strSQL="select a.sn,a.billno,a.Carno,a.BillTypeID,a.IllegalDate,a.BillMemID1,a.BillMem1,a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.BillMemID4,a.BillMem4,a.Rule1,a.Rule2,a.IllegalAddress,a.MemberStation,a.BillFillDate,a.EquipMentID,a.BillUnitiD,a.DeallineDate,a.CarSimpleID,a.SignType,a.RecordStateID,a.Recorddate,a.RecordMemberID,a.note from Billbase a where a.illegaldate between to_date('"&gOutDT(request("date1"))&" 0:0:0"&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&gOutDT(request("date2"))&" 23:59:59"&"','YYYY/MM/DD/HH24/MI/SS')" &_
	" and a.billtypeid = '2' "
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)


%>告發單號,到案處所,告發類別,舉發單狀態,入案日期,違歸時間,舉發員警,違反法條,違規路段,是否郵寄,郵寄日期,郵寄序號,簽收狀況,違規人證號,違規人姓名,違規人住址,車號,車主姓名,車主住址,填單日期,詳細車種,舉發單位,到案日期,簡式車種,建檔日期,操作人員,入案檔名,入案批號,入案狀態,上傳檔名,簽收/寄存批號,上傳狀態,簽收/寄存日期,簽收人,簽收/寄存原因,撤銷送達日期,寄存郵局,退件上傳檔名,退件批號,退件上傳狀態,退件郵寄日期,退件郵寄序號,退回日期,退件原因,二次郵寄地址,寄存送達上傳檔名,寄存送達批號,寄存送達狀態,寄存送達書號,寄存送達日,寄存送達生效(完成)日,寄存送達退件原因,寄存送達退回日期,公示送達上傳檔名,公示送達批號,公示送達上傳狀態,公示送達書號,公示送達生效日,發文監理站日期,公示送達原因,備註,法條二,舉發人代碼
<%  
					If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
					
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
					%>,<%
					'到案處所
					if trim(rsfound("BillTypeID"))="2" then
						StationName=DciReturnStation
					else
						StationName=trim(rsfound("MemberStation"))
					end if
					response.write StationName
					%>,<%'告發類別
					if trim(rsfound("BillTypeID"))="2" then
						response.write "逕舉"
					else
						response.write "攔停"
					end if
					%>,<%'舉發單狀態
					if trim(rsfound("RecordStateID"))="-1" then
						response.write "已刪除"
					else
						response.write "正常"
					end if
					%>,<%'入案日期
					response.write CaseInDate
					%>,<%'違歸時間
					if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
						response.write trim(rsfound("IllegalDate"))
					end if	
					%>,<%'舉發員警
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
					%>,<%'違反法條
					response.write trim(rsfound("Rule1"))
					%>,<%'違規路段
					response.write trim(rsfound("IllegalAddress"))
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
					%>,<%'郵寄序號
					response.write MailNumber
					%>,<%'簽收狀況
					if trim(rsfound("SignType"))<>"" and not isnull(rsfound("SignType")) then
						if rsfound("SignType")="A" then response.write "簽收"
						if rsfound("SignType")="U" then 
							strR2="select SignStateID from BillUserSignDate where billsn=" & trim(rsfound("sn"))
							set rsR2=conn.execute(strR2)
							if not rsR2.eof then 
								if rsR2("SignStateID")="2" then response.write "拒簽已收"
								if rsR2("SignStateID")="3" then response.write "已簽拒收"							
							else 
								response.write "拒簽收"
							end if
							rsR2.close
							set rsR2=nothing																
						end if				
					else
							strR2="select SignStateID from BillUserSignDate where billsn=" & trim(rsfound("sn"))
							set rsR2=conn.execute(strR2)
							if not rsR2.eof then 
								if rsR2("SignStateID")="5" then response.write "補開單"
							end if
							rsR2.close
							set rsR2=nothing															
					end if
					%>,<%'違規人證號
					
					response.write IllegalMemID
					%>,<%'違規人姓名
					response.write IllegalMem
					%>,<%'違規人住址
					response.write IllegalAddress
					%>,<%'車號
					response.write trim(rsfound("Carno"))
					
					%>,<%'車主姓名
					response.write OwnerName
					%>,<%'車主住址
					response.write OwnerAddress
					%>,<%'填單日期
					if trim(rsfound("BillFillDate"))<>"" and not isnull(rsfound("BillFillDate")) then
						response.write trim(rsfound("BillFillDate"))
					end if	
					%>,<%'詳細車種
'					if DciCarTypeID<>"" and not isnull(DciCarTypeID) then
'						strCarType="select Content from DciCode where TypeID=5 and ID='"&DciCarTypeID&"'"
'						set rsCarType=conn.execute(strCarType)
'						if not rsCarType.eof then
'							DciCarType=trim(rsCarType("Content"))
'						end if
'						rsCarType.close
'						set rsCarType=nothing
'					end if
					response.write DciCarTypeID
					%>,<%'舉發單位
					response.write trim(rsfound("BillUnitiD"))
					%>,<%'到案日期
					if trim(rsfound("DeallineDate"))<>"" and not isnull(rsfound("DeallineDate")) then
						response.write trim(rsfound("DeallineDate"))
					end if	
					%>,<%'簡式車種
					if trim(rsfound("CarSimpleID"))<>"" and not isnull(rsfound("CarSimpleID")) then
						if trim(rsfound("CarSimpleID"))="1" then
							response.write "汽車"
						elseif trim(rsfound("CarSimpleID"))="2" then
							response.write "拖車"
						elseif trim(rsfound("CarSimpleID"))="3" then
							response.write "重機"
						elseif trim(rsfound("CarSimpleID"))="4" then
							response.write "輕機"
						elseif trim(rsfound("CarSimpleID"))="6" then
							response.write "簡式車種"
						end if
					end if	
					%>,<%'建檔日期
					if trim(rsfound("Recorddate"))<>"" and not isnull(rsfound("Recorddate")) then
						response.write trim(rsfound("Recorddate"))
					end if	
					%>,<%'操作人員
					strRecMem="select ChName from MemberData where MemberID='"&trim(rsfound("RecordMemberID"))&"'"
					set rsRecMem=conn.execute(strRecMem)
					if not rsRecMem.eof then
						response.write trim(rsRecMem("ChName"))
					end if
					rsRecMem.close
					set rsRecMem=nothing
					%>,<%'入案檔名
					CaseInBatchnumber=""
					CaseInStatusID=""
					strW="select * from dcilog where billsn="&trim(rsfound("Sn"))&" and exchangeTypeid='W' order by exchangedate desc"
					set rsW=conn.execute(strW)
					if not rsW.eof then
						response.write rsW("FileName")
						CaseInBatchnumber=rsW("Batchnumber")
						CaseInStatusID=rsW("DcireturnstatusID")
					end if
					rsW.close
					set rsW=nothing

					%>,<%'入案批號
					response.write CaseInBatchnumber
					%>,<%'入案狀態
					response.write CaseInStatusID
					%>,<%'上傳檔名
					Batchnumber7=""
					StatusID7=""
				if trim(rsfound("EquipMentID"))<>"-1" then
					str7="select * from Dcilog where billsn="&trim(rsfound("Sn"))&" and exchangeTypeid='N' and ReturnmarkType='7' order by exchangedate desc"
					set rs7=conn.execute(str7)
					if not rs7.eof then
						response.write rs7("FileName")
						Batchnumber7=rs7("Batchnumber")
						StatusID7=rs7("DcireturnstatusID")
					end if
					rs7.close
					set rs7=nothing
				end if
					%>,<%'簽收/寄存批號
					response.write Batchnumber7
					%>,<%'上傳狀態
					response.write StatusID7
					%>,<%'簽收/寄存日期
					response.write SignDate
					%>,<%'簽收人
					response.write SignMem
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
					%>,<%'撤銷送達日期
				if trim(rsfound("EquipMentID"))<>"-1" then
					CancalSendDate=""   '撤銷送達日
					strCaseIn="select * from dcilog where " &_
						" BillSn=" & trim(rsfound("SN")) & " and BillNo='"&trim(rsfound("BillNo")) & "' and ExchangeTypeID='N' and ReturnMarkType='Y' and DCIRETURNSTATUSID='S' order by Exchangedate desc" 
					set rsCaseIn=conn.execute(strCaseIn)
					'response.write strCaseIn
					if not rsCaseIn.eof then
						CancalSendDate=trim(rsCaseIn("Exchangedate"))	
					end if
					rsCaseIn.close
					set rsCaseIn=nothing
					response.write CancalSendDate
				end if
					%>,<%'寄存郵局
					response.write MailStation
					%>,<%'退件上傳檔名
					BatchnumberN=""
					StatusIDN=""
				if trim(rsfound("EquipMentID"))<>"-1" then
					strN="select * from Dcilog where billsn="&trim(rsfound("Sn"))&" and exchangeTypeid='N' and ReturnmarkType='3' order by exchangedate desc"
					set rsN=conn.execute(strN)
					if not rsN.eof then
						response.write rsN("FileName")
						BatchnumberN=rsN("Batchnumber")
						StatusIDN=rsN("DcireturnstatusID")
					end if
					rsN.close
					set rsN=nothing
				end if
					%>,<%'退件批號
					response.write BatchnumberN
					%>,<%'退件上傳狀態
					response.write StatusIDN
					%>,<%'退件郵寄日期
					response.write ReturnSendMailDate
					%>,<%'退件郵寄序號
					response.write ReturnMailNumber
					%>,<%'退回日期
					response.write ReturnMailDate
					%>,<%'退件原因
					strReturnReason="select Content from DciCode where TypeID=7 and ID='"&trim(ReturnReason)&"'"
					set rsRR=conn.execute(strReturnReason)
					if not rsRR.eof then
						response.write trim(rsRR("Content"))
					end if
					rsRR.close			
					set rsRR=nothing
					%>,<%'二次郵寄地址
					Batchnumber4=""
					StatusID4=""
					FileName4=""
				if trim(rsfound("EquipMentID"))<>"-1" then
					strN="select * from Dcilog where billsn="&trim(rsfound("Sn"))&" and exchangeTypeid='N' and ReturnmarkType='4' order by exchangedate desc"
					set rsN=conn.execute(strN)
					if not rsN.eof then
						FileName4=rsN("FileName")
						Batchnumber4=rsN("Batchnumber")
						StatusID4=rsN("DcireturnstatusID")

						response.write OwnerName&"--"
						response.write SecondAddress
					end if
					rsN.close
					set rsN=nothing
				end if
					%>,<%'寄存送達上傳檔名
					response.write FileName4
					%>,<%'寄存送達批號
					response.write Batchnumber4
					%>,<%'寄存送達狀態
					response.write StatusID4
					%>,<%'寄存送達書號
					response.write StoreAndSendGovNumber
					%>,<%'寄存送達日
					response.write StoreAndSendEffectDate
					%>,<%'寄存送達生效(完成)日
					response.write StoreAndSendEndDate
					%>,<%'寄存送達退件原因
				if StoreAndSendReason<>"" then
					strSReason="select Content from DciCode where TypeID=7 and ID='"&trim(StoreAndSendReason)&"'"
					set rsSR=conn.execute(strSReason)
					if not rsSR.eof then
						response.write trim(rsSR("Content"))
					end if
					rsSR.close
					set rsSR=nothing
				end if
					%>,<%'寄存送達退回日期
					
					response.write StoreAndSendDate
					%>,<%'公示送達上傳檔名
					FileName5=""
					Batchnumber5=""
					StatusID5=""
				if trim(rsfound("EquipMentID"))<>"-1" then
					strN="select * from Dcilog where billsn="&trim(rsfound("Sn"))&" and exchangeTypeid='N' and ReturnmarkType='5' order by exchangedate desc"
					set rsN=conn.execute(strN)
					if not rsN.eof then
						FileName5=rsN("FileName")
						Batchnumber5=rsN("Batchnumber")
						StatusID5=rsN("DcireturnstatusID")
					end if
					rsN.close
					set rsN=nothing
					response.write FileName5
				end if
					%>,<%'公示送達批號
					response.write Batchnumber5
					%>,<%'公示送達上傳狀態
					response.write StatusID5
					%>,<%'公示送達書號
					response.write OpenGovGovNumber
					%>,<%'公示送達生效日
					response.write OpenGovEffectDate
					%>,<%'發文監理站日期
					response.write ReturnSendDate
					%>,<%'公示送達原因
				if ogReturnReason<>"" then
					strReturnReason="select Content from DciCode where TypeID=7 and ID='"&trim(ogReturnReason)&"'"
					set rsRR=conn.execute(strReturnReason)
					if not rsRR.eof then	
						response.write trim(rsRR("Content"))			
					end if				
					rsRR.close			
					set rsRR=nothing	
				end if
					%>,<%'備註
					response.write rsfound("note")
					%>,<%'法條二
					response.write rsfound("rule2")
					%>,<%'舉發人代碼
					if trim(rsfound("BillMemID1"))<>"" and not isnull(rsfound("BillMemID1")) then
						strM1="select loginid from memberdata where memberid="&trim(rsfound("BillMemID1"))
						set rsM1=conn.execute(strM1)
						if not rsM1.eof then
							response.write rsM1("loginid")
						end if
						rsM1.close 
						set rsM1=nothing
					end if
					if trim(rsfound("BillMemID2"))<>"" and not isnull(rsfound("BillMemID2")) then
						strM1="select loginid from memberdata where memberid="&trim(rsfound("BillMemID2"))
						set rsM1=conn.execute(strM1)
						if not rsM1.eof then
							response.write "/"&rsM1("loginid")
						end if
						rsM1.close 
						set rsM1=nothing
					end if
					if trim(rsfound("BillMemID3"))<>"" and not isnull(rsfound("BillMemID3")) then
						strM1="select loginid from memberdata where memberid="&trim(rsfound("BillMemID3"))
						set rsM1=conn.execute(strM1)
						if not rsM1.eof then
							response.write "/"&rsM1("loginid")
						end if
						rsM1.close 
						set rsM1=nothing
					end if
					if trim(rsfound("BillMemID4"))<>"" and not isnull(rsfound("BillMemID4")) then
						strM1="select loginid from memberdata where memberid="&trim(rsfound("BillMemID4"))
						set rsM1=conn.execute(strM1)
						if not rsM1.eof then
							response.write "/"&rsM1("loginid")
						end if
						rsM1.close 
						set rsM1=nothing
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