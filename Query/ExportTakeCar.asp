<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%

loginid=Replace(request("loginid"),",","','")
RecordMemberID=""
		strSQL="select memberid from memberdata where loginid in ('"&loginid&"') and recordstateid=0 and accountstateid=0"

	set rsfound=conn.execute(strSQL)
	While not rsfound.eof
		If RecordMemberID<>"" then
			RecordMemberID=RecordMemberID & "','" & rsfound("memberid")
		Else
			RecordMemberID=rsfound("memberid")
		End if
	  rsfound.MoveNext	
	wend
	RecordMemberID="'" & RecordMemberID & "'"


	fMnoth=month(now)
	if fMnoth<10 then fMnoth="0"&fMnoth
	fDay=day(now)
	if fDay<10 then	fDay="0"&fDay


	sYear = year(now)-1911
If Session("Unit_ID")="0861" Then 
	tmpNum="1"
ElseIf Session("Unit_ID")="0862" Then 
	tmpNum="2"
ElseIf Session("Unit_ID")="0863" Then 
	tmpNum="3"
ElseIf Session("Unit_ID")="0864" Then 
	tmpNum="4"
ElseIf Session("Unit_ID")="0871" Then 
	tmpNum="5"
ElseIf Session("Unit_ID")="0872" Then 
	tmpNum="6"
ElseIf Session("Unit_ID")="0873" Then 
	tmpNum="7"
ElseIf Session("Unit_ID")="0874" Then 
	tmpNum="8"
End if



'	fname= sYear & fMnoth & fDay & ".dat"
	fname= "BB" & sYear & fMnoth & fDay & "1." & tmpNum & ".T"

	Response.Buffer = true
	Response.AddHeader  "Content-Disposition","attachment;filename=" & fname    
	Response.CharSet  =  "MS950"    
	Response.ContentType = "application/vnd.ms-txt"
	Server.ScriptTimeout = 800
	Response.flush
'	RecordMemberID="819"

	strwhere=" and f.RecordDate between TO_DATE('"&gOutDT(request("StartDate"))&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&gOutDT(request("EndDate"))&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"

	strwhere=" and f.RecordDate between TO_DATE('"&gOutDT(request("StartDate"))&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&gOutDT(request("EndDate"))&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"

	If RecordMemberID<>"" Then strwhere=strwhere & " and f.RecordMemberID in ("&RecordMemberID&")"

		
		strSQL="select count(*) as cnt  from BillBase f where f.BillStatus=9 and f.RecordStateID=0 " & strwhere
		


	set rsfound=conn.execute(strSQL)
	cnt=rsfound("cnt")

		strSQL="select f.SN,f.BillNo,f.CarNo,f.CarSimpleID,g.Loginid,f.IllegalDate,f.IllegalAddress,f.RecordDate,f.ForFeit1,f.Rule1,f.Rule2,f.Rule3,f.Rule4,f.BillUnitID,f.DealLineDate" &_
			",f.DriverID,h.Loginid as Loginid2 from BillBase f ,memberdata g,memberdata h where f.BillStatus=9 and f.RecordStateID=0 " & strwhere & " and f.Billmemid1=h.memberid and f.RecordMemberID=g.memberid and g.recordstateid=0 and g.accountstateid=0  order by f.RecordMemberID,f.RecordDate"
	set rsfound=conn.execute(strSQL)


'00001 8827-JQ  981109 2048       981124 B04191912                   000 5610101  0        0        0        0864 D 3   B 00 X 0 BD0017                                                                                                                                                                                                               0    W80015 30           旺盛街               900    0      0      0      
	
Function GetSpace(tfields,length)

tmpInt=0
  For i=1 To Len(tfields)
	If asc(Mid(tfields,i,1))>127 or asc(Mid(tfields,i,1))<0 Then 
		tmpInt=tmpInt+2
	Else
		tmpInt=tmpInt+1
	End If
  next

temp=""
 For i=1 To CDbl(length)-cdbl(tmpInt)
	temp=temp&" "
 Next 
GetSpace=tfields&temp

End function
					If Not rsfound.Bof Then rsfound.MoveFirst 
					PrintSN=0

					While Not rsfound.Eof
					 	
						PrintSN=PrintSN+1
						tmp0=""
						For a=1 To 5-Len(PrintSN) 
							  tmp0=tmp0&"0"
						Next
						'1 流水號
						response.write GetSpace(tmp0&PrintSN,6)
						'2 車號
						response.write GetSpace(rsfound("CarNo"),9)
						'3 違規日
						'response.write GetSpace(ginitdt(rsfound("IllegalDate")),7)
						 response.write GetSpace(Right("0"&ginitdt(rsfound("IllegalDate")),7),8)
						'4 違規時間
						response.write GetSpace(Right("0"&Hour(rsfound("IllegalDate")),2)&Right("0"&minute(rsfound("IllegalDate")),2),5)
						'5 違規地點代碼 放空白
						response.write GetSpace("",6)
						'6 應到案日 DealLineDate
'						response.write GetSpace(ginitdt(rsfound("DealLineDate")),7)
						 response.write GetSpace(Right("0"&ginitdt(rsfound("DealLineDate")),7),8)
						'7 違規單號
						response.write GetSpace(rsfound("Billno"),10)
						'8 駕駛人證號
						response.write GetSpace("",11)
						'9 駕駛人生日
						response.write GetSpace("",8)
						'10 代管物件代碼'  
						response.write GetSpace("000",4)
						'11 款條1 (法條)
						response.write GetSpace(rsfound("Rule1"),9)
						'12 款條2 (法條)
						response.write GetSpace("0",9)
						'13 款條3 (法條)
						response.write GetSpace("0",9)
						'14 款條4 (法條)
						response.write GetSpace("0",9)
						'15 告發單位代碼 BillUnitID
						response.write GetSpace(rsfound("BillUnitID"),5)					
						'16 告發類型
						response.write GetSpace("D",2)
						'17 簡式車種
						response.write GetSpace(rsfound("CarSimpleID"),4)
						'18 異動別 固定放 J
						response.write GetSpace("J",2)
						'19 櫃別 固定放 00
						response.write GetSpace("32",3)
						'20 車籍狀態 固定放 X
						response.write GetSpace("X",2)
						'21 駕籍狀態 固定放 0
						response.write GetSpace("0",2)
						'22 建檔人員臂章號碼 Loginid
						response.write GetSpace(rsfound("Loginid"),7)
						'23 車主姓名
						response.write GetSpace("",31)
						'24 車主地址
						response.write GetSpace("",81)
						'25 駕駛人姓名
						response.write GetSpace("",13)
						'26 駕駛人地址
						response.write GetSpace("",81)
						'27 顏色代碼
						response.write GetSpace("0",5)
						'28 員警代碼 
						response.write GetSpace(Right("0"&ginitdt(rsfound("RecordDate")),7),8)
						'29 所站代號
						response.write GetSpace("30",3)
						'30 是否有保險證
						response.write GetSpace("0",2)
						'31 車主郵遞區號
						response.write GetSpace("",4)
						'32 駕駛人郵遞區號
						response.write GetSpace("",4)
						'33 違規中文地點
						response.write GetSpace(Mid(rsfound("IllegalAddress")&"",1,9),19)
						'34 違反牌照稅註記
						response.write GetSpace("",2)
						'35 條款1 金額 ForFeit1
						response.write GetSpace(rsfound("ForFeit1"),7)
						'36 條款2 金額 0
						response.write GetSpace("0",7)
						'37 條款3 金額 0
						response.write GetSpace("0",7)
						'38 條款4 金額 0
						response.write GetSpace("0",7)


					
							
												
						If CDbl(PrintSN)<>CDbl(cnt) Then response.write(vbCrLf)					
					rsfound.MoveNext

					Wend

					rsfound.close
					set rsfound=nothing
				
conn.close
set conn=Nothing
%>