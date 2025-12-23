<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html onkeydown="KeyDown()">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--

.style4 {
	color: #FF0000;
	font-size: 16px
	}

-->
</style>
<title>舉發單入案</title>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<% Server.ScriptTimeout = 800 %>
<%

'入案
function funcBillToDCICaseIn(conn,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate,RecordMemberID,theBatchTime,sysCity)
	RecordDateTemp=year(RecordDate)&"/"&month(RecordDate)&"/"&day(RecordDate)&" "&hour(RecordDate)&":"&minute(RecordDate)&":"&second(RecordDate)

				'thirdNo=right(year(date)-1911,1)
				thirdNo="C"
				'response.write thirdNo
				if trim(Session("Credit_ID"))="T220933992" then
					UserStartNo="BD"&thirdNo&"0"
					UserSeq="bill0807BD0"
				elseif trim(Session("Credit_ID"))="E223625931" then
					UserStartNo="BD"&thirdNo&"3"
					UserSeq="bill0807BD3"
				elseif trim(Session("Credit_ID"))="E121011955" then
					UserStartNo="BD"&thirdNo&"5"
					UserSeq="bill0807BD5"
				elseif trim(Session("Credit_ID"))="E120003931" then
					UserStartNo="BB"&thirdNo&"0"
					UserSeq="bill0807BB0"
				elseif trim(Session("Credit_ID"))="T220359567" then
					UserStartNo="BB"&thirdNo&"3"
					UserSeq="bill0807BB3"
				elseif trim(Session("Credit_ID"))="E220912204" then
					UserStartNo="BB"&thirdNo&"5"
					UserSeq="bill0807BB5"				
				elseif trim(Session("Credit_ID"))="T220988040" then
					UserStartNo="BB"&thirdNo&"7"
					UserSeq="bill0807BB7"
				elseif trim(Session("Credit_ID"))="S220060233" then
					UserStartNo="BB"&thirdNo&"9"
					UserSeq="bill0807BB9"
				elseif trim(Session("Credit_ID"))="E222126171" then
					UserStartNo="BC"&thirdNo&"0"
					UserSeq="bill0807BC0"
				elseif trim(Session("Credit_ID"))="T121457177" then
					UserStartNo="BC"&thirdNo&"3"
					UserSeq="bill0807BC3"
				elseif trim(Session("Credit_ID"))="R120228634" Or trim(Session("Credit_ID"))="E221201933" then
					UserStartNo="BC"&thirdNo&"5"
					UserSeq="bill0807BC5"
				else
					if trim(Session("Unit_ID"))="0807" then
						sysUserUnit="0807"
					else
						strUserUnit="select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
						set rsUserUnit=conn.execute(strUserUnit)
						if not rsUserUnit.eof then
							sysUserUnit=trim(rsUserUnit("UnitTypeID"))
						end if
						rsUserUnit.close
						set rsUserUnit=nothing
					end if
					'抓 起始碼 跟 Seq Name
					strSeq="select * from GetBillNo where UnitID='"&sysUserUnit&"'"
					set rsSeq=conn.execute(strSeq)
					if not rsSeq.eof Then
						If Len(trim(rsSeq("BillStartVocab")))=3 then
							UserStartNo=trim(rsSeq("BillStartVocab"))
						Else
							UserStartNo=trim(rsSeq("BillStartVocab"))&thirdNo
						End If 
						UserSeq=trim(rsSeq("SeqNoName"))
					end if
					rsSeq.close
					set rsSeq=nothing	
				end if

			'抓單號最大值
'			if len(UserStartNo)=2 then
'				strSQLSn = "select LPAD("&UserSeq&".nextval,7,'0') as SN from Dual"
'			elseif len(UserStartNo)=3 then
'				strSQLSn = "select LPAD("&UserSeq&".nextval,6,'0') as SN from Dual"
'			elseif len(UserStartNo)=4 then
'				strSQLSn = "select LPAD("&UserSeq&".nextval,5,'0') as SN from Dual"
'			elseif len(UserStartNo)=5 then
'				strSQLSn = "select LPAD("&UserSeq&".nextval,4,'0') as SN from Dual"
'			end if
'			set rsBNo = Conn.execute(strSQLSn)
'			if not rsBNo.EOF then
'				sMaxSN = rsBNo("SN")
'			end if
'			rsBNo.close
'			set rsBNo = nothing

			NewBillNo=UserStartNo&sMaxSN

		response.write Session("Unit_ID")&"<br>"&UserSeq&"<br>"&NewBillNo
end function

RecordDate=split(gInitDT(date),"-")

'抓縣市
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing


		funcBillToDCICaseIn conn,000,"123456",2,"0000","8D00","2012/8/30 15:00",0,"101D0001",sys_City



%>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">

</table>

</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<%
conn.close
set conn=nothing
%>