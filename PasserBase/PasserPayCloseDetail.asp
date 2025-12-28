<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">
<%

strSQL="select UnitName from Unitinfo where Unitid=(select UnitTypeID from Unitinfo where Unitid='"&Session("Unit_ID")&"')"

set rsUnitName=conn.execute(strSQL)
if not rsUnitName.eof then
	TitleUnitName2=trim(rsUnitName("UnitName"))
end if
rsUnitName.close
set rsUnitName=nothing

strUnitName="select Value from ApConfigure where ID=40"
set rsUnitName=conn.execute(strUnitName)
if not rsUnitName.eof then
	TitleUnitName=trim(rsUnitName("value"))&" "&TitleUnitName2
end if
rsUnitName.close
set rsUnitName=nothing

	nowdate=""

	if request("RecordDate1")<>"" and request("RecordDate2")<>""then
		ArgueDate1=gOutDT(request("RecordDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("RecordDate2"))&" 23:59:59"

		chk_year=year(ArgueDate2)-1911-6
		now_year=year(ArgueDate2)-1911
		now_month=month(ArgueDate2)

		nowdate=ArgueDate2
	end if

	if request("BillFillDate1")<>"" and request("BillFillDate2")<>""then
		ArgueDate1=gOutDT(request("BillFillDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("BillFillDate2"))&" 23:59:59"
		
		chk_year=year(ArgueDate2)-1911-6
		now_year=year(ArgueDate2)-1911
		now_month=month(ArgueDate2)

		nowdate=ArgueDate2
	end if

	if request("IllegalDate1")<>"" and request("IllegalDate2")<>""then
		ArgueDate1=gOutDT(request("IllegalDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("IllegalDate2"))&" 23:59:59"
		
		chk_year=year(ArgueDate2)-1911-6
		now_year=year(ArgueDate2)-1911
		now_month=month(ArgueDate2)

		nowdate=ArgueDate2
	end if

	if request("UrgeDate1")<>"" and request("UrgeDate2")<>""then
		ArgueDate1=gOutDT(request("UrgeDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("UrgeDate2"))&" 23:59:59"

		chk_year=year(ArgueDate2)-1911-6
		now_year=year(ArgueDate2)-1911
		now_month=month(ArgueDate2)

		nowdate=ArgueDate2
	end if

	if request("PayDate1")<>"" and request("PayDate2")<>""then
		ArgueDate1=gOutDT(request("PayDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("PayDate2"))&" 23:59:59"

		chk_year=year(ArgueDate2)-1911-6
		now_year=year(ArgueDate2)-1911
		now_month=month(ArgueDate2)

		nowdate=ArgueDate2
	end if

	if request("JudeDate1")<>"" and request("JudeDate2")<>""then
		ArgueDate1=gOutDT(request("JudeDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("JudeDate2"))&" 23:59:59"

		chk_year=year(ArgueDate2)-1911-6
		now_year=year(ArgueDate2)-1911
		now_month=month(ArgueDate2)

		nowdate=ArgueDate2
	end if

	if request("SendDate1")<>"" and request("SendDate2")<>""then
		ArgueDate1=gOutDT(request("SendDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("SendDate2"))&" 23:59:59"

		chk_year=year(ArgueDate2)-1911-6
		now_year=year(ArgueDate2)-1911
		now_month=month(ArgueDate2)

		nowdate=ArgueDate2
	end if

	if request("DeallIneDate1")<>"" and request("DeallIneDate2")<>""then
		ArgueDate1=gOutDT(request("DeallIneDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("DeallIneDate2"))&" 23:59:59"
		
		chk_year=year(ArgueDate2)-1911-6
		now_year=year(ArgueDate2)-1911
		now_month=month(ArgueDate2)

		nowdate=ArgueDate2
	end if

	if request("Sys_SendDetailDate1")<>"" and request("Sys_SendDetailDate2")<>""then
		ArgueDate1=gOutDT(request("Sys_SendDetailDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("Sys_SendDetailDate2"))&" 23:59:59"

		chk_year=year(ArgueDate2)-1911-6
		now_year=year(ArgueDate2)-1911
		now_month=month(ArgueDate2)

		nowdate=ArgueDate2
	end if

	if request("MakeSureDate1")<>"" and request("MakeSureDate2")<>""then
		ArgueDate1=gOutDT(request("MakeSureDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("MakeSureDate2"))&" 23:59:59"
		
		chk_year=year(ArgueDate2)-1911-6
		now_year=year(ArgueDate2)-1911
		now_month=month(ArgueDate2)

		nowdate=ArgueDate2
	end If 
	
	if request("CaseCloseDate1")<>"" and request("CaseCloseDate2")<>""then
		ArgueDate1=gOutDT(request("CaseCloseDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("CaseCloseDate2"))&" 23:59:59"

		chk_year=year(ArgueDate2)-1911-6
		now_year=year(ArgueDate2)-1911
		now_month=month(ArgueDate2)

		nowdate=ArgueDate2
	end If 

	If trim(request("Sys_PetitionDate1"))<>"" And trim(request("Sys_PetitionDate2"))<>"" Then
		ArgueDate1=gOutDT(request("Sys_PetitionDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("Sys_PetitionDate2"))&" 23:59:59"

		chk_year=year(ArgueDate2)-1911-6
		now_year=year(ArgueDate2)-1911
		now_month=month(ArgueDate2)

		nowdate=ArgueDate2
	End If 
	If ifnull(chk_year) Then
		
		chk_year=year(now)-1911-6
		now_year=year(now)-1911
		now_month=month(now)

		nowdate=year(now)&"/"&month(now)&"/"&day(now)&" 23:59:59"
	End if 

	nowdate="to_date('"&nowdate&"','YYYY/MM/DD/HH24/MI/SS')"

If Not ifnull(request("Sys_SendBillSN")) Then

	sys_billsn=request("Sys_SendBillSN")
elseif Not ifnull(request("hd_BillSN")) Then

	sys_billsn=request("hd_BillSN")
else

	sys_billsn=request("BillSN")
End If 

tmp_billsn=split(sys_billsn,",")

sys_billsn=""

For i = 0 to Ubound(tmp_billsn)

	If i >0 then

		If i mod 100 = 0 Then

			sys_billsn=sys_billsn&"@"
		elseif sys_billsn<>"" then

			sys_billsn=sys_billsn&","
		end If 
	end if

	sys_billsn=sys_billsn&tmp_billsn(i)

Next

tmpSQL=""

If Ubound(tmp_billsn) >= 100 Then

	sys_billsn=split(sys_billsn,"@")
	
	For i = 0 to Ubound(sys_billsn)
		
		If tmpSQL <>"" Then tmpSQL=tmpSQL&" union all "
		
		tmpSQL=tmpSQL&"select sn from passerbase where sn in("&sys_billsn(i)&")"
	Next

else

	tmpSQL="select sn from passerbase where sn in("&sys_billsn&")"

End if 

BasSQL="("&tmpSQL&") tmpPasser"

strSQL="select to_number(to_char(IllegalDate,'YYYY'))-1911 ilegalYear," & _
		"(case when " & _
		"	(to_number(to_char(IllegalDate,'YYYY'))-1911)<"&chk_year&" then "&now_year & _
		" else (to_number(to_char(IllegalDate,'YYYY'))-1911) " & _
		"end) payYear," & _
		"IllegalDate,(select unitname from unitinfo where unitid=a.billunitid) unitname,"&_
		"BillNo,Rule1,Rule2,Driver,"&_
		"(select max(paydate) from passerpay where paydate <= "&nowdate&" and billsn=a.sn) paydate," & _
		"nvl(forfeit1,0)+nvl(forfeit2,0) forfeit," & _
		"(select max(nvl(PayAmount,0)+nvl(MIDDLEMONEY,0)) from passerpay where paydate=(" & _
		"	select max(paydate) from passerpay where paydate <= "&nowdate&" and billsn=a.sn" & _
		") and billsn=a.sn) PayAmount1," & _
		"(select max(nvl(PayAmount,0)) from passerpay where paydate=(" & _
		"	select max(paydate) from passerpay where paydate <= "&nowdate&" and billsn=a.sn" & _
		") and billsn=a.sn) PayAmount2," & _
		"(select nvl(max(MIDDLEMONEY),0) from passerpay where paydate=(" & _
		"	select max(paydate) from passerpay where paydate <= "&nowdate&" and billsn=a.sn" & _
		") and billsn=a.sn) PayAmount3," & _
		"(select max(PayNo) from passerpay where paydate=(" & _
		"	select max(paydate) from passerpay where paydate <= "&nowdate&" and billsn=a.sn" & _
		") and billsn=a.sn) PayNo," & _
		"(case when " & _
		"	(select count(1) cnt from PasserSend where billsn=a.sn)=0 " & _
		"	and " & _
		"	(select max(PayTypeID) from passerpay where paydate=( " & _
		"		select max(paydate) from passerpay where paydate <= "&nowdate&" and billsn=a.sn" & _
		"	) and billsn=a.sn)=1 then '到場繳款'" & _
		" when " & _
		"	(select count(1) cnt from PasserSend where billsn=a.sn)=0 " & _
		"	and " & _
		"	(select max(PayTypeID) from passerpay where paydate=(" & _
		"		select max(paydate) from passerpay where paydate <= "&nowdate&" and billsn=a.sn" & _
		"	) and billsn=a.sn)=2 then '劃撥繳款'" & _
		" when " & _
		"	(select count(1) cnt from PasserSend where billsn=a.sn)>0 then '強制執行'" & _
		" end) PayType1," & _
		"(case when " & _
		"	(select count(1) cnt from PasserSend where billsn=a.sn)=0" & _
		"	 and " & _
		"	(select max(IsLate) from passerpay where paydate <= "&nowdate&" and billsn=a.sn)=0 then '未逾期'" & _
		" when " & _
		"	(select count(1) cnt from PasserSend where billsn=a.sn)=0" & _
		"	 and " & _
		"	(select max(IsLate) from passerpay where paydate <= "&nowdate&" and billsn=a.sn)=1 then '逾期'" & _
		" when " & _
		"	(select count(1) cnt from PasserSend where billsn=a.sn)>0 then '移送執行'" & _
		" end) PayType2," & _				
		"(case when BillStatus='9' then 1 else 0 end) BillStatus" & _
		" from PasserBase a where a.RecordStateID=0 and Exists(select 'Y' from "&BasSQL&" where SN=a.SN) and exists(select 'Y' from PasserPay where a.sn=BillSn) order by payno"

set rsdata=conn.execute(strSQL)

If rsdata.eof Then Response.End

Set YearObj = Server.CreateObject("Scripting.Dictionary")

strYear="":tmpYear=""

cntSQL="select distinct to_number(to_char(IllegalDate,'YYYY'))-1911 ilegalYear" & _
		" from PasserBase a where a.RecordStateID=0 and Exists(select 'Y' from "&BasSQL&" where SN=a.SN) and exists(select 'Y' from PasserPay where a.sn=BillSn) order by ilegalYear"

		set rscnt=conn.execute(cntSQL)
		While Not rscnt.eof
			
			If tmpYear <>"" Then tmpYear=tmpYear&","
			tmpYear=tmpYear&rscnt("ilegalYear")

			YearObj.Add rscnt("ilegalYear") & "_1",0
			YearObj.Add rscnt("ilegalYear") & "_2",0
			YearObj.Add rscnt("ilegalYear") & "_3",0

			rscnt.movenext
		Wend
		rscnt.close

		strYear=split(tmpYear,",")

		totalCnt=0:totalMoneyA=0:totalMoneyB=0
		A1_1=0:A1_2=0:B1_1=0:B1_2=0
		C1_1=0:C1_2=0:D1_1=0:D1_2=0
		E1_1=0:E1_2=0:F1_1=0:F1_2=0
%>
<head>
<meta http-equiv=Content-Type content="text/html; charset=big5">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 14">
<link rel=File-List href="PasserPayCloseDetail.files/filelist.xml">
<style id="PasserPayCloseDetail_21218_Styles">
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
.font521218
	{color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:新細明體, serif;
	mso-font-charset:136;}
.font621218
	{color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;}
.font721218
	{color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Wingdings 2", serif;
	mso-font-charset:2;}
.xl6321218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6421218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6521218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:121;
	vertical-align:121;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl6621218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:121;
	vertical-align:121;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl6721218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6821218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl6921218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl7021218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl7121218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl7221218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl7321218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Wingdings 2", serif;
	mso-font-charset:2;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl7421218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl7521218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0";
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7621218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7721218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7821218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:right;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7921218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl8021218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:right;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl8121218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl8221218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl8321218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl8421218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:13.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl8521218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:16.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl8621218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:16.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl8721218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:16.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid black;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl8821218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;
	layout-flow:vertical-ideographic;}
.xl8921218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;
	layout-flow:vertical-ideographic;}
.xl9021218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:121;
	vertical-align:121;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl9121218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:121;
	vertical-align:121;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl9221218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl9321218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl9421218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl9521218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl9621218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl9721218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl9821218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl9921218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl10021218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl10121218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl10221218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl10321218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl10421218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl10521218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl10621218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl10721218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl10821218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0";
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl10921218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl11021218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl11121218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl11221218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl11321218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl11421218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl11521218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl11621218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl11721218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl11821218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl11921218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl12021218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl12121218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl12221218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl12321218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl12421218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl12521218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl12621218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\@";
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl12721218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:14.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;
	layout-flow:vertical-ideographic;}
.xl12821218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:14.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;
	layout-flow:vertical-ideographic;}
.xl12921218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:14.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;
	layout-flow:vertical-ideographic;}
.xl13021218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:14.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl13121218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:14.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl13221218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:14.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid black;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl13321218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl13421218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl13521218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl13621218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl13721218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl13821218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid black;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl13921218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl14021218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl14121218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid black;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl14221218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl14321218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl14421218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid black;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl14521218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid black;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl14621218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid black;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl14721218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0_\)\;\[Red\]\\\(\#\,\#\#0\\\)";
	text-align:left;
	vertical-align:middle;
	border-top:none;
	border-right:.5pt solid black;
	border-bottom:.5pt solid black;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl14821218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:"\#\,\#\#0";
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid black;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl14921218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl15021218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl15121218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid black;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl15221218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl15321218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl15421218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:10.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:left;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid black;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl15521218
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:windowtext;
	font-size:14.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:標楷體, cursive;
	mso-font-charset:136;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
ruby
	{ruby-align:left;}
rt
	{color:windowtext;
	font-size:9.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:新細明體, serif;
	mso-font-charset:136;
	mso-char-type:none;}
-->
</style>
</head>

<body>
<!--[if !excel]>　　<![endif]-->
<!--下列資訊是由 Microsoft Excel 網頁發佈精靈所產生。-->
<!--如果由 Excel 重新發佈相同的項目時，在 DIV 標籤間的所有資訊將會被取代。-->
<!----------------------------->
<!--START OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD -->
<!----------------------------->

<div id="PasserPayCloseDetail_21218" align=center x:publishsource="Excel"><!--下列資訊是由 Microsoft Excel 網頁發佈精靈所產生。--><!--如果由 Excel 重新發佈相同的項目時，在 DIV 標籤間的所有資訊將會被取代。--><!-----------------------------><!--START OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD --><!----------------------------->

<table border=0 cellpadding=0 cellspacing=0 width=938 style='border-collapse:collapse;table-layout:fixed;width:707pt'>
 <col class=xl6321218 width=35 style='mso-width-source:userset;mso-width-alt:
 1203;width:26pt'>
 <col class=xl6321218 width=53 style='mso-width-source:userset;mso-width-alt:
 1843;width:40pt'>
 <col class=xl6321218 width=69 style='mso-width-source:userset;mso-width-alt:
 2406;width:52pt'>
 <col class=xl6321218 width=70 style='mso-width-source:userset;mso-width-alt:
 2432;width:53pt'>
 <col class=xl6421218 width=66 style='mso-width-source:userset;mso-width-alt:
 2304;width:50pt'>
 <col class=xl6321218 width=62 style='mso-width-source:userset;mso-width-alt:
 2150;width:47pt'>
 <col class=xl6321218 width=63 style='mso-width-source:userset;mso-width-alt:
 2176;width:47pt'>
 <col class=xl6321218 width=64 style='mso-width-source:userset;mso-width-alt:
 2227;width:48pt'>
 <col class=xl6321218 width=62 style='mso-width-source:userset;mso-width-alt:
 2150;width:47pt'>
 <col class=xl6321218 width=59 style='mso-width-source:userset;mso-width-alt:
 2048;width:44pt'>
 <col class=xl6321218 width=71 style='mso-width-source:userset;mso-width-alt:
 2457;width:53pt'>
 <col class=xl6321218 width=66 span=4 style='mso-width-source:userset;
 mso-width-alt:2304;width:50pt'>
	<tr class=xl6321218 height=38 style='mso-height-source:userset;height:29.25pt'>
		<td colspan=11 height=38 class=xl8521218 width=674 style='border-right:.5pt solid black;height:29.25pt;width:507pt'>
			臺中市政府警察局<%=TitleUnitName2%>辦理<%=now_year%>年<%=now_month%>月份交通違規罰鍰繳款明細表
		</td>
		<td class=xl6321218 width=66 style='width:50pt'></td>
		<td class=xl6321218 width=66 style='width:50pt'></td>
		<td class=xl6321218 width=66 style='width:50pt'></td>
		<td class=xl6321218 width=66 style='width:50pt'></td>
	</tr>
	<tr class=xl6321218 height=22 style='mso-height-source:userset;height:16.9pt'>
		<td rowspan=2 height=44 class=xl8821218 align=center style='border-bottom:.5pt solid black;height:33.8pt;border-top:none'>
			編號
		</td>
		<td rowspan=2 class=xl9021218 width=53 style='border-bottom:.5pt solid black;border-top:none;width:40pt'>
			違規<span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>日期
		</td>
		<td rowspan=2 class=xl9021218 width=69 style='border-bottom:.5pt solid black;border-top:none;width:52pt'>
			舉發<span style='mso-spacerun:yes'>&nbsp;&nbsp;</span>單位
		</td>
		<td rowspan=2 class=xl9021218 width=70 style='border-bottom:.5pt solid black;border-top:none;width:53pt'>
			舉發單<span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>字號
		</td>
		<td class=xl6621218 width=66 style='width:50pt'>
			法條
		</td>
		<td rowspan=2 class=xl9021218 width=62 style='border-bottom:.5pt solid black;border-top:none;width:47pt'>
			違規<span style='mso-spacerun:yes'>&nbsp; </span>人名
		</td>
		<td class=xl6621218 width=63 style='width:47pt'>繳款</td>
		<td class=xl6621218 width=64 style='width:48pt'>罰鍰</td>
		<td class=xl6621218 width=62 style='width:47pt'>收據</td>
		<td class=xl6621218 width=59 style='width:44pt'>繳費</td>
		<td rowspan=2 class=xl9021218 width=71 style='border-bottom:.5pt solid black;border-top:none;width:53pt'>
			備考
		</td>
		<td rowspan=2 class=xl9221218>　</td>
		<td rowspan=2 class=xl6321218></td>
		<td rowspan=2 class=xl6321218></td>
		<td rowspan=2 class=xl6321218></td>
	</tr>
	<tr class=xl6321218 height=22 style='mso-height-source:userset;height:16.9pt'>
		<td height=22 class=xl6521218 width=66 style='height:16.9pt;width:50pt'>代碼</td>
		<td class=xl6521218 width=63 style='width:47pt'>日期</td>
		<td class=xl6521218 width=64 style='width:48pt'>實收</td>
		<td class=xl6521218 width=62 style='width:47pt'>編號</td>
		<td class=xl6521218 width=59 style='width:44pt'>方式</td>
	</tr>
	<%
		While not rsdata.eof

			cntfile=cntfile+1

			YearObj.Item(rsdata("ilegalYear") & "_1")=YearObj.Item(rsdata("ilegalYear") & "_1")+1
			YearObj.Item(rsdata("payYear") & "_2")=YearObj.Item(rsdata("payYear") & "_2")+cdbl(rsdata("PayAmount2"))
			YearObj.Item(rsdata("payYear") & "_3")=YearObj.Item(rsdata("payYear") & "_3")+cdbl(rsdata("PayAmount3"))

			totalCnt=totalCnt+1
			totalMoneyA=totalMoneyA+cdbl(rsdata("PayAmount2"))
			totalMoneyB=totalMoneyB+cdbl(rsdata("PayAmount3"))

			If trim(rsdata("PayType1")) = "到場繳款" and trim(rsdata("PayType2")) = "未逾期" and cdbl(rsdata("PayAmount2")) >0 Then

				A1_1=A1_1+1

				If trim(rsdata("BillStatus")) = "0" Then

					A1_2=A1_2+1
				End if 
			End if 

			If trim(rsdata("PayType1")) = "到場繳款" and trim(rsdata("PayType2")) = "逾期" and cdbl(rsdata("PayAmount2")) >0 Then

				B1_1=B1_1+1

				If trim(rsdata("BillStatus")) = "0" Then

					B1_2=B1_2+1
				End if 
			End if 

			If trim(rsdata("PayType1")) = "劃撥繳款" and trim(rsdata("PayType2")) = "未逾期" and cdbl(rsdata("PayAmount2")) >0 Then

				C1_1=C1_1+1

				If trim(rsdata("BillStatus")) = "0" Then

					C1_2=C1_2+1
				End if 
			End if 

			
			If trim(rsdata("PayType1")) = "劃撥繳款" and trim(rsdata("PayType2")) = "逾期" and cdbl(rsdata("PayAmount2")) >0 Then

				D1_1=D1_1+1

				If trim(rsdata("BillStatus")) = "0" Then

					D1_2=D1_2+1
				End if 
			End if 

			If trim(rsdata("PayType1")) = "強制執行" and cdbl(rsdata("PayAmount2")) >0 Then

				E1_1=E1_1+1

				If trim(rsdata("BillStatus")) = "0" Then

					E1_2=E1_2+1
				End if 
			End if 

			
			If trim(rsdata("BillStatus")) = "1" and cdbl(rsdata("PayAmount2"))=0 Then

				G1_1=G1_1+1

			End if 

			Response.Write "<tr class=xl6321218 height=21 style='mso-height-source:userset;height:16.1pt'>"

			Response.Write "<td rowspan=4 height=84 class=xl9421218 style='border-bottom:.5pt solid black;height:64.4pt;border-top:none'>"&cntfile&"</td>"
			
			illegaldate=year(rsdata("IllegalDate"))-1911
			illegaldate=illegaldate&right("00"&month(rsdata("IllegalDate")),2)
			illegaldate=illegaldate&right("00"&day(rsdata("IllegalDate")),2)

			Response.Write "<td rowspan=4 class=xl9621218 width=53 style='border-bottom:.5pt solid black;border-top:none;width:40pt'>"&illegaldate&"</td>"

			Response.Write "<td rowspan=4 class=xl9621218 width=69 style='border-bottom:.5pt solid black;border-top:none;width:52pt'>"&rsdata("unitname")&"</td>"

			Response.Write "<td rowspan=4 class=xl9621218 width=70 style='border-bottom:.5pt solid black;border-top:none;width:53pt'>"&rsdata("BillNo")&"</td>"

			Response.Write "<td class=xl6321218></td>"

			Response.Write "<td rowspan=4 class=xl9421218 style='border-bottom:.5pt solid black;border-top:none'>"&rsdata("Driver")&"</td>"

			paydate=year(rsdata("paydate"))-1911
			paydate=paydate&right("00"&month(rsdata("paydate")),2)
			paydate=paydate&right("00"&day(rsdata("paydate")),2)

			Response.Write "<td rowspan=4 class=xl9921218 style='border-bottom:.5pt solid black;border-top:none'>"&paydate&"</td>"

			Response.Write "<td rowspan=4 class=xl9921218 style='border-bottom:.5pt solid black;border-top:none'>"&rsdata("forfeit")&"/"&rsdata("PayAmount1")&"</td>"

			Response.Write "<td rowspan=4 class=xl9921218 style='border-bottom:.5pt solid black;border-top:none'>"&rsdata("PayNo")&"</td>"

			Response.Write "<td class=xl11821218 style='border-top:none;border-left:none'>　</td>"
			Response.Write "<td class=xl7221218 width=71 style='width:53pt'>是否結案</td>"
			Response.Write "<td rowspan=4 class=xl9221218>　</td>"
			Response.Write "<td rowspan=4 class=xl6321218></td>"
			Response.Write "<td rowspan=4 class=xl6321218></td>"
			Response.Write "<td rowspan=4 class=xl6321218></td>"
			Response.Write "</tr>"

			Response.Write "<tr class=xl6321218 height=21 style='mso-height-source:userset;height:16.1pt'>"

			Response.Write "<td height=21 class=xl6821218 width=66 style='height:16.1pt;width:50pt'>"&rsdata("Rule1")&"</td>"

			Response.Write "<td class=xl7021218 width=59 style='width:44pt'>"&rsdata("PayType1")&"</td>"

			Response.Write "<td class=xl7321218 width=71 style='width:53pt'><font class=""font621218"">"
				
				If trim(rsdata("BillStatus")) = "1" Then
					Response.Write "■是/□否"
				else
					Response.Write "□是/■否"
				End If 
			Response.Write "</font></td>"

			Response.Write "</tr>"

			Response.Write "<tr class=xl6321218 height=21 style='mso-height-source:userset;height:16.1pt'>"
			Response.Write "<td height=21 class=xl6821218 width=66 style='height:16.1pt;width:50pt'>"&rsdata("Rule2")&"</td>"
			Response.Write "<td class=xl7021218 width=59 style='width:44pt'>"&rsdata("PayType2")&"</td>"
			Response.Write "<td class=xl7221218 width=71 style='width:53pt'>含執行費</td>"
			Response.Write "</tr>"


			Response.Write "<tr class=xl6321218 height=21 style='mso-height-source:userset;height:16.1pt'>"
			Response.Write "<td height=21 class=xl6921218 width=66 style='height:16.1pt;width:50pt'>　</td>"
			Response.Write "<td class=xl7121218 width=59 style='width:44pt'>　</td>"
			Response.Write "<td class=xl7421218 width=71 style='width:53pt'>"&rsdata("PayAmount3")&"元</td>"
			Response.Write "</tr>"


			rsdata.movenext
		Wend
		rsdata.close
	%>
	
 <tr class=xl6321218 height=38 style='mso-height-source:userset;height:29.25pt'>
  <td rowspan=15 height=470 class=xl12821218 style='border-bottom:.5pt solid black;
  height:353.75pt;border-top:none'>分項統計與註解</td>
  <td class=xl7621218>年度</td>
  <td class=xl7621218>件數</td>
  <td class=xl7621218>金額(元)</td>
  <td class=xl7721218>含執行費</td>
  <td class=xl7621218>金額(元)</td>
  <td colspan=5 class=xl13121218 style='border-right:.5pt solid black;
  border-left:none'>附註</td>
  <td class=xl7521218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
 </tr>
 <tr class=xl6321218 height=21 style='mso-height-source:userset;height:16.1pt'>
  <td rowspan=3 height=63 class=xl13321218 style='border-bottom:.5pt solid black;
  height:48.3pt;border-top:none'><%
	If ubound(strYear) >=0 Then Response.Write strYear(0)
  %></td>
  <td rowspan=3 class=xl13321218 style='border-bottom:.5pt solid black;
  border-top:none'><%
	If ubound(strYear) >=0 Then Response.Write YearObj.Item(strYear(0) & "_1")
  %></td>
  <td rowspan=3 class=xl13321218 style='border-bottom:.5pt solid black;
  border-top:none'><%
	If ubound(strYear) >=0 Then Response.Write YearObj.Item(strYear(0) & "_2")
  %></td>
  <td rowspan=3 class=xl13321218 style='border-bottom:.5pt solid black;
  border-top:none'>含執行費</td>
  <td rowspan=3 class=xl13321218 style='border-bottom:.5pt solid black;
  border-top:none'><%
	If ubound(strYear) >=0 Then Response.Write YearObj.Item(strYear(0) & "_3")
  %></td>
  <td colspan=5 class=xl13621218 width=319 style='border-right:.5pt solid black;
  border-left:none;width:239pt'>一、案件年度計算係以建檔日期為準。</td>
  <td rowspan=3 class=xl14821218>　</td>
  <td rowspan=3 class=xl6321218></td>
  <td rowspan=3 class=xl6321218></td>
  <td rowspan=3 class=xl6321218></td>
 </tr>
 <tr class=xl6321218 height=21 style='mso-height-source:userset;height:16.1pt'>
  <td colspan=5 height=21 class=xl13921218 width=319 style='border-right:.5pt solid black;
  height:16.1pt;border-left:none;width:239pt'>二、100年以前(不含101年以後)收繳罰鍰金額</td>
 </tr>
 <tr class=xl6321218 height=21 style='mso-height-source:userset;height:16.1pt'>
  <td colspan=5 height=21 class=xl13921218 width=319 style='border-right:.5pt solid black;
  height:16.1pt;border-left:none;width:239pt'><span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp; </span>(不含執行費)，直接併入今(107)年計算。</td>
 </tr>
 <tr class=xl6321218 height=32 style='mso-height-source:userset;height:23.95pt'>
  <td height=32 class=xl7721218 style='height:23.95pt'><%
	If ubound(strYear) >=1 Then Response.Write strYear(1)
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=1 Then Response.Write YearObj.Item(strYear(1) & "_1")
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=1 Then Response.Write YearObj.Item(strYear(1) & "_2")
  %></td>
  <td class=xl7721218>含執行費</td>
  <td class=xl7721218><%
	If ubound(strYear) >=1 Then Response.Write YearObj.Item(strYear(1) & "_3")
  %></td>
  <td colspan=5 class=xl14221218 style='border-right:.5pt solid black;
  border-left:none'>三、件數計算以舉發單號為主。</td>
  <td class=xl7521218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
 </tr>
 <tr class=xl6321218 height=32 style='mso-height-source:userset;height:23.95pt'>
  <td height=32 class=xl7721218 style='height:23.95pt'><%
	If ubound(strYear) >=2 Then Response.Write strYear(2)
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=2 Then Response.Write YearObj.Item(strYear(2) & "_1")
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=2 Then Response.Write YearObj.Item(strYear(2) & "_2")
  %></td>
  <td class=xl7721218>含執行費</td>
  <td class=xl7721218><%
	If ubound(strYear) >=2 Then Response.Write YearObj.Item(strYear(2) & "_3")
  %></td>
  <td colspan=5 class=xl14221218 style='border-right:.5pt solid black;
  border-left:none'>四、公危罪及撤銷(經判決或申訴)案件，仍須</td>
  <td class=xl7521218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
 </tr>
 <tr class=xl6321218 height=32 style='mso-height-source:userset;height:23.95pt'>
  <td height=32 class=xl7721218 style='height:23.95pt'><%
	If ubound(strYear) >=3 Then Response.Write strYear(3)
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=3 Then Response.Write YearObj.Item(strYear(3) & "_1")
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=3 Then Response.Write YearObj.Item(strYear(3) & "_2")
  %></td>
  <td class=xl7721218>含執行費</td>
  <td class=xl7721218><%
	If ubound(strYear) >=3 Then Response.Write YearObj.Item(strYear(3) & "_3")
  %></td>
  <td colspan=5 class=xl14521218 style='border-right:.5pt solid black;
  border-left:none'><span style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;
  </span>逐件顯示於上列表格，俾利後續審驗作業。</td>
  <td class=xl7521218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
 </tr>
 <tr class=xl6321218 height=32 style='mso-height-source:userset;height:23.95pt'>
  <td height=32 class=xl7721218 style='height:23.95pt'><%
	If ubound(strYear) >=4 Then Response.Write strYear(4)
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=4 Then Response.Write YearObj.Item(strYear(4) & "_1")
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=4 Then Response.Write YearObj.Item(strYear(4) & "_2")
  %></td>
  <td class=xl7721218>含執行費</td>
  <td class=xl7721218><%
	If ubound(strYear) >=4 Then Response.Write YearObj.Item(strYear(4) & "_3")
  %></td>
  <td class=xl7721218>到場繳款</td>
  <td class=xl7821218><%=A1_1%>件</td>
  <td class=xl7421218 width=62 style='width:47pt'>未逾期</td>
  <td class=xl7421218 width=59 style='width:44pt'>含未結</td>
  <td class=xl8021218 width=71 style='width:53pt'><%=A1_2%>件</td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
 </tr>
 <tr class=xl6321218 height=32 style='mso-height-source:userset;height:23.95pt'>
  <td height=32 class=xl7721218 style='height:23.95pt'><%
	If ubound(strYear) >=5 Then Response.Write strYear(5)
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=5 Then Response.Write YearObj.Item(strYear(5) & "_1")
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=5 Then Response.Write YearObj.Item(strYear(5) & "_2")
  %></td>
  <td class=xl7721218>含執行費</td>
  <td class=xl7721218><%
	If ubound(strYear) >=5 Then Response.Write YearObj.Item(strYear(5) & "_3")
  %></td>
  <td class=xl7721218>到場繳款</td>
  <td class=xl7821218><%=B1_1%>件</td>
  <td class=xl7421218 width=62 style='width:47pt'>逾期</td>
  <td class=xl7421218 width=59 style='width:44pt'>含未結</td>
  <td class=xl8021218 width=71 style='width:53pt'><%=B1_2%>件</td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
 </tr>
 <tr class=xl6321218 height=32 style='mso-height-source:userset;height:23.95pt'>
  <td height=32 class=xl7721218 style='height:23.95pt'><%
	If ubound(strYear) >=6 Then Response.Write strYear(6)
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=6 Then Response.Write YearObj.Item(strYear(6) & "_1")
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=6 Then Response.Write YearObj.Item(strYear(6) & "_2")
  %></td>
  <td class=xl7721218>含執行費</td>
  <td class=xl7721218><%
	If ubound(strYear) >=6 Then Response.Write YearObj.Item(strYear(6) & "_3")
  %></td>
  <td class=xl7721218>劃撥繳款</td>
  <td class=xl7821218><%=C1_1%>件</td>
  <td class=xl7421218 width=62 style='width:47pt'>未逾期</td>
  <td class=xl7421218 width=59 style='width:44pt'>含未結</td>
  <td class=xl8021218 width=71 style='width:53pt'><%=C1_2%>件</td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
 </tr>
 <tr class=xl6321218 height=32 style='mso-height-source:userset;height:23.95pt'>
  <td height=32 class=xl7721218 style='height:23.95pt'><%
	If ubound(strYear) >=7 Then Response.Write strYear(7)
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=7 Then Response.Write YearObj.Item(strYear(7) & "_1")
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=7 Then Response.Write YearObj.Item(strYear(7) & "_2")
  %></td>
  <td class=xl7721218>含執行費</td>
  <td class=xl7721218><%
	If ubound(strYear) >=7 Then Response.Write YearObj.Item(strYear(7) & "_3")
  %></td>
  <td class=xl7721218>劃撥繳款</td>
  <td class=xl7821218><%=D1_1%>件</td>
  <td class=xl7421218 width=62 style='width:47pt'>逾期</td>
  <td class=xl7421218 width=59 style='width:44pt'>含未結</td>
  <td class=xl8021218 width=71 style='width:53pt'><%=D1_2%>件</td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
 </tr>
 <tr class=xl6321218 height=32 style='mso-height-source:userset;height:23.95pt'>
  <td height=32 class=xl7721218 style='height:23.95pt'><%
	If ubound(strYear) >=8 Then Response.Write strYear(8)
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=8 Then Response.Write YearObj.Item(strYear(8) & "_1")
  %></td>
  <td class=xl7721218><%
	If ubound(strYear) >=8 Then Response.Write YearObj.Item(strYear(8) & "_2")
  %></td>
  <td class=xl7721218>含執行費</td>
  <td class=xl7721218><%
	If ubound(strYear) >=8 Then Response.Write YearObj.Item(strYear(8) & "_3")
  %></td>
  <td class=xl7721218>強制扣款</td>
  <td class=xl7821218><%=E1_1%>件</td>
  <td class=xl7421218 width=62 style='width:47pt'>移送執行</td>
  <td class=xl7421218 width=59 style='width:44pt'>含未結</td>
  <td class=xl8021218 width=71 style='width:53pt'><%=E1_2%>件</td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
 </tr>
 <tr class=xl6321218 height=32 style='mso-height-source:userset;height:23.95pt'>
  <td colspan=3 height=32 class=xl15021218 width=192 style='border-right:.5pt solid black;
  height:23.95pt;border-left:none;width:145pt'>經法院判決或通過申訴</td>
  <td class=xl7621218><%=G1_1%></td>
  <td class=xl7621218>件</td>
  <td class=xl7721218>匯票繳款</td>
  <td class=xl7821218><%=F1_1%>件</td>
  <td class=xl7421218 width=62 style='width:47pt'>移送執行</td>
  <td class=xl7421218 width=59 style='width:44pt'>含未結</td>
  <td class=xl8021218 width=71 style='width:53pt'><%=F1_2%>件</td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
 </tr>
 <tr class=xl6321218 height=32 style='mso-height-source:userset;height:23.95pt'>
  <td height=32 class=xl7921218 width=53 style='height:23.95pt;width:40pt'>小計</td>
  <td class=xl8121218 width=69 style='width:52pt'><%=totalCnt%></td>
  <td class=xl8121218 width=70 style='width:53pt'>件</td>
  <td colspan=2 class=xl15221218><%=totalMoneyA%></td>
  <td class=xl8121218 width=63 style='width:47pt'>元</td>
  <td colspan=4 class=xl15321218 style='border-right:.5pt solid black'>(不含執行費)</td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
 </tr>
 <tr class=xl6321218 height=49 style='mso-height-source:userset;height:36.7pt'>
  <td height=49 class=xl7921218 width=53 style='height:36.7pt;width:40pt'>合計</td>
  <td class=xl8121218 width=69 style='width:52pt'><%=totalCnt%></td>
  <td class=xl8121218 width=70 style='width:53pt'>件</td>
  <td colspan=2 class=xl15221218><%=totalMoneyA+totalMoneyB%></td>
  <td class=xl8121218 width=63 style='width:47pt'>元</td>
  <td class=xl8221218>含執行費</td>
  <td class=xl8321218><%=totalMoneyB%></td>
  <td class=xl8121218 width=59 style='width:44pt'>元</td>
  <td class=xl6721218>　</td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl8421218 width=66 style='width:50pt'></td>
 </tr>
 <tr class=xl6321218 height=37 style='mso-height-source:userset;height:27.7pt'>
  <td colspan=11 height=37 class=xl15521218 style='height:27.7pt'>承辦人：<span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span>會計室：<span
  style='mso-spacerun:yes'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  </span>分局長：</td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218></td>
  <td class=xl6321218><!-----------------------------><!--END OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD--><!-----------------------------></td>
 </tr>
 <![if supportMisalignedColumns]>
 <tr height=0 style='display:none'>
  <td width=35 style='width:26pt'></td>
  <td width=53 style='width:40pt'></td>
  <td width=69 style='width:52pt'></td>
  <td width=70 style='width:53pt'></td>
  <td width=66 style='width:50pt'></td>
  <td width=62 style='width:47pt'></td>
  <td width=63 style='width:47pt'></td>
  <td width=64 style='width:48pt'></td>
  <td width=62 style='width:47pt'></td>
  <td width=59 style='width:44pt'></td>
  <td width=71 style='width:53pt'></td>
  <td width=66 style='width:50pt'></td>
  <td width=66 style='width:50pt'></td>
  <td width=66 style='width:50pt'></td>
  <td width=66 style='width:50pt'></td>
 </tr>
 <![endif]>
</table>

</div>


<!----------------------------->
<!--END OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD-->
<!----------------------------->
</body>

</html>

<%

fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_交通違規罰鍰繳款明細表"
Response.AddHeader "Content-Disposition", "filename="&fname&".xls"
response.contenttype="application/x-msexcel; charset=MS950"
%>	
