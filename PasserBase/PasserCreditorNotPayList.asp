<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
'fMnoth=month(now)
'if fMnoth<10 then fMnoth="0"&fMnoth
'fDay=day(now)
'if fDay<10 then	fDay="0"&fDay
'fname=year(now)&fMnoth&fDay&"_債權憑證清冊.xls"
'Response.AddHeader "Content-Disposition", "filename="&fname
'response.contenttype="application/x-msexcel; charset=MS950"

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

strwhere=""

if trim(request("sys_CreditorTypeID"))<>"" Or (trim(request("Sys_PetitionDate1"))<>"" And trim(request("Sys_PetitionDate2"))<>"") Then


	If trim(request("sys_CreditorTypeID")) <> "-1" Then
		If trim(request("sys_CreditorTypeID"))<>"" Then
			strwhere=strwhere&" and CreditorTypeID in('"&trim(request("sys_CreditorTypeID"))&"')"
		End If 
	end If 
	
	If trim(request("Sys_PetitionDate1"))<>"" And trim(request("Sys_PetitionDate2"))<>"" Then
		PetitionDate1=gOutDT(request("Sys_PetitionDate1"))&" 0:0:0"
		PetitionDate2=gOutDT(request("Sys_PetitionDate2"))&" 23:59:59"

		strwhere=strwhere&" and PetitionDate between TO_DATE('"&PetitionDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&PetitionDate2&"','YYYY/MM/DD/HH24/MI/SS')"

	End If

	If trim(request("sys_CreditorTypeID")) = "-1" Then
		strwhere=strwhere&" and CreditorTypeID not in('0','1')"
	End if		
	
end if 

strSQL="select " &_	
	"(Select billno from PasserBase where sn=pb.billsn) billno," &_
	"(Select driver from PasserBase where sn=pb.billsn) driver," &_
	"(Select driverid from PasserBase where sn=pb.billsn) driverid," &_
	"(Select limitdate from PasserSend where billsn=pb.billsn) limitdate," &_
	"pb.OpenGovNumber,nvl(pb.RemainNT,0) RemainNT" &_
	" from PasserCreditor pb where exists(select 'Y' from PasserSend where BillSN=pb.billsn) and exists(select 'Y' from PasserSendDetail where BillSN=pb.billsn) and Exists(select 'Y' from "&BasSQL&" where SN=pb.billsn)"&strwhere&" order by DriverID,LimitDate,OpenGovNumber"

set rs=conn.execute(strSQL)

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>債權憑證清冊</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="700" border="0">
	<tr><td>
		<table width="100%" border="0" cellpadding="4" cellspacing="1">
			<tr>
				<td colspan="4" align="left">列印日期：<%=gInitDt(date)%></td>
			</tr>
			<tr>
				<td colspan="4" align="left">處理時間：<%
					Response.Write request("Sys_PetitionDate1")&"～"&request("Sys_PetitionDate2")
				%></td>
			</tr>
			<tr>
				<td colspan="4" align="left">登入者：<%=Session("Ch_Name")%></td>
			</tr>
			<tr>
				<td colspan="9" align="center"><strong>債權憑證清冊</strong></td>
			</tr>
			<tr>
				<td colspan="9" align="center"><strong>　</strong></td>
			</tr>
		</table>
	</td></tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="1">
				<tr>
					<td>編號</td>
					<td>義務人</td>
					<td>身份證字號</td>
					<td>憑證字號</td>
					<td>舉發單號</td>
					<td>限繳日期</td>
					<td>件數</td>
					<td>金額</td>
					<td>備註</td>
				</tr>
				<%
				tmp_DriverID="":tmp_LimitDate="":tmp_OpenGovNumber="":sum_RemainNT=0:sum_Count=0
				tmp_BillNo="":tmp_Driver="":fileNum=0:chkfile=1
				While not rs.eof
					If tmp_DriverID="" then
						tmp_Driver=trim(rs("Driver"))
						tmp_DriverID=trim(rs("DriverID"))
						tmp_LimitDate=trim(rs("LimitDate"))
						tmp_OpenGovNumber=trim(rs("OpenGovNumber"))
						sum_RemainNT=cdbl(rs("RemainNT"))
						sum_Count=1
						tmp_BillNo=trim(rs("BillNo"))

						rs.movenext	

						If rs.eof Then
							chkfile=2
							fileNum=fileNum+1
							Response.Write "<tr>"
							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write fileNum
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write tmp_Driver
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write tmp_DriverID
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write tmp_OpenGovNumber
							Response.Write "</td>"

							tmp_BillNo=split(tmp_BillNo,",")
							Response.Write "<td>"
							Response.Write tmp_BillNo(0)
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write gInitDT(tmp_LimitDate)
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write sum_Count
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write sum_RemainNT
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">　</td>"
							Response.Write "</tr>"

							If Ubound(tmp_BillNo) > 0 Then
								For i = 1 to Ubound(tmp_BillNo)
									Response.Write "<tr><td>"&tmp_BillNo(i)&"</td></tr>"
								Next
							End if
						end if

					elseif tmp_DriverID=trim(rs("DriverID")) and tmp_LimitDate=trim(rs("LimitDate")) and tmp_OpenGovNumber=trim(rs("OpenGovNumber")) then
						
						sum_RemainNT=sum_RemainNT+cdbl(rs("RemainNT"))
						sum_Count=sum_Count+1
						tmp_BillNo=tmp_BillNo&","&trim(rs("BillNo"))

						rs.movenext

						If rs.eof Then
							chkfile=2

							fileNum=fileNum+1
							Response.Write "<tr>"
							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write fileNum
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write tmp_Driver
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write tmp_DriverID
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write tmp_OpenGovNumber
							Response.Write "</td>"

							tmp_BillNo=split(tmp_BillNo,",")
							Response.Write "<td>"
							Response.Write tmp_BillNo(0)
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write gInitDT(tmp_LimitDate)
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write sum_Count
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write sum_RemainNT
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">　</td>"
							Response.Write "</tr>"

							If Ubound(tmp_BillNo) > 0 Then
								For i = 1 to Ubound(tmp_BillNo)
									Response.Write "<tr><td>"&tmp_BillNo(i)&"</td></tr>"
								Next
							End if
						end if
					else
						fileNum=fileNum+1
						Response.Write "<tr>"
						Response.Write "<td rowspan="""&sum_Count&""">"
						Response.Write fileNum
						Response.Write "</td>"

						Response.Write "<td rowspan="""&sum_Count&""">"
						Response.Write tmp_Driver
						Response.Write "</td>"

						Response.Write "<td rowspan="""&sum_Count&""">"
						Response.Write tmp_DriverID
						Response.Write "</td>"

						Response.Write "<td rowspan="""&sum_Count&""">"
						Response.Write tmp_OpenGovNumber
						Response.Write "</td>"

						tmp_BillNo=split(tmp_BillNo,",")
						Response.Write "<td>"
						Response.Write tmp_BillNo(0)
						Response.Write "</td>"

						Response.Write "<td rowspan="""&sum_Count&""">"
						Response.Write gInitDT(tmp_LimitDate)
						Response.Write "</td>"

						Response.Write "<td rowspan="""&sum_Count&""">"
						Response.Write sum_Count
						Response.Write "</td>"

						Response.Write "<td rowspan="""&sum_Count&""">"
						Response.Write sum_RemainNT
						Response.Write "</td>"

						Response.Write "<td rowspan="""&sum_Count&""">　</td>"
						Response.Write "</tr>"

						If Ubound(tmp_BillNo) > 0 Then
							For i = 1 to Ubound(tmp_BillNo)
								Response.Write "<tr><td>"&tmp_BillNo(i)&"</td></tr>"
							Next
						End if
						tmp_BillNo=""
						tmp_Driver=trim(rs("Driver"))
						tmp_DriverID=trim(rs("DriverID"))
						tmp_LimitDate=trim(rs("LimitDate"))
						tmp_OpenGovNumber=trim(rs("OpenGovNumber"))
						sum_RemainNT=cdbl(rs("RemainNT"))
						sum_Count=1
						tmp_BillNo=trim(rs("BillNo"))

						rs.movenext

						If rs.eof Then
							chkfile=2

							fileNum=fileNum+1
							Response.Write "<tr>"
							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write fileNum
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write tmp_Driver
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write tmp_DriverID
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write tmp_OpenGovNumber
							Response.Write "</td>"

							tmp_BillNo=split(tmp_BillNo,",")
							Response.Write "<td>"
							Response.Write tmp_BillNo(0)
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write gInitDT(tmp_LimitDate)
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write sum_Count
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">"
							Response.Write sum_RemainNT
							Response.Write "</td>"

							Response.Write "<td rowspan="""&sum_Count&""">　</td>"
							Response.Write "</tr>"

							If Ubound(tmp_BillNo) > 0 Then
								For i = 1 to Ubound(tmp_BillNo)
									Response.Write "<tr><td>"&tmp_BillNo(i)&"</td></tr>"
								Next
							End if
						end if
						
					end if
				Wend
				rs.close
				set rsfound=nothing
				%>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%
conn.close
set conn=nothing
%>