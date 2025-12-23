<!--#include virtual="Traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="Traffic/Common/OlddbAccessKao.ini"-->
<!--#include virtual="Traffic/Common/AllFunction.inc"-->
<% 

function GetDate(tDate)
	if len(tDate)=7 then
		GetDate=left(tDate,3)&"年"& mid(tDate,4,2)&"月"& Right(tDate,2)&"日"
	else
		GetDate="&nbsp;"
	end if
end function

function GetTime(ttime)
  W=""
  H=""
  H=left(ttime,2)
  N=right(ttime,2)
  if cdbl(H)=12 then
    W="中午"
  elseif cdbl(H)<6  then
    W="凌晨"
  elseif cdbl(H)>5 and cdbl(H)<12 then
    W="早上"
  elseif cdbl(H)>12 and cdbl(H)<18 then
    W="下午"
  elseif cdbl(H)>17 then
    W="晚上"
  end if

  SH=0

  if H>12 then SH=cdbl(H)-12 else SH=H

	if len(ttime)=4 then
		GetTime=W&" "&right("00"&SH,2)&"點"&N&"分"
	else
		GetTime="&nbsp;"
	end if
end function

function getDciCodeN(code)
  if code="S" then
    getDciCodeN="送達註記	成功註記"
  elseif code="N" then
    getDciCodeN="送達註記	找不到此筆資料"
  elseif code="n" then
    getDciCodeN="送達註記	已經結案"
  elseif code="k" then
    getDciCodeN="送達註記	已送達不可做未送達註記"
  elseif code="h" then
    getDciCodeN="送達註記	已開裁決書"
  elseif code="B" then
    getDciCodeN="送達註記	無此車號/無此證號"
  elseif code="E" then
    getDciCodeN="送達註記	日期錯誤"
  else
    getDciCodeN="&nbsp;"
  end if
end Function
 
function getDciCode(code)
  if code="00" then
    getDciCode="入案  未寫入資料庫"
  elseif code="Y" Then
    getDciCode="入案  寫入資料庫"
  elseif code="N" then
    getDciCode="入案	未寫入資料庫"
  elseif code="S" then
    getDciCode="結案	違規人已經繳費"
  elseif code="L" then
    getDciCode="入案	已經入案過"
  elseif code="n" then
    getDciCode="入案	監理單位已經入案"
  elseif code="0" then
    getDciCode="入案	正常"
  elseif code="1" then
    getDciCode="入案錯誤車號不全"
  elseif code="2" then
    getDciCode="入案錯誤	扣件不符"
  elseif code="3" then
    getDciCode="入案錯誤	車號不正確"
  elseif code="4" then
    getDciCode="入案錯誤	證號不正確"
  elseif code="5" then
    getDciCode="入案錯誤	處所不明"
  elseif code="6" then
    getDciCode="入案錯誤	違警案件"
  elseif code="7" then
    getDciCode="入案錯誤	未簽名"
  elseif code="8" then
    getDciCode="入案錯誤	未通知違規人"
  elseif code="9" then
    getDciCode="入案錯誤	時間不明確"
  elseif code="a" then
    getDciCode="入案錯誤	事實不明確"
  elseif code="b" then
    getDciCode="入案錯誤	欠缺駕照證號"
  elseif code="c" then
    getDciCode="入案錯誤	公文移轉"
  elseif code="d" then
    getDciCode="入案錯誤	碰結案"
  elseif code="e" then
    getDciCode="入案錯誤	非管轄碰結案"
  elseif code="f" then
    getDciCode="入案錯誤	碰未結"
  elseif code="g" then
    getDciCode="入案錯誤	移出"
  elseif code="h" then
    getDciCode="入案錯誤	舉類錯誤"
  elseif code="i" then
    getDciCode="入案錯誤	攔停需指定所站"
  elseif code="j" then
    getDciCode="入案錯誤	單號不足9位"
  elseif code="k" then
    getDciCode="入案錯誤	攔停車駕條款一起"
  elseif code="l" then
    getDciCode="入案錯誤	無此單號剔退"
  elseif code="m" then
    getDciCode="入案錯誤	條款與車別不符"
  elseif code="z" then
    getDciCode="入案錯誤	道安已完成記點"
  elseif code="A" then
    getDciCode="入案錯誤	條款錯誤"
  elseif code="B" then
    getDciCode="入案錯誤	車籍無記錄"
  elseif code="C" then
    getDciCode="入案錯誤	駕籍無紀錄"
  elseif code="D" then
    getDciCode="入案錯誤	過戶前案"
  elseif code="E" then
    getDciCode="入案錯誤	繳註銷前案"
  elseif code="F" then
    getDciCode="入案錯誤	繳註銷後案"
  elseif code="G" then
    getDciCode="入案錯誤	吊扣銷中案"
  elseif code="H" then
    getDciCode="入案錯誤	證號重號剔退"
  elseif code="I" then
    getDciCode="入案錯誤	達記點吊扣"
  elseif code="J" then
    getDciCode="入案錯誤	達記點吊銷"
  elseif code="K" then
    getDciCode="入案錯誤	單號+車號重覆"
  elseif code="L" then
    getDciCode="入案錯誤	重覆入銷案剔退"
  elseif code="M" then
    getDciCode="入案錯誤	無照駕駛"
  elseif code="N" then
    getDciCode="入案錯誤	未知,找不到"
  elseif code="O" then
    getDciCode="入案錯誤車駕非管轄"
  elseif code="P" then
    getDciCode="入案錯誤	照類不符"
  elseif code="Q" then
    getDciCode="入案錯誤	前車號違規"
  elseif code="q" then
    getDciCode="入案錯誤	已定期換牌"
  elseif code="S" then
    getDciCode="入案錯誤	非管轄"
  elseif code="T" then
    getDciCode="入案錯誤	問題車牌"
  elseif code="U" then
    getDciCode="入案錯誤	未異動"
  elseif code="V" then
    getDciCode="入案錯誤	失竊註銷"
  elseif code="X" then
    getDciCode="入案錯誤	未新增"
  elseif code="Y" then
    getDciCode="入案錯誤	資料庫錯誤"
  elseif code="Z" then
    getDciCode="入案錯誤	道安已完成開單"
  elseif code="*" then
    getDciCode="入案錯誤	刪除不入案"
  elseif code="y" then
    getDciCode="入案錯誤	行照過期"
  elseif code="x" then
    getDciCode="入案錯誤	駕照過期"
  elseif code="R" then
    getDciCode="入案錯誤	非管轄"
  else
    getDciCode="&nbsp;"
  end if

end function
sql="select *,'' as SEQNO from FMaster where FSEQ='"&request("BillNo")&"'"
   set rs1=conn1.execute(sql)
   if rs1.eof Then   
	   set rs1=conn2.execute(sql)
	   if rs1.eof Then
		set rs1=conn3.execute(sql)
		   if rs1.eof Then
			set rs1=conn4.execute(sql)
			   if rs1.eof Then
				set rs1=conn5.execute(sql)
				   if rs1.eof Then
					set rs1=conn6.execute(sql)
					   if rs1.eof Then
						set rs1=conn7.execute(sql)
						   if rs1.eof Then
							   set rs1=conn1.execute(sql)
							   If rs1.eof Then
										strSQL="select *,'' as IRCODE,'' as INSUCERT,'' as AccUSeCode,'' as FACTG1,'' as OWNAME,'' as OWADDR,'' as CDType,'' as RBType,'' as FinDate,'' as MailDate,'' as Errst_SV,'' as MailSEQNO,'' as ErrST_SD,'' as SendPrnDate,'' as AMT1_Fin,'' as AMT2_Fin,'' as AMT3_Fin,'' as AMT4_Fin,'' as TAXFLAG,'' as PUBFLAG,'' as RULEG1,EnDate from FMaster_S where FSEQ='"&request("BillNo")&"'"
										set rs1=conn1.execute(strSQL)
								End If
							End if
					   end If

				   end If

			   end If

		   end If

	   end If
   end If


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舊告發綜合查詢詳細資料</title>
</head>
<body>
<script type="text/javascript" src="../js/date.js"></script>
	<table width='100%' border='1' cellpadding="2" id="table1">
		<tr bgcolor="#FFCC33">
			<td><strong>舊告發綜合查詢詳細資料</strong></td>
		</tr>
		</table>
	<table width='100%' border='1' cellpadding="2" id="table2">
	<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="35"><b>告發單資料</b></td>
		<tr>
			<td bgcolor="#FFFF99" width="12%" align="right"><strong>告發單編號</strong></td>
			<td align="left" width="19%"><%
			if trim(rs1("FSEQ"))<>""  and not isnull(rs1("FSEQ")) then
				response.write trim(rs1("FSEQ"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#FFFF99" width="18%" align="right"><strong>入案狀態</strong></td>
			<td align="left" width="15%">
			<%
			if trim(rs1("FStatus"))<>"" and not isnull(rs1("FStatus")) then
				response.write trim(rs1("FStatus"))&"&nbsp;"&getDciCode(trim(rs1("FStatus")))
			else
				response.write "&nbsp;"
			end if
			%>
			</td>
			<td bgcolor="#FFFF99" align="right"><strong>流水號</strong></td>
			<td align="left">
			<%
			if trim(rs1("SEQNO"))<>"" and not isnull(rs1("SEQNO")) then
				response.write trim(rs1("SEQNO"))
			else
				response.write "&nbsp;"
			end if
			%>
			</td>
		</tr>
			<td bgcolor="#FFFF99" align="right"><strong>填單日期</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("RBDate"))<>"" and not isnull(rs1("RBDate")) then
				response.write GetDate(rs1("RBDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>			
			<td bgcolor="#FFFF99" align="right"><strong>員警</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("PCode1"))<>"" and not isnull(rs1("PCode1")) then
				response.write trim(rs1("PCode1"))

				sql="select PNAME from Police where PCODE ='"&rs1("PCode1")&"'"
                set rsPCode=conn1.execute(sql)
				if not rspcode.eof then response.write " "&rsPCode("PName")
                set rsPcode=nothing

                if trim(rs1("PCode2"))<>"" and not isnull(rs1("PCode2")) then
					response.write ","&trim(rs1("PCode2"))
					sql="select PNAME from Police where PCODE ='"&rs1("PCode2")&"'"
	                set rsPCode=conn1.execute(sql)

					if not rspcode.eof then response.write " "&rsPCode("PName")
					set rsPcode=nothing
                end if

				if trim(rs1("PCode3"))<>"" and not isnull(rs1("PCode3")) then
					response.write ","&trim(rs1("PCode3"))
					sql="select PNAME from Police where PCODE ='"&rs1("PCode3")&"'"
	                set rsPCode=conn1.execute(sql)
					if not rspcode.eof then response.write " "&rsPCode("PName")
					set rsPcode=nothing
				end if

				if trim(rs1("PCode4"))<>"" and not isnull(rs1("PCode4")) then
					response.write ","&trim(rs1("PCode4"))
					sql="select PNAME from Police where PCODE ='"&rs1("PCode4")&"'"
					set rsPCode=conn1.execute(sql)
					if not rspcode.eof then response.write " "&rsPCode("PName")
					set rsPcode=nothing
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td bgcolor="#FFFF99" align="right"><strong>違規車號</strong></td>
			<td align="left"><%
			if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
				response.write trim(rs1("CarNo"))
			else
				response.write "&nbsp;"
			end if%>
			<td bgcolor="#FFFF99" align="right"><strong>違規日期</strong></td>
			<td align="left" width="17%"><%
			if trim(rs1("IDate"))<>"" and not isnull(rs1("IDate")) then
				response.write GetDate(rs1("IDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#FFFF99" align="right" width="12%"><strong>違規時間</strong></td>
			<td align="left"><%
			if trim(rs1("ITime"))<>"" and not isnull(rs1("ITime")) then
				response.write GetTime(rs1("ITime")) &" = "&left(rs1("ITime"),2)&":"&right(rs1("ITime"),2)
			else
				response.write "&nbsp;"
			end if
			%></td>
	  <tr>
			<td bgcolor="#FFFF99" align="right"><strong>違規地點</strong></td>
			<td align="left" colspan="5"><%
		    	if trim(rs1("IRCODE"))<>"" and not isnull(rs1("IRCODE")) then
		    		response.write trim(rs1("IRCODE"))
		    	else
		    		response.write "&nbsp;"
		    	end if
		    	if trim(rs1("IRNAME"))<>"" and not isnull(rs1("IRNAME")) then
		    		response.write " "&trim(rs1("IRNAME"))
		    	else
		    		response.write "&nbsp;"
		    	end if
			%></td>
        <tr>
			<td bgcolor="#FFFF99" align="right"><strong>簡式車種代碼</strong></td>
			<td align="left"><%
		    	if trim(rs1("CDKind"))<>"" and not isnull(rs1("CDKind")) then
						if rs1("CDKIND")="1" then
							CDKIND="汽車"
						elseif rs1("CDKIND")="2" then
							CDKIND="拖車"
						elseif rs1("CDKIND")="3" then
							CDKIND="重機"
						elseif rs1("CDKIND")="4" then
							CDKIND="輕機"
						end if
		    		response.write trim(rs1("CDKind")) & " " & CDKIND
		    	else
		    		response.write "&nbsp;"
		    	end if
			%>
			<td bgcolor="#FFFF99" align="right"><strong>保險證狀態</strong></td>
			<td align="left" colspan="3"><%
		    	if trim(rs1("INSUCERT"))<>"" and not isnull(rs1("INSUCERT")) then
						if rs1("INSUCERT")="0" then
							INSUCERT="正常"
						elseif rs1("INSUCERT")="1" then
							INSUCERT="未帶"
						elseif rs1("INSUCERT")="2" then
							INSUCERT="肇事且未帶"
						elseif rs1("INSUCERT")="3" then
							INSUCERT="過期或未保"
						elseif rs1("INSUCERT")="4" then
							INSUCERT="肇事且過期或未保"
						end if
		    		response.write trim(rs1("INSUCERT")) & " " & INSUCERT
		    	else
		    		response.write "&nbsp;"
		    	end if
				%></td>
        <tr>
		<td bgcolor="#FFFF99" align="right"><strong>告發類別</strong></td>
			<td align="left" colspan="5"><%
		    	if trim(rs1("AccUSeCode"))<>"" and not isnull(rs1("AccUSeCode")) then
						if rs1("AccUSeCode")="1" then 
						  AccUSeCode="攔停"
						elseif rs1("AccUSeCode")="2" then 
						  AccUSeCode="逕舉"
						elseif rs1("AccUSeCode")="8" then 
						  AccUSeCode="行人攤販"
						elseif rs1("AccUSeCode")="3" then 
						  AccUSeCode="肇事"
						elseif rs1("AccUSeCode")="4" then 
						  AccUSeCode="拖吊"
						elseif rs1("AccUSeCode")="5" then 
						  AccUSeCode="戴運砂石土方"
						elseif rs1("AccUSeCode")="A" then 
						  AccUSeCode="違規營業"
						elseif rs1("AccUSeCode")="B" then 
						  AccUSeCode="違規重標"
						elseif rs1("AccUSeCode")="N" then 
						  AccUSeCode="未知"
						end if 
		    		response.write trim(rs1("AccUSeCode")) & " " &AccUSeCode
		    	else
		    		response.write "&nbsp;"
		    	end if
				%></td>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>應到案日期</strong></td>
			<td align="left"><%
			if trim(rs1("ARVDATE"))<>"" and not isnull(rs1("ARVDATE")) then
				response.write GetDate(rs1("ARVDATE"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>應到案處所</strong></td>
			<td align="left" colspan="3"><%
			if trim(rs1("SPRVSNNO"))<>"" and not isnull(rs1("SPRVSNNO")) then
		    		response.write trim(rs1("SPRVSNNO")) & " " 
  				rstemp="select SPNAME from SPRVSN where SPRVSNNO='"&trim(rs1("SPRVSNNO"))&"'"
				set rstemp=conn1.execute(rstemp)
				response.write  trim(rstemp("SPNAME"))
				set rstemp=nothing
           else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>代保管物件</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("HOLDCODE1"))<>"" and not isnull(rs1("HOLDCODE1")) and trim(rs1("HOLDCODE1"))<>"0" then
				response.write trim(rs1("HOLDCODE1"))

  				rstemp="select HOLDName from HOLD where HOLDCode='"&trim(rs1("HOLDCODE1"))&"'"
                set rstemp=conn1.execute(rstemp)
					If Not rstemp.eof Then  
						response.write   "&nbsp;"&trim(rstemp("HOLDName")&"")
					Else
						response.write   "&nbsp;"
					End if
				set rstemp=nothing

				if trim(rs1("HOLDCODE2"))<>"" and not isnull(rs1("HOLDCODE2")) and trim(rs1("HOLDCODE2"))<>"0" then
					response.write "<br>"&trim(rs1("HOLDCODE2"))

					rstemp="select HOLDName from HOLD where HOLDCode='"&trim(rs1("HOLDCODE2"))&"'"
	                set rstemp=conn1.execute(rstemp)
					If Not rstemp.eof Then  
						response.write   "&nbsp;"&trim(rstemp("HOLDName")&"")
					Else
						response.write   "&nbsp;"
					End if
					set rstemp=nothing
				end if

				if trim(rs1("HOLDCODE3"))<>"" and not isnull(rs1("HOLDCODE3")) and trim(rs1("HOLDCODE3"))<>"0" then
					response.write "<br>"&trim(rs1("HOLDCODE3"))

					rstemp="select HOLDName from HOLD where HOLDCode='"&trim(rs1("HOLDCODE3"))&"'"
					set rstemp=conn1.execute(rstemp)
					If Not rstemp.eof Then  
						response.write   "&nbsp;"&trim(rstemp("HOLDName")&"")
					Else
						response.write   "&nbsp;"
					End if
					set rstemp=nothing
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>違規人姓名</strong></td>
			<td align="left"><%
			if trim(rs1("IName"))<>"" and not isnull(rs1("IName")) then
				response.write trim(rs1("IName"))
             else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>身份證字號</strong></td>
			<td align="left"><%
			if trim(rs1("IIDNO"))<>"" and not isnull(rs1("IIDNO")) then
				response.write trim(rs1("IIDNO"))
             else
				response.write "&nbsp;"
			end if
			%></td>
                <td align="right" bgcolor="#FFFF99" width="12%"><strong>出生日期</strong></td>
                <td align="left"><%
			if trim(rs1("IBIRTH"))<>"" and not isnull(rs1("IBIRTH")) then
				response.write GetDate(rs1("IBIRTH"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>違規人地址</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("IADDR"))<>"" and not isnull(rs1("IADDR")) then
				response.write trim(rs1("IZIP"))& " " &trim(rs1("IADDR"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>法條代碼<br>告發項目代碼</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("RULEF1"))<>"" and not isnull(rs1("RULEF1")) then
				response.write trim(rs1("RULEF1")) 

				sql="select RULENAME from RULEF where RULECODE ='"&rs1("RULEF1")&"'"
				set rsPCode=conn1.execute(sql)

				if not rspcode.eof then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rsPCode("RULENAME")
				set rsPcode=nothing
				response.write "<br>"&trim(rs1("RULEF1"))&"					"&trim(rs1("FACTG1"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>法條代碼<br>告發項目代碼</strong></td>	
			<td align="left" colspan="5"><%
			if trim(rs1("RULEF2"))<>"" and not isnull(rs1("RULEF2")) then
				response.write trim(rs1("RULEF2")) 

				sql="select RULENAME from rulef where RULECODE ='"&rs1("RULEF2")&"'"
				set rsPCode=conn1.execute(sql)
				if not rspcode.eof then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rsPCode("RULENAME")
				set rsPcode=nothing
				response.write "<br>"&trim(rs1("RULEF2"))&"					"&trim(rs1("FACTG2"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>法條代碼<br>告發項目代碼</strong></td>	
			<td align="left" colspan="5"><%
			if trim(rs1("RULEF3"))<>"" and not isnull(rs1("RULEF3")) then
				response.write trim(rs1("RULEF3")) 

				sql="select RULENAME from rulef where RULECODE ='"&rs1("RULEF3")&"'"
				set rsPCode=conn1.execute(sql)
				if not rspcode.eof then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rsPCode("RULENAME")
				set rsPcode=nothing

				response.write "<br>"&trim(rs1("RULEF3"))&"					"&trim(rs1("FACTG3"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>法條代碼<br>告發項目代碼</strong></td>	
			<td align="left" colspan="5"><%
			if trim(rs1("RULEF4"))<>"" and not isnull(rs1("RULEF4")) then
				response.write trim(rs1("RULEF4")) 

				sql="select RULENAME from rulef where RULECODE ='"&rs1("RULEF4")&"'"
				set rsPCode=conn1.execute(sql)
				if not rspcode.eof then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rsPCode("RULENAME")
				set rsPcode=nothing
				response.write "<br>"&trim(rs1("RULEF4"))&"					"&trim(rs1("FACTG4"))
			else
				response.write "&nbsp;"
			end if
			%></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#FFFF99"><strong>備註</strong></td>	
			<td align="left" colspan="5"><%
			if trim(rs1("Note"))<>"" and not isnull(rs1("Note")) then
				response.write trim(rs1("Note")) 
			else
				response.write "&nbsp;"			
			End if
			%></td>
		</tr>

	<tr>	
			<td colspan="6" bgcolor="#00FFFF" height="32"><b>入案資料</b></td>
     <tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>車主姓名</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("OWNAME"))<>"" and not isnull(rs1("OWNAME")) then
				response.write trim(rs1("OWNAME"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>車主地址</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("OWADDR"))<>"" and not isnull(rs1("OWADDR")) then
				response.write trim(rs1("OWZIP"))&"  " & trim(rs1("OWADDR"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>詳細車種代碼</strong></td>
			<td align="left"><%
			if trim(rs1("CDType"))<>"" and not isnull(rs1("CDType")) then
				response.write trim(rs1("CDType")) 

				sql="select CDName from CARKIND where CDType ='"&rs1("CDType")&"'"
				set rsPCode=conn1.execute(sql)
				if not rspcode.eof then response.write "&nbsp;"&rsPCode("CDName")
				set rsPcode=nothing

			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>單位代碼</strong></td>
			<td align="left" colspan="3"><%
			if trim(rs1("PBCode"))<>"" and not isnull(rs1("PBCode")) then
				response.write trim(rs1("PBCode")) 

				sql="select PBName from PBLIST where PBCode ='"&rs1("PBCode")&"'"
				set rsPCode=conn1.execute(sql)
				if not rspcode.eof then response.write "&nbsp;"&rsPCode("PBName")
				set rsPcode=nothing

			else
				response.write "&nbsp;"
			end if
			%></td>
			  <tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>告發單種類</strong></td>
			<td align="left"><%
			if trim(rs1("RBType"))<>"" and not isnull(rs1("RBType")) then	
				response.write trim(rs1("RBType")) 
				if trim(rs1("RBType")) ="1" then 
				response.write "&nbsp;電腦製單"
				elseif trim(rs1("RBType")) ="2" then 
				response.write "&nbsp;手開單"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>操作人員</strong></td>
			<td align="left" colspan="3"><%
			if trim(rs1("OPCODE"))<>"" and not isnull(rs1("OPCODE")) then
				response.write trim(rs1("OPCODE")) 

				sql="select OPName from OPER where OPCode ='"&rs1("OPCODE")&"'"
				set rsPCode=conn1.execute(sql)
				if not rspcode.eof then response.write "&nbsp;"&rsPCode("OPName")
				set rsPcode=nothing

			else
				response.write "&nbsp;"
			end if
			%></td>
			  <tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>入案日期</strong></td>
			<td align="left"><%
			if trim(rs1("FinDate"))<>"" and not isnull(rs1("FinDate")) then
				response.write GetDate(rs1("FinDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>操作日期</strong></td>
			<td align="left" colspan="3"><%
			if trim(rs1("OPDate"))<>"" and not isnull(rs1("OPDate")) then
				response.write GetDate(rs1("OPDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			  <tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>入案檔名</strong></td>
			<td align="left"><%
			if trim(rs1("batChNo"))<>"" and not isnull(rs1("batChNo")) then
				response.write rs1("batChNo") 
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>郵局日期</strong></td>
			<td align="left" colspan="3"><%
			if trim(rs1("MailDate"))<>"" and not isnull(rs1("MailDate")) then
				response.write GetDate(rs1("MailDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			  <tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>車籍狀態</strong></td>
			<td align="left"><%
			if trim(rs1("Errst_SV"))<>"" and not isnull(rs1("Errst_SV")) then
				response.write rs1("Errst_SV") 

				sql="select ERRName from ErrCode where ErrCode ='"&rs1("Errst_SV")&"'"
				set rsPCode=conn1.execute(sql)

				if not rspcode.eof then response.write "&nbsp;"&rsPCode("ERRName")
				set rsPcode=nothing

			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>掛號號碼</strong></td>
			<td align="left" colspan="3"><%
			if trim(rs1("MailSEQNO"))<>"" and not isnull(rs1("MailSEQNO")) then
				response.write rs1("MailSEQNO") 
			else
				response.write "&nbsp;"
			end if
			%></td>
			  <tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>駕籍狀態</strong></td>
			<td align="left"><%
			if trim(rs1("ErrST_SD"))<>"" and not isnull(rs1("ErrST_SD")) then
				response.write rs1("ErrST_SD") 
				sql="select ERRName from ErrCode where ErrCode ='"&rs1("ErrST_SD")&"'"
				set rsPCode=conn1.execute(sql)
				if not rspcode.eof then response.write "&nbsp;"&rsPCode("ERRName")
				set rsPcode=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td align="right" bgcolor="#FFFF99"><strong>移送日期</strong></td>
			<td align="left" colspan="3"><%
			if trim(rs1("SendPrnDate"))<>"" and not isnull(rs1("SendPrnDate")) then
				response.write GetDate(rs1("SendPrnDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			  <tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>操作日期</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("OPDate"))<>"" and not isnull(rs1("OPDate")) then
				response.write GetDate(rs1("OPDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>操作人員</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("OPCODE"))<>"" and not isnull(rs1("OPCODE")) then
				response.write trim(rs1("OPCODE")) 

				sql="select OPName from OPER where OPCode ='"&rs1("OPCODE")&"'"
				set rsPCode=conn1.execute(sql)

				if not rspcode.eof then response.write "&nbsp;"&rsPCode("OPName")
				set rsPcode=nothing

			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>單位代碼</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("PBCode"))<>"" and not isnull(rs1("PBCode")) then
				response.write trim(rs1("PBCode")) 

				sql="select PBName from PBLIST where PBCode ='"&rs1("PBCode")&"'"
				set rsPCode=conn1.execute(sql)

				if not rspcode.eof then response.write "&nbsp;"&rsPCode("PBName")
				set rsPcode=nothing

			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>法條金額１</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("AMT1_Fin"))<>"" and not isnull(rs1("AMT1_Fin")) and trim(rs1("AMT1_Fin"))<>"0" then
				response.write rs1("AMT1_Fin") 
			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>法條金額２</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("AMT2_Fin"))<>"" and not isnull(rs1("AMT2_Fin")) and trim(rs1("AMT2_Fin"))<>"0" then
				response.write rs1("AMT2_Fin") 
			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>法條金額３</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("AMT3_Fin"))<>"" and not isnull(rs1("AMT3_Fin")) and trim(rs1("AMT3_Fin"))<>"0" then
				response.write rs1("AMT3_Fin") 
			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>法條金額４</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("AMT4_Fin"))<>"" and not isnull(rs1("AMT4_Fin")) and trim(rs1("AMT4_Fin"))<>"0" then
				response.write rs1("AMT4_Fin") 
			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>違反牌照稅註記</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("TAXFLAG"))<>"" and not isnull(rs1("TAXFLAG")) then
				response.write rs1("TAXFLAG") 
				if rs1("TAXFLAG") ="0" then
				  response.write "&nbsp;正常(無違反)"
				elseif rs1("TAXFLAG") ="1" then
				  response.write "&nbsp;違反牌照稅註記"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>送達註記　</strong></td>
			<td align="left" colspan="5"><%
			if trim(rs1("PUBFLAG"))<>"" and not isnull(rs1("PUBFLAG")) then
				response.write rs1("PUBFLAG") 
				if rs1("PUBFLAG") ="1" then
				  response.write "&nbsp;公示送達"
				elseif rs1("PUBFLAG") ="2" then
				  response.write "&nbsp;寄存送達"
				elseif rs1("PUBFLAG") ="3" then
				  response.write "&nbsp;留置送達"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
	<tr>	
	<% sql="select * from FinDel where FSEQ='"&rs1("FSEQ")&"' order by Del_OPDate"
	delint=0
			    set rsDel=conn1.execute(sql)

   		while Not rsDel.eof
		delint=delint+1
	%>
			<td colspan="6" bgcolor="#00FFFF" height="32"><b>刪除資料(<%=delint%>)</b></td>
     <tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>刪除狀態</strong></td>
			<td align="left" colspan="5"><%
			if trim(rsDel("Del_FStatus"))<>"" and not isnull(rsDel("Del_FStatus")) then
				response.write trim(rsDel("Del_FStatus"))
				if rsDel("Del_FStatus") ="0" then
				  response.write "&nbsp;未上傳"
				elseif rsDel("Del_FStatus") ="2" then
				  response.write "&nbsp;已上傳"
				elseif rsDel("Del_FStatus") ="S" then
				  response.write "&nbsp;上傳成功"
				elseif rsDel("Del_FStatus") ="N" then
				  response.write "&nbsp;無此資料"
				elseif rsDel("Del_FStatus") ="n" then
				  response.write "&nbsp;已結案不做刪除"
				elseif rsDel("Del_FStatus") ="B" then
				  response.write "&nbsp;無此車號/無此證號"
				elseif rsDel("Del_FStatus") ="Z" then
				  response.write "&nbsp;不可刪除"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>刪除人員</strong></td>
			<td align="left" colspan="5"><%
			if trim(rsDel("Del_OPCode"))<>"" and not isnull(rsDel("Del_OPCode")) then
				response.write trim(rsDel("Del_OPCode"))

				sql="select OPName from OPER where OPCode ='"&rsDel("Del_OPCode")&"'"
				set rsPCode=conn1.execute(sql)

				if not rspcode.eof then response.write "&nbsp;"&rsPCode("OPName")
				set rsPcode=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>刪除日期</strong></td>
			<td align="left" colspan="5"><%
			if trim(rsDel("Del_OPDate"))<>"" and not isnull(rsDel("Del_OPDate")) then
				response.write GetDate(rsDel("Del_OPDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>刪除上傳檔名</strong></td>
			<td align="left" colspan="5"><%
			if trim(rsDel("Del_BatChNo"))<>"" and not isnull(rsDel("Del_BatChNo")) then
				response.write trim(rsDel("Del_BatChNo"))
			else
				response.write "&nbsp;"
			end if
			%></td>
<%
			rsDel.movenext
		wend
%>
<% sql="select * from FinBack where FSEQ='"&rs1("FSEQ")&"'"
'sql="select * from FinBack where FSEQ='IA5066244'"

					set rsBack=conn1.execute(sql)
   		while Not rsBack.eof
%>
			<td colspan="6" bgcolor="#00FFFF" height="32"><b>退件資料</b></td>
			<tr>
			<td colspan="6" bgcolor="#FFFF99" height="20"><b>第一次退件資料</b></td>
     <tr>
	 		<td bgcolor="#FFFF99" align="right"><strong>退件日期</strong></td>
			<td align="left"><%
			if trim(rsBack("BackDate"))<>"" and not isnull(rsBack("BackDate")) then
				response.write GetDate(rsBack("BackDate"))
			else
				response.write "&nbsp;"
			end if%>
			<td bgcolor="#FFFF99" align="right"><strong>郵寄日期</strong></td>
			<td align="left" width="17%"><%
			if trim(rsBack("MailDate"))<>"" and not isnull(rsBack("MailDate")) then
				response.write GetDate(rsBack("MailDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#FFFF99" align="right" width="12%"><strong>貼條號碼</strong></td>
			<td align="left"><%
			if trim(rsBack("MailNo"))<>"" and not isnull(rsBack("MailNo")) then
				response.write trim(rsBack("MailNo"))
			else
				response.write "&nbsp;"
			end if
			%></td>
     <tr>
	 		<td bgcolor="#FFFF99" align="right"><strong>郵寄序號</strong></td>
			<td align="left"><%
			if trim(rsBack("MailSEQNo"))<>"" and not isnull(rsBack("MailSEQNo")) then
				response.write trim(rsBack("MailSEQNo"))
			else
				response.write "&nbsp;"
			end if%>
			<td bgcolor="#FFFF99" align="right"><strong>移送日期</strong></td>
			<td align="left" width="17%"><%
			if trim(rsBack("SendPrnDate"))<>"" and not isnull(rsBack("SendPrnDate")) then
				response.write GetDate(rsBack("SendPrnDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#FFFF99" align="right" width="12%"><strong>二次郵寄日期</strong></td>
			<td align="left"><%
			if trim(rsBack("MailDate2"))<>"" and not isnull(rsBack("MailDate2")) then
				response.write GetDate(rsBack("MailDate2"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
	 			<td align="right" bgcolor="#FFFF99"><strong>退件原因</strong></td>
			<td align="left" colspan="5"><%
			if trim(rsBack("BackCode"))<>"" and not isnull(rsBack("BackCode")) then
				response.write trim(rsBack("BackCode"))
				sql="select BACKName from BACKCODE where BACKCODE ='"&rsBack("BackCode")&"'"
				set rsPCode=conn1.execute(sql)
				if not rspcode.eof then response.write "&nbsp;"&rsPCode("BACKName")
				set rsPcode=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
			<tr>
			<td colspan="6" bgcolor="#FFFF99" height="20"><b>第二次退件資料</b></td>
     <tr>
	 		<td bgcolor="#FFFF99" align="right"><strong>退件日期</strong></td>
			<td align="left"><%
			if trim(rsBack("BackDate2"))<>"" and not isnull(rsBack("BackDate2")) then
				response.write GetDate(rsBack("BackDate2"))
			else
				response.write "&nbsp;"
			end if%>
			<td bgcolor="#FFFF99" align="right" width="12%"><strong>貼條號碼</strong></td>
			<td align="left"><%
			if trim(rsBack("MailNo2"))<>"" and not isnull(rsBack("MailNo2")) then
				response.write trim(rsBack("MailNo2"))
			else
				response.write "&nbsp;"
			end if
			%></td>
			<td bgcolor="#FFFF99" align="right"><strong>送達日期</strong></td>
			<td align="left" width="17%"><%
			if trim(rsBack("ShowDate"))<>"" and not isnull(rsBack("ShowDate")) then
				response.write GetDate(rsBack("ShowDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
     <tr>
	 		<td bgcolor="#FFFF99" align="right"><strong>送達完成日期</strong></td>
			<td align="left"><%
			if trim(rsBack("CloseDate"))<>"" and not isnull(rsBack("CloseDate")) then
				response.write GetDate(rsBack("CloseDate"))
			else
				response.write "&nbsp;"
			end if%>
			<td bgcolor="#FFFF99" align="right" width="12%"><strong>退件原因</strong></td>
			<td align="left" colspan="3"><%
			if trim(rsBack("BackCode2"))<>"" and not isnull(rsBack("BackCode2")) then
				response.write trim(rsBack("BackCode2"))
				sql="select BACKName from BACKCODE where BACKCODE ='"&rsBack("BackCode2")&"'"
				set rsPCode=conn1.execute(sql)
				if not rspcode.eof then response.write "&nbsp;"&rsPCode("BACKName")
				set rsPcode=nothing
			else
				response.write "&nbsp;"
			end if
			%></td>
    <tr>
			<td colspan="6" bgcolor="#00FFFF" height="32"><b>退件／送達上傳資料</b></td>
     <tr>
 			<td align="right" bgcolor="#FFFF99"><strong>資料類別</strong></td>
			<td align="left" colspan="3"><%
			if trim(rsBack("PubType"))<>"" and not isnull(rsBack("PubType")) then
				response.write trim(rsBack("PubType"))
				if rsBack("PubType") ="1" then
				  response.write "&nbsp;公示送達"
				elseif rsBack("PubType") ="2" then
				  response.write "&nbsp;寄存送達"
				end if
			else
				response.write "&nbsp;"
			end if
			%></td>
 			<td align="right" bgcolor="#FFFF99"><strong>操作日期</strong></td>
			<td align="left"><%
			if trim(rsBack("Opdate"))<>"" and not isnull(rsBack("Opdate")) then
				response.write GetDate(rsBack("Opdate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
	<tr>
 			<td align="right" bgcolor="#FFFF99"><strong>退件上傳註記</strong></td>
			<td align="left" colspan="5"><%
			if trim(rsBack("FStatus"))<>"" and not isnull(rsBack("FStatus")) then
				response.write trim(rsBack("FStatus"))&"&nbsp;"&getDciCodeN(rsBack("FStatus"))
			else
				response.write "&nbsp;"
			end if
			%></td>
	<tr>
 			<td align="right" bgcolor="#FFFF99"><strong>退件上傳檔名</strong></td>
			<td align="left" colspan="5"><%
			if trim(rsBack("BatChNo"))<>"" and not isnull(rsBack("BatChNo")) then
				response.write trim(rsBack("BatChNo"))
			else
				response.write "&nbsp;"
			end if
			%></td>
	<tr>
 			<td align="right" bgcolor="#FFFF99"><strong>送達上傳註記</strong></td>
			<td align="left" colspan="5"><%
			if trim(rsBack("FStatus_P"))<>"" and not isnull(rsBack("FStatus_P")) then
				response.write trim(rsBack("FStatus_P"))&"&nbsp;"&getDciCodeN(rsBack("FStatus_P"))
			else
				response.write "&nbsp;"
			end if
			%></td>
	<tr>
 			<td align="right" bgcolor="#FFFF99"><strong>送達上傳檔名</strong></td>
			<td align="left"><%
			if trim(rsBack("BatChNo_P"))<>"" and not isnull(rsBack("BatChNo_P")) then
				response.write trim(rsBack("BatChNo_P"))
			else
				response.write "&nbsp;"
			end if
			%></td>
 			<td align="right" bgcolor="#FFFF99"><strong>送達書號</strong></td>
			<td align="left"><%
			if trim(rsBack("SEQNO"))<>"" and not isnull(rsBack("SEQNO")) then
				response.write trim(rsBack("SEQNO"))
			else
				response.write "&nbsp;"
			end if
			%></td>
 			<td align="right" bgcolor="#FFFF99"><strong>送達生效日</strong></td>
			<td align="left"><%
			if trim(rsBack("PubDate"))<>"" and not isnull(rsBack("PubDate")) then
				response.write GetDate(rsBack("PubDate"))
			else
				response.write "&nbsp;"
			end if
			%></td>
<%
			rsBack.movenext
		wend
%>
		</table>
</body>
<script language="javascript">
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
		var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
		win.focus();
		return win;
}
function OpenImageWin(ImgFileName){
	urlstr=ImgFileName;
	newWin(urlstr,'MyDetail',1000,600,0,0,"yes","no","yes","no");
}
</script>
</html>

<%
				if sys_City<>"台中縣" and sys_City<>"台中市" and sys_City<>"南投縣" and sys_City<>"基隆市" And sys_City<>"澎湖縣" Then
					conn1.close
				else
					conn.close
				end if
%>