<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
' 檔案名稱： getDelBillNoToList.asp
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	BillSn=0
	BillContent=""
	ErrorCode=0
	strBill="select * from BillBaseView where BillNo='"&trim(request("sysBillNo"))&"' and RecordStateID=0"
	set rsfound=conn.execute(strBill)
	if not rsfound.eof then
		if (checkIsAllowDel(sys_City,trim(rsfound("BillTypeID")))=true) or (trim(rsfound("imagefilenameb"))<>"") or ( (Instr(rsfound("Rule1"),"56")>0) and (Instr(rsfound("Note"),"txt")>0) and (sys_City="花蓮縣") ) then
			if trim(rsfound("BillBaseTypeID"))="0" then
				BillSn="0-"&trim(rsfound("Sn"))
				BillContent=trim(rsfound("BillNo"))&" / "&trim(rsfound("CarNo"))&" / "&gInitDT(trim(rsfound("IllegalDate")))&" "&right("00"&hour(rsfound("IllegalDate")),2)&":"&right("00"&Minute(rsfound("IllegalDate")),2)
				
				'檢查是不是有上傳入案還未回傳的
				strCType2="select DciReturnStatusID from DciLog a where a.BillSn='"&trim(trim(rsfound("Sn")))&"' and ExchangeTypeID='W' Order by ExchangeDate Desc"
				set rsCType2=conn.execute(strCType2)
				if not rsCType2.eof then
					if isnull(rsCType2("DciReturnStatusID")) or trim(rsCType2("DciReturnStatusID"))="" then
						ErrorCode=3	'上傳入案還沒有回傳
					end if
				end if
				rsCType2.close
				set rsCType2=nothing
			else
				BillSn="1-"&trim(rsfound("Sn"))
				BillContent=trim(rsfound("BillNo"))&" / "&gInitDT(trim(rsfound("IllegalDate")))&" "&right("00"&hour(rsfound("IllegalDate")),2)&":"&right("00"&Minute(rsfound("IllegalDate")),2)
			end if
		else
			ErrorCode=2 '沒有刪除的權限
		end if
	else
		ErrorCode=1 '找不到未刪除的舉發單
	end if
	rsfound.close
	set rsfound=nothing
%>
addDelBillNoToList('<%=BillSn%>','<%=BillContent%>','<%=ErrorCode%>');
<%	
conn.close
set conn=nothing
%>
