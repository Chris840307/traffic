<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
fname="station.txt"
Response.AddHeader "Content-Disposition","attachment;filename=" &fname
Response.ContentType = "application/vnd.ms-txt"
Server.ScriptTimeout=60000
DateType="IllegalDate"
illegalDate1="2011/01/01 00:00:00"
illegalDate2="2011/12/31 23:59:59"

strQuery="select a.BillTypeName,a.IllegalDate,a.BillNo,a.CarNo,a.CarSimpleName,b.DriverID,b.Driver,b.Owner,b.DciReturnCarName,a.BillMem,d.UnitName,a.IllegalAddress,a.Rule1,a.Rule2,b.FORFEIT1,b.FORFEIT2,a.BillFillDate,a.DeallineDate,b.DCISTATIONNAME,a.Recorddate,c.FastenerTypeName,a.RecordMem from (select SN,decode(billtypeid,1,'攔停','逕舉') BillTypeName,Illegaldate,BillNo,CarNo,Decode(CarSimpleID,1,'汽車',2,'拖車',3,'重機',4,'輕機',6,'臨時車牌','其它') CarSimpleName,(Select ChName from memberdata where MemberID=BillBase.BillMemID1) BillMem,BillUnitID,IllegalAddress,Rule1,Rule2,BillFillDate,DeallineDate,Recorddate,(Select ChName from memberdata where MemberID=BillBase.RecordMemberID) RecordMem from billbase where "&DateType&" between to_date('"&illegalDate1&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&illegalDate2&"','YYYY/MM/DD/HH24/MI/SS') and Recordstateid=0 and billno is not null) a,(select BillNo,CarNo,DriverID,Driver,Owner,(select Content from DCICODE where ID=BillBaseDcireturn.DciReturnCarType and TypeID=5) DciReturnCarName,FORFEIT1,FORFEIT2,(select DCISTATIONNAME from Station where DCIStationID=BillBaseDcireturn.DCIReturnStation) DCISTATIONNAME from BillBaseDcireturn where ExchangetypeID='W') b,(select Max((select Content from DciCode where TypeID=6 and ID=BillFastenerDetail.FastenerTypeID)) FastenerTypeName,BillSN from BillFastenerDetail group by BillSN) c,UnitInfo d where a.BillNo=b.BillNo(+) and a.CarNo=b.CarNo(+) and a.SN=c.BillSN(+) and a.BillUnitID=d.UnitID"

strQuery=strQuery&" Union all "

strQuery=strQuery&"select '行人慢車' BillTypeName,IllegalDate,BillNo,CarNo,''CarSimpleName,DriverID,Driver,'' Owner,null DciReturnCarName,(select chName from MemberData where MemberID=PasserBase.BillMemID1) BillMem,(Select UnitName from Unitinfo where UnitID=PasserBase.BillUnitID) UnitName,IllegalAddress,Rule1,Rule2,Forfeit1,Forfeit2,BillFillDate,DeallineDate,(Select UnitName from UnitInfo where UnitID=PasserBase.MemberStation) DCISTATIONNAME,RecordDate,null FastenerTypeName,(Select ChName from memberdata where MemberID=PasserBase.RecordMemberID) RecordMem from PasserBase where "&DateType&" between to_date('"&illegalDate1&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&illegalDate2&"','YYYY/MM/DD/HH24/MI/SS') and Recordstateid=0 and billno is not null order by BillTypeName,BillNo"

set rsfound=conn.execute(strQuery)

Response.Write "類別,違規日期,違規時間,舉發單號,車號,簡式車種,駕駛人ID,駕駛人姓名,車主姓名,詳細車種,舉發員警,舉發單位,違規地點,法條一,法條二,罰款一,罰款二,填單日期,應到案日期,應到案處所,建檔日期,入案日期,代保管物件,操作人員"&vbnewline

While Not rsfound.Eof
	response.write rsfound("BillTypeName")&","
	response.write gInitDT(rsfound("IllegalDate"))&","
	response.write right("0"&hour(rsfound("IllegalDate")),2)
	Response.Write right("0"&minute(rsfound("IllegalDate")),2)&","
	response.write rsfound("BillNo")&","
	response.write rsfound("CarNo")&","
	response.write rsfound("CarSimpleName")&","
	response.write rsfound("DriverID")&","
	response.write rsfound("Driver")&","
	response.write rsfound("Owner")&","
	response.write rsfound("DciReturnCarName")&","
	response.write rsfound("BillMem")&","
	response.write rsfound("UnitName")&","
	response.write rsfound("IllegalAddress")&","
	response.write rsfound("Rule1")&","
	response.write rsfound("Rule2")&","
	response.write rsfound("FORFEIT1")&","
	response.write rsfound("FORFEIT2")&","
	response.write gInitDT(rsfound("BillFillDate"))&","
	response.write gInitDT(rsfound("DeallineDate"))&","
	response.write rsfound("DCISTATIONNAME")&","
	response.write gInitDT(rsfound("Recorddate"))&","
	response.write gInitDT(rsfound("Recorddate"))&","
	response.write rsfound("FastenerTypeName")&","
	response.write rsfound("RecordMem")&vbnewline
	response.flush
	rsfound.MoveNext
Wend
rsfound.close
set rsfound=nothing
conn.close
set conn=nothing
%>