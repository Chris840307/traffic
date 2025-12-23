<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->

<%
Server.ScriptTimeout=12000
log_start=now
ConnExecute Request.ServerVariables ("SCRIPT_NAME")&"||"& now &"||"& Request("startDate_q") & "~" & Request("endDate_q") ,362
hasDate = False
strDate=request("strDate")
UserId = Session("User_ID")
ReportId = "REPORT0017Plus"
rptHead1 = Trim(Request("rptHead1"))
rptHead2 = Trim(Request("rptHead2"))
startDate_q = Trim(Request("startDate_q"))
endDate_q = Trim(Request("endDate_q"))
unitList=trim(request("unitList"))
MemberList=trim(request("MemberList"))
ReportName=Request("rptHead2")
PageType=trim(request("PageType"))

strRul="select Value from Apconfigure where ID=3"
set rsRul=conn.execute(strRul)
RuleVer=trim(rsRul("Value"))
rsRul.Close

strRul="select Value from Apconfigure where ID=2"
set rsRul=conn.execute(strRul)
chkBillno=left(rsRul("Value"),1)
rsRul.Close

Conn.BeginTrans
   sqlDel = "Delete From UserRptInfo Where UserId=" & UserId & " And ReportId='" & ReportId & "' and ReportName is null"
   Conn.Execute(sqlDel)
   sqlDel = "Delete From UserRptInfo Where UserId=" & UserId & " And ReportId='" & ReportId & "' and ReportName='"&ReportName&"'"
   Conn.Execute(sqlDel)
   sqlDel = "Delete From UserLawInfo Where UserId=" & UserId & " And ReportId='" & ReportId & "' "
   Conn.Execute(sqlDel)
   sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType,ReportName) Values (" & UserId & ",'" & ReportId & "','rptHead1','" & rptHead1 & "','TEXT','"&ReportName&"')"
   Conn.Execute(sqlAdd)
   sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType,ReportName) Values (" & UserId & ",'" & ReportId & "','rptHead2','" & rptHead2 & "','TEXT','"&ReportName&"')"
   Conn.Execute(sqlAdd)

   If startDate_q <> "" And endDate_q <> "" Then
   	  hasDate = True
      sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType,ReportName) Values (" & UserId & ",'" & ReportId & "','startDate_q','" & startDate_q & "','TEXT','"&ReportName&"')"
      Conn.Execute(sqlAdd)   
      sqlAdd = "Insert Into UserRptInfo (UserId,ReportId,FieldName,FieldValue,FieldType,ReportName) Values (" & UserId & ",'" & ReportId & "','endDate_q','" & endDate_q & "','TEXT','"&ReportName&"')"
      Conn.Execute(sqlAdd)         	
   End If

   if err.number = 0 then
   	 Conn.CommitTrans
   else    	
     Conn.RollbackTrans
   end if
   
   tmpSql = ""
   If hasDate Then
   	  tmpSql = tmpSql & " And "&strDate&" Between To_Date('" & gOutDT(startDate_q)&" 0:0:0" & "','YYYY/MM/DD/HH24/MI/SS') And To_Date('" & gOutDT(endDate_q)&" 23:59:59" & "','YYYY/MM/DD/HH24/MI/SS')"
   End If
	
	If not ifnull(MemberList) Then
		Sys_MemberID = replace(MemberList,"~",",")

		Mem1 = " and BillMemID1 in (select MemberID from MemberData where chName in(select chName from MemberData where MemberID in(" & Sys_MemberID & ")) and UnitID in(select UnitID from MemberData where MemberID in(" & Sys_MemberID & ")))"
		Mem2 = " and BillMemID2 in (select MemberID from MemberData where chName in(select chName from MemberData where MemberID in(" & Sys_MemberID & ")) and UnitID in(select UnitID from MemberData where MemberID in(" & Sys_MemberID & ")))"
		Mem3 = " and BillMemID3 in (select MemberID from MemberData where chName in(select chName from MemberData where MemberID in(" & Sys_MemberID & ")) and UnitID in(select UnitID from MemberData where MemberID in(" & Sys_MemberID & ")))"
		Mem4 = " and BillMemID4 in (select MemberID from MemberData where chName in(select chName from MemberData where MemberID in(" & Sys_MemberID & ")) and UnitID in(select UnitID from MemberData where MemberID in(" & Sys_MemberID & ")))"
	elseif not ifnull(unitList) Then
		Sys_UnitList = replace(unitList,"~","','")

		Mem1 = " and BillMemID1 in (select MemberID from MemberData where unitid in('" & Sys_UnitList & "'))"
		Mem2 = " and BillMemID2 in (select MemberID from MemberData where unitid in('" & Sys_UnitList & "'))"
		Mem3 = " and BillMemID3 in (select MemberID from MemberData where unitid in('" & Sys_UnitList & "'))"
		Mem4 = " and BillMemID4 in (select MemberID from MemberData where unitid in('" & Sys_UnitList & "'))"
	End if 
	

qsql=" Group By BillMemID,Other3,Rule,BillTypeID,BillScore"
qsq2=" Group By BillMemID,Other3,Rule,BillTypeID"
tmpStr = ""

dim tempTable(11)
dim strtab_a(3,3,3)
dim billunitID(10):dim Other3(10):dim billtypeid(10):dim BillMemID(10):dim LoginID(10):dim ChName(10):dim UnitName(10):dim BillScore(10):dim Rule(10):dim BillTotal(10)
chkBillMemID="":chkUnitName="":chkLoginID="":chkchName="":chkRule=""

otherScore=" and Other3 not in('砂石車違規','計程車違規','違規拖吊') and Not(LawItem like '53100%' and Other3='違反路權規定')"

otherRule1=" and not (NVL(CarAddID,0) in(1,2,3,4,6) and rule1 in(select LawItem From LawScore where CountyOrnpa="&PageType &" and Other3='砂石車違規')) and not (NVL(CarAddID,0)=10 and rule1 in(select LawItem From LawScore where CountyOrnpa="&PageType &" and Other3='計程車違規')) and not (NVL(CarAddID,0)=8 and rule1 in(select LawItem From LawScore where CountyOrnpa="&PageType &" and Other3='違規拖吊')) and not (Rule1 like '53100%' and (BillTypeID=1 or UseTool=8))"

otherRule2=" and not (NVL(CarAddID,0) in(1,2,3,4,6) and rule2 in(select LawItem From LawScore where CountyOrnpa="&PageType &" and Other3='砂石車違規')) and not (NVL(CarAddID,0)=10 and rule2 in(select LawItem From LawScore where CountyOrnpa="&PageType &" and Other3='計程車違規')) and not (NVL(CarAddID,0)=8 and rule2 in(select LawItem From LawScore where CountyOrnpa="&PageType &" and Other3='違規拖吊')) and not (Rule2 like '53100%' and (BillTypeID=1 or UseTool=8))"

otherRule3=" and not (NVL(CarAddID,0) in(1,2,3,4,6) and rule3 in(select LawItem From LawScore where CountyOrnpa="&PageType &" and Other3='砂石車違規')) and not (NVL(CarAddID,0)=10 and rule3 in(select LawItem From LawScore where CountyOrnpa="&PageType &" and Other3='計程車違規')) and not (NVL(CarAddID,0)=8 and rule3 in(select LawItem From LawScore where CountyOrnpa="&PageType &" and Other3='違規拖吊')) and not (Rule3 like '53100%' and (BillTypeID=1 or UseTool=8))"

otherRule4=" and not (NVL(CarAddID,0) in(1,2,3,4,6) and rule4 in(select LawItem From LawScore where CountyOrnpa="&PageType &" and Other3='砂石車違規')) and not (NVL(CarAddID,0)=10 and rule4 in(select LawItem From LawScore where CountyOrnpa="&PageType &" and Other3='計程車違規')) and not (NVL(CarAddID,0)=8 and rule4 in(select LawItem From LawScore where CountyOrnpa="&PageType &" and Other3='違規拖吊')) and not (Rule4 like '53100%' and (BillTypeID=1 or UseTool=8))"

BillbaseViewqry="select CarAddID,rule1,rule2,rule3,nvl2(translate(rule4,'\1234567890','\'),'',rule4) rule4,BillTypeID,BillMemID1,BillMemID2,BillMemID3,BillMemID4,UseTool from billbase where BillNo like '"&chkBillno&"%' and recordstateid=0"& tmpSql &" union all select CarAddID,rule1,rule2,rule3,nvl2(translate(rule4,'\1234567890','\'),'',rule4) rule4,BillTypeID,BillMemID1,BillMemID2,BillMemID3,BillMemID4,UseTool from PasserBase where BillNo like '"&chkBillno&"%' and recordstateid=0"& tmpSql

strtab_a(0,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is null" & otherRule1 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(1,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is null" & otherRule2 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(2,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is null" & otherRule3 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(3,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is null" & otherRule4 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(0,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is not null and BillMemID3 is null" & otherRule1 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(1,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is not null and BillMemID3 is null" & otherRule2 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(2,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is not null and BillMemID3 is null" & otherRule3 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(3,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is not null and BillMemID3 is null" & otherRule4 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql


strtab_a(0,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null" & otherRule1 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(1,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null" & otherRule2 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(2,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null" & otherRule3 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(3,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null" & otherRule4 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(0,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null" & otherRule1 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(1,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null" & otherRule2 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(2,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null" & otherRule3 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(3,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null" & otherRule4 & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(0,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is not null and BillMemID3 is null" & otherRule1 & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(1,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is not null and BillMemID3 is null" & otherRule2 & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(2,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is not null and BillMemID3 is null" & otherRule3 & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(3,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is not null and BillMemID3 is null" & otherRule4 & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(0,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null" & otherRule1 & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(1,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null" & otherRule2 & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(2,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null" & otherRule3 & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(3,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null" & otherRule4 & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(0,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null" & otherRule1 & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(1,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null" & otherRule2 & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(2,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null" & otherRule3 & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(3,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null" & otherRule4 & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(0,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null" & otherRule1 & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(1,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null" & otherRule2 & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql


strtab_a(2,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null" & otherRule3 & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql


strtab_a(3,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null" & otherRule4 & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql


strtab_a(0,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null" & otherRule1 & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql


strtab_a(1,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null" & otherRule2 & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql


strtab_a(2,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null" & otherRule3 & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql


strtab_a(3,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null" & otherRule4 & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql


strtab_a(0,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null" & otherRule1 & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(1,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null" & otherRule2 & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(2,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null" & otherRule3 & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_a(3,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null" & otherRule4 & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawVerSion="&RuleVer& otherScore &") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

tempTable(0)=""

for i=0 to ubound(strtab_a,1)
	for j=0 to Ubound(strtab_a,2)
		for k=0 to Ubound(strtab_a,3)
			if trim(strtab_a(i,j,k))<>"" then
				if trim(tempTable(0))<>"" then tempTable(0)=tempTable(0) & " Union all "
				tempTable(0)=tempTable(0) & strtab_a(i,j,k)
			end if
		next
	next
next

strTable="select BillMemID,Other3,Rule,BillTypeID,sum(TotalNum) TotalNum,sum(TotalSum) TotalSum from ("&tempTable(0)&")"&qsq2
sql1 = "Select BillCount.*,Decode(BillCount.Other3,'違反路權規定','01','違反行人路權','02','砂石車違規','03','計程車違規','04','高速公路違規','05','清除道路障礙','06','違規拖吊','07','其他交通重點違規','08','09') OtherOrder,MemberData.ChName,MemberData.LoginID,MemberData.UnitID,UnitInfo.UnitName from ("&strTable&") BillCount,MemberData,UnitInfo where BillCount.BillMemID=MemberData.MemberID and MemberData.UnitID=UnitInfo.UnitID"
sql = sql1 & " order by UnitID,LoginID,OtherOrder,Other3,Rule,BillTypeID"

Set RsTemp=conn.execute(sql)

while Not RsTemp.eof
	if trim(RsTemp("BillMemID"))<>"" and trim(RsTemp("UnitID"))<>"" then
		if BillMemID(0)<>"" then
			billunitID(0)=billunitID(0)&"||"
			billtypeid(0)=billtypeid(0)&"||"
			BillMemID(0)=BillMemID(0)&"||"
			LoginID(0)=LoginID(0)&"||"
			ChName(0)=ChName(0)&"||"
			UnitName(0)=UnitName(0)&"||"
			BillScore(0)=BillScore(0)&"||"
			BillTotal(0)=BillTotal(0)&"||"
			Rule(0)=Rule(0)&"||"
			Other3(0)=Other3(0)&"||"
		end if
		billunitID(0)=billunitID(0)&RsTemp("UnitID")
		billtypeid(0)=billtypeid(0)&RsTemp("billtypeid")
		BillMemID(0)=BillMemID(0)&RsTemp("BillMemID")
		LoginID(0)=LoginID(0)&RsTemp("LoginID")
		ChName(0)=ChName(0)&RsTemp("ChName")
		UnitName(0)=UnitName(0)&RsTemp("UnitName")
		BillScore(0)=BillScore(0)&RsTemp("TotalSum")
		BillTotal(0)=BillTotal(0)&RsTemp("TotalNum")
		Rule(0)=Rule(0)&RsTemp("Rule")
		Other3(0)=Other3(0)&RsTemp("Other3")
	end if
	RsTemp.movenext
wend
RsTemp.close

billunitID(0)=split(billunitID(0),"||")
billtypeid(0)=split(billtypeid(0),"||")
BillMemID(0)=split(BillMemID(0),"||")
LoginID(0)=split(LoginID(0),"||")
ChName(0)=split(ChName(0),"||")
UnitName(0)=split(UnitName(0),"||")
BillScore(0)=split(BillScore(0),"||")
BillTotal(0)=split(BillTotal(0),"||")
Rule(0)=split(Rule(0),"||")
Other3(0)=split(Other3(0),"||")

For i=0 to Ubound(BillMemID(0))
	If chkBillMemID<>"" Then
		chkBillMemID=chkBillMemID&","
		chkUnitName=chkUnitName&","
		chkLoginID=chkLoginID&","
		chkchName=chkchName&","
		chkRule=chkRule&","
	end if
	chkBillMemID=chkBillMemID&BillMemID(0)(i)
	chkUnitName=chkUnitName&UnitName(0)(i)
	chkLoginID=chkLoginID&LoginID(0)(i)
	chkchName=chkchName&ChName(0)(i)
	chkRule=chkRule&Rule(0)(i)
Next

'==================================================================================================================

dim strtab_b(3,3,3)

strtab_b(0,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(1,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(2,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(3,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(0,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is not null and BillMemID3 is null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(1,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is not null and BillMemID3 is null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(2,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is not null and BillMemID3 is null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(3,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is not null and BillMemID3 is null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql


strtab_b(0,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(1,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(2,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(3,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(0,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(1,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(2,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(3,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and (BillTypeID=1 or UseTool=8) and NVL(CarAddID,0) not in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(0,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(1,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(2,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(3,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(0,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(1,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(2,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(3,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(0,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(1,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(2,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(3,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(0,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(1,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(2,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(3,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(0,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(1,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(2,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(3,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(0,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(1,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(2,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_b(3,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0) not in(1,2,3,4,6) and (BillTypeID=1 or UseTool=8)" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and LawItem like '53100%' and Other3='違反路權規定' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

tempTable(1)=""
for i=0 to ubound(strtab_b,1)
	for j=0 to Ubound(strtab_b,2)
		for k=0 to Ubound(strtab_b,3)
			if trim(strtab_b(i,j,k))<>"" then
				if trim(tempTable(1))<>"" then tempTable(1)=tempTable(1) & " Union all "
				tempTable(1)=tempTable(1) & strtab_b(i,j,k)
			end if
		next
	next
next

strTable="select BillMemID,Other3,Rule,BillTypeID,sum(TotalNum) TotalNum,sum(TotalSum) TotalSum from ("&tempTable(1)&")"&qsq2
sql1 = "Select BillCount.*,Decode(BillCount.Other3,'違反路權規定','01','違反行人路權','02','砂石車違規','03','計程車違規','04','高速公路違規','05','清除道路障礙','06','違規拖吊','07','其他交通重點違規','08','09') OtherOrder,MemberData.ChName,MemberData.LoginID,MemberData.UnitID,UnitInfo.UnitName from ("&strTable&") BillCount,MemberData,UnitInfo where BillCount.BillMemID=MemberData.MemberID and MemberData.UnitID=UnitInfo.UnitID"
sql = sql1 & " order by UnitID,LoginID,OtherOrder,Other3,Rule,BillTypeID"

Set RsTemp=conn.execute(sql)

while Not RsTemp.eof
	if trim(RsTemp("BillMemID"))<>"" and trim(RsTemp("UnitID"))<>"" then
		if BillMemID(1)<>"" then
			billunitID(1)=billunitID(1)&"||"
			billtypeid(1)=billtypeid(1)&"||"
			BillMemID(1)=BillMemID(1)&"||"
			LoginID(1)=LoginID(1)&"||"
			ChName(1)=ChName(1)&"||"
			UnitName(1)=UnitName(1)&"||"
			BillScore(1)=BillScore(1)&"||"
			BillTotal(1)=BillTotal(1)&"||"
			Rule(1)=Rule(1)&"||"
			Other3(1)=Other3(1)&"||"
		end if
		billunitID(1)=billunitID(1)&RsTemp("UnitID")
		billtypeid(1)=billtypeid(1)&RsTemp("billtypeid")
		BillMemID(1)=BillMemID(1)&RsTemp("BillMemID")
		LoginID(1)=LoginID(1)&RsTemp("LoginID")
		ChName(1)=ChName(1)&RsTemp("ChName")
		UnitName(1)=UnitName(1)&RsTemp("UnitName")
		BillScore(1)=BillScore(1)&RsTemp("TotalSum")
		BillTotal(1)=BillTotal(1)&RsTemp("TotalNum")
		Rule(1)=Rule(1)&RsTemp("Rule")
		Other3(1)=Other3(1)&RsTemp("Other3")
	end if
	RsTemp.movenext
wend
RsTemp.close
billunitID(1)=split(billunitID(1),"||")
billtypeid(1)=split(billtypeid(1),"||")
BillMemID(1)=split(BillMemID(1),"||")
LoginID(1)=split(LoginID(1),"||")
ChName(1)=split(ChName(1),"||")
UnitName(1)=split(UnitName(1),"||")
BillScore(1)=split(BillScore(1),"||")
BillTotal(1)=split(BillTotal(1),"||")
Rule(1)=split(Rule(1),"||")
Other3(1)=split(Other3(1),"||")

For i=0 to Ubound(BillMemID(1))
	If chkBillMemID<>"" Then
		chkBillMemID=chkBillMemID&","
		chkUnitName=chkUnitName&","
		chkLoginID=chkLoginID&","
		chkchName=chkchName&","
		chkRule=chkRule&","
	end if
	chkBillMemID=chkBillMemID&BillMemID(1)(i)
	chkUnitName=chkUnitName&UnitName(1)(i)
	chkLoginID=chkLoginID&LoginID(1)(i)
	chkchName=chkchName&ChName(1)(i)
	chkRule=chkRule&Rule(1)(i)
Next

'=================================================================================================================


dim strtab_c(3,3,3)

strtab_c(0,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(1,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(2,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(3,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(0,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(1,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(2,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(3,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql


strtab_c(0,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(1,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(2,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(3,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(0,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(1,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(2,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(3,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(0,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(1,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(2,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(3,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(0,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(1,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(2,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(3,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(0,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(1,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(2,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(3,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(0,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(1,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(2,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(3,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(0,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(1,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(2,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(3,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(0,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(1,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(2,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_c(3,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0) in(1,2,3,4,6)" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='砂石車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

tempTable(2)=""
for i=0 to ubound(strtab_c,1)
	for j=0 to Ubound(strtab_c,2)
		for k=0 to Ubound(strtab_c,3)
			if trim(strtab_c(i,j,k))<>"" then
				if trim(tempTable(2))<>"" then tempTable(2)=tempTable(2) & " Union all "
				tempTable(2)=tempTable(2) & strtab_c(i,j,k)
			end if
		next
	next
next

strTable="select BillMemID,Other3,Rule,BillTypeID,sum(TotalNum) TotalNum,sum(TotalSum) TotalSum from ("&tempTable(2)&")"&qsq2

sql1 = "Select BillCount.*,Decode(BillCount.Other3,'違反路權規定','01','違反行人路權','02','砂石車違規','03','計程車違規','04','高速公路違規','05','清除道路障礙','06','違規拖吊','07','其他交通重點違規','08','09') OtherOrder,MemberData.ChName,MemberData.LoginID,MemberData.UnitID,UnitInfo.UnitName from ("&strTable&") BillCount,MemberData,UnitInfo where BillCount.BillMemID=MemberData.MemberID and MemberData.UnitID=UnitInfo.UnitID"

sql = sql1 & " order by UnitID,LoginID,OtherOrder,Other3,Rule,BillTypeID"
Set RsTemp=conn.execute(sql)

while Not RsTemp.eof
	if trim(RsTemp("BillMemID"))<>"" and trim(RsTemp("UnitID"))<>"" then
		if BillMemID(2)<>"" then
			billunitID(2)=billunitID(2)&"||"
			billtypeid(2)=billtypeid(2)&"||"
			BillMemID(2)=BillMemID(2)&"||"
			LoginID(2)=LoginID(2)&"||"
			ChName(2)=ChName(2)&"||"
			UnitName(2)=UnitName(2)&"||"
			BillScore(2)=BillScore(2)&"||"
			BillTotal(2)=BillTotal(2)&"||"
			Rule(2)=Rule(2)&"||"
			Other3(2)=Other3(2)&"||"
		end if
		billunitID(2)=billunitID(2)&RsTemp("UnitID")
		billtypeid(2)=billtypeid(2)&RsTemp("billtypeid")
		BillMemID(2)=BillMemID(2)&RsTemp("BillMemID")
		LoginID(2)=LoginID(2)&RsTemp("LoginID")
		ChName(2)=ChName(2)&RsTemp("ChName")
		UnitName(2)=UnitName(2)&RsTemp("UnitName")
		BillScore(2)=BillScore(2)&RsTemp("TotalSum")
		BillTotal(2)=BillTotal(2)&RsTemp("TotalNum")
		Rule(2)=Rule(2)&RsTemp("Rule")
		Other3(2)=Other3(2)&RsTemp("Other3")
	end if
	RsTemp.movenext
wend
RsTemp.close
billunitID(2)=split(billunitID(2),"||")
billtypeid(2)=split(billtypeid(2),"||")
BillMemID(2)=split(BillMemID(2),"||")
LoginID(2)=split(LoginID(2),"||")
ChName(2)=split(ChName(2),"||")
UnitName(2)=split(UnitName(2),"||")
BillScore(2)=split(BillScore(2),"||")
BillTotal(2)=split(BillTotal(2),"||")
Rule(2)=split(Rule(2),"||")
Other3(2)=split(Other3(2),"||")

For i=0 to Ubound(BillMemID(2))
	If chkBillMemID<>"" Then
		chkBillMemID=chkBillMemID&","
		chkUnitName=chkUnitName&","
		chkLoginID=chkLoginID&","
		chkchName=chkchName&","
		chkRule=chkRule&","
	end if
	chkBillMemID=chkBillMemID&BillMemID(2)(i)
	chkUnitName=chkUnitName&UnitName(2)(i)
	chkLoginID=chkLoginID&LoginID(2)(i)
	chkchName=chkchName&ChName(2)(i)
	chkRule=chkRule&Rule(2)(i)
Next

'=================================================================================================================

dim strtab_d(3,3,3)

strtab_d(0,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(1,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(2,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(3,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(0,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(1,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(2,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(3,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql


strtab_d(0,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(1,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(2,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(3,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(0,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(1,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(2,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(3,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(0,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=10" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(1,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=10" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(2,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=10" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(3,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=10" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(0,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=10" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(1,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=10" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(2,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=10" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(3,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=10" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(0,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(1,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(2,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(3,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(0,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=10" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(1,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=10" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(2,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=10" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(3,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=10" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(0,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(1,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(2,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(3,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(0,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(1,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(2,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_d(3,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0)=10" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='計程車違規' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

tempTable(3)=""
for i=0 to ubound(strtab_d,1)
	for j=0 to Ubound(strtab_d,2)
		for k=0 to Ubound(strtab_d,3)
			if trim(strtab_d(i,j,k))<>"" then
				if trim(tempTable(3))<>"" then tempTable(3)=tempTable(3) & " Union all "
				tempTable(3)=tempTable(3) & strtab_d(i,j,k)
			end if
		next
	next
next

strTable="select BillMemID,Other3,Rule,BillTypeID,sum(TotalNum) TotalNum,sum(TotalSum) TotalSum from ("&tempTable(3)&")"&qsq2

sql1 = "Select BillCount.*,Decode(BillCount.Other3,'違反路權規定','01','違反行人路權','02','砂石車違規','03','計程車違規','04','高速公路違規','05','清除道路障礙','06','違規拖吊','07','其他交通重點違規','08','09') OtherOrder,MemberData.ChName,MemberData.LoginID,MemberData.UnitID,UnitInfo.UnitName from ("&strTable&") BillCount,MemberData,UnitInfo where BillCount.BillMemID=MemberData.MemberID and MemberData.UnitID=UnitInfo.UnitID"

sql = sql1 & " order by UnitID,LoginID,OtherOrder,Other3,Rule,BillTypeID"
Set RsTemp=conn.execute(sql)

while Not RsTemp.eof
	if trim(RsTemp("BillMemID"))<>"" and trim(RsTemp("UnitID"))<>"" then
		if BillMemID(3)<>"" then
			billunitID(3)=billunitID(3)&"||"
			billtypeid(3)=billtypeid(3)&"||"
			BillMemID(3)=BillMemID(3)&"||"
			LoginID(3)=LoginID(3)&"||"
			ChName(3)=ChName(3)&"||"
			UnitName(3)=UnitName(3)&"||"
			BillScore(3)=BillScore(3)&"||"
			BillTotal(3)=BillTotal(3)&"||"
			Rule(3)=Rule(3)&"||"
			Other3(3)=Other3(3)&"||"
		end if
		billunitID(3)=billunitID(3)&RsTemp("UnitID")
		billtypeid(3)=billtypeid(3)&RsTemp("billtypeid")
		BillMemID(3)=BillMemID(3)&RsTemp("BillMemID")
		LoginID(3)=LoginID(3)&RsTemp("LoginID")
		ChName(3)=ChName(3)&RsTemp("ChName")
		UnitName(3)=UnitName(3)&RsTemp("UnitName")
		BillScore(3)=BillScore(3)&RsTemp("TotalSum")
		BillTotal(3)=BillTotal(3)&RsTemp("TotalNum")
		Rule(3)=Rule(3)&RsTemp("Rule")
		Other3(3)=Other3(3)&RsTemp("Other3")
	end if
	RsTemp.movenext
wend
RsTemp.close
billunitID(3)=split(billunitID(3),"||")
billtypeid(3)=split(billtypeid(3),"||")
BillMemID(3)=split(BillMemID(3),"||")
LoginID(3)=split(LoginID(3),"||")
ChName(3)=split(ChName(3),"||")
UnitName(3)=split(UnitName(3),"||")
BillScore(3)=split(BillScore(3),"||")
BillTotal(3)=split(BillTotal(3),"||")
Rule(3)=split(Rule(3),"||")
Other3(3)=split(Other3(3),"||")

For i=0 to Ubound(BillMemID(3))
	If chkBillMemID<>"" Then
		chkBillMemID=chkBillMemID&","
		chkUnitName=chkUnitName&","
		chkLoginID=chkLoginID&","
		chkchName=chkchName&","
		chkRule=chkRule&","
	end if
	chkBillMemID=chkBillMemID&BillMemID(3)(i)
	chkUnitName=chkUnitName&UnitName(3)(i)
	chkLoginID=chkLoginID&LoginID(3)(i)
	chkchName=chkchName&ChName(3)(i)
	chkRule=chkRule&Rule(3)(i)
Next

'========================================================================================================

dim strtab_e(3,3,3)

strtab_e(0,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(1,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(2,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(3,0,0)="select BillMemID,Other3,Rule,BillTypeID,count(*) TotalNum,Count(*)*BillScore TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(0,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(1,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(2,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(3,0,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql


strtab_e(0,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(1,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(2,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(3,0,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.34 TotalNum,Count(*)*BillScore*0.34 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(0,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(1,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(2,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(3,0,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID1 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem1 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(0,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=8" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(1,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=8" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(2,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=8" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(3,1,1)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.5 TotalNum,Count(*)*BillScore*0.5 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID2 is not null and BillMemID3 is null and NVL(CarAddID,0)=8" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(0,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=8" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(1,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=8" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(2,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=8" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(3,1,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=8" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(0,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(1,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(2,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(3,1,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID2 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem2 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(0,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=8" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(1,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=8" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(2,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=8" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(3,2,2)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.33 TotalNum,Count(*)*BillScore*0.33 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID3 is not null and BillMemID4 is null and NVL(CarAddID,0)=8" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(0,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(1,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(2,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(3,2,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID3 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem3 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(0,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule1 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule1 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(1,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule2 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule2 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(2,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule3 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule3 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

strtab_e(3,3,3)="select BillMemID,Other3,Rule,BillTypeID,count(*)*0.25 TotalNum,Count(*)*BillScore*0.25 TotalSum from (select BillMemID,Other3,Rule,BillTypeID,DeCode(BillTypeID,1,BillType1Score,2,BillType2Score) BillScore from (select rule4 Rule,BillTypeID,BillMemID4 BillMemID from ("&BillbaseViewqry&") BillViewqry where rule4 is not null and BillMemID4 is not null and NVL(CarAddID,0)=8" & Mem4 &") Bill,(select LawItem,BillType1Score,BillType2Score,Other3 From LawScore where CountyOrnpa="&PageType&" and Other3='違規拖吊' and LawVerSion="&RuleVer&") LawScore where Bill.Rule=LawScore.LawItem) BillBasePurge" & qsql

tempTable(4)=""
for i=0 to ubound(strtab_e,1)
	for j=0 to Ubound(strtab_e,2)
		for k=0 to Ubound(strtab_e,3)
			if trim(strtab_e(i,j,k))<>"" then
				if trim(tempTable(4))<>"" then tempTable(4)=tempTable(4) & " Union all "
				tempTable(4)=tempTable(4) & strtab_e(i,j,k)
			end if
		next
	next
next

strTable="select BillMemID,Other3,Rule,BillTypeID,sum(TotalNum) TotalNum,sum(TotalSum) TotalSum from ("&tempTable(4)&")"&qsq2


sql1 = "Select BillCount.*,Decode(BillCount.Other3,'違反路權規定','01','違反行人路權','02','砂石車違規','03','計程車違規','04','高速公路違規','05','清除道路障礙','06','違規拖吊','07','其他交通重點違規','08','09') OtherOrder,MemberData.ChName,MemberData.LoginID,MemberData.UnitID,UnitInfo.UnitName from ("&strTable&") BillCount,MemberData,UnitInfo where BillCount.BillMemID=MemberData.MemberID and MemberData.UnitID=UnitInfo.UnitID"

sql = sql1 & " order by UnitID,LoginID,OtherOrder,Other3,Rule,BillTypeID"
Set RsTemp=conn.execute(sql)

while Not RsTemp.eof
	if trim(RsTemp("BillMemID"))<>"" and trim(RsTemp("UnitID"))<>"" then
		if BillMemID(4)<>"" then
			billunitID(4)=billunitID(4)&"||"
			billtypeid(4)=billtypeid(4)&"||"
			BillMemID(4)=BillMemID(4)&"||"
			LoginID(4)=LoginID(4)&"||"
			ChName(4)=ChName(4)&"||"
			UnitName(4)=UnitName(4)&"||"
			BillScore(4)=BillScore(4)&"||"
			BillTotal(4)=BillTotal(4)&"||"
			Rule(4)=Rule(4)&"||"
			Other3(4)=Other3(4)&"||"
		end if
		billunitID(4)=billunitID(4)&RsTemp("UnitID")
		billtypeid(4)=billtypeid(4)&RsTemp("billtypeid")
		BillMemID(4)=BillMemID(4)&RsTemp("BillMemID")
		LoginID(4)=LoginID(4)&RsTemp("LoginID")
		ChName(4)=ChName(4)&RsTemp("ChName")
		UnitName(4)=UnitName(4)&RsTemp("UnitName")
		BillScore(4)=BillScore(4)&RsTemp("TotalSum")
		BillTotal(4)=BillTotal(4)&RsTemp("TotalNum")
		Rule(4)=Rule(4)&RsTemp("Rule")
		Other3(4)=Other3(4)&RsTemp("Other3")
	end if
	RsTemp.movenext
wend
RsTemp.close
billunitID(4)=split(billunitID(4),"||")
billtypeid(4)=split(billtypeid(4),"||")
BillMemID(4)=split(BillMemID(4),"||")
LoginID(4)=split(LoginID(4),"||")
ChName(4)=split(ChName(4),"||")
UnitName(4)=split(UnitName(4),"||")
BillScore(4)=split(BillScore(4),"||")
BillTotal(4)=split(BillTotal(4),"||")
Rule(4)=split(Rule(4),"||")
Other3(4)=split(Other3(4),"||")

For i=0 to Ubound(BillMemID(4))
	If chkBillMemID<>"" Then
		chkBillMemID=chkBillMemID&","
		chkUnitName=chkUnitName&","
		chkLoginID=chkLoginID&","
		chkchName=chkchName&","
		chkRule=chkRule&","
	end if
	chkBillMemID=chkBillMemID&BillMemID(4)(i)
	chkUnitName=chkUnitName&UnitName(4)(i)
	chkLoginID=chkLoginID&LoginID(4)(i)
	chkchName=chkchName&ChName(4)(i)
	chkRule=chkRule&Rule(4)(i)
Next

tempLawID=split(chkRule,",")
chkRule=""
For i=0 to Ubound(tempLawID)
	For j=0 to i-1
		If trim(tempLawID(i))=trim(tempLawID(j)) Then exit for
	Next
	If j>i-1 Then
		If trim(chkRule)<>"" Then
			chkRule=chkRule&"||"
		end if
		chkRule=chkRule&tempLawID(i)
	End if
Next
chkRule=split(chkRule,"||")

tempChName=split(chkchName,","):tempBillMemID=split(chkBillMemID,","):tempUnitName=split(chkUnitName,",")
tempLoginID=split(chkLoginID,",")

chkBillMemID="":chkUnitName="":chkLoginID="":chkchName=""

For i=0 to Ubound(tempBillMemID)
	For j=0 to i-1
		If trim(tempBillMemID(i))=trim(tempBillMemID(j)) Then exit for
	Next
	If j>i-1 Then
		If trim(chkBillMemID)<>"" Then
			chkBillMemID=chkBillMemID&"||"
			chkUnitName=chkUnitName&"||"
			chkLoginID=chkLoginID&"||"
			chkchName=chkchName&"||"
		end if
		chkBillMemID=chkBillMemID&tempBillMemID(i)
		chkUnitName=chkUnitName&tempUnitName(i)
		chkLoginID=chkLoginID&tempLoginID(i)
		chkchName=chkchName&tempchName(i)
	End if
Next
chkBillMemID=split(chkBillMemID,"||")
chkUnitName=split(chkUnitName,"||")
chkLoginID=split(chkLoginID,"||")
chkchName=split(chkchName,"||")

set RSSystem=Server.CreateObject("ADODB.RecordSet")

sql = "select sysdate from Dual"
Set RSSystem = Conn.Execute(sql)
DBDate = RSSystem("sysdate")

sql = "select UnitName from UnitInfo where UnitID= '" & Session("Unit_ID") & "'"
Set RSSystem = Conn.Execute(sql)
if Not RSSystem.Eof Then
	printUnit = RSSystem("UnitName")
End If	

selectUnit = ""
If unit="y" Then
   sql = "Select UnitName , UnitID from UnitInfo Where UnitID= '" & UnitID_q & "'"
   Set RSSystem = Conn.Execute(sql)
   if Not RSSystem.Eof Then
   	  selectUnit = RSSystem("UnitName")
   End If
End If

strKindName=split("違反路權規定,違反行人路權,砂石車違規,計程車違規,高速公路違規,清除道路障礙,違規拖吊,其他交通重點違規",",")
%>
<html>   
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>ExportBase</title>
<style type="text/css">
<!--
body {font-family:標楷體;font-size:12pt}
.style1 {font-family:標楷體;font-size:14pt}
-->
</style>
</head>	 
<body>    
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" >
		<tr>
			<td colspan=2>
				  列印時間: <%=gInitDT(DBDate)%> <br>
			    列印單位: <%=printUnit%> <br>
			    列印人員: <%=Session("Ch_Name")%>
			</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>	  
	</table>
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" align="center" >
		<tr>
			<tr>
				<td>&nbsp;</td>
				<td colspan=3><center><span class="style1"><u><b><%=rptHead2%></b></u></span></center></td>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td colspan=3><center>統計期間: <%=startDate_q%> 至 <%=endDate_q%></center></td>
				<td>&nbsp;</td>
			</tr>		
			<tr>
				<td>&nbsp;</td>
				<td colspan=3><center>舉發單種類: 全部資料 &nbsp;&nbsp;&nbsp;&nbsp;舉發單類別: 全部資料</center></td>
				<td>&nbsp;</td>
			</tr>					
		</tr>
	</table>
	<br>
	<%
countsum1=0
countsum2=0
for h=0 to ubound(chkchName)
	response.write "<table border=0>"
	Response.Write "<tr><td>單位：</td><td>"&chkUnitName(h)&"&nbsp;</td></tr>"
	response.write "<tr><td>人員：</td><td>"&chkLoginID(h)&"&nbsp;</td><td>"&chkchName(h) & "&nbsp;</td></tr>"
	response.write "</table>"
	For j=0 to ubound(chkRule)
		chktable=false:BillScoreSum=0:BillScoretal=0
		For i = 0 To 4
			For t=0 to Ubound(BillMemID(i))
				if trim(LoginID(i)(t))=trim(chkLoginID(h)) and trim(Rule(i)(t))=trim(chkRule(j)) then
					If chktable=false Then
						response.write "法條："&chkRule(j)%>
						<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#808080" >	
						<tr>
							<td><B><center>專案名稱</center></B></td>
							<td><B><center>舉發類別</center></B></td>
							<td><B><center>件數</center></B></td>
							<td><B><center>配分</center></B></td>
						</tr><%
						chktable=true
					End if
					response.write "<tr>"
					response.write "<td>"&Other3(i)(t)&"</td>"
					If billtypeid(i)(t)=1 Then
						response.write "<td>攔停</td>"
					else
						response.write "<td>逕舉</td>"
					End if
					response.write "<td>"&BillTotal(i)(t)&"</td>"
					response.write "<td>"&BillScore(i)(t)&"</td>"
					response.write "</tr>"
					BillScoreSum=BillScoreSum+BillScore(i)(t)
					BillScoretal=BillScoretal+BillTotal(i)(t)
				end if
			next
		next
		If chktable=true Then
			response.write "<tr>"
			response.write "<td colspan=2>小計</td>"
			response.write "<td>"&BillScoretal&"</td>"
			response.write "<td>"&BillScoreSum&"</td>"
			response.write "</tr></table><br><br>"
			countsum1=countsum1+BillScoreSum
			countsum2=countsum2+BillScoretal
		end if
	next
	response.write "<br>"
	response.write "總件數："&countsum2
	response.write "  總積分："&countsum1
	response.write "<br>"

	countsum2=0:countsum1=0
next


ConnExecute Request.ServerVariables ("SCRIPT_NAME")&"||"& DateDiff("s",log_start,now) &"||"& startDate_q & "~" & endDate_q ,361

fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_"&ReportName
Response.AddHeader "Content-Disposition", "filename="&fname&".xls"
response.contenttype="application/x-msexcel; charset=MS950"
%>	 
 </body>
</html>