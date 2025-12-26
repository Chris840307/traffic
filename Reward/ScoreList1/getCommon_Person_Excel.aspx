<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<%  LoginCheck()
    
    Dim fMnoth = Month(Now)
    If fMnoth < 10 Then fMnoth = "0" & fMnoth
    Dim fDay = Day(Now)
    If fDay < 10 Then fDay = "0" & fDay
    Dim fname = Year(Now) & fMnoth & fDay & ".xls"
    Response.AddHeader("Content-Disposition", "filename=" & fname)
    If Trim(Request("sMemID")) = "" Then
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8")
    End If
    Response.ContentType = "application/ms-excel"
    
    Server.ScriptTimeout = 86400
    Response.Flush()
%>
<object id="factory" style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://localhost/traffic/smsx.cab#Version=6,1,432,1">
</object>

<script runat="server">
    Public PersonPoint1, PointTotal, MoneyTotal As Decimal
    
    '將民國yymmdd轉換為yyyy/mm/dd
    Public Function gOutDT(ByVal iDate)
        Dim DatetTemp As String
        If iDate IsNot DBNull.Value Then
            DatetTemp = DateSerial(Left(iDate, Len(iDate) - 4) + 1911, Mid(iDate, Len(iDate) - 3, 2), Right(iDate, 2))
            gOutDT = DatetTemp
        Else
            gOutDT = ""
        End If
    End Function
    
    '將民國yymmdd轉換為民國yy/mm/dd
    Public Function gOutDT2(ByVal iDate)
        Dim DatetTemp As String
        If iDate IsNot DBNull.Value Then
            DatetTemp = Left(iDate, Len(iDate) - 4) & "/" & Mid(iDate, Len(iDate) - 3, 2) & "/" & Right(iDate, 2)
            gOutDT2 = DatetTemp
        Else
            gOutDT2 = ""
        End If
    End Function
    
    Sub LoginCheck()
        If (Request.Cookies("UserFunction") IsNot Nothing) Then
            Dim FuncCookie As HttpCookie = Request.Cookies("UserFunction")
            If Trim(FuncCookie.Values("FuncID")) = "" Then
                Response.Redirect("/traffic/Reward/Login.aspx?ErrMsg=1")
            End If
        Else
            Response.Redirect("/traffic/Reward/Login.aspx?ErrMsg=1")
        End If
    End Sub
    
    Public Function RewardMonthData(ByVal strCreditID, ByVal strUnit, ByVal strLoginID, ByVal strMemberID, ByVal strChName, ByVal ShouldGetMoney, ByVal RealGetMoney, ByVal UserID)
        '****每月應領實領****
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
        
        Dim strDel As String = "delete from RewardMonthData where DirectOrTogether='0'"
        strDel = strDel & " and YearMonth=TO_DATE('" & gOutDT(Trim(Request("Date1"))) & "','YYYY/MM/DD')"
        strDel = strDel & " and UnitId='" & Trim(strUnit) & "' and LoginId='" & Trim(strLoginID) & "'"
        strDel = strDel & " and MemberId=" & Trim(strMemberID)

        Dim cmdDel As New Data.OracleClient.OracleCommand()
        cmdDel.CommandText = strDel
        cmdDel.Connection = conn
        cmdDel.ExecuteNonQuery()
                                    
        Dim strInsert As String = "insert into RewardMonthData(DirectOrTogether,YearMonth,UnitId,LoginId,ChName"
        strInsert = strInsert & ",MemberId,CreditID,ShouldGetMoney,RealGetMoney,RecordDate,RecordMemberID)"
        strInsert = strInsert & " values('0',TO_DATE('" & gOutDT(Trim(Request("Date1"))) & "','YYYY/MM/DD')"
        strInsert = strInsert & ",'" & Trim(strUnit) & "','" & Trim(strLoginID) & "'"
        strInsert = strInsert & ",'" & Trim(strChName) & "'," & Trim(strMemberID) & ",'" & strCreditID & "'," & ShouldGetMoney
        strInsert = strInsert & "," & RealGetMoney & ",sysdate," & UserID
        strInsert = strInsert & ")"
        Dim cmdInsert As New Data.OracleClient.OracleCommand()
        cmdInsert.CommandText = strInsert
        cmdInsert.Connection = conn
        cmdInsert.ExecuteNonQuery()
        
        conn.Close()
        '*********************
    End Function
        
</script>

<script language="JavaScript">
	window.focus();
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<style type="text/css">
<!--
body {font-family:新細明體;font-size:12pt }

.style1 {font-family:新細明體; font-size: 14pt}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
    <title>共同人員支領獎勵金核發清冊</title>
</head>
<body>
    <form id="form1" runat="server">
        <%
            
            '取得 Web.config 檔的資料連接設定
            Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
            '建立 Connection 物件
            Dim conn As New Data.OracleClient.OracleConnection()
            conn.ConnectionString = setting.ConnectionString
            '開啟資料連接
            conn.Open()
            
            '要用填單或建檔日統計
            Dim theDateType As String = Trim(Request("DateType"))
            '================================================
            '獎勵金總額
            Dim getMoneyTotal, DBReward28 As Decimal
            If Trim(Request("AllAnalyzeMoney")) = "" Then
                getMoneyTotal = 0
            Else
                getMoneyTotal = CDec(Request("AllAnalyzeMoney"))
            End If
            
            '取得獎勵金28%
            If Trim(Request("AnalyzeMoney")) = "" Then
                DBReward28 = 0
            Else
                DBReward28 = CDec(Request("AnalyzeMoney"))
            End If
            
            '個人薪資百分比
            Dim getPayPercent As Decimal
            Dim strPayPercent As String = "select value from Apconfigure where ID=47"
            Dim CmdPayPercent As New Data.OracleClient.OracleCommand(strPayPercent, conn)
            Dim rdPayPercent As Data.OracleClient.OracleDataReader = CmdPayPercent.ExecuteReader()
            If rdPayPercent.HasRows Then
                rdPayPercent.Read()
                        
                getPayPercent = CDec(rdPayPercent("value") / 100)
                'Response.Write(getPayPercent)
            End If
            rdPayPercent.Close()
            
            Dim sys_City As String = ""
            Dim strCity = "select Value from ApConfigure where ID=31"
            Dim CmdCity As New Data.OracleClient.OracleCommand(strCity, conn)
            Dim rdCity As Data.OracleClient.OracleDataReader = CmdCity.ExecuteReader()
            If rdCity.HasRows Then
                rdCity.Read()
                sys_City = Trim(rdCity("Value"))
            End If
            rdCity.Close()
            
            '取得四類別百分比
            Dim ShareGroup1, ShareGroup2, ShareGroup3, ShareGroup4 As Decimal
            Dim strGroupReward As String
            strGroupReward = "select * from CommonShareReward where ShareGroupID=0 order by SN"
            Dim CmdGroupReward As New Data.OracleClient.OracleCommand(strGroupReward, conn)
            Dim rdGroupReward As Data.OracleClient.OracleDataReader = CmdGroupReward.ExecuteReader()
            If rdGroupReward.HasRows Then
                While rdGroupReward.Read()
                    If Trim(rdGroupReward("CommonShareUnit")) = "1" Then
                        ShareGroup1 = CDec(rdGroupReward("SharePercent"))
                    ElseIf Trim(rdGroupReward("CommonShareUnit")) = "2" Then
                        ShareGroup2 = CDec(rdGroupReward("SharePercent"))
                    ElseIf Trim(rdGroupReward("CommonShareUnit")) = "3" Then
                        ShareGroup3 = CDec(rdGroupReward("SharePercent"))
                    ElseIf Trim(rdGroupReward("CommonShareUnit")) = "4" Then
                        ShareGroup4 = CDec(rdGroupReward("SharePercent"))
                    End If
                End While
            End If
            rdGroupReward.Close()
            '====================計算每點多少錢=======================
            Dim getPointTotal, getPointTota2, PageSum As Decimal
            getPointTotal = 0
            PageSum = 0
            Dim FlagPointTotal As String
            FlagPointTotal = Trim(Request("AnalyzeType"))
            
            '交通隊代碼
            Dim AnalyzeUnitID As String = ""
            Dim strUID = "select Value from Apconfigure where ID=49"
            Dim CmdUID As New Data.OracleClient.OracleCommand(strUID, conn)
            Dim rsUID As Data.OracleClient.OracleDataReader = CmdUID.ExecuteReader()
            If rsUID.HasRows Then
                rsUID.Read()
                AnalyzeUnitID = Trim(rsUID("Value"))
            End If
            rsUID.Close()
            
            '縣警局名稱
            Dim CityUnitName As String = ""
            Dim strCUName = "select Value from Apconfigure where ID=40"
            Dim CmdCUName As New Data.OracleClient.OracleCommand(strCUName, conn)
            Dim rsCUName As Data.OracleClient.OracleDataReader = CmdCUName.ExecuteReader()
            If rsCUName.HasRows Then
                rsCUName.Read()
                CityUnitName = Trim(rsCUName("Value"))
            End If
            rsCUName.Close()
            
            Dim UserCookie As HttpCookie = Request.Cookies("RewardUser")
            Dim AnalyzeUnitID2 As String = ""
            AnalyzeUnitID2 = Trim(UserCookie.Values("UnitID"))
            Dim UserID, UserName As String
            UserID = Trim(UserCookie.Values("MemberID"))
            
            Dim strUserName = "select * from MemberData where MemberID=" & UserID
            Dim CmdUserName As New Data.OracleClient.OracleCommand(strUserName, conn)
            Dim rdUserName As Data.OracleClient.OracleDataReader = CmdUserName.ExecuteReader()
            If rdUserName.HasRows Then
                rdUserName.Read()
                If rdUserName("ChName") IsNot DBNull.Value Then
                    UserName = Trim(rdUserName("ChName"))
                Else
                    UserName = ""
                End If
            End If
            rdUserName.Close()
            
            Dim UnitFlag As String
            If sys_City = "台中縣" Or sys_City = "台中市" Then
                UnitFlag = " and a.BillUnitID<>'" & AnalyzeUnitID & "'"
            Else
                If FlagPointTotal = 1 Then  '總分只抓交通隊
                    UnitFlag = " and a.BillUnitID='" & AnalyzeUnitID & "'"
                ElseIf FlagPointTotal = 2 Then  '總分只抓分局
                    UnitFlag = " and a.BillUnitID in (select UnitID from UnitInfo where UnitID='" & AnalyzeUnitID2 & "' or UnitTypeID='" & AnalyzeUnitID2 & "')"
                Else
                    UnitFlag = ""
                End If
            End If
               
            '攔停點數
            Dim strPointT1 As String = "select sum(b.BillType1Score) as cnt from BillBaseViewReward a"
            strPointT1 = strPointT1 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
            strPointT1 = strPointT1 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
            strPointT1 = strPointT1 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion " & UnitFlag
            strPointT1 = strPointT1 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
            strPointT1 = strPointT1 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
            strPointT1 = strPointT1 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPointT1 = strPointT1 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdPointT1 As New Data.OracleClient.OracleCommand(strPointT1, conn)
            Dim rdPointT1 As Data.OracleClient.OracleDataReader = CmdPointT1.ExecuteReader()
            If rdPointT1.HasRows Then
                rdPointT1.Read()
                If rdPointT1("cnt") Is DBNull.Value Then
                    getPointTotal = 0
                Else
                    getPointTotal = CDec(rdPointT1("cnt"))
                End If
            End If
            rdPointT1.Close()

            '逕舉點數
            Dim strPointT2 As String = "select sum(b.BillType2Score) as cnt from BillBase a"
            strPointT2 = strPointT2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
            strPointT2 = strPointT2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
            strPointT2 = strPointT2 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion " & UnitFlag
            strPointT2 = strPointT2 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
            strPointT2 = strPointT2 & " and a.BillTypeID='2' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
            strPointT2 = strPointT2 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPointT2 = strPointT2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdPointT2 As New Data.OracleClient.OracleCommand(strPointT2, conn)
            Dim rdPointT2 As Data.OracleClient.OracleDataReader = CmdPointT2.ExecuteReader()
            If rdPointT2.HasRows Then
                rdPointT2.Read()
                If rdPointT2("cnt") Is DBNull.Value Then
                    getPointTotal = getPointTotal + 0
                Else
                    getPointTotal = getPointTotal + CDec(rdPointT2("cnt"))
                End If
            End If
            rdPointT2.Close()
            
            'A1點數
            Dim strPointT3 As String = "select count(*) as cnt from BillBase a,UnitInfo c"
            strPointT3 = strPointT3 & " where a.BillUnitID=c.UnitID " & UnitFlag
            strPointT3 = strPointT3 & " and a.RecordStateID=0 and (a.TrafficAccidentType='1') and a.BillTypeID='1'"
            strPointT3 = strPointT3 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPointT3 = strPointT3 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdPointT3 As New Data.OracleClient.OracleCommand(strPointT3, conn)
            Dim rdPointT3 As Data.OracleClient.OracleDataReader = CmdPointT3.ExecuteReader()
            If rdPointT3.HasRows Then
                rdPointT3.Read()
                getPointTotal = getPointTotal + (CDec(rdPointT3("cnt")) * 100)
                    
            End If
            rdPointT3.Close()
            
            'A2點數
            Dim strPointT4 As String = "select count(*) as cnt from BillBase a,UnitInfo c"
            strPointT4 = strPointT4 & " where a.BillUnitID=c.UnitID " & UnitFlag
            strPointT4 = strPointT4 & " and a.RecordStateID=0 and (a.TrafficAccidentType='2') and a.BillTypeID='1'"
            strPointT4 = strPointT4 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPointT4 = strPointT4 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdPointT4 As New Data.OracleClient.OracleCommand(strPointT4, conn)
            Dim rdPointT4 As Data.OracleClient.OracleDataReader = CmdPointT4.ExecuteReader()
            If rdPointT4.HasRows Then
                rdPointT4.Read()
                getPointTotal = getPointTotal + (CDec(rdPointT4("cnt")) * 50)

            End If
            rdPointT4.Close()
                    
            'A3點數
            Dim strPointT5 As String = "select count(*) as cnt from BillBase a,UnitInfo c"
            strPointT5 = strPointT5 & " where a.BillUnitID=c.UnitID " & UnitFlag
            strPointT5 = strPointT5 & " and a.RecordStateID=0 and (a.TrafficAccidentType='3') and a.BillTypeID='1'"
            strPointT5 = strPointT5 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPointT5 = strPointT5 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdPointT5 As New Data.OracleClient.OracleCommand(strPointT5, conn)
            Dim rdPointT5 As Data.OracleClient.OracleDataReader = CmdPointT5.ExecuteReader()
            If rdPointT5.HasRows Then
                rdPointT5.Read()
                getPointTotal = getPointTotal + (CDec(rdPointT5("cnt")) * 20)
                        
            End If
            rdPointT5.Close()
                      
            '拖吊
            If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "台中市" Then
                Dim strPointT1a As String = "select count(*) cnt from BillBase a"
                strPointT1a = strPointT1a & " ,UnitInfo c"
                strPointT1a = strPointT1a & " where a.BillUnitID=c.UnitID " & UnitFlag
                strPointT1a = strPointT1a & " and a.RecordStateID=0"
                strPointT1a = strPointT1a & " and a.ProjectID='A5'"
                strPointT1a = strPointT1a & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPointT1a = strPointT1a & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPointT1a As New Data.OracleClient.OracleCommand(strPointT1a, conn)
                Dim rdPointT1a As Data.OracleClient.OracleDataReader = CmdPointT1a.ExecuteReader()
                If rdPointT1a.HasRows Then
                    rdPointT1a.Read()
                    If rdPointT1a("cnt") Is DBNull.Value Then
                        getPointTotal = getPointTotal
                    Else
                        
                        getPointTotal = getPointTotal + (CDec(rdPointT1a("cnt")) * 20)
                    End If
                End If
                rdPointT1a.Close()
                
                Dim strPointT1b As String = "select count(*) cnt from BillBase a"
                strPointT1b = strPointT1b & " ,UnitInfo c"
                strPointT1b = strPointT1b & " where a.BillUnitID=c.UnitID " & UnitFlag
                strPointT1b = strPointT1b & " and a.RecordStateID=0"
                strPointT1b = strPointT1b & " and a.ProjectID='A6'"
                strPointT1b = strPointT1b & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPointT1b = strPointT1b & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPointT1b As New Data.OracleClient.OracleCommand(strPointT1b, conn)
                Dim rdPointT1b As Data.OracleClient.OracleDataReader = CmdPointT1b.ExecuteReader()
                If rdPointT1b.HasRows Then
                    rdPointT1b.Read()
                    If rdPointT1b("cnt") Is DBNull.Value Then
                        getPointTotal = getPointTotal
                    Else
                        
                        getPointTotal = getPointTotal + (CDec(rdPointT1b("cnt")) * 50)
                    End If
                End If
                rdPointT1b.Close()
            End If
            
            'Dim PointMoney As Decimal
            'If getPointTotal = 0 Then
            '    PointMoney = 0
            'Else
            '    PointMoney = getMoneyTotal / getPointTotal
            'End If
            If sys_City = "台中縣" Or sys_City = "台中市" Then
                Dim UnitFlag2 As String
                UnitFlag2 = " and a.BillUnitID<>'" & AnalyzeUnitID & "' and a.BillUnitID not in (Select UnitID from UnitInfo where ShowOrder=0)"
                
                '攔停點數
                Dim strPointT12 As String = "select sum(b.BillType1Score) as cnt from BillBaseViewReward a"
                strPointT12 = strPointT12 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                strPointT12 = strPointT12 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
                strPointT12 = strPointT12 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion " & UnitFlag2
                strPointT12 = strPointT12 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                strPointT12 = strPointT12 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                strPointT12 = strPointT12 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPointT12 = strPointT12 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPointT12 As New Data.OracleClient.OracleCommand(strPointT12, conn)
                Dim rdPointT12 As Data.OracleClient.OracleDataReader = CmdPointT12.ExecuteReader()
                If rdPointT12.HasRows Then
                    rdPointT12.Read()
                    If rdPointT12("cnt") Is DBNull.Value Then
                        getPointTota2 = 0
                    Else
                        getPointTota2 = CDec(rdPointT12("cnt"))
                    End If
                End If
                rdPointT12.Close()

                '逕舉點數
                Dim strPointT22 As String = "select sum(b.BillType2Score) as cnt from BillBase a"
                strPointT22 = strPointT22 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                strPointT22 = strPointT22 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
                strPointT22 = strPointT22 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion " & UnitFlag2
                strPointT22 = strPointT22 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                strPointT22 = strPointT22 & " and a.BillTypeID='2' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                strPointT22 = strPointT22 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPointT22 = strPointT22 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                'Response.Write(strPointT22)
                'Response.End()
                Dim CmdPointT22 As New Data.OracleClient.OracleCommand(strPointT22, conn)
                Dim rdPointT22 As Data.OracleClient.OracleDataReader = CmdPointT22.ExecuteReader()
                If rdPointT22.HasRows Then
                    rdPointT22.Read()
                    If rdPointT22("cnt") Is DBNull.Value Then
                        getPointTota2 = getPointTota2 + 0
                    Else
                        getPointTota2 = getPointTota2 + CDec(rdPointT22("cnt"))
                    End If
                End If
                rdPointT22.Close()
            
                'A1點數
                Dim strPointT32 As String = "select count(*) as cnt from BillBase a,UnitInfo c"
                strPointT32 = strPointT32 & " where a.BillUnitID=c.UnitID " & UnitFlag2
                strPointT32 = strPointT32 & " and a.RecordStateID=0 and (a.TrafficAccidentType='1') and a.BillTypeID='1'"
                strPointT32 = strPointT32 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPointT32 = strPointT32 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPointT32 As New Data.OracleClient.OracleCommand(strPointT32, conn)
                Dim rdPointT32 As Data.OracleClient.OracleDataReader = CmdPointT32.ExecuteReader()
                If rdPointT32.HasRows Then
                    rdPointT32.Read()
                    getPointTota2 = getPointTota2 + (CDec(rdPointT32("cnt")) * 100)
                
                End If
                rdPointT32.Close()
            
                'A2點數
                Dim strPointT42 As String = "select count(*) as cnt from BillBase a,UnitInfo c"
                strPointT42 = strPointT42 & " where a.BillUnitID=c.UnitID " & UnitFlag2
                strPointT42 = strPointT42 & " and a.RecordStateID=0 and (a.TrafficAccidentType='2') and a.BillTypeID='1'"
                strPointT42 = strPointT42 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPointT42 = strPointT42 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPointT42 As New Data.OracleClient.OracleCommand(strPointT42, conn)
                Dim rdPointT42 As Data.OracleClient.OracleDataReader = CmdPointT42.ExecuteReader()
                If rdPointT42.HasRows Then
                    rdPointT42.Read()
                    getPointTota2 = getPointTota2 + (CDec(rdPointT42("cnt")) * 50)

                End If
                rdPointT42.Close()
                    
                'A3點數
                Dim strPointT52 As String = "select count(*) as cnt from BillBase a,UnitInfo c"
                strPointT52 = strPointT52 & " where a.BillUnitID=c.UnitID " & UnitFlag2
                strPointT52 = strPointT52 & " and a.RecordStateID=0 and (a.TrafficAccidentType='3') and a.BillTypeID='1'"
                strPointT52 = strPointT52 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPointT52 = strPointT52 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPointT52 As New Data.OracleClient.OracleCommand(strPointT52, conn)
                Dim rdPointT52 As Data.OracleClient.OracleDataReader = CmdPointT52.ExecuteReader()
                If rdPointT52.HasRows Then
                    rdPointT52.Read()
                    getPointTota2 = getPointTota2 + (CDec(rdPointT52("cnt")) * 20)
                        
                End If
                rdPointT52.Close()
                      
                '拖吊
                If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "台中市" Then
                    Dim strPointT12a As String = "select count(*) cnt from BillBase a"
                    strPointT12a = strPointT12a & " ,UnitInfo c"
                    strPointT12a = strPointT12a & " where a.BillUnitID=c.UnitID " & UnitFlag2
                    strPointT12a = strPointT12a & " and a.RecordStateID=0"
                    strPointT12a = strPointT12a & " and a.ProjectID='A5'"
                    strPointT12a = strPointT12a & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPointT12a = strPointT12a & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPointT12a As New Data.OracleClient.OracleCommand(strPointT12a, conn)
                    Dim rdPointT12a As Data.OracleClient.OracleDataReader = CmdPointT12a.ExecuteReader()
                    If rdPointT12a.HasRows Then
                        rdPointT12a.Read()
                        If rdPointT12a("cnt") Is DBNull.Value Then
                            getPointTota2 = getPointTota2
                        Else
                        
                            getPointTota2 = getPointTota2 + (CDec(rdPointT12a("cnt")) * 20)
                        End If
                    End If
                    rdPointT12a.Close()
                
                    Dim strPointT12b As String = "select count(*) cnt from BillBase a"
                    strPointT12b = strPointT12b & " ,UnitInfo c"
                    strPointT12b = strPointT12b & " where a.BillUnitID=c.UnitID " & UnitFlag2
                    strPointT12b = strPointT12b & " and a.RecordStateID=0"
                    strPointT12b = strPointT12b & " and a.ProjectID='A6'"
                    strPointT12b = strPointT12b & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPointT12b = strPointT12b & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPointT12b As New Data.OracleClient.OracleCommand(strPointT12b, conn)
                    Dim rdPointT12b As Data.OracleClient.OracleDataReader = CmdPointT12b.ExecuteReader()
                    If rdPointT12b.HasRows Then
                        rdPointT12b.Read()
                        If rdPointT12b("cnt") Is DBNull.Value Then
                            getPointTota2 = getPointTota2
                        Else
                        
                            getPointTota2 = getPointTota2 + (CDec(rdPointT12b("cnt")) * 50)
                        End If
                    End If
                    rdPointT12b.Close()
                End If
            End If
            'Response.Write(PointMoney & " " & getPointTotal)
            '******************列出清冊內容********************
            '===================1.先跑交通隊==============
            Dim PageCount, PageNo, i As Integer
            Dim Type1Money, Type2Money, Type3Money, Type4Money, MemMoney, PersonMoney, ShouldGetMoney As Decimal
            Dim UnitName, CreditIDtmp, UnitIDtmp2, LoginIDtmp, MemberIDtmp, ChNametmp, BankIDtmp As String
            If sys_City = "台中縣" Or sys_City = "台中市" Then
                PageCount = 19
                Type1Money = DBReward28 * ShareGroup1
                If getPointTotal <> 0 Then
                    Type2Money = (DBReward28 * ShareGroup2) / getPointTotal
                Else
                    Type2Money = 0
                End If
                If getPointTota2 <> 0 Then
                    Type3Money = (DBReward28 * ShareGroup3) / getPointTota2
                Else
                    Type3Money = 0
                End If
                Type4Money = DBReward28 * ShareGroup4
            Else
                PageCount = 14
                Type1Money = DBReward28 * ShareGroup1
                If getPointTotal <> 0 Then
                    Type2Money = (DBReward28 * ShareGroup2) / getPointTotal
                    Type3Money = (DBReward28 * ShareGroup3) / getPointTotal
                Else
                    Type2Money = 0
                    Type3Money = 0
                End If
                Type4Money = DBReward28 * ShareGroup4
            End If

            If InStr(Trim(Request("sUnitID")), AnalyzeUnitID) > 0 Then
                '-----------先算交通隊點數------------
                '攔停點數
                Dim strPoint1 As String = "select sum(b.BillType1Score) as cnt,count(*) as BillCnt from BillBaseViewReward a"
                strPoint1 = strPoint1 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                strPoint1 = strPoint1 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                strPoint1 = strPoint1 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & AnalyzeUnitID & "') and a.RuleVer=b.LawVersion"
                strPoint1 = strPoint1 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                strPoint1 = strPoint1 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                strPoint1 = strPoint1 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPoint1 = strPoint1 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPoint1 As New Data.OracleClient.OracleCommand(strPoint1, conn)
                Dim rdPoint1 As Data.OracleClient.OracleDataReader = CmdPoint1.ExecuteReader()
                If rdPoint1.HasRows Then
                    rdPoint1.Read()
                    If rdPoint1("cnt") Is DBNull.Value Then
                        PersonPoint1 = PersonPoint1 + 0
                    Else
                        PersonPoint1 = PersonPoint1 + CDec(rdPoint1("cnt"))
                    End If
                End If
                rdPoint1.Close()
                    
                '逕舉點數
                Dim strPoint2 As String = "select sum(b.BillType2Score) as cnt,count(*) as BillCnt from BillBaseViewReward a"
                strPoint2 = strPoint2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                strPoint2 = strPoint2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                strPoint2 = strPoint2 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & AnalyzeUnitID & "') and a.RuleVer=b.LawVersion"
                strPoint2 = strPoint2 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                strPoint2 = strPoint2 & " and a.BillTypeID='2' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                strPoint2 = strPoint2 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPoint2 = strPoint2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPoint2 As New Data.OracleClient.OracleCommand(strPoint2, conn)
                Dim rdPoint2 As Data.OracleClient.OracleDataReader = CmdPoint2.ExecuteReader()
                If rdPoint2.HasRows Then
                    rdPoint2.Read()
                    If rdPoint2("cnt") Is DBNull.Value Then
                        PersonPoint1 = PersonPoint1 + 0
                    Else
                        PersonPoint1 = PersonPoint1 + CDec(rdPoint2("cnt"))
                    End If
                End If
                rdPoint2.Close()
                    
                'A1點數
                Dim strPoint3 As String = "select count(*) as cnt from BillBase a"
                strPoint3 = strPoint3 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & AnalyzeUnitID & "') "
                strPoint3 = strPoint3 & " and a.RecordStateID=0 and (a.TrafficAccidentType='1') and a.BillTypeID='1'"
                strPoint3 = strPoint3 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPoint3 = strPoint3 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPoint3 As New Data.OracleClient.OracleCommand(strPoint3, conn)
                Dim rdPoint3 As Data.OracleClient.OracleDataReader = CmdPoint3.ExecuteReader()
                If rdPoint3.HasRows Then
                    rdPoint3.Read()
                       
                    PersonPoint1 = PersonPoint1 + (CDec(rdPoint3("cnt")) * 100)
     
                End If
                rdPoint3.Close()
                    
                'A2點數
                Dim strPoint4 As String = "select count(*) as cnt from BillBase a"
                strPoint4 = strPoint4 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & AnalyzeUnitID & "') "
                strPoint4 = strPoint4 & " and a.RecordStateID=0 and (a.TrafficAccidentType='2') and a.BillTypeID='1'"
                strPoint4 = strPoint4 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPoint4 = strPoint4 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPoint4 As New Data.OracleClient.OracleCommand(strPoint4, conn)
                Dim rdPoint4 As Data.OracleClient.OracleDataReader = CmdPoint4.ExecuteReader()
                If rdPoint4.HasRows Then
                    rdPoint4.Read()
                    PersonPoint1 = PersonPoint1 + (CDec(rdPoint4("cnt")) * 50)

                End If
                rdPoint4.Close()
                    
                'A3點數
                Dim strPoint5 As String = "select count(*) as cnt from BillBase a"
                strPoint5 = strPoint5 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & AnalyzeUnitID & "') "
                strPoint5 = strPoint5 & " and a.RecordStateID=0 and (a.TrafficAccidentType='3') and a.BillTypeID='1'"
                strPoint5 = strPoint5 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPoint5 = strPoint5 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPoint5 As New Data.OracleClient.OracleCommand(strPoint5, conn)
                Dim rdPoint5 As Data.OracleClient.OracleDataReader = CmdPoint5.ExecuteReader()
                If rdPoint5.HasRows Then
                    rdPoint5.Read()

                    PersonPoint1 = PersonPoint1 + (CDec(rdPoint5("cnt")) * 20)

                End If
                rdPoint5.Close()
                
                '拖吊
                If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "台中市" Then
                    Dim strPointT1a As String = "select count(*) cnt from BillBase a"
                    strPointT1a = strPointT1a & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & AnalyzeUnitID & "') and a.RecordStateID=0"
                    strPointT1a = strPointT1a & " and a.ProjectID='A5'"
                    strPointT1a = strPointT1a & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPointT1a = strPointT1a & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPointT1a As New Data.OracleClient.OracleCommand(strPointT1a, conn)
                    Dim rdPointT1a As Data.OracleClient.OracleDataReader = CmdPointT1a.ExecuteReader()
                    If rdPointT1a.HasRows Then
                        rdPointT1a.Read()
                        If rdPointT1a("cnt") Is DBNull.Value Then
                            PersonPoint1 = PersonPoint1
                        Else
                        
                            PersonPoint1 = PersonPoint1 + (CDec(rdPointT1a("cnt")) * 20)
                        End If
                    End If
                    rdPointT1a.Close()
                
                    Dim strPointT1b As String = "select count(*) cnt from BillBase a"
                    strPointT1b = strPointT1b & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & AnalyzeUnitID & "') and a.RecordStateID=0"
                    strPointT1b = strPointT1b & " and a.ProjectID='A6'"
                    strPointT1b = strPointT1b & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPointT1b = strPointT1b & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPointT1b As New Data.OracleClient.OracleCommand(strPointT1b, conn)
                    Dim rdPointT1b As Data.OracleClient.OracleDataReader = CmdPointT1b.ExecuteReader()
                    If rdPointT1b.HasRows Then
                        rdPointT1b.Read()
                        If rdPointT1b("cnt") Is DBNull.Value Then
                            PersonPoint1 = PersonPoint1
                        Else
                        
                            PersonPoint1 = PersonPoint1 + (CDec(rdPointT1b("cnt")) * 50)
                        End If
                    End If
                    rdPointT1b.Close()
                End If
                '--------------------------------------------
                PageNo = 1

                Dim strUnit1 As String = "select * from UnitInfo where UnitID='" & AnalyzeUnitID & "'"
                Dim CmdUnit1 As New Data.OracleClient.OracleCommand(strUnit1, conn)
                Dim rdUnit1 As Data.OracleClient.OracleDataReader = CmdUnit1.ExecuteReader()
                If rdUnit1.HasRows Then
                    rdUnit1.Read()
                    UnitName = Trim(rdUnit1("UnitName"))
                End If
                rdUnit1.Close()
                
                Dim strType1 As String = "select a.UnitID,a.CommonShareUnit,a.ShareGroupID,a.SharePercent,a.ChName"
                strType1 = strType1 & " from CommonShareReward a"
                strType1 = strType1 & " where (a.ShareGroupID in (1,4) or a.UnitID='" & AnalyzeUnitID & "')"
                strType1 = strType1 & " order by sn"
                Dim CmdType1 As New Data.OracleClient.OracleCommand(strType1, conn)
                Dim rdType1 As Data.OracleClient.OracleDataReader = CmdType1.ExecuteReader()
                If rdType1.HasRows Then
                    While rdType1.Read()
                        CreditIDtmp = ""
                        UnitIDtmp2 = ""
                        LoginIDtmp = ""
                        MemberIDtmp = ""
                        ShouldGetMoney = 0
                        MemMoney = 0
                        BankIDtmp = ""
                        PageSum = 0
                        Response.Write("<table width=""640"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                        Response.Write("<tr>")
                        Response.Write("<td align=""center"" colspan=""6""><span class=""style1"">" & CityUnitName & UnitName & "</span></td>")
                        Response.Write("</tr>")
                        Response.Write("<tr>")
                        Response.Write("<td align=""right"" colspan=""6""><span class=""style1"">" & gOutDT2(Trim(Request("Date1"))) & "~" & gOutDT2(Trim(Request("Date2"))) & "&nbsp; &nbsp;共同人員交通安全獎金印領清冊</span>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 頁次：" & PageNo & "</td>")
                        Response.Write("</tr>")
                        Response.Write("<tr>")
                        Response.Write("<td colspan=""3"">列印日期：" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "</td>")
                        Response.Write("<td align=""right"" colspan=""3"">列印人員：" & UserName & "</td>")
                        Response.Write("</tr>")
                        Response.Write("</table>")
                        
                    
                        Response.Write("<table width=""640"" border=""1"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                        Response.Write("<tr>")
                        Response.Write("<td height=""35"" align=""center"" width=""8%"">單位</td>")
                        Response.Write("<td align=""center"" width=""12%"">職位</td>")
                        Response.Write("<td align=""center"" width=""12%"">姓名</td>")
                        Response.Write("<td align=""center"" width=""12%"">實領金額</td>")
                        If sys_City <> "台中縣" And sys_City <> "台中市" Then
                            Response.Write("<td align=""center"" width=""13%"">身份證號</td>")
                        End If
                        'Response.Write("<td align=""center"" width=""13%"">郵局局號</td>")
                        Response.Write("<td align=""center"" width=""22%"">郵局帳號</td>")
                        Response.Write("<td align=""center"" width=""12%"">備註</td>")
                        Response.Write("</tr>")
                        
                        Response.Write("<tr>")
                        '單位
                        Response.Write("<td align=""center"" height=""35"">")
                        If rdType1("ChName") IsNot DBNull.Value Then
                            Dim strMoney1 As String = "select * from MemberData where ChName='" & Trim(rdType1("ChName")) & "' order by RecordDate Desc"
                            Dim CmdMoney1 As New Data.OracleClient.OracleCommand(strMoney1, conn)
                            Dim rdMoney1 As Data.OracleClient.OracleDataReader = CmdMoney1.ExecuteReader()
                            If rdMoney1.HasRows Then
                                rdMoney1.Read()
                                If rdMoney1("Money") IsNot DBNull.Value Then
                                    MemMoney = Decimal.Truncate(rdMoney1("Money") * getPayPercent)
                                Else
                                    MemMoney = 0
                                End If
                                
                                If rdMoney1("CreditID") IsNot DBNull.Value Then
                                    CreditIDtmp = Trim(rdMoney1("CreditID"))
                                Else
                                    CreditIDtmp = ""
                                End If
                                If rdMoney1("LoginID") IsNot DBNull.Value Then
                                    LoginIDtmp = Trim(rdMoney1("LoginID"))
                                Else
                                    LoginIDtmp = ""
                                End If
                                If rdMoney1("UnitID") IsNot DBNull.Value Then
                                    UnitIDtmp2 = Trim(rdMoney1("UnitID"))
                                Else
                                    UnitIDtmp2 = ""
                                End If
                                If rdMoney1("MemberID") IsNot DBNull.Value Then
                                    MemberIDtmp = Trim(rdMoney1("MemberID"))
                                Else
                                    MemberIDtmp = ""
                                End If
                                If rdMoney1("BankName") IsNot DBNull.Value Then
                                    BankIDtmp = Trim(rdMoney1("BankName"))
                                Else
                                    BankIDtmp = ""
                                End If
                                If rdMoney1("BankID") IsNot DBNull.Value Then
                                    BankIDtmp = BankIDtmp & Trim(rdMoney1("BankID"))
                                Else
                                    BankIDtmp = BankIDtmp
                                End If
                                If rdMoney1("BankAccount") IsNot DBNull.Value Then
                                    BankIDtmp = BankIDtmp & Trim(rdMoney1("BankAccount"))
                                Else
                                    BankIDtmp = BankIDtmp
                                End If
                            End If
                            rdMoney1.Close()
                        Else
                            MemMoney = 0
                        End If
                        
                        If UnitIDtmp2 <> "" Then
                            Response.Write(UnitIDtmp2 & "&nbsp;")
                        Else
                            Response.Write("&nbsp;")
                        End If
                        Response.Write("</td>")
                        '職位
                        Response.Write("<td align=""center"">")
                        If rdType1("CommonShareUnit") IsNot DBNull.Value Then
                            If Trim(rdType1("CommonShareUnit")) = "" Then
                                Response.Write("&nbsp;")
                            Else
                                Response.Write(rdType1("CommonShareUnit"))
                            End If
                        Else
                            Response.Write("&nbsp;")
                        End If
                        Response.Write("</td>")
                        '姓名
                        Response.Write("<td align=""center"">")
                        If rdType1("ChName") IsNot DBNull.Value Then
                            Response.Write(rdType1("ChName"))
                        Else
                            Response.Write("&nbsp;")
                        End If
                        Response.Write("</td>")
                        '實領金額
                        Response.Write("<td align=""center"">")
                        
                        If MemMoney = 0 Then
                            If Trim(rdType1("ShareGroupID")) = 1 Then
                                PersonMoney = Decimal.Round(Type1Money * rdType1("SharePercent"))
                            ElseIf Trim(rdType1("ShareGroupID")) = 2 Then
                                PersonMoney = Decimal.Round(PersonPoint1 * Type2Money * rdType1("SharePercent"))
                            ElseIf Trim(rdType1("ShareGroupID")) = 3 Then
                                PersonMoney = Decimal.Round(PersonPoint1 * Type3Money * rdType1("SharePercent"))
                            ElseIf Trim(rdType1("ShareGroupID")) = 4 Then
                                PersonMoney = Decimal.Round(Type4Money * rdType1("SharePercent"))
                            End If
                            ShouldGetMoney = PersonMoney
                        Else
                            If Trim(rdType1("ShareGroupID")) = 1 Then
                                PersonMoney = Decimal.Round(Type1Money * rdType1("SharePercent"))
                            ElseIf Trim(rdType1("ShareGroupID")) = 2 Then
                                PersonMoney = Decimal.Round(PersonPoint1 * Type2Money * rdType1("SharePercent"))
                            ElseIf Trim(rdType1("ShareGroupID")) = 3 Then
                                PersonMoney = Decimal.Round(PersonPoint1 * Type3Money * rdType1("SharePercent"))
                            ElseIf Trim(rdType1("ShareGroupID")) = 4 Then
                                PersonMoney = Decimal.Round(Type4Money * rdType1("SharePercent"))
                            End If
                            ShouldGetMoney = PersonMoney
                            If PersonMoney > MemMoney Then
                                PersonMoney = MemMoney
                            End If
                        End If
                        If MemberIDtmp <> "" And rdType1("ChName") IsNot DBNull.Value Then
                            RewardMonthData(CreditIDtmp, UnitIDtmp2, LoginIDtmp, MemberIDtmp, Trim(rdType1("ChName")), ShouldGetMoney, PersonMoney, UserID)
                        End If
                        MoneyTotal = MoneyTotal + PersonMoney
                        PageSum = PageSum + PersonMoney
                        Response.Write(PersonMoney)
                        Response.Write("</td>")
                        If sys_City <> "台中縣" And sys_City <> "台中市" Then
                            Response.Write("<td>")
                            '身分證號
                            'If sys_City = "台中縣" Then
                            '    If CreditIDtmp <> "" Then
                            '        Response.Write(CreditIDtmp)
                            '    Else
                            '        Response.Write("&nbsp;")
                            '    End If
                            'Else
                            Response.Write("&nbsp;")
                            'End If
                            Response.Write("</td>")
                        End If
                        Response.Write("<td align=""center"">")
                        '郵局帳號
                        If sys_City = "台中縣" Or sys_City = "台中市" Then
                            If BankIDtmp <> "" Then
                                Response.Write(BankIDtmp & "&nbsp;")
                            Else
                                Response.Write("&nbsp;")
                            End If
                        Else
                            Response.Write("&nbsp;")
                        End If
                        Response.Write("</td>")
                        Response.Write("<td>&nbsp;</td>")
                        Response.Write("</tr>")
                        For i = 1 To PageCount
                            If rdType1.Read() = True Then
                                CreditIDtmp = ""
                                UnitIDtmp2 = ""
                                LoginIDtmp = ""
                                MemberIDtmp = ""
                                BankIDtmp = ""
                                ShouldGetMoney = 0
                                MemMoney = 0
                                Response.Write("<tr>")
                                '單位
                                Response.Write("<td align=""center"" height=""35"">")
                                If rdType1("ChName") IsNot DBNull.Value Then
                                    Dim strMoney1b As String = "select * from MemberData where ChName='" & Trim(rdType1("ChName")) & "' order by RecordDate Desc"
                                    Dim CmdMoney1b As New Data.OracleClient.OracleCommand(strMoney1b, conn)
                                    Dim rdMoney1b As Data.OracleClient.OracleDataReader = CmdMoney1b.ExecuteReader()
                                    If rdMoney1b.HasRows Then
                                        rdMoney1b.Read()
                                        If rdMoney1b("Money") IsNot DBNull.Value Then
                                            MemMoney = Decimal.Truncate(rdMoney1b("Money") * getPayPercent)
                                        Else
                                            MemMoney = 0
                                        End If
                                        
                                        If rdMoney1b("CreditID") IsNot DBNull.Value Then
                                            CreditIDtmp = Trim(rdMoney1b("CreditID"))
                                        Else
                                            CreditIDtmp = ""
                                        End If
                                        If rdMoney1b("LoginID") IsNot DBNull.Value Then
                                            LoginIDtmp = Trim(rdMoney1b("LoginID"))
                                        Else
                                            LoginIDtmp = ""
                                        End If
                                        If rdMoney1b("UnitID") IsNot DBNull.Value Then
                                            UnitIDtmp2 = Trim(rdMoney1b("UnitID"))
                                        Else
                                            UnitIDtmp2 = ""
                                        End If
                                        If rdMoney1b("MemberID") IsNot DBNull.Value Then
                                            MemberIDtmp = Trim(rdMoney1b("MemberID"))
                                        Else
                                            MemberIDtmp = ""
                                        End If
                                        If rdMoney1b("BankName") IsNot DBNull.Value Then
                                            BankIDtmp = Trim(rdMoney1b("BankName"))
                                        Else
                                            BankIDtmp = ""
                                        End If
                                        If rdMoney1b("BankID") IsNot DBNull.Value Then
                                            BankIDtmp = BankIDtmp & Trim(rdMoney1b("BankID"))
                                        Else
                                            BankIDtmp = BankIDtmp
                                        End If
                                        If rdMoney1b("BankAccount") IsNot DBNull.Value Then
                                            BankIDtmp = BankIDtmp & Trim(rdMoney1b("BankAccount"))
                                        Else
                                            BankIDtmp = BankIDtmp
                                        End If
                                    End If
                                    rdMoney1b.Close()
                                Else
                                    MemMoney = 0
                                End If
                                If UnitIDtmp2 <> "" Then
                                    Response.Write(UnitIDtmp2 & "&nbsp;")
                                Else
                                    Response.Write("&nbsp;")
                                End If
                                Response.Write("</td>")
                                '職位
                                Response.Write("<td align=""center"">")
                                If rdType1("CommonShareUnit") IsNot DBNull.Value Then
                                    If Trim(rdType1("CommonShareUnit")) = "" Then
                                        Response.Write("&nbsp;")
                                    Else
                                        Response.Write(rdType1("CommonShareUnit"))
                                    End If
                                Else
                                    Response.Write("&nbsp;")
                                End If
                                Response.Write("</td>")
                                '姓名
                                Response.Write("<td align=""center"">")
                                If rdType1("ChName") IsNot DBNull.Value Then
                                    Response.Write(rdType1("ChName"))
                                Else
                                    Response.Write("&nbsp;")
                                End If
                                Response.Write("</td>")
                                '實領金額
                                Response.Write("<td align=""center"">")
                                
                                If Trim(rdType1("ShareGroupID")) = 1 Then
                                    PersonMoney = Decimal.Round(Type1Money * rdType1("SharePercent"))
                                ElseIf Trim(rdType1("ShareGroupID")) = 2 Then
                                    PersonMoney = Decimal.Round(PersonPoint1 * Type2Money * rdType1("SharePercent"))
                                ElseIf Trim(rdType1("ShareGroupID")) = 3 Then
                                    PersonMoney = Decimal.Round(PersonPoint1 * Type3Money * rdType1("SharePercent"))
                                ElseIf Trim(rdType1("ShareGroupID")) = 4 Then
                                    PersonMoney = Decimal.Round(Type4Money * rdType1("SharePercent"))
                                End If
                                ShouldGetMoney = PersonMoney
                                If MemMoney > 0 Then
                                    If PersonMoney > MemMoney Then
                                        PersonMoney = MemMoney
                                    End If
                                End If
                                If MemberIDtmp <> "" And rdType1("ChName") IsNot DBNull.Value Then
                                    RewardMonthData(CreditIDtmp, UnitIDtmp2, LoginIDtmp, MemberIDtmp, Trim(rdType1("ChName")), ShouldGetMoney, PersonMoney, UserID)
                                End If
                                MoneyTotal = MoneyTotal + PersonMoney
                                PageSum = PageSum + PersonMoney
                                Response.Write(PersonMoney)
                                Response.Write("</td>")
                                If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                    Response.Write("<td>")
                                    '身分證號
                                    'If sys_City = "台中縣" Then
                                    '    If CreditIDtmp <> "" Then
                                    '        Response.Write(CreditIDtmp)
                                    '    Else
                                    '        Response.Write("&nbsp;")
                                    '    End If
                                    'Else
                                    Response.Write("&nbsp;")
                                    'End If
                                    Response.Write("</td>")
                                End If
                                Response.Write("<td align=""center"">")
                                '郵局帳號
                                If sys_City = "台中縣" Or sys_City = "台中市" Then
                                    If BankIDtmp <> "" Then
                                        Response.Write(BankIDtmp & "&nbsp;")
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                Else
                                    Response.Write("&nbsp;")
                                End If
                                Response.Write("</td>")
                                Response.Write("<td>&nbsp;</td>")
                                Response.Write("</tr>")
                            Else
                                Exit For
                            End If
                        Next
                        Response.Write("<tr>")
                        Response.Write("<td align=""center"" height=""35"">小計</td>")
                        Response.Write("<td colspan=""2"">&nbsp;</td>")
                        Response.Write("<td align=""center"">")
                        Response.Write(PageSum)
                        Response.Write("</td>")
                        If sys_City <> "台中縣" And sys_City <> "台中市" Then
                            Response.Write("<td colspan=""3"">&nbsp;</td>")
                        Else
                            Response.Write("<td colspan=""2"">&nbsp;</td>")
                        End If
                        Response.Write("</tr>")
                        Response.Write("</table>")
                        Response.Write("<table width=""640"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                        Response.Write("<tr><td colspan=""6""><strong>承辦單位：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                        If sys_City <> "台中縣" And sys_City <> "台中市" Then
                            '    Response.Write("人事：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                            '    Response.Write("出納：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                            '    Response.Write("會計：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                            '    Response.Write("機關主官：</strong></td>")
                            'Else
                            Response.Write("秘書室：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                            Response.Write("會計室：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                            Response.Write("局長：</strong></td>")
                        End If
                        Response.Write("</tr>")
                        Response.Write("</table>")
                        PageNo = PageNo + 1
                        Response.Write("<div class=""PageNext""></div>")
                    End While
                End If
                rdType1.Close()
            End If

            '=====================其他ShowOrder=0的單位==========================
            Dim strUnit As String
            strUnit = "select * from UnitInfo where UnitID in (" & Trim(Request("sUnitID")) & ") and UnitID<>'" & AnalyzeUnitID & "' and ShowOrder=0 order by UnitTypeID,UnitID"
            Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
            Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
            If rdUnit.HasRows Then
                While rdUnit.Read()
                    PersonPoint1 = 0
                    PageNo = 1
                    '------------------------計算單位點數-----------------------
                    '攔停點數
                    Dim strPoint1 As String = "select sum(b.BillType1Score) as cnt,count(*) as BillCnt from BillBaseViewReward a"
                    strPoint1 = strPoint1 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                    strPoint1 = strPoint1 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                    strPoint1 = strPoint1 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "') and a.RuleVer=b.LawVersion"
                    strPoint1 = strPoint1 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                    strPoint1 = strPoint1 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                    strPoint1 = strPoint1 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint1 = strPoint1 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint1 As New Data.OracleClient.OracleCommand(strPoint1, conn)
                    Dim rdPoint1 As Data.OracleClient.OracleDataReader = CmdPoint1.ExecuteReader()
                    If rdPoint1.HasRows Then
                        rdPoint1.Read()
                        If rdPoint1("cnt") Is DBNull.Value Then
                            PersonPoint1 = PersonPoint1 + 0
                        Else
                            PersonPoint1 = PersonPoint1 + CDec(rdPoint1("cnt"))
                        End If
                    End If
                    rdPoint1.Close()
                    
                    '逕舉點數
                    Dim strPoint2 As String = "select sum(b.BillType2Score) as cnt,count(*) as BillCnt from BillBase a"
                    strPoint2 = strPoint2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                    strPoint2 = strPoint2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                    strPoint2 = strPoint2 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "') and a.RuleVer=b.LawVersion"
                    strPoint2 = strPoint2 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                    strPoint2 = strPoint2 & " and a.BillTypeID='2' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                    strPoint2 = strPoint2 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint2 = strPoint2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint2 As New Data.OracleClient.OracleCommand(strPoint2, conn)
                    Dim rdPoint2 As Data.OracleClient.OracleDataReader = CmdPoint2.ExecuteReader()
                    If rdPoint2.HasRows Then
                        rdPoint2.Read()
                        If rdPoint2("cnt") Is DBNull.Value Then
                            PersonPoint1 = PersonPoint1 + 0
                        Else
                            PersonPoint1 = PersonPoint1 + CDec(rdPoint2("cnt"))
                        End If
                    End If
                    rdPoint2.Close()
                    
                    'A1點數
                    Dim strPoint3 As String = "select count(*) as cnt from BillBase a"
                    strPoint3 = strPoint3 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "') "
                    strPoint3 = strPoint3 & " and a.RecordStateID=0 and (a.TrafficAccidentType='1') and a.BillTypeID='1'"
                    strPoint3 = strPoint3 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint3 = strPoint3 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint3 As New Data.OracleClient.OracleCommand(strPoint3, conn)
                    Dim rdPoint3 As Data.OracleClient.OracleDataReader = CmdPoint3.ExecuteReader()
                    If rdPoint3.HasRows Then
                        rdPoint3.Read()

                        PersonPoint1 = PersonPoint1 + (CDec(rdPoint3("cnt")) * 100)
                        
                    End If
                    rdPoint3.Close()
                    
                    'A2點數
                    Dim strPoint4 As String = "select count(*) as cnt from BillBase a"
                    strPoint4 = strPoint4 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "') "
                    strPoint4 = strPoint4 & " and a.RecordStateID=0 and (a.TrafficAccidentType='2') and a.BillTypeID='1'"
                    strPoint4 = strPoint4 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint4 = strPoint4 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint4 As New Data.OracleClient.OracleCommand(strPoint4, conn)
                    Dim rdPoint4 As Data.OracleClient.OracleDataReader = CmdPoint4.ExecuteReader()
                    If rdPoint4.HasRows Then
                        rdPoint4.Read()

                        PersonPoint1 = PersonPoint1 + (CDec(rdPoint4("cnt")) * 50)

                    End If
                    rdPoint4.Close()
                    
                    'A3點數
                    Dim strPoint5 As String = "select count(*) as cnt from BillBase a"
                    strPoint5 = strPoint5 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "') "
                    strPoint5 = strPoint5 & " and a.RecordStateID=0 and (a.TrafficAccidentType='3') and a.BillTypeID='1'"
                    strPoint5 = strPoint5 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint5 = strPoint5 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint5 As New Data.OracleClient.OracleCommand(strPoint5, conn)
                    Dim rdPoint5 As Data.OracleClient.OracleDataReader = CmdPoint5.ExecuteReader()
                    If rdPoint5.HasRows Then
                        rdPoint5.Read()

                        PersonPoint1 = PersonPoint1 + (CDec(rdPoint5("cnt")) * 20)
                        
                    End If
                    rdPoint5.Close()
                    
                    '拖吊
                    If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "台中市" Then
                        Dim strPointT1a As String = "select count(*) cnt from BillBase a"
                        strPointT1a = strPointT1a & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "') and a.RecordStateID=0"
                        strPointT1a = strPointT1a & " and a.ProjectID='A5'"
                        strPointT1a = strPointT1a & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strPointT1a = strPointT1a & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdPointT1a As New Data.OracleClient.OracleCommand(strPointT1a, conn)
                        Dim rdPointT1a As Data.OracleClient.OracleDataReader = CmdPointT1a.ExecuteReader()
                        If rdPointT1a.HasRows Then
                            rdPointT1a.Read()
                            If rdPointT1a("cnt") Is DBNull.Value Then
                                PersonPoint1 = PersonPoint1
                            Else
                        
                                PersonPoint1 = PersonPoint1 + (CDec(rdPointT1a("cnt")) * 20)
                            End If
                        End If
                        rdPointT1a.Close()
                
                        Dim strPointT1b As String = "select count(*) cnt from BillBase a"
                        strPointT1b = strPointT1b & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "') and a.RecordStateID=0"
                        strPointT1b = strPointT1b & " and a.ProjectID='A6'"
                        strPointT1b = strPointT1b & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strPointT1b = strPointT1b & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdPointT1b As New Data.OracleClient.OracleCommand(strPointT1b, conn)
                        Dim rdPointT1b As Data.OracleClient.OracleDataReader = CmdPointT1b.ExecuteReader()
                        If rdPointT1b.HasRows Then
                            rdPointT1b.Read()
                            If rdPointT1b("cnt") Is DBNull.Value Then
                                PersonPoint1 = PersonPoint1
                            Else
                        
                                PersonPoint1 = PersonPoint1 + (CDec(rdPointT1b("cnt")) * 50)
                            End If
                        End If
                        rdPointT1b.Close()
                    End If
                    '---------------------------------------------------------
                    Dim strUnit2 As String = "select * from UnitInfo where UnitID='" & Trim(rdUnit("UnitID")) & "'"
                    Dim CmdUnit2 As New Data.OracleClient.OracleCommand(strUnit2, conn)
                    Dim rdUnit2 As Data.OracleClient.OracleDataReader = CmdUnit2.ExecuteReader()
                    If rdUnit2.HasRows Then
                        rdUnit2.Read()
                        UnitName = Trim(rdUnit2("UnitName"))
                    End If
                    rdUnit2.Close()
                    
                    Dim strType2 As String = "select a.UnitID,a.CommonShareUnit,a.ShareGroupID,a.SharePercent,a.ChName"
                    strType2 = strType2 & " from CommonShareReward a"
                    strType2 = strType2 & " where a.UnitID='" & Trim(rdUnit("UnitID")) & "'"
                    strType2 = strType2 & " order by sn"
                    Dim CmdType2 As New Data.OracleClient.OracleCommand(strType2, conn)
                    Dim rdType2 As Data.OracleClient.OracleDataReader = CmdType2.ExecuteReader()
                    If rdType2.HasRows Then
                        While rdType2.Read()
                            CreditIDtmp = ""
                            UnitIDtmp2 = ""
                            LoginIDtmp = ""
                            MemberIDtmp = ""
                            BankIDtmp = ""
                            ShouldGetMoney = 0
                            MemMoney = 0
                            PageSum = 0
                            Response.Write("<table width=""640"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                            Response.Write("<tr>")
                            Response.Write("<td align=""center"" colspan=""6""><span class=""style1"">" & CityUnitName & UnitName & "</span></td>")
                            Response.Write("</tr>")
                            Response.Write("<tr>")
                            Response.Write("<td align=""right"" colspan=""6""><span class=""style1"">" & gOutDT2(Trim(Request("Date1"))) & "~" & gOutDT2(Trim(Request("Date2"))) & "&nbsp; &nbsp;共同人員交通安全獎金印領清冊</span>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 頁次：" & PageNo & "</td>")
                            Response.Write("</tr>")
                            Response.Write("<tr>")
                            Response.Write("<td colspan=""3"">列印日期：" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "</td>")
                            Response.Write("<td align=""right"" colspan=""3"">列印人員：" & UserName & "</td>")
                            Response.Write("</tr>")
                            Response.Write("</table>")
                        
                    
                            Response.Write("<table width=""640"" border=""1"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                            Response.Write("<tr>")
                            Response.Write("<td height=""35"" align=""center"" width=""8%"">單位</td>")
                            Response.Write("<td align=""center"" width=""12%"">職位</td>")
                            Response.Write("<td align=""center"" width=""12%"">姓名</td>")
                            Response.Write("<td align=""center"" width=""12%"">實領金額</td>")
                            If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                Response.Write("<td align=""center"" width=""13%"">身份證號</td>")
                            End If
                            Response.Write("<td align=""center"" width=""22%"">郵局帳號</td>")
                            Response.Write("<td align=""center"" width=""12%"">備註</td>")
                            Response.Write("</tr>")
                        
                            Response.Write("<tr>")
                            '單位
                            Response.Write("<td align=""center"" height=""35"">")
                            If rdType2("UnitID") IsNot DBNull.Value Then
                                Response.Write(rdType2("UnitID") & "&nbsp;")
                            Else
                                Response.Write("&nbsp;")
                            End If
                            Response.Write("</td>")
                            '職位
                            Response.Write("<td align=""center"">")
                            If rdType2("CommonShareUnit") IsNot DBNull.Value Then
                                If Trim(rdType2("CommonShareUnit")) = "" Then
                                    Response.Write("&nbsp;")
                                Else
                                    Response.Write(rdType2("CommonShareUnit"))
                                End If
                            Else
                                Response.Write("&nbsp;")
                            End If
                            Response.Write("</td>")
                            '姓名
                            Response.Write("<td align=""center"">")
                            If rdType2("ChName") IsNot DBNull.Value Then
                                Response.Write(rdType2("ChName"))
                            Else
                                Response.Write("&nbsp;")
                            End If
                            Response.Write("</td>")
                            '實領金額
                            Response.Write("<td align=""center"">")
                            If rdType2("ChName") IsNot DBNull.Value Then
                                Dim strMoney1 As String = "select * from MemberData where ChName='" & Trim(rdType2("ChName")) & "' order by RecordDate Desc"
                                Dim CmdMoney1 As New Data.OracleClient.OracleCommand(strMoney1, conn)
                                Dim rdMoney1 As Data.OracleClient.OracleDataReader = CmdMoney1.ExecuteReader()
                                If rdMoney1.HasRows Then
                                    rdMoney1.Read()
                                    If rdMoney1("Money") IsNot DBNull.Value Then
                                        MemMoney = Decimal.Truncate(rdMoney1("Money") * getPayPercent)
                                    Else
                                        MemMoney = 0
                                    End If
                                    
                                    If rdMoney1("CreditID") IsNot DBNull.Value Then
                                        CreditIDtmp = Trim(rdMoney1("CreditID"))
                                    Else
                                        CreditIDtmp = ""
                                    End If
                                    If rdMoney1("LoginID") IsNot DBNull.Value Then
                                        LoginIDtmp = Trim(rdMoney1("LoginID"))
                                    Else
                                        LoginIDtmp = ""
                                    End If
                                    If rdMoney1("UnitID") IsNot DBNull.Value Then
                                        UnitIDtmp2 = Trim(rdMoney1("UnitID"))
                                    Else
                                        UnitIDtmp2 = ""
                                    End If
                                    If rdMoney1("MemberID") IsNot DBNull.Value Then
                                        MemberIDtmp = Trim(rdMoney1("MemberID"))
                                    Else
                                        MemberIDtmp = ""
                                    End If
                                    If rdMoney1("BankName") IsNot DBNull.Value Then
                                        BankIDtmp = Trim(rdMoney1("BankName"))
                                    Else
                                        BankIDtmp = ""
                                    End If
                                    If rdMoney1("BankID") IsNot DBNull.Value Then
                                        BankIDtmp = BankIDtmp & Trim(rdMoney1("BankID"))
                                    Else
                                        BankIDtmp = BankIDtmp
                                    End If
                                    If rdMoney1("BankAccount") IsNot DBNull.Value Then
                                        BankIDtmp = BankIDtmp & Trim(rdMoney1("BankAccount"))
                                    Else
                                        BankIDtmp = BankIDtmp
                                    End If
                                End If
                                rdMoney1.Close()
                            Else
                                MemMoney = 0
                            End If
                            If MemMoney = 0 Then
                                If Trim(rdType2("ShareGroupID")) = 1 Then
                                    PersonMoney = Decimal.Round(Type1Money * rdType2("SharePercent"))
                                ElseIf Trim(rdType2("ShareGroupID")) = 2 Then
                                    PersonMoney = Decimal.Round(PersonPoint1 * Type2Money * rdType2("SharePercent"))
                                ElseIf Trim(rdType2("ShareGroupID")) = 3 Then
                                    PersonMoney = Decimal.Round(PersonPoint1 * Type3Money * rdType2("SharePercent"))
                                ElseIf Trim(rdType2("ShareGroupID")) = 4 Then
                                    PersonMoney = Decimal.Round(Type4Money * rdType2("SharePercent"))
                                End If
                                ShouldGetMoney = PersonMoney
                            Else
                                If Trim(rdType2("ShareGroupID")) = 1 Then
                                    PersonMoney = Decimal.Round(Type1Money * rdType2("SharePercent"))
                                ElseIf Trim(rdType2("ShareGroupID")) = 2 Then
                                    PersonMoney = Decimal.Round(PersonPoint1 * Type2Money * rdType2("SharePercent"))
                                ElseIf Trim(rdType2("ShareGroupID")) = 3 Then
                                    PersonMoney = Decimal.Round(PersonPoint1 * Type3Money * rdType2("SharePercent"))
                                ElseIf Trim(rdType2("ShareGroupID")) = 4 Then
                                    PersonMoney = Decimal.Round(Type4Money * rdType2("SharePercent"))
                                End If
                                ShouldGetMoney = PersonMoney
                                If PersonMoney > MemMoney Then
                                    PersonMoney = MemMoney
                                End If
                            End If
                            If MemberIDtmp <> "" And rdType2("ChName") IsNot DBNull.Value Then
                                RewardMonthData(CreditIDtmp, UnitIDtmp2, LoginIDtmp, MemberIDtmp, Trim(rdType2("ChName")), ShouldGetMoney, PersonMoney, UserID)
                            End If
                            MoneyTotal = MoneyTotal + PersonMoney
                            PageSum = PageSum + PersonMoney
                            Response.Write(PersonMoney)
                            Response.Write("</td>")
                            If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                Response.Write("<td>")
                                '身分證號
                                'If sys_City = "台中縣" Then
                                '    If CreditIDtmp <> "" Then
                                '        Response.Write(CreditIDtmp)
                                '    Else
                                '        Response.Write("&nbsp;")
                                '    End If
                                'Else
                                '    Response.Write("&nbsp;")
                                'End If
                                Response.Write("</td>")
                            End If
                            Response.Write("<td align=""center"">")
                            '郵局帳號
                            If sys_City = "台中縣" Or sys_City = "台中市" Then
                                If BankIDtmp <> "" Then
                                    Response.Write(BankIDtmp & "&nbsp;")
                                Else
                                    Response.Write("&nbsp;")
                                End If
                            Else
                                Response.Write("&nbsp;")
                            End If
                            Response.Write("</td>")
                            Response.Write("<td>&nbsp;</td>")
                            Response.Write("</tr>")
                            For i = 1 To PageCount
                                CreditIDtmp = ""
                                UnitIDtmp2 = ""
                                LoginIDtmp = ""
                                MemberIDtmp = ""
                                BankIDtmp = ""
                                ShouldGetMoney = 0
                                MemMoney = 0
                                If rdType2.Read() = True Then
                                    Response.Write("<tr>")
                                    '單位
                                    Response.Write("<td align=""center"" height=""35"">")
                                    If rdType2("UnitID") IsNot DBNull.Value Then
                                        Response.Write(rdType2("UnitID") & "&nbsp;")
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                    '職位
                                    Response.Write("<td align=""center"">")
                                    If rdType2("CommonShareUnit") IsNot DBNull.Value Then
                                        If Trim(rdType2("CommonShareUnit")) = "" Then
                                            Response.Write("&nbsp;")
                                        Else
                                            Response.Write(rdType2("CommonShareUnit"))
                                        End If
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                    '姓名
                                    Response.Write("<td align=""center"">")
                                    If rdType2("ChName") IsNot DBNull.Value Then
                                        Response.Write(rdType2("ChName"))
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                    '實領金額
                                    Response.Write("<td align=""center"">")
                                    If rdType2("ChName") IsNot DBNull.Value Then
                                        Dim strMoney1b As String = "select * from MemberData where ChName='" & Trim(rdType2("ChName")) & "' order by RecordDate Desc"
                                        Dim CmdMoney1b As New Data.OracleClient.OracleCommand(strMoney1b, conn)
                                        Dim rdMoney1b As Data.OracleClient.OracleDataReader = CmdMoney1b.ExecuteReader()
                                        If rdMoney1b.HasRows Then
                                            rdMoney1b.Read()
                                            If rdMoney1b("Money") IsNot DBNull.Value Then
                                                MemMoney = Decimal.Truncate(rdMoney1b("Money") * getPayPercent)
                                            Else
                                                MemMoney = 0
                                            End If
                                            
                                            If rdMoney1b("CreditID") IsNot DBNull.Value Then
                                                CreditIDtmp = Trim(rdMoney1b("CreditID"))
                                            Else
                                                CreditIDtmp = ""
                                            End If
                                            If rdMoney1b("LoginID") IsNot DBNull.Value Then
                                                LoginIDtmp = Trim(rdMoney1b("LoginID"))
                                            Else
                                                LoginIDtmp = ""
                                            End If
                                            If rdMoney1b("UnitID") IsNot DBNull.Value Then
                                                UnitIDtmp2 = Trim(rdMoney1b("UnitID"))
                                            Else
                                                UnitIDtmp2 = ""
                                            End If
                                            If rdMoney1b("MemberID") IsNot DBNull.Value Then
                                                MemberIDtmp = Trim(rdMoney1b("MemberID"))
                                            Else
                                                MemberIDtmp = ""
                                            End If
                                            If rdMoney1b("BankName") IsNot DBNull.Value Then
                                                BankIDtmp = Trim(rdMoney1b("BankName"))
                                            Else
                                                BankIDtmp = ""
                                            End If
                                            If rdMoney1b("BankID") IsNot DBNull.Value Then
                                                BankIDtmp = BankIDtmp & Trim(rdMoney1b("BankID"))
                                            Else
                                                BankIDtmp = BankIDtmp
                                            End If
                                            If rdMoney1b("BankAccount") IsNot DBNull.Value Then
                                                BankIDtmp = BankIDtmp & Trim(rdMoney1b("BankAccount"))
                                            Else
                                                BankIDtmp = BankIDtmp
                                            End If
                                        End If
                                        rdMoney1b.Close()
                                    Else
                                        MemMoney = 0
                                    End If
                                    If Trim(rdType2("ShareGroupID")) = 1 Then
                                        PersonMoney = Decimal.Round(Type1Money * rdType2("SharePercent"))
                                    ElseIf Trim(rdType2("ShareGroupID")) = 2 Then
                                        PersonMoney = Decimal.Round(PersonPoint1 * Type2Money * rdType2("SharePercent"))
                                    ElseIf Trim(rdType2("ShareGroupID")) = 3 Then
                                        PersonMoney = Decimal.Round(PersonPoint1 * Type3Money * rdType2("SharePercent"))
                                    ElseIf Trim(rdType2("ShareGroupID")) = 4 Then
                                        PersonMoney = Decimal.Round(Type4Money * rdType2("SharePercent"))
                                    End If
                                    ShouldGetMoney = PersonMoney
                                    If MemMoney > 0 Then
                                        If PersonMoney > MemMoney Then
                                            PersonMoney = MemMoney
                                        End If
                                    End If
                                    If MemberIDtmp <> "" And rdType2("ChName") IsNot DBNull.Value Then
                                        RewardMonthData(CreditIDtmp, UnitIDtmp2, LoginIDtmp, MemberIDtmp, Trim(rdType2("ChName")), ShouldGetMoney, PersonMoney, UserID)
                                    End If
                                    MoneyTotal = MoneyTotal + PersonMoney
                                    PageSum = PageSum + PersonMoney
                                    Response.Write(PersonMoney)
                                    Response.Write("</td>")
                                    If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                        Response.Write("<td>")
                                        '身分證號
                                        'If sys_City = "台中縣" Then
                                        '    If CreditIDtmp <> "" Then
                                        '        Response.Write(CreditIDtmp)
                                        '    Else
                                        '        Response.Write("&nbsp;")
                                        '    End If
                                        'Else
                                        Response.Write("&nbsp;")
                                        'End If
                                        Response.Write("</td>")
                                    End If
                                    Response.Write("<td align=""center"">")
                                    '郵局帳號
                                    If sys_City = "台中縣" Or sys_City = "台中市" Then
                                        If BankIDtmp <> "" Then
                                            Response.Write(BankIDtmp & "&nbsp;")
                                        Else
                                            Response.Write("&nbsp;")
                                        End If
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                    Response.Write("<td>&nbsp;</td>")
                                    Response.Write("</tr>")
                                Else
                                    Exit For
                                End If
                            Next
                            Response.Write("<tr>")
                            Response.Write("<td align=""center"" height=""35"">小計</td>")
                            Response.Write("<td colspan=""2"">&nbsp;</td>")
                            Response.Write("<td align=""center"">")
                            Response.Write(PageSum)
                            Response.Write("</td>")
                            If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                Response.Write("<td colspan=""3"">&nbsp;</td>")
                            Else
                                Response.Write("<td colspan=""2"">&nbsp;</td>")
                            End If
                            Response.Write("</tr>")
                            Response.Write("</table>")
                            Response.Write("<table width=""640"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                            If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                '    Response.Write("<tr><td colspan=""6""><strong>承辦單位：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")

                                '    Response.Write("人事：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                '    Response.Write("出納：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                '    Response.Write("會計：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                '    Response.Write("機關主官：</strong></td>")
                                'Else
                                Response.Write("<tr><td colspan=""6""><strong>承辦單位：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")

                                Response.Write("秘書室：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                Response.Write("會計室：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                Response.Write("局長：</strong></td>")
                            End If
                            Response.Write("</tr>")
                            Response.Write("</table>")
                            PageNo = PageNo + 1
                            Response.Write("<div class=""PageNext""></div>")
                            
                        End While
                        rdType2.Close()
                    End If
                    
                End While
            End If
            rdUnit.Close()
            
            '=====================分局==========================
            Dim SubUnitPoint As Decimal
            Dim UnitSN, j As Integer
            Dim strSubUnit, strUnitScore, UnitIDTmp, UnitPointTmp As String
            strSubUnit = "select * from UnitInfo where UnitID in (" & Trim(Request("sUnitID")) & ") and UnitID<>'" & AnalyzeUnitID & "' and ShowOrder=1 order by UnitTypeID,UnitID"
            Dim CmdSubUnit As New Data.OracleClient.OracleCommand(strSubUnit, conn)
            Dim rdSubUnit As Data.OracleClient.OracleDataReader = CmdSubUnit.ExecuteReader()
            If rdSubUnit.HasRows Then
                While rdSubUnit.Read()
                    PageNo = 1
                    '-----------計算分局及底下派出所總分-----------
                    UnitIDTmp = ""
                    UnitPointTmp = ""
                    PersonPoint1 = 0
                    strUnitScore = "select * from UnitInfo where UnitID='" & Trim(rdSubUnit("UnitID")) & "' or UnitTypeID='" & Trim(rdSubUnit("UnitID")) & "' order by UnitTypeID,UnitID"
                    Dim CmdUnitScore As New Data.OracleClient.OracleCommand(strUnitScore, conn)
                    Dim rdUnitScore As Data.OracleClient.OracleDataReader = CmdUnitScore.ExecuteReader()
                    If rdUnitScore.HasRows Then
                        While rdUnitScore.Read()
                            SubUnitPoint = 0
                            '攔停點數
                            Dim strPoint1 As String = "select sum(b.BillType1Score) as cnt,count(*) as BillCnt from BillBaseViewReward a"
                            strPoint1 = strPoint1 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                            strPoint1 = strPoint1 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                            strPoint1 = strPoint1 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnitScore("UnitID")) & "') and a.RuleVer=b.LawVersion"
                            strPoint1 = strPoint1 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                            strPoint1 = strPoint1 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                            strPoint1 = strPoint1 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                            strPoint1 = strPoint1 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                            Dim CmdPoint1 As New Data.OracleClient.OracleCommand(strPoint1, conn)
                            Dim rdPoint1 As Data.OracleClient.OracleDataReader = CmdPoint1.ExecuteReader()
                            If rdPoint1.HasRows Then
                                rdPoint1.Read()
                                If rdPoint1("cnt") Is DBNull.Value Then
                                    PersonPoint1 = PersonPoint1 + 0
                                    SubUnitPoint = SubUnitPoint + 0
                                Else
                                    PersonPoint1 = PersonPoint1 + CDec(rdPoint1("cnt"))
                                    SubUnitPoint = SubUnitPoint + CDec(rdPoint1("cnt"))
                                End If
                            End If
                            rdPoint1.Close()
                    
                            '逕舉點數
                            Dim strPoint2 As String = "select sum(b.BillType2Score) as cnt,count(*) as BillCnt from BillBase a"
                            strPoint2 = strPoint2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                            strPoint2 = strPoint2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                            strPoint2 = strPoint2 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnitScore("UnitID")) & "') and a.RuleVer=b.LawVersion"
                            strPoint2 = strPoint2 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                            strPoint2 = strPoint2 & " and a.BillTypeID='2' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                            strPoint2 = strPoint2 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                            strPoint2 = strPoint2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                            Dim CmdPoint2 As New Data.OracleClient.OracleCommand(strPoint2, conn)
                            Dim rdPoint2 As Data.OracleClient.OracleDataReader = CmdPoint2.ExecuteReader()
                            If rdPoint2.HasRows Then
                                rdPoint2.Read()
                                If rdPoint2("cnt") Is DBNull.Value Then
                                    PersonPoint1 = PersonPoint1 + 0
                                    SubUnitPoint = SubUnitPoint + 0
                                Else
                                    PersonPoint1 = PersonPoint1 + CDec(rdPoint2("cnt"))
                                    SubUnitPoint = SubUnitPoint + CDec(rdPoint2("cnt"))
                                End If
                            End If
                            rdPoint2.Close()
                    
                            'A1點數
                            Dim strPoint3 As String = "select count(*) as cnt from BillBase a"
                            strPoint3 = strPoint3 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnitScore("UnitID")) & "') "
                            strPoint3 = strPoint3 & " and a.RecordStateID=0 and (a.TrafficAccidentType='1') and a.BillTypeID='1'"
                            strPoint3 = strPoint3 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                            strPoint3 = strPoint3 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                            Dim CmdPoint3 As New Data.OracleClient.OracleCommand(strPoint3, conn)
                            Dim rdPoint3 As Data.OracleClient.OracleDataReader = CmdPoint3.ExecuteReader()
                            If rdPoint3.HasRows Then
                                rdPoint3.Read()
    
                                PersonPoint1 = PersonPoint1 + (CDec(rdPoint3("cnt")) * 100)
                                SubUnitPoint = SubUnitPoint + (CDec(rdPoint3("cnt")) * 100)

                            End If
                            rdPoint3.Close()
                    
                            'A2點數
                            Dim strPoint4 As String = "select count(*) as cnt from BillBase a"
                            strPoint4 = strPoint4 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnitScore("UnitID")) & "') "
                            strPoint4 = strPoint4 & " and a.RecordStateID=0 and (a.TrafficAccidentType='2') and a.BillTypeID='1'"
                            strPoint4 = strPoint4 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                            strPoint4 = strPoint4 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                            Dim CmdPoint4 As New Data.OracleClient.OracleCommand(strPoint4, conn)
                            Dim rdPoint4 As Data.OracleClient.OracleDataReader = CmdPoint4.ExecuteReader()
                            If rdPoint4.HasRows Then
                                rdPoint4.Read()
        
                                PersonPoint1 = PersonPoint1 + (CDec(rdPoint4("cnt")) * 50)
                                SubUnitPoint = SubUnitPoint + (CDec(rdPoint4("cnt")) * 50)
       
                            End If
                            rdPoint4.Close()
                    
                            'A3點數
                            Dim strPoint5 As String = "select count(*) as cnt from BillBase a"
                            strPoint5 = strPoint5 & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnitScore("UnitID")) & "') "
                            strPoint5 = strPoint5 & " and a.RecordStateID=0 and (a.TrafficAccidentType='3') and a.BillTypeID='1'"
                            strPoint5 = strPoint5 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                            strPoint5 = strPoint5 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                            Dim CmdPoint5 As New Data.OracleClient.OracleCommand(strPoint5, conn)
                            Dim rdPoint5 As Data.OracleClient.OracleDataReader = CmdPoint5.ExecuteReader()
                            If rdPoint5.HasRows Then
                                rdPoint5.Read()

                                PersonPoint1 = PersonPoint1 + (CDec(rdPoint5("cnt")) * 20)
                                SubUnitPoint = SubUnitPoint + (CDec(rdPoint5("cnt")) * 20)

                            End If
                            rdPoint5.Close()
                            
                            '拖吊
                            If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "台中市" Then
                                Dim strPointT1a As String = "select count(*) cnt from BillBase a"
                                strPointT1a = strPointT1a & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnitScore("UnitID")) & "') and a.RecordStateID=0"
                                strPointT1a = strPointT1a & " and a.ProjectID='A5'"
                                strPointT1a = strPointT1a & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPointT1a = strPointT1a & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPointT1a As New Data.OracleClient.OracleCommand(strPointT1a, conn)
                                Dim rdPointT1a As Data.OracleClient.OracleDataReader = CmdPointT1a.ExecuteReader()
                                If rdPointT1a.HasRows Then
                                    rdPointT1a.Read()
                                    If rdPointT1a("cnt") Is DBNull.Value Then
                                        PersonPoint1 = PersonPoint1
                                        SubUnitPoint = SubUnitPoint
                                    Else
                        
                                        PersonPoint1 = PersonPoint1 + (CDec(rdPointT1a("cnt")) * 20)
                                        SubUnitPoint = SubUnitPoint + (CDec(rdPointT1a("cnt")) * 20)
                                    End If
                                End If
                                rdPointT1a.Close()
                
                                Dim strPointT1b As String = "select count(*) cnt from BillBase a"
                                strPointT1b = strPointT1b & " where a.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdUnitScore("UnitID")) & "') and a.RecordStateID=0"
                                strPointT1b = strPointT1b & " and a.ProjectID='A6'"
                                strPointT1b = strPointT1b & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPointT1b = strPointT1b & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPointT1b As New Data.OracleClient.OracleCommand(strPointT1b, conn)
                                Dim rdPointT1b As Data.OracleClient.OracleDataReader = CmdPointT1b.ExecuteReader()
                                If rdPointT1b.HasRows Then
                                    rdPointT1b.Read()
                                    If rdPointT1b("cnt") Is DBNull.Value Then
                                        PersonPoint1 = PersonPoint1
                                        SubUnitPoint = SubUnitPoint
                                    Else
                        
                                        PersonPoint1 = PersonPoint1 + (CDec(rdPointT1b("cnt")) * 50)
                                        SubUnitPoint = SubUnitPoint + (CDec(rdPointT1b("cnt")) * 50)
                                    End If
                                End If
                                rdPointT1b.Close()
                            End If
                            
                            If Trim(rdUnitScore("ShowOrder")) = "2" And SubUnitPoint > 0 Then
                                If UnitIDTmp = "" Then
                                    UnitIDTmp = Trim(rdUnitScore("UnitID"))
                                Else
                                    UnitIDTmp = UnitIDTmp & "," & Trim(rdUnitScore("UnitID"))
                                End If
                                If UnitPointTmp = "" Then
                                    UnitPointTmp = SubUnitPoint
                                Else
                                    UnitPointTmp = UnitPointTmp & "," & SubUnitPoint
                                End If
                            End If
                            
                        End While
                    End If
                    rdUnitScore.Close()
                    '-----------先算分局----------------------------
                    Dim strUnit2 As String = "select * from UnitInfo where UnitID='" & Trim(rdSubUnit("UnitID")) & "'"
                    Dim CmdUnit2 As New Data.OracleClient.OracleCommand(strUnit2, conn)
                    Dim rdUnit2 As Data.OracleClient.OracleDataReader = CmdUnit2.ExecuteReader()
                    If rdUnit2.HasRows Then
                        rdUnit2.Read()
                        UnitName = Trim(rdUnit2("UnitName"))
                    End If
                    rdUnit2.Close()
                    
                    Dim strType3 As String = "select a.UnitID,a.CommonShareUnit,a.ShareGroupID,a.SharePercent,a.ChName"
                    strType3 = strType3 & " from CommonShareReward a"
                    strType3 = strType3 & " where a.UnitID='" & Trim(rdSubUnit("UnitID")) & "' and a.ShareGroupID=2"
                    strType3 = strType3 & " order by sn"
                    Dim CmdType3 As New Data.OracleClient.OracleCommand(strType3, conn)
                    Dim rdType3 As Data.OracleClient.OracleDataReader = CmdType3.ExecuteReader()
                    If rdType3.HasRows Then
                        While rdType3.Read()
                            CreditIDtmp = ""
                            UnitIDtmp2 = ""
                            LoginIDtmp = ""
                            MemberIDtmp = ""
                            BankIDtmp = ""
                            ShouldGetMoney = 0
                            MemMoney = 0
                            PageSum = 0
                            Response.Write("<table width=""640"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                            Response.Write("<tr>")
                            Response.Write("<td align=""center"" colspan=""6""><span class=""style1"">" & CityUnitName & UnitName & "</span></td>")
                            Response.Write("</tr>")
                            Response.Write("<tr>")
                            Response.Write("<td align=""right"" colspan=""6""><span class=""style1"">" & gOutDT2(Trim(Request("Date1"))) & "~" & gOutDT2(Trim(Request("Date2"))) & "&nbsp; &nbsp;共同人員交通安全獎金印領清冊</span>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 頁次：" & PageNo & "</td>")
                            Response.Write("</tr>")
                            Response.Write("<tr>")
                            Response.Write("<td colspan=""3"">列印日期：" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "</td>")
                            Response.Write("<td align=""right"" colspan=""3"">列印人員：" & UserName & "</td>")
                            Response.Write("</tr>")
                            Response.Write("</table>")
                        
                    
                            Response.Write("<table width=""640"" border=""1"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                            Response.Write("<tr>")
                            Response.Write("<td height=""35"" align=""center"" width=""8%"">單位</td>")
                            Response.Write("<td align=""center"" width=""12%"">職位</td>")
                            Response.Write("<td align=""center"" width=""12%"">姓名</td>")
                            Response.Write("<td align=""center"" width=""12%"">實領金額</td>")
                            If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                Response.Write("<td align=""center"" width=""13%"">身份證號</td>")
                            End If
                            Response.Write("<td align=""center"" width=""22%"">郵局帳號</td>")
                            Response.Write("<td align=""center"" width=""12%"">備註</td>")
                            Response.Write("</tr>")
                        
                            Response.Write("<tr>")
                            '單位
                            Response.Write("<td align=""center"" height=""35"">")
                            If rdType3("UnitID") IsNot DBNull.Value Then
                                Response.Write(rdType3("UnitID") & "&nbsp;")
                            Else
                                Response.Write("&nbsp;")
                            End If
                            Response.Write("</td>")
                            '職位
                            Response.Write("<td align=""center"">")
                            If rdType3("CommonShareUnit") IsNot DBNull.Value Then
                                If Trim(rdType3("CommonShareUnit")) = "" Then
                                    Response.Write("&nbsp;")
                                Else
                                    Response.Write(rdType3("CommonShareUnit"))
                                End If
                            Else
                                Response.Write("&nbsp;")
                            End If
                            Response.Write("</td>")
                            '姓名
                            Response.Write("<td align=""center"">")
                            If rdType3("ChName") IsNot DBNull.Value Then
                                Response.Write(rdType3("ChName"))
                            Else
                                Response.Write("&nbsp;")
                            End If
                            Response.Write("</td>")
                            '實領金額
                            Response.Write("<td align=""center"">")
                            If rdType3("ChName") IsNot DBNull.Value Then
                                Dim strMoney1 As String = "select * from MemberData where ChName='" & Trim(rdType3("ChName")) & "' order by RecordDate Desc"
                                Dim CmdMoney1 As New Data.OracleClient.OracleCommand(strMoney1, conn)
                                Dim rdMoney1 As Data.OracleClient.OracleDataReader = CmdMoney1.ExecuteReader()
                                If rdMoney1.HasRows Then
                                    rdMoney1.Read()
                                    If rdMoney1("Money") IsNot DBNull.Value Then
                                        MemMoney = Decimal.Truncate(rdMoney1("Money") * getPayPercent)
                                    Else
                                        MemMoney = 0
                                    End If
                                    
                                    If rdMoney1("CreditID") IsNot DBNull.Value Then
                                        CreditIDtmp = Trim(rdMoney1("CreditID"))
                                    Else
                                        CreditIDtmp = ""
                                    End If
                                    If rdMoney1("LoginID") IsNot DBNull.Value Then
                                        LoginIDtmp = Trim(rdMoney1("LoginID"))
                                    Else
                                        LoginIDtmp = ""
                                    End If
                                    If rdMoney1("UnitID") IsNot DBNull.Value Then
                                        UnitIDtmp2 = Trim(rdMoney1("UnitID"))
                                    Else
                                        UnitIDtmp2 = ""
                                    End If
                                    If rdMoney1("MemberID") IsNot DBNull.Value Then
                                        MemberIDtmp = Trim(rdMoney1("MemberID"))
                                    Else
                                        MemberIDtmp = ""
                                    End If
                                    If rdMoney1("BankName") IsNot DBNull.Value Then
                                        BankIDtmp = Trim(rdMoney1("BankName"))
                                    Else
                                        BankIDtmp = ""
                                    End If
                                    If rdMoney1("BankID") IsNot DBNull.Value Then
                                        BankIDtmp = BankIDtmp & Trim(rdMoney1("BankID"))
                                    Else
                                        BankIDtmp = BankIDtmp
                                    End If
                                    If rdMoney1("BankAccount") IsNot DBNull.Value Then
                                        BankIDtmp = BankIDtmp & Trim(rdMoney1("BankAccount"))
                                    Else
                                        BankIDtmp = BankIDtmp
                                    End If
                                End If
                                rdMoney1.Close()
                            Else
                                MemMoney = 0
                            End If
                            If MemMoney = 0 Then
                                If Trim(rdType3("ShareGroupID")) = 1 Then
                                    PersonMoney = Decimal.Round(Type1Money * rdType3("SharePercent"))
                                ElseIf Trim(rdType3("ShareGroupID")) = 2 Then
                                    PersonMoney = Decimal.Round(PersonPoint1 * Type2Money * rdType3("SharePercent"))
                                ElseIf Trim(rdType3("ShareGroupID")) = 3 Then
                                    PersonMoney = Decimal.Round(PersonPoint1 * Type3Money * rdType3("SharePercent"))
                                ElseIf Trim(rdType3("ShareGroupID")) = 4 Then
                                    PersonMoney = Decimal.Round(Type4Money * rdType3("SharePercent"))
                                End If
                                ShouldGetMoney = PersonMoney
                            Else
                                If Trim(rdType3("ShareGroupID")) = 1 Then
                                    PersonMoney = Decimal.Round(Type1Money * rdType3("SharePercent"))
                                ElseIf Trim(rdType3("ShareGroupID")) = 2 Then
                                    PersonMoney = Decimal.Round(PersonPoint1 * Type2Money * rdType3("SharePercent"))
                                ElseIf Trim(rdType3("ShareGroupID")) = 3 Then
                                    PersonMoney = Decimal.Round(PersonPoint1 * Type3Money * rdType3("SharePercent"))
                                ElseIf Trim(rdType3("ShareGroupID")) = 4 Then
                                    PersonMoney = Decimal.Round(Type4Money * rdType3("SharePercent"))
                                End If
                                ShouldGetMoney = PersonMoney
                                If PersonMoney > MemMoney Then
                                    PersonMoney = MemMoney
                                End If
                            End If
                            If MemberIDtmp <> "" And rdType3("ChName") IsNot DBNull.Value Then
                                RewardMonthData(CreditIDtmp, UnitIDtmp2, LoginIDtmp, MemberIDtmp, Trim(rdType3("ChName")), ShouldGetMoney, PersonMoney, UserID)
                            End If
                            MoneyTotal = MoneyTotal + PersonMoney
                            PageSum = PageSum + PersonMoney
                            Response.Write(PersonMoney)
                            Response.Write("</td>")
                            If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                Response.Write("<td>")
                                '身分證號
                                'If sys_City = "台中縣" Then
                                '    If CreditIDtmp <> "" Then
                                '        Response.Write(CreditIDtmp)
                                '    Else
                                '        Response.Write("&nbsp;")
                                '    End If
                                'Else
                                Response.Write("&nbsp;")
                                'End If
                                Response.Write("</td>")
                            End If
                            Response.Write("<td align=""center"">")
                            '郵局帳號
                            If sys_City = "台中縣" Or sys_City = "台中市" Then
                                If BankIDtmp <> "" Then
                                    Response.Write(BankIDtmp & "&nbsp;")
                                Else
                                    Response.Write("&nbsp;")
                                End If
                            Else
                                Response.Write("&nbsp;")
                            End If
                            Response.Write("</td>")
                            Response.Write("<td>&nbsp;</td>")
                            Response.Write("</tr>")
                            For i = 1 To PageCount
                                CreditIDtmp = ""
                                UnitIDtmp2 = ""
                                LoginIDtmp = ""
                                MemberIDtmp = ""
                                BankIDtmp = ""
                                ShouldGetMoney = 0
                                MemMoney = 0
                                If rdType3.Read() = True Then
                                    Response.Write("<tr>")
                                    '單位
                                    Response.Write("<td align=""center"" height=""35"">")
                                    If rdType3("UnitID") IsNot DBNull.Value Then
                                        Response.Write(rdType3("UnitID") & "&nbsp;")
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                    '職位
                                    Response.Write("<td align=""center"">")
                                    If rdType3("CommonShareUnit") IsNot DBNull.Value Then
                                        If Trim(rdType3("CommonShareUnit")) = "" Then
                                            Response.Write("&nbsp;")
                                        Else
                                            Response.Write(rdType3("CommonShareUnit"))
                                        End If
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                    '姓名
                                    Response.Write("<td align=""center"">")
                                    If rdType3("ChName") IsNot DBNull.Value Then
                                        Response.Write(rdType3("ChName"))
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                    '實領金額
                                    Response.Write("<td align=""center"">")
                                    If rdType3("ChName") IsNot DBNull.Value Then
                                        Dim strMoney1b As String = "select * from MemberData where ChName='" & Trim(rdType3("ChName")) & "' order by RecordDate Desc"
                                        Dim CmdMoney1b As New Data.OracleClient.OracleCommand(strMoney1b, conn)
                                        Dim rdMoney1b As Data.OracleClient.OracleDataReader = CmdMoney1b.ExecuteReader()
                                        If rdMoney1b.HasRows Then
                                            rdMoney1b.Read()
                                            If rdMoney1b("Money") IsNot DBNull.Value Then
                                                MemMoney = Decimal.Truncate(rdMoney1b("Money") * getPayPercent)
                                            Else
                                                MemMoney = 0
                                            End If
                                            
                                            If rdMoney1b("CreditID") IsNot DBNull.Value Then
                                                CreditIDtmp = Trim(rdMoney1b("CreditID"))
                                            Else
                                                CreditIDtmp = ""
                                            End If
                                            If rdMoney1b("LoginID") IsNot DBNull.Value Then
                                                LoginIDtmp = Trim(rdMoney1b("LoginID"))
                                            Else
                                                LoginIDtmp = ""
                                            End If
                                            If rdMoney1b("UnitID") IsNot DBNull.Value Then
                                                UnitIDtmp2 = Trim(rdMoney1b("UnitID"))
                                            Else
                                                UnitIDtmp2 = ""
                                            End If
                                            If rdMoney1b("MemberID") IsNot DBNull.Value Then
                                                UnitIDtmp2 = Trim(rdMoney1b("MemberID"))
                                            Else
                                                UnitIDtmp2 = ""
                                            End If
                                            If rdMoney1b("BankName") IsNot DBNull.Value Then
                                                BankIDtmp = Trim(rdMoney1b("BankName"))
                                            Else
                                                BankIDtmp = ""
                                            End If
                                            If rdMoney1b("BankID") IsNot DBNull.Value Then
                                                BankIDtmp = BankIDtmp & Trim(rdMoney1b("BankID"))
                                            Else
                                                BankIDtmp = BankIDtmp
                                            End If
                                            If rdMoney1b("BankAccount") IsNot DBNull.Value Then
                                                BankIDtmp = BankIDtmp & Trim(rdMoney1b("BankAccount"))
                                            Else
                                                BankIDtmp = BankIDtmp
                                            End If
                                        End If
                                        rdMoney1b.Close()
                                    Else
                                        MemMoney = 0
                                    End If
                                    If Trim(rdType3("ShareGroupID")) = 1 Then
                                        PersonMoney = Decimal.Round(Type1Money * rdType3("SharePercent"))
                                    ElseIf Trim(rdType3("ShareGroupID")) = 2 Then
                                        PersonMoney = Decimal.Round(PersonPoint1 * Type2Money * rdType3("SharePercent"))
                                    ElseIf Trim(rdType3("ShareGroupID")) = 3 Then
                                        PersonMoney = Decimal.Round(PersonPoint1 * Type3Money * rdType3("SharePercent"))
                                    ElseIf Trim(rdType3("ShareGroupID")) = 4 Then
                                        PersonMoney = Decimal.Round(Type4Money * rdType3("SharePercent"))
                                    End If
                                    If MemMoney > 0 Then
                                        If PersonMoney > MemMoney Then
                                            PersonMoney = MemMoney
                                        End If
                                    End If
                                    ShouldGetMoney = PersonMoney
                                    If MemberIDtmp <> "" And rdType3("ChName") IsNot DBNull.Value Then
                                        RewardMonthData(CreditIDtmp, UnitIDtmp2, LoginIDtmp, MemberIDtmp, Trim(rdType3("ChName")), ShouldGetMoney, PersonMoney, UserID)
                                    End If
                                    MoneyTotal = MoneyTotal + PersonMoney
                                    PageSum = PageSum + PersonMoney
                                    Response.Write(PersonMoney)
                                    Response.Write("</td>")
                                    If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                        Response.Write("<td>")
                                        '身分證號
                                        'If sys_City = "台中縣" Then
                                        '    If CreditIDtmp <> "" Then
                                        '        Response.Write(CreditIDtmp)
                                        '    Else
                                        '        Response.Write("&nbsp;")
                                        '    End If
                                        'Else
                                        Response.Write("&nbsp;")
                                        'End If
                                        Response.Write("</td>")
                                    End If
                                    Response.Write("<td align=""center"">")
                                    '郵局帳號
                                    If sys_City = "台中縣" Or sys_City = "台中市" Then
                                        If BankIDtmp <> "" Then
                                            Response.Write(BankIDtmp & "&nbsp;")
                                        Else
                                            Response.Write("&nbsp;")
                                        End If
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                    Response.Write("<td>&nbsp;</td>")
                                    Response.Write("</tr>")
                                Else
                                    Exit For
                                End If
                            Next
                            Response.Write("<tr>")
                            Response.Write("<td align=""center"" height=""35"">小計</td>")
                            Response.Write("<td colspan=""2"">&nbsp;</td>")
                            Response.Write("<td align=""center"">")
                            Response.Write(PageSum)
                            Response.Write("</td>")
                            If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                Response.Write("<td colspan=""3"">&nbsp;</td>")
                            Else
                                Response.Write("<td colspan=""2"">&nbsp;</td>")
                            End If
                            Response.Write("</tr>")
                            Response.Write("</table>")
                            Response.Write("<table width=""640"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                            Response.Write("<tr><td colspan=""6""><strong>承辦單位：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                            If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                '    Response.Write("人事：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                '    Response.Write("出納：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                '    Response.Write("會計：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                '    Response.Write("機關主官：</strong></td>")
                                'Else
                                Response.Write("秘書室：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                Response.Write("會計室：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                Response.Write("局長：</strong></td>")
                            End If
                            Response.Write("</tr>")
                            Response.Write("</table>")
                            PageNo = PageNo + 1
                            Response.Write("<div class=""PageNext""></div>")
                            
                        End While
                        rdType3.Close()
                        
                        '----------算分局下派出所----------------
                        Dim Type3Cnt As Integer = 0
                        Dim strType3Cnt As String = "select count(*) as cnt"
                        strType3Cnt = strType3Cnt & " from CommonShareReward a"
                        strType3Cnt = strType3Cnt & " where a.UnitID='" & Trim(rdSubUnit("UnitID")) & "' and a.ShareGroupID=3"
                        Dim CmdType3Cnt As New Data.OracleClient.OracleCommand(strType3Cnt, conn)
                        Dim rdType3Cnt As Data.OracleClient.OracleDataReader = CmdType3Cnt.ExecuteReader()
                        If rdType3Cnt.HasRows Then
                            rdType3Cnt.Read()
                            Type3Cnt = rdType3Cnt("cnt")
                        End If
                        rdType3Cnt.Close()
                        UnitSN = 0
                        Dim UnitIDArray = Split(UnitIDTmp, ",")
                        Dim UnitPointArray = Split(UnitPointTmp, ",")
                        For i = 0 To UBound(UnitIDArray)
                            j = 0
                            Dim strType3b As String = "select a.UnitID,a.CommonShareUnit,a.ShareGroupID,a.SharePercent,a.ChName"
                            strType3b = strType3b & " from CommonShareReward a"
                            strType3b = strType3b & " where a.UnitID='" & Trim(rdSubUnit("UnitID")) & "' and a.ShareGroupID=3"
                            strType3b = strType3b & " order by sn"
                            Dim CmdType3b As New Data.OracleClient.OracleCommand(strType3b, conn)
                            Dim rdType3b As Data.OracleClient.OracleDataReader = CmdType3b.ExecuteReader()
                            If rdType3b.HasRows Then
                                While rdType3b.Read()
                                    ChNametmp = ""
                                    CreditIDtmp = ""
                                    UnitIDtmp2 = ""
                                    LoginIDtmp = ""
                                    MemberIDtmp = ""
                                    BankIDtmp = ""
                                    ShouldGetMoney = 0
                                    MemMoney = 0
                                    If UnitSN = 0 Then
                                        PageSum = 0
                                        Response.Write("<table width=""640"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                                        Response.Write("<tr>")
                                        Response.Write("<td align=""center"" colspan=""6""><span class=""style1"">" & CityUnitName & UnitName & "</span></td>")
                                        Response.Write("</tr>")
                                        Response.Write("<tr>")
                                        Response.Write("<td align=""right"" colspan=""6""><span class=""style1"">" & gOutDT2(Trim(Request("Date1"))) & "~" & gOutDT2(Trim(Request("Date2"))) & "&nbsp; &nbsp;共同人員交通安全獎金印領清冊</span>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 頁次：" & PageNo & "</td>")
                                        Response.Write("</tr>")
                                        Response.Write("<tr>")
                                        Response.Write("<td colspan=""3"">列印日期：" & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & "</td>")
                                        Response.Write("<td align=""right"" colspan=""3"">列印人員：" & UserName & "</td>")
                                        Response.Write("</tr>")
                                        Response.Write("</table>")
                        
                                        Response.Write("<table width=""640"" border=""1"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                                        Response.Write("<tr>")
                                        Response.Write("<td height=""35"" align=""center"" width=""8%"">單位</td>")
                                        Response.Write("<td align=""center"" width=""12%"">職位</td>")
                                        Response.Write("<td align=""center"" width=""12%"">姓名</td>")
                                        Response.Write("<td align=""center"" width=""12%"">實領金額</td>")
                                        If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                            Response.Write("<td align=""center"" width=""13%"">身份證號</td>")
                                        End If
                                        Response.Write("<td align=""center"" width=""22%"">郵局帳號</td>")
                                        Response.Write("<td align=""center"" width=""12%"">備註</td>")
                                        Response.Write("</tr>")
                                    End If
                                    
                                    Response.Write("<tr>")
                                    '單位
                                    Response.Write("<td align=""center"" height=""35"">")
                                    If rdType3b("UnitID") IsNot DBNull.Value Then
                                        Response.Write(UnitIDArray(i) & "&nbsp;")
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                    '職位
                                    Response.Write("<td align=""center"">")
                                    If rdType3b("CommonShareUnit") IsNot DBNull.Value Then
                                        If Trim(rdType3b("CommonShareUnit")) = "" Then
                                            Response.Write("&nbsp;")
                                        Else
                                            Response.Write(rdType3b("CommonShareUnit"))
                                        End If
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                    '姓名
                                    Response.Write("<td align=""center"">")
                                    Dim strMoney1b As String
                                    If rdType3b("CommonShareUnit") IsNot DBNull.Value Then
                                        If Trim(rdType3b("CommonShareUnit")) = "所長" Then
                                            strMoney1b = "select * from MemberData where UnitID='" & UnitIDArray(i) & "' and JobID=314 order by RecordDate Desc"
                                        ElseIf Trim(rdType3b("CommonShareUnit")) = "副所長" Then
                                            strMoney1b = "select * from MemberData where UnitID='" & UnitIDArray(i) & "' and JobID=315 order by RecordDate Desc"
                                        Else
                                            strMoney1b = "select * from MemberData where UnitID='" & UnitIDArray(i) & "' and JobID=314 order by RecordDate Desc"
                                        End If
                                    Else
                                        strMoney1b = "select * from MemberData where UnitID='" & UnitIDArray(i) & "' and JobID=314 order by RecordDate Desc"
                                    End If
                                    Dim CmdMoney1b As New Data.OracleClient.OracleCommand(strMoney1b, conn)
                                    Dim rdMoney1b As Data.OracleClient.OracleDataReader = CmdMoney1b.ExecuteReader()
                                    If rdMoney1b.HasRows Then
                                        rdMoney1b.Read()
                                        If rdMoney1b("Money") IsNot DBNull.Value Then
                                            MemMoney = Decimal.Truncate(rdMoney1b("Money") * getPayPercent)
                                        Else
                                            MemMoney = 0
                                        End If
                                        If rdMoney1b("ChName") IsNot DBNull.Value Then
                                            Response.Write(rdMoney1b("ChName"))
                                        Else
                                            Response.Write("&nbsp;")
                                        End If
                                        
                                        If rdMoney1b("CreditID") IsNot DBNull.Value Then
                                            CreditIDtmp = Trim(rdMoney1b("CreditID"))
                                        Else
                                            CreditIDtmp = ""
                                        End If
                                        If rdMoney1b("LoginID") IsNot DBNull.Value Then
                                            LoginIDtmp = Trim(rdMoney1b("LoginID"))
                                        Else
                                            LoginIDtmp = ""
                                        End If
                                        If rdMoney1b("UnitID") IsNot DBNull.Value Then
                                            UnitIDtmp2 = Trim(rdMoney1b("UnitID"))
                                        Else
                                            UnitIDtmp2 = ""
                                        End If
                                        If rdMoney1b("MemberID") IsNot DBNull.Value Then
                                            MemberIDtmp = Trim(rdMoney1b("MemberID"))
                                        Else
                                            MemberIDtmp = ""
                                        End If
                                        If rdMoney1b("ChName") IsNot DBNull.Value Then
                                            ChNametmp = Trim(rdMoney1b("ChName"))
                                        Else
                                            ChNametmp = ""
                                        End If
                                        If rdMoney1b("BankName") IsNot DBNull.Value Then
                                            BankIDtmp = Trim(rdMoney1b("BankName"))
                                        Else
                                            BankIDtmp = ""
                                        End If
                                        If rdMoney1b("BankID") IsNot DBNull.Value Then
                                            BankIDtmp = BankIDtmp & Trim(rdMoney1b("BankID"))
                                        Else
                                            BankIDtmp = BankIDtmp
                                        End If
                                        If rdMoney1b("BankAccount") IsNot DBNull.Value Then
                                            BankIDtmp = BankIDtmp & Trim(rdMoney1b("BankAccount"))
                                        Else
                                            BankIDtmp = BankIDtmp
                                        End If
                                    Else
                                        MemMoney = 0
                                        Response.Write("&nbsp;")
                                    End If
                                    rdMoney1b.Close()

                                    Response.Write("</td>")
                                    '實領金額
                                    Response.Write("<td align=""center"">")
                                    If Trim(rdType3b("ShareGroupID")) = 1 Then
                                        PersonMoney = Decimal.Round(Type1Money * rdType3b("SharePercent"))
                                    ElseIf Trim(rdType3b("ShareGroupID")) = 2 Then
                                        PersonMoney = Decimal.Round(UnitPointArray(i) * Type2Money * rdType3b("SharePercent"))
                                    ElseIf Trim(rdType3b("ShareGroupID")) = 3 Then
                                        PersonMoney = Decimal.Round(UnitPointArray(i) * Type3Money * rdType3b("SharePercent"))
                                    ElseIf Trim(rdType3b("ShareGroupID")) = 4 Then
                                        PersonMoney = Decimal.Round(Type4Money * rdType3b("SharePercent"))
                                    End If
                                    ShouldGetMoney = PersonMoney
                                    If MemMoney > 0 Then
                                        If PersonMoney > MemMoney Then
                                            PersonMoney = MemMoney
                                        End If
                                    End If
                                    If MemberIDtmp <> "" Then
                                        RewardMonthData(CreditIDtmp, UnitIDtmp2, LoginIDtmp, MemberIDtmp, ChNametmp, ShouldGetMoney, PersonMoney, UserID)
                                    End If
                                    MoneyTotal = MoneyTotal + PersonMoney
                                    PageSum = PageSum + PersonMoney
                                    Response.Write(PersonMoney)
                                    Response.Write("</td>")
                                    If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                        Response.Write("<td>")
                                        '身分證號
                                        'If sys_City = "台中縣" Then
                                        '    If CreditIDtmp <> "" Then
                                        '        Response.Write(CreditIDtmp)
                                        '    Else
                                        '        Response.Write("&nbsp;")
                                        '    End If
                                        'Else
                                        Response.Write("&nbsp;")
                                        'End If
                                        Response.Write("</td>")
                                    End If
                                    Response.Write("<td align=""center"">")
                                    '郵局帳號
                                    If sys_City = "台中縣" Or sys_City = "台中市" Then
                                        If BankIDtmp <> "" Then
                                            Response.Write(BankIDtmp & "&nbsp;")
                                        Else
                                            Response.Write("&nbsp;")
                                        End If
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                    Response.Write("<td>&nbsp;</td>")
                                    Response.Write("</tr>")
                                    
                                    
                                    UnitSN = UnitSN + 1
                                    j = j + 1
                                    If UnitSN = PageCount + 1 Or (i = UBound(UnitIDArray) And Type3Cnt = j) Then
                                        Response.Write("<tr>")
                                        Response.Write("<td align=""center"" height=""35"">小計</td>")
                                        Response.Write("<td colspan=""2"">&nbsp;</td>")
                                        Response.Write("<td align=""center"">")
                                        Response.Write(PageSum)
                                        Response.Write("</td>")
                                        If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                            Response.Write("<td colspan=""3"">&nbsp;</td>")
                                        Else
                                            Response.Write("<td colspan=""2"">&nbsp;</td>")
                                        End If
                                        Response.Write("</tr>")
                                        Response.Write("</table>")
                                        Response.Write("<table width=""640"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")

                                        If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                            '    Response.Write("<tr><td colspan=""6""><strong>承辦單位：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")

                                            '    Response.Write("人事：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                            '    Response.Write("出納：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                            '    Response.Write("會計：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                            '    Response.Write("機關主官：</strong></td>")
                                            'Else
                                            Response.Write("<tr><td colspan=""6""><strong>承辦單位：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")

                                            Response.Write("秘書室：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                            Response.Write("會計室：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                                            Response.Write("局長：</strong></td>")
                                        End If
                                        Response.Write("</tr>")
                                        Response.Write("</table>")
                                        PageNo = PageNo + 1
                                        Response.Write("<div class=""PageNext""></div>")
                                        
                                        UnitSN = 0
                                    End If
                                    
                                    
                                End While
                            End If
                            rdType3b.Close()
                        Next

                    End If
                End While
            End If
            rdSubUnit.Close()
            
            Response.Write("<table width=""640"" border=""1"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
            Response.Write("<tr>")
                                    
            Response.Write("<td style=""height:35px"" colspan=""2"" ALIGN=""center"">&nbsp;</td>")
            Response.Write("<td  ALIGN=""center"" colspan=""2"">實領獎金</td>")
            Response.Write("<td  ALIGN=""center"" colspan=""2"">備註</td>")
            Response.Write("</tr>")
            Response.Write("<tr>")

            Response.Write("<td style=""height:35px"" colspan=""2"" ALIGN=""center"">總計</td>")
            Response.Write("<td ALIGN=""center"" colspan=""2"">" & MoneyTotal & "</td>")
            Response.Write("<td ALIGN=""center"" colspan=""2"">&nbsp;</td>")

            Response.Write("</tr></table>")
            If sys_City = "台中縣" Or sys_City = "台中市" Then
                Response.Write("<tr><td colspan=""6""><strong>承辦單位：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")

                Response.Write("人事：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                Response.Write("出納：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                Response.Write("會計：&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ")
                Response.Write("機關主官：</strong></td>")

            End If
            'Response.Write(Request("sUnitID"))
            conn.Close()
        %>
    
        
    </form>
</body>
<script type="text/javascript" src="../form.js"></script>
<script language="JavaScript">
	printWindow(true,5.08,5.08,5.08,5.08);
</script>
</html>