<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%  
    LoginCheck()
%>
<script runat="server">
    Public UnitPoint, PointTotal, MoneyTotal, MemMoney, MemPay, UnitMoney As Decimal
    
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
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim fMnoth = Month(Now)
        If fMnoth < 10 Then fMnoth = "0" & fMnoth
        Dim fDay = Day(Now)
        If fDay < 10 Then fDay = "0" & fDay
        Dim fname = Year(Now) & fMnoth & fDay & ".xls"
        Response.AddHeader("Content-Disposition", "filename=" & fname)
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8")
        Response.ContentType = "application/x-msexcel; charset=MS950"
    End Sub
    
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
</script>
<script language="JavaScript">
	window.focus();
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<style type="text/css">
<!--
body {font-family:新細明體;font-size:10pt; FONT-WEIGHT: bold }

.style1 {font-family:新細明體; font-size: 11pt}
-->
</style>
    <title>支領獎勵金核發清冊</title>
</head>
<body>
    <form id="form1" runat="server">
    <table width="680" border="1" cellpadding="3" cellspacing="0" align="center">
        <tr>
        <td align="center" colspan="4"><span class="style1"><strong>支&nbsp; 領&nbsp; 獎 &nbsp;勵&nbsp; 金&nbsp; 核&nbsp; 發&nbsp; 清&nbsp;冊</strong></span></td>
        </tr>
        <tr>
        <td style="width: 40%" align="left">單位名稱</td>
        <td style="width: 20%" align="right">舉發件數</td>
        <td style="width: 20%" align="right">點數</td>
        <td style="width: 20%" align="right">金額</td>
        </tr>
        <%
            '取得 Web.config 檔的資料連接設定
            Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
            '建立 Connection 物件
            Dim conn As New Data.OracleClient.OracleConnection()
            conn.ConnectionString = setting.ConnectionString
            '開啟資料連接
            conn.Open()
            
            Dim sys_City As String
            '要用填單或建檔日統計
            Dim theDateType As String = Trim(Request("DateType"))
            '================================================
            '獎勵金總額
            Dim getMoneyTotal As Decimal
            If Trim(Request("AnalyzeMoney")) = "" Then
                getMoneyTotal = 0
            Else
                getMoneyTotal = CDec(Request("AnalyzeMoney"))
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
            
            sys_City = ""
            Dim strCity = "select Value from ApConfigure where ID=31"
            Dim CmdCity As New Data.OracleClient.OracleCommand(strCity, conn)
            Dim rdCity As Data.OracleClient.OracleDataReader = CmdCity.ExecuteReader()
            If rdCity.HasRows Then
                rdCity.Read()
                sys_City = Trim(rdCity("Value"))
            End If
            rdCity.Close()
            '====================計算每點多少錢=======================
            Dim getPointTotal As Decimal
            getPointTotal = 0
            Dim FlagPointTotal As String
            FlagPointTotal = Trim(Request("AnalyzeType"))
            Dim AnalyzeUnitID As String = ""
            
            Dim strUID = "select Value from Apconfigure where ID=49"
            Dim CmdUID As New Data.OracleClient.OracleCommand(strUID, conn)
            Dim rsUID As Data.OracleClient.OracleDataReader = CmdUID.ExecuteReader()
            If rsUID.HasRows Then
                rsUID.Read()
                AnalyzeUnitID = Trim(rsUID("Value"))
            End If
            rsUID.Close()
            
            Dim UserCookie As HttpCookie = Request.Cookies("RewardUser")
            Dim AnalyzeUnitID2 As String = ""
            AnalyzeUnitID2 = Trim(UserCookie.Values("UnitID"))
            
            Dim UnitFlag, strPointT6Plus, strPointT6Plus2 As String
            If FlagPointTotal = 1 Then  '總分只抓交通隊
                UnitFlag = " and a.BillUnitID='" & AnalyzeUnitID & "'"
            ElseIf FlagPointTotal = 2 Then  '總分只抓分局
                UnitFlag = " and a.BillUnitID in (select UnitID from UnitInfo where UnitID='" & AnalyzeUnitID2 & "' or UnitTypeID='" & AnalyzeUnitID2 & "')"
            Else
                UnitFlag = ""
            End If

            If sys_City = "台東縣" Then
                strPointT6Plus = " and ((a.CarSimpleID in (3,4) and b.CarSimpleID=3) or(a.CarSimpleID=1 and b.CarSimpleID=5) or(a.CarSimpleID=2 and b.CarSimpleID=6))"
                strPointT6Plus2 = ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed,CarSimpleID from LawScore"

            Else
                strPointT6Plus = ""
                strPointT6Plus2 = ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
            End If
            
            '雲林拖吊案件分數較高
            If sys_City = "雲林縣" Or sys_City = "台東縣" Or sys_City = "宜蘭縣" Then
                '攔停點數
                Dim strPointT1 As String = "select sum(b.BillType1Score) as cnt from BillBaseViewReward a"
                strPointT1 = strPointT1 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
                strPointT1 = strPointT1 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
                strPointT1 = strPointT1 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion" & UnitFlag
                strPointT1 = strPointT1 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                strPointT1 = strPointT1 & " and (a.CarAddID<>8 or a.CarAddID is null) and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                strPointT1 = strPointT1 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPointT1 = strPointT1 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPointT1 As New Data.OracleClient.OracleCommand(strPointT1, conn)
                Dim rdPointT1 As Data.OracleClient.OracleDataReader = CmdPointT1.ExecuteReader()
                If rdPointT1.HasRows Then
                    rdPointT1.Read()
                    If rdPointT1("cnt") Is DBNull.Value Then
                        getPointTotal = getPointTotal + 0
                    Else
                        getPointTotal = getPointTotal + CDec(rdPointT1("cnt"))
                    End If
                End If
                rdPointT1.Close()

                '逕舉點數
                Dim strPointT2 As String = "select sum(b.BillType2Score) as cnt from BillBaseViewReward a"
                strPointT2 = strPointT2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
                strPointT2 = strPointT2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
                strPointT2 = strPointT2 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion" & UnitFlag
                strPointT2 = strPointT2 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                strPointT2 = strPointT2 & " and (a.CarAddID<>8 or a.CarAddID is null) and a.BillTypeID='2' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
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
                    
                '拖吊點數
                Dim strPointT6 As String = "select sum(b.Other1) as cnt from BillBase a" & strPointT6Plus2
                strPointT6 = strPointT6 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
                strPointT6 = strPointT6 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion" & UnitFlag
                strPointT6 = strPointT6 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                strPointT6 = strPointT6 & " and a.CarAddID=8 and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                strPointT6 = strPointT6 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPointT6 = strPointT6 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')" & strPointT6Plus
                Dim CmdPointT6 As New Data.OracleClient.OracleCommand(strPointT6, conn)
                Dim rdPointT6 As Data.OracleClient.OracleDataReader = CmdPointT6.ExecuteReader()
                If rdPointT6.HasRows Then
                    rdPointT6.Read()
                    If rdPointT6("cnt") Is DBNull.Value Then
                        getPointTotal = getPointTotal + 0
                    Else
                        getPointTotal = getPointTotal + CDec(rdPointT6("cnt"))
                    End If
                End If
                rdPointT6.Close()
            Else
                '攔停點數
                Dim strPointT1 As String = "select sum(b.BillType1Score) as cnt from BillBaseViewReward a"
                strPointT1 = strPointT1 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
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
                        getPointTotal = getPointTotal + 0
                    Else
                        getPointTotal = getPointTotal + CDec(rdPointT1("cnt"))
                    End If
                End If
                rdPointT1.Close()

                '逕舉點數
                Dim strPointT2 As String = "select sum(b.BillType2Score) as cnt from BillBaseViewReward a"
                strPointT2 = strPointT2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
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
            End If
                
            'A1點數
            Dim strPointT3 As String = "select b.A1Score,b.BillType1Score from BillBaseViewReward a"
            strPointT3 = strPointT3 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
            strPointT3 = strPointT3 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
            strPointT3 = strPointT3 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion " & UnitFlag
            strPointT3 = strPointT3 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
            strPointT3 = strPointT3 & " and a.RecordStateID=0 and (a.TrafficAccidentType='1') and a.BillTypeID='1'"
            strPointT3 = strPointT3 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPointT3 = strPointT3 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdPointT3 As New Data.OracleClient.OracleCommand(strPointT3, conn)
            Dim rdPointT3 As Data.OracleClient.OracleDataReader = CmdPointT3.ExecuteReader()
            If rdPointT3.HasRows Then
                While rdPointT3.Read()
                    If rdPointT3("BillType1Score") > rdPointT3("A1Score") Then
                        If rdPointT3("BillType1Score") Is DBNull.Value Then
                            getPointTotal = getPointTotal + 0
                        Else
                            getPointTotal = getPointTotal + CDec(rdPointT3("BillType1Score"))
                        End If
                    Else
                        If rdPointT3("A1Score") Is DBNull.Value Then
                            getPointTotal = getPointTotal + 0
                        Else
                            getPointTotal = getPointTotal + CDec(rdPointT3("A1Score"))
                        End If
                    End If
                End While
            End If
            rdPointT3.Close()
            
            'A2點數
            Dim strPointT4 As String = "select b.A2Score,b.BillType1Score from BillBaseViewReward a"
            strPointT4 = strPointT4 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
            strPointT4 = strPointT4 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
            strPointT4 = strPointT4 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion " & UnitFlag
            strPointT4 = strPointT4 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
            strPointT4 = strPointT4 & " and a.RecordStateID=0 and (a.TrafficAccidentType='2') and a.BillTypeID='1'"
            strPointT4 = strPointT4 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPointT4 = strPointT4 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdPointT4 As New Data.OracleClient.OracleCommand(strPointT4, conn)
            Dim rdPointT4 As Data.OracleClient.OracleDataReader = CmdPointT4.ExecuteReader()
            If rdPointT4.HasRows Then
                While rdPointT4.Read()
                    If rdPointT4("BillType1Score") > rdPointT4("A2Score") Then
                        If rdPointT4("BillType1Score") Is DBNull.Value Then
                            getPointTotal = getPointTotal + 0
                        Else
                            getPointTotal = getPointTotal + CDec(rdPointT4("BillType1Score"))
                        End If
                    Else
                        If rdPointT4("A2Score") Is DBNull.Value Then
                            getPointTotal = getPointTotal + 0
                        Else
                            getPointTotal = getPointTotal + CDec(rdPointT4("A2Score"))
                        End If
                    End If
                End While
            End If
            rdPointT4.Close()
                    
            'A3點數
            Dim strPointT5 As String = "select b.A3Score,b.BillType1Score from BillBaseViewReward a"
            strPointT5 = strPointT5 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
            strPointT5 = strPointT5 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
            strPointT5 = strPointT5 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion " & UnitFlag
            strPointT5 = strPointT5 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
            strPointT5 = strPointT5 & " and a.RecordStateID=0 and (a.TrafficAccidentType='3') and a.BillTypeID='1'"
            strPointT5 = strPointT5 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPointT5 = strPointT5 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdPointT5 As New Data.OracleClient.OracleCommand(strPointT5, conn)
            Dim rdPointT5 As Data.OracleClient.OracleDataReader = CmdPointT5.ExecuteReader()
            If rdPointT5.HasRows Then
                While rdPointT5.Read()
                    If rdPointT5("BillType1Score") > rdPointT5("A3Score") Then
                        If rdPointT5("BillType1Score") Is DBNull.Value Then
                            getPointTotal = getPointTotal + 0
                        Else
                            getPointTotal = getPointTotal + CDec(rdPointT5("BillType1Score"))
                        End If
                    Else
                        If rdPointT5("A3Score") Is DBNull.Value Then
                            getPointTotal = getPointTotal + 0
                        Else
                            getPointTotal = getPointTotal + CDec(rdPointT5("A3Score"))
                        End If
                    End If
                        
                End While
            End If
            rdPointT5.Close()

            Dim PointMoney As Decimal

            If getPointTotal = 0 Then
                PointMoney = 0
            Else
                PointMoney = getMoneyTotal / getPointTotal
            End If
            
            'Response.Write(PointMoney & "," & getPointTotal)

            '===================列出清冊內容========================
            '---------------所有單位-----------------
            Dim PersonPoint, A1ScoreTmp, A2ScoreTmp, A3ScoreTmp, BillCnt, BillCntTotal As Decimal
            Dim OverFlag As String
            PointTotal = 0
            MoneyTotal = 0
            BillCntTotal = 0
            Dim strUnit = "select * from UnitInfo where UnitID in (" & Trim(Request("sUnitID")) & ") order by UnitID"
            Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
            Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
            If rdUnit.HasRows Then
                While rdUnit.Read()
                    Response.Write("<tr>")
                    Response.Write("<td>" & Trim(rdUnit("UnitID")) & "&nbsp; " & Trim(rdUnit("UnitName")) & "</td>")
                    '抓出此單位所有人來算,只抓直接人員
                    UnitPoint = 0
                    UnitMoney = 0
                    BillCnt = 0
                    Dim strPer = "select MemberID,Money from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "'"
                    Dim CmdPer As New Data.OracleClient.OracleCommand(strPer, conn)
                    Dim rdPer As Data.OracleClient.OracleDataReader = CmdPer.ExecuteReader()
                    If rdPer.HasRows Then
                        While rdPer.Read()
                            OverFlag = ""
                            PersonPoint = 0
                            MemMoney = 0
                            A1ScoreTmp = 0
                            A2ScoreTmp = 0
                            A3ScoreTmp = 0
                            '雲林拖吊案件分數較高
                            If sys_City = "雲林縣" Or sys_City = "台東縣" Or sys_City = "宜蘭縣" Then
                                '攔停點數
                                Dim strPoint1 As String = "select b.BillType1Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBaseViewReward a"
                                strPoint1 = strPoint1 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
                                strPoint1 = strPoint1 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                                strPoint1 = strPoint1 & " where (a.BillMemID1=" & Trim(rdPer("MemberID")) & " or a.BillMemID2=" & Trim(rdPer("MemberID")) & " or a.BillMemID3=" & Trim(rdPer("MemberID")) & " or a.BillMemID4=" & Trim(rdPer("MemberID")) & ")"
                                strPoint1 = strPoint1 & " and a.RuleVer=b.LawVersion"
                                strPoint1 = strPoint1 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                                strPoint1 = strPoint1 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                                strPoint1 = strPoint1 & " and (a.CarAddID<>8 or a.CarAddID is null) and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPoint1 = strPoint1 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPoint1 As New Data.OracleClient.OracleCommand(strPoint1, conn)
                                Dim rdPoint1 As Data.OracleClient.OracleDataReader = CmdPoint1.ExecuteReader()
                                If rdPoint1.HasRows Then
                                    While rdPoint1.Read()
                                        If rdPoint1("BillType1Score") Is DBNull.Value Then
                                            PersonPoint = PersonPoint + 0
                                            BillCnt = BillCnt + 0
                                        Else
                                            If rdPoint1("BillMemID1") IsNot DBNull.Value And rdPoint1("BillMemID2") Is DBNull.Value And rdPoint1("BillMemID3") Is DBNull.Value And rdPoint1("BillMemID4") Is DBNull.Value Then
                                                PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score"))
                                                BillCnt = BillCnt + 1
                                            ElseIf rdPoint1("BillMemID1") IsNot DBNull.Value And rdPoint1("BillMemID2") IsNot DBNull.Value And rdPoint1("BillMemID3") Is DBNull.Value And rdPoint1("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdPoint1("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") / 2)
                                                    BillCnt = BillCnt + 0.5
                                                End If
                                                If Trim(rdPoint1("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") / 2)
                                                    BillCnt = BillCnt + 0.5
                                                End If
                                            ElseIf rdPoint1("BillMemID1") IsNot DBNull.Value And rdPoint1("BillMemID2") IsNot DBNull.Value And rdPoint1("BillMemID3") IsNot DBNull.Value And rdPoint1("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdPoint1("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") * 0.34)
                                                    BillCnt = BillCnt + 0.34
                                                End If
                                                If Trim(rdPoint1("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") * 0.33)
                                                    BillCnt = BillCnt + 0.33
                                                End If
                                                If Trim(rdPoint1("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") * 0.33)
                                                    BillCnt = BillCnt + 0.33
                                                End If
                                            ElseIf rdPoint1("BillMemID1") IsNot DBNull.Value And rdPoint1("BillMemID2") IsNot DBNull.Value And rdPoint1("BillMemID3") IsNot DBNull.Value And rdPoint1("BillMemID4") IsNot DBNull.Value Then
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint1("BillMemID1")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint1("BillMemID2")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint1("BillMemID3")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint1("BillMemID4")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                            End If
                                        End If
                                    End While
                                End If
                                rdPoint1.Close()
                                
                                '逕舉點數
                                Dim strPoint2 As String = "select b.BillType2Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBaseViewReward a"
                                strPoint2 = strPoint2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
                                strPoint2 = strPoint2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                                strPoint2 = strPoint2 & " where (a.BillMemID1=" & Trim(rdPer("MemberID")) & " or a.BillMemID2=" & Trim(rdPer("MemberID")) & " or a.BillMemID3=" & Trim(rdPer("MemberID")) & " or a.BillMemID4=" & Trim(rdPer("MemberID")) & ")"
                                strPoint2 = strPoint2 & " and a.RuleVer=b.LawVersion"
                                strPoint2 = strPoint2 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                                strPoint2 = strPoint2 & " and (a.CarAddID<>8 or a.CarAddID is null) and a.BillTypeID='2' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                                strPoint2 = strPoint2 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPoint2 = strPoint2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPoint2 As New Data.OracleClient.OracleCommand(strPoint2, conn)
                                Dim rdPoint2 As Data.OracleClient.OracleDataReader = CmdPoint2.ExecuteReader()
                                If rdPoint2.HasRows Then
                                    While rdPoint2.Read()
                                        If rdPoint2("BillType2Score") Is DBNull.Value Then
                                            PersonPoint = PersonPoint + 0
                                            BillCnt = BillCnt + 0
                                        Else
                                            If rdPoint2("BillMemID1") IsNot DBNull.Value And rdPoint2("BillMemID2") Is DBNull.Value And rdPoint2("BillMemID3") Is DBNull.Value And rdPoint2("BillMemID4") Is DBNull.Value Then
                                                PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score"))
                                                BillCnt = BillCnt + 1
                                            ElseIf rdPoint2("BillMemID1") IsNot DBNull.Value And rdPoint2("BillMemID2") IsNot DBNull.Value And rdPoint2("BillMemID3") Is DBNull.Value And rdPoint2("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdPoint2("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") / 2)
                                                    BillCnt = BillCnt + 0.5
                                                End If
                                                If Trim(rdPoint2("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") / 2)
                                                    BillCnt = BillCnt + 0.5
                                                End If
                                            ElseIf rdPoint2("BillMemID1") IsNot DBNull.Value And rdPoint2("BillMemID2") IsNot DBNull.Value And rdPoint2("BillMemID3") IsNot DBNull.Value And rdPoint2("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdPoint2("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") * 0.34)
                                                    BillCnt = BillCnt + 0.34
                                                End If
                                                If Trim(rdPoint2("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") * 0.33)
                                                    BillCnt = BillCnt + 0.33
                                                End If
                                                If Trim(rdPoint2("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") * 0.33)
                                                    BillCnt = BillCnt + 0.33
                                                End If
                                            ElseIf rdPoint2("BillMemID1") IsNot DBNull.Value And rdPoint2("BillMemID2") IsNot DBNull.Value And rdPoint2("BillMemID3") IsNot DBNull.Value And rdPoint2("BillMemID4") IsNot DBNull.Value Then
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint2("BillMemID1")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint2("BillMemID2")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint2("BillMemID3")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint2("BillMemID4")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                            End If
                                        End If
                                    End While
                                End If
                                rdPoint2.Close()
                                
                                '拖吊點數
                                Dim strPoint6 As String = "select b.Other1,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBase a" & strPointT6Plus2
                                strPoint6 = strPoint6 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                                strPoint6 = strPoint6 & " where (a.BillMemID1=" & Trim(rdPer("MemberID")) & " or a.BillMemID2=" & Trim(rdPer("MemberID")) & " or a.BillMemID3=" & Trim(rdPer("MemberID")) & " or a.BillMemID4=" & Trim(rdPer("MemberID")) & ")"
                                strPoint6 = strPoint6 & " and a.RuleVer=b.LawVersion"
                                strPoint6 = strPoint6 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                                strPoint6 = strPoint6 & " and a.CarAddID=8 and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                                strPoint6 = strPoint6 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPoint6 = strPoint6 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')" & strPointT6Plus
                                Dim CmdPoint6 As New Data.OracleClient.OracleCommand(strPoint6, conn)
                                Dim rdPoint6 As Data.OracleClient.OracleDataReader = CmdPoint6.ExecuteReader()
                                If rdPoint6.HasRows Then
                                    While rdPoint6.Read()
                                        If rdPoint6("Other1") Is DBNull.Value Then
                                            PersonPoint = PersonPoint + 0
                                            BillCnt = BillCnt + 0
                                        Else
                                            If rdPoint6("BillMemID1") IsNot DBNull.Value And rdPoint6("BillMemID2") Is DBNull.Value And rdPoint6("BillMemID3") Is DBNull.Value And rdPoint6("BillMemID4") Is DBNull.Value Then
                                                PersonPoint = PersonPoint + CDec(rdPoint6("Other1"))
                                                BillCnt = BillCnt + 1
                                            ElseIf rdPoint6("BillMemID1") IsNot DBNull.Value And rdPoint6("BillMemID2") IsNot DBNull.Value And rdPoint6("BillMemID3") Is DBNull.Value And rdPoint6("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdPoint6("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint6("Other1") / 2)
                                                    BillCnt = BillCnt + 0.5
                                                End If
                                                If Trim(rdPoint6("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint6("Other1") / 2)
                                                    BillCnt = BillCnt + 0.5
                                                End If
                                            ElseIf rdPoint6("BillMemID1") IsNot DBNull.Value And rdPoint6("BillMemID2") IsNot DBNull.Value And rdPoint6("BillMemID3") IsNot DBNull.Value And rdPoint6("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdPoint6("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint6("Other1") * 0.34)
                                                    BillCnt = BillCnt + 0.34
                                                End If
                                                If Trim(rdPoint6("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint6("Other1") * 0.33)
                                                    BillCnt = BillCnt + 0.33
                                                End If
                                                If Trim(rdPoint6("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint6("Other1") * 0.33)
                                                    BillCnt = BillCnt + 0.33
                                                End If
                                            ElseIf rdPoint6("BillMemID1") IsNot DBNull.Value And rdPoint6("BillMemID2") IsNot DBNull.Value And rdPoint6("BillMemID3") IsNot DBNull.Value And rdPoint6("BillMemID4") IsNot DBNull.Value Then
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint6("BillMemID1")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint6("Other1") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint6("BillMemID2")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint6("Other1") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint6("BillMemID3")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint6("Other1") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint6("BillMemID4")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint6("Other1") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                            End If
                                        End If
                                    End While
                                End If
                                rdPoint6.Close()
                            Else
                                '攔停點數
                                Dim strPoint1 As String = "select b.BillType1Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBaseViewReward a"
                                strPoint1 = strPoint1 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
                                strPoint1 = strPoint1 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                                strPoint1 = strPoint1 & " where (a.BillMemID1=" & Trim(rdPer("MemberID")) & " or a.BillMemID2=" & Trim(rdPer("MemberID")) & " or a.BillMemID3=" & Trim(rdPer("MemberID")) & " or a.BillMemID4=" & Trim(rdPer("MemberID")) & ")"
                                strPoint1 = strPoint1 & " and a.RuleVer=b.LawVersion"
                                strPoint1 = strPoint1 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                                strPoint1 = strPoint1 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                                strPoint1 = strPoint1 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPoint1 = strPoint1 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPoint1 As New Data.OracleClient.OracleCommand(strPoint1, conn)
                                Dim rdPoint1 As Data.OracleClient.OracleDataReader = CmdPoint1.ExecuteReader()
                                If rdPoint1.HasRows Then
                                    While rdPoint1.Read()
                                        If rdPoint1("BillType1Score") Is DBNull.Value Then
                                            PersonPoint = PersonPoint + 0
                                            BillCnt = BillCnt + 0
                                        Else
                                            If rdPoint1("BillMemID1") IsNot DBNull.Value And rdPoint1("BillMemID2") Is DBNull.Value And rdPoint1("BillMemID3") Is DBNull.Value And rdPoint1("BillMemID4") Is DBNull.Value Then
                                                PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score"))
                                                BillCnt = BillCnt + 1
                                            ElseIf rdPoint1("BillMemID1") IsNot DBNull.Value And rdPoint1("BillMemID2") IsNot DBNull.Value And rdPoint1("BillMemID3") Is DBNull.Value And rdPoint1("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdPoint1("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") / 2)
                                                    BillCnt = BillCnt + 0.5
                                                End If
                                                If Trim(rdPoint1("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") / 2)
                                                    BillCnt = BillCnt + 0.5
                                                End If
                                            ElseIf rdPoint1("BillMemID1") IsNot DBNull.Value And rdPoint1("BillMemID2") IsNot DBNull.Value And rdPoint1("BillMemID3") IsNot DBNull.Value And rdPoint1("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdPoint1("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") * 0.34)
                                                    BillCnt = BillCnt + 0.34
                                                End If
                                                If Trim(rdPoint1("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") * 0.33)
                                                    BillCnt = BillCnt + 0.33
                                                End If
                                                If Trim(rdPoint1("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") * 0.33)
                                                    BillCnt = BillCnt + 0.33
                                                End If
                                            ElseIf rdPoint1("BillMemID1") IsNot DBNull.Value And rdPoint1("BillMemID2") IsNot DBNull.Value And rdPoint1("BillMemID3") IsNot DBNull.Value And rdPoint1("BillMemID4") IsNot DBNull.Value Then
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint1("BillMemID1")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint1("BillMemID2")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint1("BillMemID3")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint1("BillMemID4")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint1("BillType1Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                            End If
                                        End If
                                    End While
                                End If
                                rdPoint1.Close()
                    
                                '逕舉點數
                                Dim strPoint2 As String = "select b.BillType2Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBaseViewReward a"
                                strPoint2 = strPoint2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
                                strPoint2 = strPoint2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                                strPoint2 = strPoint2 & " where (a.BillMemID1=" & Trim(rdPer("MemberID")) & " or a.BillMemID2=" & Trim(rdPer("MemberID")) & " or a.BillMemID3=" & Trim(rdPer("MemberID")) & " or a.BillMemID4=" & Trim(rdPer("MemberID")) & ")"
                                strPoint2 = strPoint2 & " and a.RuleVer=b.LawVersion"
                                strPoint2 = strPoint2 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                                strPoint2 = strPoint2 & " and a.BillTypeID='2' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                                strPoint2 = strPoint2 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPoint2 = strPoint2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPoint2 As New Data.OracleClient.OracleCommand(strPoint2, conn)
                                Dim rdPoint2 As Data.OracleClient.OracleDataReader = CmdPoint2.ExecuteReader()
                                If rdPoint2.HasRows Then
                                    While rdPoint2.Read()
                                        If rdPoint2("BillType2Score") Is DBNull.Value Then
                                            PersonPoint = PersonPoint + 0
                                            BillCnt = BillCnt + 0
                                        Else
                                            If rdPoint2("BillMemID1") IsNot DBNull.Value And rdPoint2("BillMemID2") Is DBNull.Value And rdPoint2("BillMemID3") Is DBNull.Value And rdPoint2("BillMemID4") Is DBNull.Value Then
                                                PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score"))
                                                BillCnt = BillCnt + 1
                                            ElseIf rdPoint2("BillMemID1") IsNot DBNull.Value And rdPoint2("BillMemID2") IsNot DBNull.Value And rdPoint2("BillMemID3") Is DBNull.Value And rdPoint2("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdPoint2("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") / 2)
                                                    BillCnt = BillCnt + 0.5
                                                End If
                                                If Trim(rdPoint2("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") / 2)
                                                    BillCnt = BillCnt + 0.5
                                                End If
                                            ElseIf rdPoint2("BillMemID1") IsNot DBNull.Value And rdPoint2("BillMemID2") IsNot DBNull.Value And rdPoint2("BillMemID3") IsNot DBNull.Value And rdPoint2("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdPoint2("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") * 0.34)
                                                    BillCnt = BillCnt + 0.34
                                                End If
                                                If Trim(rdPoint2("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") * 0.33)
                                                    BillCnt = BillCnt + 0.33
                                                End If
                                                If Trim(rdPoint2("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") * 0.33)
                                                    BillCnt = BillCnt + 0.33
                                                End If
                                            ElseIf rdPoint2("BillMemID1") IsNot DBNull.Value And rdPoint2("BillMemID2") IsNot DBNull.Value And rdPoint2("BillMemID3") IsNot DBNull.Value And rdPoint2("BillMemID4") IsNot DBNull.Value Then
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint2("BillMemID1")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint2("BillMemID2")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint2("BillMemID3")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdPoint2("BillMemID4")) Then
                                                    PersonPoint = PersonPoint + CDec(rdPoint2("BillType2Score") / 4)
                                                    BillCnt = BillCnt + 0.25
                                                End If
                                            End If
                                        End If
                                    End While
                                End If
                                rdPoint2.Close()
                            End If
                            'A1點數
                            Dim strPoint3 As String = "select b.BillType1Score,b.A1Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBaseViewReward a"
                            strPoint3 = strPoint3 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
                            strPoint3 = strPoint3 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                            strPoint3 = strPoint3 & " where (a.BillMemID1=" & Trim(rdPer("MemberID")) & " or a.BillMemID2=" & Trim(rdPer("MemberID")) & " or a.BillMemID3=" & Trim(rdPer("MemberID")) & " or a.BillMemID4=" & Trim(rdPer("MemberID")) & ")"
                            strPoint3 = strPoint3 & " and a.RuleVer=b.LawVersion"
                            strPoint3 = strPoint3 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                            strPoint3 = strPoint3 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType='1')"
                            strPoint3 = strPoint3 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                            strPoint3 = strPoint3 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                            Dim CmdPoint3 As New Data.OracleClient.OracleCommand(strPoint3, conn)
                            Dim rdPoint3 As Data.OracleClient.OracleDataReader = CmdPoint3.ExecuteReader()
                            If rdPoint3.HasRows Then
                                While rdPoint3.Read()
                                    If rdPoint3("BillType1Score") > rdPoint3("A1Score") Then
                                        If rdPoint3("BillType1Score") Is DBNull.Value Then
                                            A1ScoreTmp = 0
                                        Else
                                            A1ScoreTmp = CDec(rdPoint3("BillType1Score"))
                                        End If
                                    Else
                                        If rdPoint3("A1Score") Is DBNull.Value Then
                                            A1ScoreTmp = 0
                                        Else
                                            A1ScoreTmp = CDec(rdPoint3("A1Score"))
                                        End If
                                    End If
                                    
                                    If A1ScoreTmp = 0 Then
                                        PersonPoint = PersonPoint + 0
                                        BillCnt = BillCnt + 0
                                    Else
                                        If rdPoint3("BillMemID1") IsNot DBNull.Value And rdPoint3("BillMemID2") Is DBNull.Value And rdPoint3("BillMemID3") Is DBNull.Value And rdPoint3("BillMemID4") Is DBNull.Value Then
                                            PersonPoint = PersonPoint + CDec(A1ScoreTmp)
                                            BillCnt = BillCnt + 1
                                        ElseIf rdPoint3("BillMemID1") IsNot DBNull.Value And rdPoint3("BillMemID2") IsNot DBNull.Value And rdPoint3("BillMemID3") Is DBNull.Value And rdPoint3("BillMemID4") Is DBNull.Value Then
                                            If Trim(rdPoint3("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A1ScoreTmp / 2)
                                                BillCnt = BillCnt + 0.5
                                            End If
                                            If Trim(rdPoint3("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A1ScoreTmp / 2)
                                                BillCnt = BillCnt + 0.5
                                            End If
                                        ElseIf rdPoint3("BillMemID1") IsNot DBNull.Value And rdPoint3("BillMemID2") IsNot DBNull.Value And rdPoint3("BillMemID3") IsNot DBNull.Value And rdPoint3("BillMemID4") Is DBNull.Value Then
                                            If Trim(rdPoint3("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A1ScoreTmp * 0.34)
                                                BillCnt = BillCnt + 0.34
                                            End If
                                            If Trim(rdPoint3("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A1ScoreTmp * 0.33)
                                                BillCnt = BillCnt + 0.33
                                            End If
                                            If Trim(rdPoint3("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A1ScoreTmp * 0.33)
                                                BillCnt = BillCnt + 0.33
                                            End If
                                        ElseIf rdPoint3("BillMemID1") IsNot DBNull.Value And rdPoint3("BillMemID2") IsNot DBNull.Value And rdPoint3("BillMemID3") IsNot DBNull.Value And rdPoint3("BillMemID4") IsNot DBNull.Value Then
                                            If Trim(rdPer("MemberID")) = Trim(rdPoint3("BillMemID1")) Then
                                                PersonPoint = PersonPoint + CDec(A1ScoreTmp / 4)
                                                BillCnt = BillCnt + 0.25
                                            End If
                                            If Trim(rdPer("MemberID")) = Trim(rdPoint3("BillMemID2")) Then
                                                PersonPoint = PersonPoint + CDec(A1ScoreTmp / 4)
                                                BillCnt = BillCnt + 0.25
                                            End If
                                            If Trim(rdPer("MemberID")) = Trim(rdPoint3("BillMemID3")) Then
                                                PersonPoint = PersonPoint + CDec(A1ScoreTmp / 4)
                                                BillCnt = BillCnt + 0.25
                                            End If
                                            If Trim(rdPer("MemberID")) = Trim(rdPoint3("BillMemID4")) Then
                                                PersonPoint = PersonPoint + CDec(A1ScoreTmp / 4)
                                                BillCnt = BillCnt + 0.25
                                            End If
                                        End If
                                    End If
                                End While
                            End If
                            rdPoint3.Close()
                    
                            'A2點數
                            Dim strPoint4 As String = "select b.BillType1Score,b.A2Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBaseViewReward a"
                            strPoint4 = strPoint4 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
                            strPoint4 = strPoint4 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                            strPoint4 = strPoint4 & " where (a.BillMemID1=" & Trim(rdPer("MemberID")) & " or a.BillMemID2=" & Trim(rdPer("MemberID")) & " or a.BillMemID3=" & Trim(rdPer("MemberID")) & " or a.BillMemID4=" & Trim(rdPer("MemberID")) & ")"
                            strPoint4 = strPoint4 & " and a.RuleVer=b.LawVersion"
                            strPoint4 = strPoint4 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                            strPoint4 = strPoint4 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType='2')"
                            strPoint4 = strPoint4 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                            strPoint4 = strPoint4 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                            Dim CmdPoint4 As New Data.OracleClient.OracleCommand(strPoint4, conn)
                            Dim rdPoint4 As Data.OracleClient.OracleDataReader = CmdPoint4.ExecuteReader()
                            If rdPoint4.HasRows Then
                                While rdPoint4.Read()
                                    If rdPoint4("BillType1Score") > rdPoint4("A2Score") Then
                                        If rdPoint4("BillType1Score") Is DBNull.Value Then
                                            A2ScoreTmp = 0
                                        Else
                                            A2ScoreTmp = CDec(rdPoint4("BillType1Score"))
                                        End If
                                    Else
                                        If rdPoint4("A2Score") Is DBNull.Value Then
                                            A2ScoreTmp = 0
                                        Else
                                            A2ScoreTmp = CDec(rdPoint4("A2Score"))
                                        End If
                                    End If
                                    
                                    If A2ScoreTmp = 0 Then
                                        PersonPoint = PersonPoint + 0
                                        BillCnt = BillCnt + 0
                                    Else
                                        If rdPoint4("BillMemID1") IsNot DBNull.Value And rdPoint4("BillMemID2") Is DBNull.Value And rdPoint4("BillMemID3") Is DBNull.Value And rdPoint4("BillMemID4") Is DBNull.Value Then
                                            PersonPoint = PersonPoint + CDec(A2ScoreTmp)
                                            BillCnt = BillCnt + 1
                                        ElseIf rdPoint4("BillMemID1") IsNot DBNull.Value And rdPoint4("BillMemID2") IsNot DBNull.Value And rdPoint4("BillMemID3") Is DBNull.Value And rdPoint4("BillMemID4") Is DBNull.Value Then
                                            If Trim(rdPoint4("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A2ScoreTmp / 2)
                                                BillCnt = BillCnt + 0.5
                                            End If
                                            If Trim(rdPoint4("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A2ScoreTmp / 2)
                                                BillCnt = BillCnt + 0.5
                                            End If
                                        ElseIf rdPoint4("BillMemID1") IsNot DBNull.Value And rdPoint4("BillMemID2") IsNot DBNull.Value And rdPoint4("BillMemID3") IsNot DBNull.Value And rdPoint4("BillMemID4") Is DBNull.Value Then
                                            If Trim(rdPoint4("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A2ScoreTmp * 0.34)
                                                BillCnt = BillCnt + 0.34
                                            End If
                                            If Trim(rdPoint4("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A2ScoreTmp * 0.33)
                                                BillCnt = BillCnt + 0.33
                                            End If
                                            If Trim(rdPoint4("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A2ScoreTmp * 0.33)
                                                BillCnt = BillCnt + 0.33
                                            End If
                                        ElseIf rdPoint4("BillMemID1") IsNot DBNull.Value And rdPoint4("BillMemID2") IsNot DBNull.Value And rdPoint4("BillMemID3") IsNot DBNull.Value And rdPoint4("BillMemID4") IsNot DBNull.Value Then
                                            If Trim(rdPer("MemberID")) = Trim(rdPoint4("BillMemID1")) Then
                                                PersonPoint = PersonPoint + CDec(A2ScoreTmp / 4)
                                                BillCnt = BillCnt + 0.25
                                            End If
                                            If Trim(rdPer("MemberID")) = Trim(rdPoint4("BillMemID2")) Then
                                                PersonPoint = PersonPoint + CDec(A2ScoreTmp / 4)
                                                BillCnt = BillCnt + 0.25
                                            End If
                                            If Trim(rdPer("MemberID")) = Trim(rdPoint4("BillMemID3")) Then
                                                PersonPoint = PersonPoint + CDec(A2ScoreTmp / 4)
                                                BillCnt = BillCnt + 0.25
                                            End If
                                            If Trim(rdPer("MemberID")) = Trim(rdPoint4("BillMemID4")) Then
                                                PersonPoint = PersonPoint + CDec(A2ScoreTmp / 4)
                                                BillCnt = BillCnt + 0.25
                                            End If
                                        End If
                                    End If
                                End While
                            End If
                            rdPoint4.Close()
                    
                            'A3點數
                            Dim strPoint5 As String = "select b.BillType1Score,b.A3Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBaseViewReward a"
                            strPoint5 = strPoint5 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
                            strPoint5 = strPoint5 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                            strPoint5 = strPoint5 & " where (a.BillMemID1=" & Trim(rdPer("MemberID")) & " or a.BillMemID2=" & Trim(rdPer("MemberID")) & " or a.BillMemID3=" & Trim(rdPer("MemberID")) & " or a.BillMemID4=" & Trim(rdPer("MemberID")) & ")"
                            strPoint5 = strPoint5 & " and a.RuleVer=b.LawVersion"
                            strPoint5 = strPoint5 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                            strPoint5 = strPoint5 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType='3')"
                            strPoint5 = strPoint5 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                            strPoint5 = strPoint5 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                            Dim CmdPoint5 As New Data.OracleClient.OracleCommand(strPoint5, conn)
                            Dim rdPoint5 As Data.OracleClient.OracleDataReader = CmdPoint5.ExecuteReader()
                            If rdPoint5.HasRows Then
                                While rdPoint5.Read()
                                    If rdPoint5("BillType1Score") > rdPoint5("A3Score") Then
                                        If rdPoint5("BillType1Score") Is DBNull.Value Then
                                            A3ScoreTmp = 0
                                        Else
                                            A3ScoreTmp = CDec(rdPoint5("BillType1Score"))
                                        End If
                                    Else
                                        If rdPoint5("A3Score") Is DBNull.Value Then
                                            A3ScoreTmp = 0
                                        Else
                                            A3ScoreTmp = CDec(rdPoint5("A3Score"))
                                        End If
                                    End If
                                    
                                    If A3ScoreTmp = 0 Then
                                        PersonPoint = PersonPoint + 0
                                        BillCnt = BillCnt + 0
                                    Else
                                        If rdPoint5("BillMemID1") IsNot DBNull.Value And rdPoint5("BillMemID2") Is DBNull.Value And rdPoint5("BillMemID3") Is DBNull.Value And rdPoint5("BillMemID4") Is DBNull.Value Then
                                            PersonPoint = PersonPoint + CDec(A3ScoreTmp)
                                            BillCnt = BillCnt + 1
                                        ElseIf rdPoint5("BillMemID1") IsNot DBNull.Value And rdPoint5("BillMemID2") IsNot DBNull.Value And rdPoint5("BillMemID3") Is DBNull.Value And rdPoint5("BillMemID4") Is DBNull.Value Then
                                            If Trim(rdPoint5("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A3ScoreTmp / 2)
                                                BillCnt = BillCnt + 0.5
                                            End If
                                            If Trim(rdPoint5("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A3ScoreTmp / 2)
                                                BillCnt = BillCnt + 0.5
                                            End If
                                        ElseIf rdPoint5("BillMemID1") IsNot DBNull.Value And rdPoint5("BillMemID2") IsNot DBNull.Value And rdPoint5("BillMemID3") IsNot DBNull.Value And rdPoint5("BillMemID4") Is DBNull.Value Then
                                            If Trim(rdPoint5("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A3ScoreTmp * 0.34)
                                                BillCnt = BillCnt + 0.34
                                            End If
                                            If Trim(rdPoint5("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A3ScoreTmp * 0.33)
                                                BillCnt = BillCnt + 0.33
                                            End If
                                            If Trim(rdPoint5("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                PersonPoint = PersonPoint + CDec(A3ScoreTmp * 0.33)
                                                BillCnt = BillCnt + 0.33
                                            End If
                                        ElseIf rdPoint5("BillMemID1") IsNot DBNull.Value And rdPoint5("BillMemID2") IsNot DBNull.Value And rdPoint5("BillMemID3") IsNot DBNull.Value And rdPoint5("BillMemID4") IsNot DBNull.Value Then
                                            If Trim(rdPer("MemberID")) = Trim(rdPoint5("BillMemID1")) Then
                                                PersonPoint = PersonPoint + CDec(A3ScoreTmp / 4)
                                                BillCnt = BillCnt + 0.25
                                            End If
                                            If Trim(rdPer("MemberID")) = Trim(rdPoint5("BillMemID2")) Then
                                                PersonPoint = PersonPoint + CDec(A3ScoreTmp / 4)
                                                BillCnt = BillCnt + 0.25
                                            End If
                                            If Trim(rdPer("MemberID")) = Trim(rdPoint5("BillMemID3")) Then
                                                PersonPoint = PersonPoint + CDec(A3ScoreTmp / 4)
                                                BillCnt = BillCnt + 0.25
                                            End If
                                            If Trim(rdPer("MemberID")) = Trim(rdPoint5("BillMemID4")) Then
                                                PersonPoint = PersonPoint + CDec(A3ScoreTmp / 4)
                                                BillCnt = BillCnt + 0.25
                                            End If
                                        End If
                                    End If
                                End While
                            End If
                            rdPoint5.Close()
                                
                            If rdPer("Money") Is DBNull.Value Then
                                MemMoney = Decimal.Truncate(PointMoney * PersonPoint)
                            Else
                                If rdPer("Money") = "0" Then
                                    MemMoney = Decimal.Truncate(PointMoney * PersonPoint)
                                Else
                                    MemPay = Decimal.Truncate(CDec(rdPer("Money")) * getPayPercent)
                                    If Decimal.Truncate(PointMoney * PersonPoint) > MemPay Then
                                        MemMoney = MemPay
                                    Else
                                        MemMoney = Decimal.Truncate(PointMoney * PersonPoint)
                                    End If
                                End If
                            End If
                            UnitPoint = UnitPoint + PersonPoint
                            UnitMoney = UnitMoney + MemMoney
                        End While
                    End If
                    rdPer.Close()
                    PointTotal = PointTotal + UnitPoint
                    MoneyTotal = MoneyTotal + UnitMoney
                    BillCntTotal = BillCntTotal + BillCnt
                    Response.Write("<td align=""right"">" & Format(BillCnt, "##,##0.00") & "</td>")
                    Response.Write("<td align=""right"">" & Format(UnitPoint, "##,##0.00") & "</td>")
                    Response.Write("<td align=""right"">" & Format(UnitMoney, "##,##0") & "</td>")
                    Response.Write("</tr>")
                End While
            End If
            rdUnit.Close()
            conn.Close()
        %>
        <tr>
        <td>總計</td>
        <td align="right"><%=Format(BillCntTotal, "##,##0.00")%></td>
        <td align="right"><%=Format(PointTotal, "##,##0.00")%></td>
        <td align="right"><%=Format(MoneyTotal, "##,##0")%></td>
        </tr>
        <%--<tr>
        <td><span class="style1"><strong>製表</strong></span></td>
        <td><span class="style1"><strong>組長</strong></span></td>
        <td><span class="style1"><strong>副隊長</strong></span></td>
        <td><span class="style1"><strong>隊長</strong></span></td>
        </tr>--%>
    </table>
    </form>
</body>
</html>
