<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%  
    LoginCheck()
%>
<%
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

    'Response.Clear()
    'Response.Buffer = True
    'Response.Charset = "utf-8"
    ''//下面这行很重要， attachment 参数表示作为附件下载，您可以改成 online在线打开 
    ''//filename=FileFlow.xls 指定输出文件的名称，注意其扩展名和指定文件类型相符，可以为：.doc 　　 .xls 　　 .txt 　　.htm　　 
    'Response.AppendHeader("Content-Disposition", "attachment;filename=FileFlow.xls")
    'Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8")
    ''//Response.ContentType指定文件类型 可以为application/ms-excel 　　 application/ms-word 　　 application/ms-txt 　　 application/ms-html 　　 或其他浏览器可直接支持文档　 
    'Response.ContentType = "application/ms-excel"
    'Me.EnableViewState = False

    Server.ScriptTimeout = 86400
    Response.Flush()
 %>
<script runat="server">
    Public PersonPoint, PointTotal, MoneyTotal As Decimal
    
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

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>處理道路交通安全人員獎勵金發放一覽表</title>
    <style type="text/css">
<!--
    .style1 {
        font-size: 18px;
        font-family: "新細明體"
    }
    .style2 {
        font-size: 15px;
        font-family: "新細明體"
    }
    .style3 {
        font-size: 17px;
        font-family: "新細明體";
        font-weight: bold;
    }
    .style4 {
        font-size: 18px;
        font-family: "新細明體";
        font-weight: bold;
    }


-->
</style>
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
            
    Dim sys_City As String
    '要用填單或建檔日統計
    Dim theDateType As String = Trim(Request("DateType"))
    '統計期間
    Dim TheAnaDate1, TheAnaDate2 As String
    TheAnaDate1 = (Trim(Request("Year1")) + 1911) & "/" & Trim(Request("Month1")) & "/1"
    TheAnaDate2 = DateAdd("d", -1, DateAdd("m", 1, (Trim(Request("Year1")) + 1911) & "/" & Trim(Request("Month1")) & "/1"))
    
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
            
    '表頭單位
    Dim getTUnit As String = ""
    Dim strTUnit As String = "select value from Apconfigure where ID=35"
    Dim CmdTUnit As New Data.OracleClient.OracleCommand(strTUnit, conn)
    Dim rdTUnit As Data.OracleClient.OracleDataReader = CmdTUnit.ExecuteReader()
    If rdTUnit.HasRows Then
        rdTUnit.Read()
                        
        getTUnit = Trim(rdTUnit("value"))
    End If
    rdTUnit.Close()
    
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
    Dim FlagPointTotal, strPointT6Plus2, strPointT6Plus As String
    FlagPointTotal = Trim(Request("AnalyzeType"))
    
               
    ''攔停點數
    'Dim strPointT1 As String = "select sum(b.BillType1Score) as cnt from BillBaseViewReward a"
    'strPointT1 = strPointT1 & ",(select distinct LawVersion,LawItem,BillType1Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
    'strPointT1 = strPointT1 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
    'strPointT1 = strPointT1 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion "
    'strPointT1 = strPointT1 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
    'strPointT1 = strPointT1 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
    'strPointT1 = strPointT1 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
    'strPointT1 = strPointT1 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
    'Dim CmdPointT1 As New Data.OracleClient.OracleCommand(strPointT1, conn)
    'Dim rdPointT1 As Data.OracleClient.OracleDataReader = CmdPointT1.ExecuteReader()
    'If rdPointT1.HasRows Then
    '    rdPointT1.Read()
    '    If rdPointT1("cnt") Is DBNull.Value Then
    '        getPointTotal = 0
    '    Else
    '        getPointTotal = CDec(rdPointT1("cnt"))
    '    End If
    'End If
    'rdPointT1.Close()

    ''逕舉點數
    'Dim strPointT2 As String = "select sum(b.BillType2Score) as cnt from BillBase a"
    'strPointT2 = strPointT2 & ",(select distinct LawVersion,LawItem,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
    'strPointT2 = strPointT2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
    'strPointT2 = strPointT2 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion "
    'strPointT2 = strPointT2 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
    'strPointT2 = strPointT2 & " and a.BillTypeID='2' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
    'strPointT2 = strPointT2 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
    'strPointT2 = strPointT2 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
    'Dim CmdPointT2 As New Data.OracleClient.OracleCommand(strPointT2, conn)
    'Dim rdPointT2 As Data.OracleClient.OracleDataReader = CmdPointT2.ExecuteReader()
    'If rdPointT2.HasRows Then
    '    rdPointT2.Read()
    '    If rdPointT2("cnt") Is DBNull.Value Then
    '        getPointTotal = getPointTotal + 0
    '    Else
    '        getPointTotal = getPointTotal + CDec(rdPointT2("cnt"))
    '    End If
    'End If
    'rdPointT2.Close()
            
    ''A1點數
    'Dim strPointT3 As String = "select b.A1Score,b.BillType1Score from BillBase a"
    'strPointT3 = strPointT3 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
    'strPointT3 = strPointT3 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
    'strPointT3 = strPointT3 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion "
    'strPointT3 = strPointT3 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
    'strPointT3 = strPointT3 & " and a.RecordStateID=0 and (a.TrafficAccidentType='1') and a.BillTypeID='1'"
    'strPointT3 = strPointT3 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
    'strPointT3 = strPointT3 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
    'Dim CmdPointT3 As New Data.OracleClient.OracleCommand(strPointT3, conn)
    'Dim rdPointT3 As Data.OracleClient.OracleDataReader = CmdPointT3.ExecuteReader()
    'If rdPointT3.HasRows Then
    '    While rdPointT3.Read()
    '        If CDec(rdPointT3("BillType1Score")) > CDec(rdPointT3("A1Score")) Then
    '            If rdPointT3("BillType1Score") Is DBNull.Value Then
    '                getPointTotal = getPointTotal + 0
    '            Else
    '                getPointTotal = getPointTotal + CDec(rdPointT3("BillType1Score"))
    '            End If
    '        Else
    '            If rdPointT3("A1Score") Is DBNull.Value Then
    '                getPointTotal = getPointTotal + 0
    '            Else
    '                getPointTotal = getPointTotal + CDec(rdPointT3("A1Score"))
    '            End If
    '        End If
    '    End While
    'End If
    'rdPointT3.Close()
            
    ''A2點數
    'Dim strPointT4 As String = "select b.A2Score,b.BillType1Score from BillBase a"
    'strPointT4 = strPointT4 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
    'strPointT4 = strPointT4 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
    'strPointT4 = strPointT4 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion "
    'strPointT4 = strPointT4 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
    'strPointT4 = strPointT4 & " and a.RecordStateID=0 and (a.TrafficAccidentType='2') and a.BillTypeID='1'"
    'strPointT4 = strPointT4 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
    'strPointT4 = strPointT4 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
    'Dim CmdPointT4 As New Data.OracleClient.OracleCommand(strPointT4, conn)
    'Dim rdPointT4 As Data.OracleClient.OracleDataReader = CmdPointT4.ExecuteReader()
    'If rdPointT4.HasRows Then
    '    While rdPointT4.Read()
    '        If CDec(rdPointT4("BillType1Score")) > CDec(rdPointT4("A2Score")) Then
    '            If rdPointT4("BillType1Score") Is DBNull.Value Then
    '                getPointTotal = getPointTotal + 0
    '            Else
    '                getPointTotal = getPointTotal + CDec(rdPointT4("BillType1Score"))
    '            End If
    '        Else
    '            If rdPointT4("A2Score") Is DBNull.Value Then
    '                getPointTotal = getPointTotal + 0
    '            Else
    '                getPointTotal = getPointTotal + CDec(rdPointT4("A2Score"))
    '            End If
    '        End If
    '    End While

    'End If
    'rdPointT4.Close()
                    
    ''A3點數
    'Dim strPointT5 As String = "select b.A3Score,b.BillType1Score from BillBase a"
    'strPointT5 = strPointT5 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
    'strPointT5 = strPointT5 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b,UnitInfo c"
    'strPointT5 = strPointT5 & " where a.BillUnitID=c.UnitID and a.RuleVer=b.LawVersion "
    'strPointT5 = strPointT5 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
    'strPointT5 = strPointT5 & " and a.RecordStateID=0 and (a.TrafficAccidentType='3') and a.BillTypeID='1'"
    'strPointT5 = strPointT5 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
    'strPointT5 = strPointT5 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
    'Dim CmdPointT5 As New Data.OracleClient.OracleCommand(strPointT5, conn)
    'Dim rdPointT5 As Data.OracleClient.OracleDataReader = CmdPointT5.ExecuteReader()
    'If rdPointT5.HasRows Then
    '    While rdPointT5.Read()
    '        If CDec(rdPointT5("BillType1Score")) > CDec(rdPointT5("A3Score")) Then
    '            If rdPointT5("BillType1Score") Is DBNull.Value Then
    '                getPointTotal = getPointTotal + 0
    '            Else
    '                getPointTotal = getPointTotal + CDec(rdPointT5("BillType1Score"))
    '            End If
    '        Else
    '            If rdPointT5("A3Score") Is DBNull.Value Then
    '                getPointTotal = getPointTotal + 0
    '            Else
    '                getPointTotal = getPointTotal + CDec(rdPointT5("A3Score"))
    '            End If
    '        End If
                        
    '    End While
    'End If
    'rdPointT5.Close()
    
    ''拖吊
    'If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "台中市" Then
    '    Dim strPointT1a As String = "select count(*) cnt from BillBase a"
    '    strPointT1a = strPointT1a & " ,UnitInfo c"
    '    strPointT1a = strPointT1a & " where a.BillUnitID=c.UnitID "
    '    strPointT1a = strPointT1a & " and a.RecordStateID=0"
    '    strPointT1a = strPointT1a & " and a.ProjectID='A5'"
    '    strPointT1a = strPointT1a & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
    '    strPointT1a = strPointT1a & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
    '    Dim CmdPointT1a As New Data.OracleClient.OracleCommand(strPointT1a, conn)
    '    Dim rdPointT1a As Data.OracleClient.OracleDataReader = CmdPointT1a.ExecuteReader()
    '    If rdPointT1a.HasRows Then
    '        rdPointT1a.Read()
    '        If rdPointT1a("cnt") Is DBNull.Value Then
    '            getPointTotal = getPointTotal
    '        Else
                        
    '            getPointTotal = getPointTotal + (CDec(rdPointT1a("cnt")) * 20)
    '        End If
    '    End If
    '    rdPointT1a.Close()
                
    '    Dim strPointT1b As String = "select count(*) cnt from BillBase a"
    '    strPointT1b = strPointT1b & " ,UnitInfo c"
    '    strPointT1b = strPointT1b & " where a.BillUnitID=c.UnitID "
    '    strPointT1b = strPointT1b & " and a.RecordStateID=0"
    '    strPointT1b = strPointT1b & " and a.ProjectID='A6'"
    '    strPointT1b = strPointT1b & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
    '    strPointT1b = strPointT1b & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
    '    Dim CmdPointT1b As New Data.OracleClient.OracleCommand(strPointT1b, conn)
    '    Dim rdPointT1b As Data.OracleClient.OracleDataReader = CmdPointT1b.ExecuteReader()
    '    If rdPointT1b.HasRows Then
    '        rdPointT1b.Read()
    '        If rdPointT1b("cnt") Is DBNull.Value Then
    '            getPointTotal = getPointTotal
    '        Else
                        
    '            getPointTotal = getPointTotal + (CDec(rdPointT1b("cnt")) * 50)
    '        End If
    '    End If
    '    rdPointT1b.Close()
    'End If
    
    'Dim PointMoney As Decimal
    'If getPointTotal = 0 Then
    '    PointMoney = 0
    'Else
    '    PointMoney = getMoneyTotal / getPointTotal
    'End If
    
    '產生報表----------------------------------------------------------------    
    Response.Write("<table width=""1035"" border=""1"" cellpadding=""3"" cellspacing=""0"">")
    Response.Write("<tr>")
    Response.Write("<td style=""height:30px"" align=""center"" colspan=""4""><span class=""style4"">")
    Response.Write(Trim(Request("Year1")) & "年" & Trim(Request("Month1")) & "月份交通安全任務直接執行人員支領獎勵金核發清冊(單位別)")
    Response.Write("</span></td>")
    Response.Write("</tr>")
    Response.Write("<tr>")
    Response.Write("<td style=""width: 200px; height:70px"" ALIGN=""center""></td>")
    Response.Write("<td style=""width: 205px"" ALIGN=""center""><span class=""style2"">舉發件數</span></td>")
    Response.Write("<td style=""width: 205px"" ALIGN=""center""><span class=""style2"">點數</span></td>")
    Response.Write("<td style=""width: 205px"" ALIGN=""center""><span class=""style2"">備考</span></td>")
    Response.Write("</tr>")
    
    '交通隊、保安隊
    Dim UnitMoney, SubUnitScoreTotal, AllUnitScoreTotal, AllMoneyTotal, Money72Total, Money28Total, A1ScoreTmp, A2ScoreTmp, A3ScoreTmp As Decimal
    Dim UnitReward, UnitScore, UnitCount, AllUnitCount, SubUnitCount As Decimal
    Dim GroupMoney2Total As Decimal = 0
    Dim GroupMoney3Total As Decimal = 0
    Dim strUnit As String
    SubUnitScoreTotal = 0
    AllUnitScoreTotal = 0
    AllMoneyTotal = 0
    Money72Total = 0
    Money28Total = 0
    AllMoneyTotal = AllMoneyTotal + Decimal.Round(getMoneyTotal * 0.28 * ShareGroup1)
    Money28Total = Money28Total + Decimal.Round(getMoneyTotal * 0.28 * ShareGroup1)

    '先跑交通隊
    strUnit = "select * from UnitInfo where ShowOrder=0 order by UnitTypeID,UnitID"
    Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
    Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
    If rdUnit.HasRows Then
        While rdUnit.Read()
            UnitScore = 0
            UnitReward = 0
            UnitMoney = 0
            UnitCount = 0
            
            '攔停點數
            Dim strPoint1 As String = "select sum(b.BillType1Score) as cnt,count(*) as BillCnt from BILLBASEVIEWReward a"
            strPoint1 = strPoint1 & ",(select distinct LawVersion,LawItem,BillType1Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
            strPoint1 = strPoint1 & " where IsUsed=1 and CountyOrNpa=0) b"
            strPoint1 = strPoint1 & " where a.RuleVer=b.LawVersion"
            strPoint1 = strPoint1 & " and (a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit("UnitID")) & "'))"
            strPoint1 = strPoint1 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
            strPoint1 = strPoint1 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
            strPoint1 = strPoint1 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPoint1 = strPoint1 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdPoint1 As New Data.OracleClient.OracleCommand(strPoint1, conn)
            Dim rdPoint1 As Data.OracleClient.OracleDataReader = CmdPoint1.ExecuteReader()
            If rdPoint1.HasRows Then
                rdPoint1.Read()
                If rdPoint1("cnt") Is DBNull.Value Then
                    UnitScore = UnitScore + 0
                    UnitCount = UnitCount
                Else
                    UnitScore = UnitScore + CDec(rdPoint1("cnt"))
                    UnitCount = UnitCount + CDec(rdPoint1("BillCnt"))
                End If
            End If
            rdPoint1.Close()
                    
            '逕舉點數
            Dim strPoint2 As String = "select sum(b.BillType2Score) as cnt,count(*) as BillCnt from BILLBASE a"
            strPoint2 = strPoint2 & ",(select distinct LawVersion,LawItem,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
            strPoint2 = strPoint2 & " where IsUsed=1 and CountyOrNpa=0) b"
            strPoint2 = strPoint2 & " where a.RuleVer=b.LawVersion"
            strPoint2 = strPoint2 & " and (a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit("UnitID")) & "'))"
            strPoint2 = strPoint2 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
            strPoint2 = strPoint2 & " and a.BillTypeID='2' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
            strPoint2 = strPoint2 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPoint2 = strPoint2 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdPoint2 As New Data.OracleClient.OracleCommand(strPoint2, conn)
            Dim rdPoint2 As Data.OracleClient.OracleDataReader = CmdPoint2.ExecuteReader()
            If rdPoint2.HasRows Then
                rdPoint2.Read()
                If rdPoint2("cnt") Is DBNull.Value Then
                    UnitScore = UnitScore + 0
                    UnitCount = UnitCount
                Else
                    UnitScore = UnitScore + CDec(rdPoint2("cnt"))
                    UnitCount = UnitCount + CDec(rdPoint2("BillCnt"))
                End If
            End If
            rdPoint2.Close()
                    
            'A1點數
            Dim strPoint3 As String = "select b.A1Score,b.BillType1Score from BILLBASE a"
            strPoint3 = strPoint3 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
            strPoint3 = strPoint3 & " where IsUsed=1 and CountyOrNpa=0) b"
            strPoint3 = strPoint3 & " where a.RuleVer=b.LawVersion"
            strPoint3 = strPoint3 & " and (a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit("UnitID")) & "'))"
            strPoint3 = strPoint3 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
            strPoint3 = strPoint3 & " and a.RecordStateID=0 and (a.TrafficAccidentType='1') and a.BillTypeID='1'"
            strPoint3 = strPoint3 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPoint3 = strPoint3 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdPoint3 As New Data.OracleClient.OracleCommand(strPoint3, conn)
            Dim rdPoint3 As Data.OracleClient.OracleDataReader = CmdPoint3.ExecuteReader()
            If rdPoint3.HasRows Then
                While rdPoint3.Read()
                    If CDec(rdPoint3("BillType1Score")) > CDec(rdPoint3("A1Score")) Then
                        If rdPoint3("BillType1Score") Is DBNull.Value Then
                            UnitScore = UnitScore + 0
                            UnitCount = UnitCount
                        Else
                            UnitScore = UnitScore + CDec(rdPoint3("BillType1Score"))
                            UnitCount = UnitCount + 1
                        End If
                    Else
                        If rdPoint3("A1Score") Is DBNull.Value Then
                            UnitScore = UnitScore + 0
                            UnitCount = UnitCount
                        Else
                            UnitScore = UnitScore + CDec(rdPoint3("A1Score"))
                            UnitCount = UnitCount + 1
                        End If
                    End If

                End While
            End If
            rdPoint3.Close()
                    
            'A2點數
            Dim strPoint4 As String = "select b.A2Score,b.BillType1Score from BILLBASE a"
            strPoint4 = strPoint4 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
            strPoint4 = strPoint4 & " where IsUsed=1 and CountyOrNpa=0) b"
            strPoint4 = strPoint4 & " where a.RuleVer=b.LawVersion"
            strPoint4 = strPoint4 & " and (a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit("UnitID")) & "'))"
            strPoint4 = strPoint4 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
            strPoint4 = strPoint4 & " and a.RecordStateID=0 and (a.TrafficAccidentType='2') and a.BillTypeID='1'"
            strPoint4 = strPoint4 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPoint4 = strPoint4 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdPoint4 As New Data.OracleClient.OracleCommand(strPoint4, conn)
            Dim rdPoint4 As Data.OracleClient.OracleDataReader = CmdPoint4.ExecuteReader()
            If rdPoint4.HasRows Then
                While rdPoint4.Read()
                    If CDec(rdPoint4("BillType1Score")) > CDec(rdPoint4("A2Score")) Then
                        If rdPoint4("BillType1Score") Is DBNull.Value Then
                            UnitScore = UnitScore + 0
                            UnitCount = UnitCount
                        Else
                            UnitScore = UnitScore + CDec(rdPoint4("BillType1Score"))
                            UnitCount = UnitCount + 1
                        End If
                    Else
                        If rdPoint4("A2Score") Is DBNull.Value Then
                            UnitScore = UnitScore + 0
                            UnitCount = UnitCount
                        Else
                            UnitScore = UnitScore + CDec(rdPoint4("A2Score"))
                            UnitCount = UnitCount + 1
                        End If
                    End If
                                
                End While
            End If
            rdPoint4.Close()
                    
            'A3點數
            Dim strPoint5 As String = "select b.A3Score,b.BillType1Score from BILLBASEVIEWReward a"
            strPoint5 = strPoint5 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
            strPoint5 = strPoint5 & " where IsUsed=1 and CountyOrNpa=0) b"
            strPoint5 = strPoint5 & " where a.RuleVer=b.LawVersion"
            strPoint5 = strPoint5 & " and (a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit("UnitID")) & "'))"
            strPoint5 = strPoint5 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
            strPoint5 = strPoint5 & " and a.RecordStateID=0 and (a.TrafficAccidentType='3') and a.BillTypeID='1'"
            strPoint5 = strPoint5 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPoint5 = strPoint5 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
            Dim CmdPoint5 As New Data.OracleClient.OracleCommand(strPoint5, conn)
            Dim rdPoint5 As Data.OracleClient.OracleDataReader = CmdPoint5.ExecuteReader()
            If rdPoint5.HasRows Then
                While rdPoint5.Read()
                    If CDec(rdPoint5("BillType1Score")) > CDec(rdPoint5("A3Score")) Then
                        If rdPoint5("BillType1Score") Is DBNull.Value Then
                            UnitScore = UnitScore + 0
                            UnitCount = UnitCount
                        Else
                            UnitScore = UnitScore + CDec(rdPoint5("BillType1Score"))
                            UnitCount = UnitCount + 1
                        End If
                    Else
                        If rdPoint5("A3Score") Is DBNull.Value Then
                            UnitScore = UnitScore + 0
                            UnitCount = UnitCount
                        Else
                            UnitScore = UnitScore + CDec(rdPoint5("A3Score"))
                            UnitCount = UnitCount + 1
                        End If
                    End If
      
                End While
            End If
            rdPoint5.Close()
            
            '拖吊
            If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "台中市" Then
                Dim strPointT1a As String = "select count(*) cnt from BillBase a"
                strPointT1a = strPointT1a & " where a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit("UnitID")) & "') and a.RecordStateID=0"
                strPointT1a = strPointT1a & " and a.ProjectID='A5'"
                strPointT1a = strPointT1a & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPointT1a = strPointT1a & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPointT1a As New Data.OracleClient.OracleCommand(strPointT1a, conn)
                Dim rdPointT1a As Data.OracleClient.OracleDataReader = CmdPointT1a.ExecuteReader()
                If rdPointT1a.HasRows Then
                    rdPointT1a.Read()
                    If rdPointT1a("cnt") Is DBNull.Value Then
                        UnitScore = UnitScore
                    Else
                        UnitScore = UnitScore + (CDec(rdPointT1a("cnt")) * 20)
                    End If
                End If
                rdPointT1a.Close()
                
                Dim strPointT1b As String = "select count(*) cnt from BillBase a"
                strPointT1b = strPointT1b & " where a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit("UnitID")) & "') and a.RecordStateID=0"
                strPointT1b = strPointT1b & " and a.ProjectID='A6'"
                strPointT1b = strPointT1b & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPointT1b = strPointT1b & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                Dim CmdPointT1b As New Data.OracleClient.OracleCommand(strPointT1b, conn)
                Dim rdPointT1b As Data.OracleClient.OracleDataReader = CmdPointT1b.ExecuteReader()
                If rdPointT1b.HasRows Then
                    rdPointT1b.Read()
                    If rdPointT1b("cnt") Is DBNull.Value Then
                        UnitScore = UnitScore
                    Else
                        
                        UnitScore = UnitScore + (CDec(rdPointT1b("cnt")) * 50)
                    End If
                End If
                rdPointT1b.Close()
            End If
            
            AllUnitCount = AllUnitCount + UnitCount
            AllUnitScoreTotal = AllUnitScoreTotal + UnitScore
            Response.Write("<tr>")
            Response.Write("<td style=""height:30px""><span class=""style1"">" & rdUnit("UnitName") & "</span></td>")
            Response.Write("<td><span class=""style1"">" & UnitCount & "</span></td>")
            Response.Write("<td><span class=""style1"">" & UnitScore & "</span></td>")
            Response.Write("<td><span class=""style1"">&nbsp;</span></td>")
            Response.Write("</tr>")
                    
        End While
    End If
    rdUnit.Close()

    '分局
    '跑分局及底下派出所，分局由派出所加總
    Dim strSubUnit As String
    strSubUnit = "select UnitName,UnitID from UnitInfo where ShowOrder=1 order by UnitTypeID,UnitID"
    Dim CmdSubUnit As New Data.OracleClient.OracleCommand(strSubUnit, conn)
    Dim rdSubUnit As Data.OracleClient.OracleDataReader = CmdSubUnit.ExecuteReader()
    If rdSubUnit.HasRows Then
        While rdSubUnit.Read()
            Dim SubUnitScore As Integer = 0
            SubUnitCount = 0
            Dim strUnit2 As String
            strUnit2 = "select * from UnitInfo where (UnitTypeID='" & Trim(rdSubUnit("UnitID")) & "' and ShowOrder=2) or UnitID='" & Trim(rdSubUnit("UnitID")) & "' order by UnitID"
            Dim CmdUnit2 As New Data.OracleClient.OracleCommand(strUnit2, conn)
            Dim rdUnit2 As Data.OracleClient.OracleDataReader = CmdUnit2.ExecuteReader()
            If rdUnit2.HasRows Then
                While rdUnit2.Read()
                    UnitScore = 0
                    UnitReward = 0
                    
                    '攔停點數
                    Dim strPoint1 As String = "select sum(b.BillType1Score) as cnt,count(*) as BillCnt from BILLBASEVIEWReward a"
                    strPoint1 = strPoint1 & ",(select distinct LawVersion,LawItem,BillType1Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                    strPoint1 = strPoint1 & " where IsUsed=1 and CountyOrNpa=0) b"
                    strPoint1 = strPoint1 & " where a.RuleVer=b.LawVersion"
                    strPoint1 = strPoint1 & " and (a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit2("UnitID")) & "'))"
                    strPoint1 = strPoint1 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                    strPoint1 = strPoint1 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                    strPoint1 = strPoint1 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint1 = strPoint1 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint1 As New Data.OracleClient.OracleCommand(strPoint1, conn)
                    Dim rdPoint1 As Data.OracleClient.OracleDataReader = CmdPoint1.ExecuteReader()
                    If rdPoint1.HasRows Then
                        rdPoint1.Read()
                        If rdPoint1("cnt") Is DBNull.Value Then
                            UnitScore = UnitScore + 0
                            SubUnitCount = SubUnitCount
                        Else
                            UnitScore = UnitScore + CDec(rdPoint1("cnt"))
                            SubUnitCount = SubUnitCount + CDec(rdPoint1("BillCnt"))
                        End If
                    End If
                    rdPoint1.Close()
                    
                    '逕舉點數
                    Dim strPoint2 As String = "select sum(b.BillType2Score) as cnt,count(*) as BillCnt from BILLBASE a"
                    strPoint2 = strPoint2 & ",(select distinct LawVersion,LawItem,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                    strPoint2 = strPoint2 & " where IsUsed=1 and CountyOrNpa=0) b"
                    strPoint2 = strPoint2 & " where a.RuleVer=b.LawVersion"
                    strPoint2 = strPoint2 & " and (a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit2("UnitID")) & "'))"
                    strPoint2 = strPoint2 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                    strPoint2 = strPoint2 & " and a.BillTypeID='2' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                    strPoint2 = strPoint2 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint2 = strPoint2 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint2 As New Data.OracleClient.OracleCommand(strPoint2, conn)
                    Dim rdPoint2 As Data.OracleClient.OracleDataReader = CmdPoint2.ExecuteReader()
                    If rdPoint2.HasRows Then
                        rdPoint2.Read()
                        If rdPoint2("cnt") Is DBNull.Value Then
                            UnitScore = UnitScore + 0
                            SubUnitCount = SubUnitCount
                        Else
                            UnitScore = UnitScore + CDec(rdPoint2("cnt"))
                            SubUnitCount = SubUnitCount + CDec(rdPoint2("BillCnt"))
                        End If
                    End If
                    rdPoint2.Close()
                    
                    'A1點數
                    Dim strPoint3 As String = "select b.A1Score,b.BillType1Score from BILLBASE a"
                    strPoint3 = strPoint3 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                    strPoint3 = strPoint3 & " where IsUsed=1 and CountyOrNpa=0) b"
                    strPoint3 = strPoint3 & " where a.RuleVer=b.LawVersion"
                    strPoint3 = strPoint3 & " and (a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit2("UnitID")) & "'))"
                    strPoint3 = strPoint3 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                    strPoint3 = strPoint3 & " and a.RecordStateID=0 and (a.TrafficAccidentType='1') and a.BillTypeID='1'"
                    strPoint3 = strPoint3 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint3 = strPoint3 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint3 As New Data.OracleClient.OracleCommand(strPoint3, conn)
                    Dim rdPoint3 As Data.OracleClient.OracleDataReader = CmdPoint3.ExecuteReader()
                    If rdPoint3.HasRows Then
                        While rdPoint3.Read()
                            If CDec(rdPoint3("BillType1Score")) > CDec(rdPoint3("A1Score")) Then
                                If rdPoint3("BillType1Score") Is DBNull.Value Then
                                    UnitScore = UnitScore + 0
                                    SubUnitCount = SubUnitCount
                                Else
                                    UnitScore = UnitScore + CDec(rdPoint3("BillType1Score"))
                                    SubUnitCount = SubUnitCount + 1
                                End If
                            Else
                                If rdPoint3("A1Score") Is DBNull.Value Then
                                    UnitScore = UnitScore + 0
                                    SubUnitCount = SubUnitCount
                                Else
                                    UnitScore = UnitScore + CDec(rdPoint3("A1Score"))
                                    SubUnitCount = SubUnitCount + 1
                                End If
                            End If

                        End While
                    End If
                    rdPoint3.Close()
                    
                    'A2點數
                    Dim strPoint4 As String = "select b.A2Score,b.BillType1Score from BILLBASE a"
                    strPoint4 = strPoint4 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                    strPoint4 = strPoint4 & " where IsUsed=1 and CountyOrNpa=0) b"
                    strPoint4 = strPoint4 & " where a.RuleVer=b.LawVersion"
                    strPoint4 = strPoint4 & " and (a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit2("UnitID")) & "'))"
                    strPoint4 = strPoint4 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                    strPoint4 = strPoint4 & " and a.RecordStateID=0 and (a.TrafficAccidentType='2') and a.BillTypeID='1'"
                    strPoint4 = strPoint4 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint4 = strPoint4 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint4 As New Data.OracleClient.OracleCommand(strPoint4, conn)
                    Dim rdPoint4 As Data.OracleClient.OracleDataReader = CmdPoint4.ExecuteReader()
                    If rdPoint4.HasRows Then
                        While rdPoint4.Read()
                            If CDec(rdPoint4("BillType1Score")) > CDec(rdPoint4("A2Score")) Then
                                If rdPoint4("BillType1Score") Is DBNull.Value Then
                                    UnitScore = UnitScore + 0
                                    SubUnitCount = SubUnitCount
                                Else
                                    UnitScore = UnitScore + CDec(rdPoint4("BillType1Score"))
                                    SubUnitCount = SubUnitCount + 1
                                End If
                            Else
                                If rdPoint4("A2Score") Is DBNull.Value Then
                                    UnitScore = UnitScore + 0
                                    SubUnitCount = SubUnitCount
                                Else
                                    UnitScore = UnitScore + CDec(rdPoint4("A2Score"))
                                    SubUnitCount = SubUnitCount + 1
                                End If
                            End If
                                
                        End While
                    End If
                    rdPoint4.Close()
                    
                    'A3點數
                    Dim strPoint5 As String = "select b.A3Score,b.BillType1Score from BILLBASEVIEWReward a"
                    strPoint5 = strPoint5 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                    strPoint5 = strPoint5 & " where IsUsed=1 and CountyOrNpa=0) b"
                    strPoint5 = strPoint5 & " where a.RuleVer=b.LawVersion"
                    strPoint5 = strPoint5 & " and (a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit2("UnitID")) & "'))"
                    strPoint5 = strPoint5 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                    strPoint5 = strPoint5 & " and a.RecordStateID=0 and (a.TrafficAccidentType='3') and a.BillTypeID='1'"
                    strPoint5 = strPoint5 & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint5 = strPoint5 & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint5 As New Data.OracleClient.OracleCommand(strPoint5, conn)
                    Dim rdPoint5 As Data.OracleClient.OracleDataReader = CmdPoint5.ExecuteReader()
                    If rdPoint5.HasRows Then
                        While rdPoint5.Read()
                            If CDec(rdPoint5("BillType1Score")) > CDec(rdPoint5("A3Score")) Then
                                If rdPoint5("BillType1Score") Is DBNull.Value Then
                                    UnitScore = UnitScore + 0
                                    SubUnitCount = SubUnitCount
                                Else
                                    UnitScore = UnitScore + CDec(rdPoint5("BillType1Score"))
                                    SubUnitCount = SubUnitCount + 1
                                End If
                            Else
                                If rdPoint5("A3Score") Is DBNull.Value Then
                                    UnitScore = UnitScore + 0
                                    SubUnitCount = SubUnitCount
                                Else
                                    UnitScore = UnitScore + CDec(rdPoint5("A3Score"))
                                    SubUnitCount = SubUnitCount + 1
                                End If
                            End If
      
                        End While
                    End If
                    rdPoint5.Close()
                    
                                
                    '拖吊
                    If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "台中市" Then
                        Dim strPointT1a As String = "select count(*) cnt from BillBase a"
                        strPointT1a = strPointT1a & " where a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit2("UnitID")) & "') and a.RecordStateID=0"
                        strPointT1a = strPointT1a & " and a.ProjectID='A5'"
                        strPointT1a = strPointT1a & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strPointT1a = strPointT1a & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdPointT1a As New Data.OracleClient.OracleCommand(strPointT1a, conn)
                        Dim rdPointT1a As Data.OracleClient.OracleDataReader = CmdPointT1a.ExecuteReader()
                        If rdPointT1a.HasRows Then
                            rdPointT1a.Read()
                            If rdPointT1a("cnt") Is DBNull.Value Then
                                UnitScore = UnitScore
                            Else
                                UnitScore = UnitScore + (CDec(rdPointT1a("cnt")) * 20)
                            End If
                        End If
                        rdPointT1a.Close()
                
                        Dim strPointT1b As String = "select count(*) cnt from BillBase a"
                        strPointT1b = strPointT1b & " where a.BillMemID1 in (select MemberID from Memberdata where UnitID='" & Trim(rdUnit2("UnitID")) & "') and a.RecordStateID=0"
                        strPointT1b = strPointT1b & " and a.ProjectID='A6'"
                        strPointT1b = strPointT1b & " and a." & theDateType & " between TO_DATE('" & TheAnaDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strPointT1b = strPointT1b & " and TO_DATE('" & TheAnaDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdPointT1b As New Data.OracleClient.OracleCommand(strPointT1b, conn)
                        Dim rdPointT1b As Data.OracleClient.OracleDataReader = CmdPointT1b.ExecuteReader()
                        If rdPointT1b.HasRows Then
                            rdPointT1b.Read()
                            If rdPointT1b("cnt") Is DBNull.Value Then
                                UnitScore = UnitScore
                            Else
                        
                                UnitScore = UnitScore + (CDec(rdPointT1b("cnt")) * 50)
                            End If
                        End If
                        rdPointT1b.Close()
                    End If
                    
                    SubUnitScore = SubUnitScore + UnitScore
                End While
            End If
            rdUnit2.Close()

            AllUnitCount = AllUnitCount + SubUnitCount
            AllUnitScoreTotal = AllUnitScoreTotal + SubUnitScore
            SubUnitScoreTotal = SubUnitScoreTotal + SubUnitScore
            Response.Write("<tr>")
            Response.Write("<td style=""height:30px""><span class=""style1"">" & rdSubUnit("UnitName") & "</span></td>")
            Response.Write("<td><span class=""style1"">" & SubUnitCount & "</span></td>")
            Response.Write("<td><span class=""style1"">" & SubUnitScore & "</span></td>")
            Response.Write("<td><span class=""style1"">&nbsp;</span></td>")
            Response.Write("</tr>")
        End While
    End If
    rdSubUnit.Close()
    

    '合計
    Response.Write("<tr>")
    Response.Write("<td style=""height:30px""><span class=""style1"">合計</span></td>")
    Response.Write("<td><span class=""style1"">" & AllUnitCount & "</span></td>")
    Response.Write("<td><span class=""style1"">" & AllUnitScoreTotal & "</span></td>")
    Response.Write("<td><span class=""style1"">&nbsp;</span></td>")

    Response.Write("</tr>")
    
    Response.Write("</table>")
%>
    </form>
</body>
</html>

