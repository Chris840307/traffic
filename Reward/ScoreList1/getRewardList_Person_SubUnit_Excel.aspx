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
    
    Function GetMemberID(ByVal LoginID, ByVal UnitID)
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
        Dim MemString As String = ""
        Dim strMData As String = "Select MemberID from MemberData where LoginID='" & Trim(LoginID) & "'"
        strMData = strMData & " and UnitID in (Select UnitID from UnitInfo where UnitID='" & UnitID & "' or (UnitTypeID='" & UnitID & "' and ShowOrder=2))"
        Dim CmdMData As New Data.OracleClient.OracleCommand(strMData, conn)
        Dim rdMData As Data.OracleClient.OracleDataReader = CmdMData.ExecuteReader()
        If rdMData.HasRows Then
            While rdMData.Read()
                If MemString = "" Then
                    MemString = Trim(rdMData("MemberID"))
                Else
                    MemString = MemString & "," & Trim(rdMData("MemberID"))
                End If
            End While
            conn.Close()
        End If
        rdMData.Close()
        GetMemberID = MemString
    End Function
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
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
    <title>支領獎勵金核發清冊</title>
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
            
            Dim sys_City As String = ""
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
            Dim FlagPointTotal, UserName As String
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
            
            Dim UserCookie As HttpCookie = Request.Cookies("RewardUser")
            Dim AnalyzeUnitID2 As String = ""
            AnalyzeUnitID2 = Trim(UserCookie.Values("UnitID"))
            Dim UserID As String
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
            If FlagPointTotal = 1 Then  '總分只抓交通隊
                UnitFlag = " and a.BillUnitID='" & AnalyzeUnitID & "'"
            ElseIf FlagPointTotal = 2 Then  '總分只抓分局
                UnitFlag = " and a.BillUnitID in (select UnitID from UnitInfo where UnitID='" & AnalyzeUnitID2 & "' or UnitTypeID='" & AnalyzeUnitID2 & "')"
            Else
                UnitFlag = ""
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
            Dim strPointT2 As String = "select sum(b.BillType2Score) as cnt from BillBaseViewReward a"
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
            Dim strPointT3 As String = "select b.A1Score,b.BillType1Score from BillBaseViewReward a"
            strPointT3 = strPointT3 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
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
                    If CDec(rdPointT3("BillType1Score")) > CDec(rdPointT3("A1Score")) Then
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
            strPointT4 = strPointT4 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
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
                    If CDec(rdPointT4("BillType1Score")) > CDec(rdPointT4("A2Score")) Then
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
            strPointT5 = strPointT5 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
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
                    If CDec(rdPointT5("BillType1Score")) > CDec(rdPointT5("A3Score")) Then
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
            
            Dim PointMoney As Decimal
            If getPointTotal = 0 Then
                PointMoney = 0
            Else
                PointMoney = getMoneyTotal / getPointTotal
            End If
            'Response.Write(PointMoney)
            'Response.End()
            '=================列出清冊內容=========================
            '----------------單位所有人----------------------
            'Dim PersonPay As String
            Dim PersonMoney, PersonPay, BillCnt, BillCntTotal As Decimal
            Dim PersonTotal As Integer = 0
            Dim MaxMoney As Integer = 0
            Dim MinMoney As Integer = 999999
            Dim OverFlag, strDel, strInsert, strCreditID, MemIDList, strPer, strMemID1, LoginIDTmp As String
            If Trim(Request("sMemID")) = "" Then
                
                Dim strU = "select UnitID,UnitName from UnitInfo where UnitID in (" & Trim(Request("sUnitID")) & ") order by UnitID"
                Dim CmdU As New Data.OracleClient.OracleCommand(strU, conn)
                Dim rdU As Data.OracleClient.OracleDataReader = CmdU.ExecuteReader()
                If rdU.HasRows Then
                    While rdU.Read()
                        BillCntTotal = 0
                        PointTotal = 0
                        MoneyTotal = 0
                        Response.Write("<table width=""100%"" border=""1"" cellpadding=""3"" cellspacing=""0"">")
                        Response.Write("<tr>")
                        Response.Write("<td align=""center"" colspan=""5"">")

                        Response.Write("<span class=""style1""><strong>")
                        '統計單位
                        Dim strUnit As String = "select UnitName from UnitInfo where UnitID='" & Trim(rdU("UnitID")) & "'"
                        Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
                        Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
                        If rdUnit.HasRows Then
                            rdUnit.Read()
                        
                            Response.Write(Trim(rdUnit("UnitName")))
                        End If
                        rdUnit.Close()
                        Response.Write("&nbsp; 交通安全任務直接執行人員支領獎勵金核發清冊</strong></span>")
                        Response.Write("<br>")
                        If theDateType = "BillFillDate" Then
                            Response.Write("計算期間(填單日期)：")
                        Else
                            Response.Write("計算期間(建檔日期)：")
                        End If
                        Response.Write(Year(gOutDT(Trim(Request("Date1")))) - 1911 & "/" & Month(gOutDT(Trim(Request("Date1")))) & "/" & Day(gOutDT(Trim(Request("Date1")))))
                        Response.Write(" 至 ")
                        Response.Write(Year(gOutDT(Trim(Request("Date2")))) - 1911 & "/" & Month(gOutDT(Trim(Request("Date2")))) & "/" & Day(gOutDT(Trim(Request("Date2")))))
                        
                        Response.Write("</td>")
                        Response.Write("</tr><tr>")
                        Response.Write("<td style=""width: 300px"">姓名</td>")
                        Response.Write("<td style=""width: 80px"" ALIGN=""right"">舉發件數</td>")
                        Response.Write("<td style=""width: 80px"" ALIGN=""right"">點數</td>")
                        Response.Write("<td style=""width: 80px"" ALIGN=""right"">金額</td>")
                        Response.Write("<td style=""width: 80px"" ALIGN=""right"">核章</td>")
                        Response.Write("</tr>")
        
                        Dim strUnitPerson As String
                        If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "彰化縣" Or sys_City = "台中市" Then
                            strUnitPerson = "select distinct(LoginID) from MemberData where UnitID in (select UnitID from UnitInfo where UnitID='" & Trim(rdU("UnitID")) & "' or (UnitTypeID='" & Trim(rdU("UnitID")) & "' and ShowOrder=2)) order by LoginId"
                        Else
                            strUnitPerson = "select MemberID,LoginID,CHName,Money,CreditID,UnitID,JobID,BankName,BankID,BankAccount from MemberData where UnitID in (select UnitID from UnitInfo where UnitID='" & Trim(rdU("UnitID")) & "' or (UnitTypeID='" & Trim(rdU("UnitID")) & "' and ShowOrder=2)) order by UnitID"
                        End If
                        Dim CmdUnitPerson As New Data.OracleClient.OracleCommand(strUnitPerson, conn)
                        Dim rdUnitPerson As Data.OracleClient.OracleDataReader = CmdUnitPerson.ExecuteReader()
                        If rdUnitPerson.HasRows Then
                            While rdUnitPerson.Read()
                                
                                OverFlag = ""
                                PersonPoint = 0
                                BillCnt = 0
                                
                                If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "彰化縣" Or sys_City = "台中市" Then
                                    MemIDList = GetMemberID(Trim(rdUnitPerson("LoginID")), Trim(rdU("UnitID")))
                                Else
                                    MemIDList = Trim(rdUnitPerson("MemberID"))
                                End If
                                
                                '攔停點數
                                Dim strPoint1 As String = "select sum(b.BillType1Score) as cnt,count(*) as BillCnt from BillBaseViewReward a"
                                strPoint1 = strPoint1 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                                strPoint1 = strPoint1 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                                strPoint1 = strPoint1 & " where a.BillMemID1 in (" & MemIDList & ") and a.RuleVer=b.LawVersion"
                                strPoint1 = strPoint1 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                                strPoint1 = strPoint1 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                                strPoint1 = strPoint1 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPoint1 = strPoint1 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPoint1 As New Data.OracleClient.OracleCommand(strPoint1, conn)
                                Dim rdPoint1 As Data.OracleClient.OracleDataReader = CmdPoint1.ExecuteReader()
                                If rdPoint1.HasRows Then
                                    rdPoint1.Read()
                                    If rdPoint1("cnt") Is DBNull.Value Then
                                        PersonPoint = 0
                                        BillCnt = BillCnt + rdPoint1("BillCnt")
                                    Else
                                        PersonPoint = CDec(rdPoint1("cnt"))
                                        BillCnt = BillCnt + rdPoint1("BillCnt")
                                    End If
                                End If
                                rdPoint1.Close()
                        
                                '逕舉點數
                                Dim strPoint2 As String = "select sum(b.BillType2Score) as cnt,count(*) as BillCnt from BillBaseViewReward a"
                                strPoint2 = strPoint2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                                strPoint2 = strPoint2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                                strPoint2 = strPoint2 & " where a.BillMemID1 in (" & MemIDList & ") and a.RuleVer=b.LawVersion"
                                strPoint2 = strPoint2 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                                strPoint2 = strPoint2 & " and a.BillTypeID='2' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                                strPoint2 = strPoint2 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPoint2 = strPoint2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPoint2 As New Data.OracleClient.OracleCommand(strPoint2, conn)
                                Dim rdPoint2 As Data.OracleClient.OracleDataReader = CmdPoint2.ExecuteReader()
                                If rdPoint2.HasRows Then
                                    rdPoint2.Read()
                                    If rdPoint2("cnt") Is DBNull.Value Then
                                        PersonPoint = PersonPoint + 0
                                        BillCnt = BillCnt + rdPoint2("BillCnt")
                                    Else
                                        PersonPoint = PersonPoint + CDec(rdPoint2("cnt"))
                                        BillCnt = BillCnt + rdPoint2("BillCnt")
                                    End If
                                End If
                                rdPoint2.Close()
                    
                                'A1點數
                                Dim strPoint3 As String = "select b.A1Score,b.BillType1Score from BillBaseViewReward a"
                                strPoint3 = strPoint3 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                                strPoint3 = strPoint3 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                                strPoint3 = strPoint3 & " where a.BillMemID1 in (" & MemIDList & ") and a.RuleVer=b.LawVersion"
                                strPoint3 = strPoint3 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                                strPoint3 = strPoint3 & " and a.RecordStateID=0 and (a.TrafficAccidentType='1') and a.BillTypeID='1'"
                                strPoint3 = strPoint3 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPoint3 = strPoint3 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPoint3 As New Data.OracleClient.OracleCommand(strPoint3, conn)
                                Dim rdPoint3 As Data.OracleClient.OracleDataReader = CmdPoint3.ExecuteReader()
                                If rdPoint3.HasRows Then
                                    While rdPoint3.Read()
                                        If CDec(rdPoint3("BillType1Score")) > CDec(rdPoint3("A1Score")) Then
                                            If rdPoint3("BillType1Score") Is DBNull.Value Then
                                                PersonPoint = PersonPoint + 0
                                                BillCnt = BillCnt + 0
                                            Else
                                                PersonPoint = PersonPoint + CDec(rdPoint3("BillType1Score"))
                                                BillCnt = BillCnt + 1
                                            End If
                                        Else
                                            If rdPoint3("A1Score") Is DBNull.Value Then
                                                PersonPoint = PersonPoint + 0
                                                BillCnt = BillCnt + 0
                                            Else
                                                PersonPoint = PersonPoint + CDec(rdPoint3("A1Score"))
                                                BillCnt = BillCnt + 1
                                            End If
                                        End If
                                    
                                    End While
                                End If
                                rdPoint3.Close()
                                
                                'A2點數
                                Dim strPoint4 As String = "select b.A2Score,b.BillType1Score from BillBaseViewReward a"
                                strPoint4 = strPoint4 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                                strPoint4 = strPoint4 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                                strPoint4 = strPoint4 & " where a.BillMemID1 in (" & MemIDList & ") and a.RuleVer=b.LawVersion"
                                strPoint4 = strPoint4 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                                strPoint4 = strPoint4 & " and a.RecordStateID=0 and (a.TrafficAccidentType='2') and a.BillTypeID='1'"
                                strPoint4 = strPoint4 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPoint4 = strPoint4 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPoint4 As New Data.OracleClient.OracleCommand(strPoint4, conn)
                                Dim rdPoint4 As Data.OracleClient.OracleDataReader = CmdPoint4.ExecuteReader()
                                If rdPoint4.HasRows Then
                                    While rdPoint4.Read()
                                        If CDec(rdPoint4("BillType1Score")) > CDec(rdPoint4("A2Score")) Then
                                            If rdPoint4("BillType1Score") Is DBNull.Value Then
                                                PersonPoint = PersonPoint + 0
                                                BillCnt = BillCnt + 0
                                            Else
                                                PersonPoint = PersonPoint + CDec(rdPoint4("BillType1Score"))
                                                BillCnt = BillCnt + 1
                                            End If
                                        Else
                                            If rdPoint4("A2Score") Is DBNull.Value Then
                                                PersonPoint = PersonPoint + 0
                                                BillCnt = BillCnt + 0
                                            Else
                                                PersonPoint = PersonPoint + CDec(rdPoint4("A2Score"))
                                                BillCnt = BillCnt + 1
                                            End If
                                        End If
                                    End While
                                    
                                End If
                                rdPoint4.Close()
                    
                                'A3點數
                                Dim strPoint5 As String = "select b.A3Score,b.BillType1Score from BillBaseViewReward a"
                                strPoint5 = strPoint5 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                                strPoint5 = strPoint5 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                                strPoint5 = strPoint5 & " where a.BillMemID1 in (" & MemIDList & ") and a.RuleVer=b.LawVersion"
                                strPoint5 = strPoint5 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                                strPoint5 = strPoint5 & " and a.RecordStateID=0 and (a.TrafficAccidentType='3') and a.BillTypeID='1'"
                                strPoint5 = strPoint5 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPoint5 = strPoint5 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPoint5 As New Data.OracleClient.OracleCommand(strPoint5, conn)
                                Dim rdPoint5 As Data.OracleClient.OracleDataReader = CmdPoint5.ExecuteReader()
                                If rdPoint5.HasRows Then
                                    While rdPoint5.Read()
                                        If CDec(rdPoint5("BillType1Score")) > CDec(rdPoint5("A3Score")) Then
                                            If rdPoint5("BillType1Score") Is DBNull.Value Then
                                                PersonPoint = PersonPoint + 0
                                                BillCnt = BillCnt + 0
                                            Else
                                                PersonPoint = PersonPoint + CDec(rdPoint5("BillType1Score"))
                                                BillCnt = BillCnt + 1
                                            End If
                                        Else
                                            If rdPoint5("A3Score") Is DBNull.Value Then
                                                PersonPoint = PersonPoint + 0
                                                BillCnt = BillCnt + 0
                                            Else
                                                PersonPoint = PersonPoint + CDec(rdPoint5("A3Score"))
                                                BillCnt = BillCnt + 1
                                            End If
                                        End If
                                    
                                    End While
                                End If
                                rdPoint5.Close()
                                
                                '拖吊
                                If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "台中市" Then
                                    Dim strPointT1a As String = "select count(*) cnt from BillBase a"
                                    strPointT1a = strPointT1a & " where a.BillMemID1 in (" & MemIDList & ") and a.RecordStateID=0"
                                    strPointT1a = strPointT1a & " and a.ProjectID='A5'"
                                    strPointT1a = strPointT1a & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                    strPointT1a = strPointT1a & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                    Dim CmdPointT1a As New Data.OracleClient.OracleCommand(strPointT1a, conn)
                                    Dim rdPointT1a As Data.OracleClient.OracleDataReader = CmdPointT1a.ExecuteReader()
                                    If rdPointT1a.HasRows Then
                                        rdPointT1a.Read()
                                        If rdPointT1a("cnt") Is DBNull.Value Then
                                            PersonPoint = PersonPoint
                                        Else
                        
                                            PersonPoint = PersonPoint + (CDec(rdPointT1a("cnt")) * 20)
                                        End If
                                    End If
                                    rdPointT1a.Close()
                
                                    Dim strPointT1b As String = "select count(*) cnt from BillBase a"
                                    strPointT1b = strPointT1b & " where a.BillMemID1 in (" & MemIDList & ") and a.RecordStateID=0"
                                    strPointT1b = strPointT1b & " and a.ProjectID='A6'"
                                    strPointT1b = strPointT1b & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                    strPointT1b = strPointT1b & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                    Dim CmdPointT1b As New Data.OracleClient.OracleCommand(strPointT1b, conn)
                                    Dim rdPointT1b As Data.OracleClient.OracleDataReader = CmdPointT1b.ExecuteReader()
                                    If rdPointT1b.HasRows Then
                                        rdPointT1b.Read()
                                        If rdPointT1b("cnt") Is DBNull.Value Then
                                            PersonPoint = PersonPoint
                                        Else
                        
                                            PersonPoint = PersonPoint + (CDec(rdPointT1b("cnt")) * 50)
                                        End If
                                    End If
                                    rdPointT1b.Close()
                                End If

                                Dim MemChName = ""
                                Dim MemMoney = 0
                                If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "彰化縣" Or sys_City = "台中市" Then
                                    strPer = "select Money,ChName,LoginID,MemberID,CreditID,JobID,BankName,BankID,BankAccount from MemberData where RecordStateID=0 and LoginID='" & Trim(rdUnitPerson("LoginID")) & "'"
                                Else
                                    strPer = "select Money,ChName,LoginID,MemberID,CreditID,JobID,BankName,BankID,BankAccount from MemberData where MemberID=" & Trim(rdUnitPerson("MemberID"))
                                End If
                                Dim CmdPer As New Data.OracleClient.OracleCommand(strPer, conn)
                                Dim rdPer As Data.OracleClient.OracleDataReader = CmdPer.ExecuteReader()
                                If rdPer.HasRows Then
                                    rdPer.Read()
                                    MemChName = Trim(rdPer("ChName"))
                                    strMemID1 = Trim(rdPer("MemberID"))
                                    LoginIDTmp = Trim(rdPer("LoginID"))
                                    If rdPer("CreditID") IsNot DBNull.Value Then
                                        strCreditID = Trim(rdPer("CreditID"))
                                    Else
                                        strCreditID = ""
                                    End If
                                    If rdPer("Money") Is DBNull.Value Then
                                        MemMoney = 0
                                    Else
                                        MemMoney = CDec(rdPer("Money"))
                                    End If
                                End If
                                rdPer.Close()

                                If MemMoney = 0 Then
                                    MoneyTotal = MoneyTotal + Decimal.Round(PointMoney * PersonPoint)
                                    PersonMoney = Decimal.Round(PointMoney * PersonPoint)
                                Else
                                    PersonPay = Decimal.Truncate(MemMoney * getPayPercent)
                                    If Decimal.Round(PointMoney * PersonPoint) > PersonPay Then
                                        OverFlag = "＊"
                                        PersonMoney = PersonPay
                                    Else
                                        PersonMoney = Decimal.Round(PointMoney * PersonPoint)
                                    End If
                                    MoneyTotal = MoneyTotal + PersonMoney
                                End If
                                
                                If PersonMoney > 0 Then
                                    '****每月應領實領****
                                    strDel = "delete from RewardMonthData where DirectOrTogether='1'"
                                    strDel = strDel & " and YearMonth=TO_DATE('" & gOutDT(Trim(Request("Date1"))) & "','YYYY/MM/DD')"
                                    strDel = strDel & " and UnitId='" & Trim(rdU("UnitID")) & "' and LoginId='" & LoginIDTmp & "'"
                                    strDel = strDel & " and MemberId=" & strMemID1

                                    Dim cmdDel As New Data.OracleClient.OracleCommand()
                                    cmdDel.CommandText = strDel
                                    cmdDel.Connection = conn
                                    cmdDel.ExecuteNonQuery()
                                    
                                    strInsert = "insert into RewardMonthData(DirectOrTogether,YearMonth,UnitId,LoginId,ChName"
                                    strInsert = strInsert & ",MemberId,CreditID,ShouldGetMoney,RealGetMoney,RecordDate,RecordMemberID)"
                                    strInsert = strInsert & " values('1',TO_DATE('" & gOutDT(Trim(Request("Date1"))) & "','YYYY/MM/DD')"
                                    strInsert = strInsert & ",'" & Trim(rdU("UnitID")) & "','" & LoginIDTmp & "'"
                                    strInsert = strInsert & ",'" & MemChName & "'," & strMemID1 & ",'" & strCreditID & "'," & Decimal.Truncate(PointMoney * PersonPoint)
                                    strInsert = strInsert & "," & PersonMoney & ",sysdate," & UserID
                                    strInsert = strInsert & ")"
                                    Dim cmdInsert As New Data.OracleClient.OracleCommand()
                                    cmdInsert.CommandText = strInsert
                                    cmdInsert.Connection = conn
                                    cmdInsert.ExecuteNonQuery()
                                    '*********************
                                End If
                                
                                If PersonPoint > 0 Then
                                    PointTotal = PointTotal + PersonPoint
                                    BillCntTotal = BillCntTotal + BillCnt
                                
                                    PersonTotal = PersonTotal + 1
                                    If MinMoney > PersonMoney Then
                                        MinMoney = PersonMoney
                                    End If
                                    If MaxMoney < PersonMoney Then
                                        MaxMoney = PersonMoney
                                    End If
                                    Response.Write("<tr>")
                                    Response.Write("<td style=""height: 50px"">" & LoginIDTmp & "&nbsp; " & MemChName & "</td>")
                                    Response.Write("<td ALIGN=""right"">" & Format(BillCnt, "##,##0.00") & "</td>")
                                    Response.Write("<td ALIGN=""right"">" & Format(PersonPoint, "##,##0.00") & "</td>")
                                    Response.Write("<td ALIGN=""right"">" & Format(PersonMoney, "##,##0.00") & "</td>")
                                    Response.Write("<td ALIGN=""center"">" & OverFlag & "</td>")
                                    Response.Write("</tr>")
                                End If
                            End While
                        End If
                        rdUnitPerson.Close()
                
                        Response.Write("<tr>")
                        Response.Write("<td>總計</td>")
                        Response.Write("<td ALIGN=""right"">" & Format(BillCntTotal, "##,##0.00") & "</td>")
                        Response.Write("<td ALIGN=""right"">" & Format(PointTotal, "##,##0.00") & "</td>")
                        Response.Write("<td ALIGN=""right"">" & Format(MoneyTotal, "##,##0.00") & "</td>")
                        Response.Write("<td align=""center""></td>")
                        Response.Write("</tr>")
                        'Response.Write("<tr><td><span class=""style1""><strong>製表</strong></span></td>")
                        'Response.Write("<td><span class=""style1""><strong>組長</strong></span></td>")
                        'Response.Write("<td><span class=""style1""><strong>副隊長</strong></span></td>")
                        'Response.Write("<td><span class=""style1""><strong>隊長</strong></span></td>")
                        'Response.Write("</tr>")
                        Response.Write("<tr><td colspan=""5""></td>")
                        Response.Write("</tr>")
                        Response.Write("</table>")
                        
                        'Response.Write("<div class=""PageNext""></div>")
        
    
                    End While
                End If
                rdU.Close()

            End If
            '======================================================
            conn.Close()
        %>
    
        
    </form>
</body>
<script type="text/javascript" src="../form.js"></script>
<script language="JavaScript">
	//printWindow(true,5.08,5.08,5.08,5.08);
</script>
</html>
