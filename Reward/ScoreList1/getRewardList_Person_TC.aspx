<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%  LoginCheck()
    Server.ScriptTimeout = 86400
    Response.Flush()
%>
<object id="factory" style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://localhost/traffic/smsx.cab#Version=6,1,432,1">
</object>

<script runat="server">
    Public PersonPoint, PointTotal, MoneyTotal As Decimal
    Public sys_City As String = ""

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
        Dim strMData As String = "Select MemberID from MemberData where LoginID='" & Trim(LoginID) & "' and UnitID='" & Trim(Replace(UnitID, "'", "")) & "'"
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
        conn.Close()
    End Function
    
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
        
        Dim strCity = "select Value from ApConfigure where ID=31"
        Dim CmdCity As New Data.OracleClient.OracleCommand(strCity, conn)
        Dim rdCity As Data.OracleClient.OracleDataReader = CmdCity.ExecuteReader()
        If rdCity.HasRows Then
            rdCity.Read()
            sys_City = Trim(rdCity("Value"))
        End If
        rdCity.Close()
        conn.Close()
    End Sub
</script>

<script language="JavaScript">
	window.focus();
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
<style type="text/css">
<!--
<%
if sys_City="台中縣" or sys_City = "台中市" then
    response.write("body {font-family:新細明體;font-size:11pt }")
    response.write(".style1 {font-family:新細明體; font-size: 12pt}")
else
    response.write("body {font-family:新細明體;font-size:12pt }")
    response.write(".style1 {font-family:新細明體; font-size: 13pt}")
end if

 %>



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
            
            Dim RewardMonth As String
            RewardMonth = Left(Trim(Request("Date1")), Len(Trim(Request("Date1"))) - 4) & Mid(Trim(Request("Date1")), Len(Trim(Request("Date1"))) - 3, 2)

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
            strPointT1 = strPointT1 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS') and c.ShowOrder<>-1"
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
            strPointT2 = strPointT2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS') and c.ShowOrder<>-1"
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
            strPointT3 = strPointT3 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS') and c.ShowOrder<>-1"
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
            strPointT4 = strPointT4 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS') and c.ShowOrder<>-1"
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
            strPointT5 = strPointT5 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS') and c.ShowOrder<>-1"
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
                strPointT1a = strPointT1a & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS') and c.ShowOrder<>-1"
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
                strPointT1b = strPointT1b & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS') and c.ShowOrder<>-1"
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
           
            'Response.Write(PointMoney & " " & getPointTotal)
            '=================列出清冊內容=========================
            '----------------單位所有人----------------------
            'Dim PersonPay As String
            Dim PersonMoney, PersonPay, BillCnt, BillCntTotal As Decimal
            Dim PersonTotal As Integer = 0
            Dim MaxMoney As Integer = 0
            Dim MinMoney As Integer = 999999
            Dim OverFlag, strCreditID, strDel, strInsert, MemIDList, strUnitPerson, strBank, JobIDtmp As String
            Dim PageNo As Integer = 1
            Dim PageSum As Integer = 0
            Dim PageBillCnt, PagePoint As Decimal
            Dim PageCount, i, UnitCnt, UnitCntTmp, PageTmp As Integer
            Dim JobIDPlus As String
            If sys_City = "台中縣" Or sys_City = "台中市" Then
                JobIDPlus = ""
                PageCount = 19
            Else
                JobIDPlus = ""
                PageCount = 14
            End If
            If Trim(Request("sMemID")) = "" Then
                PageTmp = 0
                Dim strU = "select UnitID,UnitName from UnitInfo where UnitID in (" & Trim(Request("sUnitID")) & ") order by UnitID"
                Dim CmdU As New Data.OracleClient.OracleCommand(strU, conn)
                Dim rdU As Data.OracleClient.OracleDataReader = CmdU.ExecuteReader()
                If rdU.HasRows Then
                    While rdU.Read()
                        BillCntTotal = 0
                        PointTotal = 0
                        MoneyTotal = 0
                        
                        UnitCntTmp = 0
                        UnitCnt = 0
                        
                        Dim strCnt As String
                        strCnt = "select Count(distinct(LoginID)) as cnt from MemberData where UnitID='" & Trim(rdU("UnitID")) & "' " & JobIDPlus
                        Dim CmdCnt As New Data.OracleClient.OracleCommand(strCnt, conn)
                        Dim rdCnt As Data.OracleClient.OracleDataReader = CmdCnt.ExecuteReader()
                        If rdCnt.HasRows Then
                            rdCnt.Read()
                            UnitCnt = rdCnt("cnt")
                        End If
                        rdCnt.Close()

                        If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "彰化縣" Or sys_City = "台中市" Then
                            strUnitPerson = "select distinct(LoginID) from MemberData where UnitID='" & Trim(rdU("UnitID")) & "' " & JobIDPlus & " order by LoginID"
                        Else
                            strUnitPerson = "select MemberID,LoginID,CHName,Money,CreditID from MemberData where UnitID='" & Trim(rdU("UnitID")) & "' " & JobIDPlus & "  order by LoginID"
                        End If
                        Dim CmdUnitPerson As New Data.OracleClient.OracleCommand(strUnitPerson, conn)
                        Dim rdUnitPerson As Data.OracleClient.OracleDataReader = CmdUnitPerson.ExecuteReader()
                        If rdUnitPerson.HasRows Then
                            While rdUnitPerson.Read()
                                '--------------------------------------------
                                OverFlag = ""
                                PersonPoint = 0
                                BillCnt = 0
                                UnitCntTmp = UnitCntTmp + 1
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
                                Dim strPoint2 As String = "select sum(b.BillType2Score) as cnt,count(*) as BillCnt from BillBase a"
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
                                Dim strPoint3 As String = "select count(*) as cnt from BillBase a"
                                strPoint3 = strPoint3 & " where a.BillMemID1 in (" & MemIDList & ") "
                                strPoint3 = strPoint3 & " and a.RecordStateID=0 and (a.TrafficAccidentType='1') and a.BillTypeID='1'"
                                strPoint3 = strPoint3 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPoint3 = strPoint3 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPoint3 As New Data.OracleClient.OracleCommand(strPoint3, conn)
                                Dim rdPoint3 As Data.OracleClient.OracleDataReader = CmdPoint3.ExecuteReader()
                                If rdPoint3.HasRows Then
                                    rdPoint3.Read()
                                    PersonPoint = PersonPoint + (CDec(rdPoint3("cnt")) * 100)
                                    BillCnt = BillCnt + CDec(rdPoint3("cnt"))
                                    
                                End If
                                rdPoint3.Close()
                    
                                'A2點數
                                Dim strPoint4 As String = "select count(*) as cnt from BillBase a"
                                strPoint4 = strPoint4 & " where a.BillMemID1 in (" & MemIDList & ") "
                                strPoint4 = strPoint4 & " and a.RecordStateID=0 and (a.TrafficAccidentType='2') and a.BillTypeID='1'"
                                strPoint4 = strPoint4 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPoint4 = strPoint4 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPoint4 As New Data.OracleClient.OracleCommand(strPoint4, conn)
                                Dim rdPoint4 As Data.OracleClient.OracleDataReader = CmdPoint4.ExecuteReader()
                                If rdPoint4.HasRows Then
                                    rdPoint4.Read()
                                    PersonPoint = PersonPoint + (CDec(rdPoint4("cnt")) * 50)
                                    BillCnt = BillCnt + CDec(rdPoint4("cnt"))
                                    
                                End If
                                rdPoint4.Close()
                    
                                'A3點數
                                Dim strPoint5 As String = "select count(*) as cnt from BillBase a"
                                strPoint5 = strPoint5 & " where a.BillMemID1 in (" & MemIDList & ") "
                                strPoint5 = strPoint5 & " and a.RecordStateID=0 and (a.TrafficAccidentType='3') and a.BillTypeID='1'"
                                strPoint5 = strPoint5 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strPoint5 = strPoint5 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdPoint5 As New Data.OracleClient.OracleCommand(strPoint5, conn)
                                Dim rdPoint5 As Data.OracleClient.OracleDataReader = CmdPoint5.ExecuteReader()
                                If rdPoint5.HasRows Then
                                    rdPoint5.Read()
                                    PersonPoint = PersonPoint + (CDec(rdPoint5("cnt")) * 20)
                                    BillCnt = BillCnt + CDec(rdPoint5("cnt"))
                                    
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
                                
                                Dim strPer, MemChName, strMemID1 As String
                                Dim MemMoney As Decimal
                                If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "彰化縣" Or sys_City = "台中市" Then
                                    strPer = "select Money,ChName,MemberID,CreditID,JobID,BankName,BankID,BankAccount from MemberData where LoginID='" & Trim(rdUnitPerson("LoginID")) & "' order by RecordDate Desc"
                                Else
                                    strPer = "select Money,ChName,MemberID,CreditID,JobID,BankName,BankID,BankAccount from MemberData where MemberID=" & Trim(rdUnitPerson("MemberID"))
                                End If
                                Dim CmdPer As New Data.OracleClient.OracleCommand(strPer, conn)
                                Dim rdPer As Data.OracleClient.OracleDataReader = CmdPer.ExecuteReader()
                                If rdPer.HasRows Then
                                    rdPer.Read()
                                    MemChName = Trim(rdPer("ChName"))
                                    strMemID1 = Trim(rdPer("MemberID"))
                                    JobIDtmp = Trim(rdPer("JobID"))
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
                                    
                                    If rdPer("BankName") IsNot DBNull.Value Then
                                        strBank = rdPer("BankName")
                                    Else
                                        strBank = "&nbsp;"
                                    End If
                                    If rdPer("BankID") IsNot DBNull.Value Then
                                        strBank = strBank & rdPer("BankID")
                                    Else
                                        strBank = strBank
                                    End If
                                    If rdPer("BankAccount") IsNot DBNull.Value Then
                                        strBank = strBank & rdPer("BankAccount")
                                    Else
                                        strBank = strBank
                                    End If
                                End If
                                rdPer.Close()

                                If MemMoney = 0 Then
                                    PersonMoney = Decimal.Round(PointMoney * PersonPoint)
                                Else
                                    PersonPay = Decimal.Truncate(MemMoney * getPayPercent)
                                    If Decimal.Round(PointMoney * PersonPoint) > PersonPay Then
                                        OverFlag = "＊"
                                        PersonMoney = PersonPay
                                    Else
                                        PersonMoney = Decimal.Round(PointMoney * PersonPoint)
                                    End If
                                End If
                                '扣趴判斷
                                Dim strLMoney1 As String = "select * from RewardSpecMember where LoginID='" & Trim(rdUnitPerson("LoginID")) & "' and ROWNUM <=1 and YearMonth='" & RewardMonth & "'"
                                Dim CmdLMoney1 As New Data.OracleClient.OracleCommand(strLMoney1, conn)
                                Dim rdLMoney1 As Data.OracleClient.OracleDataReader = CmdLMoney1.ExecuteReader()
                                If rdLMoney1.HasRows Then
                                    rdLMoney1.Read()
                                    PersonMoney = Decimal.Truncate(PersonMoney * (rdLMoney1("DePercent") / 100))
                                End If
                                rdLMoney1.Close()
                                
                                MoneyTotal = MoneyTotal + PersonMoney
                                If PersonMoney > 0 Then
                                    PageSum = PageSum + PersonMoney
                                    '****每月應領實領****
                                    strDel = "delete from RewardMonthData where DirectOrTogether='1'"
                                    strDel = strDel & " and YearMonth=TO_DATE('" & gOutDT(Trim(Request("Date1"))) & "','YYYY/MM/DD')"
                                    strDel = strDel & " and UnitId='" & Trim(rdU("UnitID")) & "' and LoginId='" & Trim(rdUnitPerson("LoginID")) & "'"
                                    strDel = strDel & " and MemberId=" & strMemID1

                                    Dim cmdDel As New Data.OracleClient.OracleCommand()
                                    cmdDel.CommandText = strDel
                                    cmdDel.Connection = conn
                                    cmdDel.ExecuteNonQuery()
                                    
                                    strInsert = "insert into RewardMonthData(DirectOrTogether,YearMonth,UnitId,LoginId,ChName"
                                    strInsert = strInsert & ",MemberId,CreditID,ShouldGetMoney,RealGetMoney,RecordDate,RecordMemberID)"
                                    strInsert = strInsert & " values('1',TO_DATE('" & gOutDT(Trim(Request("Date1"))) & "','YYYY/MM/DD')"
                                    strInsert = strInsert & ",'" & Trim(rdU("UnitID")) & "','" & Trim(rdUnitPerson("LoginID")) & "'"
                                    strInsert = strInsert & ",'" & MemChName & "'," & strMemID1 & ",'" & strCreditID & "'," & Decimal.Round(PointMoney * PersonPoint)
                                    strInsert = strInsert & "," & PersonMoney & ",sysdate," & UserID
                                    strInsert = strInsert & ")"
                                    Dim cmdInsert As New Data.OracleClient.OracleCommand()
                                    cmdInsert.CommandText = strInsert
                                    cmdInsert.Connection = conn
                                    cmdInsert.ExecuteNonQuery()
                                    '*********************
                                End If
                                
                                If PersonMoney > 0 Then
                                    PageBillCnt = PageBillCnt + BillCnt
                                    PagePoint = PagePoint + PersonPoint
                                    PointTotal = PointTotal + PersonPoint
                                    BillCntTotal = BillCntTotal + BillCnt
                                
                                    PersonTotal = PersonTotal + 1
                                    If MinMoney > PersonMoney And PersonMoney > 0 Then
                                        MinMoney = PersonMoney
                                    End If
                                    If MaxMoney < PersonMoney Then
                                        MaxMoney = PersonMoney
                                    End If
                                    If PageTmp = 1 Then
                                        Response.Write("<div class=""PageNext"">&nbsp;</div>")
                                    End If
                                    PageTmp = 1
                                    Response.Write("<table width=""680"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                                    Response.Write("<td align=""center"" colspan=""3""><span class=""style1""><strong>")
                                    '統計單位
                                    Dim strUnit As String = "select UnitName from UnitInfo where UnitID='" & Trim(rdU("UnitID")) & "'"
                                    Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
                                    Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
                                    If rdUnit.HasRows Then
                                        rdUnit.Read()
                                        Response.Write(CityUnitName & Trim(rdUnit("UnitName")) & "<br>")
                                    End If
                                    rdUnit.Close()
                                    If theDateType = "BillFillDate" Then
                                        Response.Write("計算期間(填單日期)：")
                                    Else
                                        Response.Write("計算期間(建檔日期)：")
                                    End If
                                    Response.Write(Year(gOutDT(Trim(Request("Date1")))) - 1911 & "/" & Month(gOutDT(Trim(Request("Date1")))) & "/" & Day(gOutDT(Trim(Request("Date1")))))
                                    Response.Write(" 至 ")
                                    Response.Write(Year(gOutDT(Trim(Request("Date2")))) - 1911 & "/" & Month(gOutDT(Trim(Request("Date2")))) & "/" & Day(gOutDT(Trim(Request("Date2")))))
                                    If sys_City = "台中縣" Then
                                        Response.Write("&nbsp; 直接執行人員交通安全獎金請領清冊</strong></span>")
                                    Else
                                        Response.Write("&nbsp; 直接執行人員交通績效獎金請領清冊</strong></span>")
                                    End If
                                    Response.Write("<br>")
                                    Response.Write("</td></tr>")
                                    Response.Write("<tr>")
                                    Response.Write("<td width=""34%"">")
                                    Response.Write("列印日期：" & Year(Now) & "/" & Month(Now) & "/" & Day(Now))
                                    Response.Write("</td>")
                                    Response.Write("<td width=""33%"">")
                                    Response.Write("列印人員：" & UserName)
                                    Response.Write("</td>")
                                    Response.Write("<td width=""33%"">")
                                    Response.Write("頁次：" & PageNo)
                                    Response.Write("</td>")
                                    Response.Write("</tr>")
                                    Response.Write("</table>")
                        
                                    Response.Write("<table width=""680"" border=""1"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                                    Response.Write("<tr>")
                                    
                                    If sys_City = "台中縣" Or sys_City = "台中市" Then
                                        Response.Write("<td style=""width: 8%;height:35px"" ALIGN=""center"">單位</td>")
                                        Response.Write("<td style=""width: 12%"" ALIGN=""center"">職位</td>")
                                        Response.Write("<td style=""width: 12%"" ALIGN=""center"">姓名</td>")
                                        Response.Write("<td style=""width: 10%"" ALIGN=""center"">件數</td>")
                                        Response.Write("<td style=""width: 10%"" ALIGN=""center"">點數</td>")
                                        Response.Write("<td style=""width: 12%"" ALIGN=""center"">實領獎金</td>")
                                        Response.Write("<td style=""width: 25%"" ALIGN=""center"">郵局帳號</td>")
                                        Response.Write("<td style=""width: 12%"" ALIGN=""center"">備註</td>")
                                    Else
                                        Response.Write("<td style=""width: 8%;height:35px"" ALIGN=""center"">單位</td>")
                                        Response.Write("<td style=""width: 15%"" ALIGN=""center"">職位</td>")
                                        Response.Write("<td style=""width: 15%"" ALIGN=""center"">姓名</td>")
                                        Response.Write("<td style=""width: 12%"" ALIGN=""center"">實領獎金</td>")
                                        Response.Write("<td style=""width: 13%"" ALIGN=""center"">身份證號</td>")
                                        Response.Write("<td style=""width: 25%"" ALIGN=""center"">郵局帳號</td>")
                                        Response.Write("<td style=""width: 12%"" ALIGN=""center"">備註</td>")
                                    End If
                                    Response.Write("</tr>")

                                    Response.Write("<tr>")
                                    Response.Write("<td style=""height: 33px"" ALIGN=""center"">" & rdU("UnitID") & "</td>")
                                    Response.Write("<td ALIGN=""center"">")
                                    Dim strJob As String = "select * from code where typeid=4 and id=" & Trim(JobIDtmp)
                                    Dim CmdJob As New Data.OracleClient.OracleCommand(strJob, conn)
                                    Dim rdJob As Data.OracleClient.OracleDataReader = CmdJob.ExecuteReader()
                                    If rdJob.HasRows Then
                                        rdJob.Read()
                                        Response.Write(rdJob("Content"))
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    rdJob.Close()
                                    Response.Write("</td>")
                                    Response.Write("<td ALIGN=""center"">")
                                    If MemChName <> "" Then
                                        Response.Write(MemChName)
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                    If sys_City = "台中縣" Or sys_City = "台中市" Then
                                        Response.Write("<td ALIGN=""center"">")
                                        Response.Write(BillCnt)
                                        Response.Write("</td>")
                                        Response.Write("<td ALIGN=""center"">")
                                        Response.Write(PersonPoint)
                                        Response.Write("</td>")
                                    End If
                                    Response.Write("<td ALIGN=""center"">" & PersonMoney & "</td>")
                                    If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                        Response.Write("<td ALIGN=""center"">")
                                        'If sys_City = "台中縣" Then
                                        '    If rdUnitPerson("CreditID") IsNot DBNull.Value Then
                                        '        Response.Write(rdUnitPerson("CreditID"))
                                        '    Else
                                        '        Response.Write("&nbsp;")
                                        '    End If
                                        'Else
                                        Response.Write("&nbsp;")
                                        'End If
                                        Response.Write("</td>")
                                    End If
                                    Response.Write("<td ALIGN=""center"">")
                                    If sys_City = "台中縣" Or sys_City = "台中市" Then
                                        Response.Write(strBank)
                                    Else
                                        Response.Write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                    Response.Write("<td ALIGN=""center"">&nbsp;</td>")
                                    Response.Write("</tr>")
                                 
                                    For i = 1 To PageCount
                                        OverFlag = ""
                                        PersonPoint = 0
                                        BillCnt = 0

                                        If rdUnitPerson.Read() = True Then
                                            UnitCntTmp = UnitCntTmp + 1

                                            If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "彰化縣" Or sys_City = "台中市" Then
                                                MemIDList = GetMemberID(Trim(rdUnitPerson("LoginID")), Trim(rdU("UnitID")))
                                            Else
                                                MemIDList = Trim(rdUnitPerson("MemberID"))
                                            End If

                                
                                            '攔停點數
                                            Dim strPoint1b As String = "select sum(b.BillType1Score) as cnt,count(*) as BillCnt from BillBaseViewReward a"
                                            strPoint1b = strPoint1b & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                                            strPoint1b = strPoint1b & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                                            strPoint1b = strPoint1b & " where a.BillMemID1 in (" & MemIDList & ") and a.RuleVer=b.LawVersion"
                                            strPoint1b = strPoint1b & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                                            strPoint1b = strPoint1b & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                                            strPoint1b = strPoint1b & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                            strPoint1b = strPoint1b & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                            Dim CmdPoint1b As New Data.OracleClient.OracleCommand(strPoint1b, conn)
                                            Dim rdPoint1b As Data.OracleClient.OracleDataReader = CmdPoint1b.ExecuteReader()
                                            If rdPoint1b.HasRows Then
                                                rdPoint1b.Read()
                                                If rdPoint1b("cnt") Is DBNull.Value Then
                                                    PersonPoint = 0
                                                    BillCnt = BillCnt + rdPoint1b("BillCnt")
                                                Else
                                                    PersonPoint = CDec(rdPoint1b("cnt"))
                                                    BillCnt = BillCnt + rdPoint1b("BillCnt")
                                                End If
                                            End If
                                            rdPoint1b.Close()
                        
                                            '逕舉點數
                                            Dim strPoint2b As String = "select sum(b.BillType2Score) as cnt,count(*) as BillCnt from BillBase a"
                                            strPoint2b = strPoint2b & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                                            strPoint2b = strPoint2b & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                                            strPoint2b = strPoint2b & " where a.BillMemID1 in (" & MemIDList & ") and a.RuleVer=b.LawVersion"
                                            strPoint2b = strPoint2b & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                                            strPoint2b = strPoint2b & " and a.BillTypeID='2' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                                            strPoint2b = strPoint2b & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                            strPoint2b = strPoint2b & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                            Dim CmdPoint2b As New Data.OracleClient.OracleCommand(strPoint2b, conn)
                                            Dim rdPoint2b As Data.OracleClient.OracleDataReader = CmdPoint2b.ExecuteReader()
                                            If rdPoint2b.HasRows Then
                                                rdPoint2b.Read()
                                                If rdPoint2b("cnt") Is DBNull.Value Then
                                                    PersonPoint = PersonPoint + 0
                                                    BillCnt = BillCnt + rdPoint2b("BillCnt")
                                                Else
                                                    PersonPoint = PersonPoint + CDec(rdPoint2b("cnt"))
                                                    BillCnt = BillCnt + rdPoint2b("BillCnt")
                                                End If
                                            End If
                                            rdPoint2b.Close()
                    
                                            'A1點數
                                            Dim strPoint3b As String = "select count(*) as cnt from BillBase a"
                                            strPoint3b = strPoint3b & " where a.BillMemID1 in (" & MemIDList & ") "
                                            strPoint3b = strPoint3b & " and a.RecordStateID=0 and (a.TrafficAccidentType='1') and a.BillTypeID='1'"
                                            strPoint3b = strPoint3b & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                            strPoint3b = strPoint3b & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                            Dim CmdPoint3b As New Data.OracleClient.OracleCommand(strPoint3b, conn)
                                            Dim rdPoint3b As Data.OracleClient.OracleDataReader = CmdPoint3b.ExecuteReader()
                                            If rdPoint3b.HasRows Then
                                                rdPoint3b.Read()
                                                PersonPoint = PersonPoint + (CDec(rdPoint3b("cnt")) * 100)
                                                BillCnt = BillCnt + CDec(rdPoint3b("cnt"))
                                    
                                            End If
                                            rdPoint3b.Close()
                    
                                            'A2點數
                                            Dim strPoint4b As String = "select count(*) as cnt from BillBase a"
                                            strPoint4b = strPoint4b & " where a.BillMemID1 in (" & MemIDList & ")"
                                            strPoint4b = strPoint4b & " and a.RecordStateID=0 and (a.TrafficAccidentType='2') and a.BillTypeID='1'"
                                            strPoint4b = strPoint4b & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                            strPoint4b = strPoint4b & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                            Dim CmdPoint4b As New Data.OracleClient.OracleCommand(strPoint4b, conn)
                                            Dim rdPoint4b As Data.OracleClient.OracleDataReader = CmdPoint4b.ExecuteReader()
                                            If rdPoint4b.HasRows Then
                                                rdPoint4b.Read()
                                                PersonPoint = PersonPoint + (CDec(rdPoint4b("cnt")) * 50)
                                                BillCnt = BillCnt + CDec(rdPoint4b("cnt"))
                                    
                                            End If
                                            rdPoint4b.Close()
                    
                                            'A3點數
                                            Dim strPoint5b As String = "select count(*) as cnt from BillBase a"
                                            strPoint5b = strPoint5b & " where a.BillMemID1 in (" & MemIDList & ") "
                                            strPoint5b = strPoint5b & " and a.RecordStateID=0 and (a.TrafficAccidentType='3') and a.BillTypeID='1'"
                                            strPoint5b = strPoint5b & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                            strPoint5b = strPoint5b & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                            Dim CmdPoint5b As New Data.OracleClient.OracleCommand(strPoint5b, conn)
                                            Dim rdPoint5b As Data.OracleClient.OracleDataReader = CmdPoint5b.ExecuteReader()
                                            If rdPoint5b.HasRows Then
                                                rdPoint5b.Read()
                                                PersonPoint = PersonPoint + (CDec(rdPoint5b("cnt")) * 20)
                                                BillCnt = BillCnt + CDec(rdPoint5b("cnt"))
                                    
                                            End If
                                            rdPoint5b.Close()
                                            
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
                                            
                                            Dim strPer2 As String
                                            If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "彰化縣" Or sys_City = "台中市" Then
                                                strPer2 = "select Money,ChName,MemberID,CreditID,JobID,BankName,BankID,BankAccount from MemberData where LoginID='" & Trim(rdUnitPerson("LoginID")) & "' order by RecordDate Desc"
                                            Else
                                                strPer2 = "select Money,ChName,MemberID,CreditID,JobID,BankName,BankID,BankAccount from MemberData where MemberID=" & Trim(rdUnitPerson("MemberID"))
                                            End If
                                            Dim CmdPer2 As New Data.OracleClient.OracleCommand(strPer2, conn)
                                            Dim rdPer2 As Data.OracleClient.OracleDataReader = CmdPer2.ExecuteReader()
                                            If rdPer2.HasRows Then
                                                rdPer2.Read()
                                                MemChName = Trim(rdPer2("ChName"))
                                                strMemID1 = Trim(rdPer2("MemberID"))
                                                JobIDtmp = Trim(rdPer2("JobID"))
                                                If rdPer2("CreditID") IsNot DBNull.Value Then
                                                    strCreditID = Trim(rdPer2("CreditID"))
                                                Else
                                                    strCreditID = ""
                                                End If
                                                If rdPer2("Money") Is DBNull.Value Then
                                                    MemMoney = 0
                                                Else
                                                    MemMoney = CDec(rdPer2("Money"))
                                                End If

                                                If rdPer2("BankName") IsNot DBNull.Value Then
                                                    strBank = rdPer2("BankName")
                                                Else
                                                    strBank = "&nbsp;"
                                                End If
                                                If rdPer2("BankID") IsNot DBNull.Value Then
                                                    strBank = strBank & rdPer2("BankID")
                                                Else
                                                    strBank = strBank
                                                End If
                                                If rdPer2("BankAccount") IsNot DBNull.Value Then
                                                    strBank = strBank & rdPer2("BankAccount")
                                                Else
                                                    strBank = strBank
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
                                
                                            '扣趴判斷
                                            Dim strLMoney1b As String = "select * from RewardSpecMember where LoginID='" & Trim(rdUnitPerson("LoginID")) & "' and ROWNUM <=1 and YearMonth='" & RewardMonth & "'"
                                            Dim CmdLMoney1b As New Data.OracleClient.OracleCommand(strLMoney1b, conn)
                                            Dim rdLMoney1b As Data.OracleClient.OracleDataReader = CmdLMoney1b.ExecuteReader()
                                            If rdLMoney1b.HasRows Then
                                                rdLMoney1b.Read()
                                                PersonMoney = Decimal.Truncate(PersonMoney * (rdLMoney1b("DePercent") / 100))
                                            End If
                                            rdLMoney1b.Close()
                                            
                                            If PersonMoney > 0 Then
                                                PageSum = PageSum + PersonMoney
                                                '****每月應領實領****
                                                strDel = "delete from RewardMonthData where DirectOrTogether='1'"
                                                strDel = strDel & " and YearMonth=TO_DATE('" & gOutDT(Trim(Request("Date1"))) & "','YYYY/MM/DD')"
                                                strDel = strDel & " and UnitId='" & Trim(rdU("UnitID")) & "' and LoginId='" & Trim(rdUnitPerson("LoginID")) & "'"
                                                strDel = strDel & " and MemberId=" & strMemID1

                                                Dim cmdDel As New Data.OracleClient.OracleCommand()
                                                cmdDel.CommandText = strDel
                                                cmdDel.Connection = conn
                                                cmdDel.ExecuteNonQuery()
                                    
                                                strInsert = "insert into RewardMonthData(DirectOrTogether,YearMonth,UnitId,LoginId,ChName"
                                                strInsert = strInsert & ",MemberId,CreditID,ShouldGetMoney,RealGetMoney,RecordDate,RecordMemberID)"
                                                strInsert = strInsert & " values('1',TO_DATE('" & gOutDT(Trim(Request("Date1"))) & "','YYYY/MM/DD')"
                                                strInsert = strInsert & ",'" & Trim(rdU("UnitID")) & "','" & Trim(rdUnitPerson("LoginID")) & "'"
                                                strInsert = strInsert & ",'" & MemChName & "'," & strMemID1 & ",'" & strCreditID & "'," & Decimal.Round(PointMoney * PersonPoint)
                                                strInsert = strInsert & "," & PersonMoney & ",sysdate," & UserID
                                                strInsert = strInsert & ")"
                                                Dim cmdInsert As New Data.OracleClient.OracleCommand()
                                                cmdInsert.CommandText = strInsert
                                                cmdInsert.Connection = conn
                                                cmdInsert.ExecuteNonQuery()
                                                '*********************
                                            End If
                                
                                            If PersonMoney > 0 Then
                                                PageBillCnt = PageBillCnt + BillCnt
                                                PagePoint = PagePoint + PersonPoint
                                                PointTotal = PointTotal + PersonPoint
                                                BillCntTotal = BillCntTotal + BillCnt
                                
                                                PersonTotal = PersonTotal + 1
                                                If MinMoney > PersonMoney And PersonMoney > 0 Then
                                                    MinMoney = PersonMoney
                                                End If
                                                If MaxMoney < PersonMoney Then
                                                    MaxMoney = PersonMoney
                                                End If

                                                Response.Write("<tr>")
                                                Response.Write("<td style=""height: 33px"" ALIGN=""center"">" & rdU("UnitID") & "</td>")
                                                Response.Write("<td ALIGN=""center"">")
                                                Dim strJob2 As String = "select * from code where typeid=4 and id=" & Trim(JobIDtmp)
                                                Dim CmdJob2 As New Data.OracleClient.OracleCommand(strJob2, conn)
                                                Dim rdJob2 As Data.OracleClient.OracleDataReader = CmdJob2.ExecuteReader()
                                                If rdJob2.HasRows Then
                                                    rdJob2.Read()
                                                    Response.Write(rdJob2("Content"))
                                                Else
                                                    Response.Write("&nbsp;")
                                                End If
                                                rdJob2.Close()
                                                Response.Write("</td>")
                                                Response.Write("<td ALIGN=""center"">")
                                                If MemChName <> "" Then
                                                    Response.Write(MemChName)
                                                Else
                                                    Response.Write("&nbsp;")
                                                End If
                                                Response.Write("</td>")
                                                If sys_City = "台中縣" Or sys_City = "台中市" Then
                                                    Response.Write("<td ALIGN=""center"">")
                                                    Response.Write(BillCnt)
                                                    Response.Write("</td>")
                                                    Response.Write("<td ALIGN=""center"">")
                                                    Response.Write(PersonPoint)
                                                    Response.Write("</td>")
                                                End If
                                                Response.Write("<td ALIGN=""center"">" & PersonMoney & "</td>")
                                                If sys_City <> "台中縣" And sys_City <> "台中市" Then
                                                    Response.Write("<td ALIGN=""center"">")
                                                    'If sys_City = "台中縣" Then
                                                    '    If rdUnitPerson("CreditID") IsNot DBNull.Value Then
                                                    '        Response.Write(rdUnitPerson("CreditID"))
                                                    '    Else
                                                    '        Response.Write("&nbsp;")
                                                    '    End If
                                                    'Else
                                                    Response.Write("&nbsp;")
                                                    'End If
                                                    Response.Write("</td>")
                                                End If
                                                Response.Write("<td ALIGN=""center"">")
                                                If sys_City = "台中縣" Or sys_City = "台中市" Then
                                                    Response.Write(strBank)
                                                Else
                                                    Response.Write("&nbsp;")
                                                End If
                                                Response.Write("</td>")
                                                Response.Write("<td ALIGN=""center"">&nbsp;</td>")
                                                Response.Write("</tr>")
                                            Else
                                                i = i - 1
                                            End If
                                        End If
                                    Next
                                
                                    PageNo = PageNo + 1
                                    Response.Write("<tr>")
                                    If sys_City = "台中縣" Or sys_City = "台中市" Then
                                        Response.Write("<td align=""center"" height=""35"">小計</td>")
                                        Response.Write("<td colspan=""2"">&nbsp;</td>")
                                        Response.Write("<td align=""center"">" & PageBillCnt & "</td>")
                                        PageBillCnt = 0
                                        Response.Write("<td align=""center"">" & PagePoint & "</td>")
                                        PagePoint = 0
                                        Response.Write("<td align=""center"">")
                                        Response.Write(PageSum)
                                        PageSum = 0
                                        Response.Write("</td>")
                                        Response.Write("<td colspan=""3"">&nbsp;</td>")

                                    Else
                                        Response.Write("<td align=""center"" height=""35"">小計</td>")
                                        Response.Write("<td colspan=""2"">&nbsp;</td>")
                                        Response.Write("<td align=""center"">")
                                        Response.Write(PageSum)
                                        PageSum = 0
                                        Response.Write("</td>")
                                        Response.Write("<td colspan=""4"">&nbsp;</td>")
                                        
                                    End If
                                    Response.Write("</tr>")
                                    Response.Write("</table>")
                                    Response.Write("<table width=""680"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                                    Response.Write("<tr>")
                                    Response.Write("</table>")
                                    
                                End If
                                
                                If (UnitCnt = UnitCntTmp - 1 Or UnitCnt = UnitCntTmp) And BillCntTotal > 0 Then
                                    Response.Write("<table width=""680"" border=""1"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                                    Response.Write("<tr>")
                                    If sys_City = "台中縣" Or sys_City = "台中市" Then
                                        
                                        Response.Write("<td style=""width: 32%;height:35px"" ALIGN=""center"">總計</td>")
                                        Response.Write("<td style=""width: 10%"" ALIGN=""center"">" & BillCntTotal & "</td>")
                                        Response.Write("<td style=""width: 10%"" ALIGN=""center"">" & PointTotal & "</td>")
                                        Response.Write("<td style=""width: 49%"" ALIGN=""center"">實領總金額&nbsp;" & MoneyTotal & "</td>")
                                    End If
                                    Response.Write("</tr></table>")
                                    Response.Write("</br>")
                                    If sys_City = "台中縣" Or sys_City = "台中市" Then
                                        Response.Write("<table width=""680"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                                        Response.Write("<tr><td style=""width: 20%""><span class=""style1""><strong>承辦單位：</strong></span></td>")
                                        Response.Write("<td style=""width: 20%""><span class=""style1""><strong>人事：</strong></span></td>")
                                        Response.Write("<td style=""width: 20%""><span class=""style1""><strong>出納：</strong></span></td>")
                                        Response.Write("<td style=""width: 20%""><span class=""style1""><strong>會計：</strong></span></td>")
                                        Response.Write("<td style=""width: 20%""><span class=""style1""><strong>機關主官：</strong></span></td>")
                                        Response.Write("</tr></table>")
                                    Else
                                        Response.Write("<table width=""680"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                                        Response.Write("<tr><td style=""width: 25%""><span class=""style1""><strong>承辦單位：</strong></span></td>")
                                        Response.Write("<td style=""width: 25%""><span class=""style1""><strong>秘書室：</strong></span></td>")
                                        Response.Write("<td style=""width: 25%""><span class=""style1""><strong>會計室：</strong></span></td>")
                                        Response.Write("<td style=""width: 25%""><span class=""style1""><strong>局長：</strong></span></td>")
                                        Response.Write("</tr></table>")
                                    End If
                                End If
                            End While
                        End If
                        rdUnitPerson.Close()
    
                    End While
                End If
                rdU.Close()
                
                '****儲存該月份獎勵金分配統計表********
                If Trim(Request("SaveFlag")) = "1" Then
                    Dim allMoney As Integer = 0
                    Dim strAllMoney As String = "select value from Apconfigure where ID=46"
                    Dim CmdAllMoney As New Data.OracleClient.OracleCommand(strAllMoney, conn)
                    Dim rdAllMoney As Data.OracleClient.OracleDataReader = CmdAllMoney.ExecuteReader()
                    If rdAllMoney.HasRows Then
                        rdAllMoney.Read()
                        allMoney = Trim(rdAllMoney("value"))
                    End If
                    rdAllMoney.Close()
                    
                    Dim strchk As String = "select * from RewardAnalyze where BeginDate=TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS') and EndDate=TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim Cmdchk As New Data.OracleClient.OracleCommand(strchk, conn)
                    Dim rdchk As Data.OracleClient.OracleDataReader = Cmdchk.ExecuteReader()
                    If rdchk.HasRows Then
                        rdchk.Read()
                        Dim cmdUpd As New Data.OracleClient.OracleCommand()
                        Dim strSql = "Update RewardAnalyze set RewardTotal=" & allMoney
                        strSql = strSql + " ,PeopleCount=" & PersonTotal & ",MaxMoney=" & MaxMoney & ",MinMoney=" & MinMoney
                        strSql = strSql + " where BeginDate=TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS') and EndDate=TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        cmdUpd.CommandText = strSql
                        cmdUpd.Connection = conn
                        cmdUpd.ExecuteNonQuery()
                    
                    Else
                        Dim cmdUpd As New Data.OracleClient.OracleCommand()
                        Dim strSql = "insert into RewardAnalyze(BeginDate,EndDate,RewardTotal,PeopleCount,MaxMoney,MinMoney)"
                        strSql = strSql + " values(TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS'),TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        strSql = strSql + "," & allMoney & "," & PersonTotal & "," & MaxMoney & "," & MinMoney
                        strSql = strSql + ")"
                        cmdUpd.CommandText = strSql
                        cmdUpd.Connection = conn
                        cmdUpd.ExecuteNonQuery()
                    End If
                    rdchk.Close()
                    rdchk = Nothing
                
                End If
            
                '**************************************
                '-------------------一人------------------------
            Else
                
                BillCntTotal = 0

                Response.Write("<table width=""680"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                Response.Write("<td align=""center"" colspan=""3""><span class=""style1""><strong>")
                '統計單位
                Dim strUnit As String = "select UnitName from UnitInfo where UnitID=" & Trim(Request("sUnitID")) & ""
                Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
                Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
                If rdUnit.HasRows Then
                    rdUnit.Read()
                        
                    Response.Write(CityUnitName & Trim(rdUnit("UnitName")) & "<br>")
                End If
                rdUnit.Close()
                If theDateType = "BillFillDate" Then
                    Response.Write("計算期間(填單日期)：")
                Else
                    Response.Write("計算期間(建檔日期)：")
                End If
                Response.Write(Year(gOutDT(Trim(Request("Date1")))) - 1911 & "/" & Month(gOutDT(Trim(Request("Date1")))) & "/" & Day(gOutDT(Trim(Request("Date1")))))
                Response.Write(" 至 ")
                Response.Write(Year(gOutDT(Trim(Request("Date2")))) - 1911 & "/" & Month(gOutDT(Trim(Request("Date2")))) & "/" & Day(gOutDT(Trim(Request("Date2")))))
                If sys_City = "台中縣" Then
                    Response.Write("&nbsp; 直接執行人員交通安全獎金請領清冊</strong></span>")
  
                Else
                    Response.Write("&nbsp; 直接執行人員交通績效獎金請領清冊</strong></span>")

                End If
                Response.Write("<br>")
                Response.Write("</td></tr>")
                Response.Write("<tr>")
                Response.Write("<td width=""34%"">")
                Response.Write("列印日期：" & Year(Now) & "/" & Month(Now) & "/" & Day(Now))
                Response.Write("</td>")
                Response.Write("<td width=""33%"">")
                Response.Write("列印人員：" & UserName)
                Response.Write("</td>")
                Response.Write("<td width=""33%"">")
                Response.Write("頁次：" & PageNo)
                Response.Write("</td>")
                Response.Write("</tr>")
                Response.Write("</table>")
                        
                Response.Write("<table width=""680"" border=""1"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                Response.Write("<tr>")
                If sys_City = "台中縣" Or sys_City = "台中市" Then
                    Response.Write("<td style=""width: 8%;height:35px"" ALIGN=""center"">單位</td>")
                    Response.Write("<td style=""width: 12%"" ALIGN=""center"">職位</td>")
                    Response.Write("<td style=""width: 12%"" ALIGN=""center"">姓名</td>")
                    Response.Write("<td style=""width: 10%"" ALIGN=""center"">件數</td>")
                    Response.Write("<td style=""width: 10%"" ALIGN=""center"">點數</td>")
                    Response.Write("<td style=""width: 12%"" ALIGN=""center"">實領獎金</td>")
                    Response.Write("<td style=""width: 25%"" ALIGN=""center"">郵局帳號</td>")
                    Response.Write("<td style=""width: 12%"" ALIGN=""center"">備註</td>")
                Else
                    Response.Write("<td style=""width: 8%;height:35px"" ALIGN=""center"">單位</td>")
                    Response.Write("<td style=""width: 15%"" ALIGN=""center"">職位</td>")
                    Response.Write("<td style=""width: 15%"" ALIGN=""center"">姓名</td>")
                    Response.Write("<td style=""width: 12%"" ALIGN=""center"">實領獎金</td>")
                    Response.Write("<td style=""width: 13%"" ALIGN=""center"">身份證號</td>")
                    Response.Write("<td style=""width: 25%"" ALIGN=""center"">郵局帳號</td>")
                    Response.Write("<td style=""width: 12%"" ALIGN=""center"">備註</td>")
                End If
                Response.Write("</tr>")
                
                Dim LoginIDTmp, ChNameTmp, strMemID1 As String
                Dim MemMoney As Decimal
                MemIDList = ""
                strUnitPerson = "select MemberID,LoginID,CHName,Money,CreditID,JobID,BankName,BankID,BankAccount from MemberData where UnitID=" & Trim(Request("sUnitID")) & " and MemberID=" & Trim(Request("sMemID")) & " " & JobIDPlus & " order by LoginID"
                Dim CmdUnitPerson As New Data.OracleClient.OracleCommand(strUnitPerson, conn)
                Dim rdUnitPerson As Data.OracleClient.OracleDataReader = CmdUnitPerson.ExecuteReader()
                If rdUnitPerson.HasRows Then
                    While rdUnitPerson.Read()
                        If sys_City = "台中縣" Or sys_City = "南投縣" Or sys_City = "彰化縣" Or sys_City = "台中市" Then
                            MemIDList = GetMemberID(Trim(rdUnitPerson("LoginID")), Trim(Request("sUnitID")))
                        Else
                            MemIDList = Trim(rdUnitPerson("MemberID"))
                        End If
                        LoginIDTmp = Trim(rdUnitPerson("LoginID"))
                        ChNameTmp = Trim(rdUnitPerson("CHName"))
                        strMemID1 = Trim(rdUnitPerson("MemberID"))
                        If rdUnitPerson("CreditID") IsNot DBNull.Value Then
                            strCreditID = Trim(rdUnitPerson("CreditID"))
                        Else
                            strCreditID = ""
                        End If
                        If rdUnitPerson("Money") Is DBNull.Value Then
                            MemMoney = 0
                        Else
                            MemMoney = CDec(rdUnitPerson("Money"))
                        End If
                        JobIDtmp = Trim(rdUnitPerson("JobID"))
                        
                        If rdUnitPerson("BankName") IsNot DBNull.Value Then
                            strBank = rdUnitPerson("BankName")
                        Else
                            strBank = "&nbsp;"
                        End If
                        If rdUnitPerson("BankID") IsNot DBNull.Value Then
                            strBank = strBank & rdUnitPerson("BankID")
                        Else
                            strBank = strBank
                        End If
                        If rdUnitPerson("BankAccount") IsNot DBNull.Value Then
                            strBank = strBank & rdUnitPerson("BankAccount")
                        Else
                            strBank = strBank
                        End If
                    End While
                End If
                rdUnitPerson.Close()
                
                If MemIDList <> "" Then
                    OverFlag = ""
                    PersonPoint = 0
                    BillCnt = 0
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
                    Dim strPoint2 As String = "select sum(b.BillType2Score) as cnt,count(*) as BillCnt from BillBase a"
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
                    Dim strPoint3 As String = "select count(*) cnt from BillBase a"
                    strPoint3 = strPoint3 & " where a.BillMemID1 in (" & MemIDList & ") "
                    strPoint3 = strPoint3 & " and a.RecordStateID=0 and (a.TrafficAccidentType='1') and a.BillTypeID='1'"
                    strPoint3 = strPoint3 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint3 = strPoint3 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint3 As New Data.OracleClient.OracleCommand(strPoint3, conn)
                    Dim rdPoint3 As Data.OracleClient.OracleDataReader = CmdPoint3.ExecuteReader()
                    If rdPoint3.HasRows Then
                        rdPoint3.Read()
                        PersonPoint = PersonPoint + (CDec(rdPoint3("cnt")) * 100)
                        BillCnt = BillCnt + CDec(rdPoint3("cnt"))
                            
                    End If
                    rdPoint3.Close()
                    
                    'A2點數
                    Dim strPoint4 As String = "select count(*) cnt from BillBase a"
                    strPoint4 = strPoint4 & " where a.BillMemID1 in (" & MemIDList & ") "
                    strPoint4 = strPoint4 & " and a.RecordStateID=0 and (a.TrafficAccidentType='2') and a.BillTypeID='1'"
                    strPoint4 = strPoint4 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint4 = strPoint4 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint4 As New Data.OracleClient.OracleCommand(strPoint4, conn)
                    Dim rdPoint4 As Data.OracleClient.OracleDataReader = CmdPoint4.ExecuteReader()
                    If rdPoint4.HasRows Then
                        rdPoint4.Read()
                        PersonPoint = PersonPoint + (CDec(rdPoint4("cnt")) * 50)
                        BillCnt = BillCnt + CDec(rdPoint4("cnt"))
                            
                    End If
                    rdPoint4.Close()
                    
                    'A3點數
                    Dim strPoint5 As String = "select count(*) cnt from BillBase a"
                    strPoint5 = strPoint5 & " where a.BillMemID1 in (" & MemIDList & ") "
                    strPoint5 = strPoint5 & " and a.RecordStateID=0 and (a.TrafficAccidentType='3') and a.BillTypeID='1'"
                    strPoint5 = strPoint5 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint5 = strPoint5 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint5 As New Data.OracleClient.OracleCommand(strPoint5, conn)
                    Dim rdPoint5 As Data.OracleClient.OracleDataReader = CmdPoint5.ExecuteReader()
                    If rdPoint5.HasRows Then
                        rdPoint5.Read()
                        PersonPoint = PersonPoint + (CDec(rdPoint5("cnt")) * 20)
                        BillCnt = BillCnt + CDec(rdPoint5("cnt"))
                            
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

                    If MemMoney = 0 Then
                        PersonMoney = Decimal.Round(PointMoney * PersonPoint)
                    Else
                        PersonPay = Decimal.Truncate(MemMoney * getPayPercent)
                        If Decimal.Round(PointMoney * PersonPoint) > PersonPay Then
                            OverFlag = "＊"
                            PersonMoney = PersonPay
                        Else
                            PersonMoney = Decimal.Round(PointMoney * PersonPoint)
                        End If
                    End If
                    
                    Dim strLMoney1 As String = "select * from RewardSpecMember where LoginID='" & LoginIDTmp & "' and ROWNUM <=1 and YearMonth='" & RewardMonth & "'"
                    Dim CmdLMoney1 As New Data.OracleClient.OracleCommand(strLMoney1, conn)
                    Dim rdLMoney1 As Data.OracleClient.OracleDataReader = CmdLMoney1.ExecuteReader()
                    If rdLMoney1.HasRows Then
                        rdLMoney1.Read()
                        PersonMoney = Decimal.Truncate(PersonMoney * (rdLMoney1("DePercent") / 100))
                    End If
                    rdLMoney1.Close()
                    
                    MoneyTotal = MoneyTotal + PersonMoney

                    If PersonMoney > 0 Then
                        PointTotal = PointTotal + PersonPoint
                        BillCntTotal = BillCntTotal + BillCnt

                        '****每月應領實領****
                        strDel = "delete from RewardMonthData where DirectOrTogether='1'"
                        strDel = strDel & " and YearMonth=TO_DATE('" & gOutDT(Trim(Request("Date1"))) & "','YYYY/MM/DD')"
                        strDel = strDel & " and UnitId=" & Trim(Request("sUnitID")) & " and LoginId='" & LoginIDTmp & "'"
                        strDel = strDel & " and MemberId=" & strMemID1

                        Dim cmdDel As New Data.OracleClient.OracleCommand()
                        cmdDel.CommandText = strDel
                        cmdDel.Connection = conn
                        cmdDel.ExecuteNonQuery()
                        
                        strInsert = "insert into RewardMonthData(DirectOrTogether,YearMonth,UnitId,LoginId,ChName"
                        strInsert = strInsert & ",MemberId,CreditID,ShouldGetMoney,RealGetMoney,RecordDate,RecordMemberID)"
                        strInsert = strInsert & " values('1',TO_DATE('" & gOutDT(Trim(Request("Date1"))) & "','YYYY/MM/DD')"
                        strInsert = strInsert & "," & Trim(Request("sUnitID")) & ",'" & LoginIDTmp & "'"
                        strInsert = strInsert & ",'" & ChNameTmp & "'," & strMemID1 & ",'" & strCreditID & "'," & Decimal.Round(PointMoney * PersonPoint)
                        strInsert = strInsert & "," & PersonMoney & ",sysdate," & UserID
                        strInsert = strInsert & ")"
                        
                        Dim cmdInsert As New Data.OracleClient.OracleCommand()
                        cmdInsert.CommandText = strInsert
                        cmdInsert.Connection = conn
                        cmdInsert.ExecuteNonQuery()
                        '*********************
                        PageSum = PageSum + PersonMoney
                        Response.Write("<tr>")
                        Response.Write("<td style=""height: 33px"" ALIGN=""center"">" & Replace(Trim(Request("sUnitID")), "'", "") & "</td>")
                        Response.Write("<td ALIGN=""center"">")
                        Dim strJob2 As String = "select * from code where typeid=4 and id=" & Trim(JobIDtmp)
                        Dim CmdJob2 As New Data.OracleClient.OracleCommand(strJob2, conn)
                        Dim rdJob2 As Data.OracleClient.OracleDataReader = CmdJob2.ExecuteReader()
                        If rdJob2.HasRows Then
                            rdJob2.Read()
                            Response.Write(rdJob2("Content"))
                        Else
                            Response.Write("&nbsp;")
                        End If
                        rdJob2.Close()
                        Response.Write("</td>")
                        Response.Write("<td ALIGN=""center"">")
                        If ChNameTmp <> "" Then
                            Response.Write(ChNameTmp)
                        Else
                            Response.Write("&nbsp;")
                        End If
                        Response.Write("</td>")
                        If sys_City = "台中縣" Or sys_City = "台中市" Then
                            Response.Write("<td ALIGN=""center"">")
                            Response.Write(BillCnt)
                            Response.Write("</td>")
                            Response.Write("<td ALIGN=""center"">")
                            Response.Write(PersonPoint)
                            Response.Write("</td>")
                        End If
                        Response.Write("<td ALIGN=""center"">" & PersonMoney & "</td>")
                        If sys_City <> "台中縣" And sys_City <> "台中市" Then
                            Response.Write("<td ALIGN=""center"">")
                            'If sys_City = "台中縣" Then
                            '    If rdUnitPerson("CreditID") IsNot DBNull.Value Then
                            '        Response.Write(rdUnitPerson("CreditID"))
                            '    Else
                            '        Response.Write("&nbsp;")
                            '    End If
                            'Else
                            Response.Write("&nbsp;")
                            'End If
                            Response.Write("</td>")
                        End If
                        Response.Write("<td ALIGN=""center"">")
                        If sys_City = "台中縣" Or sys_City = "台中市" Then
                            Response.Write(strBank)
                        Else
                            Response.Write("&nbsp;")
                        End If
                        Response.Write("</td>")
                        Response.Write("<td ALIGN=""center"">&nbsp;</td>")
                        Response.Write("</tr>")

                    End If

                End If
                rdUnitPerson.Close()
                
                Response.Write("<tr>")
                If sys_City = "台中縣" Or sys_City = "台中市" Then
                    Response.Write("<td align=""center"" height=""35"">小計</td>")
                    Response.Write("<td colspan=""2"">&nbsp;</td>")
                    Response.Write("<td align=""center"">" & BillCntTotal & "</td>")
                    PageBillCnt = 0
                    Response.Write("<td align=""center"">" & PointTotal & "</td>")
                    PagePoint = 0
                    Response.Write("<td align=""center"">")
                    Response.Write(PageSum)
                    PageSum = 0
                    Response.Write("</td>")
                    Response.Write("<td colspan=""2"">&nbsp;</td>")

                Else
                    Response.Write("<td align=""center"" height=""35"">小計</td>")
                    Response.Write("<td colspan=""2"">&nbsp;</td>")
                    Response.Write("<td align=""center"">")
                    Response.Write(PageSum)
                    PageSum = 0
                    Response.Write("</td>")
                    Response.Write("<td colspan=""3"">&nbsp;</td>")
                                        
                End If
                Response.Write("</tr>")
                Response.Write("</table>")
                Response.Write("<table width=""680"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                Response.Write("<tr>")
                Response.Write("</table>")
                        
                Response.Write("<table width=""680"" border=""0"" cellpadding=""3"" cellspacing=""0"" align=""center"">")
                Response.Write("<tr><td style=""width: 25%""><span class=""style1""><strong>承辦單位：</strong></span></td>")
                Response.Write("<td style=""width: 25%""><span class=""style1""><strong>秘書室：</strong></span></td>")
                Response.Write("<td style=""width: 25%""><span class=""style1""><strong>會計室：</strong></span></td>")
                Response.Write("<td style=""width: 25%""><span class=""style1""><strong>局長：</strong></span></td>")
                Response.Write("</tr></table>")
            End If
            
            conn.Close()
        %>
    
        
    </form>
</body>
<script type="text/javascript" src="../form.js"></script>
<script language="JavaScript">
	printWindow(true,5.08,5.08,5.08,5.08);
</script>
</html>