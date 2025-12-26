<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%  
    LoginCheck()
    Server.ScriptTimeout = 86400
    Response.Flush()
%>
<object id="factory" style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://10.104.10.246/traffic/smsx.cab#Version=6,1,432,1">
</object>
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
    <table width="680" border="0" cellpadding="3" cellspacing="0" align="center">
        <tr>
        <td align="center" colspan="4"><span class="style1"><strong>
        <%
            Dim theDateType As String = Trim(Request("DateType"))
            Response.Write("交通安全任務直接執行人員支領獎勵金核發清冊(單位別)</strong></span>")
            Response.Write("<br>")
            If theDateType = "BillFillDate" Then
                Response.Write("計算期間(填單日期)：")
            Else
                Response.Write("計算期間(建檔日期)：")
            End If
            Response.Write(Year(gOutDT(Trim(Request("Date1")))) - 1911 & "/" & Month(gOutDT(Trim(Request("Date1")))) & "/" & Day(gOutDT(Trim(Request("Date1")))))
            Response.Write(" 至 ")
            Response.Write(Year(gOutDT(Trim(Request("Date2")))) - 1911 & "/" & Month(gOutDT(Trim(Request("Date2")))) & "/" & Day(gOutDT(Trim(Request("Date2")))))
        %></strong></span></td>
        </tr>
        <tr>
        <td style="width: 40%">單位名稱</td>
        <td style="width: 30%" align="center">舉發件數</td>
        <td style="width: 30%" align="center">點數</td>
        </tr>
    </table>
    <hr size="3" />
    <table width="680" border="0" cellpadding="3" cellspacing="0" align="center">
        <%
            '取得 Web.config 檔的資料連接設定
            Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
            '建立 Connection 物件
            Dim conn As New Data.OracleClient.OracleConnection()
            conn.ConnectionString = setting.ConnectionString
            '開啟資料連接
            conn.Open()
            
            '要用填單或建檔日統計
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
            
            Dim UserCookie As HttpCookie = Request.Cookies("RewardUser")
            Dim AnalyzeUnitID2 As String = ""
            AnalyzeUnitID2 = Trim(UserCookie.Values("UnitID"))
            
            Dim sys_City As String = ""
            Dim strCity = "select Value from ApConfigure where ID=31"
            Dim CmdCity As New Data.OracleClient.OracleCommand(strCity, conn)
            Dim rdCity As Data.OracleClient.OracleDataReader = CmdCity.ExecuteReader()
            If rdCity.HasRows Then
                rdCity.Read()
                sys_City = Trim(rdCity("Value"))
            End If
            rdCity.Close()
            '===================列出清冊內容========================
            '---------------所有單位-----------------
            Dim PersonPoint, BillCnt, BillCntTotal As Decimal
            Dim OverFlag As String
            Dim JobIDPlus As String
            If sys_City = "台中縣" Or sys_City = "台中市" Then
                JobIDPlus = ""
            Else
                JobIDPlus = ""
            End If
            PointTotal = 0
            MoneyTotal = 0
            BillCntTotal = 0
            Dim strUnit = "select * from UnitInfo where UnitID in (" & Trim(Request("sUnitID")) & ") order by UnitID"
            Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
            Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
            If rdUnit.HasRows Then
                While rdUnit.Read()
                    Response.Write("<tr>")
                    Response.Write("<td style=""width: 40%"">" & Trim(rdUnit("UnitID")) & "&nbsp; " & Trim(rdUnit("UnitName")) & "</td>")
                    '抓出此單位所有人來算,只抓直接人員
                    UnitPoint = 0
                    UnitMoney = 0
                    BillCnt = 0
                    
                    OverFlag = ""
                    PersonPoint = 0
                    MemMoney = 0
                    '攔停點數
                    Dim strPoint1 As String = "select sum(b.BillType1Score) as cnt,count(*) as BillCnt from BillBaseViewReward a"
                    strPoint1 = strPoint1 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                    strPoint1 = strPoint1 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                    strPoint1 = strPoint1 & " where a.BillMemID1 in (select memberid from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "' " & JobIDPlus & ") and a.RuleVer=b.LawVersion"
                    strPoint1 = strPoint1 & " and (a.Rule1=b.LawItem or a.Rule2=b.LawItem or a.Rule3=b.LawItem or a.Rule4=b.LawItem)"
                    strPoint1 = strPoint1 & " and a.BillTypeID='1' and a.RecordStateID=0 and (a.TrafficAccidentType is null or a.TrafficAccidentType='')"
                    strPoint1 = strPoint1 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                    strPoint1 = strPoint1 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                    Dim CmdPoint1 As New Data.OracleClient.OracleCommand(strPoint1, conn)
                    Dim rdPoint1 As Data.OracleClient.OracleDataReader = CmdPoint1.ExecuteReader()
                    If rdPoint1.HasRows Then
                        rdPoint1.Read()
                        If rdPoint1("cnt") Is DBNull.Value Then
                            PersonPoint = PersonPoint + 0
                            BillCnt = BillCnt + 0
                        Else
                            PersonPoint = PersonPoint + CDec(rdPoint1("cnt"))
                            BillCnt = BillCnt + CDec(rdPoint1("BillCnt"))
                        End If
                    End If
                    rdPoint1.Close()
                    
                    '逕舉點數
                    Dim strPoint2 As String = "select sum(b.BillType2Score) as cnt,count(*) as BillCnt from BillBase a"
                    strPoint2 = strPoint2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                    strPoint2 = strPoint2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                    strPoint2 = strPoint2 & " where a.BillMemID1 in (select memberid from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "' " & JobIDPlus & ") and a.RuleVer=b.LawVersion"
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
                            BillCnt = BillCnt + 0
                        Else
                            PersonPoint = PersonPoint + CDec(rdPoint2("cnt"))
                            BillCnt = BillCnt + CDec(rdPoint2("BillCnt"))
                        End If
                    End If
                    rdPoint2.Close()
                    
                    'A1點數
                    Dim strPoint3 As String = "select b.A1Score,b.BillType1Score from BillBase a"
                    strPoint3 = strPoint3 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                    strPoint3 = strPoint3 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                    strPoint3 = strPoint3 & " where a.BillMemID1 in (select memberid from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "' " & JobIDPlus & ") and a.RuleVer=b.LawVersion"
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
                    Dim strPoint4 As String = "select b.A2Score,b.BillType1Score from BillBase a"
                    strPoint4 = strPoint4 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                    strPoint4 = strPoint4 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                    strPoint4 = strPoint4 & " where a.BillMemID1 in (select memberid from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "' " & JobIDPlus & ") and a.RuleVer=b.LawVersion"
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
                    Dim strPoint5 As String = "select b.A3Score,b.BillType1Score from BillBase a"
                    strPoint5 = strPoint5 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                    strPoint5 = strPoint5 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("sCountyOrNpa")) & ") b"
                    strPoint5 = strPoint5 & " where a.BillMemID1 in (select memberid from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "' " & JobIDPlus & ") and a.RuleVer=b.LawVersion"
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
                        strPointT1a = strPointT1a & " where a.BillMemID1 in (select memberid from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "' " & JobIDPlus & ") and a.RecordStateID=0"
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
                        strPointT1b = strPointT1b & " where a.BillMemID1 in (select memberid from MemberData where UnitID='" & Trim(rdUnit("UnitID")) & "' " & JobIDPlus & ") and a.RecordStateID=0"
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
                    
                    UnitMoney = UnitMoney + MemMoney
                    
                    PointTotal = PointTotal + PersonPoint
                    BillCntTotal = BillCntTotal + BillCnt
                    Response.Write("<td style=""width: 20%""  align=""right"">" & Format(BillCnt, "##,##0.00") & "</td>")
                    Response.Write("<td style=""width: 28%""  align=""right"">" & Format(PersonPoint, "##,##0.00") & "</td>")
                    Response.Write("<td style=""width: 12%""  align=""center""></td>")
                    Response.Write("</tr>")
                End While
            End If
            rdUnit.Close()
            conn.Close()
        %>
    </table>
        <hr size="3" />
    <table width="680" border="0" cellpadding="3" cellspacing="0" align="center">
        <tr>
        <td style="width: 40%" align="center"></td>
        <td style="width: 20%" align="right"><%=Format(BillCntTotal, "##,##0.00")%></td>
        <td style="width: 28%" align="right"><%=Format(PointTotal, "##,##0.00")%></td>
        <td style="width: 12%" ></td>
        </tr>
    </table>
        <br />
        <br />
        <br />
        <br />
        
    <table width="680" border="0" cellpadding="3" cellspacing="0" align="center">
        <tr>
        <td style="width: 25%"><span class="style1"><strong>製表</strong></span></td>
        <td style="width: 25%"><span class="style1"><strong>組長</strong></span></td>
        <td style="width: 25%"><span class="style1"><strong>副隊長</strong></span></td>
        <td style="width: 25%"><span class="style1"><strong>隊長</strong></span></td>
        </tr>
    </table>
    </form>
</body>
<script type="text/javascript" src="../form.js"></script>
<script language="JavaScript">
	printWindow(true,5.08,5.08,5.08,5.08);
</script>
</html>

