<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%  
    LoginCheck()
%>
<%

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
<object id="factory" style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://localhost/traffic/smsx.cab#Version=6,1,432,1">
</object>
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
<script language="JavaScript">
	window.focus();
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>直接人員每點點數金額</title>
</head>
<body style="text-align: center">
    <form id="form1" runat="server">
        <span style="font-size: 18pt">
        <% 
            '取得 Web.config 檔的資料連接設定
            Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
            '建立 Connection 物件
            Dim conn As New Data.OracleClient.OracleConnection()
            conn.ConnectionString = setting.ConnectionString
            '開啟資料連接
            conn.Open()
                
            Dim sys_Title As String = ""
            Dim strTitle = "select Value from ApConfigure where ID=40"
            Dim CmdTitle As New Data.OracleClient.OracleCommand(strTitle, conn)
            Dim rdTitle As Data.OracleClient.OracleDataReader = CmdTitle.ExecuteReader()
            If rdTitle.HasRows Then
                rdTitle.Read()
                sys_Title = Trim(rdTitle("Value"))
            End If
            rdTitle.Close()
            
            Response.Write(sys_Title)
        %></span>
         <br />
         <br />
        <span style="font-size: 16pt"><strong>
        <%
            Response.Write(Trim(Request("Date1")))
        %> 年 <%
            Response.Write(Trim(Request("Date2")))
        %> 月 直接人員每點點數金額(應領)<br />
            <br />
            <br />
        </strong>總金額 &nbsp; &nbsp;&nbsp; * &nbsp;&nbsp; 分配比率 &nbsp; &nbsp;/ &nbsp; &nbsp;總點數
            &nbsp; &nbsp;= &nbsp; &nbsp;每點金額
            </span>
            <br />
            <span style="font-size: 12pt">
        <%
            '要用填單或建檔日統計
            Dim AnalyzeDate1, AnalyzeDate2 As String
            Dim theDateType As String = Trim(Request("DateType"))
            AnalyzeDate1 = (Trim(Request("Date1")) + 1911) & "/" & Trim(Request("Date2")) & "/1"
            AnalyzeDate2 = DateAdd("d", -1, DateAdd("m", 1, (Trim(Request("Date1")) + 1911) & "/" & Trim(Request("Date2")) & "/1"))
            'Response.Write(AnalyzeDate1 & ",,,,,," & AnalyzeDate2)
            'Response.End()
            '================================================
            '獎勵金總額
            Dim getMoneyTotal As Decimal
            If Trim(Request("AllAnalyzeMoney")) = "" Then
                getMoneyTotal = 0
            Else
                getMoneyTotal = CDec(Request("AllAnalyzeMoney"))
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
            Dim UserID As String
            UserID = Trim(UserCookie.Values("MemberID"))
            
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
            strPointT1 = strPointT1 & " and a." & theDateType & " between TO_DATE('" & AnalyzeDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPointT1 = strPointT1 & " and TO_DATE('" & AnalyzeDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
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
            strPointT2 = strPointT2 & " and a." & theDateType & " between TO_DATE('" & AnalyzeDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPointT2 = strPointT2 & " and TO_DATE('" & AnalyzeDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
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
            strPointT3 = strPointT3 & " and a." & theDateType & " between TO_DATE('" & AnalyzeDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPointT3 = strPointT3 & " and TO_DATE('" & AnalyzeDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
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
            strPointT4 = strPointT4 & " and a." & theDateType & " between TO_DATE('" & AnalyzeDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPointT4 = strPointT4 & " and TO_DATE('" & AnalyzeDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
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
            strPointT5 = strPointT5 & " and a." & theDateType & " between TO_DATE('" & AnalyzeDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strPointT5 = strPointT5 & " and TO_DATE('" & AnalyzeDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
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
                strPointT1a = strPointT1a & " and a." & theDateType & " between TO_DATE('" & AnalyzeDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPointT1a = strPointT1a & " and TO_DATE('" & AnalyzeDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
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
                strPointT1b = strPointT1b & " and a." & theDateType & " between TO_DATE('" & AnalyzeDate1 & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                strPointT1b = strPointT1b & " and TO_DATE('" & AnalyzeDate2 & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
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
            Response.Write(Format(getMoneyTotal, "##,##0.00") & "&nbsp; &nbsp; &nbsp;*")
            Response.Write("&nbsp; &nbsp; &nbsp;0.72&nbsp; &nbsp; &nbsp;/")
            Response.Write("&nbsp; &nbsp; &nbsp;" & Format(getPointTotal, "##,##0.00"))
            Response.Write("&nbsp; &nbsp; &nbsp;=&nbsp; &nbsp; &nbsp;" & Format(PointMoney, "##,##0.000000000000"))
            conn.Close()

        %>
        </span>

    </form>
</body>
<script type="text/javascript" src="../form.js"></script>
<script language="JavaScript">
	printWindow(true,5.08,5.08,5.08,5.08);
</script>
</html>
