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
    body {font-family:新細明體;font-size:10pt; }

    .style1 {font-family:新細明體; font-size: 11pt}
    -->
    </style>
    <style media=print>
    .Noprint{display:none;}
    .PageNext{page-break-after: always;}
    </style>
    <title>單位舉發件數暨績分統計明細表</title>
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
    
    Dim strMem, strLaw, strScore, strScore2, strScoreA, strScoreB, strIll, strDeatil As String
    Dim ReportValueCnt, ReportValueScore, ReportTotalCnt, TotalScore As Decimal
    Dim StopValueCnt, StopValueScore, StopTotalCnt As Decimal
    
    sys_City = ""
    Dim strCity = "select Value from ApConfigure where ID=31"
    Dim CmdCity As New Data.OracleClient.OracleCommand(strCity, conn)
    Dim rdCity As Data.OracleClient.OracleDataReader = CmdCity.ExecuteReader()
    If rdCity.HasRows Then
        rdCity.Read()
        sys_City = Trim(rdCity("Value"))
    End If
    rdCity.Close()
    
    strMem = "select UnitID,UnitName from UnitInfo where UnitID in (" & Trim(Request("sUnitID")) & ")"
    strMem = strMem & " order by UnitID"

    Dim CmdMem As New Data.OracleClient.OracleCommand(strMem, conn)
    Dim rdMem As Data.OracleClient.OracleDataReader = CmdMem.ExecuteReader()
    If rdMem.HasRows Then
        While rdMem.Read()
            strDeatil = ""
            ReportTotalCnt = 0
            StopTotalCnt = 0
            TotalScore = 0

            '全部法條 or 1~68條 or 68條以後
            Dim strLawRange As String = ""
            If Trim(Request("LawRange")) = "0" Then
                strLawRange = ""
            ElseIf Trim(Request("LawRange")) = "1" Then
                strLawRange = " and substr(a.ItemID,1,2) between '1' and '68'"
            ElseIf Trim(Request("LawRange")) = "2" Then
                strLawRange = " and substr(a.ItemID,1,2) > '68'"
            End If
            
            strLaw = "select distinct(ItemID) from Law a,BillBaseViewReward b where (a.ItemID=b.Rule1 or a.ItemID=b.Rule2 or a.ItemID=b.Rule3 or a.ItemID=b.Rule4)"
            strLaw = strLaw & " and b.BillUnitID='" & Trim(rdMem("UnitID")) & "' and b.RecordstateID=0 " & strLawRange
            strLaw = strLaw & " and b." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strLaw = strLaw & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS') order by ItemID"
            Dim CmdLaw As New Data.OracleClient.OracleCommand(strLaw, conn)
            Dim rdLaw As Data.OracleClient.OracleDataReader = CmdLaw.ExecuteReader()
            If rdLaw.HasRows Then
                While rdLaw.Read()
                        
                    StopValueCnt = 0
                    StopValueScore = 0
                    ReportValueCnt = 0
                    ReportValueScore = 0
                    strDeatil = strDeatil & "<tr>"
                    strDeatil = strDeatil & "<td width=""10%"">" & Trim(rdLaw("ItemID")) & "</td>"
                    '雲林拖吊案件分數較高
                    If sys_City = "雲林縣" Then
                        '抓逕舉件數及績分
                        strScore = "select count(*) as cnt,sum(b.BillType2Score) as ScoreSum from BillBaseViewReward a "
                        strScore = strScore & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                        strScore = strScore & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                        strScore = strScore & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                        strScore = strScore & " and a.BillUnitID='" & Trim(rdMem("UnitID")) & "' and a.BillBaseTypeID='0' and a.BillTypeID='2'"
                        strScore = strScore & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule1)"
                        strScore = strScore & " and (a.CarAddID<>8 or a.CarAddID is null) and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                        strScore = strScore & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strScore = strScore & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdScore As New Data.OracleClient.OracleCommand(strScore, conn)
                        Dim rdScore As Data.OracleClient.OracleDataReader = CmdScore.ExecuteReader()
                        If rdScore.HasRows Then
                            rdScore.Read()
                            ReportValueCnt = rdScore("cnt")
                            ReportTotalCnt = ReportTotalCnt + rdScore("cnt")
                            If (rdScore("ScoreSum")) Is DBNull.Value Then
                                ReportValueScore = 0
                                TotalScore = TotalScore
                            Else
                                ReportValueScore = rdScore("ScoreSum")
                                TotalScore = TotalScore + rdScore("ScoreSum")
                            End If
                        End If
                        rdScore.Close()
                    
                        '抓逕舉件數及績分(拖吊)
                        strScoreA = "select count(*) as cnt,sum(b.Other1) as ScoreSum from BillBaseViewReward a "
                        strScoreA = strScoreA & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,Other1,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                        strScoreA = strScoreA & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                        strScoreA = strScoreA & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                        strScoreA = strScoreA & " and a.BillUnitID='" & Trim(rdMem("UnitID")) & "' and a.BillBaseTypeID='0' and a.BillTypeID='2'"
                        strScoreA = strScoreA & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule1)"
                        strScoreA = strScoreA & " and a.CarAddID=8 and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                        strScoreA = strScoreA & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strScoreA = strScoreA & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdScoreA As New Data.OracleClient.OracleCommand(strScoreA, conn)
                        Dim rdScoreA As Data.OracleClient.OracleDataReader = CmdScoreA.ExecuteReader()
                        If rdScoreA.HasRows Then
                            rdScoreA.Read()
                            ReportValueCnt = ReportValueCnt + rdScoreA("cnt")
                            ReportTotalCnt = ReportTotalCnt + rdScoreA("cnt")
                            If (rdScoreA("ScoreSum")) Is DBNull.Value Then
                                ReportValueScore = ReportValueScore + 0
                                TotalScore = TotalScore
                            Else
                                ReportValueScore = ReportValueScore + rdScoreA("ScoreSum")
                                TotalScore = TotalScore + rdScoreA("ScoreSum")
                            End If
                        End If
                        rdScoreA.Close()
                        
                        '抓攔停件數及績分
                        strScore2 = "select count(*) as cnt,sum(b.BillType1Score) as ScoreSum from BillBaseViewReward a "
                        strScore2 = strScore2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                        strScore2 = strScore2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                        strScore2 = strScore2 & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                        strScore2 = strScore2 & " and a.BillUnitID='" & Trim(rdMem("UnitID")) & "' and ((a.BillBaseTypeID='0' and a.BillTypeID='1') or (a.BillBaseTypeID='1'))"
                        strScore2 = strScore2 & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule1)"
                        strScore2 = strScore2 & " and (a.CarAddID<>8 or a.CarAddID is null) and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                        strScore2 = strScore2 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strScore2 = strScore2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdScore2 As New Data.OracleClient.OracleCommand(strScore2, conn)
                        Dim rdScore2 As Data.OracleClient.OracleDataReader = CmdScore2.ExecuteReader()
                        If rdScore2.HasRows Then
                            rdScore2.Read()
                            StopValueCnt = rdScore2("cnt")
                            StopTotalCnt = StopTotalCnt + rdScore2("cnt")
                            If (rdScore2("ScoreSum")) Is DBNull.Value Then
                                StopValueScore = 0
                                TotalScore = TotalScore
                            Else
                                StopValueScore = rdScore2("ScoreSum")
                                TotalScore = TotalScore + rdScore2("ScoreSum")
                            End If
                        End If
                        rdScore2.Close()
                        
                        '抓攔停件數及績分(拖吊)
                        strScoreB = "select count(*) as cnt,sum(b.Other1) as ScoreSum from BillBaseViewReward a "
                        strScoreB = strScoreB & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,Other1,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                        strScoreB = strScoreB & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                        strScoreB = strScoreB & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                        strScoreB = strScoreB & " and a.BillUnitID='" & Trim(rdMem("UnitID")) & "' and ((a.BillBaseTypeID='0' and a.BillTypeID='1') or (a.BillBaseTypeID='1'))"
                        strScoreB = strScoreB & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule1)"
                        strScoreB = strScoreB & " and a.CarAddID=8 and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                        strScoreB = strScoreB & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strScoreB = strScoreB & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdScoreB As New Data.OracleClient.OracleCommand(strScoreB, conn)
                        Dim rdScoreB As Data.OracleClient.OracleDataReader = CmdScoreB.ExecuteReader()
                        If rdScoreB.HasRows Then
                            rdScoreB.Read()
                            StopValueCnt = StopValueCnt + rdScoreB("cnt")
                            StopTotalCnt = StopTotalCnt + rdScoreB("cnt")
                            If (rdScoreB("ScoreSum")) Is DBNull.Value Then
                                StopValueScore = StopValueScore + 0
                                TotalScore = TotalScore
                            Else
                                StopValueScore = StopValueScore + rdScoreB("ScoreSum")
                                TotalScore = TotalScore + rdScoreB("ScoreSum")
                            End If
                        End If
                        rdScoreB.Close()
                    Else
                        '抓逕舉件數及績分
                        strScore = "select count(*) as cnt,sum(b.BillType2Score) as ScoreSum from BillBaseViewReward a "
                        strScore = strScore & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                        strScore = strScore & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                        strScore = strScore & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                        strScore = strScore & " and a.BillUnitID='" & Trim(rdMem("UnitID")) & "' and a.BillBaseTypeID='0' and a.BillTypeID='2'"
                        strScore = strScore & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule1)"
                        strScore = strScore & " and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                        strScore = strScore & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strScore = strScore & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdScore As New Data.OracleClient.OracleCommand(strScore, conn)
                        Dim rdScore As Data.OracleClient.OracleDataReader = CmdScore.ExecuteReader()
                        If rdScore.HasRows Then
                            rdScore.Read()
                            ReportValueCnt = rdScore("cnt")
                            ReportTotalCnt = ReportTotalCnt + rdScore("cnt")
                            If (rdScore("ScoreSum")) Is DBNull.Value Then
                                ReportValueScore = 0
                                TotalScore = TotalScore
                            Else
                                ReportValueScore = rdScore("ScoreSum")
                                TotalScore = TotalScore + rdScore("ScoreSum")
                            End If
                        End If
                        rdScore.Close()
                    
                        '抓攔停件數及績分
                        strScore2 = "select count(*) as cnt,sum(b.BillType1Score) as ScoreSum from BillBaseViewReward a "
                        strScore2 = strScore2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                        strScore2 = strScore2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                        strScore2 = strScore2 & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                        strScore2 = strScore2 & " and a.BillUnitID='" & Trim(rdMem("UnitID")) & "' and ((a.BillBaseTypeID='0' and a.BillTypeID='1') or (a.BillBaseTypeID='1'))"
                        strScore2 = strScore2 & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule1)"
                        strScore2 = strScore2 & " and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                        strScore2 = strScore2 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strScore2 = strScore2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdScore2 As New Data.OracleClient.OracleCommand(strScore2, conn)
                        Dim rdScore2 As Data.OracleClient.OracleDataReader = CmdScore2.ExecuteReader()
                        If rdScore2.HasRows Then
                            rdScore2.Read()
                            StopValueCnt = rdScore2("cnt")
                            StopTotalCnt = StopTotalCnt + rdScore2("cnt")
                            If (rdScore2("ScoreSum")) Is DBNull.Value Then
                                StopValueScore = 0
                                TotalScore = TotalScore
                            Else
                                StopValueScore = rdScore2("ScoreSum")
                                TotalScore = TotalScore + rdScore2("ScoreSum")
                            End If
                        End If
                        rdScore2.Close()
                    End If
                    '抓法條內容
                    strDeatil = strDeatil & "<td width=""60%"">"
                    strIll = "select IllegalRule from Law where ItemID='" & Trim(rdLaw("ItemID")) & "'"
                    Dim CmdIll As New Data.OracleClient.OracleCommand(strIll, conn)
                    Dim rdIll As Data.OracleClient.OracleDataReader = CmdIll.ExecuteReader()
                    If rdIll.HasRows Then
                        rdIll.Read()
                        
                        strDeatil = strDeatil & Trim(rdIll("IllegalRule"))
                    End If
                    rdIll.Close()
                    strDeatil = strDeatil & "</td>"
                    strDeatil = strDeatil & "<td width=""10%"" align=""right"">" & Format(ReportValueCnt, "##,##0") & "</td>"
                    strDeatil = strDeatil & "<td width=""10%"" align=""right"">" & Format(StopValueCnt, "##,##0") & "</td>"
                    strDeatil = strDeatil & "<td width=""10%"" align=""right"">" & Format(ReportValueScore + StopValueScore, "##,##0.#") & "</td>"
                    strDeatil = strDeatil & "</tr>"
                
                        
                End While
            End If
            rdLaw.Close()
            If ReportTotalCnt + StopTotalCnt > 0 Then
                Response.Write("<table width=""680"" border=""0"" cellpadding=""1"" cellspacing=""0"" align=""center"">")
                Response.Write("<tr><td align=""center"" height=""35""><strong><span class=""style1"">單 位 舉 發 件 數 暨 績 分 統 計 明 細 表</span></strong></td></tr>")
                Response.Write("<tr><td>舉發單位：&nbsp;" & rdMem("UnitID") & "&nbsp;" & rdMem("UnitName") & "</td></tr>")
                Response.Write("<tr><td>查詢期間：&nbsp;" & Request("Date1") & " ~ " & Request("Date2") & "</td></tr>")
                Response.Write("</table>")
                Response.Write("<hr>")
                Response.Write("<table width=""680"" border=""0"" cellpadding=""1"" cellspacing=""0"" align=""center"">")
                Response.Write("<tr>")
                Response.Write("<td width=""10%"">違規法條</td>")
                Response.Write("<td width=""60%"">違規事實</td>")
                Response.Write("<td width=""10%"" align=""right"">逕舉件數</td>")
                Response.Write("<td width=""10%"" align=""right"">攔停件數</td>")
                Response.Write("<td width=""10%"" align=""right"">舉發績分</td>")
                Response.Write("</tr>")
                Response.Write("</table>")
                Response.Write("<hr>")
                Response.Write("<table width=""680"" border=""0"" cellpadding=""1"" cellspacing=""0"" align=""center"">")
            
                Response.Write(strDeatil)
            
                Response.Write("</table>")
                Response.Write("<hr>")
                Response.Write("<table width=""680"" border=""0"" cellpadding=""1"" cellspacing=""0"" align=""center"">")
                Response.Write("<tr>")
                Response.Write("<td width=""15%"">總舉發件數</td>")
                Response.Write("<td width=""55%""></td>")
                Response.Write("<td width=""10%"" align=""right"">" & Format(ReportTotalCnt, "##,##0") & "</td>")
                Response.Write("<td width=""10%"" align=""right"">" & Format(StopTotalCnt, "##,##0") & "</td>")
                Response.Write("<td width=""10%"" align=""right"">" & Format(TotalScore, "##,##0.#") & "</td>")
                Response.Write("</tr>")
                Response.Write("</table>")
            
                Response.Write("<div class=""PageNext""></div>")
            End If
        End While
    End If
    rdMem.Close()
    
    conn.Close()
%>
    </form>
</body>
<script type="text/javascript" src="../form.js"></script>
<script language="JavaScript">
	printWindow(true,5.08,5.08,5.08,5.08);
</script>
</html>
