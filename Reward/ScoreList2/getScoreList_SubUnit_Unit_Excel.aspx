<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%  
    LoginCheck()
    Dim fMnoth = Month(Now)
    If fMnoth < 10 Then fMnoth = "0" & fMnoth
    Dim fDay = Day(Now)
    If fDay < 10 Then fDay = "0" & fDay
    Dim fname = Year(Now) & fMnoth & fDay & ".xls"
    Response.AddHeader("Content-Disposition", "filename=" & fname)
    'If Trim(Request("sMemID")) = "" Then
    'Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8")
    'End If
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

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <style type="text/css">
    <!--
    body {font-family:新細明體;font-size:10pt; }

    .style1 {font-family:新細明體; font-size: 11pt}
    -->
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
    
    Dim strMem, strLaw, strScore, strScore2, strScoreA, strScoreB, strIll, strDeatil, strPointT6Plus2, strPointT6Plus As String
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
    
    If sys_City = "台東縣" Then
        strPointT6Plus = " and ((a.CarSimpleID in (3,4) and b.CarSimpleID=3) or(a.CarSimpleID in (1,6) and b.CarSimpleID=5) or(a.CarSimpleID=2 and b.CarSimpleID=6))"
        strPointT6Plus2 = ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed,CarSimpleID from LawScore"

    Else
        strPointT6Plus = ""
        strPointT6Plus2 = ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
    End If
    
    strMem = "select UnitID,UnitName from UnitInfo where UnitID in (" & Trim(Request("sUnitID")) & ")"
    strMem = strMem & " order by UnitID"
    Response.Write("<table width=""660"" border=""1"" cellpadding=""1"" cellspacing=""0"" align=""center"">")

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
            strLaw = strLaw & " and (b.BillMemID1 in (select MemberID from MemberData where UnitID='" & Trim(rdMem("UnitID")) & "')"
            strLaw = strLaw & " or b.BillMemID2 in (select MemberID from MemberData where UnitID='" & Trim(rdMem("UnitID")) & "')"
            strLaw = strLaw & " or b.BillMemID3 in (select MemberID from MemberData where UnitID='" & Trim(rdMem("UnitID")) & "')"
            strLaw = strLaw & " or b.BillMemID4 in (select MemberID from MemberData where UnitID='" & Trim(rdMem("UnitID")) & "')"
            strLaw = strLaw & ")"
            strLaw = strLaw & " and b.RecordstateID=0 " & strLawRange
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
                    strDeatil = strDeatil & "<td width=""10%"" align=""left"">" & Trim(rdLaw("ItemID")) & "</td>"
                    
                    Dim strPer = "select MemberID,Money from MemberData where UnitID='" & Trim(rdMem("UnitID")) & "'"
                    Dim CmdPer As New Data.OracleClient.OracleCommand(strPer, conn)
                    Dim rdPer As Data.OracleClient.OracleDataReader = CmdPer.ExecuteReader()
                    If rdPer.HasRows Then
                        While rdPer.Read()
                            '雲林拖吊案件分數較高
                            If sys_City = "雲林縣" Or sys_City = "台東縣" Or sys_City = "宜蘭縣" Then
                                '抓逕舉件數及績分
                                strScore = "select b.BillType2Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBaseViewReward a "
                                strScore = strScore & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                                strScore = strScore & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                                strScore = strScore & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                                strScore = strScore & " and (a.BillMemID1='" & Trim(rdPer("MemberID")) & "'"
                                strScore = strScore & " or a.BillMemID2='" & Trim(rdPer("MemberID")) & "'"
                                strScore = strScore & " or a.BillMemID3='" & Trim(rdPer("MemberID")) & "'"
                                strScore = strScore & " or a.BillMemID4='" & Trim(rdPer("MemberID")) & "'"
                                strScore = strScore & ")"
                                strScore = strScore & " and a.BillBaseTypeID='0' and a.BillTypeID='2'"
                                strScore = strScore & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule4)"
                                strScore = strScore & " and (a.CarAddID<>8 or a.CarAddID is null) and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                                strScore = strScore & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strScore = strScore & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdScore As New Data.OracleClient.OracleCommand(strScore, conn)
                                Dim rdScore As Data.OracleClient.OracleDataReader = CmdScore.ExecuteReader()
                                If rdScore.HasRows Then
                                    While rdScore.Read()
                                        If rdScore("BillType2Score") Is DBNull.Value Then
                                            ReportValueScore = ReportValueScore + 0
                                            ReportTotalCnt = ReportTotalCnt + 0
                                            TotalScore = TotalScore
                                            ReportValueCnt = ReportValueCnt
                                        Else
                                            If rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") Is DBNull.Value And rdScore("BillMemID3") Is DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score"))
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score"))
                                                ReportTotalCnt = ReportTotalCnt + 1
                                                ReportValueCnt = ReportValueCnt + 1
                                            ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") Is DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdScore("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 2)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 2)
                                                    ReportTotalCnt = ReportTotalCnt + 0.5
                                                    ReportValueCnt = ReportValueCnt + 0.5
                                                End If
                                                If Trim(rdScore("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 2)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 2)
                                                    ReportTotalCnt = ReportTotalCnt + 0.5
                                                    ReportValueCnt = ReportValueCnt + 0.5
                                                End If
                                            ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") IsNot DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdScore("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") * 0.34)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") * 0.34)
                                                    ReportValueCnt = ReportValueCnt + 0.34
                                                    ReportTotalCnt = ReportTotalCnt + 0.34
                                                End If
                                                If Trim(rdScore("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") * 0.33)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") * 0.33)
                                                    ReportValueCnt = ReportValueCnt + 0.33
                                                    ReportTotalCnt = ReportTotalCnt + 0.33
                                                End If
                                                If Trim(rdScore("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") * 0.33)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") * 0.33)
                                                    ReportValueCnt = ReportValueCnt + 0.33
                                                    ReportTotalCnt = ReportTotalCnt + 0.33
                                                End If
                                            ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") IsNot DBNull.Value And rdScore("BillMemID4") IsNot DBNull.Value Then
                                                If Trim(rdPer("MemberID")) = Trim(rdScore("BillMemID1")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                    ReportValueCnt = ReportValueCnt + 0.25
                                                    ReportTotalCnt = ReportTotalCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScore("BillMemID2")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                    ReportValueCnt = ReportValueCnt + 0.25
                                                    ReportTotalCnt = ReportTotalCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScore("BillMemID3")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                    ReportValueCnt = ReportValueCnt + 0.25
                                                    ReportTotalCnt = ReportTotalCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScore("BillMemID4")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                    ReportValueCnt = ReportValueCnt + 0.25
                                                    ReportTotalCnt = ReportTotalCnt + 0.25
                                                End If
                                            End If
                                        End If
                                    End While
                                End If
                                rdScore.Close()

                                '抓逕舉件數及績分(拖吊)
                                strScoreA = "select b.Other1,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBase a " & strPointT6Plus2
                                strScoreA = strScoreA & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                                strScoreA = strScoreA & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                                strScoreA = strScoreA & " and (a.BillMemID1='" & Trim(rdPer("MemberID")) & "'"
                                strScoreA = strScoreA & " or a.BillMemID2='" & Trim(rdPer("MemberID")) & "'"
                                strScoreA = strScoreA & " or a.BillMemID3='" & Trim(rdPer("MemberID")) & "'"
                                strScoreA = strScoreA & " or a.BillMemID4='" & Trim(rdPer("MemberID")) & "'"
                                strScoreA = strScoreA & ")"
                                strScoreA = strScoreA & " and a.BillBaseTypeID='0' and a.BillTypeID='2'"
                                strScoreA = strScoreA & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule4)"
                                strScoreA = strScoreA & " and a.CarAddID=8 and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                                strScoreA = strScoreA & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strScoreA = strScoreA & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')" & strPointT6Plus
                                Dim CmdScoreA As New Data.OracleClient.OracleCommand(strScoreA, conn)
                                Dim rdScoreA As Data.OracleClient.OracleDataReader = CmdScoreA.ExecuteReader()
                                If rdScoreA.HasRows Then
                                    While rdScoreA.Read()
                                        If rdScoreA("BillType2Score") Is DBNull.Value Then
                                            ReportValueScore = ReportValueScore + 0
                                            ReportTotalCnt = ReportTotalCnt + 0
                                            TotalScore = TotalScore
                                            ReportValueCnt = ReportValueCnt
                                        Else
                                            If rdScoreA("BillMemID1") IsNot DBNull.Value And rdScoreA("BillMemID2") Is DBNull.Value And rdScoreA("BillMemID3") Is DBNull.Value And rdScoreA("BillMemID4") Is DBNull.Value Then
                                                ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1"))
                                                TotalScore = TotalScore + CDec(rdScoreA("Other1"))
                                                ReportTotalCnt = ReportTotalCnt + 1
                                                ReportValueCnt = ReportValueCnt + 1
                                            ElseIf rdScoreA("BillMemID1") IsNot DBNull.Value And rdScoreA("BillMemID2") IsNot DBNull.Value And rdScoreA("BillMemID3") Is DBNull.Value And rdScoreA("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdScoreA("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") / 2)
                                                    TotalScore = TotalScore + CDec(rdScoreA("Other1") / 2)
                                                    ReportTotalCnt = ReportTotalCnt + 0.5
                                                    ReportValueCnt = ReportValueCnt + 0.5
                                                End If
                                                If Trim(rdScoreA("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") / 2)
                                                    TotalScore = TotalScore + CDec(rdScoreA("Other1") / 2)
                                                    ReportTotalCnt = ReportTotalCnt + 0.5
                                                    ReportValueCnt = ReportValueCnt + 0.5
                                                End If
                                            ElseIf rdScoreA("BillMemID1") IsNot DBNull.Value And rdScoreA("BillMemID2") IsNot DBNull.Value And rdScoreA("BillMemID3") IsNot DBNull.Value And rdScoreA("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdScoreA("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") * 0.34)
                                                    TotalScore = TotalScore + CDec(rdScoreA("Other1") * 0.34)
                                                    ReportTotalCnt = ReportTotalCnt + 0.34
                                                    ReportValueCnt = ReportValueCnt + 0.34
                                                End If
                                                If Trim(rdScoreA("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") * 0.33)
                                                    TotalScore = TotalScore + CDec(rdScoreA("Other1") * 0.33)
                                                    ReportTotalCnt = ReportTotalCnt + 0.33
                                                    ReportValueCnt = ReportValueCnt + 0.33
                                                End If
                                                If Trim(rdScoreA("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") * 0.33)
                                                    TotalScore = TotalScore + CDec(rdScoreA("Other1") * 0.33)
                                                    ReportTotalCnt = ReportTotalCnt + 0.33
                                                    ReportValueCnt = ReportValueCnt + 0.33
                                                End If
                                                
                                            ElseIf rdScoreA("BillMemID1") IsNot DBNull.Value And rdScoreA("BillMemID2") IsNot DBNull.Value And rdScoreA("BillMemID3") IsNot DBNull.Value And rdScoreA("BillMemID4") IsNot DBNull.Value Then
                                                If Trim(rdPer("MemberID")) = Trim(rdScoreA("BillMemID1")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") / 4)
                                                    TotalScore = TotalScore + CDec(rdScoreA("Other1") / 4)
                                                    ReportTotalCnt = ReportTotalCnt + 0.25
                                                    ReportValueCnt = ReportValueCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScoreA("BillMemID2")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") / 4)
                                                    TotalScore = TotalScore + CDec(rdScoreA("Other1") / 4)
                                                    ReportTotalCnt = ReportTotalCnt + 0.25
                                                    ReportValueCnt = ReportValueCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScoreA("BillMemID3")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") / 4)
                                                    TotalScore = TotalScore + CDec(rdScoreA("Other1") / 4)
                                                    ReportTotalCnt = ReportTotalCnt + 0.25
                                                    ReportValueCnt = ReportValueCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScoreA("BillMemID4")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") / 4)
                                                    TotalScore = TotalScore + CDec(rdScoreA("Other1") / 4)
                                                    ReportTotalCnt = ReportTotalCnt + 0.25
                                                    ReportValueCnt = ReportValueCnt + 0.25
                                                End If
                                            End If
                                        End If
                                    End While
                                End If
                                rdScoreA.Close()

                                '抓攔停件數及績分
                                strScore2 = "select b.BillType1Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBaseViewReward a "
                                strScore2 = strScore2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                                strScore2 = strScore2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                                strScore2 = strScore2 & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                                strScore2 = strScore2 & " and (a.BillMemID1='" & Trim(rdPer("MemberID")) & "'"
                                strScore2 = strScore2 & " or a.BillMemID2='" & Trim(rdPer("MemberID")) & "'"
                                strScore2 = strScore2 & " or a.BillMemID3='" & Trim(rdPer("MemberID")) & "'"
                                strScore2 = strScore2 & " or a.BillMemID4='" & Trim(rdPer("MemberID")) & "'"
                                strScore2 = strScore2 & ")"
                                strScore2 = strScore2 & " and ((a.BillBaseTypeID='0' and a.BillTypeID='1') or (a.BillBaseTypeID='1'))"
                                strScore2 = strScore2 & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule4)"
                                strScore2 = strScore2 & " and (a.CarAddID<>8 or a.CarAddID is null) and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                                strScore2 = strScore2 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strScore2 = strScore2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdScore2 As New Data.OracleClient.OracleCommand(strScore2, conn)
                                Dim rdScore2 As Data.OracleClient.OracleDataReader = CmdScore2.ExecuteReader()
                                If rdScore2.HasRows Then
                                    While rdScore2.Read()
                                        If rdScore2("BillType1Score") Is DBNull.Value Then
                                            StopValueScore = StopValueScore + 0
                                            StopTotalCnt = StopTotalCnt + 0
                                            TotalScore = TotalScore
                                            StopValueCnt = StopValueCnt
                                        Else
                                            If rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") Is DBNull.Value And rdScore2("BillMemID3") Is DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score"))
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score"))
                                                StopTotalCnt = StopTotalCnt + 1
                                                StopValueCnt = StopValueCnt + 1
                                            ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") Is DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdScore2("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 2)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 2)
                                                    StopTotalCnt = StopTotalCnt + 0.5
                                                    StopValueCnt = StopValueCnt + 0.5
                                                End If
                                                If Trim(rdScore2("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 2)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 2)
                                                    StopTotalCnt = StopTotalCnt + 0.5
                                                    StopValueCnt = StopValueCnt + 0.5
                                                End If
                                            ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") IsNot DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdScore2("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") * 0.34)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") * 0.34)
                                                    StopTotalCnt = StopTotalCnt + 0.34
                                                    StopValueCnt = StopValueCnt + 0.34
                                                End If
                                                If Trim(rdScore2("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                    StopTotalCnt = StopTotalCnt + 0.33
                                                    StopValueCnt = StopValueCnt + 0.33
                                                End If
                                                If Trim(rdScore2("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                    StopTotalCnt = StopTotalCnt + 0.33
                                                    StopValueCnt = StopValueCnt + 0.33
                                                End If
                                            ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") IsNot DBNull.Value And rdScore2("BillMemID4") IsNot DBNull.Value Then
                                                If Trim(rdPer("MemberID")) = Trim(rdScore2("BillMemID1")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                    StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                    StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScore2("BillMemID2")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                    StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                    StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScore2("BillMemID3")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                    StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                    StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScore2("BillMemID4")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                    StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                    StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                                End If
                                            End If
                                            
                                        End If
                                    End While
                                End If
                                rdScore2.Close()
                        
                                '抓攔停件數及績分(拖吊)
                                strScoreB = "select b.BillType1Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBase a " & strPointT6Plus2
                                strScoreB = strScoreB & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                                strScoreB = strScoreB & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                                strScoreB = strScoreB & " and (a.BillMemID1='" & Trim(rdPer("MemberID")) & "'"
                                strScoreB = strScoreB & " or a.BillMemID2='" & Trim(rdPer("MemberID")) & "'"
                                strScoreB = strScoreB & " or a.BillMemID3='" & Trim(rdPer("MemberID")) & "'"
                                strScoreB = strScoreB & " or a.BillMemID4='" & Trim(rdPer("MemberID")) & "'"
                                strScoreB = strScoreB & ")"
                                strScoreB = strScoreB & " and ((a.BillBaseTypeID='0' and a.BillTypeID='1') or (a.BillBaseTypeID='1'))"
                                strScoreB = strScoreB & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule4)"
                                strScoreB = strScoreB & " and a.CarAddID=8 and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                                strScoreB = strScoreB & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strScoreB = strScoreB & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')" & strPointT6Plus
                                Dim CmdScoreB As New Data.OracleClient.OracleCommand(strScoreB, conn)
                                Dim rdScoreB As Data.OracleClient.OracleDataReader = CmdScoreB.ExecuteReader()
                                If rdScoreB.HasRows Then
                                    While rdScoreB.Read()
                                        If rdScoreB("BillType1Score") Is DBNull.Value Then
                                            StopValueScore = StopValueScore + 0
                                            StopTotalCnt = StopTotalCnt + 0
                                            TotalScore = TotalScore
                                            StopValueCnt = StopValueCnt
                                        Else
                                            If rdScoreB("BillMemID1") IsNot DBNull.Value And rdScoreB("BillMemID2") Is DBNull.Value And rdScoreB("BillMemID3") Is DBNull.Value And rdScoreB("BillMemID4") Is DBNull.Value Then
                                                StopValueScore = StopValueScore + CDec(rdScoreB("Other1"))
                                                TotalScore = TotalScore + CDec(rdScoreB("Other1"))
                                                StopTotalCnt = StopTotalCnt + 1
                                                StopValueCnt = StopValueCnt + 1
                                            ElseIf rdScoreB("BillMemID1") IsNot DBNull.Value And rdScoreB("BillMemID2") IsNot DBNull.Value And rdScoreB("BillMemID3") Is DBNull.Value And rdScoreB("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdScoreB("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScoreB("Other1") / 2)
                                                    TotalScore = TotalScore + CDec(rdScoreB("Other1") / 2)
                                                    StopTotalCnt = StopTotalCnt + 0.5
                                                    StopValueCnt = StopValueCnt + 0.5
                                                End If
                                                If Trim(rdScoreB("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScoreB("Other1") / 2)
                                                    TotalScore = TotalScore + CDec(rdScoreB("Other1") / 2)
                                                    StopTotalCnt = StopTotalCnt + 0.5
                                                    StopValueCnt = StopValueCnt + 0.5
                                                End If
                                            ElseIf rdScoreB("BillMemID1") IsNot DBNull.Value And rdScoreB("BillMemID2") IsNot DBNull.Value And rdScoreB("BillMemID3") IsNot DBNull.Value And rdScoreB("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdScoreB("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScoreB("Other1") * 0.34)
                                                    TotalScore = TotalScore + CDec(rdScoreB("Other1") * 0.34)
                                                    StopTotalCnt = StopTotalCnt + 0.34
                                                    StopValueCnt = StopValueCnt + 0.34
                                                End If
                                                If Trim(rdScoreB("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScoreB("Other1") * 0.33)
                                                    TotalScore = TotalScore + CDec(rdScoreB("Other1") * 0.33)
                                                    StopTotalCnt = StopTotalCnt + 0.33
                                                    StopValueCnt = StopValueCnt + 0.33
                                                End If
                                                If Trim(rdScoreB("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScoreB("Other1") * 0.33)
                                                    TotalScore = TotalScore + CDec(rdScoreB("Other1") * 0.33)
                                                    StopTotalCnt = StopTotalCnt + 0.33
                                                    StopValueCnt = StopValueCnt + 0.33
                                                End If
                                            ElseIf rdScoreB("BillMemID1") IsNot DBNull.Value And rdScoreB("BillMemID2") IsNot DBNull.Value And rdScoreB("BillMemID3") IsNot DBNull.Value And rdScoreB("BillMemID4") IsNot DBNull.Value Then
                                                If Trim(rdPer("MemberID")) = Trim(rdScoreB("BillMemID1")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScoreB("Other1") / 4)
                                                    TotalScore = TotalScore + CDec(rdScoreB("Other1") / 4)
                                                    StopTotalCnt = StopTotalCnt + 0.25
                                                    StopValueCnt = StopValueCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScoreB("BillMemID2")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScoreB("Other1") / 4)
                                                    TotalScore = TotalScore + CDec(rdScoreB("Other1") / 4)
                                                    StopTotalCnt = StopTotalCnt + 0.25
                                                    StopValueCnt = StopValueCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScoreB("BillMemID3")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScoreB("Other1") / 4)
                                                    TotalScore = TotalScore + CDec(rdScoreB("Other1") / 4)
                                                    StopTotalCnt = StopTotalCnt + 0.25
                                                    StopValueCnt = StopValueCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScoreB("BillMemID4")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScoreB("Other1") / 4)
                                                    TotalScore = TotalScore + CDec(rdScoreB("Other1") / 4)
                                                    StopTotalCnt = StopTotalCnt + 0.25
                                                    StopValueCnt = StopValueCnt + 0.25
                                                End If
                                            End If
                                        End If
                                    End While
                            
                                End If
                                rdScoreB.Close()
                            Else
                                '抓逕舉件數及績分
                                strScore = "select b.BillType2Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBaseViewReward a "
                                strScore = strScore & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                                strScore = strScore & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                                strScore = strScore & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                                strScore = strScore & " and (a.BillMemID1='" & Trim(rdPer("MemberID")) & "'"
                                strScore = strScore & " or a.BillMemID2='" & Trim(rdPer("MemberID")) & "'"
                                strScore = strScore & " or a.BillMemID3='" & Trim(rdPer("MemberID")) & "'"
                                strScore = strScore & " or a.BillMemID4='" & Trim(rdPer("MemberID")) & "'"
                                strScore = strScore & ")"
                                strScore = strScore & " and a.BillBaseTypeID='0' and a.BillTypeID='2'"
                                strScore = strScore & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule4)"
                                strScore = strScore & " and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                                strScore = strScore & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strScore = strScore & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdScore As New Data.OracleClient.OracleCommand(strScore, conn)
                                Dim rdScore As Data.OracleClient.OracleDataReader = CmdScore.ExecuteReader()
                                If rdScore.HasRows Then
                                    While rdScore.Read()
                                        If rdScore("BillType2Score") Is DBNull.Value Then
                                            ReportValueScore = ReportValueScore + 0
                                            ReportTotalCnt = ReportTotalCnt + 0
                                            TotalScore = TotalScore
                                            ReportValueCnt = ReportValueCnt
                                        Else
                                            If rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") Is DBNull.Value And rdScore("BillMemID3") Is DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score"))
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score"))
                                                ReportTotalCnt = ReportTotalCnt + 1
                                                ReportValueCnt = ReportValueCnt + 1
                                            ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") Is DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdScore("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 2)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 2)
                                                    ReportTotalCnt = ReportTotalCnt + 0.5
                                                    ReportValueCnt = ReportValueCnt + 0.5
                                                End If
                                                If Trim(rdScore("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 2)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 2)
                                                    ReportTotalCnt = ReportTotalCnt + 0.5
                                                    ReportValueCnt = ReportValueCnt + 0.5
                                                End If
                                            ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") IsNot DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdScore("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") * 0.34)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") * 0.34)
                                                    ReportTotalCnt = ReportTotalCnt + 0.34
                                                    ReportValueCnt = ReportValueCnt + 0.34
                                                End If
                                                If Trim(rdScore("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") * 0.33)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") * 0.33)
                                                    ReportTotalCnt = ReportTotalCnt + 0.33
                                                    ReportValueCnt = ReportValueCnt + 0.33
                                                End If
                                                If Trim(rdScore("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") * 0.33)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") * 0.33)
                                                    ReportTotalCnt = ReportTotalCnt + 0.33
                                                    ReportValueCnt = ReportValueCnt + 0.33
                                                End If
                                            ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") IsNot DBNull.Value And rdScore("BillMemID4") IsNot DBNull.Value Then
                                                If Trim(rdPer("MemberID")) = Trim(rdScore("BillMemID1")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                    ReportTotalCnt = ReportTotalCnt + 0.25
                                                    ReportValueCnt = ReportValueCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScore("BillMemID2")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                    ReportTotalCnt = ReportTotalCnt + 0.25
                                                    ReportValueCnt = ReportValueCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScore("BillMemID3")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                    ReportTotalCnt = ReportTotalCnt + 0.25
                                                    ReportValueCnt = ReportValueCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScore("BillMemID4")) Then
                                                    ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                    ReportTotalCnt = ReportTotalCnt + 0.25
                                                    ReportValueCnt = ReportValueCnt + 0.25
                                                End If
                                            End If
                                        End If
                                    End While
                            
                                End If
                                rdScore.Close()
                    
                                '抓攔停件數及績分
                                strScore2 = "select b.BillType1Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBaseViewReward a "
                                strScore2 = strScore2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                                strScore2 = strScore2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                                strScore2 = strScore2 & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                                strScore2 = strScore2 & " and (a.BillMemID1='" & Trim(rdPer("MemberID")) & "'"
                                strScore2 = strScore2 & " or a.BillMemID2='" & Trim(rdPer("MemberID")) & "'"
                                strScore2 = strScore2 & " or a.BillMemID3='" & Trim(rdPer("MemberID")) & "'"
                                strScore2 = strScore2 & " or a.BillMemID4='" & Trim(rdPer("MemberID")) & "'"
                                strScore2 = strScore2 & ")"
                                strScore2 = strScore2 & " and ((a.BillBaseTypeID='0' and a.BillTypeID='1') or (a.BillBaseTypeID='1'))"
                                strScore2 = strScore2 & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule4)"
                                strScore2 = strScore2 & " and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                                strScore2 = strScore2 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                                strScore2 = strScore2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                                Dim CmdScore2 As New Data.OracleClient.OracleCommand(strScore2, conn)
                                Dim rdScore2 As Data.OracleClient.OracleDataReader = CmdScore2.ExecuteReader()
                                If rdScore2.HasRows Then
                                    While rdScore2.Read()
                                        If rdScore2("BillType1Score") Is DBNull.Value Then
                                            StopValueScore = StopValueScore + 0
                                            StopTotalCnt = StopTotalCnt + 0
                                            TotalScore = TotalScore
                                            StopValueCnt = StopValueCnt
                                        Else
                                            If rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") Is DBNull.Value And rdScore2("BillMemID3") Is DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score"))
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score"))
                                                StopTotalCnt = StopTotalCnt + 1
                                                StopValueCnt = StopValueCnt + 1
                                            ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") Is DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdScore2("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 2)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 2)
                                                    StopTotalCnt = StopTotalCnt + 0.5
                                                    StopValueCnt = StopValueCnt + 0.5
                                                End If
                                                If Trim(rdScore2("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 2)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 2)
                                                    StopTotalCnt = StopTotalCnt + 0.5
                                                    StopValueCnt = StopValueCnt + 0.5
                                                End If
                                            ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") IsNot DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                                If Trim(rdScore2("BillMemID1")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") * 0.34)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") * 0.34)
                                                    StopTotalCnt = StopTotalCnt + 0.34
                                                    StopValueCnt = StopValueCnt + 0.34
                                                End If
                                                If Trim(rdScore2("BillMemID2")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                    StopTotalCnt = StopTotalCnt + 0.33
                                                    StopValueCnt = StopValueCnt + 0.33
                                                End If
                                                If Trim(rdScore2("BillMemID3")) = Trim(rdPer("MemberID")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                    StopTotalCnt = StopTotalCnt + 0.33
                                                    StopValueCnt = StopValueCnt + 0.33
                                                End If
                                            ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") IsNot DBNull.Value And rdScore2("BillMemID4") IsNot DBNull.Value Then
                                                If Trim(rdPer("MemberID")) = Trim(rdScore2("BillMemID1")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                    StopTotalCnt = StopTotalCnt + 0.25
                                                    StopValueCnt = StopValueCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScore2("BillMemID2")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                    StopTotalCnt = StopTotalCnt + 0.25
                                                    StopValueCnt = StopValueCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScore2("BillMemID3")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                    StopTotalCnt = StopTotalCnt + 0.25
                                                    StopValueCnt = StopValueCnt + 0.25
                                                End If
                                                If Trim(rdPer("MemberID")) = Trim(rdScore2("BillMemID4")) Then
                                                    StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                    TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                    StopTotalCnt = StopTotalCnt + 0.25
                                                    StopValueCnt = StopValueCnt + 0.25
                                                End If
                                            End If
                                        End If
                                    End While

                                End If
                                rdScore2.Close()
                        
                            End If
                        End While
                    End If
                    rdPer.Close()
                    '抓法條內容
                    strDeatil = strDeatil & "<td width=""64%"">"
                    strIll = "select IllegalRule from Law where ItemID='" & Trim(rdLaw("ItemID")) & "'"
                    Dim CmdIll As New Data.OracleClient.OracleCommand(strIll, conn)
                    Dim rdIll As Data.OracleClient.OracleDataReader = CmdIll.ExecuteReader()
                    If rdIll.HasRows Then
                        rdIll.Read()
                        
                        strDeatil = strDeatil & Trim(rdIll("IllegalRule"))
                    End If
                    rdIll.Close()
                    strDeatil = strDeatil & "</td>"
                    strDeatil = strDeatil & "<td width=""8%"" align=""right"">" & Format(ReportValueCnt, "##,##0.##") & "</td>"
                    strDeatil = strDeatil & "<td width=""8%"" align=""right"">" & Format(StopValueCnt, "##,##0.##") & "</td>"
                    strDeatil = strDeatil & "<td width=""10%"" align=""right"">" & Format(ReportValueScore + StopValueScore, "##,##0.##") & "</td>"
                    strDeatil = strDeatil & "</tr>"
                
                        
                End While
            End If
            rdLaw.Close()
            
            If ReportTotalCnt + StopTotalCnt > 0 Then
                Response.Write("<tr><td align=""center"" height=""35"" colspan=""5""><strong><span class=""style1"">單 位 舉 發 件 數 暨 績 分 統 計 明 細 表</span></strong></td></tr>")
                Response.Write("<tr><td colspan=""5"">舉發單位：&nbsp; " & rdMem("UnitID") & "&nbsp; " & rdMem("UnitName") & "</td></tr>")
                Response.Write("<tr><td colspan=""5"">查詢期間：&nbsp; " & Request("Date1") & " ~ " & Request("Date2") & "</td></tr>")

                Response.Write("<tr>")
                Response.Write("<td width=""10%"">違規法條</td>")
                Response.Write("<td width=""64%"">違規事實</td>")
                Response.Write("<td width=""8%"" align=""right"">逕舉件數</td>")
                Response.Write("<td width=""8%"" align=""right"">攔停件數</td>")
                Response.Write("<td width=""10%"" align=""right"">舉發績分</td>")
                Response.Write("</tr>")
            
                Response.Write(strDeatil)
            
                Response.Write("<tr>")
                Response.Write("<td width=""10%"" colspan=""2"">總舉發件數</td>")
                'Response.Write("<td width=""70%""></td>")
                Response.Write("<td width=""8%"" align=""right"">" & Format(ReportTotalCnt, "##,##0.##") & "</td>")
                Response.Write("<td width=""8%"" align=""right"">" & Format(StopTotalCnt, "##,##0.##") & "</td>")
                Response.Write("<td width=""10%"" align=""right"">" & Format(TotalScore, "##,##0.##") & "</td>")
                Response.Write("</tr>")
            
                'Response.Write("<div class=""PageNext""></div>")
            End If
        End While
    End If
    rdMem.Close()
    Response.Write("</table>")

    conn.Close()
%>
    </form>
</body>
</html>
