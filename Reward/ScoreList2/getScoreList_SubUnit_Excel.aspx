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
    
    Function GetMemberID(ByVal LoginID)
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
        Dim MemString As String = ""
        Dim strMData As String = "Select MemberID from MemberData where LoginID='" & Trim(LoginID) & "'"
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

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <style type="text/css">
    <!--
    body {font-family:新細明體;font-size:10pt; }

    .style1 {font-family:新細明體; font-size: 11pt}
    -->
    </style>
    <title>個別員警舉發件數暨績分統計明細表</title>
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
    
    Dim strMem, strLaw, strScore, strScore2, strScoreA, strScoreB, strIll, strDeatil, MemIDList As String
    Dim ReportValueCnt, ReportValueScore, ReportTotalCnt, TotalScore As Decimal
    Dim StopValueCnt, StopValueScore, StopTotalCnt, A1ScoreTmp As Decimal
    Dim sys_City, strPointT6Plus, strPointT6Plus2 As String
    '抓縣市
    sys_City = ""
    Dim strCity = "select Value from ApConfigure where ID=31"
    Dim CmdCity As New Data.OracleClient.OracleCommand(strCity, conn)
    Dim rdCity As Data.OracleClient.OracleDataReader = CmdCity.ExecuteReader()
    If rdCity.HasRows Then
        rdCity.Read()
        sys_City = Trim(rdCity("Value"))
    End If
    rdCity.Close()
    
    '拖吊點數
    If sys_City = "台東縣" Then
        strPointT6Plus = " and ((a.CarSimpleID in (3,4) and b.CarSimpleID=3) or(a.CarSimpleID in (1,6) and b.CarSimpleID=5) or(a.CarSimpleID=2 and b.CarSimpleID=6))"
        strPointT6Plus2 = ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed,CarSimpleID from LawScore"

    Else
        strPointT6Plus = ""
        strPointT6Plus2 = ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,Other1,CountyOrNpa,IsUsed from LawScore"
    End If
    
    If Trim(Request("MemLoginID")) <> "" Then
        If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Then
            strMem = "select distinct(LoginID) from MemberData where LoginID='" & Trim(Request("MemLoginID")) & "'"
            strMem = strMem & " and AccountStateID<>-1 and RecordStateID<>-1 order by LoginID"
        Else
            strMem = "select LoginID,MemberID,ChName,CreditID,UnitID from MemberData where LoginID='" & Trim(Request("MemLoginID")) & "'"
            strMem = strMem & " and RecordStateID<>-1 order by UnitID,ChName"
        End If
    Else
        If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Then
            strMem = "select distinct(LoginID) from MemberData where UnitID in (" & Trim(Request("sUnitID")) & ")"
            strMem = strMem & " and AccountStateID<>-1 and RecordStateID<>-1 order by LoginID"
        Else
            strMem = "select LoginID,MemberID,ChName,CreditID,UnitID from MemberData where UnitID in (" & Trim(Request("sUnitID")) & ")"
            strMem = strMem & " and RecordStateID<>-1 order by UnitID,ChName"
        End If
    End If
    Response.Write("<table width=""660"" border=""1"" cellpadding=""1"" cellspacing=""0"" align=""center"">")

    Dim CmdMem As New Data.OracleClient.OracleCommand(strMem, conn)
    Dim rdMem As Data.OracleClient.OracleDataReader = CmdMem.ExecuteReader()
    If rdMem.HasRows Then
        While rdMem.Read()
            strDeatil = ""
            ReportTotalCnt = 0
            StopTotalCnt = 0
            TotalScore = 0

            If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Then
                MemIDList = GetMemberID(Trim(rdMem("LoginID")))
            Else
                MemIDList = Trim(rdMem("MemberID"))
            End If
            
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
            strLaw = strLaw & " and (b.BillMemID1 in (" & MemIDList & ") or b.BillMemID2 in (" & MemIDList & ") or b.BillMemID3 in (" & MemIDList & ") or b.BillMemID4 in (" & MemIDList & "))"
            strLaw = strLaw & " and b.RecordstateID=0 " & strLawRange
            strLaw = strLaw & " and b." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
            strLaw = strLaw & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS') order by ItemID"
            Dim CmdLaw As New Data.OracleClient.OracleCommand(strLaw, conn)
            Dim rdLaw As Data.OracleClient.OracleDataReader = CmdLaw.ExecuteReader()
            If rdLaw.HasRows Then
                While rdLaw.Read()
                        
                    ReportValueCnt = 0
                    ReportValueScore = 0
                    StopValueCnt = 0
                    StopValueScore = 0
                    strDeatil = strDeatil & "<tr>"
                    strDeatil = strDeatil & "<td width=""10%"" align=""left"">" & Trim(rdLaw("ItemID")) & "</td>"

                    
                    '雲林拖吊案件分數較高
                    If sys_City = "雲林縣" Or sys_City = "台東縣" Or sys_City = "宜蘭縣" Then
                        '抓逕舉件數及績分
                        strScore = "select b.BillType2Score,b.A1Score,b.A2Score,b.A3Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4,a.TrafficAccidentType from BillBaseViewReward a "
                        strScore = strScore & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                        strScore = strScore & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                        strScore = strScore & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                        strScore = strScore & " and (a.BillMemID1 in (" & MemIDList & ") or a.BillMemID2 in (" & MemIDList & ") or a.BillMemID3 in (" & MemIDList & ") or a.BillMemID4 in (" & MemIDList & "))"
                        strScore = strScore & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule4)" & " and a.BillBaseTypeID='0' and a.BillTypeID='2'"
                        strScore = strScore & " and (a.CarAddID<>8 or a.CarAddID is null) and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                        strScore = strScore & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strScore = strScore & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdScore As New Data.OracleClient.OracleCommand(strScore, conn)
                        Dim rdScore As Data.OracleClient.OracleDataReader = CmdScore.ExecuteReader()
                        If rdScore.HasRows Then
                            While rdScore.Read()
                                If rdScore("TrafficAccidentType") Is DBNull.Value Then
                                    '===================非交通事故======================
                                    If (rdScore("BillType2Score")) Is DBNull.Value Then
                                        ReportValueScore = ReportValueScore
                                        TotalScore = TotalScore
                                    Else
                                        If rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") Is DBNull.Value And rdScore("BillMemID3") Is DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                            ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score"))
                                            TotalScore = TotalScore + CDec(rdScore("BillType2Score"))
                                            ReportValueCnt = ReportValueCnt + 1
                                            ReportTotalCnt = ReportTotalCnt + 1
                                        ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") Is DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore("BillMemID1"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 2)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 2)
                                                ReportValueCnt = ReportValueCnt + CDec(1 / 2)
                                                ReportTotalCnt = ReportTotalCnt + CDec(1 / 2)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID2"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 2)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 2)
                                                ReportValueCnt = ReportValueCnt + CDec(1 / 2)
                                                ReportTotalCnt = ReportTotalCnt + CDec(1 / 2)
                                            End If
                                        ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") IsNot DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore("BillMemID1"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") * 0.34)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") * 0.34)
                                                ReportValueCnt = ReportValueCnt + 0.34
                                                ReportTotalCnt = ReportTotalCnt + 0.34
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID2"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") * 0.33)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") * 0.33)
                                                ReportValueCnt = ReportValueCnt + 0.33
                                                ReportTotalCnt = ReportTotalCnt + 0.33
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID3"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") * 0.33)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") * 0.33)
                                                ReportValueCnt = ReportValueCnt + 0.33
                                                ReportTotalCnt = ReportTotalCnt + 0.33
                                            End If
                                        ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") IsNot DBNull.Value And rdScore("BillMemID4") IsNot DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore("BillMemID1"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID2"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID3"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID4"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                        End If
                                    End If
                                    '==================================================
                                Else
                                    '***************交通事故A1,2,3***********************
                                    If Trim(rdScore("TrafficAccidentType")) = "1" Then
                                        If rdScore("BillType2Score") > rdScore("A1Score") Then
                                            If rdScore("BillType2Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore("BillType2Score"))
                                            End If
                                        Else
                                            If rdScore("A1Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore("A1Score"))
                                            End If
                                        End If
                                    ElseIf Trim(rdScore("TrafficAccidentType")) = "2" Then
                                        If rdScore("BillType2Score") > rdScore("A2Score") Then
                                            If rdScore("BillType2Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore("BillType2Score"))
                                            End If
                                        Else
                                            If rdScore("A2Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore("A2Score"))
                                            End If
                                        End If
                                    ElseIf Trim(rdScore("TrafficAccidentType")) = "3" Then
                                        If rdScore("BillType2Score") > rdScore("A3Score") Then
                                            If rdScore("BillType2Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore("BillType2Score"))
                                            End If
                                        Else
                                            If rdScore("A3Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore("A3Score"))
                                            End If
                                        End If
                                    End If
                                    
                                        
                                    If A1ScoreTmp = 0 Then
                                        ReportValueScore = ReportValueScore
                                        TotalScore = TotalScore
                                    Else
                                        If rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") Is DBNull.Value And rdScore("BillMemID3") Is DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                            ReportValueScore = ReportValueScore + CDec(A1ScoreTmp)
                                            TotalScore = TotalScore + CDec(A1ScoreTmp)
                                            ReportValueCnt = ReportValueCnt + 1
                                            ReportTotalCnt = ReportTotalCnt + 1
                                        ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") Is DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore("BillMemID1"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp / 2)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 2)
                                                ReportValueCnt = ReportValueCnt + CDec(1 / 2)
                                                ReportTotalCnt = ReportTotalCnt + CDec(1 / 2)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID2"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp / 2)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 2)
                                                ReportValueCnt = ReportValueCnt + CDec(1 / 2)
                                                ReportTotalCnt = ReportTotalCnt + CDec(1 / 2)
                                            End If
                                        ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") IsNot DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore("BillMemID1"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp * 0.33)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp * 0.34)
                                                ReportValueCnt = ReportValueCnt + 0.34
                                                ReportTotalCnt = ReportTotalCnt + 0.34
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID2"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp * 0.33)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp * 0.33)
                                                ReportValueCnt = ReportValueCnt + 0.33
                                                ReportTotalCnt = ReportTotalCnt + 0.33
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID3"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp * 0.33)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp * 0.33)
                                                ReportValueCnt = ReportValueCnt + 0.33
                                                ReportTotalCnt = ReportTotalCnt + 0.33
                                            End If
                                        ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") IsNot DBNull.Value And rdScore("BillMemID4") IsNot DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore("BillMemID1"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID2"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID3"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID4"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                        End If
                                    End If
                                    '***********************************************************
                                End If
                            End While
                        End If
                        rdScore.Close()
                      
                        '抓逕舉件數及績分(拖吊)
                        strScoreA = "select b.Other1,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBase a " & strPointT6Plus2
                        strScoreA = strScoreA & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                        strScoreA = strScoreA & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                        strScoreA = strScoreA & " and (a.BillMemID1 in (" & MemIDList & ") or a.BillMemID2 In (" & MemIDList & ") or a.BillMemID3 In (" & MemIDList & ") or a.BillMemID4 In (" & MemIDList & "))"
                        strScoreA = strScoreA & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule4)" & " and a.BillBaseTypeID='0' and a.BillTypeID='2'"
                        strScoreA = strScoreA & " and a.CarAddID=8 and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                        strScoreA = strScoreA & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strScoreA = strScoreA & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')" & strPointT6Plus
                        Dim CmdScoreA As New Data.OracleClient.OracleCommand(strScoreA, conn)
                        Dim rdScoreA As Data.OracleClient.OracleDataReader = CmdScoreA.ExecuteReader()
                        If rdScoreA.HasRows Then
                            While rdScoreA.Read()
                                If (rdScoreA("Other1")) Is DBNull.Value Then
                                    ReportValueScore = ReportValueScore
                                    TotalScore = TotalScore
                                Else
                                    If rdScoreA("BillMemID1") IsNot DBNull.Value And rdScoreA("BillMemID2") Is DBNull.Value And rdScoreA("BillMemID3") Is DBNull.Value And rdScoreA("BillMemID4") Is DBNull.Value Then
                                        ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1"))
                                        TotalScore = TotalScore + CDec(rdScoreA("Other1"))
                                        ReportValueCnt = ReportValueCnt + 1
                                        ReportTotalCnt = ReportTotalCnt + 1
                                    ElseIf rdScoreA("BillMemID1") IsNot DBNull.Value And rdScoreA("BillMemID2") IsNot DBNull.Value And rdScoreA("BillMemID3") Is DBNull.Value And rdScoreA("BillMemID4") Is DBNull.Value Then
                                        If InStr(MemIDList, Trim(rdScoreA("BillMemID1"))) > 0 Then
                                            ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") / 2)
                                            TotalScore = TotalScore + CDec(rdScoreA("Other1") / 2)
                                            ReportValueCnt = ReportValueCnt + CDec(1 / 2)
                                            ReportTotalCnt = ReportTotalCnt + CDec(1 / 2)
                                        End If
                                        If InStr(MemIDList, Trim(rdScoreA("BillMemID2"))) > 0 Then
                                            ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") / 2)
                                            TotalScore = TotalScore + CDec(rdScoreA("Other1") / 2)
                                            ReportValueCnt = ReportValueCnt + CDec(1 / 2)
                                            ReportTotalCnt = ReportTotalCnt + CDec(1 / 2)
                                        End If

                                    ElseIf rdScoreA("BillMemID1") IsNot DBNull.Value And rdScoreA("BillMemID2") IsNot DBNull.Value And rdScoreA("BillMemID3") IsNot DBNull.Value And rdScoreA("BillMemID4") Is DBNull.Value Then
                                        If InStr(MemIDList, Trim(rdScoreA("BillMemID1"))) > 0 Then
                                            ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") * 0.34)
                                            TotalScore = TotalScore + CDec(rdScoreA("Other1") * 0.34)
                                            ReportValueCnt = ReportValueCnt + 0.34
                                            ReportTotalCnt = ReportTotalCnt + 0.34
                                        End If
                                        If InStr(MemIDList, Trim(rdScoreA("BillMemID2"))) > 0 Then
                                            ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") * 0.33)
                                            TotalScore = TotalScore + CDec(rdScoreA("Other1") * 0.33)
                                            ReportValueCnt = ReportValueCnt + 0.33
                                            ReportTotalCnt = ReportTotalCnt + 0.33
                                        End If
                                        If InStr(MemIDList, Trim(rdScoreA("BillMemID3"))) > 0 Then
                                            ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") * 0.33)
                                            TotalScore = TotalScore + CDec(rdScoreA("Other1") * 0.33)
                                            ReportValueCnt = ReportValueCnt + 0.33
                                            ReportTotalCnt = ReportTotalCnt + 0.33
                                        End If
                                    ElseIf rdScoreA("BillMemID1") IsNot DBNull.Value And rdScoreA("BillMemID2") IsNot DBNull.Value And rdScoreA("BillMemID3") IsNot DBNull.Value And rdScoreA("BillMemID4") IsNot DBNull.Value Then
                                        If InStr(MemIDList, Trim(rdScoreA("BillMemID1"))) > 0 Then
                                            ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") / 4)
                                            TotalScore = TotalScore + CDec(rdScoreA("Other1") / 4)
                                            ReportValueCnt = ReportValueCnt + CDec(1 / 4)
                                            ReportTotalCnt = ReportTotalCnt + CDec(1 / 4)
                                        End If
                                        If InStr(MemIDList, Trim(rdScoreA("BillMemID2"))) > 0 Then
                                            ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") / 4)
                                            TotalScore = TotalScore + CDec(rdScoreA("Other1") / 4)
                                            ReportValueCnt = ReportValueCnt + CDec(1 / 4)
                                            ReportTotalCnt = ReportTotalCnt + CDec(1 / 4)
                                        End If
                                        If InStr(MemIDList, Trim(rdScoreA("BillMemID3"))) > 0 Then
                                            ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") / 4)
                                            TotalScore = TotalScore + CDec(rdScoreA("Other1") / 4)
                                            ReportValueCnt = ReportValueCnt + CDec(1 / 4)
                                            ReportTotalCnt = ReportTotalCnt + CDec(1 / 4)
                                        End If
                                        If InStr(MemIDList, Trim(rdScoreA("BillMemID4"))) > 0 Then
                                            ReportValueScore = ReportValueScore + CDec(rdScoreA("Other1") / 4)
                                            TotalScore = TotalScore + CDec(rdScoreA("Other1") / 4)
                                            ReportValueCnt = ReportValueCnt + CDec(1 / 4)
                                            ReportTotalCnt = ReportTotalCnt + CDec(1 / 4)
                                        End If
                                    End If
                                End If

                            End While
                        End If
                        rdScoreA.Close()
                        
                        '抓攔停件數及績分
                        strScore2 = "select b.BillType1Score,b.A1Score,b.A2Score,b.A3Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4,a.TrafficAccidentType,a.BillBaseTypeID from BillBaseViewReward a "
                        strScore2 = strScore2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                        strScore2 = strScore2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                        strScore2 = strScore2 & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                        strScore2 = strScore2 & " and (a.BillMemID1 in (" & MemIDList & ") or a.BillMemID2 In (" & MemIDList & ") or a.BillMemID3 In (" & MemIDList & ") or a.BillMemID4 In (" & MemIDList & "))"
                        strScore2 = strScore2 & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule4)" & " and ((a.BillBaseTypeID='0' and a.BillTypeID='1') or (a.BillBaseTypeID='1'))"
                        strScore2 = strScore2 & " and (a.CarAddID<>8 or a.CarAddID is null) and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                        strScore2 = strScore2 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strScore2 = strScore2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdScore2 As New Data.OracleClient.OracleCommand(strScore2, conn)
                        Dim rdScore2 As Data.OracleClient.OracleDataReader = CmdScore2.ExecuteReader()
                        If rdScore2.HasRows Then
                            While rdScore2.Read()
                                If rdScore2("TrafficAccidentType") Is DBNull.Value Or Trim(rdScore2("BillBaseTypeID")) = "1" Then
                                    '===========非交通事故及行人攤販================
                                    If (rdScore2("BillType1Score")) Is DBNull.Value Then
                                        ReportValueScore = ReportValueScore
                                        TotalScore = TotalScore
                                    Else
                                        If rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") Is DBNull.Value And rdScore2("BillMemID3") Is DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                            StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score"))
                                            TotalScore = TotalScore + CDec(rdScore2("BillType1Score"))
                                            StopValueCnt = StopValueCnt + 1
                                            StopTotalCnt = StopTotalCnt + 1
                                        ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") Is DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID1"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 2)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 2)
                                                StopValueCnt = StopValueCnt + CDec(1 / 2)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 2)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID2"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 2)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 2)
                                                StopValueCnt = StopValueCnt + CDec(1 / 2)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 2)
                                            End If
                                        ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") IsNot DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID1"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") * 0.34)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") * 0.34)
                                                StopValueCnt = StopValueCnt + 0.34
                                                StopTotalCnt = StopTotalCnt + 0.34
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID2"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                StopValueCnt = StopValueCnt + 0.33
                                                StopTotalCnt = StopTotalCnt + 0.33
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID3"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                StopValueCnt = StopValueCnt + 0.33
                                                StopTotalCnt = StopTotalCnt + 0.33
                                            End If
                                    
                                        ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") IsNot DBNull.Value And rdScore2("BillMemID4") IsNot DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID1"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID2"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID3"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID4"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                            End If
                                        End If
                                    End If
                                    '=========================================
                                Else
                                    '***************交通事故類別**************
                                    If Trim(rdScore2("TrafficAccidentType")) = "1" Then
                                        If rdScore2("BillType1Score") > rdScore2("A1Score") Then
                                            If rdScore2("BillType1Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore2("BillType1Score"))
                                            End If
                                        Else
                                            If rdScore2("A1Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore2("A1Score"))
                                            End If
                                        End If
                                    ElseIf Trim(rdScore2("TrafficAccidentType")) = "2" Then
                                        If rdScore2("BillType1Score") > rdScore2("A2Score") Then
                                            If rdScore2("BillType1Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore2("BillType1Score"))
                                            End If
                                        Else
                                            If rdScore2("A2Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore2("A2Score"))
                                            End If
                                        End If
                                    ElseIf Trim(rdScore2("TrafficAccidentType")) = "3" Then
                                        If rdScore2("BillType1Score") > rdScore2("A3Score") Then
                                            If rdScore2("BillType1Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore2("BillType1Score"))
                                            End If
                                        Else
                                            If rdScore2("A3Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore2("A3Score"))
                                            End If
                                        End If
                                    End If
                                    
                                        
                                    If A1ScoreTmp = 0 Then
                                        StopValueScore = StopValueScore
                                        TotalScore = TotalScore
                                    Else
                                        If rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") Is DBNull.Value And rdScore2("BillMemID3") Is DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                            StopValueScore = StopValueScore + CDec(A1ScoreTmp)
                                            TotalScore = TotalScore + CDec(A1ScoreTmp)
                                            StopValueCnt = StopValueCnt + 1
                                            StopTotalCnt = StopTotalCnt + 1
                                        ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") Is DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID1"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp / 2)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 2)
                                                StopValueCnt = StopValueCnt + CDec(1 / 2)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 2)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID2"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp / 2)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 2)
                                                StopValueCnt = StopValueCnt + CDec(1 / 2)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 2)
                                            End If
                                        ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") IsNot DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID1"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp * 0.34)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp * 0.34)
                                                StopValueCnt = StopValueCnt + 0.34
                                                StopTotalCnt = StopTotalCnt + 0.34
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID2"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp * 0.33)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp * 0.33)
                                                StopValueCnt = StopValueCnt + 0.34
                                                StopTotalCnt = StopTotalCnt + 0.34
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID3"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp * 0.33)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp * 0.33)
                                                StopValueCnt = StopValueCnt + 0.34
                                                StopTotalCnt = StopTotalCnt + 0.34
                                            End If
                                        ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") IsNot DBNull.Value And rdScore2("BillMemID4") IsNot DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID1"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID2"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID3"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID4"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                            End If
                                        End If
                                    End If
                                    '*****************************************
                                End If

                            End While
                        End If
                        rdScore2.Close()
                        
                        '抓攔停件數及績分(拖吊)
                        strScoreB = "select b.Other1,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4 from BillBase a " & strPointT6Plus2
                        strScoreB = strScoreB & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                        strScoreB = strScoreB & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                        strScoreB = strScoreB & " and (a.BillMemID1 in (" & MemIDList & ") or a.BillMemID2 In (" & MemIDList & ") or a.BillMemID3 In (" & MemIDList & ") or a.BillMemID4 In (" & MemIDList & "))"
                        strScoreB = strScoreB & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule4)" & " and ((a.BillBaseTypeID='0' and a.BillTypeID='1') or (a.BillBaseTypeID='1'))"
                        strScoreB = strScoreB & " and a.CarAddID=8 and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                        strScoreB = strScoreB & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strScoreB = strScoreB & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')" & strPointT6Plus
                        Dim CmdScoreB As New Data.OracleClient.OracleCommand(strScoreB, conn)
                        Dim rsScoreB As Data.OracleClient.OracleDataReader = CmdScoreB.ExecuteReader()
                        If rsScoreB.HasRows Then
                            While rsScoreB.Read()
                                If (rsScoreB("Other1")) Is DBNull.Value Then
                                    ReportValueScore = ReportValueScore
                                    TotalScore = TotalScore
                                Else
                                    If rsScoreB("BillMemID1") IsNot DBNull.Value And rsScoreB("BillMemID2") Is DBNull.Value And rsScoreB("BillMemID3") Is DBNull.Value And rsScoreB("BillMemID4") Is DBNull.Value Then
                                        StopValueScore = StopValueScore + CDec(rsScoreB("Other1"))
                                        TotalScore = TotalScore + CDec(rsScoreB("Other1"))
                                        StopValueCnt = StopValueCnt + 1
                                        StopTotalCnt = StopTotalCnt + 1
                                    ElseIf rsScoreB("BillMemID1") IsNot DBNull.Value And rsScoreB("BillMemID2") IsNot DBNull.Value And rsScoreB("BillMemID3") Is DBNull.Value And rsScoreB("BillMemID4") Is DBNull.Value Then
                                        If InStr(MemIDList, Trim(rsScoreB("BillMemID1"))) > 0 Then
                                            StopValueScore = StopValueScore + CDec(rsScoreB("Other1") / 2)
                                            TotalScore = TotalScore + CDec(rsScoreB("Other1") / 2)
                                            StopValueCnt = StopValueCnt + CDec(1 / 2)
                                            StopTotalCnt = StopTotalCnt + CDec(1 / 2)
                                        End If
                                        If InStr(MemIDList, Trim(rsScoreB("BillMemID2"))) > 0 Then
                                            StopValueScore = StopValueScore + CDec(rsScoreB("Other1") / 2)
                                            TotalScore = TotalScore + CDec(rsScoreB("Other1") / 2)
                                            StopValueCnt = StopValueCnt + CDec(1 / 2)
                                            StopTotalCnt = StopTotalCnt + CDec(1 / 2)
                                        End If
                                    ElseIf rsScoreB("BillMemID1") IsNot DBNull.Value And rsScoreB("BillMemID2") IsNot DBNull.Value And rsScoreB("BillMemID3") IsNot DBNull.Value And rsScoreB("BillMemID4") Is DBNull.Value Then
                                        If InStr(MemIDList, Trim(rsScoreB("BillMemID1"))) > 0 Then
                                            StopValueScore = StopValueScore + CDec(rsScoreB("Other1") * 0.34)
                                            TotalScore = TotalScore + CDec(rsScoreB("Other1") * 0.34)
                                            StopValueCnt = StopValueCnt + 0.34
                                            StopTotalCnt = StopTotalCnt + 0.34
                                        End If
                                        If InStr(MemIDList, Trim(rsScoreB("BillMemID2"))) > 0 Then
                                            StopValueScore = StopValueScore + CDec(rsScoreB("Other1") * 0.33)
                                            TotalScore = TotalScore + CDec(rsScoreB("Other1") * 0.33)
                                            StopValueCnt = StopValueCnt + 0.33
                                            StopTotalCnt = StopTotalCnt + 0.33
                                        End If
                                        If InStr(MemIDList, Trim(rsScoreB("BillMemID3"))) > 0 Then
                                            StopValueScore = StopValueScore + CDec(rsScoreB("Other1") * 0.33)
                                            TotalScore = TotalScore + CDec(rsScoreB("Other1") * 0.33)
                                            StopValueCnt = StopValueCnt + 0.33
                                            StopTotalCnt = StopTotalCnt + 0.33
                                        End If
                                    ElseIf rsScoreB("BillMemID1") IsNot DBNull.Value And rsScoreB("BillMemID2") IsNot DBNull.Value And rsScoreB("BillMemID3") IsNot DBNull.Value And rsScoreB("BillMemID4") IsNot DBNull.Value Then
                                        If InStr(MemIDList, Trim(rsScoreB("BillMemID1"))) > 0 Then
                                            StopValueScore = StopValueScore + CDec(rsScoreB("Other1") / 4)
                                            TotalScore = TotalScore + CDec(rsScoreB("Other1") / 4)
                                            StopValueCnt = StopValueCnt + CDec(1 / 4)
                                            StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                        End If
                                        If InStr(MemIDList, Trim(rsScoreB("BillMemID2"))) > 0 Then
                                            StopValueScore = StopValueScore + CDec(rsScoreB("Other1") / 4)
                                            TotalScore = TotalScore + CDec(rsScoreB("Other1") / 4)
                                            StopValueCnt = StopValueCnt + CDec(1 / 4)
                                            StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                        End If
                                        If InStr(MemIDList, Trim(rsScoreB("BillMemID3"))) > 0 Then
                                            StopValueScore = StopValueScore + CDec(rsScoreB("Other1") / 4)
                                            TotalScore = TotalScore + CDec(rsScoreB("Other1") / 4)
                                            StopValueCnt = StopValueCnt + CDec(1 / 4)
                                            StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                        End If
                                        If InStr(MemIDList, Trim(rsScoreB("BillMemID4"))) > 0 Then
                                            StopValueScore = StopValueScore + CDec(rsScoreB("Other1") / 4)
                                            TotalScore = TotalScore + CDec(rsScoreB("Other1") / 4)
                                            StopValueCnt = StopValueCnt + CDec(1 / 4)
                                            StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                        End If
                                    End If
                                End If

                            End While
                        End If
                        rsScoreB.Close()
                    Else
                        '抓逕舉件數及績分
                        strScore = "select b.BillType2Score,b.A1Score,b.A2Score,b.A3Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4,a.TrafficAccidentType from BillBaseViewReward a "
                        strScore = strScore & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                        strScore = strScore & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                        strScore = strScore & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                        strScore = strScore & " and (a.BillMemID1 in (" & MemIDList & ") or a.BillMemID2 In (" & MemIDList & ") or a.BillMemID3 In (" & MemIDList & ") or a.BillMemID4 In (" & MemIDList & "))"
                        strScore = strScore & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule4)" & " and a.BillBaseTypeID='0' and a.BillTypeID='2'"
                        strScore = strScore & " and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                        strScore = strScore & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strScore = strScore & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdScore As New Data.OracleClient.OracleCommand(strScore, conn)
                        Dim rdScore As Data.OracleClient.OracleDataReader = CmdScore.ExecuteReader()
                        If rdScore.HasRows Then
                            While rdScore.Read()
                                If rdScore("TrafficAccidentType") Is DBNull.Value Then
                                    '===================非交通事故======================
                                    If (rdScore("BillType2Score")) Is DBNull.Value Then
                                        ReportValueScore = ReportValueScore
                                        TotalScore = TotalScore
                                    Else
                                        If rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") Is DBNull.Value And rdScore("BillMemID3") Is DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                            ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score"))
                                            TotalScore = TotalScore + CDec(rdScore("BillType2Score"))
                                            ReportValueCnt = ReportValueCnt + 1
                                            ReportTotalCnt = ReportTotalCnt + 1
                                        ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") Is DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore("BillMemID1"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 2)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 2)
                                                ReportValueCnt = ReportValueCnt + CDec(1 / 2)
                                                ReportTotalCnt = ReportTotalCnt + CDec(1 / 2)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID2"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 2)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 2)
                                                ReportValueCnt = ReportValueCnt + CDec(1 / 2)
                                                ReportTotalCnt = ReportTotalCnt + CDec(1 / 2)
                                            End If
                                        ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") IsNot DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore("BillMemID1"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") * 0.34)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") * 0.34)
                                                ReportValueCnt = ReportValueCnt + 0.34
                                                ReportTotalCnt = ReportTotalCnt + 0.34
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID2"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") * 0.33)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") * 0.33)
                                                ReportValueCnt = ReportValueCnt + 0.33
                                                ReportTotalCnt = ReportTotalCnt + 0.33
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID3"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") * 0.33)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") * 0.33)
                                                ReportValueCnt = ReportValueCnt + 0.33
                                                ReportTotalCnt = ReportTotalCnt + 0.33
                                            End If

                                        ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") IsNot DBNull.Value And rdScore("BillMemID4") IsNot DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore("BillMemID1"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID2"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID3"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID4"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(rdScore("BillType2Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore("BillType2Score") / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                        End If
                                    End If
                                    '========================================
                                Else
                                    '*******交通事故類別A1,2,3***************
                                    If Trim(rdScore("TrafficAccidentType")) = "1" Then
                                        If rdScore("BillType2Score") > rdScore("A1Score") Then
                                            If rdScore("BillType2Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore("BillType2Score"))
                                            End If
                                        Else
                                            If rdScore("A1Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore("A1Score"))
                                            End If
                                        End If
                                    ElseIf Trim(rdScore("TrafficAccidentType")) = "2" Then
                                        If rdScore("BillType2Score") > rdScore("A2Score") Then
                                            If rdScore("BillType2Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore("BillType2Score"))
                                            End If
                                        Else
                                            If rdScore("A2Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore("A2Score"))
                                            End If
                                        End If
                                    ElseIf Trim(rdScore("TrafficAccidentType")) = "3" Then
                                        If rdScore("BillType2Score") > rdScore("A3Score") Then
                                            If rdScore("BillType2Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore("BillType2Score"))
                                            End If
                                        Else
                                            If rdScore("A3Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore("A3Score"))
                                            End If
                                        End If
                                    End If
                                    
                                        
                                    If A1ScoreTmp = 0 Then
                                        ReportValueScore = ReportValueScore
                                        TotalScore = TotalScore
                                    Else
                                        If rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") Is DBNull.Value And rdScore("BillMemID3") Is DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                            ReportValueScore = ReportValueScore + CDec(A1ScoreTmp)
                                            TotalScore = TotalScore + CDec(A1ScoreTmp)
                                            ReportValueCnt = ReportValueCnt + 1
                                            ReportTotalCnt = ReportTotalCnt + 1
                                        ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") Is DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore("BillMemID1"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp / 2)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 2)
                                                ReportValueCnt = ReportValueCnt + CDec(1 / 2)
                                                ReportTotalCnt = ReportTotalCnt + CDec(1 / 2)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID2"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp / 2)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 2)
                                                ReportValueCnt = ReportValueCnt + CDec(1 / 2)
                                                ReportTotalCnt = ReportTotalCnt + CDec(1 / 2)
                                            End If
                                        ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") IsNot DBNull.Value And rdScore("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore("BillMemID1"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp * 0.33)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp * 0.34)
                                                ReportValueCnt = ReportValueCnt + 0.34
                                                ReportTotalCnt = ReportTotalCnt + 0.34
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID2"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp * 0.33)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp * 0.33)
                                                ReportValueCnt = ReportValueCnt + 0.33
                                                ReportTotalCnt = ReportTotalCnt + 0.33
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID3"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp * 0.33)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp * 0.33)
                                                ReportValueCnt = ReportValueCnt + 0.33
                                                ReportTotalCnt = ReportTotalCnt + 0.33
                                            End If
                                        ElseIf rdScore("BillMemID1") IsNot DBNull.Value And rdScore("BillMemID2") IsNot DBNull.Value And rdScore("BillMemID3") IsNot DBNull.Value And rdScore("BillMemID4") IsNot DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore("BillMemID1"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID2"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID3"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore("BillMemID4"))) > 0 Then
                                                ReportValueScore = ReportValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                ReportValueCnt = ReportValueCnt + 0.25
                                                ReportTotalCnt = ReportTotalCnt + 0.25
                                            End If
                                        End If
                                    End If
                                    '****************************************
                                End If
                            End While
                        End If
                        rdScore.Close()
                    
                        '抓攔停件數及績分
                        strScore2 = "select b.BillType1Score,b.A1Score,b.A2Score,b.A3Score,a.BillMemID1,a.BillMemID2,a.BillMemID3,a.BillMemID4,a.TrafficAccidentType,a.BillBaseTypeID from BillBaseViewReward a "
                        strScore2 = strScore2 & ",(select distinct LawVersion,LawItem,BillType1Score,BillType2Score,A1Score,A2Score,A3Score,CountyOrNpa,IsUsed from LawScore"
                        strScore2 = strScore2 & " where IsUsed=1 and CountyOrNpa=" & Trim(Request("theCountyOrNpa")) & ") b"
                        strScore2 = strScore2 & " where b.LawItem='" & Trim(rdLaw("ItemID")) & "'"
                        strScore2 = strScore2 & " and (a.BillMemID1 in (" & MemIDList & ") or a.BillMemID2 In (" & MemIDList & ") or a.BillMemID3 In (" & MemIDList & ") or a.BillMemID4 In (" & MemIDList & "))"
                        strScore2 = strScore2 & " and (b.LawItem=a.Rule1 or b.LawItem=a.Rule2 or b.LawItem=a.Rule3 or b.LawItem=a.Rule4)" & " and ((a.BillBaseTypeID='0' and a.BillTypeID='1') or (a.BillBaseTypeID='1'))"
                        strScore2 = strScore2 & " and b.LawVersion=a.RuleVer and a.RecordStateID=0"
                        strScore2 = strScore2 & " and a." & theDateType & " between TO_DATE('" & gOutDT(Trim(Request("Date1"))) & " 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
                        strScore2 = strScore2 & " and TO_DATE('" & gOutDT(Trim(Request("Date2"))) & " 23:59:59','YYYY/MM/DD/HH24/MI/SS')"
                        Dim CmdScore2 As New Data.OracleClient.OracleCommand(strScore2, conn)
                        Dim rdScore2 As Data.OracleClient.OracleDataReader = CmdScore2.ExecuteReader()
                        If rdScore2.HasRows Then
                            While rdScore2.Read()
                                If rdScore2("TrafficAccidentType") Is DBNull.Value Or Trim(rdScore2("BillBaseTypeID")) = "1" Then
                                    '===========非交通事故及行人攤販================
                                    If (rdScore2("BillType1Score")) Is DBNull.Value Then
                                        ReportValueScore = ReportValueScore
                                        TotalScore = TotalScore
                                    Else
                                        If rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") Is DBNull.Value And rdScore2("BillMemID3") Is DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                            StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score"))
                                            TotalScore = TotalScore + CDec(rdScore2("BillType1Score"))
                                            StopValueCnt = StopValueCnt + 1
                                            StopTotalCnt = StopTotalCnt + 1
                                        ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") Is DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID1"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 2)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 2)
                                                StopValueCnt = StopValueCnt + CDec(1 / 2)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 2)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID2"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 2)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 2)
                                                StopValueCnt = StopValueCnt + CDec(1 / 2)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 2)
                                            End If
                                        ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") IsNot DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID1"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") * 0.34)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") * 0.34)
                                                StopValueCnt = StopValueCnt + 0.34
                                                StopTotalCnt = StopTotalCnt + 0.34
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID2"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                StopValueCnt = StopValueCnt + 0.33
                                                StopTotalCnt = StopTotalCnt + 0.33
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID3"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") * 0.33)
                                                StopValueCnt = StopValueCnt + 0.33
                                                StopTotalCnt = StopTotalCnt + 0.33
                                            End If
                                        ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") IsNot DBNull.Value And rdScore2("BillMemID4") IsNot DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID1"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                StopValueCnt = StopValueCnt + 0.25
                                                StopTotalCnt = StopTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID2"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                StopValueCnt = StopValueCnt + 0.25
                                                StopTotalCnt = StopTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID3"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                StopValueCnt = StopValueCnt + 0.25
                                                StopTotalCnt = StopTotalCnt + 0.25
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID4"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(rdScore2("BillType1Score") / 4)
                                                TotalScore = TotalScore + CDec(rdScore2("BillType1Score") / 4)
                                                StopValueCnt = StopValueCnt + 0.25
                                                StopTotalCnt = StopTotalCnt + 0.25
                                            End If

                                        End If
                                    End If
                                    '=========================================
                                Else
                                    '*******交通事故類別A1,2,3****************
                                    If Trim(rdScore2("TrafficAccidentType")) = "1" Then
                                        If rdScore2("BillType1Score") > rdScore2("A1Score") Then
                                            If rdScore2("BillType1Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore2("BillType1Score"))
                                            End If
                                        Else
                                            If rdScore2("A1Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore2("A1Score"))
                                            End If
                                        End If
                                    ElseIf Trim(rdScore2("TrafficAccidentType")) = "2" Then
                                        If rdScore2("BillType1Score") > rdScore2("A2Score") Then
                                            If rdScore2("BillType1Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore2("BillType1Score"))
                                            End If
                                        Else
                                            If rdScore2("A2Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore2("A2Score"))
                                            End If
                                        End If
                                    ElseIf Trim(rdScore2("TrafficAccidentType")) = "3" Then
                                        If rdScore2("BillType1Score") > rdScore2("A3Score") Then
                                            If rdScore2("BillType1Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore2("BillType1Score"))
                                            End If
                                        Else
                                            If rdScore2("A3Score") Is DBNull.Value Then
                                                A1ScoreTmp = 0
                                            Else
                                                A1ScoreTmp = CDec(rdScore2("A3Score"))
                                            End If
                                        End If
                                    End If
                                    
                                        
                                    If A1ScoreTmp = 0 Then
                                        StopValueScore = StopValueScore
                                        TotalScore = TotalScore
                                    Else
                                        If rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") Is DBNull.Value And rdScore2("BillMemID3") Is DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                            StopValueScore = StopValueScore + CDec(A1ScoreTmp)
                                            TotalScore = TotalScore + CDec(A1ScoreTmp)
                                            StopValueCnt = StopValueCnt + 1
                                            StopTotalCnt = StopTotalCnt + 1
                                        ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") Is DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID1"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp / 2)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 2)
                                                StopValueCnt = StopValueCnt + CDec(1 / 2)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 2)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID2"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp / 2)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 2)
                                                StopValueCnt = StopValueCnt + CDec(1 / 2)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 2)
                                            End If
                                        ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") IsNot DBNull.Value And rdScore2("BillMemID4") Is DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID1"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp * 0.34)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp * 0.34)
                                                StopValueCnt = StopValueCnt + 0.34
                                                StopTotalCnt = StopTotalCnt + 0.34
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID2"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp * 0.33)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp * 0.33)
                                                StopValueCnt = StopValueCnt + 0.33
                                                StopTotalCnt = StopTotalCnt + 0.33
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID3"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp * 0.33)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp * 0.33)
                                                StopValueCnt = StopValueCnt + 0.33
                                                StopTotalCnt = StopTotalCnt + 0.33
                                            End If
                                        ElseIf rdScore2("BillMemID1") IsNot DBNull.Value And rdScore2("BillMemID2") IsNot DBNull.Value And rdScore2("BillMemID3") IsNot DBNull.Value And rdScore2("BillMemID4") IsNot DBNull.Value Then
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID1"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID2"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID3"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                            End If
                                            If InStr(MemIDList, Trim(rdScore2("BillMemID4"))) > 0 Then
                                                StopValueScore = StopValueScore + CDec(A1ScoreTmp / 4)
                                                TotalScore = TotalScore + CDec(A1ScoreTmp / 4)
                                                StopValueCnt = StopValueCnt + CDec(1 / 4)
                                                StopTotalCnt = StopTotalCnt + CDec(1 / 4)
                                            End If
                                        End If
                                    End If
                                    '*********************************************
                                End If
                            End While
                        End If
                        rdScore2.Close()
                    End If
                    
                    
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
                    strDeatil = strDeatil & "<td width=""8%"" align=""right"">" & Format(ReportValueCnt, "##,##0.00") & "</td>"
                    strDeatil = strDeatil & "<td width=""8%"" align=""right"">" & Format(StopValueCnt, "##,##0.00") & "</td>"
                    strDeatil = strDeatil & "<td width=""10%"" align=""right"">" & Format(ReportValueScore + StopValueScore, "##,##0.00") & "</td>"
                    strDeatil = strDeatil & "</tr>"
                
                        
                End While
            End If
            rdLaw.Close()
            If ReportTotalCnt + StopTotalCnt > 0 Then
                Dim sys_Mem1, sys_CreditID As String
                If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Then
                    Dim strMem1 = "select * from Memberdata where LoginID='" & Trim(rdMem("LoginID")) & "' and RecordStateID=0"
                    Dim CmdMem1 As New Data.OracleClient.OracleCommand(strMem1, conn)
                    Dim rdMem1 As Data.OracleClient.OracleDataReader = CmdMem1.ExecuteReader()
                    If rdMem1.HasRows Then
                        rdMem1.Read()
                        sys_Mem1 = Trim(rdMem1("ChName"))
                        If rdMem1("CreditID") IsNot DBNull.Value Then
                            sys_CreditID = Trim(rdMem1("CreditID"))
                        Else
                            sys_CreditID = ""
                        End If
                    End If
                    rdCity.Close()
                Else
                    sys_Mem1 = rdMem("ChName")
                    If rdMem("CreditID") IsNot DBNull.Value Then
                        sys_CreditID = Trim(rdMem("CreditID"))
                    Else
                        sys_CreditID = ""
                    End If
                End If

                
                Response.Write("<tr><td align=""center"" height=""35"" colspan=""5""><strong><span class=""style1"">個 別 員 警 舉 發 件 數 暨 績 分 統 計 明 細 表</span></strong></td></tr>")
                Response.Write("<tr><td colspan=""5"">舉發員警：&nbsp; " & rdMem("LoginID") & "&nbsp; " & sys_Mem1 & "&nbsp; " & sys_CreditID & "</td></tr>")
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
                Response.Write("<td width=""8%"" align=""right"">" & Format(ReportTotalCnt, "##,##0.00") & "</td>")
                Response.Write("<td width=""8%"" align=""right"">" & Format(StopTotalCnt, "##,##0.00") & "</td>")
                Response.Write("<td width=""10%"" align=""right"">" & Format(TotalScore, "##,##0.00") & "</td>")
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
